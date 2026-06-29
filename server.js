import makeWASocket, { useMultiFileAuthState, DisconnectReason, fetchLatestBaileysVersion } from '@whiskeysockets/baileys';
import { spawn, execSync } from 'child_process';
import pino from 'pino';
import express from 'express';
import cors from 'cors';
import QRCode from 'qrcode';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import nodemailer from 'nodemailer';
import XLSX from 'xlsx';
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());

// Allow localtunnel requests to bypass the tunnel password/reminder page
app.use((req, res, next) => {
    res.setHeader('bypass-tunnel-reminder', 'true');
    next();
});

// Serve front-end static files (index.html, style.css, app.js, etc.)
app.use(express.static(__dirname));

const authPath = path.join(__dirname, 'session_auth_info');

let sock = null;
let qrCodeBase64 = null;
let connectionStatus = 'DISCONNECTED';
let incomingMessages = [];
let startTime = Date.now(); // used to skip history-load appends on startup
let mySentMessageIds = new Set(); // Track message IDs sent by the CRM to avoid self-messaging loops

// Helper to extract text from various WhatsApp message formats
function getMessageText(message) {
    if (!message) return '';
    if (message.conversation) return message.conversation;
    if (message.extendedTextMessage?.text) return message.extendedTextMessage.text;
    if (message.imageMessage) return message.imageMessage.caption || '[Image]';
    if (message.videoMessage) return message.videoMessage.caption || '[Video]';
    if (message.audioMessage) return '[Audio/Voice Note]';
    if (message.documentMessage) return message.documentMessage.fileName || '[Document]';
    if (message.stickerMessage) return '[Sticker]';
    if (message.reactionMessage) return ''; // ignore reactions
    if (message.protocolMessage) return ''; // ignore protocol messages
    if (message.ephemeralMessage) return getMessageText(message.ephemeralMessage.message);
    if (message.viewOnceMessage) return getMessageText(message.viewOnceMessage.message);
    if (message.viewOnceMessageV2) return getMessageText(message.viewOnceMessageV2.message);
    if (message.buttonsResponseMessage) return message.buttonsResponseMessage.selectedDisplayText || '[Button Response]';
    if (message.listResponseMessage) return message.listResponseMessage.title || '[List Response]';
    return '';
}

// Check if a JID is a real 1-on-1 chat (not group/broadcast/newsletter)
function isDirectChat(jid) {
    if (!jid) return false;
    if (jid === 'status@broadcast') return false;
    if (jid.endsWith('@g.us')) return false;
    if (jid.endsWith('@broadcast')) return false;
    if (jid.endsWith('@newsletter')) return false;
    // Accept both @s.whatsapp.net (standard) and @lid (new WhatsApp linked ID format)
    return jid.endsWith('@s.whatsapp.net') || jid.endsWith('@lid');
}

// Helper to resolve LID JIDs (like 221946897764392@lid) to phone number JIDs (like 97430738570@s.whatsapp.net) using Baileys session files
function resolveJid(jid) {
    if (!jid) return jid;
    if (jid.endsWith('@lid')) {
        const lidId = jid.split('@')[0];
        const mappingPath = path.join(authPath, `lid-mapping-${lidId}_reverse.json`);
        if (fs.existsSync(mappingPath)) {
            try {
                const pn = JSON.parse(fs.readFileSync(mappingPath, 'utf8'));
                if (pn) {
                    const resolved = `${pn}@s.whatsapp.net`;
                    console.log(`[LID RESOLVER] Successfully resolved LID ${jid} to Phone JID ${resolved}`);
                    return resolved;
                }
            } catch (e) {
                console.warn(`[LID RESOLVER] Failed to parse mapping file for ${lidId}:`, e.message);
            }
        }
    }
    return jid;
}

// Initialize and Connect WhatsApp Socket
async function connectToWhatsApp() {
    console.log('Initializing WhatsApp Baileys socket...');
    startTime = Date.now(); // reset so history-load window is fresh

    const { state, saveCreds } = await useMultiFileAuthState(authPath);
    const { version } = await fetchLatestBaileysVersion();
    console.log(`Using Baileys WA version: ${version.join('.')}`);

    sock = makeWASocket({
        version,
        auth: state,
        logger: pino({ level: 'silent' }),
        getMessage: async (key) => {
            // Required: return empty message for retry requests
            return { conversation: '' };
        },
        syncFullHistory: false,
        markOnlineOnConnect: true,
        generateHighQualityLinkPreview: false,
        defaultQueryTimeoutMs: 60000,
        connectTimeoutMs: 60000,
        keepAliveIntervalMs: 15000,
        retryRequestDelayMs: 500,
        maxMsgRetryCount: 5
    });

    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update;

        if (qr) {
            connectionStatus = 'CONNECTING';
            try {
                qrCodeBase64 = await QRCode.toDataURL(qr);
                console.log('New QR code generated — scan it in the CRM browser.');
            } catch (err) {
                console.error('Failed to generate QR code:', err);
            }
        }

        if (connection === 'close') {
            const statusCode = lastDisconnect?.error?.output?.statusCode;
            const isLoggedOut = statusCode === DisconnectReason.loggedOut;
            console.log(`Connection closed [${statusCode}] — LoggedOut: ${isLoggedOut}`);
            qrCodeBase64 = null;
            connectionStatus = 'DISCONNECTED';

            if (isLoggedOut) {
                // Session is invalid — clear auth files and show a fresh QR
                console.log('WhatsApp session logged out. Clearing session files and generating new QR...');
                try {
                    const files = fs.readdirSync(authPath);
                    for (const f of files) fs.rmSync(path.join(authPath, f), { recursive: true, force: true });
                    console.log('Session files cleared. Reconnecting for new QR...');
                } catch (e) { console.warn('Could not clear session files:', e.message); }
            }
            // Always reconnect so a new QR is shown
            setTimeout(connectToWhatsApp, 3000);
        } else if (connection === 'open') {
            console.log('✅ WhatsApp CONNECTED — ready to send and receive messages!');
            qrCodeBase64 = null;
            connectionStatus = 'CONNECTED';
        }
    });

    // ─── LIVE MESSAGES ─────────────────────────────────────────────────────────
    sock.ev.on('messages.upsert', async ({ messages, type }) => {
        console.log(`[messages.upsert] type=${type}, count=${messages.length}`);

        // Process both 'notify' (real-time) and 'append' (some Baileys versions use this for live messages)
        // Skip 'append' only if it came in during the first 10 seconds (history sync window)
        const isHistoryLoad = (type === 'append') && (Date.now() - startTime < 10000);
        if (isHistoryLoad) {
            console.log(`  ℹ Skipping history-load append batch`);
            return;
        }

        for (const msg of messages) {
            if (!msg?.key) continue;

            console.log("----------------- INCOMING MESSAGE METADATA -----------------");
            console.log("msg keys:", Object.keys(msg));
            console.log("msg.key keys:", Object.keys(msg.key));
            console.log("remoteJid:", msg.key.remoteJid);
            if (msg.sender) console.log("sender:", msg.sender);
            if (msg.senderPn) console.log("senderPn:", msg.senderPn);
            if (msg.key.participant) console.log("participant:", msg.key.participant);
            if (msg.key.remoteJidAlt) console.log("remoteJidAlt:", msg.key.remoteJidAlt);
            if (msg.key.participantAlt) console.log("participantAlt:", msg.key.participantAlt);
            console.log("-------------------------------------------------------------");

            const fromJid = resolveJid(msg.key.remoteJid);
            const isFromMe = msg.key.fromMe;

            console.log(`  → fromMe=${isFromMe} jid=${fromJid}`);

            // Filter out non-direct JIDs (groups, broadcasts, newsletters)
            if (!isDirectChat(fromJid)) {
                console.log(`  ✗ Filtered out: ${fromJid}`);
                continue;
            }

            // Skip messages sent BY ME (outgoing) — only want incoming
            // But allow self-messages (messaging yourself) for testing purposes
            const myId = sock?.user?.id ? sock.user.id.split(':')[0].replace(/[^0-9]/g, '') : '';
            const myLid = sock?.user?.lid ? sock.user.lid.split(':')[0].replace(/[^0-9]/g, '') : '';
            const targetId = fromJid.split(':')[0].replace(/[^0-9]/g, '');
            const isSelfMessage = (myId && targetId && myId === targetId) || (myLid && targetId && myLid === targetId);

            // Skip messages sent BY ME (outgoing) — unless it's a self-message NOT sent by the CRM itself
            const isCrmSent = msg.key.id && mySentMessageIds.has(msg.key.id);
            if (isFromMe && (!isSelfMessage || isCrmSent)) {
                console.log(`  ℹ Own outgoing message — skipping`);
                continue;
            }

            const text = getMessageText(msg.message);
            if (!text) {
                console.log(`  ✗ No text content`);
                continue;
            }

            const pushName = msg.pushName || 'WhatsApp Contact';
            const rawId = fromJid.split('@')[0];
            const phoneNumber = msg.key.participant ? msg.key.participant.split('@')[0] : rawId;

            console.log(`  ✅ INCOMING from ${pushName} (+${phoneNumber}): ${text.substring(0, 100)}`);

            const exists = incomingMessages.some(m => m.id === msg.key.id);
            if (!exists) {
                incomingMessages.push({
                    id: msg.key.id,
                    from: fromJid, // Store full JID (e.g. 265463808393312@lid or 97430544802@s.whatsapp.net)
                    name: pushName,
                    text: text,
                    time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                    timestamp: Date.now()
                });
                if (incomingMessages.length > 150) incomingMessages.shift();
            }
        }
    });

    // ─── HISTORY SYNC ──────────────────────────────────────────────────────────
    sock.ev.on('messaging-history.set', ({ messages }) => {
        const count = messages ? messages.length : 0;
        console.log(`History sync: ${count} messages`);
        if (!messages || count === 0) return;

        let added = 0;
        for (const msg of messages) {
            if (!msg?.key || msg.key.fromMe) continue;
            const fromJid = msg.key.remoteJid;
            if (!isDirectChat(fromJid)) continue;

            const text = getMessageText(msg.message);
            if (!text) continue;

            const exists = incomingMessages.some(m => m.id === msg.key.id);
            if (!exists) {
                incomingMessages.push({
                    id: msg.key.id,
                    from: fromJid, // Store full JID (e.g. 265463808393312@lid or 97430544802@s.whatsapp.net)
                    name: msg.pushName || 'WhatsApp Contact',
                    text: text,
                    time: new Date((msg.messageTimestamp || 0) * 1000).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                    timestamp: msg.messageTimestamp || 0
                });
                added++;
            }
        }

        // Sort by timestamp and keep last 150
        incomingMessages.sort((a, b) => a.timestamp - b.timestamp);
        if (incomingMessages.length > 150) {
            incomingMessages = incomingMessages.slice(-150);
        }
        console.log(`History sync: added ${added} messages. Total: ${incomingMessages.length}`);
    });
}

// ─── REST API ───────────────────────────────────────────────────────────────

// Google Sheets CSV proxy — fetches a public Google Sheets export URL server-side to avoid CORS
app.get('/api/google-sheets-import', async (req, res) => {
    const { csvUrl } = req.query;
    if (!csvUrl) return res.status(400).json({ error: 'Missing csvUrl parameter' });

    try {
        const response = await fetch(csvUrl);
        if (!response.ok) {
            return res.status(502).json({ error: `Google Sheets responded with HTTP ${response.status}. Make sure the sheet is publicly accessible (Anyone with the link can view).` });
        }
        const csvText = await response.text();
        res.setHeader('Content-Type', 'text/csv');
        res.send(csvText);
    } catch (err) {
        console.error('Google Sheets proxy error:', err);
        res.status(500).json({ error: 'Failed to fetch Google Sheets: ' + err.message });
    }
});

// WhatsApp connection status + QR code
app.get('/api/status', (req, res) => {
    res.json({ status: connectionStatus, qr: qrCodeBase64 });
});

// All received messages
app.get('/api/messages', (req, res) => {
    // Dynamically resolve any unresolved LID JIDs in incomingMessages
    incomingMessages.forEach(msg => {
        if (msg.from && msg.from.endsWith('@lid')) {
            const resolved = resolveJid(msg.from);
            if (resolved && !resolved.endsWith('@lid')) {
                msg.from = resolved;
            }
        }
    });
    res.json({ messages: incomingMessages });
});

// Webhook receiver endpoint (System Webhook Receiver)
app.post('/api/webhooks/incoming', (req, res) => {
    console.log(`\n🔔 [WEBHOOK RECEIVER] Received event: "${req.body?.event || 'Unknown'}"`);
    console.log(`Payload: ${JSON.stringify(req.body, null, 2)}\n`);
    res.json({ success: true, received: true });
});

// ── TEST ENDPOINT: inject a fake incoming message (for testing without 2nd phone)
app.post('/api/test-incoming', (req, res) => {
    const testMsg = {
        id: 'test_' + Date.now(),
        from: req.body?.from || '97450000001',
        name: req.body?.name || 'Test Customer',
        text: req.body?.text || 'Hello! I need a quote for AC maintenance.',
        time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
        timestamp: Date.now()
    };
    incomingMessages.push(testMsg);
    console.log(`[TEST] Injected fake incoming message from ${testMsg.name} (${testMsg.from}): "${testMsg.text}"`);
    res.json({ success: true, message: testMsg });
});

// Send a WhatsApp message
app.post('/api/send', async (req, res) => {
    const { to, text } = req.body;

    if (!to || !text) {
        return res.status(400).json({ error: "Missing 'to' or 'text' parameters." });
    }
    if (connectionStatus !== 'CONNECTED' || !sock) {
        return res.status(503).json({ error: 'WhatsApp not connected. Please scan the QR code first.' });
    }

    try {
        let jid = to;
        if (!jid.includes('@')) {
            const cleanNum = to.replace(/[^0-9]/g, '');
            if (cleanNum.length < 7) {
                return res.status(400).json({ error: 'Invalid phone number — too short.' });
            }
            jid = `${cleanNum}@s.whatsapp.net`;
        }
        console.log(`Sending to ${jid}: ${text.substring(0, 80)}`);
        const sentMsg = await sock.sendMessage(jid, { text });
        if (sentMsg?.key?.id) {
            mySentMessageIds.add(sentMsg.key.id);
            if (mySentMessageIds.size > 1000) {
                const first = mySentMessageIds.values().next().value;
                mySentMessageIds.delete(first);
            }
        }
        res.json({ success: true });
    } catch (err) {
        console.error('Send failed:', err);
        res.status(500).json({ error: err.message });
    }
});

// Disconnect and reset session (forces fresh QR scan)
app.post('/api/disconnect', async (req, res) => {
    console.log('Disconnecting WhatsApp session...');
    try {
        if (sock && connectionStatus === 'CONNECTED') {
            await sock.logout().catch(() => {});
        }
        sock = null;
        if (fs.existsSync(authPath)) {
            fs.rmSync(authPath, { recursive: true, force: true });
        }
        connectionStatus = 'DISCONNECTED';
        qrCodeBase64 = null;
        incomingMessages = [];
        connectToWhatsApp();
        res.json({ success: true });
    } catch (err) {
        console.error('Disconnect error:', err);
        res.status(500).json({ error: err.message });
    }
});

// Send email via SMTP
app.post('/api/send-email', async (req, res) => {
    const { to, subject, body, smtp } = req.body;
    if (!to || !subject || !body || !smtp?.host || !smtp?.port || !smtp?.user || !smtp?.pass) {
        return res.status(400).json({ error: 'Missing required fields or SMTP config.' });
    }
    try {
        console.log(`Sending email to ${to} via ${smtp.host}:${smtp.port}...`);
        const transporter = nodemailer.createTransport({
            host: smtp.host,
            port: parseInt(smtp.port),
            secure: smtp.port == 465,
            auth: { user: smtp.user, pass: smtp.pass },
            tls: { rejectUnauthorized: false }
        });
        const info = await transporter.sendMail({
            from: smtp.user,
            to,
            subject,
            text: body
        });
        console.log(`Email sent! ID: ${info.messageId}`);
        res.json({ success: true, messageId: info.messageId });
    } catch (err) {
        console.error('Email failed:', err);
        res.status(500).json({ error: err.message });
    }
});

// Fetch email inbox messages via IMAP
app.post('/api/fetch-emails', async (req, res) => {
    const { smtp } = req.body;
    
    if (!smtp?.host || !smtp?.user || !smtp?.pass) {
        return res.status(400).json({ error: 'Missing required credentials or SMTP/IMAP config.' });
    }

    let imapHost = smtp.host;
    if (imapHost === 'smtp.gmail.com') {
        imapHost = 'imap.gmail.com';
    } else if (imapHost === 'smtp.office365.com') {
        imapHost = 'outlook.office365.com';
    } else if (imapHost === 'smtp.mail.yahoo.com') {
        imapHost = 'imap.mail.yahoo.com';
    } else if (imapHost.startsWith('smtp.')) {
        imapHost = 'imap.' + imapHost.substring(5);
    }

    console.log(`Connecting to IMAP server ${imapHost}:993 for user ${smtp.user}...`);

    const client = new ImapFlow({
        host: imapHost,
        port: 993,
        secure: true,
        auth: {
            user: smtp.user,
            pass: smtp.pass
        },
        logger: false
    });

    try {
        await client.connect();
        let lock = await client.getMailboxLock('INBOX');
        try {
            let status = await client.status('INBOX', { messages: true });
            const totalMessages = status.messages;
            const messages = [];

            if (totalMessages > 0) {
                // Fetch the latest 10 messages (or fewer if total is less than 10)
                const fetchCount = Math.min(10, totalMessages);
                const startRange = Math.max(1, totalMessages - fetchCount + 1);
                const range = `${startRange}:${totalMessages}`;

                for await (let msg of client.fetch(range, { envelope: true, source: true })) {
                    const parsed = await simpleParser(msg.source);
                    
                    // Simple text extraction and sanitization
                    let textBody = parsed.text || "";
                    if (!textBody && parsed.html) {
                        textBody = parsed.html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
                    }

                    // Extract first name or display name
                    let senderName = msg.envelope.from[0]?.name || 'Unknown Sender';
                    if (!msg.envelope.from[0]?.name && msg.envelope.from[0]?.address) {
                        senderName = msg.envelope.from[0].address.split('@')[0];
                        // Capitalize
                        senderName = senderName.charAt(0).toUpperCase() + senderName.slice(1);
                    }

                    messages.push({
                        id: msg.uid || msg.seq,
                        senderName: senderName,
                        senderEmail: msg.envelope.from[0]?.address || '',
                        subject: msg.envelope.subject || '(No Subject)',
                        date: msg.envelope.date ? msg.envelope.date.toISOString().replace('T', ' ').substring(0, 16) : new Date().toISOString().replace('T', ' ').substring(0, 16),
                        body: textBody,
                        read: true
                    });
                }
            }

            // Return in reverse chronological order (newest first)
            res.json({ success: true, emails: messages.reverse() });
        } finally {
            lock.release();
        }
        await client.logout();
    } catch (err) {
        console.error('IMAP fetch failed:', err);
        res.status(500).json({ error: 'Failed to fetch emails: ' + err.message });
    }
});

// ── EXCEL SPREADSHEET SYSTEM ROUTING ───────────────────────────────────────
app.get('/api/excel/read', (req, res) => {
    const xlsPath = path.join(__dirname, 'crm.xls');
    
    const defaultData = [];

    if (!fs.existsSync(xlsPath)) {
        try {
            console.log("crm.xls does not exist. Initializing with default data...");
            const worksheet = XLSX.utils.json_to_sheet(defaultData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
            XLSX.writeFile(workbook, xlsPath, { bookType: 'biff8' });
            return res.json({ success: true, data: defaultData });
        } catch (err) {
            console.error("Failed to initialize crm.xls:", err);
            return res.json({ success: true, data: defaultData, warning: "Initial write failed, returning defaults." });
        }
    }

    try {
        const workbook = XLSX.readFile(xlsPath);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        if (!sheet || !sheet['!ref']) {
            try {
                const worksheet = XLSX.utils.json_to_sheet(defaultData);
                const workbook2 = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook2, worksheet, "Sheet1");
                XLSX.writeFile(workbook2, xlsPath, { bookType: 'biff8' });
                return res.json({ success: true, data: defaultData });
            } catch (werr) {
                return res.json({ success: true, data: defaultData });
            }
        }
        const data = XLSX.utils.sheet_to_json(sheet);
        res.json({ success: true, data });
    } catch (err) {
        console.error('Error reading excel file:', err);
        res.status(500).json({ error: 'Failed to read Excel file: ' + err.message });
    }
});

app.post('/api/excel/save', (req, res) => {
    const { data } = req.body;
    if (!Array.isArray(data)) {
        return res.status(400).json({ error: "Invalid parameters: 'data' must be an array." });
    }

    const xlsPath = path.join(__dirname, 'crm.xls');
    const tmpPath = path.join(__dirname, 'crm_tmp.xls');
    const csvPath = path.join(__dirname, 'crm.csv');

    try {
        console.log(`Saving ${data.length} rows to crm.xls...`);
        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

        // Write to a temp file first to avoid EBUSY if crm.xls is open in Excel
        XLSX.writeFile(workbook, tmpPath, { bookType: 'biff8' });

        // Rename temp → crm.xls (atomic swap)
        if (fs.existsSync(xlsPath)) fs.unlinkSync(xlsPath);
        fs.renameSync(tmpPath, xlsPath);

        // Also write a CSV copy for easy access from any app
        try {
            const csvContent = XLSX.utils.sheet_to_csv(worksheet);
            fs.writeFileSync(csvPath, csvContent, 'utf8');
        } catch (csvErr) {
            console.warn('CSV write warning (non-fatal):', csvErr.message);
        }

        console.log(`Saved successfully to ${xlsPath}`);
        res.json({ success: true });
    } catch (err) {
        console.error('Failed to write Excel file:', err);
        // Clean up temp file if it exists
        try { if (fs.existsSync(tmpPath)) fs.unlinkSync(tmpPath); } catch (_) {}
        if (err.code === 'EBUSY') {
            return res.status(500).json({ error: 'The crm.xls file is currently open in Microsoft Excel. Please close it in Excel first to save changes.' });
        }
        res.status(500).json({ error: 'Failed to write Excel file: ' + err.message });
    }
});


// ─── TUNNEL MANAGEMENT ─────────────────────────────────────────────────────────
let tunnelProcess = null;
let tunnelUrl = null;

app.get('/api/tunnel-status', (req, res) => {
    res.json({
        running: tunnelProcess !== null && !tunnelProcess.killed,
        url: tunnelUrl || null
    });
});

app.post('/api/start-tunnel', async (req, res) => {
    // Kill existing tunnel if running
    if (tunnelProcess && !tunnelProcess.killed) {
        tunnelProcess.kill();
        tunnelProcess = null;
        tunnelUrl = null;
    }

    console.log('[Tunnel] Starting localtunnel...');
    res.json({ success: true, message: 'Tunnel starting... check /api/tunnel-status in a few seconds.' });

    // Start tunnel as background process
    const tunnel = spawn('npx', ['localtunnel', '--port', '3000', '--subdomain', 'smartvision-crm'], {
        cwd: __dirname,
        shell: true,
        stdio: ['ignore', 'pipe', 'pipe']
    });
    tunnelProcess = tunnel;

    tunnel.stdout.on('data', (data) => {
        const line = data.toString();
        console.log('[Tunnel]', line.trim());
        const match = line.match(/your url is: (.+)/);
        if (match) {
            tunnelUrl = match[1].trim();
            console.log('[Tunnel] Public URL:', tunnelUrl);

            // Update config.json with the new URL
            const configPath = path.join(__dirname, 'config.json');
            try {
                fs.writeFileSync(configPath, JSON.stringify({ backendUrl: tunnelUrl }, null, 2));
                console.log('[Tunnel] config.json updated with:', tunnelUrl);

                // Push to GitHub
                try {
                    execSync('git add config.json', { cwd: __dirname });
                    execSync(`git commit -m "chore: Update live tunnel URL to ${tunnelUrl}"`, { cwd: __dirname });
                    execSync('git push origin main', { cwd: __dirname });
                    console.log('[Tunnel] config.json pushed to GitHub Pages.');
                } catch (gitErr) {
                    console.warn('[Tunnel] Git push failed (non-fatal):', gitErr.message);
                }
            } catch (fileErr) {
                console.error('[Tunnel] Failed to write config.json:', fileErr.message);
            }
        }
    });

    tunnel.stderr.on('data', (data) => {
        const line = data.toString();
        if (!line.includes('npm warn')) console.log('[Tunnel stderr]', line.trim());
    });

    tunnel.on('close', (code) => {
        console.log('[Tunnel] Process exited with code', code);
        if (tunnelProcess === tunnel) {
            tunnelProcess = null;
            tunnelUrl = null;
        }
    });
});

app.post('/api/stop-tunnel', (req, res) => {
    if (tunnelProcess && !tunnelProcess.killed) {
        tunnelProcess.kill();
        tunnelProcess = null;
        tunnelUrl = null;
        res.json({ success: true, message: 'Tunnel stopped.' });
    } else {
        res.json({ success: false, message: 'No tunnel running.' });
    }
});



// Start server — bind 0.0.0.0 so Android on same Wi-Fi can reach it at http://192.168.1.59:3000
app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n🚀 Smart Vision CRM Server → http://localhost:${PORT}  |  http://192.168.1.59:${PORT}\n`);
    connectToWhatsApp();
});
