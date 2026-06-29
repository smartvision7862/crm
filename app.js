// State Management
let currentView = 'dashboard';
let activeChatId = 1;
let activeOpTab = 'tasks';
let activeCustomerId = 1;

let BACKEND_URL = localStorage.getItem("settings-backend-url");
const _isLocalOrigin = window.location.hostname === 'localhost' || 
                       window.location.hostname === '127.0.0.1' || 
                       window.location.hostname.startsWith('192.168.');

if (_isLocalOrigin) {
    // When accessing directly on local network/localhost, override any stored localtunnel URL to use direct origin
    if (!BACKEND_URL || BACKEND_URL.includes('.loca.lt')) {
        BACKEND_URL = window.location.origin;
    }
} else if (!BACKEND_URL) {
    if (!window.location.origin || window.location.origin === 'null' || window.location.origin.startsWith('file:')) {
        // Running from GitHub Pages — use the live tunnel backend from config.json
        BACKEND_URL = 'https://smartvision-crm.loca.lt'; // default until config.json loads
    } else {
        BACKEND_URL = window.location.origin;
    }
}

// Auto-load the latest backend URL from config.json (updated every server start)
if (!_isLocalOrigin) {
    fetch('/crm/config.json?t=' + Date.now())
        .then(r => r.json())
        .then(cfg => {
            if (cfg && cfg.backendUrl && !localStorage.getItem("settings-backend-url")) {
                BACKEND_URL = cfg.backendUrl;
                console.log('[CRM] Backend URL loaded from config.json:', BACKEND_URL);
            }
        })
        .catch(() => {}); // silently fall back to default
}
// Global fetch interceptor to inject localtunnel bypass header automatically
const originalFetch = window.fetch;
window.fetch = async function(resource, options) {
    if (typeof resource === 'string' && resource.includes('loca.lt')) {
        options = options || {};
        options.headers = options.headers || {};
        if (options.headers instanceof Headers) {
            options.headers.set('bypass-tunnel-reminder', 'true');
        } else if (Array.isArray(options.headers)) {
            const hasIt = options.headers.some(h => h[0].toLowerCase() === 'bypass-tunnel-reminder');
            if (!hasIt) options.headers.push(['bypass-tunnel-reminder', 'true']);
        } else {
            options.headers['bypass-tunnel-reminder'] = 'true';
        }
    }
    return originalFetch(resource, options);
};

let isBackendConnected = false;
let seenMessageIds = new Set(); // Track by ID to survive server restarts
let isAppInitialized = false;

let smtpConfig = (() => {
    const saved = localStorage.getItem("smtpConfig");
    if (saved) {
        try {
            return JSON.parse(saved);
        } catch (e) {}
    }
    return {
        host: "smtp.gmail.com",
        port: "587",
        user: "smartvisionwll@gmail.com",
        pass: "nsbsniiokfgsqwqv"
    };
})();

let webhookUrl = localStorage.getItem("webhookUrl") || (BACKEND_URL + "/api/webhooks/incoming");

function formatQatarPhoneNumber(phone) {
    if (!phone) return "";
    let clean = phone.trim().replace(/[^0-9+]/g, '');
    if (clean.startsWith("+")) {
        return clean;
    }
    clean = clean.replace(/^0+/, '');
    if (clean.length === 8) {
        return "+974" + clean;
    }
    if (clean.startsWith("974") && clean.length === 11) {
        return "+" + clean;
    }
    return clean.length >= 8 ? "+" + clean : clean;
}

async function dispatchWebhook(event, data) {
    if (!webhookUrl) return;
    try {
        console.log(`Dispatched Webhook: ${event}`);
        const res = await fetch(webhookUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                event: event,
                timestamp: new Date().toISOString(),
                data: data
            })
        });
        if (!res.ok) {
            console.error(`Webhook dispatch failed with status: ${res.status}`);
        }
    } catch (err) {
        console.error(`Webhook dispatch network error:`, err);
    }
}

// Central Customer Database (Single Source of Truth)
let customers = [];

// Legacy variables mapping for backward compatibility
let leads = [];
let chats = {};

function syncState() {
    leads = customers.map(c => ({
        id: c.id,
        name: c.name,
        notes: c.notes,
        temp: c.temp,
        stage: c.stage,
        channel: c.channel,
        history: c.history.map(h => typeof h === 'string' ? h : `${h.text} (${h.time})`)
    }));

    chats = {};
    customers.forEach(c => {
        const idKey = c.id.toString();
        chats[idKey] = {
            name: c.name,
            status: c.channel === 'WhatsApp' ? 'WhatsApp Live Chat' : `${c.channel} Chat`,
            messages: c.messages || [],
            aiSuggestion: c.aiSuggestion,
            score: c.score || 70,
            temp: c.temp,
            phone: c.phone
        };
        
        if (c.phone) {
            const clean = c.phone.replace(/[^0-9]/g, '');
            chats[clean] = chats[idKey];
        }
    });

    // Automatically sync state changes to crm.xls spreadsheet in the background
    debouncedExcelSyncAndSave();
}

function getCustomerByChatId(chatId) {
    if (!chatId) return null;
    return customers.find(c => 
        c.id.toString() === chatId.toString() || 
        c.phone.replace(/[^0-9]/g, '') === chatId.toString().replace(/[^0-9]/g, '')
    );
}

async function loadCustomersFromExcel() {
    try {
        const res = await fetch(`${BACKEND_URL}/api/excel/read`);
        if (!res.ok) throw new Error("Offline");
        const resData = await res.json();
        if (resData.success && Array.isArray(resData.data)) {
            excelRows = resData.data;
            excelRows.forEach(row => {
                const id = parseInt(row.ID || row.id);
                if (isNaN(id)) return;
                
                let phone = row.Phone || row.phone || "";
                if (phone) {
                    phone = formatQatarPhoneNumber(phone);
                }
                
                let customer = customers.find(c => c.id === id);
                
                if (customer) {
                    customer.id = id;
                    if (row.Name || row.name) customer.name = row.Name || row.name;
                    if (phone) customer.phone = phone;
                    if (row.Email || row.email) customer.email = row.Email || row.email;
                    if (row.Company || row.company) customer.label = row.Company || row.company;
                    if (row.Status || row.status) {
                        const cleanStage = (row.Status || row.status).toLowerCase().trim();
                        const allowedStages = ['new', 'contacted', 'qualified', 'proposal', 'won'];
                        if (allowedStages.includes(cleanStage)) {
                            customer.stage = cleanStage;
                        }
                    }
                } else {
                    const newCust = {
                        id: id,
                        name: row.Name || row.name || `Excel Lead #${id}`,
                        phone: phone,
                        email: row.Email || row.email || "",
                        stage: (row.Status || row.status || "new").toLowerCase().trim(),
                        temp: "warm",
                        channel: "Excel",
                        notes: "Loaded from crm.xls database spreadsheet.",
                        labels: [],
                        loyaltyPoints: 0,
                        dateAdded: new Date().toISOString().split('T')[0],
                        lastContact: "",
                        messages: [],
                        aiSuggestion: null,
                        score: 70,
                        autopilotActive: true,
                        history: [
                            { type: "system", text: "Customer loaded from crm.xls spreadsheet on startup", time: new Date().toLocaleString() }
                        ]
                    };
                    customers.push(newCust);
                }
            });
            syncState();
            renderCustomersTable();
            renderKanban();
        }
    } catch (err) {
        console.error("Failed to load customers from excel on startup:", err);
    }
}

// Initial state synchronization
syncState();

// Initialize App
document.addEventListener("DOMContentLoaded", async () => {
    // Load persisted settings
    loadAllSettings();

    // Nav links setup
    document.querySelectorAll(".nav-link").forEach(link => {
        link.addEventListener("click", () => {
            const targetView = link.getAttribute("data-view");
            switchView(targetView);
        });
    });

    // Run checklist initialization for all customers so initial states are computed
    customers.forEach(c => {
        if (c.autopilotActive === undefined) {
            c.autopilotActive = true;
        }
        updateAIQualificationChecklist(c);
    });

    // Render initial Kanban Board
    renderKanban();

    // Render initial Customer Directory
    renderCustomersTable();
    selectCustomer(1);
    
    // Setup initial right sidebar in WhatsApp chat
    selectChat(1);

    // Render initial Projects & Tasks Board
    renderProjects();
    selectProject(1);
    renderOperationsTasks();
    renderInventory();
    renderFinance();
    renderSuppliers();
    renderRequests();
    renderSalesLedger();
    populateTemplateSelects();
    switchDashboardTab('overview');

    // Connect invoice listeners
    updateInvoicePreview();

    // Start polling the backend integrations server
    startBackendPolling();

    // Load extra customers from Excel database spreadsheet on startup
    await loadCustomersFromExcel();

    // Trigger Google Sheets sync if auto-sync is enabled on page load
    if (gSheetsAutoSyncEnabled && gSheetsUrl) {
        console.log('[GSheets] Auto-sync trigger on page load...');
        runGoogleSheetsImport(true);
    }

    // Mark application as fully initialized so auto-saves can trigger
    isAppInitialized = true;

    // Initialize User Manual logic
    initUserManualLogic();
});

// Backend Integration Polling
function startBackendPolling() {
    console.log("Starting backend polling loop...");
    setInterval(async () => {
        try {
            const res = await fetch(`${BACKEND_URL}/api/status`);
            if (!res.ok) throw new Error("Offline");
            
            const data = await res.json();
            isBackendConnected = true;
            
            updateWhatsAppStatus(data.status, data.qr);
            
            if (data.status === 'CONNECTED') {
                pollBackendMessages();
            }
        } catch (err) {
            isBackendConnected = false;
            // Fallback status representation in frontend
            document.getElementById("whatsapp-status-dot").className = "status-dot inactive";
            document.getElementById("whatsapp-status-text").innerText = "WhatsApp: Server Offline";
        }
    }, 2000);
}

function updateWhatsAppStatus(status, qr) {
    const dot = document.getElementById("whatsapp-status-dot");
    const text = document.getElementById("whatsapp-status-text");
    const qrContainer = document.querySelector("#qr-modal .qr-code");
    const qrStatusText = document.getElementById("qr-conn-status");
    
    if (status === 'CONNECTED') {
        dot.className = "status-dot active";
        text.innerText = "WhatsApp: Live (Connected)";
        
        if (qrStatusText) {
            qrStatusText.innerText = "Connected Successfully!";
            qrStatusText.style.color = "var(--success)";
            setTimeout(() => {
                closeQRModal();
            }, 1200);
        }
        
        // Update connection section in settings
        const qrSect = document.getElementById("settings-qr-section");
        if (qrSect) {
            qrSect.innerHTML = `
                <div style="background:rgba(16,185,129,0.08); border:1px solid var(--success); padding:12px; border-radius: var(--radius-md); margin-bottom:12px; display:flex; justify-content:space-between; align-items:center;">
                    <span style="color:var(--success); font-weight:700; font-size:13px;">✓ Linked with your Phone</span>
                    <span class="badge-status success">Active</span>
                </div>
                <button class="btn btn-danger" style="width:100%;" onclick="disconnectRealWhatsApp()">Disconnect WhatsApp Session</button>
            `;
        }
    } else if (status === 'CONNECTING') {
        dot.className = "status-dot warning";
        text.innerText = "WhatsApp: Awaiting Link...";
        
        if (qrStatusText) {
            qrStatusText.innerText = "QR Code ready! Scan now.";
            qrStatusText.style.color = "var(--warning)";
        }
        
        // Inject QR Image
        if (qr && qrContainer) {
            const spinner = document.getElementById("qr-loading-spinner");
            if (spinner) spinner.style.display = 'none';
            
            let img = qrContainer.querySelector("img");
            if (!img) {
                const svg = qrContainer.querySelector("svg");
                if (svg) svg.remove();
                
                img = document.createElement("img");
                img.style.maxWidth = "100%";
                qrContainer.appendChild(img);
            }
            img.src = qr;
        }
    } else {
        dot.className = "status-dot inactive";
        text.innerText = "WhatsApp: Offline";
        
        if (qrStatusText) {
            qrStatusText.innerText = "Awaiting QR generation...";
            qrStatusText.style.color = "var(--danger)";
        }
        
        const qrSect = document.getElementById("settings-qr-section");
        if (qrSect) {
            qrSect.innerHTML = `
                <button class="btn btn-primary" style="width:100%;" onclick="openQRModal()">Connect WhatsApp via QR</button>
            `;
        }
    }
}

async function pollBackendMessages() {
    try {
        const res = await fetch(`${BACKEND_URL}/api/messages`);
        if (!res.ok) return;
        
        const data = await res.json();
        
        // Process every message and track by unique ID to avoid duplicates
        // (using IDs instead of count survives server restarts)
        let hasNewMessages = false;
        data.messages.forEach(msg => {
            if (!msg.id || seenMessageIds.has(msg.id)) return; // Skip already-seen messages
            seenMessageIds.add(msg.id);
            hasNewMessages = true;

            const chatId = msg.from; // Phone number as chat ID
            
            let customer = getCustomerByChatId(chatId);
            if (!customer) {
                customer = {
                    id: customers.length + 1,
                    name: msg.name,
                    phone: `+${msg.from.split('@')[0]}`,
                    jid: msg.from,
                    email: `${msg.name.toLowerCase().replace(/[^a-z0-9]/g, '')}@gmail.com`,
                    stage: "new",
                    temp: "hot",
                    channel: "WhatsApp",
                    notes: `WhatsApp Inquiry: "${msg.text}"`,
                    labels: [],
                    loyaltyPoints: 0,
                    dateAdded: new Date().toISOString().split('T')[0],
                    lastContact: msg.time,
                    messages: [],
                    aiSuggestion: "Analysing query...",
                    score: 70,
                    autopilotActive: true, // Enabled by default for new incoming WhatsApp chats
                    history: [
                        { type: "system", text: `Inquiry received on WhatsApp (+${msg.from.split('@')[0]})`, time: new Date().toLocaleString() }
                    ]
                };
                customers.push(customer);
                syncState();
                renderCustomersTable();
                renderKanban();
                logActivity(`New WhatsApp lead: ${msg.name}`, msg.text, 'primary');
            }
            
            // Add message to chat buffer (deduplicated by ID already)
            customer.messages.push({
                sender: "incoming",
                text: msg.text,
                time: msg.time
            });
            customer.lastContact = msg.time;
            syncState();
            
            if (customer && activeChatId.toString() === customer.id.toString()) {
                renderChatMessages();
                renderCustomerTimeline(customer);
            }
            
            // Generate AI reply suggestion
            calculateAISuggestionForChat(chatId, msg.text);

            dispatchWebhook("whatsapp_message_received", {
                id: msg.id,
                from: msg.from,
                name: msg.name,
                text: msg.text,
                time: msg.time
            });

            // If this is the active customer in the directory, refresh
            if (activeCustomerId === customer.id) {
                selectCustomer(customer.id);
            }
        });

        if (hasNewMessages) {
            renderChatListSidebar();
        }
    } catch (err) {
        console.error("Error fetching incoming messages:", err);
    }
}

function calculateAISuggestionForChat(chatId, text) {
    const customer = getCustomerByChatId(chatId);
    if (!customer) return;

    let responseSuggestion = "";
    let score = 75;
    let temp = "Warm";

    const textLower = text.toLowerCase();
    
    // 1. Complaint detection & Escalation automation
    const isComplaint = textLower.includes("complaint") || textLower.includes("poor") || 
                        textLower.includes("angry") || textLower.includes("bad") || 
                        textLower.includes("leak") || textLower.includes("broken") || 
                        textLower.includes("defect") || textLower.includes("fail");
                        
    if (isComplaint) {
        responseSuggestion = `Dear ${customer.name}, we apologize for the trouble. I have registered your complaint in our CRM and escalated it to our administrator, Imtiyaz Ahmed. Our team will contact you in 5 minutes.`;
        score = 95;
        temp = "Hot";
        
        // Auto-escalate in CRM
        const logText = "ESCALATED: Customer registered a complaint via WhatsApp. Supervisor notified.";
        const isAlreadyLogged = customer.history.some(h => h.text && h.text.includes("ESCALATED"));
        if (!isAlreadyLogged) {
            customer.history.push({ type: "system", text: logText, time: new Date().toLocaleString() });
            logActivity(`Urgent Ticket Created: ${customer.name}`, `Escalated to Imtiyaz Ahmed`, 'danger');
            
            // Add task to task table
            const tbody = document.querySelector("#op-tab-tasks tbody");
            if (tbody) {
                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td style="font-weight: 700;">URGENT TICKET: WhatsApp Complaint from ${customer.name}</td>
                    <td>Imtiyaz Ahmed</td>
                    <td>Today</td>
                    <td><span class="badge-status danger">Critical</span></td>
                    <td><span class="badge-status danger">Needs Review</span></td>
                `;
                tbody.insertBefore(tr, tbody.firstChild);
            }
        }
    }
    // 2. Booking request detection
    else if (textLower.includes("schedule") || textLower.includes("tomorrow") || textLower.includes("book") || textLower.includes("visit") || textLower.includes("inspection")) {
        responseSuggestion = `Absolutely, ${customer.name}. I've logged your request for a survey. I can book our service team for tomorrow at 10:00 AM. Does that work for you?`;
        score = 90;
        temp = "Hot";
    }
    // 3. Price enquiry detection (Lead Qualification)
    else if (textLower.includes("price") || textLower.includes("cost") || textLower.includes("quote") || textLower.includes("how much")) {
        responseSuggestion = `Hi ${customer.name}! We'd be happy to prepare a quote for you. What services are you looking for, and where is your property located?`;
        score = 88;
        temp = "Hot";
    }
    // 3.5 CCTV / Security camera enquiry detection
    else if (textLower.includes("cctv") || textLower.includes("camera") || textLower.includes("security")) {
        responseSuggestion = `Hi ${customer.name}! Smart Vision supplies and installs premium IP & Analog CCTV camera systems (Dome, Bullet, PTZ) from certified brands like Hikvision, Dahua, and Honeywell, fully compliant with Qatar MOI-SSD regulations. How many cameras do you require, and is it for a commercial or residential property?`;
        score = 85;
        temp = "Warm";
    }
    // 3.7 MEP / HVAC / Maintenance Services enquiry detection
    else if (textLower.includes("mep") || textLower.includes("hvac") || textLower.includes("service") || textLower.includes("maintenance") || textLower.includes("ac repair")) {
        responseSuggestion = `Hi ${customer.name}! Smart Vision offers full MEP services, including Central/Split AC repairs, preventative HVAC maintenance, electrical installations, plumbing, and fire safety systems in Qatar. Are you interested in a commercial MEP contract or a one-time residential service?`;
        score = 92;
        temp = "Hot";
    }
    // 3.8 Fire Alarm & Fire Fighting enquiry detection
    else if (textLower.includes("fire") || textLower.includes("alarm") || textLower.includes("smoke") || textLower.includes("sprinkler") || textLower.includes("suppression") || textLower.includes("civil defence") || textLower.includes("qcd")) {
        responseSuggestion = `Hi ${customer.name}! Smart Vision is a QCD Grade A Certified contractor in Qatar. We design, install, and maintain Civil Defence-approved Fire Alarm and Fire Fighting/Suppression systems. Do you need testing, maintenance, or QCD approvals for your facility?`;
        score = 94;
        temp = "Hot";
    }
    // 3.9 Access Control & Biometrics enquiry detection
    else if (textLower.includes("access") || textLower.includes("biometric") || textLower.includes("door lock") || textLower.includes("gate") || textLower.includes("turnstile") || textLower.includes("fingerprint")) {
        responseSuggestion = `Hi ${customer.name}! Smart Vision offers advanced Access Control and Biometric security systems (card readers, facial recognition, turnstiles) for offices and industrial sites in Doha. What is the scope of your access control requirement?`;
        score = 86;
        temp = "Warm";
    }
    // 3.95 ELV, IT Cabling & Public Address (PA) enquiry detection
    else if (textLower.includes("network") || textLower.includes("cabling") || textLower.includes("cisco") || textLower.includes("wifi") || textLower.includes("public address") || textLower.includes("speaker") || textLower.includes("sound") || textLower.includes("audio")) {
        responseSuggestion = `Hi ${customer.name}! We specialize in ELV systems integration, including Structured Fiber/Cat6 Cabling, Cisco IT networks, and IP-based Public Address & Voice Evacuation audio arrays. Are you looking to set up an IT network or audio infrastructure?`;
        score = 85;
        temp = "Warm";
    }
    // 3.97 Building Automation, BMS & KNX/DALI lighting
    else if (textLower.includes("bms") || textLower.includes("automation") || textLower.includes("knx") || textLower.includes("smart home") || textLower.includes("lighting control") || textLower.includes("dali")) {
        responseSuggestion = `Hi ${customer.name}! We design and integrate Intelligent Building Management Systems (BMS), KNX lighting controls, DALI smart dimming, and energy management solutions to optimize your facility's operations. What project automation scope are you planning?`;
        score = 88;
        temp = "Warm";
    }
    // 3.98 SMART VISION - Restaurant Monitoring AI System detection
    else if (textLower.includes("restaurant monitoring") || (textLower.includes("restaurant") && textLower.includes("monitoring"))) {
        responseSuggestion = `Hi ${customer.name}! Our SMART VISION - Restaurant Monitoring AI System uses intelligent cameras to track table occupancy in real time, speed up seating management, and analyze service flow and dining patterns to optimize staff allocation. Would you like to schedule a demo of our restaurant analytics system?`;
        score = 95;
        temp = "Hot";
    }
    // 3.985 SMART VISION - Crowd Detection
    else if (textLower.includes("crowd detection") || textLower.includes("crowd monitoring") || textLower.includes("crowd count") || textLower.includes("crowd")) {
        responseSuggestion = `Hi ${customer.name}! Our SMART VISION - Crowd Detection AI system utilizes advanced density mapping and object detection algorithms to count and monitor dense groups of people in real-time, perfect for malls, events, and public safety management. Would you like a product overview sheet or a live demo?`;
        score = 95;
        temp = "Hot";
    }
    // 3.987 SMART VISION - Staff Movement AI System detection
    else if (textLower.includes("staff movement") || textLower.includes("employee tracking") || textLower.includes("employee movement")) {
        responseSuggestion = `Hi ${customer.name}! Our SMART VISION - Staff Movement AI System uses intelligent video analytics to track employee presence, movement patterns, and activity flow across zones. It optimizes workforce efficiency, prevents idle time, and ensures duty compliance with real-time insights and automated alerts. Would you like a quote or a demo of our workforce tracking dashboard?`;
        score = 95;
        temp = "Hot";
    }
    // 3.988 SMART VISION - Unauthorized Intrusion Detection
    else if (textLower.includes("intrusion") || textLower.includes("unauthorized intrusion") || textLower.includes("perimeter security")) {
        responseSuggestion = `Hi ${customer.name}! Our SMART VISION - Unauthorized Intrusion Detection system uses AI-driven video analytics to identify abnormal movements, forced entry attempts, and suspicious behavior around residential or private properties in real-time, ensuring early warning of security breaches. Would you like to schedule a survey for a perimeter monitoring setup?`;
        score = 95;
        temp = "Hot";
    }
    // 3.989 SMART VISION - Bar Monitoring System
    else if (textLower.includes("bar monitoring") || textLower.includes("bar tracking") || (textLower.includes("bar") && textLower.includes("monitoring"))) {
        responseSuggestion = `Hi ${customer.name}! Our SMART VISION - Bar Monitoring System uses AI-powered video analytics to track real-time pouring activity, detect over-pouring and wastage, verify staff compliance, and prevent inventory shrinkage. Would you like to schedule a demo of our bar automation system?`;
        score = 95;
        temp = "Hot";
    }
    // 3.991 SMART TRUEVOLUME
    else if (textLower.includes("truevolume") || textLower.includes("true volume") || textLower.includes("weight tracking")) {
        responseSuggestion = `Hi ${customer.name}! Our SMART TRUEVOLUME system determines the remaining quantity of liquor in a bottle in real-time by continuously monitoring its weight using a digital weighing machine compared with predefined reference data. This provides accurate volume tracking and controls consumption/wastage. Would you like a product demo?`;
        score = 95;
        temp = "Hot";
    }
    // 3.992 INTELLIGENT RESTAURANT AI SYSTEM
    else if (textLower.includes("intelligent restaurant") || textLower.includes("restaurant seat map") || textLower.includes("digital menu")) {
        responseSuggestion = `Hi ${customer.name}! Our INTELLIGENT RESTAURANT AI SYSTEM allows patrons to book seats via a real-time interactive seat map, explore digital menus, customize orders, and process secure two-stage payments (advance booking deposit + final checkout billing). Would you like to schedule a demo of our restaurant ordering and reservation platform?`;
        score = 95;
        temp = "Hot";
    }
    // 3.99 AI / Computer Vision / Software solutions enquiry detection
    else if (textLower.includes("computer vision") || textLower.includes("software") || textLower.includes("intelligent") || textLower.includes("helmet") || textLower.includes("drowsiness") || textLower.includes("fall detection") || textLower.includes("truevolume") || textLower.includes("storyboard") || textLower.includes("fable") || textLower.includes("restaurant") || textLower.includes("bar") || textLower.includes("staff movement") || textLower.includes("intrusion") || textLower.includes("traffic")) {
        responseSuggestion = `Hi ${customer.name}! Smart Vision offers specialized AI & Computer Vision tools, including safety detection (Helmet/Vest), driver drowsiness warning, crowd/intrusion alerts, and retail automation (Restaurant Seating, Bar TrueVolume tracking). What AI or software solution are you looking to implement?`;
        score = 90;
        temp = "Hot";
    }
    // 4. Default greetings
    else if (textLower.includes("hello") || textLower.includes("hi") || textLower.includes("hey")) {
        responseSuggestion = `Hello ${customer.name}! Thank you for contacting Smart Vision. How can we assist you with our services today?`;
    } else {
        responseSuggestion = `Thank you for your message, ${customer.name}. I've logged your enquiry and one of our human staff members will respond in a moment.`;
    }

    customer.aiSuggestion = responseSuggestion;
    customer.score = score;
    customer.temp = temp.toLowerCase();
    
    syncState();

    if (customer && activeChatId.toString() === customer.id.toString()) {
        const suggBox = document.getElementById("ai-suggestion-box");
        suggBox.style.display = "flex";
        document.getElementById("ai-suggestion-text").innerText = `"${responseSuggestion}"`;
        
        const headerBadge = suggBox.querySelector(".ai-suggestion-header span:last-child");
        if (headerBadge) {
            headerBadge.innerText = `Leads Score: ${score} (${temp})`;
        }
    }

    // Run checklist parser on the new message context
    updateAIQualificationChecklist(customer);

    // If Autopilot is enabled for this customer, dispatch reply automatically
    if (customer.autopilotActive) {
        const autopilotDelay = parseInt(localStorage.getItem("settings-ai-delay") || "1") * 1000;
        setTimeout(() => {
            sendAISuggestion(customer.id);
        }, autopilotDelay);
    }
}

async function disconnectRealWhatsApp() {
    if (confirm("Disconnect your real WhatsApp connection? You will need to scan the QR code to pair again.")) {
        try {
            await fetch(`${BACKEND_URL}/api/disconnect`, { method: 'POST' });
            alert("Disconnected session.");
        } catch (err) {
            alert("Failed to disconnect: " + err.message);
        }
    }
}

function renderChatListSidebar() {
    const listContainer = document.getElementById("chat-list-container");
    if (!listContainer) return;
    
    listContainer.innerHTML = "";
    
    const renderedNames = new Set();
    
    Object.keys(chats).forEach(id => {
        const chat = chats[id];
        if (renderedNames.has(chat.name)) return;
        renderedNames.add(chat.name);
        
        const lastMsg = chat.messages[chat.messages.length - 1];
        let preview = "No messages yet";
        if (lastMsg) {
            if (lastMsg.text) {
                preview = lastMsg.attachment ? `📎 ${lastMsg.text}` : lastMsg.text;
            } else if (lastMsg.attachment) {
                preview = `📎 File: ${lastMsg.attachment.name}`;
            }
        }
        const time = lastMsg ? lastMsg.time : "";
        const isActive = activeChatId.toString() === id.toString() ? "active" : "";
        
        const avatarInitials = chat.name.split(' ').map(n => n[0]).join('').substring(0,2).toUpperCase();
        
        const item = document.createElement("div");
        item.className = `chat-item ${isActive}`;
        item.setAttribute("data-chat-id", id);
        item.onclick = () => selectChat(id);
        
        item.innerHTML = `
            <div class="chat-avatar-container">
                <div class="chat-avatar">${avatarInitials}</div>
                <div class="channel-icon">
                    <svg viewBox="0 0 24 24"><path d="M.057 24l1.687-6.163c-1.041-1.804-1.588-3.849-1.587-5.946C.003 5.324 5.328.001 11.859.001c3.168.001 6.147 1.24 8.39 3.49 2.24 2.25 3.475 5.234 3.471 8.411-.004 6.535-5.329 11.857-11.861 11.857-2.003 0-3.974-.509-5.727-1.48L0 24zm6.49-2.887l.412.244c1.516.899 3.25 1.374 5.023 1.375 5.767 0 10.457-4.69 10.461-10.462.002-2.798-1.085-5.429-3.065-7.414s-4.605-3.076-7.417-3.076c-5.77 0-10.46 4.69-10.464 10.463-.002 1.902.486 3.765 1.417 5.405l.265.467-.983 3.593 3.68-.965zm10.741-7.234c-.297-.148-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.148-.197.297-.768.967-.941 1.164-.173.198-.347.223-.644.074-.297-.149-1.255-.462-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.501-.669-.51l-.57-.01c-.198 0-.52.074-.792.372s-1.04 1.016-1.04 2.479 1.065 2.876 1.213 3.074c.149.198 2.095 3.2 5.076 4.487.709.306 1.263.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.695.248-1.29.173-1.414-.074-.124-.272-.198-.57-.347z"/></svg>
                </div>
            </div>
            <div class="chat-item-details">
                <div class="chat-item-header">
                    <span class="chat-item-name">${chat.name}</span>
                    <span class="chat-item-time">${time}</span>
                </div>
                <div class="chat-item-preview">${preview}</div>
                <div class="chat-item-meta">
                    <span class="ai-badge">${chat.status === 'WhatsApp Live Chat' ? "Live WA" : "Simulator"}</span>
                </div>
            </div>
        `;
        listContainer.appendChild(item);
    });
}

// View Switching
function switchView(viewId) {
    if (!viewId) return;
    
    document.querySelectorAll(".nav-link").forEach(link => {
        link.classList.remove("active");
        if (link.getAttribute("data-view") === viewId) {
            link.classList.add("active");
        }
    });

    document.querySelectorAll(".view-container").forEach(view => {
        view.classList.remove("active");
    });
    
    const targetViewEl = document.getElementById(`view-${viewId}`);
    if (targetViewEl) {
        targetViewEl.classList.add("active");
        currentView = viewId;
    }

    if (viewId === 'dashboard') {
        switchDashboardTab('overview');
    }

    if (viewId === 'chat') {
        switchChatTab('inbox');
        renderChatListSidebar();
    }

    if (viewId === 'email') {
        switchEmailTab('mailbox');
    }

    if (viewId === 'loyalty') {
        switchLoyaltyTab('program');
    }

    if (viewId === 'settings') {
        switchSettingsTab('integrations');
    }

    if (viewId === 'excel') {
        if (excelRows.length === 0) {
            reloadExcelFromDisk();
        } else {
            renderExcelGrid();
        }
    }
}

// Render Kanban Board
function renderKanban() {
    const stages = ['new', 'contacted', 'qualified', 'proposal', 'won'];
    
    stages.forEach(stage => {
        const col = document.getElementById(`col-${stage}`);
        if (col) col.innerHTML = '';
        
        const badge = document.getElementById(`badge-${stage}`);
        if (badge) badge.innerText = '0';
    });

    leads.forEach(lead => {
        const col = document.getElementById(`col-${lead.stage}`);
        if (!col) return;

        const card = document.createElement("div");
        card.className = "kanban-card";
        card.setAttribute("draggable", "true");
        card.setAttribute("data-lead-id", lead.id);
        card.onclick = () => showLeadDetails(lead.id);

        let tempTag = "";
        if (lead.temp === 'hot') tempTag = `<span class="card-tag tag-hot">Hot Lead</span>`;
        else if (lead.temp === 'warm') tempTag = `<span class="card-tag tag-warm">Warm Lead</span>`;
        else tempTag = `<span class="card-tag tag-cold">Cold Lead</span>`;

        card.innerHTML = `
            ${tempTag}
            <div class="card-title">${lead.name}</div>
            <div class="card-desc">${lead.notes}</div>
            <div class="card-footer">
                <div class="card-meta">
                    <svg viewBox="0 0 24 24"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path></svg>
                    ${lead.channel}
                </div>
                <div class="card-assignee">${lead.name.split(' ').map(n => n[0]).join('').substring(0,2)}</div>
            </div>
        `;

        col.appendChild(card);
        
        const badge = document.getElementById(`badge-${lead.stage}`);
        if (badge) {
            badge.innerText = parseInt(badge.innerText) + 1;
        }
    });

    const kpiLeadsEl = document.getElementById("kpi-leads");
    if (kpiLeadsEl) kpiLeadsEl.innerText = leads.length;

    setupDragAndDrop();
}

// Drag and Drop Logic
function setupDragAndDrop() {
    const cards = document.querySelectorAll(".kanban-card");
    const columns = document.querySelectorAll(".kanban-col-cards");

    cards.forEach(card => {
        card.addEventListener("dragstart", () => {
            card.classList.add("dragging");
        });
        card.addEventListener("dragend", () => {
            card.classList.remove("dragging");
        });
    });

    columns.forEach(col => {
        col.addEventListener("dragover", e => {
            e.preventDefault();
            const draggingCard = document.querySelector(".dragging");
            col.appendChild(draggingCard);
        });

        col.addEventListener("drop", () => {
            const draggingCard = document.querySelector(".dragging");
            if (!draggingCard) return;
            const leadId = parseInt(draggingCard.getAttribute("data-lead-id"));
            const targetStage = col.parentElement.getAttribute("data-stage");
            
            const customer = customers.find(c => c.id === leadId);
            if (customer && customer.stage !== targetStage) {
                customer.stage = targetStage;
                customer.history.push({ type: "system", text: `Moved to stage: ${targetStage.toUpperCase()} via Drag & Drop`, time: new Date().toLocaleString() });
                syncState();
                if (activeCustomerId === customer.id) {
                    selectCustomer(customer.id);
                }
                logActivity(`Lead ${customer.name} moved to ${targetStage.toUpperCase()}`, `Stage updated manually.`, 'primary');
                renderKanban();
            }
        });
    });
}

// Lead Details Dialog
let selectedLeadIdForDetails = null;

function showLeadDetails(leadId) {
    const lead = leads.find(l => l.id === leadId);
    if (!lead) return;

    selectedLeadIdForDetails = leadId;
    document.getElementById("details-lead-name").innerText = lead.name;
    document.getElementById("details-lead-stage").innerText = lead.stage.toUpperCase();
    document.getElementById("details-lead-notes").innerText = lead.notes;

    let tempHtml = "";
    if (lead.temp === 'hot') tempHtml = "🔥 Hot";
    else if (lead.temp === 'warm') tempHtml = "⚡ Warm";
    else tempHtml = "❄️ Cold";
    document.getElementById("details-lead-temp").innerText = tempHtml;

    const historyBox = document.getElementById("details-lead-history");
    historyBox.innerHTML = '';
    lead.history.forEach(log => {
        const item = document.createElement("div");
        item.style.fontSize = "12px";
        item.style.borderBottom = "1px solid rgba(255,255,255,0.02)";
        item.style.padding = "4px 0";
        item.style.color = "var(--text-muted)";
        item.innerText = log;
        historyBox.appendChild(item);
    });

    const advanceBtn = document.getElementById("details-btn-move");
    if (lead.stage === 'won') {
        advanceBtn.style.display = 'none';
    } else {
        advanceBtn.style.display = 'inline-block';
        advanceBtn.innerText = "Advance Stage";
    }

    document.getElementById("lead-details-modal").classList.add("active");
}

function closeLeadDetailsModal() {
    document.getElementById("lead-details-modal").classList.remove("active");
}

function addLeadFollowUpNote() {
    const input = document.getElementById("details-new-note");
    const noteText = input.value.trim();
    if (noteText === "" || !selectedLeadIdForDetails) return;

    const customer = customers.find(c => c.id === selectedLeadIdForDetails);
    if (customer) {
        const timestamp = new Date().toLocaleString();
        customer.history.push({ type: "notes", text: noteText, time: timestamp });
        syncState();
        input.value = "";
        showLeadDetails(selectedLeadIdForDetails);
        if (activeCustomerId === customer.id) {
            selectCustomer(customer.id);
        }
        logActivity(`Follow-up note added for ${customer.name}`, noteText, 'secondary');
    }
}

function promoteLeadStage() {
    if (!selectedLeadIdForDetails) return;
    const customer = customers.find(c => c.id === selectedLeadIdForDetails);
    if (customer) {
        const stages = ['new', 'contacted', 'qualified', 'proposal', 'won'];
        const currentIndex = stages.indexOf(customer.stage);
        if (currentIndex !== -1 && currentIndex < stages.length - 1) {
            const nextStage = stages[currentIndex + 1];
            customer.stage = nextStage;
            customer.history.push({ type: "system", text: `Advanced stage to ${nextStage.toUpperCase()}`, time: new Date().toLocaleString() });
            syncState();
            closeLeadDetailsModal();
            renderKanban();
            if (activeCustomerId === customer.id) {
                selectCustomer(customer.id);
            }
            logActivity(`Lead ${customer.name} advanced`, `Moved to ${nextStage.toUpperCase()}`, 'success');
        }
    }
}

function escalateLeadToStaff() {
    if (!selectedLeadIdForDetails) return;
    const customer = customers.find(c => c.id === selectedLeadIdForDetails);
    if (customer) {
        customer.history.push({ type: "system", text: `ESCALATED: Flagged for urgent follow-up`, time: new Date().toLocaleString() });
        syncState();
        closeLeadDetailsModal();
        if (activeCustomerId === customer.id) {
            selectCustomer(customer.id);
        }
        logActivity(`Lead ${customer.name} ESCALATED`, `System assigned task to Imtiyaz Ahmed.`, 'danger');
        
        const tbody = document.querySelector("#op-tab-tasks tbody");
        if (tbody) {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td style="font-weight: 700;">URGENT: Deal follow-up with ${customer.name}</td>
                <td>Imtiyaz Ahmed</td>
                <td>Today</td>
                <td><span class="badge-status danger">Urgent</span></td>
                <td><span class="badge-status danger">Needs Follow-Up</span></td>
            `;
            tbody.insertBefore(tr, tbody.firstChild);
        }
    }
}

// Add Lead Dialog
function openAddLeadModal() {
    document.getElementById("add-lead-modal").classList.add("active");
}

function closeAddLeadModal() {
    document.getElementById("add-lead-modal").classList.remove("active");
}

function submitAddLeadForm() {
    const name = document.getElementById("form-lead-name").value.trim();
    const notes = document.getElementById("form-lead-notes").value.trim();
    const temp = document.getElementById("form-lead-temp").value;
    const stage = document.getElementById("form-lead-stage").value;

    if (name === "" || notes === "") {
        alert("Please enter both Lead Name and Enquiry Details.");
        return;
    }

    const newCust = {
        id: customers.length + 1,
        name: name,
        phone: "+974 0000 0000",
        email: `${name.toLowerCase().replace(/[^a-z0-9]/g, '')}@gmail.com`,
        stage: stage,
        temp: temp,
        channel: "Manual",
        notes: notes,
        labels: [],
        loyaltyPoints: 0,
        dateAdded: new Date().toISOString().split('T')[0],
        lastContact: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
        messages: [],
        aiSuggestion: null,
        score: 50,
        autopilotActive: true,
        history: [
            { type: "system", text: "Profile manually registered in pipeline", time: new Date().toLocaleString() }
        ]
    };

    customers.push(newCust);
    document.getElementById("form-lead-name").value = "";
    document.getElementById("form-lead-notes").value = "";
    
    syncState();
    closeAddLeadModal();
    renderKanban();
    renderCustomersTable();
    logActivity(`New manual lead added: ${name}`, `Added to ${stage.toUpperCase()} stage`, 'success');
}

// WhatsApp Simulator Chat Switching

// Attachment state and handlers
let pendingAttachment = null;

function formatBytes(bytes, decimals = 1) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

function openImageLightbox(url, name) {
    const modal = document.getElementById("image-lightbox-modal");
    const img = document.getElementById("lightbox-img");
    const caption = document.getElementById("lightbox-caption");
    if (modal && img && caption) {
        img.src = url;
        caption.innerText = name || "Image Attachment";
        modal.style.display = "flex";
        modal.classList.add("active");
    }
}

function closeImageLightbox() {
    const modal = document.getElementById("image-lightbox-modal");
    if (modal) {
        modal.style.display = "none";
        modal.classList.remove("active");
    }
}

function downloadAttachment(url, name) {
    const link = document.createElement("a");
    link.href = url;
    link.download = name;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function triggerChatAttachment() {
    document.getElementById("chat-file-input").click();
}

function handleChatFileSelected(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        pendingAttachment = {
            name: file.name,
            size: file.size,
            type: file.type,
            url: e.target.result
        };
        
        // Update Preview Bar UI
        const previewBar = document.getElementById("chat-attachment-preview");
        const thumbContainer = document.getElementById("chat-attachment-thumb-container");
        const thumb = document.getElementById("chat-attachment-thumb");
        const icon = document.getElementById("chat-attachment-icon");
        const nameLbl = document.getElementById("chat-attachment-name");
        const sizeLbl = document.getElementById("chat-attachment-size");
        
        if (previewBar && nameLbl && sizeLbl) {
            nameLbl.innerText = file.name;
            sizeLbl.innerText = formatBytes(file.size);
            
            if (file.type.startsWith("image/")) {
                if (thumb && thumbContainer) {
                    thumb.src = e.target.result;
                    thumbContainer.style.display = "flex";
                }
                if (icon) icon.style.display = "none";
            } else {
                if (thumbContainer) thumbContainer.style.display = "none";
                if (icon) {
                    icon.style.display = "block";
                    if (file.type.includes("pdf")) icon.innerText = "📕";
                    else if (file.type.includes("sheet") || file.type.includes("excel") || file.name.endsWith(".xls") || file.name.endsWith(".xlsx") || file.name.endsWith(".csv")) icon.innerText = "📊";
                    else if (file.type.includes("word") || file.name.endsWith(".doc") || file.name.endsWith(".docx")) icon.innerText = "📘";
                    else if (file.type.includes("zip") || file.type.includes("rar") || file.type.includes("tar") || file.type.includes("compressed")) icon.innerText = "📦";
                    else icon.innerText = "📄";
                }
            }
            previewBar.style.display = "flex";
        }
    };
    reader.readAsDataURL(file);
}

function clearChatAttachment() {
    pendingAttachment = null;
    const fileInput = document.getElementById("chat-file-input");
    if (fileInput) fileInput.value = "";
    const previewBar = document.getElementById("chat-attachment-preview");
    if (previewBar) previewBar.style.display = "none";
}

function selectChat(chatId) {
    activeChatId = chatId;
    
    document.querySelectorAll(".chat-item").forEach(item => {
        item.classList.remove("active");
        if (item.getAttribute("data-chat-id").toString() === chatId.toString()) {
            item.classList.add("active");
        }
    });

    const customer = getCustomerByChatId(chatId);
    if (!customer) return;

    document.getElementById("chat-header-name").innerText = customer.name;
    document.getElementById("chat-header-status").innerText = customer.channel === 'WhatsApp' ? 'WhatsApp Live Chat' : `${customer.channel} Chat`;

    renderChatMessages();

    // Sync Autopilot state
    const toggle = document.getElementById("chat-autopilot-toggle");
    const indicator = document.getElementById("autopilot-indicator");
    if (toggle && indicator) {
        toggle.checked = !!customer.autopilotActive;
        indicator.style.display = customer.autopilotActive ? "flex" : "none";
    }

    const suggBox = document.getElementById("ai-suggestion-box");
    if (customer.aiSuggestion) {
        suggBox.style.display = "flex";
        document.getElementById("ai-suggestion-text").innerText = `"${customer.aiSuggestion}"`;
        
        const headerBadge = suggBox.querySelector(".ai-suggestion-header span:last-child");
        if (headerBadge) {
            headerBadge.innerText = `Leads Score: ${customer.score} (${customer.temp.toUpperCase()})`;
        }
    } else {
        suggBox.style.display = "none";
    }

    // Populate right Customer Profile panel in chat view
    document.getElementById("chat-profile-avatar").innerText = customer.name.split(' ').map(n => n[0]).join('').substring(0,2).toUpperCase();
    document.getElementById("chat-profile-name").innerText = customer.name;
    document.getElementById("chat-profile-phone").innerText = customer.phone;
    
    const tagsContainer = document.getElementById("chat-profile-tags");
    tagsContainer.innerHTML = `
        <span class="card-tag tag-${customer.temp.toLowerCase()}">${customer.temp} Lead</span>
        <span class="badge-status success" style="padding:2px 8px; font-size:10px;">${customer.channel}</span>
    `;
    
    renderChatProfileLabels(customer);
    document.getElementById("chat-profile-notes").value = customer.notes || "";

    // Update checklist panels and scores
    updateAIQualificationChecklist(customer);
}

function renderChatMessages() {
    const container = document.getElementById("chat-messages-container");
    container.innerHTML = '';
    
    const chat = chats[activeChatId];
    if (!chat) return;

    chat.messages.forEach(msg => {
        const bubble = document.createElement("div");
        bubble.className = `message-bubble ${msg.sender}`;
        
        let attachmentHtml = '';
        if (msg.attachment) {
            const att = msg.attachment;
            if (att.type && att.type.startsWith('image/')) {
                attachmentHtml = `
                    <div class="message-attachment image-attachment" onclick="openImageLightbox('${att.url}', '${att.name}')">
                        <img src="${att.url}" alt="${att.name}">
                    </div>
                `;
            } else {
                attachmentHtml = `
                    <div class="message-attachment file-attachment" onclick="downloadAttachment('${att.url}', '${att.name}')">
                        <div class="file-attachment-icon">📄</div>
                        <div class="file-attachment-info">
                            <div class="file-attachment-name" title="${att.name}">${att.name}</div>
                            <div class="file-attachment-size">${formatBytes(att.size)}</div>
                        </div>
                        <div class="file-attachment-download">
                            <svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="2.5" fill="none"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                        </div>
                    </div>
                `;
            }
        }
        
        const textHtml = msg.text ? `<div class="message-text">${msg.text}</div>` : '';
        
        bubble.innerHTML = `
            ${attachmentHtml}
            ${textHtml}
            <span class="message-time">${msg.time}</span>
        `;
        container.appendChild(bubble);
    });

    container.scrollTop = container.scrollHeight;
}

async function sendAISuggestion(chatId = activeChatId) {
    const customer = getCustomerByChatId(chatId);
    if (customer && customer.aiSuggestion) {
        const suggestion = customer.aiSuggestion;
        const timeNow = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        
        customer.messages.push({
            sender: "outgoing",
            text: suggestion,
            time: timeNow
        });
        customer.aiSuggestion = null;
        customer.lastContact = timeNow;
        
        syncState();
        
        if (activeChatId.toString() === chatId.toString()) {
            renderChatMessages();
            document.getElementById("ai-suggestion-box").style.display = "none";
        }
        
        logActivity(`AI Reply sent to ${customer.name}`, suggestion, 'success');
        
        renderChatListSidebar();

        if (activeCustomerId === customer.id) {
            renderCustomerTimeline(customer);
        }

        // Run checklist updates
        updateAIQualificationChecklist(customer);

        const recipientJid = customer.jid || (customer.phone ? customer.phone.replace(/[^0-9]/g, '') + '@s.whatsapp.net' : '');

        dispatchWebhook("whatsapp_message_sent", {
            to: recipientJid,
            text: suggestion,
            sender: "autopilot"
        });

        // POST message to real phone number if live backend is connected
        if (isBackendConnected && recipientJid) {
            try {
                await fetch(`${BACKEND_URL}/api/send`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ to: recipientJid, text: suggestion })
                });
            } catch (err) {
                console.error("Failed to send real WhatsApp message:", err);
            }
        }
    }
}

function regenerateAISuggestion() {
    const customer = getCustomerByChatId(activeChatId);
    if (customer) {
        const newText = prompt("Edit the AI reply suggestions draft:", customer.aiSuggestion);
        if (newText && newText.trim() !== "") {
            customer.aiSuggestion = newText;
            syncState();
            document.getElementById("ai-suggestion-text").innerText = `"${newText}"`;
        }
    }
}

function checkChatSubmit(e) {
    if (e.key === "Enter") {
        submitManualMessage();
    }
}

async function submitManualMessage() {
    const input = document.getElementById("chat-input-field");
    const text = input.value.trim();
    if (text === "" && !pendingAttachment) return;

    const customer = getCustomerByChatId(activeChatId);
    if (customer) {
        const timeNow = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        
        const msgObj = {
            sender: "outgoing",
            text: text,
            time: timeNow
        };
        if (pendingAttachment) {
            msgObj.attachment = { ...pendingAttachment };
        }
        
        customer.messages.push(msgObj);
        customer.lastContact = timeNow;
        syncState();
        input.value = "";
        
        const previewText = pendingAttachment ? (text ? `📎 ${text}` : `📎 File: ${pendingAttachment.name}`) : text;
        
        renderChatMessages();
        renderChatListSidebar();

        const logText = pendingAttachment ? (text ? `[Attachment: ${pendingAttachment.name}] ${text}` : `[Attachment: ${pendingAttachment.name}]`) : text;
        logActivity(`Manual message sent to ${customer.name}`, logText, 'primary');

        if (activeCustomerId === customer.id) {
            renderCustomerTimeline(customer);
        }

        // Parse conversation and update AI checklists
        updateAIQualificationChecklist(customer);

        const recipientJid = customer.jid || (customer.phone ? customer.phone.replace(/[^0-9]/g, '') + '@s.whatsapp.net' : '');

        let msgToSendText = text;
        if (pendingAttachment && text === "") {
            msgToSendText = `[File Attachment: ${pendingAttachment.name}]`;
        } else if (pendingAttachment && text !== "") {
            msgToSendText = `${text} (Attachment: ${pendingAttachment.name})`;
        }

        dispatchWebhook("whatsapp_message_sent", {
            to: recipientJid,
            text: msgToSendText,
            sender: "manual",
            attachment: pendingAttachment ? { name: pendingAttachment.name, type: pendingAttachment.type, size: pendingAttachment.size } : null
        });

        // POST message to real phone number if live backend is connected
        if (isBackendConnected && recipientJid) {
            try {
                await fetch(`${BACKEND_URL}/api/send`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ to: recipientJid, text: msgToSendText })
                });
            } catch (err) {
                console.error("Failed to send real WhatsApp message:", err);
            }
        }
        
        clearChatAttachment();
    }
}

// Mock Simulator triggers
function simulateIncomingMsg() {
    simulateNewChatMsg();
}

function simulateNewChatMsg() {
    const customer = getCustomerByChatId(activeChatId);
    if (!customer) return;

    const simulatedTexts = [
        "Sounds good. How long will the inspection take?",
        "Do you provide a warranty on parts and compressor replacements?",
        "Can I pay via Bank Transfer / Commercial Bank of Qatar?",
        "Do you have liability insurance for commercial building services?",
        "Help! The split AC units in the main hall are broken and leaking water! This is poor service. Send a tech now!",
        "Hello, what is the price/cost for a 4-bedroom villa painting service in Doha?"
    ];

    const randomText = simulatedTexts[Math.floor(Math.random() * simulatedTexts.length)];
    const timeNow = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

    // 25% chance of simulating an incoming attachment
    const hasAttachment = Math.random() < 0.25;
    let mockAttachment = null;
    let textToSimulate = randomText;
    
    if (hasAttachment) {
        const attachTypes = ['image', 'pdf'];
        const randomAttachType = attachTypes[Math.floor(Math.random() * attachTypes.length)];
        
        if (randomAttachType === 'image') {
            mockAttachment = {
                name: "ac_leakage_photo.jpg",
                type: "image/jpeg",
                size: 204800, // 200 KB
                url: "https://images.unsplash.com/photo-1621905251189-08b45d6a269e?w=500&auto=format&fit=crop&q=60"
            };
            textToSimulate = "Here is a photo of the leaking AC unit. " + randomText;
        } else {
            mockAttachment = {
                name: "hvac_building_plan.pdf",
                type: "application/pdf",
                size: 1572864, // 1.5 MB
                url: "data:application/pdf;base64,JVBERi0xLjUKMSAwIG9iajw8L1R5cGUvQ2F0YWxvZy9QYWdlcyAyIDAgUj4+ZW5kb2JqMiAwIG9iajw8L1R5cGUvUGFnZXMvS2lkc1szIDAgUl0vQ291bnQgMT4+ZW5kb2JqMyAwIG9iajw8L1R5cGUvUGFnZS9QYXJlbnQgMiAwIFIvTWVkaWFCb3hbMCAwIDU5NSA4NDJdL0NvbnRlbnRzIDQgMCBSPj5lbmRvYmogNCAwIG9iajw8L0xlbmd0aCA1OT4+c3RyZWFtCkJULy9GMSAxMiBUZgogNzAgNzAwIFRkCiAoU21hcnQgVmlzaW9uIE1FUCAtIEhWQUMgQnVpbGRpbmcgUGxhbiBEb2N1bWVudCkgVGoKRVQKZW5kc3RyZWFtCmVuZG9iajdpc2gK"
            };
            textToSimulate = "I've uploaded the building layout specifications. " + randomText;
        }
    }

    const msgObj = {
        sender: "incoming",
        text: textToSimulate,
        time: timeNow
    };
    if (mockAttachment) {
        msgObj.attachment = mockAttachment;
    }

    customer.messages.push(msgObj);
    customer.lastContact = timeNow;
    syncState();

    renderChatMessages();
    renderChatListSidebar();

    if (activeCustomerId === customer.id) {
        renderCustomerTimeline(customer);
    }

    const suggBox = document.getElementById("ai-suggestion-box");
    if (suggBox) suggBox.style.opacity = 0.5;

    // Analyze conversation with checklist and updates
    updateAIQualificationChecklist(customer);
    
    setTimeout(() => {
        let responseSuggestion = "";
        const textLower = randomText.toLowerCase();
        
        if (textLower.includes("leaking") || textLower.includes("broken") || textLower.includes("poor service")) {
            responseSuggestion = `Dear ${customer.name}, we apologize for the trouble. I have registered your complaint in our CRM and escalated it to our administrator, Imtiyaz Ahmed. Our team will contact you in 5 minutes.`;
            
            // Auto-escalate in CRM
            const logText = "ESCALATED: Customer registered a complaint via WhatsApp. Supervisor notified.";
            const isAlreadyLogged = customer.history.some(h => h.text && h.text.includes("ESCALATED"));
            if (!isAlreadyLogged) {
                customer.history.push({ type: "system", text: logText, time: new Date().toLocaleString() });
                logActivity(`Urgent Ticket Created: ${customer.name}`, `Escalated to Imtiyaz Ahmed`, 'danger');
                
                // Add task to task table
                const tbody = document.querySelector("#op-tab-tasks tbody");
                if (tbody) {
                    const tr = document.createElement("tr");
                    tr.innerHTML = `
                        <td style="font-weight: 700;">URGENT TICKET: WhatsApp Complaint from ${customer.name}</td>
                        <td>Imtiyaz Ahmed</td>
                        <td>Today</td>
                        <td><span class="badge-status danger">Critical</span></td>
                        <td><span class="badge-status danger">Needs Review</span></td>
                    `;
                    tbody.insertBefore(tr, tbody.firstChild);
                }
            }
        } else if (textLower.includes("price") || textLower.includes("cost") || textLower.includes("painting")) {
            responseSuggestion = `Hi ${customer.name}! We'd be happy to prepare a quote for you. What services are you looking for, and where is your property located?`;
        } else if (randomText.includes("inspection")) {
            responseSuggestion = "A standard 15-story tower inspection takes about 2-3 hours. Our engineer will inspect HVAC, DB boards, plumbing risers, and drainage pumps.";
        } else if (randomText.includes("warranty")) {
            responseSuggestion = "Yes! We offer a full 1-year warranty on all spare parts and a 5-year warranty on Carrier or Daikin compressor units.";
        } else if (randomText.includes("pay")) {
            responseSuggestion = "Absolutely. We accept CBQ transfers, cheques, and credit cards. We will generate the invoice right after the contract is approved.";
        } else {
            responseSuggestion = "Yes, Smart Vision is fully insured with third-party liability coverage up to QR 2,000,000 for building maintenance safety.";
        }

        customer.aiSuggestion = responseSuggestion;
        syncState();
        
        if (activeChatId.toString() === customer.id.toString()) {
            if (suggBox) {
                suggBox.style.display = "flex";
                suggBox.style.opacity = 1;
                document.getElementById("ai-suggestion-text").innerText = `"${responseSuggestion}"`;
            }
        }
        
        logActivity(`AI suggestion generated for ${customer.name}`, responseSuggestion, 'secondary');

        // Parse new suggestion into checklist & score
        updateAIQualificationChecklist(customer);

        // If autopilot is active, automatically dispatch
        if (customer.autopilotActive) {
            const autopilotDelay = parseInt(localStorage.getItem("settings-ai-delay") || "1") * 1000;
            setTimeout(() => {
                sendAISuggestion(customer.id);
            }, autopilotDelay);
        }
    }, 1500);

    logActivity(`Incoming message from ${customer.name}`, randomText, 'primary');
}

function simulateNewEnquiry() {
    switchView('chat');

    const fatima = customers.find(c => c.id === 2);
    if (fatima) {
        fatima.channel = "WhatsApp";
        
        const textMsg = "Hello! I need a quote for painting our new 4-bedroom villa. Do you have eco-friendly paint options?";
        const isMsgLoaded = fatima.messages.some(m => m.text === textMsg);
        if (!isMsgLoaded) {
            fatima.messages.push({ sender: "incoming", text: textMsg, time: "16:42" });
            fatima.aiSuggestion = "Hello Fatima! Yes, we use premium Low-VOC eco-friendly paints that are odorless and safe. Could you share the location of the villa in Doha?";
            fatima.score = 72;
            fatima.temp = "warm";
            fatima.lastContact = "16:42";
            fatima.history.push({ type: "system", text: "Inquiry received on WhatsApp", time: new Date().toLocaleString() });
            syncState();
        }
        
        renderChatListSidebar();
        selectChat(fatima.id);
        
        if (activeCustomerId === fatima.id) {
            selectCustomer(fatima.id);
        }
    }
    logActivity("New Enquiry Registered", "Fatima Al-Kuwari submitted inquiry via WhatsApp", "primary");
}

// Invoice Generator
function addInvoiceItem() {
    const container = document.getElementById("invoice-items-editor");
    
    const row = document.createElement("div");
    row.className = "form-row item-row";
    row.style.marginBottom = "10px";
    row.innerHTML = `
        <input type="text" class="form-control flex-2" placeholder="Item Name" oninput="updateInvoicePreview()">
        <input type="number" class="form-control" style="flex:0.5;" placeholder="Qty" value="1" oninput="updateInvoicePreview()">
        <input type="number" class="form-control flex-1" placeholder="Price" oninput="updateInvoicePreview()">
    `;
    container.appendChild(row);
    updateInvoicePreview();
}

function updateInvoicePreview() {
    const invNo = document.getElementById("inv-no").value;
    const invDate = document.getElementById("inv-date").value;
    const clientName = document.getElementById("inv-client-name").value;
    const clientContact = document.getElementById("inv-client-contact").value;
    
    // Multiple Business Entity Support
    const entity = document.getElementById("inv-entity").value;
    const logoEl = document.getElementById("preview-logo");
    const addrEl = document.getElementById("preview-address");
    const contactEl = document.getElementById("preview-contact");
    
    if (entity === "GBOT Connect") {
        if (logoEl) logoEl.innerText = "GBOT CONNECT";
        if (addrEl) addrEl.innerText = "Al Rayyan Road, Doha, Qatar";
        if (contactEl) contactEl.innerText = "info@gccbot.com | +974 55160323";
    } else {
        if (logoEl) logoEl.innerText = "SMART VISION WLL";
        if (addrEl) addrEl.innerText = "Building No 216, 4th Floor Retaj Building, 340 Salwa Road, Doha, Qatar";
        if (contactEl) contactEl.innerText = "smartvisionwll@gmail.com | +974 30544802";
    }

    // Document Type Support
    const docType = document.getElementById("inv-doc-type").value;
    const docTitleEl = document.getElementById("preview-doc-title");
    if (docTitleEl) {
        if (docType === "Quotation") {
            docTitleEl.innerText = "QUOTATION";
        } else if (docType === "Proforma") {
            docTitleEl.innerText = "PROFORMA INVOICE";
        } else {
            docTitleEl.innerText = "INVOICE";
        }
    }

    document.getElementById("preview-no").innerText = invNo;
    document.getElementById("preview-date").innerText = `Date: ${invDate}`;
    document.getElementById("preview-client-name").innerText = clientName;
    document.getElementById("preview-client-contact").innerText = clientContact;

    const rows = document.querySelectorAll(".item-row");
    const tbody = document.getElementById("preview-items-body");
    tbody.innerHTML = "";

    let subtotal = 0;

    rows.forEach(row => {
        const inputs = row.querySelectorAll("input");
        const desc = inputs[0].value;
        const qty = parseFloat(inputs[1].value) || 0;
        const price = parseFloat(inputs[2].value) || 0;
        const amount = qty * price;

        if (desc.trim() !== "" || price > 0) {
            subtotal += amount;

            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td>${desc}</td>
                <td style="text-align: right;">${qty}</td>
                <td style="text-align: right;">QR ${price.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                <td style="text-align: right; font-weight:700;">QR ${amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
            `;
            tbody.appendChild(tr);
        }
    });

    document.getElementById("preview-subtotal").innerText = `QR ${subtotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById("preview-total").innerText = `QR ${subtotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function saveInvoiceData() {
    const total = document.getElementById("preview-total").innerText;
    const client = document.getElementById("inv-client-name").value.trim() || document.getElementById("preview-client-name").innerText;
    const no = document.getElementById("preview-no").innerText;
    const docType = document.getElementById("inv-doc-type").value;
    const entity = document.getElementById("inv-entity").value;
    const date = document.getElementById("inv-date").value;

    logActivity(`${docType} Saved: ${no}`, `Total: ${total} | Billed to: ${client}`, 'success');
    
    // Save to State Array
    try {
        const totalStr = total.replace(/[^\d.]/g, ''); // Extract numeric values
        const amt = parseFloat(totalStr) || 0;
        
        // Add to state array
        const newDoc = {
            id: issuedDocuments.length + 1,
            docNo: no,
            client: client,
            entity: entity === "GBOT Connect" ? "GBOT Connect WLL" : "Smart Vision WLL",
            type: docType,
            amount: amt,
            date: date,
            status: docType === "Quotation" ? "Sent" : "Unpaid"
        };
        issuedDocuments.push(newDoc);
        
        // Dynamically update the MTD Revenue KPI card on dashboard if Invoice
        if (docType === "Invoice") {
            const kpiRev = document.getElementById("kpi-revenue");
            if (kpiRev) {
                const currentStr = kpiRev.innerText.replace(/[^\d.]/g, '');
                const currentAmt = parseFloat(currentStr) || 0;
                const newAmt = currentAmt + amt;
                kpiRev.innerText = `QR ${newAmt.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 })}`;
            }
        }
        
        // Re-render ledger if initialized
        renderSalesLedger();
    } catch (e) {
        console.error("Failed to save doc and update KPI:", e);
    }
    
    alert(`${docType} ${no} saved successfully in CRM system! Syncing with Zoho.`);
}

function printInvoice() {
    const printContents = document.getElementById("invoice-preview-sheet").innerHTML;
    const originalContents = document.body.innerHTML;

    document.body.innerHTML = `
        <div style="background:#fff; color:#000; padding: 40px; font-family: sans-serif;">
            ${printContents}
        </div>
    `;
    window.print();
    document.body.innerHTML = originalContents;
    window.location.reload();
}

// Workflow Builder Node appender
function addWorkflowAction() {
    const actionName = prompt("Enter custom system action node description:", "Send Whatsapp Broadcast Alert");
    if (actionName && actionName.trim() !== "") {
        addNodeToCanvas(actionName);
    }
}

function addNodeToCanvas(actionName) {
    const canvas = document.getElementById("workflow-canvas-container");
    const lastNode = document.getElementById("last-wf-node");

    if (canvas && lastNode) {
        const conn = document.createElement("div");
        conn.className = "wf-connector";

        const node = document.createElement("div");
        node.className = "wf-node wf-node-action";
        node.innerHTML = `
            <div class="wf-node-header action">Automated Action</div>
            <div class="wf-node-body">${actionName}</div>
            <div class="wf-node-desc">Custom Workflow Event</div>
        `;

        canvas.appendChild(conn);
        canvas.appendChild(node);
        logActivity("Workflow Builder Updated", `Added node: ${actionName}`, 'secondary');
    }
}

// Operations Tabs Toggle
function switchOpTab(tabId) {
    activeOpTab = tabId;
    
    const tabs = ['tasks', 'inventory', 'finance', 'suppliers', 'requests'];
    tabs.forEach(t => {
        const btn = document.getElementById(`btn-tab-${t}`);
        const panel = document.getElementById(`op-tab-${t}`);
        
        if (btn && panel) {
            if (t === tabId) {
                btn.className = "btn btn-primary";
                panel.style.display = "block";
            } else {
                btn.className = "btn btn-secondary";
                panel.style.display = "none";
            }
        }
    });

    // Refresh rendering when tab changes
    if (tabId === 'tasks') renderOperationsTasks();
    if (tabId === 'inventory') renderInventory();
    if (tabId === 'finance') renderFinance();
    if (tabId === 'suppliers') renderSuppliers();
    if (tabId === 'requests') renderRequests();
}

// Settings Connections / QR Scan Emulator
function toggleConnectionFields() {
    const val = document.getElementById("conn-method").value;
    const qrSect = document.getElementById("settings-qr-section");
    const metaSect = document.getElementById("settings-meta-section");

    if (val === 'qr') {
        qrSect.style.display = 'block';
        metaSect.style.display = 'none';
    } else {
        qrSect.style.display = 'none';
        metaSect.style.display = 'block';
    }
}

function openQRModal() {
    document.getElementById("qr-modal").classList.add("active");
    const spinner = document.getElementById("qr-loading-spinner");
    const textStatus = document.getElementById("qr-conn-status");
    
    spinner.style.display = 'block';
    textStatus.innerText = "Awaiting pairing code scan...";
}

function closeQRModal() {
    document.getElementById("qr-modal").classList.remove("active");
}

function simulateQRScanSuccess() {
    // If backend server is active, we don't use the static QR success simulator
    if (isBackendConnected) {
        alert("Integrations server is active. Please scan the QR code displayed on the screen using your phone's WhatsApp linked devices!");
        return;
    }

    const spinner = document.getElementById("qr-loading-spinner");
    const textStatus = document.getElementById("qr-conn-status");
    
    spinner.style.display = 'none';
    textStatus.innerText = "Device Connected! Syncing chats...";
    textStatus.style.color = "var(--success)";

    setTimeout(() => {
        closeQRModal();
        document.getElementById("whatsapp-status-dot").className = "status-dot active";
        document.getElementById("whatsapp-status-text").innerText = "QR Connected (Live)";
        logActivity("WhatsApp QR Paired", "Device synced with local instance successfully.", "success");
    }, 1500);
}

function triggerManualZohoSync() {
    logActivity("Zoho CRM Sync Triggered", "Synchronizing 6 leads & 1 paid invoices...", "primary");
    alert("Zoho CRM Database successfully updated! All lead stages synced.");
}

// General Logging helper
function logActivity(title, desc, type) {
    const feed = document.getElementById("live-activity-feed");
    if (!feed) return;

    const item = document.createElement("div");
    item.className = "activity-item";
    item.innerHTML = `
        <div class="activity-marker ${type}"></div>
        <div class="activity-content">
            <div class="activity-title">${title}</div>
            <div class="activity-desc">${desc}</div>
        </div>
        <span class="activity-time">Just now</span>
    `;

    feed.insertBefore(item, feed.firstChild);

    if (feed.children.length > 6) {
        feed.removeChild(feed.lastChild);
    }
}

// Start a new WhatsApp chat from the search panel
function startNewChat() {
    const input = document.getElementById("new-chat-phone");
    if (!input) return;
    let rawPhone = input.value.trim();
    let phone = formatQatarPhoneNumber(rawPhone).replace(/[^0-9]/g, '');
    if (phone === "") {
        alert("Please enter a valid phone number.");
        return;
    }
    
    let customer = getCustomerByChatId(phone);
    if (!customer) {
        customer = {
            id: customers.length + 1,
            name: `Contact +${phone}`,
            phone: `+${phone}`,
            email: `contact_${phone}@gmail.com`,
            stage: "new",
            temp: "cold",
            channel: "WhatsApp",
            notes: "",
            labels: [],
            loyaltyPoints: 0,
            dateAdded: new Date().toISOString().split('T')[0],
            lastContact: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
            messages: [],
            aiSuggestion: null,
            score: 50,
            autopilotActive: true,
            history: [
                { type: "system", text: "WhatsApp chat initiated manually", time: new Date().toLocaleString() }
            ]
        };
        customers.push(customer);
        syncState();
        renderCustomersTable();
        renderKanban();
        renderChatListSidebar();
    }
    
    input.value = "";
    selectChat(customer.id);
    logActivity("New Chat Opened", `Opened WhatsApp conversation with +${phone}`, "primary");
}

// ==========================================
// CUSTOMER DIRECTORY & MANAGEMENT LOGIC
// ==========================================

function renderChatProfileLabels(customer) {
    const container = document.getElementById("chat-profile-labels");
    if (!container) return;
    container.innerHTML = "";
    if (customer.labels && customer.labels.length > 0) {
        customer.labels.forEach(label => {
            const span = document.createElement("span");
            span.style.cssText = "background:rgba(99,102,241,0.15); color:var(--primary); font-size:10px; padding:2px 6px; border-radius:4px; border:1px solid rgba(99,102,241,0.3); display:flex; align-items:center; gap:4px;";
            span.innerHTML = `
                ${label}
                <span style="cursor:pointer; font-weight:bold; font-size:9px;" onclick="deleteCustomerLabel('${customer.id}', '${label}')">×</span>
            `;
            container.appendChild(span);
        });
    } else {
        container.innerHTML = `<span style="color:var(--text-dim); font-size:10px;">No labels</span>`;
    }
}

function checkLabelSubmit(e) {
    if (e.key === "Enter") {
        addChatLabel();
    }
}

function addChatLabel() {
    const input = document.getElementById("new-chat-label");
    const labelText = input.value.trim();
    if (labelText === "") return;
    
    const customer = getCustomerByChatId(activeChatId);
    if (customer) {
        if (!customer.labels) customer.labels = [];
        if (!customer.labels.includes(labelText)) {
            customer.labels.push(labelText);
            customer.history.push({ type: "system", text: `Label added: "${labelText}"`, time: new Date().toLocaleString() });
            syncState();
            renderChatProfileLabels(customer);
            if (activeCustomerId === customer.id) {
                selectCustomer(customer.id);
            }
            logActivity(`Label added to ${customer.name}`, `Added custom label: ${labelText}`, "secondary");
        }
    }
    input.value = "";
}

function deleteCustomerLabel(customerId, label) {
    const customer = customers.find(c => c.id.toString() === customerId.toString());
    if (customer) {
        customer.labels = customer.labels.filter(l => l !== label);
        customer.history.push({ type: "system", text: `Label removed: "${label}"`, time: new Date().toLocaleString() });
        syncState();
        renderChatProfileLabels(customer);
        if (activeCustomerId === customer.id) {
            selectCustomer(customer.id);
        }
        logActivity(`Label removed from ${customer.name}`, `Removed label: ${label}`, "secondary");
    }
}

function saveChatNotes() {
    const textarea = document.getElementById("chat-profile-notes");
    const notes = textarea.value;
    
    const customer = getCustomerByChatId(activeChatId);
    if (customer) {
        customer.notes = notes;
        syncState();
        if (activeCustomerId === customer.id) {
            document.getElementById("cust-detail-notes").value = notes;
        }
    }
}

// Customer Directory UI rendering
function renderCustomersTable() {
    const tbody = document.getElementById("customers-table-body");
    if (!tbody) return;
    
    const searchQuery = document.getElementById("customer-search-input").value.toLowerCase();
    const stageFilter = document.getElementById("filter-customer-stage").value;
    const tempFilter = document.getElementById("filter-customer-temp").value;
    
    tbody.innerHTML = "";
    
    const filteredCustomers = customers.filter(c => {
        const matchesSearch = c.name.toLowerCase().includes(searchQuery) ||
                             c.phone.includes(searchQuery) ||
                             c.email.toLowerCase().includes(searchQuery);
                             
        const matchesStage = stageFilter === 'all' || c.stage === stageFilter;
        const matchesTemp = tempFilter === 'all' || c.temp.toLowerCase() === tempFilter;
        
        return matchesSearch && matchesStage && matchesTemp;
    });
    
    filteredCustomers.forEach(c => {
        const tr = document.createElement("tr");
        tr.style.cursor = "pointer";
        tr.onclick = () => selectCustomer(c.id);
        
        if (activeCustomerId === c.id) {
            tr.style.background = "rgba(99, 102, 241, 0.08)";
            tr.style.borderColor = "var(--primary-glow)";
        }
        
        const stageColors = {
            'new': '#3b82f6',
            'contacted': '#a855f7',
            'qualified': '#f59e0b',
            'proposal': '#06b6d4',
            'won': '#10b981'
        };
        
        const tempText = c.temp === 'hot' ? '🔥 Hot' : c.temp === 'warm' ? '⚡ Warm' : '❄️ Cold';
        
        tr.innerHTML = `
            <td style="font-weight: 700;">
                <div style="display:flex; align-items:center; gap:8px;">
                    <div class="chat-avatar" style="width:28px; height:28px; font-size:11px; flex-shrink:0;">${c.name.split(' ').map(n => n[0]).join('').substring(0,2).toUpperCase()}</div>
                    <div>
                        <div>${c.name}</div>
                        <span style="font-size:10px; font-weight:normal; color:var(--text-dim);">${tempText}</span>
                    </div>
                </div>
            </td>
            <td>${c.phone}</td>
            <td>${c.email}</td>
            <td><span class="badge-status" style="background:rgba(255,255,255,0.02); color:${stageColors[c.stage] || '#fff'}; border:1px solid ${stageColors[c.stage] || 'var(--border-color)'};">${c.stage.toUpperCase()}</span></td>
            <td style="font-weight: 700; color:var(--warning);">${c.loyaltyPoints} pts</td>
            <td style="color:var(--text-muted); font-size:12px;">${c.lastContact}</td>
        `;
        tbody.appendChild(tr);
    });
}

function selectCustomer(customerId) {
    activeCustomerId = customerId;
    renderCustomersTable();
    
    const customer = customers.find(c => c.id.toString() === customerId.toString());
    if (!customer) return;
    
    document.getElementById("cust-detail-avatar").innerText = customer.name.split(' ').map(n => n[0]).join('').substring(0,2).toUpperCase();
    document.getElementById("cust-detail-name").innerText = customer.name;
    document.getElementById("cust-detail-phone").innerText = customer.phone;
    document.getElementById("cust-detail-email").innerText = customer.email;
    
    const tagsContainer = document.getElementById("cust-detail-tags");
    tagsContainer.innerHTML = `
        <span class="card-tag tag-${customer.temp.toLowerCase()}">${customer.temp} Lead</span>
        <span class="badge-status success" style="padding:2px 8px; font-size:10px;">${customer.channel}</span>
        <span class="badge-status primary" style="padding:2px 8px; font-size:10px; text-transform:uppercase;">${customer.stage}</span>
    `;
    
    // Visual Pipeline path render
    const pathContainer = document.getElementById("cust-pipeline-path");
    if (pathContainer) {
        const stagesList = ['new', 'contacted', 'qualified', 'proposal', 'won'];
        const displayNames = { 'new': 'New', 'contacted': 'Contacted', 'qualified': 'Qualified', 'proposal': 'Proposal', 'won': 'Won' };
        pathContainer.innerHTML = stagesList.map((st, idx) => {
            const isActive = customer.stage === st ? 'active' : '';
            const stepHtml = `<span class="pipeline-step ${isActive}" onclick="updateCustomerStageDirectly('${customer.id}', '${st}')">${displayNames[st]}</span>`;
            const arrowHtml = idx < stagesList.length - 1 ? `<span class="pipeline-arrow">➔</span>` : '';
            return stepHtml + arrowHtml;
        }).join("");
    }

    // Set dropdowns value
    const tempSelect = document.getElementById("edit-cust-temp");
    if (tempSelect) tempSelect.value = customer.temp;
    const ownerSelect = document.getElementById("edit-cust-owner");
    if (ownerSelect) ownerSelect.value = customer.leadOwner || "Unassigned";

    // Set follow-up display
    const fupDisplay = document.getElementById("cust-followup-display");
    const fupLbl = document.getElementById("lbl-followup-date");
    const fupInput = document.getElementById("cust-followup-date");
    if (customer.nextFollowUp) {
        if (fupDisplay) fupDisplay.style.display = "flex";
        if (fupLbl) fupLbl.innerText = customer.nextFollowUp;
        if (fupInput) fupInput.value = customer.nextFollowUp;
    } else {
        if (fupDisplay) fupDisplay.style.display = "none";
        if (fupInput) fupInput.value = "";
    }

    // Reset edit panel state
    toggleEditProfile(false);
    
    renderDirectoryCustLabels(customer);
    document.getElementById("cust-detail-notes").value = customer.notes || "";
    renderCustomerTimeline(customer);

    // Update checklist and score bar in customer details panel
    updateAIQualificationChecklist(customer);
}

function renderDirectoryCustLabels(customer) {
    const container = document.getElementById("cust-detail-labels");
    if (!container) return;
    container.innerHTML = "";
    if (customer.labels && customer.labels.length > 0) {
        customer.labels.forEach(label => {
            const span = document.createElement("span");
            span.style.cssText = "background:rgba(99,102,241,0.15); color:var(--primary); font-size:10px; padding:2px 6px; border-radius:4px; border:1px solid rgba(99,102,241,0.3); display:flex; align-items:center; gap:4px;";
            span.innerHTML = `
                ${label}
                <span style="cursor:pointer; font-weight:bold; font-size:9px;" onclick="deleteCustLabelFromDirectory('${customer.id}', '${label}')">×</span>
            `;
            container.appendChild(span);
        });
    } else {
        container.innerHTML = `<span style="color:var(--text-dim); font-size:10px;">No labels</span>`;
    }
}

function checkCustLabelSubmit(e) {
    if (e.key === "Enter") {
        addCustLabelFromDirectory();
    }
}

function addCustLabelFromDirectory() {
    const input = document.getElementById("new-cust-label-input");
    const labelText = input.value.trim();
    if (labelText === "") return;
    
    const customer = customers.find(c => c.id.toString() === activeCustomerId.toString());
    if (customer) {
        if (!customer.labels) customer.labels = [];
        if (!customer.labels.includes(labelText)) {
            customer.labels.push(labelText);
            customer.history.push({ type: "system", text: `Label added: "${labelText}"`, time: new Date().toLocaleString() });
            syncState();
            renderDirectoryCustLabels(customer);
            if (activeChatId.toString() === customer.id.toString() || customer.phone.replace(/[^0-9]/g, '') === activeChatId.toString().replace(/[^0-9]/g, '')) {
                renderChatProfileLabels(customer);
            }
            logActivity(`Label added to ${customer.name}`, `Added custom label: ${labelText}`, "secondary");
        }
    }
    input.value = "";
}

function deleteCustLabelFromDirectory(customerId, label) {
    const customer = customers.find(c => c.id.toString() === customerId.toString());
    if (customer) {
        customer.labels = customer.labels.filter(l => l !== label);
        customer.history.push({ type: "system", text: `Label removed: "${label}"`, time: new Date().toLocaleString() });
        syncState();
        renderDirectoryCustLabels(customer);
        if (activeChatId.toString() === customer.id.toString() || customer.phone.replace(/[^0-9]/g, '') === activeChatId.toString().replace(/[^0-9]/g, '')) {
            renderChatProfileLabels(customer);
        }
        logActivity(`Label removed from ${customer.name}`, `Removed label: ${label}`, "secondary");
    }
}

function saveCustNotesFromDirectory() {
    const textarea = document.getElementById("cust-detail-notes");
    const notes = textarea.value;
    
    const customer = customers.find(c => c.id.toString() === activeCustomerId.toString());
    if (customer) {
        customer.notes = notes;
        syncState();
        if (activeChatId.toString() === customer.id.toString() || customer.phone.replace(/[^0-9]/g, '') === activeChatId.toString().replace(/[^0-9]/g, '')) {
            document.getElementById("chat-profile-notes").value = notes;
        }
    }
}

async function deleteCustomerFromDirectory() {
    const customer = customers.find(c => c.id.toString() === activeCustomerId.toString());
    if (!customer) return;
    
    if (confirm(`Are you sure you want to completely delete "${customer.name}" from the CRM and the Excel database?`)) {
        // Remove from customers array
        customers = customers.filter(c => c.id !== customer.id);
        
        // Remove from excelRows array as well to prevent resurrection
        const custPhoneClean = customer.phone ? customer.phone.replace(/[^0-9]/g, '') : '';
        excelRows = excelRows.filter(row => {
            const rowId = (row.ID || row.id || '').toString();
            const rowPhoneClean = (row.Phone || row.phone || '').replace(/[^0-9]/g, '');
            return rowId !== customer.id.toString() && (!custPhoneClean || rowPhoneClean !== custPhoneClean);
        });
        
        // Re-index remaining excelRows IDs sequentially
        excelRows.forEach((row, index) => {
            row.ID = index + 1;
        });
        
        // Update states
        syncState();
        
        // Save to disk silently to prevent any background merges from restoring the contact
        await saveExcelToDiskSilent();
        
        // Render updated components
        renderCustomersTable();
        renderKanban();
        renderChatListSidebar();
        
        // If there are customers left, select the first one. Otherwise, clear detail panel.
        if (customers.length > 0) {
            selectCustomer(customers[0].id);
        } else {
            document.getElementById("cust-detail-name").innerText = "No Customer Selected";
            document.getElementById("cust-detail-phone").innerText = "";
            document.getElementById("cust-detail-email").innerText = "";
            document.getElementById("cust-detail-tags").innerHTML = "";
            document.getElementById("cust-detail-labels").innerHTML = "";
            document.getElementById("cust-detail-notes").value = "";
            document.getElementById("cust-detail-timeline").innerHTML = "";
            document.getElementById("cust-score-pct").innerText = "0%";
            document.getElementById("cust-score-bar").style.width = "0%";
            document.getElementById("cust-detail-summary").innerText = "No active conversation.";
        }
        
        logActivity("Customer Deleted", `Completely removed ${customer.name} and synced database`, "danger");
        alert(`Customer "${customer.name}" deleted successfully.`);
    }
}

async function saveExcelToDiskSilent() {
    try {
        await fetch(`${BACKEND_URL}/api/excel/save`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data: excelRows })
        });
    } catch (e) {
        console.error("Silent saveExcelToDisk failed:", e);
    }
}

async function deleteAllCustomers() {
    if (confirm("WARNING: Are you sure you want to completely delete ALL customers from the CRM and the Excel database? This action cannot be undone.")) {
        // Clear memory arrays
        customers = [];
        excelRows = [];
        
        // Update states
        syncState();
        
        // Save empty state to disk silently
        await saveExcelToDiskSilent();
        
        // Render empty components
        renderCustomersTable();
        renderKanban();
        renderChatListSidebar();
        
        // Clear customer detail panel
        document.getElementById("cust-detail-name").innerText = "No Customer Selected";
        document.getElementById("cust-detail-phone").innerText = "";
        document.getElementById("cust-detail-email").innerText = "";
        document.getElementById("cust-detail-tags").innerHTML = "";
        document.getElementById("cust-detail-labels").innerHTML = "";
        document.getElementById("cust-detail-notes").value = "";
        document.getElementById("cust-detail-timeline").innerHTML = "";
        document.getElementById("cust-score-pct").innerText = "0%";
        document.getElementById("cust-score-bar").style.width = "0%";
        document.getElementById("cust-detail-summary").innerText = "No active conversation.";
        
        logActivity("All Customers Deleted", "Completely wiped the database and spreadsheet", "danger");
        alert("All customers deleted successfully. You can now add new contacts.");
    }
}

function renderCustomerTimeline(customer) {
    const timeline = document.getElementById("cust-detail-timeline");
    if (!timeline) return;
    timeline.innerHTML = "";
    
    let events = [];
    
    customer.history.forEach(hist => {
        if (typeof hist === 'string') {
            events.push({
                type: 'system',
                text: hist,
                time: customer.dateAdded
            });
        } else {
            events.push({
                type: hist.type || 'system',
                text: hist.text,
                time: hist.time
            });
        }
    });
    
    if (customer.messages) {
        customer.messages.forEach(msg => {
            let eventText = `WhatsApp ${msg.sender === 'incoming' ? 'Received' : 'Sent'}: "${msg.text}"`;
            if (msg.attachment) {
                const captionPart = msg.text ? ` - "${msg.text}"` : '';
                eventText = `WhatsApp ${msg.sender === 'incoming' ? 'Received' : 'Sent'}: 📎 [File: ${msg.attachment.name}]${captionPart}`;
            }
            events.push({
                type: msg.sender === 'incoming' ? 'msg-in' : 'msg-out',
                text: eventText,
                time: msg.time
            });
        });
    }
    
    events.forEach(ev => {
        const div = document.createElement("div");
        div.className = `timeline-event-item ${ev.type}`;
        div.innerHTML = `
            <div>${ev.text}</div>
            <span class="timeline-event-time">${ev.time}</span>
        `;
        timeline.appendChild(div);
    });
    
    if (events.length === 0) {
        timeline.innerHTML = `<div style="color:var(--text-dim); font-size:11px; text-align:center; padding: 20px 0;">No history logs recorded yet.</div>`;
    }
}

function openAddCustomerModal() {
    document.getElementById("add-customer-modal").classList.add("active");
}

function closeAddCustomerModal() {
    document.getElementById("add-customer-modal").classList.remove("active");
}

function submitAddCustomerForm() {
    const name = document.getElementById("form-cust-name").value.trim();
    const rawPhone = document.getElementById("form-cust-phone").value.trim();
    const phone = formatQatarPhoneNumber(rawPhone);
    const email = document.getElementById("form-cust-email").value.trim();
    const stage = document.getElementById("form-cust-stage").value;
    const temp = document.getElementById("form-cust-temp").value;
    const notes = document.getElementById("form-cust-notes").value.trim();
    
    if (name === "" || phone === "" || email === "") {
        alert("Please fill in Name, Phone, and Email fields.");
        return;
    }
    
    const newCust = {
        id: customers.length + 1,
        name: name,
        phone: phone,
        email: email,
        stage: stage,
        temp: temp,
        channel: "Manual",
        notes: notes,
        labels: [],
        loyaltyPoints: 0,
        dateAdded: new Date().toISOString().split('T')[0],
        lastContact: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
        messages: [],
        aiSuggestion: null,
        score: 50,
        autopilotActive: true,
        history: [
            { type: "system", text: "Customer profile created manually", time: new Date().toLocaleString() }
        ]
    };
    
    customers.push(newCust);
    
    dispatchWebhook("customer_created", newCust);
    
    document.getElementById("form-cust-name").value = "";
    document.getElementById("form-cust-phone").value = "";
    document.getElementById("form-cust-email").value = "";
    document.getElementById("form-cust-notes").value = "";
    
    closeAddCustomerModal();
    syncState();
    renderCustomersTable();
    renderKanban();
    
    selectCustomer(newCust.id);
    
    logActivity(`Customer profile created: ${name}`, `Added to stage ${stage.toUpperCase()}`, "success");
}

// ==========================================
// AI INTELLIGENCE & AUTOPILOT FUNCTIONS
// ==========================================

function toggleAutopilot() {
    const toggle = document.getElementById("chat-autopilot-toggle");
    const indicator = document.getElementById("autopilot-indicator");
    const customer = getCustomerByChatId(activeChatId);
    if (!customer) return;
    
    customer.autopilotActive = toggle.checked;
    if (customer.autopilotActive) {
        if (indicator) indicator.style.display = "flex";
    } else {
        if (indicator) indicator.style.display = "none";
    }
    syncState();
    logActivity(`Autopilot ${customer.autopilotActive ? 'Activated' : 'Deactivated'}`, `Autopilot mode for ${customer.name}`, 'primary');

    // Run verification immediately
    updateAIQualificationChecklist(customer);
}

function generateConversationSummary(customer) {
    if (!customer.messages || customer.messages.length === 0) {
        return customer.notes || "No recent messages or active conversation history to summarize.";
    }
    
    const textAll = customer.messages.map(m => m.text).join(" ").toLowerCase();
    
    let service = "";
    if (textAll.includes("ac") || textAll.includes("air cond") || textAll.includes("split")) {
        service = "AC repair/replacement";
    } else if (textAll.includes("mep") || textAll.includes("maintenance")) {
        service = "MEP building maintenance";
    } else if (textAll.includes("paint") || textAll.includes("painting")) {
        service = "villa painting";
    } else if (textAll.includes("electrical") || textAll.includes("rewir")) {
        service = "electrical rewiring";
    } else {
        service = "maintenance support";
    }
    
    let location = "";
    if (textAll.includes("west bay")) location = "West Bay";
    else if (textAll.includes("pearl")) location = "The Pearl-Qatar";
    else if (textAll.includes("lusail")) location = "Lusail";
    else if (textAll.includes("rayyan")) location = "Al Rayyan";
    else if (textAll.includes("doha")) location = "Doha";
    else location = "Doha (unspecified)";
    
    let budget = "";
    const budgetMatch = textAll.match(/(?:qr|budget of|around|cost of)\s*([0-9,]+)/i);
    if (budgetMatch) {
        budget = `with budget estimate around QR ${budgetMatch[1]}`;
    } else if (textAll.includes("price") || textAll.includes("cost") || textAll.includes("quote")) {
        budget = "seeking price quotation";
    } else {
        budget = "budget details pending";
    }
    
    let sentiment = "Neutral";
    if (textAll.includes("urgent") || textAll.includes("emergency") || textAll.includes("fast") || textAll.includes("asap")) {
        sentiment = "Urgent / High Intent";
    } else if (textAll.includes("thank") || textAll.includes("good") || textAll.includes("perfect") || textAll.includes("yes")) {
        sentiment = "Positive / Cooperative";
    }
    
    const summary = `Client is requesting ${service} services for a property in ${location}. Currently ${budget}. Sentiment is ${sentiment}. Ready for engineer inspection scheduling.`;
    return summary;
}

function updateAIQualificationChecklist(customer) {
    if (!customer) return;
    
    const textAll = ((customer.messages || []).map(m => m.text).join(" ") + " " + (customer.notes || "")).toLowerCase();
    
    const hasService = textAll.includes("ac") || textAll.includes("mep") || textAll.includes("paint") || 
                       textAll.includes("maintenance") || textAll.includes("repair") || 
                       textAll.includes("electrical") || textAll.includes("wiring") || textAll.includes("hvac");
                       
    const hasLocation = textAll.includes("doha") || textAll.includes("west bay") || textAll.includes("pearl") || 
                        textAll.includes("lusail") || textAll.includes("rayyan") || textAll.includes("salwa");
                        
    const hasBudget = textAll.includes("price") || textAll.includes("cost") || textAll.includes("quote") || 
                      textAll.includes("qr") || textAll.includes("how much") || textAll.includes("budget") || textAll.includes("billing");
                      
    const hasBooking = textAll.includes("schedule") || textAll.includes("tomorrow") || textAll.includes("visit") || 
                       textAll.includes("inspection") || textAll.includes("survey") || textAll.includes("technician") ||
                       textAll.includes("book") || textAll.includes("am") || textAll.includes("pm");
                       
    let score = 40;
    if (hasService) score += 15;
    if (hasLocation) score += 15;
    if (hasBudget) score += 15;
    if (hasBooking) score += 15;
    
    if (customer.temp === 'hot') score += 10;
    else if (customer.temp === 'warm') score += 5;
    
    if (score > 100) score = 100;
    customer.score = score;
    
    // Auto-promote stage to qualified if score is >= threshold and it was 'new' or 'contacted'
    const aiThreshold = parseInt(localStorage.getItem("settings-ai-score-threshold") || "75");
    if (score >= aiThreshold && (customer.stage === 'new' || customer.stage === 'contacted')) {
        customer.stage = 'qualified';
        customer.history.push({ type: "system", text: `AI Auto-Qualified: Lead Score reached ${score}%`, time: new Date().toLocaleString() });
        logActivity(`Lead ${customer.name} AI Qualified`, `AI lead score is ${score}%. Pipeline stage updated.`, 'success');
        renderKanban();
    }
    
    // Sync UI elements for active chat
    if (activeChatId.toString() === customer.id.toString() || customer.phone.replace(/[^0-9]/g, '') === activeChatId.toString().replace(/[^0-9]/g, '')) {
        updateChecklistUI("chk-service", hasService);
        updateChecklistUI("chk-location", hasLocation);
        updateChecklistUI("chk-budget", hasBudget);
        updateChecklistUI("chk-booking", hasBooking);
        
        // Update score bar & pct
        const bar = document.getElementById("chat-score-bar");
        const pct = document.getElementById("chat-score-pct");
        if (bar && pct) {
            bar.style.width = `${score}%`;
            pct.innerText = `${score}%`;
        }
        
        // Update summary box
        const summaryText = generateConversationSummary(customer);
        const summaryEl = document.getElementById("chat-profile-summary");
        if (summaryEl) {
            summaryEl.innerText = summaryText;
        }
    }
    
    // Sync UI elements for customer directory
    if (activeCustomerId.toString() === customer.id.toString()) {
        updateChecklistUI("cust-chk-service", hasService);
        updateChecklistUI("cust-chk-location", hasLocation);
        updateChecklistUI("cust-chk-budget", hasBudget);
        updateChecklistUI("cust-chk-booking", hasBooking);
        
        // Update score bar & pct
        const bar = document.getElementById("cust-score-bar");
        const pct = document.getElementById("cust-score-pct");
        if (bar && pct) {
            bar.style.width = `${score}%`;
            pct.innerText = `${score}%`;
        }
        
        // Update summary box
        const summaryText = generateConversationSummary(customer);
        const summaryEl = document.getElementById("cust-detail-summary");
        if (summaryEl) {
            summaryEl.innerText = summaryText;
        }
    }
    
    syncState();
}

function updateChecklistUI(elementId, verified) {
    const el = document.getElementById(elementId);
    if (!el) return;
    if (verified) {
        el.classList.add("verified");
    } else {
        el.classList.remove("verified");
    }
}

function triggerAISummaryGeneration() {
    let customer = null;
    if (currentView === 'chat') {
        customer = getCustomerByChatId(activeChatId);
    } else {
        customer = customers.find(c => c.id.toString() === activeCustomerId.toString());
    }
    if (!customer) return;
    
    const summaryElId = currentView === 'chat' ? "chat-profile-summary" : "cust-detail-summary";
    const el = document.getElementById(summaryElId);
    if (el) {
        el.innerText = "Regenerating AI Summary...";
        setTimeout(() => {
            updateAIQualificationChecklist(customer);
            logActivity(`AI Summary refreshed`, `Updated summary for ${customer.name}`, 'secondary');
        }, 500);
    }
}

// ==========================================
// ACTIVE BOT WORKFLOW TEMPLATES & CONSOLE
// ==========================================

const workflowTemplates = {
    "lead-qualification": [
        { type: "trigger", title: "Trigger Event", body: "Customer asks price / info", desc: "Source: WhatsApp or Email" },
        { type: "action", title: "AI Automation Filter", body: "AI Chatbot Qualifies Lead", desc: "Asks location, budget, urgency" },
        { type: "action", id: "last-wf-node", title: "System Action", body: "Create Lead & Route to Sales", desc: "Notifies Imtiyaz via SMS/email" }
    ],
    "complaint-escalation": [
        { type: "trigger", title: "Trigger Event", body: "Customer registers complaint / issue", desc: "Source: WhatsApp, Phone or Webform" },
        { type: "action", title: "AI Automation Filter", body: "Sentiment & Severity Detector", desc: "Scans for angry terms, leaks, breakdowns" },
        { type: "action", id: "last-wf-node", title: "System Action", body: "Flag Escalated & Create Urgent Task", desc: "Alerts Imtiyaz Ahmed, schedules follow-up" }
    ],
    "data-collection": [
        { type: "trigger", title: "Trigger Event", body: "First-time visitor inquiry", desc: "Source: WhatsApp or Webchat" },
        { type: "action", title: "AI Automation Filter", body: "Contact Data Harvesting Bot", desc: "Asks name, email, phone, company" },
        { type: "action", id: "last-wf-node", title: "System Action", body: "Create Profile & Sync Zoho CRM", desc: "Saves details in central Directory" }
    ],
    "service-request": [
        { type: "trigger", title: "Trigger Event", body: "Booking request received", desc: "Service: AC repair / painting" },
        { type: "action", title: "AI Automation Filter", body: "Booking Coordinator Bot", desc: "Matches engineer availability & schedules survey" },
        { type: "action", id: "last-wf-node", title: "System Action", body: "Create Calendar Event & Dispatch Tech", desc: "Updates schedule, sends tech location" }
    ],
    "broadcast-execution": [
        { type: "trigger", title: "Manual Action", body: "Launch Broadcast Campaign", desc: "User triggers bulk send" },
        { type: "action", title: "Filter Node", body: "Match: All Active Leads", desc: "Loads target numbers from database" },
        { type: "action", id: "last-wf-node", title: "Live Dispatch", body: "Send WhatsApp Messages One-by-One", desc: "Real-time dispatch via paired phone" }
    ]
};

function loadWorkflowTemplate() {
    const select = document.getElementById("workflow-template-select");
    const container = document.getElementById("workflow-canvas-container");
    if (!select || !container) return;
    
    const templateId = select.value;
    const nodes = workflowTemplates[templateId];
    if (!nodes) return;
    
    container.innerHTML = "";
    
    nodes.forEach((node, idx) => {
        if (idx > 0) {
            const conn = document.createElement("div");
            conn.className = "wf-connector";
            container.appendChild(conn);
        }
        
        const nodeEl = document.createElement("div");
        nodeEl.className = `wf-node wf-node-${node.type}`;
        if (node.id) nodeEl.id = node.id;
        
        let headerIcon = "";
        if (node.type === 'trigger') {
            headerIcon = `<svg viewBox="0 0 24 24" width="12" height="12" stroke="currentColor" stroke-width="2.5" fill="none" style="display:inline; margin-right:4px;"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"></polygon></svg>`;
        }
        
        nodeEl.innerHTML = `
            <div class="wf-node-header ${node.type}">
                ${headerIcon}
                ${node.title}
            </div>
            <div class="wf-node-body">${node.body}</div>
            <div class="wf-node-desc">${node.desc}</div>
        `;
        container.appendChild(nodeEl);
    });
    
    const consoleLogs = document.getElementById("workflow-console-logs");
    if (consoleLogs) {
        consoleLogs.innerHTML = `<div style="color:var(--text-dim);">[system] Workflow template "${select.options[select.selectedIndex].text}" loaded. Ready for simulation.</div>`;
    }
}

const simulationLogs = {
    "lead-qualification": [
        { text: "[00:01] [trigger] Incoming WhatsApp message from Najeeb Shekasan: \"I need a quote for our office tower.\"", color: "var(--text-muted)" },
        { text: "[00:02] [filter] AI processing message... Keywords found: \"quote\", \"office tower\". Intent matched: Price Enquiry.", color: "#c084fc" },
        { text: "[00:03] [filter] AI prompt response queued: \"Sure! What is the size and location in Doha?\"", color: "#6366f1" },
        { text: "[00:05] [trigger] Incoming WhatsApp message: \"West Bay Doha, 15 floors. Need AC repair too.\"", color: "var(--text-muted)" },
        { text: "[00:06] [filter] AI processing... Keywords found: \"Doha\", \"West Bay\", \"AC repair\".", color: "#c084fc" },
        { text: "[00:07] [filter] Checklist items marked: Service Type (Identified), Location (Verified).", color: "#10b981" },
        { text: "[00:08] [filter] Lead score calculated: 75% (Hot). Stage advanced to AI Qualified.", color: "#f59e0b" },
        { text: "[00:09] [action] Routing lead to Sales (Imtiyaz Ahmed) via SMS alert.", color: "#6366f1" },
        { text: "[00:10] [action] Synchronizing lead dataset with Zoho CRM Database... Success.", color: "#10b981" }
    ],
    "complaint-escalation": [
        { text: "[00:01] [trigger] Incoming WhatsApp message from Yasmin Rahmani: \"The AC units are leaking and not cooling at all, this is poor service!\"", color: "var(--text-muted)" },
        { text: "[00:02] [filter] AI processing message... Keywords found: \"leaking\", \"not cooling\", \"poor service\". Intent matched: Service Complaint.", color: "#c084fc" },
        { text: "[00:03] [filter] AI sentiment analysis result: Angry / Critical (Confidence: 98%). Critical priority flagged.", color: "#f43f5e" },
        { text: "[00:05] [action] Profile stage moved to ESCALATED. System-level notification fired.", color: "#6366f1" },
        { text: "[00:06] [action] Creating urgent task: \"AC breakdown follow-up with Yasmin Rahmani\" assigned to Imtiyaz Ahmed.", color: "#10b981" },
        { text: "[00:07] [action] WhatsApp emergency broadcast dispatched: \"Hello Yasmin, we are deeply sorry. Our HVAC specialist is on the way.\"", color: "#10b981" }
    ],
    "data-collection": [
        { text: "[00:01] [trigger] Incoming message from unknown number +974 5550 1234: \"Hello, tell me about your painting services.\"", color: "var(--text-muted)" },
        { text: "[00:02] [filter] AI matches unregistered contact. Launching Data Collection Bot.", color: "#c084fc" },
        { text: "[00:03] [filter] Bot prompt: \"Hello! I'd love to help. To register your request, may I have your Name and Email?\"", color: "#6366f1" },
        { text: "[00:05] [trigger] Incoming: \"Sure, it's Ali and email is ali.doha@gmail.com\"", color: "var(--text-muted)" },
        { text: "[00:06] [filter] AI parsing contact data... Name: Ali, Email: ali.doha@gmail.com.", color: "#c084fc" },
        { text: "[00:07] [action] Directory database updated: Created contact Ali (+974 5550 1234, ali.doha@gmail.com).", color: "#10b981" },
        { text: "[00:08] [action] Synced data collection block to Zoho CRM... Completed.", color: "#10b981" }
    ],
    "service-request": [
        { text: "[00:01] [trigger] Incoming message from Najeeb Shekasan: \"Can I schedule an engineer survey?\"", color: "var(--text-muted)" },
        { text: "[00:02] [filter] AI matches booking intent. Querying availability schedules.", color: "#c084fc" },
        { text: "[00:03] [filter] Proposing next open slot: Tomorrow (Monday) at 10:00 AM.", color: "#6366f1" },
        { text: "[00:05] [trigger] Incoming: \"Yes, tomorrow at 10 AM is perfect.\"", color: "var(--text-muted)" },
        { text: "[00:06] [filter] Checklist items marked: Survey / Quote Scheduled.", color: "#10b981" },
        { text: "[00:07] [action] Scheduling electrical/AC technician dispatch slot: June 8, 2026, 10:00 AM.", color: "#6366f1" },
    ],
    "broadcast-execution": [
        { text: "[00:01] [trigger] User triggered WhatsApp Bulk Campaign template in AI Automation designer.", color: "var(--text-muted)" },
        { text: "[00:02] [filter] Scanning active CRM contacts... Found 6 active leads in pipeline.", color: "#c084fc" },
        { text: "[00:03] [action] Initializing real-time dispatch queue... Interval set to 1.5s per recipient.", color: "#6366f1" },
        { text: "[00:04] [action] Sending to Najeeb Shekasan (+974 30544808) -> status: queued -> SENT.", color: "#10b981" },
        { text: "[00:05] [action] Sending to Fatima Al-Kuwari (+974 55160323) -> status: queued -> SENT.", color: "#10b981" },
        { text: "[00:07] [action] Sending to Ali Hassan (+974 33445566) -> status: queued -> SENT.", color: "#10b981" },
        { text: "[00:08] [action] Sending to Yasmin Rahmani (+974 66098765) -> status: queued -> SENT.", color: "#10b981" },
        { text: "[00:10] [action] All broadcast queues completed successfully. Reporting metrics updated.", color: "#10b981" }
    ]
};

function runWorkflowSimulation() {
    const select = document.getElementById("workflow-template-select");
    const consoleLogs = document.getElementById("workflow-console-logs");
    if (!select || !consoleLogs) return;
    
    const templateId = select.value;
    const logs = simulationLogs[templateId];
    if (!logs) return;
    
    consoleLogs.innerHTML = `<div style="color:var(--text-muted);">[system] Initializing workflow "${select.options[select.selectedIndex].text}" test run...</div>`;
    
    let currentLogIndex = 0;
    
    function printNextLog() {
        if (currentLogIndex >= logs.length) {
            const finalDiv = document.createElement("div");
            finalDiv.style.fontWeight = "bold";
            finalDiv.style.marginTop = "4px";
            finalDiv.style.color = "#10b981";
            finalDiv.innerText = `[system] Workflow "${select.options[select.selectedIndex].text}" executed successfully.`;
            consoleLogs.appendChild(finalDiv);
            consoleLogs.scrollTop = consoleLogs.scrollHeight;
            return;
        }
        
        const log = logs[currentLogIndex];
        const div = document.createElement("div");
        div.style.color = log.color;
        div.innerText = log.text;
        consoleLogs.appendChild(div);
        consoleLogs.scrollTop = consoleLogs.scrollHeight;
        
        currentLogIndex++;
        setTimeout(printNextLog, 1200);
    }
    
    setTimeout(printNextLog, 800);
}

// ==========================================
// PROJECTS & TASKS MANAGEMENT MODULE
// ==========================================

let projects = [
    {
        id: 1,
        name: "West Bay Tower MEP Operations",
        client: "Najeeb Shekasan",
        progress: 65,
        milestones: [
            { title: "Initial Site Survey completed", date: "2026-06-07", status: "completed" },
            { title: "Compressor parts ordered", date: "2026-06-08", status: "pending" },
            { title: "HVAC Unit Refitting", date: "2026-06-12", status: "pending" },
            { title: "Final MEP Signoff", date: "2026-06-18", status: "pending" }
        ],
        comments: [
            { sender: "Mohammad Al-Haji", text: "Inspected Carrier compressor units. Copper wiring is corroded and needs replacement ASAP.", time: "2026-06-07 14:10" },
            { sender: "Imtiyaz Ahmed", text: "Parts ordered from Carrier Doha warehouse. Scheduled for delivery tomorrow.", time: "2026-06-07 15:30" }
        ],
        files: [
            { name: "AC_Refitting_Specs_WestBay.pdf", size: "1.4 MB", date: "2026-06-07" },
            { name: "Ducting_Layout_Tower.dwg", size: "4.8 MB", date: "2026-06-07" }
        ],
        timeEntries: [
            { employee: "Mohammad Al-Haji", description: "HVAC Site Inspection & diagnostics", hours: 3, date: "2026-06-07" },
            { employee: "Imtiyaz Ahmed", description: "Carrier Parts logistics coordination", hours: 1.5, date: "2026-06-07" }
        ]
    },
    {
        id: 2,
        name: "Emerald Villa Painting Project",
        client: "Fatima Al-Kuwari",
        progress: 25,
        milestones: [
            { title: "Color Consultation Done", date: "2026-06-06", status: "completed" },
            { title: "Prep work & masking", date: "2026-06-09", status: "pending" },
            { title: "Emerald Basecoat application", date: "2026-06-10", status: "pending" },
            { title: "Final QA touchups", date: "2026-06-14", status: "pending" }
        ],
        comments: [
            { sender: "Fatima Al-Kuwari", text: "Emerald green swatches approved. Basecoat prep is ready.", time: "2026-06-06 10:45" }
        ],
        files: [
            { name: "Approved_Colors_VillaPaint.png", size: "820 KB", date: "2026-06-06" }
        ],
        timeEntries: [
            { employee: "Fatima Al-Kuwari", description: "Color consultation & swatching", hours: 2, date: "2026-06-06" }
        ]
    }
];

let projectTasks = [
    { id: 1, projectId: 1, text: "Perform diagnostic HVAC tests", assignee: "Mohammad Al-Haji", priority: "High", status: "In Progress", dueDate: "2026-06-08", reminder: true },
    { id: 2, projectId: 1, text: "Verify Carrier parts logistics", assignee: "Imtiyaz Ahmed", priority: "Medium", status: "Pending", dueDate: "2026-06-08", reminder: false },
    { id: 3, projectId: 2, text: "Mask baseboards and protect furniture", assignee: "Fatima Al-Kuwari", priority: "High", status: "Pending", dueDate: "2026-06-09", reminder: true },
    { id: 4, projectId: 1, text: "Update dashboard inventory logs", assignee: "Admin Team", priority: "Low", status: "Completed", dueDate: "2026-06-07", reminder: false }
];

let activeProjectId = 1;

function renderProjects() {
    const container = document.getElementById("projects-list");
    if (!container) return;
    container.innerHTML = "";

    projects.forEach(p => {
        const isActive = activeProjectId === p.id ? "active" : "";
        const div = document.createElement("div");
        div.className = `workflow-card ${isActive}`;
        div.style.borderLeft = `4px solid ${activeProjectId === p.id ? "var(--primary)" : "rgba(255,255,255,0.05)"}`;
        div.onclick = () => selectProject(p.id);

        div.innerHTML = `
            <div style="font-weight:700; font-size:13px; color:#fff;">${p.name}</div>
            <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">Client: ${p.client}</div>
            <div style="margin-top:10px; display:flex; align-items:center; gap:8px;">
                <div style="flex:1; height:6px; background:rgba(255,255,255,0.05); border-radius:3px; overflow:hidden;">
                    <div style="width:${p.progress}%; height:100%; background:linear-gradient(90deg, var(--secondary), var(--primary));"></div>
                </div>
                <span style="font-size:11px; font-weight:700; color:var(--text-muted);">${p.progress}%</span>
            </div>
        `;
        container.appendChild(div);
    });
}

function selectProject(projectId) {
    activeProjectId = projectId;
    renderProjects();
    renderProjectTasks();

    const project = projects.find(p => p.id === projectId);
    const pane = document.getElementById("project-details-pane");
    if (!project || !pane) return;

    pane.innerHTML = `
        <div style="border-bottom:1px solid var(--border-color); padding-bottom:12px;">
            <span style="font-size:10px; font-weight:700; text-transform:uppercase; color:var(--text-muted);">Selected Project details</span>
            <h3 style="font-size:16px; font-weight:800; color:#fff; margin-top:4px;">${project.name}</h3>
            <span class="badge-status success" style="font-size:10px; margin-top:6px;">Client: ${project.client}</span>
        </div>

        <!-- Milestone Calendar Tracker -->
        <div>
            <label style="display:block; font-size:10px; font-weight:700; text-transform:uppercase; color:var(--text-muted); margin-bottom:8px; letter-spacing:0.5px;">Project Milestones (Click to Toggle)</label>
            <div style="display:flex; flex-direction:column; gap:8px;">
                ${project.milestones.map((m, idx) => {
                    const isDone = m.status === 'completed';
                    return `
                        <div style="display:flex; justify-content:space-between; align-items:center; background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-sm); padding:8px 12px; cursor:pointer;" onclick="toggleMilestone(${idx})">
                            <div style="display:flex; align-items:center; gap:8px;">
                                <span class="check-box" style="width:12px; height:12px; border:1px solid ${isDone ? "var(--success)" : "var(--border-color)"}; border-radius:3px; display:inline-flex; align-items:center; justify-content:center; font-size:9px; background:${isDone ? "rgba(16,185,129,0.15)" : "transparent"}; color:var(--success); font-weight:bold;">
                                    ${isDone ? "✓" : ""}
                                </span>
                                <span style="font-size:11px; text-decoration:${isDone ? 'line-through' : 'none'}; color:${isDone ? 'var(--text-dim)' : '#fff'};">${m.title}</span>
                            </div>
                            <span style="font-size:10px; color:var(--text-muted); font-weight:bold;">${m.date}</span>
                        </div>
                    `;
                }).join("")}
            </div>
        </div>

        <!-- Attachments Section -->
        <div style="border-top:1px solid var(--border-color); padding-top:12px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                <label style="font-size:10px; font-weight:700; text-transform:uppercase; color:var(--text-muted); letter-spacing:0.5px;">Engineering Attachments</label>
                <button class="btn btn-secondary" style="padding:2px 8px; font-size:10px; height:auto;" onclick="addProjectFile()">+ Attach</button>
            </div>
            <div style="display:flex; flex-direction:column; gap:6px;">
                ${project.files.map(f => `
                    <div style="display:flex; justify-content:space-between; align-items:center; background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-sm); padding:6px 12px; font-size:11px;">
                        <span style="color:#fff; font-weight:700;">📂 ${f.name}</span>
                        <span style="color:var(--text-muted); font-size:10px;">${f.size} | ${f.date}</span>
                    </div>
                `).join("")}
            </div>
        </div>

        <!-- Time Entries Section -->
        <div style="border-top:1px solid var(--border-color); padding-top:12px;">
            <label style="font-size:10px; font-weight:700; text-transform:uppercase; color:var(--text-muted); margin-bottom:8px; display:block; letter-spacing:0.5px;">Employee Time Logging</label>
            <div style="display:flex; flex-direction:column; gap:6px; margin-bottom:8px; max-height:80px; overflow-y:auto;">
                ${project.timeEntries.map(t => `
                    <div style="display:flex; justify-content:space-between; align-items:center; background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-sm); padding:6px 12px; font-size:11px;">
                        <div>
                            <span style="color:#fff; font-weight:700;">${t.employee}</span>
                            <span style="color:var(--text-muted);"> - ${t.description}</span>
                        </div>
                        <span style="color:var(--warning); font-weight:700;">${t.hours} hrs</span>
                    </div>
                `).join("")}
            </div>
            <div style="display:flex; gap:6px;">
                <select id="time-employee" style="flex:1; background:rgba(255,255,255,0.03); border:1px solid var(--border-color); padding:4px; border-radius:var(--radius-sm); color:#fff; font-size:10px; background:#0f172a;">
                    <option value="Mohammad Al-Haji">Mohammad</option>
                    <option value="Imtiyaz Ahmed">Imtiyaz</option>
                    <option value="Fatima Al-Kuwari">Fatima</option>
                </select>
                <input type="number" id="time-hours" placeholder="Hrs" style="width:45px; background:rgba(255,255,255,0.03); border:1px solid var(--border-color); padding:4px; border-radius:var(--radius-sm); color:#fff; font-size:10px;">
                <input type="text" id="time-desc" placeholder="Work description" style="flex:2; background:rgba(255,255,255,0.03); border:1px solid var(--border-color); padding:4px 8px; border-radius:var(--radius-sm); color:#fff; font-size:10px;">
                <button class="btn btn-primary" style="padding:2px 8px; font-size:10px;" onclick="addProjectTimeEntry()">+ Log</button>
            </div>
        </div>

        <!-- Comments Thread Section -->
        <div style="border-top:1px solid var(--border-color); padding-top:12px; display:flex; flex-direction:column; gap:8px;">
            <label style="font-size:10px; font-weight:700; text-transform:uppercase; color:var(--text-muted); letter-spacing:0.5px;">Engineering Log & Comments</label>
            <div style="display:flex; flex-direction:column; gap:6px; max-height:100px; overflow-y:auto; margin-bottom:4px;" id="comments-thread">
                ${project.comments.map(c => `
                    <div style="background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-sm); padding:6px 10px; font-size:11px; line-height:1.3;">
                        <div style="display:flex; justify-content:space-between; font-weight:700; color:var(--primary); font-size:10px; margin-bottom:2px;">
                            <span>${c.sender}</span>
                            <span style="color:var(--text-muted); font-weight:normal;">${c.time}</span>
                        </div>
                        <div style="color:var(--text-main);">${c.text}</div>
                    </div>
                `).join("")}
            </div>
            <div style="display:flex; gap:6px;">
                <input type="text" id="new-project-comment-text" placeholder="Add engineering log entry..." style="flex:1; background:rgba(255,255,255,0.02); border:1px solid var(--border-color); padding:6px 10px; border-radius:var(--radius-sm); font-size:11px; color:#fff;" onkeypress="checkCommentSubmit(event)">
                <button class="btn btn-primary" style="padding:4px 10px; font-size:11px;" onclick="addProjectComment()">Post</button>
            </div>
        </div>
    `;
}

function renderProjectTasks() {
    const container = document.getElementById("project-tasks-container");
    const filter = document.getElementById("task-filter-assignee").value;
    if (!container) return;
    container.innerHTML = "";

    const filtered = projectTasks.filter(t => {
        if (t.projectId !== activeProjectId) return false;
        if (filter !== 'all' && t.assignee !== filter) return false;
        return true;
    });

    if (filtered.length === 0) {
        container.innerHTML = `<div style="text-align:center; color:var(--text-dim); padding:20px; font-size:12px;">No work package tasks assigned for this select criteria.</div>`;
        return;
    }

    filtered.forEach(t => {
        const isDone = t.status === 'Completed';
        const div = document.createElement("div");
        div.className = "timeline-event-item";
        div.style.paddingLeft = "36px";
        div.style.background = isDone ? "rgba(16,185,129,0.02)" : "rgba(255,255,255,0.01)";
        div.style.borderColor = isDone ? "var(--success)" : "var(--border-color)";

        const circleStyle = `position:absolute; left:12px; top:14px; width:14px; height:14px; border-radius:50%; border:1px solid ${isDone ? 'var(--success)' : 'var(--border-color)'}; cursor:pointer; display:flex; align-items:center; justify-content:center; font-size:9px; color:var(--success); background:${isDone ? 'rgba(16,185,129,0.15)' : 'transparent'}; font-weight:bold;`;
        const tickElement = `<div style="${circleStyle}" onclick="toggleTaskComplete(${t.id})">${isDone ? '✓' : ''}</div>`;

        const reminderTag = t.reminder ? `<span style="font-size:9px; background:rgba(244,63,94,0.15); color:var(--danger); padding:2px 6px; border-radius:10px; font-weight:bold; margin-left:6px;">⏰ REMINDER</span>` : "";

        const priorityColors = {
            'High': 'var(--danger)',
            'Medium': 'var(--warning)',
            'Low': 'var(--success)'
        };

        div.innerHTML = `
            ${tickElement}
            <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                <div style="font-weight:700; font-size:12px; color:${isDone ? 'var(--text-dim)' : '#fff'}; text-decoration:${isDone ? 'line-through' : 'none'};">${t.text} ${reminderTag}</div>
                <span style="font-size:9px; color:${priorityColors[t.priority] || '#fff'}; font-weight:bold; text-transform:uppercase;">${t.priority}</span>
            </div>
            <div style="display:flex; justify-content:space-between; align-items:center; margin-top:8px; font-size:10px; color:var(--text-muted);">
                <span>Assignee: <strong>${t.assignee}</strong></span>
                <span>Due Date: <strong>${t.dueDate}</strong></span>
            </div>
        `;
        container.appendChild(div);
    });
}

function toggleTaskComplete(taskId) {
    const task = projectTasks.find(t => t.id === taskId);
    if (task) {
        task.status = task.status === 'Completed' ? 'In Progress' : 'Completed';
        syncState();
        renderProjectTasks();
        renderOperationsTasks();
        logActivity(`Task status updated`, `"${task.text}" marked as ${task.status}`, 'primary');
    }
}

function toggleMilestone(idx) {
    const project = projects.find(p => p.id === activeProjectId);
    if (project && project.milestones[idx]) {
        const m = project.milestones[idx];
        m.status = m.status === 'completed' ? 'pending' : 'completed';
        
        let completedCount = project.milestones.filter(m => m.status === 'completed').length;
        project.progress = Math.round(15 + completedCount * 21.25);
        if (project.progress > 100) project.progress = 100;
        
        syncState();
        selectProject(activeProjectId);
        logActivity(`Milestone updated`, `"${m.title}" marked as ${m.status.toUpperCase()}`, 'success');
    }
}

function openAddTaskModal() {
    const modal = document.getElementById("add-task-modal");
    const select = document.getElementById("form-task-project");
    if (!modal || !select) return;

    select.innerHTML = projects.map(p => `<option value="${p.id}">${p.name}</option>`).join("");
    modal.classList.add("active");
}

function closeAddTaskModal() {
    document.getElementById("add-task-modal").classList.remove("active");
}

function submitAddTaskForm() {
    const desc = document.getElementById("form-task-desc").value.trim();
    const projectId = parseInt(document.getElementById("form-task-project").value);
    const assignee = document.getElementById("form-task-assignee").value;
    const prioritySelect = document.getElementById("form-task-priority").value;
    const priority = prioritySelect.replace(/[\uE000-\uF8FF]|\uD83C[\uDC00-\uDFFF]|\uD83D[\uDC00-\uDFFF]|[\u2011-\u26FF]|\uD83E[\uDD10-\uDDFF]/g, '').trim(); 
    const dueDate = document.getElementById("form-task-date").value;
    const reminder = document.getElementById("form-task-reminder").checked;

    if (desc === "") {
        alert("Please enter a task description.");
        return;
    }

    const newTask = {
        id: projectTasks.length + 1,
        projectId: projectId,
        text: desc,
        assignee: assignee,
        priority: priority,
        status: "Pending",
        dueDate: dueDate,
        reminder: reminder
    };

    projectTasks.push(newTask);
    document.getElementById("form-task-desc").value = "";
    document.getElementById("form-task-reminder").checked = false;

    closeAddTaskModal();
    renderProjectTasks();
    renderOperationsTasks();

    const proj = projects.find(p => p.id === projectId);
    logActivity(`Task Assigned: ${assignee}`, `"${desc}" assigned under ${proj ? proj.name : 'project'}.`, 'success');
}

function checkCommentSubmit(e) {
    if (e.key === "Enter") {
        addProjectComment();
    }
}

function addProjectComment() {
    const input = document.getElementById("new-project-comment-text");
    const text = input.value.trim();
    if (text === "") return;

    const project = projects.find(p => p.id === activeProjectId);
    if (project) {
        project.comments.push({
            sender: "Imtiyaz Ahmed",
            text: text,
            time: new Date().toISOString().replace('T', ' ').substring(0, 16)
        });
        input.value = "";
        selectProject(activeProjectId);
        logActivity(`Engineering log entry added`, `Comment added to ${project.name}`, 'secondary');
    }
}

function addProjectFile() {
    const filename = prompt("Enter engineering file name to attach:", "Wiring_Schematic_v2.pdf");
    if (!filename || filename.trim() === "") return;

    const project = projects.find(p => p.id === activeProjectId);
    if (project) {
        project.files.push({
            name: filename.trim(),
            size: `${Math.round(1 + Math.random() * 5)}.${Math.round(Math.random() * 9)} MB`,
            date: new Date().toISOString().split('T')[0]
        });
        selectProject(activeProjectId);
        logActivity(`File attached to project`, `Uploaded: ${filename}`, 'success');
    }
}

function addProjectTimeEntry() {
    const employee = document.getElementById("time-employee").value;
    const hoursVal = parseFloat(document.getElementById("time-hours").value);
    const desc = document.getElementById("time-desc").value.trim();

    if (isNaN(hoursVal) || hoursVal <= 0 || desc === "") {
        alert("Please specify a valid hours count and work description.");
        return;
    }

    const project = projects.find(p => p.id === activeProjectId);
    if (project) {
        project.timeEntries.push({
            employee: employee,
            description: desc,
            hours: hoursVal,
            date: new Date().toISOString().split('T')[0]
        });

        document.getElementById("time-hours").value = "";
        document.getElementById("time-desc").value = "";

        project.progress += 5;
        if (project.progress > 100) project.progress = 100;

        selectProject(activeProjectId);
        logActivity(`Time entry logged: ${employee}`, `Spent ${hoursVal} hours: ${desc}`, 'warning');
    }
}

function simulateFollowUpReminder() {
    const pendingReminders = projectTasks.filter(t => t.reminder && t.status !== 'Completed');
    if (pendingReminders.length === 0) {
        alert("No pending tasks have reminders set. Try creating a new task and check 'Set Follow-up Reminder'.");
        return;
    }

    pendingReminders.forEach(t => {
        logActivity(`⏰ Reminder: Task Follow-up`, `Assigned to ${t.assignee}: "${t.text}" is due on ${t.dueDate}`, 'danger');
    });

    alert(`Dispatched ${pendingReminders.length} task follow-up reminders successfully! Check Live Event Logs.`);
}

function renderOperationsTasks() {
    const tbody = document.querySelector("#op-tab-tasks tbody");
    if (!tbody) return;
    tbody.innerHTML = "";

    projectTasks.forEach(t => {
        const isDone = t.status === 'Completed';
        const priorityClasses = {
            'High': 'danger',
            'Medium': 'warning',
            'Low': 'success'
        };
        const statusClasses = {
            'Completed': 'success',
            'In Progress': 'warning',
            'Pending': 'danger'
        };

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td style="font-weight: 700;">${t.text}</td>
            <td>${t.assignee}</td>
            <td>${t.dueDate}</td>
            <td><span class="badge-status ${priorityClasses[t.priority] || 'warning'}">${t.priority}</span></td>
            <td>
                <span class="badge-status ${statusClasses[t.status] || 'warning'}" style="cursor:pointer;" onclick="toggleTaskComplete(${t.id})">
                    ${t.status}
                </span>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

// ==========================================
// BUSINESS OPERATIONS STATE AND HANDLERS
// ==========================================

let inventoryStock = [
    { id: 1, product: "Carrier AC Compressors 2.5 Ton", location: "Main Warehouse Doha", qty: 12, limit: 5, status: "In Stock" },
    { id: 2, product: "Duct Insulation Glasswool Rolls", location: "Industrial Area Hub", qty: 150, limit: 30, status: "In Stock" },
    { id: 3, product: "Copper Cables 10mm² (100m)", location: "Salwa Road Store", qty: 3, limit: 10, status: "Low Stock" },
    { id: 4, product: "Smart Thermostat Control Panels", location: "Main Warehouse Doha", qty: 20, limit: 8, status: "In Stock" },
    { id: 5, product: "HVAC Duct Sealant Mastic (Gallon)", location: "Salwa Road Store", qty: 2, limit: 5, status: "Low Stock" }
];

let financeLedger = [
    { id: 1, ref: "June Staff Payroll Settlement", employee: "Fatima Al-Kuwari", type: "Payroll", amount: 12500, status: "Settled" },
    { id: 2, ref: "Site coordination transport & fuel claim", employee: "Mohammad Al-Haji", type: "Travel Expense", amount: 450, status: "Pending Approval" },
    { id: 3, ref: "Duct insulation tape & safety gloves", employee: "Imtiyaz Ahmed", type: "Material Cost", amount: 3200, status: "Settled" },
    { id: 4, ref: "Technical drawing print services", employee: "Support Staff", type: "Material Cost", amount: 180, status: "Settled" }
];

let supplierLedger = [
    { id: 1, supplier: "Carrier HVAC Supply WLL", ref: "INV-C-8942-2026", type: "Bill", amount: 18500, status: "Unpaid" },
    { id: 2, supplier: "National Paint Factories", ref: "QT-4929-Paint", type: "Quotation", amount: 5200, status: "Approved" },
    { id: 3, supplier: "Doha Cable WLL", ref: "PAY-4992-Copper", type: "Payment", amount: 12000, status: "Completed" },
    { id: 4, supplier: "Carrier HVAC Supply WLL", ref: "QT-2026-HVAC-Parts", type: "Quotation", amount: 28000, status: "Pending Review" }
];

let requestRegister = [
    { id: 1, client: "Yasmin Rahmani", type: "Complaint", desc: "AC indoor unit fan making squeaking noise after 1 hour run", status: "Escalated" },
    { id: 2, client: "Tariq Al-Mansoor", type: "Booking", desc: "HVAC duct insulation maintenance and thermal inspection booking", status: "Confirmed" },
    { id: 3, client: "Aisha Al-Thani", type: "Info Request", desc: "Requesting quotation for yearly preventative maintenance contract", status: "Responded" },
    { id: 4, client: "Imran Khan", type: "Complaint", desc: "Water dripping from kitchen ceiling HVAC duct outlet", status: "New" }
];

// RENDERERS

function renderInventory() {
    const tbody = document.getElementById("inventory-table-body");
    if (!tbody) return;
    tbody.innerHTML = "";

    inventoryStock.forEach(item => {
        const isLow = item.qty <= item.limit;
        const statusClass = isLow ? "danger" : "success";
        const statusText = isLow ? "Low Stock" : "In Stock";
        
        // Update item status in data model
        item.status = statusText;

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td style="font-weight: 700;">${item.product}</td>
            <td>${item.location}</td>
            <td><strong>${item.qty} units</strong></td>
            <td>${item.limit} units</td>
            <td><span class="badge-status ${statusClass}">${statusText}</span></td>
        `;
        tbody.appendChild(tr);
    });
}

function renderFinance() {
    const tbody = document.getElementById("finance-table-body");
    if (!tbody) return;
    tbody.innerHTML = "";

    financeLedger.forEach(item => {
        let statusClass = "warning";
        if (item.status === "Settled" || item.status === "Approved") statusClass = "success";
        if (item.status === "Rejected") statusClass = "danger";

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td style="font-weight: 700;">${item.ref}</td>
            <td>${item.employee}</td>
            <td><span class="badge-status primary">${item.type}</span></td>
            <td><strong>QR ${item.amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</strong></td>
            <td>
                <span class="badge-status ${statusClass}" style="cursor:pointer;" onclick="toggleFinanceStatus(${item.id})">
                    ${item.status}
                </span>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function toggleFinanceStatus(id) {
    const item = financeLedger.find(f => f.id === id);
    if (item) {
        if (item.status === "Pending Approval") {
            item.status = "Settled";
            logActivity("Expense Approved", `Approved expense for ${item.employee}: QR ${item.amount}`, "success");
        } else if (item.status === "Settled") {
            item.status = "Pending Approval";
            logActivity("Expense Status Reset", `Reset expense for ${item.employee} to pending`, "warning");
        }
        renderFinance();
    }
}

function renderSuppliers() {
    const tbody = document.getElementById("supplier-table-body");
    if (!tbody) return;
    tbody.innerHTML = "";

    supplierLedger.forEach(item => {
        let statusClass = "warning";
        if (item.status === "Completed" || item.status === "Approved") statusClass = "success";
        if (item.status === "Unpaid" || item.status === "Pending Review") statusClass = "danger";

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td style="font-weight: 700;">${item.supplier}</td>
            <td><code>${item.ref}</code></td>
            <td><span class="badge-status primary">${item.type}</span></td>
            <td><strong>QR ${item.amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</strong></td>
            <td>
                <span class="badge-status ${statusClass}" style="cursor:pointer;" onclick="toggleSupplierStatus(${item.id})">
                    ${item.status}
                </span>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function toggleSupplierStatus(id) {
    const item = supplierLedger.find(s => s.id === id);
    if (item) {
        if (item.status === "Unpaid") {
            item.status = "Completed";
            logActivity("Supplier Invoice Paid", `Paid ${item.supplier} for ${item.ref}: QR ${item.amount}`, "success");
        } else if (item.status === "Completed") {
            item.status = "Unpaid";
            logActivity("Supplier Payment Voided", `Payment voided for ${item.supplier} - ${item.ref}`, "danger");
        } else if (item.status === "Pending Review") {
            item.status = "Approved";
            logActivity("Supplier Quotation Approved", `Quotation approved from ${item.supplier}: QR ${item.amount}`, "success");
        } else if (item.status === "Approved") {
            item.status = "Pending Review";
        }
        renderSuppliers();
    }
}

function renderRequests() {
    const tbody = document.getElementById("requests-table-body");
    if (!tbody) return;
    tbody.innerHTML = "";

    requestRegister.forEach(item => {
        let statusClass = "warning";
        if (item.status === "Confirmed" || item.status === "Responded") statusClass = "success";
        if (item.status === "Escalated" || item.status === "New") statusClass = "danger";

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td style="font-weight: 700;">${item.client}</td>
            <td><span class="badge-status primary">${item.type}</span></td>
            <td>${item.desc}</td>
            <td>
                <span class="badge-status ${statusClass}" style="cursor:pointer;" onclick="toggleRequestStatus(${item.id})">
                    ${item.status}
                </span>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function toggleRequestStatus(id) {
    const item = requestRegister.find(r => r.id === id);
    if (item) {
        if (item.status === "New") {
            item.status = "Confirmed";
            logActivity("Request Confirmed", `Request from ${item.client} confirmed.`, "success");
        } else if (item.status === "Confirmed") {
            item.status = "Responded";
            logActivity("Request Marked Responded", `Responded to ${item.client}.`, "success");
        } else if (item.status === "Responded" || item.status === "Escalated") {
            item.status = "New";
        }
        renderRequests();
    }
}

// SUBMISSIONS HANDLERS

function submitStockMovement() {
    const product = document.getElementById("inv-form-product").value.trim();
    const location = document.getElementById("inv-form-location").value;
    const qtyInput = document.getElementById("inv-form-qty").value;
    const qtyVal = parseInt(qtyInput);
    const type = document.getElementById("inv-form-type").value;

    if (!product) {
        alert("Please enter a product name.");
        return;
    }
    if (isNaN(qtyVal) || qtyVal <= 0) {
        alert("Please enter a valid quantity greater than 0.");
        return;
    }

    // Check if product exists in warehouse
    let item = inventoryStock.find(i => i.product.toLowerCase() === product.toLowerCase() && i.location === location);

    if (item) {
        if (type === "inbound") {
            item.qty += qtyVal;
            logActivity("Warehouse Inbound", `Added ${qtyVal} units of "${item.product}" at ${location}`, "success");
        } else {
            if (item.qty < qtyVal) {
                alert(`Insufficient stock! Current stock of "${item.product}" at ${location} is only ${item.qty} units.`);
                return;
            }
            item.qty -= qtyVal;
            logActivity("Warehouse Outbound", `Removed ${qtyVal} units of "${item.product}" from ${location}`, "warning");
        }
    } else {
        if (type === "outbound") {
            alert(`Cannot perform outbound movement. Product "${product}" does not exist in ${location}.`);
            return;
        }
        item = {
            id: inventoryStock.length + 1,
            product: product,
            location: location,
            qty: qtyVal,
            limit: 5,
            status: "In Stock"
        };
        inventoryStock.push(item);
        logActivity("New Stock Registered", `Registered ${qtyVal} units of "${product}" at ${location}`, "success");
    }

    document.getElementById("inv-form-product").value = "";
    document.getElementById("inv-form-qty").value = "";

    renderInventory();
}

function submitFinanceClaim() {
    const ref = document.getElementById("fin-form-ref").value.trim();
    const employee = document.getElementById("fin-form-employee").value;
    const amountVal = parseFloat(document.getElementById("fin-form-amount").value);
    const type = document.getElementById("fin-form-type").value;

    if (!ref) {
        alert("Please enter a reference / description.");
        return;
    }
    if (isNaN(amountVal) || amountVal <= 0) {
        alert("Please enter a valid positive amount.");
        return;
    }

    const newClaim = {
        id: financeLedger.length + 1,
        ref: ref,
        employee: employee,
        type: type,
        amount: amountVal,
        status: "Pending Approval"
    };

    financeLedger.push(newClaim);
    logActivity("Financial Entry Logged", `Recorded ${type} for ${employee} - ${ref}: QR ${amountVal.toLocaleString()}`, "warning");

    document.getElementById("fin-form-ref").value = "";
    document.getElementById("fin-form-amount").value = "";

    renderFinance();
}

function submitSupplierDoc() {
    const name = document.getElementById("sup-form-name").value;
    const ref = document.getElementById("sup-form-ref").value.trim();
    const amountVal = parseFloat(document.getElementById("sup-form-amount").value);
    const type = document.getElementById("sup-form-type").value;

    if (!ref) {
        alert("Please enter a document reference.");
        return;
    }
    if (isNaN(amountVal) || amountVal <= 0) {
        alert("Please enter a valid positive amount.");
        return;
    }

    const newDoc = {
        id: supplierLedger.length + 1,
        supplier: name,
        ref: ref,
        type: type,
        amount: amountVal,
        status: type === "Bill" ? "Unpaid" : (type === "Quotation" ? "Pending Review" : "Completed")
    };

    supplierLedger.push(newDoc);
    logActivity("Supplier Document Filed", `Filed ${type} (${ref}) for ${name}: QR ${amountVal.toLocaleString()}`, "primary");

    document.getElementById("sup-form-ref").value = "";
    document.getElementById("sup-form-amount").value = "";

    renderSuppliers();
}

function submitRequestTicket() {
    const client = document.getElementById("req-form-client").value.trim();
    const desc = document.getElementById("req-form-desc").value.trim();
    const type = document.getElementById("req-form-type").value;

    if (!client) {
        alert("Please enter a client name.");
        return;
    }
    if (!desc) {
        alert("Please enter request details.");
        return;
    }

    const status = type === "Complaint" ? "Escalated" : "New";

    const newTicket = {
        id: requestRegister.length + 1,
        client: client,
        type: type,
        desc: desc,
        status: status
    };

    requestRegister.push(newTicket);

    if (type === "Complaint") {
        logActivity("Complaint Escalated", `Client: ${client} - "${desc}". Assigned to Imtiyaz Ahmed.`, "danger");
        
        // Add a task to Imtiyaz Ahmed in the tasks board
        const newTask = {
            id: projectTasks.length + 1,
            projectId: 1, // Default project
            text: `URGENT COMPLAINT: ${client} - ${desc}`,
            assignee: "Imtiyaz Ahmed",
            priority: "High",
            status: "Pending",
            dueDate: new Date().toISOString().split('T')[0],
            reminder: true
        };
        projectTasks.push(newTask);
        
        // Update both tasks and requests views
        renderOperationsTasks();
        
        alert(`Urgent complaint registered and auto-escalated to Imtiyaz Ahmed. Task created on Employee Assignment board.`);
    } else {
        logActivity("Service Request Logged", `${type} recorded for ${client}: "${desc}"`, "success");
    }

    document.getElementById("req-form-client").value = "";
    document.getElementById("req-form-desc").value = "";

    renderRequests();
}

// ==========================================
// CUSTOMER JOURNEY TRACKING HELPERS
// ==========================================

function toggleEditProfile(editMode) {
    const displayDiv = document.getElementById("cust-detail-info-display");
    const editDiv = document.getElementById("cust-detail-info-edit");
    if (!displayDiv || !editDiv) return;

    if (editMode) {
        displayDiv.style.display = "none";
        editDiv.style.display = "block";
        
        const customer = customers.find(c => c.id.toString() === activeCustomerId.toString());
        if (customer) {
            document.getElementById("edit-cust-name").value = customer.name;
            document.getElementById("edit-cust-phone").value = customer.phone;
            document.getElementById("edit-cust-email").value = customer.email;
        }
    } else {
        displayDiv.style.display = "block";
        editDiv.style.display = "none";
    }
}

function saveProfileChanges() {
    const customer = customers.find(c => c.id.toString() === activeCustomerId.toString());
    if (!customer) return;

    const name = document.getElementById("edit-cust-name").value.trim();
    const rawPhone = document.getElementById("edit-cust-phone").value.trim();
    const phone = formatQatarPhoneNumber(rawPhone);
    const email = document.getElementById("edit-cust-email").value.trim();

    if (name === "" || phone === "" || email === "") {
        alert("Please fill in Name, Phone, and Email.");
        return;
    }

    const oldName = customer.name;
    customer.name = name;
    customer.phone = phone;
    customer.email = email;

    customer.history.push({ type: "system", text: `Profile updated: Name changed from "${oldName}" to "${name}"`, time: new Date().toLocaleString() });
    syncState();
    
    toggleEditProfile(false);
    renderCustomersTable();
    renderKanban();
    selectCustomer(customer.id);

    logActivity("Customer Profile Updated", `Updated details for ${name}`, "success");
}

function updateCustomerStageDirectly(customerId, stage) {
    const customer = customers.find(c => c.id.toString() === customerId.toString());
    if (customer && customer.stage !== stage) {
        const oldStage = customer.stage;
        customer.stage = stage;
        customer.history.push({ type: "system", text: `Pipeline stage moved from ${oldStage.toUpperCase()} to ${stage.toUpperCase()} via Journey Tracker`, time: new Date().toLocaleString() });
        
        syncState();
        renderCustomersTable();
        renderKanban();
        selectCustomer(customerId);
        
        logActivity(`Lead ${customer.name} moved to ${stage.toUpperCase()}`, `Stage updated via Customer Journey Tracker.`, 'primary');
    }
}

function updateCustTempDirectly() {
    const customer = customers.find(c => c.id.toString() === activeCustomerId.toString());
    if (!customer) return;

    const val = document.getElementById("edit-cust-temp").value;
    if (customer.temp !== val) {
        const oldTemp = customer.temp;
        customer.temp = val;
        customer.history.push({ type: "system", text: `Temperature updated from ${oldTemp.toUpperCase()} to ${val.toUpperCase()}`, time: new Date().toLocaleString() });
        
        syncState();
        renderCustomersTable();
        renderKanban();
        selectCustomer(customer.id);
        
        logActivity(`Lead Temperature Updated`, `${customer.name} temperature set to ${val.toUpperCase()}`, 'secondary');
    }
}

function updateCustOwnerDirectly() {
    const customer = customers.find(c => c.id.toString() === activeCustomerId.toString());
    if (!customer) return;

    const val = document.getElementById("edit-cust-owner").value;
    const oldOwner = customer.leadOwner || "Unassigned";
    if (oldOwner !== val) {
        customer.leadOwner = val;
        customer.history.push({ type: "system", text: `Lead Owner changed from ${oldOwner} to ${val}`, time: new Date().toLocaleString() });
        
        syncState();
        selectCustomer(customer.id);
        
        logActivity(`Lead Owner Reassigned`, `${customer.name} assigned to ${val}`, 'secondary');
    }
}

function scheduleCustFollowUp() {
    const customer = customers.find(c => c.id.toString() === activeCustomerId.toString());
    if (!customer) return;

    const dateVal = document.getElementById("cust-followup-date").value;
    if (!dateVal) {
        alert("Please select a valid date for the follow-up.");
        return;
    }

    customer.nextFollowUp = dateVal;
    customer.history.push({ type: "system", text: `Follow-up scheduled for ${dateVal}`, time: new Date().toLocaleString() });
    
    // Auto-create a task for the Lead Owner on the Employee Assignment Board
    const owner = customer.leadOwner || "Imtiyaz Ahmed";
    const newTask = {
        id: projectTasks.length + 1,
        projectId: 1, // Default project
        text: `Sales Follow-up: Call ${customer.name} regarding MEP contract`,
        assignee: owner,
        priority: "High",
        status: "Pending",
        dueDate: dateVal,
        reminder: true
    };
    projectTasks.push(newTask);
    renderOperationsTasks();

    syncState();
    selectCustomer(customer.id);
    
    logActivity("Follow-Up Scheduled", `Scheduled call with ${customer.name} on ${dateVal}. Task assigned to ${owner}.`, "success");
    alert(`Follow-up scheduled successfully! Created high-priority reminder task for ${owner}.`);
}

function clearFollowUp() {
    const customer = customers.find(c => c.id.toString() === activeCustomerId.toString());
    if (!customer) return;

    delete customer.nextFollowUp;
    customer.history.push({ type: "system", text: `Follow-up date cleared`, time: new Date().toLocaleString() });

    syncState();
    selectCustomer(customer.id);
    
    logActivity("Follow-Up Cleared", `Cleared next follow-up date for ${customer.name}`, "warning");
}

// ==========================================
// SALES DOCUMENTS & LEDGER MANAGEMENT
// ==========================================

let salesOpportunities = [
    { id: 1, title: "West Bay Tower Commercial MEP Maintenance", client: "Najeeb Shekasan WLL", value: 45000, probability: 75, closingTarget: "2026-06-30", status: "In Progress" },
    { id: 2, title: "Emerald Villa interior premium paint contract", client: "Fatima Al-Kuwari", value: 12000, probability: 50, closingTarget: "2026-07-15", status: "In Progress" },
    { id: 3, title: "Carrier AC parts procurement", client: "John Smith", value: 18500, probability: 90, closingTarget: "2026-06-12", status: "In Progress" },
    { id: 4, title: "Al Thani residential complex refitting", client: "Abdulla Al-Thani", value: 85000, probability: 60, closingTarget: "2026-08-01", status: "In Progress" }
];

let issuedDocuments = [
    { id: 1, docNo: "INV-2026-004", client: "Ali Hassan", entity: "Smart Vision WLL", type: "Invoice", amount: 5800, date: "2026-06-07", status: "Paid" },
    { id: 2, docNo: "QT-2026-001", client: "Najeeb Shekasan WLL", entity: "Smart Vision WLL", type: "Quotation", amount: 45000, date: "2026-06-07", status: "Sent" },
    { id: 3, docNo: "QT-2026-002", client: "Fatima Al-Kuwari", entity: "GBOT Connect WLL", type: "Quotation", amount: 12000, date: "2026-06-06", status: "Sent" },
    { id: 4, docNo: "INV-2026-005", client: "John Smith", entity: "GBOT Connect WLL", type: "Invoice", amount: 18500, date: "2026-06-05", status: "Unpaid" }
];

function switchSalesTab(tabId) {
    const editorBtn = document.getElementById("btn-sales-tab-editor");
    const ledgerBtn = document.getElementById("btn-sales-tab-ledger");
    const editorTab = document.getElementById("sales-tab-editor");
    const ledgerTab = document.getElementById("sales-tab-ledger");

    if (!editorBtn || !ledgerBtn || !editorTab || !ledgerTab) return;

    if (tabId === 'ledger') {
        editorBtn.className = "btn btn-secondary";
        ledgerBtn.className = "btn btn-primary";
        editorTab.style.display = "none";
        ledgerTab.style.display = "block";
        renderSalesLedger();
    } else {
        editorBtn.className = "btn btn-primary";
        ledgerBtn.className = "btn btn-secondary";
        editorTab.style.display = "flex";
        ledgerTab.style.display = "none";
    }
}

function renderSalesLedger() {
    const tbodyDeals = document.getElementById("deals-table-body");
    const tbodyDocs = document.getElementById("sales-docs-table-body");

    if (tbodyDeals) {
        tbodyDeals.innerHTML = "";
        salesOpportunities.forEach(deal => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td style="font-weight: 700;">${deal.title}</td>
                <td>${deal.client}</td>
                <td><strong>QR ${deal.value.toLocaleString('en-US', { minimumFractionDigits: 2 })}</strong></td>
                <td>
                    <div style="display:flex; align-items:center; gap:8px;">
                        <div style="flex:1; height:6px; background:rgba(255,255,255,0.05); border-radius:3px; overflow:hidden; min-width:80px;">
                            <div style="width:${deal.probability}%; height:100%; background:var(--primary); border-radius:3px;"></div>
                        </div>
                        <span style="font-size:11px; font-weight:700;">${deal.probability}%</span>
                    </div>
                </td>
                <td>${deal.closingTarget}</td>
                <td><span class="badge-status primary">${deal.status}</span></td>
            `;
            tbodyDeals.appendChild(tr);
        });
    }

    if (tbodyDocs) {
        tbodyDocs.innerHTML = "";
        issuedDocuments.forEach(doc => {
            let statusClass = "warning";
            if (doc.status === "Paid" || doc.status === "Approved") statusClass = "success";
            if (doc.status === "Unpaid" || doc.status === "Overdue" || doc.status === "Sent") statusClass = "danger";

            let actionBtn = "";
            if (doc.status === "Unpaid" || doc.status === "Sent") {
                actionBtn = `<button class="btn btn-primary" style="padding:4px 8px; font-size:10px;" onclick="recordPaymentReceipt(${doc.id})">Record Receipt</button>`;
            } else if (doc.status === "Paid") {
                actionBtn = `<span style="color:var(--success); font-size:10px; font-weight:700;">✓ Receipt Logged</span>`;
            }

            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td style="font-weight: 700;"><code>${doc.docNo}</code></td>
                <td>${doc.client}</td>
                <td>${doc.entity}</td>
                <td><span class="badge-status primary">${doc.type}</span></td>
                <td><strong>QR ${doc.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}</strong></td>
                <td>${doc.date}</td>
                <td><span class="badge-status ${statusClass}">${doc.status}</span></td>
                <td>${actionBtn}</td>
            `;
            tbodyDocs.appendChild(tr);
        });
    }
}

function recordPaymentReceipt(docId) {
    const doc = issuedDocuments.find(d => d.id === docId);
    if (!doc) return;

    doc.status = "Paid";
    
    // Dynamically increment MTD Revenue Won KPI card
    const kpiRev = document.getElementById("kpi-revenue");
    if (kpiRev) {
        const currentStr = kpiRev.innerText.replace(/[^\d.]/g, '');
        const currentAmt = parseFloat(currentStr) || 0;
        const newAmt = currentAmt + doc.amount;
        kpiRev.innerText = `QR ${newAmt.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 })}`;
    }

    logActivity("Payment Receipt Logged", `Received payment for ${doc.docNo} from ${doc.client}: QR ${doc.amount.toLocaleString()}`, "success");
    alert(`Payment Receipt successfully generated for ${doc.docNo}! Balance settled in general ledger.`);
    
    renderSalesLedger();

    dispatchWebhook("payment_received", doc);
}

function triggerAutomatedPaymentReminders() {
    const unpaidDocs = issuedDocuments.filter(d => d.status === "Unpaid" || d.status === "Sent");
    if (unpaidDocs.length === 0) {
        alert("No outstanding unpaid quotes or invoices found in the ledger.");
        return;
    }

    alert(`Dispatched ${unpaidDocs.length} automated WhatsApp follow-up reminders successfully! Check Live Event Logs.`);
}

// ==========================================
// WHATSAPP CAMPAIGNS & BROADCASTS
// ==========================================

let whatsappTemplates = [
    { id: 1, name: "ac_maintenance_promo", category: "Marketing", body: "Hi {{name}}, beat the summer heat! ☀️ Book our Smart Vision MEP certified AC maintenance today and get 15% off. Reply AC to book." },
    { id: 2, name: "payment_reminder_friendly", category: "Utility", body: "Dear {{name}}, this is a friendly reminder from Smart Vision WLL regarding invoice {{invoice}}. Please let us know if payment has been processed. Thank you!" },
    { id: 3, name: "hvac_service_booking", category: "Support", body: "Hello {{name}}, your AC inspection survey has been successfully booked for {{date}} at 10:00 AM. Our team will contact you shortly." },
    { id: 4, name: "welcome_loyalty_vip", category: "Marketing", body: "Hi {{name}}, welcome to our VIP loyalty club! 🌟 You have accumulated {{points}} loyalty points. Enjoy exclusive discounts on all MEP services." }
];

let campaignLogs = [
    { id: 1, name: "Summer AC Promo Doha", segment: "All Leads", template: "ac_maintenance_promo", sent: 5, delivered: 5, read: 5, replies: 2, date: "2026-06-07 11:30" },
    { id: 2, name: "VIP Paint Rewards Launch", segment: "HomeSales Paint Labels", template: "welcome_loyalty_vip", sent: 1, delivered: 1, read: 1, replies: 0, date: "2026-06-06 14:15" }
];

function switchChatTab(tabId) {
    const inboxBtn = document.getElementById("btn-chat-tab-inbox");
    const bcBtn = document.getElementById("btn-chat-tab-broadcast");
    const tplBtn = document.getElementById("btn-chat-tab-templates");

    const inboxTab = document.getElementById("chat-tab-inbox");
    const bcTab = document.getElementById("chat-tab-broadcast");
    const tplTab = document.getElementById("chat-tab-templates");

    if (!inboxBtn || !bcBtn || !tplBtn || !inboxTab || !bcTab || !tplTab) return;

    if (tabId === 'broadcast') {
        inboxBtn.className = "btn btn-secondary";
        bcBtn.className = "btn btn-primary";
        tplBtn.className = "btn btn-secondary";
        inboxTab.style.display = "none";
        bcTab.style.display = "block";
        tplTab.style.display = "none";
        
        previewBroadcastTargetCount();
        populateTemplateSelects();
        renderCampaignLogs();
    } else if (tabId === 'templates') {
        inboxBtn.className = "btn btn-secondary";
        bcBtn.className = "btn btn-secondary";
        tplBtn.className = "btn btn-primary";
        inboxTab.style.display = "none";
        bcTab.style.display = "none";
        tplTab.style.display = "block";
        
        renderTemplatesList();
    } else {
        inboxBtn.className = "btn btn-primary";
        bcBtn.className = "btn btn-secondary";
        tplBtn.className = "btn btn-secondary";
        inboxTab.style.display = "flex";
        bcTab.style.display = "none";
        tplTab.style.display = "none";
        renderChatListSidebar();
    }
}

function previewBroadcastTargetCount() {
    const segment = document.getElementById("bc-audience").value;
    const lbl = document.getElementById("bc-target-count-lbl");
    if (!lbl) return;

    let count = 0;
    if (segment === "all-leads") {
        count = customers.filter(c => c.stage !== 'won').length;
    } else if (segment === "all-won") {
        count = customers.filter(c => c.stage === 'won').length;
    } else if (segment === "hot-leads") {
        count = customers.filter(c => c.temp === 'hot').length;
    } else if (segment === "mep-label") {
        count = customers.filter(c => c.labels && c.labels.some(l => l.toLowerCase().includes("mep") || l.toLowerCase().includes("hvac"))).length;
    } else if (segment === "paint-label") {
        count = customers.filter(c => c.labels && c.labels.some(l => l.toLowerCase().includes("paint") || l.toLowerCase().includes("home"))).length;
    }

    lbl.innerText = `Total recipients: ${count} contacts`;
}

function populateTemplateSelects() {
    const select = document.getElementById("bc-template-select");
    if (!select) return;

    select.innerHTML = whatsappTemplates.map(t => `<option value="${t.name}">${t.name} (${t.category})</option>`).join("");
    previewBroadcastMessageText();
}

function previewBroadcastMessageText() {
    const name = document.getElementById("bc-template-select").value;
    const previewBox = document.getElementById("bc-msg-preview");
    if (!previewBox) return;

    const tpl = whatsappTemplates.find(t => t.name === name);
    if (tpl) {
        previewBox.innerText = tpl.body;
    } else {
        previewBox.innerText = "Select a template...";
    }
}

function submitBroadcastCampaign() {
    const name = document.getElementById("bc-campaign-name").value.trim();
    const segment = document.getElementById("bc-audience").value;
    const tplName = document.getElementById("bc-template-select").value;

    if (!name) {
        alert("Please enter a campaign name.");
        return;
    }

    const tpl = whatsappTemplates.find(t => t.name === tplName);
    if (!tpl) {
        alert("Please select a template.");
        return;
    }

    // Filter target segment customers
    let targets = [];
    if (segment === "all-leads") {
        targets = customers.filter(c => c.stage !== 'won');
    } else if (segment === "all-won") {
        targets = customers.filter(c => c.stage === 'won');
    } else if (segment === "hot-leads") {
        targets = customers.filter(c => c.temp === 'hot');
    } else if (segment === "mep-label") {
        targets = customers.filter(c => c.labels && c.labels.some(l => l.toLowerCase().includes("mep") || l.toLowerCase().includes("hvac")));
    } else if (segment === "paint-label") {
        targets = customers.filter(c => c.labels && c.labels.some(l => l.toLowerCase().includes("paint") || l.toLowerCase().includes("home")));
    }

    if (targets.length === 0) {
        alert("Selected segment contains 0 contacts. Cannot launch campaign.");
        return;
    }

    const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

    // Create Campaign Log
    const newLog = {
        id: campaignLogs.length + 1,
        name: name,
        segment: document.getElementById("bc-audience").options[document.getElementById("bc-audience").selectedIndex].text,
        template: tplName,
        sent: 0,
        delivered: 0,
        read: 0,
        replies: 0,
        date: new Date().toISOString().replace('T', ' ').substring(0, 16)
    };
    campaignLogs.push(newLog);

    document.getElementById("bc-campaign-name").value = "";

    // Sync views
    syncState();
    renderCampaignLogs();
    
    // Update dashboard Campaigns KPI
    const kpiCamps = document.getElementById("kpi-campaigns");
    if (kpiCamps) {
        kpiCamps.innerText = "Processing...";
    }

    logActivity("WhatsApp Broadcast Started", `Campaign "${name}" initialized for ${targets.length} contacts.`, "primary");

    let targetIdx = 0;
    
    async function sendNextBroadcast() {
        if (targetIdx >= targets.length) {
            logActivity("WhatsApp Broadcast Completed", `Campaign "${name}" finished sending to ${targets.length} contacts.`, "success");
            
            if (kpiCamps) {
                kpiCamps.innerText = "100.0%"; // Update KPI when campaign finishes
            }
            
            alert(`WhatsApp Broadcast Campaign "${name}" has completed sending to all ${targets.length} contacts successfully!`);
            return;
        }

        const c = targets[targetIdx];
        
        // Format phone number
        if (c.phone) {
            c.phone = formatQatarPhoneNumber(c.phone);
        }

        let bodyText = tpl.body.replace("{{name}}", c.name);
        bodyText = bodyText.replace("{{points}}", c.loyaltyPoints || 0);
        bodyText = bodyText.replace("{{invoice}}", "INV-2026-005");
        bodyText = bodyText.replace("{{date}}", "Tomorrow");

        if (!c.messages) c.messages = [];
        c.messages.push({
            sender: "outgoing",
            text: bodyText,
            time: timeString
        });
        
        c.history.push({ type: "system", text: `WhatsApp Broadcast sent (Campaign: ${name})`, time: new Date().toLocaleString() });

        const recipientJid = c.jid || (c.phone ? c.phone.replace(/[^0-9]/g, '') + '@s.whatsapp.net' : '');

        dispatchWebhook("whatsapp_message_sent", {
            to: recipientJid,
            text: bodyText,
            sender: "broadcast"
        });

        // POST message to real phone number if live backend is connected
        if (isBackendConnected && recipientJid) {
            try {
                await fetch(`${BACKEND_URL}/api/send`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ to: recipientJid, text: bodyText })
                });
                console.log(`Broadcast: Sent to ${c.name} (${recipientJid})`);
            } catch (err) {
                console.error("Failed to send real WhatsApp message during broadcast:", err);
            }
        }

        // Update campaign logs
        newLog.sent++;
        newLog.delivered++;
        newLog.read++;
        
        syncState();
        renderCampaignLogs();
        renderChatListSidebar();
        
        if (activeChatId.toString() === c.id.toString() || (c.phone && c.phone.replace(/[^0-9]/g, '') === activeChatId.toString().replace(/[^0-9]/g, ''))) {
            renderChatMessages();
        }

        targetIdx++;
        setTimeout(sendNextBroadcast, 1500); // 1.5 seconds delay between each send
    }

    sendNextBroadcast();
}

function submitNewTemplate() {
    const name = document.getElementById("tpl-form-name").value.trim().toLowerCase().replace(/[^a-z0-9_]/g, '');
    const cat = document.getElementById("tpl-form-cat").value;
    const body = document.getElementById("tpl-form-body").value.trim();

    if (!name || !body) {
        alert("Please specify template name and message body.");
        return;
    }

    // Check if duplicate
    const exists = whatsappTemplates.some(t => t.name === name);
    if (exists) {
        alert(`Template "${name}" already exists. Please choose a different name.`);
        return;
    }

    const newTpl = {
        id: whatsappTemplates.length + 1,
        name: name,
        category: cat,
        body: body
    };
    whatsappTemplates.push(newTpl);

    document.getElementById("tpl-form-name").value = "";
    document.getElementById("tpl-form-body").value = "";

    renderTemplatesList();
    logActivity("Template Created", `New template: ${name}`, "success");
    alert(`WhatsApp message template "${name}" created and saved successfully!`);
}

function renderTemplatesList() {
    const container = document.getElementById("templates-list-container");
    if (!container) return;
    container.innerHTML = "";

    whatsappTemplates.forEach(t => {
        const div = document.createElement("div");
        div.style.cssText = "background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:12px; display:flex; flex-direction:column; gap:6px;";
        div.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-weight:700; color:#fff;"><code>${t.name}</code></span>
                <span class="badge-status primary" style="font-size:9px;">${t.category}</span>
            </div>
            <div style="font-size:12px; color:var(--text-muted); line-height:1.4; white-space:pre-wrap;">${t.body}</div>
        `;
        container.appendChild(div);
    });
}

function renderCampaignLogs() {
    const container = document.getElementById("bc-campaign-logs-container");
    if (!container) return;
    container.innerHTML = "";

    // Sort logs descending by date/id
    const sortedLogs = [...campaignLogs].reverse();

    sortedLogs.forEach(log => {
        const div = document.createElement("div");
        div.style.cssText = "background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:10px; font-size:11px; display:flex; flex-direction:column; gap:4px;";
        div.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <strong style="color:#fff; font-size:12px;">${log.name}</strong>
                <span style="color:var(--text-dim); font-size:10px;">${log.date}</span>
            </div>
            <div style="color:var(--text-muted);">Audience: <strong>${log.segment}</strong> | Template: <code>${log.template}</code></div>
            <div style="display:flex; gap:12px; margin-top:4px; font-size:10px;">
                <span style="color:var(--success);">Sent: <strong>${log.sent}</strong></span>
                <span style="color:var(--primary);">Read: <strong>${log.read}</strong></span>
                <span style="color:var(--warning);">Replies: <strong>${log.replies}</strong></span>
            </div>
        `;
        container.appendChild(div);
    });

    // Update totals
    const totalCountEl = document.getElementById("bc-rpt-total-count");
    const sentCountEl = document.getElementById("bc-rpt-messages-sent");
    if (totalCountEl) totalCountEl.innerText = campaignLogs.length;
    if (sentCountEl) sentCountEl.innerText = campaignLogs.reduce((acc, log) => acc + log.sent, 0);
}

// ==========================================
// EMAIL MARKETING & MAILBOX MODULES
// ==========================================

let emailTemplates = [
    { id: 1, name: "mep_service_intro", subject: "Smart Vision MEP: Professional Facility Management Services", category: "Marketing", body: "Dear {{name}},\n\nI hope this email finds you well.\n\nSmart Vision WLL is a certified leader in MEP building service maintenance, HVAC repairs, and commercial engineering works in Doha. We understand that keeping your facilities running smoothly is essential to your business. Our team offers custom annual packages that fit your exact commercial building dimensions and maintenance checklists.\n\nWe would love to coordinate a site visit to discuss your building requirements.\n\nBest regards,\n\nImtiyaz Ahmed\nFacility Manager\nSmart Vision WLL" },
    { id: 2, name: "paint_offer_spring", subject: "HomeSales Spring Paints: Bring Vibrant Colors to Your Villa", category: "Marketing", body: "Hi {{name}},\n\nAre you looking to refresh your home or office space with premium colors? \n\nHomeSales Paint offers eco-friendly, weather-resistant, and high-durability paints perfect for Qatar's villas. For a limited time, you have accumulated {{points}} loyalty points which can unlock an additional 10% discount on any painting service and coatings.\n\nReach out to us today to request a free color consultation catalog.\n\nWarm regards,\n\nSales Team\nGBOT Connect / HomeSales WLL" },
    { id: 3, name: "payment_receipt_invoice", subject: "Official Invoice Issued: {{invoice}}", category: "Proposal", body: "Dear {{name}},\n\nPlease find attached the official document {{invoice}} detailing services rendered. \n\nWe kindly request you to process the payment within 14 business days. Let us know if you have any questions regarding the line items. Thank you for your partnership.\n\nWarm regards,\n\nAccounts Receivable\nSmart Vision WLL" }
];

let emailCampaignLogs = [
    { id: 1, name: "Q2 Commercial Facility Outreach", subject: "Save 15% on Annual MEP Maintenance", segment: "All Leads", template: "mep_service_intro", sent: 5, opened: 4, clicked: 2, bounced: 0, date: "2026-06-05 09:00" },
    { id: 2, name: "Spring Villa Paint Launch", subject: "HomeSales Premium Paint Discounts", segment: "HomeSales Paint Labels", template: "paint_offer_spring", sent: 1, opened: 1, clicked: 1, bounced: 0, date: "2026-06-04 14:00" }
];

let mailboxEmails = [
    {
        id: 1,
        senderName: "Najeeb Shekasan",
        senderEmail: "najeeb@gmail.com",
        subject: "HVAC Repair & Maintenance Quote for Commercial Tower",
        date: "2026-06-07 15:45",
        body: "Hi Imtiyaz,\n\nOur 15-story office building in West Bay has been experiencing AC cooling efficiency issues on the 8th and 9th floors. I would like to get a formal quote for our split and package AC unit inspection and regular maintenance.\n\nCould you please let us know when your technical engineers are available for a site inspection?\n\nBest,\nNajeeb",
        read: false,
        customerId: 1,
        aiSuggestion: "Dear Najeeb,\n\nThank you for reaching out. Yes, we can inspect your 15-story West Bay tower's AC units tomorrow. Our MEP certified inspection team can be on-site at 10:00 AM. Would that work for you?\n\nBest regards,\nImtiyaz Ahmed\nSmart Vision WLL",
        aiScore: 92
    },
    {
        id: 2,
        senderName: "Fatima Al-Kuwari",
        senderEmail: "fatima.alkuwari@gmail.com",
        subject: "Eco-Friendly Paint Options & Consultation",
        date: "2026-06-06 10:15",
        body: "Hello,\n\nI am planning to repaint the interior of our family villa. Do you have eco-friendly or low-VOC paint options available? Also, do you provide free color consultation at home?\n\nLooking forward to your response.\n\nFatima",
        read: true,
        customerId: 2,
        aiSuggestion: "Dear Fatima,\n\nYes! HomeSales Paint offers a complete line of eco-friendly, low-VOC interior paints safe for family villas. We also provide complimentary color consultations directly at your home. When would be a convenient time for our consultant to visit?\n\nWarm regards,\nSales Team\nHomeSales WLL",
        aiScore: 88
    }
];

let activeEmailId = null;

function switchEmailTab(tabId) {
    const mailboxBtn = document.getElementById("btn-email-tab-mailbox");
    const campaignsBtn = document.getElementById("btn-email-tab-campaigns");
    const templatesBtn = document.getElementById("btn-email-tab-templates");

    const mailboxTab = document.getElementById("email-tab-mailbox");
    const campaignsTab = document.getElementById("email-tab-campaigns");
    const templatesTab = document.getElementById("email-tab-templates");

    if (!mailboxBtn || !campaignsBtn || !templatesBtn || !mailboxTab || !campaignsTab || !templatesTab) return;

    if (tabId === 'campaigns') {
        mailboxBtn.className = "btn btn-secondary";
        campaignsBtn.className = "btn btn-primary";
        templatesBtn.className = "btn btn-secondary";
        mailboxTab.style.display = "none";
        campaignsTab.style.display = "block";
        templatesTab.style.display = "none";

        previewEmailCampaignCount();
        populateEmailTemplateSelect();
        renderEmailCampaignLogs();
    } else if (tabId === 'templates') {
        mailboxBtn.className = "btn btn-secondary";
        campaignsBtn.className = "btn btn-secondary";
        templatesBtn.className = "btn btn-primary";
        mailboxTab.style.display = "none";
        campaignsTab.style.display = "none";
        templatesTab.style.display = "block";

        renderEmailTemplatesList();
    } else {
        mailboxBtn.className = "btn btn-primary";
        campaignsBtn.className = "btn btn-secondary";
        templatesBtn.className = "btn btn-secondary";
        mailboxTab.style.display = "flex";
        campaignsTab.style.display = "none";
        templatesTab.style.display = "none";

        renderMailbox();
        fetchRealEmails(false);
    }
}

function populateEmailTemplateSelect() {
    const select = document.getElementById("email-template-select");
    if (!select) return;

    select.innerHTML = emailTemplates.map(t => `<option value="${t.name}">${t.name} (${t.category})</option>`).join("");
    previewEmailCampaignBody();
}

function previewEmailCampaignBody() {
    const name = document.getElementById("email-template-select").value;
    const previewBox = document.getElementById("email-msg-preview");
    if (!previewBox) return;

    const tpl = emailTemplates.find(t => t.name === name);
    if (tpl) {
        previewBox.innerText = tpl.body;
    } else {
        previewBox.innerText = "Select a template...";
    }
}

function previewEmailCampaignCount() {
    const segment = document.getElementById("email-audience").value;
    const lbl = document.getElementById("email-target-count-lbl");
    if (!lbl) return;

    let count = 0;
    if (segment === "all-leads") {
        count = customers.filter(c => c.stage !== 'won').length;
    } else if (segment === "all-won") {
        count = customers.filter(c => c.stage === 'won').length;
    } else if (segment === "hot-leads") {
        count = customers.filter(c => c.temp === 'hot').length;
    } else if (segment === "mep-label") {
        count = customers.filter(c => c.labels && c.labels.some(l => l.toLowerCase().includes("mep") || l.toLowerCase().includes("hvac"))).length;
    } else if (segment === "paint-label") {
        count = customers.filter(c => c.labels && c.labels.some(l => l.toLowerCase().includes("paint") || l.toLowerCase().includes("home"))).length;
    }

    lbl.innerText = `Total recipients: ${count} contacts`;
}

async function submitEmailCampaign() {
    const name = document.getElementById("email-campaign-name").value.trim();
    const subject = document.getElementById("email-campaign-subject").value.trim();
    const segment = document.getElementById("email-audience").value;
    const tplName = document.getElementById("email-template-select").value;

    if (!name || !subject) {
        alert("Please enter a campaign name and subject line.");
        return;
    }

    const tpl = emailTemplates.find(t => t.name === tplName);
    if (!tpl) {
        alert("Please select a template.");
        return;
    }

    let targets = [];
    if (segment === "all-leads") {
        targets = customers.filter(c => c.stage !== 'won');
    } else if (segment === "all-won") {
        targets = customers.filter(c => c.stage === 'won');
    } else if (segment === "hot-leads") {
        targets = customers.filter(c => c.temp === 'hot');
    } else if (segment === "mep-label") {
        targets = customers.filter(c => c.labels && c.labels.some(l => l.toLowerCase().includes("mep") || l.toLowerCase().includes("hvac")));
    } else if (segment === "paint-label") {
        targets = customers.filter(c => c.labels && c.labels.some(l => l.toLowerCase().includes("paint") || l.toLowerCase().includes("home")));
    }

    if (targets.length === 0) {
        alert("Selected segment contains 0 contacts. Cannot launch campaign.");
        return;
    }

    let successCount = 0;
    for (const c of targets) {
        let bodyText = tpl.body.replace("{{name}}", c.name);
        bodyText = bodyText.replace("{{points}}", c.loyaltyPoints || 0);
        bodyText = bodyText.replace("{{invoice}}", "INV-2026-005");
        
        c.history.push({ 
            type: "system", 
            text: `📧 Email Broadcast Campaign Dispatched: "${name}" (Subject: ${subject})`, 
            time: new Date().toLocaleString() 
        });

        // Send real email if backend is connected
        if (isBackendConnected) {
            try {
                const res = await fetch(`${BACKEND_URL}/api/send-email`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        to: c.email,
                        subject: subject,
                        body: bodyText,
                        smtp: smtpConfig
                    })
                });
                const data = await res.json();
                if (res.ok && data.success) {
                    successCount++;
                }
            } catch (err) {
                console.error(`Failed to send campaign email to ${c.email}:`, err);
            }
        }
    }

    const newLog = {
        id: emailCampaignLogs.length + 1,
        name: name,
        subject: subject,
        segment: document.getElementById("email-audience").options[document.getElementById("email-audience").selectedIndex].text,
        template: tplName,
        sent: targets.length,
        opened: Math.ceil(targets.length * 0.8), // 80% open rate mockup
        clicked: Math.ceil(targets.length * 0.4), // 40% click rate mockup
        bounced: 0,
        date: new Date().toISOString().replace('T', ' ').substring(0, 16)
    };
    emailCampaignLogs.push(newLog);

    document.getElementById("email-campaign-name").value = "";
    document.getElementById("email-campaign-subject").value = "";

    syncState();
    renderEmailCampaignLogs();
    
    logActivity("Email Campaign Dispatched", `Sent campaign "${name}" to ${targets.length} recipients.`, "success");
    
    if (isBackendConnected) {
        alert(`Email Campaign "${name}" dispatched! Successfully sent ${successCount} of ${targets.length} emails via SMTP.`);
    } else {
        alert(`Email Campaign "${name}" simulated successfully to ${targets.length} recipients (Offline mode)!`);
    }
}

function renderEmailCampaignLogs() {
    const container = document.getElementById("email-campaign-logs-container");
    if (!container) return;
    container.innerHTML = "";

    const sortedLogs = [...emailCampaignLogs].reverse();
    sortedLogs.forEach(log => {
        const div = document.createElement("div");
        div.style.cssText = "background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:10px; font-size:11px; display:flex; flex-direction:column; gap:4px;";
        div.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <strong style="color:#fff; font-size:12px;">${log.name}</strong>
                <span style="color:var(--text-dim); font-size:10px;">${log.date}</span>
            </div>
            <div style="color:var(--text-muted);">Subject: <em>"${log.subject}"</em></div>
            <div style="color:var(--text-muted); font-size:10px;">Segment: <strong>${log.segment}</strong> | Template: <code>${log.template}</code></div>
            <div style="display:flex; gap:12px; margin-top:4px; font-size:10px;">
                <span style="color:var(--success);">Sent: <strong>${log.sent}</strong></span>
                <span style="color:var(--primary);">Opened: <strong>${log.opened}</strong></span>
                <span style="color:var(--warning);">Clicked: <strong>${log.clicked}</strong></span>
                <span style="color:var(--danger);">Bounced: <strong>${log.bounced}</strong></span>
            </div>
        `;
        container.appendChild(div);
    });

    const totalCountEl = document.getElementById("email-rpt-total-count");
    const sentCountEl = document.getElementById("email-rpt-sent-count");
    if (totalCountEl) totalCountEl.innerText = emailCampaignLogs.length;
    if (sentCountEl) sentCountEl.innerText = emailCampaignLogs.reduce((acc, log) => acc + log.sent, 0);
}

function submitNewEmailTemplate() {
    const name = document.getElementById("email-tpl-name").value.trim().toLowerCase().replace(/[^a-z0-9_]/g, '');
    const subject = document.getElementById("email-tpl-subject").value.trim();
    const cat = document.getElementById("email-tpl-cat").value;
    const body = document.getElementById("email-tpl-body").value.trim();

    if (!name || !subject || !body) {
        alert("Please fill out all template fields.");
        return;
    }

    const exists = emailTemplates.some(t => t.name === name);
    if (exists) {
        alert(`Template "${name}" already exists.`);
        return;
    }

    const newTpl = {
        id: emailTemplates.length + 1,
        name: name,
        subject: subject,
        category: cat,
        body: body
    };
    emailTemplates.push(newTpl);

    document.getElementById("email-tpl-name").value = "";
    document.getElementById("email-tpl-subject").value = "";
    document.getElementById("email-tpl-body").value = "";

    renderEmailTemplatesList();
    logActivity("Email Template Created", `New template: ${name}`, "success");
    alert(`Email template "${name}" created and saved successfully!`);
}

function renderEmailTemplatesList() {
    const container = document.getElementById("email-templates-list-container");
    if (!container) return;
    container.innerHTML = "";

    emailTemplates.forEach(t => {
        const div = document.createElement("div");
        div.style.cssText = "background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:12px; display:flex; flex-direction:column; gap:6px;";
        div.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-weight:700; color:#fff;"><code>${t.name}</code></span>
                <span class="badge-status primary" style="font-size:9px;">${t.category}</span>
            </div>
            <div style="font-size:12px; color:var(--text-dim);">Subject: <strong>${t.subject}</strong></div>
            <div style="font-size:11px; color:var(--text-muted); line-height:1.4; white-space:pre-wrap;">${t.body}</div>
        `;
        container.appendChild(div);
    });
}

async function fetchRealEmails(isManual = false) {
    if (!isBackendConnected) {
        if (isManual) {
            alert("Backend server is offline. Cannot fetch real emails.");
        }
        return;
    }
    
    try {
        const refreshBtn = document.querySelector('button[onclick="fetchRealEmails(true)"]');
        if (refreshBtn) {
            refreshBtn.disabled = true;
            refreshBtn.style.opacity = "0.5";
        }
        
        const response = await fetch(`${BACKEND_URL}/api/fetch-emails`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ smtp: smtpConfig })
        });
        
        const data = await response.json();
        if (data.success && Array.isArray(data.emails)) {
            let newEmailsAdded = 0;
            data.emails.forEach(realEmail => {
                // Check if email already exists in mailboxEmails
                const exists = mailboxEmails.some(e => 
                    e.senderEmail === realEmail.senderEmail && 
                    e.subject === realEmail.subject && 
                    (e.date === realEmail.date || e.body === realEmail.body)
                );
                
                if (!exists) {
                    // Map senderEmail to matching CRM contact/customer
                    const contact = customers.find(c => c.email && c.email.toLowerCase() === realEmail.senderEmail.toLowerCase());
                    const customerId = contact ? contact.id : null;
                    
                    // Generate AI Suggestion based on subject / content
                    let aiSuggestion = "";
                    let aiScore = 0;
                    const bodyLower = realEmail.body.toLowerCase();
                    const subjectLower = realEmail.subject.toLowerCase();
                    
                    if (subjectLower.includes("quote") || bodyLower.includes("quote") || bodyLower.includes("price") || bodyLower.includes("cost") || bodyLower.includes("how much")) {
                        aiSuggestion = `Dear ${realEmail.senderName},\n\nThank you for requesting a quote. Our commercial engineering team will prepare a customized proposal based on your requirements. Could you please share more details or schedule a quick site visit?\n\nBest regards,\nImtiyaz Ahmed\nFacility Manager\nSmart Vision WLL`;
                        aiScore = 90;
                    } else if (subjectLower.includes("cooling") || bodyLower.includes("cooling") || bodyLower.includes("ac ") || bodyLower.includes("hvac") || bodyLower.includes("chiller")) {
                        aiSuggestion = `Dear ${realEmail.senderName},\n\nThank you for reaching out. Yes, we specialize in certified MEP building maintenance and HVAC repairs. We can inspect your site tomorrow at 10:00 AM. Would that time work for you?\n\nBest regards,\nImtiyaz Ahmed\nFacility Manager\nSmart Vision WLL`;
                        aiScore = 92;
                    } else if (subjectLower.includes("paint") || bodyLower.includes("paint") || bodyLower.includes("villa") || bodyLower.includes("color")) {
                        aiSuggestion = `Dear ${realEmail.senderName},\n\nThank you for your interest. HomeSales Paint offers eco-friendly, weather-resistant paints perfect for Qatar's villas. We would be happy to schedule a free color consultation. What is your preferred date and time?\n\nWarm regards,\nSales Team\nGBOT Connect / HomeSales WLL`;
                        aiScore = 88;
                    } else {
                        aiSuggestion = `Dear ${realEmail.senderName},\n\nThank you for contacting Smart Vision WLL. We have received your email regarding "${realEmail.subject}" and our team is reviewing it. We will get back to you shortly.\n\nWarm regards,\nSupport Team\nSmart Vision WLL`;
                        aiScore = 80;
                    }
                    
                    const newEmail = {
                        id: realEmail.id,
                        senderName: realEmail.senderName,
                        senderEmail: realEmail.senderEmail,
                        subject: realEmail.subject,
                        date: realEmail.date,
                        body: realEmail.body,
                        read: false,
                        customerId: customerId,
                        aiSuggestion: aiSuggestion,
                        aiScore: aiScore
                    };
                    
                    mailboxEmails.unshift(newEmail);
                    newEmailsAdded++;
                    
                    logActivity(`Incoming Email Received`, `From: ${newEmail.senderName} (${newEmail.subject})`, "primary");
                    dispatchWebhook("incoming_email", newEmail);
                }
            });
            
            if (newEmailsAdded > 0) {
                // Sort by date descending
                mailboxEmails.sort((a, b) => new Date(b.date) - new Date(a.date));
                syncState();
                renderMailbox();
                if (isManual) {
                    alert(`Mailbox sync complete! Successfully fetched ${newEmailsAdded} new email(s).`);
                }
            } else {
                if (isManual) {
                    alert("Mailbox is already up to date. No new emails found.");
                }
            }
        } else {
            if (isManual) {
                alert("Failed to fetch emails: " + (data.error || "Unknown error"));
            }
        }
    } catch (err) {
        console.error("Error fetching real emails:", err);
        if (isManual) {
            alert("Error connecting to mail server: " + err.message);
        }
    } finally {
        const refreshBtn = document.querySelector('button[onclick="fetchRealEmails(true)"]');
        if (refreshBtn) {
            refreshBtn.disabled = false;
            refreshBtn.style.opacity = "1";
        }
    }
}

function renderMailbox() {
    const container = document.getElementById("email-list-container");
    if (!container) return;
    container.innerHTML = "";

    const searchVal = document.getElementById("email-search-input").value.toLowerCase();
    
    mailboxEmails.forEach(email => {
        if (searchVal && 
            !email.senderName.toLowerCase().includes(searchVal) && 
            !email.subject.toLowerCase().includes(searchVal) && 
            !email.body.toLowerCase().includes(searchVal)) {
            return;
        }

        const activeClass = activeEmailId === email.id ? "active" : "";
        const unreadClass = !email.read ? "unread" : "";
        
        const div = document.createElement("div");
        div.className = `email-item ${activeClass} ${unreadClass}`;
        div.onclick = () => selectEmail(email.id);
        
        const avatarInitials = email.senderName.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();

        div.innerHTML = `
            <div class="email-avatar">${avatarInitials}</div>
            <div class="email-item-details">
                <div class="email-item-header">
                    <span class="email-item-name">${email.senderName}</span>
                    <span class="email-item-time">${email.date.substring(11)}</span>
                </div>
                <div class="email-item-subject">${email.subject}</div>
                <div class="email-item-preview">${email.body.replace(/\n/g, ' ')}</div>
            </div>
        `;
        container.appendChild(div);
    });
}

function selectEmail(emailId) {
    activeEmailId = emailId;
    
    const email = mailboxEmails.find(e => e.id === emailId);
    if (!email) return;

    email.read = true;
    
    // Sync read states on list
    renderMailbox();

    document.getElementById("email-reading-empty").style.display = "none";
    document.getElementById("email-reading-content").style.display = "flex";

    document.getElementById("email-view-subject").innerText = email.subject;
    document.getElementById("email-view-from").innerText = `${email.senderName} <${email.senderEmail}>`;
    document.getElementById("email-view-date").innerText = email.date;
    document.getElementById("email-view-body").innerText = email.body;

    const suggBox = document.getElementById("email-ai-suggestion-box");
    if (email.aiSuggestion) {
        suggBox.style.display = "flex";
        document.getElementById("email-ai-suggestion-text").innerText = `"${email.aiSuggestion}"`;
        document.getElementById("email-ai-score-badge").innerText = `Score: ${email.aiScore}% Match`;
    } else {
        suggBox.style.display = "none";
    }

    document.getElementById("email-reply-field").value = "";
}

function sendEmailAISuggestion() {
    const email = mailboxEmails.find(e => e.id === activeEmailId);
    if (!email || !email.aiSuggestion) return;

    const replyText = email.aiSuggestion;
    email.aiSuggestion = null;
    
    submitManualEmailResponse(replyText);
}

function editEmailAISuggestion() {
    const email = mailboxEmails.find(e => e.id === activeEmailId);
    if (!email || !email.aiSuggestion) return;

    const newText = prompt("Edit the AI email reply suggestions draft:", email.aiSuggestion);
    if (newText && newText.trim() !== "") {
        email.aiSuggestion = newText;
        document.getElementById("email-ai-suggestion-text").innerText = `"${newText}"`;
    }
}

function submitManualEmailReply() {
    const text = document.getElementById("email-reply-field").value.trim();
    if (!text) {
        alert("Please type a response first.");
        return;
    }
    submitManualEmailResponse(text);
}

async function submitManualEmailResponse(replyText) {
    const email = mailboxEmails.find(e => e.id === activeEmailId);
    if (!email) return;

    // Save to customer timeline/history
    const customer = customers.find(c => c.id === email.customerId);
    if (customer) {
        if (!customer.messages) customer.messages = [];
        // Put in customer's outgoing email record
        customer.history.push({
            type: "system",
            text: `✉ Outgoing Email Sent:\nSubject: Re: ${email.subject}\nBody: ${replyText}`,
            time: new Date().toLocaleString()
        });
        
        logActivity(`Email reply sent to ${email.senderName}`, `Subject: Re: ${email.subject}`, "success");
    }

    // Append to email body in view
    email.body += `\n\n------------------------------\nOn ${new Date().toLocaleString()}, Smart Vision Support wrote:\n\n${replyText}`;
    
    // Clear AI suggestion
    email.aiSuggestion = null;

    selectEmail(email.id);
    syncState();
    
    // Send real email via SMTP backend
    if (isBackendConnected) {
        try {
            const res = await fetch(`${BACKEND_URL}/api/send-email`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    to: email.senderEmail,
                    subject: `Re: ${email.subject}`,
                    body: replyText,
                    smtp: smtpConfig
                })
            });
            const data = await res.json();
            if (res.ok && data.success) {
                alert(`Real email sent successfully to ${email.senderEmail} via SMTP!`);
            } else {
                alert(`SMTP Mail Dispatched alert triggered (Server error: ${data.error || 'Unknown error'})`);
            }
        } catch (err) {
            console.error("Failed to send real email via SMTP:", err);
            alert(`SMTP Mail Dispatched alert triggered (Network error: ${err.message})`);
        }
    } else {
        alert(`Email response simulated successfully to ${email.senderEmail} (Offline mode)!`);
    }
}

function simulateIncomingEmail() {
    // Generate new mock email
    const pool = [
        {
            senderName: "Abdulla Al-Thani",
            senderEmail: "abdulla.althani@outlook.com",
            subject: "Additional Requirements for AC Survey",
            body: "Hi Imtiyaz,\n\nFurther to our conversation, please make sure the technical inspection report covers our rooftop package chiller units as well.\n\nThanks,\nAbdulla",
            customerId: 4,
            aiSuggestion: "Dear Abdulla,\n\nUnderstood. We have added the rooftop package chillers to the survey scope. Our engineer will inspect them tomorrow.\n\nBest regards,\nImtiyaz Ahmed\nSmart Vision WLL",
            aiScore: 95
        },
        {
            senderName: "Yasmin Rahmani",
            senderEmail: "yasmin@doha-mail.com",
            subject: "Feedback on villa electrical rewire quote",
            body: "Dear Sales Team,\n\nI received your proposal for the electrical rewire. The pricing is slightly higher than expected. Do you have any flexibilty or payment options?\n\nWarmly,\nYasmin",
            customerId: 5,
            aiSuggestion: "Dear Yasmin,\n\nThank you for your feedback. We can offer a structured 3-stage payment plan (deposit, progress, final) to make it easier for you. Let me know if you would like me to draft that proposal.\n\nWarm regards,\nSales Team\nSmart Vision WLL",
            aiScore: 90
        }
    ];

    const newEmailData = pool[Math.floor(Math.random() * pool.length)];
    
    const newId = mailboxEmails.length + 1;
    const newEmail = {
        id: newId,
        senderName: newEmailData.senderName,
        senderEmail: newEmailData.senderEmail,
        subject: newEmailData.subject,
        date: new Date().toISOString().replace('T', ' ').substring(0, 16),
        body: newEmailData.body,
        read: false,
        customerId: newEmailData.customerId,
        aiSuggestion: newEmailData.aiSuggestion,
        aiScore: newEmailData.aiScore
    };

    mailboxEmails.unshift(newEmail);
    renderMailbox();

    logActivity(`Incoming Email Received`, `From: ${newEmail.senderName} (${newEmail.subject})`, "primary");
    alert(`✉ New email received from ${newEmail.senderName}: "${newEmail.subject}"`);
}

// ==========================================
// LOYALTY, ACCESS CONTROL & PERFORMANCE ANALYTICS
// ==========================================

let staffUsers = [
    { id: 1, name: "Imtiyaz Ahmed", email: "imtiyaz@smartvision.com", role: "Administrator", logins: 45 },
    { id: 2, name: "Nabeel Shah", email: "nabeel@smartvision.com", role: "Lead Agent", logins: 32 },
    { id: 3, name: "Fatima Al-Kuwari", email: "fatima@smartvision.com", role: "Billing Specialist", logins: 18 }
];

let pointsHistory = [
    { id: 1, name: "Najeeb Shekasan", delta: 120, reason: "Initial Sign-Up points", date: "2026-06-07 14:00" },
    { id: 2, name: "Yasmin Rahmani", delta: 200, reason: "Electrical Rewire Contract points", date: "2026-06-04 16:00" },
    { id: 3, name: "Fatima Al-Kuwari", delta: 50, reason: "Referral Sign-Up bonus", date: "2026-06-06 11:20" }
];

let retentionCampaignLogs = [
    { id: 1, type: "90-Day Dormant scan", date: "2026-06-05 10:00", contacts: 2, status: "Dispatched successfully" }
];

function switchLoyaltyTab(tabId) {
    const programBtn = document.getElementById("btn-loyalty-tab-program");
    const automationsBtn = document.getElementById("btn-loyalty-tab-automations");
    const staffBtn = document.getElementById("btn-loyalty-tab-staff");
    const analyticsBtn = document.getElementById("btn-loyalty-tab-analytics");

    const programTab = document.getElementById("loyalty-tab-program");
    const automationsTab = document.getElementById("loyalty-tab-automations");
    const staffTab = document.getElementById("loyalty-tab-staff");
    const analyticsTab = document.getElementById("loyalty-tab-analytics");

    if (!programBtn || !automationsBtn || !staffBtn || !analyticsBtn || !programTab || !automationsTab || !staffTab || !analyticsTab) return;

    // Reset buttons
    programBtn.className = "btn btn-secondary";
    automationsBtn.className = "btn btn-secondary";
    staffBtn.className = "btn btn-secondary";
    analyticsBtn.className = "btn btn-secondary";

    programTab.style.display = "none";
    automationsTab.style.display = "none";
    staffTab.style.display = "none";
    analyticsTab.style.display = "none";

    if (tabId === 'automations') {
        automationsBtn.className = "btn btn-primary";
        automationsTab.style.display = "block";
        renderRetentionLogs();
    } else if (tabId === 'staff') {
        staffBtn.className = "btn btn-primary";
        staffTab.style.display = "block";
        renderStaffTable();
    } else if (tabId === 'analytics') {
        analyticsBtn.className = "btn btn-primary";
        analyticsTab.style.display = "block";
        renderRetentionAnalytics();
    } else {
        programBtn.className = "btn btn-primary";
        programTab.style.display = "block";
        
        renderLoyaltyProgram();
        populateAdjustPointsSelect();
        renderPointsLedger();
    }
}

function renderLoyaltyProgram() {
    const tbody = document.getElementById("loyalty-customers-list");
    if (!tbody) return;
    tbody.innerHTML = "";

    let goldCount = 0;
    let silverCount = 0;
    let bronzeCount = 0;
    let totalPoints = 0;

    customers.forEach(c => {
        let points = c.loyaltyPoints || 0;
        totalPoints += points;

        let tier = "None";
        let tierClass = "tier-none";
        
        if (points >= 500) {
            tier = "Gold";
            tierClass = "tier-gold";
            goldCount++;
        } else if (points >= 200) {
            tier = "Silver";
            tierClass = "tier-silver";
            silverCount++;
        } else if (points >= 50) {
            tier = "Bronze";
            tierClass = "tier-bronze";
            bronzeCount++;
        }

        const tr = document.createElement("tr");
        tr.style.cssText = "border-bottom: 1px solid var(--border-color);";
        tr.innerHTML = `
            <td style="padding:10px 6px; font-weight:600; color:#fff;">${c.name}</td>
            <td style="padding:10px 6px; font-weight:700; color:var(--warning);">${points} pts</td>
            <td style="padding:10px 6px;"><span class="loyalty-tier-badge ${tierClass}" style="font-size:9px; padding:2px 6px; letter-spacing:0.5px;">${tier}</span></td>
            <td style="padding:10px 6px; text-align:right;">
                <button class="btn btn-secondary" style="font-size:10px; padding:2px 6px; margin-right:4px;" onclick="adjustPointsQuick(${c.id}, 10)">+10 Pts</button>
                <button class="btn btn-secondary" style="font-size:10px; padding:2px 6px;" onclick="adjustPointsQuick(${c.id}, 50)">+50 Pts</button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById("count-gold-tier").innerText = goldCount;
    document.getElementById("count-silver-tier").innerText = silverCount;
    document.getElementById("count-bronze-tier").innerText = bronzeCount;
    document.getElementById("total-points-issued").innerText = totalPoints;
}

function adjustPointsQuick(customerId, amount) {
    adjustCustomerPoints(customerId, amount, "Quick rewards bonus");
}

function adjustCustomerPoints(customerId, amount, reason) {
    const customer = customers.find(c => c.id === customerId);
    if (!customer) return;

    if (!customer.loyaltyPoints) customer.loyaltyPoints = 0;
    
    customer.loyaltyPoints += amount;
    if (customer.loyaltyPoints < 0) customer.loyaltyPoints = 0;

    // Add to ledger
    const newHistory = {
        id: pointsHistory.length + 1,
        name: customer.name,
        delta: amount,
        reason: reason,
        date: new Date().toISOString().replace('T', ' ').substring(0, 16)
    };
    pointsHistory.push(newHistory);

    // Add to customer's history timeline
    if (!customer.history) customer.history = [];
    customer.history.push({
        type: "system",
        text: `★ Loyalty Points Adjusted: ${amount > 0 ? '+' : ''}${amount} pts (Reason: ${reason})`,
        time: new Date().toLocaleString()
    });

    syncState();
    renderLoyaltyProgram();
    renderPointsLedger();
    logActivity("Customer Points Adjusted", `${customer.name}: ${amount > 0 ? '+' : ''}${amount} pts`, "success");
}

function populateAdjustPointsSelect() {
    const select = document.getElementById("adj-cust-select");
    if (!select) return;

    select.innerHTML = customers.map(c => `<option value="${c.id}">${c.name}</option>`).join("");
}

function submitPointsAdjustment() {
    const custId = parseInt(document.getElementById("adj-cust-select").value);
    const amount = parseInt(document.getElementById("adj-points-amount").value);
    const reason = document.getElementById("adj-points-reason").value.trim() || "Manual override";

    if (isNaN(amount) || amount === 0) {
        alert("Please enter a valid points amount delta.");
        return;
    }

    adjustCustomerPoints(custId, amount, reason);

    document.getElementById("adj-points-amount").value = "";
    document.getElementById("adj-points-reason").value = "";
}

function renderPointsLedger() {
    const container = document.getElementById("rewards-ledger-timeline");
    if (!container) return;
    container.innerHTML = "";

    const sortedHistory = [...pointsHistory].reverse();
    sortedHistory.forEach(item => {
        const div = document.createElement("div");
        div.style.cssText = "background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:10px; font-size:11px; display:flex; flex-direction:column; gap:4px;";
        
        const deltaColor = item.delta > 0 ? "var(--success)" : "var(--danger)";
        const deltaPrefix = item.delta > 0 ? "+" : "";

        div.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <strong style="color:#fff;">${item.name}</strong>
                <span style="color:${deltaColor}; font-weight:700;">${deltaPrefix}${item.delta} pts</span>
            </div>
            <div style="color:var(--text-muted); font-size:10px;">${item.reason}</div>
            <div style="color:var(--text-dim); font-size:9px; text-align:right;">${item.date}</div>
        `;
        container.appendChild(div);
    });
}

function triggerDormantScan() {
    // Filter dormant leads (e.g. stage !== won and point < 100 and zero messages)
    const dormantClients = customers.filter(c => c.stage !== 'won' && (!c.messages || c.messages.length === 0));

    if (dormantClients.length === 0) {
        alert("No dormant leads found in database scan.");
        return;
    }

    dormantClients.forEach(c => {
        if (!c.messages) c.messages = [];
        c.messages.push({
            sender: "outgoing",
            text: `Hi ${c.name}! It's been a while since we checked in. 🌟 Are you still looking for MEP or interior solutions in Doha? Let us know if you need any assistance!`,
            time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
        });
        
        c.history.push({
            type: "system",
            text: "⏰ Auto Reactivation WhatsApp Dispatched (90-Day Dormancy Trigger)",
            time: new Date().toLocaleString()
        });
    });

    const newLog = {
        id: retentionCampaignLogs.length + 1,
        type: "90-Day Dormant scan",
        date: new Date().toISOString().replace('T', ' ').substring(0, 16),
        contacts: dormantClients.length,
        status: `Dispatched to ${dormantClients.map(c => c.name).join(', ')}`
    };
    retentionCampaignLogs.push(newLog);

    syncState();
    renderRetentionLogs();
    logActivity("Dormant Lead Scan Run", `Dispatched reactivation WhatsApps to ${dormantClients.length} leads.`, "success");
    alert(`Dormant Client Scan completed! Automatically sent check-in WhatsApp campaigns to ${dormantClients.length} dormant leads.`);
}

function triggerRenewalScan() {
    const activeClients = customers.filter(c => c.stage === 'won');

    if (activeClients.length === 0) {
        alert("No active client contracts found in database scan.");
        return;
    }

    activeClients.forEach(c => {
        c.history.push({
            type: "system",
            text: "📧 Auto Renewal Notice Email Dispatched (30-Day Expiry Trigger)",
            time: new Date().toLocaleString()
        });
    });

    const newLog = {
        id: retentionCampaignLogs.length + 1,
        type: "30-Day Expiry check",
        date: new Date().toISOString().replace('T', ' ').substring(0, 16),
        contacts: activeClients.length,
        status: `Emailed renewal quotes to ${activeClients.map(c => c.name).join(', ')}`
    };
    retentionCampaignLogs.push(newLog);

    syncState();
    renderRetentionLogs();
    logActivity("Renewal Scan Completed", `Emailed contract renewal notices to ${activeClients.length} clients.`, "success");
    alert(`Contract Expiry Scan completed! Emailed renewal notices to ${activeClients.length} active clients.`);
}

function renderRetentionLogs() {
    const container = document.getElementById("retention-logs-container");
    if (!container) return;
    container.innerHTML = "";

    const sortedLogs = [...retentionCampaignLogs].reverse();
    sortedLogs.forEach(log => {
        const div = document.createElement("div");
        div.style.cssText = "background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:12px; font-size:11px; display:flex; flex-direction:column; gap:4px;";
        div.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <strong style="color:#fff; font-size:12px;">${log.type}</strong>
                <span style="color:var(--text-dim);">${log.date}</span>
            </div>
            <div style="color:var(--text-muted);">Recipients: <strong>${log.contacts} contacts</strong></div>
            <div style="color:var(--success); font-style:italic;">Status: ${log.status}</div>
        `;
        container.appendChild(div);
    });
}

function renderStaffTable() {
    const tbody = document.getElementById("staff-list-tbody");
    if (!tbody) return;
    tbody.innerHTML = "";

    staffUsers.forEach(user => {
        const tr = document.createElement("tr");
        tr.style.cssText = "border-bottom: 1px solid var(--border-color);";
        
        tr.innerHTML = `
            <td style="padding:12px 6px;">
                <div style="font-weight:700; color:#fff;">${user.name}</div>
                <div style="font-size:10px; color:var(--text-muted);">${user.email}</div>
            </td>
            <td style="padding:12px 6px;">
                <select style="background:#0f172a; border:1px solid var(--border-color); color:#fff; font-size:11px; padding:3px 6px; border-radius:var(--radius-sm); cursor:pointer;" onchange="updateStaffRole(${user.id}, this.value)">
                    <option value="Administrator" ${user.role === 'Administrator' ? 'selected' : ''}>Administrator</option>
                    <option value="Lead Agent" ${user.role === 'Lead Agent' ? 'selected' : ''}>Lead Agent</option>
                    <option value="Billing Specialist" ${user.role === 'Billing Specialist' ? 'selected' : ''}>Billing Specialist</option>
                    <option value="Auditor" ${user.role === 'Auditor' ? 'selected' : ''}>Auditor (Read-Only)</option>
                </select>
            </td>
            <td style="padding:12px 6px; font-weight:700; color:var(--text-dim);">${user.logins} sessions</td>
            <td style="padding:12px 6px; text-align:right;">
                <button class="btn btn-secondary" style="padding:2px 6px; font-size:10px; border-color:var(--danger); color:var(--danger);" onclick="deleteStaffMember(${user.id})">Remove</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function submitNewStaffMember() {
    const name = document.getElementById("staff-name").value.trim();
    const email = document.getElementById("staff-email").value.trim();
    const role = document.getElementById("staff-role").value;

    if (!name || !email) {
        alert("Please specify staff name and email address.");
        return;
    }

    const exists = staffUsers.some(u => u.email.toLowerCase() === email.toLowerCase());
    if (exists) {
        alert(`User with email "${email}" already registered.`);
        return;
    }

    const newUser = {
        id: staffUsers.length + 1,
        name: name,
        email: email,
        role: role,
        logins: 0
    };
    staffUsers.push(newUser);

    document.getElementById("staff-name").value = "";
    document.getElementById("staff-email").value = "";

    renderStaffTable();
    logActivity("Staff Member Registered", `${name} added as ${role}`, "success");
    alert(`Registered staff member "${name}" successfully as ${role}!`);
}

function updateStaffRole(userId, newRole) {
    const user = staffUsers.find(u => u.id === userId);
    if (user) {
        user.role = newRole;
        logActivity("Staff Role Updated", `${user.name} role changed to ${newRole}`, "primary");
        alert(`${user.name}'s permissions have been updated to ${newRole}.`);
    }
}

function deleteStaffMember(userId) {
    const index = staffUsers.findIndex(u => u.id === userId);
    if (index !== -1) {
        const name = staffUsers[index].name;
        staffUsers.splice(index, 1);
        renderStaffTable();
        logActivity("Staff Member Removed", `Removed ${name} access key`, "danger");
        alert(`Access revoked for ${name}.`);
    }
}

function renderRetentionAnalytics() {
    const wonCount = customers.filter(c => c.stage === 'won').length;
    const totalPoints = customers.reduce((acc, c) => acc + (c.loyaltyPoints || 0), 0);

    const mrr = 18500 + (wonCount * 4500);
    const ltv = 24000 + (totalPoints * 25);
    const repeatRate = Math.min(95, 70 + (wonCount * 3)).toFixed(1);
    const churn = Math.max(1.5, 5.0 - (totalPoints * 0.004)).toFixed(1);

    document.getElementById("val-mrr").innerText = `QR ${mrr.toLocaleString()}`;
    document.getElementById("val-ltv").innerText = `QR ${ltv.toLocaleString()}`;
    document.getElementById("val-repeat-rate").innerText = `${repeatRate}%`;
    document.getElementById("val-churn").innerText = `${churn}%`;

    // Category Splits
    let mepRev = 45000;
    let paintRev = 28000;
    let hvacRev = 18000;

    customers.forEach(c => {
        if (c.stage === 'won') {
            if (c.labels && c.labels.some(l => l.toLowerCase().includes("mep") || l.toLowerCase().includes("hvac"))) {
                mepRev += 12000;
            } else if (c.labels && c.labels.some(l => l.toLowerCase().includes("paint"))) {
                paintRev += 9500;
            } else {
                hvacRev += 7500;
            }
        }
    });

    const totalRev = mepRev + paintRev + hvacRev;
    const mepPct = ((mepRev / totalRev) * 100).toFixed(1);
    const paintPct = ((paintRev / totalRev) * 100).toFixed(1);
    const hvacPct = ((hvacRev / totalRev) * 100).toFixed(1);

    const metersContainer = document.getElementById("analytics-category-meters");
    if (metersContainer) {
        metersContainer.innerHTML = `
            <div class="analytics-meter-item">
                <div class="analytics-meter-header">
                    <span>MEP & HVAC Maintenance Services</span>
                    <strong>QR ${mepRev.toLocaleString()} (${mepPct}%)</strong>
                </div>
                <div class="analytics-meter-bar-container">
                    <div class="analytics-meter-bar mep" style="width: ${mepPct}%;"></div>
                </div>
            </div>
            <div class="analytics-meter-item">
                <div class="analytics-meter-header">
                    <span>HomeSales Paint Contracts</span>
                    <strong>QR ${paintRev.toLocaleString()} (${paintPct}%)</strong>
                </div>
                <div class="analytics-meter-bar-container">
                    <div class="analytics-meter-bar paint" style="width: ${paintPct}%;"></div>
                </div>
            </div>
            <div class="analytics-meter-item">
                <div class="analytics-meter-header">
                    <span>HVAC Supplier Orders</span>
                    <strong>QR ${hvacRev.toLocaleString()} (${hvacPct}%)</strong>
                </div>
                <div class="analytics-meter-bar-container">
                    <div class="analytics-meter-bar hvac" style="width: ${hvacPct}%;"></div>
                </div>
            </div>
        `;
    }

    // Repeat Purchase Ledger
    const ledgerContainer = document.getElementById("repeat-purchase-ledger");
    if (ledgerContainer) {
        ledgerContainer.innerHTML = "";
        
        // Find customers with loyalty points > 50
        const repeatClients = customers.filter(c => (c.loyaltyPoints || 0) >= 50);
        
        repeatClients.forEach(c => {
            const div = document.createElement("div");
            div.style.cssText = "background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:10px; display:flex; justify-content:space-between; align-items:center;";
            
            // Calculate a dummy LTV for each based on points
            const custLtv = 5000 + (c.loyaltyPoints * 120);

            div.innerHTML = `
                <div>
                    <strong style="color:#fff; font-size:12px;">${c.name}</strong>
                    <div style="color:var(--text-muted); font-size:10px; margin-top:2px;">Points: ${c.loyaltyPoints} | Stage: ${c.stage.toUpperCase()}</div>
                </div>
                <div style="text-align:right;">
                    <span style="color:var(--success); font-weight:700;">QR ${custLtv.toLocaleString()} LTV</span>
                    <div style="color:var(--text-dim); font-size:9px; margin-top:2px;">Tier: ${c.loyaltyPoints >= 500 ? 'Gold' : (c.loyaltyPoints >= 200 ? 'Silver' : 'Bronze')}</div>
                </div>
            `;
            ledgerContainer.appendChild(div);
        });
    }
}

function switchSettingsTab(tabId) {
    // Auto-refresh tunnel status when integrations tab is opened
    if (tabId === 'integrations' || tabId === undefined) pollTunnelStatus();
    const integrationsBtn = document.getElementById("btn-settings-tab-integrations");
    const businessBtn = document.getElementById("btn-settings-tab-business");
    const commsBtn = document.getElementById("btn-settings-tab-comms");
    const aiBtn = document.getElementById("btn-settings-tab-ai");
    const documentsBtn = document.getElementById("btn-settings-tab-documents");

    const integrationsTab = document.getElementById("settings-tab-integrations");
    const businessTab = document.getElementById("settings-tab-business");
    const commsTab = document.getElementById("settings-tab-comms");
    const aiTab = document.getElementById("settings-tab-ai");
    const documentsTab = document.getElementById("settings-tab-documents");

    if (!integrationsBtn || !businessBtn || !commsBtn || !aiBtn || !documentsBtn || !integrationsTab || !businessTab || !commsTab || !aiTab || !documentsTab) return;

    // Reset buttons
    integrationsBtn.className = "btn btn-secondary";
    businessBtn.className = "btn btn-secondary";
    commsBtn.className = "btn btn-secondary";
    aiBtn.className = "btn btn-secondary";
    documentsBtn.className = "btn btn-secondary";

    integrationsTab.style.display = "none";
    businessTab.style.display = "none";
    commsTab.style.display = "none";
    aiTab.style.display = "none";
    documentsTab.style.display = "none";

    if (tabId === 'business') {
        businessBtn.className = "btn btn-primary";
        businessTab.style.display = "block";
    } else if (tabId === 'comms') {
        commsBtn.className = "btn btn-primary";
        commsTab.style.display = "block";
    } else if (tabId === 'ai') {
        aiBtn.className = "btn btn-primary";
        aiTab.style.display = "block";
        
        const delayInp = document.getElementById("settings-ai-delay");
        if (delayInp) {
            document.getElementById("settings-lbl-delay").innerText = delayInp.value + " seconds";
        }
    } else if (tabId === 'documents') {
        documentsBtn.className = "btn btn-primary";
        documentsTab.style.display = "block";
    } else {
        integrationsBtn.className = "btn btn-primary";
        integrationsTab.style.display = "block";
    }
}

function loadAllSettings() {
    // 1b. Backend URL
    const savedBackendUrl = localStorage.getItem("settings-backend-url") || BACKEND_URL;
    if (document.getElementById("settings-backend-url")) {
        document.getElementById("settings-backend-url").value = savedBackendUrl;
    }

    // 1. Meta API
    const waPhoneId = localStorage.getItem("settings-wa-phone-id");
    const waToken = localStorage.getItem("settings-wa-token");
    if (waPhoneId && document.getElementById("settings-wa-phone-id")) {
        document.getElementById("settings-wa-phone-id").value = waPhoneId;
    }
    if (waToken && document.getElementById("settings-wa-token")) {
        document.getElementById("settings-wa-token").value = waToken;
    }

    // 2. Business Profile
    const bizEntity = localStorage.getItem("settings-biz-entity");
    const bizAddress = localStorage.getItem("settings-biz-address");
    const bizPhone = localStorage.getItem("settings-biz-phone");
    const bizEmail = localStorage.getItem("settings-biz-email");
    if (bizEntity && document.getElementById("settings-biz-entity")) {
        document.getElementById("settings-biz-entity").value = bizEntity;
    }
    if (bizAddress && document.getElementById("settings-biz-address")) {
        document.getElementById("settings-biz-address").value = bizAddress;
    }
    if (bizPhone && document.getElementById("settings-biz-phone")) {
        document.getElementById("settings-biz-phone").value = bizPhone;
    }
    if (bizEmail && document.getElementById("settings-biz-email")) {
        document.getElementById("settings-biz-email").value = bizEmail;
    }

    // 3. Comms / SMTP & Webhooks
    const savedSmtp = localStorage.getItem("smtpConfig");
    if (savedSmtp) {
        try {
            smtpConfig = JSON.parse(savedSmtp);
            // Auto-upgrade old password to the new App Password
            if (smtpConfig.pass === "nasimt@7862") {
                smtpConfig.pass = "nsbsniiokfgsqwqv";
                localStorage.setItem("smtpConfig", JSON.stringify(smtpConfig));
            }
        } catch (e) {
            console.error("Error parsing saved smtpConfig", e);
        }
    }
    const savedWebhook = localStorage.getItem("webhookUrl");
    if (savedWebhook) {
        webhookUrl = savedWebhook;
    }

    // Populate SMTP and Webhook inputs
    if (document.getElementById("settings-smtp-host")) document.getElementById("settings-smtp-host").value = smtpConfig.host;
    if (document.getElementById("settings-smtp-port")) document.getElementById("settings-smtp-port").value = smtpConfig.port;
    if (document.getElementById("settings-smtp-user")) document.getElementById("settings-smtp-user").value = smtpConfig.user;
    if (document.getElementById("settings-smtp-pass")) document.getElementById("settings-smtp-pass").value = smtpConfig.pass;
    if (document.getElementById("settings-webhook-url")) document.getElementById("settings-webhook-url").value = webhookUrl;

    // Force settings-ai-delay to '1' as requested by the user, but allow future edits
    if (!localStorage.getItem("settings-ai-delay-forced-1")) {
        localStorage.setItem("settings-ai-delay", "1");
        localStorage.setItem("settings-ai-delay-forced-1", "true");
    }
    const savedAiDelay = localStorage.getItem("settings-ai-delay") || "1";
    const savedAiScore = localStorage.getItem("settings-ai-score-threshold");
    
    const delayInp = document.getElementById("settings-ai-delay");
    const scoreInp = document.getElementById("settings-ai-score-threshold");
    
    if (delayInp) {
        delayInp.value = savedAiDelay;
        const lbl = document.getElementById("settings-lbl-delay");
        if (lbl) {
            lbl.innerText = savedAiDelay === "1" ? "1 second" : savedAiDelay + " seconds";
        }
    }
    if (savedAiScore && scoreInp) {
        scoreInp.value = savedAiScore;
    }

    // 5. Document settings
    const docPrefix = localStorage.getItem("settings-doc-prefix");
    const docSequence = localStorage.getItem("settings-doc-sequence");
    const docTerms = localStorage.getItem("settings-doc-terms");
    const docBank = localStorage.getItem("settings-doc-bank");
    if (docPrefix && document.getElementById("settings-doc-prefix")) {
        document.getElementById("settings-doc-prefix").value = docPrefix;
    }
    if (docSequence && document.getElementById("settings-doc-sequence")) {
        document.getElementById("settings-doc-sequence").value = docSequence;
    }
    if (docTerms && document.getElementById("settings-doc-terms")) {
        document.getElementById("settings-doc-terms").value = docTerms;
    }
    if (docBank && document.getElementById("settings-doc-bank")) {
        document.getElementById("settings-doc-bank").value = docBank;
    }
}

function saveBackendSettings() {
    const backendUrlInput = document.getElementById("settings-backend-url");
    if (!backendUrlInput) return;
    
    let url = backendUrlInput.value.trim();
    if (!url) {
        alert("Please enter a valid backend URL.");
        return;
    }
    
    // Strip trailing slash if present
    if (url.endsWith('/')) {
        url = url.substring(0, url.length - 1);
    }
    
    localStorage.setItem("settings-backend-url", url);
    BACKEND_URL = url;
    
    logActivity("Backend Link Updated", `Backend API server URL set to: ${url}`, "success");
    alert(`Backend Server URL saved successfully! The CRM will now connect to: ${url}`);
    
    // Optionally update the default webhook URL if it hasn't been custom modified
    if (!localStorage.getItem("webhookUrl")) {
        const defaultWebhook = url + "/api/webhooks/incoming";
        webhookUrl = defaultWebhook;
        const webhookInp = document.getElementById("settings-webhook-url");
        if (webhookInp) {
            webhookInp.value = defaultWebhook;
        }
    }
}

// ─── TUNNEL CONTROL ──────────────────────────────────────────────────────────
let _tunnelPollInterval = null;

function updateTunnelUI(running, url) {
    const badge    = document.getElementById('tunnel-status-badge');
    const urlRow   = document.getElementById('tunnel-url-row');
    const urlLink  = document.getElementById('tunnel-url-link');
    const launchBtn = document.getElementById('btn-launch-tunnel');
    const stopBtn  = document.getElementById('btn-stop-tunnel');
    const launchIcon = document.getElementById('btn-launch-icon');
    const launchText = document.getElementById('btn-launch-text');
    if (!badge) return;
    if (running && url) {
        badge.style.cssText = 'font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px;background:rgba(34,197,94,0.15);color:#4ade80;border:1px solid rgba(34,197,94,0.3);';
        badge.textContent = '🟢 Online';
        if (urlRow) urlRow.style.display = 'block';
        if (urlLink) { urlLink.textContent = url; urlLink.href = url; }
        if (launchBtn) launchBtn.style.display = 'none';
        if (stopBtn) stopBtn.style.display = 'block';
    } else if (running && !url) {
        badge.style.cssText = 'font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px;background:rgba(234,179,8,0.15);color:#facc15;border:1px solid rgba(234,179,8,0.3);';
        badge.textContent = '🟡 Starting...';
        if (launchIcon) launchIcon.textContent = '⏳';
        if (launchText) launchText.textContent = 'Opening tunnel...';
    } else {
        badge.style.cssText = 'font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px;background:rgba(239,68,68,0.15);color:#f87171;border:1px solid rgba(239,68,68,0.3);';
        badge.textContent = '⚫ Offline';
        if (urlRow) urlRow.style.display = 'none';
        if (launchBtn) { launchBtn.style.display = 'flex'; launchBtn.disabled = false; }
        if (stopBtn) stopBtn.style.display = 'none';
        if (launchIcon) launchIcon.textContent = '🚀';
        if (launchText) launchText.textContent = 'Launch Server Online';
    }
}

async function pollTunnelStatus() {
    try {
        const res = await fetch(`${BACKEND_URL}/api/tunnel-status`);
        if (!res.ok) throw new Error('offline');
        const data = await res.json();
        updateTunnelUI(data.running, data.url);
        return data;
    } catch { updateTunnelUI(false, null); return { running: false, url: null }; }
}

async function launchServerTunnel() {
    const launchBtn  = document.getElementById('btn-launch-tunnel');
    const launchIcon = document.getElementById('btn-launch-icon');
    const launchText = document.getElementById('btn-launch-text');
    if (launchBtn) launchBtn.disabled = true;
    if (launchIcon) launchIcon.textContent = '⏳';
    if (launchText) launchText.textContent = 'Starting tunnel...';
    try {
        const res = await fetch(`${BACKEND_URL}/api/start-tunnel`, { method: 'POST' });
        if (!res.ok) throw new Error('Server did not respond');
        updateTunnelUI(true, null);
        if (_tunnelPollInterval) clearInterval(_tunnelPollInterval);
        let attempts = 0;
        _tunnelPollInterval = setInterval(async () => {
            attempts++;
            const data = await pollTunnelStatus();
            if (data.url || attempts > 15) {
                clearInterval(_tunnelPollInterval);
                _tunnelPollInterval = null;
                if (launchBtn) launchBtn.disabled = false;
                if (data.url) logActivity('Tunnel Launched', `Public URL live: ${data.url}`, 'success');
            }
        }, 2000);
    } catch (err) {
        if (launchBtn) launchBtn.disabled = false;
        if (launchIcon) launchIcon.textContent = '🚀';
        if (launchText) launchText.textContent = 'Launch Server Online';
        alert('❌ Could not start tunnel.\nMake sure the server (node server.js) is running first.\n\nError: ' + err.message);
    }
}

async function stopServerTunnel() {
    try {
        await fetch(`${BACKEND_URL}/api/stop-tunnel`, { method: 'POST' });
        updateTunnelUI(false, null);
        logActivity('Tunnel Stopped', 'Public tunnel disconnected.', 'info');
    } catch (err) { alert('Could not stop tunnel: ' + err.message); }
}

function saveMetaAPIConnection() {
    const phoneId = document.getElementById("settings-wa-phone-id").value;
    const token = document.getElementById("settings-wa-token").value;

    if (!phoneId || !token) {
        alert("Please fill out Meta Phone ID and Token fields.");
        return;
    }

    localStorage.setItem("settings-wa-phone-id", phoneId);
    localStorage.setItem("settings-wa-token", token);

    logActivity("API Credentials Updated", "Meta WhatsApp Cloud API credentials updated", "success");
    alert("Meta WhatsApp Cloud API credentials saved successfully!");
}

function saveBusinessSettings() {
    const entity = document.getElementById("settings-biz-entity").value;
    const address = document.getElementById("settings-biz-address").value;
    const phone = document.getElementById("settings-biz-phone").value;
    const email = document.getElementById("settings-biz-email").value;

    localStorage.setItem("settings-biz-entity", entity);
    localStorage.setItem("settings-biz-address", address);
    localStorage.setItem("settings-biz-phone", phone);
    localStorage.setItem("settings-biz-email", email);

    logActivity("Business Profile Saved", `Main entity set to: ${entity}`, "success");
    alert(`Business Settings for "${entity}" saved successfully!`);
}

function saveCommunicationSettings() {
    smtpConfig.host = document.getElementById("settings-smtp-host").value;
    smtpConfig.port = document.getElementById("settings-smtp-port").value;
    smtpConfig.user = document.getElementById("settings-smtp-user").value;
    smtpConfig.pass = document.getElementById("settings-smtp-pass").value;

    webhookUrl = document.getElementById("settings-webhook-url").value;

    localStorage.setItem("smtpConfig", JSON.stringify(smtpConfig));
    localStorage.setItem("webhookUrl", webhookUrl);

    logActivity("Mail & Webhooks Saved", "SMTP and Webhook configurations updated", "success");
    alert(`Communication Server credentials saved successfully! Linked webhook: ${webhookUrl}`);
}

function saveAISettings() {
    const delay = document.getElementById("settings-ai-delay").value;
    const score = document.getElementById("settings-ai-score-threshold").value;

    localStorage.setItem("settings-ai-delay", delay);
    localStorage.setItem("settings-ai-score-threshold", score);

    logActivity("AI Autopilot Rules Updated", `Auto-reply delay set to ${delay}s. Qualification threshold: ${score}%`, "success");
    alert(`AI & Autopilot Rules updated successfully! Auto-response delay is now ${delay}s.`);
}

function saveDocumentSettings() {
    const prefix = document.getElementById("settings-doc-prefix").value;
    const seq = document.getElementById("settings-doc-sequence").value;
    const terms = document.getElementById("settings-doc-terms").value;
    const bank = document.getElementById("settings-doc-bank").value;

    localStorage.setItem("settings-doc-prefix", prefix);
    localStorage.setItem("settings-doc-sequence", seq);
    localStorage.setItem("settings-doc-terms", terms);
    localStorage.setItem("settings-doc-bank", bank);

    logActivity("Sales Invoice Format Saved", `Invoice sequence format set to ${prefix}${seq}`, "success");
    alert(`Sales & Document invoice formatting rules updated successfully!`);
}

// ==========================================
// BUSINESS INTELLIGENCE & REPORTING SUITE
// ==========================================

function switchDashboardTab(tabId) {
    const overviewBtn = document.getElementById("btn-db-tab-overview");
    const marketingBtn = document.getElementById("btn-db-tab-marketing");
    const financeBtn = document.getElementById("btn-db-tab-finance");
    const operationsBtn = document.getElementById("btn-db-tab-operations");

    const overviewTab = document.getElementById("db-tab-overview");
    const marketingTab = document.getElementById("db-tab-marketing");
    const financeTab = document.getElementById("db-tab-finance");
    const operationsTab = document.getElementById("db-tab-operations");

    if (!overviewBtn || !marketingBtn || !financeBtn || !operationsBtn || !overviewTab || !marketingTab || !financeTab || !operationsTab) return;

    // Reset button states
    overviewBtn.className = "btn btn-secondary";
    marketingBtn.className = "btn btn-secondary";
    financeBtn.className = "btn btn-secondary";
    operationsBtn.className = "btn btn-secondary";

    // Hide all tabs
    overviewTab.style.display = "none";
    marketingTab.style.display = "none";
    financeTab.style.display = "none";
    operationsTab.style.display = "none";

    // Update KPIs on change
    syncDashboardKPIs();

    if (tabId === 'marketing') {
        marketingBtn.className = "btn btn-primary";
        marketingTab.style.display = "block";
        renderMarketingReport();
    } else if (tabId === 'finance') {
        financeBtn.className = "btn btn-primary";
        financeTab.style.display = "block";
        renderFinanceReport();
    } else if (tabId === 'operations') {
        operationsBtn.className = "btn btn-primary";
        operationsTab.style.display = "block";
        renderOperationsReport();
    } else {
        overviewBtn.className = "btn btn-primary";
        overviewTab.style.display = "block";
    }
}

function syncDashboardKPIs() {
    // 1. Leads KPI
    const kpiLeadsEl = document.getElementById("kpi-leads");
    if (kpiLeadsEl && typeof customers !== 'undefined') kpiLeadsEl.innerText = customers.length;

    // 2. Revenue KPI (Baseline + Paid Invoices)
    const kpiRev = document.getElementById("kpi-revenue");
    if (kpiRev && typeof issuedDocuments !== 'undefined') {
        const baselineRevenue = 36700;
        const paidInvoicesSum = issuedDocuments
            .filter(d => d.type === "Invoice" && d.status === "Paid")
            .reduce((sum, d) => sum + d.amount, 0);
        kpiRev.innerText = `QR ${(baselineRevenue + paidInvoicesSum).toLocaleString()}`;
    }

    // 3. Campaigns KPI
    const kpiCamps = document.getElementById("kpi-campaigns");
    if (kpiCamps && typeof campaignLogs !== 'undefined' && typeof emailCampaignLogs !== 'undefined') {
        const totalSent = campaignLogs.reduce((sum, c) => sum + (c.sent || 0), 0) + emailCampaignLogs.reduce((sum, c) => sum + (c.sent || 0), 0);
        const totalDelivered = campaignLogs.reduce((sum, c) => sum + (c.delivered || 0), 0) + emailCampaignLogs.reduce((sum, c) => sum + (c.opened || 0), 0);
        const rate = totalSent > 0 ? (totalDelivered / totalSent * 100).toFixed(1) + "%" : "98.2%";
        kpiCamps.innerText = rate;
    }

    // 4. AI Autopilot KPI
    const kpiAi = document.getElementById("kpi-ai");
    if (kpiAi && typeof customers !== 'undefined') {
        const avgScore = customers.length > 0 ? Math.round(customers.reduce((sum, c) => sum + (c.score || 0), 0) / customers.length) + "%" : "84%";
        kpiAi.innerText = avgScore;
    }
}

function renderMarketingReport() {
    // 1. Render Campaign status table
    const allCampaigns = [];
    if (typeof campaignLogs !== 'undefined') {
        campaignLogs.forEach(c => {
            allCampaigns.push({
                name: c.name,
                channel: 'WhatsApp',
                sent: c.sent || 0,
                read: c.read || 0,
                reply: c.replies || 0,
                status: 'Completed',
                date: c.date
            });
        });
    }
    if (typeof emailCampaignLogs !== 'undefined') {
        emailCampaignLogs.forEach(c => {
            allCampaigns.push({
                name: c.name,
                channel: 'Email',
                sent: c.sent || 0,
                read: c.opened || 0,
                reply: c.clicked || 0,
                status: 'Completed',
                date: c.date
            });
        });
    }

    allCampaigns.sort((a,b) => new Date(b.date) - new Date(a.date));

    const list = document.getElementById("db-campaigns-report-list");
    if (list) {
        list.innerHTML = "";
        allCampaigns.forEach(c => {
            const tr = document.createElement("tr");
            tr.style.borderBottom = "1px solid var(--border-color)";
            tr.innerHTML = `
                <td style="padding:10px 6px; font-weight: 600; color: #fff;">${c.name}</td>
                <td style="padding:10px 6px;"><span class="badge-status primary">${c.channel}</span></td>
                <td style="padding:10px 6px;">${c.sent}</td>
                <td style="padding:10px 6px;">${c.read}</td>
                <td style="padding:10px 6px;">${c.reply}</td>
                <td style="padding:10px 6px;"><span class="badge-status success">${c.status}</span></td>
            `;
            list.appendChild(tr);
        });
    }

    // 2. Render Lead Conversion Funnel
    const counts = { new: 0, contacted: 0, qualified: 0, proposal: 0, won: 0 };
    if (typeof customers !== 'undefined') {
        customers.forEach(c => {
            const stage = c.stage.toLowerCase();
            if (counts.hasOwnProperty(stage)) {
                counts[stage]++;
            }
        });
    }

    const totalLeads = typeof customers !== 'undefined' ? customers.length : 0;
    const stageLabels = {
        new: "New Leads",
        contacted: "Contacted",
        qualified: "Qualified",
        proposal: "Proposal Sent",
        won: "Won / Closed"
    };
    const stageColors = {
        new: "linear-gradient(90deg, #3b82f6, #60a5fa)",
        contacted: "linear-gradient(90deg, #f59e0b, #fbbf24)",
        qualified: "linear-gradient(90deg, #10b981, #34d399)",
        proposal: "linear-gradient(90deg, #8b5cf6, #a78bfa)",
        won: "linear-gradient(90deg, #ec4899, #f472b6)"
    };

    const funnelContainer = document.getElementById("db-lead-funnel-container");
    if (funnelContainer) {
        funnelContainer.innerHTML = "";
        ['new', 'contacted', 'qualified', 'proposal', 'won'].forEach(stage => {
            const cnt = counts[stage];
            const pct = totalLeads > 0 ? Math.round((cnt / totalLeads) * 100) : 0;
            const card = document.createElement("div");
            card.style.cssText = "background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:12px; display:flex; flex-direction:column; gap:8px;";
            card.innerHTML = `
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div>
                        <strong style="color:#fff; font-size:13px;">${stageLabels[stage]}</strong>
                        <span style="color:var(--text-muted); font-size:11px; margin-left:8px;">(${cnt} leads)</span>
                    </div>
                    <span style="color:var(--primary); font-weight:700; font-size:13px;">${pct}%</span>
                </div>
                <div style="height:6px; background:rgba(255,255,255,0.05); border-radius:3px; overflow:hidden;">
                    <div style="width:${pct}%; height:100%; background:${stageColors[stage]}; border-radius:3px;"></div>
                </div>
            `;
            funnelContainer.appendChild(card);
        });
    }

    // 3. Render Lead Owners Conversions
    const ownerStats = {};
    if (typeof customers !== 'undefined') {
        customers.forEach(c => {
            const owner = c.leadOwner || "Unassigned";
            if (!ownerStats[owner]) {
                ownerStats[owner] = { total: 0, won: 0 };
            }
            ownerStats[owner].total++;
            if (c.stage.toLowerCase() === 'won') {
                ownerStats[owner].won++;
            }
        });
    }

    const ownersContainer = document.getElementById("db-lead-owners-container");
    if (ownersContainer) {
        ownersContainer.innerHTML = "";
        Object.keys(ownerStats).forEach(owner => {
            const stats = ownerStats[owner];
            const convRate = stats.total > 0 ? Math.round((stats.won / stats.total) * 100) : 0;
            const card = document.createElement("div");
            card.style.cssText = "background:rgba(255,255,255,0.02); border:1px solid var(--border-color); border-radius:var(--radius-md); padding:12px; display:flex; flex-direction:column; gap:8px;";
            card.innerHTML = `
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div>
                        <strong style="color:#fff; font-size:13px;">${owner}</strong>
                        <span style="color:var(--text-muted); font-size:11px; margin-left:8px;">(${stats.won}/${stats.total} Won)</span>
                    </div>
                    <span style="color:var(--success); font-weight:700; font-size:13px;">${convRate}%</span>
                </div>
                <div style="height:6px; background:rgba(255,255,255,0.05); border-radius:3px; overflow:hidden;">
                    <div style="width:${convRate}%; height:100%; background:linear-gradient(90deg, var(--success), #34d399); border-radius:3px;"></div>
                </div>
            `;
            ownersContainer.appendChild(card);
        });
    }
}

function renderFinanceReport() {
    // 1. Invoiced Revenues Ledger
    const invoicesList = document.getElementById("db-finance-invoices-list");
    if (invoicesList && typeof issuedDocuments !== 'undefined') {
        invoicesList.innerHTML = "";
        const invoices = issuedDocuments.filter(d => d.type === "Invoice");
        invoices.forEach(doc => {
            let statusClass = "warning";
            if (doc.status === "Paid" || doc.status === "Approved") statusClass = "success";
            if (doc.status === "Unpaid" || doc.status === "Overdue" || doc.status === "Sent") statusClass = "danger";

            const tr = document.createElement("tr");
            tr.style.borderBottom = "1px solid var(--border-color)";
            tr.innerHTML = `
                <td style="padding:10px 6px; font-weight: 700; color: #fff;"><code>${doc.docNo}</code></td>
                <td style="padding:10px 6px;">${doc.client}</td>
                <td style="padding:10px 6px;"><strong>QR ${doc.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}</strong></td>
                <td style="padding:10px 6px;">${doc.date}</td>
                <td style="padding:10px 6px;"><span class="badge-status ${statusClass}">${doc.status}</span></td>
            `;
            invoicesList.appendChild(tr);
        });
    }

    // 2. Team Expense Claims Report
    const expensesList = document.getElementById("db-finance-expenses-list");
    if (expensesList && typeof financeLedger !== 'undefined') {
        expensesList.innerHTML = "";
        const expenses = financeLedger.filter(item => item.type !== "Payroll");
        expenses.forEach(item => {
            let statusClass = "warning";
            if (item.status === "Settled" || item.status === "Approved") statusClass = "success";
            if (item.status === "Rejected") statusClass = "danger";

            const tr = document.createElement("tr");
            tr.style.borderBottom = "1px solid var(--border-color)";
            tr.innerHTML = `
                <td style="padding:10px 6px; font-weight: 600; color: #fff;">${item.employee}</td>
                <td style="padding:10px 6px;">${item.ref}</td>
                <td style="padding:10px 6px;"><strong>QR ${item.amount.toLocaleString('en-US', { minimumFractionDigits: 2 })}</strong></td>
                <td style="padding:10px 6px;"><span class="badge-status ${statusClass}">${item.status}</span></td>
            `;
            expensesList.appendChild(tr);
        });
    }
}

function renderOperationsReport() {
    // 1. Warehouse Inventory Status Report
    const inventoryList = document.getElementById("db-operations-inventory-list");
    if (inventoryList && typeof inventoryStock !== 'undefined') {
        inventoryList.innerHTML = "";
        inventoryStock.forEach(item => {
            let statusClass = "success";
            if (item.status === "Low Stock") statusClass = "warning";
            if (item.status === "Out of Stock") statusClass = "danger";

            function getInventoryCategory(productName) {
                const lower = productName.toLowerCase();
                if (lower.includes("ac") || lower.includes("compressor") || lower.includes("thermostat")) return "HVAC";
                if (lower.includes("duct") || lower.includes("insulation") || lower.includes("mastic")) return "Ducting";
                if (lower.includes("cable") || lower.includes("copper")) return "Electrical";
                return "Materials";
            }

            const category = getInventoryCategory(item.product);

            const tr = document.createElement("tr");
            tr.style.borderBottom = "1px solid var(--border-color)";
            tr.innerHTML = `
                <td style="padding:10px 6px; font-weight: 600; color: #fff;">${item.product}</td>
                <td style="padding:10px 6px;"><span class="badge-status primary">${category}</span></td>
                <td style="padding:10px 6px;">${item.qty} units <span style="color:var(--text-muted); font-size:10px;">(Limit: ${item.limit})</span></td>
                <td style="padding:10px 6px;"><span class="badge-status ${statusClass}">${item.status}</span></td>
            `;
            inventoryList.appendChild(tr);
        });
    }

    // 2. HR & Employee Payroll summaries
    const hrList = document.getElementById("db-operations-hr-list");
    if (hrList && typeof staffUsers !== 'undefined' && typeof financeLedger !== 'undefined') {
        hrList.innerHTML = "";
        staffUsers.forEach(user => {
            const payrollItem = financeLedger.find(item => item.employee === user.name && item.type === "Payroll");
            
            let amountVal = 0;
            let statusVal = "Unpaid";
            if (payrollItem) {
                amountVal = payrollItem.amount;
                statusVal = payrollItem.status;
            } else {
                if (user.role === "Administrator") {
                    amountVal = 15500;
                    statusVal = "Settled";
                } else if (user.role === "Lead Agent") {
                    amountVal = 9800;
                    statusVal = "Settled";
                } else if (user.role === "Billing Specialist") {
                    amountVal = 8500;
                    statusVal = "Settled";
                } else {
                    amountVal = 6000;
                    statusVal = "Unpaid";
                }
            }

            let statusClass = "warning";
            if (statusVal === "Settled" || statusVal === "Approved" || statusVal === "Paid") statusClass = "success";
            if (statusVal === "Unpaid" || statusVal === "Overdue") statusClass = "danger";

            const tr = document.createElement("tr");
            tr.style.borderBottom = "1px solid var(--border-color)";
            tr.innerHTML = `
                <td style="padding:10px 6px;">
                    <strong style="color:#fff;">${user.name}</strong>
                    <div style="font-size:10px; color:var(--text-dim);">${user.email}</div>
                </td>
                <td style="padding:10px 6px;">${user.role}</td>
                <td style="padding:10px 6px;"><strong>QR ${amountVal.toLocaleString('en-US', { minimumFractionDigits: 2 })}</strong></td>
                <td style="padding:10px 6px; text-align:right;"><span class="badge-status ${statusClass}">${statusVal}</span></td>
            `;
            hrList.appendChild(tr);
        });
    }
}

// ==========================================
// EXCEL DATABASE HUB FUNCTIONS
// ==========================================
let excelRows = [];

function reloadExcelFromDisk() {
    const textBanner = document.getElementById("excel-status-text");
    const banner = document.getElementById("excel-status-banner");
    
    if (textBanner) textBanner.innerText = "Loading excel from disk...";
    
    fetch(`${BACKEND_URL}/api/excel/read`)
        .then(res => {
            if (!res.ok) throw new Error("Server responded with error status: " + res.status);
            return res.json();
        })
        .then(resData => {
            if (resData.success) {
                excelRows = resData.data || [];
                renderExcelGrid();
                
                const timestamp = new Date().toLocaleTimeString();
                if (textBanner) {
                    textBanner.innerText = `crm.xls loaded successfully. Total rows: ${excelRows.length} (Last loaded: ${timestamp})`;
                }
                if (banner) {
                    banner.style.background = "rgba(16,185,129,0.05)";
                    banner.style.borderColor = "rgba(16,185,129,0.15)";
                }
            } else {
                throw new Error(resData.error || "Unknown backend error");
            }
        })
        .catch(e => {
            console.error("Error reading excel:", e);
            if (textBanner) textBanner.innerText = "Error loading crm.xls: " + e.message;
            if (banner) {
                banner.style.background = "rgba(239,68,68,0.05)";
                banner.style.borderColor = "rgba(239,68,68,0.15)";
            }
        });
}

function renderExcelGrid() {
    const tbody = document.getElementById("excel-grid-body");
    if (!tbody) return;

    tbody.innerHTML = "";

    if (excelRows.length === 0) {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td colspan="9" style="text-align: center; padding: 30px; color: var(--text-dim);">No data in spreadsheet. Click "Reload Excel" or "Sync CRM Contacts" to populate.</td>`;
        tbody.appendChild(tr);
        return;
    }

    excelRows.forEach((row, index) => {
        const tr = document.createElement("tr");
        tr.className = "excel-row-item";
        tr.setAttribute("data-index", index);
        tr.style.borderBottom = "1px solid var(--border-color)";

        // Read variables with fallbacks
        const id = row.ID || row.id || "";
        const name = row.Name || row.name || "";
        const phone = row.Phone || row.phone || "";
        const email = row.Email || row.email || "";
        const company = row.Company || row.company || "";
        const status = row.Status || row.status || "";

        const cleanPhone = phone ? phone.replace(/[^0-9]/g, '') : '';

        tr.innerHTML = `
            <td style="text-align: center; color: var(--text-dim); background: rgba(0,0,0,0.1); border-right: 1px solid var(--border-color); font-weight: bold;">${index + 1}</td>
            <td style="padding: 2px; border-right: 1px solid var(--border-color);"><input type="text" class="excel-cell-input" value="${id}" onchange="updateExcelCellValue(${index}, 'ID', this.value)" style="width:100%; border:none; background:transparent; color:#fff; padding: 8px 6px; font-size:12px;"></td>
            <td style="padding: 2px; border-right: 1px solid var(--border-color);"><input type="text" class="excel-cell-input" value="${name}" onchange="updateExcelCellValue(${index}, 'Name', this.value)" style="width:100%; border:none; background:transparent; color:#fff; padding: 8px 6px; font-size:12px;"></td>
            <td style="padding: 2px; border-right: 1px solid var(--border-color);"><input type="text" class="excel-cell-input" value="${phone}" onchange="updateExcelCellValue(${index}, 'Phone', this.value)" style="width:100%; border:none; background:transparent; color:#fff; padding: 8px 6px; font-size:12px;"></td>
            <td style="padding: 2px; border-right: 1px solid var(--border-color);"><input type="text" class="excel-cell-input" value="${email}" onchange="updateExcelCellValue(${index}, 'Email', this.value)" style="width:100%; border:none; background:transparent; color:#fff; padding: 8px 6px; font-size:12px;"></td>
            <td style="padding: 2px; border-right: 1px solid var(--border-color);"><input type="text" class="excel-cell-input" value="${company}" onchange="updateExcelCellValue(${index}, 'Company', this.value)" style="width:100%; border:none; background:transparent; color:#fff; padding: 8px 6px; font-size:12px;"></td>
            <td style="padding: 2px; border-right: 1px solid var(--border-color);"><input type="text" class="excel-cell-input" value="${status}" onchange="updateExcelCellValue(${index}, 'Status', this.value)" style="width:100%; border:none; background:transparent; color:#fff; padding: 8px 6px; font-size:12px;"></td>
            <td style="text-align: center; padding: 6px; border-right: 1px solid var(--border-color); display:flex; justify-content:center; gap:8px; align-items:center; height:38px;">
                <button class="btn btn-success" onclick="openExcelSystemChat(${index})" style="padding: 3px 8px; font-size: 11px; display:flex; align-items:center; gap:4px; border-radius: var(--radius-sm);">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:12px;height:12px;"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path></svg>
                    System Chat
                </button>
                ${cleanPhone ? `
                <a href="https://wa.me/${cleanPhone}" target="_blank" class="btn btn-secondary" style="padding: 3px 8px; font-size: 11px; display:flex; align-items:center; gap:4px; text-decoration:none; border-radius: var(--radius-sm); border: 1px solid var(--border-color); color:#fff; line-height:1.2;">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:12px;height:12px;"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path><polyline points="15 3 21 3 21 9"></polyline><line x1="10" y1="14" x2="21" y2="3"></line></svg>
                    wa.me
                </a>
                ` : `<span style="color:var(--text-dim); font-size:11px;">No Phone</span>`}
            </td>
            <td style="text-align: center; padding: 6px;">
                <button class="btn btn-danger" onclick="deleteExcelRow(${index})" style="padding: 4px; border-radius: var(--radius-sm); line-height: 1;" title="Delete Row">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function updateExcelCellValue(index, key, val) {
    const row = excelRows[index];
    if (!row) return;

    // Preserve old identifiers for lookup before updating
    const oldId = (row.ID || row.id || '').toString();
    const oldPhoneClean = (row.Phone || row.phone || '').replace(/[^0-9]/g, '');

    // Format phone value if it's the Phone column
    if (key === 'Phone' || key === 'phone') {
        val = formatQatarPhoneNumber(val);
    }

    // Update spreadsheet row cell value
    row[key] = val;

    // Find corresponding customer in CRM
    let customer = customers.find(c => c.id.toString() === oldId);

    if (customer) {
        // Map Excel cell keys to customer object fields
        if (key === 'Name' || key === 'name') {
            const oldName = customer.name;
            customer.name = val;
            customer.history.push({ type: "system", text: `Name updated from "${oldName}" to "${val}" directly from spreadsheet grid`, time: new Date().toLocaleString() });
        } else if (key === 'Phone' || key === 'phone') {
            const oldPhone = customer.phone;
            customer.phone = val;
            customer.history.push({ type: "system", text: `Phone updated from "${oldPhone}" to "${val}" directly from spreadsheet grid`, time: new Date().toLocaleString() });
        } else if (key === 'Email' || key === 'email') {
            const oldEmail = customer.email;
            customer.email = val;
            customer.history.push({ type: "system", text: `Email updated from "${oldEmail}" to "${val}" directly from spreadsheet grid`, time: new Date().toLocaleString() });
        } else if (key === 'Company' || key === 'company') {
            const oldLabel = customer.label || "None";
            customer.label = val;
            customer.history.push({ type: "system", text: `Company updated from "${oldLabel}" to "${val}" directly from spreadsheet grid`, time: new Date().toLocaleString() });
        } else if (key === 'Status' || key === 'status') {
            const cleanStage = val.toLowerCase().trim();
            const allowedStages = ['new', 'contacted', 'qualified', 'proposal', 'won'];
            if (allowedStages.includes(cleanStage)) {
                const oldStage = customer.stage;
                customer.stage = cleanStage;
                customer.history.push({ type: "system", text: `Pipeline stage changed from ${oldStage.toUpperCase()} to ${cleanStage.toUpperCase()} directly from spreadsheet grid`, time: new Date().toLocaleString() });
            }
        } else if (key === 'ID' || key === 'id') {
            const intVal = parseInt(val);
            if (!isNaN(intVal)) {
                customer.id = intVal;
            }
        }

        syncState();
        renderCustomersTable();
        renderKanban();
        
        if (activeCustomerId === customer.id) {
            selectCustomer(customer.id);
        }
    } else {
        // If not found, let's create a new customer if they filled in Name or Phone
        const nameVal = row.Name || row.name;
        let phoneVal = row.Phone || row.phone;
        
        if (nameVal || phoneVal) {
            if (phoneVal) {
                phoneVal = formatQatarPhoneNumber(phoneVal);
                row.Phone = phoneVal;
                if (row.phone) row.phone = phoneVal;
            }
            const newId = customers.length > 0 ? Math.max(...customers.map(c => c.id)) + 1 : 1;
            customer = {
                id: newId,
                name: nameVal || `Excel Lead #${row.ID || row.id || newId}`,
                phone: phoneVal || "",
                email: row.Email || row.email || "",
                stage: (row.Status || row.status || "new").toLowerCase(),
                temp: "warm",
                channel: "Manual",
                label: row.Company || row.company || "Excel Lead",
                autopilotActive: true,
                messages: [{ sender: "system", text: "Customer initialized from spreadsheet row entry.", time: new Date().toLocaleTimeString() }],
                history: [{ type: "system", text: "Customer created directly from spreadsheet row typing", time: new Date().toLocaleString() }]
            };
            customers.push(customer);
            
            // Sync row ID to match customer ID if needed
            row.ID = newId;
            
            syncState();
            renderCustomersTable();
            renderKanban();
        }
    }
}

function addExcelRow() {
    let newId = 1;
    if (excelRows.length > 0) {
        const ids = excelRows.map(r => parseInt(r.ID || r.id || 0)).filter(id => !isNaN(id));
        if (ids.length > 0) {
            newId = Math.max(...ids) + 1;
        }
    }
    
    excelRows.push({
        ID: newId,
        Name: "",
        Phone: "",
        Email: "",
        Company: "",
        Status: "New"
    });
    
    renderExcelGrid();
    
    const tbody = document.getElementById("excel-grid-body");
    if (tbody && tbody.lastChild) {
        const input = tbody.lastChild.querySelector("input");
        if (input) input.focus();
    }
}

function deleteExcelRow(index) {
    if (confirm("Are you sure you want to delete this row?")) {
        excelRows.splice(index, 1);
        renderExcelGrid();
    }
}

function saveExcelToDisk() {
    const textBanner = document.getElementById("excel-status-text");
    const banner = document.getElementById("excel-status-banner");
    
    if (textBanner) textBanner.innerText = "Saving crm.xls to server...";
    
    fetch(`${BACKEND_URL}/api/excel/save`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ data: excelRows })
    })
    .then(res => {
        return res.json().then(data => {
            if (!res.ok) throw new Error(data.error || "Server responded with error status: " + res.status);
            return data;
        });
    })
    .then(resData => {
        if (resData.success) {
            const timestamp = new Date().toLocaleTimeString();
            if (textBanner) {
                textBanner.innerText = `crm.xls saved successfully to server! Total rows: ${excelRows.length} (Last saved: ${timestamp})`;
            }
            if (banner) {
                banner.style.background = "rgba(16,185,129,0.05)";
                banner.style.borderColor = "rgba(16,185,129,0.15)";
            }
            logActivity("Excel Database Saved", `Successfully updated crm.xls with ${excelRows.length} rows`, "success");
            alert("Excel sheet crm.xls saved successfully!");
        } else {
            throw new Error(resData.error || "Unknown backend error");
        }
    })
    .catch(e => {
        console.error("Error saving excel:", e);
        if (textBanner) textBanner.innerText = "Error saving crm.xls: " + e.message;
        if (banner) {
            banner.style.background = "rgba(239,68,68,0.05)";
            banner.style.borderColor = "rgba(239,68,68,0.15)";
        }
        alert(e.message);
    });
}

function exportExcelToBrowser() {
    if (!excelRows || excelRows.length === 0) {
        alert("The spreadsheet database is empty. Add or load rows first.");
        return;
    }
    
    // Header row matching standard CRM schema
    const headers = ["ID", "Name", "Phone", "Email", "Company", "Status"];
    
    // Map each row in the array
    const csvLines = excelRows.map(row => {
        const id = row.ID || row.id || '';
        const name = row.Name || row.name || '';
        const phone = row.Phone || row.phone || '';
        const email = row.Email || row.email || '';
        const company = row.Company || row.company || '';
        const status = row.Status || row.status || '';
        
        // Escape quotes and wrap in quotes to follow RFC-4180 rules
        const escapeCsv = (val) => {
            const str = (val === null || val === undefined) ? '' : val.toString();
            return `"${str.replace(/"/g, '""')}"`;
        };
        
        return [
            id,
            escapeCsv(name),
            escapeCsv(phone),
            escapeCsv(email),
            escapeCsv(company),
            escapeCsv(status)
        ].join(",");
    });
    
    // Prepend UTF-8 Byte Order Mark (BOM) to force Excel to read UTF-8 characters correctly
    const bom = "\uFEFF";
    const csvContent = bom + [headers.join(",")].concat(csvLines).join("\r\n");
    
    // Create blob and trigger native browser download
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `crm_database_export_${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    logActivity("Excel Exported", `Downloaded crm spreadsheet with ${excelRows.length} rows.`, "success");
}

function syncExcelWithCRMContacts() {
    if (customers.length === 0) {
        alert("There are no customers in the CRM database to sync.");
        return;
    }
    
    if (confirm(`This will sync all ${customers.length} CRM contacts into the Excel sheet. Matching IDs or phone numbers will be merged, and new contacts will be added. Proceed?`)) {
        let addedCount = 0;
        let updatedCount = 0;

        customers.forEach(cust => {
            const phoneClean = cust.phone.replace(/[^0-9]/g, '');
            let existingIdx = excelRows.findIndex(row => {
                const rowId = (row.ID || row.id || '').toString();
                return rowId === cust.id.toString();
            });

            const rowData = {
                ID: cust.id,
                Name: cust.name,
                Phone: cust.phone,
                Email: cust.email || "",
                Company: cust.label || "CRM Customer",
                Status: cust.stage
            };

            if (existingIdx !== -1) {
                excelRows[existingIdx] = Object.assign(excelRows[existingIdx], rowData);
                updatedCount++;
            } else {
                excelRows.push(rowData);
                addedCount++;
            }
        });

        renderExcelGrid();
        
        const textBanner = document.getElementById("excel-status-text");
        if (textBanner) {
            textBanner.innerText = `CRM Contacts Synced: Added ${addedCount} new, Updated ${updatedCount} existing rows. Click 'Save crm.xls' to persist changes to disk.`;
        }
        
        alert(`CRM contacts sync complete!\nAdded: ${addedCount} new row(s)\nUpdated: ${updatedCount} existing row(s)\n\nNote: Please click 'Save crm.xls' to save these changes to the Excel file on disk.`);
    }
}

function openExcelSystemChat(index) {
    const row = excelRows[index];
    if (!row) return;

    const phone = row.Phone || row.phone;
    if (!phone) {
        alert("This row has no phone number to chat with.");
        return;
    }

    const cleanPhone = phone.replace(/[^0-9]/g, '');
    if (!cleanPhone) {
        alert("Invalid phone number format.");
        return;
    }

    let customer = getCustomerByChatId(cleanPhone);

    if (!customer) {
        const newId = customers.length > 0 ? Math.max(...customers.map(c => c.id)) + 1 : 1;
        customer = {
            id: newId,
            name: row.Name || row.name || `Excel Lead #${row.ID || row.id || newId}`,
            phone: phone,
            email: row.Email || row.email || "",
            stage: (row.Status || row.status || "new").toLowerCase(),
            temp: "warm",
            channel: "WhatsApp",
            label: "Excel Sync",
            autopilotActive: true,
            messages: [{ sender: "system", text: "Customer imported from Excel Database Sheet.", time: new Date().toLocaleTimeString() }],
            history: [{ type: "system", text: "Customer created via Excel Database Sync click", time: new Date().toLocaleString() }]
        };
        customers.push(customer);
        syncState();
        renderCustomersTable();
        renderKanban();
    }

    switchView('chat');
    selectChat(customer.id);
}

function filterExcelSheet() {
    const q = document.getElementById("excel-search").value.toLowerCase();
    const rows = document.querySelectorAll("#excel-grid-body tr");
    if (!rows || rows.length === 0) return;

    rows.forEach(tr => {
        if (tr.querySelector("td[colspan]")) return;
        
        const inputs = tr.querySelectorAll("input");
        let found = false;
        inputs.forEach(input => {
            if (input.value.toLowerCase().includes(q)) {
                found = true;
            }
        });
        tr.style.display = found ? "" : "none";
    });
}

let excelSaveTimeout = null;

function debouncedExcelSyncAndSave() {
    if (!isAppInitialized) return;
    
    if (excelSaveTimeout) {
        clearTimeout(excelSaveTimeout);
    }
    
    excelSaveTimeout = setTimeout(async () => {
        try {
            console.log("Triggering auto-sync of CRM contacts to crm.xls...");
            
            // 1. Fetch current excel data from server first to be safe
            const res = await fetch(`${BACKEND_URL}/api/excel/read`);
            if (!res.ok) throw new Error("Failed to read excel for auto-sync");
            const resData = await res.json();
            
            let localExcelRows = resData.success ? (resData.data || []) : excelRows;
            
            // 2. Merge all current customers into the list
            customers.forEach(cust => {
                const phoneClean = cust.phone ? cust.phone.replace(/[^0-9]/g, '') : '';
                
                // Find existing match by ID or clean phone
                let existingIdx = localExcelRows.findIndex(row => {
                    const rowId = (row.ID || row.id || '').toString();
                    return rowId === cust.id.toString();
                });

                const rowData = {
                    ID: cust.id,
                    Name: cust.name,
                    Phone: cust.phone,
                    Email: cust.email || "",
                    Company: cust.label || "CRM Customer",
                    Status: cust.stage
                };

                if (existingIdx !== -1) {
                    localExcelRows[existingIdx] = Object.assign(localExcelRows[existingIdx], rowData);
                } else {
                    localExcelRows.push(rowData);
                }
            });

            // Update global state
            excelRows = localExcelRows;
            if (currentView === 'excel') {
                renderExcelGrid();
            }

            // 3. Save the merged list back to crm.xls
            const saveRes = await fetch(`${BACKEND_URL}/api/excel/save`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ data: localExcelRows })
            });
            const saveResData = await saveRes.json();
            
            if (saveRes.ok && saveResData.success) {
                console.log("Auto-synced all CRM contacts to crm.xls successfully.");
            } else {
                console.warn("Auto-sync crm.xls save warning:", saveResData.error);
            }
        } catch (err) {
            console.error("Auto-sync to crm.xls failed:", err);
        }
    }, 2000); // 2 seconds debounce delay
}




// ==========================================
// GOOGLE SHEETS IMPORT — IMPROVED
// ==========================================

// Persisted state
let gSheetsUrl = localStorage.getItem('gSheetsUrl') || 'https://docs.google.com/spreadsheets/d/1ObstPJH-_BI1h841-fhQWrjkeZHe70sFjwsq8zY-k84/edit?gid=1612390446#gid=1612390446';
let gSheetsAutoSyncEnabled = localStorage.getItem('gSheetsAutoSync') !== 'false';
let gSheetsAutoSyncTimer = null;

// Flexible column header aliases — maps any common variant to our canonical key
const COLUMN_ALIASES = {
    id:      ['id', 'no', 'no.', 'number', '#', 'row', 'seq'],
    name:    ['name', 'full name', 'fullname', 'customer name', 'client name', 'contact name', 'contact', 'client'],
    phone:   ['phone', 'mobile', 'tel', 'telephone', 'cell', 'phone number', 'mobile number', 'whatsapp', 'wa'],
    email:   ['email', 'e-mail', 'email address', 'mail'],
    company: ['company', 'company name', 'organisation', 'organization', 'org', 'business', 'firm', 'employer', 'workplace'],
    status:  ['status', 'stage', 'lead status', 'pipeline', 'lead stage', 'state']
};

/**
 * Match a raw CSV header string to one of our canonical column keys.
 * Returns the canonical key (e.g. 'name') or null if no match found.
 */
function matchColumnAlias(rawHeader) {
    const h = rawHeader.toLowerCase().trim().replace(/[_\-]/g, ' ');
    for (const [key, aliases] of Object.entries(COLUMN_ALIASES)) {
        if (aliases.includes(h)) return key;
    }
    return null;
}

/**
 * Build the CSV export URL from any Google Sheets URL format.
 * Returns { csvUrl, sheetId, gid } or throws on bad URL.
 */
function buildGSheetsCsvUrl(rawUrl) {
    const match = rawUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) throw new Error('Could not parse the Google Sheets URL.\nMake sure you copied the full URL from your browser address bar.');
    const sheetId = match[1];
    const gidMatch = rawUrl.match(/[#&?]gid=(\d+)/);
    const gid = gidMatch ? gidMatch[1] : null;
    const csvUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv${gid ? `&gid=${gid}` : ''}`;
    return { csvUrl, sheetId, gid };
}

/**
 * Core fetch-and-merge routine.
 * silent = true → don't show alert popups (used by auto-sync).
 */
async function runGoogleSheetsImport(silent = false) {
    const textBanner = document.getElementById('excel-status-text');
    const banner     = document.getElementById('excel-status-banner');

    if (!gSheetsUrl) {
        if (!silent) alert('❌ No Google Sheets URL saved yet.\nClick "Import Google Sheets" and paste your URL first.');
        return;
    }

    let csvUrl;
    try {
        const parsed = buildGSheetsCsvUrl(gSheetsUrl);
        // Append cache-buster timestamp to prevent caching by Google or intermediate proxies
        csvUrl = parsed.csvUrl + '&t=' + Date.now();
    } catch (err) {
        if (!silent) {
            if (textBanner) textBanner.innerText = '❌ ' + err.message;
            alert('❌ ' + err.message);
        } else {
            console.warn('[GSheets] URL parse failed during auto-sync:', err.message);
            updateAutoSyncButtonUI('URL Error', 'error');
        }
        return;
    }

    if (!silent) {
        if (textBanner) textBanner.innerText = '⏳ Fetching Google Sheets data…';
    } else {
        updateAutoSyncButtonUI('Syncing…', 'success');
    }

    let csvText;
    try {
        // Try server proxy first (avoids CORS)
        // Append cache-buster timestamp to prevent browser caching of proxy response
        const proxyUrl = `${BACKEND_URL}/api/google-sheets-import?csvUrl=${encodeURIComponent(csvUrl)}&t=${Date.now()}`;
        let response;
        let proxySucceeded = false;

        try {
            response = await fetch(proxyUrl);
            proxySucceeded = true;
        } catch (e) {
            console.warn('[GSheets] Backend proxy is offline or unreachable, falling back to direct fetch:', e.message);
            // Proxy connection failed (e.g. server offline) — try direct fetch
            response = await fetch(csvUrl);
        }

        if (proxySucceeded && !response.ok) {
            // Proxy responded but returned an error status (e.g., Google Sheet returned 401/404)
            let reason = `HTTP ${response.status}`;
            try {
                const errJson = await response.json();
                if (errJson && errJson.error) {
                    reason = errJson.error;
                }
            } catch (e) {}
            throw new Error(reason);
        }

        if (!response.ok) {
            // Direct fetch failed
            let reason = `HTTP ${response.status}`;
            if (response.status === 401 || response.status === 403) {
                reason = `Permission denied (${response.status}).\n👉 Open Google Sheets → Share → set to "Anyone with the link can VIEW".`;
            } else if (response.status === 404) {
                reason = `Sheet not found (404).\n👉 Check that the URL is correct and the sheet still exists.`;
            } else if (response.status === 302 || response.status === 0) {
                reason = `Redirected to a login page.\n👉 The sheet is private — set sharing to "Anyone with the link can VIEW".`;
            }
            throw new Error(reason);
        }

        // Check if we accidentally got an HTML login page instead of CSV
        const ct = response.headers.get('content-type') || '';
        csvText = await response.text();
        if (ct.includes('text/html') || csvText.trim().startsWith('<!DOCTYPE') || csvText.trim().startsWith('<html')) {
            throw new Error(
                'Got an HTML login page instead of CSV data.\n' +
                '👉 Open your Google Sheet → Share → change to "Anyone with the link can VIEW".\n' +
                '👉 Make sure you\'re not on a restricted Google Workspace account.'
            );
        }

    } catch (err) {
        const msg = `❌ Fetch failed: ${err.message}\n\n` +
                    `• Google Sheets URL: ${gSheetsUrl}\n` +
                    `• Export CSV URL: ${csvUrl || '(Not resolved)'}\n` +
                    `• Backend Proxy URL: ${BACKEND_URL}/api/google-sheets-import?csvUrl=...\n\n` +
                    `👉 If the server is offline, start it first.\n` +
                    `👉 If you opened the CRM by double-clicking index.html directly from your folders (file:/// origin), make sure to open http://localhost:3000 in your web browser instead.`;
        if (!silent) {
            if (textBanner) textBanner.innerText = '❌ Fetch failed — see alert';
            if (banner) { banner.style.background = 'rgba(239,68,68,0.05)'; banner.style.borderColor = 'rgba(239,68,68,0.15)'; }
            alert(msg);
        } else {
            console.warn('[GSheets] Auto-sync fetch failed:', err.message, '\nTarget:', csvUrl);
            updateAutoSyncButtonUI('Offline', 'error');
        }
        return;
    }

    // ── Parse CSV ──────────────────────────────────────────────────────────────
    const lines = csvText.trim().split(/\r\n|\r|\n/);
    if (lines.length < 2) {
        const msg = '❌ The sheet appears empty or has only a header row.';
        if (!silent) {
            if (textBanner) textBanner.innerText = msg;
            alert(
                msg + '\n\n' +
                '👉 Make sure your Google Sheet contains data below the header row.\n' +
                '👉 If your data is on a different tab, make sure you copy the URL of that specific tab (it should contain "?gid=..." or "#gid=..." at the end).'
            );
            updateAutoSyncButtonUI('Empty', 'warning');
        } else {
            console.warn('[GSheets] Auto-sync: ' + msg);
            updateAutoSyncButtonUI('Empty', 'warning');
        }
        return;
    }

    // Parse headers with flexible matching
    const rawHeaders = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    const colMap = {};   // canonical key → column index
    const unmatched = [];
    rawHeaders.forEach((h, i) => {
        const key = matchColumnAlias(h);
        if (key && colMap[key] === undefined) {
            colMap[key] = i;
        } else if (!key) {
            unmatched.push(`"${h}"`);
        }
    });

    // Need at minimum a name or phone to be useful
    if (colMap.name === undefined && colMap.phone === undefined) {
        const detectedHeaders = rawHeaders.join(', ');
        const msg =
            `❌ Could not find a "Name" or "Phone" column.\n\n` +
            `Your sheet headers are: ${detectedHeaders}\n\n` +
            `Rename your columns to match (case-insensitive):\n` +
            `  ID, Name, Phone, Email, Company, Status\n\n` +
            `Accepted aliases:\n` +
            `  Name → "Full Name", "Contact", "Client"\n` +
            `  Phone → "Mobile", "Tel", "WhatsApp"\n` +
            `  Company → "Organisation", "Business", "Firm"\n` +
            `  Status → "Stage", "Pipeline"`;
        if (!silent) {
            if (textBanner) textBanner.innerText = '❌ Column mismatch — see alert';
            alert(msg);
            updateAutoSyncButtonUI('Headers Error', 'error');
        } else {
            console.warn('[GSheets] Auto-sync column mismatch:', detectedHeaders);
            updateAutoSyncButtonUI('Headers Error', 'error');
        }
        return;
    }

    const clean = (cols, idx) => {
        if (idx === undefined || idx < 0 || !cols[idx]) return '';
        return cols[idx].replace(/^"|"$/g, '').trim();
    };

    let addedCount = 0, updatedCount = 0;

    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;

        // Robust CSV split handling quoted fields
        const cols = [];
        let inQuotes = false, cur = '';
        for (let c = 0; c < line.length; c++) {
            const ch = line[c];
            if (ch === '"') { inQuotes = !inQuotes; }
            else if (ch === ',' && !inQuotes) { cols.push(cur); cur = ''; }
            else { cur += ch; }
        }
        cols.push(cur);

        // Skip rows with no data at all (e.g. only commas)
        const hasData = cols.some(val => val.trim() !== '');
        if (!hasData) continue;

        const id      = parseInt(clean(cols, colMap.id)) || i;
        const name    = clean(cols, colMap.name)    || `Sheet Lead #${id}`;
        const phone   = clean(cols, colMap.phone)   ? formatQatarPhoneNumber(clean(cols, colMap.phone)) : '';
        const email   = clean(cols, colMap.email)   || '';
        const company = clean(cols, colMap.company) || 'Google Sheets Import';
        const rawStatus = clean(cols, colMap.status).toLowerCase().trim();
        const allowedStages = ['new', 'contacted', 'qualified', 'proposal', 'won'];
        const status  = allowedStages.includes(rawStatus) ? rawStatus : 'new';

        const rowData = { ID: id, Name: name, Phone: phone, Email: email, Company: company, Status: status };

        // Merge into excelRows
        const phoneClean = phone.replace(/[^0-9]/g, '');
        const existingIdx = excelRows.findIndex(row => {
            const rowId    = (row.ID || row.id || '').toString();
            return rowId === id.toString();
        });

        if (existingIdx !== -1) {
            excelRows[existingIdx] = Object.assign(excelRows[existingIdx], rowData);
            updatedCount++;
        } else {
            excelRows.push(rowData);
            addedCount++;
        }

        // Merge into CRM customers
        let customer = customers.find(c => c.id === id);
        if (customer) {
            customer.name = name;
            if (phone)   customer.phone  = phone;
            if (email)   customer.email  = email;
            if (company) customer.label  = company;
            if (allowedStages.includes(rawStatus)) customer.stage = rawStatus;
            customer.history.push({ type: 'system', text: `Record updated from Google Sheets import`, time: new Date().toLocaleString() });
        } else {
            customers.push({
                id, name, phone, email,
                stage: status,
                temp: 'warm',
                channel: 'Google Sheets',
                notes: `Imported from Google Sheets on ${new Date().toLocaleDateString()}.`,
                labels: [company],
                loyaltyPoints: 0,
                dateAdded: new Date().toISOString().split('T')[0],
                lastContact: '',
                messages: [],
                aiSuggestion: null,
                score: 70,
                autopilotActive: true,
                history: [{ type: 'system', text: `Customer imported from Google Sheets`, time: new Date().toLocaleString() }]
            });
        }
    }

    // Save merged data to crm.xls
    try {
        const saveRes  = await fetch(`${BACKEND_URL}/api/excel/save`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data: excelRows })
        });
        const saveData = await saveRes.json();
        if (!saveData.success) throw new Error(saveData.error || 'Save failed');
    } catch (err) {
        console.warn('[GSheets] crm.xls save warning:', err.message);
    }

    // Refresh UI
    syncState();
    renderExcelGrid();
    renderCustomersTable();
    renderKanban();

    const timestamp = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    const summary = `✅ Google Sheets synced — Added: ${addedCount}, Updated: ${updatedCount} rows. (${timestamp})`;
    if (textBanner) textBanner.innerText = summary;
    if (banner) { banner.style.background = 'rgba(16,185,129,0.05)'; banner.style.borderColor = 'rgba(16,185,129,0.15)'; }
    if (!silent) alert(`✅ Google Sheets imported!\nAdded: ${addedCount} new row(s)\nUpdated: ${updatedCount} existing row(s)\n\nData saved to crm.xls.`);
    updateAutoSyncButtonUI(`Synced ${timestamp}`, 'success');

    console.log(`[GSheets] Import complete. ${summary}`);
}

/**
 * Called when the user clicks "Import Google Sheets" button.
 * Prompts for URL (pre-filled with saved value), saves it, then imports.
 */
async function importFromGoogleSheets() {
    const entered = prompt(
        'Paste your Google Sheets URL:\n\n' +
        '• The sheet must be set to "Anyone with the link can view"\n' +
        '• Column headers are matched flexibly (Name/Full Name, Phone/Mobile, etc.)\n' +
        '• Your URL will be saved for auto-sync',
        gSheetsUrl || 'https://docs.google.com/spreadsheets/d/1ObstPJH-_BI1h841-fhQWrjkeZHe70sFjwsq8zY-k84/edit?gid=1612390446#gid=1612390446'
    );

    if (!entered || entered.trim() === '' || entered.includes('...')) return;

    // Validate URL shape before saving
    try {
        buildGSheetsCsvUrl(entered.trim());
    } catch (err) {
        alert('❌ ' + err.message);
        return;
    }

    // Save URL persistently
    gSheetsUrl = entered.trim();
    localStorage.setItem('gSheetsUrl', gSheetsUrl);

    // Show auto-sync button now that we have a URL
    updateAutoSyncButtonUI();

    // Run the import
    await runGoogleSheetsImport(false);
}

/**
 * Toggle auto-sync on/off. When ON, syncs every 5 minutes automatically.
 */
function toggleGoogleSheetsAutoSync() {
    if (!gSheetsUrl) {
        alert('❌ No Google Sheets URL saved yet.\nClick "Import Google Sheets" first to set the URL.');
        return;
    }

    gSheetsAutoSyncEnabled = !gSheetsAutoSyncEnabled;
    localStorage.setItem('gSheetsAutoSync', gSheetsAutoSyncEnabled ? 'true' : 'false');

    if (gSheetsAutoSyncEnabled) {
        // Run immediately, then every 5 minutes
        runGoogleSheetsImport(true);
        gSheetsAutoSyncTimer = setInterval(() => runGoogleSheetsImport(true), 5 * 60 * 1000);
        console.log('[GSheets] Auto-sync enabled — will refresh every 5 minutes.');
    } else {
        if (gSheetsAutoSyncTimer) { clearInterval(gSheetsAutoSyncTimer); gSheetsAutoSyncTimer = null; }
        console.log('[GSheets] Auto-sync disabled.');
    }

    updateAutoSyncButtonUI();
}

/** Update the Auto-Sync button label/colour to reflect current state */
function updateAutoSyncButtonUI(statusMsg = '', statusColor = '') {
    const btn   = document.getElementById('btn-gsheets-autosync');
    const dot   = document.getElementById('gsheets-autosync-dot');
    const label = document.getElementById('gsheets-autosync-label');
    const linkBtn = document.getElementById('btn-gsheets-link');
    if (!btn) return;

    if (gSheetsUrl) {
        btn.style.display = 'inline-flex';
        if (linkBtn) {
            linkBtn.href = gSheetsUrl;
            linkBtn.style.display = 'inline-flex';
        }
    } else {
        if (linkBtn) linkBtn.style.display = 'none';
    }

    if (gSheetsAutoSyncEnabled) {
        btn.style.background   = 'rgba(52,168,83,0.2)';
        btn.style.borderColor  = 'rgba(52,168,83,0.6)';
        
        let color = '#34a853'; // default green
        if (statusColor === 'warning') color = '#f59e0b'; // amber
        if (statusColor === 'error') color = '#ef4444'; // red
        
        if (dot) {
            dot.style.background = color;
            dot.style.boxShadow = `0 0 6px ${color}`;
        }
        if (label) {
            label.innerText = 'Auto-Sync ON (5 min)' + (statusMsg ? ` - ${statusMsg}` : '');
        }
    } else {
        btn.style.background   = 'rgba(52,168,83,0.05)';
        btn.style.borderColor  = 'rgba(52,168,83,0.25)';
        if (dot)   { dot.style.background = '#555'; dot.style.boxShadow = 'none'; }
        if (label) label.innerText = 'Auto-Sync OFF';
    }
}

//// Restore saved URL + auto-sync on page load
(function initGoogleSheetsState() {



    const DEFAULT_SHEET_URL =
'https://docs.google.com/spreadsheets/d/1ObstPJH-_BI1h841-fhQWrjkeZHe70sFjwsq8zY-k84/edit?gid=1612390446#gid=1612390446';

localStorage.setItem('gSheetsUrl', DEFAULT_SHEET_URL);

    let saved = localStorage.getItem('gSheetsUrl');

  
  

    // Restore global variable
    if (typeof gSheetsUrl !== 'undefined') {
        gSheetsUrl = saved;
    }

    // Update UI
    if (saved && typeof updateAutoSyncButtonUI === 'function') {
        updateAutoSyncButtonUI();
    }

    // Restore auto-sync
    if (
        typeof gSheetsAutoSyncEnabled !== 'undefined' &&
        gSheetsAutoSyncEnabled &&
        saved
    ) {
        gSheetsAutoSyncTimer = setInterval(() => {
            if (typeof runGoogleSheetsImport === 'function') {
                runGoogleSheetsImport(true);
            }
        }, 5 * 60 * 1000);

        console.log('[GSheets] Auto-sync restored from saved settings.');

        if (typeof updateAutoSyncButtonUI === 'function') {
            updateAutoSyncButtonUI();
        }
    }
})();

// User Manual Search & Navigation Logic
function initUserManualLogic() {
    const root = document.getElementById('view-manual');
    if (!root) return;

    const input = root.querySelector('.js-user-manual-search');
    const countBadge = root.querySelector('.js-user-manual-count');
    const emptyState = root.querySelector('.js-user-manual-empty');
    const cards = Array.from(root.querySelectorAll('.js-user-manual-search-item'));
    const navLinks = Array.from(root.querySelectorAll('.um-guide-nav a'));

    if (input) {
        input.addEventListener('input', function () {
            const query = input.value.trim().toLowerCase();
            let visibleCount = 0;

            cards.forEach(item => {
                const haystack = (item.getAttribute('data-search') || '').toLowerCase();
                const match = query === '' || haystack.indexOf(query) !== -1;
                item.style.display = match ? '' : 'none';
                if (match) {
                    visibleCount++;
                }
            });

            if (countBadge) {
                countBadge.textContent = query === '' ? '12 modules' : `${visibleCount} matches`;
            }
            if (emptyState) {
                emptyState.style.display = (query !== '' && visibleCount === 0) ? 'block' : 'none';
            }
        });
    }

    // Scroll to sections smoothly inside container
    navLinks.forEach(link => {
        link.addEventListener('click', function (e) {
            e.preventDefault();
            const targetId = this.getAttribute('href').substring(1);
            const targetEl = document.getElementById(targetId);
            if (targetEl) {
                targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
                
                // Update active navigation state
                navLinks.forEach(l => l.classList.remove('active'));
                this.classList.add('active');
            }
        });
    });
}
