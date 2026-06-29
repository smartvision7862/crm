const XLSX = require('xlsx');

async function main() {
    try {
        console.log("Fetching user spreadsheet in XLSX format...");
        const res = await fetch('https://docs.google.com/spreadsheets/d/1ObstPJH-_BI1h841-fhQWrjkeZHe70sFjwsq8zY-k84/export?format=xlsx');
        if (!res.ok) {
            throw new Error(`Failed to fetch XLSX: ${res.statusText}`);
        }
        const buffer = await res.arrayBuffer();
        console.log("Parsing XLSX...");
        const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
        
        console.log("Sheet names in workbook:", workbook.SheetNames);
        
        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet);
            console.log(`\n--- Sheet: "${sheetName}" (${data.length} rows) ---`);
            if (data.length > 0) {
                console.log("Rows:", data);
            } else {
                console.log("Sheet is empty or has only headers.");
                const headers = XLSX.utils.sheet_to_json(sheet, { header: 1 })[0];
                console.log("Headers:", headers);
            }
        });
    } catch (err) {
        console.error("Error:", err);
    }
}

main();
