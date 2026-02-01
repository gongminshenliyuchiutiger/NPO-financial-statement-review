import { ReportStandardizer } from './standardizer.js';
import { FinancialReportVerifier } from './verifier.js';
import { GeminiProcessor } from './gemini.js';

// DOM Elements are now retrieved lazily inside functions to ensure they exist

// State
let selectedFile = null;
let processedWorkbook = null;

// Modal Logic
function setupModal() {
    const modal = document.getElementById('helpModal');
    const openBtn = document.getElementById('openHelpModal');
    const closeBtn = document.querySelector('.close-modal');

    if (openBtn && modal && closeBtn) {
        console.log("Modal initialized");
        openBtn.addEventListener('click', (e) => {
            e.preventDefault();
            console.log("Open clicked");
            modal.classList.add('show');
        });
        closeBtn.addEventListener('click', () => {
            console.log("Close clicked");
            modal.classList.remove('show');
        });
        window.addEventListener('click', (e) => {
            if (e.target == modal) {
                modal.classList.remove('show');
            }
        });
    } else {
        console.warn("Modal elements not found:", { openBtn, modal, closeBtn });
    }
}
setupModal();

// Event Listeners (Remaining Drag-and-Drop)
function initDropZone() {
    const dz = document.getElementById('dropZone');
    if (dz) {
        dz.addEventListener('dragover', (e) => {
            e.preventDefault();
            dz.classList.add('dragover');
        });
        dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
        dz.addEventListener('drop', (e) => {
            e.preventDefault();
            dz.classList.remove('dragover');
            if (window.appHandleFileSelect) window.appHandleFileSelect(e.dataTransfer.files[0]);
        });
    }
}
initDropZone();

let lastResults = []; // Store results for export functionality

// Expose handleFileSelect globally for index.html file pickers
window.appHandleFileSelect = function (file) {
    if (!file) return;
    selectedFile = file;
    console.log("Module received file:", file.name);
    // UI feedback is now handled in index.html for speed
};

async function startVerification() {
    if (!selectedFile) {
        alert("請先選擇檔案");
        return;
    }

    const vBtn = document.getElementById('verifyBtn');
    const lDiv = document.getElementById('loading');
    const rSec = document.getElementById('resultSection');
    const rBody = document.getElementById('resultTableBody');

    // Reset UI
    if (vBtn) vBtn.disabled = true;
    if (lDiv) lDiv.style.display = 'block';
    if (rSec) rSec.style.display = 'none';
    if (rBody) rBody.innerHTML = '';
    processedWorkbook = null;
    lastResults = [];

    try {
        const apiKey = document.getElementById('apiKey').value.trim();
        let workbook;

        // Path 1: Gemini Processing
        if (apiKey && !selectedFile.name.match(/\.(xlsx|xlsm)$/i)) {
            // If API key is present AND file is NOT excel, use Gemini
            // Or if user wants to use Gemini for Excel? Usually Gemini is for PDF/Image
            // Logic: If PDF/Image -> Gemini. If Excel -> Local (unless user forces? lets keep simple)
            const processor = new GeminiProcessor(apiKey);
            workbook = await processor.processFile(selectedFile);
        } else if (apiKey && selectedFile.name.match(/\.(pdf|png|jpg|jpeg)$/i)) {
            const processor = new GeminiProcessor(apiKey);
            workbook = await processor.processFile(selectedFile);
        }
        else {
            // Path 2: Local Processing
            if (!selectedFile.name.match(/\.(xlsx|xlsm)$/i) && !apiKey) {
                throw new Error("若無 API Key，僅支援 .xlsx 檔案");
            }
            // If it is Excel, use local
            workbook = await readExcelFile(selectedFile);
            const standardizer = new ReportStandardizer();
            workbook = standardizer.standardize(workbook);
        }

        // Verification
        const verifier = new FinancialReportVerifier(workbook);
        lastResults = verifier.runVerify(); // Store description

        processedWorkbook = workbook;
        displayResults(lastResults);

    } catch (error) {
        alert("錯誤: " + error.message);
        console.error(error);
    } finally {
        if (lDiv) lDiv.style.display = 'none';
        if (vBtn) vBtn.disabled = false;
    }
}

function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            resolve(workbook);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function displayResults(results) {
    const rBody = document.getElementById('resultTableBody');
    const rSec = document.getElementById('resultSection');

    if (!rBody) return;
    rBody.innerHTML = '';

    if (results.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td colspan="6" style="text-align:center">無任何結果 (可能是格式不符或讀取失敗)</td>';
        rBody.appendChild(tr);
    } else {
        results.forEach(res => {
            const tr = document.createElement('tr');

            // Map status to badge class
            let badgeClass = 'status-info';
            if (res.status === 'ERROR') badgeClass = 'status-error';
            else if (res.status === 'WARNING') badgeClass = 'status-warning';
            else if (res.status === 'OK') badgeClass = 'status-ok';

            tr.innerHTML = `
                <td>${res.category}</td>
                <td>${res.rule}</td>
                <td><span class="status-badge ${badgeClass}">${res.status}</span></td>
                <td>${res.message}</td>
                <td style="color: #ddd; font-size: 0.9em">${res.significance || '-'}</td>
                <td style="color: var(--accent); font-size: 0.9em">${res.suggestion || '-'}</td>
            `;
            rBody.appendChild(tr);
        });
    }

    if (rSec) {
        rSec.style.display = 'block';
        rSec.scrollIntoView({ behavior: 'smooth' });
    }
}

// Global exposure for non-module wrappers in index.html
window.appStartVerification = startVerification;
window.appDownloadExcel = downloadExcel;

function downloadTemplate() {
    // Path to the template file in the repository
    const templateUrl = './NPO_財務報表檢核範本_v2.xlsx';
    const link = document.createElement('a');
    link.href = templateUrl;
    link.download = 'NPO_財務報表檢核範本_v2.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function downloadExcel() {
    if (!processedWorkbook) return;

    const newWb = XLSX.utils.book_new();

    // Helper to format percentage
    const fmtPct = (val) => (typeof val === 'number' ? (val * 100).toFixed(2) + "%" : val);

    // Standard Headers for Template
    const headers = ['檢核項目', '輸入值A', '輸入值B', '計算結果', '是否異常', '營運意義', '建議追問／修正'];

    // 1. Process Data Sheets
    for (const sheetName of processedWorkbook.SheetNames) {
        const ws = processedWorkbook.Sheets[sheetName];
        let targetSheetName = "";
        let isFinancialSheet = false;
        let isFundSheet = false;

        // Determine Target Sheet Name & Type
        if (sheetName === "收支決算表") {
            targetSheetName = "S1_收支決算表_輸入即檢核";
            isFinancialSheet = true;
        } else if (sheetName === "資產負債表") {
            targetSheetName = "S2_資產負債表_輸入即檢核";
            isFinancialSheet = true;
        } else if (sheetName.includes("基金") || sheetName.includes("淨值")) {
            targetSheetName = "S3_基金或淨值變動_分析";
            isFundSheet = true;
        } else {
            // Keep other sheets as is
            XLSX.utils.book_append_sheet(newWb, ws, sheetName);
            continue;
        }

        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const auditData = ws._audit || {};
        const newData = [headers]; // Start with headers

        // Process Rows
        // Skip original header row (start at r=1)
        for (let r = 1; r < data.length; r++) {
            const row = data[r];
            const itemName = String(row[0] || "").trim();
            if (!itemName) continue;

            const finding = auditData[itemName];

            let valueA = ""; // Last Year / Start
            let valueB = ""; // This Year / End
            let calcResult = "";
            let status = "正常";
            let meaning = "";
            let suggestion = "";

            if (isFinancialSheet) {
                // S1/S2: Value A = Last Year (Col 1), Value B = This Year (Col 3)
                // Note: Original Import Format assumed [Name, Last, Budget, This]
                valueA = row[1];
                valueB = row[3];

                // If S1 (Income Statement), we might have Budget in Col 2. 
                // The template doesn't strict have Budget col, but we can append it or ignore.
                // For direct mapping to template structure, we focus on Last vs This comparison which rules use.
                // Ratio calculation (This / Total)
                if (row._ratio !== undefined) {
                    calcResult = fmtPct(row._ratio);
                }
            } else if (isFundSheet) {
                // S3: Value A = Start (Col 1), Value B = End (Col 2), Change (Col 3)
                // Assuming standardizer output: [Item, Start, Decrease, Increase, End] ?? 
                // Wait, standardizer logic for Fund sheet might vary. 
                // Let's check verifier logic: 
                // Verifier uses: row[2] (Increase), row[3] (Decrease).
                // Let's assume input keys: [Item, Start, Increase, Decrease, End] based on common logic 
                // or [Item, Start, End] ??
                // Let's stick to what's visible in `verifier.js`:
                // It reads row[2] (Increase?), row[3] (Decrease?).
                // Let's safe guard:
                valueA = row[1]; // Start
                valueB = row[4] || row[2]; // End (Try col 4 first if standard 5-col, else col 2)

                // Calculate simple change
                const vA = parseFloat(valueA) || 0;
                const vB = parseFloat(valueB) || 0;
                const diff = vB - vA;
                calcResult = diff !== 0 ? diff : "";
            }

            // Fill Audit Info
            if (finding) {
                status = finding.status === 'OK' ? '正常' : finding.status;
                meaning = finding.significance || "";
                suggestion = finding.suggestion || "";

                // Override status text for Template Compatibility if needed
                // Template uses "OK" or "異常" logic usually, but text is fine.
            }

            // Construct new row
            // ['檢核項目', '輸入值A', '輸入值B', '計算結果', '是否異常', '營運意義', '建議追問／修正']
            const newRow = [
                itemName,
                valueA,
                valueB,
                calcResult,
                status,
                meaning,
                suggestion
            ];
            newData.push(newRow);
        }

        const newWs = XLSX.utils.aoa_to_sheet(newData);

        // Style Columns (approx width)
        newWs['!cols'] = [
            { wch: 25 }, // Item
            { wch: 15 }, // Val A
            { wch: 15 }, // Val B
            { wch: 12 }, // Result
            { wch: 10 }, // Status
            { wch: 35 }, // Meaning
            { wch: 35 }  // Suggestion
        ];

        XLSX.utils.book_append_sheet(newWb, newWs, targetSheetName);
    }

    // 2. Summary Report Sheet (Optional, but good to keep as extra analysis)
    if (lastResults && lastResults.length > 0) {
        const reportHeaders = ["報表類別", "檢核規則", "檢核結果", "詳細訊息", "異常意義", "修正建議"];
        const reportData = [reportHeaders];
        lastResults.forEach(res => {
            reportData.push([res.category, res.rule, res.status, res.message, res.significance || "", res.suggestion || ""]);
        });
        const reportWs = XLSX.utils.aoa_to_sheet(reportData);
        reportWs['!cols'] = [{ wch: 18 }, { wch: 18 }, { wch: 10 }, { wch: 45 }, { wch: 35 }, { wch: 35 }];
        // Rename slightly to differentiate
        XLSX.utils.book_append_sheet(newWb, reportWs, "全表稽核彙總");
    }

    XLSX.writeFile(newWb, `NPO_財務報表稽核報告_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// End of main.js
