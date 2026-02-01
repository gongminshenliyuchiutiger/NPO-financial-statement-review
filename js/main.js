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
    const headersS1S2 = ['檢核項目', '輸入值A', '輸入值B', '計算結果', '是否異常', '營運意義', '建議追問／修正'];
    const headersS3 = ['項目', '期初', '期末', '變動數', '判斷', '營運意義', '建議'];

    // 1. Process Data Sheets
    for (const sheetName of processedWorkbook.SheetNames) {
        const ws = processedWorkbook.Sheets[sheetName];
        let targetSheetName = "";
        let isFinancialSheet = false;
        let isFundSheet = false;
        let currentHeaders = [];

        // Determine Target Sheet Name & Type
        if (sheetName === "收支決算表") {
            targetSheetName = "S1_收支決算表_輸入即檢核";
            isFinancialSheet = true;
            currentHeaders = headersS1S2;
        } else if (sheetName === "資產負債表") {
            targetSheetName = "S2_資產負債表_輸入即檢核";
            isFinancialSheet = true;
            currentHeaders = headersS1S2;
        } else if (sheetName.includes("基金") || sheetName.includes("淨值")) {
            targetSheetName = "S3_基金或淨值變動_分析";
            isFundSheet = true;
            currentHeaders = headersS3;
        } else {
            // Keep other sheets as is
            XLSX.utils.book_append_sheet(newWb, ws, sheetName);
            continue;
        }

        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const auditData = ws._audit || {};
        const newData = [currentHeaders]; // Start with specific headers

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
                // S1/S2: Value A = Last Year (Col 1)
                // Note: Original Import Format assumed [Name, Last, Budget, This] for IS
                // But [Name, Last, This, Note] for BS
                valueA = row[1];

                if (sheetName === "收支決算表") {
                    valueB = row[3]; // This Year is Col 3 (0,1,2,3)
                } else if (sheetName === "資產負債表") {
                    valueB = row[2]; // This Year is Col 2 (0,1,2,3)
                } else {
                    valueB = row[3]; // Fallback
                }

                // Ratio calculation (This / Total)
                if (row._ratio !== undefined) {
                    calcResult = fmtPct(row._ratio);
                }
            } else if (isFundSheet) {
                // S3: Value A = Start (Col 1), Value B = End (Col 2 or 4)
                valueA = row[1];

                // Determine End Value (Col 4 if exists, else Col 2)
                if (row.length >= 5) {
                    valueB = row[4];
                } else {
                    valueB = (row[4] !== undefined) ? row[4] : row[2];
                }

                // Calculate simple change
                const vA = parseFloat(valueA) || 0;
                const vB = parseFloat(valueB) || 0;
                const diff = (vB - vA);
                // Simple difference
                calcResult = diff;
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

        // --- NEW LOGIC: Prepend Summary Checks (Dashboard) to match Template ---
        if (isFinancialSheet) {
            const summaryRows = [];

            // Helper to find result
            const findResult = (ruleNamePartial) => {
                return lastResults.find(r => r.category === sheetName && r.rule.includes(ruleNamePartial));
            };

            // Helper to create Summary Row
            const createSumRow = (title, ruleKey, valA = "", valB = "") => {
                const res = findResult(ruleKey);
                // If result found, use its status/message. If not, assume OK or N/A
                // However, verifier only pushes logs for Checks it ran.
                // For "Sum Check", verifier logs ERROR if mismatch.

                let status = "正常";
                let calc = "TRUE";
                let meaning = "";
                let suggest = "";

                if (res) {
                    status = res.status === 'OK' ? '正常' : res.status;
                    if (res.status !== 'OK') {
                        calc = "FALSE"; // Or specific value if available
                        meaning = res.significance;
                        suggest = res.suggestion;
                    } else if (res.message) {
                        // Sometimes we log OK messages
                        calc = "TRUE"; // Or extracted value
                    }
                } else {
                    // Logic didn't trigger? Could be OK or Skipped.
                    // For "Sum Check", if no error logged, usually implies OK in our verifier logic?
                    // Actually, verifier.js ONLY logs Sum Check if it fails (lines 116, 119 for IS; 193 for BS).
                    // WAIT: It DOES log "OK" for Balance Check (line 126, 201).
                    // For Sum Check, it only logs ERROR. So if not found, it is OK.
                    if (ruleKey === "合計正確性") status = "正常";
                }

                return [title, valA, valB, calc, status, meaning, suggest];
            };

            // Define rows based on Sheet Type
            const dashboard = [];
            if (sheetName === "收支決算表") {
                dashboard.push(createSumRow("收入加總正確", "合計正確性")); // Logic: If no error log found, assume OK
                dashboard.push(createSumRow("支出加總正確", "合計正確性"));
                dashboard.push(createSumRow("收支平衡", "報表平衡"));
                dashboard.push(createSumRow("餘絀比例", "結餘合理性"));
            } else if (sheetName === "資產負債表") {
                dashboard.push(createSumRow("資產負債平衡", "報表平衡"));
                dashboard.push(createSumRow("負債比", "財務風險"));
                dashboard.push(createSumRow("本期餘絀勾稽", "跨表勾稽"));
                dashboard.push(createSumRow("流動比率", "流動比率"));
                // Add "Fund Cross Check" if needed, usually S2
            }

            // Insert Dashboard rows at the beginning (after header)
            // newData[0] is Header. Insert at index 1.
            if (dashboard.length > 0) {
                // Add an empty divider row after dashboard? Template doesn't seem to have one but looks cleaner.
                // Template just lists them at top.
                newData.splice(1, 0, ...dashboard, ["", "", "", "", "", "", ""]); // Add separator row
            }
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
