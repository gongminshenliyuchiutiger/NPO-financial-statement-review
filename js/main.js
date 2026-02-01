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

    // 1. Create a deep copy of the processed workbook (so we don't modify original if viewed again)
    // SheetJS structures are complex, simple way is to create NEW workbook and append
    const newWb = XLSX.utils.book_new();

    // Copy existing sheets (Standardized Data)
    for (const sheetName of processedWorkbook.SheetNames) {
        const sheet = processedWorkbook.Sheets[sheetName];
        XLSX.utils.book_append_sheet(newWb, sheet, sheetName);
    }

    // 2. Generate Verification Report Sheet
    if (lastResults && lastResults.length > 0) {
        const reportHeaders = ["報表類別", "檢核規則", "檢核結果", "詳細訊息", "異常意義", "修正建議"];
        const reportData = [reportHeaders];

        lastResults.forEach(res => {
            reportData.push([
                res.category,
                res.rule,
                res.status,
                res.message,
                res.significance || "",
                res.suggestion || ""
            ]);
        });

        const reportWs = XLSX.utils.aoa_to_sheet(reportData);

        // Add simple styling (width) info to cols
        reportWs['!cols'] = [
            { wch: 18 }, { wch: 18 }, { wch: 10 }, { wch: 45 }, { wch: 35 }, { wch: 35 }
        ];

        XLSX.utils.book_append_sheet(newWb, reportWs, "稽核報告");
    }

    XLSX.writeFile(newWb, `NPO_財務報表稽核報告_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// End of main.js
