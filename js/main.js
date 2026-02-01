import { ReportStandardizer } from './standardizer.js';
import { FinancialReportVerifier } from './verifier.js';
import { GeminiProcessor } from './gemini.js';

// DOM Elements
const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');
const verifyBtn = document.getElementById('verifyBtn'); // Hidden in new design? No, strictly logic flow check
const loadingDiv = document.getElementById('loading');
const resultsSection = document.getElementById('resultSection'); // Corrected ID
const resultTableBody = document.getElementById('resultTableBody');

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

// Event Listeners
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    handleFileSelect(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', (e) => handleFileSelect(e.target.files[0]));
verifyBtn.addEventListener('click', startVerification);

let lastResults = []; // Store results for export functionality

// Auto-trigger verification on file select -> REMOVED AUTO TRIGGER
async function handleFileSelect(file) {
    if (!file) return;
    selectedFile = file;
    // UI Feedback in Dropzone
    dropZone.innerHTML = `<p style="font-weight:bold; color:var(--accent)">已選擇: ${file.name}</p>`;

    // Enable the button, but DO NOT start automatically
    verifyBtn.disabled = false;
    resultsSection.style.display = 'none';
}

async function startVerification() {
    if (!selectedFile) return;

    // Reset UI
    verifyBtn.disabled = true;
    loadingDiv.style.display = 'block';
    resultsSection.style.display = 'none';
    resultTableBody.innerHTML = '';
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
        loadingDiv.style.display = 'none';
        verifyBtn.disabled = false;
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
    resultTableBody.innerHTML = '';

    if (results.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td colspan="6" style="text-align:center">無任何結果 (可能是格式不符或讀取失敗)</td>';
        resultTableBody.appendChild(tr);
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
            resultTableBody.appendChild(tr);
        });
    }

    resultsSection.style.display = 'block';

    // Smooth scroll
    resultsSection.scrollIntoView({ behavior: 'smooth' });
}

export const main = {
    downloadExcel,
    downloadEmptyTemplate
};

function downloadEmptyTemplate() {
    const wb = XLSX.utils.book_new();

    // 1. 收支決算表
    const ws1Data = [
        ["科目名稱", "上年度決算數", "本年度預算數", "本年度決算數"],
        ["收入總額", 0, 0, 0],
        ["支出總額", 0, 0, 0],
        ["本期餘絀", 0, 0, 0]
    ];
    const ws1 = XLSX.utils.aoa_to_sheet(ws1Data);
    XLSX.utils.book_append_sheet(wb, ws1, "收支決算表");

    // 2. 資產負債表
    const ws2Data = [
        ["科目名稱", "上年度金額", "本年度金額"],
        ["資產總額", 0, 0],
        ["負債總額", 0, 0],
        ["基金及餘絀總額", 0, 0],
        ["本期餘絀", 0, 0],
        ["累積餘絀", 0, 0]
    ];
    const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
    XLSX.utils.book_append_sheet(wb, ws2, "資產負債表");

    // 3. 基金收支表
    const ws3Data = [
        ["科目名稱", "期初餘額", "本期增加", "本期減少", "期末餘額"],
        ["基金", 0, 0, 0, 0]
    ];
    const ws3 = XLSX.utils.aoa_to_sheet(ws3Data);
    XLSX.utils.book_append_sheet(wb, ws3, "基金收支表");

    // 4. 財產目錄
    const ws4Data = [
        ["資產名稱", "取得日期", "取得成本", "本期折舊", "帳面價值"],
        ["範例設備", "2023-01-01", 10000, 1000, 9000]
    ];
    const ws4 = XLSX.utils.aoa_to_sheet(ws4Data);
    XLSX.utils.book_append_sheet(wb, ws4, "財產目錄");

    XLSX.writeFile(wb, "NPO_財務報表檢核範本_v3.xlsx");
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
        const reportHeaders = ["類別", "檢核規則", "狀態", "訊息", "異常意義", "修正建議"];
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
            { wch: 15 }, { wch: 15 }, { wch: 10 }, { wch: 50 }, { wch: 30 }, { wch: 30 }
        ];

        XLSX.utils.book_append_sheet(newWb, reportWs, "檢核報告結果");
    }

    XLSX.writeFile(newWb, "NPO_財務報表檢核報告_Pro.xlsx");
}
