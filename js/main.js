import { ReportStandardizer } from './standardizer.js';
import { FinancialReportVerifier } from './verifier.js';
import { GeminiProcessor } from './gemini.js';

// DOM Elements
const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');
const fileInfo = document.getElementById('fileInfo');
const fileNameSpan = document.getElementById('fileName');
const verifyBtn = document.getElementById('verifyBtn');
const loadingDiv = document.getElementById('loading');
const statusText = document.getElementById('statusText');
const resultsSection = document.getElementById('resultsSection');
const resultTableBody = document.querySelector('#resultTable tbody');
const downloadBtn = document.getElementById('downloadBtn');

// State
let selectedFile = null;
let processedWorkbook = null;

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
downloadBtn.addEventListener('click', downloadExcel);

function handleFileSelect(file) {
    if (!file) return;
    selectedFile = file;
    fileNameSpan.textContent = file.name;
    fileInfo.classList.remove('hidden');
    verifyBtn.disabled = false;
    resultsSection.classList.add('hidden');
}

async function startVerification() {
    if (!selectedFile) return;

    // Reset UI
    verifyBtn.disabled = true;
    loadingDiv.classList.remove('hidden');
    resultsSection.classList.add('hidden');
    processedWorkbook = null;

    try {
        const apiKey = document.getElementById('apiKey').value.trim();
        let workbook;

        // Path 1: Gemini Processing
        if (apiKey) {
            statusText.textContent = "正在聯絡 Gemini AI 進行分析...";
            const processor = new GeminiProcessor(apiKey);
            workbook = await processor.processFile(selectedFile);
        }
        // Path 2: Local Processing (Excel Only)
        else {
            if (!selectedFile.name.match(/\.(xlsx|xlsm)$/i)) {
                throw new Error("若無 API Key，僅支援 .xlsx 檔案");
            }
            statusText.textContent = "正在讀取 Excel 檔案...";
            workbook = await readExcelFile(selectedFile);

            statusText.textContent = "標準化格式中...";
            const standardizer = new ReportStandardizer();
            workbook = standardizer.standardize(workbook);
        }

        // Verification
        statusText.textContent = "執行檢查規則...";
        const verifier = new FinancialReportVerifier(workbook);
        const results = verifier.runVerify();

        processedWorkbook = workbook;
        displayResults(results);

    } catch (error) {
        alert("錯誤: " + error.message);
        console.error(error);
    } finally {
        loadingDiv.classList.add('hidden');
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

    let errorCount = 0;
    let warningCount = 0;

    if (results.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td colspan="4" style="text-align:center">無任何結果 (可能是格式不符)</td>';
        resultTableBody.appendChild(tr);
    } else {
        results.forEach(res => {
            const tr = document.createElement('tr');
            const statusClass = res.status === 'ERROR' ? 'status-error' : (res.status === 'WARNING' ? 'status-warning' : 'status-ok');

            if (res.status === 'ERROR') errorCount++;
            if (res.status === 'WARNING') warningCount++;

            tr.innerHTML = `
                <td class="${statusClass}">${res.status}</td>
                <td>${res.category}</td>
                <td>${res.rule}</td>
                <td>${res.message}</td>
            `;
            resultTableBody.appendChild(tr);
        });
    }

    document.getElementById('totalChecks').textContent = results.length;
    document.getElementById('errorCount').textContent = errorCount;
    document.getElementById('warningCount').textContent = warningCount;

    resultsSection.classList.remove('hidden');
}

function downloadExcel() {
    if (!processedWorkbook) return;
    XLSX.writeFile(processedWorkbook, "標準化檢核報告.xlsx");
}
