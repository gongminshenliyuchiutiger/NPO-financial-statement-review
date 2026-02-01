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

// Auto-trigger verification on file select
async function handleFileSelect(file) {
    if (!file) return;
    selectedFile = file;
    // UI Feedback in Dropzone
    dropZone.innerHTML = `<p style="font-weight:bold; color:var(--accent)">已選擇: ${file.name}</p>`;

    // Start automatically
    await startVerification();
}

async function startVerification() {
    if (!selectedFile) return;

    // Reset UI
    loadingDiv.style.display = 'block';
    resultsSection.style.display = 'none';
    resultTableBody.innerHTML = '';
    processedWorkbook = null;

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
        const results = verifier.runVerify();

        processedWorkbook = workbook;
        displayResults(results);

    } catch (error) {
        alert("錯誤: " + error.message);
        console.error(error);
    } finally {
        loadingDiv.style.display = 'none';
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
    downloadExcel
}; // Export for inline script usage

function downloadExcel() {
    if (!processedWorkbook) return;
    XLSX.writeFile(processedWorkbook, "標準化檢核報告.xlsx");
}
