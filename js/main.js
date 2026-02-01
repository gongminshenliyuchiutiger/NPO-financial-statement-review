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

    // 1. Process Data Sheets and Buffer Them
    const sheetBuffer = {
        S1: null,
        S2: null,
        S3: null,
        others: []
    };

    const s1Name = "S1_收支決算表_輸入即檢核";
    const s2Name = "S2_資產負債表_輸入即檢核";
    const s3Name = "S3_基金或淨值變動_分析";

    for (const sheetName of processedWorkbook.SheetNames) {
        const ws = processedWorkbook.Sheets[sheetName];
        let currentHeaders = [];
        let targetType = ""; // S1, S2, S3

        // Determine Target Sheet Name & Type
        if (sheetName === "收支決算表") {
            targetType = "S1";
            currentHeaders = headersS1S2;
        } else if (sheetName === "資產負債表") {
            targetType = "S2";
            currentHeaders = headersS1S2;
        } else if (sheetName.includes("基金") || sheetName.includes("淨值")) {
            targetType = "S3";
            currentHeaders = headersS3;
        } else {
            // Keep other sheets as is (Buffer them)
            sheetBuffer.others.push({ name: sheetName, ws: ws });
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

            if (targetType === "S1" || targetType === "S2") {
                // S1/S2: Value A = Last Year (Col 1)
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
            } else if (targetType === "S3") {
                // S3: Value A = Start (Col 1), Value B = End (Col 4)
                // Assuming standard: Item, Start, Inc, Dec, End
                valueA = row[1];
                valueB = row[4];

                // Calculate simple change
                const vA = parseFloat(valueA) || 0;
                const vB = parseFloat(valueB) || 0;
                calcResult = (vB - vA);
            }

            // Fill Audit Info
            if (finding) {
                status = finding.status === 'OK' ? '正常' : finding.status;
                meaning = finding.significance || "";
                suggestion = finding.suggestion || "";
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

        // --- Prepend Summary Checks (Dashboard) ---
        if (targetType === "S1" || targetType === "S2") {
            const summaryRows = [];
            // Helper to find result logic...
            const findResult = (ruleNamePartial) => {
                return lastResults.find(r => r.category === sheetName && r.rule.includes(ruleNamePartial));
            };
            const createSumRow = (title, ruleKey) => {
                const res = findResult(ruleKey);
                let status = "正常";
                let calc = "TRUE";
                let meaning = "";
                let suggest = "";
                if (res) {
                    status = res.status === 'OK' ? '正常' : res.status;
                    if (res.status !== 'OK') {
                        calc = "FALSE";
                        meaning = res.significance;
                        suggest = res.suggestion;
                    }
                } else {
                    if (ruleKey === "合計正確性") status = "正常";
                }
                return [title, "", "", calc, status, meaning, suggest];
            };
            // Define rows based on Sheet Type
            const dashboard = [];
            const metrics = lastResults.metrics || {};

            if (targetType === "S1") {
                const m = metrics['收支決算表'] || {};
                // 1. Sum Check: Reported Total vs Manual Sum
                dashboard.push(createSumRow("收入加總正確", "合計正確性", m.totalIncome, m.manualSumIncome));
                dashboard.push(createSumRow("支出加總正確", "合計正確性", m.totalExpense, m.manualSumExpense));

                // 2. Balance: Income vs Expense + Surplus
                const calculatedBalance = (m.totalExpense || 0) + (m.surplus || 0);
                dashboard.push(createSumRow("收支平衡", "報表平衡", m.totalIncome, calculatedBalance));

                // 3. Surplus Ratio: Surplus vs Income
                dashboard.push(createSumRow("餘絀比例", "結餘合理性", m.surplus, m.totalIncome));

            } else if (targetType === "S2") {
                const m = metrics['資產負債表'] || {};
                const mIS = metrics['收支決算表'] || {};

                // 1. Balance Check: Assets vs Liab + Equity
                const equity = (m.totalNetValue || 0); // Total Net Value
                const rightSide = (m.totalLiabilities || 0) + equity;
                dashboard.push(createSumRow("資產負債平衡", "報表平衡", m.totalAssets, rightSide));

                // 2. Debt Ratio: Liabilities vs Assets
                dashboard.push(createSumRow("負債比", "財務風險", m.totalLiabilities, m.totalAssets));

                // 3. Surplus Cross Check: BS Surplus vs IS Surplus
                const isSurplus = mIS.crossCheckIncomeSurplus || mIS.surplus;
                dashboard.push(createSumRow("本期餘絀勾稽", "跨表勾稽", m.bsSurplus, isSurplus));

                // 4. Liquidity Ratio: Current Assets vs Current Liab
                dashboard.push(createSumRow("流動比率", "流動比率", m.currentAssets, m.currentLiabilities));
            }
            if (dashboard.length > 0) {
                newData.splice(1, 0, ...dashboard, ["", "", "", "", "", "", ""]);
            }
        }

        const newWs = XLSX.utils.aoa_to_sheet(newData);
        // Style Columns (approx width)
        newWs['!cols'] = [
            { wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 10 }, { wch: 35 }, { wch: 35 }
        ];

        // Store in Buffer
        sheetBuffer[targetType] = newWs;
    }

    // Force create S3 if it doesn't exist
    if (!sheetBuffer.S3) {
        let s3Data = null;

        // Fallback: Try to generate from Balance Sheet
        // Find S2 sheet buffer or raw sheet
        // We might have processed S2 already in buffer, or can read from processedWorkbook
        const bsSheet = processedWorkbook.Sheets["資產負債表"];
        if (bsSheet) {
            const rows = XLSX.utils.sheet_to_json(bsSheet, { header: 1 });
            const newRows = [];
            rows.forEach(row => {
                const name = String(row[0] || "").trim();
                // Filter for Fund/NetValue items
                if ((name.includes("基金") || name.includes("淨值") || name.includes("餘絀") || name.includes("權益")) && !name.includes("負債") && !name.includes("資產")) {
                    const start = parseFloat(row[1]) || 0; // Col 1 Last Year
                    const end = parseFloat(row[2]) || 0;   // Col 2 This Year

                    // Construct S3 Row: [Item, Start, End, Diff, Status, Meaning, Suggestion]
                    // We don't have detailed Inc/Dec, just Diff.
                    const diff = end - start;

                    // Check if there are any specific audit findings for this item from Verifier (S3 category)
                    // Verifier runs on "基金及淨值變動表" category.
                    // The fallback logic in verifier used these same names.
                    // But verifier stores results in `lastResults`.
                    // We can look up findings by Item Name.

                    // Finding might be under "基金及淨值變動表"
                    const finding = lastResults.find(r => r.category === "基金及淨值變動表" && r.targetItem === name) ||
                        (bsSheet._audit ? bsSheet._audit[name] : null); // Fallback to BS audit if any? No, separate category.

                    let status = "正常";
                    let meaning = "";
                    let suggest = "";
                    if (finding) {
                        status = finding.status === 'OK' ? '正常' : finding.status;
                        meaning = finding.significance || "";
                        suggest = finding.suggestion || "";
                    }

                    newRows.push([name, start, end, diff, status, meaning, suggest]);
                }
            });

            if (newRows.length > 0) {
                s3Data = [headersS3, ...newRows];
            }
        }

        if (!s3Data) {
            s3Data = [
                headersS3,
                ["(未偵測到基金或淨值變動表原始資料)", "", "", "", "提示", "請確認上傳檔案是否包含此表", ""]
            ];
        }

        const wsS3 = XLSX.utils.aoa_to_sheet(s3Data);
        wsS3['!cols'] = [{ wch: 30 }];
        sheetBuffer.S3 = wsS3;
    }

    // 2. Append Sheets in Specific Order
    if (sheetBuffer.S1) XLSX.utils.book_append_sheet(newWb, sheetBuffer.S1, s1Name);
    if (sheetBuffer.S2) XLSX.utils.book_append_sheet(newWb, sheetBuffer.S2, s2Name);
    if (sheetBuffer.S3) XLSX.utils.book_append_sheet(newWb, sheetBuffer.S3, s3Name);

    // Append Others (e.g., Property Catalog)
    sheetBuffer.others.forEach(item => {
        XLSX.utils.book_append_sheet(newWb, item.ws, item.name);
    });

    // 3. Summary Report Sheet
    if (lastResults && lastResults.length > 0) {
        const reportHeaders = ["報表類別", "檢核規則", "檢核結果", "詳細訊息", "異常意義", "修正建議"];
        const reportData = [reportHeaders];
        lastResults.forEach(res => {
            reportData.push([res.category, res.rule, res.status, res.message, res.significance || "", res.suggestion || ""]);
        });
        const reportWs = XLSX.utils.aoa_to_sheet(reportData);
        reportWs['!cols'] = [{ wch: 18 }, { wch: 18 }, { wch: 10 }, { wch: 45 }, { wch: 35 }, { wch: 35 }];
        XLSX.utils.book_append_sheet(newWb, reportWs, "全表稽核彙總");
    }

    XLSX.writeFile(newWb, `NPO_財務報表稽核報告_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// End of main.js
