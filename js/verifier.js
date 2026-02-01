export class FinancialReportVerifier {
    constructor(workbook) {
        this.wb = workbook;
        this.results = [];
        // Cross-check values
        this.crossCheck = {
            incomeSurplus: null,
            depreciationExp: 0,
            fixedAssets: 0,
            bsSurplus: null,
            prevAccumulatedSurplus: 0
        };
        this.metrics = {}; // Store raw calculated values for export
    }

    log(category, rule, status, message, significance = "", suggestion = "", targetItem = null) {
        const finding = { category, rule, status, message, significance, suggestion, targetItem };
        this.results.push(finding);

        // Attach to workbook row if possible
        if (targetItem && this.wb.Sheets[category]) {
            const ws = this.wb.Sheets[category];
            // Since we use AOA to sheet conversion, we can't easily find by item name without re-parsing
            // However, we can store it in a map for the exporter to pick up
            if (!ws._audit) ws._audit = {};
            ws._audit[targetItem] = finding;
        }
    }

    runVerify() {
        this.results = [];
        this.metrics = {};
        this._verifyIncomeStatement();
        this._verifyBalanceSheet();
        this._verifyFundStatement();
        this._verifyPropertyCatalog();

        // Attach metrics to the results array logic
        // We can't easily add property to array in JSON if serialized, but within JS it's fine.
        this.results.metrics = this.metrics;
        return this.results;
    }

    _getRows(sheetName) {
        if (!this.wb.Sheets[sheetName]) return null;
        const data = XLSX.utils.sheet_to_json(this.wb.Sheets[sheetName], { header: 1 });
        if (data.length < 2) return [];
        return data.slice(1);
    }

    _verifyIncomeStatement() {
        const category = "收支決算表";
        const rows = this._getRows(category);
        if (!rows) {
            this.log(category, "報表存在性", "ERROR", "找不到 '收支決算表'", "無法進行檢核", "請檢查上傳檔案之工作表名稱");
            return;
        }

        let totalIncome = 0;
        let totalExpense = 0;
        let surplus = 0;

        let manualSumIncome = 0;
        let manualSumExpense = 0;

        rows.forEach(row => {
            const name = String(row[0] || "").trim();
            const lastYear = parseFloat(row[1]) || 0;
            const budget = parseFloat(row[2]) || 0;
            const thisYear = parseFloat(row[3]) || 0;

            if (name.includes("收入") && (name.includes("總額") || name.includes("合計") || name.includes("總計"))) totalIncome = thisYear;
            if (name.includes("支出") && (name.includes("總額") || name.includes("合計") || name.includes("總計"))) totalExpense = thisYear;
            if (name.includes("本期餘絀")) {
                surplus = thisYear;
                this.crossCheck.incomeSurplus = surplus;
            }
            if (name.includes("折舊")) this.crossCheck.depreciationExp += thisYear;

            // Manual Sum Logic (Heuristic: skip if name contains "總額", "合計", "餘絀")
            const isTotalRow = name.includes("總額") || name.includes("合計") || name.includes("餘絀");
            const isIncomeItem = name.includes("收入") && !isTotalRow;
            const isExpenseItem = !name.includes("收入") && !isTotalRow && name.length > 0;

            if (isIncomeItem) manualSumIncome += thisYear;
            if (isExpenseItem) manualSumExpense += thisYear;

            // Individual Item Analysis (Ratios/Variances)
            if (!isTotalRow && name.length > 0) {
                // Attach ratio to row for export
                if (totalIncome > 0 && isIncomeItem) row._ratio = (thisYear / totalIncome);
                if (totalExpense > 0 && isExpenseItem) row._ratio = (thisYear / totalExpense);

                // Rule 5: Budget Variance (>20%)
                if (budget !== 0) {
                    const variance = Math.abs(thisYear - budget) / Math.abs(budget);
                    if (variance > 0.2) {
                        this.log(category, "預決算差異", "WARNING", `[${name}] 差異 ${(variance * 100).toFixed(1)}%`, "預決算差異超過 20%", "請說明差異原因，並檢視是否需提供預決算比較表", name);
                    }
                }

                // Rule 6: YoY Variance (>10%)
                if (lastYear !== 0) {
                    const variance = Math.abs(thisYear - lastYear) / Math.abs(lastYear);
                    if (variance > 0.1) {
                        this.log(category, "年度變動率", "WARNING", `[${name}] 增減 ${(variance * 100).toFixed(1)}%`, "兩年內金額增減幅度超過 10%", "說明其科目性質變動及金額大幅波動之原因", name);
                    }
                }

                // Rule 4: Single Expense > 20% of Total
                if (isExpenseItem && totalExpense > 0) {
                    const ratio = thisYear / totalExpense;
                    if (ratio > 0.2) {
                        this.log(category, "支出集中度", "WARNING", `[${name}] 佔支出 ${(ratio * 100).toFixed(1)}%`, "單項支出佔比過高 (>20%)", "專案或營運成本較集中，請說明其性質來源是否合理", name);
                    }
                }
            }
        });

        // Rule 1: Sum Check
        if (totalIncome > 0 && Math.abs(totalIncome - manualSumIncome) > 10) {
            this.log(category, "合計正確性", "ERROR", `收入合計(${totalIncome}) 與明細加總(${manualSumIncome.toFixed(0)}) 不符`, "報表內容數據可能漏列或加總錯誤", "請重新計算所有收入明細金額");
        }
        if (totalExpense > 0 && Math.abs(totalExpense - manualSumExpense) > 10) {
            this.log(category, "合計正確性", "ERROR", `支出合計(${totalExpense}) 與明細加總(${manualSumExpense.toFixed(0)}) 不符`, "報表內容數據可能漏列或加總錯誤", "請重新計算所有支出明細金額");
        }

        // Rule 2: Balance Check (Revenue = Expense + Surplus)
        if (Math.abs(totalIncome - (totalExpense + surplus)) > 10) {
            this.log(category, "報表平衡", "ERROR", `不平衡: 收入(${totalIncome}) != 支出(${totalExpense}) + 餘絀(${surplus})`, "基本會計恆等式錯誤", "請檢查收入明細、支出明細與餘絀計算是否有一致");
        } else if (totalIncome > 0) {
            this.log(category, "報表平衡", "OK", "收支平衡無誤", "基本會計恆等式正確", "-");
        }

        // Rule 3: Surplus Ratio (<40%)
        if (totalIncome > 0) {
            const ratio = surplus / totalIncome;
            if (ratio > 0.4) {
                this.log(category, "結餘合理性", "WARNING", `餘絀佔收入 ${(ratio * 100).toFixed(1)}%`, "結餘超過當年收入 40%", "請說明結餘之使用計畫，或是否轉列特定資產項目");
            }
        }

        // Save Metrics for Export
        this.metrics['收支決算表'] = {
            totalIncome,
            manualSumIncome,
            totalExpense,
            manualSumExpense,
            surplus,
            crossCheckIncomeSurplus: this.crossCheck.incomeSurplus
        };
    }

    _verifyBalanceSheet() {
        const category = "資產負債表";
        const rows = this._getRows(category);
        if (!rows) {
            this.log(category, "報表存在性", "ERROR", "找不到 '資產負債表'", "無法進行檢核", "-");
            return;
        }

        let totalAssets = 0;
        let totalLiabilities = 0;
        let totalNetValue = 0;
        let bsSurplus = 0;
        let accumSurplusLast = 0;
        let accumSurplusThis = 0;

        let manualSumAssets = 0;
        let manualSumLiab = 0;

        rows.forEach(row => {
            const name = String(row[0] || "").trim();
            const lastYear = parseFloat(row[1]) || 0;
            const thisYear = parseFloat(row[2]) || 0;

            if (name.includes("資產總額") || name === "資產合計") {
                totalAssets = thisYear;
                // Capture index to define Asset Section
            }
            if (name.includes("負債總額") || name === "負債合計") {
                totalLiabilities = thisYear;
            }
            if ((name.includes("淨值總額") || name.includes("權益總額") || name.includes("淨值合計") || (name.includes("基金") && name.includes("合計"))) && !name.includes("負債") && !name.includes("資產")) {
                totalNetValue = thisYear;
                this.crossCheck.totalNetValue = totalNetValue;
            }

            if (name.includes("本期餘絀") || name.includes("本期損益") || name.includes("本期結餘")) {
                bsSurplus = thisYear;
                this.crossCheck.bsSurplus = bsSurplus;
            }

            if (name.includes("累積餘絀")) {
                accumSurplusLast = lastYear;
                accumSurplusThis = thisYear;
                this.crossCheck.prevAccumulatedSurplus = accumSurplusLast;
            }

            if (name.includes("固定資產") && !name.includes("總額") && !name.includes("合計")) {
                this.crossCheck.fixedAssets += thisYear;
            }

            // Rule 6: Debt Composition (Look for specific items)
            if (["銀行借款", "暫收款", "暫付款", "短期借款", "應付票據"].some(k => name.includes(k))) {
                if (thisYear !== 0) {
                    this.log(category, "債務結構", "INFO", `標註科目: [${name}] 金額 ${thisYear}`, "了解債務組成結構及比例", "人工審閱債務償還期限與利息負擔情形");
                }
            }
        });

        // REVISED SUM CHECK LOGIC (Position based)
        // Find indices of delimiters
        let idxTotalAssets = rows.findIndex(r => String(r[0] || "").includes("資產總額") || String(r[0] || "") === "資產合計");
        let idxTotalLiabilities = rows.findIndex(r => String(r[0] || "").includes("負債總額") || String(r[0] || "") === "負債合計");

        // If not found, fallback to name-based or skip
        if (idxTotalAssets !== -1) {
            // Sum Assets: 0 to idxTotalAssets
            for (let i = 0; i < idxTotalAssets; i++) {
                const row = rows[i];
                const name = String(row[0] || "").trim();
                const val = parseFloat(row[2]) || 0;
                // Exclude subtotals
                if (!name.includes("合計") && !name.includes("總額") && !name.includes("小計")) {
                    manualSumAssets += val;
                }
            }
        } else {
            // Fallback Logic
            rows.forEach(row => {
                const name = String(row[0] || "").trim();
                const val = parseFloat(row[2]) || 0;
                if (name.includes("資產") && !name.includes("總額") && !name.includes("合計")) manualSumAssets += val;
            });
        }

        if (idxTotalLiabilities !== -1 && idxTotalAssets !== -1) {
            // Sum Liab: idxTotalAssets + 1 to idxTotalLiabilities
            for (let i = idxTotalAssets + 1; i < idxTotalLiabilities; i++) {
                const row = rows[i];
                const name = String(row[0] || "").trim();
                const val = parseFloat(row[2]) || 0;
                if (!name.includes("合計") && !name.includes("總額") && !name.includes("小計")) {
                    manualSumLiab += val;
                }
            }
        } else {
            // Fallback
            rows.forEach(row => {
                const name = String(row[0] || "").trim();
                const val = parseFloat(row[2]) || 0;
                if (name.includes("負債") && !name.includes("總額") && !name.includes("合計")) manualSumLiab += val;
            });
        }

        // Rule 1: Sum Check (Assets & Liab)
        if (totalAssets > 0 && Math.abs(totalAssets - manualSumAssets) > 100) {
            this.log(category, "合計正確性", "WARNING", `資產總額(${totalAssets.toFixed(0)}) 與科目明細加總(${manualSumAssets.toFixed(0)}) 有顯著差異`, "報表內容數據可能漏列或科目歸類不完全", "請核對資產科目明細是否完整加總");
        }

        // Rule 2: Balance Check (Assets = Liab + Equity)
        if (Math.abs(totalAssets - (totalLiabilities + totalNetValue)) > 10) {
            this.log(category, "報表平衡", "ERROR", `不平衡: 資產(${totalAssets}) != 負債(${totalLiabilities}) + 淨值(${totalNetValue})`, "不符合會計基礎平衡公式", "檢查借貸雙方加總是否正確，或是否漏列科目");
        } else if (totalAssets > 0) {
            this.log(category, "報表平衡", "OK", "資產負債平衡正確", "符合基本會計公式", "-");
        }

        // Rule 3: Cross Check Income vs BS
        if (this.crossCheck.incomeSurplus !== null && Math.abs(bsSurplus - this.crossCheck.incomeSurplus) > 10) {
            this.log(category, "跨表勾稽", "ERROR", `資負表本期餘絀(${bsSurplus}) != 收支表本期餘絀(${this.crossCheck.incomeSurplus})`, "不同報表間數值不一致", "請檢查兩表之本期餘絀計算欄位與對應邏輯");
        }

        // Rule 4: Accumulated Surplus chain check
        // Current Accum = Last Accum + This Year Surplus
        if (accumSurplusThis !== 0 && accumSurplusLast !== 0 && bsSurplus !== 0) {
            const expectedAccum = accumSurplusLast + bsSurplus;
            if (Math.abs(accumSurplusThis - expectedAccum) > 10) {
                this.log(category, "餘絀連動", "WARNING", `累積餘絀(${accumSurplusThis}) != 前一年度累積(${accumSurplusLast}) + 本期(${bsSurplus})`, "累積餘絀與本期運算不連續", "請確認是否曾有盈餘撥充、基金轉列或前期損益調整");
            }
        }

        // Rule 7 (From IS): Depreciation check
        if (this.crossCheck.fixedAssets > 0) {
            if (this.crossCheck.depreciationExp === 0) {
                this.log("收支決算表", "折舊提列情況", "WARNING", `帳列固定資產(${this.crossCheck.fixedAssets.toFixed(0)}) 但收支表未見折舊支出`, "固定資產依規定應定期提列折舊", "請檢查是否遺漏提列折舊，或科目名稱異動");
            } else {
                this.log("收支決算表", "折舊提列情況", "OK", `已提列折舊支出: ${this.crossCheck.depreciationExp}`, "符合會計折舊提列規定", "-");
            }
        }

        // Rule 7 (BS): Debt Ratio (<50%)
        if (totalAssets > 0) {
            const debtRatio = totalLiabilities / totalAssets;
            if (debtRatio > 0.5) {
                this.log(category, "財務風險", "WARNING", `負債佔總資產 ${(debtRatio * 100).toFixed(1)}%`, "負債比率超過 50%", "財務槓桿過高可能導致現金流風險，請確認還款能力或補助款運用情形");
            }
        }

        // Rule 8 (New): Liquidity Ratio (Current Assets / Current Liabilities > 1)
        if (["流動資產", "流動資產合計", "流動資產總額"].some(k => rows.find(r => (r[0] || "").includes(k)))) {
            // Find Current Assets and Current Liabilities
            // Find Current Assets and Current Liabilities
            let currentAssets = 0;
            let currentLiabilities = 0;

            rows.forEach(row => {
                const name = String(row[0] || "").trim();
                const val = parseFloat(row[2]) || 0; // This Year
                if (name === "流動資產" || name === "流動資產合計" || name === "流動資產總額") currentAssets = val;
                if (name === "流動負債" || name === "流動負債合計" || name === "流動負債總額") currentLiabilities = val;
            });

            // Store for export
            this.metrics['資產負債表'].currentAssets = currentAssets;
            this.metrics['資產負債表'].currentLiabilities = currentLiabilities;

            if (currentLiabilities > 0) {
                const liqRatio = currentAssets / currentLiabilities;
                if (liqRatio <= 1) {
                    this.log(category, "流動比率", "WARNING", `流動比率 ${(liqRatio).toFixed(2)} (流動資產/流動負債)`, "流動比率小於 1，短期償債能力有風險", "請確認現金流是否充足");
                } else {
                    this.log(category, "流動比率", "OK", `流動比率 ${(liqRatio).toFixed(2)}`, "短期償債能力正常", "-");
                }
            }
        }

        // Save Metrics for Export
        this.metrics['資產負債表'] = {
            totalAssets,
            manualSumAssets,
            totalLiabilities,
            manualSumLiab,
            totalNetValue,
            bsSurplus
        };
    }

    _verifyFundStatement() {
        const category = "基金及淨值變動表";
        // Attempt to find the sheet using varying names
        let rows = this._getRows(category);
        if (!rows) rows = this._getRows("基金收支表");
        if (!rows) rows = this._getRows("淨值變動表");

        if (!rows) {
            // S3 Logic Fallback: If S3 sheet missing, try to extract Fund/NetValue items from Balance Sheet
            // We can look at BS rows if available
            const bsRows = this._getRows("資產負債表");
            if (bsRows) {
                rows = [];
                bsRows.forEach(row => {
                    const name = String(row[0] || "").trim();
                    if (name.includes("基金") || name.includes("淨值") || name.includes("餘絀") || name.includes("權益")) {
                        if (!name.includes("負債") && !name.includes("資產")) {
                            // Construct pseudo-S3 row: [Name, LastYear(Start), Inc?, Dec?, ThisYear(End)]
                            // We only have Start (Last) and End (This). We don't know Inc/Dec.
                            // We can put: Start, 0, 0, End.
                            // But verification logic checks Math: End = Start + Inc - Dec.
                            // If we put 0, 0, then End must equal Start, which is wrong.
                            // So we construct: Start, (Diff>0?Diff:0), (Diff<0?-Diff:0), End.
                            const start = parseFloat(row[1]) || 0;
                            const end = parseFloat(row[2]) || 0;
                            const diff = end - start;
                            const inc = diff > 0 ? diff : 0;
                            const dec = diff < 0 ? -diff : 0;
                            rows.push([name, start, inc, dec, end]);
                        }
                    }
                });
                if (rows.length > 0) {
                    this.log(category, "資料來源", "INFO", "使用資產負債表資料進行變動分析", "原始檔案無專屬變動表，改以資產負債表對應科目分析", "-");
                }
            }
        }

        if (!rows || rows.length === 0) {
            // Rule: Existence - If BS has Net Assets, FS should exist
            if (this.crossCheck.bsSurplus !== null || this.crossCheck.fixedAssets > 0) {
                this.log(category, "報表存在性", "INFO", "未偵測到 '基金及淨值變動表'", "若組織有基金或淨值變動，應編製此表", "請確認是否漏傳或工作表名稱不符");
            }
            return;
        }

        // Identify Columns: Item, Start, Inc, Dec, End
        let totalEnd = 0;
        let totalStart = 0;
        let totalInc = 0;
        let totalDec = 0;

        rows.forEach((row, idx) => {
            const name = String(row[0] || "").trim();
            if (!name || name === "合計" || name.includes("總額")) return;

            const start = parseFloat(row[1]) || 0;
            const inc = parseFloat(row[2]) || 0;
            const dec = parseFloat(row[3]) || 0;
            const end = parseFloat(row[4]) || 0;

            // Update totals
            totalStart += start;
            totalInc += inc;
            totalDec += dec;
            totalEnd += end;

            // Rule 1: Math Check (End = Start + Inc - Dec)
            // Allow small float error
            const expectedEnd = start + inc - dec;
            if (Math.abs(end - expectedEnd) > 1) {
                this.log(category, "試算正確性", "ERROR", `[${name}] 期末(${end}) != 期初(${start}) + 增加(${inc}) - 減少(${dec})`, "變動金額試算不符", "請檢查該項目之增減金額是否正確");
            }

            // Rule 3: Significant Variation
            // If Increase or Decrease is significant (> 10% of Start, or absolute large amount)
            if (start > 0) {
                if (inc / start > 0.1 && inc > 1000) {
                    this.log(category, "異常變動", "WARNING", `[${name}] 增加金額佔期初 ${(inc / start * 100).toFixed(0)}%`, "當期有顯著增加", "確認是否有大額捐贈或專案結餘轉入");
                }
                if (dec / start > 0.1 && dec > 1000) {
                    this.log(category, "異常變動", "WARNING", `[${name}] 減少金額佔期初 ${(dec / start * 100).toFixed(0)}%`, "當期有顯著減少", "確認是否有用途不符或資產減損");
                }
            } else if (inc > 10000) {
                // New item large amount
                this.log(category, "異常變動", "INFO", `[${name}] 新增金額 ${inc}`, "本期新增項目", "確認來源依據");
            }

            // Rule 4: Negative Balance
            if (end < 0) {
                this.log(category, "餘額合理性", "WARNING", `[${name}] 期末餘額為負數 (${end})`, "基金或淨值通常不應為負", "請檢查是否超支或會計分錄錯誤");
            }
        });

        // Rule 2: Cross Check with BS (Total Net Assets)
        if (this.crossCheck.totalNetValue !== undefined && this.crossCheck.totalNetValue !== 0) {
            if (Math.abs(totalEnd - this.crossCheck.totalNetValue) > 10) {
                this.log(category, "跨表勾稽", "ERROR", `變動表期末合計(${totalEnd}) != 資負表淨值總額(${this.crossCheck.totalNetValue})`, "兩表淨值總額不一致", "請檢查期初餘額引用或本期損益結轉是否正確");
            } else {
                this.log(category, "跨表勾稽", "OK", "變動表與資負表淨值一致", "勾稽相符", "-");
            }
        }
    }

    _verifyPropertyCatalog() {
        const category = "財產目錄";
        const rows = this._getRows(category);
        if (!rows) {
            if (this.crossCheck.fixedAssets > 1000) {
                this.log(category, "目錄完整性", "WARNING", "偵測到固定資產但無目錄", "帳列固定資產需有對應明細清單", "請補齊財產目錄以供核對");
            }
            return;
        }

        let catTotal = 0;
        rows.forEach(row => {
            catTotal += parseFloat(row[4]) || 0; // Book Value Column
        });

        // Rule 5 (BS): Catalog vs BS Fixed Assets
        if (this.crossCheck.fixedAssets > 0 && Math.abs(this.crossCheck.fixedAssets - catTotal) > 100) {
            this.log(category, "帳實相符", "ERROR", `財產目錄總計(${catTotal.toFixed(0)}) != 資負表固定資產(${this.crossCheck.fixedAssets.toFixed(0)})`, "目錄明細與會計帳載金額不符", "重新盤點財產並核對帳載金額，調整差異");
        }
    }
}
