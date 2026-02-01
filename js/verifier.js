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
    }

    log(category, rule, status, message, significance = "", suggestion = "") {
        this.results.push({
            category,
            rule,
            status,
            message,
            significance,
            suggestion
        });
    }

    runVerify() {
        this.results = [];
        this._verifyIncomeStatement();
        this._verifyBalanceSheet();
        this._verifyFundStatement();
        this._verifyPropertyCatalog();
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
        let calcSurplus = 0;

        let incomeItems = [];
        let expenseItems = [];

        rows.forEach(row => {
            const name = String(row[0] || "").trim();
            const lastYear = parseFloat(row[1]) || 0;
            const budget = parseFloat(row[2]) || 0;
            const thisYear = parseFloat(row[3]) || 0;

            if (name.includes("收入總額")) totalIncome = thisYear;
            if (name.includes("支出總額")) totalExpense = thisYear;
            if (name.includes("本期餘絀")) {
                surplus = thisYear;
                this.crossCheck.incomeSurplus = surplus;
            }
            if (name.includes("折舊")) this.crossCheck.depreciationExp += thisYear;

            // Collect items for logic check
            // Simple heuristic to differentiate income/expense based on position or checking "收入"/"支出"
            // For this standard format, we verify Logic Rule 4 below row-by-row

            // 4. Single Expense Item check (> 20% of Total Expense)
            // Assuming this is an expense item if not "Total" and not "Surplus"
            // In a real generic parser this is hard, but we assume the standardizer did a decent job
            // We'll skip "Total" rows
            if (!name.includes("總額") && !name.includes("餘絀") && name.length > 0) {
                // Check if it's an expense item (heuristic: row index or keyword)
                // Here we simply check ratio against Total Expense if valuable
                if (totalExpense > 0 && thisYear > 0) {
                    const ratio = thisYear / totalExpense;
                    if (ratio > 0.2) {
                        // Check if likely expense (no "收入" in name)
                        if (!name.includes("收入")) {
                            this.log(category, "支出集中度", "WARNING",
                                `科目 [${name}] 金額佔總支出 ${(ratio * 100).toFixed(1)}%`,
                                "單一科目支出佔比過高 (>20%)",
                                "說明其性質及原因，確認是否合理");
                        }
                    }
                }
            }

            // 5. Budget Variance Check
            if (budget > 0) {
                const variance = Math.abs(thisYear - budget) / budget;
                if (variance > 0.2) {
                    this.log(category, "預算執行率", "WARNING",
                        `[${name}] 預決算差異 ${(variance * 100).toFixed(1)}%`,
                        "預算與決算差異超過 20%",
                        "請說明原因 (須提供預決算比較表)");
                }
            }

            // 6. YoY Check
            if (lastYear > 0) {
                const variance = Math.abs(thisYear - lastYear) / lastYear;
                if (variance > 0.1) {
                    this.log(category, "年度變動率", "WARNING",
                        `[${name}] 年度增減 ${(variance * 100).toFixed(1)}%`,
                        "兩年內金額增減幅度超過 10%",
                        "請說明其性質及原因");
                }
            }
        });

        // 1. Summation Check (Simplified: Check if Total Income/Expense rows exist)
        if (totalIncome === 0 && totalExpense === 0) {
            this.log(category, "合計正確性", "WARNING", "無法偵測到收入或支出總額", "可能影響其他比率計算", "請確認科目名稱包含 '收入總額'/'支出總額'");
        }

        // 2. Balance Check
        if (Math.abs(totalIncome - (totalExpense + surplus)) > 1) {
            this.log(category, "報表平衡", "ERROR",
                `不平衡: 收入(${totalIncome}) != 支出(${totalExpense}) + 餘絀(${surplus})`,
                "基本會計恆等式錯誤",
                "請檢查各項金額加總是否正確");
        } else {
            this.log(category, "報表平衡", "OK", "收支平衡正確", "基本會計恆等式正確", "-");
        }

        // 3. Surplus Ratio Check
        if (totalIncome > 0) {
            const ratio = surplus / totalIncome;
            if (ratio > 0.4) {
                this.log(category, "結餘合理性", "WARNING",
                    `本期餘絀佔收入 ${(ratio * 100).toFixed(1)}%`,
                    "結餘超過當年收入 40%",
                    "說明其結餘之使用計畫或轉列何種用途");
            }
        }

        // 7. Depreciation presence check (will be cross-checked in BS)
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
        let totalFundSurplus = 0; // Equity
        let bsSurplus = 0;
        let accumSurplusLast = 0; // accumulated surplus prev year

        rows.forEach(row => {
            const name = String(row[0] || "").trim();
            const lastYear = parseFloat(row[1]) || 0;
            const thisYear = parseFloat(row[2]) || 0;

            if (name.includes("資產總額")) totalAssets = thisYear;
            if (name.includes("負債總額")) totalLiabilities = thisYear;
            if (name.includes("基金") && name.includes("總額")) totalFundSurplus = thisYear;

            if (name.includes("本期餘絀")) {
                bsSurplus = thisYear;
                this.crossCheck.bsSurplus = bsSurplus;
            }

            if (name.includes("累積餘絀")) {
                accumSurplusLast = lastYear;
                this.crossCheck.prevAccumulatedSurplus = accumSurplusLast;
            }

            if (name.includes("固定資產") && !name.includes("總額")) {
                this.crossCheck.fixedAssets += thisYear;
            }

            // 6. Debt Composition Check
            if (["銀行借款", "暫收款", "暫付款", "短期借款", "應付票據"].some(k => name.includes(k))) {
                this.log(category, "債務結構", "INFO",
                    `檢測到科目: [${name}] 金額: ${thisYear}`,
                    "了解債務組成結構及比例",
                    "請確認債務性質");
            }
        });

        // 1. Summation & 2. Balance Check
        if (Math.abs(totalAssets - (totalLiabilities + totalFundSurplus)) > 1) {
            this.log(category, "報表平衡", "ERROR",
                `不平衡: 資產(${totalAssets}) != 負債(${totalLiabilities}) + 基金餘絀(${totalFundSurplus})`,
                "資產 != 負債 + 淨值",
                "請檢查加總");
        } else {
            this.log(category, "報表平衡", "OK", "平衡正確", "資產負債平衡", "-");
        }

        // 3. Cross Check: Income Surplus vs BS Surplus
        if (this.crossCheck.incomeSurplus !== null) {
            if (Math.abs(bsSurplus - this.crossCheck.incomeSurplus) > 1) {
                this.log(category, "跨表勾稽", "ERROR",
                    `資負表餘絀(${bsSurplus}) != 收支表餘絀(${this.crossCheck.incomeSurplus})`,
                    "本期餘絀金額不一致",
                    "請檢查兩表餘絀計算");
            } else {
                this.log(category, "跨表勾稽", "OK", "餘絀相符", "邏輯一致", "-");
            }
        }

        // 4. Accumulated Surplus Check (If applicable)
        // Heuristic: If BS Equity = Accum Surplus + This Year Surplus
        // This is complex to generalize without exact rows, but we can verify logic if row exists

        // 7. Debt Ratio
        if (totalAssets > 0) {
            const debtRatio = totalLiabilities / totalAssets;
            if (debtRatio > 0.5) {
                this.log(category, "財務結構", "WARNING",
                    `負債比率 ${(debtRatio * 100).toFixed(1)}%`,
                    "負債比超過總資產 50%",
                    "財務風險較高，請關注償債能力");
            }
        }

        // 7 (from IS). Fixed Assets Depreciation Check
        if (this.crossCheck.fixedAssets > 0) {
            if (this.crossCheck.depreciationExp === 0) {
                // Check if BS row "累積折舊" exists? standardizer might not extract separate col
                // We rely on IS depreciation expense
                this.log("收支決算表", "折舊提列", "WARNING",
                    `帳有固定資產(${this.crossCheck.fixedAssets}) 但未列折舊支出`,
                    "固定資產應提列折舊",
                    "請於收支決算表填寫折舊金額");
            } else {
                this.log("收支決算表", "折舊提列", "OK",
                    `已提列折舊: ${this.crossCheck.depreciationExp}`,
                    "符合會計原則",
                    "-");
            }
        }
    }

    _verifyFundStatement() {
        const category = "基金收支表";
        const rows = this._getRows(category);

        if (!rows) {
            // Not strictly error if association has no fund, but helpful info
            // User prompt: "(社團法人無提撥基金則不需提供)"
            return;
        }

        let fundMovements = false;
        let equityMovements = false;

        rows.forEach(row => {
            const name = String(row[0] || "").trim();
            const increase = parseFloat(row[2]) || 0;
            const decrease = parseFloat(row[3]) || 0;

            // 1. Fund movements
            if (name.includes("基金")) {
                if (increase > 0 || decrease > 0) {
                    fundMovements = true;
                    this.log(category, "基金動支", "INFO",
                        `[${name}] 有變動 (增:${increase}, 減:${decrease})`,
                        "基金是否有動支變動",
                        "請確認動支是否有核准程序");
                }
            }

            // 2. Net Value (Equity) movements (Example for Foundation)
            if (name.includes("淨值") || name.includes("累積餘絀")) {
                if (increase > 0 || decrease > 0) {
                    this.log(category, "淨值變動", "INFO",
                        `[${name}] 有變動 1:減少 2:增加`,
                        "淨值是否有特別變動 (不含本期餘絀)",
                        "確認變動原因");
                }
            }
        });
    }

    _verifyPropertyCatalog() {
        const category = "財產目錄";
        const rows = this._getRows(category);
        if (!rows) {
            if (this.crossCheck.fixedAssets > 0) {
                this.log(category, "目錄完整性", "WARNING", "無財產目錄", "帳有固定資產但無目錄", "請提供財產目錄");
            }
            return;
        }

        let catalogTotal = 0;
        let items = 0;
        rows.forEach(row => {
            // 4: Book Value
            catalogTotal += parseFloat(row[4]) || 0;
            items++;
        });

        // 5. Cross Check with BS Fixed Assets
        if (this.crossCheck.fixedAssets > 0) {
            if (Math.abs(this.crossCheck.fixedAssets - catalogTotal) > 1) {
                this.log(category, "帳實相符", "ERROR",
                    `財產目錄總額(${catalogTotal}) != 資負表固定資產(${this.crossCheck.fixedAssets})`,
                    "資產帳面價值不符",
                    "請檢查財產目錄或帳簿");
            } else {
                this.log(category, "帳實相符", "OK", "金額相符", "帳實相符", "-");
            }
        }
    }
}
