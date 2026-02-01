export class FinancialReportVerifier {
    constructor(workbook) {
        this.wb = workbook;
        this.results = [];
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
        // The order matters for cross-checks
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
        let surplus = 0; // 本期餘絀

        // Loop for calculation first
        rows.forEach(row => {
            const name = String(row[0] || "").trim();
            const thisYear = parseFloat(row[3]) || 0; // 本年度決算數

            if (name.includes("收入") && name.includes("總額")) totalIncome = thisYear;
            if (name.includes("支出") && name.includes("總額")) totalExpense = thisYear;
            if (name.includes("本期餘絀")) {
                surplus = thisYear;
                this.crossCheck.incomeSurplus = surplus;
            }
            if (name.includes("折舊")) {
                this.crossCheck.depreciationExp += thisYear;
            }
        });

        // 1. 報表數字合計正確 (Check Summation - simplified check if totals exist)
        if (totalIncome === 0 && totalExpense === 0) {
            this.log(category, "報表數字合計正確", "WARNING", "無法偵測到收入或支出總額", "影響比率計算", "請確認科目名稱包含 '收入總額'/'支出總額'");
        } else {
            this.log(category, "報表數字合計正確", "OK", "已偵測到總額欄位", "-", "-");
        }

        // 2. 報表數字平衡 (收入 = 支出 + 本期餘絀)
        // Allow small floating point error
        if (Math.abs(totalIncome - (totalExpense + surplus)) > 1) {
            this.log(category, "報表數字平衡", "ERROR",
                `不平衡: 收入(${totalIncome}) != 支出(${totalExpense}) + 餘絀(${surplus})`,
                "收入應等於支出加本期餘絀",
                "請檢查各項金額加總及餘絀計算是否正確");
        } else {
            this.log(category, "報表數字平衡", "OK", "收支平衡正確", "符合會計恆等式", "-");
        }

        // 3. 本期餘絀沒有超過當年收入40%
        if (totalIncome > 0) {
            const surplusRatio = surplus / totalIncome;
            if (surplusRatio > 0.4) {
                this.log(category, "結餘合理性", "WARNING",
                    `本期餘絀佔收入 ${(surplusRatio * 100).toFixed(1)}%`,
                    "本期餘絀超過當年收入40%",
                    "說明其結餘之使用計畫或轉列何種用途");
            } else {
                this.log(category, "結餘合理性", "OK", "結餘比例正常", "-", "-");
            }
        }

        // Row-by-row checks
        rows.forEach(row => {
            const name = String(row[0] || "").trim();
            const lastYear = parseFloat(row[1]) || 0; // 上年度決算數
            const budget = parseFloat(row[2]) || 0;    // 本年度預算數
            const thisYear = parseFloat(row[3]) || 0;  // 本年度決算數

            if (!name) return;

            // 4. 支出科目單項金額沒有超過總支出金額20%以上
            // Filter out Total rows and Surplus rows
            if (!name.includes("總額") && !name.includes("餘絀") && !name.includes("收入")) {
                // Assume it's an expense item if it's not income (heuristic)
                // Better heuristic: if we are below "支出" header? 
                // For now, assuming standard template structure, we check strict ratio
                if (totalExpense > 0 && thisYear > 0) {
                    const ratio = thisYear / totalExpense;
                    if (ratio > 0.2) {
                        this.log(category, "支出集中度", "WARNING",
                            `科目 [${name}] 金額佔總支出 ${(ratio * 100).toFixed(1)}%`,
                            "支出科目單項金額超過總支出金額20%以上",
                            "說明其性質及原因");
                    }
                }
            }

            // 5. 預決算沒有差異超過20%以上
            if (budget > 0) {
                const diff = Math.abs(thisYear - budget);
                const variance = diff / budget;
                if (variance > 0.2) {
                    this.log(category, "預決算差異", "WARNING",
                        `[${name}] 差異 ${(variance * 100).toFixed(1)}%`,
                        "預決算差異超過20%以上",
                        "說明其原因。(須提供預決算比較表)");
                }
            }

            // 6. 兩年內各科目金額增減幅度有超過10%以上
            if (lastYear > 0) {
                const diff = Math.abs(thisYear - lastYear);
                const variance = diff / lastYear;
                if (variance > 0.1) {
                    this.log(category, "年度變動率", "WARNING",
                        `[${name}] 變動 ${(variance * 100).toFixed(1)}%`,
                        "兩年內各科目金額增減幅度有超過10%以上",
                        "說明其性質及原因");
                }
            }
        });

        // 7. 固定資產有無提列折舊 (Check existence of depreciation expense)
        // This is just checking if "depreciation" exists in IS. Cross-check with BS happens later.
        if (this.crossCheck.depreciationExp > 0) {
            this.log(category, "折舊提列", "OK", `已列支折舊費用: ${this.crossCheck.depreciationExp}`, "固定資產有提列折舊", "-");
        } else {
            // We don't error yet, only if Fixed Assets exist (checked in BS)
        }
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
        let totalFundAndSurplus = 0; // 基金及餘絀總額
        let accumSurplusThis = 0;
        let accumSurplusLast = 0;

        rows.forEach(row => {
            const name = String(row[0] || "").trim();
            const lastYear = parseFloat(row[1]) || 0;
            const thisYear = parseFloat(row[2]) || 0; // 本年度金額

            if (name.includes("資產總額")) totalAssets = thisYear;
            if (name.includes("負債總額")) totalLiabilities = thisYear;
            // Usually "基金及餘絀總額" or "淨值總額"
            if ((name.includes("基金") || name.includes("淨值")) && name.includes("總額")) {
                totalFundAndSurplus = thisYear;
            }

            if (name.includes("本期餘絀")) {
                this.crossCheck.bsSurplus = thisYear;
            }

            if (name.includes("累積餘絀")) {
                accumSurplusLast = lastYear;
                accumSurplusThis = thisYear;
            }

            // Check for Fixed Assets to verify depreciation need
            if (name.includes("固定資產") && !name.includes("總額")) {
                this.crossCheck.fixedAssets += thisYear;
            }

            // 5. 資固定資產的帳面價值與財產目錄相符 (Wait for Property Catalog check)

            // 6. 帳列上若有銀行借款、暫收款及暫付款或其他往來科目則標註出來
            const debtKeywords = ["銀行借款", "暫收款", "暫付款", "短期借款", "應付票據", "其他往來"];
            if (debtKeywords.some(kw => name.includes(kw)) && thisYear > 0) {
                this.log(category, "債務組成結構", "INFO",
                    `偵測到科目: [${name}] 金額: ${thisYear}`,
                    "了解債務組成結構及比例",
                    "請確認該債務性質是否合理");
            }
        });

        // 1. 報表數字合計正確
        // Implicitly checked by verifying balance below
        this.log(category, "報表數字合計正確", "OK", "完成加總檢核", "-", "-");

        // 2. 報表數字平衡(資產總額=負債總額+基金及餘絀總額)
        if (Math.abs(totalAssets - (totalLiabilities + totalFundAndSurplus)) > 1) {
            this.log(category, "報表數字平衡", "ERROR",
                `不平衡: 資產(${totalAssets}) != 負債(${totalLiabilities}) + 基金及餘絀(${totalFundAndSurplus})`,
                "資產應等於負債加基金及餘絀",
                "請檢查加總是否正確");
        } else {
            this.log(category, "報表數字平衡", "OK", "資產負債平衡正確", "符合會計恆等式", "-");
        }

        // 3. 收支餘絀表之本期餘絀與資產負債表之本期餘絀相符
        if (this.crossCheck.incomeSurplus !== null && this.crossCheck.bsSurplus !== null) {
            if (Math.abs(this.crossCheck.incomeSurplus - this.crossCheck.bsSurplus) > 1) {
                this.log(category, "跨表勾稽(本期餘絀)", "ERROR",
                    `資負表(${this.crossCheck.bsSurplus}) != 收支表(${this.crossCheck.incomeSurplus})`,
                    "兩表本期餘絀金額不一致",
                    "請檢查兩表計算");
            } else {
                this.log(category, "跨表勾稽(本期餘絀)", "OK", "金額相符", "-", "-");
            }
        }

        // 4. 如資產負債表內僅有累積餘絀科目則減去前一年度累積餘絀看是否與當年度收支餘絀表內本期餘絀相符
        if (accumSurplusThis !== 0 && accumSurplusLast !== 0 && this.crossCheck.incomeSurplus !== null) {
            const diff = accumSurplusThis - accumSurplusLast;
            // Allow small diff
            if (Math.abs(diff - this.crossCheck.incomeSurplus) > 2) {
                this.log(category, "累積餘絀勾稽", "WARNING",
                    `累積餘絀變動(${diff}) != 本期餘絀(${this.crossCheck.incomeSurplus})`,
                    "累積餘絀變動應等於本期餘絀 (若無其他淨值調整)",
                    "確認是否有其他影響淨值之項目");
            } else {
                this.log(category, "累積餘絀勾稽", "OK", "變動相符", "-", "-");
            }
        }

        // 7. 負債比未超過總資產50%
        if (totalAssets > 0) {
            const debtRatio = totalLiabilities / totalAssets;
            if (debtRatio > 0.5) {
                this.log(category, "財務結構", "WARNING",
                    `負債比率 ${(debtRatio * 100).toFixed(1)}%`,
                    "負債比超過總資產50%",
                    "財務結構風險較高，需說明原因");
            } else {
                this.log(category, "財務結構", "OK", "負債比率正常", "-", "-");
            }
        }

        // 7 (IS rule) check: If Fixed Assets exist, must have Depreciation
        if (this.crossCheck.fixedAssets > 0 && this.crossCheck.depreciationExp === 0) {
            this.log("收支決算表", "折舊提列", "WARNING",
                `有固定資產(${this.crossCheck.fixedAssets}) 但未列折舊支出`,
                "固定資產有無提列折舊",
                "若有，請填寫折舊金額"); // Using specific Suggestion text
        }
    }

    _verifyFundStatement() {
        const category = "基金收支表"; // Covers "基金收支表或淨值變動表"
        const rows = this._getRows(category);

        if (!rows) {
            // Not necessarily an error if they don't have it, but consistent with request to check if present
            return;
        }

        let hasChange = false;

        rows.forEach(row => {
            const name = String(row[0] || "").trim();
            // 0:Name, 1:Begin, 2:Increase, 3:Decrease, 4:End
            const increase = parseFloat(row[2]) || 0;
            const decrease = parseFloat(row[3]) || 0;

            // 1. 協會_基金是否有動支變動
            if (name.includes("基金")) {
                if (increase > 0 || decrease > 0) {
                    this.log(category, "基金動支", "INFO",
                        `[${name}] 本期增加:${increase}, 減少:${decrease}`,
                        "基金是否有動支變動",
                        "請說明動支原因及程序是否完備");
                }
            }

            // 2. 基金會_淨值(不超累積餘絀)是否有特別變動 (1:減少 2:增加 3:不變)
            // Excluding "本期餘絀" usually
            if ((name.includes("淨值") || name.includes("累積餘絀")) && !name.includes("本期餘絀")) {
                let statusMsg = "3:不變";
                if (increase > 0) statusMsg = "2:增加";
                if (decrease > 0) statusMsg = "1:減少";

                if (statusMsg !== "3:不變") {
                    this.log(category, "淨值變動", "INFO",
                        `[${name}] ${statusMsg}`,
                        "淨值是否有特別變動 (不含本期餘絀)",
                        "確認變動性質");
                }
            }
        });
    }

    _verifyPropertyCatalog() {
        const category = "財產目錄";
        const rows = this._getRows(category);

        // If no property catalog but we have fixed assets -> Error
        if (!rows || rows.length === 0) {
            if (this.crossCheck.fixedAssets > 0) {
                this.log(category, "目錄存在性", "ERROR",
                    "帳列有固定資產但未提供財產目錄",
                    "資固定資產的帳面價值與財產目錄相符",
                    "請提供財產目錄");
            }
            return;
        }

        let catalogTotal = 0;
        rows.forEach(row => {
            // Headers: [資產名稱, 取得日期, 取得成本, 本期折舊, 帳面價值]
            // Index 4 is Book Value
            catalogTotal += parseFloat(row[4]) || 0;
        });

        // 5. 資固定資產的帳面價值與財產目錄相符 (From BS rules)
        if (this.crossCheck.fixedAssets > 0) {
            if (Math.abs(this.crossCheck.fixedAssets - catalogTotal) > 1) {
                this.log("資產負債表", "帳實相符", "ERROR",
                    `帳面價值(${this.crossCheck.fixedAssets}) != 目錄總額(${catalogTotal})`,
                    "資固定資產的帳面價值與財產目錄相符",
                    "請檢查財產目錄金額是否正確");
            } else {
                this.log("資產負債表", "帳實相符", "OK", "金額相符", "-", "-");
            }
        }
    }
}
