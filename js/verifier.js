export class FinancialReportVerifier {
    constructor(workbook) {
        this.wb = workbook;
        this.results = [];
        // Cross-check values
        this.depreciation = 0;
        this.incomeSurplus = null;
        this.fixedAssets = 0;
    }

    log(category, rule, status, message) {
        this.results.push({
            category,
            rule,
            status,
            message
        });
    }

    runVerify() {
        this.results = [];
        this._verifyIncomeStatement();
        this._verifyBalanceSheet();
        this._verifyPropertyCatalog();
        return this.results;
    }

    _getRows(sheetName) {
        // Check if sheet exists
        if (!this.wb.Sheets[sheetName]) return null;

        // Get data skipping header (assuming standard format from Standardizer)
        // Standard format: [Item, LastYear, Budget, ThisYear, Note] etc.
        const data = XLSX.utils.sheet_to_json(this.wb.Sheets[sheetName], { header: 1 });
        if (data.length < 2) return []; // Only header or empty
        return data.slice(1); // Skip header row
    }

    _verifyIncomeStatement() {
        const category = "收支決算表";
        const rows = this._getRows("收支決算表");

        let incomeTotal = 0;
        let expenseTotal = 0;
        let surplus = 0;
        this.depreciation = 0;

        if (!rows) {
            this.log(category, "工作表檢查", "ERROR", "找不到 '收支決算表' 工作表");
            return;
        }

        rows.forEach(row => {
            // Standard format: 0:Item, 1:LastYear, 2:Budget, 3:ThisYear
            const name = String(row[0] || "").trim();
            const valLast = parseFloat(row[1]) || 0;
            const valBudget = parseFloat(row[2]) || 0;
            const valFinal = parseFloat(row[3]) || 0;

            if (name.includes("收入總額")) incomeTotal = valFinal;
            if (name.includes("支出總額")) expenseTotal = valFinal;
            if (name.includes("本期餘絀")) surplus = valFinal;
            if (name.includes("折舊")) this.depreciation += valFinal;

            // Rules
            if (valBudget > 0 && Math.abs(valFinal - valBudget) / valBudget > 0.2) {
                this.log(category, "預算執行率", "WARNING", `[${name}] 決算(${valFinal.toLocaleString()}) 與預算(${valBudget.toLocaleString()}) 差異超過 20%`);
            }

            if (valLast > 0 && Math.abs(valFinal - valLast) / valLast > 0.1) {
                this.log(category, "年度變動率", "WARNING", `[${name}] 本年(${valFinal.toLocaleString()}) 與上年(${valLast.toLocaleString()}) 變動超過 10%`);
            }
        });

        // Balance
        if (Math.abs(incomeTotal - (expenseTotal + surplus)) > 1) {
            this.log(category, "報表平衡", "ERROR", `收支不平衡: 收入(${incomeTotal.toLocaleString()}) != 支出(${expenseTotal.toLocaleString()}) + 餘絀(${surplus.toLocaleString()})`);
        } else {
            this.log(category, "報表平衡", "OK", "收支平衡正確");
        }

        // Surplus Ratio
        if (incomeTotal > 0 && (surplus / incomeTotal) > 0.4) {
            this.log(category, "結餘比率", "WARNING", `本期餘絀佔收入 ${(surplus / incomeTotal * 100).toFixed(1)}% (標準: <40%) -> 請說明結餘用途`);
        } else {
            this.log(category, "結餘比率", "OK", "結餘比率正常");
        }

        this.incomeSurplus = surplus;
    }

    _verifyBalanceSheet() {
        const category = "資產負債表";
        const rows = this._getRows("資產負債表");

        let assetTotal = 0;
        let liabilityTotal = 0;
        let fundSurplusTotal = 0;
        let bsSurplus = 0;
        this.fixedAssets = 0;

        if (!rows) {
            this.log(category, "工作表檢查", "ERROR", "找不到 '資產負債表' 工作表");
            return;
        }

        rows.forEach(row => {
            // Standard format: 0:Item, 1:LastYear, 2:ThisYear
            const name = String(row[0] || "").trim();
            const valLast = parseFloat(row[1]) || 0;
            const val = parseFloat(row[2]) || 0;

            if (name.includes("資產總額")) assetTotal = val;
            if (name.includes("負債總額")) liabilityTotal = val;
            if (name.includes("基金及餘絀總額") || name.includes("基金暨餘絀總額")) fundSurplusTotal = val;
            if (name.includes("本期餘絀")) bsSurplus = val;
            if (name.includes("固定資產") && !name.includes("總額")) this.fixedAssets = val;
        });

        // Balance
        if (Math.abs(assetTotal - (liabilityTotal + fundSurplusTotal)) > 1) {
            this.log(category, "報表平衡", "ERROR", `資產負債不平衡: 資產(${assetTotal.toLocaleString()}) != 負債(${liabilityTotal.toLocaleString()}) + 基金餘絀(${fundSurplusTotal.toLocaleString()})`);
        } else {
            this.log(category, "報表平衡", "OK", "資產負債平衡正確");
        }

        // Cross Check
        if (this.incomeSurplus !== null) {
            if (Math.abs(bsSurplus - this.incomeSurplus) > 1) {
                this.log(category, "跨表勾稽", "ERROR", `資產負債表餘絀(${bsSurplus.toLocaleString()}) 與 收支決算表餘絀(${this.incomeSurplus.toLocaleString()}) 不符`);
            } else {
                this.log(category, "跨表勾稽", "OK", "餘絀金額一致");
            }
        }

        // Debt Ratio
        if (assetTotal > 0 && (liabilityTotal / assetTotal) > 0.5) {
            this.log(category, "財務結構", "WARNING", `負債比率 ${(liabilityTotal / assetTotal * 100).toFixed(1)}% (標準: <50%) -> 請關注償債能力`);
        } else {
            this.log(category, "財務結構", "OK", "負債比率正常");
        }

        // Depreciation check
        if (this.fixedAssets > 0 && this.depreciation === 0) {
            this.log(category, "折舊查核", "WARNING", `帳上有固定資產(${this.fixedAssets.toLocaleString()})，但收支表無折舊費用`);
        }
    }

    _verifyPropertyCatalog() {
        const category = "財產目錄";
        const rows = this._getRows("財產目錄");

        if (!rows) {
            this.log(category, "資料檢查", "WARNING", "無財產目錄資料");
            return;
        }

        // Standard format: 0:Name, 1:Date, 2:Cost, 3:Depreciation, 4:BookValue
        let propTotalValue = 0;
        rows.forEach(row => {
            propTotalValue += parseFloat(row[4]) || 0;
        });

        if (this.fixedAssets > 0) {
            if (Math.abs(this.fixedAssets - propTotalValue) > 1) {
                this.log(category, "目錄勾稽", "ERROR", `固定資產帳面價值(${this.fixedAssets.toLocaleString()}) 與 財產目錄加總(${propTotalValue.toLocaleString()}) 不符`);
            } else {
                this.log(category, "目錄勾稽", "OK", "金額相符");
            }
        }
    }
}
