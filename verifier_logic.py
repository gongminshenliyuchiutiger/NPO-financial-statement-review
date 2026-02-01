import openpyxl

class FinancialReportVerifier:
    def __init__(self, workbook):
        self.wb = workbook
        self.results = [] # List of {"category": str, "rule": str, "status": "OK"|"WARNING"|"ERROR", "message": str}

    def log(self, category, rule, status, message):
        self.results.append({
            "category": category,
            "rule": rule,
            "status": status,
            "message": message
        })

    def run_verify(self):
        self.results = []
        self._verify_income_statement()
        self._verify_balance_sheet()
        self._verify_property_catalog()
        return self.results

    def _get_rows(self, sheet_name):
        if sheet_name not in self.wb.sheetnames:
            return None
        ws = self.wb[sheet_name]
        return list(ws.iter_rows(min_row=2, values_only=True))

    def _verify_income_statement(self):
        category = "收支決算表"
        rows = self._get_rows("收支決算表")
        
        income_total = 0
        expense_total = 0
        surplus = 0
        self.depreciation = 0 # Store for cross-check

        if not rows:
            self.log(category, "工作表檢查", "ERROR", "找不到 '收支決算表' 工作表")
            return

        for row in rows:
            name = str(row[0]).strip() if row[0] else ""
            val_final = row[3] if isinstance(row[3], (int, float)) else 0
            val_budget = row[2] if isinstance(row[2], (int, float)) else 0
            val_last = row[1] if isinstance(row[1], (int, float)) else 0
            
            if "收入總額" in name: income_total = val_final
            if "支出總額" in name: expense_total = val_final
            if "本期餘絀" in name: surplus = val_final
            if "折舊" in name: self.depreciation += val_final

            # Rules
            if val_budget > 0 and abs(val_final - val_budget) / val_budget > 0.2:
                self.log(category, "預算執行率", "WARNING", f"[{name}] 決算({val_final:,}) 與預算({val_budget:,}) 差異超過 20%")
            
            if val_last > 0 and abs(val_final - val_last) / val_last > 0.1:
                self.log(category, "年度變動率", "WARNING", f"[{name}] 本年({val_final:,}) 與上年({val_last:,}) 變動超過 10%")

        # Balance
        if abs(income_total - (expense_total + surplus)) > 1:
            self.log(category, "報表平衡", "ERROR", f"收支不平衡: 收入({income_total:,}) != 支出({expense_total:,}) + 餘絀({surplus:,})")
        else:
            self.log(category, "報表平衡", "OK", "收支平衡正確")

        # Surplus Ratio
        if income_total > 0 and (surplus / income_total) > 0.4:
            self.log(category, "結餘比率", "WARNING", f"本期餘絀佔收入 {surplus/income_total:.1%} (標準: <40%) -> 請說明結餘用途")
        else:
            self.log(category, "結餘比率", "OK", "結餘比率正常")
        
        # Store for cross check
        self.income_surplus = surplus

    def _verify_balance_sheet(self):
        category = "資產負債表"
        rows = self._get_rows("資產負債表")
        
        asset_total = 0
        liability_total = 0
        fund_surplus_total = 0
        bs_surplus = 0
        self.fixed_assets = 0

        if not rows:
            self.log(category, "工作表檢查", "ERROR", "找不到 '資產負債表' 工作表")
            return

        for row in rows:
            name = str(row[0]).strip() if row[0] else ""
            val = row[2] if isinstance(row[2], (int, float)) else 0
            
            if "資產總額" in name: asset_total = val
            if "負債總額" in name: liability_total = val
            if "基金及餘絀總額" in name: fund_surplus_total = val
            if "本期餘絀" in name: bs_surplus = val
            if "固定資產" in name and "總額" not in name: self.fixed_assets = val

        # Balance
        if abs(asset_total - (liability_total + fund_surplus_total)) > 1:
            self.log(category, "報表平衡", "ERROR", f"資產負債不平衡: 資產({asset_total:,}) != 負債({liability_total:,}) + 基金餘絀({fund_surplus_total:,})")
        else:
            self.log(category, "報表平衡", "OK", "資產負債平衡正確")

        # Cross Check
        if hasattr(self, 'income_surplus'):
            if abs(bs_surplus - self.income_surplus) > 1:
                 self.log(category, "跨表勾稽", "ERROR", f"資產負債表餘絀({bs_surplus:,}) 與 收支決算表餘絀({self.income_surplus:,}) 不符")
            else:
                 self.log(category, "跨表勾稽", "OK", "餘絀金額一致")

        # Debt Ratio
        if asset_total > 0 and (liability_total / asset_total) > 0.5:
             self.log(category, "財務結構", "WARNING", f"負債比率 {liability_total/asset_total:.1%} (標準: <50%) -> 請關注償債能力")
        else:
             self.log(category, "財務結構", "OK", "負債比率正常")

        # Depreciation Cross Check
        if self.fixed_assets > 0 and getattr(self, 'depreciation', 0) == 0:
            self.log(category, "折舊查核", "WARNING", f"帳上有固定資產({self.fixed_assets:,})，但收支表無折舊費用")

    def _verify_property_catalog(self):
        category = "財產目錄"
        rows = self._get_rows("財產目錄")
        
        if not rows:
            self.log(category, "資料檢查", "WARNING", "無財產目錄資料")
            return

        prop_total_value = sum(row[4] for row in rows if isinstance(row[4], (int, float)))
        
        if hasattr(self, 'fixed_assets'):
            if abs(self.fixed_assets - prop_total_value) > 1:
                self.log(category, "目錄勾稽", "ERROR", f"固定資產帳面價值({self.fixed_assets:,}) 與 財產目錄加總({prop_total_value:,}) 不符")
            else:
                self.log(category, "目錄勾稽", "OK", "金額相符")
