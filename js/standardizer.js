export class ReportStandardizer {
    constructor() {
        // Keywords for identifying sheets and columns
    }

    normalizeString(s) {
        if (!s) return "";
        return String(s).trim().replace(" ", "").replace("\u3000", "");
    }

    standardize(inputWb) {
        // Create a new workbook
        const outputWb = XLSX.utils.book_new();

        this._processIncomeStatement(inputWb, outputWb);
        this._processBalanceSheet(inputWb, outputWb);
        this._processFundStatement(inputWb, outputWb); // New
        this._processPropertyCatalog(inputWb, outputWb);

        return outputWb;
    }

    _findSheetByKeyword(wb, keywords) {
        for (const sheetName of wb.SheetNames) {
            const normName = this.normalizeString(sheetName);
            if (keywords.some(k => normName.includes(k))) {
                return wb.Sheets[sheetName];
            }
        }
        return null;
    }

    _extractData(ws, requiredCols) {
        // Convert sheet to array of arrays
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        if (!data || data.length === 0) return [];

        let headerRowIdx = -1;
        let colMapping = {};

        // 1. Find Header Row (scan first 10 rows)
        for (let r = 0; r < Math.min(10, data.length); r++) {
            const rowValues = data[r];
            if (!rowValues) continue;

            let tempMapping = {};

            for (const [key, keywords] of Object.entries(requiredCols)) {
                for (let c = 0; c < rowValues.length; c++) {
                    const cellVal = this.normalizeString(rowValues[c]);
                    if (keywords.some(kw => cellVal.includes(kw))) {
                        tempMapping[key] = c;
                        break;
                    }
                }
            }

            // At least find 'item' and one amount
            if (Object.keys(tempMapping).length >= 2 && 'item' in tempMapping) {
                headerRowIdx = r;
                colMapping = tempMapping;
                break;
            }
        }

        if (headerRowIdx === -1) {
            console.warn("Could not identify headers");
            return [];
        }

        // 2. Extract Data
        const extractedData = [];
        for (let r = headerRowIdx + 1; r < data.length; r++) {
            const row = data[r];
            if (!row) continue;

            const rowData = {};
            let hasData = false;

            // Item Name
            if ('item' in colMapping) {
                const itemVal = row[colMapping['item']];
                if (itemVal) {
                    rowData['item'] = itemVal;
                    hasData = true;
                }
            }

            if (!hasData) continue;

            // Amounts
            for (const [key, colIdx] of Object.entries(colMapping)) {
                if (key === 'item') continue;
                let val = row[colIdx];

                // Clean number
                if (typeof val === 'number') {
                    rowData[key] = val;
                } else if (typeof val === 'string') {
                    // Remove commas
                    let cleanVal = val.replace(/,/g, "").trim();
                    // Handle (100) as -100
                    if (cleanVal.startsWith("(") && cleanVal.endsWith(")")) {
                        cleanVal = "-" + cleanVal.substring(1, cleanVal.length - 1);
                    }
                    const num = parseFloat(cleanVal);
                    rowData[key] = isNaN(num) ? 0 : num;
                } else {
                    rowData[key] = 0;
                }
            }
            extractedData.push(rowData);
        }
        return extractedData;
    }

    _processIncomeStatement(inputWb, outputWb) {
        const wsIn = this._findSheetByKeyword(inputWb, ["收支", "損益", "決算"]);
        const headers = ["科目名稱", "上年度決算數", "本年度預算數", "本年度決算數", "說明"];
        const outData = [headers];

        if (wsIn) {
            const data = this._extractData(wsIn, {
                "item": ["科目", "項目", "摘要", "名稱"],
                "this_year": ["本年", "本期", "決算"],
                "last_year": ["上年", "上期", "去年", "比較"],
                "budget": ["預算"]
            });

            data.forEach(row => {
                outData.push([
                    row.item || "",
                    row.last_year || 0,
                    row.budget || 0,
                    row.this_year || 0,
                    "系統轉檔"
                ]);
            });
        }

        const wsOut = XLSX.utils.aoa_to_sheet(outData);
        XLSX.utils.book_append_sheet(outputWb, wsOut, "收支決算表");
    }

    _processBalanceSheet(inputWb, outputWb) {
        const wsIn = this._findSheetByKeyword(inputWb, ["資產", "負債", "平衡"]);
        const headers = ["科目名稱", "上年度金額", "本年度金額", "說明"];
        const outData = [headers];

        if (wsIn) {
            const data = this._extractData(wsIn, {
                "item": ["科目", "項目", "摘要", "名稱"],
                "this_year": ["本年", "本期", "期末", "決算"],
                "last_year": ["上年", "上期", "期初", "去年"]
            });

            data.forEach(row => {
                outData.push([
                    row.item || "",
                    row.last_year || 0,
                    row.this_year || 0,
                    "系統轉檔"
                ]);
            });
        }

        const wsOut = XLSX.utils.aoa_to_sheet(outData);
        XLSX.utils.book_append_sheet(outputWb, wsOut, "資產負債表");
    }

    _processFundStatement(inputWb, outputWb) {
        // Keywords: 基金, 淨值, 變動
        const wsIn = this._findSheetByKeyword(inputWb, ["基金", "淨值", "變動"]);
        const headers = ["科目名稱", "期初餘額", "本期增加", "本期減少", "期末餘額"];
        const outData = [headers];

        if (wsIn) {
            const data = this._extractData(wsIn, {
                "item": ["科目", "項目", "名稱", "種類"],
                "start_balance": ["期初", "上期"],
                "increase": ["增加", "撥入"],
                "decrease": ["減少", "動支"],
                "end_balance": ["期末", "本期"]
            });

            data.forEach(row => {
                outData.push([
                    row.item || "",
                    row.start_balance || 0,
                    row.increase || 0,
                    row.decrease || 0,
                    row.end_balance || 0
                ]);
            });
        }

        const wsOut = XLSX.utils.aoa_to_sheet(outData);
        XLSX.utils.book_append_sheet(outputWb, wsOut, "基金收支表");
    }

    _processPropertyCatalog(inputWb, outputWb) {
        // Create empty placeholder if not found for now, or simple copy in future
        const headers = ["資產名稱", "取得日期", "取得成本", "本期折舊", "帳面價值"];
        const wsOut = XLSX.utils.aoa_to_sheet([headers]);
        XLSX.utils.book_append_sheet(outputWb, wsOut, "財產目錄");
    }
}
