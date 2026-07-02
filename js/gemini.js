import { GoogleGenerativeAI } from "@google/generative-ai";

export class GeminiProcessor {
    constructor(apiKey, modelName = "gemini-2.5-flash") {
        if (!apiKey) throw new Error("API Key is required");
        // Remove any non-ASCII characters that might cause Header encoding errors
        const sanitizedKey = apiKey.replace(/[^\x21-\x7E]/g, "").trim();
        this.genAI = new GoogleGenerativeAI(sanitizedKey);
        // Use the selected model (default: gemini-2.5-flash)
        this.model = this.genAI.getGenerativeModel({ model: modelName });
    }

    async processFile(file) {
        // 1. Convert file to base64
        const base64Data = await this.fileToGenerativePart(file);

        // 2. Construct Prompt
        const prompt = `
        你是專業的非營利組織財務會計師。請協助將上傳的財務報表文件（可能是圖片、PDF或Excel），轉換為標準的 JSON 格式。
        
        請提取以下三張表的數據：
        1. 收支決算表 (Income Statement): 包含科目名稱(item)、上年度決算數(last_year)、本年度預算數(budget)、本年度決算數(this_year)
        2. 資產負債表 (Balance Sheet): 包含科目名稱(item)、上年度金額(last_year)、本年度金額(this_year)
        3. 財產目錄 (Property Catalog): 包含資產名稱(item_name)、取得日期(date)、取得成本(cost)、本期折舊(depreciation)、帳面價值(book_value)

        請務必返回純 JSON 格式，不要有 Markdown 標記（如 \`\`\`json ... \`\`\`）。結構如下：
        {
            "income_statement": [
                {"item": "收入總額", "last_year": 1000, "budget": 1200, "this_year": 1100},
                ...
            ],
            "balance_sheet": [
                {"item": "資產總額", "last_year": 5000, "this_year": 5500},
                ...
            ],
            "property_catalog": [
                {"item_name": "電腦", "date": "2023-01-01", "cost": 30000, "depreciation": 10000, "book_value": 20000},
                ...
            ]
        }
        
        如果文件中缺某些欄位（例如沒有預算數），請填 0 或 null。請自動識別並對應到最接近的標準科目名稱。
        `;

        // 3. Generate Content
        try {
            const result = await this.model.generateContent([prompt, base64Data]);
            const response = await result.response;
            const text = response.text();

            return this._parseResponse(text);
        } catch (error) {
            console.error("Gemini Error:", error);
            throw new Error(`Gemini Processing Failed: ${error.message}`);
        }
    }

    async fileToGenerativePart(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => {
                const base64String = reader.result.split(',')[1];
                resolve({
                    inlineData: {
                        data: base64String,
                        mimeType: file.type
                    }
                });
            };
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    }

    _parseResponse(text) {
        try {
            // Strip markdown
            let cleanText = text.trim();
            if (cleanText.startsWith("```json")) cleanText = cleanText.slice(7);
            if (cleanText.startsWith("```")) cleanText = cleanText.slice(3);
            if (cleanText.endsWith("```")) cleanText = cleanText.slice(0, -3);

            const data = JSON.parse(cleanText);
            return this._jsonToWorkbook(data);
        } catch (e) {
            console.error("JSON Parse Error:", e);
            throw new Error("Failed to parse Gemini response as JSON");
        }
    }

    _jsonToWorkbook(data) {
        const wb = XLSX.utils.book_new();

        // 1. Income Statement
        if (data.income_statement) {
            const wsData = [["科目名稱", "上年度決算數", "本年度預算數", "本年度決算數", "說明"]];
            data.income_statement.forEach(row => {
                wsData.push([
                    row.item || "",
                    row.last_year || 0,
                    row.budget || 0,
                    row.this_year || 0,
                    "Gemini擷取"
                ]);
            });
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            XLSX.utils.book_append_sheet(wb, ws, "收支決算表");
        }

        // 2. Balance Sheet
        if (data.balance_sheet) {
            const wsData = [["科目名稱", "上年度金額", "本年度金額", "說明"]];
            data.balance_sheet.forEach(row => {
                wsData.push([
                    row.item || "",
                    row.last_year || 0,
                    row.this_year || 0,
                    "Gemini擷取"
                ]);
            });
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            XLSX.utils.book_append_sheet(wb, ws, "資產負債表");
        }

        // 3. Property Catalog
        if (data.property_catalog) {
            const wsData = [["資產名稱", "取得日期", "取得成本", "本期折舊", "帳面價值"]];
            data.property_catalog.forEach(row => {
                wsData.push([
                    row.item_name || "",
                    row.date || "",
                    row.cost || 0,
                    row.depreciation || 0,
                    row.book_value || 0
                ]);
            });
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            XLSX.utils.book_append_sheet(wb, ws, "財產目錄");
        }

        return wb;
    }
}
