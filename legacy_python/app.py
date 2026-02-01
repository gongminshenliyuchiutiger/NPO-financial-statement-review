from flask import Flask, render_template, request, send_file
import openpyxl
import os
import mimetypes
from verifier_logic import FinancialReportVerifier
from standardizer import ReportStandardizer
from gemini_processor import GeminiProcessor

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

TEMPLATE_PATH = "Full_Financial_Report_Template.xlsx"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/verify', methods=['POST'])
def verify():
    if 'file' not in request.files:
        return {"error": "No file uploaded"}, 400
    
    file = request.files['file']
    use_ai = request.form.get('use_ai') == 'true'
    api_key = request.form.get('api_key', '').strip()
    
    if file.filename == '':
        return {"error": "No file selected"}, 400

    if file:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        
        try:
            wb = None
            
            # Scenario 1: User provided API Key -> Use Gemini (Strong AI)
            if api_key:
                try:
                    mime_type, _ = mimetypes.guess_type(filepath)
                    if not mime_type: mime_type = 'application/octet-stream'
                    
                    processor = GeminiProcessor(api_key)
                    wb = processor.process_file(filepath, mime_type)
                    # Helper: Save intermediate for debugging
                    wb.save(os.path.join(app.config['UPLOAD_FOLDER'], "gemini_converted_" + file.filename + ".xlsx"))
                except Exception as e:
                     return {"error": f"Gemini API Error: {str(e)}"}, 500

            # Scenario 2: No API Key, but use_ai=true (Excel only) -> Use Local Standardizer (Weak AI)
            elif use_ai and (file.filename.endswith('.xlsx') or file.filename.endswith('.xlsm')):
                wb = openpyxl.load_workbook(filepath, data_only=True)
                standardizer = ReportStandardizer()
                wb = standardizer.standardize(wb)
                
            # Scenario 3: Standard Excel File
            elif file.filename.endswith('.xlsx') or file.filename.endswith('.xlsm'):
                wb = openpyxl.load_workbook(filepath, data_only=True)
            else:
                return {"error": "如果不使用 Gemini API，僅支援 .xlsx/.xlsm 檔案。請輸入 API Key 以支援圖片/PDF。"}, 400
            
            # Verify the resulting workbook
            verifier = FinancialReportVerifier(wb)
            results = verifier.run_verify()
            
            if 'wb' in locals() and wb: wb.close()
            return {"results": results}
            
        except Exception as e:
            return {"error": str(e)}, 500

@app.route('/download_template')
def download_template():
    if os.path.exists(TEMPLATE_PATH):
        return send_file(TEMPLATE_PATH, as_attachment=True)
    else:
        return "Template not found. Please run generate_full_template.py first.", 404

if __name__ == '__main__':
    app.run(debug=True, port=5000)
