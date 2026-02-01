import docx
import sys

try:
    doc = docx.Document('EX指令.docx')
    print("Content of EX指令.docx:")
    for para in doc.paragraphs:
        if para.text.strip():
            print(para.text)
except Exception as e:
    print(f"Error reading file: {e}")
