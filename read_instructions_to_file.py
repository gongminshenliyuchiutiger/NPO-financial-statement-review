import docx
import sys

try:
    doc = docx.Document('EX指令.docx')
    with open('instructions_utf8.txt', 'w', encoding='utf-8') as f:
        f.write("Content of EX指令.docx:\n")
        for para in doc.paragraphs:
            if para.text.strip():
                f.write(para.text + "\n")
    print("Successfully wrote instructions to instructions_utf8.txt")
except Exception as e:
    print(f"Error reading file: {e}")
