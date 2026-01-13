from flask import Flask, render_template, request, send_file
import os, zipfile
from pypdf import PdfWriter, PdfReader
from pdf2docx import Converter
from PIL import Image
import pandas as pd
from docx import Document
from pptx import Presentation

app = Flask(__name__)
UPLOAD_FOLDER = '/tmp'

@app.route('/')
def home():
    return render_template('index.html')

# --- PDF <-> WORD ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    f = request.files['file']
    path_in = os.path.join(UPLOAD_FOLDER, f.filename)
    path_out = path_in.replace('.pdf', '.docx')
    f.save(path_in)
    cv = Converter(path_in)
    cv.convert(path_out)
    cv.close()
    return send_file(path_out, as_attachment=True)

# --- EXCEL <-> WORD ---
@app.route('/excel-to-word', methods=['POST'])
def excel_to_word():
    f = request.files['file']
    df = pd.read_excel(f)
    path_out = os.path.join(UPLOAD_FOLDER, "export_excel.docx")
    doc = Document()
    t = doc.add_table(df.shape[0]+1, df.shape[1])
    doc.save(path_out)
    return send_file(path_out, as_attachment=True)

@app.route('/word-to-excel', methods=['POST'])
def word_to_excel():
    f = request.files['file']
    doc = Document(f)
    data = [[cell.text for cell in row.cells] for table in doc.tables for row in table.rows]
    path_out = os.path.join(UPLOAD_FOLDER, "export_word.xlsx")
    pd.DataFrame(data).to_excel(path_out, index=False)
    return send_file(path_out, as_attachment=True)

# --- ZIP (WINRAR STYLE) ---
@app.route('/zip', methods=['POST'])
def make_zip():
    files = request.files.getlist("files")
    path_zip = os.path.join(UPLOAD_FOLDER, "archive.zip")
    with zipfile.ZipFile(path_zip, 'w') as z:
        for f in files:
            f_path = os.path.join(UPLOAD_FOLDER, f.filename)
            f.save(f_path)
            z.write(f_path, f.filename)
    return send_file(path_zip, as_attachment=True)

# --- PDF TOOLS (FUSION/PROTECT) ---
@app.route('/merge', methods=['POST'])
def merge():
    files = request.files.getlist("pdfs")
    merger = PdfWriter()
    for f in files: merger.append(f)
    path = os.path.join(UPLOAD_FOLDER, "fusion.pdf")
    merger.write(path)
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
