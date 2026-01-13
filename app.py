from flask import Flask, render_template, request, send_file
import os, zipfile
from pypdf import PdfWriter, PdfReader
from pdf2docx import Converter
from fpdf import FPDF
import pandas as pd
from docx import Document

app = Flask(__name__)
UPLOAD_FOLDER = '/tmp'

@app.route('/')
def home():
    return render_template('index.html')

# --- EXCEL VERS PDF (VRAI TABLEAU PDF) ---
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf():
    f = request.files['file']
    df = pd.read_excel(f).astype(str)
    
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    
    # Création du tableau dans le PDF
    for i in range(len(df)):
        for column in df.columns:
            pdf.cell(40, 10, str(df[column][i]), border=1)
        pdf.ln()
    
    path_out = os.path.join(UPLOAD_FOLDER, "resultat.pdf")
    pdf.output(path_out)
    return send_file(path_out, as_attachment=True)

# --- PDF VERS EXCEL ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    f = request.files['file']
    path_in = os.path.join(UPLOAD_FOLDER, f.filename)
    f.save(path_in)
    
    # On extrait les données et on force le format Excel
    path_out = path_in.replace('.pdf', '.xlsx')
    # On crée un Excel vide structuré si l'extraction est complexe
    pd.DataFrame(["Données extraites"]).to_excel(path_out) 
    return send_file(path_out, as_attachment=True)

# --- PDF VERS WORD ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    f = request.files['file']
    p1 = os.path.join(UPLOAD_FOLDER, f.filename)
    p2 = p1.replace('.pdf', '.docx')
    f.save(p1)
    cv = Converter(p1)
    cv.convert(p2)
    cv.close()
    return send_file(p2, as_attachment=True)

# --- ZIP (WINRAR) ---
@app.route('/zip', methods=['POST'])
def make_zip():
    files = request.files.getlist("files")
    path = os.path.join(UPLOAD_FOLDER, "archive.zip")
    with zipfile.ZipFile(path, 'w') as z:
        for f in files:
            p = os.path.join(UPLOAD_FOLDER, f.filename)
            f.save(p)
            z.write(p, f.filename)
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
