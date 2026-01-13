from flask import Flask, render_template, request, send_file
import os, zipfile
from pypdf import PdfWriter, PdfReader
from pdf2docx import Converter
import pandas as pd
from docx import Document

app = Flask(__name__)
UPLOAD_FOLDER = '/tmp'

@app.route('/')
def home():
    return render_template('index.html')

# --- TOUT VERS PDF ---
@app.route('/to-pdf', methods=['POST'])
def to_pdf():
    f = request.files['file']
    ext = f.filename.split('.')[-1].lower()
    path_in = os.path.join(UPLOAD_FOLDER, f.filename)
    f.save(path_in)
    
    # Note: Sur Render (Linux), Word->PDF direct nécessite un moteur lourd. 
    # On propose ici une conversion structurée pour Excel/Word.
    if ext in ['xlsx', 'xls']:
        df = pd.read_excel(path_in)
        path_out = path_in.replace(f'.{ext}', '.html')
        df.to_html(path_out)
        return send_file(path_out, as_attachment=True, download_name="table_view.html")
    
    return "Format reçu pour conversion PDF (Aperçu HTML généré)"

# --- PDF VERS TOUT ---
@app.route('/pdf-to-any', methods=['POST'])
def pdf_to_any():
    target = request.form.get('target')
    f = request.files['file']
    path_in = os.path.join(UPLOAD_FOLDER, f.filename)
    f.save(path_in)

    if target == 'word':
        path_out = path_in.replace('.pdf', '.docx')
        cv = Converter(path_in); cv.convert(path_out); cv.close()
    elif target == 'excel':
        # Extraction des tableaux du PDF vers Excel
        path_out = path_in.replace('.pdf', '.xlsx')
        tables = pd.read_html(path_in) # Alternative via pandas
        pd.concat(tables).to_excel(path_out)
        
    return send_file(path_out, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
