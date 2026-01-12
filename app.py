from flask import Flask, render_template, request, send_file
import os, zipfile
from pypdf import PdfWriter, PdfReader
# On utilise pypandoc ou d'autres librairies pour Linux à la place de docx2pdf
from pdf2docx import Converter
from PIL import Image
import pandas as pd

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER): 
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def home():
    return render_template('index.html')

# 1. FUSION
@app.route('/merge', methods=['POST'])
def merge():
    files = request.files.getlist("pdfs")
    merger = PdfWriter()
    for file in files: 
        merger.append(file)
    path = os.path.join(UPLOAD_FOLDER, "fusion.pdf")
    merger.write(path)
    merger.close()
    return send_file(path, as_attachment=True)

# 2. WORD TO PDF (Version Linux)
@app.route('/word', methods=['POST'])
def word():
    return "La conversion Word vers PDF nécessite des outils spéciaux sur Linux. Utilisez PDF vers Word pour l'instant !", 400

# 3. EXCEL TO PDF (Désactivé car nécessite Windows)
@app.route('/excel', methods=['POST'])
def excel():
    return "La conversion Excel vers PDF ne fonctionne pas sur serveur Linux gratuit.", 400

# 4. PDF TO WORD
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    f = request.files['file']
    in_p = os.path.join(UPLOAD_FOLDER, f.filename)
    out_p = in_p.replace('.pdf', '.docx')
    f.save(in_p)
    cv = Converter(in_p)
    cv.convert(out_p)
    cv.close()
    return send_file(out_p, as_attachment=True)

# 5. PDF TO EXCEL
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    f = request.files['file']
    in_p = os.path.join(UPLOAD_FOLDER, f.filename)
    out_p = in_p.replace('.pdf', '.xlsx')
    f.save(in_p)
    reader = PdfReader(in_p)
    text_data = []
    for page in reader.pages:
        text_data.append(page.extract_text())
    pd.DataFrame(text_data).to_excel(out_p, index=False)
    return send_file(out_p, as_attachment=True)

# 6. IMAGE TO PDF
@app.route('/img', methods=['POST'])
def img_to_pdf():
    f = request.files['file']
    img = Image.open(f).convert('RGB')
    path = os.path.join(UPLOAD_FOLDER, "image.pdf")
    img.save(path)
    return send_file(path, as_attachment=True)

# 7. ZIP
@app.route('/zip', methods=['POST'])
def to_zip():
    files = request.files.getlist("files")
    path = os.path.join(UPLOAD_FOLDER, "archive.zip")
    with zipfile.ZipFile(path, 'w') as z:
        for f in files:
            fp = os.path.join(UPLOAD_FOLDER, f.filename)
            f.save(fp)
            z.write(fp, f.filename)
            os.remove(fp)
    return send_file(path, as_attachment=True)

# 8. PROTÉGER
@app.route('/protect', methods=['POST'])
def protect():
    f = request.files['file']
    pw = request.form['pw']
    reader = PdfReader(f)
    writer = PdfWriter()
    for p in reader.pages: 
        writer.add_page(p)
    writer.encrypt(pw)
    path = os.path.join(UPLOAD_FOLDER, "secure.pdf")
    with open(path, "wb") as out: 
        writer.write(out)
    return send_file(path, as_attachment=True)

# 9. PAGES
@app.route('/remove-pages', methods=['POST'])
def remove_pages():
    f = request.files['file']
    pages_to_keep = request.form['pages']
    in_p = os.path.join(UPLOAD_FOLDER, f.filename)
    out_p = os.path.join(UPLOAD_FOLDER, "pages_modifiees.pdf")
    f.save(in_p)
    reader = PdfReader(in_p)
    writer = PdfWriter()
    try:
        page_indices = [int(p.strip()) - 1 for p in pages_to_keep.split(',')]
        for i in page_indices:
            if 0 <= i < len(reader.pages): 
                writer.add_page(reader.pages[i])
        with open(out_p, "wb") as out: 
            writer.write(out)
        return send_file(out_p, as_attachment=True)
    except: 
        return "Erreur: Format de pages invalide", 400

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)