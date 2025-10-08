from flask import Flask, request, render_template, send_file, redirect, url_for
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(url_for('index'))
        file = request.files['file']
        if file.filename == '':
            return redirect(url_for('index'))
        if file and file.filename.endswith('.docx'):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename.replace('.docx', '.pdf'))
            convert_docx_to_pdf(filepath, pdf_path)
            return send_file(pdf_path, as_attachment=True, download_name='converted.pdf')
    return render_template('index.html')

def convert_docx_to_pdf(docx_path, pdf_path):
    doc = Document(docx_path)
    pdf = SimpleDocTemplate(pdf_path, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    for para in doc.paragraphs:
        if para.text.strip():  # Only add non-empty paragraphs
            story.append(Paragraph(para.text, styles['Normal']))
    pdf.build(story)
    os.remove(docx_path)  # Clean up uploaded file

if __name__ == '__main__':
    app.run(debug=True)