from flask import Flask, render_template, request, redirect, url_for, send_from_directory
from flask_bootstrap import Bootstrap
import os
import shutil
from pdf2docx import Converter
from datetime import datetime
import docx

app = Flask(__name__)
bootstrap = Bootstrap(app)

UPLOAD_FOLDER = 'in'
CONVERTED_FOLDER = 'out/convertidos'
ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER

def convert_pdf_to_word(pdf_path, output_folder):
    file_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_path = os.path.join(output_folder, f"{file_name}.docx")
    
    cv = Converter(pdf_path)
    cv.convert(output_path)
    cv.close()
    return output_path

def convert_pdf_to_doc(pdf_path, output_folder):
    file_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_path = os.path.join(output_folder, f"{file_name}.doc")
    
    # Convertir PDF a DOCX primero
    docx_path = convert_pdf_to_word(pdf_path, output_folder)
    
    # Luego, convertir DOCX a DOC
    doc = docx.Document(docx_path)
    doc.save(output_path)
    
    # Eliminar el archivo DOCX temporal
    os.remove(docx_path)
    
    return output_path

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_file_data(filename, filepath):
    size = os.path.getsize(filepath)
    size_kb = size / 1024
    modified_time = datetime.fromtimestamp(os.path.getmtime(filepath))
    return {
        'name': filename,
        'size': f"{size_kb:.2f} KB",
        'extension': filename.split('.')[-1],
        'date': modified_time.strftime("%Y-%m-%d %H:%M:%S"),
        'path': filepath
    }

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = file.filename
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            if filename.endswith('.pdf'):
                converted_file_path = convert_pdf_to_word(os.path.join(app.config['UPLOAD_FOLDER'], filename), app.config['CONVERTED_FOLDER'])
            elif filename.endswith('.doc'):
                converted_file_path = os.path.join(app.config['CONVERTED_FOLDER'], filename)
            elif filename.endswith('.docx'):
                shutil.move(os.path.join(app.config['UPLOAD_FOLDER'], filename), os.path.join(app.config['CONVERTED_FOLDER'], filename))
            return redirect(url_for('upload_file'))
    files = os.listdir(CONVERTED_FOLDER)
    converted_files = [get_file_data(f, os.path.join(CONVERTED_FOLDER, f)) for f in files]
    return render_template('index.html', files=converted_files)

@app.route('/download/<path:filename>')
def download_file(filename):
    return send_from_directory(app.config['CONVERTED_FOLDER'], filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
