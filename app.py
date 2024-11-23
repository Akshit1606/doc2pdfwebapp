import os
from flask import Flask, render_template, request, send_from_directory, jsonify
from docx import Document
from docx2pdf import convert  # Cross-platform library for DOCX to PDF conversion
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader, PdfWriter

# Initialize Flask app
app = Flask(__name__)

# Use the PORT environment variable if it's available, otherwise default to 5000
port = int(os.environ.get("PORT", 5000))

# Configure upload folder and allowed extensions
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'converted_pdfs'
ALLOWED_EXTENSIONS = {'docx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# Ensure the upload and output directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Check file extension
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Function to extract metadata from the docx file
def get_docx_metadata(file_path):
    document = Document(file_path)
    metadata = {
        'title': document.core_properties.title,
        'author': document.core_properties.author,
        'subject': document.core_properties.subject,
        'keywords': document.core_properties.keywords,
        'created': document.core_properties.created,
    }
    return metadata

# Function to add password protection to a PDF
def add_password_to_pdf(pdf_path, password):
    with open(pdf_path, 'rb') as input_pdf:
        reader = PdfReader(input_pdf)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        writer.encrypt(password)

        protected_pdf_path = pdf_path  # Overwrite the original file
        with open(protected_pdf_path, 'wb') as output_pdf:
            writer.write(output_pdf)

    return protected_pdf_path  # Return the path of the protected PDF


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part', 400
    
    file = request.files['file']
    
    if file.filename == '':
        return 'No selected file', 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        # Extract metadata
        metadata = get_docx_metadata(file_path)

        # Convert docx to pdf
        pdf_filename = filename.rsplit('.', 1)[0] + '.pdf'
        pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)

        # Convert DOCX to PDF
        convert(file_path, pdf_path)

        # Check if password is set and encrypt PDF if necessary
        password = request.form.get('password')
        if password:
            # Add password protection and overwrite the original PDF path
            pdf_path = add_password_to_pdf(pdf_path, password)

        return render_template('result.html', metadata=metadata, pdf_filename=pdf_filename, password=password)

    return 'Invalid file format', 400


@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename)

@app.route('/encrypt', methods=['POST'])
def encrypt_pdf_service():
    data = request.json
    pdf_filename = data.get('pdf_filename')
    password = data.get('password')
    
    if not pdf_filename or not password:
        return jsonify({"error": "Missing PDF filename or password"}), 400
    
    # Add password protection to the PDF
    pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)
    protected_pdf_path = add_password_to_pdf(pdf_path, password)
    
    return jsonify({"protected_pdf_url": protected_pdf_path})

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=port)

# import os
# from flask import Flask, render_template, request, send_from_directory, jsonify
# from docx import Document
# from docx2pdf import convert
# from werkzeug.utils import secure_filename
# from PyPDF2 import PdfReader, PdfWriter
# import pythoncom
# import win32com.client

# # Initialize Flask app
# app = Flask(__name__)

# # Configure upload folder and allowed extensions
# UPLOAD_FOLDER = 'uploads'
# OUTPUT_FOLDER = 'converted_pdfs'
# ALLOWED_EXTENSIONS = {'docx'}

# app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
# app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# # Ensure the upload and output directories exist
# os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# # Check file extension
# def allowed_file(filename):
#     return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# # Function to extract metadata from the docx file
# def get_docx_metadata(file_path):
#     document = Document(file_path)
#     metadata = {
#         'title': document.core_properties.title,
#         'author': document.core_properties.author,
#         'subject': document.core_properties.subject,
#         'keywords': document.core_properties.keywords,
#         'created': document.core_properties.created,
#     }
#     return metadata

# # Function to add password protection to a PDF
# def add_password_to_pdf(pdf_path, password):
#     with open(pdf_path, 'rb') as input_pdf:
#         reader = PdfReader(input_pdf)
#         writer = PdfWriter()

#         for page in reader.pages:
#             writer.add_page(page)

#         writer.encrypt(password)

#         protected_pdf_path = pdf_path  # Overwrite the original file
#         with open(protected_pdf_path, 'wb') as output_pdf:
#             writer.write(output_pdf)

#     return protected_pdf_path  # Return the path of the protected PDF


# @app.route('/')
# def index():
#     return render_template('index.html')

# @app.route('/upload', methods=['POST'])
# def upload_file():
#     if 'file' not in request.files:
#         return 'No file part', 400
    
#     file = request.files['file']
    
#     if file.filename == '':
#         return 'No selected file', 400
    
#     if file and allowed_file(file.filename):
#         filename = secure_filename(file.filename)
#         file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
#         file.save(file_path)

#         # Extract metadata
#         metadata = get_docx_metadata(file_path)

#         # Convert docx to pdf
#         pdf_filename = filename.rsplit('.', 1)[0] + '.pdf'
#         pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)

#         # Initialize COM before conversion
#         pythoncom.CoInitialize()
#         convert(file_path, pdf_path)
#         pythoncom.CoUninitialize()

#         # Check if password is set and encrypt PDF if necessary
#         password = request.form.get('password')
#         if password:
#             # Add password protection and overwrite the original PDF path
#             pdf_path = add_password_to_pdf(pdf_path, password)

#         return render_template('result.html', metadata=metadata, pdf_filename=pdf_filename, password=password)

#     return 'Invalid file format', 400


# @app.route('/download/<filename>')
# def download_file(filename):
#     return send_from_directory(app.config['OUTPUT_FOLDER'], filename)

# @app.route('/encrypt', methods=['POST'])
# def encrypt_pdf_service():
#     data = request.json
#     pdf_filename = data.get('pdf_filename')
#     password = data.get('password')
    
#     if not pdf_filename or not password:
#         return jsonify({"error": "Missing PDF filename or password"}), 400
    
#     # Add password protection to the PDF
#     pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)
#     protected_pdf_path = add_password_to_pdf(pdf_path, password)
    
#     return jsonify({"protected_pdf_url": protected_pdf_path})

# if __name__ == '__main__':
#     app.run(debug=True)
