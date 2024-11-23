import os
from flask import Flask, render_template, request, send_from_directory, jsonify
from docx import Document
import pypandoc  # Linux-compatible for DOCX to PDF conversion
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader, PdfWriter
import logging
import subprocess
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

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

# Configure logging
logging.basicConfig(level=logging.DEBUG)

# Check file extension
def allowed_file(filename):
    if '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS:
        return True
    return False

# Function to extract metadata from the docx file
def get_docx_metadata(file_path):
    try:
        document = Document(file_path)
        metadata = {
            'title': document.core_properties.title,
            'author': document.core_properties.author,
            'subject': document.core_properties.subject,
            'keywords': document.core_properties.keywords,
            'created': document.core_properties.created,
        }
        logging.debug(f"Metadata extracted: {metadata}")
        return metadata
    except Exception as e:
        logging.error(f"Error extracting metadata: {str(e)}")
        return None

# Function to add password protection to a PDF
def add_password_to_pdf(pdf_path, password):
    try:
        with open(pdf_path, 'rb') as input_pdf:
            reader = PdfReader(input_pdf)
            writer = PdfWriter()

            for page in reader.pages:
                writer.add_page(page)

            writer.encrypt(password)

            protected_pdf_path = pdf_path  # Overwrite the original file
            with open(protected_pdf_path, 'wb') as output_pdf:
                writer.write(output_pdf)

        logging.debug(f"Password protection added to PDF: {pdf_path}")
        return protected_pdf_path  # Return the path of the protected PDF
    except Exception as e:
        logging.error(f"Error adding password to PDF: {str(e)}")
        return None

# Function to convert DOCX to PDF using pypandoc
def convert_docx_to_pdf(docx_path, pdf_path):
    doc = Document(docx_path)
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter

    y_position = height - 40
    for para in doc.paragraphs:
        c.drawString(40, y_position, para.text)
        y_position -= 12
        if y_position < 40:
            c.showPage()
            y_position = height - 40

    c.save()
    logging.debug(f"Converted DOCX to PDF: {pdf_path}")
    return pdf_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            logging.warning("No file part in the request")
            return 'No file part', 400
        
        file = request.files['file']
        
        if file.filename == '':
            logging.warning("No selected file")
            return 'No selected file', 400
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            # Extract metadata
            metadata = get_docx_metadata(file_path)
            if not metadata:
                return 'Failed to extract metadata', 500

            # Convert docx to pdf
            pdf_filename = filename.rsplit('.', 1)[0] + '.pdf'
            pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)

            # Convert DOCX to PDF using pypandoc
            if not convert_docx_to_pdf(file_path, pdf_path):
                return 'Error during DOCX to PDF conversion', 500

            # Check if password is set and encrypt PDF if necessary
            password = request.form.get('password')
            if password:
                # Add password protection and overwrite the original PDF path
                protected_pdf_path = add_password_to_pdf(pdf_path, password)
                if not protected_pdf_path:
                    return 'Error adding password protection to PDF', 500
                pdf_path = protected_pdf_path

            logging.info(f"File uploaded and processed: {pdf_filename}")
            return render_template('result.html', metadata=metadata, pdf_filename=pdf_filename, password=password)

        logging.warning(f"Invalid file format: {file.filename}")
        return 'Invalid file format', 400
    except Exception as e:
        logging.error(f"Error during file upload and processing: {str(e)}")
        return 'Internal Server Error', 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        return send_from_directory(app.config['OUTPUT_FOLDER'], filename)
    except Exception as e:
        logging.error(f"Error sending file {filename}: {str(e)}")
        return 'File not found', 404

@app.route('/encrypt', methods=['POST'])
def encrypt_pdf_service():
    try:
        data = request.json
        pdf_filename = data.get('pdf_filename')
        password = data.get('password')
        
        if not pdf_filename or not password:
            logging.warning("Missing PDF filename or password")
            return jsonify({"error": "Missing PDF filename or password"}), 400
        
        # Add password protection to the PDF
        pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], pdf_filename)
        protected_pdf_path = add_password_to_pdf(pdf_path, password)
        if not protected_pdf_path:
            return jsonify({"error": "Failed to add password to PDF"}), 500
        
        logging.info(f"Password protected PDF generated: {protected_pdf_path}")
        return jsonify({"protected_pdf_url": protected_pdf_path})
    except Exception as e:
        logging.error(f"Error during PDF encryption service: {str(e)}")
        return jsonify({"error": "Internal Server Error"}), 500

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=port)
