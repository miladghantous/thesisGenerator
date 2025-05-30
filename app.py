from flask import Flask, render_template, request, send_file, jsonify
from docxtpl import DocxTemplate
import pandas as pd
import io
import zipfile
import os
import logging
from werkzeug.utils import secure_filename
from flask_cors import CORS

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# Configure maximum file size (5MB)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024

# Get the absolute path to the directory containing app.py
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Ensure the required directories exist
TEMPLATE_FOLDER = os.path.join(BASE_DIR, 'templates', 'word_templates')
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

for folder in [UPLOAD_FOLDER, TEMPLATE_FOLDER]:
    os.makedirs(folder, exist_ok=True)

def validate_csv(file):
    """Validate the uploaded CSV file."""
    if not file:
        return False, "No file uploaded"
    if file.filename == '':
        return False, "No file selected"
    if not file.filename.endswith('.csv'):
        return False, "File must be a CSV"
    return True, None

@app.route('/')
def index():
    """Render the upload form."""
    try:
        return render_template('index.html')
    except Exception as e:
        logger.error(f"Error rendering index: {str(e)}")
        return jsonify({'error': 'Internal server error'}), 500

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and document generation."""
    try:
        # Validate request has file
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        # Validate file
        is_valid, error_message = validate_csv(file)
        if not is_valid:
            return jsonify({'error': error_message}), 400
        
        try:
            # Read CSV file
            df = pd.read_csv(file)
            logger.debug(f"CSV Headers: {df.columns.tolist()}")
            
            if df.empty:
                return jsonify({'error': 'CSV file is empty'}), 400
            
            # Process templates and create ZIP
            memory_file = process_templates(df)
            
            return send_file(
                memory_file,
                mimetype='application/zip',
                as_attachment=True,
                download_name='generated_documents.zip'
            )
            
        except Exception as e:
            logger.error(f"Error processing file: {str(e)}")
            return jsonify({'error': str(e)}), 400
            
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        return jsonify({'error': 'An unexpected error occurred'}), 500

def process_templates(df):
    """Process each row of the dataframe with one template."""
    memory_file = io.BytesIO()
    filename_column = 'Student_Name'
    try:
        with zipfile.ZipFile(memory_file, 'w') as zf:
            # Process each row
            for index, row in df.iterrows():
                # Convert row to dict and clean up the data
                context = row.to_dict()
                # Convert any non-string values to strings and handle NaN
                context = {k: str(v) if pd.notnull(v) else '' for k, v in context.items()}
                logger.debug(f"Processing row {index + 1} with context: {context}")
                if filename_column not in context:
                    raise ValueError(f"Column '{filename_column}' not found in CSV headers")
                
                # NEW: Get the file identifier from the CSV column
                file_identifier = context[filename_column]
                
                template_path = os.path.join(TEMPLATE_FOLDER, 'template1.docx')
                if not os.path.exists(template_path):
                    raise FileNotFoundError(f"Template not found at {template_path}")
                
                try:
                    doc = DocxTemplate(template_path)
                    doc.render(context)
                    
                    # Save to memory
                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    
                    # Add to ZIP
                    filename = f'{file_identifier}.docx'
                    zf.writestr(filename, doc_buffer.getvalue())
                    
                except Exception as e:
                    logger.error(f"Error processing row {index + 1}: {str(e)}")
                    raise Exception(f"Error processing row {index + 1}: {str(e)}")
        
        memory_file.seek(0)
        return memory_file
        
    except Exception as e:
        logger.error(f"Process templates failed: {str(e)}")
        raise

@app.errorhandler(413)
def too_large(e):
    """Handle files that exceed the maximum size limit."""
    return jsonify({'error': 'File is too large. Maximum size is 5MB'}), 413

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
