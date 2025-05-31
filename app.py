from flask import Flask, render_template, request, send_file, jsonify
from docxtpl import DocxTemplate
import pandas as pd
import io
import zipfile
import os
import logging
from werkzeug.utils import secure_filename

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Configure maximum file size (5MB)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024

# Ensure the required directories exist
UPLOAD_FOLDER = 'uploads'
TEMPLATE_FOLDER = os.path.join('templates', 'word_templates')

for folder in [UPLOAD_FOLDER, TEMPLATE_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

def validate_excel(file):
    """Validate the uploaded Excel file."""
    if not file:
        return False, "No file uploaded"
    if file.filename == '':
        return False, "No file selected"
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        return False, "File must be an Excel file (.xlsx or .xls)"
    return True, None

def process_templates(df):
    """Process each row of the dataframe with one template."""  # Changed comment
    memory_file = io.BytesIO()
    
    try:
        # Debug CSV data
        print("\n=== CSV DEBUG INFO ===")
        print("CSV Headers:", df.columns.tolist())
        print("First row data:", df.iloc[0].to_dict() if len(df) > 0 else "No data")
        filename_column = 'Student_Name'  # CHANGE THIS to your preferred column name

        with zipfile.ZipFile(memory_file, 'w') as zf:
            # Process each row
            for index, row in df.iterrows():
                print(f"\n=== Processing Row {index + 1} ===")
                
                # Convert row to dict and clean up the data
                context = row.to_dict()
                # Convert any non-string values to strings and handle NaN
                context = {k: str(v) if pd.notnull(v) else '' for k, v in context.items()}
                print(f"Row data:", context)
                if filename_column not in context:
                    raise ValueError(f"Column '{filename_column}' not found in CSV headers")
                
                # NEW: Get the file identifier from the CSV column
                file_identifier = context[filename_column]

                # REMOVED: for i in range(1, 6):  <- This loop is gone
                
                # CHANGED: Now just looks for one template file
                template_path = os.path.join(TEMPLATE_FOLDER, 'template1.docx')  # Changed from f'template{i}.docx'
                print("\nChecking template...")  # Changed message
                
                if not os.path.exists(template_path):
                    print(f"ERROR: Template file not found: {template_path}")
                    raise FileNotFoundError(f"Template not found at {template_path}")
                
                try:
                    print("Loading template...")  # Changed message
                    doc = DocxTemplate(template_path)
                    
                    # Get template variables
                    variables = doc.get_undeclared_template_variables()
                    print(f"Template variables:", variables)  # Changed message
                    
                    # Check for missing variables
                    missing_vars = [var for var in variables if var not in context]
                    if missing_vars:
                        print(f"WARNING: Missing variables:", missing_vars)  # Changed message
                        print("Available context keys:", list(context.keys()))
                    
                    # Render the document
                    print("Rendering template...")  # Changed message
                    doc.render(context)
                    
                    # Save to memory
                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    
                    # CHANGED: Simplified filename
                    filename = f'{file_identifier}.docx'  # Changed from f'row_{index+1}_template_{i}.docx'
                    zf.writestr(filename, doc_buffer.getvalue())
                    print(f"Successfully processed document for row {index + 1}")  # Changed message
                    
                except Exception as e:
                    print(f"ERROR processing template for row {index + 1}")  # Changed message
                    print(f"Error details: {str(e)}")
                    raise Exception(f"Template processing failed: {str(e)}")  # Changed message
        
        memory_file.seek(0)
        return memory_file
        
    except Exception as e:
        print(f"Process templates failed: {str(e)}")
        raise


@app.route('/')
def index():
    """Render the upload form."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and document generation."""
    try:
        # Validate request has file
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        # Validate file
        is_valid, error_message = validate_excel(file)
        if not is_valid:
            return jsonify({'error': error_message}), 400
        
        try:
            # Read CSV file
            df = pd.read_excel(file)
            logger.debug(f"CSV Headers: {df.columns.tolist()}")
            logger.debug(f"Number of rows: {len(df)}")
            
            if df.empty:
                return jsonify({'error': 'excel file is empty'}), 400
            
        except Exception as e:
            logger.error(f"Error reading excel: {str(e)}")
            return jsonify({'error': 'Error reading excel file. Please ensure it is properly formatted.'}), 400
        
        try:
            # Process templates and create ZIP
            memory_file = process_templates(df)
            
            return send_file(
                memory_file,
                mimetype='application/zip',
                as_attachment=True,
                download_name='generated_documents.zip'
            )
            
        except Exception as e:
            logger.error(f"Error processing templates: {str(e)}")
            return jsonify({'error': 'Error processing templates. Please check template format and CSV data.'}), 500
            
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        return jsonify({'error': 'An unexpected error occurred'}), 500

@app.errorhandler(413)
def too_large(e):
    """Handle files that exceed the maximum size limit."""
    return jsonify({'error': 'File is too large. Maximum size is 5MB'}), 413

if __name__ == '__main__':
    app.run(debug=True)
