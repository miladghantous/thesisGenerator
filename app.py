from flask import Flask, render_template, request, send_file, jsonify, url_for
from docxtpl import DocxTemplate
import pandas as pd
import io
import zipfile
import os
import logging
import gc
from werkzeug.utils import secure_filename
from time import time
import threading
import uuid
from datetime import datetime, timedelta
import atexit
from concurrent.futures import ThreadPoolExecutor

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Configure maximum file size (5MB)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024

# Task storage
task_results = {}
RESULT_EXPIRY = timedelta(minutes=30)
BATCH_SIZE = 5
MAX_WORKERS = 2

# Ensure the required directories exist
UPLOAD_FOLDER = 'uploads'
TEMPLATE_FOLDER = os.path.join('templates', 'word_templates')

for folder in [UPLOAD_FOLDER, TEMPLATE_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# Global thread pool executor
executor = ThreadPoolExecutor(max_workers=MAX_WORKERS)

def validate_excel(file):
    """Validate the uploaded Excel file."""
    if not file:
        return False, "No file uploaded"
    if file.filename == '':
        return False, "No file selected"
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        return False, "File must be an Excel file (.xlsx or .xls)"
    return True, None

def process_single_row(template_path, row_data, index):
    """Process a single row of data."""
    try:
        start_time = time()
        context = {k: str(v) if pd.notnull(v) else '' for k, v in row_data.items()}
        filename_column = 'Student_Name'
        
        if filename_column not in context:
            raise ValueError(f"Column '{filename_column}' not found in headers")
        
        file_identifier = context[filename_column]
        
        # Load and process template
        doc = DocxTemplate(template_path)
        doc.render(context)
        
        # Save to memory
        doc_buffer = io.BytesIO()
        doc.save(doc_buffer)
        doc_buffer.seek(0)
        
        # Clear some memory
        del doc
        gc.collect()
        
        return {
            'success': True,
            'filename': f'{file_identifier}.docx',
            'data': doc_buffer.getvalue(),
            'index': index,
            'processing_time': time() - start_time
        }
        
    except Exception as e:
        logger.error(f"Error processing row {index}: {str(e)}")
        return {
            'success': False,
            'error': str(e),
            'index': index
        }

def process_templates_background(task_id, df):
    """Process the dataframe in batches."""
    logger.info(f"Starting background processing for task {task_id}")
    try:
        memory_file = io.BytesIO()
        template_path = os.path.join(TEMPLATE_FOLDER, 'template1.docx')
        
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template not found at {template_path}")
        
        with zipfile.ZipFile(memory_file, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
            total_rows = len(df)
            processed = 0
            
            # Process in batches
            while processed < total_rows:
                batch = df.iloc[processed:processed + BATCH_SIZE]
                batch_size = len(batch)
                
                # Process each row in the batch
                for idx, row in batch.iterrows():
                    try:
                        result = process_single_row(template_path, row.to_dict(), idx)
                        if result['success']:
                            zf.writestr(result['filename'], result['data'])
                            processed += 1
                            # Update progress
                            progress = int((processed * 100) / total_rows)
                            logger.info(f"Task {task_id}: Processed {processed}/{total_rows} ({progress}%)")
                            task_results[task_id]['progress'] = progress
                        else:
                            logger.error(f"Failed to process row {idx}: {result.get('error')}")
                    except Exception as e:
                        logger.error(f"Error processing row {idx}: {str(e)}")
                
                gc.collect()
        
        memory_file.seek(0)
        task_results[task_id].update({
            'status': 'completed',
            'result': memory_file.getvalue(),
            'progress': 100,
            'timestamp': datetime.now()
        })
        logger.info(f"Task {task_id} completed successfully")
        
    except Exception as e:
        error_msg = str(e)
        logger.error(f"Task {task_id} failed: {error_msg}")
        task_results[task_id].update({
            'status': 'failed',
            'error': error_msg,
            'timestamp': datetime.now()
        })
    finally:
        # Clean up the dataframe
        if 'data' in task_results[task_id]:
            del task_results[task_id]['data']
            gc.collect()

def clean_old_results():
    """Remove old results to prevent memory leaks"""
    current_time = datetime.now()
    expired_tasks = []
    for task_id, task in task_results.items():
        if current_time - task.get('timestamp', current_time) > RESULT_EXPIRY:
            expired_tasks.append(task_id)
    
    for task_id in expired_tasks:
        try:
            del task_results[task_id]
            logger.info(f"Cleaned up expired task {task_id}")
        except KeyError:
            pass

# Register cleanup function
atexit.register(executor.shutdown)

@app.route('/')
def index():
    """Render the upload form."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and start background processing."""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        is_valid, error_message = validate_excel(file)
        if not is_valid:
            return jsonify({'error': error_message}), 400
        
        try:
            df = pd.read_excel(file, engine='openpyxl')
            
            if df.empty:
                return jsonify({'error': 'Excel file is empty'}), 400
            
            # Create new task
            task_id = str(uuid.uuid4())
            task_results[task_id] = {
                'status': 'pending',
                'progress': 0,
                'timestamp': datetime.now(),
                'data': df
            }
            
            # Start background processing
            executor.submit(process_templates_background, task_id, df)
            
            logger.info(f"Created and started task {task_id}")
            return jsonify({
                'task_id': task_id,
                'status': 'processing',
                'message': 'File processing started'
            })
            
        except Exception as e:
            logger.error(f"Error processing file: {str(e)}")
            return jsonify({'error': str(e)}), 500
            
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        return jsonify({'error': 'An unexpected error occurred'}), 500

@app.route('/status/<task_id>')
def get_status(task_id):
    """Get the status of a processing task."""
    logger.info(f"Status check for task {task_id}")
    if task_id not in task_results:
        return jsonify({'error': 'Task not found'}), 404
    
    task = task_results[task_id]
    response = {
        'status': task['status'],
        'progress': task.get('progress', 0)
    }
    
    if task.get('error'):
        response['error'] = task['error']
    
    return jsonify(response)

@app.route('/download/<task_id>')
def download_result(task_id):
    """Download the processed files."""
    logger.info(f"Download requested for task {task_id}")
    if task_id not in task_results:
        return jsonify({'error': 'Task not found'}), 404
    
    task = task_results[task_id]
    if task['status'] != 'completed':
        return jsonify({'error': 'Task not completed'}), 400
    
    if 'result' not in task:
        return jsonify({'error': 'Result not found'}), 404
    
    try:
        memory_file = io.BytesIO(task['result'])
        memory_file.seek(0)
        
        response = send_file(
            memory_file,
            mimetype='application/zip',
            as_attachment=True,
            download_name='generated_documents.zip'
        )
        
        # Clean up after sending
        clean_old_results()
        return response
        
    except Exception as e:
        logger.error(f"Error sending file: {str(e)}")
        return jsonify({'error': 'Error sending file'}), 500

@app.errorhandler(413)
def too_large(e):
    """Handle files that exceed the maximum size limit."""
    return jsonify({'error': 'File is too large. Maximum size is 5MB'}), 413

if __name__ == '__main__':
    app.run(debug=True)
