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
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import uuid
from collections import deque
from datetime import datetime, timedelta

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Configure maximum file size (5MB)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024

# Task queue and results storage
task_queue = deque()
task_results = {}
RESULT_EXPIRY = timedelta(minutes=30)

# Ensure the required directories exist
UPLOAD_FOLDER = 'uploads'
TEMPLATE_FOLDER = os.path.join('templates', 'word_templates')
BATCH_SIZE = 5  # Process 5 documents at a time
MAX_WORKERS = 2  # Limit concurrent processing

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
    """Process the dataframe in batches in the background."""
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
                
                with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                    future_to_row = {
                        executor.submit(
                            process_single_row, 
                            template_path, 
                            row.to_dict(), 
                            idx
                        ): idx 
                        for idx, row in batch.iterrows()
                    }
                    
                    for future in as_completed(future_to_row):
                        result = future.result()
                        if result['success']:
                            zf.writestr(result['filename'], result['data'])
                            logger.info(f"Processed document {result['index'] + 1}/{total_rows}")
                            # Update task progress
                            task_results[task_id]['progress'] = int((processed + 1) * 100 / total_rows)
                        else:
                            logger.error(f"Failed to process row {result['index'] + 1}: {result.get('error')}")
                
                processed += len(batch)
                gc.collect()
        
        memory_file.seek(0)
        task_results[task_id].update({
            'status': 'completed',
            'result': memory_file.getvalue(),
            'progress': 100,
            'timestamp': datetime.now()  # Update timestamp when completed
        })
        
    except Exception as e:
        logger.error(f"Process templates failed: {str(e)}")
        task_results[task_id].update({
            'status': 'failed',
            'error': str(e),
            'timestamp': datetime.now()  # Update timestamp when failed
        })

def process_task_queue():
    """Background worker to process tasks."""
    while True:
        try:
            if task_queue:
                task_id = task_queue.popleft()
                if task_id in task_results:
                    task = task_results[task_id]
                    if task['status'] == 'pending':
                        process_templates_background(task_id, task['data'])
                        # Clear the dataframe from memory after processing
                        if 'data' in task_results[task_id]:
                            del task_results[task_id]['data']
                            gc.collect()
            else:
                clean_old_results()
                threading.Event().wait(1)  # Sleep for 1 second when queue is empty
        except Exception as e:
            logger.error(f"Task queue processing error: {str(e)}")
            threading.Event().wait(1)

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
        except KeyError:
            pass  # Task might have been deleted by another thread

# Start background worker
worker_thread = threading.Thread(target=process_task_queue, daemon=True)
worker_thread.start()

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
            task_queue.append(task_id)
            
            logger.info(f"Created task {task_id}")
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
    logger.info(f"Checking status for task {task_id}")
    if task_id not in task_results:
        return jsonify({'error': 'Task not found'}), 404
    
    task = task_results[task_id]
    return jsonify({
        'status': task['status'],
        'progress': task.get('progress', 0),
        'error': task.get('error', None)
    })

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
        
        return send_file(
            memory_file,
            mimetype='application/zip',
            as_attachment=True,
            download_name='generated_documents.zip'
        )
    except Exception as e:
        logger.error(f"Error sending file: {str(e)}")
        return jsonify({'error': 'Error sending file'}), 500

@app.errorhandler(413)
def too_large(e):
    """Handle files that exceed the maximum size limit."""
    return jsonify({'error': 'File is too large. Maximum size is 5MB'}), 413

if __name__ == '__main__':
    app.run(debug=True)
