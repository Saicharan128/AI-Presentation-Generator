from flask import Flask, render_template, request, send_from_directory, jsonify, redirect, url_for, session
import os
import uuid
import time
import hashlib
from config import Config
from models import db, FileRecord  # Import from models.py
from services.ai_service import generate_yaml_from_topic
from flask import send_file, abort
import io
import base64
from services.document_service import (
    create_pptx_from_yaml,
    create_docx_from_yaml,
    create_pdf_from_yaml,
    create_html_from_yaml
)
from services.preview_service import generate_preview_images
import yaml
import threading

app = Flask(__name__)
app.config.from_object(Config)

# Set a 16-character hex secret key for sessions
app.secret_key = 'a1b2c3d4e5f6a7b8'  # Replace this with your own secure 16-char hex string

# Add a lock for database operations
db_lock = threading.Lock()

# Initialize DB and create tables
db.init_app(app)
with app.app_context():
    db.create_all()

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# In-memory storage for request deduplication (use Redis in production)
recent_requests = {}

def generate_request_hash(email, topic, file_type, html_presentation_type, include_images):
    """Generate a hash for request deduplication"""
    request_data = f"{email}:{topic}:{file_type}:{html_presentation_type}:{include_images}"
    return hashlib.md5(request_data.encode()).hexdigest()

def is_duplicate_request(request_hash, window_seconds=10):
    """Check if this is a duplicate request within the time window"""
    current_time = time.time()
    
    # Clean up old requests (older than 1 minute)
    cutoff_time = current_time - 60
    keys_to_remove = [k for k, v in recent_requests.items() if v['timestamp'] < cutoff_time]
    for key in keys_to_remove:
        del recent_requests[key]
    
    # Check if this request was made recently
    if request_hash in recent_requests:
        time_diff = current_time - recent_requests[request_hash]['timestamp']
        if time_diff < window_seconds:
            return True, recent_requests[request_hash]['result']
    
    return False, None

def store_request_result(request_hash, result):
    """Store the result of a request for deduplication"""
    recent_requests[request_hash] = {
        'timestamp': time.time(),
        'result': result
    }

@app.route('/')
def index():
    if Config.ATOM_AUTHENTICATION:
        email = session['email']
        token = session['token']

        print(session)

        return render_template('index.html', email=email, token=token)
    else:
        return render_template('atom_authentication.html')

@app.route('/atom_auth')
def atom():
    global atom_authentication
    Config.ATOM_AUTHENTICATION = True
    token = request.args.get('token')
    email = request.args.get('email')

    session['token'] = token
    session['email'] = email

    return redirect(url_for('index'))    

@app.route('/generate', methods=['POST'])
def generate():
    email = session['email']
    token = session['token']

    topic = request.form.get('topic')
    file_type = request.form.get('file_type')
    html_presentation_type = request.form.get('html_presentation_type', 'minimalist')
    include_images = request.form.get('include_images', 'true').lower() == 'true'
    
    if not topic or not file_type:
        return jsonify({"success": False, "error": "Missing topic or file type"})
    
    if file_type == 'html' and html_presentation_type not in ['minimalist', 'modern', 'professional', 'corporate', 'executive', 'elegant', 'refined', 'glassy']:
        html_presentation_type = 'minimalist'
    
    # Generate request hash for deduplication
    request_hash = generate_request_hash(email, topic, file_type, html_presentation_type, include_images)
    
    # Use database lock for the entire process
    with db_lock:
        # Check for duplicate request in memory
        is_duplicate, cached_result = is_duplicate_request(request_hash)
        if is_duplicate:
            print(f"Returning cached result for duplicate request: {request_hash}")
            return jsonify(cached_result)
        
        # Check if file already exists in database
        existing_file = FileRecord.query.filter_by(
            user_email=email,
            topic=topic,
            file_type=file_type
        ).order_by(FileRecord.created_at.desc()).first()
        
        if existing_file:
            result = {
                "success": True,
                "file_url": f"/db_download/{existing_file.id}",
                "message": "File already exists"
            }
            try:
                if existing_file.yaml_content:
                    result["preview"] = yaml.safe_load(existing_file.yaml_content)
            except:
                pass
            
            store_request_result(request_hash, result)
            return jsonify(result)
        
        # Mark this request as being processed
        processing_result = {"success": False, "error": "Processing..."}
        store_request_result(request_hash, processing_result)
    
    # Generate file outside of lock (this can take time)
    file_id = str(uuid.uuid4())
    file_name = f"{topic.replace(' ', '_')}_{file_id}"
    
    try:
        yaml_content = generate_yaml_from_topic(topic, include_images=include_images)
        yaml_content_preview = None

        try:
            preview_data = yaml.safe_load(yaml_content)
            yaml_content_preview = preview_data
        except Exception as e:
            error_result = {"success": False, "error": f"YAML parsing error: {str(e)}"}
            store_request_result(request_hash, error_result)
            return jsonify(error_result)
        
        preview_images = {}
        if file_type == 'html':
            preview_images = generate_preview_images(yaml_content, topic, app.config['UPLOAD_FOLDER'], email)
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
        result = {"success": False, "error": "Invalid file type"}
        file_content = None

        # Generate the file
        if file_type == "pptx":
            full_path = f"{output_path}.pptx"
            result = create_pptx_from_yaml(yaml_content, full_path, topic, email)
            if result["success"]:
                if os.path.exists(full_path):
                    with open(full_path, 'rb') as f:
                        file_content = f.read()
                    try:
                        os.remove(full_path)
                    except:
                        pass

        elif file_type == "docx":
            full_path = f"{output_path}.docx"
            result = create_docx_from_yaml(yaml_content, full_path, email)
            if result["success"]:
                if os.path.exists(full_path):
                    with open(full_path, 'rb') as f:
                        file_content = f.read()
                    try:
                        os.remove(full_path)
                    except:
                        pass

        elif file_type == "pdf":
            full_path = f"{output_path}.pdf"
            result = create_pdf_from_yaml(yaml_content, full_path, email)
            if result["success"]:
                if os.path.exists(full_path):
                    with open(full_path, 'rb') as f:
                        file_content = f.read()
                    try:
                        os.remove(full_path)
                    except:
                        pass

        elif file_type == "html":
            full_path = f"{output_path}.html"
            result = create_html_from_yaml(yaml_content, full_path, topic, email, html_presentation_type=html_presentation_type)
            if result["success"]:
                if os.path.exists(full_path):
                    with open(full_path, 'rb') as f:
                        file_content = f.read()
                    try:
                        os.remove(full_path)
                    except:
                        pass
                result["preview_images"] = preview_images

        if result["success"] and file_content:
            # Save to database with duplicate handling
            with db_lock:
                try:
                    # Check one more time before inserting
                    existing_file = FileRecord.query.filter_by(
                        user_email=email,
                        topic=topic,
                        file_type=file_type
                    ).first()
                    
                    if existing_file:
                        # File was created by another request, return existing one
                        result["file_url"] = f"/db_download/{existing_file.id}"
                        result["preview"] = yaml_content_preview
                    else:
                        # Create new record
                        file_record = FileRecord(
                            user_email=email,
                            topic=topic,
                            file_type=file_type,
                            file_data=base64.b64encode(file_content).decode('utf-8'),
                            yaml_content=yaml_content
                        )
                        db.session.add(file_record)
                        db.session.commit()
                        
                        result["file_url"] = f"/db_download/{file_record.id}"
                        result["preview"] = yaml_content_preview
                    
                except Exception as db_error:
                    db.session.rollback()
                    print(f"Database error: {str(db_error)}")
                    # Check if it's a duplicate key error
                    if "UNIQUE constraint failed" in str(db_error) or "duplicate key" in str(db_error).lower():
                        # Fetch the existing record
                        existing_file = FileRecord.query.filter_by(
                            user_email=email,
                            topic=topic,
                            file_type=file_type
                        ).first()
                        if existing_file:
                            result["file_url"] = f"/db_download/{existing_file.id}"
                            result["preview"] = yaml_content_preview
                    else:
                        result["file_url"] = f"/download/{file_name}.{file_type}"
                        result["preview"] = yaml_content_preview

        # Store the final result
        store_request_result(request_hash, result)
        
    except Exception as e:
        error_result = {"success": False, "error": f"File generation error: {str(e)}"}
        store_request_result(request_hash, error_result)
        return jsonify(error_result)

    return jsonify(result)

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

@app.route('/db_download/<int:file_id>')
def db_download_file(file_id):
    file_record = FileRecord.query.get(file_id)
    if not file_record:
        abort(404)

    # Decode base64 content
    decoded_bytes = base64.b64decode(file_record.file_data)

    # Define filename
    filename = f"{file_record.topic.replace(' ', '_')}.{file_record.file_type}"

    # Send as downloadable file
    return send_file(
        io.BytesIO(decoded_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/octet-stream"
    )

@app.route('/my_files')
def my_files():
    email = session.get('email')
    if not email:
        return redirect(url_for('index'))

    files = FileRecord.query.filter_by(user_email=email).order_by(FileRecord.created_at.desc()).all()
    return render_template('my_files.html', files=files, email=email, token=session.get('token'))

@app.route('/logout')
def logout():
    Config.ATOM_AUTHENTICATION = False

    # Clear session keys
    session.pop('email', None)
    session.pop('token', None)
    session.pop('authenticated', None)  # If you have this flag

    return redirect(url_for('index'))

if __name__ == "__main__":
    app.run(debug=True,port=8099)