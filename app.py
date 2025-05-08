from flask import Flask, render_template, request, send_from_directory, jsonify
import os
import uuid
from config import Config
from services.ai_service import generate_yaml_from_topic
from services.document_service import (
    create_pptx_from_yaml,
    create_docx_from_yaml,
    create_pdf_from_yaml,
    create_html_from_yaml
)
from services.preview_service import generate_preview_images
import yaml

app = Flask(__name__)
app.config.from_object(Config)

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    topic = request.form.get('topic')
    file_type = request.form.get('file_type')
    html_presentation_type = request.form.get('html_presentation_type', 'minimalist')  # Default to minimalist
    include_images = request.form.get('include_images', 'true').lower() == 'true'  # Default to true
    
    if not topic or not file_type:
        return jsonify({"success": False, "error": "Missing topic or file type"})
    
    # Validate HTML presentation type
    if file_type == 'html' and html_presentation_type not in ['minimalist', 'modern', 'professional']:
        html_presentation_type = 'minimalist'
    
    # Generate a unique ID for this file
    file_id = str(uuid.uuid4())
    file_name = f"{topic.replace(' ', '_')}_{file_id}"
    
    # Generate YAML content with the include_images flag
    yaml_content = generate_yaml_from_topic(topic, include_images=include_images)
    yaml_content_preview = None
    
    try:
        # Parse YAML to get preview content
        preview_data = yaml.safe_load(yaml_content)
        yaml_content_preview = preview_data
    except Exception as e:
        return jsonify({"success": False, "error": f"YAML parsing error: {str(e)}"})
    
    # Generate preview images for HTML presentation types
    preview_images = {}
    if file_type == 'html':
        preview_images = generate_preview_images(yaml_content, topic, app.config['UPLOAD_FOLDER'])
    
    # Process according to file type
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
    result = {"success": False, "error": "Invalid file type"}
    
    if file_type == "pptx":
        full_path = f"{output_path}.pptx"
        result = create_pptx_from_yaml(yaml_content, full_path, topic)
        if result["success"]:
            result["file_url"] = f"/download/{file_name}.pptx"
            result["preview"] = yaml_content_preview
    
    elif file_type == "docx":
        full_path = f"{output_path}.docx"
        result = create_docx_from_yaml(yaml_content, full_path)
        if result["success"]:
            result["file_url"] = f"/download/{file_name}.docx"
            result["preview"] = yaml_content_preview
    
    elif file_type == "pdf":
        full_path = f"{output_path}.pdf"
        result = create_pdf_from_yaml(yaml_content, full_path)
        if result["success"]:
            result["file_url"] = f"/download/{file_name}.pdf"
            result["preview"] = yaml_content_preview

    elif file_type == "html":
        full_path = f"{output_path}.html"
        result = create_html_from_yaml(yaml_content, full_path, topic, html_presentation_type=html_presentation_type)
        if result["success"]:
            result["file_url"] = f"/download/{file_name}.html"
            result["preview"] = yaml_content_preview
            result["preview_images"] = preview_images
    
    return jsonify(result)

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)