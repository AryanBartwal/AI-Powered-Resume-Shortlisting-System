from flask import Blueprint, request, jsonify
from utils.parser import extract_text, clean_text
import os

parse_routes = Blueprint('parse', __name__)

@parse_routes.route('/resume', methods=['POST'])
def parse_resume():
    """Parse resume and extract text"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file provided'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'}), 400
        
        # Save temporarily
        temp_path = f"/tmp/{file.filename}"
        file.save(temp_path)
        
        # Extract text
        text = extract_text(temp_path)
        cleaned_text = clean_text(text)
        
        # Clean up temp file
        os.remove(temp_path)
        
        return jsonify({
            'success': True,
            'data': {
                'text': cleaned_text,
                'raw_text': text,
                'filename': file.filename
            }
        })
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@parse_routes.route('/job-description', methods=['POST'])
def parse_job_description():
    """Parse and clean job description text"""
    try:
        data = request.get_json()
        
        if not data or 'description' not in data:
            return jsonify({'success': False, 'error': 'No description provided'}), 400
        
        description = data['description']
        cleaned_description = clean_text(description)
        
        return jsonify({
            'success': True,
            'data': {
                'cleaned_text': cleaned_description,
                'original_text': description
            }
        })
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
