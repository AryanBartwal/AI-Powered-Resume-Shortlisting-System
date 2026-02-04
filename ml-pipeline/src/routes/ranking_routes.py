from flask import Blueprint, request, jsonify
from utils.similarity import rank_resumes, calculate_tfidf_similarity

ranking_routes = Blueprint('ranking', __name__)

@ranking_routes.route('/calculate', methods=['POST'])
def calculate_ranking():
    """Calculate resume rankings based on job description"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'success': False, 'error': 'No data provided'}), 400
        
        # Support both camelCase and snake_case
        job_description = data.get('job_description') or data.get('jobDescription', '')
        resumes = data.get('resumes', [])
        required_skills = data.get('required_skills') or data.get('requiredSkills', [])
        
        if not job_description:
            return jsonify({'success': False, 'error': 'Job description is required'}), 400
        
        if not resumes:
            return jsonify({'success': False, 'error': 'No resumes provided'}), 400
        
        # Perform ranking
        rankings = rank_resumes(job_description, resumes, required_skills)
        
        return jsonify({
            'success': True,
            'data': rankings,
            'total': len(rankings)
        })
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@ranking_routes.route('/similarity', methods=['POST'])
def calculate_similarity():
    """Calculate similarity between two texts"""
    try:
        data = request.get_json()
        
        text1 = data.get('text1', '')
        text2 = data.get('text2', '')
        
        if not text1 or not text2:
            return jsonify({'success': False, 'error': 'Both texts are required'}), 400
        
        similarity_score = calculate_tfidf_similarity(text1, text2)
        
        return jsonify({
            'success': True,
            'data': {
                'similarity_score': similarity_score * 100,
                'match_percentage': similarity_score * 100
            }
        })
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
