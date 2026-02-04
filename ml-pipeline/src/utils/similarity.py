from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np
import re

# BERT model - lazy loading to save memory
_bert_model = None

def get_bert_model():
    """Lazy load BERT model"""
    global _bert_model
    if _bert_model is None:
        try:
            from sentence_transformers import SentenceTransformer
            # Using a lightweight but accurate model
            _bert_model = SentenceTransformer('all-MiniLM-L6-v2')
        except Exception as e:
            _bert_model = None
    return _bert_model

def calculate_bert_similarity(text1, text2):
    """Calculate semantic similarity using BERT embeddings"""
    try:
        model = get_bert_model()
        if model is None:
            return None  # Fall back to TF-IDF
        
        # Generate embeddings
        embeddings = model.encode([text1, text2])
        
        # Calculate cosine similarity
        similarity = cosine_similarity([embeddings[0]], [embeddings[1]])
        return float(similarity[0][0])
    except Exception as e:
        return None  # Fall back to TF-IDF

def calculate_tfidf_similarity(text1, text2):
    """Calculate cosine similarity between two texts using TF-IDF"""
    try:
        # Use n-grams to capture multi-word skills like "Machine Learning"
        vectorizer = TfidfVectorizer(
            ngram_range=(1, 3),  # Capture 1, 2, and 3-word phrases
            max_features=5000,
            stop_words='english',
            sublinear_tf=True
        )
        tfidf_matrix = vectorizer.fit_transform([text1, text2])
        similarity = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])
        return float(similarity[0][0])
    except Exception as e:
        return 0.0

def normalize_skill(skill):
    """Normalize skill name for better matching"""
    # Remove special characters and extra spaces
    skill = re.sub(r'[^a-zA-Z0-9\s+#]', ' ', skill)
    skill = ' '.join(skill.split())
    return skill.lower().strip()

def get_skill_variations(skill):
    """Get common variations of a skill name"""
    variations = [skill.lower()]
    skill_lower = skill.lower()
    
    # Common variations mapping
    variation_map = {
        'javascript': ['js', 'javascript', 'java script'],
        'typescript': ['ts', 'typescript', 'type script'],
        'react.js': ['react', 'reactjs', 'react.js', 'react js'],
        'node.js': ['node', 'nodejs', 'node.js', 'node js'],
        'mongodb': ['mongo', 'mongodb', 'mongo db'],
        'postgresql': ['postgres', 'postgresql', 'postgre sql'],
        'machine learning': ['ml', 'machine learning', 'machinelearning'],
        'artificial intelligence': ['ai', 'artificial intelligence'],
        'data structures': ['dsa', 'data structures', 'data structure'],
        'object oriented programming': ['oop', 'object oriented', 'object-oriented'],
        'rest api': ['rest', 'restful', 'rest api', 'rest apis', 'restful api'],
        'aws': ['amazon web services', 'aws', 'amazon aws'],
        'c++': ['cpp', 'c++', 'cplusplus', 'c plus plus'],
        'c#': ['csharp', 'c#', 'c sharp']
    }
    
    # Check if skill matches any key or value in variation map
    for main_skill, vars_list in variation_map.items():
        if skill_lower in vars_list or any(var in skill_lower for var in vars_list):
            variations.extend(vars_list)
            break
    
    return list(set(variations))

def extract_skills_from_text(text, skill_keywords):
    """Extract matched skills from text with better matching"""
    text_lower = text.lower()
    # Normalize text for better matching
    text_normalized = re.sub(r'[^a-zA-Z0-9\s+#]', ' ', text_lower)
    matched_skills = []
    
    for skill in skill_keywords:
        skill_variations = get_skill_variations(skill)
        
        # Check if any variation exists in the text
        for variation in skill_variations:
            # Use word boundary to avoid partial matches
            pattern = r'\b' + re.escape(variation) + r'\b'
            if re.search(pattern, text_normalized, re.IGNORECASE) or variation in text_lower:
                matched_skills.append(skill)
                break
    
    return matched_skills

def calculate_skill_match_score(resume_text, required_skills):
    """Calculate score based on matched skills"""
    if not required_skills:
        return 0.0, []
    
    matched = extract_skills_from_text(resume_text, required_skills)
    match_percentage = (len(matched) / len(required_skills)) * 100
    return match_percentage, matched

def extract_experience_years(resume_text):
    """Extract years of experience from resume"""
    resume_lower = resume_text.lower()
    
    # Patterns for experience
    patterns = [
        r'(\d+)\+?\s*years?\s+(?:of\s+)?experience',
        r'experience[:\s]+(\d+)\+?\s*years?',
        r'(\d+)\+?\s*years?\s+(?:in|with)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, resume_lower)
        if match:
            return int(match.group(1))
    
    # Check for fresher/entry-level indicators
    if any(term in resume_lower for term in ['fresher', 'entry level', 'graduate', 'recent graduate']):
        return 0
    
    return None

def extract_education_level(resume_text):
    """Extract education level from resume"""
    resume_lower = resume_text.lower()
    
    education_scores = {
        'phd': 100,
        'ph.d': 100,
        'doctorate': 100,
        'masters': 85,
        'master': 85,
        'm.tech': 85,
        'mtech': 85,
        'm.s': 85,
        'ms in': 85,
        'mba': 80,
        'bachelor': 70,
        'b.tech': 70,
        'btech': 70,
        'b.e': 70,
        'b.s': 70,
        'bs in': 70,
        'undergraduate': 65,
        'diploma': 50,
        '12th': 30,
        'high school': 30
    }
    
    for degree, score in education_scores.items():
        if degree in resume_lower:
            return score
    
    return 50  # Default

def detect_relevant_projects(resume_text, job_description):
    """Detect if resume has projects/work experience relevant to job (generic)"""
    resume_lower = resume_text.lower()
    job_lower = job_description.lower()
    
    # Generic project/work indicators (works for all fields)
    project_indicators = [
        'project', 'built', 'developed', 'created', 'implemented', 'designed',
        'managed', 'led', 'coordinated', 'executed', 'delivered', 'achieved',
        'improved', 'optimized', 'established', 'launched', 'organized'
    ]
    
    has_projects = any(indicator in resume_lower for indicator in project_indicators)
    
    if not has_projects:
        return 0.0
    
    # Extract important keywords from job description (dynamic, not hardcoded)
    # Split job description into words, filter meaningful ones
    job_words = re.findall(r'\b[a-zA-Z]{3,}\b', job_lower)
    job_keywords = set(job_words) - set([
        'the', 'and', 'for', 'with', 'will', 'can', 'has', 'are', 'this',
        'that', 'from', 'have', 'you', 'our', 'we', 'their', 'been', 'about'
    ])
    
    # Count how many job keywords appear near project indicators in resume
    matched_count = 0
    total_keywords = len(job_keywords)
    
    if total_keywords == 0:
        return 50.0  # Neutral if no specific keywords
    
    for keyword in job_keywords:
        if keyword in resume_lower:
            matched_count += 1
    
    # Return percentage of matched keywords
    return min((matched_count / total_keywords) * 100, 100)

def detect_certifications(resume_text):
    """Detect relevant certifications in resume"""
    resume_lower = resume_text.lower()
    
    certifications = [
        'aws certified', 'azure certified', 'gcp certified', 'google cloud',
        'mongodb certified', 'oracle certified', 'cisco certified',
        'certified kubernetes', 'docker certified', 'red hat certified',
        'microsoft certified', 'comptia', 'cissp', 'pmp'
    ]
    
    found_certs = []
    for cert in certifications:
        if cert in resume_lower:
            found_certs.append(cert)
    
    return len(found_certs) > 0, found_certs

def calculate_experience_education_boost(resume_text, job_description):
    """Calculate boost based on experience and education (generic for all fields)"""
    boost_score = 0.0
    
    # Experience boost (0-20 points)
    years_exp = extract_experience_years(resume_text)
    job_lower = job_description.lower()
    
    if years_exp is not None:
        # Extract experience requirement from job description dynamically
        if re.search(r'\b0.?2\s*years?\b|\bfresher\b|\bentry.level\b', job_lower):
            # Entry level job
            if years_exp <= 2:
                boost_score += 20
            elif years_exp <= 4:
                boost_score += 12
            else:
                boost_score += 5  # Overqualified
        elif re.search(r'\b2.?5\s*years?\b|\bmid.?level\b', job_lower):
            # Mid level job
            if 2 <= years_exp <= 5:
                boost_score += 20
            elif years_exp < 2:
                boost_score += 10  # Underqualified
            else:
                boost_score += 12
        elif re.search(r'\b5\+?\s*years?\b|\bsenior\b', job_lower):
            # Senior level job
            if years_exp >= 5:
                boost_score += 20
            elif years_exp >= 3:
                boost_score += 12
            else:
                boost_score += 5
        else:
            # No specific requirement, give moderate boost
            boost_score += 10
    
    # Education boost (0-15 points) - Generic for all fields
    edu_score = extract_education_level(resume_text)
    
    # Check if job requires any degree
    if re.search(r'\bdegree\b|\bbachelor\b|\bmaster\b|\bgraduate\b|\beducation\b', job_lower):
        if edu_score >= 85:  # Master's or higher
            boost_score += 15
        elif edu_score >= 70:  # Bachelor's
            boost_score += 12
        elif edu_score >= 50:  # Diploma
            boost_score += 8
        else:
            boost_score += 3
    else:
        # No education requirement, give small boost for having degree
        if edu_score >= 70:
            boost_score += 5
    
    return min(boost_score, 35)  # Cap at 35

def calculate_keyword_boost(job_description, resume_text):
    """Calculate boost score based on important role keywords (generic for all fields)"""
    boost_score = 0.0
    job_lower = job_description.lower()
    resume_lower = resume_text.lower()
    
    # Generic role indicators (works across all industries)
    role_patterns = [
        # Experience level
        (r'\bfresher\b', r'\bfresher\b|\bgraduate\b|\bentry.level\b', 3),
        (r'\bentry.level\b', r'\bentry.level\b|\bfresher\b|\bgraduate\b', 3),
        (r'\bjunior\b', r'\bjunior\b|\b1.2.years\b|\bentry\b', 2),
        (r'\bmid.level\b', r'\bmid.level\b|\b3.5.years\b', 2),
        (r'\bsenior\b', r'\bsenior\b|\b5\+?.years\b|\blead\b', 2),
        
        # Management roles
        (r'\bmanager\b', r'\bmanager\b|\bmanagement\b|\bmanaged\b', 3),
        (r'\blead\b', r'\blead\b|\bleader\b|\bleading\b', 2),
        (r'\bdirector\b', r'\bdirector\b|\bdirect\b', 2),
        
        # General professional terms
        (r'\bprofessional\b', r'\bprofessional\b|\bexperienced\b', 1),
        (r'\bteam\b', r'\bteam\b|\bcollaborat\w+\b', 1),
    ]
    
    total_possible = 0
    matched_score = 0
    
    # Check each pattern
    for job_pattern, resume_pattern, weight in role_patterns:
        if re.search(job_pattern, job_lower):
            total_possible += weight
            if re.search(resume_pattern, resume_lower):
                matched_score += weight
    
    # Return boost as percentage (0-100)
    if total_possible > 0:
        boost_score = (matched_score / total_possible) * 100
    else:
        boost_score = 50.0  # Neutral if no patterns found
    
    return boost_score

def rank_resumes(job_description, resumes_data, required_skills=[]):
    """
    Rank resumes based on similarity to job description
    
    Args:
        job_description: str - The job description text
        resumes_data: list - List of dicts with 'id', 'text', and 'name'
        required_skills: list - List of required skill keywords
    
    Returns:
        list of dicts with ranking information
    """
    rankings = []
    
    for resume in resumes_data:
        # Get resume text from various possible field names
        resume_text = resume.get('text', '') or resume.get('parsedContent', '') or resume.get('content', '')
        if not resume_text:
            resume_text = str(resume.get('skills', [])) + ' ' + str(resume.get('experience', ''))
        
        # Try BERT similarity first (more accurate), fall back to TF-IDF
        bert_score = calculate_bert_similarity(job_description, resume_text)
        if bert_score is not None:
            semantic_score = bert_score
            semantic_method = 'BERT'
        else:
            semantic_score = calculate_tfidf_similarity(job_description, resume_text)
            semantic_method = 'TF-IDF'
        
        # Calculate skill match score
        skill_score, matched_skills = calculate_skill_match_score(resume_text, required_skills)
        
        # Calculate keyword boost (for role-specific terms)
        keyword_boost = calculate_keyword_boost(job_description, resume_text)
        
        # Calculate experience & education boost
        exp_edu_boost = calculate_experience_education_boost(resume_text, job_description)
        
        # Calculate project relevance
        project_score = detect_relevant_projects(resume_text, job_description)
        
        # Detect certifications
        has_certs, cert_list = detect_certifications(resume_text)
        cert_boost = 5 if has_certs else 0
        
        # Combined score with optimized weights (generic, works for all fields):
        # 65% Skills (most important - dynamically from job posting)
        # 12% Semantic similarity (BERT/TF-IDF - understands context)
        # 10% Experience/Education match
        # 8% Project/Work relevance
        # 3% Role keywords
        # 2% Certifications
        combined_score = (
            (skill_score / 100 * 0.65) +
            (semantic_score * 0.12) +
            (exp_edu_boost / 35 * 0.10) +
            (project_score / 100 * 0.08) +
            (keyword_boost / 100 * 0.03) +
            (cert_boost / 5 * 0.02)
        )
        
        rankings.append({
            'resume_id': resume.get('id') or resume.get('_id'),
            'candidate_name': resume.get('name', 'Unknown'),
            'match_percentage': combined_score * 100,
            'similarity_score': semantic_score * 100,
            'semantic_method': semantic_method,
            'tfidf_score': semantic_score * 100,
            'skill_match_score': skill_score,
            'keyword_boost': keyword_boost,
            'project_score': project_score,
            'experience_education_boost': exp_edu_boost,
            'certifications': cert_list if has_certs else [],
            'matched_skills': matched_skills,
            'missing_skills': list(set(required_skills) - set(matched_skills))
        })
    
    # Sort by match percentage (descending)
    rankings.sort(key=lambda x: x['match_percentage'], reverse=True)
    
    # Add rank numbers
    for idx, ranking in enumerate(rankings):
        ranking['rank'] = idx + 1
    
    return rankings
