import os
import PyPDF2
import docx
from pathlib import Path

def extract_text_from_pdf(file_path):
    """Extract text from PDF file"""
    text = ""
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
    except Exception as e:
        pass
    return text.strip()

def extract_text_from_docx(file_path):
    """Extract text from DOCX file"""
    text = ""
    try:
        doc = docx.Document(file_path)
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        pass
    return text.strip()

def extract_text(file_path):
    """Extract text based on file extension"""
    file_path = Path(file_path)
    extension = file_path.suffix.lower()
    
    if extension == '.pdf':
        return extract_text_from_pdf(str(file_path))
    elif extension in ['.docx', '.doc']:
        return extract_text_from_docx(str(file_path))
    else:
        raise ValueError(f"Unsupported file format: {extension}")

def clean_text(text):
    """Clean and normalize text"""
    # Remove extra whitespace
    text = ' '.join(text.split())
    return text.lower()
