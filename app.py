from flask import Flask, render_template, request, jsonify, send_file
import os
import json
from werkzeug.utils import secure_filename
import requests
from pathlib import Path
import PyPDF2
import docx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import chromadb
from chromadb.utils import embedding_functions
import uuid

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Ensure folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Initialize ChromaDB
chroma_client = chromadb.Client()
collection = None

# Ollama configuration
OLLAMA_URL = "http://localhost:11434/api/generate"
MODEL_NAME = "llama3.2:3b"

# Session storage for conversation history
sessions = {}

def extract_text_from_pdf(file_path):
    """Extract text from PDF file"""
    text = ""
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
    return text

def extract_text_from_docx(file_path):
    """Extract text from DOCX file"""
    doc = docx.Document(file_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text + "\n"
    return text

def extract_text_from_txt(file_path):
    """Extract text from TXT file"""
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()

def chunk_text(text, chunk_size=500, overlap=50):
    """Split text into overlapping chunks"""
    words = text.split()
    chunks = []
    for i in range(0, len(words), chunk_size - overlap):
        chunk = ' '.join(words[i:i + chunk_size])
        chunks.append(chunk)
    return chunks

def query_ollama(prompt, context=""):
    """Query Ollama model"""
    full_prompt = f"{context}\n\n{prompt}" if context else prompt
    
    payload = {
        "model": MODEL_NAME,
        "prompt": full_prompt,
        "stream": False,
        "temperature": 0.7
    }
    
    try:
        response = requests.post(OLLAMA_URL, json=payload)
        response.raise_for_status()
        return response.json()['response']
    except Exception as e:
        return f"Error querying Ollama: {str(e)}"

def create_presentation(presentation_data):
    """Create PowerPoint presentation from structured data"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Title slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = presentation_data.get('title', 'Presentation')
    subtitle.text = presentation_data.get('subtitle', '')
    
    # Content slides
    for slide_data in presentation_data.get('slides', []):
        bullet_slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(bullet_slide_layout)
        
        title = slide.shapes.title
        title.text = slide_data.get('title', '')
        
        body_shape = slide.placeholders[1]
        tf = body_shape.text_frame
        
        for point in slide_data.get('points', []):
            p = tf.add_paragraph()
            p.text = point
            p.level = 0
    
    return prs

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_documents():
    """Upload and process documents"""
    global collection
    
    if 'files' not in request.files:
        return jsonify({'error': 'No files provided'}), 400
    
    files = request.files.getlist('files')
    
    if not files:
        return jsonify({'error': 'No files selected'}), 400
    
    # Reset collection
    try:
        chroma_client.delete_collection(name="documents")
    except:
        pass
    
    collection = chroma_client.create_collection(
        name="documents",
        metadata={"hnsw:space": "cosine"}
    )
    
    all_chunks = []
    chunk_ids = []
    
    for file in files:
        if file.filename == '':
            continue
        
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Extract text based on file type
        if filename.endswith('.pdf'):
            text = extract_text_from_pdf(file_path)
        elif filename.endswith('.docx'):
            text = extract_text_from_docx(file_path)
        elif filename.endswith('.txt'):
            text = extract_text_from_txt(file_path)
        else:
            continue
        
        # Chunk the text
        chunks = chunk_text(text)
        all_chunks.extend(chunks)
        chunk_ids.extend([f"{filename}_{i}" for i in range(len(chunks))])
    
    # Add to ChromaDB
    if all_chunks:
        collection.add(
            documents=all_chunks,
            ids=chunk_ids
        )
    
    return jsonify({
        'success': True,
        'message': f'Processed {len(files)} files with {len(all_chunks)} chunks'
    })

@app.route('/generate_presentation', methods=['POST'])
def generate_presentation():
    """Generate presentation structure based on user request"""
    data = request.json
    user_request = data.get('request', '')
    session_id = data.get('session_id', str(uuid.uuid4()))
    
    if not user_request:
        return jsonify({'error': 'No request provided'}), 400
    
    # Initialize session if needed
    if session_id not in sessions:
        sessions[session_id] = {
            'history': [],
            'iterations': 0
        }
    
    # Retrieve relevant context from documents
    context = ""
    if collection:
        results = collection.query(
            query_texts=[user_request],
            n_results=5
        )
        if results['documents']:
            context = "\n\n".join(results['documents'][0])
    
    # Add previous iterations to context
    previous_context = ""
    if sessions[session_id]['history']:
        previous_context = "\n\nPrevious presentation attempts:\n"
        for i, hist in enumerate(sessions[session_id]['history'], 1):
            previous_context += f"\nAttempt {i}:\n{json.dumps(hist['structure'], indent=2)}\n"
            previous_context += f"User feedback: {hist['feedback']}\n"
    
    # Generate presentation structure
    prompt = f"""Based on the following context and user request, create a PowerPoint presentation structure.

Context from documents:
{context}

{previous_context}

User request: {user_request}

Please provide a JSON structure for the presentation with the following format:
{{
    "title": "Main presentation title",
    "subtitle": "Subtitle or tagline",
    "slides": [
        {{
            "title": "Slide title",
            "points": ["Point 1", "Point 2", "Point 3"]
        }}
    ]
}}

Provide ONLY the JSON structure, no additional text."""
    
    response = query_ollama(prompt, context)
    
    # Extract JSON from response
    try:
        # Try to find JSON in the response
        start_idx = response.find('{')
        end_idx = response.rfind('}') + 1
        json_str = response[start_idx:end_idx]
        presentation_structure = json.loads(json_str)
    except:
        # Fallback structure
        presentation_structure = {
            "title": "Generated Presentation",
            "subtitle": "Based on your request",
            "slides": [
                {
                    "title": "Overview",
                    "points": ["Point 1", "Point 2", "Point 3"]
                }
            ]
        }
    
    sessions[session_id]['iterations'] += 1
    
    return jsonify({
        'session_id': session_id,
        'structure': presentation_structure,
        'iteration': sessions[session_id]['iterations']
    })

@app.route('/confirm_presentation', methods=['POST'])
def confirm_presentation():
    """Handle user confirmation and generate or refine presentation"""
    data = request.json
    confirmed = data.get('confirmed', False)
    session_id = data.get('session_id')
    structure = data.get('structure')
    feedback = data.get('feedback', '')
    
    if not session_id or session_id not in sessions:
        return jsonify({'error': 'Invalid session'}), 400
    
    if confirmed:
        # Generate the actual PPTX file
        try:
            prs = create_presentation(structure)
            output_filename = f"presentation_{session_id}.pptx"
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
            prs.save(output_path)
            
            return jsonify({
                'success': True,
                'message': 'Presentation generated successfully',
                'download_url': f'/download/{output_filename}'
            })
        except Exception as e:
            return jsonify({'error': f'Error generating presentation: {str(e)}'}), 500
    else:
        # Store feedback for next iteration
        sessions[session_id]['history'].append({
            'structure': structure,
            'feedback': feedback
        })
        
        return jsonify({
            'success': True,
            'message': 'Feedback recorded. Please submit a new generation request.'
        })

@app.route('/download/<filename>')
def download_file(filename):
    """Download generated presentation"""
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return jsonify({'error': 'File not found'}), 404

if __name__ == '__main__':
    app.run(debug=True, port=5000)
