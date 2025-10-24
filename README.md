# RAG PowerPoint Generator - Setup Guide

## Prerequisites

1. **Python 3.8+** installed on your system
2. **Ollama** installed and running with llama3.2:3b model

## Installation Steps

### 1. Install Ollama and Model

```bash
# Install Ollama (visit https://ollama.ai for OS-specific instructions)

# Pull the llama3.2:3b model
ollama pull llama3.2:3b

# Verify it's running
ollama run llama3.2:3b
# Type "Hello" to test, then /bye to exit
```

### 2. Set Up Python Environment

```bash
# Create project directory
mkdir rag-pptx-generator
cd rag-pptx-generator

# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 3. Project Structure

Create the following folder structure:

```
rag-pptx-generator/
│
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Web interface
├── uploads/              # Uploaded documents (auto-created)
└── outputs/              # Generated presentations (auto-created)
```

### 4. Create Files

1. Save the Flask application code as `app.py`
2. Create a `templates` folder and save the HTML code as `templates/index.html`
3. Save the requirements.txt file

## Running the Application

### 1. Start Ollama Service

Make sure Ollama is running in the background:

```bash
# The Ollama service should be running automatically
# If not, start it manually (OS-dependent)
```

### 2. Start Flask Application

```bash
# Make sure virtual environment is activated
python app.py
```

The application will start on `http://127.0.0.1:5000`

### 3. Access Web Interface

Open your browser and navigate to:
```
http://127.0.0.1:5000
```

## How to Use

### Step 1: Upload Documents
1. Click on the upload area
2. Select one or more documents (PDF, DOCX, or TXT files)
3. Click "Upload Documents" button
4. Wait for confirmation message

### Step 2: Describe Your Presentation
1. In the text area, describe what you want in your presentation
2. Be specific about:
   - Topic
   - Number of slides
   - Key points to cover
   - Any specific structure you want

**Example requests:**
```
Create a 10-slide presentation about renewable energy sources, 
including solar, wind, and hydroelectric power. Include statistics 
and future trends.
```

```
Make a 5-slide presentation summarizing the key findings from 
the uploaded research papers about machine learning in healthcare.
```

### Step 3: Review Generated Structure
1. Click "Generate Presentation Structure"
2. Wait for the AI to generate the structure
3. Review the preview showing all slides and content

### Step 4: Confirm or Revise
1. If satisfied, click "✓ Yes, Generate PPTX"
   - The PowerPoint file will be generated and downloaded automatically
   
2. If not satisfied, click "✗ No, Let me revise"
   - Provide specific feedback about what to change
   - Click "Submit Feedback & Regenerate"
   - Update your request in Step 2 if needed
   - Click "Generate Presentation Structure" again
   - The AI will use your previous feedback to improve

### Iterative Refinement Process

The application supports multiple iterations:
1. Generate initial structure
2. Review and provide feedback
3. Generate improved structure (AI learns from previous attempts)
4. Repeat until satisfied
5. Generate final PPTX file

## Features

### RAG (Retrieval-Augmented Generation)
- Uploads documents are processed and chunked
- Relevant content is retrieved based on your request
- AI uses this context to generate accurate presentations

### Document Support
- **PDF**: Research papers, reports, books
- **DOCX**: Word documents, essays
- **TXT**: Plain text files, notes

### Iterative Refinement
- Multiple generation attempts
- AI remembers previous structures and feedback
- Continuous improvement until you're satisfied

### Session Management
- Each generation session is tracked
- Previous attempts inform new generations
- Feedback history is maintained

## Configuration

### Changing Ollama Model

In `app.py`, modify:
```python
MODEL_NAME = "llama3.2:3b"  # Change to any Ollama model
```

Available models:
- `llama3.2:1b` - Faster, less accurate
- `llama3.2:3b` - Balanced (recommended)
- `llama3.1:8b` - More accurate, slower

### Adjusting Chunk Size

For document processing:
```python
def chunk_text(text, chunk_size=500, overlap=50):
    # Increase chunk_size for longer contexts
    # Increase overlap for better continuity
```

### Changing Number of Retrieved Chunks

```python
results = collection.query(
    query_texts=[user_request],
    n_results=5  # Increase for more context
)
```

### Temperature Control

For creativity vs accuracy:
```python
payload = {
    "model": MODEL_NAME,
    "prompt": full_prompt,
    "stream": False,
    "temperature": 0.7  # Lower = more focused, Higher = more creative
}
```

## Troubleshooting

### Ollama Connection Error
```
Error: Connection refused to localhost:11434
```
**Solution**: Make sure Ollama is running
```bash
ollama serve
```

### Model Not Found
```
Error: model 'llama3.2:3b' not found
```
**Solution**: Pull the model
```bash
ollama pull llama3.2:3b
```

### JSON Parsing Error
If the AI returns malformed JSON, the application uses a fallback structure. To improve:
- Make your request more specific
- Try regenerating
- Lower the temperature in configuration

### File Upload Issues
- Maximum file size: 16MB
- Supported formats: PDF, DOCX, TXT only
- Check file permissions

### ChromaDB Errors
```bash
# Reinstall ChromaDB if needed
pip uninstall chromadb
pip install chromadb==0.4.22
```

## Tips for Best Results

### 1. Document Upload
- Upload relevant documents before generating
- More context = better presentations
- Keep documents focused on your topic

### 2. Request Writing
- Be specific about number of slides
- Mention key topics to cover
- Specify any required structure
- Include desired depth of content

### 3. Feedback
- Be specific in your feedback
- Mention exact slides or points to change
- Suggest improvements clearly
- Don't hesitate to iterate multiple times

### 4. Example Workflow
```
Upload: 3 research papers on climate change

Request: "Create an 8-slide presentation about climate change 
impacts on agriculture. Include current data, regional effects, 
and adaptation strategies."

Review: Structure looks good but needs more data

Feedback: "Add specific statistics to slides 3 and 4, and 
include a case study slide about drought-resistant crops"

Regenerate: Now includes requested improvements

Confirm: Generate final PPTX
```

## Advanced Usage

### Custom Prompt Engineering

Modify the prompt in `app.py` for different styles:

```python
prompt = f"""Based on the following context, create a PROFESSIONAL 
and DETAILED PowerPoint presentation structure.

Use BULLET POINTS for clarity.
Include ACTIONABLE insights.
Add relevant STATISTICS when available.

Context: {context}
Request: {user_request}

[Rest of prompt...]
"""
```

### Styling Presentations

Modify `create_presentation()` function:

```python
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN

# Add styling to text
for paragraph in text_frame.paragraphs:
    paragraph.font.size = Pt(18)
    paragraph.font.name = 'Arial'
    paragraph.alignment = PP_ALIGN.LEFT
```

### Adding Images

```python
# In create_presentation()
from pptx.util import Inches

# Add image to slide
slide.shapes.add_picture('image.jpg', 
                        Inches(1), Inches(2), 
                        width=Inches(4))
```

## Security Considerations

- Files are stored temporarily in `uploads/` folder
- Consider adding file validation
- Implement user authentication for production
- Add rate limiting for API calls
- Sanitize file names and inputs

## Production Deployment

For production use:

1. **Use a production WSGI server**:
```bash
pip install gunicorn
gunicorn -w 4 app:app
```

2. **Add environment variables**:
```python
import os
OLLAMA_URL = os.getenv('OLLAMA_URL', 'http://localhost:11434/api/generate')
```

3. **Implement proper logging**:
```python
import logging
logging.basicConfig(level=logging.INFO)
```

4. **Add file cleanup**:
```python
import schedule
def cleanup_old_files():
    # Remove files older than 24 hours
    pass
```

## Support and Resources

- **Ollama Documentation**: https://ollama.ai/docs
- **Flask Documentation**: https://flask.palletsprojects.com/
- **Python-PPTX Documentation**: https://python-pptx.readthedocs.io/
- **ChromaDB Documentation**: https://docs.trychroma.com/

## License

This is a demonstration project. Modify and use as needed for your purposes.

---

**Happy Presentation Generating! 🎯**
