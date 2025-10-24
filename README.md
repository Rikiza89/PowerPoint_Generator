# 🎯 RAG PowerPoint Generator

A sophisticated Flask-based application that generates professional PowerPoint presentations using RAG (Retrieval-Augmented Generation) with Ollama AI and document context.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Flask](https://img.shields.io/badge/Flask-3.0.0-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

## ✨ Features

### 🤖 AI-Powered Generation
- **RAG Integration**: Uses uploaded documents as context for accurate, relevant presentations
- **Ollama LLM**: Powered by llama3.2:3b model for intelligent content generation
- **Iterative Refinement**: Built-in feedback loop for continuous improvement

### 🎨 Professional Design
- **Multiple Color Schemes**: Corporate Blue, Modern Green, Elegant Purple
- **Various Slide Types**: Bullet points, two-column, numbered lists, charts, and tables
- **Advanced Layouts**: Professional styling with branded headers and consistent formatting
- **Data Visualization**: Bar charts, column charts, line charts, and pie charts

### 📊 Sophisticated Content
- **Chart Generation**: Automatic creation of data visualizations
- **Table Creation**: Styled tables with headers and alternating row colors
- **Multiple Layouts**: Flexible content presentation options
- **Professional Typography**: Calibri font with proper sizing and hierarchy

### 🔄 User-Friendly Workflow
1. Upload documents (PDF, DOCX, TXT)
2. Describe your presentation requirements
3. Review AI-generated structure
4. Provide feedback or confirm
5. Download professional PPTX file

## 📋 Prerequisites

- **Python 3.8 or higher**
- **Ollama** installed and running
- **llama3.2:3b** model downloaded

## 🚀 Installation

### 1. Install Ollama

Visit [https://ollama.ai](https://ollama.ai) and follow installation instructions for your OS.

```bash
# Pull the required model
ollama pull llama3.2:3b

# Verify installation
ollama run llama3.2:3b
# Type "Hello" to test, then /bye to exit
```

### 2. Clone or Create Project

```bash
# Create project directory
mkdir rag-pptx-generator
cd rag-pptx-generator
```

### 3. Set Up Python Environment

```bash
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

### 4. Create Project Structure

```
rag-pptx-generator/
│
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── README.md             # This file
├── LICENSE               # MIT License
├── templates/
│   └── index.html        # Web interface
├── uploads/              # Auto-created for uploaded documents
└── outputs/              # Auto-created for generated presentations
```

## 🎮 Usage

### Starting the Application

```bash
# Ensure virtual environment is activated
python app.py
```

The application will start on `http://127.0.0.1:5000`

### Web Interface Workflow

#### Step 1: Upload Documents
1. Click the upload area
2. Select PDF, DOCX, or TXT files
3. Click "Upload Documents"
4. Wait for processing confirmation

#### Step 2: Describe Your Presentation
Write a detailed description of what you want:

```
Create a 10-slide presentation about renewable energy. Include:
- Overview of solar, wind, and hydroelectric power
- Current statistics and market trends
- Environmental impact comparison
- Future projections with charts
- Case studies in table format
```

#### Step 3: Review Generated Structure
- Preview all slides and content
- Check layout types and data visualizations
- Review chart data and table structures

#### Step 4: Confirm or Refine
- **Option A**: Click "✓ Yes, Generate PPTX" to create the presentation
- **Option B**: Click "✗ No, Let me revise" to provide feedback

#### Step 5: Iterative Improvement (if needed)
1. Provide specific feedback about changes
2. Click "Submit Feedback & Regenerate"
3. Update your original request if needed
4. Click "Generate Presentation Structure" again
5. AI incorporates previous feedback for better results

#### Step 6: Download
Once confirmed, your professional PPTX file downloads automatically!

## 🎨 Color Schemes

Choose from three professional themes:

### Corporate Blue
- Primary: Dark Blue (#003366)
- Secondary: Blue (#0066CC)
- Accent: Orange (#FF9900)
- Perfect for business and corporate presentations

### Modern Green
- Primary: Forest Green (#228B22)
- Secondary: Lime Green (#32CD32)
- Accent: Gold (#FFD700)
- Ideal for environmental and sustainability topics

### Elegant Purple
- Primary: Indigo (#4B0082)
- Secondary: Medium Purple (#9370DB)
- Accent: Pink (#FFC0CB)
- Great for creative and innovative presentations

## 📊 Slide Types

### Bullet Slides
Traditional point-based content with professional bullets

### Two-Column Slides
Side-by-side content for comparisons or parallel information

### Numbered Slides
Sequential steps or prioritized information

### Chart Slides
Data visualizations including:
- **Bar Charts**: Horizontal comparisons
- **Column Charts**: Vertical comparisons
- **Line Charts**: Trends over time
- **Pie Charts**: Proportions and percentages

### Table Slides
Structured data with:
- Styled headers
- Alternating row colors
- Professional formatting

## ⚙️ Configuration

### Change Ollama Model

Edit `app.py`:
```python
MODEL_NAME = "llama3.2:3b"  # Change to any Ollama model
```

Available alternatives:
- `llama3.2:1b` - Faster, less accurate
- `llama3.1:8b` - More accurate, slower
- `mistral:7b` - Alternative model

### Adjust Document Processing

```python
# In chunk_text() function
def chunk_text(text, chunk_size=500, overlap=50):
    # Increase chunk_size for longer contexts
    # Increase overlap for better continuity
```

### Modify Retrieved Context

```python
# In generate_presentation() function
results = collection.query(
    query_texts=[user_request],
    n_results=5  # Increase for more context
)
```

### Temperature Control

```python
# In query_ollama() function
payload = {
    "temperature": 0.7  # Lower = focused, Higher = creative
}
```

## 🔧 Troubleshooting

### Ollama Connection Error
```
Error: Connection refused to localhost:11434
```
**Solution**: Start Ollama service
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
The application uses a fallback structure if JSON parsing fails.
**Solutions**:
- Be more specific in your request
- Try regenerating
- Lower the temperature in configuration

### File Upload Issues
- **Maximum file size**: 16MB
- **Supported formats**: PDF, DOCX, TXT only
- **Check**: File permissions and disk space

### ChromaDB Errors
```bash
pip uninstall chromadb
pip install chromadb==0.4.22
```

## 💡 Tips for Best Results

### Document Upload
- Upload relevant, high-quality documents
- More context leads to better presentations
- Keep documents focused on your topic

### Request Writing
✅ **Good Example**:
```
Create a 12-slide presentation about climate change impacts on 
agriculture. Include:
- Introduction with statistics
- Regional impact analysis in table format
- Crop yield comparisons in bar charts
- Adaptation strategies with case studies
- Future projections with line charts
```

❌ **Poor Example**:
```
Make a presentation about climate change
```

### Feedback
- Be specific about what to change
- Mention slide numbers or titles
- Suggest concrete improvements
- Don't hesitate to iterate multiple times

## 📚 Example Workflow

```
1. Upload: 3 research papers on artificial intelligence

2. Request: "Create a 15-slide presentation about AI in healthcare. 
   Include current applications, benefits analysis in a table, 
   adoption rates chart, challenges, and future trends."

3. Review: Structure looks good but needs more specific data

4. Feedback: "Add specific statistics to slides 4-6, include a 
   comparison chart of AI vs traditional methods, and add a case 
   study table about successful implementations"

5. Regenerate: AI incorporates feedback with improved content

6. Confirm: Generate final professional PPTX
```

## 🚀 Production Deployment

### Use Production WSGI Server
```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

### Environment Variables
```python
import os
OLLAMA_URL = os.getenv('OLLAMA_URL', 'http://localhost:11434/api/generate')
SECRET_KEY = os.getenv('SECRET_KEY', 'your-secret-key')
```

### Add Logging
```python
import logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
```

### File Cleanup
Implement scheduled cleanup for old files:
```python
import schedule
import time

def cleanup_old_files():
    # Remove files older than 24 hours
    pass

schedule.every(1).hours.do(cleanup_old_files)
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **Ollama** - For providing the LLM infrastructure
- **Flask** - Web framework
- **python-pptx** - PowerPoint generation library
- **ChromaDB** - Vector database for RAG
- **Meta** - Llama 3.2 model

## 📧 Support

For issues, questions, or suggestions:
- Open an issue on GitHub
- Check existing documentation
- Review troubleshooting section

## 🔗 Resources

- [Ollama Documentation](https://ollama.ai/docs)
- [Flask Documentation](https://flask.palletsprojects.com/)
- [python-pptx Documentation](https://python-pptx.readthedocs.io/)
- [ChromaDB Documentation](https://docs.trychroma.com/)

## 📊 Project Status

**Current Version**: 1.0.0

**Status**: Active Development

**Features in Development**:
- [ ] Image insertion in slides
- [ ] Custom template support
- [ ] Animation effects
- [ ] Video embedding
- [ ] Cloud storage integration
- [ ] Multi-user support
- [ ] API endpoints

---

**Made with ❤️ using Python, Flask, and AI**

**Star ⭐ this repository if you find it helpful!**
