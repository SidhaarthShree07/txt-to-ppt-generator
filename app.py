import os
import tempfile
import uuid
from flask import Flask, request, render_template, jsonify, send_file
from werkzeug.utils import secure_filename
import logging
from src.llm_providers import GeminiProvider
from src.ppt_analyzer import PowerPointAnalyzer
from src.slide_generator import SlideGenerator
from src.utils import validate_file, cleanup_temp_files

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

# Allowed file extensions
ALLOWED_EXTENSIONS = {'pptx', 'potx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """Render the main page"""
    return render_template('index.html')

@app.route('/api/generate', methods=['POST'])
def generate_presentation():
    """Main endpoint to generate PowerPoint presentation"""
    try:
        # Validate request
        if 'template' not in request.files:
            return jsonify({'error': 'No template file uploaded'}), 400
        
        file = request.files['template']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Only .pptx and .potx files are allowed'}), 400
        
        # Get form data
        text_content = request.form.get('text_content', '').strip()
        guidance = request.form.get('guidance', '').strip()
        api_key = request.form.get('api_key', '').strip()
        
        if not text_content:
            return jsonify({'error': 'Text content is required'}), 400
        
        if not api_key:
            return jsonify({'error': 'Gemini API key is required'}), 400
        
        # Generate unique session ID for file management
        session_id = str(uuid.uuid4())
        
        # Save uploaded template
        filename = secure_filename(file.filename)
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_template_{filename}")
        file.save(template_path)
        
        # Validate the PowerPoint file
        if not validate_file(template_path):
            os.remove(template_path)
            return jsonify({'error': 'Invalid PowerPoint file'}), 400
        
        # Initialize components
        gemini = GeminiProvider(api_key)
        analyzer = PowerPointAnalyzer()
        generator = SlideGenerator()
        
        # Process the presentation
        logger.info(f"Processing presentation for session {session_id}")
        
        # Step 1: Analyze template
        template_info = analyzer.analyze_template(template_path)
        
        # Step 2: Parse text content with Gemini
        slide_structure = gemini.parse_text_to_slides(text_content, guidance)
        
        # Step 3: Generate new presentation
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_generated.pptx")
        generator.create_presentation(slide_structure, template_info, output_path)
        
        # Clean up template file
        os.remove(template_path)
        
        # Return success response with download info
        return jsonify({
            'success': True,
            'download_url': f'/api/download/{session_id}',
            'session_id': session_id
        })
        
    except Exception as e:
        logger.error(f"Error generating presentation: {str(e)}")
        # Clean up any temp files
        cleanup_temp_files(locals().get('template_path'), locals().get('output_path'))
        return jsonify({'error': f'Failed to generate presentation: {str(e)}'}), 500

@app.route('/api/download/<session_id>')
def download_presentation(session_id):
    """Download the generated presentation"""
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_generated.pptx")
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found or expired'}), 404
        
        def cleanup_after_send():
            try:
                os.remove(file_path)
            except Exception:
                pass
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name='generated_presentation.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except Exception as e:
        logger.error(f"Error downloading file: {str(e)}")
        return jsonify({'error': 'Failed to download file'}), 500

@app.route('/api/health')
def health_check():
    """Health check endpoint"""
    return jsonify({'status': 'healthy'})

@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': 'File too large. Maximum size is 50MB.'}), 413

@app.errorhandler(500)
def internal_error(error):
    logger.error(f"Internal server error: {str(error)}")
    return jsonify({'error': 'Internal server error'}), 500

# Create upload directory if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Export app for Vercel
# Vercel will automatically detect this as the WSGI application

if __name__ == '__main__':
    # Run the app locally
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
