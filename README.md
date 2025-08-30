# Text-to-PowerPoint Generator

A web application that automatically converts text or markdown into professionally formatted PowerPoint presentations using your own templates.

## Features

- **Intelligent Text Parsing**: Uses LLM APIs to break down large text blocks into structured slide content
- **Template Style Preservation**: Automatically applies colors, fonts, layouts, and images from uploaded PowerPoint templates
- **Google Gemini Integration**: Uses Google's Gemini Pro model for intelligent text parsing
- **Secure Processing**: API keys are never stored or logged
- **Professional Output**: Generates downloadable .pptx files matching your template's look and feel

## How It Works

1. **Input**: Paste your text content and optionally provide guidance (e.g., "investor pitch deck")
2. **API Key**: Enter your Google Gemini API key (free tier available)
3. **Template Upload**: Upload a PowerPoint template (.pptx or .potx) file
4. **Processing**: The app analyzes your template's style and uses Gemini AI to structure your content
5. **Output**: Download a new presentation with your content styled to match your template

## Quick Deploy

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://github.com/SidhaarthShree07/txt-to-ppt-generator)

## Technology Stack

- **Backend**: Python Flask with python-pptx for PowerPoint manipulation
- **Frontend**: HTML, CSS, JavaScript with Bootstrap for responsive design
- **AI Integration**: Google Gemini Pro for intelligent text processing
- **File Processing**: Secure temporary file handling with automatic cleanup
- **Deployment**: Ready for Vercel, Railway, Heroku, and other platforms

## Setup and Installation

### Prerequisites
- Python 3.8+
- pip package manager

### Installation

1. Clone the repository:
```bash
git clone <https://github.com/SidhaarthShree07/txt-to-ppt-generator>
cd text-to-powerpoint
```

2. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the application:
```bash
python app.py
```

5. Open your browser and navigate to `http://localhost:5000`

## Usage

1. **Enter Your Content**: Paste your text, markdown, or prose into the main text area
2. **Add Guidance** (Optional): Provide a brief description like "sales presentation" or "technical overview"
3. **Choose LLM Provider**: Select your preferred AI provider and enter your API key
4. **Upload Template**: Choose a PowerPoint template file that defines your desired style
5. **Generate**: Click "Generate Presentation" and wait for processing
6. **Download**: Your new presentation will be ready for download

## API Provider

- **Google Gemini**: Gemini Pro model with free tier available
- **Get your API key**: [Google AI Studio](https://makersuite.google.com/app/apikey)

## File Format Support

- **Input Templates**: .pptx, .potx files
- **Output**: .pptx files compatible with PowerPoint and other presentation software

## Security and Privacy

- API keys are processed in memory only and never stored
- Uploaded files are temporarily processed and automatically deleted
- No user data is logged or retained

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

## License

This project is licensed under the MIT License - see the [LICENSE](https://github.com/SidhaarthShree07/txt-to-ppt-generator/blob/main/LICENSE) file for details.

## Architecture Overview

### Text Processing Pipeline
The application uses a multi-step approach to convert text into slides:

1. **Content Analysis**: LLM analyzes input text to identify main topics, supporting points, and logical flow
2. **Slide Structure**: Content is organized into title slides, content slides, and conclusion slides
3. **Template Mapping**: Each piece of content is mapped to appropriate template layouts

### Style Application System
Template styles are preserved through:

1. **Style Extraction**: Colors, fonts, and formatting are extracted from the uploaded template
2. **Layout Analysis**: Slide layouts and placeholders are identified and cataloged
3. **Asset Reuse**: Images and design elements from templates are reused appropriately
4. **Dynamic Application**: Extracted styles are applied to generated content maintaining visual consistency

## Development

### Project Structure
```
text-to-powerpoint/
├── app.py                      # Main Flask application (all endpoints, session management, PDF preview)
├── run.py                      # Alternate entry point for running the app
├── requirements.txt            # Python dependencies
├── static/                     # CSS, JS, and static assets
│   └── css/
│       └── style.css
├── templates/                  # HTML templates
│   └── index.html
├── src/                        # Core application modules
│   ├── content_mapper.py           # Maps AI content to best-matching template slides
│   ├── format_detector.py          # Detects content formats and placeholder capacities
│   ├── llm_providers.py            # LLM API integrations (Gemini, OpenAI, AIPipe)
│   ├── multi_placeholder_handler.py # Handles content for slides with multiple placeholders
│   ├── ppt_analyzer.py             # PowerPoint template analysis and asset extraction
│   ├── robust_pipeline.py          # Orchestrates robust, multi-step slide generation
│   ├── simple_slide_replacer.py    # Replaces placeholder text with generated content
│   ├── slide_generator.py          # Slide creation logic and enforcement of slide count
│   ├── slide_refiner.py            # Refines content to match placeholder requirements
│   ├── smart_mapper.py             # Matches AI content to template slides based on format
│   └── utils.py                    # File validation, cleanup, and helper functions
├── IMPLEMENTATION.md           # Technical implementation details
├── DEPLOYMENT.md               # Deployment instructions
├── README.md                   # This file
└── VERCEL_DEPLOY.md            # Vercel-specific deployment guide
```

## Limitations

- Maximum file size: 50MB for template uploads
- Processing time varies based on content length and LLM provider response time
- Some complex template layouts may not be perfectly preserved
- No support for animated elements or complex multimedia

## Future Enhancements

- Speaker notes generation
- Multiple presentation style modes
- Batch processing capabilities
- Advanced layout preservation
