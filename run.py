#!/usr/bin/env python3
"""
Simple run script for the Text-to-PowerPoint Generator
"""

import os
import sys
from app import app
import logging
logging.basicConfig(level=logging.DEBUG)

def main():
    """Main entry point for the application"""
    print("Starting Text-to-PowerPoint Generator...")
    print("=" * 50)
    print("🚀 Web Application will be available at: http://localhost:5000")
    print("📝 Make sure you have a Gemini API key ready!")
    print("📁 Prepare a PowerPoint template file (.pptx or .potx)")
    print("=" * 50)
    print()
    
    # Set development mode
    if '--debug' in sys.argv or '-d' in sys.argv:
        app.run(debug=True, host='0.0.0.0', port=5000)
    elif '--prod' in sys.argv or '-p' in sys.argv:
        # Production mode with gunicorn would be better
        print("For production, use: gunicorn -w 4 -b 0.0.0.0:5000 app:app")
        app.run(debug=False, host='0.0.0.0', port=5000)
    else:
        # Default development mode
        app.run(debug=True, host='0.0.0.0', port=5000)

if __name__ == '__main__':
    main()
