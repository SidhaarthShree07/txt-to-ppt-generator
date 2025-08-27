import os
import logging
import tempfile
from typing import Optional, Union
from pptx import Presentation

logger = logging.getLogger(__name__)

def validate_file(file_path: str) -> bool:
    """
    Validate that a file is a proper PowerPoint file
    
    Args:
        file_path: Path to the file to validate
        
    Returns:
        True if valid PowerPoint file, False otherwise
    """
    try:
        if not os.path.exists(file_path):
            return False
        
        # Check file size (not too large, not empty)
        file_size = os.path.getsize(file_path)
        if file_size == 0 or file_size > 50 * 1024 * 1024:  # 50MB limit
            return False
        
        # Try to load with python-pptx
        presentation = Presentation(file_path)
        
        # Basic validation - should have at least slide layouts
        if len(presentation.slide_layouts) == 0:
            return False
        
        return True
        
    except Exception as e:
        logger.warning(f"File validation failed for {file_path}: {e}")
        return False

def cleanup_temp_files(*file_paths: Optional[str]) -> None:
    """
    Clean up temporary files
    
    Args:
        *file_paths: Variable number of file paths to clean up
    """
    for file_path in file_paths:
        if file_path and isinstance(file_path, str):
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    logger.debug(f"Cleaned up temp file: {file_path}")
            except Exception as e:
                logger.warning(f"Could not clean up file {file_path}: {e}")

def get_file_extension(filename: str) -> str:
    """
    Get the file extension from a filename
    
    Args:
        filename: The filename to extract extension from
        
    Returns:
        File extension in lowercase (without the dot)
    """
    if '.' in filename:
        return filename.rsplit('.', 1)[1].lower()
    return ''

def is_allowed_file_type(filename: str) -> bool:
    """
    Check if the file type is allowed for upload
    
    Args:
        filename: The filename to check
        
    Returns:
        True if file type is allowed, False otherwise
    """
    allowed_extensions = {'pptx', 'potx'}
    return get_file_extension(filename) in allowed_extensions

def sanitize_filename(filename: str) -> str:
    """
    Sanitize a filename for safe storage
    
    Args:
        filename: The filename to sanitize
        
    Returns:
        Sanitized filename
    """
    # Remove path components
    filename = os.path.basename(filename)
    
    # Replace potentially dangerous characters
    dangerous_chars = '<>:"/\\|?*'
    for char in dangerous_chars:
        filename = filename.replace(char, '_')
    
    # Limit length
    if len(filename) > 255:
        name, ext = os.path.splitext(filename)
        filename = name[:255-len(ext)] + ext
    
    return filename

def create_temp_file(suffix: str = '.tmp') -> str:
    """
    Create a temporary file and return its path
    
    Args:
        suffix: File suffix/extension
        
    Returns:
        Path to the temporary file
    """
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)  # Close the file descriptor
    return path

def format_file_size(size_bytes: int) -> str:
    """
    Format file size in human-readable format
    
    Args:
        size_bytes: Size in bytes
        
    Returns:
        Formatted size string
    """
    if size_bytes == 0:
        return "0 B"
    
    size_names = ["B", "KB", "MB", "GB"]
    import math
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return f"{s} {size_names[i]}"

def validate_api_key(api_key: str) -> bool:
    """
    Basic validation for API key format
    
    Args:
        api_key: The API key to validate
        
    Returns:
        True if key appears valid, False otherwise
    """
    if not api_key or not isinstance(api_key, str):
        return False
    
    # Remove whitespace
    api_key = api_key.strip()
    
    # Check minimum length
    if len(api_key) < 10:
        return False
    
    # Basic format checks for common API key patterns
    # Gemini keys typically start with specific patterns
    if api_key.startswith('AIza') or api_key.startswith('gcp-'):
        return True
    
    # More lenient check - contains alphanumeric characters
    if any(c.isalnum() for c in api_key):
        return True
    
    return False

def truncate_text(text: str, max_length: int = 100) -> str:
    """
    Truncate text to a maximum length with ellipsis
    
    Args:
        text: Text to truncate
        max_length: Maximum length
        
    Returns:
        Truncated text with ellipsis if needed
    """
    if len(text) <= max_length:
        return text
    
    return text[:max_length-3] + "..."

def parse_markdown_to_text(markdown_text: str) -> str:
    """
    Basic markdown parsing to plain text
    
    Args:
        markdown_text: Markdown formatted text
        
    Returns:
        Plain text version
    """
    # Remove markdown headers
    lines = markdown_text.split('\n')
    processed_lines = []
    
    for line in lines:
        # Remove header markers
        line = line.lstrip('#').strip()
        
        # Remove bold/italic markers (basic)
        line = line.replace('**', '').replace('*', '')
        line = line.replace('__', '').replace('_', '')
        
        # Remove code markers
        line = line.replace('`', '')
        
        if line:  # Only keep non-empty lines
            processed_lines.append(line)
    
    return '\n'.join(processed_lines)

def estimate_processing_time(text_length: int) -> int:
    """
    Estimate processing time in seconds based on text length
    
    Args:
        text_length: Length of input text
        
    Returns:
        Estimated processing time in seconds
    """
    # Base processing time
    base_time = 10  # 10 seconds minimum
    
    # Additional time per 1000 characters
    additional_time = (text_length // 1000) * 2
    
    # Cap at 60 seconds
    total_time = min(base_time + additional_time, 60)
    
    return total_time
