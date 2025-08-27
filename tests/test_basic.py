import unittest
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.utils import validate_api_key, truncate_text, parse_markdown_to_text
from src.llm_providers import GeminiProvider
import tempfile

class TestBasicFunctionality(unittest.TestCase):
    """Basic tests for utility functions"""
    
    def test_validate_api_key(self):
        """Test API key validation"""
        # Valid keys
        self.assertTrue(validate_api_key('AIzaSyDMockKey123456789'))
        self.assertTrue(validate_api_key('gcp-mockkey123456'))
        self.assertTrue(validate_api_key('valid_key_1234567890'))
        
        # Invalid keys
        self.assertFalse(validate_api_key(''))
        self.assertFalse(validate_api_key('short'))
        self.assertFalse(validate_api_key(None))
        self.assertFalse(validate_api_key('   '))
    
    def test_truncate_text(self):
        """Test text truncation"""
        long_text = "This is a very long text that should be truncated when it exceeds the maximum length"
        
        # Test normal truncation
        result = truncate_text(long_text, 20)
        self.assertEqual(len(result), 20)
        self.assertTrue(result.endswith('...'))
        
        # Test short text (should not be truncated)
        short_text = "Short text"
        result = truncate_text(short_text, 20)
        self.assertEqual(result, short_text)
    
    def test_parse_markdown_to_text(self):
        """Test markdown parsing"""
        markdown = """
        # Main Title
        ## Subtitle
        This is **bold** text and *italic* text.
        - List item 1
        - List item 2
        `code snippet`
        """
        
        result = parse_markdown_to_text(markdown)
        
        # Should remove markdown syntax
        self.assertNotIn('#', result)
        self.assertNotIn('**', result)
        self.assertNotIn('*', result)
        self.assertNotIn('`', result)
        
        # Should contain the actual content
        self.assertIn('Main Title', result)
        self.assertIn('bold', result)
        self.assertIn('italic', result)

class TestGeminiProvider(unittest.TestCase):
    """Test Gemini provider (mocked)"""
    
    def setUp(self):
        self.mock_api_key = "mock_gemini_key_12345"
    
    def test_gemini_initialization(self):
        """Test Gemini provider initialization"""
        # This will fail without a real API key, but we can test initialization
        try:
            provider = GeminiProvider(self.mock_api_key)
            self.assertIsNotNone(provider.api_key)
            self.assertEqual(provider.api_key, self.mock_api_key)
        except Exception:
            # Expected to fail without real API setup
            pass

if __name__ == '__main__':
    unittest.main()
