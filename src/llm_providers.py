import json
import logging
from typing import List, Dict, Any
import google.generativeai as genai

logger = logging.getLogger(__name__)

class GeminiProvider:
    """Google Gemini API provider for text parsing and slide generation"""
    
    def __init__(self, api_key: str):
        """Initialize Gemini provider with API key"""
        self.api_key = api_key
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel('gemini-pro')
    
    def parse_text_to_slides(self, text_content: str, guidance: str = "") -> List[Dict[str, Any]]:
        """
        Parse input text into structured slide content using Gemini
        
        Args:
            text_content: The input text to be converted
            guidance: Optional guidance for tone/structure
            
        Returns:
            List of slide dictionaries with structure and content
        """
        try:
            # Build the prompt
            prompt = self._build_prompt(text_content, guidance)
            
            # Generate content with Gemini
            logger.info("Sending request to Gemini API")
            response = self.model.generate_content(prompt)
            
            # Parse the response
            slide_structure = self._parse_response(response.text)
            
            logger.info(f"Generated {len(slide_structure)} slides")
            return slide_structure
            
        except Exception as e:
            logger.error(f"Error with Gemini API: {str(e)}")
            raise Exception(f"Failed to process text with Gemini: {str(e)}")
    
    def _build_prompt(self, text_content: str, guidance: str) -> str:
        """Build the prompt for Gemini API"""
        
        base_prompt = """
You are an expert presentation designer. Your task is to analyze the provided text and convert it into a structured PowerPoint presentation.

INSTRUCTIONS:
1. Break down the text into logical slides
2. Each slide should have a clear purpose and focused content
3. Create an appropriate number of slides (typically 5-15 depending on content length)
4. Include a title slide and conclusion slide
5. Use bullet points for easy reading
6. Keep slide content concise and engaging

OUTPUT FORMAT:
Return ONLY a valid JSON array with this exact structure:

[
    {
        "slide_type": "title",
        "title": "Main presentation title",
        "subtitle": "Brief subtitle or tagline",
        "content": []
    },
    {
        "slide_type": "content",
        "title": "Slide title",
        "subtitle": "",
        "content": [
            "First bullet point",
            "Second bullet point",
            "Third bullet point"
        ]
    },
    {
        "slide_type": "conclusion",
        "title": "Conclusion",
        "subtitle": "Summary or call to action",
        "content": [
            "Key takeaway 1",
            "Key takeaway 2"
        ]
    }
]

SLIDE TYPES:
- "title": Opening slide with main title and subtitle
- "content": Regular content slide with title and bullet points
- "conclusion": Final slide with summary or call to action

CONTENT GUIDELINES:
- Keep titles under 50 characters
- Limit bullet points to 3-6 per slide
- Each bullet point should be 1-2 lines maximum
- Use clear, action-oriented language
"""

        if guidance:
            base_prompt += f"\n\nADDITIONAL GUIDANCE: {guidance}\n"
            base_prompt += "Apply this guidance to the tone, structure, and focus of the presentation.\n"
        
        base_prompt += f"\n\nTEXT TO CONVERT:\n{text_content}\n\n"
        base_prompt += "Remember: Return ONLY the JSON array, no additional text or formatting."
        
        return base_prompt
    
    def _parse_response(self, response_text: str) -> List[Dict[str, Any]]:
        """Parse and validate Gemini's JSON response"""
        try:
            # Clean the response text
            response_text = response_text.strip()
            
            # Remove markdown code blocks if present
            if response_text.startswith('```json'):
                response_text = response_text[7:]
            if response_text.startswith('```'):
                response_text = response_text[3:]
            if response_text.endswith('```'):
                response_text = response_text[:-3]
            
            response_text = response_text.strip()
            
            # Parse JSON
            slide_data = json.loads(response_text)
            
            # Validate structure
            if not isinstance(slide_data, list):
                raise ValueError("Response must be a list of slides")
            
            validated_slides = []
            for slide in slide_data:
                validated_slide = self._validate_slide(slide)
                validated_slides.append(validated_slide)
            
            return validated_slides
            
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse JSON response: {e}")
            # Fallback: create basic slides from text
            return self._create_fallback_slides(response_text)
        except Exception as e:
            logger.error(f"Error parsing response: {e}")
            raise Exception(f"Failed to parse Gemini response: {str(e)}")
    
    def _validate_slide(self, slide: Dict[str, Any]) -> Dict[str, Any]:
        """Validate and clean individual slide data"""
        validated = {
            "slide_type": slide.get("slide_type", "content"),
            "title": str(slide.get("title", ""))[:100],  # Limit title length
            "subtitle": str(slide.get("subtitle", ""))[:150],  # Limit subtitle length
            "content": []
        }
        
        # Validate slide type
        if validated["slide_type"] not in ["title", "content", "conclusion"]:
            validated["slide_type"] = "content"
        
        # Process content
        content = slide.get("content", [])
        if isinstance(content, list):
            for item in content[:6]:  # Limit to 6 bullet points
                if isinstance(item, str) and item.strip():
                    validated["content"].append(str(item).strip()[:200])  # Limit bullet length
        
        return validated
    
    def _create_fallback_slides(self, text: str) -> List[Dict[str, Any]]:
        """Create basic slides when JSON parsing fails"""
        logger.warning("Creating fallback slides due to parsing error")
        
        # Split text into chunks
        sentences = text.split('. ')
        chunks = []
        current_chunk = []
        
        for sentence in sentences:
            current_chunk.append(sentence)
            if len(current_chunk) >= 3:  # 3 sentences per slide
                chunks.append('. '.join(current_chunk))
                current_chunk = []
        
        if current_chunk:
            chunks.append('. '.join(current_chunk))
        
        slides = []
        
        # Title slide
        slides.append({
            "slide_type": "title",
            "title": "Generated Presentation",
            "subtitle": "Converted from your text content",
            "content": []
        })
        
        # Content slides
        for i, chunk in enumerate(chunks[:10]):  # Max 10 content slides
            slides.append({
                "slide_type": "content",
                "title": f"Topic {i + 1}",
                "subtitle": "",
                "content": [chunk[:200]]  # Single content item
            })
        
        # Conclusion slide
        slides.append({
            "slide_type": "conclusion",
            "title": "Summary",
            "subtitle": "Thank you",
            "content": ["Key points covered in this presentation"]
        })
        
        return slides
