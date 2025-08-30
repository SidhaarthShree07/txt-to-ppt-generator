"""
Slide refiner that rewrites content for specific placeholder requirements
Supports parallel API requests and retries
"""
import logging
import asyncio
import json
from typing import List, Dict, Any, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from .format_detector import get_content_placeholders_from_template_slide, placeholder_capacity

logger = logging.getLogger(__name__)

class SlideRefiner:
    """Refines slide content to match exact placeholder requirements"""
    
    def __init__(self, llm_provider):
        """
        Initialize with an LLM provider instance
        llm_provider should have a method to generate content
        """
        self.llm_provider = llm_provider
        self.max_retries = 3
    
    def refine_slides_parallel(self, mapped_slides: List[Dict[str, Any]], max_workers: int = 5) -> List[Dict[str, Any]]:
        """
        Refine multiple slides in parallel
        
        Args:
            mapped_slides: List of slides with template info embedded
            max_workers: Maximum parallel workers
            
        Returns:
            List of refined slides with properly formatted content
        """
        refined_slides = [None] * len(mapped_slides)
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all refinement tasks
            future_to_index = {
                executor.submit(self._refine_single_slide_with_retry, slide): i
                for i, slide in enumerate(mapped_slides)
            }
            
            # Collect results as they complete
            for future in as_completed(future_to_index):
                index = future_to_index[future]
                try:
                    refined_slide = future.result()
                    refined_slides[index] = refined_slide
                    logger.info(f"Successfully refined slide {index + 1}")
                except Exception as e:
                    logger.error(f"Failed to refine slide {index + 1}: {e}")
                    # Keep original on failure
                    refined_slides[index] = mapped_slides[index]
        
        return refined_slides
    
    def _refine_single_slide_with_retry(self, slide: Dict[str, Any]) -> Dict[str, Any]:
        """Refine a single slide with retry logic"""
        for attempt in range(self.max_retries):
            try:
                refined = self._refine_single_slide(slide)
                if self._validate_refined_content(refined):
                    return refined
                logger.warning(f"Refinement attempt {attempt + 1} produced invalid format, retrying...")
            except Exception as e:
                logger.warning(f"Refinement attempt {attempt + 1} failed: {e}")
                if attempt == self.max_retries - 1:
                    raise
        
        # Return original if all attempts fail
        return slide
    
    def _refine_single_slide(self, slide: Dict[str, Any]) -> Dict[str, Any]:
        """Refine content for a single slide based on its template requirements"""
        template_info = slide.get('_template_slide_info', {})
        if not template_info:
            return slide
        
        # Build refinement prompt
        prompt = self._build_refinement_prompt(slide, template_info)
        
        # Call LLM to refine
        try:
            if hasattr(self.llm_provider, 'model') and hasattr(self.llm_provider.model, 'generate_content'):
                # Gemini provider
                response = self.llm_provider.model.generate_content(prompt)
                refined_json = self._parse_llm_response(response.text)
            elif hasattr(self.llm_provider, 'client'):
                # OpenAI provider
                response = self.llm_provider.client.chat.completions.create(
                    model=self.llm_provider.model_name,
                    messages=[
                        {"role": "system", "content": "You are an expert at formatting presentation content. Return only valid JSON."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.3,
                    max_tokens=2000
                )
                refined_json = self._parse_llm_response(response.choices[0].message.content)
            else:
                # Generic provider
                response_text = self.llm_provider.refine_content(prompt)
                refined_json = self._parse_llm_response(response_text)
        except Exception as e:
            logger.error(f"LLM refinement failed: {e}")
            return slide
        
        # Merge refined content back into slide
        refined_slide = slide.copy()
        if 'title' in refined_json:
            refined_slide['title'] = refined_json['title']
        if 'subtitle' in refined_json:
            refined_slide['subtitle'] = refined_json['subtitle']
        if 'content' in refined_json:
            refined_slide['content'] = refined_json['content']
        
        return refined_slide
    
    def _build_refinement_prompt(self, slide: Dict[str, Any], template_info: Dict[str, Any]) -> str:
        """Build prompt to refine content for specific template requirements"""
        
        # Get placeholder information
        content_placeholders = get_content_placeholders_from_template_slide(template_info)
        
        prompt = f"""You are reformatting presentation content to fit exact template requirements.

CURRENT SLIDE CONTENT:
- Title: {slide.get('title', '')}
- Subtitle: {slide.get('subtitle', '')}
- Content items: {len(slide.get('content', []))}

TEMPLATE REQUIREMENTS:
"""
        
        # Add title/subtitle requirements
        if template_info.get('has_title'):
            title_ph = next((p for p in template_info.get('placeholders', []) 
                           if 'TITLE' in str(p.get('type', '')).upper()), None)
            if title_ph:
                max_chars = title_ph.get('max_chars_per_line', 60)
                prompt += f"- Title: Maximum {max_chars} characters\n"
        
        if template_info.get('has_subtitle'):
            subtitle_ph = next((p for p in template_info.get('placeholders', [])
                              if 'SUBTITLE' in str(p.get('type', '')).upper()), None)
            if subtitle_ph:
                max_chars = subtitle_ph.get('max_chars_per_line', 100)
                prompt += f"- Subtitle: Maximum {max_chars} characters\n"
        
        # Add content placeholder requirements
        if content_placeholders:
            prompt += f"\nCONTENT PLACEHOLDERS: {len(content_placeholders)} separate text areas\n"
            
            for i, ph in enumerate(content_placeholders):
                cap = placeholder_capacity(ph)
                prompt += f"\n[TEXT AREA {i+1}]:\n"
                prompt += f"- Format: {cap['text_format']}\n"
                prompt += f"- Capacity: {cap['suggested_lines']} lines\n"
                prompt += f"- Max chars per line: {cap['max_chars_per_line']}\n"
        
        # Add current content
        if slide.get('content'):
            prompt += f"\n\nCURRENT CONTENT TO REFORMAT:\n"
            for item in slide.get('content', []):
                prompt += f"- {item}\n"
        
        # Add instructions
        prompt += f"""

YOUR TASK:
1. Rewrite the content to fit EXACTLY within the template requirements
2. CRITICAL: You MUST generate content for ALL {len(content_placeholders)} text areas
3. Each text area MUST have at least 1-2 items of content
4. Use the format specified for each placeholder (numbered, bullets, paragraph)
5. Ensure each line fits within character limits
6. For multiple text areas, use "[PLACEHOLDER_X]" markers to separate content

ABSOLUTE REQUIREMENT: Generate content for ALL {len(content_placeholders)} text areas!
If you have limited source content, creatively expand it to fill all areas.
Each area should have relevant, meaningful content.

For multiple text areas, structure like:
{{
  "title": "Short title here",
  "subtitle": "Brief subtitle if needed",
  "content": [
    "Content for text area 1 - line 1",
    "Content for text area 1 - line 2",
    "[PLACEHOLDER_2]",
    "Content for text area 2 - line 1",
    "Content for text area 2 - line 2",
    "[PLACEHOLDER_3]",
    "Content for text area 3 - line 1"
  ]
}}

For single text area:
{{
  "title": "Short title here",
  "subtitle": "Brief subtitle if needed",
  "content": [
    "Bullet point 1",
    "Bullet point 2",
    "Bullet point 3"
  ]
}}

Return ONLY the JSON object, no other text."""
        
        # Add format-specific instructions
        if content_placeholders:
            formats = [placeholder_capacity(ph)['text_format'] for ph in content_placeholders]
            if 'numbered_list' in formats:
                prompt += "\n\nFor numbered lists, start each item with '1. ', '2. ', etc."
            if 'paragraph' in formats:
                prompt += "\n\nFor paragraph format, write flowing text without bullet points."
        
        return prompt
    
    def _parse_llm_response(self, response_text: str) -> Dict[str, Any]:
        """Parse LLM response to extract JSON"""
        try:
            # Clean response
            text = response_text.strip()
            
            # Remove markdown if present
            if '```json' in text:
                start = text.find('```json') + 7
                end = text.find('```', start)
                if end > start:
                    text = text[start:end]
            elif '```' in text:
                start = text.find('```') + 3
                end = text.find('```', start)
                if end > start:
                    text = text[start:end]
            
            text = text.strip()
            
            # Parse JSON
            return json.loads(text)
            
        except json.JSONDecodeError:
            logger.error(f"Failed to parse LLM response as JSON: {response_text[:200]}")
            # Return empty structure
            return {}
    
    def _validate_refined_content(self, slide: Dict[str, Any]) -> bool:
        """Validate that refined content has proper format markers if needed"""
        template_info = slide.get('_template_slide_info', {})
        content_placeholders = get_content_placeholders_from_template_slide(template_info)
        
        # If multiple placeholders, check for markers
        if len(content_placeholders) > 1:
            content = slide.get('content', [])
            if not content:
                return False
            
            # Check if there are placeholder markers
            markers_found = sum(1 for item in content 
                              if '[PLACEHOLDER' in str(item).upper())
            
            # We should have n-1 markers for n placeholders
            expected_markers = len(content_placeholders) - 1
            if markers_found < expected_markers:
                logger.warning(f"Expected {expected_markers} placeholder markers but found {markers_found}")
                # Still acceptable if there's at least some separation
                return markers_found > 0 or len(content) >= len(content_placeholders)
        
        return True
