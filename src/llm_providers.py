import json
import logging
from typing import List, Dict, Any
import google.generativeai as genai
try:
    import openai
except ImportError:
    openai = None

logger = logging.getLogger(__name__)

class BaseLLMProvider:
    """Base class for LLM providers"""
    
    def __init__(self, api_key: str, model_name: str = None):
        self.api_key = api_key
        self.model_name = model_name
    
    def parse_text_to_slides(self, text_content: str, guidance: str = "", template_structure: Dict[str, Any] = None, num_slides: int = None) -> List[Dict[str, Any]]:
        """Parse input text into structured slide content, enforcing text length and slide count"""
        # Enforce text character limit (same as app.py)
        MAX_TEXT_CHARS = 60000
        if text_content and len(text_content) > MAX_TEXT_CHARS:
            text_content = text_content[:MAX_TEXT_CHARS]
        # Build prompt with slide count if provided
        prompt = self._build_prompt(text_content, guidance, template_structure, num_slides=num_slides)
        # Actual LLM call is implemented in subclasses
        raise NotImplementedError("Subclasses must implement parse_text_to_slides")
    
    def _build_prompt(self, text_content: str, guidance: str, template_structure: Dict[str, Any] = None, num_slides: int = None) -> str:
        """Build the prompt for the LLM - shared across providers"""
        
        # If template structure is provided, create a customized prompt
        if template_structure and 'existing_slides' in template_structure:
            return self._build_template_aware_prompt(text_content, guidance, template_structure, num_slides)
        
        # Otherwise use the standard prompt
        base_prompt = """
You are a slide planner. Return JSON ONLY (no code fences, no Markdown), mapping the user's text into slides.

INSTRUCTIONS:
- Each slide must have a clear, specific title (10-60 characters) and, if required, a subtitle (20-100 characters).
- For each content slide, generate 4-6 bullet points. Each bullet must be a full, meaningful sentence (minimum 40 characters, ideally 60-120 characters). Do NOT use incomplete points, ellipses, or '...'.
- NEVER output any bullet or content as '...', 'None', 'N/A', or similar. Every bullet must be a real, substantive statement.
- If the source text is short, expand on the ideas with relevant context, examples, or implications. If the text is long, summarize and condense as needed.
- Do not leave any content arrays empty. If you cannot find enough points, expand with logical, relevant, and non-repetitive information.
- Do not use generic or filler statements. Avoid vague language. Be specific and actionable.
- Do not output images, tables, or notes.

CRITICAL REQUIREMENTS:
- Bullet points must be complete, detailed, and never end with '...'.
- No bullet or content should be less than 40 characters or more than 140 characters.
- Titles and subtitles must be concise, clear, and never generic.
- If a placeholder or content area is present, it must be filled with real, meaningful text.

OUTPUT FORMAT:
Return ONLY a valid JSON array with this exact structure:

[
    {
        "slide_type": "title",
        "title": "Main presentation title - be specific and engaging",
        "subtitle": "Descriptive subtitle that provides context or value proposition",
        "content": []
    },
    {
        "slide_type": "content",
        "title": "Clear, descriptive slide title that introduces the topic",
        "subtitle": "",
        "content": [
            "First detailed point that explains a key concept or provides specific information about the topic",
            "Second comprehensive point that builds on the first with examples, data, or additional context",
            "Third substantive point that provides deeper insight or another perspective on the subject",
            "Fourth important detail that supports the overall message with concrete information",
            "Fifth relevant point that adds value through implications, benefits, or related considerations"
        ]
    },
    {
        "slide_type": "conclusion",
        "title": "Conclusion: Key Takeaways and Next Steps",
        "subtitle": "Actionable insights and recommendations for moving forward",
        "content": [
            "First key takeaway: Specific insight with clear implications for the audience",
            "Second action item: Concrete step that can be taken based on the presentation",
            "Third recommendation: Strategic consideration or future opportunity to explore",
            "Fourth key point: Important reminder or critical success factor to remember"
        ]
    }
]

SLIDE TYPES:
- "title": Opening slide with main title and descriptive subtitle (both required)
- "content": Regular content slide with title and 4-6 detailed bullet points (all required)
- "conclusion": Final slide with summary title, action subtitle, and 3-5 key takeaways (all required)

CONTENT GUIDELINES:
- Titles: 10-60 characters, clear, descriptive, and engaging
- Subtitles: 20-100 characters, provide valuable context, purpose, or key message
- Bullet points: 40-140 characters each, complete sentences with substance
- Each bullet must be a complete, informative sentence (not fragments, not '...')
- Use action verbs, specific details, and concrete examples
- Include data points, percentages, or metrics when available in source text
- Avoid vague statements - be specific, informative, and actionable
- Extract and expand on concrete information from the source text
- If source text lacks detail, intelligently expand with relevant context
- Ensure every slide has enough content to be meaningful and valuable
"""

        if guidance:
            base_prompt += f"\n\nADDITIONAL GUIDANCE: {guidance}\n"
            base_prompt += "Apply this guidance to the tone, structure, and focus of the presentation.\n"
        
        if num_slides:
            base_prompt += f"\n\nREQUIRED NUMBER OF SLIDES: {num_slides} (expand or condense as needed, but output exactly {num_slides} slides)"
        base_prompt += f"\n\nTEXT TO CONVERT:\n{text_content}\n\n"
        base_prompt += "Remember: Return ONLY the JSON array, no additional text or formatting."
        return base_prompt
    
    def _build_template_aware_prompt(self, text_content: str, guidance: str, template_structure: Dict[str, Any], num_slides: int = None) -> str:
        """Build a prompt that considers the actual template structure"""
        
        slides_info = template_structure.get('existing_slides', [])
        prompt = f"""
You are an expert presentation designer. You need to create content for a PowerPoint presentation using an existing template.

TEMPLATE STRUCTURE:
The template has {len(slides_info)} existing slides that need to be populated with content.

SLIDE DETAILS:
"""
        
        # Add details about each slide
        for slide in slides_info:
            slide_idx = slide['slide_index']
            slide_type = slide.get('suggested_content_type', 'content')
            content_format = slide.get('content_format', None)

            prompt += f"\n[SLIDE {slide_idx + 1}]\n"
            prompt += f"- Suggested Type: {slide_type}\n"
            prompt += f"- Layout: {slide['layout_name']}\n"

            # Add placeholder reference text and length for LLM guidance
            for ph in slide.get('placeholders', []):
                ph_type = ph.get('type', '').upper()
                ph_text = ph.get('text', '').strip()
                ph_len = len(ph_text.split())
                ph_name = ph.get('name', ph_type)
                # Only count real content placeholders (not static markers)
                import re
                if not re.fullmatch(r'\d+|[ivxlc]+|[a-zA-Z]\.|•|\-|\–', str(ph_name).strip()):
                    if ph_text:
                        prompt += f"- Placeholder '{ph_name}': Example text: '{ph_text}' (about {ph_len} words, {len(ph_text)} characters). Match the style, tone, and keep your generated text within ±2 words or ±10 characters of this length.\n"

            if slide['has_title']:
                title_ph = next((p for p in slide['placeholders'] if 'TITLE' in p['type']), None)
                if title_ph:
                    if title_ph.get('actual_text_length', 0) > 0:
                        prompt += f"- Title: Approximately {title_ph['actual_text_length']} characters (max {title_ph['max_chars_per_line']})\n"
                    else:
                        prompt += f"- Title: Max {title_ph['max_chars_per_line']} characters\n"

            if slide['has_subtitle']:
                subtitle_ph = next((p for p in slide['placeholders'] if 'SUBTITLE' in p['type']), None)
                if subtitle_ph:
                    if subtitle_ph.get('actual_text_length', 0) > 0:
                        prompt += f"- Subtitle: Approximately {subtitle_ph['actual_text_length']} characters total\n"
                    else:
                        prompt += f"- Subtitle: Max {subtitle_ph['max_chars_per_line']} characters per line, {subtitle_ph['suggested_lines']} lines\n"

            if slide['has_content']:
                content_phs = [p for p in slide['placeholders'] if 'CONTENT' in p['type'] or 'BODY' in p['type']]
                n_content = len(content_phs)
                format_str = ''
                if n_content > 1:
                    prompt += f"- This slide has {n_content} content placeholders. YOU MUST generate exactly {n_content} content items for this slide.\n"
                    prompt += f"- If you generate fewer or more than {n_content} content items, your output will be rejected and extra items will be discarded.\n"
                    prompt += f"- Each content placeholder should have a unique, meaningful point, and the length of each item must closely match the example/placeholder text.\n"
                    prompt += f"- Structure your content array with '[NEXT_PLACEHOLDER]' as a separator between each content area.\n"
                else:
                    prompt += f"- This slide has 1 content placeholder. Generate exactly 1 content item, matching the length and style of the placeholder text.\n"
                if content_phs:
                    for idx, ph in enumerate(content_phs):
                        if ph.get('text_format'):
                            format_str = f" Format: {ph['text_format']}."
                        if ph.get('line_count', 0) > 0:
                            prompt += f"- Content for area {idx+1}:{format_str} {ph['line_count']} items/lines, max {ph['max_chars_per_line']} chars/line\n"
                        else:
                            prompt += f"- Content for area {idx+1}:{format_str} Max {ph['max_chars_per_line']} chars/line, up to {ph['suggested_lines']} lines\n"

            # Add format-specific guidance
            if content_format == 'numbered_list':
                prompt += "  FORMAT: Use numbered list (1. 2. 3. etc.)\n"
            elif content_format == 'bullet_list':
                prompt += "  FORMAT: Use bullet points\n"
            elif content_format == 'paragraph':
                prompt += "  FORMAT: Use paragraph text (not bullet points)\n"
        
        if num_slides:
            prompt += f"\nREQUIRED NUMBER OF SLIDES: {num_slides} (expand or condense as needed, but output exactly {num_slides} slides)\n"
        prompt += f"""

YOUR TASK:
1. Create content for EXACTLY {len(slides_info)} slides
2. Each slide must match the constraints and format patterns listed above
3. Keep titles concise and impactful (fit within character limits)
4. Match the detected format (numbered list, bullet list, or paragraph) for each slide
5. Do not exceed the character limits for each placeholder
6. If a placeholder should be empty (no subtitle needed), set it to null or empty string

CRITICAL REQUIREMENTS:
- Generate content for ALL {len(slides_info)} slides
- Respect the EXACT character and line limits for each placeholder
- Follow the FORMAT specified for each slide (numbered/bullet/paragraph)
- Use the suggested slide types (title, content, conclusion) appropriately
- Make titles short and punchy to avoid text overflow
- For numbered lists: Start each item with "1. ", "2. ", etc.
- For bullet lists: Create concise bullet points
- For paragraphs: Write flowing text without bullet points
- Set content to null or empty if a placeholder shouldn't have content

MULTIPLE TEXT PLACEHOLDERS:
- If a slide has multiple content/text areas, use '[NEXT_PLACEHOLDER]' to separate content
- Example for a slide with 3 text areas:
  "content": [
    "Point 1 for first text area",
    "Point 2 for first text area",
    "[NEXT_PLACEHOLDER]",
    "Point 1 for second text area",
    "Point 2 for second text area",
    "[NEXT_PLACEHOLDER]",
    "Point 1 for third text area",
    "Point 2 for third text area"
  ]

OUTPUT FORMAT:
Return a JSON array with EXACTLY {len(slides_info)} slide objects:

[
    {{
        "slide_number": 1,
        "slide_type": "title/content/conclusion",
        "title": "Short title (respect char limit)",
        "subtitle": "Brief subtitle or null if not needed",
        "content": ["Point 1", "Point 2", ...] or null
    }},
    ...
]
"""
        
        if guidance:
            prompt += f"\nADDITIONAL GUIDANCE: {guidance}\n"
        
        prompt += f"\nTEXT TO CONVERT:\n{text_content}\n\n"
        prompt += "Return ONLY the JSON array, no additional text."
        
        return prompt
    
    def _build_initial_content_prompt(self, text_content: str, guidance: str, num_slides: int = None) -> str:
        """Build prompt for initial content generation with optional slide count constraint"""
        
        if num_slides:
            prompt = f"""
You are an expert presentation designer. Convert the provided text into EXACTLY {num_slides} presentation slides.

You MUST generate exactly {num_slides} slides total, including:
- 1 title slide (with engaging title and descriptive subtitle)
- {num_slides - 2} content slides (covering key points from the text)
- 1 conclusion slide (with key takeaways)

Adjust the content distribution to fit exactly {num_slides} slides. If the text is short, expand on ideas. If long, condense appropriately.
"""
        else:
            prompt = """
You are an expert presentation designer. Convert the provided text into presentation slides.

Generate a comprehensive presentation with:
- A title slide with engaging title and descriptive subtitle
- Multiple content slides covering all key points from the text
- A conclusion slide with key takeaways
"""
        
        prompt += """

Focus on extracting important information from the text.

IMPORTANT: Return ONLY a valid JSON array with this EXACT structure:
[
    {
        "slide_type": "title",
        "title": "Your title here (10-60 chars)",
        "subtitle": "Your subtitle here (20-100 chars)",
        "content": []  // MUST be empty for title slides - no bullet points!
    },
    {
        "slide_type": "content",
        "title": "Slide title (10-60 chars)",
        "subtitle": "",
        "content": ["Bullet point 1 (40-120 chars)", "Bullet point 2 (40-120 chars)", "Bullet point 3 (40-120 chars)"]
    },
    {
        "slide_type": "conclusion",
        "title": "Conclusion title (10-60 chars)",
        "subtitle": "Conclusion subtitle (20-100 chars)",
        "content": ["Key point 1 (40-120 chars)", "Key point 2 (40-120 chars)"]
    }
]

CHARACTER LIMITS:
- Titles: 10-60 characters
- Subtitles: 20-100 characters  
- Bullet points: 40-120 characters each

DO NOT include any text before or after the JSON array.
DO NOT use markdown formatting.
DO NOT add comments.
Just the JSON array.
"""
        if guidance:
            prompt += f"\nGUIDANCE: {guidance}\n"
        
        if num_slides:
            prompt += f"\nREMEMBER: You MUST generate EXACTLY {num_slides} slides total.\n"
        
        prompt += f"\nTEXT TO CONVERT:\n{text_content}\n\n"
        prompt += "Return ONLY the JSON array."
        return prompt
    
    def _build_refinement_prompt(self, mapped_content: List[Dict[str, Any]], 
                                template_structure: Dict[str, Any],
                                selected_indices: List[int]) -> str:
        """Build prompt to refine content for specific template slides"""
        prompt = """You need to refine presentation content to perfectly fit template constraints.

For each slide below, adjust the content to match the exact requirements:

"""
        
        for i, (content, idx) in enumerate(zip(mapped_content, selected_indices)):
            template_slide = template_structure['existing_slides'][idx]
            prompt += f"\n[SLIDE {i+1}]\n"
            prompt += f"Current content:\n"
            prompt += f"- Title: {content.get('title', '')}\n"
            prompt += f"- Subtitle: {content.get('subtitle', '')}\n"
            prompt += f"- Content items: {len(content.get('content', []))}\n"
            
            prompt += f"\nTemplate requirements:\n"
            for ph in template_slide.get('placeholders', []):
                if 'TITLE' in ph.get('type', ''):
                    prompt += f"- Title: max {ph.get('max_chars_per_line', 60)} chars\n"
                elif 'SUBTITLE' in ph.get('type', ''):
                    prompt += f"- Subtitle: max {ph.get('max_chars_per_line', 100)} chars\n"
                elif 'CONTENT' in ph.get('type', '') or 'BODY' in ph.get('type', ''):
                    format_type = ph.get('text_format', 'bullet_list')
                    prompt += f"- Content: {format_type}, max {ph.get('suggested_lines', 5)} items, "
                    prompt += f"{ph.get('max_chars_per_line', 80)} chars/line\n"
                    
                    if format_type == 'numbered_list':
                        prompt += "  FORMAT: Use numbered list (1. 2. 3.)\n"
                    elif format_type == 'paragraph':
                        prompt += "  FORMAT: Use paragraph text, not bullets\n"
        
        prompt += """\n\nReturn a JSON array with refined content for each slide.
Ensure all text fits within the constraints and follows the specified format.
Return ONLY the JSON array.
"""
        return prompt
    
    def _parse_response(self, response_text: str) -> List[Dict[str, Any]]:
        """Parse and validate Gemini's JSON response with better error handling"""
        try:
            # Store original for debugging
            original_response = response_text
            
            # Clean the response text more thoroughly
            response_text = response_text.strip()
            
            # Remove markdown code blocks if present
            if '```json' in response_text:
                start = response_text.find('```json') + 7
                end = response_text.find('```', start)
                if end > start:
                    response_text = response_text[start:end]
            elif '```' in response_text:
                start = response_text.find('```') + 3
                end = response_text.find('```', start)
                if end > start:
                    response_text = response_text[start:end]
            
            response_text = response_text.strip()
            
            # Remove any trailing commas that might cause JSON errors
            import re
            # Fix trailing commas in arrays
            response_text = re.sub(r',\s*]', ']', response_text)
            # Fix trailing commas in objects
            response_text = re.sub(r',\s*}', '}', response_text)
            
            # Fix common JSON formatting issues
            # Replace single quotes with double quotes if needed
            if "'" in response_text and '"' not in response_text:
                response_text = response_text.replace("'", '"')
            
            # Try to find JSON array in the response
            if not response_text.startswith('['):
                # Look for the first '[' in the response
                array_start = response_text.find('[')
                if array_start != -1:
                    response_text = response_text[array_start:]
            
            if not response_text.endswith(']'):
                # Look for the last ']' in the response
                array_end = response_text.rfind(']')
                if array_end != -1:
                    response_text = response_text[:array_end + 1]
            
            # Parse JSON
            slide_data = json.loads(response_text)
            
            # Validate structure
            if not isinstance(slide_data, list):
                raise ValueError("Response must be a list of slides")
            
            if len(slide_data) == 0:
                raise ValueError("Response contains no slides")
            
            validated_slides = []
            for slide in slide_data:
                validated_slide = self._validate_slide(slide)
                validated_slides.append(validated_slide)
            
            logger.info(f"Successfully parsed {len(validated_slides)} slides")
            return validated_slides
            
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse JSON response: {e}")
            logger.error(f"Problematic response text (first 500 chars): {response_text[:500]}")
            
            # Try one more time with aggressive cleaning
            try:
                # Extract JSON array more aggressively
                import re
                json_match = re.search(r'\[.*\]', response_text, re.DOTALL)
                if json_match:
                    cleaned_json = json_match.group(0)
                    slide_data = json.loads(cleaned_json)
                    validated_slides = []
                    for slide in slide_data:
                        validated_slide = self._validate_slide(slide)
                        validated_slides.append(validated_slide)
                    logger.info(f"Successfully parsed {len(validated_slides)} slides after aggressive cleaning")
                    return validated_slides
            except:
                pass
            
            # Last resort: fallback
            logger.warning("Using fallback slide generation")
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

class GeminiProvider:
    """Google Gemini API provider for text parsing and slide generation"""
    
    def __init__(self, api_key: str, model_name: str = 'gemini-2.5-pro'):
        """Initialize Gemini provider with API key and model name"""
        self.api_key = api_key
        self.model_name = model_name or 'gemini-2.5-pro'
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel(self.model_name)
    
    def parse_text_to_slides(self, text_content: str, guidance: str = "", template_structure: Dict[str, Any] = None, num_slides: int = None) -> List[Dict[str, Any]]:
        """
        Parse input text into structured slide content using Gemini with two-pass generation
        
        Args:
            text_content: The input text to be converted
            guidance: Optional guidance for tone/structure
            template_structure: Optional template structure information
            num_slides: Target number of slides to generate
            
        Returns:
            List of slide dictionaries with structure and content
        """
        try:
            # First pass: Generate initial content
            if template_structure and 'existing_slides' in template_structure:
                # If we have a template, do intelligent two-pass generation
                logger.info("Starting two-pass content generation with template awareness")
                # Pass 1: Generate initial content WITH num_slides constraint
                initial_prompt = self._build_initial_content_prompt(text_content, guidance, num_slides)
                logger.info(f"Pass 1: Generating initial content (target: {num_slides} slides)" if num_slides else "Pass 1: Generating initial content")
                initial_response = self.model.generate_content(initial_prompt)
                initial_slides = self._parse_response(initial_response.text)
                
                # Map content to best template slides
                from .content_mapper import ContentMapper
                mapper = ContentMapper()
                mapped_content, selected_indices = mapper.map_content_to_template(
                    initial_slides, template_structure
                )
                
                # Pass 2: Refine content for selected slides
                if mapped_content:
                    logger.info("Pass 2: Refining content for selected template slides")
                    refined_prompt = self._build_refinement_prompt(
                        mapped_content, template_structure, selected_indices
                    )
                    refined_response = self.model.generate_content(refined_prompt)
                    final_slides = self._parse_response(refined_response.text)
                    
                    # Merge refinements with mapped content, keeping original content if refinement fails
                    for i, slide in enumerate(final_slides[:len(mapped_content)]):
                        if i < len(mapped_content):
                            # Keep the original content if the refined version is empty
                            if slide.get('content') or mapped_content[i].get('content') is None:
                                mapped_content[i]['content'] = slide.get('content', mapped_content[i].get('content', []))
                            if slide.get('title'):
                                mapped_content[i]['title'] = slide.get('title')
                            if slide.get('subtitle'):
                                mapped_content[i]['subtitle'] = slide.get('subtitle')
                    
                    # Log the actual content being returned
                    for idx, content in enumerate(mapped_content):
                        logger.debug(f"Slide {idx+1} final content: title='{content.get('title')}', "
                                   f"subtitle='{content.get('subtitle')}', "
                                   f"content={content.get('content', [])}")
                    
                    logger.info(f"Generated and refined {len(mapped_content)} slides")
                    return mapped_content
                else:
                    return initial_slides
            else:
                # Single pass for non-template generation
                prompt = self._build_prompt(text_content, guidance, template_structure, num_slides=num_slides)
                logger.info("Sending request to Gemini API")
                response = self.model.generate_content(prompt)
                slide_structure = self._parse_response(response.text)
                logger.info(f"Generated {len(slide_structure)} slides")
                return slide_structure
            
        except Exception as e:
            logger.error(f"Error with Gemini API: {str(e)}")
            raise Exception(f"Failed to process text with Gemini: {str(e)}")
    
    def _build_prompt(self, text_content: str, guidance: str, template_structure: Dict[str, Any] = None, num_slides: int = None) -> str:
        """Build the prompt for Gemini API"""
        
        # If template structure is provided, use template-aware prompt
        if template_structure and 'existing_slides' in template_structure:
            return self._build_template_aware_prompt(text_content, guidance, template_structure, num_slides)
        
        base_prompt = """
You are an expert presentation designer. Your task is to analyze the provided text and convert it into a structured PowerPoint presentation.

INSTRUCTIONS:
1. Break down the text into logical slides
2. Each slide MUST have meaningful content - never leave content arrays empty
3. Create an appropriate number of slides (typically 7-15 depending on content length)
4. Include a title slide with subtitle and conclusion slide with key takeaways
5. Use bullet points for easy reading
6. Keep slide content concise but substantive
7. Extract and expand key points from the source text

CRITICAL REQUIREMENTS:
- EVERY content slide MUST have at least 4-6 bullet points
- EVERY bullet point must be a complete sentence with substantive information (15-80 words each)
- Title slides MUST have both title (10-60 characters) AND subtitle (20-100 characters)
- Subtitles should provide context, explain the presentation's purpose, or set expectations
- Conclusion slides MUST have at least 3-5 key takeaways or action items
- Each takeaway must be actionable and specific (not generic statements)
- Never leave content arrays empty - always provide meaningful, detailed content
- If the source text is short, expand on the ideas with relevant context and implications

OUTPUT FORMAT:
Return ONLY a valid JSON array with this exact structure:

[
    {
        "slide_type": "title",
        "title": "Main presentation title - be specific and engaging",
        "subtitle": "Descriptive subtitle that provides context or value proposition",
        "content": []
    },
    {
        "slide_type": "content",
        "title": "Clear, descriptive slide title that introduces the topic",
        "subtitle": "",
        "content": [
            "First detailed point that explains a key concept or provides specific information about the topic",
            "Second comprehensive point that builds on the first with examples, data, or additional context",
            "Third substantive point that provides deeper insight or another perspective on the subject",
            "Fourth important detail that supports the overall message with concrete information",
            "Fifth relevant point that adds value through implications, benefits, or related considerations"
        ]
    },
    {
        "slide_type": "conclusion",
        "title": "Conclusion: Key Takeaways and Next Steps",
        "subtitle": "Actionable insights and recommendations for moving forward",
        "content": [
            "First key takeaway: Specific insight with clear implications for the audience",
            "Second action item: Concrete step that can be taken based on the presentation",
            "Third recommendation: Strategic consideration or future opportunity to explore",
            "Fourth key point: Important reminder or critical success factor to remember"
        ]
    }
]

SLIDE TYPES:
- "title": Opening slide with main title and descriptive subtitle (both required)
- "content": Regular content slide with title and 4-6 detailed bullet points (all required)
- "conclusion": Final slide with summary title, action subtitle, and 3-5 key takeaways (all required)

CONTENT GUIDELINES:
- Titles: 10-60 characters, clear, descriptive, and engaging
- Subtitles: 20-100 characters, provide valuable context, purpose, or key message
- Bullet points: 15-120 characters each, complete sentences with substance
- Each bullet must be a complete, informative sentence (not fragments)
- Use action verbs, specific details, and concrete examples
- Include data points, percentages, or metrics when available in source text
- Avoid vague statements - be specific, informative, and actionable
- Extract and expand on concrete information from the source text
- If source text lacks detail, intelligently expand with relevant context
- Ensure every slide has enough content to be meaningful and valuable
"""

        if guidance:
            base_prompt += f"\n\nADDITIONAL GUIDANCE: {guidance}\n"
            base_prompt += "Apply this guidance to the tone, structure, and focus of the presentation.\n"
        
        if num_slides:
            base_prompt += f"\n\nREQUIRED NUMBER OF SLIDES: {num_slides} (expand or condense as needed, but output exactly {num_slides} slides)\n"
        
        base_prompt += f"\n\nTEXT TO CONVERT:\n{text_content}\n\n"
        base_prompt += "Remember: Return ONLY the JSON array, no additional text or formatting."
        
        return base_prompt
    
    def _build_template_aware_prompt(self, text_content: str, guidance: str, template_structure: Dict[str, Any], num_slides: int = None) -> str:
        """Build a prompt that considers the actual template structure"""
        
        slides_info = template_structure.get('existing_slides', [])
        
        prompt = f"""
You are an expert presentation designer. You need to create content for a PowerPoint presentation using an existing template.

TEMPLATE STRUCTURE:
The template has {len(slides_info)} existing slides that need to be populated with content.

SLIDE DETAILS:
"""
        
        # Add details about each slide
        for slide in slides_info:
            slide_idx = slide['slide_index']
            slide_type = slide.get('suggested_content_type', 'content')
            content_format = slide.get('content_format', None)
            
            prompt += f"\n[SLIDE {slide_idx + 1}]\n"
            prompt += f"- Suggested Type: {slide_type}\n"
            prompt += f"- Layout: {slide['layout_name']}\n"
            
            if slide['has_title']:
                title_ph = next((p for p in slide['placeholders'] if 'TITLE' in p['type']), None)
                if title_ph:
                    # Use actual text length if available, otherwise use calculated max
                    if title_ph.get('actual_text_length', 0) > 0:
                        prompt += f"- Title: Approximately {title_ph['actual_text_length']} characters (max {title_ph['max_chars_per_line']})\n"
                    else:
                        prompt += f"- Title: Max {title_ph['max_chars_per_line']} characters\n"
            
            if slide['has_subtitle']:
                subtitle_ph = next((p for p in slide['placeholders'] if 'SUBTITLE' in p['type']), None)
                if subtitle_ph:
                    if subtitle_ph.get('actual_text_length', 0) > 0:
                        prompt += f"- Subtitle: Approximately {subtitle_ph['actual_text_length']} characters total\n"
                    else:
                        prompt += f"- Subtitle: Max {subtitle_ph['max_chars_per_line']} characters per line, {subtitle_ph['suggested_lines']} lines\n"
            
            if slide['has_content']:
                content_ph = next((p for p in slide['placeholders'] if 'CONTENT' in p['type'] or 'BODY' in p['type']), None)
                if content_ph:
                    format_str = ""
                    if content_ph.get('text_format'):
                        format_str = f" Format: {content_ph['text_format']}."
                    
                    if content_ph.get('line_count', 0) > 0:
                        prompt += f"- Content:{format_str} {content_ph['line_count']} items/lines, "
                        prompt += f"max {content_ph['max_chars_per_line']} chars/line\n"
                    else:
                        prompt += f"- Content:{format_str} Max {content_ph['max_chars_per_line']} chars/line, "
                        prompt += f"up to {content_ph['suggested_lines']} lines\n"
            
            # Add format-specific guidance
            if content_format == 'numbered_list':
                prompt += "  FORMAT: Use numbered list (1. 2. 3. etc.)\n"
            elif content_format == 'bullet_list':
                prompt += "  FORMAT: Use bullet points\n"
            elif content_format == 'paragraph':
                prompt += "  FORMAT: Use paragraph text (not bullet points)\n"
        
        # Add num_slides requirement if specified
        if num_slides:
            prompt += f"\n\nIMPORTANT: You must generate content for EXACTLY {num_slides} slides from the available {len(slides_info)} template slides.\n"
            prompt += f"Select the most appropriate {num_slides} slides from the template to match your content.\n"
        
        prompt += f"""

YOUR TASK:
1. Create content for EXACTLY {num_slides if num_slides else len(slides_info)} slides
2. Each slide must match the constraints and format patterns listed above
3. Keep titles concise and impactful (fit within character limits)
4. Match the detected format (numbered list, bullet list, or paragraph) for each slide
5. Do not exceed the character limits for each placeholder
6. If a placeholder should be empty (no subtitle needed), set it to null or empty string

CRITICAL REQUIREMENTS:
- Generate content for {num_slides if num_slides else 'ALL ' + str(len(slides_info))} slides
- Respect the EXACT character and line limits for each placeholder
- Follow the FORMAT specified for each slide (numbered/bullet/paragraph)
- Use the suggested slide types (title, content, conclusion) appropriately
- Make titles short and punchy to avoid text overflow
- For numbered lists: Start each item with "1. ", "2. ", etc.
- For bullet lists: Create concise bullet points
- For paragraphs: Write flowing text without bullet points
- Set content to null or empty if a placeholder shouldn't have content

OUTPUT FORMAT:
Return a JSON array with EXACTLY {num_slides if num_slides else len(slides_info)} slide objects:

[
    {{
        "slide_number": 1,
        "slide_type": "title/content/conclusion",
        "title": "Short title (respect char limit)",
        "subtitle": "Brief subtitle or null if not needed",
        "content": ["Point 1", "Point 2", ...] or null
    }},
    ...
]
"""
        
        if guidance:
            prompt += f"\nADDITIONAL GUIDANCE: {guidance}\n"
        
        prompt += f"\nTEXT TO CONVERT:\n{text_content}\n\n"
        prompt += "Return ONLY the JSON array, no additional text."
        
        return prompt
    
    def _build_initial_content_prompt(self, text_content: str, guidance: str, num_slides: int = None) -> str:
        """Build prompt for initial content generation with optional slide count constraint"""
        
        if num_slides:
            prompt = f"""
You are an expert presentation designer. Convert the provided text into EXACTLY {num_slides} presentation slides.

You MUST generate exactly {num_slides} slides total, including:
- 1 title slide (with engaging title and descriptive subtitle)
- {num_slides - 2} content slides (covering key points from the text)
- 1 conclusion slide (with key takeaways)

Adjust the content distribution to fit exactly {num_slides} slides. If the text is short, expand on ideas. If long, condense appropriately.
"""
        else:
            prompt = """
You are an expert presentation designer. Convert the provided text into presentation slides.

Generate a comprehensive presentation with:
- A title slide with engaging title and descriptive subtitle
- Multiple content slides covering all key points from the text
- A conclusion slide with key takeaways
"""
        
        prompt += """

Focus on extracting important information from the text.

IMPORTANT: Return ONLY a valid JSON array with this EXACT structure:
[
    {
        "slide_type": "title",
        "title": "Your title here (10-60 chars)",
        "subtitle": "Your subtitle here (20-100 chars)",
        "content": []  // MUST be empty for title slides - no bullet points!
    },
    {
        "slide_type": "content",
        "title": "Slide title (10-60 chars)",
        "subtitle": "",
        "content": ["Bullet point 1 (40-120 chars)", "Bullet point 2 (40-120 chars)", "Bullet point 3 (40-120 chars)"]
    },
    {
        "slide_type": "conclusion",
        "title": "Conclusion title (10-60 chars)",
        "subtitle": "Conclusion subtitle (20-100 chars)",
        "content": ["Key point 1 (40-120 chars)", "Key point 2 (40-120 chars)"]
    }
]

CHARACTER LIMITS:
- Titles: 10-60 characters
- Subtitles: 20-100 characters  
- Bullet points: 40-120 characters each

DO NOT include any text before or after the JSON array.
DO NOT use markdown formatting.
DO NOT add comments.
Just the JSON array.
"""
        if guidance:
            prompt += f"\nGUIDANCE: {guidance}\n"
        
        if num_slides:
            prompt += f"\nREMEMBER: You MUST generate EXACTLY {num_slides} slides total.\n"
        
        prompt += f"\nTEXT TO CONVERT:\n{text_content}\n\n"
        prompt += "Return ONLY the JSON array."
        return prompt
    
    def _build_refinement_prompt(self, mapped_content: List[Dict[str, Any]], 
                                template_structure: Dict[str, Any],
                                selected_indices: List[int]) -> str:
        """Build prompt to refine content for specific template slides"""
        prompt = """You need to refine presentation content to perfectly fit template constraints.

For each slide below, adjust the content to match the exact requirements:

"""
        
        for i, (content, idx) in enumerate(zip(mapped_content, selected_indices)):
            template_slide = template_structure['existing_slides'][idx]
            prompt += f"\n[SLIDE {i+1}]\n"
            prompt += f"Current content:\n"
            prompt += f"- Title: {content.get('title', '')}\n"
            prompt += f"- Subtitle: {content.get('subtitle', '')}\n"
            prompt += f"- Content items: {len(content.get('content', []))}\n"
            
            prompt += f"\nTemplate requirements:\n"
            for ph in template_slide.get('placeholders', []):
                if 'TITLE' in ph.get('type', ''):
                    prompt += f"- Title: max {ph.get('max_chars_per_line', 60)} chars\n"
                elif 'SUBTITLE' in ph.get('type', ''):
                    prompt += f"- Subtitle: max {ph.get('max_chars_per_line', 100)} chars\n"
                elif 'CONTENT' in ph.get('type', '') or 'BODY' in ph.get('type', ''):
                    format_type = ph.get('text_format', 'bullet_list')
                    prompt += f"- Content: {format_type}, max {ph.get('suggested_lines', 5)} items, "
                    prompt += f"{ph.get('max_chars_per_line', 80)} chars/line\n"
                    
                    if format_type == 'numbered_list':
                        prompt += "  FORMAT: Use numbered list (1. 2. 3.)\n"
                    elif format_type == 'paragraph':
                        prompt += "  FORMAT: Use paragraph text, not bullets\n"
        
        prompt += """\n\nReturn a JSON array with refined content for each slide.
Ensure all text fits within the constraints and follows the specified format.
Return ONLY the JSON array.
"""
        return prompt
    
    def _parse_response(self, response_text: str) -> List[Dict[str, Any]]:
        """Parse and validate Gemini's JSON response with better error handling"""
        try:
            # Store original for debugging
            original_response = response_text
            
            # Clean the response text more thoroughly
            response_text = response_text.strip()
            
            # Remove markdown code blocks if present
            if '```json' in response_text:
                start = response_text.find('```json') + 7
                end = response_text.find('```', start)
                if end > start:
                    response_text = response_text[start:end]
            elif '```' in response_text:
                start = response_text.find('```') + 3
                end = response_text.find('```', start)
                if end > start:
                    response_text = response_text[start:end]
            
            response_text = response_text.strip()
            
            # Remove any trailing commas that might cause JSON errors
            import re
            # Fix trailing commas in arrays
            response_text = re.sub(r',\s*]', ']', response_text)
            # Fix trailing commas in objects
            response_text = re.sub(r',\s*}', '}', response_text)
            
            # Fix common JSON formatting issues
            # Replace single quotes with double quotes if needed
            if "'" in response_text and '"' not in response_text:
                response_text = response_text.replace("'", '"')
            
            # Try to find JSON array in the response
            if not response_text.startswith('['):
                # Look for the first '[' in the response
                array_start = response_text.find('[')
                if array_start != -1:
                    response_text = response_text[array_start:]
            
            if not response_text.endswith(']'):
                # Look for the last ']' in the response
                array_end = response_text.rfind(']')
                if array_end != -1:
                    response_text = response_text[:array_end + 1]
            
            # Parse JSON
            slide_data = json.loads(response_text)
            
            # Validate structure
            if not isinstance(slide_data, list):
                raise ValueError("Response must be a list of slides")
            
            if len(slide_data) == 0:
                raise ValueError("Response contains no slides")
            
            validated_slides = []
            for slide in slide_data:
                validated_slide = self._validate_slide(slide)
                validated_slides.append(validated_slide)
            
            logger.info(f"Successfully parsed {len(validated_slides)} slides")
            return validated_slides
            
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse JSON response: {e}")
            logger.error(f"Problematic response text (first 500 chars): {response_text[:500]}")
            
            # Try one more time with aggressive cleaning
            try:
                # Extract JSON array more aggressively
                import re
                json_match = re.search(r'\[.*\]', response_text, re.DOTALL)
                if json_match:
                    cleaned_json = json_match.group(0)
                    slide_data = json.loads(cleaned_json)
                    validated_slides = []
                    for slide in slide_data:
                        validated_slide = self._validate_slide(slide)
                        validated_slides.append(validated_slide)
                    logger.info(f"Successfully parsed {len(validated_slides)} slides after aggressive cleaning")
                    return validated_slides
            except:
                pass
            
            # Last resort: fallback
            logger.warning("Using fallback slide generation")
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
        
        # IMPORTANT: Title slides should NOT have content - only title and subtitle
        if validated["slide_type"] == "title":
            # For title slides, ensure content is empty
            validated["content"] = []
            return validated
        
        # Process content for non-title slides - Check if there are separator markers
        content = slide.get("content", [])
        has_separators = False
        if isinstance(content, list):
            # Check for separator markers
            separator_markers = ["[NEXT_PLACEHOLDER]", "[PLACEHOLDER]", "---", "###", "[TEXT_AREA]"]
            for item in content:
                if any(sep in str(item).upper() for sep in separator_markers):
                    has_separators = True
                    break
            
            # If has separators, don't limit content (for multi-placeholder support)
            if has_separators:
                for item in content:  # Keep all items for multi-placeholder
                    if isinstance(item, str):
                        validated["content"].append(str(item).strip()[:200])  # Still limit individual item length
            else:
                # Regular validation with limits
                for item in content[:6]:  # Limit to 6 bullet points for single placeholder
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

class AIPipeProvider(BaseLLMProvider):
    """AI Pipe provider for accessing multiple LLMs through unified API"""
    
    AVAILABLE_MODELS = [
        "openai/gpt-4o",
        "openai/gpt-4o-mini",
        "openai/gpt-4-turbo",
        "openai/gpt-3.5-turbo",
        "anthropic/claude-3-5-sonnet",
        "anthropic/claude-3-opus",
        "anthropic/claude-3-haiku",
        "google/gemini-2.0-flash-exp",
        "google/gemini-pro-1.5",
        "meta-llama/llama-3.1-70b-instruct",
        "mistralai/mixtral-8x7b-instruct"
    ]
    
    def __init__(self, api_key: str, model_name: str = "openai/gpt-4o-mini"):
        """Initialize AI Pipe provider with API token and model"""
        super().__init__(api_key, model_name)
        
        # Validate model
        if model_name not in self.AVAILABLE_MODELS:
            logger.warning(f"Model {model_name} not in known models list. Proceeding anyway.")
    
    def parse_text_to_slides(self, text_content: str, guidance: str = "", template_structure: Dict[str, Any] = None, num_slides: int = None) -> List[Dict[str, Any]]:
        """
        Parse input text into structured slide content using AI Pipe
        
        Args:
            text_content: The input text to be converted
            guidance: Optional guidance for tone/structure
            template_structure: Optional template structure information
            num_slides: Target number of slides to generate
            
        Returns:
            List of slide dictionaries with structure and content
        """
        try:
            import requests
            
            # Build the prompt
            prompt = self._build_prompt(text_content, guidance, template_structure, num_slides=num_slides)
            
            # Generate content with AI Pipe
            logger.info(f"Sending request to AI Pipe API ({self.model_name})")
            
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            
            payload = {
                "model": self.model_name,
                "messages": [
                    {"role": "system", "content": "You are an expert presentation designer. Always respond with valid JSON only."},
                    {"role": "user", "content": prompt}
                ],
                "max_tokens": 4000,
                "temperature": 0.3
            }
            
            response = requests.post(
                "https://aipipe.org/openrouter/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=60
            )
            response.raise_for_status()
            
            result = response.json()
            response_text = result['choices'][0]['message']['content']
            
            # Parse the response
            slide_structure = self._parse_response(response_text)
            
            logger.info(f"Generated {len(slide_structure)} slides using AI Pipe ({self.model_name})")
            return slide_structure
            
        except Exception as e:
            logger.error(f"Error with AI Pipe API: {str(e)}")
            raise Exception(f"Failed to process text with AI Pipe ({self.model_name}): {str(e)}")
    
    def refine_content(self, mapped_content, template_structure, selected_indices):
        """
        Refine slide content for selected template slides using the LLM (AI Pipe).
        Args:
            mapped_content: List of slide dicts to refine
            template_structure: Template structure info
            selected_indices: Indices of selected slides
        Returns:
            List of refined slide dicts
        """
        import requests
        try:
            # Build a refinement prompt (reuse Gemini logic for now)
            from .llm_providers import GeminiProvider
            gemini = GeminiProvider(self.api_key)
            prompt = gemini._build_refinement_prompt(mapped_content, template_structure, selected_indices)
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            payload = {
                "model": self.model_name,
                "messages": [
                    {"role": "system", "content": "You are an expert presentation designer. Always respond with valid JSON only."},
                    {"role": "user", "content": prompt}
                ],
                "max_tokens": 4000,
                "temperature": 0.3
            }
            response = requests.post(
                "https://aipipe.org/openrouter/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=60
            )
            response.raise_for_status()
            result = response.json()
            response_text = result['choices'][0]['message']['content']
            refined_slides = self._parse_response(response_text)
            return refined_slides
        except Exception as e:
            logger.error(f"AI Pipe refinement failed: {str(e)}")
            return mapped_content

    @classmethod
    def get_available_models(cls) -> List[str]:
        """Get list of available AI Pipe models"""
        return cls.AVAILABLE_MODELS.copy()

class OpenAIProvider(BaseLLMProvider):
    """OpenAI API provider for text parsing and slide generation"""
    
    AVAILABLE_MODELS = [
        "gpt-4o",
        "gpt-4o-mini", 
        "gpt-4-turbo",
        "gpt-4",
        "gpt-3.5-turbo"
    ]
    
    def __init__(self, api_key: str, model_name: str = "gpt-4o-mini"):
        """Initialize OpenAI provider with API key and model"""
        if openai is None:
            raise ImportError("OpenAI library not installed. Install it with: pip install openai")
        
        super().__init__(api_key, model_name)
        self.client = openai.OpenAI(api_key=api_key)
        
        # Validate model
        if model_name not in self.AVAILABLE_MODELS:
            logger.warning(f"Model {model_name} not in known models list. Proceeding anyway.")
    
    def parse_text_to_slides(self, text_content: str, guidance: str = "", template_structure: Dict[str, Any] = None, num_slides: int = None) -> List[Dict[str, Any]]:
        """
        Parse input text into structured slide content using OpenAI
        
        Args:
            text_content: The input text to be converted
            guidance: Optional guidance for tone/structure
            template_structure: Optional template structure information
            
        Returns:
            List of slide dictionaries with structure and content
        """
        try:
            # Build the prompt
            prompt = self._build_prompt(text_content, guidance, template_structure, num_slides=num_slides)
            
            # Generate content with OpenAI
            logger.info(f"Sending request to OpenAI API ({self.model_name})")
            
            response = self.client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {"role": "system", "content": "You are an expert presentation designer. Always respond with valid JSON only."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=4000,
                temperature=0.3  # Lower temperature for more consistent formatting
            )
            
            response_text = response.choices[0].message.content
            
            # Parse the response
            slide_structure = self._parse_response(response_text)
            
            logger.info(f"Generated {len(slide_structure)} slides using {self.model_name}")
            return slide_structure
            
        except Exception as e:
            logger.error(f"Error with OpenAI API: {str(e)}")
            raise Exception(f"Failed to process text with OpenAI ({self.model_name}): {str(e)}")
    
    def refine_content(self, mapped_content, template_structure, selected_indices):
        """
        Refine slide content for selected template slides using the LLM (OpenAI).
        Args:
            mapped_content: List of slide dicts to refine
            template_structure: Template structure info
            selected_indices: Indices of selected slides
        Returns:
            List of refined slide dicts
        """
        try:
            from .llm_providers import GeminiProvider
            gemini = GeminiProvider(self.api_key)
            prompt = gemini._build_refinement_prompt(mapped_content, template_structure, selected_indices)
            response = self.client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {"role": "system", "content": "You are an expert presentation designer. Always respond with valid JSON only."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=4000,
                temperature=0.3
            )
            response_text = response.choices[0].message.content
            refined_slides = self._parse_response(response_text)
            return refined_slides
        except Exception as e:
            logger.error(f"OpenAI refinement failed: {str(e)}")
            return mapped_content

    @classmethod
    def get_available_models(cls) -> List[str]:
        """Get list of available OpenAI models"""
        return cls.AVAILABLE_MODELS.copy()
