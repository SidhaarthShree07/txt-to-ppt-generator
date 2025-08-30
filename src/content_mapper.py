import logging
from typing import Dict, List, Any, Tuple, Optional
import re

logger = logging.getLogger(__name__)

class ContentMapper:
    """Maps AI-generated content to best-matching template slides"""
    
    def __init__(self):
        self.template_info = None
        self.content_slides = None
    
    def map_content_to_template(self, ai_content: List[Dict[str, Any]], 
                               template_info: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], List[int]]:
        """
        Map AI-generated content to best template slides
        
        Args:
            ai_content: List of slide content from AI
            template_info: Template analysis results
            
        Returns:
            Tuple of (mapped_content, selected_slide_indices)
        """
        self.template_info = template_info
        self.content_slides = ai_content
        
        existing_slides = template_info.get('existing_slides', [])
        
        # Create mapping between content and template slides
        mapping = []
        used_indices = set()
        
        # SPECIAL CASE: Always try to use slide 0 for the title slide if we have one
        title_slides = [s for s in ai_content if s.get('slide_type') == 'title']
        if title_slides and len(existing_slides) > 0:
            # Force first slide to be used for title
            mapping.append({
                'content': title_slides[0],
                'template_slide_index': 0,
                'template_slide': existing_slides[0]
            })
            used_indices.add(0)
            # Remove the title slide from further processing
            ai_content = [s for s in ai_content if s != title_slides[0]]
        
        # Map remaining slides
        for content_slide in ai_content:
            best_match = self._find_best_template_match(
                content_slide, 
                existing_slides, 
                used_indices
            )
            
            if best_match is not None:
                mapping.append({
                    'content': content_slide,
                    'template_slide_index': best_match,
                    'template_slide': existing_slides[best_match]
                })
                used_indices.add(best_match)
        
        # Prepare final content with template constraints
        final_content = []
        selected_indices = []
        
        for item in mapping:
            adjusted_content = self._adjust_content_for_template(
                item['content'],
                item['template_slide']
            )
            final_content.append(adjusted_content)
            selected_indices.append(item['template_slide_index'])
        
        return final_content, selected_indices
    
    def _find_best_template_match(self, content_slide: Dict[str, Any], 
                                 template_slides: List[Dict[str, Any]], 
                                 used_indices: set) -> Optional[int]:
        """Find the best template slide for given content"""
        
        slide_type = content_slide.get('slide_type', 'content')
        has_subtitle = bool(content_slide.get('subtitle'))
        content_items = content_slide.get('content', [])
        content_count = len(content_items) if content_items else 0
        
        best_score = -1
        best_index = None
        
        for idx, template_slide in enumerate(template_slides):
            if idx in used_indices:
                continue
            
            score = 0
            
            # CRITICAL: For title slides, strongly prefer the first slide (index 0)
            if slide_type == 'title':
                if idx == 0:
                    score += 100  # Heavily weight the first slide for title
                elif idx < 3:
                    score += 20  # Still consider early slides but much lower
                    
                # Also check if it's marked as a title slide
                if template_slide.get('suggested_content_type') == 'title':
                    score += 50
                    
                # Title slides should have title and subtitle placeholders
                if template_slide.get('has_title') and template_slide.get('has_subtitle'):
                    score += 30
            else:
                # For non-title slides, match slide type
                if template_slide.get('suggested_content_type') == slide_type:
                    score += 20
                
                # Avoid using the first slide for non-title content
                if idx == 0:
                    score -= 50
            
            # Match title capability
            if template_slide.get('has_title'):
                score += 5
            
            # Match subtitle capability
            if has_subtitle and template_slide.get('has_subtitle'):
                score += 5
            elif not has_subtitle and not template_slide.get('has_subtitle'):
                score += 3
            
            # Match content capability for content slides
            if slide_type != 'title' and template_slide.get('has_content'):
                # Check if content format matches
                content_format = template_slide.get('content_format')
                
                if content_format == 'numbered_list' and self._has_numbered_content(content_items):
                    score += 8
                elif content_format == 'bullet_list' and content_count > 0:
                    score += 6
                elif content_format == 'paragraph' and content_count <= 2:
                    score += 4
                
                # Check content capacity
                content_placeholders = [p for p in template_slide.get('placeholders', []) 
                                       if 'CONTENT' in p.get('type', '') or 'BODY' in p.get('type', '')]
                
                if content_placeholders:
                    placeholder = content_placeholders[0]
                    suggested_lines = placeholder.get('suggested_lines', 5)
                    
                    if abs(suggested_lines - content_count) <= 2:
                        score += 5
                    elif suggested_lines >= content_count:
                        score += 3
            
            # Prefer later slides for conclusion
            if slide_type == 'conclusion' and idx >= len(template_slides) - 3:
                score += 10
            
            if score > best_score:
                best_score = score
                best_index = idx
        
        return best_index
    
    def _has_numbered_content(self, content_items: List[str]) -> bool:
        """Check if content has numbered list pattern"""
        if not content_items:
            return False
        
        numbered_pattern = re.compile(r'^(\d+[\.\)]\s|[a-z][\.\)]\s)', re.IGNORECASE)
        numbered_count = sum(1 for item in content_items if numbered_pattern.match(str(item)))
        
        return numbered_count >= len(content_items) / 2
    
    def _adjust_content_for_template(self, content: Dict[str, Any], 
                                    template_slide: Dict[str, Any]) -> Dict[str, Any]:
        """Adjust content to fit template slide constraints"""
        
        adjusted = content.copy()
        
        # Adjust title length
        if template_slide.get('has_title'):
            title_placeholder = next(
                (p for p in template_slide.get('placeholders', []) if 'TITLE' in p.get('type', '')),
                None
            )
            if title_placeholder:
                max_chars = title_placeholder.get('max_chars_per_line', 60)
                if adjusted.get('title') and len(adjusted['title']) > max_chars:
                    adjusted['title'] = adjusted['title'][:max_chars-3] + '...'
        
        # Adjust subtitle
        if not template_slide.get('has_subtitle'):
            adjusted['subtitle'] = None
        elif template_slide.get('has_subtitle') and adjusted.get('subtitle'):
            subtitle_placeholder = next(
                (p for p in template_slide.get('placeholders', []) if 'SUBTITLE' in p.get('type', '')),
                None
            )
            if subtitle_placeholder:
                max_chars = subtitle_placeholder.get('max_chars_per_line', 100)
                if len(adjusted['subtitle']) > max_chars:
                    adjusted['subtitle'] = adjusted['subtitle'][:max_chars-3] + '...'
        
        # Adjust content format and length
        if template_slide.get('has_content') and adjusted.get('content'):
            content_placeholder = next(
                (p for p in template_slide.get('placeholders', []) 
                 if 'CONTENT' in p.get('type', '') or 'BODY' in p.get('type', '')),
                None
            )
            
            if content_placeholder:
                suggested_lines = content_placeholder.get('suggested_lines', 5)
                max_chars_per_line = content_placeholder.get('max_chars_per_line', 80)
                text_format = content_placeholder.get('text_format')
                
                # Adjust content format
                content_items = adjusted['content']
                
                if text_format == 'numbered_list':
                    # Ensure numbered format
                    adjusted['content'] = [
                        f"{i+1}. {item.lstrip('0123456789.-) ')}" 
                        if not re.match(r'^\d+[\.\)]', item) else item
                        for i, item in enumerate(content_items[:suggested_lines])
                    ]
                elif text_format == 'bullet_list':
                    # Remove any numbering
                    adjusted['content'] = [
                        re.sub(r'^\d+[\.\)]\s*', '', item)
                        for item in content_items[:suggested_lines]
                    ]
                elif text_format == 'paragraph':
                    # Combine into paragraph if multiple items
                    if len(content_items) > 1:
                        adjusted['content'] = [' '.join(content_items[:2])]
                
                # Trim to character limits
                adjusted['content'] = [
                    item[:max_chars_per_line-3] + '...' if len(item) > max_chars_per_line else item
                    for item in adjusted['content']
                ]
        
        # Add template-specific metadata
        adjusted['_template_slide_index'] = template_slide['slide_index']
        adjusted['_template_format'] = template_slide.get('content_format')
        
        return adjusted
    
    def refine_content_with_ai(self, mapped_content: List[Dict[str, Any]], 
                              template_info: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Prepare refinement instructions for AI to adjust content
        
        Args:
            mapped_content: Content already mapped to template slides
            template_info: Template analysis results
            
        Returns:
            Refinement instructions for AI
        """
        refinements = []
        
        for content in mapped_content:
            slide_idx = content.get('_template_slide_index')
            template_slide = template_info['existing_slides'][slide_idx]
            
            refinement = {
                'original_content': content,
                'slide_number': slide_idx + 1,
                'constraints': {}
            }
            
            # Add specific constraints
            for placeholder in template_slide.get('placeholders', []):
                ph_type = placeholder.get('type', '')
                
                if 'TITLE' in ph_type:
                    refinement['constraints']['title'] = {
                        'max_chars': placeholder.get('max_chars_per_line', 60),
                        'current_length': len(content.get('title', ''))
                    }
                elif 'SUBTITLE' in ph_type:
                    refinement['constraints']['subtitle'] = {
                        'max_chars': placeholder.get('max_chars_per_line', 100),
                        'required': template_slide.get('has_subtitle', False)
                    }
                elif 'CONTENT' in ph_type or 'BODY' in ph_type:
                    refinement['constraints']['content'] = {
                        'format': placeholder.get('text_format', 'bullet_list'),
                        'max_lines': placeholder.get('suggested_lines', 5),
                        'max_chars_per_line': placeholder.get('max_chars_per_line', 80),
                        'current_items': len(content.get('content', []))
                    }
            
            refinements.append(refinement)
        
        return refinements
