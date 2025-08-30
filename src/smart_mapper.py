"""
Smart mapper to match AI generated content to template slides based on format matching
"""
import logging
from typing import List, Dict, Any, Optional, Tuple
from .format_detector import detect_content_format, get_content_placeholders_from_template_slide

logger = logging.getLogger(__name__)

class SmartMapper:
    """Maps AI-generated content to template slides based on format matching"""
    
    def map_content_to_template(self, 
                               ai_slides: List[Dict[str, Any]], 
                               template_info: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], List[int]]:
        """
        Map AI-generated slides to template slides based on format matching
        
        Returns:
            - List of mapped content (with template slide info embedded)
            - List of template slide indices that were selected
        """
        template_slides = template_info.get('existing_slides', [])
        if not template_slides:
            return ai_slides, list(range(len(ai_slides)))
        
        # Analyze formats
        ai_formats = [self._analyze_ai_slide(slide) for slide in ai_slides]
        template_formats = [self._analyze_template_slide(slide) for slide in template_slides]
        
        # Map slides
        mappings = []
        used_template_indices = set()
        
        # First pass: Map title slide (always first)
        if ai_slides and ai_slides[0].get('slide_type') == 'title':
            # Find best title slide in template
            title_idx = self._find_title_slide(template_slides)
            if title_idx is not None:
                mappings.append({
                    'ai_slide': ai_slides[0],
                    'ai_idx': 0,
                    'template_idx': title_idx,
                    'template_slide': template_slides[title_idx]
                })
                used_template_indices.add(title_idx)
        
        # Second pass: Map content slides by format matching
        for i, ai_slide in enumerate(ai_slides):
            if i == 0 and mappings:  # Skip if already mapped as title
                continue
            
            ai_fmt = ai_formats[i]
            best_match_idx = self._find_best_format_match(
                ai_fmt, template_formats, used_template_indices, template_slides
            )
            
            if best_match_idx is not None:
                mappings.append({
                    'ai_slide': ai_slide,
                    'ai_idx': i,
                    'template_idx': best_match_idx,
                    'template_slide': template_slides[best_match_idx]
                })
                used_template_indices.add(best_match_idx)
        
        # Third pass: Map remaining AI slides to unused template slides
        unused_template_indices = [i for i in range(len(template_slides)) 
                                  if i not in used_template_indices]
        unmapped_ai_indices = [i for i in range(len(ai_slides))
                              if not any(m['ai_idx'] == i for m in mappings)]
        
        for ai_idx, template_idx in zip(unmapped_ai_indices, unused_template_indices):
            mappings.append({
                'ai_slide': ai_slides[ai_idx],
                'ai_idx': ai_idx,
                'template_idx': template_idx,
                'template_slide': template_slides[template_idx]
            })
            used_template_indices.add(template_idx)
        
        # Sort mappings by template index to maintain order
        mappings.sort(key=lambda x: x['template_idx'])
        
        # Build final mapped content
        mapped_content = []
        selected_indices = []
        
        for mapping in mappings:
            content = mapping['ai_slide'].copy()
            content['_template_slide_index'] = mapping['template_idx']
            content['_template_slide_info'] = mapping['template_slide']
            mapped_content.append(content)
            selected_indices.append(mapping['template_idx'])
        
        logger.info(f"Mapped {len(mapped_content)} AI slides to template slides: {selected_indices}")
        
        return mapped_content, selected_indices
    
    def _analyze_ai_slide(self, slide: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze format of AI-generated slide"""
        content = slide.get('content', [])
        return {
            'slide_type': slide.get('slide_type', 'content'),
            'has_title': bool(slide.get('title')),
            'has_subtitle': bool(slide.get('subtitle')),
            'content_format': detect_content_format(content),
            'content_count': len([x for x in content if x and str(x).strip()]),
            'is_title_slide': slide.get('slide_type') == 'title',
            'is_conclusion': slide.get('slide_type') == 'conclusion'
        }
    
    def _analyze_template_slide(self, slide: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze format capabilities of template slide"""
        content_phs = get_content_placeholders_from_template_slide(slide)
        
        # Aggregate format from content placeholders
        content_format = 'bullet_list'
        total_capacity = 0
        
        for ph in content_phs:
            if 'text_format' in ph:
                content_format = ph['text_format']
            total_capacity += ph.get('suggested_lines', 5)
        
        return {
            'slide_type': slide.get('suggested_content_type', 'content'),
            'has_title': slide.get('has_title', False),
            'has_subtitle': slide.get('has_subtitle', False),
            'content_format': slide.get('content_format', content_format),
            'content_placeholder_count': len(content_phs),
            'total_content_capacity': total_capacity,
            'is_title_capable': slide.get('has_title') and slide.get('has_subtitle'),
            'layout_name': slide.get('layout_name', '')
        }
    
    def _find_title_slide(self, template_slides: List[Dict[str, Any]]) -> Optional[int]:
        """Find best title slide in template"""
        for i, slide in enumerate(template_slides):
            if slide.get('suggested_content_type') == 'title':
                return i
            if slide.get('has_title') and slide.get('has_subtitle') and not slide.get('has_content'):
                return i
            if 'title' in slide.get('layout_name', '').lower():
                return i
        
        # Fallback: first slide with title and subtitle
        for i, slide in enumerate(template_slides):
            if slide.get('has_title') and slide.get('has_subtitle'):
                return i
        
        return 0  # Default to first slide
    
    def _find_best_format_match(self, 
                                ai_format: Dict[str, Any],
                                template_formats: List[Dict[str, Any]], 
                                used_indices: set,
                                template_slides: List[Dict[str, Any]]) -> Optional[int]:
        """Find best matching template slide for AI content format"""
        best_score = -1
        best_idx = None
        
        for i, template_fmt in enumerate(template_formats):
            if i in used_indices:
                continue
            
            score = 0
            
            # Format matching
            if ai_format['content_format'] == template_fmt['content_format']:
                score += 10
            
            # Slide type matching
            if ai_format['slide_type'] == template_fmt['slide_type']:
                score += 5
            
            # Conclusion slides should map to later slides
            if ai_format['is_conclusion'] and i >= len(template_formats) - 3:
                score += 8
            
            # Capacity matching
            if template_fmt['total_content_capacity'] >= ai_format['content_count']:
                score += 3
            
            # Title/subtitle matching
            if ai_format['has_title'] == template_fmt['has_title']:
                score += 2
            if ai_format['has_subtitle'] == template_fmt['has_subtitle']:
                score += 2
            
            # Multiple placeholders bonus for content with separators
            if template_fmt['content_placeholder_count'] > 1 and ai_format['content_count'] > 3:
                score += 4
            
            if score > best_score:
                best_score = score
                best_idx = i
        
        return best_idx
