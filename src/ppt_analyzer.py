import logging
from typing import Dict, List, Any, Tuple, Optional
from pptx import Presentation
from pptx.slide import Slide
from pptx.shapes.base import BaseShape
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
import tempfile
import os

logger = logging.getLogger(__name__)

class PowerPointAnalyzer:
    """Analyzes PowerPoint templates to extract styles, layouts, and assets"""
    
    def __init__(self):
        self.template_info = {}
    
    def analyze_template(self, template_path: str) -> Dict[str, Any]:
        """
        Analyze a PowerPoint template file and extract styling information
        
        Args:
            template_path: Path to the PowerPoint template file
            
        Returns:
            Dictionary containing template analysis results
        """
        try:
            logger.info(f"Analyzing template: {template_path}")
            
            # Load the presentation
            presentation = Presentation(template_path)
            
            analysis = {
                'slide_layouts': self._analyze_slide_layouts(presentation),
                'theme_colors': self._extract_theme_colors(presentation),
                'fonts': self._extract_fonts(presentation),
                'images': self._extract_images(presentation),
                'slide_count': len(presentation.slides),
                'layout_count': len(presentation.slide_layouts),
                'master_slide': self._analyze_slide_master(presentation),
                'presentation_object': presentation  # Keep reference for generation
            }
            
            logger.info(f"Template analysis complete: {analysis['slide_count']} slides, {analysis['layout_count']} layouts")
            return analysis
            
        except Exception as e:
            logger.error(f"Error analyzing template: {str(e)}")
            raise Exception(f"Failed to analyze PowerPoint template: {str(e)}")
    
    def _analyze_slide_layouts(self, presentation: Presentation) -> List[Dict[str, Any]]:
        """Analyze available slide layouts"""
        layouts = []
        
        for i, layout in enumerate(presentation.slide_layouts):
            layout_info = {
                'index': i,
                'name': layout.name,
                'placeholders': self._analyze_placeholders(layout),
                'width': presentation.slide_width,
                'height': presentation.slide_height
            }
            layouts.append(layout_info)
        
        return layouts
    
    def _analyze_placeholders(self, layout) -> List[Dict[str, Any]]:
        """Analyze placeholders in a slide layout"""
        placeholders = []
        
        for placeholder in layout.placeholders:
            try:
                placeholder_info = {
                    'index': placeholder.placeholder_format.idx,
                    'type': str(placeholder.placeholder_format.type),
                    'name': getattr(placeholder, 'name', f'Placeholder {placeholder.placeholder_format.idx}'),
                    'left': placeholder.left,
                    'top': placeholder.top,
                    'width': placeholder.width,
                    'height': placeholder.height
                }
                placeholders.append(placeholder_info)
            except Exception as e:
                logger.warning(f"Could not analyze placeholder: {e}")
                continue
        
        return placeholders
    
    def _extract_theme_colors(self, presentation: Presentation) -> Dict[str, Any]:
        """Extract theme colors from the presentation"""
        try:
            theme_part = presentation.part.package.part_related_by(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
            )
            
            # Basic color extraction - this is simplified
            # In a full implementation, you'd parse the theme XML
            colors = {
                'has_theme': theme_part is not None,
                'extracted_colors': []
            }
            
            # Try to extract colors from existing slides
            if presentation.slides:
                sample_colors = self._sample_colors_from_slides(presentation)
                colors['sample_colors'] = sample_colors
            
            return colors
            
        except Exception as e:
            logger.warning(f"Could not extract theme colors: {e}")
            return {'has_theme': False, 'extracted_colors': []}
    
    def _sample_colors_from_slides(self, presentation: Presentation) -> List[str]:
        """Sample colors from existing slides"""
        colors = set()
        
        try:
            for slide in list(presentation.slides)[:3]:  # Sample first 3 slides
                for shape in slide.shapes:
                    try:
                        if hasattr(shape, 'fill') and shape.fill.type is not None:
                            if hasattr(shape.fill, 'fore_color'):
                                color = shape.fill.fore_color.rgb
                                if color:
                                    colors.add(f"#{color}")
                    except Exception:
                        continue
        except Exception as e:
            logger.warning(f"Error sampling colors: {e}")
        
        return list(colors)[:10]  # Return up to 10 colors
    
    def _extract_fonts(self, presentation: Presentation) -> Dict[str, Any]:
        """Extract font information from the presentation"""
        fonts = set()
        
        try:
            # Sample fonts from existing slides
            for slide in list(presentation.slides)[:3]:  # Sample first 3 slides
                for shape in slide.shapes:
                    if hasattr(shape, 'text_frame'):
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if run.font.name:
                                    fonts.add(run.font.name)
                                    
        except Exception as e:
            logger.warning(f"Error extracting fonts: {e}")
        
        return {
            'fonts_found': list(fonts),
            'default_font': list(fonts)[0] if fonts else 'Calibri'
        }
    
    def _extract_images(self, presentation: Presentation) -> List[Dict[str, Any]]:
        """Extract image information from the presentation"""
        images = []
        
        try:
            for slide_num, slide in enumerate(presentation.slides):
                slide_images = self._extract_images_from_slide(slide, slide_num)
                images.extend(slide_images)
                
        except Exception as e:
            logger.warning(f"Error extracting images: {e}")
        
        return images
    
    def _extract_images_from_slide(self, slide: Slide, slide_num: int) -> List[Dict[str, Any]]:
        """Extract images from a specific slide"""
        slide_images = []
        
        try:
            for shape_num, shape in enumerate(slide.shapes):
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    try:
                        image_info = {
                            'slide_index': slide_num,
                            'shape_index': shape_num,
                            'left': shape.left,
                            'top': shape.top,
                            'width': shape.width,
                            'height': shape.height,
                            'image_data': shape.image.blob,
                            'filename': getattr(shape.image, 'filename', f'image_{slide_num}_{shape_num}.png')
                        }
                        slide_images.append(image_info)
                    except Exception as e:
                        logger.warning(f"Could not extract image from slide {slide_num}, shape {shape_num}: {e}")
                        continue
                        
        except Exception as e:
            logger.warning(f"Error processing slide {slide_num} for images: {e}")
        
        return slide_images
    
    def _analyze_slide_master(self, presentation: Presentation) -> Dict[str, Any]:
        """Analyze the slide master for default styling"""
        try:
            slide_master = presentation.slide_master
            
            master_info = {
                'width': presentation.slide_width,
                'height': presentation.slide_height,
                'background': self._analyze_background(slide_master),
                'placeholders': self._analyze_placeholders(slide_master)
            }
            
            return master_info
            
        except Exception as e:
            logger.warning(f"Could not analyze slide master: {e}")
            return {}
    
    def _analyze_background(self, slide_master) -> Dict[str, Any]:
        """Analyze background styling"""
        try:
            background_info = {
                'has_background': hasattr(slide_master, 'background'),
                'fill_type': None
            }
            
            if hasattr(slide_master, 'background') and slide_master.background:
                if hasattr(slide_master.background, 'fill'):
                    background_info['fill_type'] = str(slide_master.background.fill.type)
            
            return background_info
            
        except Exception as e:
            logger.warning(f"Error analyzing background: {e}")
            return {'has_background': False}
    
    def get_best_layout_for_slide_type(self, template_info: Dict[str, Any], slide_type: str) -> int:
        """
        Determine the best layout index for a given slide type
        
        Args:
            template_info: Template analysis results
            slide_type: Type of slide (title, content, conclusion)
            
        Returns:
            Index of the best matching layout
        """
        layouts = template_info.get('slide_layouts', [])
        
        if not layouts:
            return 0
        
        # Mapping preferences for slide types
        layout_preferences = {
            'title': ['title', 'Title Slide', 'Title Only'],
            'content': ['content', 'Title and Content', 'Two Content', 'Content with Caption'],
            'conclusion': ['title', 'Title and Content', 'Title Only']
        }
        
        preferred_names = layout_preferences.get(slide_type, ['content'])
        
        # Find best matching layout by name
        for pref_name in preferred_names:
            for i, layout in enumerate(layouts):
                if pref_name.lower() in layout['name'].lower():
                    return i
        
        # Fallback: return first available layout or default
        return 0 if layouts else 0
