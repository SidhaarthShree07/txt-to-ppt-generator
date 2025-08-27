import logging
from typing import Dict, List, Any, Optional
from pptx import Presentation
from pptx.slide import Slide
from pptx.shapes.placeholder import SlidePlaceholder
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
import io
import random

logger = logging.getLogger(__name__)

class SlideGenerator:
    """Generates PowerPoint slides from structured content and template information"""
    
    def __init__(self):
        self.presentation = None
        self.template_info = None
    
    def create_presentation(self, slide_structure: List[Dict[str, Any]], 
                          template_info: Dict[str, Any], 
                          output_path: str) -> None:
        """
        Create a new PowerPoint presentation from slide structure and template info
        
        Args:
            slide_structure: List of slide data from LLM parsing
            template_info: Template analysis results
            output_path: Path where to save the generated presentation
        """
        try:
            logger.info(f"Creating presentation with {len(slide_structure)} slides")
            
            # Use the template presentation as base
            self.presentation = template_info['presentation_object']
            self.template_info = template_info
            
            # Clear existing slides (keep layouts and master)
            self._clear_existing_slides()
            
            # Generate slides based on structure
            for slide_data in slide_structure:
                self._create_slide(slide_data)
            
            # Save the presentation
            self.presentation.save(output_path)
            logger.info(f"Presentation saved to: {output_path}")
            
        except Exception as e:
            logger.error(f"Error creating presentation: {str(e)}")
            raise Exception(f"Failed to create presentation: {str(e)}")
    
    def _clear_existing_slides(self):
        """Remove existing slides while preserving layouts and master slides"""
        try:
            # Get the slide part manager
            slide_ids = [slide.slide_id for slide in self.presentation.slides]
            
            # Remove slides in reverse order to maintain indices
            for slide_id in reversed(slide_ids):
                slide_part = self.presentation.part.related_parts[slide_id]
                del self.presentation.part.related_parts[slide_id]
                
            # Clear the slides collection
            self.presentation.slides._sldIdLst.clear()
            
            logger.info("Cleared existing slides from template")
            
        except Exception as e:
            logger.warning(f"Could not clear existing slides: {e}")
            # If we can't clear, we'll work with existing slides
    
    def _create_slide(self, slide_data: Dict[str, Any]) -> Slide:
        """Create a single slide from slide data"""
        try:
            # Get the appropriate layout
            layout_index = self._get_layout_for_slide(slide_data['slide_type'])
            slide_layout = self.presentation.slide_layouts[layout_index]
            
            # Add slide
            slide = self.presentation.slides.add_slide(slide_layout)
            
            # Populate slide content
            self._populate_slide_content(slide, slide_data)
            
            # Apply template styling
            self._apply_template_styling(slide, slide_data)
            
            # Add images if appropriate
            self._add_template_images(slide, slide_data)
            
            logger.debug(f"Created slide: {slide_data['title']}")
            return slide
            
        except Exception as e:
            logger.error(f"Error creating slide '{slide_data.get('title', 'Unknown')}': {e}")
            raise
    
    def _get_layout_for_slide(self, slide_type: str) -> int:
        """Get the best layout index for a slide type"""
        layouts = self.template_info.get('slide_layouts', [])
        
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
    
    def _populate_slide_content(self, slide: Slide, slide_data: Dict[str, Any]):
        """Populate slide with text content"""
        try:
            slide_type = slide_data['slide_type']
            
            # Handle different slide types
            if slide_type == 'title':
                self._populate_title_slide(slide, slide_data)
            elif slide_type == 'content':
                self._populate_content_slide(slide, slide_data)
            elif slide_type == 'conclusion':
                self._populate_conclusion_slide(slide, slide_data)
            else:
                # Default to content slide
                self._populate_content_slide(slide, slide_data)
                
        except Exception as e:
            logger.warning(f"Error populating slide content: {e}")
            # Create basic text if content population fails
            self._create_fallback_content(slide, slide_data)
    
    def _populate_title_slide(self, slide: Slide, slide_data: Dict[str, Any]):
        """Populate a title slide"""
        try:
            # Find title placeholder
            title_placeholder = None
            subtitle_placeholder = None
            
            for shape in slide.shapes:
                if shape.is_placeholder:
                    placeholder_type = str(shape.placeholder_format.type)
                    if 'TITLE' in placeholder_type.upper():
                        title_placeholder = shape
                    elif 'SUBTITLE' in placeholder_type.upper() or 'CONTENT' in placeholder_type.upper():
                        subtitle_placeholder = shape
            
            # Set title
            if title_placeholder and slide_data['title']:
                title_placeholder.text = slide_data['title']
                self._format_title_text(title_placeholder)
            
            # Set subtitle
            if subtitle_placeholder and slide_data['subtitle']:
                subtitle_placeholder.text = slide_data['subtitle']
                self._format_subtitle_text(subtitle_placeholder)
                
        except Exception as e:
            logger.warning(f"Error populating title slide: {e}")
    
    def _populate_content_slide(self, slide: Slide, slide_data: Dict[str, Any]):
        """Populate a content slide with title and bullet points"""
        try:
            title_placeholder = None
            content_placeholder = None
            
            # Find placeholders
            for shape in slide.shapes:
                if shape.is_placeholder:
                    placeholder_type = str(shape.placeholder_format.type)
                    if 'TITLE' in placeholder_type.upper():
                        title_placeholder = shape
                    elif 'CONTENT' in placeholder_type.upper() or 'BODY' in placeholder_type.upper():
                        content_placeholder = shape
            
            # Set title
            if title_placeholder and slide_data['title']:
                title_placeholder.text = slide_data['title']
                self._format_title_text(title_placeholder)
            
            # Set content
            if content_placeholder and slide_data['content']:
                self._populate_bullet_points(content_placeholder, slide_data['content'])
                
        except Exception as e:
            logger.warning(f"Error populating content slide: {e}")
    
    def _populate_conclusion_slide(self, slide: Slide, slide_data: Dict[str, Any]):
        """Populate a conclusion slide"""
        # Conclusion slides are similar to content slides
        self._populate_content_slide(slide, slide_data)
    
    def _populate_bullet_points(self, placeholder, content_list: List[str]):
        """Add bullet points to a content placeholder"""
        try:
            if not placeholder.has_text_frame:
                return
            
            text_frame = placeholder.text_frame
            text_frame.clear()  # Clear existing content
            
            for i, bullet_text in enumerate(content_list):
                if i == 0:
                    # First paragraph
                    paragraph = text_frame.paragraphs[0]
                else:
                    # Add new paragraphs
                    paragraph = text_frame.add_paragraph()
                
                paragraph.text = bullet_text
                paragraph.level = 0  # Top level bullet
                
                # Format the paragraph
                self._format_bullet_text(paragraph)
                
        except Exception as e:
            logger.warning(f"Error adding bullet points: {e}")
    
    def _format_title_text(self, placeholder):
        """Apply formatting to title text"""
        try:
            if not placeholder.has_text_frame:
                return
                
            for paragraph in placeholder.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(44)  # Large title size
                    run.font.bold = True
                    # Apply template font if available
                    default_font = self.template_info.get('fonts', {}).get('default_font', 'Calibri')
                    run.font.name = default_font
                    
        except Exception as e:
            logger.warning(f"Error formatting title text: {e}")
    
    def _format_subtitle_text(self, placeholder):
        """Apply formatting to subtitle text"""
        try:
            if not placeholder.has_text_frame:
                return
                
            for paragraph in placeholder.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(24)  # Medium subtitle size
                    run.font.bold = False
                    # Apply template font if available
                    default_font = self.template_info.get('fonts', {}).get('default_font', 'Calibri')
                    run.font.name = default_font
                    
        except Exception as e:
            logger.warning(f"Error formatting subtitle text: {e}")
    
    def _format_bullet_text(self, paragraph):
        """Apply formatting to bullet point text"""
        try:
            for run in paragraph.runs:
                run.font.size = Pt(18)  # Body text size
                run.font.bold = False
                # Apply template font if available
                default_font = self.template_info.get('fonts', {}).get('default_font', 'Calibri')
                run.font.name = default_font
                
        except Exception as e:
            logger.warning(f"Error formatting bullet text: {e}")
    
    def _apply_template_styling(self, slide: Slide, slide_data: Dict[str, Any]):
        """Apply template colors and styling to the slide"""
        try:
            # This is where you would apply colors, fonts, and other styling
            # from the template analysis. For now, we rely on the layout's
            # built-in styling which is preserved from the template.
            
            sample_colors = self.template_info.get('theme_colors', {}).get('sample_colors', [])
            if sample_colors and random.random() > 0.7:  # Occasionally apply accent colors
                self._apply_accent_color(slide, sample_colors[0])
                
        except Exception as e:
            logger.warning(f"Error applying template styling: {e}")
    
    def _apply_accent_color(self, slide: Slide, color_hex: str):
        """Apply an accent color to text elements"""
        try:
            # Parse hex color
            if color_hex.startswith('#'):
                color_hex = color_hex[1:]
            
            if len(color_hex) == 6:
                r = int(color_hex[0:2], 16)
                g = int(color_hex[2:4], 16)  
                b = int(color_hex[4:6], 16)
                rgb_color = RGBColor(r, g, b)
                
                # Apply to some text elements
                for shape in slide.shapes:
                    if shape.has_text_frame and random.random() > 0.8:  # Occasionally
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if run.text.strip():  # Only if there's text
                                    run.font.color.rgb = rgb_color
                                    break  # Only first run
                            break  # Only first paragraph
                            
        except Exception as e:
            logger.warning(f"Error applying accent color: {e}")
    
    def _add_template_images(self, slide: Slide, slide_data: Dict[str, Any]):
        """Add images from template to appropriate slides"""
        try:
            template_images = self.template_info.get('images', [])
            
            # For title slides, try to add a representative image
            if slide_data['slide_type'] == 'title' and template_images:
                self._add_background_image(slide, template_images[0])
            
            # For content slides, occasionally add small decorative images
            elif slide_data['slide_type'] == 'content' and template_images and random.random() > 0.7:
                self._add_decorative_image(slide, random.choice(template_images))
                
        except Exception as e:
            logger.warning(f"Error adding template images: {e}")
    
    def _add_background_image(self, slide: Slide, image_info: Dict[str, Any]):
        """Add an image as a background or large element"""
        try:
            # Add image to slide
            image_stream = io.BytesIO(image_info['image_data'])
            
            # Position it in the background (bottom-right corner, smaller)
            left = Inches(8)  # Right side
            top = Inches(5)   # Bottom area
            width = Inches(2) # Smaller size
            height = Inches(1.5)
            
            slide.shapes.add_picture(image_stream, left, top, width, height)
            
        except Exception as e:
            logger.warning(f"Error adding background image: {e}")
    
    def _add_decorative_image(self, slide: Slide, image_info: Dict[str, Any]):
        """Add a small decorative image to the slide"""
        try:
            image_stream = io.BytesIO(image_info['image_data'])
            
            # Position it as a small decorative element
            left = Inches(8.5)  # Far right
            top = Inches(1)     # Top area
            width = Inches(1)   # Small size
            height = Inches(0.75)
            
            slide.shapes.add_picture(image_stream, left, top, width, height)
            
        except Exception as e:
            logger.warning(f"Error adding decorative image: {e}")
    
    def _create_fallback_content(self, slide: Slide, slide_data: Dict[str, Any]):
        """Create basic text content if normal population fails"""
        try:
            # Add a text box with the slide content
            left = Inches(1)
            top = Inches(1.5)
            width = Inches(8)
            height = Inches(5)
            
            textbox = slide.shapes.add_textbox(left, top, width, height)
            text_frame = textbox.text_frame
            
            # Add title
            if slide_data['title']:
                title_paragraph = text_frame.paragraphs[0]
                title_paragraph.text = slide_data['title']
                title_paragraph.font.size = Pt(32)
                title_paragraph.font.bold = True
            
            # Add content
            for content_item in slide_data.get('content', []):
                paragraph = text_frame.add_paragraph()
                paragraph.text = f"• {content_item}"
                paragraph.font.size = Pt(18)
                
        except Exception as e:
            logger.error(f"Error creating fallback content: {e}")
