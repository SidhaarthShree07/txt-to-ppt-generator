"""
Robust pipeline orchestrating end-to-end generation, mapping, refinement, and replacement
"""
import logging
from typing import Dict, Any, List, Tuple
from .smart_mapper import SmartMapper
from .slide_refiner import SlideRefiner
from .format_detector import has_separators

try:
    from .simple_slide_replacer import replace_slide_content_simple, clear_all_placeholder_text
except ImportError:
    from simple_slide_replacer import replace_slide_content_simple, clear_all_placeholder_text

logger = logging.getLogger(__name__)

class RobustSlidePipeline:
    """End-to-end pipeline that follows the robust multi-step process"""

    def __init__(self, llm_provider):
        self.llm_provider = llm_provider

    def run(self, input_text: str, template_info: Dict[str, Any], presentation, output_path: str, guidance: str = "", num_slides: int = None, reuse_images: bool = False) -> Tuple[List[Dict[str, Any]], List[int]]:
        """
        Execute the robust pipeline and write to the provided presentation object
        Returns the final mapped content and selected indices
        """
        # 1) Initial generation
        logger.info("Generating initial slides from LLM")
        try:
            initial_slides = self.llm_provider.parse_text_to_slides(input_text, guidance, None, num_slides=num_slides)
        except Exception as e:
            logger.error(f"Initial generation failed: {e}")
            raise

        # Ensure first slide is a title slide
        if not initial_slides or initial_slides[0].get('slide_type') != 'title':
            title = (initial_slides[0].get('title') if initial_slides else 'Presentation')
            initial_slides = [{
                'slide_type': 'title',
                'title': title or 'Presentation',
                'subtitle': 'Overview'
            }] + initial_slides

        # 2) Smart mapping to template
        mapper = SmartMapper()
        mapped_content, selected_indices = mapper.map_content_to_template(initial_slides, template_info)

        # 3) Parallel per-slide refinement to match placeholder capacities
        refiner = SlideRefiner(self.llm_provider)
        refined_slides = refiner.refine_slides_parallel(mapped_content)

        # 4) Apply content to presentation slides
        existing_slides = list(presentation.slides)
        used_indices = set(selected_indices)

        for content, slide_idx in zip(refined_slides, selected_indices):
            if slide_idx < len(existing_slides):
                slide = existing_slides[slide_idx]
                clear_all_placeholder_text(slide)
                ok = replace_slide_content_simple(slide, content)
                if not ok:
                    logger.warning(f"Fallback replacer failed for slide {slide_idx+1}")

        # 5) Delete unused slides
        for i in range(len(existing_slides) - 1, -1, -1):
            if i not in used_indices:
                try:
                    presentation.slides._sldIdLst.remove(presentation.slides._sldIdLst[i])
                except Exception:
                    pass

        # 6) Validate all placeholders got filled
        self._validate_and_fill_missing(presentation)

        # 7) Save
        presentation.save(output_path)
        logger.info(f"Saved presentation to {output_path}")

        return refined_slides, selected_indices

    def _validate_and_fill_missing(self, presentation):
        for slide in presentation.slides:
            for shape in slide.shapes:
                try:
                    if getattr(shape, 'is_placeholder', False) and getattr(shape, 'has_text_frame', False):
                        text = (shape.text or '').strip()
                        if not text:
                            # Fill with a minimal placeholder to avoid empty boxes
                            shape.text = " "
                except Exception:
                    continue

