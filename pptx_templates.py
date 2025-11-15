"""
PowerPoint Template Management
Work with existing templates and create professional layouts
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from typing import Dict, List, Optional, Tuple
import json
from pathlib import Path
from PIL import Image
from io import BytesIO


class TemplateManager:
    """Manage PowerPoint templates"""

    def __init__(self, template_path: Optional[str] = None):
        """
        Initialize template manager

        Args:
            template_path: Path to existing .pptx template (None for blank)
        """
        if template_path:
            self.prs = Presentation(template_path)
        else:
            self.prs = Presentation()

    def get_slide_inventory(self) -> List[Dict]:
        """
        Get inventory of all slides with metadata

        Returns:
            List of dicts with slide information
        """
        inventory = []

        for idx, slide in enumerate(self.prs.slides):
            slide_info = {
                'index': idx,
                'layout_name': slide.slide_layout.name,
                'shapes': []
            }

            for shape in slide.shapes:
                shape_info = {
                    'name': shape.name,
                    'type': shape.shape_type,
                    'left': shape.left.pt,
                    'top': shape.top.pt,
                    'width': shape.width.pt,
                    'height': shape.height.pt
                }

                # Extract text if present
                if hasattr(shape, 'text'):
                    shape_info['text'] = shape.text

                slide_info['shapes'].append(shape_info)

            inventory.append(slide_info)

        return inventory

    def save_inventory(self, output_path: str):
        """Save slide inventory to JSON"""
        inventory = self.get_slide_inventory()
        with open(output_path, 'w') as f:
            json.dump(inventory, f, indent=2)

    def duplicate_slide(self, slide_index: int) -> int:
        """
        Duplicate a slide

        Args:
            slide_index: Index of slide to duplicate

        Returns:
            Index of new slide
        """
        source_slide = self.prs.slides[slide_index]
        layout = source_slide.slide_layout

        # Create new slide with same layout
        new_slide = self.prs.slides.add_slide(layout)

        # Copy shapes
        for shape in source_slide.shapes:
            el = shape.element
            newel = el.__class__()
            newel._element = el._element
            new_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')

        return len(self.prs.slides) - 1

    def rearrange_slides(self, order: List[int]):
        """
        Rearrange slides in specified order

        Args:
            order: List of slide indices in desired order
        """
        # This is complex in python-pptx, requires XML manipulation
        # For now, we'll document this limitation
        print("Note: Slide rearrangement requires manual XML manipulation")
        print(f"Desired order: {order}")

    def replace_text(self, replacements: Dict[str, str], slide_indices: Optional[List[int]] = None):
        """
        Replace text across slides

        Args:
            replacements: Dict of {old_text: new_text}
            slide_indices: Specific slides to modify (None for all)
        """
        slides = [self.prs.slides[i] for i in slide_indices] if slide_indices else self.prs.slides

        for slide in slides:
            for shape in slide.shapes:
                if hasattr(shape, 'text_frame'):
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            for old_text, new_text in replacements.items():
                                if old_text in run.text:
                                    run.text = run.text.replace(old_text, new_text)

    def save(self, output_path: str):
        """Save presentation"""
        self.prs.save(output_path)


class ModernTemplates:
    """Create modern, professional slide templates"""

    @staticmethod
    def create_title_slide(prs: Presentation, title: str, subtitle: str,
                          color_scheme: Tuple[int, int, int] = (79, 70, 229)) -> None:
        """Create modern title slide"""
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # Background with gradient (will be added via image)
        from pptx_images import ImageProcessor
        gradient = ImageProcessor.create_gradient_image(
            int(prs.slide_width.inches * 150),
            int(prs.slide_height.inches * 150),
            color_scheme,
            tuple(max(0, c - 50) for c in color_scheme),
            angle=135
        )

        img_stream = BytesIO()
        gradient.save(img_stream, format='PNG')
        img_stream.seek(0)

        slide.shapes.add_picture(img_stream, 0, 0,
                                prs.slide_width, prs.slide_height)

        # Title
        title_box = slide.shapes.add_textbox(
            Inches(1), Inches(2),
            prs.slide_width - Inches(2), Inches(1.5)
        )
        title_frame = title_box.text_frame
        title_frame.text = title
        title_para = title_frame.paragraphs[0]
        title_para.alignment = PP_ALIGN.CENTER
        title_para.font.size = Pt(54)
        title_para.font.bold = True
        title_para.font.color.rgb = RGBColor(255, 255, 255)

        # Subtitle
        subtitle_box = slide.shapes.add_textbox(
            Inches(1), Inches(3.5),
            prs.slide_width - Inches(2), Inches(1)
        )
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = subtitle
        subtitle_para = subtitle_frame.paragraphs[0]
        subtitle_para.alignment = PP_ALIGN.CENTER
        subtitle_para.font.size = Pt(28)
        subtitle_para.font.color.rgb = RGBColor(255, 255, 255)

        # Remove borders
        title_box.line.fill.background()
        subtitle_box.line.fill.background()

    @staticmethod
    def create_section_slide(prs: Presentation, section_title: str,
                            color_scheme: Tuple[int, int, int] = (79, 70, 229)) -> None:
        """Create section divider slide"""
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # Simple solid background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*color_scheme)

        # Section title
        title_box = slide.shapes.add_textbox(
            Inches(1), Inches(2),
            prs.slide_width - Inches(2), Inches(2)
        )
        title_frame = title_box.text_frame
        title_frame.text = section_title
        title_para = title_frame.paragraphs[0]
        title_para.alignment = PP_ALIGN.CENTER
        title_para.font.size = Pt(60)
        title_para.font.bold = True
        title_para.font.color.rgb = RGBColor(255, 255, 255)

        title_box.line.fill.background()

    @staticmethod
    def create_content_slide(prs: Presentation, title: str, content: List[str],
                            color_scheme: Tuple[int, int, int] = (79, 70, 229),
                            layout: str = 'single_column') -> None:
        """
        Create content slide with various layouts

        Args:
            prs: Presentation object
            title: Slide title
            content: List of content items
            color_scheme: RGB color tuple
            layout: 'single_column', 'two_column', or 'three_column'
        """
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # Title bar with accent color
        title_shape = slide.shapes.add_shape(
            1,  # Rectangle
            0, 0,
            prs.slide_width, Inches(1)
        )
        title_shape.fill.solid()
        title_shape.fill.fore_color.rgb = RGBColor(*color_scheme)
        title_shape.line.fill.background()

        # Title text
        title_frame = title_shape.text_frame
        title_frame.text = title
        title_para = title_frame.paragraphs[0]
        title_para.alignment = PP_ALIGN.LEFT
        title_para.font.size = Pt(36)
        title_para.font.bold = True
        title_para.font.color.rgb = RGBColor(255, 255, 255)
        title_frame.margin_left = Inches(0.5)
        title_frame.vertical_anchor = 1  # Middle

        # Content area
        content_top = Inches(1.2)
        content_height = prs.slide_height - content_top - Inches(0.5)

        if layout == 'single_column':
            ModernTemplates._add_content_column(
                slide, Inches(0.5), content_top,
                prs.slide_width - Inches(1), content_height,
                content
            )
        elif layout == 'two_column':
            col_width = (prs.slide_width - Inches(1.5)) / 2
            mid_point = len(content) // 2

            ModernTemplates._add_content_column(
                slide, Inches(0.5), content_top,
                col_width, content_height,
                content[:mid_point]
            )
            ModernTemplates._add_content_column(
                slide, Inches(0.5) + col_width + Inches(0.5), content_top,
                col_width, content_height,
                content[mid_point:]
            )
        elif layout == 'three_column':
            col_width = (prs.slide_width - Inches(2)) / 3
            third = len(content) // 3

            for i in range(3):
                left = Inches(0.5) + i * (col_width + Inches(0.5))
                start_idx = i * third
                end_idx = start_idx + third if i < 2 else len(content)

                ModernTemplates._add_content_column(
                    slide, left, content_top,
                    col_width, content_height,
                    content[start_idx:end_idx]
                )

    @staticmethod
    def _add_content_column(slide, left, top, width, height, items: List[str]):
        """Add content column to slide"""
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.word_wrap = True

        for i, item in enumerate(items):
            if i > 0:
                text_frame.add_paragraph()

            paragraph = text_frame.paragraphs[-1]
            paragraph.text = f"• {item}"
            paragraph.font.size = Pt(18)
            paragraph.space_after = Pt(12)

        textbox.line.fill.background()

    @staticmethod
    def create_image_slide(prs: Presentation, title: str, image_path: str,
                          caption: str = "",
                          color_scheme: Tuple[int, int, int] = (79, 70, 229)) -> None:
        """Create slide with image"""
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # Title
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.5),
            prs.slide_width - Inches(1), Inches(0.7)
        )
        title_frame = title_box.text_frame
        title_frame.text = title
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(36)
        title_para.font.bold = True
        title_para.font.color.rgb = RGBColor(*color_scheme)
        title_box.line.fill.background()

        # Image
        img_top = Inches(1.5)
        img_height = prs.slide_height - img_top - Inches(1.5 if caption else 0.5)

        # Load image to get aspect ratio
        try:
            img = Image.open(image_path)
            img_aspect = img.width / img.height
            slide_aspect = prs.slide_width.inches / img_height.inches

            if img_aspect > slide_aspect:
                # Image is wider
                img_width = prs.slide_width - Inches(1)
                actual_height = img_width.inches / img_aspect
                pic = slide.shapes.add_picture(
                    image_path,
                    Inches(0.5),
                    img_top,
                    img_width,
                    Inches(actual_height)
                )
            else:
                # Image is taller
                actual_width = img_height.inches * img_aspect
                pic = slide.shapes.add_picture(
                    image_path,
                    (prs.slide_width - Inches(actual_width)) / 2,
                    img_top,
                    Inches(actual_width),
                    img_height
                )

            # Caption
            if caption:
                caption_box = slide.shapes.add_textbox(
                    Inches(0.5),
                    prs.slide_height - Inches(1),
                    prs.slide_width - Inches(1),
                    Inches(0.5)
                )
                caption_frame = caption_box.text_frame
                caption_frame.text = caption
                caption_para = caption_frame.paragraphs[0]
                caption_para.alignment = PP_ALIGN.CENTER
                caption_para.font.size = Pt(16)
                caption_para.font.italic = True
                caption_box.line.fill.background()

        except Exception as e:
            print(f"Failed to add image: {e}")

    @staticmethod
    def create_quote_slide(prs: Presentation, quote: str, author: str,
                          color_scheme: Tuple[int, int, int] = (79, 70, 229)) -> None:
        """Create quote slide"""
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # Background
        from pptx_images import ImageProcessor
        gradient = ImageProcessor.create_gradient_image(
            int(prs.slide_width.inches * 150),
            int(prs.slide_height.inches * 150),
            tuple(c + 30 for c in color_scheme),
            tuple(max(0, c - 30) for c in color_scheme),
            angle=45
        )

        img_stream = BytesIO()
        gradient.save(img_stream, format='PNG')
        img_stream.seek(0)

        slide.shapes.add_picture(img_stream, 0, 0,
                                prs.slide_width, prs.slide_height)

        # Quote
        quote_box = slide.shapes.add_textbox(
            Inches(2), Inches(2),
            prs.slide_width - Inches(4), Inches(2.5)
        )
        quote_frame = quote_box.text_frame
        quote_frame.text = f'"{quote}"'
        quote_para = quote_frame.paragraphs[0]
        quote_para.alignment = PP_ALIGN.CENTER
        quote_para.font.size = Pt(32)
        quote_para.font.italic = True
        quote_para.font.color.rgb = RGBColor(255, 255, 255)
        quote_box.line.fill.background()

        # Author
        author_box = slide.shapes.add_textbox(
            Inches(2), Inches(4.5),
            prs.slide_width - Inches(4), Inches(0.5)
        )
        author_frame = author_box.text_frame
        author_frame.text = f"— {author}"
        author_para = author_frame.paragraphs[0]
        author_para.alignment = PP_ALIGN.RIGHT
        author_para.font.size = Pt(24)
        author_para.font.color.rgb = RGBColor(255, 255, 255)
        author_box.line.fill.background()


if __name__ == '__main__':
    # Demo template creation
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    color = (79, 70, 229)  # Modern purple

    ModernTemplates.create_title_slide(
        prs, "Amazing Presentation", "Created with Python", color
    )

    ModernTemplates.create_section_slide(prs, "Introduction", color)

    ModernTemplates.create_content_slide(
        prs, "Key Features",
        ["Beautiful gradients", "Professional layouts", "Easy to use", "Pure Python"],
        color, 'two_column'
    )

    ModernTemplates.create_quote_slide(
        prs, "Python is the best language for PowerPoint automation",
        "Happy Developer", color
    )

    prs.save('template_demo.pptx')
    print("Template demo created: template_demo.pptx")
