"""
Python HTML to PowerPoint Converter
A pure Python replacement for the Node.js html2pptx with advanced styling support
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from bs4 import BeautifulSoup
import re
from typing import Dict, List, Tuple, Optional
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import cssselect
import tinycss2


class CSSParser:
    """Parse and apply CSS styles"""

    @staticmethod
    def parse_style(style_str: str) -> Dict[str, str]:
        """Parse inline CSS style string"""
        if not style_str:
            return {}

        styles = {}
        declarations = tinycss2.parse_declaration_list(style_str)

        for declaration in declarations:
            if hasattr(declaration, 'name') and hasattr(declaration, 'value'):
                name = declaration.name
                value = ''.join(token.serialize() for token in declaration.value).strip()
                styles[name] = value

        return styles

    @staticmethod
    def parse_color(color_str: str) -> Optional[RGBColor]:
        """Parse CSS color to RGBColor"""
        if not color_str:
            return None

        color_str = color_str.strip().lower()

        # Hex colors
        if color_str.startswith('#'):
            hex_color = color_str[1:]
            if len(hex_color) == 3:
                hex_color = ''.join([c*2 for c in hex_color])
            if len(hex_color) == 6:
                r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
                return RGBColor(r, g, b)

        # RGB/RGBA
        rgb_match = re.match(r'rgba?\((\d+),\s*(\d+),\s*(\d+)', color_str)
        if rgb_match:
            r, g, b = map(int, rgb_match.groups())
            return RGBColor(r, g, b)

        # Named colors
        color_map = {
            'white': (255, 255, 255), 'black': (0, 0, 0), 'red': (255, 0, 0),
            'green': (0, 128, 0), 'blue': (0, 0, 255), 'yellow': (255, 255, 0),
            'cyan': (0, 255, 255), 'magenta': (255, 0, 255), 'gray': (128, 128, 128),
            'grey': (128, 128, 128), 'orange': (255, 165, 0), 'purple': (128, 0, 128),
            'navy': (0, 0, 128), 'teal': (0, 128, 128), 'lime': (0, 255, 0),
        }

        if color_str in color_map:
            r, g, b = color_map[color_str]
            return RGBColor(r, g, b)

        return None

    @staticmethod
    def parse_dimension(dim_str: str, reference_size: float = 720.0) -> float:
        """Parse CSS dimension to points"""
        if not dim_str:
            return 0.0

        dim_str = dim_str.strip().lower()

        # Points
        if dim_str.endswith('pt'):
            return float(dim_str[:-2])

        # Pixels (assume 96 DPI)
        if dim_str.endswith('px'):
            return float(dim_str[:-2]) * 0.75

        # Inches
        if dim_str.endswith('in'):
            return float(dim_str[:-2]) * 72

        # Percentage
        if dim_str.endswith('%'):
            return (float(dim_str[:-1]) / 100.0) * reference_size

        # Try direct number
        try:
            return float(dim_str)
        except ValueError:
            return 0.0


class GradientRenderer:
    """Render CSS gradients as images"""

    @staticmethod
    def render_gradient(gradient_str: str, width: int, height: int) -> Image.Image:
        """Render a CSS gradient to a PIL Image"""
        gradient_str = gradient_str.strip()

        # Linear gradient
        if gradient_str.startswith('linear-gradient'):
            return GradientRenderer._render_linear_gradient(gradient_str, width, height)

        # Radial gradient
        if gradient_str.startswith('radial-gradient'):
            return GradientRenderer._render_radial_gradient(gradient_str, width, height)

        # Default solid color
        return Image.new('RGBA', (width, height), (255, 255, 255, 255))

    @staticmethod
    def _parse_gradient_colors(gradient_str: str) -> List[Tuple[int, int, int, float]]:
        """Parse gradient color stops"""
        # Extract content between parentheses
        match = re.search(r'\((.*)\)', gradient_str)
        if not match:
            return [(255, 255, 255, 0.0), (0, 0, 0, 1.0)]

        content = match.group(1)

        # Remove angle/position if present
        parts = re.split(r',(?![^(]*\))', content)

        # Filter out direction/angle
        color_stops = []
        for part in parts:
            part = part.strip()
            # Skip if it's a direction or angle
            if any(keyword in part for keyword in ['deg', 'to top', 'to bottom', 'to left', 'to right']):
                continue

            # Parse color
            rgb_match = re.match(r'rgba?\((\d+),\s*(\d+),\s*(\d+)', part)
            hex_match = re.match(r'#([0-9a-fA-F]{6})', part)

            if rgb_match:
                r, g, b = map(int, rgb_match.groups())
                # Extract position if present
                pos_match = re.search(r'(\d+)%', part)
                position = float(pos_match.group(1)) / 100.0 if pos_match else len(color_stops) / max(1, len(parts) - 1)
                color_stops.append((r, g, b, position))
            elif hex_match:
                hex_color = hex_match.group(1)
                r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
                pos_match = re.search(r'(\d+)%', part)
                position = float(pos_match.group(1)) / 100.0 if pos_match else len(color_stops) / max(1, len(parts) - 1)
                color_stops.append((r, g, b, position))

        if not color_stops:
            return [(255, 255, 255, 0.0), (0, 0, 0, 1.0)]

        return color_stops

    @staticmethod
    def _render_linear_gradient(gradient_str: str, width: int, height: int) -> Image.Image:
        """Render linear gradient"""
        img = Image.new('RGBA', (width, height))
        draw = ImageDraw.Draw(img)

        color_stops = GradientRenderer._parse_gradient_colors(gradient_str)

        # Default to vertical gradient
        angle = 180  # Bottom to top

        # Parse angle
        angle_match = re.search(r'(\d+)deg', gradient_str)
        if angle_match:
            angle = int(angle_match.group(1))
        elif 'to bottom' in gradient_str:
            angle = 180
        elif 'to top' in gradient_str:
            angle = 0
        elif 'to right' in gradient_str:
            angle = 90
        elif 'to left' in gradient_str:
            angle = 270

        # Simple vertical or horizontal gradient
        if angle in [0, 180]:
            for y in range(height):
                position = y / height
                color = GradientRenderer._interpolate_color(color_stops, position)
                draw.line([(0, y), (width, y)], fill=color)
        elif angle in [90, 270]:
            for x in range(width):
                position = x / width
                color = GradientRenderer._interpolate_color(color_stops, position)
                draw.line([(x, 0), (x, height)], fill=color)
        else:
            # Approximate diagonal gradients
            for y in range(height):
                position = y / height
                color = GradientRenderer._interpolate_color(color_stops, position)
                draw.line([(0, y), (width, y)], fill=color)

        return img

    @staticmethod
    def _render_radial_gradient(gradient_str: str, width: int, height: int) -> Image.Image:
        """Render radial gradient"""
        img = Image.new('RGBA', (width, height))
        draw = ImageDraw.Draw(img)

        color_stops = GradientRenderer._parse_gradient_colors(gradient_str)

        center_x, center_y = width // 2, height // 2
        max_radius = max(width, height)

        for radius in range(max_radius, 0, -1):
            position = 1 - (radius / max_radius)
            color = GradientRenderer._interpolate_color(color_stops, position)
            draw.ellipse([
                center_x - radius, center_y - radius,
                center_x + radius, center_y + radius
            ], fill=color)

        return img

    @staticmethod
    def _interpolate_color(color_stops: List[Tuple[int, int, int, float]], position: float) -> Tuple[int, int, int, int]:
        """Interpolate color at given position"""
        if not color_stops:
            return (255, 255, 255, 255)

        # Sort by position
        sorted_stops = sorted(color_stops, key=lambda x: x[3])

        # Before first stop
        if position <= sorted_stops[0][3]:
            r, g, b, _ = sorted_stops[0]
            return (r, g, b, 255)

        # After last stop
        if position >= sorted_stops[-1][3]:
            r, g, b, _ = sorted_stops[-1]
            return (r, g, b, 255)

        # Between stops
        for i in range(len(sorted_stops) - 1):
            stop1 = sorted_stops[i]
            stop2 = sorted_stops[i + 1]

            if stop1[3] <= position <= stop2[3]:
                # Linear interpolation
                t = (position - stop1[3]) / (stop2[3] - stop1[3])
                r = int(stop1[0] + t * (stop2[0] - stop1[0]))
                g = int(stop1[1] + t * (stop2[1] - stop1[1]))
                b = int(stop1[2] + t * (stop2[2] - stop1[2]))
                return (r, g, b, 255)

        return (255, 255, 255, 255)


class HTML2PPTX:
    """Convert HTML to PowerPoint with advanced styling"""

    # Web-safe fonts mapping
    FONT_MAP = {
        'arial': 'Arial',
        'helvetica': 'Helvetica',
        'times': 'Times New Roman',
        'times new roman': 'Times New Roman',
        'georgia': 'Georgia',
        'courier': 'Courier New',
        'courier new': 'Courier New',
        'verdana': 'Verdana',
        'tahoma': 'Tahoma',
        'trebuchet': 'Trebuchet MS',
        'trebuchet ms': 'Trebuchet MS',
        'impact': 'Impact',
        'comic sans': 'Comic Sans MS',
        'comic sans ms': 'Comic Sans MS',
    }

    def __init__(self, aspect_ratio: str = '16:9'):
        """Initialize converter with aspect ratio"""
        self.prs = Presentation()

        # Set slide dimensions based on aspect ratio
        dimensions = {
            '16:9': (Inches(10), Inches(5.625)),
            '4:3': (Inches(10), Inches(7.5)),
            '16:10': (Inches(10), Inches(6.25)),
        }

        width, height = dimensions.get(aspect_ratio, dimensions['16:9'])
        self.prs.slide_width = width
        self.prs.slide_height = height

        self.slide_width_pt = width.pt
        self.slide_height_pt = height.pt

        self.css_parser = CSSParser()
        self.gradient_renderer = GradientRenderer()

    def add_slide_from_html(self, html: str) -> None:
        """Add a slide from HTML content"""
        soup = BeautifulSoup(html, 'html.parser')

        # Create blank slide
        blank_layout = self.prs.slide_layouts[6]  # Blank layout
        slide = self.prs.slides.add_slide(blank_layout)

        # Find the main container (should have width/height matching slide)
        container = soup.find('div')
        if not container:
            return

        # Parse container styles
        container_styles = self.css_parser.parse_style(container.get('style', ''))

        # Add background if present
        self._add_background(slide, container_styles)

        # Process child elements
        self._process_elements(slide, container, 0, 0)

    def _add_background(self, slide, styles: Dict[str, str]) -> None:
        """Add background to slide"""
        background = styles.get('background') or styles.get('background-color')

        if not background:
            return

        # Check for gradient
        if 'gradient' in background:
            # Render gradient as image
            width_px = int(self.slide_width_pt * 1.33)  # Convert pt to px
            height_px = int(self.slide_height_pt * 1.33)

            gradient_img = self.gradient_renderer.render_gradient(background, width_px, height_px)

            # Save to BytesIO
            img_stream = BytesIO()
            gradient_img.save(img_stream, format='PNG')
            img_stream.seek(0)

            # Add as background
            left = top = Inches(0)
            pic = slide.shapes.add_picture(img_stream, left, top,
                                          width=self.prs.slide_width,
                                          height=self.prs.slide_height)

            # Send to back
            slide.shapes._spTree.remove(pic._element)
            slide.shapes._spTree.insert(2, pic._element)
        else:
            # Solid color
            color = self.css_parser.parse_color(background)
            if color:
                fill = slide.background.fill
                fill.solid()
                fill.fore_color.rgb = color

    def _process_elements(self, slide, element, offset_x: float, offset_y: float) -> None:
        """Process HTML elements and add to slide"""
        for child in element.children:
            if isinstance(child, str):
                continue

            if child.name in ['div']:
                self._process_div(slide, child, offset_x, offset_y)
            elif child.name in ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                self._process_text(slide, child, offset_x, offset_y)
            elif child.name in ['ul', 'ol']:
                self._process_list(slide, child, offset_x, offset_y)
            elif child.name == 'img':
                self._process_image(slide, child, offset_x, offset_y)

    def _process_div(self, slide, div, parent_offset_x: float, parent_offset_y: float) -> None:
        """Process div element with positioning"""
        styles = self.css_parser.parse_style(div.get('style', ''))

        # Get position
        position = styles.get('position', 'static')
        left = self.css_parser.parse_dimension(styles.get('left', '0'), self.slide_width_pt)
        top = self.css_parser.parse_dimension(styles.get('top', '0'), self.slide_height_pt)
        width = self.css_parser.parse_dimension(styles.get('width', '100pt'), self.slide_width_pt)
        height = self.css_parser.parse_dimension(styles.get('height', '100pt'), self.slide_height_pt)

        offset_x = parent_offset_x + left
        offset_y = parent_offset_y + top

        # Add shape for div if it has background/border
        if any(key in styles for key in ['background', 'background-color', 'border', 'box-shadow']):
            shape = slide.shapes.add_shape(
                1,  # Rectangle
                Pt(offset_x), Pt(offset_y),
                Pt(width), Pt(height)
            )

            # Apply background
            background = styles.get('background') or styles.get('background-color')
            if background and 'gradient' in background:
                # Render gradient
                width_px = int(width * 1.33)
                height_px = int(height * 1.33)
                gradient_img = self.gradient_renderer.render_gradient(background, width_px, height_px)

                img_stream = BytesIO()
                gradient_img.save(img_stream, format='PNG')
                img_stream.seek(0)

                # Remove shape and add picture instead
                slide.shapes._spTree.remove(shape._element)
                slide.shapes.add_picture(img_stream, Pt(offset_x), Pt(offset_y), Pt(width), Pt(height))
            elif background:
                color = self.css_parser.parse_color(background)
                if color:
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = color
            else:
                shape.fill.background()

            # Apply border
            border = styles.get('border')
            if border:
                self._apply_border(shape, border)

            # Line should be invisible if no border
            if not border:
                shape.line.fill.background()

        # Process child elements
        self._process_elements(slide, div, offset_x, offset_y)

    def _process_text(self, slide, element, offset_x: float, offset_y: float) -> None:
        """Process text elements (p, h1-h6)"""
        styles = self.css_parser.parse_style(element.get('style', ''))

        # Get dimensions
        left = offset_x + self.css_parser.parse_dimension(styles.get('left', '0'), self.slide_width_pt)
        top = offset_y + self.css_parser.parse_dimension(styles.get('top', '0'), self.slide_height_pt)
        width = self.css_parser.parse_dimension(styles.get('width', '200pt'), self.slide_width_pt)
        height = self.css_parser.parse_dimension(styles.get('height', '50pt'), self.slide_height_pt)

        # Create text box
        textbox = slide.shapes.add_textbox(Pt(left), Pt(top), Pt(width), Pt(height))
        text_frame = textbox.text_frame
        text_frame.word_wrap = True

        # Clear default paragraph
        text_frame.clear()

        # Extract text content with formatting
        self._add_formatted_text(text_frame, element, styles)

        # Remove text box border
        textbox.line.fill.background()

    def _add_formatted_text(self, text_frame, element, parent_styles: Dict[str, str]) -> None:
        """Add formatted text to text frame"""
        paragraph = text_frame.paragraphs[0] if text_frame.paragraphs else text_frame.add_paragraph()

        # Apply paragraph-level styles
        text_align = parent_styles.get('text-align', 'left')
        if text_align == 'center':
            paragraph.alignment = PP_ALIGN.CENTER
        elif text_align == 'right':
            paragraph.alignment = PP_ALIGN.RIGHT
        elif text_align == 'justify':
            paragraph.alignment = PP_ALIGN.JUSTIFY
        else:
            paragraph.alignment = PP_ALIGN.LEFT

        # Process inline content
        self._process_inline_content(paragraph, element, parent_styles)

    def _process_inline_content(self, paragraph, element, parent_styles: Dict[str, str]) -> None:
        """Process inline content with formatting"""
        for child in element.descendants:
            if isinstance(child, str):
                text = child.strip()
                if not text:
                    continue

                run = paragraph.add_run()
                run.text = text

                # Apply parent styles
                self._apply_text_styles(run, parent_styles)
            elif child.name in ['b', 'strong']:
                text = child.get_text()
                run = paragraph.add_run()
                run.text = text
                run.font.bold = True
                self._apply_text_styles(run, parent_styles)
            elif child.name in ['i', 'em']:
                text = child.get_text()
                run = paragraph.add_run()
                run.text = text
                run.font.italic = True
                self._apply_text_styles(run, parent_styles)
            elif child.name == 'u':
                text = child.get_text()
                run = paragraph.add_run()
                run.text = text
                run.font.underline = True
                self._apply_text_styles(run, parent_styles)
            elif child.name == 'span':
                span_styles = self.css_parser.parse_style(child.get('style', ''))
                merged_styles = {**parent_styles, **span_styles}
                text = child.get_text()
                run = paragraph.add_run()
                run.text = text
                self._apply_text_styles(run, merged_styles)
            elif child.name == 'br':
                paragraph.add_run().text = '\n'

    def _apply_text_styles(self, run, styles: Dict[str, str]) -> None:
        """Apply text styles to run"""
        # Font family
        font_family = styles.get('font-family', '').lower()
        for font_key, font_name in self.FONT_MAP.items():
            if font_key in font_family:
                run.font.name = font_name
                break
        else:
            run.font.name = 'Arial'  # Default

        # Font size
        font_size = styles.get('font-size', '')
        if font_size:
            size_pt = self.css_parser.parse_dimension(font_size)
            if size_pt > 0:
                run.font.size = Pt(size_pt)

        # Font color
        color = styles.get('color', '')
        if color:
            rgb_color = self.css_parser.parse_color(color)
            if rgb_color:
                run.font.color.rgb = rgb_color

        # Font weight
        font_weight = styles.get('font-weight', '')
        if font_weight in ['bold', '700', '800', '900']:
            run.font.bold = True

        # Font style
        font_style = styles.get('font-style', '')
        if font_style == 'italic':
            run.font.italic = True

    def _process_list(self, slide, list_element, offset_x: float, offset_y: float) -> None:
        """Process list elements"""
        styles = self.css_parser.parse_style(list_element.get('style', ''))

        left = offset_x + self.css_parser.parse_dimension(styles.get('left', '0'), self.slide_width_pt)
        top = offset_y + self.css_parser.parse_dimension(styles.get('top', '0'), self.slide_height_pt)
        width = self.css_parser.parse_dimension(styles.get('width', '300pt'), self.slide_width_pt)
        height = self.css_parser.parse_dimension(styles.get('height', '200pt'), self.slide_height_pt)

        textbox = slide.shapes.add_textbox(Pt(left), Pt(top), Pt(width), Pt(height))
        text_frame = textbox.text_frame
        text_frame.clear()

        # Process list items
        for li in list_element.find_all('li', recursive=False):
            paragraph = text_frame.add_paragraph()
            paragraph.text = li.get_text()
            paragraph.level = 0

            # Apply list styles
            self._apply_text_styles(paragraph.runs[0] if paragraph.runs else paragraph.add_run(), styles)

        textbox.line.fill.background()

    def _process_image(self, slide, img_element, offset_x: float, offset_y: float) -> None:
        """Process image elements"""
        styles = self.css_parser.parse_style(img_element.get('style', ''))
        src = img_element.get('src', '')

        if not src:
            return

        left = offset_x + self.css_parser.parse_dimension(styles.get('left', '0'), self.slide_width_pt)
        top = offset_y + self.css_parser.parse_dimension(styles.get('top', '0'), self.slide_height_pt)
        width = self.css_parser.parse_dimension(styles.get('width', '100pt'), self.slide_width_pt)
        height = self.css_parser.parse_dimension(styles.get('height', '100pt'), self.slide_height_pt)

        # Add image (src should be file path or URL)
        try:
            slide.shapes.add_picture(src, Pt(left), Pt(top), Pt(width), Pt(height))
        except Exception as e:
            print(f"Failed to add image {src}: {e}")

    def _apply_border(self, shape, border_style: str) -> None:
        """Apply border to shape"""
        # Parse border: "1px solid #000"
        parts = border_style.split()
        if len(parts) < 3:
            return

        width_pt = self.css_parser.parse_dimension(parts[0])
        color = self.css_parser.parse_color(parts[2])

        if width_pt > 0:
            shape.line.width = Pt(width_pt)

        if color:
            shape.line.color.rgb = color

    def save(self, filename: str) -> None:
        """Save presentation to file"""
        self.prs.save(filename)


def create_presentation_from_html(html_slides: List[str], output_file: str, aspect_ratio: str = '16:9') -> None:
    """
    Create a PowerPoint presentation from HTML slides

    Args:
        html_slides: List of HTML strings, one per slide
        output_file: Output .pptx file path
        aspect_ratio: Slide aspect ratio ('16:9', '4:3', or '16:10')
    """
    converter = HTML2PPTX(aspect_ratio=aspect_ratio)

    for html in html_slides:
        converter.add_slide_from_html(html)

    converter.save(output_file)


if __name__ == '__main__':
    # Example usage
    sample_html = """
    <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
        <h1 style="position: absolute; left: 50pt; top: 150pt; width: 620pt; color: white;
                   font-size: 48pt; font-family: Arial; text-align: center;">
            Welcome to Python PowerPoint
        </h1>
        <p style="position: absolute; left: 50pt; top: 250pt; width: 620pt; color: white;
                  font-size: 24pt; font-family: Arial; text-align: center;">
            Creating beautiful presentations without Node.js
        </p>
    </div>
    """

    create_presentation_from_html([sample_html], 'demo.pptx', '16:9')
    print("Demo presentation created: demo.pptx")
