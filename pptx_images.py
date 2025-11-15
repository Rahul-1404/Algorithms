"""
Advanced Image Processing for PowerPoint
Handle images, icons, gradients, and visual effects
"""

from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageEnhance
from io import BytesIO
from typing import Tuple, Optional, List
import math


class ImageProcessor:
    """Process and optimize images for PowerPoint"""

    @staticmethod
    def create_gradient_image(width: int, height: int, color1: Tuple[int, int, int],
                             color2: Tuple[int, int, int], angle: int = 90,
                             gradient_type: str = 'linear') -> Image.Image:
        """
        Create a gradient image

        Args:
            width, height: Image dimensions in pixels
            color1, color2: RGB tuples for start and end colors
            angle: Gradient angle in degrees (0=horizontal, 90=vertical)
            gradient_type: 'linear' or 'radial'
        """
        img = Image.new('RGBA', (width, height))

        if gradient_type == 'radial':
            return ImageProcessor._create_radial_gradient(img, color1, color2)
        else:
            return ImageProcessor._create_linear_gradient(img, color1, color2, angle)

    @staticmethod
    def _create_linear_gradient(img: Image.Image, color1: Tuple[int, int, int],
                               color2: Tuple[int, int, int], angle: int) -> Image.Image:
        """Create linear gradient"""
        width, height = img.size
        draw = ImageDraw.Draw(img)

        # Normalize angle
        angle = angle % 360

        if angle in [0, 180]:
            # Horizontal gradient
            for x in range(width):
                ratio = x / width if angle == 0 else 1 - x / width
                r = int(color1[0] + (color2[0] - color1[0]) * ratio)
                g = int(color1[1] + (color2[1] - color1[1]) * ratio)
                b = int(color1[2] + (color2[2] - color1[2]) * ratio)
                draw.line([(x, 0), (x, height)], fill=(r, g, b, 255))
        elif angle in [90, 270]:
            # Vertical gradient
            for y in range(height):
                ratio = y / height if angle == 90 else 1 - y / height
                r = int(color1[0] + (color2[0] - color1[0]) * ratio)
                g = int(color1[1] + (color2[1] - color1[1]) * ratio)
                b = int(color1[2] + (color2[2] - color1[2]) * ratio)
                draw.line([(0, y), (width, y)], fill=(r, g, b, 255))
        else:
            # Diagonal gradient (approximate)
            max_distance = math.sqrt(width**2 + height**2)
            angle_rad = math.radians(angle)

            for y in range(height):
                for x in range(width):
                    # Calculate distance along gradient direction
                    distance = (x * math.cos(angle_rad) + y * math.sin(angle_rad))
                    ratio = min(1.0, max(0.0, distance / max_distance))

                    r = int(color1[0] + (color2[0] - color1[0]) * ratio)
                    g = int(color1[1] + (color2[1] - color1[1]) * ratio)
                    b = int(color1[2] + (color2[2] - color1[2]) * ratio)
                    draw.point((x, y), fill=(r, g, b, 255))

        return img

    @staticmethod
    def _create_radial_gradient(img: Image.Image, color1: Tuple[int, int, int],
                               color2: Tuple[int, int, int]) -> Image.Image:
        """Create radial gradient"""
        width, height = img.size
        draw = ImageDraw.Draw(img)

        center_x, center_y = width // 2, height // 2
        max_radius = max(width, height) // 2

        for y in range(height):
            for x in range(width):
                distance = math.sqrt((x - center_x)**2 + (y - center_y)**2)
                ratio = min(1.0, distance / max_radius)

                r = int(color1[0] + (color2[0] - color1[0]) * ratio)
                g = int(color1[1] + (color2[1] - color1[1]) * ratio)
                b = int(color1[2] + (color2[2] - color1[2]) * ratio)
                draw.point((x, y), fill=(r, g, b, 255))

        return img

    @staticmethod
    def create_icon(icon_type: str, size: int, color: Tuple[int, int, int],
                   background: Optional[Tuple[int, int, int]] = None) -> Image.Image:
        """
        Create simple vector-style icons

        Args:
            icon_type: Type of icon ('circle', 'square', 'triangle', 'star', 'checkmark', 'arrow')
            size: Icon size in pixels
            color: Icon color RGB tuple
            background: Background color (transparent if None)
        """
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0) if not background else (*background, 255))
        draw = ImageDraw.Draw(img)

        padding = size // 10
        inner_size = size - 2 * padding

        if icon_type == 'circle':
            draw.ellipse([padding, padding, size - padding, size - padding],
                        fill=(*color, 255), outline=(*color, 255), width=2)

        elif icon_type == 'square':
            draw.rectangle([padding, padding, size - padding, size - padding],
                          fill=(*color, 255), outline=(*color, 255), width=2)

        elif icon_type == 'triangle':
            points = [
                (size // 2, padding),
                (padding, size - padding),
                (size - padding, size - padding)
            ]
            draw.polygon(points, fill=(*color, 255), outline=(*color, 255))

        elif icon_type == 'star':
            points = ImageProcessor._get_star_points(size // 2, size // 2, size // 2 - padding, 5)
            draw.polygon(points, fill=(*color, 255), outline=(*color, 255))

        elif icon_type == 'checkmark':
            # Draw checkmark
            width = max(2, size // 15)
            points = [
                (padding + inner_size * 0.2, size // 2),
                (padding + inner_size * 0.4, size - padding - inner_size * 0.2),
                (size - padding, padding + inner_size * 0.1)
            ]
            for i in range(len(points) - 1):
                draw.line([points[i], points[i + 1]], fill=(*color, 255), width=width)

        elif icon_type == 'arrow':
            # Right arrow
            width = max(2, size // 15)
            # Arrow shaft
            draw.line([(padding, size // 2), (size - padding - inner_size * 0.3, size // 2)],
                     fill=(*color, 255), width=width)
            # Arrow head
            points = [
                (size - padding - inner_size * 0.3, padding + inner_size * 0.2),
                (size - padding, size // 2),
                (size - padding - inner_size * 0.3, size - padding - inner_size * 0.2)
            ]
            draw.polygon(points, fill=(*color, 255))

        return img

    @staticmethod
    def _get_star_points(cx: int, cy: int, radius: int, points: int = 5) -> List[Tuple[int, int]]:
        """Calculate star polygon points"""
        star_points = []
        angle = math.pi / 2  # Start from top
        angle_step = 2 * math.pi / points

        for i in range(points * 2):
            r = radius if i % 2 == 0 else radius * 0.4
            x = cx + r * math.cos(angle)
            y = cy - r * math.sin(angle)
            star_points.append((int(x), int(y)))
            angle += angle_step / 2

        return star_points

    @staticmethod
    def apply_shadow(img: Image.Image, offset: Tuple[int, int] = (5, 5),
                    blur_radius: int = 10, shadow_color: Tuple[int, int, int] = (0, 0, 0),
                    opacity: int = 128) -> Image.Image:
        """
        Apply drop shadow to image

        Args:
            img: Input image
            offset: Shadow offset (x, y)
            blur_radius: Shadow blur radius
            shadow_color: Shadow color RGB
            opacity: Shadow opacity (0-255)
        """
        # Create shadow layer
        shadow = Image.new('RGBA', (img.width + abs(offset[0]) + blur_radius * 2,
                                   img.height + abs(offset[1]) + blur_radius * 2),
                          (0, 0, 0, 0))

        # Create shadow shape
        shadow_img = Image.new('RGBA', img.size, (*shadow_color, opacity))
        shadow_img.putalpha(img.split()[3] if img.mode == 'RGBA' else Image.new('L', img.size, 255))

        # Blur shadow
        shadow_img = shadow_img.filter(ImageFilter.GaussianBlur(blur_radius))

        # Paste shadow
        shadow_x = blur_radius + max(0, offset[0])
        shadow_y = blur_radius + max(0, offset[1])
        shadow.paste(shadow_img, (shadow_x, shadow_y), shadow_img)

        # Paste original image
        img_x = blur_radius + max(0, -offset[0])
        img_y = blur_radius + max(0, -offset[1])
        shadow.paste(img, (img_x, img_y), img if img.mode == 'RGBA' else None)

        return shadow

    @staticmethod
    def create_rounded_rectangle(width: int, height: int, radius: int,
                                 fill_color: Tuple[int, int, int],
                                 border_color: Optional[Tuple[int, int, int]] = None,
                                 border_width: int = 0) -> Image.Image:
        """Create rounded rectangle"""
        img = Image.new('RGBA', (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Draw rounded rectangle
        draw.rounded_rectangle([0, 0, width, height], radius=radius,
                              fill=(*fill_color, 255),
                              outline=(*border_color, 255) if border_color else None,
                              width=border_width)

        return img

    @staticmethod
    def optimize_image(img: Image.Image, max_width: int = 1920, max_height: int = 1080,
                      quality: int = 85) -> BytesIO:
        """
        Optimize image for PowerPoint

        Args:
            img: Input PIL Image
            max_width, max_height: Maximum dimensions
            quality: JPEG quality (1-100)
        """
        # Resize if needed
        if img.width > max_width or img.height > max_height:
            img.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)

        # Convert to RGB if needed for JPEG
        if img.mode in ('RGBA', 'P'):
            # Keep transparency for PNG
            output = BytesIO()
            img.save(output, format='PNG', optimize=True)
        else:
            if img.mode != 'RGB':
                img = img.convert('RGB')
            output = BytesIO()
            img.save(output, format='JPEG', quality=quality, optimize=True)

        output.seek(0)
        return output

    @staticmethod
    def create_text_image(text: str, width: int, height: int,
                         font_size: int = 24, font_color: Tuple[int, int, int] = (0, 0, 0),
                         background: Tuple[int, int, int] = (255, 255, 255),
                         align: str = 'center') -> Image.Image:
        """
        Create image with text

        Args:
            text: Text to render
            width, height: Image dimensions
            font_size: Font size in points
            font_color: Text color RGB
            background: Background color RGB
            align: Text alignment ('left', 'center', 'right')
        """
        img = Image.new('RGB', (width, height), background)
        draw = ImageDraw.Draw(img)

        # Try to use a nice font, fallback to default
        try:
            font = ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', font_size)
        except:
            font = ImageFont.load_default()

        # Calculate text position
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]

        if align == 'center':
            x = (width - text_width) // 2
        elif align == 'right':
            x = width - text_width - 10
        else:
            x = 10

        y = (height - text_height) // 2

        draw.text((x, y), text, font=font, fill=font_color)

        return img

    @staticmethod
    def create_multi_color_gradient(width: int, height: int,
                                   colors: List[Tuple[int, int, int]],
                                   angle: int = 90) -> Image.Image:
        """
        Create multi-color gradient

        Args:
            width, height: Image dimensions
            colors: List of RGB color tuples
            angle: Gradient angle
        """
        if len(colors) < 2:
            colors = [(255, 255, 255), (0, 0, 0)]

        img = Image.new('RGBA', (width, height))
        draw = ImageDraw.Draw(img)

        # Vertical gradient
        if angle in [90, 270]:
            segment_height = height / (len(colors) - 1)

            for y in range(height):
                # Determine which color segment
                segment = min(len(colors) - 2, int(y / segment_height))
                local_ratio = (y - segment * segment_height) / segment_height

                color1 = colors[segment]
                color2 = colors[segment + 1]

                r = int(color1[0] + (color2[0] - color1[0]) * local_ratio)
                g = int(color1[1] + (color2[1] - color1[1]) * local_ratio)
                b = int(color1[2] + (color2[2] - color1[2]) * local_ratio)

                draw.line([(0, y), (width, y)], fill=(r, g, b, 255))
        else:
            # Horizontal gradient
            segment_width = width / (len(colors) - 1)

            for x in range(width):
                segment = min(len(colors) - 2, int(x / segment_width))
                local_ratio = (x - segment * segment_width) / segment_width

                color1 = colors[segment]
                color2 = colors[segment + 1]

                r = int(color1[0] + (color2[0] - color1[0]) * local_ratio)
                g = int(color1[1] + (color2[1] - color1[1]) * local_ratio)
                b = int(color1[2] + (color2[2] - color1[2]) * local_ratio)

                draw.line([(x, 0), (x, height)], fill=(r, g, b, 255))

        return img


if __name__ == '__main__':
    # Demo various image processing capabilities
    processor = ImageProcessor()

    # Create gradient
    gradient = processor.create_gradient_image(800, 600, (102, 126, 234), (118, 75, 162), 45)
    gradient.save('gradient_demo.png')

    # Create icons
    icons = ['circle', 'square', 'triangle', 'star', 'checkmark', 'arrow']
    for icon_type in icons:
        icon = processor.create_icon(icon_type, 200, (31, 119, 180))
        icon.save(f'icon_{icon_type}_demo.png')

    # Multi-color gradient
    multi_grad = processor.create_multi_color_gradient(
        800, 600,
        [(102, 126, 234), (118, 75, 162), (237, 100, 166), (255, 198, 93)]
    )
    multi_grad.save('multi_gradient_demo.png')

    print("Image processing demos created!")
