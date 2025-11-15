# Python PowerPoint Creation Suite

**Professional PowerPoint presentations without Node.js dependencies**

A comprehensive, pure-Python solution for creating visually stunning PowerPoint presentations that rivals Node.js-based tools. Features advanced HTML-to-PowerPoint conversion, beautiful chart generation, gradient rendering, and modern template designs.

## 🌟 Key Features

- **HTML to PowerPoint Conversion** - Convert HTML/CSS to PowerPoint with precise positioning and styling
- **Beautiful Gradient Support** - Render CSS gradients (linear, radial, multi-color) as high-quality images
- **Professional Charts** - Matplotlib-powered charts with 6 gorgeous color schemes
- **Modern Templates** - Pre-built professional slide layouts
- **Advanced Image Processing** - Create icons, gradients, and visual effects
- **Template Management** - Work with existing templates, extract content, and manage slides
- **Validation & Utilities** - Comprehensive validation, thumbnail generation, and helper functions

## 🚀 Quick Start

### Installation

```bash
pip install -r requirements.txt
```

### Basic Usage

```python
from html2pptx import create_presentation_from_html

slides = [
    """
    <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
        <h1 style="position: absolute; left: 50pt; top: 150pt; width: 620pt; color: white;
                   font-size: 54pt; font-family: Arial; text-align: center;">
            Beautiful Presentation
        </h1>
    </div>
    """
]

create_presentation_from_html(slides, 'output.pptx', '16:9')
```

## 📚 Module Overview

### 1. `html2pptx.py` - HTML to PowerPoint Converter

Convert HTML with CSS styling to PowerPoint slides.

**Features:**
- CSS parsing and application
- Precise positioning with absolute/relative layouts
- Linear and radial gradient rendering
- Web-safe font support
- Border, background, and text styling
- Lists, images, and complex layouts

**Example:**

```python
from html2pptx import HTML2PPTX

converter = HTML2PPTX(aspect_ratio='16:9')
converter.add_slide_from_html(html_content)
converter.save('presentation.pptx')
```

**Supported CSS Properties:**
- `position`, `left`, `top`, `width`, `height`
- `background`, `background-color` (including gradients)
- `color`, `font-family`, `font-size`, `font-weight`, `font-style`
- `text-align`, `border`
- `linear-gradient`, `radial-gradient`

### 2. `pptx_charts.py` - Professional Chart Generation

Create stunning charts with matplotlib integration.

**Chart Types:**
- Bar charts (clustered and stacked)
- Line charts (with/without markers)
- Pie charts
- Area charts
- Scatter plots

**Color Schemes:**
- `professional` - Classic business colors
- `vibrant` - Bold, energetic colors
- `pastel` - Soft, gentle colors
- `corporate` - Conservative blues and grays
- `modern` - Contemporary gradient-ready colors

**Example:**

```python
from pptx_charts import ChartGenerator
from pptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

ChartGenerator.add_matplotlib_chart(
    slide, 1, 1, 8, 4,
    ChartGenerator.create_beautiful_bar_chart,
    categories=['Q1', 'Q2', 'Q3', 'Q4'],
    data_series=[
        {'name': 'Sales', 'values': [100, 120, 140, 160]},
        {'name': 'Revenue', 'values': [90, 110, 130, 150]}
    ],
    title='Quarterly Performance',
    color_scheme='modern'
)

prs.save('charts.pptx')
```

### 3. `pptx_images.py` - Advanced Image Processing

Create and manipulate images for PowerPoint.

**Capabilities:**
- Gradient generation (linear, radial, multi-color)
- Icon creation (circle, square, triangle, star, checkmark, arrow)
- Drop shadows
- Rounded rectangles
- Text rendering
- Image optimization

**Example:**

```python
from pptx_images import ImageProcessor

processor = ImageProcessor()

# Create gradient
gradient = processor.create_gradient_image(800, 600, (102, 126, 234), (118, 75, 162), 45)
gradient.save('gradient.png')

# Create icon
icon = processor.create_icon('star', 200, (79, 70, 229))
icon.save('star_icon.png')

# Multi-color gradient
multi_grad = processor.create_multi_color_gradient(
    800, 600,
    [(102, 126, 234), (118, 75, 162), (237, 100, 166), (255, 198, 93)]
)
multi_grad.save('multi_gradient.png')
```

### 4. `pptx_templates.py` - Modern Template Library

Pre-built professional slide templates.

**Template Types:**
- Title slides with gradients
- Section divider slides
- Content slides (1, 2, or 3 column layouts)
- Image slides with captions
- Quote slides
- Custom layouts

**Example:**

```python
from pptx_templates import ModernTemplates
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)

color = (79, 70, 229)  # Modern indigo

ModernTemplates.create_title_slide(
    prs, "My Presentation", "Subtitle Here", color
)

ModernTemplates.create_content_slide(
    prs, "Key Points",
    ["Point 1", "Point 2", "Point 3", "Point 4"],
    color, 'two_column'
)

ModernTemplates.create_quote_slide(
    prs, "Inspiring quote here", "Author Name", color
)

prs.save('modern_presentation.pptx')
```

**Template Management:**

```python
from pptx_templates import TemplateManager

# Work with existing template
manager = TemplateManager('template.pptx')

# Get inventory of all slides
inventory = manager.get_slide_inventory()
manager.save_inventory('inventory.json')

# Replace text across slides
manager.replace_text({
    '{{COMPANY}}': 'Acme Corp',
    '{{DATE}}': '2024-01-15'
})

# Duplicate slides
new_slide_idx = manager.duplicate_slide(0)

manager.save('updated_presentation.pptx')
```

### 5. `pptx_utils.py` - Utilities and Validation

Helper functions for PowerPoint operations.

**Features:**
- Presentation validation
- File size checking
- Thumbnail generation
- Slide grid creation
- Text extraction
- Statistics gathering
- Slide numbering
- Presentation merging

**Example:**

```python
from pptx_utils import Validator, ThumbnailGenerator, PPTXHelper

# Validate presentation
validation = Validator.validate_presentation('presentation.pptx')
print(f"Valid: {validation['valid']}")
print(f"Errors: {validation['errors']}")
print(f"Warnings: {validation['warnings']}")

# Check file size
size_info = Validator.check_file_size('presentation.pptx', max_size_mb=50)
print(f"File size: {size_info['file_size_mb']} MB")

# Generate thumbnails
thumbnails = ThumbnailGenerator.generate_slide_thumbnails(
    'presentation.pptx', 'output_dir'
)

# Create slide grid
grid_path = ThumbnailGenerator.create_slide_grid(
    'presentation.pptx', 'grid.png', columns=3
)

# Get statistics
stats = PPTXHelper.get_presentation_stats('presentation.pptx')
print(stats)

# Extract all text
text_data = PPTXHelper.extract_text('presentation.pptx')

# Add slide numbers
PPTXHelper.add_slide_numbers('input.pptx', 'output.pptx', start_number=1)
```

## 🎨 Complete Examples

Run the comprehensive examples:

```bash
python examples.py
```

This generates 5 example presentations demonstrating all features:

1. **example_1_html.pptx** - HTML conversion with beautiful gradients
2. **example_2_charts.pptx** - Professional charts and data visualization
3. **example_3_templates.pptx** - Modern template designs
4. **example_4_images.pptx** - Advanced image processing and effects
5. **example_5_complete.pptx** - Complete business presentation

## 🎯 Best Practices

### HTML to PowerPoint

1. **Use web-safe fonts**: Arial, Helvetica, Times New Roman, Georgia, Courier New, Verdana, Tahoma, Trebuchet MS
2. **Always wrap text**: Use `<p>`, `<h1>`-`<h6>`, `<ul>`, or `<ol>` tags
3. **Use points (pt) for dimensions**: Matches PowerPoint's native units
4. **Gradient performance**: Pre-render complex gradients for better performance

### Chart Design

1. **Limit data series**: 3-5 series for bar/line charts
2. **Use consistent color schemes**: Stick to one scheme per presentation
3. **Add context**: Always include titles and labels
4. **Optimize size**: Use appropriate chart dimensions for readability

### Performance Optimization

1. **Image size**: Keep images under 1MB when possible
2. **Gradient resolution**: Use 96-150 DPI for gradients
3. **Slide count**: Break large presentations into multiple files
4. **Validation**: Always validate after creation

## 🆚 Comparison with Node.js Version

| Feature | Python Version | Node.js Version |
|---------|---------------|-----------------|
| HTML to PPTX | ✅ Full support | ✅ Full support |
| CSS Gradients | ✅ Rendered as images | ✅ Native rendering |
| Charts | ✅ Matplotlib (beautiful) | ✅ PptxGenJS |
| Image Processing | ✅ Pillow | ✅ Sharp |
| Template Management | ✅ Full support | ✅ Full support |
| Cross-platform | ✅ Pure Python | ⚠️ Requires Node.js |
| Dependencies | 10 Python packages | 15+ npm packages |
| Performance | Fast | Very fast |
| Learning Curve | Low (Python) | Medium (Node.js) |

## 📋 Requirements

```
python-pptx>=0.6.21
Pillow>=10.0.0
matplotlib>=3.7.0
beautifulsoup4>=4.12.0
lxml>=4.9.0
cairosvg>=2.7.0
cssselect>=1.2.0
tinycss2>=1.2.0
reportlab>=4.0.0
svglib>=1.5.0
```

## 🔧 Advanced Usage

### Custom Color Schemes

```python
from pptx_charts import ChartGenerator

# Add custom color scheme
ChartGenerator.COLOR_SCHEMES['mycustom'] = [
    (255, 99, 71),   # Tomato
    (60, 179, 113),  # Medium Sea Green
    (106, 90, 205),  # Slate Blue
    (255, 215, 0),   # Gold
]

# Use it
ChartGenerator.add_matplotlib_chart(
    slide, 1, 1, 8, 4,
    ChartGenerator.create_beautiful_bar_chart,
    categories=categories,
    data_series=data_series,
    color_scheme='mycustom'
)
```

### Complex Layouts

```python
from html2pptx import HTML2PPTX

html = """
<div style="width: 720pt; height: 405pt; background: white;">
    <!-- Header -->
    <div style="position: absolute; left: 0; top: 0; width: 720pt; height: 60pt;
                background: linear-gradient(90deg, #4f46e5 0%, #7c3aed 100%);"></div>

    <!-- Title -->
    <h1 style="position: absolute; left: 30pt; top: 10pt; color: white;
               font-size: 36pt; font-family: Arial;">Complex Layout</h1>

    <!-- Left column -->
    <div style="position: absolute; left: 30pt; top: 80pt; width: 320pt; height: 300pt;
                background: #f3f4f6; border: 2px solid #e5e7eb;">
        <h2 style="position: absolute; left: 20pt; top: 20pt; font-size: 24pt;">Column 1</h2>
        <p style="position: absolute; left: 20pt; top: 60pt; width: 280pt; font-size: 16pt;">
            Content here
        </p>
    </div>

    <!-- Right column -->
    <div style="position: absolute; left: 370pt; top: 80pt; width: 320pt; height: 300pt;
                background: #f3f4f6; border: 2px solid #e5e7eb;">
        <h2 style="position: absolute; left: 20pt; top: 20pt; font-size: 24pt;">Column 2</h2>
        <p style="position: absolute; left: 20pt; top: 60pt; width: 280pt; font-size: 16pt;">
            More content
        </p>
    </div>
</div>
"""

converter = HTML2PPTX()
converter.add_slide_from_html(html)
converter.save('complex_layout.pptx')
```

## 🐛 Troubleshooting

### Common Issues

**Issue**: Gradients look pixelated
- **Solution**: Increase gradient image resolution (multiply width/height by 1.5-2)

**Issue**: Fonts not displaying correctly
- **Solution**: Use only web-safe fonts from the supported list

**Issue**: Text overflow
- **Solution**: Adjust text box dimensions or reduce font size

**Issue**: Large file size
- **Solution**: Optimize images before adding, reduce gradient resolution

**Issue**: Charts not displaying
- **Solution**: Ensure matplotlib backend is set to 'Agg'

## 📄 License

This project is part of the Algorithms repository.

## 🤝 Contributing

Contributions welcome! This is a pure-Python implementation designed to match and exceed Node.js PowerPoint capabilities.

## 📚 Additional Resources

- [python-pptx Documentation](https://python-pptx.readthedocs.io/)
- [Matplotlib Gallery](https://matplotlib.org/stable/gallery/)
- [Pillow Documentation](https://pillow.readthedocs.io/)

## ✨ What Makes This Special

Unlike the Node.js version that requires Chromium, Sharp, and multiple dependencies, this Python implementation:

1. **Pure Python** - No external binaries required
2. **Beautiful by Default** - Professional color schemes and modern templates
3. **Powerful Charts** - Matplotlib integration for publication-quality visualizations
4. **Gradient Mastery** - Render any CSS gradient perfectly
5. **Easy to Deploy** - Simple pip install, works anywhere Python runs
6. **Well Documented** - Comprehensive examples and clear API

## 🎉 Success Stories

This toolkit enables you to:
- Create investor pitch decks with stunning visuals
- Generate automated reports with dynamic charts
- Build template-based presentations at scale
- Convert HTML slides to PowerPoint for offline use
- Design modern, gradient-rich presentations

---

**Ready to create amazing PowerPoint presentations with Python? Start with `examples.py`!**
