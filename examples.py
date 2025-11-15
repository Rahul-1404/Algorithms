"""
Comprehensive Examples for Python PowerPoint Creation
Demonstrates all features and capabilities
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from html2pptx import HTML2PPTX, create_presentation_from_html
from pptx_charts import ChartGenerator
from pptx_images import ImageProcessor
from pptx_templates import ModernTemplates, TemplateManager
from pptx_utils import Validator, ThumbnailGenerator, PPTXHelper
from io import BytesIO


def example_1_html_to_pptx():
    """Example 1: Create presentation from HTML with gradients and styling"""
    print("Example 1: HTML to PowerPoint with Beautiful Gradients")

    slides = [
        # Title slide with gradient
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
            <h1 style="position: absolute; left: 50pt; top: 150pt; width: 620pt; color: white;
                       font-size: 54pt; font-family: Arial; text-align: center; font-weight: bold;">
                Python PowerPoint Mastery
            </h1>
            <p style="position: absolute; left: 50pt; top: 250pt; width: 620pt; color: white;
                      font-size: 28pt; font-family: Arial; text-align: center;">
                Professional Presentations Without Node.js
            </p>
        </div>
        """,

        # Content slide with multiple sections
        """
        <div style="width: 720pt; height: 405pt; background: white;">
            <div style="position: absolute; left: 0; top: 0; width: 720pt; height: 80pt;
                       background: linear-gradient(90deg, #4f46e5 0%, #7c3aed 100%);"></div>

            <h2 style="position: absolute; left: 30pt; top: 15pt; color: white;
                       font-size: 36pt; font-family: Arial; font-weight: bold;">
                Key Features
            </h2>

            <div style="position: absolute; left: 50pt; top: 120pt; width: 300pt; height: 250pt;
                       background: #f3f4f6; border: 2px solid #e5e7eb;">
                <h3 style="position: absolute; left: 20pt; top: 20pt; width: 260pt;
                           font-size: 24pt; font-family: Arial; color: #4f46e5;">
                    HTML Conversion
                </h3>
                <p style="position: absolute; left: 20pt; top: 70pt; width: 260pt;
                          font-size: 16pt; font-family: Arial; color: #374151;">
                    Convert HTML/CSS to PowerPoint with precise positioning and styling support.
                </p>
            </div>

            <div style="position: absolute; left: 370pt; top: 120pt; width: 300pt; height: 250pt;
                       background: #f3f4f6; border: 2px solid #e5e7eb;">
                <h3 style="position: absolute; left: 20pt; top: 20pt; width: 260pt;
                           font-size: 24pt; font-family: Arial; color: #7c3aed;">
                    Beautiful Charts
                </h3>
                <p style="position: absolute; left: 20pt; top: 70pt; width: 260pt;
                          font-size: 16pt; font-family: Arial; color: #374151;">
                    Professional charts with matplotlib integration and custom styling.
                </p>
            </div>
        </div>
        """,

        # Gradient showcase
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(45deg, #fa7c91 0%, #f8cd4a 50%, #56ccf2 100%);">
            <h1 style="position: absolute; left: 50pt; top: 160pt; width: 620pt; color: white;
                       font-size: 48pt; font-family: Arial; text-align: center; font-weight: bold;
                       text-shadow: 2px 2px 4px rgba(0,0,0,0.3);">
                Beautiful Gradients Rendered as Images
            </h1>
        </div>
        """
    ]

    create_presentation_from_html(slides, 'example_1_html.pptx', '16:9')
    print("✓ Created: example_1_html.pptx\n")


def example_2_charts_showcase():
    """Example 2: Beautiful charts and data visualization"""
    print("Example 2: Professional Charts and Visualizations")

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    color_scheme = (79, 70, 229)  # Modern indigo

    # Title
    ModernTemplates.create_title_slide(
        prs, "Data Visualization", "Professional Charts with Python", color_scheme
    )

    # Bar chart slide
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    categories = ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024']
    data_series = [
        {'name': 'Revenue', 'values': [150, 180, 220, 250]},
        {'name': 'Profit', 'values': [45, 60, 75, 90]},
        {'name': 'Growth', 'values': [30, 40, 55, 70]}
    ]

    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Quarterly Performance Analysis"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_box.line.fill.background()

    # Beautiful matplotlib chart
    ChartGenerator.add_matplotlib_chart(
        slide, 0.5, 1.0, 9, 4,
        ChartGenerator.create_beautiful_bar_chart,
        categories=categories,
        data_series=data_series,
        title='',
        color_scheme='modern'
    )

    # Line chart slide
    slide2 = prs.slides.add_slide(blank_layout)

    title_box = slide2.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Growth Trends Over Time"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_box.line.fill.background()

    ChartGenerator.add_matplotlib_chart(
        slide2, 0.5, 1.0, 9, 4,
        ChartGenerator.create_beautiful_line_chart,
        categories=categories,
        data_series=data_series[:2],
        title='',
        color_scheme='vibrant'
    )

    # Pie chart slide
    slide3 = prs.slides.add_slide(blank_layout)

    title_box = slide3.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Market Share Distribution"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_box.line.fill.background()

    ChartGenerator.add_matplotlib_chart(
        slide3, 2, 1.0, 6, 4,
        ChartGenerator.create_beautiful_pie_chart,
        categories=['Product A', 'Product B', 'Product C', 'Product D'],
        values=[35, 25, 20, 20],
        title='',
        color_scheme='professional'
    )

    prs.save('example_2_charts.pptx')
    print("✓ Created: example_2_charts.pptx\n")


def example_3_modern_templates():
    """Example 3: Modern template designs"""
    print("Example 3: Modern Professional Templates")

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    color = (236, 72, 153)  # Modern pink

    # Title slide
    ModernTemplates.create_title_slide(
        prs, "Modern Presentation Design", "Sleek and Professional", color
    )

    # Section divider
    ModernTemplates.create_section_slide(prs, "Part 1: Introduction", color)

    # Content slides with different layouts
    ModernTemplates.create_content_slide(
        prs, "Single Column Layout",
        ["First key point with detailed explanation",
         "Second important insight",
         "Third crucial element",
         "Fourth supporting detail",
         "Final summary point"],
        color, 'single_column'
    )

    ModernTemplates.create_content_slide(
        prs, "Two Column Layout",
        ["Point 1", "Point 2", "Point 3", "Point 4", "Point 5", "Point 6"],
        color, 'two_column'
    )

    ModernTemplates.create_content_slide(
        prs, "Three Column Layout",
        ["A1", "A2", "A3", "B1", "B2", "B3", "C1", "C2", "C3"],
        color, 'three_column'
    )

    # Quote slide
    ModernTemplates.create_quote_slide(
        prs,
        "The best way to predict the future is to create it",
        "Peter Drucker",
        color
    )

    prs.save('example_3_templates.pptx')
    print("✓ Created: example_3_templates.pptx\n")


def example_4_image_processing():
    """Example 4: Advanced image processing and effects"""
    print("Example 4: Advanced Image Processing")

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    processor = ImageProcessor()

    # Gradient backgrounds slide
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Multi-color gradient background
    gradient = processor.create_multi_color_gradient(
        1500, 844,
        [(102, 126, 234), (118, 75, 162), (237, 100, 166)]
    )

    img_stream = BytesIO()
    gradient.save(img_stream, format='PNG')
    img_stream.seek(0)

    slide.shapes.add_picture(img_stream, 0, 0, prs.slide_width, prs.slide_height)

    # Add title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = "Advanced Gradient Effects"
    title_para = title_frame.paragraphs[0]
    title_para.alignment = 1  # Center
    title_para.font.size = Pt(54)
    title_para.font.bold = True
    title_para.font.color.rgb = (255, 255, 255)
    title_box.line.fill.background()

    # Icons slide
    slide2 = prs.slides.add_slide(blank_layout)

    # Background
    fill = slide2.background.fill
    fill.solid()
    fill.fore_color.rgb = (248, 250, 252)

    # Title
    title_box = slide2.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.7))
    title_frame = title_box.text_frame
    title_frame.text = "Custom Vector Icons"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_box.line.fill.background()

    # Add various icons
    icons = ['circle', 'square', 'triangle', 'star', 'checkmark', 'arrow']
    colors = [(79, 70, 229), (236, 72, 153), (251, 146, 60), (34, 197, 94), (14, 165, 233), (168, 85, 247)]

    for i, (icon_type, color) in enumerate(zip(icons, colors)):
        icon_img = processor.create_icon(icon_type, 200, color)

        icon_stream = BytesIO()
        icon_img.save(icon_stream, format='PNG')
        icon_stream.seek(0)

        x = 1.5 + (i % 3) * 2.5
        y = 1.8 + (i // 3) * 2

        slide2.shapes.add_picture(icon_stream, Inches(x), Inches(y), Inches(1.5), Inches(1.5))

    prs.save('example_4_images.pptx')
    print("✓ Created: example_4_images.pptx\n")


def example_5_complete_presentation():
    """Example 5: Complete professional presentation"""
    print("Example 5: Complete Professional Business Presentation")

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    primary_color = (79, 70, 229)
    accent_color = (236, 72, 153)

    # Slide 1: Title
    ModernTemplates.create_title_slide(
        prs,
        "Q4 2024 Business Review",
        "Strategic Insights & Performance Analysis",
        primary_color
    )

    # Slide 2: Agenda
    ModernTemplates.create_content_slide(
        prs, "Agenda",
        [
            "Executive Summary",
            "Financial Performance",
            "Market Analysis",
            "Key Achievements",
            "Future Outlook",
            "Q&A"
        ],
        primary_color, 'two_column'
    )

    # Slide 3: Section - Financial Performance
    ModernTemplates.create_section_slide(prs, "Financial Performance", accent_color)

    # Slide 4: Revenue Chart
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Revenue Growth Analysis"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_box.line.fill.background()

    ChartGenerator.add_matplotlib_chart(
        slide, 0.5, 1.0, 9, 4,
        ChartGenerator.create_beautiful_bar_chart,
        categories=['Q1', 'Q2', 'Q3', 'Q4'],
        data_series=[
            {'name': '2023', 'values': [120, 135, 148, 165]},
            {'name': '2024', 'values': [150, 175, 195, 220]}
        ],
        title='',
        color_scheme='modern'
    )

    # Slide 5: Market Share
    slide5 = prs.slides.add_slide(blank_layout)

    title_box = slide5.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Market Share Distribution"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_box.line.fill.background()

    ChartGenerator.add_matplotlib_chart(
        slide5, 2, 1.0, 6, 4,
        ChartGenerator.create_beautiful_pie_chart,
        categories=['Our Company', 'Competitor A', 'Competitor B', 'Others'],
        values=[40, 25, 20, 15],
        title='',
        color_scheme='professional'
    )

    # Slide 6: Key Achievements
    ModernTemplates.create_content_slide(
        prs, "Key Achievements",
        [
            "Exceeded revenue targets by 25%",
            "Launched 3 new product lines",
            "Expanded to 5 new markets",
            "Increased customer satisfaction to 95%",
            "Reduced operational costs by 15%",
            "Grew team by 40 talented individuals"
        ],
        primary_color, 'two_column'
    )

    # Slide 7: Quote
    ModernTemplates.create_quote_slide(
        prs,
        "Success is not final, failure is not fatal: it is the courage to continue that counts",
        "Winston Churchill",
        primary_color
    )

    # Slide 8: Thank You
    ModernTemplates.create_title_slide(
        prs,
        "Thank You",
        "Questions?",
        accent_color
    )

    prs.save('example_5_complete.pptx')
    print("✓ Created: example_5_complete.pptx\n")

    # Validate and generate stats
    stats = PPTXHelper.get_presentation_stats('example_5_complete.pptx')
    print(f"Presentation stats: {stats}")

    validation = Validator.validate_presentation('example_5_complete.pptx')
    print(f"Validation: Valid={validation['valid']}, Warnings={len(validation['warnings'])}")


def run_all_examples():
    """Run all examples"""
    print("=" * 60)
    print("Python PowerPoint Creation - Comprehensive Examples")
    print("=" * 60)
    print()

    example_1_html_to_pptx()
    example_2_charts_showcase()
    example_3_modern_templates()
    example_4_image_processing()
    example_5_complete_presentation()

    print("=" * 60)
    print("All examples completed successfully!")
    print("=" * 60)
    print("\nGenerated files:")
    print("  • example_1_html.pptx - HTML conversion with gradients")
    print("  • example_2_charts.pptx - Professional charts")
    print("  • example_3_templates.pptx - Modern templates")
    print("  • example_4_images.pptx - Image processing")
    print("  • example_5_complete.pptx - Complete business presentation")


if __name__ == '__main__':
    run_all_examples()
