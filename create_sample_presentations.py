"""
Create 5 Professional Sample Presentations
Each with 7 slides demonstrating various features
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx_charts import ChartGenerator
from pptx_templates import ModernTemplates
from pptx_images import ImageProcessor
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from io import BytesIO


def create_business_quarterly_review():
    """Sample 1: Business Quarterly Review Presentation"""
    print("\n1. Creating Business Quarterly Review...")

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    color = (59, 130, 246)  # Professional blue

    # Slide 1: Title
    ModernTemplates.create_title_slide(
        prs, "Q4 2024 Business Review", "Strategic Performance Analysis", color
    )

    # Slide 2: Executive Summary
    ModernTemplates.create_content_slide(
        prs, "Executive Summary",
        [
            "Revenue exceeded targets by 23%, reaching $245M",
            "Customer acquisition increased 35% year-over-year",
            "Launched 4 new product lines successfully",
            "Expanded into 3 new international markets",
            "Employee satisfaction rating improved to 4.7/5",
            "Operating margin improved to 28%"
        ],
        color, 'two_column'
    )

    # Slide 3: Revenue Performance
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Quarterly Revenue Performance"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(*color)
    title_box.line.fill.background()

    ChartGenerator.add_matplotlib_chart(
        slide, 0.75, 1.1, 8.5, 4,
        ChartGenerator.create_beautiful_bar_chart,
        categories=['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024'],
        data_series=[
            {'name': 'Revenue ($M)', 'values': [185, 205, 225, 245]},
            {'name': 'Target ($M)', 'values': [180, 200, 220, 240]},
        ],
        title='',
        color_scheme='corporate'
    )

    # Slide 4: Customer Growth
    slide2 = prs.slides.add_slide(blank_layout)

    title_box = slide2.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Customer Acquisition Trends"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(*color)
    title_box.line.fill.background()

    ChartGenerator.add_matplotlib_chart(
        slide2, 0.75, 1.1, 8.5, 4,
        ChartGenerator.create_beautiful_line_chart,
        categories=['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        data_series=[
            {'name': 'New Customers', 'values': [450, 520, 580, 650, 720, 800]},
            {'name': 'Active Users', 'values': [2800, 3100, 3400, 3700, 4100, 4500]}
        ],
        title='',
        color_scheme='corporate'
    )

    # Slide 5: Key Achievements
    ModernTemplates.create_content_slide(
        prs, "Major Achievements",
        [
            "Product Innovation: Launched AI-powered analytics platform",
            "Market Expansion: Entered APAC and EMEA regions",
            "Partnership: Strategic alliance with Fortune 500 companies",
            "Team Growth: Hired 150+ talented professionals",
            "Recognition: Named Industry Leader by Gartner",
            "Sustainability: Achieved carbon-neutral operations"
        ],
        color, 'two_column'
    )

    # Slide 6: Strategic Priorities
    ModernTemplates.create_content_slide(
        prs, "2025 Strategic Priorities",
        [
            "Accelerate AI/ML product development",
            "Scale operations in new markets",
            "Enhance customer success programs",
            "Invest in talent development",
            "Drive operational efficiency",
            "Strengthen competitive positioning"
        ],
        color, 'two_column'
    )

    # Slide 7: Closing
    ModernTemplates.create_quote_slide(
        prs,
        "Excellence is not a destination; it is a continuous journey that never ends",
        "Brian Tracy",
        color
    )

    prs.save('samples/sample_1_business_review.pptx')
    print("   ✓ Created: sample_1_business_review.pptx")


def create_product_launch():
    """Sample 2: Product Launch Presentation"""
    print("\n2. Creating Product Launch Presentation...")

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    color = (236, 72, 153)  # Vibrant pink
    processor = ImageProcessor()

    # Slide 1: Title with gradient
    gradient = processor.create_multi_color_gradient(
        1500, 844,
        [(236, 72, 153), (168, 85, 247), (59, 130, 246)]
    )
    img_stream = BytesIO()
    gradient.save(img_stream, format='PNG')
    img_stream.seek(0)

    blank_layout = prs.slide_layouts[6]
    slide1 = prs.slides.add_slide(blank_layout)
    slide1.shapes.add_picture(img_stream, 0, 0, prs.slide_width, prs.slide_height)

    title_box = slide1.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = "Introducing CloudSync Pro"
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_ALIGN.CENTER
    title_para.font.size = Pt(60)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_box.line.fill.background()

    subtitle_box = slide1.shapes.add_textbox(Inches(1), Inches(3.5), Inches(8), Inches(0.8))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "Next-Generation Cloud Storage Solution"
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.alignment = PP_ALIGN.CENTER
    subtitle_para.font.size = Pt(28)
    subtitle_para.font.color.rgb = RGBColor(255, 255, 255)
    subtitle_box.line.fill.background()

    # Slide 2: The Problem
    ModernTemplates.create_content_slide(
        prs, "The Problem We're Solving",
        [
            "Data fragmentation across multiple platforms",
            "Security concerns with current solutions",
            "Slow sync speeds affecting productivity",
            "Limited collaboration features",
            "High costs for enterprise storage",
            "Complex user interfaces"
        ],
        color, 'two_column'
    )

    # Slide 3: Our Solution
    ModernTemplates.create_content_slide(
        prs, "CloudSync Pro Solution",
        [
            "Unified storage platform with AI organization",
            "Military-grade end-to-end encryption",
            "10x faster sync with edge computing",
            "Real-time collaboration tools built-in",
            "60% cost reduction vs. competitors",
            "Intuitive, award-winning interface"
        ],
        color, 'two_column'
    )

    # Slide 4: Key Features
    slide4 = prs.slides.add_slide(blank_layout)
    fill = slide4.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(248, 250, 252)

    title_box = slide4.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.7))
    title_frame = title_box.text_frame
    title_frame.text = "Revolutionary Features"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(*color)
    title_box.line.fill.background()

    features = [
        ("🚀", "Lightning Fast", "10x faster sync speed"),
        ("🔒", "Ultra Secure", "Military-grade encryption"),
        ("🤖", "AI-Powered", "Smart file organization"),
        ("💰", "Cost Effective", "60% savings guaranteed"),
        ("🌐", "Global CDN", "Access from anywhere"),
        ("👥", "Team Ready", "Built for collaboration")
    ]

    for i, (icon, title, desc) in enumerate(features):
        row = i // 3
        col = i % 3
        x = 0.5 + col * 3.2
        y = 1.5 + row * 1.8

        icon_box = slide4.shapes.add_textbox(Inches(x), Inches(y), Inches(0.5), Inches(0.5))
        icon_frame = icon_box.text_frame
        icon_frame.text = icon
        icon_para = icon_frame.paragraphs[0]
        icon_para.font.size = Pt(36)
        icon_box.line.fill.background()

        title_box = slide4.shapes.add_textbox(Inches(x + 0.6), Inches(y), Inches(2.4), Inches(0.4))
        title_frame = title_box.text_frame
        title_frame.text = title
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(16)
        title_para.font.bold = True
        title_box.line.fill.background()

        desc_box = slide4.shapes.add_textbox(Inches(x + 0.6), Inches(y + 0.35), Inches(2.4), Inches(0.6))
        desc_frame = desc_box.text_frame
        desc_frame.text = desc
        desc_para = desc_frame.paragraphs[0]
        desc_para.font.size = Pt(12)
        desc_box.line.fill.background()

    # Slide 5: Market Opportunity
    slide5 = prs.slides.add_slide(blank_layout)

    title_box = slide5.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Market Opportunity"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(*color)
    title_box.line.fill.background()

    ChartGenerator.add_matplotlib_chart(
        slide5, 2, 1.1, 6, 4,
        ChartGenerator.create_beautiful_pie_chart,
        categories=['Enterprise (42%)', 'SMB (28%)', 'Individual (18%)', 'Education (12%)'],
        values=[42, 28, 18, 12],
        title='',
        color_scheme='vibrant'
    )

    # Slide 6: Pricing
    ModernTemplates.create_content_slide(
        prs, "Simple, Transparent Pricing",
        [
            "Starter: $9/month - 100GB, 3 users",
            "Professional: $29/month - 1TB, 10 users",
            "Business: $99/month - 10TB, unlimited users",
            "Enterprise: Custom - Unlimited everything",
            "All plans include: 24/7 support, 99.9% uptime SLA",
            "30-day money-back guarantee"
        ],
        color, 'two_column'
    )

    # Slide 7: Call to Action
    gradient2 = processor.create_gradient_image(1500, 844, color, (168, 85, 247), 135)
    img_stream2 = BytesIO()
    gradient2.save(img_stream2, format='PNG')
    img_stream2.seek(0)

    slide7 = prs.slides.add_slide(blank_layout)
    slide7.shapes.add_picture(img_stream2, 0, 0, prs.slide_width, prs.slide_height)

    cta_box = slide7.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
    cta_frame = cta_box.text_frame
    cta_frame.text = "Start Your Free Trial Today"
    cta_para = cta_frame.paragraphs[0]
    cta_para.alignment = PP_ALIGN.CENTER
    cta_para.font.size = Pt(54)
    cta_para.font.bold = True
    cta_para.font.color.rgb = RGBColor(255, 255, 255)
    cta_box.line.fill.background()

    link_box = slide7.shapes.add_textbox(Inches(1), Inches(3.7), Inches(8), Inches(0.8))
    link_frame = link_box.text_frame
    link_frame.text = "www.cloudsyncpro.com/trial"
    link_para = link_frame.paragraphs[0]
    link_para.alignment = PP_ALIGN.CENTER
    link_para.font.size = Pt(32)
    link_para.font.color.rgb = RGBColor(255, 255, 255)
    link_box.line.fill.background()

    prs.save('samples/sample_2_product_launch.pptx')
    print("   ✓ Created: sample_2_product_launch.pptx")


def create_marketing_strategy():
    """Sample 3: Marketing Strategy Presentation"""
    print("\n3. Creating Marketing Strategy Presentation...")

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    color = (34, 197, 94)  # Green

    # Slide 1: Title
    ModernTemplates.create_title_slide(
        prs, "2025 Marketing Strategy", "Digital-First Growth Plan", color
    )

    # Slide 2: Current Situation
    ModernTemplates.create_content_slide(
        prs, "Current Market Position",
        [
            "Market Share: 18% (up from 12% last year)",
            "Brand Awareness: 67% in target demographic",
            "Digital Presence: 2.5M social media followers",
            "Website Traffic: 500K monthly visitors",
            "Email List: 180K engaged subscribers",
            "Customer NPS Score: 68 (Industry: 45)"
        ],
        color, 'two_column'
    )

    # Slide 3: Channel Performance
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Marketing Channel Performance"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(*color)
    title_box.line.fill.background()

    ChartGenerator.add_matplotlib_chart(
        slide, 0.75, 1.1, 8.5, 4,
        ChartGenerator.create_beautiful_bar_chart,
        categories=['Social Media', 'Email', 'SEO', 'Paid Ads', 'Content'],
        data_series=[
            {'name': 'ROI (%)', 'values': [340, 420, 380, 280, 450]},
            {'name': 'Engagement', 'values': [85, 92, 78, 65, 95]}
        ],
        title='',
        color_scheme='modern'
    )

    # Slide 4: Strategic Objectives
    ModernTemplates.create_content_slide(
        prs, "2025 Strategic Objectives",
        [
            "Increase market share to 25% by Q4",
            "Grow digital audience to 5M followers",
            "Achieve 1M monthly website visitors",
            "Launch in 3 new international markets",
            "Increase customer lifetime value by 40%",
            "Maintain NPS above 70"
        ],
        color, 'two_column'
    )

    # Slide 5: Key Initiatives
    ModernTemplates.create_content_slide(
        prs, "Key Marketing Initiatives",
        [
            "Content Hub: Launch weekly video series and podcast",
            "Influencer Program: Partner with 50+ micro-influencers",
            "SEO Expansion: Target 200 new high-value keywords",
            "Marketing Automation: Implement AI-driven personalization",
            "Community Building: Create customer advocacy program",
            "Brand Refresh: Update visual identity and messaging"
        ],
        color, 'two_column'
    )

    # Slide 6: Budget Allocation
    slide6 = prs.slides.add_slide(blank_layout)

    title_box = slide6.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Budget Allocation ($5M Total)"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(*color)
    title_box.line.fill.background()

    ChartGenerator.add_matplotlib_chart(
        slide6, 2, 1.1, 6, 4,
        ChartGenerator.create_beautiful_pie_chart,
        categories=['Content Creation', 'Paid Media', 'Tools & Tech', 'Events', 'Team'],
        values=[30, 25, 20, 15, 10],
        title='',
        color_scheme='professional'
    )

    # Slide 7: Expected Results
    ModernTemplates.create_quote_slide(
        prs,
        "Marketing is no longer about the stuff you make, but the stories you tell",
        "Seth Godin",
        color
    )

    prs.save('samples/sample_3_marketing_strategy.pptx')
    print("   ✓ Created: sample_3_marketing_strategy.pptx")


def create_technology_overview():
    """Sample 4: Technology Overview Presentation"""
    print("\n4. Creating Technology Overview Presentation...")

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    color = (99, 102, 241)  # Indigo
    processor = ImageProcessor()

    # Slide 1: Title
    gradient = processor.create_gradient_image(1500, 844, color, (139, 92, 246), 135)
    img_stream = BytesIO()
    gradient.save(img_stream, format='PNG')
    img_stream.seek(0)

    blank_layout = prs.slide_layouts[6]
    slide1 = prs.slides.add_slide(blank_layout)
    slide1.shapes.add_picture(img_stream, 0, 0, prs.slide_width, prs.slide_height)

    title_box = slide1.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = "AI & Machine Learning"
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_ALIGN.CENTER
    title_para.font.size = Pt(60)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_box.line.fill.background()

    subtitle_box = slide1.shapes.add_textbox(Inches(1), Inches(3.5), Inches(8), Inches(0.8))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "Technology Overview & Implementation Guide"
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.alignment = PP_ALIGN.CENTER
    subtitle_para.font.size = Pt(26)
    subtitle_para.font.color.rgb = RGBColor(255, 255, 255)
    subtitle_box.line.fill.background()

    # Slide 2: What is AI/ML?
    ModernTemplates.create_content_slide(
        prs, "Understanding AI & Machine Learning",
        [
            "AI: Computer systems that mimic human intelligence",
            "ML: Algorithms that learn from data patterns",
            "Deep Learning: Neural networks with multiple layers",
            "Natural Language Processing: Understanding human language",
            "Computer Vision: Analyzing and interpreting images",
            "Reinforcement Learning: Learning through trial and error"
        ],
        color, 'two_column'
    )

    # Slide 3: Current Applications
    ModernTemplates.create_content_slide(
        prs, "Real-World Applications Today",
        [
            "Healthcare: Disease diagnosis and drug discovery",
            "Finance: Fraud detection and algorithmic trading",
            "Retail: Personalized recommendations and inventory",
            "Transportation: Autonomous vehicles and route optimization",
            "Manufacturing: Quality control and predictive maintenance",
            "Entertainment: Content creation and recommendation systems"
        ],
        color, 'two_column'
    )

    # Slide 4: Technology Stack
    slide4 = prs.slides.add_slide(blank_layout)
    fill = slide4.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(248, 250, 252)

    title_box = slide4.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.7))
    title_frame = title_box.text_frame
    title_frame.text = "Core Technology Stack"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(*color)
    title_box.line.fill.background()

    stack = [
        ("🐍", "Python", "Primary language"),
        ("🧠", "TensorFlow", "Deep learning"),
        ("⚡", "PyTorch", "Research & prod"),
        ("☁️", "AWS/Azure", "Cloud platform"),
        ("🗄️", "PostgreSQL", "Data storage"),
        ("📊", "Jupyter", "Development")
    ]

    for i, (icon, tech, desc) in enumerate(stack):
        row = i // 3
        col = i % 3
        x = 0.5 + col * 3.2
        y = 1.5 + row * 1.8

        icon_box = slide4.shapes.add_textbox(Inches(x), Inches(y), Inches(0.5), Inches(0.5))
        icon_frame = icon_box.text_frame
        icon_frame.text = icon
        icon_para = icon_frame.paragraphs[0]
        icon_para.font.size = Pt(36)
        icon_box.line.fill.background()

        tech_box = slide4.shapes.add_textbox(Inches(x + 0.6), Inches(y), Inches(2.4), Inches(0.4))
        tech_frame = tech_box.text_frame
        tech_frame.text = tech
        tech_para = tech_frame.paragraphs[0]
        tech_para.font.size = Pt(16)
        tech_para.font.bold = True
        tech_box.line.fill.background()

        desc_box = slide4.shapes.add_textbox(Inches(x + 0.6), Inches(y + 0.35), Inches(2.4), Inches(0.6))
        desc_frame = desc_box.text_frame
        desc_frame.text = desc
        desc_para = desc_frame.paragraphs[0]
        desc_para.font.size = Pt(12)
        desc_box.line.fill.background()

    # Slide 5: Performance Metrics
    slide5 = prs.slides.add_slide(blank_layout)

    title_box = slide5.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Model Performance Improvements"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(*color)
    title_box.line.fill.background()

    ChartGenerator.add_matplotlib_chart(
        slide5, 0.75, 1.1, 8.5, 4,
        ChartGenerator.create_beautiful_line_chart,
        categories=['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        data_series=[
            {'name': 'Accuracy (%)', 'values': [78, 82, 86, 89, 92, 95]},
            {'name': 'Speed (ms)', 'values': [450, 380, 320, 280, 230, 180]}
        ],
        title='',
        color_scheme='modern'
    )

    # Slide 6: Implementation Roadmap
    ModernTemplates.create_content_slide(
        prs, "6-Month Implementation Roadmap",
        [
            "Month 1-2: Data collection and preprocessing",
            "Month 2-3: Model architecture design and testing",
            "Month 3-4: Training and hyperparameter tuning",
            "Month 4-5: Validation and performance optimization",
            "Month 5-6: Deployment and monitoring setup",
            "Month 6+: Continuous improvement and scaling"
        ],
        color, 'two_column'
    )

    # Slide 7: Future Vision
    ModernTemplates.create_quote_slide(
        prs,
        "Artificial Intelligence is the new electricity",
        "Andrew Ng",
        color
    )

    prs.save('samples/sample_4_technology_overview.pptx')
    print("   ✓ Created: sample_4_technology_overview.pptx")


def create_training_module():
    """Sample 5: Educational Training Module"""
    print("\n5. Creating Educational Training Module...")

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    color = (251, 146, 60)  # Orange

    # Slide 1: Title
    ModernTemplates.create_title_slide(
        prs, "Effective Leadership", "Essential Skills for Modern Managers", color
    )

    # Slide 2: Learning Objectives
    ModernTemplates.create_content_slide(
        prs, "What You'll Learn Today",
        [
            "Understand core leadership principles and styles",
            "Develop effective communication strategies",
            "Master the art of delegation and empowerment",
            "Build high-performing, motivated teams",
            "Navigate difficult conversations with confidence",
            "Create a culture of continuous improvement"
        ],
        color, 'two_column'
    )

    # Slide 3: Leadership Styles
    blank_layout = prs.slide_layouts[6]
    slide3 = prs.slides.add_slide(blank_layout)
    fill = slide3.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(248, 250, 252)

    title_box = slide3.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.7))
    title_frame = title_box.text_frame
    title_frame.text = "Five Leadership Styles"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(*color)
    title_box.line.fill.background()

    styles = [
        ("👑", "Autocratic", "Quick decisions, clear direction"),
        ("🤝", "Democratic", "Team input, collaborative"),
        ("🎯", "Transformational", "Inspire & innovate"),
        ("📋", "Transactional", "Structure & rewards"),
        ("🌱", "Servant", "Team-first approach"),
        ("🚀", "Visionary", "Future-focused")
    ]

    for i, (icon, style, desc) in enumerate(styles):
        row = i // 3
        col = i % 3
        x = 0.5 + col * 3.2
        y = 1.5 + row * 1.8

        icon_box = slide3.shapes.add_textbox(Inches(x), Inches(y), Inches(0.5), Inches(0.5))
        icon_frame = icon_box.text_frame
        icon_frame.text = icon
        icon_para = icon_frame.paragraphs[0]
        icon_para.font.size = Pt(36)
        icon_box.line.fill.background()

        style_box = slide3.shapes.add_textbox(Inches(x + 0.6), Inches(y), Inches(2.4), Inches(0.4))
        style_frame = style_box.text_frame
        style_frame.text = style
        style_para = style_frame.paragraphs[0]
        style_para.font.size = Pt(16)
        style_para.font.bold = True
        style_box.line.fill.background()

        desc_box = slide3.shapes.add_textbox(Inches(x + 0.6), Inches(y + 0.35), Inches(2.4), Inches(0.6))
        desc_frame = desc_box.text_frame
        desc_frame.text = desc
        desc_para = desc_frame.paragraphs[0]
        desc_para.font.size = Pt(12)
        desc_box.line.fill.background()

    # Slide 4: Communication Skills
    ModernTemplates.create_content_slide(
        prs, "Essential Communication Skills",
        [
            "Active Listening: Give full attention, ask clarifying questions",
            "Clear Messaging: Be concise, specific, and purposeful",
            "Empathy: Understand perspectives and emotions",
            "Feedback: Provide constructive, timely input",
            "Non-Verbal: Master body language and tone",
            "Adaptability: Adjust style to audience and situation"
        ],
        color, 'two_column'
    )

    # Slide 5: Team Performance
    slide5 = prs.slides.add_slide(blank_layout)

    title_box = slide5.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "Impact of Leadership on Team Performance"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(30)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(*color)
    title_box.line.fill.background()

    ChartGenerator.add_matplotlib_chart(
        slide5, 0.75, 1.1, 8.5, 4,
        ChartGenerator.create_beautiful_bar_chart,
        categories=['Productivity', 'Engagement', 'Retention', 'Innovation'],
        data_series=[
            {'name': 'Poor Leadership', 'values': [55, 42, 38, 35]},
            {'name': 'Good Leadership', 'values': [92, 88, 85, 78]}
        ],
        title='',
        color_scheme='vibrant'
    )

    # Slide 6: Action Plan
    ModernTemplates.create_content_slide(
        prs, "Your 30-Day Action Plan",
        [
            "Week 1: Schedule 1-on-1s with each team member",
            "Week 2: Implement daily stand-ups or check-ins",
            "Week 3: Delegate one major project to develop others",
            "Week 4: Gather 360-degree feedback on your leadership",
            "Ongoing: Read one leadership book per month",
            "Ongoing: Find a mentor or join a leadership community"
        ],
        color, 'two_column'
    )

    # Slide 7: Inspiration
    ModernTemplates.create_quote_slide(
        prs,
        "Leadership is not about being in charge. It's about taking care of those in your charge",
        "Simon Sinek",
        color
    )

    prs.save('samples/sample_5_training_module.pptx')
    print("   ✓ Created: sample_5_training_module.pptx")


if __name__ == '__main__':
    print("="*70)
    print("Creating 5 Professional Sample Presentations")
    print("Each with 7 slides and rich text content")
    print("="*70)

    create_business_quarterly_review()
    create_product_launch()
    create_marketing_strategy()
    create_technology_overview()
    create_training_module()

    print("\n" + "="*70)
    print("✓ All 5 sample presentations created successfully!")
    print("="*70)
    print("\nFiles created in samples/:")
    print("  1. sample_1_business_review.pptx - Quarterly business review")
    print("  2. sample_2_product_launch.pptx - Product launch presentation")
    print("  3. sample_3_marketing_strategy.pptx - Marketing strategy deck")
    print("  4. sample_4_technology_overview.pptx - AI/ML technology overview")
    print("  5. sample_5_training_module.pptx - Leadership training module")
