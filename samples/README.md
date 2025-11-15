# Sample Presentations

This directory contains example PowerPoint presentations created with the Python PowerPoint Suite to demonstrate the capabilities of the library.

## 📊 Quick Start Samples

### 🌟 DEMO_PRESENTATION.pptx (87KB)
**The main showcase - 8 professional slides**

Demonstrates all features:
- **Slide 1**: Title with multi-color gradient background
- **Slide 2**: Features grid with 6 feature boxes and icons
- **Slide 3**: Bar chart - Sales performance analysis (3 data series)
- **Slide 4**: Line chart - Growth trends over 6 months
- **Slide 5**: Pie chart - Market share distribution
- **Slide 6**: Two-column content layout
- **Slide 7**: Quote slide with gradient background
- **Slide 8**: Thank you slide with gradient

**Features Showcased**:
- ✅ Multi-color CSS gradients rendered as images
- ✅ Professional matplotlib charts (bar, line, pie)
- ✅ Modern slide layouts and templates
- ✅ Rich text formatting and styling
- ✅ Consistent color scheme and branding
- ✅ Icon integration

---

### 🚀 quickstart_html.pptx (31KB)
**HTML to PowerPoint conversion demo - 2 slides**

Shows how to create slides from HTML/CSS:
- Gradient backgrounds from CSS
- Precise positioning with absolute layout
- Rich text formatting
- Web-safe fonts

**Code Example**:
```python
from html2pptx import create_presentation_from_html

slides = ["""
    <div style="width: 720pt; height: 405pt;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
        <h1 style="position: absolute; left: 50pt; top: 150pt;
                   color: white; font-size: 54pt; text-align: center;">
            Beautiful Slide
        </h1>
    </div>
"""]

create_presentation_from_html(slides, 'output.pptx', '16:9')
```

---

### 📈 quickstart_charts.pptx (44KB)
**Chart generation demo - 2 slides**

Demonstrates professional data visualization:
- Title slide with gradient
- Bar chart with multiple data series
- Professional color scheme
- Clean, modern layout

**Code Example**:
```python
from pptx_charts import ChartGenerator

ChartGenerator.add_matplotlib_chart(
    slide, 1, 1, 8, 4,
    ChartGenerator.create_beautiful_bar_chart,
    categories=['Q1', 'Q2', 'Q3', 'Q4'],
    data_series=[
        {'name': 'Sales', 'values': [100, 120, 140, 160]},
        {'name': 'Revenue', 'values': [90, 110, 130, 150]}
    ],
    color_scheme='modern'
)
```

---

### 🎨 quickstart_templates.pptx (43KB)
**Modern templates demo - 3 slides**

Shows pre-built professional templates:
- Title slide with gradient background
- Two-column content layout
- Quote slide
- Consistent styling throughout

**Code Example**:
```python
from pptx_templates import ModernTemplates

ModernTemplates.create_title_slide(
    prs, "My Presentation", "Subtitle", color=(79, 70, 229)
)

ModernTemplates.create_content_slide(
    prs, "Key Points",
    ["Point 1", "Point 2", "Point 3"],
    color=(79, 70, 229),
    layout='two_column'
)
```

---

## 🎯 Professional Sample Presentations

### 📈 Sample 1: Business Quarterly Review (125KB)
**7 slides - Complete business performance presentation**

Perfect for: Quarterly business reviews, executive presentations, board meetings

Content includes:
- Executive summary with key metrics
- Revenue performance charts (bar chart, Q1-Q4)
- Customer growth trends (line chart, 6 months)
- Major achievements and milestones
- Strategic priorities for next quarter
- Inspirational closing quote

**Features**: Corporate color scheme, professional charts, two-column layouts

---

### 🚀 Sample 2: Product Launch (80KB)
**7 slides - New product introduction**

Perfect for: Product launches, startup pitches, feature announcements

Content includes:
- Eye-catching title with multi-color gradient
- Problem statement and market pain points
- Solution overview and value proposition
- Revolutionary features with icons
- Market opportunity (pie chart)
- Transparent pricing tiers
- Strong call-to-action

**Features**: Vibrant gradients, icon integration, engaging visuals

---

### 📱 Sample 3: Marketing Strategy (112KB)
**7 slides - Digital marketing plan**

Perfect for: Marketing planning, strategy presentations, campaign proposals

Content includes:
- Current market position analysis
- Channel performance comparison (bar chart)
- 2025 strategic objectives
- Key marketing initiatives
- Budget allocation breakdown (pie chart)
- Expected results and KPIs
- Inspirational marketing quote

**Features**: Green theme, data-driven insights, ROI focus

---

### 💻 Sample 4: Technology Overview (100KB)
**7 slides - AI/ML technology presentation**

Perfect for: Tech talks, training sessions, architecture reviews

Content includes:
- Stunning gradient title slide
- AI/ML fundamentals explained
- Real-world applications across industries
- Complete technology stack with icons
- Performance improvement trends (line chart)
- 6-month implementation roadmap
- Future vision quote

**Features**: Tech-focused design, indigo/purple gradients, clear explanations

---

### 👥 Sample 5: Training Module (80KB)
**7 slides - Leadership development**

Perfect for: Training sessions, workshops, professional development

Content includes:
- Learning objectives overview
- Five leadership styles with descriptions
- Essential communication skills
- Team performance impact (comparison bar chart)
- 30-day action plan
- Key takeaways
- Motivational leadership quote

**Features**: Educational layout, orange theme, practical frameworks

---

## 🎯 How to Use These Samples

### View the Presentations
Open with any of these applications:
- Microsoft PowerPoint
- Google Slides
- LibreOffice Impress
- Apple Keynote

### Recreate or Modify
All samples can be recreated or customized:

```bash
# Install dependencies
pip install -r ../requirements.txt

# Run quick start to generate fresh versions
cd ..
python quick_start.py

# Run comprehensive examples
python examples.py
```

### Learn from the Code
Each sample demonstrates specific features:
- Check `quick_start.py` for the code that creates these
- See `examples.py` for more advanced examples
- Read `README.md` for full documentation

## 📚 What Makes These Special

Unlike Node.js-based PPT tools, these presentations were created with:
- **Pure Python** - No Node.js, Chromium, or Sharp required
- **Native Gradients** - CSS gradients rendered directly, no browser needed
- **Publication-Quality Charts** - Matplotlib integration
- **Cross-Platform** - Works anywhere Python runs
- **Simple Deployment** - Just `pip install`

## 🔧 Technical Details

All presentations use:
- **Aspect Ratio**: 16:9 (10" × 5.625")
- **Format**: Standard .pptx (Office Open XML)
- **Compatibility**: PowerPoint 2007 and later
- **Color Depth**: 24-bit RGB
- **Chart Rendering**: 150 DPI matplotlib images
- **Gradient Quality**: High-resolution PNG backgrounds

## 💡 Next Steps

1. **Download and view** these samples
2. **Run the scripts** to see how they're created
3. **Modify the code** to create your own presentations
4. **Read the docs** in the main README.md
5. **Explore modules** for advanced features

---

**Created with Python PowerPoint Suite** - Professional presentations without Node.js
