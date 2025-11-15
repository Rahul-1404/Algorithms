"""
Advanced Chart Generation for PowerPoint
Pure Python charts with matplotlib for beautiful visualizations
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
from io import BytesIO
from typing import List, Dict, Tuple, Optional
import numpy as np


class ChartGenerator:
    """Generate beautiful charts for PowerPoint"""

    # Professional color schemes
    COLOR_SCHEMES = {
        'professional': [
            (31, 119, 180), (255, 127, 14), (44, 160, 44),
            (214, 39, 40), (148, 103, 189), (140, 86, 75),
            (227, 119, 194), (127, 127, 127), (188, 189, 34), (23, 190, 207)
        ],
        'vibrant': [
            (102, 126, 234), (118, 75, 162), (237, 100, 166),
            (255, 198, 93), (87, 204, 153), (255, 107, 107)
        ],
        'pastel': [
            (179, 226, 221), (253, 205, 172), (203, 153, 201),
            (161, 196, 253), (255, 179, 186), (255, 223, 186)
        ],
        'corporate': [
            (0, 51, 102), (0, 102, 204), (0, 153, 255),
            (102, 178, 255), (204, 229, 255), (128, 128, 128)
        ],
        'modern': [
            (79, 70, 229), (236, 72, 153), (251, 146, 60),
            (34, 197, 94), (14, 165, 233), (168, 85, 247)
        ]
    }

    @staticmethod
    def add_bar_chart(slide, left: float, top: float, width: float, height: float,
                     categories: List[str], data_series: List[Dict],
                     title: str = "", color_scheme: str = 'professional',
                     show_legend: bool = True) -> None:
        """
        Add a bar chart to slide

        Args:
            slide: PowerPoint slide object
            left, top, width, height: Position and size in inches
            categories: List of category labels
            data_series: List of dicts with 'name' and 'values' keys
            title: Chart title
            color_scheme: Color scheme name
            show_legend: Whether to show legend
        """
        chart_data = CategoryChartData()
        chart_data.categories = categories

        for series in data_series:
            chart_data.add_series(series['name'], series['values'])

        x, y, cx, cy = Inches(left), Inches(top), Inches(width), Inches(height)
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
        ).chart

        chart.has_title = bool(title)
        if title:
            chart.chart_title.text_frame.text = title

        chart.has_legend = show_legend
        if show_legend:
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.include_in_layout = False

        # Apply colors
        ChartGenerator._apply_colors(chart, color_scheme)

    @staticmethod
    def add_line_chart(slide, left: float, top: float, width: float, height: float,
                      categories: List[str], data_series: List[Dict],
                      title: str = "", color_scheme: str = 'professional',
                      show_legend: bool = True, smooth: bool = True) -> None:
        """Add a line chart to slide"""
        chart_data = CategoryChartData()
        chart_data.categories = categories

        for series in data_series:
            chart_data.add_series(series['name'], series['values'])

        x, y, cx, cy = Inches(left), Inches(top), Inches(width), Inches(height)
        chart_type = XL_CHART_TYPE.LINE_MARKERS if not smooth else XL_CHART_TYPE.LINE_MARKERS
        chart = slide.shapes.add_chart(
            chart_type, x, y, cx, cy, chart_data
        ).chart

        chart.has_title = bool(title)
        if title:
            chart.chart_title.text_frame.text = title

        chart.has_legend = show_legend
        if show_legend:
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.include_in_layout = False

        ChartGenerator._apply_colors(chart, color_scheme)

    @staticmethod
    def add_pie_chart(slide, left: float, top: float, width: float, height: float,
                     categories: List[str], values: List[float],
                     title: str = "", color_scheme: str = 'professional',
                     show_legend: bool = True) -> None:
        """Add a pie chart to slide"""
        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series('Series 1', values)

        x, y, cx, cy = Inches(left), Inches(top), Inches(width), Inches(height)
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.PIE, x, y, cx, cy, chart_data
        ).chart

        chart.has_title = bool(title)
        if title:
            chart.chart_title.text_frame.text = title

        chart.has_legend = show_legend
        if show_legend:
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.include_in_layout = False

        ChartGenerator._apply_colors(chart, color_scheme)

    @staticmethod
    def add_matplotlib_chart(slide, left: float, top: float, width: float, height: float,
                            plot_func, **kwargs) -> None:
        """
        Add a custom matplotlib chart as an image

        Args:
            slide: PowerPoint slide object
            left, top, width, height: Position and size in inches
            plot_func: Function that takes a matplotlib axis and creates the plot
            **kwargs: Additional arguments passed to plot_func
        """
        # Create figure with high DPI for quality
        dpi = 150
        fig_width = width
        fig_height = height

        # Create figure with proper sizing and margins
        fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=dpi)

        # Adjust subplot to leave room for labels
        fig.subplots_adjust(left=0.1, right=0.95, top=0.92, bottom=0.12)

        # Call custom plot function
        plot_func(ax, **kwargs)

        # Save to BytesIO
        img_stream = BytesIO()
        plt.savefig(img_stream, format='png', dpi=dpi, bbox_inches='tight',
                   pad_inches=0.1, facecolor='white')
        plt.close(fig)
        img_stream.seek(0)

        # Add to slide
        slide.shapes.add_picture(img_stream, Inches(left), Inches(top),
                                Inches(width), Inches(height))

    @staticmethod
    def create_beautiful_bar_chart(ax, categories: List[str], data_series: List[Dict],
                                   title: str = "", color_scheme: str = 'professional'):
        """Create a beautiful bar chart with matplotlib"""
        colors = ChartGenerator.COLOR_SCHEMES.get(color_scheme, ChartGenerator.COLOR_SCHEMES['professional'])

        x = np.arange(len(categories))
        width = 0.8 / len(data_series)

        max_value = 0
        for i, series in enumerate(data_series):
            offset = (i - len(data_series) / 2) * width + width / 2
            color_rgb = [c / 255 for c in colors[i % len(colors)]]

            bars = ax.bar(x + offset, series['values'], width,
                         label=series['name'], color=color_rgb, alpha=0.9,
                         edgecolor='white', linewidth=0.5)

            # Track max value
            max_value = max(max_value, max(series['values']))

            # Add value labels on bars
            for bar in bars:
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width() / 2., height,
                       f'{int(height)}',
                       ha='center', va='bottom', fontsize=8, fontweight='bold')

        # Set y-axis limit to give room for labels
        ax.set_ylim(0, max_value * 1.12)

        ax.set_ylabel('Values', fontweight='bold', fontsize=11)
        if title:
            ax.set_title(title, fontweight='bold', fontsize=13, pad=10)
        ax.set_xticks(x)
        ax.set_xticklabels(categories, fontsize=9)
        ax.legend(frameon=False, loc='upper left', fontsize=9)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(axis='y', alpha=0.3, linestyle='--', linewidth=0.5)
        ax.set_axisbelow(True)

        # Ensure tick labels are readable
        ax.tick_params(axis='both', labelsize=9)

    @staticmethod
    def create_beautiful_line_chart(ax, categories: List[str], data_series: List[Dict],
                                   title: str = "", color_scheme: str = 'professional'):
        """Create a beautiful line chart with matplotlib"""
        colors = ChartGenerator.COLOR_SCHEMES.get(color_scheme, ChartGenerator.COLOR_SCHEMES['professional'])

        x = np.arange(len(categories))
        max_value = 0

        for i, series in enumerate(data_series):
            color_rgb = [c / 255 for c in colors[i % len(colors)]]

            ax.plot(x, series['values'], marker='o', linewidth=2.5,
                   markersize=7, label=series['name'], color=color_rgb,
                   markeredgecolor='white', markeredgewidth=1.5)

            max_value = max(max_value, max(series['values']))

            # Add value labels
            for j, value in enumerate(series['values']):
                ax.text(x[j], value, f'{int(value)}',
                       ha='center', va='bottom', fontsize=8, fontweight='bold')

        # Set y-axis limit to give room for labels
        ax.set_ylim(0, max_value * 1.12)

        ax.set_ylabel('Values', fontweight='bold', fontsize=11)
        if title:
            ax.set_title(title, fontweight='bold', fontsize=13, pad=10)
        ax.set_xticks(x)
        ax.set_xticklabels(categories, fontsize=9)
        ax.legend(frameon=False, loc='upper left', fontsize=9)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(alpha=0.3, linestyle='--', linewidth=0.5)
        ax.set_axisbelow(True)
        ax.tick_params(axis='both', labelsize=9)

    @staticmethod
    def create_beautiful_pie_chart(ax, categories: List[str], values: List[float],
                                   title: str = "", color_scheme: str = 'professional'):
        """Create a beautiful pie chart with matplotlib"""
        colors = ChartGenerator.COLOR_SCHEMES.get(color_scheme, ChartGenerator.COLOR_SCHEMES['professional'])
        colors_normalized = [[c / 255 for c in color] for color in colors[:len(categories)]]

        wedges, texts, autotexts = ax.pie(values, labels=categories, autopct='%1.1f%%',
                                          colors=colors_normalized, startangle=90,
                                          textprops={'fontsize': 10, 'weight': 'bold'})

        # Make percentage text white
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontsize(11)

        ax.set_title(title, fontweight='bold', fontsize=14, pad=20)

    @staticmethod
    def create_area_chart(ax, categories: List[str], data_series: List[Dict],
                         title: str = "", color_scheme: str = 'professional'):
        """Create a beautiful stacked area chart"""
        colors = ChartGenerator.COLOR_SCHEMES.get(color_scheme, ChartGenerator.COLOR_SCHEMES['professional'])

        x = np.arange(len(categories))
        y_stack = None

        for i, series in enumerate(data_series):
            color_rgb = [c / 255 for c in colors[i % len(colors)]]
            values = np.array(series['values'])

            if y_stack is None:
                ax.fill_between(x, 0, values, label=series['name'],
                               color=color_rgb, alpha=0.7)
                y_stack = values
            else:
                ax.fill_between(x, y_stack, y_stack + values,
                               label=series['name'], color=color_rgb, alpha=0.7)
                y_stack += values

        ax.set_xlabel('Categories', fontweight='bold')
        ax.set_ylabel('Values', fontweight='bold')
        ax.set_title(title, fontweight='bold', fontsize=14, pad=20)
        ax.set_xticks(x)
        ax.set_xticklabels(categories)
        ax.legend(frameon=False, loc='upper left')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        ax.set_axisbelow(True)

    @staticmethod
    def create_scatter_plot(ax, data_points: List[Tuple[float, float]], labels: List[str] = None,
                          title: str = "", color_scheme: str = 'professional'):
        """Create a beautiful scatter plot"""
        colors = ChartGenerator.COLOR_SCHEMES.get(color_scheme, ChartGenerator.COLOR_SCHEMES['professional'])

        if labels:
            unique_labels = list(set(labels))
            for i, label in enumerate(unique_labels):
                indices = [j for j, l in enumerate(labels) if l == label]
                points = [data_points[j] for j in indices]
                x_vals = [p[0] for p in points]
                y_vals = [p[1] for p in points]

                color_rgb = [c / 255 for c in colors[i % len(colors)]]
                ax.scatter(x_vals, y_vals, s=100, alpha=0.7,
                          color=color_rgb, label=label, edgecolors='white', linewidth=1.5)

            ax.legend(frameon=False)
        else:
            x_vals = [p[0] for p in data_points]
            y_vals = [p[1] for p in data_points]
            color_rgb = [c / 255 for c in colors[0]]
            ax.scatter(x_vals, y_vals, s=100, alpha=0.7,
                      color=color_rgb, edgecolors='white', linewidth=1.5)

        ax.set_xlabel('X Values', fontweight='bold')
        ax.set_ylabel('Y Values', fontweight='bold')
        ax.set_title(title, fontweight='bold', fontsize=14, pad=20)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(alpha=0.3, linestyle='--')
        ax.set_axisbelow(True)

    @staticmethod
    def _apply_colors(chart, color_scheme: str):
        """Apply color scheme to native PowerPoint chart"""
        colors = ChartGenerator.COLOR_SCHEMES.get(color_scheme, ChartGenerator.COLOR_SCHEMES['professional'])

        try:
            for i, series in enumerate(chart.series):
                color = colors[i % len(colors)]
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = RGBColor(*color)
        except Exception as e:
            print(f"Could not apply colors: {e}")


if __name__ == '__main__':
    from pptx import Presentation

    # Demo
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Add matplotlib bar chart
    categories = ['Q1', 'Q2', 'Q3', 'Q4']
    data_series = [
        {'name': 'Sales', 'values': [100, 120, 140, 160]},
        {'name': 'Revenue', 'values': [90, 110, 130, 150]}
    ]

    ChartGenerator.add_matplotlib_chart(
        slide, 0.5, 0.5, 4.5, 3.5,
        ChartGenerator.create_beautiful_bar_chart,
        categories=categories,
        data_series=data_series,
        title='Quarterly Performance',
        color_scheme='modern'
    )

    prs.save('chart_demo.pptx')
    print("Chart demo created: chart_demo.pptx")
