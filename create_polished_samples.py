"""
Create 5 POLISHED, Visually Stunning Presentations
Perfect text alignment, spacing, and professional layouts
"""

from html2pptx import create_presentation_from_html


def create_startup_pitch_polished():
    """Sample 1: Startup Pitch - Perfectly Aligned"""
    print("\n1. Creating Polished Startup Pitch...")

    slides = [
        # Slide 1: Clean Bold Title
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
            <h1 style="position: absolute; left: 60pt; top: 150pt; width: 600pt;
                       font-size: 68pt; font-family: Arial; font-weight: bold; color: white;
                       text-align: left; line-height: 1.1;">
                TechFlow AI
            </h1>

            <p style="position: absolute; left: 60pt; top: 250pt; width: 500pt;
                      font-size: 24pt; font-family: Arial; color: white;
                      text-align: left; line-height: 1.4;">
                Revolutionizing Enterprise Automation<br/>with Artificial Intelligence
            </p>

            <div style="position: absolute; left: 60pt; top: 330pt; width: 180pt; height: 50pt;
                       background: rgba(255,255,255,0.2);">
                <p style="position: absolute; left: 0pt; top: 13pt; width: 180pt;
                          font-size: 16pt; font-family: Arial; font-weight: bold; color: white;
                          text-align: center;">
                    SERIES A PITCH
                </p>
            </div>
        </div>
        """,

        # Slide 2: Problem - Clean Grid
        """
        <div style="width: 720pt; height: 405pt; background: #ffffff;">
            <h1 style="position: absolute; left: 50pt; top: 40pt; width: 620pt;
                       font-size: 44pt; font-family: Arial; font-weight: bold; color: #1a1a1a;
                       text-align: left;">
                The Problem
            </h1>

            <p style="position: absolute; left: 50pt; top: 100pt; width: 620pt;
                      font-size: 16pt; font-family: Arial; color: #666666;
                      text-align: left; line-height: 1.5;">
                Companies waste billions on manual processes that could be automated
            </p>

            <!-- Card 1 -->
            <div style="position: absolute; left: 50pt; top: 160pt; width: 200pt; height: 210pt;
                       background: #f8f9fa; border-left: 6pt solid #ef4444;">
                <p style="position: absolute; left: 20pt; top: 30pt; width: 160pt;
                          font-size: 48pt; font-family: Arial; font-weight: bold; color: #ef4444;
                          text-align: left;">
                    87%
                </p>
                <p style="position: absolute; left: 20pt; top: 100pt; width: 160pt;
                          font-size: 20pt; font-family: Arial; font-weight: bold; color: #1a1a1a;
                          text-align: left;">
                    Manual Work
                </p>
                <p style="position: absolute; left: 20pt; top: 140pt; width: 160pt;
                          font-size: 14pt; font-family: Arial; color: #666666;
                          text-align: left; line-height: 1.4;">
                    Tasks still done manually despite automation tools
                </p>
            </div>

            <!-- Card 2 -->
            <div style="position: absolute; left: 260pt; top: 160pt; width: 200pt; height: 210pt;
                       background: #f8f9fa; border-left: 6pt solid #f59e0b;">
                <p style="position: absolute; left: 20pt; top: 30pt; width: 160pt;
                          font-size: 42pt; font-family: Arial; font-weight: bold; color: #f59e0b;
                          text-align: left;">
                    $450B
                </p>
                <p style="position: absolute; left: 20pt; top: 100pt; width: 160pt;
                          font-size: 20pt; font-family: Arial; font-weight: bold; color: #1a1a1a;
                          text-align: left;">
                    Wasted Cost
                </p>
                <p style="position: absolute; left: 20pt; top: 140pt; width: 160pt;
                          font-size: 14pt; font-family: Arial; color: #666666;
                          text-align: left; line-height: 1.4;">
                    Annual waste on inefficient processes
                </p>
            </div>

            <!-- Card 3 -->
            <div style="position: absolute; left: 470pt; top: 160pt; width: 200pt; height: 210pt;
                       background: #f8f9fa; border-left: 6pt solid #8b5cf6;">
                <p style="position: absolute; left: 20pt; top: 30pt; width: 160pt;
                          font-size: 48pt; font-family: Arial; font-weight: bold; color: #8b5cf6;
                          text-align: left;">
                    40%
                </p>
                <p style="position: absolute; left: 20pt; top: 100pt; width: 160pt;
                          font-size: 20pt; font-family: Arial; font-weight: bold; color: #1a1a1a;
                          text-align: left;">
                    Time Lost
                </p>
                <p style="position: absolute; left: 20pt; top: 140pt; width: 160pt;
                          font-size: 14pt; font-family: Arial; color: #666666;
                          text-align: left; line-height: 1.4;">
                    Employee time on repetitive tasks
                </p>
            </div>
        </div>
        """,

        # Slide 3: Solution - Dark Clean
        """
        <div style="width: 720pt; height: 405pt; background: #1a1a1a;">
            <h1 style="position: absolute; left: 50pt; top: 40pt; width: 620pt;
                       font-size: 44pt; font-family: Arial; font-weight: bold; color: #ffffff;
                       text-align: left;">
                Our Solution
            </h1>

            <!-- Feature 1 -->
            <div style="position: absolute; left: 50pt; top: 120pt; width: 610pt; height: 70pt;
                       background: #2a2a2a;">
                <p style="position: absolute; left: 25pt; top: 22pt; width: 560pt;
                          font-size: 22pt; font-family: Arial; font-weight: bold; color: #ffffff;
                          text-align: left;">
                    🤖  AI Agent Platform
                </p>
                <p style="position: absolute; left: 80pt; top: 48pt; width: 505pt;
                          font-size: 14pt; font-family: Arial; color: #999999;
                          text-align: left;">
                    Deploy autonomous AI agents that learn and improve continuously
                </p>
            </div>

            <!-- Feature 2 -->
            <div style="position: absolute; left: 50pt; top: 200pt; width: 610pt; height: 70pt;
                       background: #2a2a2a;">
                <p style="position: absolute; left: 25pt; top: 22pt; width: 560pt;
                          font-size: 22pt; font-family: Arial; font-weight: bold; color: #ffffff;
                          text-align: left;">
                    ⚡  Real-Time Learning
                </p>
                <p style="position: absolute; left: 80pt; top: 48pt; width: 505pt;
                          font-size: 14pt; font-family: Arial; color: #999999;
                          text-align: left;">
                    Agents adapt to your business processes in real-time
                </p>
            </div>

            <!-- Feature 3 -->
            <div style="position: absolute; left: 50pt; top: 280pt; width: 295pt; height: 70pt;
                       background: #2a2a2a;">
                <p style="position: absolute; left: 25pt; top: 22pt; width: 245pt;
                          font-size: 22pt; font-family: Arial; font-weight: bold; color: #ffffff;
                          text-align: left;">
                    💰  80% Savings
                </p>
                <p style="position: absolute; left: 25pt; top: 48pt; width: 245pt;
                          font-size: 14pt; font-family: Arial; color: #999999;
                          text-align: left;">
                    Reduce costs dramatically
                </p>
            </div>

            <!-- Feature 4 -->
            <div style="position: absolute; left: 365pt; top: 280pt; width: 295pt; height: 70pt;
                       background: #2a2a2a;">
                <p style="position: absolute; left: 25pt; top: 22pt; width: 245pt;
                          font-size: 22pt; font-family: Arial; font-weight: bold; color: #ffffff;
                          text-align: left;">
                    🚀  5min Setup
                </p>
                <p style="position: absolute; left: 25pt; top: 48pt; width: 245pt;
                          font-size: 14pt; font-family: Arial; color: #999999;
                          text-align: left;">
                    Deploy instantly
                </p>
            </div>
        </div>
        """,

        # Slide 4: Market - Clean Numbers
        """
        <div style="width: 720pt; height: 405pt; background: #ffffff;">
            <h1 style="position: absolute; left: 50pt; top: 40pt; width: 620pt;
                       font-size: 44pt; font-family: Arial; font-weight: bold; color: #1a1a1a;
                       text-align: left;">
                Market Opportunity
            </h1>

            <!-- Stat 1 -->
            <div style="position: absolute; left: 50pt; top: 140pt; width: 200pt; height: 200pt;
                       background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
                <p style="position: absolute; left: 0pt; top: 60pt; width: 200pt;
                          font-size: 56pt; font-family: Arial; font-weight: bold; color: #ffffff;
                          text-align: center;">
                    $50B
                </p>
                <p style="position: absolute; left: 20pt; top: 140pt; width: 160pt;
                          font-size: 14pt; font-family: Arial; color: #ffffff;
                          text-align: center; line-height: 1.3;">
                    Total Addressable<br/>Market
                </p>
            </div>

            <!-- Stat 2 -->
            <div style="position: absolute; left: 260pt; top: 140pt; width: 200pt; height: 200pt;
                       background: linear-gradient(135deg, #fbbf24 0%, #f59e0b 100%);">
                <p style="position: absolute; left: 0pt; top: 60pt; width: 200pt;
                          font-size: 56pt; font-family: Arial; font-weight: bold; color: #ffffff;
                          text-align: center;">
                    65%
                </p>
                <p style="position: absolute; left: 20pt; top: 140pt; width: 160pt;
                          font-size: 14pt; font-family: Arial; color: #ffffff;
                          text-align: center; line-height: 1.3;">
                    Annual Growth<br/>Rate
                </p>
            </div>

            <!-- Stat 3 -->
            <div style="position: absolute; left: 470pt; top: 140pt; width: 200pt; height: 200pt;
                       background: linear-gradient(135deg, #10b981 0%, #059669 100%);">
                <p style="position: absolute; left: 0pt; top: 60pt; width: 200pt;
                          font-size: 56pt; font-family: Arial; font-weight: bold; color: #ffffff;
                          text-align: center;">
                    5,000
                </p>
                <p style="position: absolute; left: 20pt; top: 140pt; width: 160pt;
                          font-size: 14pt; font-family: Arial; color: #ffffff;
                          text-align: center; line-height: 1.3;">
                    Enterprise<br/>Prospects
                </p>
            </div>
        </div>
        """,

        # Slide 5: Traction - Clean Timeline
        """
        <div style="width: 720pt; height: 405pt; background: #f8f9fa;">
            <h1 style="position: absolute; left: 50pt; top: 40pt; width: 620pt;
                       font-size: 44pt; font-family: Arial; font-weight: bold; color: #1a1a1a;
                       text-align: left;">
                Our Traction
            </h1>

            <p style="position: absolute; left: 50pt; top: 95pt; width: 620pt;
                      font-size: 16pt; font-family: Arial; color: #666666;
                      text-align: left;">
                From zero to $3.5M MRR in 11 months
            </p>

            <!-- Timeline -->
            <div style="position: absolute; left: 80pt; top: 160pt; width: 560pt; height: 3pt;
                       background: #d1d5db;"></div>

            <!-- Jan 2024 -->
            <div style="position: absolute; left: 80pt; top: 147pt; width: 20pt; height: 20pt;
                       background: #667eea;"></div>
            <div style="position: absolute; left: 45pt; top: 185pt; width: 100pt;">
                <p style="position: absolute; left: 0pt; top: 0pt; width: 100pt;
                          font-size: 14pt; font-family: Arial; font-weight: bold; color: #1a1a1a;
                          text-align: left;">
                    Jan 2024
                </p>
                <p style="position: absolute; left: 0pt; top: 25pt; width: 100pt;
                          font-size: 12pt; font-family: Arial; color: #666666;
                          text-align: left; line-height: 1.4;">
                    Founded<br/>Raised $3M
                </p>
            </div>

            <!-- Apr 2024 -->
            <div style="position: absolute; left: 220pt; top: 147pt; width: 20pt; height: 20pt;
                       background: #fbbf24;"></div>
            <div style="position: absolute; left: 185pt; top: 185pt; width: 100pt;">
                <p style="position: absolute; left: 0pt; top: 0pt; width: 100pt;
                          font-size: 14pt; font-family: Arial; font-weight: bold; color: #1a1a1a;
                          text-align: left;">
                    Apr 2024
                </p>
                <p style="position: absolute; left: 0pt; top: 25pt; width: 100pt;
                          font-size: 12pt; font-family: Arial; color: #666666;
                          text-align: left; line-height: 1.4;">
                    100 customers<br/>$250K MRR
                </p>
            </div>

            <!-- Aug 2024 -->
            <div style="position: absolute; left: 360pt; top: 147pt; width: 20pt; height: 20pt;
                       background: #10b981;"></div>
            <div style="position: absolute; left: 325pt; top: 185pt; width: 100pt;">
                <p style="position: absolute; left: 0pt; top: 0pt; width: 100pt;
                          font-size: 14pt; font-family: Arial; font-weight: bold; color: #1a1a1a;
                          text-align: left;">
                    Aug 2024
                </p>
                <p style="position: absolute; left: 0pt; top: 25pt; width: 100pt;
                          font-size: 12pt; font-family: Arial; color: #666666;
                          text-align: left; line-height: 1.4;">
                    500 customers<br/>$1M MRR
                </p>
            </div>

            <!-- Nov 2024 -->
            <div style="position: absolute; left: 500pt; top: 147pt; width: 20pt; height: 20pt;
                       background: #ec4899;"></div>
            <div style="position: absolute; left: 465pt; top: 185pt; width: 100pt;">
                <p style="position: absolute; left: 0pt; top: 0pt; width: 100pt;
                          font-size: 14pt; font-family: Arial; font-weight: bold; color: #1a1a1a;
                          text-align: left;">
                    Nov 2024
                </p>
                <p style="position: absolute; left: 0pt; top: 25pt; width: 100pt;
                          font-size: 12pt; font-family: Arial; color: #666666;
                          text-align: left; line-height: 1.4;">
                    1,500 customers<br/>$3.5M MRR
                </p>
            </div>
        </div>
        """,

        # Slide 6: Team
        """
        <div style="width: 720pt; height: 405pt; background: #1a1a1a;">
            <h1 style="position: absolute; left: 50pt; top: 40pt; width: 620pt;
                       font-size: 44pt; font-family: Arial; font-weight: bold; color: #ffffff;
                       text-align: left;">
                Leadership Team
            </h1>

            <!-- Person 1 -->
            <div style="position: absolute; left: 50pt; top: 130pt; width: 200pt;">
                <div style="position: absolute; left: 0pt; top: 0pt; width: 200pt; height: 140pt;
                           background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);"></div>
                <p style="position: absolute; left: 0pt; top: 160pt; width: 200pt;
                          font-size: 18pt; font-family: Arial; font-weight: bold; color: #ffffff;
                          text-align: left;">
                    Dr. Sarah Chen
                </p>
                <p style="position: absolute; left: 0pt; top: 185pt; width: 200pt;
                          font-size: 14pt; font-family: Arial; color: #fbbf24;
                          text-align: left;">
                    CEO & Co-Founder
                </p>
                <p style="position: absolute; left: 0pt; top: 210pt; width: 200pt;
                          font-size: 12pt; font-family: Arial; color: #999999;
                          text-align: left; line-height: 1.3;">
                    Ex-Google DeepMind<br/>Stanford PhD AI
                </p>
            </div>

            <!-- Person 2 -->
            <div style="position: absolute; left: 260pt; top: 130pt; width: 200pt;">
                <div style="position: absolute; left: 0pt; top: 0pt; width: 200pt; height: 140pt;
                           background: linear-gradient(135deg, #fbbf24 0%, #f59e0b 100%);"></div>
                <p style="position: absolute; left: 0pt; top: 160pt; width: 200pt;
                          font-size: 18pt; font-family: Arial; font-weight: bold; color: #ffffff;
                          text-align: left;">
                    Marcus Johnson
                </p>
                <p style="position: absolute; left: 0pt; top: 185pt; width: 200pt;
                          font-size: 14pt; font-family: Arial; color: #fbbf24;
                          text-align: left;">
                    CTO & Co-Founder
                </p>
                <p style="position: absolute; left: 0pt; top: 210pt; width: 200pt;
                          font-size: 12pt; font-family: Arial; color: #999999;
                          text-align: left; line-height: 1.3;">
                    Ex-Amazon AWS<br/>MIT Computer Science
                </p>
            </div>

            <!-- Person 3 -->
            <div style="position: absolute; left: 470pt; top: 130pt; width: 200pt;">
                <div style="position: absolute; left: 0pt; top: 0pt; width: 200pt; height: 140pt;
                           background: linear-gradient(135deg, #10b981 0%, #059669 100%);"></div>
                <p style="position: absolute; left: 0pt; top: 160pt; width: 200pt;
                          font-size: 18pt; font-family: Arial; font-weight: bold; color: #ffffff;
                          text-align: left;">
                    Priya Sharma
                </p>
                <p style="position: absolute; left: 0pt; top: 185pt; width: 200pt;
                          font-size: 14pt; font-family: Arial; color: #fbbf24;
                          text-align: left;">
                    Head of Product
                </p>
                <p style="position: absolute; left: 0pt; top: 210pt; width: 200pt;
                          font-size: 12pt; font-family: Arial; color: #999999;
                          text-align: left; line-height: 1.3;">
                    Ex-Meta AI<br/>Berkeley MBA
                </p>
            </div>
        </div>
        """,

        # Slide 7: The Ask
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
            <h1 style="position: absolute; left: 60pt; top: 80pt; width: 600pt;
                       font-size: 56pt; font-family: Arial; font-weight: bold; color: #ffffff;
                       text-align: left;">
                The Ask
            </h1>

            <!-- Amount -->
            <div style="position: absolute; left: 60pt; top: 180pt; width: 280pt; height: 140pt;
                       background: #ffffff;">
                <p style="position: absolute; left: 0pt; top: 25pt; width: 280pt;
                          font-size: 18pt; font-family: Arial; color: #666666;
                          text-align: center;">
                    Raising Series A
                </p>
                <p style="position: absolute; left: 0pt; top: 55pt; width: 280pt;
                          font-size: 60pt; font-family: Arial; font-weight: bold; color: #1a1a1a;
                          text-align: center;">
                    $15M
                </p>
            </div>

            <!-- Use -->
            <div style="position: absolute; left: 360pt; top: 180pt; width: 300pt; height: 140pt;
                       background: rgba(255,255,255,0.15);">
                <p style="position: absolute; left: 25pt; top: 20pt; width: 250pt;
                          font-size: 16pt; font-family: Arial; font-weight: bold; color: #ffffff;
                          text-align: left;">
                    Use of Funds
                </p>
                <p style="position: absolute; left: 25pt; top: 50pt; width: 250pt;
                          font-size: 13pt; font-family: Arial; color: #ffffff;
                          text-align: left; line-height: 1.5;">
                    45% Engineering & AI R&D<br/>
                    30% Sales & Marketing<br/>
                    15% Operations<br/>
                    10% Reserve
                </p>
            </div>

            <!-- Contact -->
            <p style="position: absolute; left: 60pt; top: 345pt; width: 600pt;
                      font-size: 16pt; font-family: Arial; color: #ffffff;
                      text-align: left;">
                sarah@techflow-ai.com  •  (415) 555-0123  •  techflow-ai.com
            </p>
        </div>
        """
    ]

    create_presentation_from_html(slides, 'samples/polished_startup_pitch.pptx', '16:9')
    print("   ✓ Created: polished_startup_pitch.pptx")


if __name__ == '__main__':
    print("="*70)
    print("Creating POLISHED Visually Stunning Presentation")
    print("Perfect alignment, spacing, and professional layouts")
    print("="*70)

    create_startup_pitch_polished()

    print("\n" + "="*70)
    print("✓ Polished presentation created with perfect alignment!")
    print("="*70)
