"""
Create 5 VISUALLY STUNNING Sample Presentations
Using advanced HTML/CSS features for beautiful designs
"""

from html2pptx import create_presentation_from_html


def create_modern_startup_pitch():
    """Sample 1: Modern Startup Pitch - Sleek and Visual"""
    print("\n1. Creating Modern Startup Pitch (Visually Stunning)...")

    slides = [
        # Slide 1: Bold Title with Overlay
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
            <!-- Overlay for depth -->
            <div style="position: absolute; left: 0; top: 0; width: 720pt; height: 405pt;
                       background: linear-gradient(45deg, rgba(0,0,0,0.3) 0%, rgba(0,0,0,0) 100%);"></div>

            <!-- Main Title -->
            <h1 style="position: absolute; left: 60pt; top: 140pt; width: 600pt; color: white;
                       font-size: 72pt; font-family: Arial; font-weight: bold; line-height: 1.2;">
                TechFlow AI
            </h1>

            <!-- Subtitle with accent -->
            <div style="position: absolute; left: 60pt; top: 260pt; width: 8pt; height: 60pt;
                       background: #fbbf24;"></div>
            <h2 style="position: absolute; left: 80pt; top: 250pt; width: 500pt; color: white;
                       font-size: 32pt; font-family: Arial;">
                Revolutionizing Business Intelligence<br/>with Real-Time AI
            </h2>

            <!-- Bottom tag -->
            <div style="position: absolute; left: 60pt; top: 350pt; width: 200pt; height: 40pt;
                       background: rgba(255,255,255,0.2); border: 2px solid white;">
                <p style="position: absolute; left: 20pt; top: 8pt; color: white;
                          font-size: 16pt; font-family: Arial; font-weight: bold;">
                    Series A Pitch Deck
                </p>
            </div>
        </div>
        """,

        # Slide 2: Problem Statement with Visual Cards
        """
        <div style="width: 720pt; height: 405pt; background: #f8fafc;">
            <!-- Header Bar -->
            <div style="position: absolute; left: 0; top: 0; width: 720pt; height: 100pt;
                       background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);"></div>

            <h1 style="position: absolute; left: 50pt; top: 25pt; width: 620pt; color: white;
                       font-size: 42pt; font-family: Arial; font-weight: bold;">
                The Problem
            </h1>

            <!-- Problem Card 1 -->
            <div style="position: absolute; left: 50pt; top: 130pt; width: 200pt; height: 220pt;
                       background: white; border: 3px solid #e2e8f0;">
                <div style="position: absolute; left: 0; top: 0; width: 200pt; height: 8pt; background: #ef4444;"></div>

                <h2 style="position: absolute; left: 20pt; top: 30pt; width: 160pt;
                           font-size: 24pt; font-family: Arial; font-weight: bold; color: #1e293b;">
                    Data Silos
                </h2>

                <p style="position: absolute; left: 20pt; top: 80pt; width: 160pt;
                          font-size: 14pt; font-family: Arial; color: #475569; line-height: 1.5;">
                    Companies store data across 15+ different platforms, making insights impossible
                </p>

                <p style="position: absolute; left: 20pt; top: 170pt; width: 160pt;
                          font-size: 32pt; font-family: Arial; font-weight: bold; color: #ef4444;">
                    87%
                </p>
            </div>

            <!-- Problem Card 2 -->
            <div style="position: absolute; left: 260pt; top: 130pt; width: 200pt; height: 220pt;
                       background: white; border: 3px solid #e2e8f0;">
                <div style="position: absolute; left: 0; top: 0; width: 200pt; height: 8pt; background: #f59e0b;"></div>

                <h2 style="position: absolute; left: 20pt; top: 30pt; width: 160pt;
                           font-size: 24pt; font-family: Arial; font-weight: bold; color: #1e293b;">
                    Slow Analysis
                </h2>

                <p style="position: absolute; left: 20pt; top: 80pt; width: 160pt;
                          font-size: 14pt; font-family: Arial; color: #475569; line-height: 1.5;">
                    Traditional BI tools take weeks to generate insights from data
                </p>

                <p style="position: absolute; left: 20pt; top: 170pt; width: 160pt;
                          font-size: 32pt; font-family: Arial; font-weight: bold; color: #f59e0b;">
                    3-4 weeks
                </p>
            </div>

            <!-- Problem Card 3 -->
            <div style="position: absolute; left: 470pt; top: 130pt; width: 200pt; height: 220pt;
                       background: white; border: 3px solid #e2e8f0;">
                <div style="position: absolute; left: 0; top: 0; width: 200pt; height: 8pt; background: #8b5cf6;"></div>

                <h2 style="position: absolute; left: 20pt; top: 30pt; width: 160pt;
                           font-size: 24pt; font-family: Arial; font-weight: bold; color: #1e293b;">
                    High Costs
                </h2>

                <p style="position: absolute; left: 20pt; top: 80pt; width: 160pt;
                          font-size: 14pt; font-family: Arial; color: #475569; line-height: 1.5;">
                    Enterprise BI solutions cost $500K+ annually for mid-size teams
                </p>

                <p style="position: absolute; left: 20pt; top: 170pt; width: 160pt;
                          font-size: 28pt; font-family: Arial; font-weight: bold; color: #8b5cf6;">
                    $500K+
                </p>
            </div>
        </div>
        """,

        # Slide 3: Solution with Visual Showcase
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #1e293b 0%, #334155 100%);">
            <!-- Accent Shape -->
            <div style="position: absolute; left: -50pt; top: -50pt; width: 300pt; height: 300pt;
                       background: radial-gradient(circle, rgba(102, 126, 234, 0.3) 0%, rgba(102, 126, 234, 0) 70%);"></div>

            <h1 style="position: absolute; left: 50pt; top: 40pt; width: 620pt; color: white;
                       font-size: 48pt; font-family: Arial; font-weight: bold;">
                Our Solution: TechFlow AI
            </h1>

            <!-- Feature Box 1 -->
            <div style="position: absolute; left: 50pt; top: 120pt; width: 300pt; height: 100pt;
                       background: rgba(255, 255, 255, 0.1); border: 2px solid rgba(255, 255, 255, 0.2);">
                <div style="position: absolute; left: 20pt; top: 15pt; width: 50pt; height: 50pt;
                           background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                           border: 2px solid white;">
                    <p style="position: absolute; left: 12pt; top: 8pt; color: white;
                              font-size: 28pt; font-family: Arial; font-weight: bold;">
                        AI
                    </p>
                </div>

                <h3 style="position: absolute; left: 85pt; top: 20pt; width: 200pt;
                           font-size: 20pt; font-family: Arial; font-weight: bold; color: white;">
                    Unified Data Layer
                </h3>
                <p style="position: absolute; left: 85pt; top: 52pt; width: 200pt;
                          font-size: 13pt; font-family: Arial; color: #cbd5e1;">
                    Connect all your data sources instantly
                </p>
            </div>

            <!-- Feature Box 2 -->
            <div style="position: absolute; left: 370pt; top: 120pt; width: 300pt; height: 100pt;
                       background: rgba(255, 255, 255, 0.1); border: 2px solid rgba(255, 255, 255, 0.2);">
                <div style="position: absolute; left: 20pt; top: 15pt; width: 50pt; height: 50pt;
                           background: linear-gradient(135deg, #f59e0b 0%, #ef4444 100%);
                           border: 2px solid white;">
                    <p style="position: absolute; left: 8pt; top: 8pt; color: white;
                              font-size: 28pt; font-family: Arial; font-weight: bold;">
                        ⚡
                    </p>
                </div>

                <h3 style="position: absolute; left: 85pt; top: 20pt; width: 200pt;
                           font-size: 20pt; font-family: Arial; font-weight: bold; color: white;">
                    Real-Time Insights
                </h3>
                <p style="position: absolute; left: 85pt; top: 52pt; width: 200pt;
                          font-size: 13pt; font-family: Arial; color: #cbd5e1;">
                    Get answers in seconds, not weeks
                </p>
            </div>

            <!-- Feature Box 3 -->
            <div style="position: absolute; left: 50pt; top: 240pt; width: 300pt; height: 100pt;
                       background: rgba(255, 255, 255, 0.1); border: 2px solid rgba(255, 255, 255, 0.2);">
                <div style="position: absolute; left: 20pt; top: 15pt; width: 50pt; height: 50pt;
                           background: linear-gradient(135deg, #10b981 0%, #059669 100%);
                           border: 2px solid white;">
                    <p style="position: absolute; left: 10pt; top: 8pt; color: white;
                              font-size: 28pt; font-family: Arial; font-weight: bold;">
                        $
                    </p>
                </div>

                <h3 style="position: absolute; left: 85pt; top: 20pt; width: 200pt;
                           font-size: 20pt; font-family: Arial; font-weight: bold; color: white;">
                    90% Cost Reduction
                </h3>
                <p style="position: absolute; left: 85pt; top: 52pt; width: 200pt;
                          font-size: 13pt; font-family: Arial; color: #cbd5e1;">
                    Pay only $50K/year for enterprise features
                </p>
            </div>

            <!-- Feature Box 4 -->
            <div style="position: absolute; left: 370pt; top: 240pt; width: 300pt; height: 100pt;
                       background: rgba(255, 255, 255, 0.1); border: 2px solid rgba(255, 255, 255, 0.2);">
                <div style="position: absolute; left: 20pt; top: 15pt; width: 50pt; height: 50pt;
                           background: linear-gradient(135deg, #ec4899 0%, #8b5cf6 100%);
                           border: 2px solid white;">
                    <p style="position: absolute; left: 10pt; top: 8pt; color: white;
                              font-size: 28pt; font-family: Arial; font-weight: bold;">
                        🚀
                    </p>
                </div>

                <h3 style="position: absolute; left: 85pt; top: 20pt; width: 200pt;
                           font-size: 20pt; font-family: Arial; font-weight: bold; color: white;">
                    Zero Setup Time
                </h3>
                <p style="position: absolute; left: 85pt; top: 52pt; width: 200pt;
                          font-size: 13pt; font-family: Arial; color: #cbd5e1;">
                    Deploy in minutes with our AI assistant
                </p>
            </div>
        </div>
        """,

        # Slide 4: Market Opportunity with Visual Data
        """
        <div style="width: 720pt; height: 405pt; background: #ffffff;">
            <!-- Gradient Header -->
            <div style="position: absolute; left: 0; top: 0; width: 720pt; height: 120pt;
                       background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
                <h1 style="position: absolute; left: 50pt; top: 35pt; width: 620pt; color: white;
                           font-size: 48pt; font-family: Arial; font-weight: bold;">
                    Market Opportunity
                </h1>
            </div>

            <!-- Big Number 1 -->
            <div style="position: absolute; left: 50pt; top: 150pt; width: 200pt; height: 200pt;
                       background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
                <p style="position: absolute; left: 20pt; top: 40pt; width: 160pt; color: white;
                          font-size: 64pt; font-family: Arial; font-weight: bold; text-align: center;">
                    $12B
                </p>
                <p style="position: absolute; left: 20pt; top: 130pt; width: 160pt; color: white;
                          font-size: 18pt; font-family: Arial; text-align: center; line-height: 1.3;">
                    Total Addressable Market
                </p>
            </div>

            <!-- Big Number 2 -->
            <div style="position: absolute; left: 260pt; top: 150pt; width: 200pt; height: 200pt;
                       background: linear-gradient(135deg, #f59e0b 0%, #ef4444 100%);">
                <p style="position: absolute; left: 20pt; top: 40pt; width: 160pt; color: white;
                          font-size: 64pt; font-family: Arial; font-weight: bold; text-align: center;">
                    45%
                </p>
                <p style="position: absolute; left: 20pt; top: 130pt; width: 160pt; color: white;
                          font-size: 18pt; font-family: Arial; text-align: center; line-height: 1.3;">
                    Annual Growth Rate
                </p>
            </div>

            <!-- Big Number 3 -->
            <div style="position: absolute; left: 470pt; top: 150pt; width: 200pt; height: 200pt;
                       background: linear-gradient(135deg, #10b981 0%, #059669 100%);">
                <p style="position: absolute; left: 20pt; top: 40pt; width: 160pt; color: white;
                          font-size: 64pt; font-family: Arial; font-weight: bold; text-align: center;">
                    2.5K
                </p>
                <p style="position: absolute; left: 20pt; top: 130pt; width: 160pt; color: white;
                          font-size: 18pt; font-family: Arial; text-align: center; line-height: 1.3;">
                    Enterprise Customers Ready
                </p>
            </div>
        </div>
        """,

        # Slide 5: Traction with Timeline
        """
        <div style="width: 720pt; height: 405pt; background: #f8fafc;">
            <h1 style="position: absolute; left: 50pt; top: 30pt; width: 620pt; color: #1e293b;
                       font-size: 48pt; font-family: Arial; font-weight: bold;">
                Our Traction
            </h1>

            <!-- Timeline Line -->
            <div style="position: absolute; left: 50pt; top: 140pt; width: 620pt; height: 4pt; background: #cbd5e1;"></div>

            <!-- Milestone 1 -->
            <div style="position: absolute; left: 50pt; top: 125pt; width: 20pt; height: 20pt;
                       background: #667eea; border: 4pt solid white;"></div>
            <div style="position: absolute; left: 20pt; top: 160pt; width: 100pt; height: 150pt;
                       background: white; border: 3px solid #667eea;">
                <p style="position: absolute; left: 10pt; top: 15pt; width: 80pt;
                          font-size: 14pt; font-family: Arial; font-weight: bold; color: #667eea;">
                    Jan 2024
                </p>
                <p style="position: absolute; left: 10pt; top: 45pt; width: 80pt;
                          font-size: 12pt; font-family: Arial; color: #475569; line-height: 1.4;">
                    Founded<br/>Seed funded<br/>$2M raised
                </p>
            </div>

            <!-- Milestone 2 -->
            <div style="position: absolute; left: 215pt; top: 125pt; width: 20pt; height: 20pt;
                       background: #f59e0b; border: 4pt solid white;"></div>
            <div style="position: absolute; left: 185pt; top: 160pt; width: 100pt; height: 150pt;
                       background: white; border: 3px solid #f59e0b;">
                <p style="position: absolute; left: 10pt; top: 15pt; width: 80pt;
                          font-size: 14pt; font-family: Arial; font-weight: bold; color: #f59e0b;">
                    Apr 2024
                </p>
                <p style="position: absolute; left: 10pt; top: 45pt; width: 80pt;
                          font-size: 12pt; font-family: Arial; color: #475569; line-height: 1.4;">
                    Beta launch<br/>50 customers<br/>$100K MRR
                </p>
            </div>

            <!-- Milestone 3 -->
            <div style="position: absolute; left: 380pt; top: 125pt; width: 20pt; height: 20pt;
                       background: #10b981; border: 4pt solid white;"></div>
            <div style="position: absolute; left: 350pt; top: 160pt; width: 100pt; height: 150pt;
                       background: white; border: 3px solid #10b981;">
                <p style="position: absolute; left: 10pt; top: 15pt; width: 80pt;
                          font-size: 14pt; font-family: Arial; font-weight: bold; color: #10b981;">
                    Aug 2024
                </p>
                <p style="position: absolute; left: 10pt; top: 45pt; width: 80pt;
                          font-size: 12pt; font-family: Arial; color: #475569; line-height: 1.4;">
                    Series A<br/>250 customers<br/>$500K MRR
                </p>
            </div>

            <!-- Milestone 4 -->
            <div style="position: absolute; left: 545pt; top: 125pt; width: 20pt; height: 20pt;
                       background: #ec4899; border: 4pt solid white;"></div>
            <div style="position: absolute; left: 515pt; top: 160pt; width: 100pt; height: 150pt;
                       background: white; border: 3px solid #ec4899;">
                <p style="position: absolute; left: 10pt; top: 15pt; width: 80pt;
                          font-size: 14pt; font-family: Arial; font-weight: bold; color: #ec4899;">
                    Dec 2024
                </p>
                <p style="position: absolute; left: 10pt; top: 45pt; width: 80pt;
                          font-size: 12pt; font-family: Arial; color: #475569; line-height: 1.4;">
                    Growth<br/>1000 customers<br/>$2M MRR
                </p>
            </div>
        </div>
        """,

        # Slide 6: Team with Photo Placeholders
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #1e293b 0%, #334155 100%);">
            <h1 style="position: absolute; left: 50pt; top: 30pt; width: 620pt; color: white;
                       font-size: 48pt; font-family: Arial; font-weight: bold;">
                World-Class Team
            </h1>

            <!-- Team Member 1 -->
            <div style="position: absolute; left: 50pt; top: 120pt; width: 150pt; height: 240pt;
                       background: rgba(255, 255, 255, 0.1); border: 2px solid rgba(255, 255, 255, 0.2);">
                <div style="position: absolute; left: 25pt; top: 20pt; width: 100pt; height: 100pt;
                           background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);"></div>
                <h3 style="position: absolute; left: 15pt; top: 135pt; width: 120pt;
                           font-size: 18pt; font-family: Arial; font-weight: bold; color: white; text-align: center;">
                    Sarah Chen
                </h3>
                <p style="position: absolute; left: 15pt; top: 165pt; width: 120pt;
                          font-size: 13pt; font-family: Arial; color: #cbd5e1; text-align: center;">
                    CEO, Founder
                </p>
                <p style="position: absolute; left: 15pt; top: 190pt; width: 120pt;
                          font-size: 11pt; font-family: Arial; color: #94a3b8; text-align: center; line-height: 1.3;">
                    Ex-Google ML<br/>Stanford PhD<br/>10+ years AI
                </p>
            </div>

            <!-- Team Member 2 -->
            <div style="position: absolute; left: 220pt; top: 120pt; width: 150pt; height: 240pt;
                       background: rgba(255, 255, 255, 0.1); border: 2px solid rgba(255, 255, 255, 0.2);">
                <div style="position: absolute; left: 25pt; top: 20pt; width: 100pt; height: 100pt;
                           background: linear-gradient(135deg, #f59e0b 0%, #ef4444 100%);"></div>
                <h3 style="position: absolute; left: 15pt; top: 135pt; width: 120pt;
                           font-size: 18pt; font-family: Arial; font-weight: bold; color: white; text-align: center;">
                    Marcus Johnson
                </h3>
                <p style="position: absolute; left: 15pt; top: 165pt; width: 120pt;
                          font-size: 13pt; font-family: Arial; color: #cbd5e1; text-align: center;">
                    CTO, Co-Founder
                </p>
                <p style="position: absolute; left: 15pt; top: 190pt; width: 120pt;
                          font-size: 11pt; font-family: Arial; color: #94a3b8; text-align: center; line-height: 1.3;">
                    Ex-Amazon<br/>MIT CS<br/>Infra expert
                </p>
            </div>

            <!-- Team Member 3 -->
            <div style="position: absolute; left: 390pt; top: 120pt; width: 150pt; height: 240pt;
                       background: rgba(255, 255, 255, 0.1); border: 2px solid rgba(255, 255, 255, 0.2);">
                <div style="position: absolute; left: 25pt; top: 20pt; width: 100pt; height: 100pt;
                           background: linear-gradient(135deg, #10b981 0%, #059669 100%);"></div>
                <h3 style="position: absolute; left: 15pt; top: 135pt; width: 120pt;
                           font-size: 18pt; font-family: Arial; font-weight: bold; color: white; text-align: center;">
                    Priya Sharma
                </h3>
                <p style="position: absolute; left: 15pt; top: 165pt; width: 120pt;
                          font-size: 13pt; font-family: Arial; color: #cbd5e1; text-align: center;">
                    Head of Product
                </p>
                <p style="position: absolute; left: 15pt; top: 190pt; width: 120pt;
                          font-size: 11pt; font-family: Arial; color: #94a3b8; text-align: center; line-height: 1.3;">
                    Ex-Meta<br/>Berkeley MBA<br/>Product visionary
                </p>
            </div>
        </div>
        """,

        # Slide 7: The Ask - Bold CTA
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
            <!-- Decorative Circle -->
            <div style="position: absolute; left: 500pt; top: -100pt; width: 400pt; height: 400pt;
                       background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, rgba(255,255,255,0) 70%);"></div>

            <h1 style="position: absolute; left: 50pt; top: 80pt; width: 620pt; color: white;
                       font-size: 64pt; font-family: Arial; font-weight: bold;">
                The Ask
            </h1>

            <!-- Amount Box -->
            <div style="position: absolute; left: 50pt; top: 180pt; width: 300pt; height: 140pt;
                       background: rgba(255, 255, 255, 0.95); border: 4px solid #fbbf24;">
                <p style="position: absolute; left: 30pt; top: 25pt; width: 240pt;
                          font-size: 18pt; font-family: Arial; color: #64748b;">
                    Raising
                </p>
                <p style="position: absolute; left: 30pt; top: 55pt; width: 240pt;
                          font-size: 56pt; font-family: Arial; font-weight: bold; color: #1e293b;">
                    $10M
                </p>
            </div>

            <!-- Use of Funds -->
            <div style="position: absolute; left: 370pt; top: 180pt; width: 300pt; height: 140pt;
                       background: rgba(255, 255, 255, 0.15); border: 2px solid rgba(255, 255, 255, 0.3);">
                <p style="position: absolute; left: 20pt; top: 15pt; width: 260pt;
                          font-size: 16pt; font-family: Arial; font-weight: bold; color: white;">
                    Use of Funds:
                </p>
                <p style="position: absolute; left: 20pt; top: 50pt; width: 260pt;
                          font-size: 13pt; font-family: Arial; color: #e2e8f0; line-height: 1.6;">
                    • 40% Engineering & Product<br/>
                    • 30% Sales & Marketing<br/>
                    • 20% Operations & Scale<br/>
                    • 10% Reserve
                </p>
            </div>

            <!-- Contact -->
            <div style="position: absolute; left: 50pt; top: 340pt; width: 620pt; height: 50pt;
                       background: rgba(255, 255, 255, 0.1); border: 2px solid rgba(255, 255, 255, 0.3);">
                <p style="position: absolute; left: 30pt; top: 12pt; width: 560pt;
                          font-size: 18pt; font-family: Arial; color: white;">
                    📧 sarah@techflow.ai  •  📞 (415) 555-0123  •  🌐 techflow.ai
                </p>
            </div>
        </div>
        """
    ]

    create_presentation_from_html(slides, 'samples/visual_1_startup_pitch.pptx', '16:9')
    print("   ✓ Created: visual_1_startup_pitch.pptx")


# Run the creation
if __name__ == '__main__':
    print("="*70)
    print("Creating VISUALLY STUNNING Presentations")
    print("="*70)

    create_modern_startup_pitch()

    print("\n✓ Created visually appealing startup pitch!")
