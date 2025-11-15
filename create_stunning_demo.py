"""
Create ONE Complete Visually Stunning Presentation
7 slides with rich, beautiful design
"""

from html2pptx import create_presentation_from_html


def create_complete_startup_deck():
    """Complete Startup Pitch Deck - Visually Stunning"""
    print("\nCreating Complete Visually Stunning Startup Deck...")

    slides = [
        # Slide 1: Bold Title with Geometric Shapes
        """
        <div style="width: 720pt; height: 405pt; background: #0f172a;">
            <!-- Large Circle Background -->
            <div style="position: absolute; left: 400pt; top: -150pt; width: 500pt; height: 500pt;
                       background: radial-gradient(circle, rgba(99, 102, 241, 0.4) 0%, rgba(139, 92, 246, 0) 70%);"></div>

            <!-- Accent Rectangle -->
            <div style="position: absolute; left: -30pt; top: 200pt; width: 250pt; height: 250pt;
                       background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);"></div>

            <!-- Main Title -->
            <h1 style="position: absolute; left: 80pt; top: 120pt; width: 560pt; color: white;
                       font-size: 72pt; font-family: Arial; font-weight: bold; line-height: 1;">
                NOVA AI
            </h1>

            <div style="position: absolute; left: 80pt; top: 210pt; width: 5pt; height: 60pt; background: #fbbf24;"></div>

            <h2 style="position: absolute; left: 100pt; top: 215pt; width: 500pt; color: #cbd5e1;
                       font-size: 28pt; font-family: Arial;">
                Revolutionizing Enterprise<br/>Automation with AI
            </h2>

            <!-- Bottom Badge -->
            <div style="position: absolute; left: 80pt; top: 320pt; width: 220pt; height: 55pt;
                       background: rgba(251, 191, 36, 0.15); border: 2pt solid #fbbf24;">
                <p style="position: absolute; left: 25pt; top: 14pt; color: #fbbf24;
                          font-size: 18pt; font-family: Arial; font-weight: bold;">
                    SERIES A PITCH DECK
                </p>
            </div>
        </div>
        """,

        # Slide 2: Problem - Visual Cards
        """
        <div style="width: 720pt; height: 405pt; background: #f8fafc;">
            <!-- Header -->
            <div style="position: absolute; left: 0; top: 0; width: 10pt; height: 405pt;
                       background: linear-gradient(180deg, #6366f1 0%, #8b5cf6 100%);"></div>

            <h1 style="position: absolute; left: 50pt; top: 40pt; width: 620pt; color: #0f172a;
                       font-size: 48pt; font-family: Arial; font-weight: bold;">
                The Problem
            </h1>

            <p style="position: absolute; left: 50pt; top: 100pt; width: 550pt; color: #64748b;
                      font-size: 18pt; font-family: Arial; line-height: 1.5;">
                Enterprises waste $450B annually on inefficient processes
            </p>

            <!-- Problem Card 1 -->
            <div style="position: absolute; left: 50pt; top: 160pt; width: 190pt; height: 200pt;
                       background: white; border: 3pt solid #ef4444;">
                <div style="position: absolute; left: 0; top: 0; width: 190pt; height: 10pt; background: #ef4444;"></div>

                <div style="position: absolute; left: 20pt; top: 30pt; width: 60pt; height: 60pt;
                           background: linear-gradient(135deg, #fecaca 0%, #fca5a5 100%);">
                    <p style="position: absolute; left: 15pt; top: 10pt; color: #dc2626;
                              font-size: 36pt; font-family: Arial; font-weight: bold;">
                        87%
                    </p>
                </div>

                <h3 style="position: absolute; left: 20pt; top: 110pt; width: 150pt;
                           font-size: 20pt; font-family: Arial; font-weight: bold; color: #0f172a;">
                    Manual Work
                </h3>

                <p style="position: absolute; left: 20pt; top: 145pt; width: 150pt;
                          font-size: 13pt; font-family: Arial; color: #64748b; line-height: 1.4;">
                    Tasks still done manually despite having automation tools
                </p>
            </div>

            <!-- Problem Card 2 -->
            <div style="position: absolute; left: 265pt; top: 160pt; width: 190pt; height: 200pt;
                       background: white; border: 3pt solid #f59e0b;">
                <div style="position: absolute; left: 0; top: 0; width: 190pt; height: 10pt; background: #f59e0b;"></div>

                <div style="position: absolute; left: 20pt; top: 30pt; width: 60pt; height: 60pt;
                           background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%);">
                    <p style="position: absolute; left: 8pt; top: 10pt; color: #d97706;
                              font-size: 32pt; font-family: Arial; font-weight: bold;">
                        $450B
                    </p>
                </div>

                <h3 style="position: absolute; left: 20pt; top: 110pt; width: 150pt;
                           font-size: 20pt; font-family: Arial; font-weight: bold; color: #0f172a;">
                    Wasted Cost
                </h3>

                <p style="position: absolute; left: 20pt; top: 145pt; width: 150pt;
                          font-size: 13pt; font-family: Arial; color: #64748b; line-height: 1.4;">
                    Annual waste on inefficient business processes globally
                </p>
            </div>

            <!-- Problem Card 3 -->
            <div style="position: absolute; left: 480pt; top: 160pt; width: 190pt; height: 200pt;
                       background: white; border: 3pt solid #8b5cf6;">
                <div style="position: absolute; left: 0; top: 0; width: 190pt; height: 10pt; background: #8b5cf6;"></div>

                <div style="position: absolute; left: 20pt; top: 30pt; width: 60pt; height: 60pt;
                           background: linear-gradient(135deg, #ddd6fe 0%, #c4b5fd 100%);">
                    <p style="position: absolute; left: 12pt; top: 10pt; color: #7c3aed;
                              font-size: 36pt; font-family: Arial; font-weight: bold;">
                        40%
                    </p>
                </div>

                <h3 style="position: absolute; left: 20pt; top: 110pt; width: 150pt;
                           font-size: 20pt; font-family: Arial; font-weight: bold; color: #0f172a;">
                    Time Lost
                </h3>

                <p style="position: absolute; left: 20pt; top: 145pt; width: 150pt;
                          font-size: 13pt; font-family: Arial; color: #64748b; line-height: 1.4;">
                    Employee time spent on repetitive tasks
                </p>
            </div>
        </div>
        """,

        # Slide 3: Solution - Dark with Glowing Cards
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);">
            <!-- Glowing Accent -->
            <div style="position: absolute; left: 500pt; top: -100pt; width: 400pt; height: 400pt;
                       background: radial-gradient(circle, rgba(99, 102, 241, 0.2) 0%, rgba(99, 102, 241, 0) 70%);"></div>

            <h1 style="position: absolute; left: 50pt; top: 40pt; width: 620pt; color: white;
                       font-size: 52pt; font-family: Arial; font-weight: bold;">
                Our Solution
            </h1>

            <!-- Solution Feature 1 -->
            <div style="position: absolute; left: 50pt; top: 130pt; width: 300pt; height: 110pt;
                       background: rgba(99, 102, 241, 0.1); border: 2pt solid rgba(99, 102, 241, 0.3);">
                <div style="position: absolute; left: 20pt; top: 15pt; width: 60pt; height: 60pt;
                           background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);">
                    <p style="position: absolute; left: 18pt; top: 12pt; color: white;
                              font-size: 32pt; font-family: Arial;">
                        🤖
                    </p>
                </div>

                <h3 style="position: absolute; left: 95pt; top: 20pt; width: 190pt; color: white;
                           font-size: 22pt; font-family: Arial; font-weight: bold;">
                    AI Agent Platform
                </h3>
                <p style="position: absolute; left: 95pt; top: 55pt; width: 190pt; color: #cbd5e1;
                          font-size: 14pt; font-family: Arial; line-height: 1.3;">
                    Deploy autonomous AI agents in minutes
                </p>
            </div>

            <!-- Solution Feature 2 -->
            <div style="position: absolute; left: 370pt; top: 130pt; width: 300pt; height: 110pt;
                       background: rgba(251, 191, 36, 0.1); border: 2pt solid rgba(251, 191, 36, 0.3);">
                <div style="position: absolute; left: 20pt; top: 15pt; width: 60pt; height: 60pt;
                           background: linear-gradient(135deg, #fbbf24 0%, #f59e0b 100%);">
                    <p style="position: absolute; left: 18pt; top: 12pt; color: #0f172a;
                              font-size: 32pt; font-family: Arial;">
                        ⚡
                    </p>
                </div>

                <h3 style="position: absolute; left: 95pt; top: 20pt; width: 190pt; color: white;
                           font-size: 22pt; font-family: Arial; font-weight: bold;">
                    Real-Time Learning
                </h3>
                <p style="position: absolute; left: 95pt; top: 55pt; width: 190pt; color: #cbd5e1;
                          font-size: 14pt; font-family: Arial; line-height: 1.3;">
                    Agents learn and improve continuously
                </p>
            </div>

            <!-- Solution Feature 3 -->
            <div style="position: absolute; left: 50pt; top: 260pt; width: 300pt; height: 110pt;
                       background: rgba(16, 185, 129, 0.1); border: 2pt solid rgba(16, 185, 129, 0.3);">
                <div style="position: absolute; left: 20pt; top: 15pt; width: 60pt; height: 60pt;
                           background: linear-gradient(135deg, #10b981 0%, #059669 100%);">
                    <p style="position: absolute; left: 18pt; top: 12pt; color: white;
                              font-size: 32pt; font-family: Arial;">
                        💰
                    </p>
                </div>

                <h3 style="position: absolute; left: 95pt; top: 20pt; width: 190pt; color: white;
                           font-size: 22pt; font-family: Arial; font-weight: bold;">
                    80% Cost Savings
                </h3>
                <p style="position: absolute; left: 95pt; top: 55pt; width: 190pt; color: #cbd5e1;
                          font-size: 14pt; font-family: Arial; line-height: 1.3;">
                    Reduce operational costs dramatically
                </p>
            </div>

            <!-- Solution Feature 4 -->
            <div style="position: absolute; left: 370pt; top: 260pt; width: 300pt; height: 110pt;
                       background: rgba(236, 72, 153, 0.1); border: 2pt solid rgba(236, 72, 153, 0.3);">
                <div style="position: absolute; left: 20pt; top: 15pt; width: 60pt; height: 60pt;
                           background: linear-gradient(135deg, #ec4899 0%, #db2777 100%);">
                    <p style="position: absolute; left: 18pt; top: 12pt; color: white;
                              font-size: 32pt; font-family: Arial;">
                        🚀
                    </p>
                </div>

                <h3 style="position: absolute; left: 95pt; top: 20pt; width: 190pt; color: white;
                           font-size: 22pt; font-family: Arial; font-weight: bold;">
                    Zero Setup
                </h3>
                <p style="position: absolute; left: 95pt; top: 55pt; width: 190pt; color: #cbd5e1;
                          font-size: 14pt; font-family: Arial; line-height: 1.3;">
                    Deploy in under 5 minutes
                </p>
            </div>
        </div>
        """,

        # Slide 4: Market - Big Numbers
        """
        <div style="width: 720pt; height: 405pt; background: white;">
            <!-- Top Gradient Bar -->
            <div style="position: absolute; left: 0; top: 0; width: 720pt; height: 120pt;
                       background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);">
                <h1 style="position: absolute; left: 50pt; top: 35pt; width: 620pt; color: white;
                           font-size: 52pt; font-family: Arial; font-weight: bold;">
                    Market Opportunity
                </h1>
            </div>

            <!-- Big Number 1 -->
            <div style="position: absolute; left: 50pt; top: 160pt; width: 190pt; height: 210pt;
                       background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);">
                <p style="position: absolute; left: 20pt; top: 50pt; width: 150pt; color: white;
                          font-size: 72pt; font-family: Arial; font-weight: bold; text-align: center;">
                    $50B
                </p>
                <div style="position: absolute; left: 20pt; top: 140pt; width: 150pt; height: 3pt; background: rgba(255,255,255,0.3);"></div>
                <p style="position: absolute; left: 20pt; top: 155pt; width: 150pt; color: white;
                          font-size: 16pt; font-family: Arial; text-align: center; line-height: 1.3;">
                    Total Addressable Market
                </p>
            </div>

            <!-- Big Number 2 -->
            <div style="position: absolute; left: 265pt; top: 160pt; width: 190pt; height: 210pt;
                       background: linear-gradient(135deg, #fbbf24 0%, #f59e0b 100%);">
                <p style="position: absolute; left: 20pt; top: 50pt; width: 150pt; color: white;
                          font-size: 72pt; font-family: Arial; font-weight: bold; text-align: center;">
                    65%
                </p>
                <div style="position: absolute; left: 20pt; top: 140pt; width: 150pt; height: 3pt; background: rgba(255,255,255,0.3);"></div>
                <p style="position: absolute; left: 20pt; top: 155pt; width: 150pt; color: white;
                          font-size: 16pt; font-family: Arial; text-align: center; line-height: 1.3;">
                    Annual Growth Rate
                </p>
            </div>

            <!-- Big Number 3 -->
            <div style="position: absolute; left: 480pt; top: 160pt; width: 190pt; height: 210pt;
                       background: linear-gradient(135deg, #10b981 0%, #059669 100%);">
                <p style="position: absolute; left: 20pt; top: 50pt; width: 150pt; color: white;
                          font-size: 72pt; font-family: Arial; font-weight: bold; text-align: center;">
                    5K
                </p>
                <div style="position: absolute; left: 20pt; top: 140pt; width: 150pt; height: 3pt; background: rgba(255,255,255,0.3);"></div>
                <p style="position: absolute; left: 20pt; top: 155pt; width: 150pt; color: white;
                          font-size: 16pt; font-family: Arial; text-align: center; line-height: 1.3;">
                    Enterprise Customers Ready
                </p>
            </div>
        </div>
        """,

        # Slide 5: Traction - Timeline
        """
        <div style="width: 720pt; height: 405pt; background: #f8fafc;">
            <h1 style="position: absolute; left: 50pt; top: 35pt; width: 620pt; color: #0f172a;
                       font-size: 52pt; font-family: Arial; font-weight: bold;">
                Our Traction
            </h1>

            <!-- Timeline Line -->
            <div style="position: absolute; left: 70pt; top: 150pt; width: 580pt; height: 4pt; background: #cbd5e1;"></div>

            <!-- Milestone 1 -->
            <div style="position: absolute; left: 70pt; top: 133pt; width: 24pt; height: 24pt;
                       background: #6366f1; border: 5pt solid white;"></div>
            <div style="position: absolute; left: 40pt; top: 180pt; width: 100pt; height: 160pt;
                       background: white; border: 3pt solid #6366f1;">
                <p style="position: absolute; left: 15pt; top: 20pt; width: 70pt;
                          font-size: 15pt; font-family: Arial; font-weight: bold; color: #6366f1;">
                    Jan 2024
                </p>
                <div style="position: absolute; left: 15pt; top: 50pt; width: 70pt; height: 2pt; background: #e2e8f0;"></div>
                <p style="position: absolute; left: 15pt; top: 65pt; width: 70pt;
                          font-size: 13pt; font-family: Arial; color: #64748b; line-height: 1.4;">
                    Founded<br/>Raised $3M seed
                </p>
            </div>

            <!-- Milestone 2 -->
            <div style="position: absolute; left: 210pt; top: 133pt; width: 24pt; height: 24pt;
                       background: #fbbf24; border: 5pt solid white;"></div>
            <div style="position: absolute; left: 180pt; top: 180pt; width: 100pt; height: 160pt;
                       background: white; border: 3pt solid #fbbf24;">
                <p style="position: absolute; left: 15pt; top: 20pt; width: 70pt;
                          font-size: 15pt; font-family: Arial; font-weight: bold; color: #f59e0b;">
                    Apr 2024
                </p>
                <div style="position: absolute; left: 15pt; top: 50pt; width: 70pt; height: 2pt; background: #e2e8f0;"></div>
                <p style="position: absolute; left: 15pt; top: 65pt; width: 70pt;
                          font-size: 13pt; font-family: Arial; color: #64748b; line-height: 1.4;">
                    100 customers<br/>$250K MRR
                </p>
            </div>

            <!-- Milestone 3 -->
            <div style="position: absolute; left: 350pt; top: 133pt; width: 24pt; height: 24pt;
                       background: #10b981; border: 5pt solid white;"></div>
            <div style="position: absolute; left: 320pt; top: 180pt; width: 100pt; height: 160pt;
                       background: white; border: 3pt solid #10b981;">
                <p style="position: absolute; left: 15pt; top: 20pt; width: 70pt;
                          font-size: 15pt; font-family: Arial; font-weight: bold; color: #059669;">
                    Aug 2024
                </p>
                <div style="position: absolute; left: 15pt; top: 50pt; width: 70pt; height: 2pt; background: #e2e8f0;"></div>
                <p style="position: absolute; left: 15pt; top: 65pt; width: 70pt;
                          font-size: 13pt; font-family: Arial; color: #64748b; line-height: 1.4;">
                    500 customers<br/>$1M MRR
                </p>
            </div>

            <!-- Milestone 4 -->
            <div style="position: absolute; left: 490pt; top: 133pt; width: 24pt; height: 24pt;
                       background: #ec4899; border: 5pt solid white;"></div>
            <div style="position: absolute; left: 460pt; top: 180pt; width: 100pt; height: 160pt;
                       background: white; border: 3pt solid #ec4899;">
                <p style="position: absolute; left: 15pt; top: 20pt; width: 70pt;
                          font-size: 15pt; font-family: Arial; font-weight: bold; color: #db2777;">
                    Nov 2024
                </p>
                <div style="position: absolute; left: 15pt; top: 50pt; width: 70pt; height: 2pt; background: #e2e8f0;"></div>
                <p style="position: absolute; left: 15pt; top: 65pt; width: 70pt;
                          font-size: 13pt; font-family: Arial; color: #64748b; line-height: 1.4;">
                    1500 customers<br/>$3.5M MRR
                </p>
            </div>

            <!-- Future Milestone -->
            <div style="position: absolute; left: 630pt; top: 133pt; width: 24pt; height: 24pt;
                       background: #94a3b8; border: 5pt solid white;"></div>
            <div style="position: absolute; left: 600pt; top: 180pt; width: 100pt; height: 160pt;
                       background: white; border: 3pt solid #cbd5e1;">
                <p style="position: absolute; left: 15pt; top: 20pt; width: 70pt;
                          font-size: 15pt; font-family: Arial; font-weight: bold; color: #475569;">
                    Q2 2025
                </p>
                <div style="position: absolute; left: 15pt; top: 50pt; width: 70pt; height: 2pt; background: #e2e8f0;"></div>
                <p style="position: absolute; left: 15pt; top: 65pt; width: 70pt;
                          font-size: 13pt; font-family: Arial; color: #64748b; line-height: 1.4;">
                    5K customers<br/>$12M MRR
                </p>
            </div>
        </div>
        """,

        # Slide 6: Team
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);">
            <h1 style="position: absolute; left: 50pt; top: 35pt; width: 620pt; color: white;
                       font-size: 52pt; font-family: Arial; font-weight: bold;">
                World-Class Team
            </h1>

            <!-- Team Member 1 -->
            <div style="position: absolute; left: 50pt; top: 120pt; width: 150pt; height: 240pt;">
                <div style="position: absolute; left: 0; top: 0; width: 150pt; height: 150pt;
                           background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);
                           border: 4pt solid rgba(255, 255, 255, 0.1);"></div>

                <h3 style="position: absolute; left: 0; top: 170pt; width: 150pt; color: white;
                           font-size: 19pt; font-family: Arial; font-weight: bold; text-align: center;">
                    Dr. Sarah Chen
                </h3>
                <p style="position: absolute; left: 0; top: 200pt; width: 150pt; color: #fbbf24;
                          font-size: 14pt; font-family: Arial; text-align: center;">
                    CEO & Co-Founder
                </p>
                <p style="position: absolute; left: 5pt; top: 223pt; width: 140pt; color: #94a3b8;
                          font-size: 11pt; font-family: Arial; text-align: center; line-height: 1.2;">
                    Ex-Google DeepMind<br/>Stanford PhD AI
                </p>
            </div>

            <!-- Team Member 2 -->
            <div style="position: absolute; left: 220pt; top: 120pt; width: 150pt; height: 240pt;">
                <div style="position: absolute; left: 0; top: 0; width: 150pt; height: 150pt;
                           background: linear-gradient(135deg, #fbbf24 0%, #f59e0b 100%);
                           border: 4pt solid rgba(255, 255, 255, 0.1);"></div>

                <h3 style="position: absolute; left: 0; top: 170pt; width: 150pt; color: white;
                           font-size: 19pt; font-family: Arial; font-weight: bold; text-align: center;">
                    Marcus Johnson
                </h3>
                <p style="position: absolute; left: 0; top: 200pt; width: 150pt; color: #fbbf24;
                          font-size: 14pt; font-family: Arial; text-align: center;">
                    CTO & Co-Founder
                </p>
                <p style="position: absolute; left: 5pt; top: 223pt; width: 140pt; color: #94a3b8;
                          font-size: 11pt; font-family: Arial; text-align: center; line-height: 1.2;">
                    Ex-Amazon AWS<br/>MIT Computer Science
                </p>
            </div>

            <!-- Team Member 3 -->
            <div style="position: absolute; left: 390pt; top: 120pt; width: 150pt; height: 240pt;">
                <div style="position: absolute; left: 0; top: 0; width: 150pt; height: 150pt;
                           background: linear-gradient(135deg, #10b981 0%, #059669 100%);
                           border: 4pt solid rgba(255, 255, 255, 0.1);"></div>

                <h3 style="position: absolute; left: 0; top: 170pt; width: 150pt; color: white;
                           font-size: 19pt; font-family: Arial; font-weight: bold; text-align: center;">
                    Priya Sharma
                </h3>
                <p style="position: absolute; left: 0; top: 200pt; width: 150pt; color: #fbbf24;
                          font-size: 14pt; font-family: Arial; text-align: center;">
                    Head of Product
                </p>
                <p style="position: absolute; left: 5pt; top: 223pt; width: 140pt; color: #94a3b8;
                          font-size: 11pt; font-family: Arial; text-align: center; line-height: 1.2;">
                    Ex-Meta AI<br/>Berkeley MBA
                </p>
            </div>
        </div>
        """,

        # Slide 7: The Ask - Bold CTA
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);">
            <!-- Decorative Elements -->
            <div style="position: absolute; left: 500pt; top: -150pt; width: 500pt; height: 500pt;
                       background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, rgba(255,255,255,0) 70%);"></div>

            <h1 style="position: absolute; left: 60pt; top: 70pt; width: 600pt; color: white;
                       font-size: 68pt; font-family: Arial; font-weight: bold;">
                The Ask
            </h1>

            <!-- Amount Box -->
            <div style="position: absolute; left: 60pt; top: 180pt; width: 320pt; height: 150pt;
                       background: white; border-left: 10pt solid #fbbf24;">
                <p style="position: absolute; left: 30pt; top: 25pt; width: 270pt;
                          font-size: 20pt; font-family: Arial; color: #64748b;">
                    Raising Series A
                </p>
                <p style="position: absolute; left: 30pt; top: 60pt; width: 270pt;
                          font-size: 68pt; font-family: Arial; font-weight: bold; color: #0f172a;">
                    $15M
                </p>
            </div>

            <!-- Use of Funds -->
            <div style="position: absolute; left: 400pt; top: 180pt; width: 260pt; height: 150pt;
                       background: rgba(255, 255, 255, 0.15); border: 2pt solid rgba(255, 255, 255, 0.3);">
                <p style="position: absolute; left: 25pt; top: 20pt; width: 210pt;
                          font-size: 18pt; font-family: Arial; font-weight: bold; color: white;">
                    Use of Funds
                </p>
                <p style="position: absolute; left: 25pt; top: 55pt; width: 210pt;
                          font-size: 14pt; font-family: Arial; color: #e2e8f0; line-height: 1.6;">
                    • 45% Engineering & AI R&D<br/>
                    • 30% Sales & Marketing<br/>
                    • 15% Operations<br/>
                    • 10% Reserve
                </p>
            </div>

            <!-- Contact Bar -->
            <div style="position: absolute; left: 60pt; top: 345pt; width: 600pt; height: 45pt;
                       background: rgba(255, 255, 255, 0.1); border: 1pt solid rgba(255, 255, 255, 0.2);">
                <p style="position: absolute; left: 25pt; top: 11pt; width: 550pt; color: white;
                          font-size: 17pt; font-family: Arial;">
                    📧 sarah@nova-ai.com  •  📞 (415) 555-0199  •  🌐 nova-ai.com
                </p>
            </div>
        </div>
        """
    ]

    create_presentation_from_html(slides, 'samples/visual_startup_deck.pptx', '16:9')
    print("   ✓ Created: visual_startup_deck.pptx (7 slides)")


if __name__ == '__main__':
    print("="*70)
    print("Creating COMPLETE Visually Stunning Presentation")
    print("="*70)

    create_complete_startup_deck()

    print("\n" + "="*70)
    print("✓ Complete visually stunning 7-slide deck created!")
    print("="*70)
