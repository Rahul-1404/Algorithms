"""
Create 4 Complete Visually Stunning Presentations
Each with 7 slides and unique modern design
"""

from html2pptx import create_presentation_from_html


def create_creative_agency():
    """Sample 2: Creative Agency Portfolio - Bold & Artistic"""
    print("\n2. Creating Creative Agency Portfolio...")

    slides = [
        # Slide 1: Split Diagonal Design
        """
        <div style="width: 720pt; height: 405pt; background: #0a0a0a;">
            <!-- Diagonal Accent -->
            <div style="position: absolute; left: -100pt; top: -50pt; width: 500pt; height: 500pt;
                       background: linear-gradient(135deg, #ff6b6b 0%, #f06595 100%);
                       transform: rotate(-15deg);"></div>

            <h1 style="position: absolute; left: 60pt; top: 100pt; width: 400pt; color: white;
                       font-size: 68pt; font-family: Arial; font-weight: bold; line-height: 0.95;">
                CREATIVE<br/>STUDIO
            </h1>

            <div style="position: absolute; left: 60pt; top: 250pt; width: 300pt; height: 80pt;
                       background: rgba(255, 255, 255, 0.95);">
                <p style="position: absolute; left: 25pt; top: 20pt; width: 250pt; color: #0a0a0a;
                          font-size: 20pt; font-family: Arial; font-weight: bold;">
                    Bold Ideas. Beautiful Execution.
                </p>
            </div>

            <p style="position: absolute; left: 500pt; top: 340pt; width: 180pt; color: white;
                      font-size: 16pt; font-family: Arial; text-align: right;">
                Portfolio 2024
            </p>
        </div>
        """,

        # Slide 2: Masonry Grid Layout
        """
        <div style="width: 720pt; height: 405pt; background: #f8f9fa;">
            <h1 style="position: absolute; left: 40pt; top: 30pt; width: 640pt; color: #212529;
                       font-size: 44pt; font-family: Arial; font-weight: bold;">
                Our Services
            </h1>

            <!-- Service Box 1 - Tall -->
            <div style="position: absolute; left: 40pt; top: 100pt; width: 150pt; height: 270pt;
                       background: linear-gradient(180deg, #ff6b6b 0%, #f06595 100%);">
                <h2 style="position: absolute; left: 20pt; top: 200pt; width: 110pt; color: white;
                           font-size: 24pt; font-family: Arial; font-weight: bold;">
                    Brand<br/>Identity
                </h2>
            </div>

            <!-- Service Box 2 - Short -->
            <div style="position: absolute; left: 210pt; top: 100pt; width: 150pt; height: 130pt;
                       background: linear-gradient(180deg, #4ecdc4 0%, #44a08d 100%);">
                <h2 style="position: absolute; left: 20pt; top: 60pt; width: 110pt; color: white;
                           font-size: 24pt; font-family: Arial; font-weight: bold;">
                    Web<br/>Design
                </h2>
            </div>

            <!-- Service Box 3 - Medium -->
            <div style="position: absolute; left: 210pt; top: 250pt; width: 150pt; height: 120pt;
                       background: linear-gradient(180deg, #feca57 0%, #ff9f43 100%);">
                <h2 style="position: absolute; left: 20pt; top: 50pt; width: 110pt; color: white;
                           font-size: 24pt; font-family: Arial; font-weight: bold;">
                    Motion<br/>Graphics
                </h2>
            </div>

            <!-- Service Box 4 - Tall -->
            <div style="position: absolute; left: 380pt; top: 100pt; width: 150pt; height: 270pt;
                       background: linear-gradient(180deg, #a29bfe 0%, #6c5ce7 100%);">
                <h2 style="position: absolute; left: 20pt; top: 200pt; width: 110pt; color: white;
                           font-size: 24pt; font-family: Arial; font-weight: bold;">
                    UI/UX<br/>Design
                </h2>
            </div>

            <!-- Service Box 5 - Square -->
            <div style="position: absolute; left: 550pt; top: 100pt; width: 130pt; height: 130pt;
                       background: linear-gradient(180deg, #fd79a8 0%, #e84393 100%);">
                <h2 style="position: absolute; left: 15pt; top: 60pt; width: 100pt; color: white;
                           font-size: 22pt; font-family: Arial; font-weight: bold;">
                    Social<br/>Media
                </h2>
            </div>

            <!-- Service Box 6 - Rectangle -->
            <div style="position: absolute; left: 550pt; top: 250pt; width: 130pt; height: 120pt;
                       background: linear-gradient(180deg, #74b9ff 0%, #0984e3 100%);">
                <h2 style="position: absolute; left: 15pt; top: 50pt; width: 100pt; color: white;
                           font-size: 22pt; font-family: Arial; font-weight: bold;">
                    Print<br/>Design
                </h2>
            </div>
        </div>
        """,

        # Slide 3: Featured Work Showcase
        """
        <div style="width: 720pt; height: 405pt; background: #1a1a1a;">
            <!-- Large Image Placeholder with Gradient Overlay -->
            <div style="position: absolute; left: 0; top: 0; width: 720pt; height: 405pt;
                       background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
            </div>

            <div style="position: absolute; left: 0; top: 0; width: 720pt; height: 405pt;
                       background: linear-gradient(0deg, rgba(0,0,0,0.7) 0%, rgba(0,0,0,0) 50%);"></div>

            <!-- Project Info Card -->
            <div style="position: absolute; left: 50pt; top: 250pt; width: 350pt; height: 130pt;
                       background: white;">
                <div style="position: absolute; left: 0; top: 0; width: 6pt; height: 130pt; background: #ff6b6b;"></div>

                <p style="position: absolute; left: 30pt; top: 20pt; width: 300pt;
                          font-size: 14pt; font-family: Arial; color: #999;">
                    Featured Project
                </p>

                <h2 style="position: absolute; left: 30pt; top: 45pt; width: 300pt;
                           font-size: 32pt; font-family: Arial; font-weight: bold; color: #212529;">
                    Nike Air Max
                </h2>

                <p style="position: absolute; left: 30pt; top: 95pt; width: 300pt;
                          font-size: 13pt; font-family: Arial; color: #666;">
                    Brand campaign • 2024 • Photography
                </p>
            </div>
        </div>
        """,

        # Slide 4: Client Testimonial with Large Quote
        """
        <div style="width: 720pt; height: 405pt; background: #f5f5f5;">
            <!-- Quote Mark Background -->
            <p style="position: absolute; left: 50pt; top: 20pt; width: 200pt; color: rgba(255,107,107,0.15);
                      font-size: 180pt; font-family: Arial; font-weight: bold;">
                "
            </p>

            <!-- Testimonial Text -->
            <p style="position: absolute; left: 100pt; top: 120pt; width: 520pt; color: #212529;
                      font-size: 28pt; font-family: Arial; line-height: 1.4;">
                Working with this team transformed our brand. The results exceeded all expectations.
            </p>

            <!-- Client Info Box -->
            <div style="position: absolute; left: 100pt; top: 280pt; width: 400pt; height: 80pt;
                       background: white; border-left: 4pt solid #ff6b6b;">
                <h3 style="position: absolute; left: 25pt; top: 20pt; width: 350pt;
                           font-size: 20pt; font-family: Arial; font-weight: bold; color: #212529;">
                    Sarah Johnson
                </h3>
                <p style="position: absolute; left: 25pt; top: 50pt; width: 350pt;
                          font-size: 15pt; font-family: Arial; color: #666;">
                    CEO, TechCorp Inc.
                </p>
            </div>
        </div>
        """,

        # Slide 5: Team Grid
        """
        <div style="width: 720pt; height: 405pt; background: #ffffff;">
            <h1 style="position: absolute; left: 50pt; top: 35pt; width: 620pt; color: #212529;
                       font-size: 48pt; font-family: Arial; font-weight: bold;">
                Meet The Team
            </h1>

            <!-- Team Member 1 -->
            <div style="position: absolute; left: 50pt; top: 120pt; width: 140pt; height: 220pt;">
                <div style="position: absolute; left: 0; top: 0; width: 140pt; height: 140pt;
                           background: linear-gradient(135deg, #ff6b6b 0%, #f06595 100%);"></div>
                <h3 style="position: absolute; left: 0; top: 155pt; width: 140pt;
                           font-size: 18pt; font-family: Arial; font-weight: bold; color: #212529;">
                    Alex Chen
                </h3>
                <p style="position: absolute; left: 0; top: 185pt; width: 140pt;
                          font-size: 13pt; font-family: Arial; color: #666;">
                    Creative Director
                </p>
            </div>

            <!-- Team Member 2 -->
            <div style="position: absolute; left: 210pt; top: 120pt; width: 140pt; height: 220pt;">
                <div style="position: absolute; left: 0; top: 0; width: 140pt; height: 140pt;
                           background: linear-gradient(135deg, #4ecdc4 0%, #44a08d 100%);"></div>
                <h3 style="position: absolute; left: 0; top: 155pt; width: 140pt;
                           font-size: 18pt; font-family: Arial; font-weight: bold; color: #212529;">
                    Maria Garcia
                </h3>
                <p style="position: absolute; left: 0; top: 185pt; width: 140pt;
                          font-size: 13pt; font-family: Arial; color: #666;">
                    Art Director
                </p>
            </div>

            <!-- Team Member 3 -->
            <div style="position: absolute; left: 370pt; top: 120pt; width: 140pt; height: 220pt;">
                <div style="position: absolute; left: 0; top: 0; width: 140pt; height: 140pt;
                           background: linear-gradient(135deg, #a29bfe 0%, #6c5ce7 100%);"></div>
                <h3 style="position: absolute; left: 0; top: 155pt; width: 140pt;
                           font-size: 18pt; font-family: Arial; font-weight: bold; color: #212529;">
                    James Wilson
                </h3>
                <p style="position: absolute; left: 0; top: 185pt; width: 140pt;
                          font-size: 13pt; font-family: Arial; color: #666;">
                    Lead Designer
                </p>
            </div>

            <!-- Team Member 4 -->
            <div style="position: absolute; left: 530pt; top: 120pt; width: 140pt; height: 220pt;">
                <div style="position: absolute; left: 0; top: 0; width: 140pt; height: 140pt;
                           background: linear-gradient(135deg, #feca57 0%, #ff9f43 100%);"></div>
                <h3 style="position: absolute; left: 0; top: 155pt; width: 140pt;
                           font-size: 18pt; font-family: Arial; font-weight: bold; color: #212529;">
                    Emma Davis
                </h3>
                <p style="position: absolute; left: 0; top: 185pt; width: 140pt;
                          font-size: 13pt; font-family: Arial; color: #666;">
                    UX Designer
                </p>
            </div>
        </div>
        """,

        # Slide 6: Process Timeline
        """
        <div style="width: 720pt; height: 405pt; background: #0a0a0a;">
            <h1 style="position: absolute; left: 50pt; top: 30pt; width: 620pt; color: white;
                       font-size: 48pt; font-family: Arial; font-weight: bold;">
                Our Process
            </h1>

            <!-- Step 1 -->
            <div style="position: absolute; left: 50pt; top: 110pt; width: 130pt; height: 250pt;">
                <div style="position: absolute; left: 0; top: 0; width: 60pt; height: 60pt;
                           background: #ff6b6b;">
                    <p style="position: absolute; left: 18pt; top: 12pt; color: white;
                              font-size: 32pt; font-family: Arial; font-weight: bold;">
                        1
                    </p>
                </div>
                <h3 style="position: absolute; left: 0; top: 80pt; width: 130pt; color: white;
                           font-size: 20pt; font-family: Arial; font-weight: bold;">
                    Discover
                </h3>
                <p style="position: absolute; left: 0; top: 115pt; width: 130pt; color: #999;
                          font-size: 13pt; font-family: Arial; line-height: 1.4;">
                    Research and understand your brand goals
                </p>
            </div>

            <!-- Step 2 -->
            <div style="position: absolute; left: 200pt; top: 110pt; width: 130pt; height: 250pt;">
                <div style="position: absolute; left: 0; top: 0; width: 60pt; height: 60pt;
                           background: #4ecdc4;">
                    <p style="position: absolute; left: 18pt; top: 12pt; color: white;
                              font-size: 32pt; font-family: Arial; font-weight: bold;">
                        2
                    </p>
                </div>
                <h3 style="position: absolute; left: 0; top: 80pt; width: 130pt; color: white;
                           font-size: 20pt; font-family: Arial; font-weight: bold;">
                    Design
                </h3>
                <p style="position: absolute; left: 0; top: 115pt; width: 130pt; color: #999;
                          font-size: 13pt; font-family: Arial; line-height: 1.4;">
                    Create beautiful concepts and iterations
                </p>
            </div>

            <!-- Step 3 -->
            <div style="position: absolute; left: 350pt; top: 110pt; width: 130pt; height: 250pt;">
                <div style="position: absolute; left: 0; top: 0; width: 60pt; height: 60pt;
                           background: #a29bfe;">
                    <p style="position: absolute; left: 18pt; top: 12pt; color: white;
                              font-size: 32pt; font-family: Arial; font-weight: bold;">
                        3
                    </p>
                </div>
                <h3 style="position: absolute; left: 0; top: 80pt; width: 130pt; color: white;
                           font-size: 20pt; font-family: Arial; font-weight: bold;">
                    Develop
                </h3>
                <p style="position: absolute; left: 0; top: 115pt; width: 130pt; color: #999;
                          font-size: 13pt; font-family: Arial; line-height: 1.4;">
                    Build and refine the final product
                </p>
            </div>

            <!-- Step 4 -->
            <div style="position: absolute; left: 500pt; top: 110pt; width: 130pt; height: 250pt;">
                <div style="position: absolute; left: 0; top: 0; width: 60pt; height: 60pt;
                           background: #feca57;">
                    <p style="position: absolute; left: 18pt; top: 12pt; color: #0a0a0a;
                              font-size: 32pt; font-family: Arial; font-weight: bold;">
                        4
                    </p>
                </div>
                <h3 style="position: absolute; left: 0; top: 80pt; width: 130pt; color: white;
                           font-size: 20pt; font-family: Arial; font-weight: bold;">
                    Deliver
                </h3>
                <p style="position: absolute; left: 0; top: 115pt; width: 130pt; color: #999;
                          font-size: 13pt; font-family: Arial; line-height: 1.4;">
                    Launch and support your success
                </p>
            </div>
        </div>
        """,

        # Slide 7: Contact CTA
        """
        <div style="width: 720pt; height: 405pt; background: linear-gradient(135deg, #ff6b6b 0%, #f06595 100%);">
            <h1 style="position: absolute; left: 60pt; top: 120pt; width: 600pt; color: white;
                       font-size: 64pt; font-family: Arial; font-weight: bold; line-height: 1.1;">
                Let's Create<br/>Something<br/>Amazing
            </h1>

            <div style="position: absolute; left: 60pt; top: 320pt; width: 400pt; height: 60pt;
                       background: white;">
                <p style="position: absolute; left: 30pt; top: 17pt; width: 340pt; color: #0a0a0a;
                          font-size: 20pt; font-family: Arial; font-weight: bold;">
                    hello@creativestudio.com
                </p>
            </div>
        </div>
        """
    ]

    create_presentation_from_html(slides, 'samples/visual_2_creative_agency.pptx', '16:9')
    print("   ✓ Created: visual_2_creative_agency.pptx")


# Create function for Sample 3, 4, 5 following same pattern...

if __name__ == '__main__':
    print("="*70)
    print("Creating 4 More VISUALLY STUNNING Presentations")
    print("="*70)

    create_creative_agency()

    print("\n✓ Created visually stunning presentation!")
