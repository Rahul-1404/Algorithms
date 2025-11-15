#!/usr/bin/env python3
"""
Create 4 additional polished, visually stunning presentations with perfect text alignment.
Each presentation has 7 slides with professional layouts and consistent spacing.
"""

from html2pptx import create_presentation_from_html


def create_creative_agency():
    """Creative Agency Portfolio - Vibrant coral/teal theme"""
    slides = [
        # Slide 1: Title
        """
        <div style="width: 720pt; height: 405pt;
                    background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);">
            <h1 style="position: absolute; left: 60pt; top: 140pt; width: 600pt;
                       font-size: 72pt; font-family: Arial; font-weight: bold; color: white;
                       text-align: left; line-height: 1.1;">
                Creative Studio
            </h1>
            <p style="position: absolute; left: 60pt; top: 250pt; width: 600pt;
                      font-size: 28pt; font-family: Arial; color: rgba(255,255,255,0.95);
                      text-align: left; line-height: 1.3;">
                Transforming Ideas Into Visual Masterpieces
            </p>
        </div>
        """,

        # Slide 2: Our Expertise
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <div style="position: absolute; left: 60pt; top: 40pt; width: 600pt; height: 60pt;
                        background: linear-gradient(90deg, #f093fb 0%, #f5576c 100%); border-radius: 8pt;">
                <h2 style="position: absolute; left: 30pt; top: 16pt;
                           font-size: 36pt; font-family: Arial; font-weight: bold; color: white;
                           text-align: left;">
                    Our Expertise
                </h2>
            </div>

            <!-- Expertise Cards -->
            <div style="position: absolute; left: 60pt; top: 130pt; width: 280pt; height: 100pt;
                        background-color: #fff5f7; border-left: 6pt solid #f5576c; padding: 20pt;">
                <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #f5576c;
                           text-align: left; margin: 0; padding: 0;">
                    Brand Identity
                </h3>
                <p style="font-size: 14pt; font-family: Arial; color: #333333;
                          text-align: left; margin-top: 8pt; line-height: 1.4;">
                    Creating memorable brands that stand out and connect with audiences
                </p>
            </div>

            <div style="position: absolute; left: 380pt; top: 130pt; width: 280pt; height: 100pt;
                        background-color: #fef3ff; border-left: 6pt solid #f093fb; padding: 20pt;">
                <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #f093fb;
                           text-align: left; margin: 0; padding: 0;">
                    Digital Design
                </h3>
                <p style="font-size: 14pt; font-family: Arial; color: #333333;
                          text-align: left; margin-top: 8pt; line-height: 1.4;">
                    Modern web and app interfaces that deliver exceptional user experiences
                </p>
            </div>

            <div style="position: absolute; left: 60pt; top: 260pt; width: 280pt; height: 100pt;
                        background-color: #fff5f7; border-left: 6pt solid #f5576c; padding: 20pt;">
                <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #f5576c;
                           text-align: left; margin: 0; padding: 0;">
                    Motion Graphics
                </h3>
                <p style="font-size: 14pt; font-family: Arial; color: #333333;
                          text-align: left; margin-top: 8pt; line-height: 1.4;">
                    Engaging animations and videos that bring stories to life
                </p>
            </div>

            <div style="position: absolute; left: 380pt; top: 260pt; width: 280pt; height: 100pt;
                        background-color: #fef3ff; border-left: 6pt solid #f093fb; padding: 20pt;">
                <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #f093fb;
                           text-align: left; margin: 0; padding: 0;">
                    Photography
                </h3>
                <p style="font-size: 14pt; font-family: Arial; color: #333333;
                          text-align: left; margin-top: 8pt; line-height: 1.4;">
                    Professional photography that captures your brand's essence
                </p>
            </div>
        </div>
        """,

        # Slide 3: Portfolio Highlights
        """
        <div style="width: 720pt; height: 405pt; background-color: #fafafa;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #f5576c;
                       text-align: left;">
                Portfolio Highlights
            </h2>

            <!-- Project Grid -->
            <div style="position: absolute; left: 60pt; top: 110pt; width: 195pt; height: 120pt;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 8pt;">
                <p style="position: absolute; left: 20pt; bottom: 15pt;
                          font-size: 18pt; font-family: Arial; font-weight: bold; color: white;
                          text-align: left;">
                    Tech Startup<br/>Branding
                </p>
            </div>

            <div style="position: absolute; left: 265pt; top: 110pt; width: 195pt; height: 120pt;
                        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); border-radius: 8pt;">
                <p style="position: absolute; left: 20pt; bottom: 15pt;
                          font-size: 18pt; font-family: Arial; font-weight: bold; color: white;
                          text-align: left;">
                    E-commerce<br/>Platform
                </p>
            </div>

            <div style="position: absolute; left: 470pt; top: 110pt; width: 195pt; height: 120pt;
                        background: linear-gradient(135deg, #fa709a 0%, #fee140 100%); border-radius: 8pt;">
                <p style="position: absolute; left: 20pt; bottom: 15pt;
                          font-size: 18pt; font-family: Arial; font-weight: bold; color: white;
                          text-align: left;">
                    Restaurant<br/>Identity
                </p>
            </div>

            <div style="position: absolute; left: 60pt; top: 245pt; width: 195pt; height: 120pt;
                        background: linear-gradient(135deg, #30cfd0 0%, #330867 100%); border-radius: 8pt;">
                <p style="position: absolute; left: 20pt; bottom: 15pt;
                          font-size: 18pt; font-family: Arial; font-weight: bold; color: white;
                          text-align: left;">
                    Fashion<br/>Campaign
                </p>
            </div>

            <div style="position: absolute; left: 265pt; top: 245pt; width: 195pt; height: 120pt;
                        background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%); border-radius: 8pt;">
                <p style="position: absolute; left: 20pt; bottom: 15pt;
                          font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                          text-align: left;">
                    Mobile App<br/>UI/UX
                </p>
            </div>

            <div style="position: absolute; left: 470pt; top: 245pt; width: 195pt; height: 120pt;
                        background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%); border-radius: 8pt;">
                <p style="position: absolute; left: 20pt; bottom: 15pt;
                          font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                          text-align: left;">
                    Corporate<br/>Video
                </p>
            </div>
        </div>
        """,

        # Slide 4: Our Process
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #f5576c;
                       text-align: left;">
                Our Creative Process
            </h2>

            <!-- Process Steps -->
            <div style="position: absolute; left: 80pt; top: 130pt; width: 560pt;">
                <!-- Step 1 -->
                <div style="position: relative; margin-bottom: 25pt;">
                    <div style="float: left; width: 50pt; height: 50pt; border-radius: 25pt;
                                background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
                                text-align: center; line-height: 50pt;">
                        <span style="font-size: 24pt; font-family: Arial; font-weight: bold; color: white;">1</span>
                    </div>
                    <div style="margin-left: 70pt; padding-top: 5pt;">
                        <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #333333;
                                   text-align: left; margin: 0;">
                            Discovery & Research
                        </h3>
                        <p style="font-size: 14pt; font-family: Arial; color: #666666;
                                  text-align: left; margin-top: 5pt; line-height: 1.3;">
                            Understanding your brand, audience, and objectives
                        </p>
                    </div>
                </div>

                <!-- Step 2 -->
                <div style="position: relative; margin-bottom: 25pt;">
                    <div style="float: left; width: 50pt; height: 50pt; border-radius: 25pt;
                                background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
                                text-align: center; line-height: 50pt;">
                        <span style="font-size: 24pt; font-family: Arial; font-weight: bold; color: white;">2</span>
                    </div>
                    <div style="margin-left: 70pt; padding-top: 5pt;">
                        <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #333333;
                                   text-align: left; margin: 0;">
                            Concept Development
                        </h3>
                        <p style="font-size: 14pt; font-family: Arial; color: #666666;
                                  text-align: left; margin-top: 5pt; line-height: 1.3;">
                            Creating multiple creative directions and mood boards
                        </p>
                    </div>
                </div>

                <!-- Step 3 -->
                <div style="position: relative; margin-bottom: 25pt;">
                    <div style="float: left; width: 50pt; height: 50pt; border-radius: 25pt;
                                background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
                                text-align: center; line-height: 50pt;">
                        <span style="font-size: 24pt; font-family: Arial; font-weight: bold; color: white;">3</span>
                    </div>
                    <div style="margin-left: 70pt; padding-top: 5pt;">
                        <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #333333;
                                   text-align: left; margin: 0;">
                            Design & Refinement
                        </h3>
                        <p style="font-size: 14pt; font-family: Arial; color: #666666;
                                  text-align: left; margin-top: 5pt; line-height: 1.3;">
                            Iterative design process with client feedback
                        </p>
                    </div>
                </div>

                <!-- Step 4 -->
                <div style="position: relative;">
                    <div style="float: left; width: 50pt; height: 50pt; border-radius: 25pt;
                                background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
                                text-align: center; line-height: 50pt;">
                        <span style="font-size: 24pt; font-family: Arial; font-weight: bold; color: white;">4</span>
                    </div>
                    <div style="margin-left: 70pt; padding-top: 5pt;">
                        <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #333333;
                                   text-align: left; margin: 0;">
                            Delivery & Support
                        </h3>
                        <p style="font-size: 14pt; font-family: Arial; color: #666666;
                                  text-align: left; margin-top: 5pt; line-height: 1.3;">
                            Final assets and ongoing brand support
                        </p>
                    </div>
                </div>
            </div>
        </div>
        """,

        # Slide 5: Client Success
        """
        <div style="width: 720pt; height: 405pt; background-color: #fafafa;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #f5576c;
                       text-align: left;">
                Client Success Stories
            </h2>

            <!-- Stats Grid -->
            <div style="position: absolute; left: 60pt; top: 130pt; width: 150pt; height: 150pt;
                        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
                        border-radius: 8pt; text-align: center; padding-top: 35pt;">
                <h1 style="font-size: 54pt; font-family: Arial; font-weight: bold; color: white;
                           margin: 0; line-height: 1;">
                    150+
                </h1>
                <p style="font-size: 16pt; font-family: Arial; color: rgba(255,255,255,0.9);
                          margin-top: 10pt;">
                    Projects<br/>Completed
                </p>
            </div>

            <div style="position: absolute; left: 230pt; top: 130pt; width: 150pt; height: 150pt;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        border-radius: 8pt; text-align: center; padding-top: 35pt;">
                <h1 style="font-size: 54pt; font-family: Arial; font-weight: bold; color: white;
                           margin: 0; line-height: 1;">
                    98%
                </h1>
                <p style="font-size: 16pt; font-family: Arial; color: rgba(255,255,255,0.9);
                          margin-top: 10pt;">
                    Client<br/>Satisfaction
                </p>
            </div>

            <div style="position: absolute; left: 400pt; top: 130pt; width: 150pt; height: 150pt;
                        background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
                        border-radius: 8pt; text-align: center; padding-top: 35pt;">
                <h1 style="font-size: 54pt; font-family: Arial; font-weight: bold; color: white;
                           margin: 0; line-height: 1;">
                    12
                </h1>
                <p style="font-size: 16pt; font-family: Arial; color: rgba(255,255,255,0.9);
                          margin-top: 10pt;">
                    Industry<br/>Awards
                </p>
            </div>

            <!-- Testimonial -->
            <div style="position: absolute; left: 60pt; top: 310pt; width: 600pt; height: 60pt;
                        background-color: white; border-left: 6pt solid #f5576c; padding: 15pt;">
                <p style="font-size: 15pt; font-family: Arial; font-style: italic; color: #333333;
                          text-align: left; line-height: 1.4; margin: 0;">
                    "Creative Studio transformed our brand identity beyond our expectations. Their attention to detail and creative vision is unmatched."
                </p>
            </div>
        </div>
        """,

        # Slide 6: Why Choose Us
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #f5576c;
                       text-align: left;">
                Why Creative Studio?
            </h2>

            <!-- Benefits List -->
            <div style="position: absolute; left: 60pt; top: 120pt; width: 300pt;">
                <div style="margin-bottom: 30pt;">
                    <h3 style="font-size: 22pt; font-family: Arial; font-weight: bold; color: #f093fb;
                               text-align: left; margin: 0;">
                        ✦ Award-Winning Team
                    </h3>
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Industry-recognized designers with 10+ years experience
                    </p>
                </div>

                <div style="margin-bottom: 30pt;">
                    <h3 style="font-size: 22pt; font-family: Arial; font-weight: bold; color: #f093fb;
                               text-align: left; margin: 0;">
                        ✦ Strategic Approach
                    </h3>
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Design backed by research and data-driven insights
                    </p>
                </div>

                <div>
                    <h3 style="font-size: 22pt; font-family: Arial; font-weight: bold; color: #f093fb;
                               text-align: left; margin: 0;">
                        ✦ Full-Service Studio
                    </h3>
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        From concept to delivery, all under one roof
                    </p>
                </div>
            </div>

            <div style="position: absolute; left: 380pt; top: 120pt; width: 300pt;">
                <div style="margin-bottom: 30pt;">
                    <h3 style="font-size: 22pt; font-family: Arial; font-weight: bold; color: #f5576c;
                               text-align: left; margin: 0;">
                        ✦ Fast Turnaround
                    </h3>
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Quick delivery without compromising on quality
                    </p>
                </div>

                <div style="margin-bottom: 30pt;">
                    <h3 style="font-size: 22pt; font-family: Arial; font-weight: bold; color: #f5576c;
                               text-align: left; margin: 0;">
                        ✦ Flexible Pricing
                    </h3>
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Transparent pricing tailored to your budget
                    </p>
                </div>

                <div>
                    <h3 style="font-size: 22pt; font-family: Arial; font-weight: bold; color: #f5576c;
                               text-align: left; margin: 0;">
                        ✦ Ongoing Support
                    </h3>
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Continued partnership after project completion
                    </p>
                </div>
            </div>
        </div>
        """,

        # Slide 7: Let's Create Together
        """
        <div style="width: 720pt; height: 405pt;
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
            <h1 style="position: absolute; left: 60pt; top: 120pt; width: 600pt;
                       font-size: 58pt; font-family: Arial; font-weight: bold; color: white;
                       text-align: left; line-height: 1.2;">
                Let's Create<br/>Something Amazing
            </h1>

            <div style="position: absolute; left: 60pt; top: 280pt; width: 600pt;">
                <p style="font-size: 22pt; font-family: Arial; color: rgba(255,255,255,0.95);
                          text-align: left; margin: 0; line-height: 1.5;">
                    📧 hello@creativestudio.com<br/>
                    📱 +1 (555) 123-4567<br/>
                    🌐 www.creativestudio.com
                </p>
            </div>
        </div>
        """
    ]

    create_presentation_from_html(slides, 'samples/polished_creative_agency.pptx', '16:9')
    print("✓ Created: polished_creative_agency.pptx")


def create_modern_saas():
    """Modern SaaS Platform - Clean blue/indigo theme"""
    slides = [
        # Slide 1: Title
        """
        <div style="width: 720pt; height: 405pt;
                    background: linear-gradient(135deg, #4f46e5 0%, #0ea5e9 100%);">
            <h1 style="position: absolute; left: 60pt; top: 130pt; width: 600pt;
                       font-size: 68pt; font-family: Arial; font-weight: bold; color: white;
                       text-align: left; line-height: 1.1;">
                CloudFlow
            </h1>
            <p style="position: absolute; left: 60pt; top: 240pt; width: 600pt;
                      font-size: 30pt; font-family: Arial; color: rgba(255,255,255,0.95);
                      text-align: left; line-height: 1.3;">
                The All-in-One Platform for Modern Teams
            </p>
            <div style="position: absolute; left: 60pt; top: 320pt; width: 180pt; height: 50pt;
                        background-color: white; border-radius: 6pt; text-align: center; line-height: 50pt;">
                <span style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #4f46e5;">
                    Get Started
                </span>
            </div>
        </div>
        """,

        # Slide 2: The Problem
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #4f46e5;
                       text-align: left;">
                The Modern Work Challenge
            </h2>

            <!-- Problem Cards -->
            <div style="position: absolute; left: 60pt; top: 130pt; width: 600pt;">
                <div style="margin-bottom: 25pt; padding: 20pt; background-color: #fef2f2;
                            border-left: 6pt solid #ef4444;">
                    <h3 style="font-size: 22pt; font-family: Arial; font-weight: bold; color: #ef4444;
                               text-align: left; margin: 0;">
                        Fragmented Tools
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Teams juggle 10+ different applications daily, wasting time switching contexts and losing productivity
                    </p>
                </div>

                <div style="margin-bottom: 25pt; padding: 20pt; background-color: #fef2f2;
                            border-left: 6pt solid #ef4444;">
                    <h3 style="font-size: 22pt; font-family: Arial; font-weight: bold; color: #ef4444;
                               text-align: left; margin: 0;">
                        Data Silos
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Critical information scattered across platforms creates confusion and delays decision-making
                    </p>
                </div>

                <div style="padding: 20pt; background-color: #fef2f2;
                            border-left: 6pt solid #ef4444;">
                    <h3 style="font-size: 22pt; font-family: Arial; font-weight: bold; color: #ef4444;
                               text-align: left; margin: 0;">
                        Rising Costs
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Multiple subscriptions drain budgets while IT struggles to manage licensing and access
                    </p>
                </div>
            </div>
        </div>
        """,

        # Slide 3: The Solution
        """
        <div style="width: 720pt; height: 405pt; background-color: #fafafa;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #4f46e5;
                       text-align: left;">
                One Platform, Infinite Possibilities
            </h2>

            <!-- Feature Grid -->
            <div style="position: absolute; left: 60pt; top: 130pt; width: 185pt; height: 110pt;
                        background-color: white; border-radius: 8pt; padding: 20pt; box-shadow: 0 2pt 8pt rgba(0,0,0,0.1);">
                <div style="width: 40pt; height: 40pt; background-color: #ede9fe; border-radius: 6pt;
                            text-align: center; line-height: 40pt; margin-bottom: 12pt;">
                    <span style="font-size: 24pt;">💬</span>
                </div>
                <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                           text-align: left; margin: 0;">
                    Team Chat
                </h3>
                <p style="font-size: 13pt; font-family: Arial; color: #666666;
                          text-align: left; margin-top: 6pt; line-height: 1.3;">
                    Real-time messaging and collaboration
                </p>
            </div>

            <div style="position: absolute; left: 265pt; top: 130pt; width: 185pt; height: 110pt;
                        background-color: white; border-radius: 8pt; padding: 20pt; box-shadow: 0 2pt 8pt rgba(0,0,0,0.1);">
                <div style="width: 40pt; height: 40pt; background-color: #dbeafe; border-radius: 6pt;
                            text-align: center; line-height: 40pt; margin-bottom: 12pt;">
                    <span style="font-size: 24pt;">📋</span>
                </div>
                <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                           text-align: left; margin: 0;">
                    Task Management
                </h3>
                <p style="font-size: 13pt; font-family: Arial; color: #666666;
                          text-align: left; margin-top: 6pt; line-height: 1.3;">
                    Organize work with boards and timelines
                </p>
            </div>

            <div style="position: absolute; left: 470pt; top: 130pt; width: 185pt; height: 110pt;
                        background-color: white; border-radius: 8pt; padding: 20pt; box-shadow: 0 2pt 8pt rgba(0,0,0,0.1);">
                <div style="width: 40pt; height: 40pt; background-color: #ede9fe; border-radius: 6pt;
                            text-align: center; line-height: 40pt; margin-bottom: 12pt;">
                    <span style="font-size: 24pt;">📊</span>
                </div>
                <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                           text-align: left; margin: 0;">
                    Analytics
                </h3>
                <p style="font-size: 13pt; font-family: Arial; color: #666666;
                          text-align: left; margin-top: 6pt; line-height: 1.3;">
                    Track performance with insights
                </p>
            </div>

            <div style="position: absolute; left: 60pt; top: 265pt; width: 185pt; height: 110pt;
                        background-color: white; border-radius: 8pt; padding: 20pt; box-shadow: 0 2pt 8pt rgba(0,0,0,0.1);">
                <div style="width: 40pt; height: 40pt; background-color: #dbeafe; border-radius: 6pt;
                            text-align: center; line-height: 40pt; margin-bottom: 12pt;">
                    <span style="font-size: 24pt;">📁</span>
                </div>
                <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                           text-align: left; margin: 0;">
                    File Storage
                </h3>
                <p style="font-size: 13pt; font-family: Arial; color: #666666;
                          text-align: left; margin-top: 6pt; line-height: 1.3;">
                    Secure cloud storage for all files
                </p>
            </div>

            <div style="position: absolute; left: 265pt; top: 265pt; width: 185pt; height: 110pt;
                        background-color: white; border-radius: 8pt; padding: 20pt; box-shadow: 0 2pt 8pt rgba(0,0,0,0.1);">
                <div style="width: 40pt; height: 40pt; background-color: #ede9fe; border-radius: 6pt;
                            text-align: center; line-height: 40pt; margin-bottom: 12pt;">
                    <span style="font-size: 24pt;">🎥</span>
                </div>
                <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                           text-align: left; margin: 0;">
                    Video Calls
                </h3>
                <p style="font-size: 13pt; font-family: Arial; color: #666666;
                          text-align: left; margin-top: 6pt; line-height: 1.3;">
                    HD video meetings built-in
                </p>
            </div>

            <div style="position: absolute; left: 470pt; top: 265pt; width: 185pt; height: 110pt;
                        background-color: white; border-radius: 8pt; padding: 20pt; box-shadow: 0 2pt 8pt rgba(0,0,0,0.1);">
                <div style="width: 40pt; height: 40pt; background-color: #dbeafe; border-radius: 6pt;
                            text-align: center; line-height: 40pt; margin-bottom: 12pt;">
                    <span style="font-size: 24pt;">🔗</span>
                </div>
                <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                           text-align: left; margin: 0;">
                    Integrations
                </h3>
                <p style="font-size: 13pt; font-family: Arial; color: #666666;
                          text-align: left; margin-top: 6pt; line-height: 1.3;">
                    Connect all your favorite tools
                </p>
            </div>
        </div>
        """,

        # Slide 4: Key Benefits
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #4f46e5;
                       text-align: left;">
                Transform Your Team's Workflow
            </h2>

            <!-- Benefits with Icons -->
            <div style="position: absolute; left: 60pt; top: 130pt; width: 280pt;">
                <div style="margin-bottom: 35pt;">
                    <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #0ea5e9;
                               text-align: left; margin: 0;">
                        ⚡ 40% Faster Collaboration
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Eliminate tool-switching and streamline communication
                    </p>
                </div>

                <div style="margin-bottom: 35pt;">
                    <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #0ea5e9;
                               text-align: left; margin: 0;">
                        💰 60% Cost Reduction
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Replace 10+ tools with one affordable platform
                    </p>
                </div>

                <div>
                    <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #0ea5e9;
                               text-align: left; margin: 0;">
                        🔒 Enterprise Security
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        SOC 2 compliant with end-to-end encryption
                    </p>
                </div>
            </div>

            <div style="position: absolute; left: 380pt; top: 130pt; width: 280pt;">
                <div style="margin-bottom: 35pt;">
                    <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #4f46e5;
                               text-align: left; margin: 0;">
                        📈 Improved Productivity
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        AI-powered insights help teams work smarter
                    </p>
                </div>

                <div style="margin-bottom: 35pt;">
                    <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #4f46e5;
                               text-align: left; margin: 0;">
                        🌍 Work From Anywhere
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Cloud-based platform accessible on any device
                    </p>
                </div>

                <div>
                    <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #4f46e5;
                               text-align: left; margin: 0;">
                        ⏱️ 5-Minute Setup
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Get your team onboarded in minutes, not weeks
                    </p>
                </div>
            </div>
        </div>
        """,

        # Slide 5: Customer Success
        """
        <div style="width: 720pt; height: 405pt; background-color: #fafafa;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #4f46e5;
                       text-align: left;">
                Trusted by 50,000+ Teams
            </h2>

            <!-- Stats -->
            <div style="position: absolute; left: 90pt; top: 130pt; width: 540pt; height: 80pt;
                        background: linear-gradient(90deg, #4f46e5 0%, #0ea5e9 100%);
                        border-radius: 8pt; padding: 20pt;">
                <div style="float: left; width: 160pt; text-align: center;">
                    <h1 style="font-size: 44pt; font-family: Arial; font-weight: bold; color: white; margin: 0;">
                        10M+
                    </h1>
                    <p style="font-size: 14pt; font-family: Arial; color: rgba(255,255,255,0.9); margin-top: 5pt;">
                        Active Users
                    </p>
                </div>
                <div style="float: left; width: 160pt; text-align: center;">
                    <h1 style="font-size: 44pt; font-family: Arial; font-weight: bold; color: white; margin: 0;">
                        99.9%
                    </h1>
                    <p style="font-size: 14pt; font-family: Arial; color: rgba(255,255,255,0.9); margin-top: 5pt;">
                        Uptime SLA
                    </p>
                </div>
                <div style="float: left; width: 160pt; text-align: center;">
                    <h1 style="font-size: 44pt; font-family: Arial; font-weight: bold; color: white; margin: 0;">
                        4.9★
                    </h1>
                    <p style="font-size: 14pt; font-family: Arial; color: rgba(255,255,255,0.9); margin-top: 5pt;">
                        User Rating
                    </p>
                </div>
            </div>

            <!-- Testimonials -->
            <div style="position: absolute; left: 60pt; top: 250pt; width: 600pt;">
                <div style="margin-bottom: 20pt; padding: 20pt; background-color: white;
                            border-radius: 8pt; box-shadow: 0 2pt 6pt rgba(0,0,0,0.1);">
                    <p style="font-size: 15pt; font-family: Arial; font-style: italic; color: #333333;
                              text-align: left; line-height: 1.5; margin: 0;">
                        "CloudFlow reduced our software costs by 65% while improving team productivity. Best decision we made this year."
                    </p>
                    <p style="font-size: 13pt; font-family: Arial; color: #0ea5e9;
                              text-align: left; margin-top: 10pt; font-weight: bold;">
                        — Sarah Chen, CTO at TechCorp
                    </p>
                </div>

                <div style="padding: 20pt; background-color: white;
                            border-radius: 8pt; box-shadow: 0 2pt 6pt rgba(0,0,0,0.1);">
                    <p style="font-size: 15pt; font-family: Arial; font-style: italic; color: #333333;
                              text-align: left; line-height: 1.5; margin: 0;">
                        "The all-in-one platform our remote team needed. Setup was incredibly fast and adoption was seamless."
                    </p>
                    <p style="font-size: 13pt; font-family: Arial; color: #4f46e5;
                              text-align: left; margin-top: 10pt; font-weight: bold;">
                        — Michael Rodriguez, Head of Operations at StartupX
                    </p>
                </div>
            </div>
        </div>
        """,

        # Slide 6: Pricing
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #4f46e5;
                       text-align: left;">
                Simple, Transparent Pricing
            </h2>

            <!-- Pricing Cards -->
            <div style="position: absolute; left: 60pt; top: 130pt; width: 180pt; height: 220pt;
                        background-color: #fafafa; border-radius: 8pt; padding: 20pt; border: 2pt solid #e5e7eb;">
                <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #333333;
                           text-align: left; margin: 0;">
                    Starter
                </h3>
                <h1 style="font-size: 42pt; font-family: Arial; font-weight: bold; color: #4f46e5;
                           text-align: left; margin-top: 15pt; margin-bottom: 5pt;">
                    $10
                </h1>
                <p style="font-size: 14pt; font-family: Arial; color: #666666;
                          text-align: left; margin: 0; margin-bottom: 20pt;">
                    per user/month
                </p>
                <p style="font-size: 13pt; font-family: Arial; color: #333333;
                          text-align: left; line-height: 1.8;">
                    ✓ Up to 10 users<br/>
                    ✓ 100 GB storage<br/>
                    ✓ Basic integrations<br/>
                    ✓ Email support
                </p>
            </div>

            <div style="position: absolute; left: 265pt; top: 120pt; width: 180pt; height: 240pt;
                        background: linear-gradient(135deg, #4f46e5 0%, #0ea5e9 100%);
                        border-radius: 8pt; padding: 20pt; box-shadow: 0 4pt 12pt rgba(79,70,229,0.3);">
                <div style="position: absolute; right: 20pt; top: 20pt; background-color: #fbbf24;
                            padding: 4pt 12pt; border-radius: 4pt;">
                    <span style="font-size: 11pt; font-family: Arial; font-weight: bold; color: #333333;">
                        POPULAR
                    </span>
                </div>
                <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: white;
                           text-align: left; margin: 0;">
                    Professional
                </h3>
                <h1 style="font-size: 42pt; font-family: Arial; font-weight: bold; color: white;
                           text-align: left; margin-top: 15pt; margin-bottom: 5pt;">
                    $25
                </h1>
                <p style="font-size: 14pt; font-family: Arial; color: rgba(255,255,255,0.9);
                          text-align: left; margin: 0; margin-bottom: 20pt;">
                    per user/month
                </p>
                <p style="font-size: 13pt; font-family: Arial; color: white;
                          text-align: left; line-height: 1.8;">
                    ✓ Unlimited users<br/>
                    ✓ 1 TB storage<br/>
                    ✓ All integrations<br/>
                    ✓ Priority support<br/>
                    ✓ Advanced analytics
                </p>
            </div>

            <div style="position: absolute; left: 470pt; top: 130pt; width: 180pt; height: 220pt;
                        background-color: #fafafa; border-radius: 8pt; padding: 20pt; border: 2pt solid #e5e7eb;">
                <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #333333;
                           text-align: left; margin: 0;">
                    Enterprise
                </h3>
                <h1 style="font-size: 42pt; font-family: Arial; font-weight: bold; color: #4f46e5;
                           text-align: left; margin-top: 15pt; margin-bottom: 5pt;">
                    Custom
                </h1>
                <p style="font-size: 14pt; font-family: Arial; color: #666666;
                          text-align: left; margin: 0; margin-bottom: 20pt;">
                    contact sales
                </p>
                <p style="font-size: 13pt; font-family: Arial; color: #333333;
                          text-align: left; line-height: 1.8;">
                    ✓ Custom features<br/>
                    ✓ Unlimited storage<br/>
                    ✓ Dedicated support<br/>
                    ✓ SLA guarantee
                </p>
            </div>
        </div>
        """,

        # Slide 7: Call to Action
        """
        <div style="width: 720pt; height: 405pt;
                    background: linear-gradient(135deg, #0ea5e9 0%, #4f46e5 100%);">
            <h1 style="position: absolute; left: 60pt; top: 110pt; width: 600pt;
                       font-size: 62pt; font-family: Arial; font-weight: bold; color: white;
                       text-align: left; line-height: 1.2;">
                Ready to Transform<br/>Your Workflow?
            </h1>

            <p style="position: absolute; left: 60pt; top: 270pt; width: 500pt;
                      font-size: 22pt; font-family: Arial; color: rgba(255,255,255,0.95);
                      text-align: left; line-height: 1.5;">
                Start your 14-day free trial today.<br/>
                No credit card required.
            </p>

            <div style="position: absolute; left: 60pt; top: 345pt; width: 200pt; height: 50pt;
                        background-color: white; border-radius: 6pt; text-align: center; line-height: 50pt;">
                <span style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #4f46e5;">
                    Start Free Trial
                </span>
            </div>

            <p style="position: absolute; left: 280pt; top: 358pt;
                      font-size: 16pt; font-family: Arial; color: rgba(255,255,255,0.9);
                      text-align: left;">
                or schedule a demo →
            </p>
        </div>
        """
    ]

    create_presentation_from_html(slides, 'samples/polished_modern_saas.pptx', '16:9')
    print("✓ Created: polished_modern_saas.pptx")


def create_ecommerce_growth():
    """E-commerce Growth Strategy - Vibrant orange/purple theme"""
    slides = [
        # Slide 1: Title
        """
        <div style="width: 720pt; height: 405pt;
                    background: linear-gradient(135deg, #ff6b6b 0%, #ee5a6f 50%, #c44569 100%);">
            <h1 style="position: absolute; left: 60pt; top: 140pt; width: 600pt;
                       font-size: 68pt; font-family: Arial; font-weight: bold; color: white;
                       text-align: left; line-height: 1.1;">
                E-Commerce<br/>Growth Strategy
            </h1>
            <p style="position: absolute; left: 60pt; top: 290pt; width: 600pt;
                      font-size: 26pt; font-family: Arial; color: rgba(255,255,255,0.95);
                      text-align: left; line-height: 1.3;">
                Scaling from $1M to $10M in 12 Months
            </p>
        </div>
        """,

        # Slide 2: Current State
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #c44569;
                       text-align: left;">
                Where We Are Today
            </h2>

            <!-- Current Metrics -->
            <div style="position: absolute; left: 60pt; top: 130pt; width: 180pt; height: 130pt;
                        background: linear-gradient(135deg, #ff6b6b 0%, #ee5a6f 100%);
                        border-radius: 8pt; text-align: center; padding-top: 30pt;">
                <h1 style="font-size: 52pt; font-family: Arial; font-weight: bold; color: white;
                           margin: 0; line-height: 1;">
                    $1.2M
                </h1>
                <p style="font-size: 16pt; font-family: Arial; color: rgba(255,255,255,0.95);
                          margin-top: 12pt;">
                    Annual Revenue<br/>(2024)
                </p>
            </div>

            <div style="position: absolute; left: 260pt; top: 130pt; width: 180pt; height: 130pt;
                        background: linear-gradient(135deg, #c44569 0%, #a82e54 100%);
                        border-radius: 8pt; text-align: center; padding-top: 30pt;">
                <h1 style="font-size: 52pt; font-family: Arial; font-weight: bold; color: white;
                           margin: 0; line-height: 1;">
                    15K
                </h1>
                <p style="font-size: 16pt; font-family: Arial; color: rgba(255,255,255,0.95);
                          margin-top: 12pt;">
                    Monthly Active<br/>Customers
                </p>
            </div>

            <div style="position: absolute; left: 460pt; top: 130pt; width: 180pt; height: 130pt;
                        background: linear-gradient(135deg, #ee5a6f 0%, #c44569 100%);
                        border-radius: 8pt; text-align: center; padding-top: 30pt;">
                <h1 style="font-size: 52pt; font-family: Arial; font-weight: bold; color: white;
                           margin: 0; line-height: 1;">
                    $78
                </h1>
                <p style="font-size: 16pt; font-family: Arial; color: rgba(255,255,255,0.95);
                          margin-top: 12pt;">
                    Average Order<br/>Value
                </p>
            </div>

            <!-- Key Challenges -->
            <div style="position: absolute; left: 60pt; top: 290pt; width: 600pt;">
                <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #333333;
                           text-align: left; margin: 0; margin-bottom: 15pt;">
                    Key Challenges:
                </h3>
                <p style="font-size: 15pt; font-family: Arial; color: #666666;
                          text-align: left; line-height: 1.8; margin: 0;">
                    • Limited traffic channels (85% from paid ads)<br/>
                    • Low conversion rate (1.8%)<br/>
                    • High cart abandonment (72%)
                </p>
            </div>
        </div>
        """,

        # Slide 3: Growth Strategy
        """
        <div style="width: 720pt; height: 405pt; background-color: #fafafa;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #c44569;
                       text-align: left;">
                5-Pillar Growth Strategy
            </h2>

            <!-- Strategy Pillars -->
            <div style="position: absolute; left: 60pt; top: 120pt; width: 600pt;">
                <!-- Pillar 1 -->
                <div style="margin-bottom: 18pt; padding: 18pt; background-color: white;
                            border-left: 6pt solid #ff6b6b; border-radius: 4pt;">
                    <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #ff6b6b;
                               text-align: left; margin: 0;">
                        1. SEO & Content Marketing
                    </h3>
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Reduce ad dependency through organic traffic channels and high-value content
                    </p>
                </div>

                <!-- Pillar 2 -->
                <div style="margin-bottom: 18pt; padding: 18pt; background-color: white;
                            border-left: 6pt solid #ee5a6f; border-radius: 4pt;">
                    <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #ee5a6f;
                               text-align: left; margin: 0;">
                        2. Conversion Rate Optimization
                    </h3>
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        A/B testing, UX improvements, and personalization to boost conversions to 3.5%
                    </p>
                </div>

                <!-- Pillar 3 -->
                <div style="margin-bottom: 18pt; padding: 18pt; background-color: white;
                            border-left: 6pt solid #c44569; border-radius: 4pt;">
                    <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #c44569;
                               text-align: left; margin: 0;">
                        3. Email & SMS Automation
                    </h3>
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Recover abandoned carts and increase customer lifetime value through automation
                    </p>
                </div>

                <!-- Pillar 4 -->
                <div style="margin-bottom: 18pt; padding: 18pt; background-color: white;
                            border-left: 6pt solid #ee5a6f; border-radius: 4pt;">
                    <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #ee5a6f;
                               text-align: left; margin: 0;">
                        4. Social Commerce
                    </h3>
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Leverage Instagram Shopping, TikTok Shop, and influencer partnerships
                    </p>
                </div>

                <!-- Pillar 5 -->
                <div style="padding: 18pt; background-color: white;
                            border-left: 6pt solid #ff6b6b; border-radius: 4pt;">
                    <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #ff6b6b;
                               text-align: left; margin: 0;">
                        5. Customer Retention
                    </h3>
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 8pt; line-height: 1.4;">
                        Loyalty program and subscription model to maximize repeat purchases
                    </p>
                </div>
            </div>
        </div>
        """,

        # Slide 4: 12-Month Roadmap
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #c44569;
                       text-align: left;">
                12-Month Roadmap
            </h2>

            <!-- Timeline -->
            <div style="position: absolute; left: 80pt; top: 130pt; width: 560pt;">
                <!-- Q1 -->
                <div style="margin-bottom: 22pt;">
                    <div style="float: left; width: 80pt; height: 50pt; background-color: #ffe5e5;
                                border-radius: 6pt; text-align: center; padding-top: 10pt;">
                        <span style="font-size: 16pt; font-family: Arial; font-weight: bold; color: #ff6b6b;">
                            Q1
                        </span>
                        <p style="font-size: 12pt; font-family: Arial; color: #666666; margin-top: 2pt;">
                            Jan-Mar
                        </p>
                    </div>
                    <div style="margin-left: 100pt; padding-top: 8pt;">
                        <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                                   text-align: left; margin: 0;">
                            Foundation & Testing
                        </h3>
                        <p style="font-size: 13pt; font-family: Arial; color: #666666;
                                  text-align: left; margin-top: 4pt;">
                            Launch CRO tests, set up email automation, begin SEO strategy
                        </p>
                    </div>
                </div>

                <!-- Q2 -->
                <div style="margin-bottom: 22pt;">
                    <div style="float: left; width: 80pt; height: 50pt; background-color: #ffe5e5;
                                border-radius: 6pt; text-align: center; padding-top: 10pt;">
                        <span style="font-size: 16pt; font-family: Arial; font-weight: bold; color: #ee5a6f;">
                            Q2
                        </span>
                        <p style="font-size: 12pt; font-family: Arial; color: #666666; margin-top: 2pt;">
                            Apr-Jun
                        </p>
                    </div>
                    <div style="margin-left: 100pt; padding-top: 8pt;">
                        <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                                   text-align: left; margin: 0;">
                            Scale & Optimize
                        </h3>
                        <p style="font-size: 13pt; font-family: Arial; color: #666666;
                                  text-align: left; margin-top: 4pt;">
                            Launch social commerce, optimize top performers, expand content
                        </p>
                    </div>
                </div>

                <!-- Q3 -->
                <div style="margin-bottom: 22pt;">
                    <div style="float: left; width: 80pt; height: 50pt; background-color: #ffe5e5;
                                border-radius: 6pt; text-align: center; padding-top: 10pt;">
                        <span style="font-size: 16pt; font-family: Arial; font-weight: bold; color: #c44569;">
                            Q3
                        </span>
                        <p style="font-size: 12pt; font-family: Arial; color: #666666; margin-top: 2pt;">
                            Jul-Sep
                        </p>
                    </div>
                    <div style="margin-left: 100pt; padding-top: 8pt;">
                        <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                                   text-align: left; margin: 0;">
                            Retention Focus
                        </h3>
                        <p style="font-size: 13pt; font-family: Arial; color: #666666;
                                  text-align: left; margin-top: 4pt;">
                            Launch loyalty program, introduce subscription tiers
                        </p>
                    </div>
                </div>

                <!-- Q4 -->
                <div>
                    <div style="float: left; width: 80pt; height: 50pt; background-color: #ffe5e5;
                                border-radius: 6pt; text-align: center; padding-top: 10pt;">
                        <span style="font-size: 16pt; font-family: Arial; font-weight: bold; color: #ee5a6f;">
                            Q4
                        </span>
                        <p style="font-size: 12pt; font-family: Arial; color: #666666; margin-top: 2pt;">
                            Oct-Dec
                        </p>
                    </div>
                    <div style="margin-left: 100pt; padding-top: 8pt;">
                        <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #333333;
                                   text-align: left; margin: 0;">
                            Holiday Push & Review
                        </h3>
                        <p style="font-size: 13pt; font-family: Arial; color: #666666;
                                  text-align: left; margin-top: 4pt;">
                            Maximize Q4 sales, analyze annual results, plan for Year 2
                        </p>
                    </div>
                </div>
            </div>
        </div>
        """,

        # Slide 5: Revenue Projections
        """
        <div style="width: 720pt; height: 405pt; background-color: #fafafa;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #c44569;
                       text-align: left;">
                Revenue Projections
            </h2>

            <!-- Monthly Growth Targets -->
            <div style="position: absolute; left: 60pt; top: 130pt; width: 280pt;">
                <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #333333;
                           text-align: left; margin-bottom: 20pt;">
                    Monthly Targets:
                </h3>

                <div style="margin-bottom: 15pt; padding: 12pt; background-color: white; border-radius: 6pt;">
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin: 0;">
                        <span style="font-weight: bold; color: #ff6b6b;">Month 3:</span> $150K/mo
                    </p>
                </div>

                <div style="margin-bottom: 15pt; padding: 12pt; background-color: white; border-radius: 6pt;">
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin: 0;">
                        <span style="font-weight: bold; color: #ee5a6f;">Month 6:</span> $400K/mo
                    </p>
                </div>

                <div style="margin-bottom: 15pt; padding: 12pt; background-color: white; border-radius: 6pt;">
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin: 0;">
                        <span style="font-weight: bold; color: #c44569;">Month 9:</span> $650K/mo
                    </p>
                </div>

                <div style="padding: 12pt; background-color: white; border-radius: 6pt;">
                    <p style="font-size: 14pt; font-family: Arial; color: #666666;
                              text-align: left; margin: 0;">
                        <span style="font-weight: bold; color: #ff6b6b;">Month 12:</span> $900K/mo
                    </p>
                </div>
            </div>

            <!-- Year-End Goals -->
            <div style="position: absolute; left: 380pt; top: 130pt; width: 280pt;">
                <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #333333;
                           text-align: left; margin-bottom: 20pt;">
                    Year-End Goals:
                </h3>

                <div style="margin-bottom: 20pt; padding: 20pt;
                            background: linear-gradient(135deg, #ff6b6b 0%, #c44569 100%);
                            border-radius: 8pt; text-align: center;">
                    <h1 style="font-size: 48pt; font-family: Arial; font-weight: bold; color: white;
                               margin: 0;">
                        $10M
                    </h1>
                    <p style="font-size: 16pt; font-family: Arial; color: rgba(255,255,255,0.95);
                              margin-top: 8pt;">
                        Annual Revenue
                    </p>
                </div>

                <div style="padding: 15pt; background-color: white; border-radius: 6pt; margin-bottom: 12pt;">
                    <p style="font-size: 15pt; font-family: Arial; color: #333333;
                              text-align: left; margin: 0; font-weight: bold;">
                        3.5% Conversion Rate
                    </p>
                    <p style="font-size: 12pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 4pt;">
                        +94% improvement
                    </p>
                </div>

                <div style="padding: 15pt; background-color: white; border-radius: 6pt;">
                    <p style="font-size: 15pt; font-family: Arial; color: #333333;
                              text-align: left; margin: 0; font-weight: bold;">
                        50K Active Customers
                    </p>
                    <p style="font-size: 12pt; font-family: Arial; color: #666666;
                              text-align: left; margin-top: 4pt;">
                        +233% growth
                    </p>
                </div>
            </div>
        </div>
        """,

        # Slide 6: Investment Required
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <h2 style="position: absolute; left: 60pt; top: 40pt; width: 600pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #c44569;
                       text-align: left;">
                Investment Required
            </h2>

            <!-- Budget Breakdown -->
            <div style="position: absolute; left: 60pt; top: 130pt; width: 600pt;">
                <div style="margin-bottom: 20pt; padding: 20pt; background-color: #fafafa; border-radius: 8pt;">
                    <div style="margin-bottom: 15pt;">
                        <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #ff6b6b;
                                   text-align: left; margin: 0; float: left;">
                            Marketing & Advertising
                        </h3>
                        <span style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #ff6b6b;
                                     float: right;">
                            $2.5M
                        </span>
                        <div style="clear: both;"></div>
                    </div>
                    <p style="font-size: 13pt; font-family: Arial; color: #666666;
                              text-align: left; margin: 0;">
                        Paid ads, influencer partnerships, content creation
                    </p>
                </div>

                <div style="margin-bottom: 20pt; padding: 20pt; background-color: #fafafa; border-radius: 8pt;">
                    <div style="margin-bottom: 15pt;">
                        <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #ee5a6f;
                                   text-align: left; margin: 0; float: left;">
                            Technology & Tools
                        </h3>
                        <span style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #ee5a6f;
                                     float: right;">
                            $500K
                        </span>
                        <div style="clear: both;"></div>
                    </div>
                    <p style="font-size: 13pt; font-family: Arial; color: #666666;
                              text-align: left; margin: 0;">
                        CRO platform, email/SMS tools, analytics, CRM
                    </p>
                </div>

                <div style="margin-bottom: 20pt; padding: 20pt; background-color: #fafafa; border-radius: 8pt;">
                    <div style="margin-bottom: 15pt;">
                        <h3 style="font-size: 18pt; font-family: Arial; font-weight: bold; color: #c44569;
                                   text-align: left; margin: 0; float: left;">
                            Team Expansion
                        </h3>
                        <span style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #c44569;
                                     float: right;">
                            $800K
                        </span>
                        <div style="clear: both;"></div>
                    </div>
                    <p style="font-size: 13pt; font-family: Arial; color: #666666;
                              text-align: left; margin: 0;">
                        Growth marketer, CRO specialist, content team
                    </p>
                </div>

                <div style="padding: 25pt; background: linear-gradient(135deg, #ff6b6b 0%, #c44569 100%);
                            border-radius: 8pt; text-align: center;">
                    <h3 style="font-size: 18pt; font-family: Arial; color: rgba(255,255,255,0.9);
                               margin: 0; margin-bottom: 10pt;">
                        Total Investment
                    </h3>
                    <h1 style="font-size: 56pt; font-family: Arial; font-weight: bold; color: white;
                               margin: 0;">
                        $3.8M
                    </h1>
                    <p style="font-size: 15pt; font-family: Arial; color: rgba(255,255,255,0.95);
                              margin-top: 10pt;">
                        Expected ROI: 163% | Payback Period: 8 months
                    </p>
                </div>
            </div>
        </div>
        """,

        # Slide 7: Next Steps
        """
        <div style="width: 720pt; height: 405pt;
                    background: linear-gradient(135deg, #ee5a6f 0%, #c44569 100%);">
            <h1 style="position: absolute; left: 60pt; top: 100pt; width: 600pt;
                       font-size: 58pt; font-family: Arial; font-weight: bold; color: white;
                       text-align: left; line-height: 1.2;">
                Let's Accelerate<br/>Your Growth
            </h1>

            <div style="position: absolute; left: 60pt; top: 250pt; width: 600pt;">
                <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold;
                           color: rgba(255,255,255,0.95); text-align: left; margin-bottom: 20pt;">
                    Next Steps:
                </h3>
                <p style="font-size: 18pt; font-family: Arial; color: rgba(255,255,255,0.95);
                          text-align: left; line-height: 1.8; margin: 0;">
                    1. Approve budget and roadmap<br/>
                    2. Hire key team members (Weeks 1-4)<br/>
                    3. Launch Q1 initiatives (Month 1)<br/>
                    4. Monthly progress reviews
                </p>
            </div>
        </div>
        """
    ]

    create_presentation_from_html(slides, 'samples/polished_ecommerce_growth.pptx', '16:9')
    print("✓ Created: polished_ecommerce_growth.pptx")


def create_minimal_product():
    """Minimal Product Design - Clean black/white/gray theme"""
    slides = [
        # Slide 1: Title
        """
        <div style="width: 720pt; height: 405pt; background-color: #000000;">
            <h1 style="position: absolute; left: 60pt; top: 160pt; width: 600pt;
                       font-size: 72pt; font-family: Arial; font-weight: bold; color: white;
                       text-align: left; line-height: 1.1;">
                Minimal<br/>Design System
            </h1>
            <div style="position: absolute; left: 60pt; top: 320pt; width: 4pt; height: 50pt;
                        background-color: white;">
            </div>
            <p style="position: absolute; left: 80pt; top: 325pt; width: 500pt;
                      font-size: 20pt; font-family: Arial; color: #999999;
                      text-align: left; letter-spacing: 2pt;">
                SIMPLICITY AT SCALE
            </p>
        </div>
        """,

        # Slide 2: Philosophy
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <div style="position: absolute; left: 60pt; top: 40pt; width: 4pt; height: 40pt;
                        background-color: #000000;">
            </div>
            <h2 style="position: absolute; left: 80pt; top: 40pt; width: 560pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #000000;
                       text-align: left;">
                Design Philosophy
            </h2>

            <!-- Philosophy Grid -->
            <div style="position: absolute; left: 60pt; top: 130pt; width: 600pt;">
                <div style="margin-bottom: 40pt;">
                    <h1 style="font-size: 72pt; font-family: Arial; font-weight: bold; color: #f0f0f0;
                               text-align: left; margin: 0; line-height: 1;">
                        01
                    </h1>
                    <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #000000;
                               text-align: left; margin-top: 10pt; margin-bottom: 8pt;">
                        Less is More
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; line-height: 1.5; margin: 0;">
                        Strip away the unnecessary. Every element must serve a purpose. Clean interfaces lead to better user experiences.
                    </p>
                </div>

                <div style="margin-bottom: 40pt;">
                    <h1 style="font-size: 72pt; font-family: Arial; font-weight: bold; color: #f0f0f0;
                               text-align: left; margin: 0; line-height: 1;">
                        02
                    </h1>
                    <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #000000;
                               text-align: left; margin-top: 10pt; margin-bottom: 8pt;">
                        Form Follows Function
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; line-height: 1.5; margin: 0;">
                        Design should solve problems, not create them. Prioritize usability over decoration.
                    </p>
                </div>

                <div>
                    <h1 style="font-size: 72pt; font-family: Arial; font-weight: bold; color: #f0f0f0;
                               text-align: left; margin: 0; line-height: 1;">
                        03
                    </h1>
                    <h3 style="font-size: 24pt; font-family: Arial; font-weight: bold; color: #000000;
                               text-align: left; margin-top: 10pt; margin-bottom: 8pt;">
                        Consistency Creates Clarity
                    </h3>
                    <p style="font-size: 15pt; font-family: Arial; color: #666666;
                              text-align: left; line-height: 1.5; margin: 0;">
                        A unified system reduces cognitive load and builds user confidence through familiarity.
                    </p>
                </div>
            </div>
        </div>
        """,

        # Slide 3: Color System
        """
        <div style="width: 720pt; height: 405pt; background-color: #fafafa;">
            <div style="position: absolute; left: 60pt; top: 40pt; width: 4pt; height: 40pt;
                        background-color: #000000;">
            </div>
            <h2 style="position: absolute; left: 80pt; top: 40pt; width: 560pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #000000;
                       text-align: left;">
                Color System
            </h2>

            <!-- Color Swatches -->
            <div style="position: absolute; left: 60pt; top: 130pt; width: 600pt;">
                <!-- Primary Colors -->
                <h3 style="font-size: 16pt; font-family: Arial; font-weight: bold; color: #666666;
                           text-align: left; margin-bottom: 20pt; letter-spacing: 1pt;">
                    PRIMARY PALETTE
                </h3>

                <div style="margin-bottom: 50pt;">
                    <div style="float: left; width: 140pt; margin-right: 15pt;">
                        <div style="width: 140pt; height: 100pt; background-color: #000000;"></div>
                        <p style="font-size: 14pt; font-family: Arial; color: #333333;
                                  text-align: left; margin-top: 10pt; font-weight: bold;">
                            Black
                        </p>
                        <p style="font-size: 12pt; font-family: Arial; color: #999999;
                                  text-align: left; margin-top: 4pt;">
                            #000000
                        </p>
                    </div>

                    <div style="float: left; width: 140pt; margin-right: 15pt;">
                        <div style="width: 140pt; height: 100pt; background-color: #333333;"></div>
                        <p style="font-size: 14pt; font-family: Arial; color: #333333;
                                  text-align: left; margin-top: 10pt; font-weight: bold;">
                            Dark Gray
                        </p>
                        <p style="font-size: 12pt; font-family: Arial; color: #999999;
                                  text-align: left; margin-top: 4pt;">
                            #333333
                        </p>
                    </div>

                    <div style="float: left; width: 140pt; margin-right: 15pt;">
                        <div style="width: 140pt; height: 100pt; background-color: #999999;"></div>
                        <p style="font-size: 14pt; font-family: Arial; color: #333333;
                                  text-align: left; margin-top: 10pt; font-weight: bold;">
                            Medium Gray
                        </p>
                        <p style="font-size: 12pt; font-family: Arial; color: #999999;
                                  text-align: left; margin-top: 4pt;">
                            #999999
                        </p>
                    </div>

                    <div style="float: left; width: 140pt;">
                        <div style="width: 140pt; height: 100pt; background-color: #f0f0f0;
                                    border: 1pt solid #e0e0e0;"></div>
                        <p style="font-size: 14pt; font-family: Arial; color: #333333;
                                  text-align: left; margin-top: 10pt; font-weight: bold;">
                            Light Gray
                        </p>
                        <p style="font-size: 12pt; font-family: Arial; color: #999999;
                                  text-align: left; margin-top: 4pt;">
                            #F0F0F0
                        </p>
                    </div>
                    <div style="clear: both;"></div>
                </div>

                <!-- Usage Guide -->
                <div style="padding: 20pt; background-color: white; border-left: 3pt solid #000000;">
                    <p style="font-size: 14pt; font-family: Arial; color: #333333;
                              text-align: left; line-height: 1.6; margin: 0;">
                        <span style="font-weight: bold;">Usage:</span> Black for text and UI elements.
                        Grays for hierarchy and borders. Light gray for backgrounds and disabled states.
                    </p>
                </div>
            </div>
        </div>
        """,

        # Slide 4: Typography
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <div style="position: absolute; left: 60pt; top: 40pt; width: 4pt; height: 40pt;
                        background-color: #000000;">
            </div>
            <h2 style="position: absolute; left: 80pt; top: 40pt; width: 560pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #000000;
                       text-align: left;">
                Typography
            </h2>

            <!-- Type Scale -->
            <div style="position: absolute; left: 60pt; top: 140pt; width: 600pt;">
                <div style="margin-bottom: 25pt;">
                    <h1 style="font-size: 56pt; font-family: Arial; font-weight: bold; color: #000000;
                               text-align: left; margin: 0; line-height: 1;">
                        Display
                    </h1>
                    <p style="font-size: 12pt; font-family: Arial; color: #999999;
                              text-align: left; margin-top: 6pt;">
                        56pt / Bold / Line height 1.0
                    </p>
                </div>

                <div style="margin-bottom: 25pt;">
                    <h1 style="font-size: 42pt; font-family: Arial; font-weight: bold; color: #000000;
                               text-align: left; margin: 0; line-height: 1.1;">
                        Heading 1
                    </h1>
                    <p style="font-size: 12pt; font-family: Arial; color: #999999;
                              text-align: left; margin-top: 6pt;">
                        42pt / Bold / Line height 1.1
                    </p>
                </div>

                <div style="margin-bottom: 25pt;">
                    <h2 style="font-size: 28pt; font-family: Arial; font-weight: bold; color: #000000;
                               text-align: left; margin: 0; line-height: 1.2;">
                        Heading 2
                    </h2>
                    <p style="font-size: 12pt; font-family: Arial; color: #999999;
                              text-align: left; margin-top: 6pt;">
                        28pt / Bold / Line height 1.2
                    </p>
                </div>

                <div style="margin-bottom: 25pt;">
                    <p style="font-size: 16pt; font-family: Arial; color: #000000;
                              text-align: left; margin: 0; line-height: 1.5;">
                        Body text at 16pt provides optimal readability across all devices and screen sizes.
                    </p>
                    <p style="font-size: 12pt; font-family: Arial; color: #999999;
                              text-align: left; margin-top: 6pt;">
                        16pt / Regular / Line height 1.5
                    </p>
                </div>

                <div>
                    <p style="font-size: 13pt; font-family: Arial; color: #666666;
                              text-align: left; margin: 0; line-height: 1.4;">
                        Small text for captions and secondary information
                    </p>
                    <p style="font-size: 12pt; font-family: Arial; color: #999999;
                              text-align: left; margin-top: 6pt;">
                        13pt / Regular / Line height 1.4
                    </p>
                </div>
            </div>
        </div>
        """,

        # Slide 5: Spacing System
        """
        <div style="width: 720pt; height: 405pt; background-color: #fafafa;">
            <div style="position: absolute; left: 60pt; top: 40pt; width: 4pt; height: 40pt;
                        background-color: #000000;">
            </div>
            <h2 style="position: absolute; left: 80pt; top: 40pt; width: 560pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #000000;
                       text-align: left;">
                Spacing System
            </h2>

            <!-- Spacing Scale -->
            <div style="position: absolute; left: 60pt; top: 140pt; width: 600pt;">
                <p style="font-size: 15pt; font-family: Arial; color: #666666;
                          text-align: left; margin-bottom: 30pt; line-height: 1.5;">
                    Consistent spacing creates visual rhythm. Use multiples of 8pt for all spacing decisions.
                </p>

                <!-- Spacing Examples -->
                <div style="margin-bottom: 20pt;">
                    <div style="float: left; width: 8pt; height: 8pt; background-color: #000000;
                                margin-right: 15pt; margin-top: 5pt;"></div>
                    <p style="font-size: 16pt; font-family: Arial; color: #000000;
                              text-align: left; margin: 0; float: left;">
                        <span style="font-weight: bold;">XS</span> — 8pt
                    </p>
                    <p style="font-size: 14pt; font-family: Arial; color: #999999;
                              text-align: left; margin: 0; margin-left: 200pt;">
                        Tight spacing between related items
                    </p>
                    <div style="clear: both;"></div>
                </div>

                <div style="margin-bottom: 20pt;">
                    <div style="float: left; width: 16pt; height: 16pt; background-color: #000000;
                                margin-right: 15pt; margin-top: 5pt;"></div>
                    <p style="font-size: 16pt; font-family: Arial; color: #000000;
                              text-align: left; margin: 0; float: left;">
                        <span style="font-weight: bold;">S</span> — 16pt
                    </p>
                    <p style="font-size: 14pt; font-family: Arial; color: #999999;
                              text-align: left; margin: 0; margin-left: 200pt;">
                        Standard padding for components
                    </p>
                    <div style="clear: both;"></div>
                </div>

                <div style="margin-bottom: 20pt;">
                    <div style="float: left; width: 24pt; height: 24pt; background-color: #000000;
                                margin-right: 15pt; margin-top: 5pt;"></div>
                    <p style="font-size: 16pt; font-family: Arial; color: #000000;
                              text-align: left; margin: 0; float: left;">
                        <span style="font-weight: bold;">M</span> — 24pt
                    </p>
                    <p style="font-size: 14pt; font-family: Arial; color: #999999;
                              text-align: left; margin: 0; margin-left: 200pt;">
                        Spacing between component groups
                    </p>
                    <div style="clear: both;"></div>
                </div>

                <div style="margin-bottom: 20pt;">
                    <div style="float: left; width: 32pt; height: 32pt; background-color: #000000;
                                margin-right: 15pt; margin-top: 5pt;"></div>
                    <p style="font-size: 16pt; font-family: Arial; color: #000000;
                              text-align: left; margin: 0; float: left;">
                        <span style="font-weight: bold;">L</span> — 32pt
                    </p>
                    <p style="font-size: 14pt; font-family: Arial; color: #999999;
                              text-align: left; margin: 0; margin-left: 200pt;">
                        Section spacing and large gaps
                    </p>
                    <div style="clear: both;"></div>
                </div>

                <div style="margin-bottom: 20pt;">
                    <div style="float: left; width: 48pt; height: 48pt; background-color: #000000;
                                margin-right: 15pt; margin-top: 5pt;"></div>
                    <p style="font-size: 16pt; font-family: Arial; color: #000000;
                              text-align: left; margin: 0; float: left;">
                        <span style="font-weight: bold;">XL</span> — 48pt
                    </p>
                    <p style="font-size: 14pt; font-family: Arial; color: #999999;
                              text-align: left; margin: 0; margin-left: 200pt;">
                        Major layout divisions
                    </p>
                    <div style="clear: both;"></div>
                </div>
            </div>
        </div>
        """,

        # Slide 6: Components
        """
        <div style="width: 720pt; height: 405pt; background-color: #ffffff;">
            <div style="position: absolute; left: 60pt; top: 40pt; width: 4pt; height: 40pt;
                        background-color: #000000;">
            </div>
            <h2 style="position: absolute; left: 80pt; top: 40pt; width: 560pt;
                       font-size: 42pt; font-family: Arial; font-weight: bold; color: #000000;
                       text-align: left;">
                Core Components
            </h2>

            <!-- Component Examples -->
            <div style="position: absolute; left: 60pt; top: 140pt; width: 600pt;">
                <!-- Buttons -->
                <div style="margin-bottom: 35pt;">
                    <p style="font-size: 13pt; font-family: Arial; color: #999999;
                              text-align: left; margin-bottom: 12pt; letter-spacing: 1pt;">
                        BUTTONS
                    </p>

                    <div style="display: inline-block; padding: 12pt 32pt; background-color: #000000;
                                margin-right: 15pt;">
                        <span style="font-size: 14pt; font-family: Arial; font-weight: bold; color: white;">
                            Primary Action
                        </span>
                    </div>

                    <div style="display: inline-block; padding: 12pt 32pt; background-color: white;
                                border: 2pt solid #000000; margin-right: 15pt;">
                        <span style="font-size: 14pt; font-family: Arial; font-weight: bold; color: #000000;">
                            Secondary Action
                        </span>
                    </div>

                    <div style="display: inline-block; padding: 12pt 32pt;">
                        <span style="font-size: 14pt; font-family: Arial; color: #999999;
                                     text-decoration: underline;">
                            Text Link
                        </span>
                    </div>
                </div>

                <!-- Input Fields -->
                <div style="margin-bottom: 35pt;">
                    <p style="font-size: 13pt; font-family: Arial; color: #999999;
                              text-align: left; margin-bottom: 12pt; letter-spacing: 1pt;">
                        INPUT FIELDS
                    </p>

                    <div style="padding: 12pt 16pt; background-color: white;
                                border: 1pt solid #e0e0e0; width: 280pt;">
                        <span style="font-size: 14pt; font-family: Arial; color: #999999;">
                            Email address
                        </span>
                    </div>
                </div>

                <!-- Cards -->
                <div>
                    <p style="font-size: 13pt; font-family: Arial; color: #999999;
                              text-align: left; margin-bottom: 12pt; letter-spacing: 1pt;">
                        CARDS
                    </p>

                    <div style="padding: 24pt; background-color: #fafafa; width: 520pt;">
                        <h3 style="font-size: 20pt; font-family: Arial; font-weight: bold; color: #000000;
                                   text-align: left; margin: 0; margin-bottom: 12pt;">
                            Card Title
                        </h3>
                        <p style="font-size: 14pt; font-family: Arial; color: #666666;
                                  text-align: left; margin: 0; line-height: 1.5;">
                            Cards contain content and actions about a single subject. Use minimal borders and subtle backgrounds to maintain the clean aesthetic.
                        </p>
                    </div>
                </div>
            </div>
        </div>
        """,

        # Slide 7: Summary
        """
        <div style="width: 720pt; height: 405pt; background-color: #000000;">
            <div style="position: absolute; left: 60pt; top: 60pt; width: 600pt;">
                <h1 style="font-size: 58pt; font-family: Arial; font-weight: bold; color: white;
                           text-align: left; margin: 0; line-height: 1.2; margin-bottom: 40pt;">
                    Design is not<br/>what it looks like.
                </h1>

                <h1 style="font-size: 58pt; font-family: Arial; font-weight: bold; color: white;
                           text-align: left; margin: 0; line-height: 1.2; margin-bottom: 50pt;">
                    Design is<br/>how it works.
                </h1>

                <div style="width: 4pt; height: 60pt; background-color: white; margin-bottom: 20pt;"></div>

                <p style="font-size: 16pt; font-family: Arial; color: #999999;
                          text-align: left; letter-spacing: 2pt;">
                    MINIMAL DESIGN SYSTEM 2025
                </p>
            </div>
        </div>
        """
    ]

    create_presentation_from_html(slides, 'samples/polished_minimal_design.pptx', '16:9')
    print("✓ Created: polished_minimal_design.pptx")


if __name__ == '__main__':
    print("Creating 4 additional polished presentations...\n")

    create_creative_agency()
    create_modern_saas()
    create_ecommerce_growth()
    create_minimal_product()

    print("\n✅ All 4 polished presentations created successfully!")
    print("\nGenerated presentations:")
    print("  • samples/polished_creative_agency.pptx")
    print("  • samples/polished_modern_saas.pptx")
    print("  • samples/polished_ecommerce_growth.pptx")
    print("  • samples/polished_minimal_design.pptx")
