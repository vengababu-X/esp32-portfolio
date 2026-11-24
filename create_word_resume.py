import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_resume():
    doc = Document()

    # Set page margins (0.6 in on all sides for clean, dense, professional 2-page fit)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)

    COLOR_PRIMARY = RGBColor(15, 44, 89)     # Deep Navy #0F2C59
    COLOR_ACCENT = RGBColor(30, 64, 175)     # Vibrant Blue #1E40AF
    COLOR_DARK = RGBColor(30, 41, 59)        # Slate 800 #1E293B
    COLOR_MUTED = RGBColor(71, 85, 105)      # Slate 600 #475569
    COLOR_BORDER_HEX = "1E3A8A"

    def set_cell_background(cell, fill_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title.upper())
        run.bold = True
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        run.font.color.rgb = COLOR_PRIMARY
        
        # Add bottom border under section header
        pBrd = parse_xml(f'<w:pBrd {nsdecls("w")}><w:bottom w:val="single" w:sz="8" w:space="2" w:color="{COLOR_BORDER_HEX}"/></w:pBrd>')
        p._p.get_or_add_pPr().append(pBrd)

    # HEADER
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(1)
    run_name = p_name.add_run("BOLIGORLA VENGABABU")
    run_name.bold = True
    run_name.font.name = 'Calibri'
    run_name.font.size = Pt(20)
    run_name.font.color.rgb = COLOR_PRIMARY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(2)
    run_sub = p_sub.add_run("ELECTRONICS & COMMUNICATION ENGINEERING UNDERGRADUATE")
    run_sub.bold = True
    run_sub.font.name = 'Calibri'
    run_sub.font.size = Pt(10.5)
    run_sub.font.color.rgb = COLOR_ACCENT

    p_domains = doc.add_paragraph()
    p_domains.paragraph_format.space_before = Pt(0)
    p_domains.paragraph_format.space_after = Pt(3)
    run_domains = p_domains.add_run("Embedded Systems  |  Railway Electronics  |  Communication Systems  |  Robotics  |  AI & Prompt Engineering")
    run_domains.font.name = 'Calibri'
    run_domains.font.size = Pt(9.5)
    run_domains.font.color.rgb = COLOR_MUTED

    p_contact = doc.add_paragraph()
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(6)
    run_contact = p_contact.add_run("Kavali, Andhra Pradesh, India   •   Phone: +91 91218 43829   •   Email: boligorlavengababu@gmail.com\nLinkedIn: linkedin.com/in/boligorla-vengababu-503729349/   •   GitHub: github.com/vengababu-X")
    run_contact.font.name = 'Calibri'
    run_contact.font.size = Pt(9)
    run_contact.font.color.rgb = COLOR_DARK

    # DUAL TARGET OBJECTIVE
    add_section_header("Dual-Target Career Objective (Medha Servo Drives & ECIL Hyderabad)")
    p_obj = doc.add_paragraph()
    p_obj.paragraph_format.space_before = Pt(2)
    p_obj.paragraph_format.space_after = Pt(4)
    r_obj = p_obj.add_run(
        "Dedicated Electronics and Communication Engineering undergraduate seeking an entry-level engineering role in Locomotive & Railway Control Systems (Medha Servo Drives) and Strategic Electronics, Defense Communications & Instrumentation (ECIL Hyderabad). Offering proven hands-on capability in embedded C/C++ firmware, railway rolling-stock monitoring concepts, CAN bus/RS-485 industrial protocols, faculty-guided communication systems research, and Verilog RTL digital design."
    )
    r_obj.font.name = 'Calibri'
    r_obj.font.size = Pt(9.5)
    r_obj.font.color.rgb = COLOR_DARK

    # PROFESSIONAL SUMMARY
    add_section_header("Professional Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(2)
    p_sum.paragraph_format.space_after = Pt(4)
    r_sum = p_sum.add_run(
        "Motivated Electronics and Communication Engineering undergraduate (2023–2027) with hands-on experience in embedded systems, microcontroller firmware development, railway technology, robotics, and sensor integration. Proficient in ESP32, ESP8266, Arduino, Embedded C/C++, Python, digital electronics, and hardware-software integration. Actively collaborating with a faculty mentor on communication engineering research and technical manuscript development. Experienced in AI-assisted engineering workflows, prompt engineering, and structured technical problem-solving. Seeking opportunities in embedded electronics, communication systems, railway control technologies, and industrial automation."
    )
    r_sum.font.name = 'Calibri'
    r_sum.font.size = Pt(9.5)
    r_sum.font.color.rgb = COLOR_DARK

    # EDUCATION
    add_section_header("Education")
    p_edu = doc.add_paragraph()
    p_edu.paragraph_format.space_before = Pt(2)
    p_edu.paragraph_format.space_after = Pt(1)
    r_deg = p_edu.add_run("Bachelor of Technology (B.Tech) – Electronics and Communication Engineering")
    r_deg.bold = True
    r_deg.font.name = 'Calibri'
    r_deg.font.size = Pt(10)
    r_deg.font.color.rgb = COLOR_DARK

    p_col = doc.add_paragraph()
    p_col.paragraph_format.space_before = Pt(0)
    p_col.paragraph_format.space_after = Pt(4)
    r_col = p_col.add_run("PBR Visvodaya Institute of Technology & Science (PBRVITS), Kavali, Andhra Pradesh\n2023 – 2027   |   CGPA: 8.42 / 10.0 (First Class with Distinction • No Active Backlogs)")
    r_col.font.name = 'Calibri'
    r_col.font.size = Pt(9.5)
    r_col.font.color.rgb = COLOR_MUTED

    # TECHNICAL SKILLS
    add_section_header("Technical Skills")
    skills = [
        ("1. Programming & Software Development: ", "Embedded C, C++, Python, HTML, Data Structures, Problem-Solving & Bare-Metal Firmware Logic, AI-Assisted Coding, Debugging & Technical Documentation."),
        ("2. Embedded Systems & Microcontrollers: ", "ESP32 (Dual-Core Xtensa LX6), ESP8266, Arduino Uno, NodeMCU, Microcontroller Programming, Hardware Timers & Interrupt Handling, Hardware-Software Integration & Sensor Interfacing."),
        ("3. Electronics & Hardware Engineering: ", "Digital & Analog Electronics, Electronic Prototyping & Circuit Fundamentals, PWM Motor & Servo Control, PCB Design Fundamentals (KiCad / EasyEDA), Multimeter & Oscilloscope Debugging Basics."),
        ("4. Communication Technologies & Industrial Protocols: ", "UART, I2C, SPI, ESP-NOW, CAN Bus & RS-485 Fundamentals (Railway & Industrial Protocols), Communication Systems & Wireless Communication Concepts."),
        ("5. Robotics & Industrial Automation: ", "ROS 2 Fundamentals, PCA9685 16-Channel PWM Servo Controller, Ultrasonic Obstacle Detection & Multi-Servo Synchronized Gait Control, Computer Vision (OpenCV / MediaPipe) & Motor Drive Control."),
        ("6. VLSI & Digital Design (Strategic / ECIL Relevance): ", "Verilog HDL & RTL Design Concepts, Digital Logic Simulation & Verification, Sky130 PDK & OpenLane ASIC Flow."),
        ("7. AI & Prompt Engineering: ", "Structured Prompt Design & Context Management, LLM Interaction & AI-Assisted Engineering Workflows, Workflow Automation Concepts (n8n Fundamentals)."),
        ("8. Development Tools & Platforms: ", "Arduino IDE, Visual Studio Code, Git and GitHub, Supabase.")
    ]
    for cat, desc in skills:
        p_sk = doc.add_paragraph()
        p_sk.paragraph_format.space_before = Pt(0)
        p_sk.paragraph_format.space_after = Pt(1.5)
        p_sk.paragraph_format.left_indent = Inches(0.15)
        r_c = p_sk.add_run(cat)
        r_c.bold = True
        r_c.font.name = 'Calibri'
        r_c.font.size = Pt(9)
        r_c.font.color.rgb = COLOR_DARK
        r_d = p_sk.add_run(desc)
        r_d.font.name = 'Calibri'
        r_d.font.size = Pt(9)
        r_d.font.color.rgb = COLOR_DARK

    # RESEARCH EXPERIENCE
    add_section_header("Research Experience")
    p_res = doc.add_paragraph()
    p_res.paragraph_format.space_before = Pt(2)
    p_res.paragraph_format.space_after = Pt(1)
    r_rt = p_res.add_run("Undergraduate Research Collaboration – Communication Systems")
    r_rt.bold = True
    r_rt.font.name = 'Calibri'
    r_rt.font.size = Pt(10)
    r_rt.font.color.rgb = COLOR_DARK

    p_rm = doc.add_paragraph()
    p_rm.paragraph_format.space_before = Pt(0)
    p_rm.paragraph_format.space_after = Pt(2)
    r_rm = p_rm.add_run("Faculty-Guided Research | PBRVITS, Kavali   •   Research Status: Manuscript in Preparation")
    r_rm.italic = True
    r_rm.font.name = 'Calibri'
    r_rm.font.size = Pt(9)
    r_rm.font.color.rgb = COLOR_MUTED

    res_bullets = [
        "Collaborating with a faculty mentor on technical research in the communication and signal processing domain.",
        "Conducting technical literature reviews, channel characterization, system analysis, and performance evaluations.",
        "Authoring technical documentation and preparing a research manuscript for publication."
    ]
    for b in res_bullets:
        p_b = doc.add_paragraph(style='List Bullet')
        p_b.paragraph_format.space_before = Pt(0)
        p_b.paragraph_format.space_after = Pt(1)
        p_b.paragraph_format.left_indent = Inches(0.25)
        r_b = p_b.add_run(b)
        r_b.font.name = 'Calibri'
        r_b.font.size = Pt(9)
        r_b.font.color.rgb = COLOR_DARK

    # ENGINEERING PROJECTS
    add_section_header("Engineering Projects")

    projects = [
        (
            "1. RAILGUARD-X – Railway Security Robotics System",
            "Railway Technology | Robotics | Embedded Systems | ROS 2",
            [
                "Engineered a quadruped robotic prototype designed for railway track surveillance and perimeter security.",
                "Interfaced ESP8266 with a PCA9685 16-channel PWM servo driver over I2C for synchronized multi-servo actuation and gait stability.",
                "Integrated ultrasonic sensors for real-time obstacle detection and emergency halt fail-safe.",
                "Designed a modular architecture utilizing ROS 2 fundamentals for robotic communication, remote command handling, and sensor telemetry."
            ]
        ),
        (
            "2. ITMS – Intelligent Train Monitoring System",
            "Railway Technology | Rolling Stock Monitoring | Industrial Automation | Safety Telemetry",
            [
                "Designed an intelligent railway monitoring concept focused on improving rolling-stock safety, bearing/bogie thermal supervision, and operational awareness.",
                "Evaluated embedded sensor integration, fault-alert logic, and telemetry workflows for railway infrastructure.",
                "Applied industrial automation and electronics principles to enhance real-time train condition monitoring with CAN Bus and RS-485 bus topologies."
            ]
        ),
        (
            "3. Smart Wearable Safety Monitoring System",
            "IoT | Embedded Electronics | Sensor Integration | Supabase",
            [
                "Developed an ESP32-based wearable safety prototype integrating environmental, motion (MPU6050 accelerometer/gyroscope), and physiological sensor modules.",
                "Implemented low-latency ESP-NOW peer-to-peer wireless communication between ESP32 and ESP8266 hardware units.",
                "Designed automated alert mechanisms and connected cloud telemetry logging using Supabase backend."
            ]
        ),
        (
            "4. Eye-Controlled Smart Wheelchair System",
            "Assistive Robotics | Embedded Systems | Computer Vision & IR Sensing",
            [
                "Engineered a hands-free eye-controlled wheelchair prototype for individuals with severe mobility impairments.",
                "Processed directional eye-gaze and blink patterns via OpenCV/MediaPipe and translated gestures into dual-motor differential drive commands.",
                "Built hardware safety interlocks with ultrasonic sensors to trigger immediate emergency braking when obstacles are detected within 30 cm."
            ]
        ),
        (
            "5. Y2K 11-Bit Synchronous Counter (VLSI / Digital Design)",
            "Verilog HDL | Digital Electronics | OpenLane ASIC Simulation",
            [
                "Designed, synthesized, and verified an 11-bit synchronous digital counter in Verilog HDL.",
                "Implemented synchronous clear, clock-enable gating, and overflow ripple-carry detection flags.",
                "Simulated functional waveforms with zero setup/hold timing violations for custom digital architectures."
            ]
        )
    ]

    for title, tag, bullets in projects:
        p_pt = doc.add_paragraph()
        p_pt.paragraph_format.space_before = Pt(3)
        p_pt.paragraph_format.space_after = Pt(0)
        p_pt.paragraph_format.keep_with_next = True
        r_pt = p_pt.add_run(title)
        r_pt.bold = True
        r_pt.font.name = 'Calibri'
        r_pt.font.size = Pt(9.5)
        r_pt.font.color.rgb = COLOR_DARK

        p_tg = doc.add_paragraph()
        p_tg.paragraph_format.space_before = Pt(0)
        p_tg.paragraph_format.space_after = Pt(1.5)
        p_tg.paragraph_format.keep_with_next = True
        r_tg = p_tg.add_run(f"Domain: {tag}")
        r_tg.italic = True
        r_tg.font.name = 'Calibri'
        r_tg.font.size = Pt(8.5)
        r_tg.font.color.rgb = COLOR_ACCENT

        for b in bullets:
            p_b = doc.add_paragraph(style='List Bullet')
            p_b.paragraph_format.space_before = Pt(0)
            p_b.paragraph_format.space_after = Pt(1)
            p_b.paragraph_format.left_indent = Inches(0.25)
            r_b = p_b.add_run(b)
            r_b.font.name = 'Calibri'
            r_b.font.size = Pt(9)
            r_b.font.color.rgb = COLOR_DARK

    # HONORS & CERTIFICATIONS
    add_section_header("Honors, Awards & Certifications")
    honors = [
        "Paper Presentation Winner (1st Prize): National-Level Technical Symposium for innovative railway monitoring architecture.",
        "Top 10 Finalist: Regional Embedded & Robotics Hackathon (Track inspection robotics).",
        "Certifications: Embedded Systems with ARM & ESP32  •  Digital VLSI Design Fundamentals (NPTEL)  •  Industrial IoT & Protocols (I2C, SPI, CAN)."
    ]
    for h in honors:
        p_h = doc.add_paragraph(style='List Bullet')
        p_h.paragraph_format.space_before = Pt(0)
        p_h.paragraph_format.space_after = Pt(1.5)
        p_h.paragraph_format.left_indent = Inches(0.25)
        r_h = p_h.add_run(h)
        r_h.font.name = 'Calibri'
        r_h.font.size = Pt(9)
        r_h.font.color.rgb = COLOR_DARK

    # PERSONAL DETAILS & DECLARATION
    add_section_header("Personal Details & Declaration")
    p_per = doc.add_paragraph()
    p_per.paragraph_format.space_before = Pt(2)
    p_per.paragraph_format.space_after = Pt(3)
    r_per = p_per.add_run(
        "Father's Name: Boligorla Venkata Ramanaiah   •   Date of Birth: 15th August 2005\n"
        "Languages Known: English, Telugu, Hindi   •   Permanent Address: Kavali, SPSR Nellore District, Andhra Pradesh – 524201, India"
    )
    r_per.font.name = 'Calibri'
    r_per.font.size = Pt(9)
    r_per.font.color.rgb = COLOR_DARK

    p_dec = doc.add_paragraph()
    p_dec.paragraph_format.space_before = Pt(3)
    p_dec.paragraph_format.space_after = Pt(8)
    r_dec = p_dec.add_run("Declaration: I hereby declare that the information provided above is true and accurate to the best of my knowledge.")
    r_dec.italic = True
    r_dec.font.name = 'Calibri'
    r_dec.font.size = Pt(8.5)
    r_dec.font.color.rgb = COLOR_MUTED

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(0)
    p_sig.paragraph_format.space_after = Pt(0)
    r_pl = p_sig.add_run("Place: Kavali, Andhra Pradesh                                                                 ")
    r_pl.font.name = 'Calibri'
    r_pl.font.size = Pt(9)
    r_pl.font.color.rgb = COLOR_DARK
    r_sn = p_sig.add_run("BOLIGORLA VENGABABU")
    r_sn.bold = True
    r_sn.font.name = 'Calibri'
    r_sn.font.size = Pt(9.5)
    r_sn.font.color.rgb = COLOR_PRIMARY

    output_path = r"C:\Users\Bolig\.gemini\antigravity\scratch\esp32-portfolio\Boligorla_Vengababu_Resume.docx"
    doc.save(output_path)
    print(f"Resume saved successfully at {output_path}")

if __name__ == "__main__":
    create_resume()
