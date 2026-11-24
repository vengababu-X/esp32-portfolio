import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def build_executive_resume():
    doc = Document()

    # Page Margins: 0.5 in top/bottom, 0.6 in left/right for an elegant, dense, high-impact layout
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # Color Palette - Executive Modern Engineering Theme
    COLOR_NAVY = RGBColor(11, 37, 69)        # #0B2545 Deep Executive Navy
    COLOR_STEEL = RGBColor(19, 64, 116)      # #134074 Steel Blue
    COLOR_AZURE = RGBColor(0, 119, 182)      # #0077B6 Bright Azure Accent
    COLOR_DARK = RGBColor(30, 41, 59)        # #1E293B Body Text
    COLOR_MUTED = RGBColor(71, 85, 105)      # #475569 Subtitles & Meta
    COLOR_WHITE = RGBColor(255, 255, 255)

    HEX_PRIMARY = "0B2545"
    HEX_ACCENT_BG = "EEF4F8"
    HEX_BORDER = "134074"
    HEX_LIGHT_BORDER = "CBD5E1"
    HEX_CARD_BG = "F8FAFC"
    HEX_MEDHA_BG = "F0FDF4"
    HEX_ECIL_BG = "EFF6FF"

    def set_cell_background(cell, fill_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        cell._tc.get_or_add_tcPr().append(tcMar)

    def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
        tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
        borders = {'top': top, 'bottom': bottom, 'left': left, 'right': right}
        for b_name, b_val in borders.items():
            if b_val:
                b_elm = parse_xml(f'<w:{b_name} {nsdecls("w")} w:val="{b_val.get("val", "single")}" w:sz="{b_val.get("sz", "6")}" w:space="0" w:color="{b_val.get("color", "CBD5E1")}"/>')
                tcBorders.append(b_elm)
            else:
                b_elm = parse_xml(f'<w:{b_name} {nsdecls("w")} w:val="none"/>')
                tcBorders.append(b_elm)
        cell._tc.get_or_add_tcPr().append(tcBorders)

    # ==========================================
    # 1. HEADER CARD (Two-Column Executive Banner)
    # ==========================================
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False

    cell_left = header_table.cell(0, 0)
    cell_right = header_table.cell(0, 1)

    cell_left.width = Inches(4.5)
    cell_right.width = Inches(2.8)

    set_cell_background(cell_left, HEX_ACCENT_BG)
    set_cell_background(cell_right, HEX_ACCENT_BG)
    set_cell_margins(cell_left, top=140, bottom=140, left=160, right=140)
    set_cell_margins(cell_right, top=140, bottom=140, left=140, right=160)
    
    border_spec = {"val": "single", "sz": "10", "color": HEX_BORDER}
    set_cell_border(cell_left, top=border_spec, bottom=border_spec, left=border_spec, right=None)
    set_cell_border(cell_right, top=border_spec, bottom=border_spec, left=None, right=border_spec)

    # Left Column: Name & Title
    p_name = cell_left.paragraphs[0]
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(2)
    r_name = p_name.add_run("BOLIGORLA VENGABABU")
    r_name.bold = True
    r_name.font.name = 'Calibri'
    r_name.font.size = Pt(20)
    r_name.font.color.rgb = COLOR_NAVY

    p_sub = cell_left.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(2)
    r_sub = p_sub.add_run("ELECTRONICS & COMMUNICATION ENGINEERING UNDERGRADUATE")
    r_sub.bold = True
    r_sub.font.name = 'Calibri'
    r_sub.font.size = Pt(9.5)
    r_sub.font.color.rgb = COLOR_STEEL

    p_spec = cell_left.add_paragraph()
    p_spec.paragraph_format.space_before = Pt(0)
    p_spec.paragraph_format.space_after = Pt(0)
    r_spec = p_spec.add_run("Embedded Systems  |  Railway Electronics  |  Communication Systems  |  Robotics  |  VLSI")
    r_spec.font.name = 'Calibri'
    r_spec.font.size = Pt(8.5)
    r_spec.font.color.rgb = COLOR_MUTED

    # Right Column: Contact Details & Links
    p_c1 = cell_right.paragraphs[0]
    p_c1.paragraph_format.space_before = Pt(0)
    p_c1.paragraph_format.space_after = Pt(1.5)
    p_c1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_c1 = p_c1.add_run("📍 Kavali, Andhra Pradesh, India")
    r_c1.font.name = 'Calibri'
    r_c1.font.size = Pt(8.5)
    r_c1.font.color.rgb = COLOR_DARK

    p_c2 = cell_right.add_paragraph()
    p_c2.paragraph_format.space_before = Pt(0)
    p_c2.paragraph_format.space_after = Pt(1.5)
    p_c2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_c2 = p_c2.add_run("📞 +91 91218 43829")
    r_c2.bold = True
    r_c2.font.name = 'Calibri'
    r_c2.font.size = Pt(8.5)
    r_c2.font.color.rgb = COLOR_DARK

    p_c3 = cell_right.add_paragraph()
    p_c3.paragraph_format.space_before = Pt(0)
    p_c3.paragraph_format.space_after = Pt(1.5)
    p_c3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_c3 = p_c3.add_run("✉ boligorlavengababu@gmail.com")
    r_c3.font.name = 'Calibri'
    r_c3.font.size = Pt(8.5)
    r_c3.font.color.rgb = COLOR_STEEL

    p_c4 = cell_right.add_paragraph()
    p_c4.paragraph_format.space_before = Pt(0)
    p_c4.paragraph_format.space_after = Pt(0)
    p_c4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_c4 = p_c4.add_run("🔗 linkedin.com/in/boligorla-vengababu-503729349/\n💻 github.com/vengababu-X")
    r_c4.font.name = 'Calibri'
    r_c4.font.size = Pt(8)
    r_c4.font.color.rgb = COLOR_MUTED

    # Spacer
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(4)
    p_sp.paragraph_format.space_after = Pt(0)

    # ==========================================
    # 2. DUAL TARGET APPLICATION CARD (Medha & ECIL)
    # ==========================================
    dual_table = doc.add_table(rows=1, cols=2)
    dual_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    dual_table.autofit = False

    c_medha = dual_table.cell(0, 0)
    c_ecil = dual_table.cell(0, 1)
    c_medha.width = Inches(3.65)
    c_ecil.width = Inches(3.65)

    set_cell_background(c_medha, HEX_MEDHA_BG)
    set_cell_background(c_ecil, HEX_ECIL_BG)
    set_cell_margins(c_medha, top=80, bottom=80, left=120, right=120)
    set_cell_margins(c_ecil, top=80, bottom=80, left=120, right=120)

    border_medha = {"val": "single", "sz": "6", "color": "15803D"}
    border_ecil = {"val": "single", "sz": "6", "color": "1D4ED8"}
    set_cell_border(c_medha, top=border_medha, bottom=border_medha, left=border_medha, right=border_medha)
    set_cell_border(c_ecil, top=border_ecil, bottom=border_ecil, left=border_ecil, right=border_ecil)

    # Medha Box
    p_mb_title = c_medha.paragraphs[0]
    p_mb_title.paragraph_format.space_before = Pt(0)
    p_mb_title.paragraph_format.space_after = Pt(2)
    r_mbt = p_mb_title.add_run("🎯 TARGET: MEDHA SERVO DRIVES PVT. LTD.")
    r_mbt.bold = True
    r_mbt.font.name = 'Calibri'
    r_mbt.font.size = Pt(8.5)
    r_mbt.font.color.rgb = RGBColor(21, 128, 61)

    p_mb_body = c_medha.add_paragraph()
    p_mb_body.paragraph_format.space_before = Pt(0)
    p_mb_body.paragraph_format.space_after = Pt(0)
    r_mbb = p_mb_body.add_run("Specialization in Railway Technology & Rolling Stock Monitoring (ITMS, RAILGUARD-X), 16-channel PWM servo/motor control (PCA9685), CAN Bus & RS-485 industrial protocols, and safety-critical fail-safe interlocks.")
    r_mbb.font.name = 'Calibri'
    r_mbb.font.size = Pt(8)
    r_mbb.font.color.rgb = COLOR_DARK

    # ECIL Box
    p_ec_title = c_ecil.paragraphs[0]
    p_ec_title.paragraph_format.space_before = Pt(0)
    p_ec_title.paragraph_format.space_after = Pt(2)
    r_ect = p_ec_title.add_run("🎯 TARGET: ECIL HYDERABAD (DEFENSE & STRATEGIC)")
    r_ect.bold = True
    r_ect.font.name = 'Calibri'
    r_ect.font.size = Pt(8.5)
    r_ect.font.color.rgb = RGBColor(29, 78, 216)

    p_ec_body = c_ecil.add_paragraph()
    p_ec_body.paragraph_format.space_before = Pt(0)
    p_ec_body.paragraph_format.space_after = Pt(0)
    r_ecb = p_ec_body.add_run("Specialization in Faculty-Guided Communication Systems Research, bare-metal C/C++ firmware, digital instrumentation, and VLSI RTL design & verification (Verilog HDL, Sky130 PDK, OpenLane ASIC flow).")
    r_ecb.font.name = 'Calibri'
    r_ecb.font.size = Pt(8)
    r_ecb.font.color.rgb = COLOR_DARK

    # Helper function for Section Titles
    def add_section_header(title, icon="▪"):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(7)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        
        r_icon = p.add_run(f"{icon} ")
        r_icon.bold = True
        r_icon.font.name = 'Calibri'
        r_icon.font.size = Pt(11)
        r_icon.font.color.rgb = COLOR_AZURE

        run = p.add_run(title.upper())
        run.bold = True
        run.font.name = 'Calibri'
        run.font.size = Pt(10.5)
        run.font.color.rgb = COLOR_NAVY

        # Subtle bottom border
        pBrd = parse_xml(f'<w:pBrd {nsdecls("w")}><w:bottom w:val="single" w:sz="8" w:space="2" w:color="{HEX_BORDER}"/></w:pBrd>')
        p._p.get_or_add_pPr().append(pBrd)

    # ==========================================
    # 3. PROFESSIONAL SUMMARY
    # ==========================================
    add_section_header("Professional Summary", "📌")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(2)
    p_sum.paragraph_format.space_after = Pt(4)
    p_sum.paragraph_format.line_spacing = 1.15
    r_sum = p_sum.add_run(
        "Motivated Electronics and Communication Engineering undergraduate (2023–2027) with hands-on experience in embedded systems, microcontroller firmware development, railway technology, robotics, and sensor integration. Proficient in ESP32, ESP8266, Arduino, Embedded C/C++, Python, digital electronics, and hardware-software integration. Actively collaborating with a faculty mentor on communication engineering research and technical manuscript development. Experienced in AI-assisted engineering workflows, prompt engineering, and structured technical problem-solving. Seeking opportunities in embedded electronics, communication systems, railway control technologies, and industrial automation."
    )
    r_sum.font.name = 'Calibri'
    r_sum.font.size = Pt(9.5)
    r_sum.font.color.rgb = COLOR_DARK

    # ==========================================
    # 4. EDUCATION
    # ==========================================
    add_section_header("Education", "🎓")
    edu_table = doc.add_table(rows=1, cols=2)
    edu_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    edu_table.autofit = False
    c_e1 = edu_table.cell(0, 0)
    c_e2 = edu_table.cell(0, 1)
    c_e1.width = Inches(5.3)
    c_e2.width = Inches(2.0)
    set_cell_background(c_e1, HEX_CARD_BG)
    set_cell_background(c_e2, HEX_CARD_BG)
    set_cell_margins(c_e1, top=60, bottom=60, left=100, right=100)
    set_cell_margins(c_e2, top=60, bottom=60, left=100, right=100)
    set_cell_border(c_e1, bottom={"val": "single", "sz": "4", "color": HEX_LIGHT_BORDER})
    set_cell_border(c_e2, bottom={"val": "single", "sz": "4", "color": HEX_LIGHT_BORDER})

    p_e1 = c_e1.paragraphs[0]
    p_e1.paragraph_format.space_before = Pt(0)
    p_e1.paragraph_format.space_after = Pt(1)
    r_e1 = p_e1.add_run("Bachelor of Technology (B.Tech) – Electronics and Communication Engineering")
    r_e1.bold = True
    r_e1.font.name = 'Calibri'
    r_e1.font.size = Pt(9.5)
    r_e1.font.color.rgb = COLOR_DARK

    p_e1_sub = c_e1.add_paragraph()
    p_e1_sub.paragraph_format.space_before = Pt(0)
    p_e1_sub.paragraph_format.space_after = Pt(0)
    r_e1_sub = p_e1_sub.add_run("PBR Visvodaya Institute of Technology & Science (PBRVITS), Kavali, Andhra Pradesh")
    r_e1_sub.font.name = 'Calibri'
    r_e1_sub.font.size = Pt(8.5)
    r_e1_sub.font.color.rgb = COLOR_MUTED

    p_e2 = c_e2.paragraphs[0]
    p_e2.paragraph_format.space_before = Pt(0)
    p_e2.paragraph_format.space_after = Pt(1)
    p_e2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_e2 = p_e2.add_run("2023 – 2027")
    r_e2.bold = True
    r_e2.font.name = 'Calibri'
    r_e2.font.size = Pt(9.5)
    r_e2.font.color.rgb = COLOR_STEEL

    p_e2_sub = c_e2.add_paragraph()
    p_e2_sub.paragraph_format.space_before = Pt(0)
    p_e2_sub.paragraph_format.space_after = Pt(0)
    p_e2_sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_e2_sub = p_e2_sub.add_run("CGPA: 8.42 / 10.0\n(First Class with Distinction • No Backlogs)")
    r_e2_sub.font.name = 'Calibri'
    r_e2_sub.font.size = Pt(8)
    r_e2_sub.font.color.rgb = RGBColor(21, 128, 61)
    r_e2_sub.bold = True

    # ==========================================
    # 5. TECHNICAL SKILLS
    # ==========================================
    add_section_header("Technical Skills", "⚡")
    
    skills_data = [
        ("1. Programming & Software: ", "Embedded C, C++, Python, HTML, Data Structures, Problem-Solving & Bare-Metal Firmware Logic, AI-Assisted Coding, Debugging & Technical Documentation."),
        ("2. Embedded Microcontrollers: ", "ESP32 (Dual-Core Xtensa LX6 @ 240MHz), ESP8266, Arduino Uno, NodeMCU, Microcontroller Programming, Hardware Timers & Interrupt Handling, Sensor Interfacing."),
        ("3. Electronics & Hardware: ", "Digital & Analog Electronics, Electronic Prototyping & Circuit Fundamentals, PWM Motor & Servo Control, PCB Design Fundamentals (KiCad / EasyEDA), Multimeter & Oscilloscope Debugging Basics."),
        ("4. Communication Protocols: ", "UART, I2C, SPI, ESP-NOW, CAN Bus & RS-485 Fundamentals (Railway & Industrial Protocols), Communication Systems & Wireless Communication Concepts."),
        ("5. Robotics & Automation: ", "ROS 2 Fundamentals, PCA9685 16-Channel PWM Servo Controller, Ultrasonic Obstacle Detection & Multi-Servo Synchronized Gait Control, Computer Vision (OpenCV / MediaPipe) & Motor Drive Control."),
        ("6. VLSI & Digital Design: ", "Verilog HDL & RTL Design Concepts, Digital Logic Simulation & Verification, Sky130 PDK & OpenLane ASIC Flow."),
        ("7. AI & Prompt Engineering: ", "Structured Prompt Design & Context Management, LLM Interaction & AI-Assisted Engineering Workflows, Workflow Automation Concepts (n8n Fundamentals)."),
        ("8. Development Tools: ", "Arduino IDE, Visual Studio Code, Git and GitHub, Supabase.")
    ]

    for label, desc in skills_data:
        p_sk = doc.add_paragraph()
        p_sk.paragraph_format.space_before = Pt(0)
        p_sk.paragraph_format.space_after = Pt(1.5)
        p_sk.paragraph_format.left_indent = Inches(0.12)
        r_lbl = p_sk.add_run(label)
        r_lbl.bold = True
        r_lbl.font.name = 'Calibri'
        r_lbl.font.size = Pt(9)
        r_lbl.font.color.rgb = COLOR_NAVY
        r_desc = p_sk.add_run(desc)
        r_desc.font.name = 'Calibri'
        r_desc.font.size = Pt(9)
        r_desc.font.color.rgb = COLOR_DARK

    # ==========================================
    # 6. RESEARCH EXPERIENCE
    # ==========================================
    add_section_header("Research Experience", "🔬")
    p_rt = doc.add_paragraph()
    p_rt.paragraph_format.space_before = Pt(1)
    p_rt.paragraph_format.space_after = Pt(0)
    p_rt.paragraph_format.keep_with_next = True
    r_rt = p_rt.add_run("Undergraduate Research Collaboration – Communication Systems")
    r_rt.bold = True
    r_rt.font.name = 'Calibri'
    r_rt.font.size = Pt(9.5)
    r_rt.font.color.rgb = COLOR_DARK

    p_rm = doc.add_paragraph()
    p_rm.paragraph_format.space_before = Pt(0)
    p_rm.paragraph_format.space_after = Pt(2)
    p_rm.paragraph_format.keep_with_next = True
    r_rm = p_rm.add_run("Faculty-Guided Research | PBRVITS, Kavali   •   Research Status: Manuscript in Preparation")
    r_rm.italic = True
    r_rm.font.name = 'Calibri'
    r_rm.font.size = Pt(8.5)
    r_rm.font.color.rgb = COLOR_STEEL

    res_items = [
        "Collaborating with a faculty mentor on technical research in the communication and signal processing domain.",
        "Conducting technical literature reviews, channel characterization, system analysis, and performance evaluations.",
        "Authoring technical documentation and preparing a research manuscript for publication."
    ]
    for item in res_items:
        p_i = doc.add_paragraph(style='List Bullet')
        p_i.paragraph_format.space_before = Pt(0)
        p_i.paragraph_format.space_after = Pt(1)
        p_i.paragraph_format.left_indent = Inches(0.25)
        r_i = p_i.add_run(item)
        r_i.font.name = 'Calibri'
        r_i.font.size = Pt(9)
        r_i.font.color.rgb = COLOR_DARK

    # ==========================================
    # 7. ENGINEERING PROJECTS
    # ==========================================
    add_section_header("Engineering Projects", "🛠")

    projects_list = [
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

    for p_title, p_domain, p_bullets in projects_list:
        p_head = doc.add_paragraph()
        p_head.paragraph_format.space_before = Pt(3)
        p_head.paragraph_format.space_after = Pt(0)
        p_head.paragraph_format.keep_with_next = True
        r_head = p_head.add_run(p_title)
        r_head.bold = True
        r_head.font.name = 'Calibri'
        r_head.font.size = Pt(9.5)
        r_head.font.color.rgb = COLOR_DARK

        p_dom = doc.add_paragraph()
        p_dom.paragraph_format.space_before = Pt(0)
        p_dom.paragraph_format.space_after = Pt(1.5)
        p_dom.paragraph_format.keep_with_next = True
        r_dom = p_dom.add_run(f"Domain: {p_domain}")
        r_dom.italic = True
        r_dom.font.name = 'Calibri'
        r_dom.font.size = Pt(8.5)
        r_dom.font.color.rgb = COLOR_STEEL

        for bullet in p_bullets:
            p_bl = doc.add_paragraph(style='List Bullet')
            p_bl.paragraph_format.space_before = Pt(0)
            p_bl.paragraph_format.space_after = Pt(1)
            p_bl.paragraph_format.left_indent = Inches(0.25)
            r_bl = p_bl.add_run(bullet)
            r_bl.font.name = 'Calibri'
            r_bl.font.size = Pt(9)
            r_bl.font.color.rgb = COLOR_DARK

    # ==========================================
    # 8. HONORS & CERTIFICATIONS
    # ==========================================
    add_section_header("Honors, Awards & Certifications", "🏆")
    honors_items = [
        "Paper Presentation Winner (1st Prize): National-Level Technical Symposium for innovative railway monitoring architecture.",
        "Top 10 Finalist: Regional Embedded & Robotics Hackathon (Track inspection robotics).",
        "Certifications: Embedded Systems with ARM & ESP32  •  Digital VLSI Design Fundamentals (NPTEL)  •  Industrial IoT & Protocols (I2C, SPI, CAN)."
    ]
    for h in honors_items:
        p_h = doc.add_paragraph(style='List Bullet')
        p_h.paragraph_format.space_before = Pt(0)
        p_h.paragraph_format.space_after = Pt(1.5)
        p_h.paragraph_format.left_indent = Inches(0.25)
        r_h = p_h.add_run(h)
        r_h.font.name = 'Calibri'
        r_h.font.size = Pt(9)
        r_h.font.color.rgb = COLOR_DARK

    # ==========================================
    # 9. PERSONAL DETAILS & FORMAL DECLARATION
    # ==========================================
    add_section_header("Personal Details & Declaration", "📋")
    
    p_per = doc.add_paragraph()
    p_per.paragraph_format.space_before = Pt(2)
    p_per.paragraph_format.space_after = Pt(2)
    r_per = p_per.add_run(
        "Father's Name: Boligorla Venkata Ramanaiah   •   Date of Birth: 15th August 2005\n"
        "Languages Known: English, Telugu, Hindi   •   Permanent Address: Kavali, SPSR Nellore District, Andhra Pradesh – 524201, India"
    )
    r_per.font.name = 'Calibri'
    r_per.font.size = Pt(9)
    r_per.font.color.rgb = COLOR_DARK

    p_dec = doc.add_paragraph()
    p_dec.paragraph_format.space_before = Pt(3)
    p_dec.paragraph_format.space_after = Pt(6)
    r_dec = p_dec.add_run("Declaration: I hereby declare that the information provided above is true and accurate to the best of my knowledge.")
    r_dec.italic = True
    r_dec.font.name = 'Calibri'
    r_dec.font.size = Pt(8.5)
    r_dec.font.color.rgb = COLOR_MUTED

    # Signature Block Table
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False
    c_s1 = sig_table.cell(0, 0)
    c_s2 = sig_table.cell(0, 1)
    c_s1.width = Inches(4.0)
    c_s2.width = Inches(3.3)
    set_cell_margins(c_s1, top=40, bottom=40, left=40, right=40)
    set_cell_margins(c_s2, top=40, bottom=40, left=40, right=40)

    p_s1 = c_s1.paragraphs[0]
    p_s1.paragraph_format.space_before = Pt(0)
    p_s1.paragraph_format.space_after = Pt(0)
    r_s1 = p_s1.add_run("Place: Kavali, Andhra Pradesh\nDate: 24th November 2026")
    r_s1.font.name = 'Calibri'
    r_s1.font.size = Pt(9)
    r_s1.font.color.rgb = COLOR_DARK

    p_s2 = c_s2.paragraphs[0]
    p_s2.paragraph_format.space_before = Pt(0)
    p_s2.paragraph_format.space_after = Pt(0)
    p_s2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_s2_sig = p_s2.add_run("BOLIGORLA VENGABABU\n")
    r_s2_sig.bold = True
    r_s2_sig.font.name = 'Calibri'
    r_s2_sig.font.size = Pt(10)
    r_s2_sig.font.color.rgb = COLOR_NAVY
    r_s2_sub = p_s2.add_run("(Applicant Signature)")
    r_s2_sub.italic = True
    r_s2_sub.font.name = 'Calibri'
    r_s2_sub.font.size = Pt(8)
    r_s2_sub.font.color.rgb = COLOR_MUTED

    desktop_path = r"C:\Users\Bolig\Desktop\Boligorla_Vengababu_Resume.docx"
    doc.save(desktop_path)
    print(f"Executive Template Resume saved to Desktop at: {desktop_path}")

    # Also keep a copy in the portfolio project workspace
    workspace_path = r"C:\Users\Bolig\.gemini\antigravity\scratch\esp32-portfolio\Boligorla_Vengababu_Resume.docx"
    doc.save(workspace_path)
    print(f"Also updated in workspace: {workspace_path}")

if __name__ == "__main__":
    build_executive_resume()
