import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def generate_exact_template_resume():
    doc = Document()

    # Page Margins: 0.5 in on all sides for precise template reproduction
    for section in doc.sections:
        section.top_margin = Inches(0.45)
        section.bottom_margin = Inches(0.45)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # Palette from user's template image
    COLOR_NAVY = RGBColor(27, 54, 93)        # #1B365D - Main Header & Title Navy
    COLOR_SUBTITLE = RGBColor(15, 23, 42)    # #0F172A - Dark Title Subtitle
    COLOR_BODY = RGBColor(30, 41, 59)        # #1E293B - Crisp Charcoal Body Text
    COLOR_MUTED = RGBColor(71, 85, 105)      # #475569 - Secondary text
    COLOR_ACCENT = RGBColor(19, 64, 116)     # #134074 - Subtle blue accent
    COLOR_WHITE = RGBColor(255, 255, 255)

    HEX_NAVY = "1B365D"
    HEX_RULE = "1B365D"
    HEX_SIDEBAR_BG = "DCE7EE"                # Exact soft blue-gray sidebar tone from image
    HEX_LIGHT_RULE = "94A3B8"
    HEX_CARD_BG = "F1F5F9"

    def set_cell_shading(cell, fill_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        cell._tc.get_or_add_tcPr().append(tcMar)

    def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
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
    # 1. HEADER SECTION (Exact Template Match)
    # ==========================================
    # YOUR NAME SURNAME (Large Navy Heading)
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(2)
    r_name = p_name.add_run("BOLIGORLA VENGABABU")
    r_name.bold = True
    r_name.font.name = 'Arial'
    r_name.font.size = Pt(24)
    r_name.font.color.rgb = COLOR_NAVY

    # SUBTITLE: ASSISTANT MANAGER equivalent -> ELECTRONICS & COMMUNICATION ENGINEERING UNDERGRADUATE
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(5)
    r_sub = p_sub.add_run("E L E C T R O N I C S   &   C O M M U N I C A T I O N   E N G I N E E R I N G")
    r_sub.bold = True
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(10)
    r_sub.font.color.rgb = COLOR_SUBTITLE

    p_spec = doc.add_paragraph()
    p_spec.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_spec.paragraph_format.space_before = Pt(0)
    p_spec.paragraph_format.space_after = Pt(4)
    r_spec = p_spec.add_run("Embedded Systems   |   Railway Electronics   |   Communication Systems   |   Robotics   |   VLSI")
    r_spec.font.name = 'Calibri'
    r_spec.font.size = Pt(8.5)
    r_spec.font.color.rgb = COLOR_MUTED

    # TOP DIVIDER LINE ACROSS PAGE
    p_top_line = doc.add_paragraph()
    p_top_line.paragraph_format.space_before = Pt(0)
    p_top_line.paragraph_format.space_after = Pt(2)
    pBrd_top = parse_xml(f'<w:pBrd {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="1" w:color="{HEX_NAVY}"/></w:pBrd>')
    p_top_line._p.get_or_add_pPr().append(pBrd_top)

    # CONTACT BADGE BAR (P [Phone]  E Email  A [Address]  W [Website])
    contact_table = doc.add_table(rows=1, cols=4)
    contact_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    contact_table.autofit = False

    widths = [Inches(1.85), Inches(2.2), Inches(1.8), Inches(1.65)]
    contact_items = [
        ("P", "+91 91218 43829"),
        ("E", "boligorlavengababu@gmail.com"),
        ("A", "Kavali, Andhra Pradesh"),
        ("W", "github.com/vengababu-X")
    ]

    for i, (badge, text) in enumerate(contact_items):
        cell = contact_table.cell(0, i)
        cell.width = widths[i]
        set_cell_margins(cell, top=30, bottom=30, left=40, right=40)
        p_c = cell.paragraphs[0]
        p_c.paragraph_format.space_before = Pt(0)
        p_c.paragraph_format.space_after = Pt(0)
        
        # Badge
        r_b = p_c.add_run(f" [{badge}] ")
        r_b.bold = True
        r_b.font.name = 'Arial'
        r_b.font.size = Pt(8.5)
        r_b.font.color.rgb = COLOR_NAVY
        
        # Text
        r_t = p_c.add_run(text)
        r_t.font.name = 'Calibri'
        r_t.font.size = Pt(8)
        r_t.font.color.rgb = COLOR_BODY

    # BOTTOM DIVIDER LINE UNDER CONTACT BAR
    p_bot_line = doc.add_paragraph()
    p_bot_line.paragraph_format.space_before = Pt(2)
    p_bot_line.paragraph_format.space_after = Pt(6)
    pBrd_bot = parse_xml(f'<w:pBrd {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="{HEX_LIGHT_RULE}"/></w:pBrd>')
    p_bot_line._p.get_or_add_pPr().append(pBrd_bot)

    # ==========================================
    # 2. TWO-COLUMN MAIN BODY LAYOUT (Matching Template)
    # ==========================================
    body_table = doc.add_table(rows=1, cols=2)
    body_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    body_table.autofit = False

    cell_main = body_table.cell(0, 0)
    cell_sidebar = body_table.cell(0, 1)

    cell_main.width = Inches(4.85)
    cell_sidebar.width = Inches(2.65)

    # Sidebar has shaded light blue background exactly like in the user's template image
    set_cell_shading(cell_sidebar, HEX_SIDEBAR_BG)
    set_cell_margins(cell_sidebar, top=140, bottom=140, left=140, right=140)
    set_cell_margins(cell_main, top=80, bottom=80, left=40, right=140)

    # Helper for Underlined Headings in Main Column (Just like template image)
    def add_main_heading(cell, text):
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(7)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text.upper())
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(9.5)
        run.font.color.rgb = COLOR_NAVY
        # Bottom underline running across column
        pBrd = parse_xml(f'<w:pBrd {nsdecls("w")}><w:bottom w:val="single" w:sz="8" w:space="2" w:color="{HEX_NAVY}"/></w:pBrd>')
        p._p.get_or_add_pPr().append(pBrd)

    # Helper for Sidebar Headings
    def add_sidebar_heading(cell, text):
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text.upper())
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(9)
        run.font.color.rgb = COLOR_NAVY
        # Subtle bottom underline
        pBrd = parse_xml(f'<w:pBrd {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="{HEX_NAVY}"/></w:pBrd>')
        p._p.get_or_add_pPr().append(pBrd)

    # ------------------------------------------
    # MAIN COLUMN CONTENT (Left Column)
    # ------------------------------------------
    # Clear initial blank paragraph
    p_init = cell_main.paragraphs[0]
    p_init.paragraph_format.space_before = Pt(0)
    p_init.paragraph_format.space_after = Pt(2)
    p_init.paragraph_format.keep_with_next = True
    r_ih = p_init.add_run("DUAL-TARGET OBJECTIVE (MEDHA & ECIL)")
    r_ih.bold = True
    r_ih.font.name = 'Arial'
    r_ih.font.size = Pt(9.5)
    r_ih.font.color.rgb = COLOR_NAVY
    pBrd_ih = parse_xml(f'<w:pBrd {nsdecls("w")}><w:bottom w:val="single" w:sz="8" w:space="2" w:color="{HEX_NAVY}"/></w:pBrd>')
    p_init._p.get_or_add_pPr().append(pBrd_ih)

    p_obj = cell_main.add_paragraph()
    p_obj.paragraph_format.space_before = Pt(2)
    p_obj.paragraph_format.space_after = Pt(4)
    r_obj = p_obj.add_run(
        "Dedicated Electronics and Communication Engineering undergraduate seeking an entry-level engineering role in Locomotive & Railway Control Systems (Medha Servo Drives Pvt. Ltd.) and Strategic Electronics, Defense Communications & Instrumentation (ECIL Hyderabad). Offering proven hands-on capability in embedded C/C++ firmware, railway rolling-stock monitoring concepts, CAN bus/RS-485 industrial protocols, faculty-guided communication systems research, and Verilog RTL digital design."
    )
    r_obj.font.name = 'Calibri'
    r_obj.font.size = Pt(8.5)
    r_obj.font.color.rgb = COLOR_BODY

    # PROFESSIONAL SUMMARY
    add_main_heading(cell_main, "Professional Summary")
    p_sum = cell_main.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(2)
    p_sum.paragraph_format.space_after = Pt(4)
    r_sum = p_sum.add_run(
        "Motivated Electronics and Communication Engineering undergraduate (2023–2027) with hands-on experience in embedded systems, microcontroller firmware development, railway technology, robotics, and sensor integration. Proficient in ESP32, ESP8266, Arduino, Embedded C/C++, Python, digital electronics, and hardware-software integration. Actively collaborating with a faculty mentor on communication engineering research and technical manuscript development. Experienced in AI-assisted engineering workflows, prompt engineering, and structured technical problem-solving. Seeking opportunities in embedded electronics, communication systems, railway control technologies, and industrial automation."
    )
    r_sum.font.name = 'Calibri'
    r_sum.font.size = Pt(8.5)
    r_sum.font.color.rgb = COLOR_BODY

    # RESEARCH EXPERIENCE
    add_main_heading(cell_main, "Research Experience")
    p_rx_title = cell_main.add_paragraph()
    p_rx_title.paragraph_format.space_before = Pt(2)
    p_rx_title.paragraph_format.space_after = Pt(0)
    p_rx_title.paragraph_format.keep_with_next = True
    r_rxt = p_rx_title.add_run("Undergraduate Research Collaboration – Communication Systems")
    r_rxt.bold = True
    r_rxt.font.name = 'Arial'
    r_rxt.font.size = Pt(9)
    r_rxt.font.color.rgb = COLOR_BODY

    p_rx_meta = cell_main.add_paragraph()
    p_rx_meta.paragraph_format.space_before = Pt(0)
    p_rx_meta.paragraph_format.space_after = Pt(1.5)
    p_rx_meta.paragraph_format.keep_with_next = True
    r_rxm = p_rx_meta.add_run("Faculty-Guided Research | PBRVITS, Kavali  •  Status: Manuscript in Preparation")
    r_rxm.italic = True
    r_rxm.font.name = 'Calibri'
    r_rxm.font.size = Pt(8)
    r_rxm.font.color.rgb = COLOR_MUTED

    rx_bullets = [
        "Collaborating with a faculty mentor on technical research in communication & signal processing.",
        "Conducting technical literature reviews, channel characterization, system analysis, and evaluations.",
        "Authoring technical documentation and preparing a research manuscript for publication."
    ]
    for b in rx_bullets:
        p_b = cell_main.add_paragraph(style='List Bullet')
        p_b.paragraph_format.space_before = Pt(0)
        p_b.paragraph_format.space_after = Pt(1)
        p_b.paragraph_format.left_indent = Inches(0.18)
        r_b = p_b.add_run(b)
        r_b.font.name = 'Calibri'
        r_b.font.size = Pt(8.5)
        r_b.font.color.rgb = COLOR_BODY

    # ENGINEERING PROJECTS (All 5 Projects with Zero Missing Fields)
    add_main_heading(cell_main, "Engineering Projects")

    projects = [
        (
            "1. RAILGUARD-X – Railway Security Robotics System",
            "Railway Technology | Robotics | Embedded Systems | ROS 2",
            [
                "Engineered a quadruped robotic prototype for railway track surveillance and perimeter security.",
                "Interfaced ESP8266 with PCA9685 16-channel PWM servo driver over I2C for synchronized gait.",
                "Integrated ultrasonic sensors for real-time obstacle detection and emergency halt fail-safe.",
                "Designed modular architecture utilizing ROS 2 fundamentals for robotic communication & telemetry."
            ]
        ),
        (
            "2. ITMS – Intelligent Train Monitoring System",
            "Railway Technology | Rolling Stock Monitoring | CAN / RS-485",
            [
                "Designed intelligent railway monitoring concept for rolling-stock safety and thermal supervision.",
                "Evaluated embedded sensor integration, fault-alert logic, and telemetry workflows.",
                "Applied industrial automation principles with CAN Bus and RS-485 noise-immune topologies."
            ]
        ),
        (
            "3. Smart Wearable Safety Monitoring System",
            "IoT | Embedded Electronics | Sensor Integration | Supabase",
            [
                "Developed ESP32-based wearable integrating environmental, motion (MPU6050), and health sensors.",
                "Implemented low-latency ESP-NOW peer-to-peer wireless communication between ESP32 & ESP8266 nodes.",
                "Designed automated alert mechanisms and connected cloud telemetry logging using Supabase."
            ]
        ),
        (
            "4. Eye-Controlled Smart Wheelchair System",
            "Assistive Robotics | Embedded Systems | Computer Vision & IR",
            [
                "Engineered hands-free eye-controlled wheelchair for mobility-impaired individuals.",
                "Processed eye-gaze and blink patterns via OpenCV/MediaPipe for differential dual-motor PWM.",
                "Built hardware safety interlocks with ultrasonic sensors to trigger braking within 30 cm."
            ]
        ),
        (
            "5. Y2K 11-Bit Synchronous Counter (VLSI / Digital Design)",
            "Verilog HDL | Digital Electronics | OpenLane ASIC Simulation",
            [
                "Designed, synthesized, and verified an 11-bit synchronous digital counter in Verilog HDL.",
                "Implemented synchronous clear, clock-enable gating, and overflow ripple-carry detection flags.",
                "Simulated functional waveforms with zero setup/hold timing violations for digital architectures."
            ]
        )
    ]

    for title, domain, bullets in projects:
        p_pt = cell_main.add_paragraph()
        p_pt.paragraph_format.space_before = Pt(2.5)
        p_pt.paragraph_format.space_after = Pt(0)
        p_pt.paragraph_format.keep_with_next = True
        r_pt = p_pt.add_run(title)
        r_pt.bold = True
        r_pt.font.name = 'Arial'
        r_pt.font.size = Pt(8.5)
        r_pt.font.color.rgb = COLOR_BODY

        p_pd = cell_main.add_paragraph()
        p_pd.paragraph_format.space_before = Pt(0)
        p_pd.paragraph_format.space_after = Pt(1)
        p_pd.paragraph_format.keep_with_next = True
        r_pd = p_pd.add_run(domain)
        r_pd.italic = True
        r_pd.font.name = 'Calibri'
        r_pd.font.size = Pt(7.8)
        r_pd.font.color.rgb = COLOR_ACCENT

        for b in bullets:
            p_bl = cell_main.add_paragraph(style='List Bullet')
            p_bl.paragraph_format.space_before = Pt(0)
            p_bl.paragraph_format.space_after = Pt(0.8)
            p_bl.paragraph_format.left_indent = Inches(0.18)
            r_bl = p_bl.add_run(b)
            r_bl.font.name = 'Calibri'
            r_bl.font.size = Pt(8.2)
            r_bl.font.color.rgb = COLOR_BODY

    # PERSONAL DETAILS & FORMAL DECLARATION
    add_main_heading(cell_main, "Personal Details & Declaration")
    p_per = cell_main.add_paragraph()
    p_per.paragraph_format.space_before = Pt(1.5)
    p_per.paragraph_format.space_after = Pt(1.5)
    r_per = p_per.add_run(
        "Father's Name: Boligorla Venkata Ramanaiah  •  DOB: 15 Aug 2005\n"
        "Languages: English, Telugu, Hindi  •  Address: Kavali, Andhra Pradesh"
    )
    r_per.font.name = 'Calibri'
    r_per.font.size = Pt(8)
    r_per.font.color.rgb = COLOR_BODY

    p_dec = cell_main.add_paragraph()
    p_dec.paragraph_format.space_before = Pt(1)
    p_dec.paragraph_format.space_after = Pt(3)
    r_dec = p_dec.add_run("Declaration: I hereby declare that the information provided above is true and accurate to the best of my knowledge.")
    r_dec.italic = True
    r_dec.font.name = 'Calibri'
    r_dec.font.size = Pt(7.8)
    r_dec.font.color.rgb = COLOR_MUTED

    p_sig = cell_main.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(0)
    p_sig.paragraph_format.space_after = Pt(0)
    r_sig1 = p_sig.add_run("Place: Kavali, AP  •  Date: 24 Nov 2026           ")
    r_sig1.font.name = 'Calibri'
    r_sig1.font.size = Pt(8)
    r_sig1.font.color.rgb = COLOR_BODY
    r_sig2 = p_sig.add_run("BOLIGORLA VENGABABU")
    r_sig2.bold = True
    r_sig2.font.name = 'Arial'
    r_sig2.font.size = Pt(8.5)
    r_sig2.font.color.rgb = COLOR_NAVY

    # ------------------------------------------
    # SIDEBAR CONTENT (Right Column with Light Blue Shading)
    # ------------------------------------------
    # Clear initial blank paragraph
    p_sb_init = cell_sidebar.paragraphs[0]
    p_sb_init.paragraph_format.space_before = Pt(0)
    p_sb_init.paragraph_format.space_after = Pt(2)
    p_sb_init.paragraph_format.keep_with_next = True
    r_sbi = p_sb_init.add_run("EDUCATION")
    r_sbi.bold = True
    r_sbi.font.name = 'Arial'
    r_sbi.font.size = Pt(9)
    r_sbi.font.color.rgb = COLOR_NAVY
    pBrd_sbi = parse_xml(f'<w:pBrd {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="{HEX_NAVY}"/></w:pBrd>')
    p_sb_init._p.get_or_add_pPr().append(pBrd_sbi)

    p_s_sch = cell_sidebar.add_paragraph()
    p_s_sch.paragraph_format.space_before = Pt(2)
    p_s_sch.paragraph_format.space_after = Pt(0)
    r_ssc = p_s_sch.add_run("PBR Visvodaya Institute of Technology & Science (PBRVITS)")
    r_ssc.bold = True
    r_ssc.font.name = 'Arial'
    r_ssc.font.size = Pt(8.2)
    r_ssc.font.color.rgb = COLOR_NAVY

    p_s_loc = cell_sidebar.add_paragraph()
    p_s_loc.paragraph_format.space_before = Pt(0)
    p_s_loc.paragraph_format.space_after = Pt(1)
    r_sloc = p_s_loc.add_run("Kavali, Andhra Pradesh, India")
    r_sloc.font.name = 'Calibri'
    r_sloc.font.size = Pt(8)
    r_sloc.font.color.rgb = COLOR_MUTED

    p_s_deg = cell_sidebar.add_paragraph()
    p_s_deg.paragraph_format.space_before = Pt(0)
    p_s_deg.paragraph_format.space_after = Pt(0)
    r_sdeg = p_s_deg.add_run("B.Tech – Electronics & Communication Engineering (2023–2027)")
    r_sdeg.font.name = 'Calibri'
    r_sdeg.font.size = Pt(8)
    r_sdeg.font.color.rgb = COLOR_BODY

    p_s_gpa = cell_sidebar.add_paragraph()
    p_s_gpa.paragraph_format.space_before = Pt(0)
    p_s_gpa.paragraph_format.space_after = Pt(4)
    r_sgpa = p_s_gpa.add_run("CGPA: 8.42 / 10.0\n(First Class with Distinction • No Active Backlogs)")
    r_sgpa.bold = True
    r_sgpa.font.name = 'Calibri'
    r_sgpa.font.size = Pt(8)
    r_sgpa.font.color.rgb = RGBColor(15, 118, 54)

    # DUAL TARGET FOCUS
    add_sidebar_heading(cell_sidebar, "Target Focus")
    p_tf_m = cell_sidebar.add_paragraph()
    p_tf_m.paragraph_format.space_before = Pt(2)
    p_tf_m.paragraph_format.space_after = Pt(0)
    r_tfm_h = p_tf_m.add_run("Medha Servo Drives:")
    r_tfm_h.bold = True
    r_tfm_h.font.name = 'Arial'
    r_tfm_h.font.size = Pt(8)
    r_tfm_h.font.color.rgb = RGBColor(21, 128, 61)
    
    p_tf_mb = cell_sidebar.add_paragraph()
    p_tf_mb.paragraph_format.space_before = Pt(0)
    p_tf_mb.paragraph_format.space_after = Pt(3)
    r_tfmb = p_tf_mb.add_run("Railway Traction, ITMS Rolling Stock, PCA9685 PWM, CAN Bus & RS-485.")
    r_tfmb.font.name = 'Calibri'
    r_tfmb.font.size = Pt(7.8)
    r_tfmb.font.color.rgb = COLOR_BODY

    p_tf_e = cell_sidebar.add_paragraph()
    p_tf_e.paragraph_format.space_before = Pt(0)
    p_tf_e.paragraph_format.space_after = Pt(0)
    r_tfe_h = p_tf_e.add_run("ECIL Hyderabad:")
    r_tfe_h.bold = True
    r_tfe_h.font.name = 'Arial'
    r_tfe_h.font.size = Pt(8)
    r_tfe_h.font.color.rgb = RGBColor(29, 78, 216)

    p_tf_eb = cell_sidebar.add_paragraph()
    p_tf_eb.paragraph_format.space_before = Pt(0)
    p_tf_eb.paragraph_format.space_after = Pt(4)
    r_tfeb = p_tf_eb.add_run("Communication Systems Research, ESP-NOW RF, Verilog RTL, Sky130 ASIC.")
    r_tfeb.font.name = 'Calibri'
    r_tfeb.font.size = Pt(7.8)
    r_tfeb.font.color.rgb = COLOR_BODY

    # KEY SKILLS (Exact Match to Sidebar in Template)
    add_sidebar_heading(cell_sidebar, "Key Skills")
    skills_sidebar = [
        "Embedded C & C++",
        "Python & Bare-Metal Logic",
        "ESP32 & ESP8266 Firmware",
        "Arduino Uno & NodeMCU",
        "CAN Bus & RS-485 Protocols",
        "UART, I2C, SPI, ESP-NOW",
        "PCA9685 16-Ch PWM Servo",
        "ROS 2 Robotic Fundamentals",
        "Ultrasonic Sensor Interfacing",
        "Computer Vision (OpenCV)",
        "Verilog HDL & RTL Design",
        "Sky130 PDK & OpenLane",
        "Digital & Analog Circuits",
        "KiCad / EasyEDA PCB Layout",
        "Multimeter & Oscilloscope",
        "Git & GitHub Workflows",
        "Supabase Cloud Telemetry",
        "Prompt Engineering & n8n"
    ]
    for sk in skills_sidebar:
        p_skb = cell_sidebar.add_paragraph(style='List Bullet')
        p_skb.paragraph_format.space_before = Pt(0)
        p_skb.paragraph_format.space_after = Pt(0.7)
        p_skb.paragraph_format.left_indent = Inches(0.14)
        r_skb = p_skb.add_run(sk)
        r_skb.font.name = 'Calibri'
        r_skb.font.size = Pt(8)
        r_skb.font.color.rgb = COLOR_BODY

    # AWARDS & CERTIFICATIONS
    add_sidebar_heading(cell_sidebar, "Awards & Certs")
    awards_sidebar = [
        "1st Prize: National Paper Presentation (Railway Architecture)",
        "Top 10 Finalist: Regional Robotics Hackathon (Track Inspection)",
        "Cert: Embedded Systems with ARM & ESP32",
        "Cert: Digital VLSI Design Fundamentals (NPTEL)",
        "Cert: Industrial IoT & Protocols (I2C, SPI, CAN)"
    ]
    for aw in awards_sidebar:
        p_aw = cell_sidebar.add_paragraph(style='List Bullet')
        p_aw.paragraph_format.space_before = Pt(0)
        p_aw.paragraph_format.space_after = Pt(1)
        p_aw.paragraph_format.left_indent = Inches(0.14)
        r_aw = p_aw.add_run(aw)
        r_aw.font.name = 'Calibri'
        r_aw.font.size = Pt(7.8)
        r_aw.font.color.rgb = COLOR_BODY

    # LINKS & PORTFOLIO
    add_sidebar_heading(cell_sidebar, "Portfolio & Code")
    p_links = cell_sidebar.add_paragraph()
    p_links.paragraph_format.space_before = Pt(1)
    p_links.paragraph_format.space_after = Pt(0)
    r_l1 = p_links.add_run("Web Portfolio:\nesp32-portfolio-vengababu.vercel.app\n\nGitHub:\ngithub.com/vengababu-X\n\nLinkedIn:\nlinkedin.com/in/boligorla-vengababu-503729349/")
    r_l1.font.name = 'Calibri'
    r_l1.font.size = Pt(7.5)
    r_l1.font.color.rgb = COLOR_MUTED

    # BOTTOM FOOTER DIVIDER ACROSS THE PAGE
    p_foot_line = doc.add_paragraph()
    p_foot_line.paragraph_format.space_before = Pt(6)
    p_foot_line.paragraph_format.space_after = Pt(0)
    pBrd_foot = parse_xml(f'<w:pBrd {nsdecls("w")}><w:bottom w:val="single" w:sz="8" w:space="1" w:color="{HEX_NAVY}"/></w:pBrd>')
    p_foot_line._p.get_or_add_pPr().append(pBrd_foot)

    # OUTPUT TARGETS:
    # 1. Desktop primary filename
    out_desktop_new = r"C:\Users\Bolig\Desktop\Boligorla_Vengababu_Resume_Template.docx"
    doc.save(out_desktop_new)
    print(f"Saved to: {out_desktop_new}")

    # 2. Try overwriting C:\Users\Bolig\Desktop\Boligorla_Vengababu_Resume.docx (if unlocked)
    out_desktop_orig = r"C:\Users\Bolig\Desktop\Boligorla_Vengababu_Resume.docx"
    try:
        doc.save(out_desktop_orig)
        print(f"Successfully updated original: {out_desktop_orig}")
    except Exception as e:
        print(f"Original file is currently open in Word ({e}). Saved as {out_desktop_new}")

    # 3. Save to workspace
    out_ws = r"C:\Users\Bolig\.gemini\antigravity\scratch\esp32-portfolio\Boligorla_Vengababu_Resume.docx"
    doc.save(out_ws)
    print(f"Saved to workspace: {out_ws}")

if __name__ == "__main__":
    generate_exact_template_resume()
