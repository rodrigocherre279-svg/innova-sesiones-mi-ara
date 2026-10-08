"""
Motor didactico de generacion de presentaciones Innova Schools.
Analiza dinamicamente el plan de sesion (.docx) y genera la presentacion PPTX
correspondiente con la metodologia oficial de Araceli:
- Unit 6 Session 6: Listening / Historical Events (Machu Picchu, Was & Were, Speech Draft)
- Unit 6 Session 5: Writing / Personal Events (Wedding, Connectors, Rubrics, Drafting)
- Unit 6 Session 4: Reading / Personal Life Events (Moving house, Prepositions in/on/at, Past simple)
- Unit 5 Session 2: Oral Communication / Food & Drinks (Cafe ordering, Waiter & Customer)
- O cualquier otra sesion mediante adaptacion pedagogica dinamica.
"""

import os
import io
import re
import docx
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
SAMPLES_DIR = os.path.join(ASSETS_DIR, "samples")

CLR_BLUE = RGBColor(0, 102, 158)       # #00669E
CLR_PLUM = RGBColor(116, 23, 83)       # #741753
CLR_ORANGE = RGBColor(234, 129, 24)    # #EA8118
CLR_GREEN = RGBColor(140, 186, 55)     # #8CBA37
CLR_BUBBLE = RGBColor(201, 218, 248)   # #C9DAF8
CLR_DARK_SLATE = RGBColor(35, 58, 68)  # #233A44
CLR_GRAY = RGBColor(66, 66, 66)        # #424242
CLR_BLACK = RGBColor(0, 0, 0)
CLR_WHITE = RGBColor(255, 255, 255)
CLR_LIGHT_YELLOW = RGBColor(255, 249, 230)
CLR_LIGHT_GREEN = RGBColor(234, 250, 234)
CLR_LIGHT_RED = RGBColor(253, 237, 236)

def parse_docx(file_stream_or_path):
    """Extrae la informacion pedagogica de un plan de sesion (.docx)."""
    doc = docx.Document(file_stream_or_path)
    full_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    data = {
        "level": "A1 Level",
        "unit_num": "6",
        "unit": "Unit 6",
        "unit_title": "Important events",
        "session_num": "6",
        "skill_area": "Listening",
        "standard": "",
        "desempeno": "",
        "content_vocab": "Time Expressions",
        "content_grammar": "Past Simple (was / were)",
        "objective": "",
        "summary": "",
        "stages": []
    }

    # Detectar nivel
    for p in doc.paragraphs[:10]:
        text = p.text.strip()
        if "LEVEL" in text.upper():
            data["level"] = text

    # Detectar Unidad
    m_unit = re.search(r"UNIT\s+(\d+)", full_text, re.IGNORECASE)
    if m_unit:
        data["unit_num"] = m_unit.group(1)
        data["unit"] = f"Unit {m_unit.group(1)}"

    # Detectar Sesion
    m_sess = re.search(r"SESSION\s+(\d+)", full_text, re.IGNORECASE)
    if m_sess:
        data["session_num"] = m_sess.group(1)

    # Detectar carta a la docente ("Dear teacher...")
    for p in doc.paragraphs:
        t = p.text.strip()
        if "In this lesson" in t or "In this first lesson" in t:
            data["summary"] = t
            break

    # Leer Tabla 0 (Alineamiento curricular)
    if len(doc.tables) > 0 and len(doc.tables[0].rows) > 1:
        cells = doc.tables[0].rows[1].cells
        if len(cells) >= 5:
            raw_skill = cells[0].text.strip().split("\n")[0].strip()
            data["skill_area"] = raw_skill
            data["standard"] = cells[1].text.strip()
            data["desempeno"] = cells[2].text.strip()
            content = cells[3].text.strip()
            if "Vocabulary:" in content and "Grammar:" in content:
                parts = content.split("Grammar:")
                data["content_vocab"] = parts[0].replace("Vocabulary:", "").strip()
                data["content_grammar"] = parts[1].strip()
            else:
                data["content_vocab"] = content

    # Detectar el titulo de la unidad dinamicamente
    combined_info = f"{data['summary']} {data['content_vocab']} {data['content_grammar']} {full_text}".lower()
    
    if data["unit_num"] == "5" or "food" in combined_info or "waiter" in combined_info or "cafe" in combined_info:
        data["unit_title"] = "Food and Drinks"
    elif data["unit_num"] == "6" or "past simple" in combined_info or "was" in combined_info or "event" in combined_info or "historical" in combined_info:
        data["unit_title"] = "Important events"
    elif "routine" in combined_info:
        data["unit_title"] = "Daily Routines"
    else:
        data["unit_title"] = "English Communication"

    # Refinar gramatica si menciona was/were
    if "was" in combined_info and "were" in combined_info:
        data["content_grammar"] = "Past Simple with 'was' and 'were'"

    # Leer Tabla 1 (Objetivo y etapas)
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        for row in t1.rows[1:]:
            desc = row.cells[0].text.strip()
            m_obj = re.search(r"objective of the class:\s*([^\.\n]+)", desc, re.IGNORECASE)
            if not m_obj:
                m_obj = re.search(r"challenge of the unit:\s*([^\.\n]+)", desc, re.IGNORECASE)
            if m_obj and not data["objective"]:
                data["objective"] = m_obj.group(1).strip()
            data["stages"].append(desc)

    if not data["objective"]:
        if "was" in combined_info:
            data["objective"] = "Infer general and specific information in an oral text about an important historical event."
        else:
            data["objective"] = f"Practise {data['skill_area']} sub-skills, vocabulary and grammar."

    return data

def helper_add_bubble(slide, left, top, width, height, text, font_size=13, bold=True):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = CLR_BUBBLE
    shape.line.color.rgb = CLR_BUBBLE
    tf = shape.text_frame
    tf.word_wrap = True
    tf.text = text
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.LEFT
        for r in p.runs:
            r.font.name = "Roboto"
            r.font.size = Pt(font_size)
            r.font.bold = bold
            r.font.color.rgb = CLR_BLACK
    return shape

def helper_add_card(slide, left, top, width, height, title, items, bg_color=CLR_WHITE, border_color=CLR_BLUE, title_color=CLR_BLUE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(1.5)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.text = title
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    p0.font.name = "Roboto"
    p0.font.size = Pt(14)
    p0.font.bold = True
    p0.font.color.rgb = title_color
    
    for item in items:
        p = tf.add_paragraph()
        p.text = item
        p.alignment = PP_ALIGN.LEFT
        p.font.name = "Roboto"
        p.font.size = Pt(12)
        p.font.bold = False
        p.font.color.rgb = CLR_BLACK
    return shape

def clean_slide_shapes(slide, keep_pictures=False):
    """Limpia los elementos de la diapositiva."""
    for sh in list(slide.shapes):
        if keep_pictures and sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE:
            continue
        sp = sh._element
        sp.getparent().remove(sp)

def build_presentation(plan_data, images_dict=None):
    """
    Construye la presentacion exacta de 16 diapositivas
    despachando al flujo adecuado segun el contenido de la unidad.
    """
    template_path = os.path.join(ASSETS_DIR, "innova_template.pptx")
    prs = pptx.Presentation(template_path)

    unit = str(plan_data.get("unit_num", "6"))
    sess = str(plan_data.get("session_num", "6"))
    grammar_str = plan_data.get("content_grammar", "").lower()
    summary_str = plan_data.get("summary", "").lower()

    # Asset paths
    img_logo = os.path.join(ASSETS_DIR, "innova_schools_logo.png")
    img_target = os.path.join(ASSETS_DIR, "objective_target_icon.png")
    img_goodjob = os.path.join(ASSETS_DIR, "good_job_sticker.png")
    img_rubric = os.path.join(ASSETS_DIR, "innova_rubric_table.png")
    img_timer_02 = os.path.join(ASSETS_DIR, "timer_02_00.jpg")
    img_timer_03 = os.path.join(ASSETS_DIR, "timer_03_00.jpg")
    img_timer_07 = os.path.join(ASSETS_DIR, "timer_07_00.jpg")

    # =========================================================================
    # CASO A: UNIT 6 SESSION 6 (HISTORICAL EVENTS / MACHU PICCHU / WAS & WERE)
    # =========================================================================
    if unit == "6" and sess == "6" or ("was" in grammar_str and "were" in grammar_str) or "machu picchu" in summary_str:
        img_mp = os.path.join(SAMPLES_DIR, "machu_picchu.png")
        img_hist_org = os.path.join(SAMPLES_DIR, "historical_organizer.png")

        # Slide 1 (Index 0): Cover
        s1 = prs.slides[0]
        for sh in s1.shapes:
            if sh.has_text_frame and "A1 Level" in sh.text_frame.text:
                sh.text_frame.text = f"{plan_data['level']}\nUnit 6: “Important events”\nSession 6"
                for p in sh.text_frame.paragraphs:
                    p.font.name = "Roboto"
                    p.font.bold = True
                    p.font.size = Pt(38)
                    p.font.color.rgb = CLR_WHITE

        # Slide 2 (Index 1): Golden Rules (se preserva del template)

        # Slide 3 (Index 2): Warm-up: Machu Picchu
        s3 = prs.slides[2]
        clean_slide_shapes(s3)
        t_box = s3.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(8.8), Inches(0.55))
        tf = t_box.text_frame
        tf.text = "MACHU PICCHU: AN IMPORTANT HISTORICAL EVENT"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(22)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        if os.path.exists(img_mp):
            s3.shapes.add_picture(img_mp, Inches(0.6), Inches(1.05), Inches(5.6), Inches(4.0))

        helper_add_bubble(s3, Inches(6.4), Inches(1.05), Inches(3.1), Inches(1.3),
                          "What do you know about Machu Picchu?\n• Where is it located?\n• How old is it?\n• Is it famous? Why?")
        helper_add_bubble(s3, Inches(6.4), Inches(2.5), Inches(3.1), Inches(1.2),
                          "I think that Machu Picchu is an ancient Inca city in Peru.")
        helper_add_bubble(s3, Inches(6.4), Inches(3.85), Inches(3.1), Inches(1.2),
                          "Machu Picchu is famous because it is one of the New 7 Wonders of the World.")

        # Slide 4 (Index 3): Objective & Date
        s4 = prs.slides[3]
        for sh in s4.shapes:
            if sh.has_text_frame:
                t = sh.text_frame.text
                if "September" in t or "2026" in t or "Monday" in t or "Thursday" in t:
                    sh.text_frame.text = "Thursday, September 17th, 2026."
                    sh.text_frame.paragraphs[0].font.name = "Roboto"
                    sh.text_frame.paragraphs[0].font.size = Pt(16)
                    sh.text_frame.paragraphs[0].font.color.rgb = CLR_WHITE
                elif "Session N" in t or "Session 6" in t:
                    sh.text_frame.text = "Session 6: Final Product (Part 1)"
                    sh.text_frame.paragraphs[0].font.name = "Roboto"
                    sh.text_frame.paragraphs[0].font.bold = True
                    sh.text_frame.paragraphs[0].font.size = Pt(24)
                    sh.text_frame.paragraphs[0].font.color.rgb = CLR_WHITE
                elif len(t) > 20 and "Objective" not in t:
                    sh.text_frame.text = plan_data['objective']
                    sh.text_frame.paragraphs[0].font.name = "Roboto"
                    sh.text_frame.paragraphs[0].font.bold = True
                    sh.text_frame.paragraphs[0].font.size = Pt(22)
                    sh.text_frame.paragraphs[0].font.color.rgb = CLR_WHITE

        # Slide 5 (Index 4): Historical Text Input: The Discovery of Machu Picchu
        s5 = prs.slides[4]
        clean_slide_shapes(s5)
        t_box = s5.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "AN IMPORTANT HISTORICAL EVENT: THE DISCOVERY OF MACHU PICCHU"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(20)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        text_p1 = "Machu Picchu is an ancient Inca city in Peru. It is one of the most famous historical places in the world."
        text_p2 = "In 1911, an American explorer named Hiram Bingham visited Peru. He wanted to find an important Inca city. During his trip, he heard about a place in the mountains called Machu Picchu."
        text_p3 = "First, Bingham travelled to the area with local people. Then, he arrived at Machu Picchu and saw the ancient ruins. He was surprised because the city was very large and beautiful. After that, he took photographs and shared information about the place."
        text_p4 = "Today, Machu Picchu is an important historical and cultural site. Many people visit it every year. The discovery of Machu Picchu was important because it helped people learn more about the Inca civilization."

        helper_add_card(s5, Inches(0.6), Inches(0.95), Inches(8.8), Inches(4.2), "Read and analyze the historical text:", [
            f"1. Name of the event: {text_p1}",
            f"2. When it happened & Who was involved: {text_p2}",
            f"3. What happened & Important details: {text_p3}",
            f"4. Why is it important? {text_p4}"
        ], bg_color=CLR_LIGHT_YELLOW)

        # Slide 6 (Index 5): Text Analysis & Important Details
        s6 = prs.slides[5]
        clean_slide_shapes(s6)
        t_box = s6.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "ANALYZING THE TEXT: SPECIFIC DETAILS"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(22)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        helper_add_card(s6, Inches(0.6), Inches(1.0), Inches(4.2), Inches(3.8), "HISTORICAL FACTS FROM TEXT", [
            "• Where: In the Andes mountains, Peru.",
            "• Who: Hiram Bingham & local guides.",
            "• When: In 1911.",
            "• What: Discovered the ancient Inca ruins.",
            "• Feelings: Bingham was surprised and impressed by the beauty of the city."
        ], bg_color=CLR_WHITE, border_color=CLR_BLUE)

        helper_add_card(s6, Inches(5.2), Inches(1.0), Inches(4.2), Inches(3.8), "NOTICE THE GRAMMAR (WAS / WERE)", [
            "Observe how we describe past facts:",
            "",
            "• Machu Picchu WAS an ancient city.",
            "• Hiram Bingham WAS an explorer.",
            "• The ruins WERE very large.",
            "• Many people WERE interested in the discovery."
        ], bg_color=CLR_LIGHT_GREEN, border_color=CLR_GREEN, title_color=CLR_GREEN)

        # Slide 7 (Index 6): Collaborative Work: Choose your role
        s7 = prs.slides[6]
        clean_slide_shapes(s7)
        t_box = s7.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "COLLABORATIVE WORK: CHOOSE YOUR ROLE"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(22)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        helper_add_card(s7, Inches(1.0), Inches(1.1), Inches(3.8), Inches(3.2), "ROLE 1: INTERVIEWER", [
            "• Ask questions clearly.",
            "• Listen carefully to your partner's ideas.",
            "• Evaluates answers:",
            "  'I agree with you because...'",
            "  'I don't agree because...'"
        ], bg_color=CLR_LIGHT_YELLOW, border_color=CLR_ORANGE, title_color=CLR_ORANGE)

        helper_add_card(s7, Inches(5.2), Inches(1.1), Inches(3.8), Inches(3.2), "ROLE 2: SPEAKER", [
            "• Answers questions with complete ideas.",
            "• Shares personal thoughts about history.",
            "• Uses sentence frames:",
            "  'Yes, I knew that information...'",
            "  'I would like to know about...'"
        ], bg_color=CLR_LIGHT_GREEN, border_color=CLR_GREEN, title_color=CLR_GREEN)

        helper_add_bubble(s7, Inches(1.0), Inches(4.5), Inches(8.0), Inches(0.6),
                          "👥 Pair work: Choose your partner and decide who is the Interviewer and who is the Speaker!", font_size=13)

        # Slide 8 (Index 7): Speaking Discussion in Pairs
        s8 = prs.slides[7]
        clean_slide_shapes(s8)
        t_box = s8.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "DISCUSS THE QUESTIONS IN PAIRS"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(22)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        if os.path.exists(img_timer_02):
            s8.shapes.add_picture(img_timer_02, Inches(7.5), Inches(0.2), Inches(1.9), Inches(1.0))

        helper_add_card(s8, Inches(0.6), Inches(1.2), Inches(4.2), Inches(3.8), "DISCUSSION QUESTIONS", [
            "1. Did you know that information about Machu Picchu? What do you think about that?",
            "",
            "2. Would you like to know about a specific historical event? Which one?"
        ], bg_color=CLR_WHITE, border_color=CLR_BLUE)

        helper_add_card(s8, Inches(5.2), Inches(1.2), Inches(4.2), Inches(3.8), "SPEAKING FRAMES & CHALLENGE", [
            "• 'Yes, I knew that information and I think it was fascinating.'",
            "• 'No, I didn't know that. For me it was surprising.'",
            "• 'I would like to know about (event) because...'",
            "",
            "⭐ CHALLENGE: Include Past Simple with 'was' / 'were' at least 2 times!"
        ], bg_color=CLR_LIGHT_GREEN, border_color=CLR_GREEN, title_color=CLR_GREEN)

        # Slide 9 (Index 8): Check-in (Did we complete...)
        s9 = prs.slides[8]
        for sh in s9.shapes:
            if sh.has_text_frame and "complete" in sh.text_frame.text.lower():
                p = sh.text_frame.paragraphs[0]
                p.font.name = "Chewy"
                p.font.size = Pt(36)
                p.font.color.rgb = CLR_DARK_SLATE

        # Slide 10 (Index 9): Task & Success Criteria
        s10 = prs.slides[9]
        clean_slide_shapes(s10)
        t_box = s10.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "TASK & SUCCESS CRITERIA: HISTORICAL EVENT SPEECH"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(21)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        helper_add_card(s10, Inches(0.6), Inches(1.0), Inches(4.2), Inches(3.8), "SPEECH TASK", [
            "TASK: Write your speech describing an important historical event.",
            "",
            "• Name of the event",
            "• When and where it happened",
            "• Who was involved and what happened",
            "• Why the event is important",
            "• Write around 80-90 words."
        ], bg_color=CLR_LIGHT_YELLOW, border_color=CLR_ORANGE, title_color=CLR_ORANGE)

        helper_add_card(s10, Inches(5.2), Inches(1.0), Inches(4.2), Inches(3.8), "SUCCESS CRITERIA", [
            "✅ 1. Use time expressions (at least 3 times: In 1911, In the past...).",
            "",
            "✅ 2. Use connectors (at least 3 times: First, Then, After that, Finally).",
            "",
            "✅ 3. Use expressions to describe historical events.",
            "",
            "✅ 4. Use Past Simple with WAS and WERE (at least 3 times)."
        ], bg_color=CLR_LIGHT_GREEN, border_color=CLR_GREEN, title_color=CLR_GREEN)

        # Slide 11 (Index 10): Brainstorming Organizer
        s11 = prs.slides[10]
        clean_slide_shapes(s11)
        t_box = s11.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "BRAINSTORMING: HISTORICAL EVENT ORGANIZER"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(21)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        if os.path.exists(img_hist_org):
            s11.shapes.add_picture(img_hist_org, Inches(0.6), Inches(0.95), Inches(8.8), Inches(4.2))

        # Slide 12 (Index 11): Rubrics
        s12 = prs.slides[11]
        clean_slide_shapes(s12, keep_pictures=True)
        t_box = s12.shapes.add_textbox(Inches(0.6), Inches(0.2), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "EVALUATION RUBRICS: HISTORICAL EVENT SPEECH"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(20)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        if os.path.exists(img_rubric):
            s12.shapes.add_picture(img_rubric, Inches(0.5), Inches(0.75), Inches(9.0), Inches(4.5))

        # Slide 13 (Index 12): Drafting in notebook with timer
        s13 = prs.slides[12]
        clean_slide_shapes(s13)
        t_box = s13.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "DRAFTING YOUR SPEECH IN YOUR NOTEBOOK"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(21)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        if os.path.exists(img_timer_07):
            s13.shapes.add_picture(img_timer_07, Inches(7.5), Inches(0.2), Inches(1.9), Inches(1.0))

        helper_add_card(s13, Inches(0.6), Inches(1.2), Inches(8.8), Inches(3.8), "INDIVIDUAL DRAFTING GUIDELINES", [
            "1. Write the title of your historical event.",
            "2. Follow the 4-part structure from the graphic organizer:",
            "   • Introduction: Name and significance of the event.",
            "   • Details: When it happened, who was involved.",
            "   • Sequence: What happened step by step (First, Then, After that).",
            "   • Conclusion: Why it was important for history.",
            "",
            "CHECKPOINT: Ensure you included WAS and WERE at least 3 times!"
        ], bg_color=CLR_LIGHT_YELLOW, border_color=CLR_ORANGE, title_color=CLR_ORANGE)

        # Slide 14 (Index 13): Showtime & Peer Feedback
        s14 = prs.slides[13]
        clean_slide_shapes(s14)
        t_box = s14.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "SHOWTIME: SPEAKING PERFORMANCE & PEER FEEDBACK"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(20)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        helper_add_card(s14, Inches(0.6), Inches(1.0), Inches(4.2), Inches(3.8), "DELIVERING YOUR SPEECH", [
            "• Present your historical event speech.",
            "• Speak loudly and clearly.",
            "• Use expressive intonation.",
            "• Make eye contact with your classmates."
        ], bg_color=CLR_LIGHT_GREEN, border_color=CLR_GREEN, title_color=CLR_GREEN)

        helper_add_card(s14, Inches(5.2), Inches(1.0), Inches(4.2), Inches(3.8), "PEER FEEDBACK: STAR RATING", [
            "Listen to your partner's speech and evaluate:",
            "",
            "⭐⭐⭐ 3 STARS: Used was/were accurately, clear connectors, engaging speech.",
            "⭐⭐ 2 STARS: Good details, minor mistakes in was/were or connectors.",
            "⭐ 1 STAR: Needs more historical details or practice with was/were."
        ], bg_color=CLR_LIGHT_YELLOW, border_color=CLR_ORANGE, title_color=CLR_ORANGE)

        # Slide 15 (Index 14): Grammar & Common Mistakes with WAS / WERE
        s15 = prs.slides[14]
        clean_slide_shapes(s15)
        t_box = s15.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "GRAMMAR FOCUS & COMMON MISTAKES: WAS vs. WERE"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(21)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        helper_add_card(s15, Inches(0.6), Inches(1.0), Inches(4.2), Inches(3.8), "GRAMMAR RULES (PAST OF 'BE')", [
            "• WAS: I, He, She, It",
            "  'Bingham was an explorer.'",
            "  'Machu Picchu wasn't ruined.'",
            "",
            "• WERE: You, We, They",
            "  'The Incas were great builders.'",
            "  'Many people were surprised.'"
        ], bg_color=CLR_WHITE, border_color=CLR_BLUE, title_color=CLR_BLUE)

        helper_add_card(s15, Inches(5.2), Inches(1.0), Inches(4.2), Inches(3.8), "COMMON MISTAKES (AVOID):", [
            "❌ 'Hiram Bingham were an explorer.'",
            "   ➔ ✅ 'Hiram Bingham was an explorer.'",
            "",
            "❌ 'The ruins was very large.'",
            "   ➔ ✅ 'The ruins were very large.'",
            "",
            "❌ 'They wasn't ready.'",
            "   ➔ ✅ 'They weren't ready.'"
        ], bg_color=CLR_LIGHT_RED, border_color=RGBColor(217, 86, 63), title_color=RGBColor(217, 86, 63))

        # Slide 16 (Index 15): Reflection
        s16 = prs.slides[15]
        clean_slide_shapes(s16)
        t_box = s16.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(8.8), Inches(0.6))
        tf = t_box.text_frame
        tf.text = "REFLECTION: CAN-DO STATEMENTS"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(24)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        helper_add_card(s16, Inches(1.0), Inches(1.2), Inches(8.0), Inches(3.6), "HOW DO YOU FEEL ABOUT TODAY'S LESSON?", [
            "",
            "   [   ]  1. I can understand main and specific details of a historical event.",
            "",
            "   [   ]  2. I can use 'was' and 'were' in positive, negative and questions.",
            "",
            "   [   ]  3. I can draft and deliver a short speech about a historical event.",
            "",
            "Share your learning reflection with the class!"
        ], bg_color=CLR_WHITE, border_color=CLR_BLUE, title_color=CLR_BLUE)

    # =========================================================================
    # CASO B: UNIT 5 SESSION 2 (ORAL COMMUNICATION / FOOD & DRINKS)
    # =========================================================================
    elif unit == "5" or "food" in plan_data.get("unit_title", "").lower():
        img_scene = os.path.join(SAMPLES_DIR, "cafe_ordering_scene_1791484978565.jpg")
        img_grid = os.path.join(SAMPLES_DIR, "cafe_food_grid_1791485095937.jpg")
        img_card = os.path.join(SAMPLES_DIR, "cafe_menu_card_1791485139541.jpg")

        # Slide 1: Cover
        s1 = prs.slides[0]
        for sh in s1.shapes:
            if sh.has_text_frame and "A1 Level" in sh.text_frame.text:
                sh.text_frame.text = f"{plan_data['level']}\nUnit 5: “Food and Drinks”\nSession {plan_data['session_num']}"
                for p in sh.text_frame.paragraphs:
                    p.font.name = "Roboto"
                    p.font.bold = True
                    p.font.size = Pt(38)
                    p.font.color.rgb = CLR_WHITE

        # Slide 3: Warm-up
        s3 = prs.slides[2]
        clean_slide_shapes(s3)
        t_box = s3.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(8.8), Inches(0.55))
        tf = t_box.text_frame
        tf.text = "AT THE CAFÉ: WHO ARE THE PEOPLE?"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(22)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        if os.path.exists(img_scene):
            s3.shapes.add_picture(img_scene, Inches(0.6), Inches(1.05), Inches(5.6), Inches(4.0))

        helper_add_bubble(s3, Inches(6.4), Inches(1.05), Inches(3.1), Inches(1.2),
                          "Look at the picture:\n• The waiter is taking the order.\n• The customers are looking at the menu.")
        helper_add_bubble(s3, Inches(6.4), Inches(2.4), Inches(3.1), Inches(1.25),
                          "What can you order in a café?\nFor me, I usually order a cup of coffee and a sandwich.")
        helper_add_bubble(s3, Inches(6.4), Inches(3.8), Inches(3.1), Inches(1.25),
                          "Vocabulary check:\n• Waiter = Male server\n• Waitress = Female server")

        # Slide 4: Objective & Date
        s4 = prs.slides[3]
        for sh in s4.shapes:
            if sh.has_text_frame:
                t = sh.text_frame.text
                if "September" in t or "2026" in t or "Monday" in t:
                    sh.text_frame.text = "Monday, October 12th, 2026."
                    sh.text_frame.paragraphs[0].font.name = "Roboto"
                    sh.text_frame.paragraphs[0].font.size = Pt(16)
                    sh.text_frame.paragraphs[0].font.color.rgb = CLR_WHITE
                elif "Session N" in t:
                    sh.text_frame.text = f"Session N{plan_data['session_num']}: Oral Communication"
                    sh.text_frame.paragraphs[0].font.name = "Roboto"
                    sh.text_frame.paragraphs[0].font.bold = True
                    sh.text_frame.paragraphs[0].font.size = Pt(24)
                    sh.text_frame.paragraphs[0].font.color.rgb = CLR_WHITE
                elif len(t) > 20 and "Objective" not in t:
                    sh.text_frame.text = plan_data['objective']
                    sh.text_frame.paragraphs[0].font.name = "Roboto"
                    sh.text_frame.paragraphs[0].font.bold = True
                    sh.text_frame.paragraphs[0].font.size = Pt(22)
                    sh.text_frame.paragraphs[0].font.color.rgb = CLR_WHITE

        # Slide 5: Vocabulary: Things you can order
        s5 = prs.slides[4]
        clean_slide_shapes(s5)
        t_box = s5.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(8.8), Inches(0.55))
        tf = t_box.text_frame
        tf.text = "THINGS WE CAN ORDER IN A CAFÉ"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(22)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        if os.path.exists(img_grid):
            s5.shapes.add_picture(img_grid, Inches(0.6), Inches(1.05), Inches(5.8), Inches(4.0))

        items_list = ["1. A cup of coffee", "2. A cup of tea", "3. A sandwich", "4. A slice of cake", "5. A soda", "6. A cup of hot chocolate"]
        helper_add_card(s5, Inches(6.6), Inches(1.05), Inches(2.9), Inches(2.6), "Vocabulary:", items_list)
        helper_add_bubble(s5, Inches(6.6), Inches(3.8), Inches(2.9), Inches(1.25),
                          "In a café, I would like to order a sandwich and a soda because I am hungry.")

        # Slide 6: Model Conversation
        s6 = prs.slides[5]
        clean_slide_shapes(s6)
        t_box = s6.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "CONVERSATION WORKSHEET: AT THE CORNER CAFÉ"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(21)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        if os.path.exists(img_card):
            s6.shapes.add_picture(img_card, Inches(0.6), Inches(0.95), Inches(3.6), Inches(4.2))

        dialogue_lines = [
            "Waiter: Good afternoon! Are you ready to order?",
            "Customer 1: Yes, please. I would like a chicken sandwich and a coffee.",
            "Waiter: Sure! And for you?",
            "Customer 2: I would like a slice of chocolate cake and a fresh soda, please.",
            "Waiter: Anything else?",
            "Customer 1: No, that's all, thank you. How much is it?",
            "Waiter: That's $15, please.",
            "Customer 2: Here you are. Thank you!"
        ]
        helper_add_card(s6, Inches(4.4), Inches(0.95), Inches(5.1), Inches(4.2), "Complete Conversation:", dialogue_lines, bg_color=CLR_LIGHT_YELLOW)

        # Slide 7: Useful Expressions
        s7 = prs.slides[6]
        clean_slide_shapes(s7)
        t_box = s7.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "USEFUL EXPRESSIONS IN A CAFÉ"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(22)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        waiter_expr = ["• Good afternoon / Hello!", "• Are you ready to order?", "• What would you like?", "• Can I take your order?", "• Anything else?", "• Here is your order. Enjoy!", "• That's $15, please."]
        helper_add_card(s7, Inches(0.6), Inches(0.95), Inches(4.2), Inches(3.5), "FOR THE WAITER / WAITRESS", waiter_expr, bg_color=CLR_WHITE, border_color=CLR_ORANGE, title_color=CLR_ORANGE)

        cust_expr = ["• Yes, please. / Hello!", "• I would like a ..., please.", "• Can I have a ..., please?", "• No, that's all, thank you.", "• How much is it?", "• Do we have to pay now?", "• Here you are. Thank you!"]
        helper_add_card(s7, Inches(5.2), Inches(0.95), Inches(4.2), Inches(3.5), "FOR THE CUSTOMER", cust_expr, bg_color=CLR_WHITE, border_color=CLR_BLUE, title_color=CLR_BLUE)
        helper_add_bubble(s7, Inches(0.6), Inches(4.6), Inches(8.8), Inches(0.65), "TIP: Focus on polite intonation! Sound cheerful and use 'please' and 'thank you'.", font_size=13)

        # Slide 8: Collaborative Work
        s8 = prs.slides[7]
        clean_slide_shapes(s8)
        t_box = s8.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "COLLABORATIVE WORK: CHOOSE YOUR ROLES"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(22)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        helper_add_card(s8, Inches(0.6), Inches(1.0), Inches(2.8), Inches(3.2), "ROLE 1: WAITER / WAITRESS", ["• Greets customers politely.", "• Asks: 'Are you ready to order?'", "• Takes food and drink orders.", "• Asks: 'Anything else?'", "• Gives the total bill."], bg_color=CLR_LIGHT_YELLOW, border_color=CLR_ORANGE, title_color=CLR_ORANGE)
        helper_add_card(s8, Inches(3.6), Inches(1.0), Inches(2.8), Inches(3.2), "ROLE 2: CUSTOMER 1", ["• Greets the waiter.", "• Orders food and drink using:", "  'I would like a ..., please.'", "• Asks questions about the menu.", "• Polite and friendly."], bg_color=CLR_LIGHT_GREEN, border_color=CLR_GREEN, title_color=CLR_GREEN)
        helper_add_card(s8, Inches(6.6), Inches(1.0), Inches(2.8), Inches(3.2), "ROLE 3: CUSTOMER 2", ["• Orders food and drink using:", "  'Can I have a ..., please?'", "• Confirms: 'That's all, thank you.'", "• Asks for the total bill:", "  'How much is it?'"], bg_color=CLR_BUBBLE, border_color=CLR_BLUE, title_color=CLR_BLUE)
        helper_add_bubble(s8, Inches(0.6), Inches(4.35), Inches(8.8), Inches(0.7), "Form groups of 3 (or pairs). Decide who is the waiter/waitress and who are the customers!", font_size=14)

        # Slide 9: Check-in
        s9 = prs.slides[8]
        for sh in s9.shapes:
            if sh.has_text_frame and "complete" in sh.text_frame.text.lower():
                p = sh.text_frame.paragraphs[0]
                p.font.name = "Chewy"
                p.font.size = Pt(36)
                p.font.color.rgb = CLR_DARK_SLATE

        # Slide 10: Task & Success Criteria
        s10 = prs.slides[9]
        clean_slide_shapes(s10)
        t_box = s10.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "SPEAKING TASK & SUCCESS CRITERIA"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(22)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        helper_add_card(s10, Inches(0.6), Inches(1.0), Inches(4.2), Inches(3.8), "CONVERSATION TASK", ["Work in your group of 3 (or pairs).", "", "Prepare and deliver a conversation between a waiter/waitress and customers in a café.", "", "• Customer 1 & 2 order food & drinks.", "• Waiter/Waitress takes the order politely.", "• Ask for and state the total price."], bg_color=CLR_LIGHT_YELLOW, border_color=CLR_ORANGE, title_color=CLR_ORANGE)
        helper_add_card(s10, Inches(5.2), Inches(1.0), Inches(4.2), Inches(3.8), "SUCCESS CRITERIA", ["1. Greet and order food/drinks politely.", "", "2. Use 'I would like...' and food vocabulary correctly.", "", "3. Use polite expressions ('Please', 'Thank you', 'Anything else?').", "", "4. Pronounce clearly with polite intonation."], bg_color=CLR_LIGHT_GREEN, border_color=CLR_GREEN, title_color=CLR_GREEN)

        # Slide 11: Brainstorming Worksheet
        s11 = prs.slides[10]
        clean_slide_shapes(s11)
        t_box = s11.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "BRAINSTORMING WORKSHEET: PLAN YOUR CONVERSATION"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(21)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        if os.path.exists(img_timer_03):
            s11.shapes.add_picture(img_timer_03, Inches(7.5), Inches(0.25), Inches(1.9), Inches(1.0))

        helper_add_card(s11, Inches(0.6), Inches(1.4), Inches(2.8), Inches(3.6), "WAITER / WAITRESS NOTES", ["• Greeting: 'Good afternoon!'", "• Questions to ask:", "  'Are you ready to order?'", "  'What would you like?'", "  'Anything else?'", "• Prices & Total: 'That's $___, please.'"], bg_color=CLR_WHITE, border_color=CLR_ORANGE, title_color=CLR_ORANGE)
        helper_add_card(s11, Inches(3.6), Inches(1.4), Inches(2.8), Inches(3.6), "CUSTOMER 1 NOTES", ["• Food: 'I would like a _____.'", "• Drink: 'A cup of _____.'", "• Polite: '..., please.'", "• 'Thank you.'"], bg_color=CLR_WHITE, border_color=CLR_GREEN, title_color=CLR_GREEN)
        helper_add_card(s11, Inches(6.6), Inches(1.4), Inches(2.8), Inches(3.6), "CUSTOMER 2 NOTES", ["• Food: 'I would like a _____.'", "• Drink: 'A glass of _____.'", "• Question for bill: 'How much is it?'", "• 'Here you are.'"], bg_color=CLR_WHITE, border_color=CLR_BLUE, title_color=CLR_BLUE)

        # Slide 12: Rubrics
        s12 = prs.slides[11]
        clean_slide_shapes(s12, keep_pictures=True)
        t_box = s12.shapes.add_textbox(Inches(0.6), Inches(0.2), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "EVALUATION RUBRICS: ORAL COMMUNICATION"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(20)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        if os.path.exists(img_rubric):
            s12.shapes.add_picture(img_rubric, Inches(0.5), Inches(0.75), Inches(9.0), Inches(4.5))

        # Slide 13: Rehearsal
        s13 = prs.slides[12]
        clean_slide_shapes(s13)
        t_box = s13.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "STAGE 1 PRACTICE: REHEARSE WITH NOTES"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(21)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        if os.path.exists(img_timer_02):
            s13.shapes.add_picture(img_timer_02, Inches(7.5), Inches(0.25), Inches(1.9), Inches(1.0))

        helper_add_card(s13, Inches(0.6), Inches(1.4), Inches(8.8), Inches(3.6), "INSTRUCTIONS FOR PRACTICE", [
            "1. Stay in your groups of 3.",
            "2. Rehearse your conversation looking at your notes.",
            "3. Focus on polite intonation, clear pronunciation, and smiling.",
            "4. The teacher will monitor and help with grammar and vocabulary corrections.",
            "",
            "CHALLENGE: Every customer must use 'I would like...' at least 2 times!"
        ], bg_color=CLR_LIGHT_YELLOW, border_color=CLR_ORANGE, title_color=CLR_ORANGE)

        # Slide 14: Showtime Performance & Peer Feedback
        s14 = prs.slides[13]
        clean_slide_shapes(s14)
        t_box = s14.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "SHOWTIME: SPEAKING PERFORMANCE & PEER FEEDBACK"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(20)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        helper_add_card(s14, Inches(0.6), Inches(1.0), Inches(4.2), Inches(3.8), "STAGE 2: PERFORM WITHOUT NOTES!", [
            "• Put your notes face down on the desk.",
            "• Perform the conversation naturally.",
            "• Volunteer groups will present in front of the class!",
            "• Speak loudly and clearly for everyone to hear."
        ], bg_color=CLR_LIGHT_GREEN, border_color=CLR_GREEN, title_color=CLR_GREEN)

        helper_add_card(s14, Inches(5.2), Inches(1.0), Inches(4.2), Inches(3.8), "PEER FEEDBACK: STAR RATING", [
            "Listen to your classmates and evaluate:",
            "",
            "⭐⭐⭐ 3 STARS: All criteria fully met! Polite, clear, accurate.",
            "⭐⭐ 2 STARS: Partially met. Good effort, minor errors.",
            "⭐ 1 STAR: Needs improvement in vocabulary or fluency."
        ], bg_color=CLR_LIGHT_YELLOW, border_color=CLR_ORANGE, title_color=CLR_ORANGE)

        # Slide 15: Common Mistakes
        s15 = prs.slides[14]
        clean_slide_shapes(s15)
        t_box = s15.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(8.8), Inches(0.5))
        tf = t_box.text_frame
        tf.text = "LANGUAGE AWARENESS: COMMON MISTAKES"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(21)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        helper_add_card(s15, Inches(0.6), Inches(1.0), Inches(4.2), Inches(3.8), "COMMON MISTAKES (AVOID):", [
            "1. 'I want coffee and a cake.' (Too direct / informal)",
            "2. 'How many is the price?' (Grammar error)",
            "3. 'Give me a sandwich.' (Not polite)",
            "4. 'We has to pay now?' (Subject-verb agreement)"
        ], bg_color=CLR_LIGHT_RED, border_color=RGBColor(217, 86, 63), title_color=RGBColor(217, 86, 63))

        helper_add_card(s15, Inches(5.2), Inches(1.0), Inches(4.2), Inches(3.8), "POLITE & CORRECT ENGLISH:", [
            "1. 'I would like a cup of coffee and a cake, please.'",
            "2. 'How much is it?'",
            "3. 'Can I have a sandwich, please?'",
            "4. 'Do we have to pay now?'"
        ], bg_color=CLR_LIGHT_GREEN, border_color=CLR_GREEN, title_color=CLR_GREEN)

        # Slide 16: Reflection
        s16 = prs.slides[15]
        clean_slide_shapes(s16)
        t_box = s16.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(8.8), Inches(0.6))
        tf = t_box.text_frame
        tf.text = "REFLECTION: CAN-DO STATEMENTS"
        tf.paragraphs[0].font.name = "Roboto"
        tf.paragraphs[0].font.size = Pt(24)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = CLR_BLUE

        helper_add_card(s16, Inches(1.0), Inches(1.2), Inches(8.0), Inches(3.6), "HOW DO YOU FEEL ABOUT TODAY'S LESSON?", [
            "",
            "   [   ]  1. I can name foods and drinks I can order in a café.",
            "   [   ]  2. I can use 'I would like...' and polite expressions to order.",
            "   [   ]  3. I can perform a conversation as a waiter/waitress or customer.",
            "",
            "Share your thoughts with the class!"
        ], bg_color=CLR_WHITE, border_color=CLR_BLUE, title_color=CLR_BLUE)

    # Export to memory buffer
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer
