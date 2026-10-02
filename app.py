import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image, ImageDraw
import bidi.algorithm
import arabic_reshaper

st.set_page_config(page_title="منظومة توليد الامتحانات", layout="wide")

st.title("منظومة توليد الامتحانات المباشرة")
st.write("مدرسة الذاريات الابتدائية المختلطة")

# 1. بيانات الامتحان المولد بالكامل (يمكنك ربطه بمخرجات منظومتك)
header_info = {
    "title": "منظومة توليد الامتحانات المباشرة",
    "school": "مدرسة الذاريات الابتدائية المختلطة",
    "details": "الصف: الخامس الابتدائي | العام الدراسي: 2026 / 2027 م | معلم المادة: حيدر محمد عبد الكريم",
    "subject": "المادة: القرآن الكريم والتربية الإسلامية | الزمن: ساعة واحدة"
}

# قائمة الأسئلة الكاملة مع جميع الفروع
questions = [
    {
        "title": "السؤال الأول: أحكام التلاوة والحفظ (20 درجة)",
        "subs": [
            "أ) اكتب ما تحفظه من سورة الأعلى من قوله تعالى: (سَبِّحِ اسْمَ رَبِّكَ الأَعْلَى) إلى قوله تعالى: (فَنَسَى).",
            "ب) بين أحكام الإظهار والإدغام الواردة في الآيات المذكورة."
        ]
    },
    {
        "title": "السؤال الثاني: الفهم والتفسير (20 درجة)",
        "subs": [
            "أ) اذكر معاني الكلمات الآتية: (سَوَّى - قَدَّرَ فَهَدَى - الغُثَاء - أَحْوَى).",
            "ب) ما هي أهم الدروس والعبر المستفادة من سورة الأعلى؟"
        ]
    },
    {
        "title": "السؤال الثالث: الحديث الشريف والأخلاق (20 درجة)",
        "subs": [
            "أ) اذكر حديثاً نبوياً شريفاً يحث على الصدق وأهميته في حياة المسلم.",
            "ب) وضح كيف يساهم التعاون والتراحم بين الطلاب في بناء بيئة مدرسية ناجحة."
        ]
    }
]

# دالة ضبط الاتجاه لليمين لليسار في Word
def set_paragraph_rtl(paragraph, align=WD_ALIGN_PARAGRAPH.RIGHT):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.alignment = align

# --- 1. إنشاء ملف Word كامل بدون اقتطاع ---
def generate_word():
    doc = Document()
    
    # ضبط الهوامش
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # الترويسة العليا
    p_head = doc.add_paragraph()
    set_paragraph_rtl(p_head, WD_ALIGN_PARAGRAPH.CENTER)
    r_head = p_head.add_run(f"{header_info['title']}\n{header_info['school']}")
    r_head.font.name = 'Arial'
    r_head.font.size = Pt(15)
    r_head.bold = True

    p_det = doc.add_paragraph()
    set_paragraph_rtl(p_det, WD_ALIGN_PARAGRAPH.CENTER)
    r_det = p_det.add_run(f"{header_info['details']}\n{header_info['subject']}")
    r_det.font.name = 'Arial'
    r_det.font.size = Pt(11)

    # خط فاصل
    p_div = doc.add_paragraph()
    set_paragraph_rtl(p_div, WD_ALIGN_PARAGRAPH.CENTER)
    r_div = p_div.add_run("--------------------------------------------------------------------------------")
    r_div.font.name = 'Arial'

    # إضافة كل الأسئلة والفرعيات
    for q in questions:
        p_q = doc.add_paragraph()
        set_paragraph_rtl(p_q)
        r_q = p_q.add_run(q["title"])
        r_q.font.name = 'Arial'
        r_q.font.size = Pt(13)
        r_q.bold = True
        
        for sub in q["subs"]:
            p_sub = doc.add_paragraph()
            set_paragraph_rtl(p_sub)
            r_sub = p_sub.add_run(sub)
            r_sub.font.name = 'Arial'
            r_sub.font.size = Pt(11)
            
        doc.add_paragraph() # مسافة بين الأسئلة

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- 2. إنشاء ملف PDF كامل ---
def generate_pdf():
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    y = height - 50

    lines = [
        header_info['title'],
        header_info['school'],
        header_info['details'],
        header_info['subject'],
        "--------------------------------------------------"
    ]
    
    for q in questions:
        lines.append(q["title"])
        for sub in q["subs"]:
            lines.append(sub)
        lines.append("") # مسافة

    for line in lines:
        if line.strip():
            reshaped_text = arabic_reshaper.reshape(line)
            bidi_text = bidi.algorithm.get_display(reshaped_text)
            p.drawString(40, y, bidi_text)
        y -= 22
        if y < 50:
            p.showPage()
            y = height - 50

    p.save()
    buffer.seek(0)
    return buffer

# --- 3. إنشاء صورة PNG كاملة ---
def generate_png():
    img = Image.new('RGB', (850, 900), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    y = 40

    lines = [
        header_info['title'],
        header_info['school'],
        header_info['details'],
        header_info['subject'],
        "--------------------------------------------------"
    ]
    for q in questions:
        lines.append(q["title"])
        for sub in q["subs"]:
            lines.append(sub)
        lines.append("")

    for line in lines:
        if line.strip():
            reshaped_text = arabic_reshaper.reshape(line)
            bidi_text = bidi.algorithm.get_display(reshaped_text)
            d.text((40, y), bidi_text, fill=(0, 0, 0))
        y += 28

    buffer = BytesIO()
    img.save(buffer, format="PNG")
    buffer.seek(0)
    return buffer

# --- الواجهة والأزرار ---
st.subheader("تحميل ورقة الامتحان الكاملة")

col1, col2, col3 = st.columns(3)

with col1:
    st.download_button("📄 تحميل PDF", data=generate_pdf(), file_name="exam_full.pdf", mime="application/pdf")

with col2:
    st.download_button("📝 تحميل Word", data=generate_word(), file_name="exam_full.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with col3:
    st.download_button("🖼️ تحميل PNG", data=generate_png(), file_name="exam_full.png", mime="image/png")
