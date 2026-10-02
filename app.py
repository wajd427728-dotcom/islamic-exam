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

st.title("منظومة توليد الامتحانات المباشرة")
st.write("مدرسة الذاريات الابتدائية المختلطة")

# قائمة الأسئلة والترويسة كعناصر منفصلة ومرتبة
school_title = "منظومة توليد الامتحانات المباشرة\nمدرسة الذاريات الابتدائية المختلطة"
exam_info = "الصف: الخامس الابتدائي | العام الدراسي: 2026 / 2027 م | معلم المادة: حيدر محمد عبد الكريم\nالمادة: القرآن الكريم والتربية الإسلامية"
divider = "--------------------------------------------------"
q1 = "السؤال الأول: أجب عن الأسئلة الآتية:"
sub_q1 = "1. ما هي سور القرآن المكية؟"
sub_q2 = "2. اذكر أحكام النون الساكنة والتنوين."

# معاينة نصية في التطبيق
full_text = f"{school_title}\n{exam_info}\n{divider}\n{q1}\n{sub_q1}\n{sub_q2}"
st.text_area("معاينة الأسئلة:", full_text, height=180)

# دالة لضبط اتجاه النص من اليمين لليسار في ملفات Word
def set_paragraph_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

# 1. تصدير Word (.docx) بشكل مرتب لكل فقرة على حدة
def generate_word():
    doc = Document()
    
    # ضبط الهوامش القياسية
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # فقرة العنوان والترويسة
    p1 = doc.add_paragraph()
    set_paragraph_rtl(p1)
    r1 = p1.add_run(school_title)
    r1.font.name = 'Arial'
    r1.font.size = Pt(14)
    r1.bold = True

    # فقرة معلومات الامتحان
    p2 = doc.add_paragraph()
    set_paragraph_rtl(p2)
    r2 = p2.add_run(exam_info)
    r2.font.name = 'Arial'
    r2.font.size = Pt(11)

    # فاصل
    p3 = doc.add_paragraph()
    set_paragraph_rtl(p3)
    r3 = p3.add_run(divider)
    r3.font.name = 'Arial'

    # فقرات الأسئلة بترتيبها الصحيح
    questions_list = [q1, sub_q1, sub_q2]
    for text in questions_list:
        pq = doc.add_paragraph()
        set_paragraph_rtl(pq)
        rq = pq.add_run(text)
        rq.font.name = 'Arial'
        rq.font.size = Pt(12)
        if "السؤال" in text:
            rq.bold = True

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# 2. تصدير PDF
def generate_pdf():
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    
    lines = [school_title, exam_info, divider, q1, sub_q1, sub_q2]
    y = height - 50
    for block in lines:
        for line in block.split('\n'):
            if line.strip():
                reshaped_text = arabic_reshaper.reshape(line)
                bidi_text = bidi.algorithm.get_display(reshaped_text)
                p.drawString(40, y, bidi_text)
            y -= 25
            if y < 50:
                p.showPage()
                y = height - 50
            
    p.save()
    buffer.seek(0)
    return buffer

# 3. تصدير PNG (صورة)
def generate_png():
    img = Image.new('RGB', (800, 600), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    
    lines = [school_title, exam_info, divider, q1, sub_q1, sub_q2]
    y = 40
    for block in lines:
        for line in block.split('\n'):
            if line.strip():
                reshaped_text = arabic_reshaper.reshape(line)
                bidi_text = bidi.algorithm.get_display(reshaped_text)
                d.text((40, y), bidi_text, fill=(0, 0, 0))
            y += 30
        
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    buffer.seek(0)
    return buffer

# --- أزرار التنزيل في الواجهة ---
st.subheader("تحميل وفتح الملفات بالصيغ المختلفة")

col1, col2, col3 = st.columns(3)

with col1:
    pdf_data = generate_pdf()
    st.download_button("📄 تحميل PDF", data=pdf_data, file_name="exam.pdf", mime="application/pdf")

with col2:
    word_data = generate_word()
    st.download_button("📝 تحميل Word", data=word_data, file_name="exam.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with col3:
    png_data = generate_png()
    st.download_button("🖼 تحميل PNG", data=png_data, file_name="exam.png", mime="image/png")
