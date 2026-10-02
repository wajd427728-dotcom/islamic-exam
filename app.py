import streamlit as st
from docx import Document
from docx.shared import Pt
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

# نص الامتحان
exam_content = """منظومة توليد الامتحانات المباشرة
مدرسة الذاريات الابتدائية المختلطة
الصف: الخامس الابتدائي | العام الدراسي: 2026 / 2027 م | معلم المادة: حيدر محمد عبد الكريم
المادة: القرآن الكريم والتربية الإسلامية
--------------------------------------------------
السؤال الأول: أجب عن الأسئلة الآتية:
1. ما هي سور القرآن المكية؟
2. اذكر أحكام النون الساكنة والتنوين.
"""

st.text_area("معاينة الأسئلة:", exam_content, height=180)

# دالة لضبط اتجاه النص من اليمين لليسار في ملفات Word
def set_paragraph_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

# 1. تصدير Word (.docx)
def generate_word(text):
    doc = Document()
    for line in text.split('\n'):
        p = doc.add_paragraph()
        set_paragraph_rtl(p)
        run = p.add_run(line)
        run.font.name = 'Arial'
        run.font.size = Pt(14)
        
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# 2. تصدير PDF
def generate_pdf(text):
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    
    y = height - 50
    for line in text.split('\n'):
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
def generate_png(text):
    img = Image.new('RGB', (800, 600), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    
    y = 40
    for line in text.split('\n'):
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
    pdf_data = generate_pdf(exam_content)
    st.download_button("📄 تحميل PDF", data=pdf_data, file_name="exam.pdf", mime="application/pdf")

with col2:
    word_data = generate_word(exam_content)
    st.download_button("📝 تحميل Word", data=word_data, file_name="exam.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with col3:
    png_data = generate_png(exam_content)
    st.download_button("🖼️ تحميل PNG", data=png_data, file_name="exam.png", mime="image/png")
