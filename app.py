import streamlit as st
from docx import Document
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image, ImageDraw, ImageFont

st.title("منظومة توليد الامتحانات المباشرة")
st.write("مدرسة الذاريات الابتدائية المختلطة")

# نص الامتحان الوهمي أو المولد
exam_content = """
الصف: الخامس الابتدائي | العام الدراسي: 2026 / 2027 م | معلم المادة: حيدر محمد عبد الكريم
المادة: القرآن الكريم وتربية الإسلامية
--------------------------------------------------
السؤال الأول: أجب عن الأسئلة الآتية:
1. ما هي سور القرآن المكية؟
2. اذكر أحكام النون السكنية والتنوين.
"""

st.text_area("معاينة الأسئلة:", exam_content, height=150)

# --- وظائف التصدير ---

# 1. تصدير Word (.docx)
def generate_word(text):
    doc = Document()
    doc.add_heading('منظومة توليد الامتحانات المباشرة', level=1)
    doc.add_paragraph(text)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# 2. تصدير PDF
def generate_pdf(text):
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    
    # كتابة النص بشكل مبسط في ملف الـ PDF
    text_object = p.beginText(40, height - 50)
    text_object.setFont("Helvetica", 12)
    
    for line in text.split('\n'):
        text_object.textLine(line)
        
    p.drawText(text_object)
    p.showPage()
    p.save()
    buffer.seek(0)
    return buffer

# 3. تصدير PNG (صورة)
def generate_png(text):
    # إنشاء صورة بيضاء وكتابة النص عليها
    img = Image.new('RGB', (800, 600), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    
    # رسم النص على الصورة (ملاحظة: لغة بايثون الافتراضية قد تحتاج خط يدعم العربية، أو يمكن استخدام خط افتراضي)
    d.text((40, 40), text, fill=(0, 0, 0))
    
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    buffer.seek(0)
    return buffer

# --- أزرار الواجهة للتنزيل ---
st.subheader("تحميل وفتح الملفات")

col1, col2, col3 = st.columns(3)

with col1:
    pdf_data = generate_pdf(exam_content)
    st.download_button(
        label="📄 تحميل PDF",
        data=pdf_data,
        file_name="exam.pdf",
        mime="application/pdf"
    )

with col2:
    word_data = generate_word(exam_content)
    st.download_button(
        label="📝 تحميل Word",
        data=word_data,
        file_name="exam.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

with col3:
    png_data = generate_png(exam_content)
    st.download_button(
        label="🖼️️ تحميل PNG",
        data=png_data,
        file_name="exam.png",
        mime="image/png"
    )
