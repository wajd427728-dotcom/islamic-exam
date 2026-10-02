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

# دالة لضبط اتجاه النص من اليمين لليسار في Word
def set_paragraph_rtl(paragraph, align=WD_ALIGN_PARAGRAPH.RIGHT):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.alignment = align

def generate_word():
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # الترويسة الثلاثية للنموذج الرسمي
    table = doc.add_table(rows=1, cols=3)
    table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width = Inches(2.2)
    table.columns[1].width = Inches(3.0)
    table.columns[2].width = Inches(2.3)

    cell_right = table.cell(0, 0)
    cell_center = table.cell(0, 1)
    cell_left = table.cell(0, 2)

    p_r = cell_right.paragraphs[0]
    set_p_rtl(p_r, WD_ALIGN_PARAGRAPH.RIGHT)
    r = p_r.add_run("المادة : التربية الإسلامية\nالصف : الخامس الابتدائي\nالزمن : ساعتان")
    r.font.name = 'Arial'
    r.font.size = Pt(10)
    r.bold = True

    p_c = cell_center.paragraphs[0]
    set_p_rtl(p_c, WD_ALIGN_PARAGRAPH.CENTER)
    rc = p_c.add_run("بسم الله الرحمن الرحيم\nأسئلة امتحانات نهاية السنة\nللعام الدراسي 2026/2025 الدور الثاني")
    rc.font.name = 'Arial'
    rc.font.size = Pt(11)
    rc.bold = True

    p_l = cell_left.paragraphs[0]
    set_p_rtl(p_l, WD_ALIGN_PARAGRAPH.LEFT)
    rl = p_l.add_run("ادارة\nمدرسة الذاريات\nالابتدائية المختلطة")
    rl.font.name = 'Arial'
    rl.font.size = Pt(10)
    rl.bold = True

    doc.add_paragraph()

    def add_section_header(title_text):
        p = doc.add_paragraph()
        set_p_rtl(p, WD_ALIGN_PARAGRAPH.RIGHT)
        r = p.add_run(title_text)
        r.font.name = 'Arial'
        r.font.size = Pt(12)
        r.bold = True

    # Section 1
    add_section_header("القرآن الكريم : ( 20 درجة )")
    p = doc.add_paragraph()
    set_p_rtl(p)
    p.add_run("س1 : اجب عن احد الفرعين :").bold = True
    p.runs[0].font.name = 'Arial'
    
    p = doc.add_paragraph()
    set_p_rtl(p)
    p.add_run("أ / اكتب ما تحفظه من سورة ( الملك ) من قوله تعالى ( تَبَارَكَ الَّذِي بِيَدِهِ الْمُلْكُ ) إلى قوله تعالى ( عَذَابَ جَهَنَّمَ وَبِئْسَ الْمَصِيرُ )").font.name = 'Arial'
    
    p = doc.add_paragraph()
    set_p_rtl(p)
    p.add_run("ب / اكتب ما تحفظه من سورة ( البلد ) من قوله تعالى ( لَا أُقْسِمُ بِهَذَا الْبَلَدِ ) إلى قوله تعالى ( وَهَدَيْنَاهُ النَّجْدَيْنِ )").font.name = 'Arial'

    # Section 2
    add_section_header("المعاني والتفسير : ( 10 درجات )")
    p = doc.add_paragraph()
    set_p_rtl(p)
    p.add_run("س2 : اجب عن ما يلي :").bold = True
    p.runs[0].font.name = 'Arial'

    p = doc.add_paragraph()
    set_p_rtl(p)
    p.add_run("أ / أعط معاني لخمس من الكلمات الآتيين :").font.name = 'Arial'

    p = doc.add_paragraph()
    set_p_rtl(p)
    p.add_run("(1- مشفقون   2- وما يسطرون   3- طباقا   4- حل   5- كرتين   6- هلوعا )").font.name = 'Arial'

    p = doc.add_paragraph()
    set_p_rtl(p)
    p.add_run("ب / ما المعنى العام للآية القرآنية التالية : بسم الله الرحمن الرحيم ( الَّذِينَ هُمْ عَلَى صَلَاتِهِمْ دَائِمُونَ ) .").font.name = 'Arial'

    # Section 3
    add_section_header("الحديث الشريف : ( 15 درجة )")
    p = doc.add_paragraph()
    set_p_rtl(p)
    p.add_run("س3 : الإجابة عن احد الفرعين :").bold = True
    p.runs[0].font.name = 'Arial'

    p = doc.add_paragraph()
    set_p_rtl(p)
    p.add_run("أ / اكتب حديثاً نبوياً شريفاً في ( التوبة ) ؟           ب / اكتب حديثاً نبوياً شريفاً في ( حفظ اللسان ) ؟").font.name = 'Arial'

    # Section 4
    add_section_header("العقائد والعبادات : ( 15 درجة )")
    p = doc.add_paragraph()
    set_p_rtl(p)
    p.add_run("س4 : اجب عن الإجابة الصحيحة بكلمة ( صح ) وعن الإجابة الخاطئة بكلمة ( خطا ) :").bold = True
    p.runs[0].font.name = 'Arial'

    q_list_4 = [
        "1- الإنجيل هو الكتاب المنزل على النبي يوسف (ع) .",
        "2- آمن علماء اليهود والنصاري بالنبي محمد (ص) .",
        "3- التوراة هو الكتاب المنزل على النبي إبراهيم (ع) .",
        "4- من أسماء الله الحسنى المنتقم والودود و الشكور .",
        "5- يتصف جميع الأنبياء بالصدق والحكمة والصبر ومكارم الأخلاق .",
        "6- حارب الأنبياء الطواغيت والحكام الظالمين لنصرة المستضعفين وتحريرهم من سيطرة الطواغيت ."
    ]
    for q in q_list_4:
        pq = doc.add_paragraph()
        set_p_rtl(pq)
        pq.add_run(q).font.name = 'Arial'

    # Section 5
    add_section_header("السيرة النبوية والآداب الإسلامية : ( 20 درجة )")
    p = doc.add_paragraph()
    set_p_rtl(p)
    p.add_run("س5 : املأ الفراغات الآتية :").bold = True
    p.runs[0].font.name = 'Arial'

    q_list_5 = [
        "1- مرت الدعوة الإسلامية بمرحلتين ______ و ______ .",
        "2- يرجع نسب النبي أيوب (ع) إلى النبي ______ .",
        "3- أول الآيات التي نزلت على النبي محمد (ص) كانت من سورة ______ .",
        "4- كان اسم المدينة المنورة قبل مجيء الرسول اليها تسمى ______ .",
        "5- سمى القرآن الكريم يوم معركة بدر بيوم ______ .",
        "6- من ابرز شهداء معركة احد مصعب بن عمير و ______ .",
        "7- تبعد المدينة المنورة عن مكة ______ .",
        "8- سميت السور التي نزلت بمكة بالسور ______ التي نزلت بالمدينة بالسور ______ ."
    ]
    for q in q_list_5:
        pq = doc.add_paragraph()
        set_p_rtl(pq)
        pq.add_run(q).font.name = 'Arial'

    # Signature
    p_sig = doc.add_paragraph()
    set_p_rtl(p_sig, WD_ALIGN_PARAGRAPH.LEFT)
    rs = p_sig.add_run("معلم المادة\nحيدر محمد عبد الكريم")
    rs.font.name = 'Arial'
    rs.bold = True

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# زر التحميل في واجهة Streamlit
st.subheader("تحميل ورقة الامتحان الرسمية المطابقة للنموذج")
if st.button("📝 تحميل ملف Word للنموذج الرسمي"):
    word_file = generate_word()
    st.download_button(
        label="اضغط هنا لتنزيل الملف",
        data=word_file,
        file_name="official_exam_template.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
