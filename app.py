import streamlit as st
import random
from io import BytesIO
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

st.set_page_config(page_title="منظومة توليد الامتحانات", layout="wide")

st.title("منظومة توليد الامتحانات المباشرة - مدرسة الذاريات")
st.write("الصف الخامس الابتدائي | المعلم: حيدر محمد عبد الكريم")

# ---------------------------------------------------------
# 1. إعدادات القوائم المنسدلة
# ---------------------------------------------------------
st.sidebar.header("📋 إعدادات نموذج الامتحان")

exam_type = st.sidebar.selectbox(
    "اختر نوع الامتحان:",
    [
        "أسئلة امتحانات الشهر الأول",
        "أسئلة امتحانات الشهر الثاني",
        "أسئلة امتحانات نصف السنة",
        "أسئلة امتحانات نهاية السنة - الدور الأول",
        "أسئلة امتحانات نهاية السنة - الدور الثاني"
    ]
)

academic_year = st.sidebar.selectbox(
    "اختر العام الدراسي:",
    ["2026/2025", "2027/2026", "2025/2024"]
)

time_limit = st.sidebar.selectbox(
    "اختر زمن الامتحان:",
    ["ساعة واحدة", "ساعتان", "ساعة ونصف"]
)

# ---------------------------------------------------------
# 2. بنك الأسئلة
# ---------------------------------------------------------
QUESTION_BANK = {
    "q1_quran": [
        "سورة ( الملك ) من قوله تعالى ( تَبَارَكَ الَّذِي بِيَدِهِ الْمُلْكُ ) إلى قوله تعالى ( عَذَابَ جَهَنَّمَ وَبِئْسَ الْمَصِيرُ )",
        "سورة ( البلد ) من قوله تعالى ( لَا أُقْسِمُ بِهَذَا الْبَلَدِ ) إلى قوله تعالى ( وَهَدَيْنَاهُ النَّجْدَيْنِ )",
        "سورة ( الأعلى ) من قوله تعالى ( سَبِّحِ اسْمَ رَبِّكَ الْأَعْلَى ) إلى قوله تعالى ( فَنَسَى )"
    ],
    "q2_meanings": [
        "مشفقون", "وما يسطرون", "طباقا", "حل", "كرتين", "هلوعا", 
        "سوى", "قدر فهدى", "الغثاء", "أحوى", "فلا تنسى", "النجدين"
    ],
    "q2_tafseer": [
        "ما المعنى العام للآية الكريمة : ( الَّذِينَ هُمْ عَلَى صَلَاتِهِمْ دَائِمُونَ ) ؟",
        "ما المعنى العام للآية الكريمة : ( وَالَّذِينَ فِي أَمْوَالِهِمْ حَقٌّ مَّعْلُومٌ ) ؟"
    ],
    "q3_hadith": [
        ("اكتب حديثاً نبوياً شريفاً في ( التوبة ) ؟", "اكتب حديثاً نبوياً شريفاً في ( حفظ اللسان ) ؟")
    ],
    "q4_beliefs": [
        "آمن علماء اليهود والنصارى بالنبي محمد ( ص ) .",
        "الإنجيل هو الكتاب المنزل على النبي يوسف ( ع ) .",
        "التوراة هو الكتاب المنزل على النبي إبراهيم ( ع ) .",
        "من أسماء الله الحسنى المنتقم والودود والشكور .",
        "يتصف جميع الأنبياء بالصدق والحكمة والصبر ومكارم الأخلاق .",
        "حارب الأنبياء الطواغيت والحكام الظالمين لنصرة المستضعفين وتحريرهم ."
    ],
    "q5_seerah": [
        "كان اسم المدينة المنورة قبل مجيء الرسول إليها يسمى ______________ .",
        "سميت السور التي نزلت بمكة بالسور ______________ والتي نزلت بالمدينة بالسور ______________ .",
        "يرجع نسب النبي أيوب ( ع ) إلى النبي ______________ .",
        "أول الآيات التي نزلت على النبي محمد ( ص ) كانت من سورة ______________ .",
        "مرت الدعوة الإسلامية بمرحلتين ______________ و ______________ .",
        "من أبرز شهداء معركة أحد مصعب بن عمير و ______________ .",
        "تبعد المدينة المنورة عن مكة المكرمة مسافة ______________ .",
        "سمى القرآن الكريم يوم معركة بدر بيوم ______________ ."
    ]
}

# ---------------------------------------------------------
# 3. دالة توليد ملف Word
# ---------------------------------------------------------
def generate_word(selected_type, selected_year, selected_time):
    q1_samples = random.sample(QUESTION_BANK["q1_quran"], 2)
    words = random.sample(QUESTION_BANK["q2_meanings"], 6)
    words_str = "    ".join([f"{i+1}- {w}" for i, w in enumerate(words)])
    tafseer = random.choice(QUESTION_BANK["q2_tafseer"])
    hadith = random.choice(QUESTION_BANK["q3_hadith"])
    beliefs = random.sample(QUESTION_BANK["q4_beliefs"], 6)
    seerah = random.sample(QUESTION_BANK["q5_seerah"], 8)

    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    def set_p_rtl(p, align=WD_ALIGN_PARAGRAPH.RIGHT):
        pPr = p._p.get_or_add_pPr()
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)
        p.alignment = align

    # الترويسة
    table = doc.add_table(rows=1, cols=3)
    table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width = Inches(2.2)
    table.columns[1].width = Inches(3.0)
    table.columns[2].width = Inches(2.3)

    p_r = table.cell(0, 0).paragraphs[0]
    set_p_rtl(p_r, WD_ALIGN_PARAGRAPH.RIGHT)
    r = p_r.add_run(f"المادة : التربية الإسلامية\nالصف : الخامس الابتدائي\nالزمن : {selected_time}")
    r.font.name = 'Arial'; r.font.size = Pt(10); r.bold = True

    p_c = table.cell(0, 1).paragraphs[0]
    set_p_rtl(p_c, WD_ALIGN_PARAGRAPH.CENTER)
    rc = p_c.add_run(f"بسم الله الرحمن الرحيم\n{selected_type}\nللعام الدراسي {selected_year}")
    rc.font.name = 'Arial'; rc.font.size = Pt(11); rc.bold = True

    p_l = table.cell(0, 2).paragraphs[0]
    set_p_rtl(p_l, WD_ALIGN_PARAGRAPH.LEFT)
    rl = p_l.add_run("إدارة\nمدرسة الذاريات\nالابتدائية المختلطة")
    rl.font.name = 'Arial'; rl.font.size = Pt(10); rl.bold = True

    doc.add_paragraph()

    def add_section_header(title_text):
        p = doc.add_paragraph()
        set_p_rtl(p, WD_ALIGN_PARAGRAPH.RIGHT)
        r = p.add_run(title_text)
        r.font.name = 'Arial'; r.font.size = Pt(12); r.bold = True

    add_section_header("القرآن الكريم : ( 20 درجة )")
    p = doc.add_paragraph(); set_p_rtl(p)
    p.add_run("س1 : أجب عن أحد الفرعين :").bold = True

    p_a = doc.add_paragraph(); set_p_rtl(p_a)
    p_a.add_run(f"أ / اكتب ما تحفظه من {q1_samples[0]}").font.name = 'Arial'

    p_b = doc.add_paragraph(); set_p_rtl(p_b)
    p_b.add_run(f"ب / اكتب ما تحفظه من {q1_samples[1]}").font.name = 'Arial'

    add_section_header("المعاني والتفسير : ( 10 درجات )")
    p = doc.add_paragraph(); set_p_rtl(p)
    p.add_run("س2 : أجب عن ما يلي :").bold = True

    p_w_title = doc.add_paragraph(); set_p_rtl(p_w_title)
    p_w_title.add_run("أ / أعط معاني لخمس من الكلمات الآتية :").font.name = 'Arial'
    
    p_w = doc.add_paragraph(); set_p_rtl(p_w)
    p_w.add_run(f"( {words_str} )").font.name = 'Arial'

    p_t = doc.add_paragraph(); set_p_rtl(p_t)
    p_t.add_run(f"ب / {tafseer}").font.name = 'Arial'

    add_section_header("الحديث الشريف : ( 15 درجة )")
    p = doc.add_paragraph(); set_p_rtl(p)
    p.add_run("س3 : الإجابة عن أحد الفرعين :").bold = True

    p_h = doc.add_paragraph(); set_p_rtl(p_h)
    p_h.add_run(f"أ / {hadith[0]}           ب / {hadith[1]}").font.name = 'Arial'

    add_section_header("العقائد والعبادات : ( 15 درجة )")
    p = doc.add_paragraph(); set_p_rtl(p)
    p.add_run("س4 : أجب بكلمة ( صح ) عن العبارة الصحيحة وكلمة ( خطأ ) عن العبارة الخاطئة :").bold = True

    for idx, b in enumerate(beliefs, 1):
        pq = doc.add_paragraph(); set_p_rtl(pq)
        pq.add_run(f"{idx}- {b}").font.name = 'Arial'

    add_section_header("السيرة النبوية والآداب الإسلامية : ( 20 درجة )")
    p = doc.add_paragraph(); set_p_rtl(p)
    p.add_run("س5 : املأ الفراغات الآتية :").bold = True

    for idx, s in enumerate(seerah, 1):
        pq = doc.add_paragraph(); set_p_rtl(pq)
        pq.add_run(f"{idx}- {s}").font.name = 'Arial'

    p_sig = doc.add_paragraph()
    set_p_rtl(p_sig, WD_ALIGN_PARAGRAPH.LEFT)
    rs = p_sig.add_run("معلم المادة\nحيدر محمد عبد الكريم")
    rs.font.name = 'Arial'; rs.bold = True

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# 4. توليد المحتوى وتنزيله في الواجهة
# ---------------------------------------------------------
if st.button("🎲 توليد نموذج الامتحان الجديد"):
    st.session_state['exam_word'] = generate_word(exam_type, academic_year, time_limit)
    st.success("تم توليد نموذج الأسئلة بنجاح!")

if 'exam_word' in st.session_state:
    st.download_button(
        label="📥 تحميل ملف Word (قابل للتعديل)",
        data=st.session_state['exam_word'],
        file_name=f"Exam_{exam_type}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
