import streamlit as st
import random
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import docx

st.set_page_config(page_title="منظومة توليد الامتحانات", layout="wide")

st.title("منظومة توليد الامتحانات المباشرة - مدرسة الذاريات")
st.write("الصف الخامس الابتدائي | المعلم: حيدر محمد عبد الكريم")

# ---------------------------------------------------------
# 1. القوائم المنسدلة تحديد إعدادات الامتحان
# ---------------------------------------------------------
st.sidebar.header("📋 إعدادات نموذج الامتحان")

# قائمة منسدلة لنوع الامتحان
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

# قائمة منسدلة للسنة الدراسية
academic_year = st.sidebar.selectbox(
    "اختر العام الدراسي:",
    [
        "2026/2025",
        "2027/2026",
        "2025/2024"
    ]
)

# قائمة منسدلة لزمن الامتحان
time_limit = st.sidebar.selectbox(
    "اختر زمن الامتحان:",
    [
        "ساعتان",
        "ساعة ونصف",
        "ساعة واحدة"
    ]
)

# ---------------------------------------------------------
# 2. بنك الأسئلة الشامل (مضبوط الأقواس والرموز)
# ---------------------------------------------------------
QUESTION_BANK = {
    "q1_quran": [
        "سورة ( الملك ) من قوله تعالى \u200f( تَبَارَكَ الَّذِي بِيَدِهِ الْمُلْكُ )\u200f إلى قوله تعالى \u200f( عَذَابَ جَهَنَّمَ وَبِئْسَ الْمَصِيرُ )\u200f",
        "سورة ( البلد ) من قوله تعالى \u200f( لَا أُقْسِمُ بِهَذَا الْبَلَدِ )\u200f إلى قوله تعالى \u200f( وَهَدَيْنَاهُ النَّجْدَيْنِ )\u200f",
        "سورة ( الأعلى ) من قوله تعالى \u200f( سَبِّحِ اسْمَ رَبِّكَ الْأَعْلَى )\u200f إلى قوله تعالى \u200f( فَنَسَى )\u200f",
        "سورة ( المعارج ) من قوله تعالى \u200f( سَأَلَ سَائِلٌ بِعَذَابٍ وَاقِعٍ )\u200f إلى قوله تعالى \u200f( لَّيْسَ لَهُ دَافِعٌ )\u200f"
    ],
    "q2_meanings": [
        "مشفقون", "وما يسطرون", "طباقا", "حل", "كرتين", "هلوعا", 
        "سوى", "قدر فهدى", "الغثاء", "أحوى", "فلا تنسى", "النجدين"
    ],
    "q2_tafseer": [
        "ما المعنى العام للآية الكريمة : \u200f( الَّذِينَ هُمْ عَلَى صَلَاتِهِمْ دَائِمُونَ )\u200f ؟",
        "ما المعنى العام للآية الكريمة : \u200f( وَالَّذِينَ فِي أَمْوَالِهِمْ حَقٌّ مَّعْلُومٌ )\u200f ؟",
        "ما المعنى العام للآية الكريمة : \u200f( الَّذِي خَلَقَ فَسَوَّى )\u200f ؟"
    ],
    "q3_hadith": [
        ("اكتب حديثاً نبوياً شريفاً في ( التوبة ) ؟", "اكتب حديثاً نبوياً شريفاً في ( حفظ اللسان ) ؟"),
        ("اكتب حديثاً نبوياً شريفاً في ( إعادة العارية ) ؟", "اكتب حديثاً نبوياً شريفاً في ( الحث على الصدق ) ؟")
    ],
    "q4_beliefs": [
        "الإنجيل هو الكتاب المنزل على النبي يوسف \u200f(ع)\u200f .",
        "آمن علماء اليهود والنصارى بالنبي محمد \u200f(ص)\u200f .",
        "التوراة هو الكتاب المنزل على النبي إبراهيم \u200f(ع)\u200f .",
        "من أسماء الله الحسنى المنتقم والودود والشكور .",
        "يتصف جميع الأنبياء بالصدق والحكمة والصبر ومكارم الأخلاق .",
        "حارب الأنبياء الطواغيت والحكام الظالمين لنصرة المستضعفين وتحريرهم ."
    ],
    "q5_seerah": [
        "مرت الدعوة الإسلامية بمرحلتين ______________ و ______________ .",
        "يرجع نسب النبي أيوب \u200f(ع)\u200f إلى النبي ______________ .",
        "أول الآيات التي نزلت على النبي محمد \u200f(ص)\u200f كانت من سورة ______________ .",
        "كان اسم المدينة المنورة قبل مجيء الرسول إليها يسمى ______________ .",
        "سمى القرآن الكريم يوم معركة بدر بيوم ______________ .",
        "من أبرز شهداء معركة أحد مصعب بن عمير و ______________ .",
        "تبعد المدينة المنورة عن مكة المكرمة مسافة ______________ .",
        "سميت السور التي نزلت بمكة بالسور ______________ والتي نزلت بالمدينة بالسور ______________ ."
    ]
}

# ---------------------------------------------------------
# 3. دالة محاذاة اتجاه اليمين (RTL)
# ---------------------------------------------------------
def set_paragraph_rtl(paragraph, align=WD_ALIGN_PARAGRAPH.RIGHT):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.alignment = align

# ---------------------------------------------------------
# 4. دالة توليد الامتحان بأسئلة عشوائية
# ---------------------------------------------------------
def generate_exam(selected_type, selected_year, selected_time):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # الترويسة الثلاثية الرسمية
    table = doc.add_table(rows=1, cols=3)
    table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width = Inches(2.2)
    table.columns[1].width = Inches(3.0)
    table.columns[2].width = Inches(2.3)

    # اليمين
    p_r = table.cell(0, 0).paragraphs[0]
    set_paragraph_rtl(p_r, WD_ALIGN_PARAGRAPH.RIGHT)
    r = p_r.add_run(f"المادة : التربية الإسلامية\nالصف : الخامس الابتدائي\nالزمن : {selected_time}")
    r.font.name = 'Arial'; r.font.size = Pt(10); r.bold = True

    # الوسط (النوع المختار والمناظر للسنة)
    p_c = table.cell(0, 1).paragraphs[0]
    set_paragraph_rtl(p_c, WD_ALIGN_PARAGRAPH.CENTER)
    rc = p_c.add_run(f"بسم الله الرحمن الرحيم\n{selected_type}\nللعام الدراسي {selected_year}")
    rc.font.name = 'Arial'; rc.font.size = Pt(11); rc.bold = True

    # اليسار
    p_l = table.cell(0, 2).paragraphs[0]
    set_paragraph_rtl(p_l, WD_ALIGN_PARAGRAPH.LEFT)
    rl = p_l.add_run("إدارة\nمدرسة الذاريات\nالابتدائية المختلطة")
    rl.font.name = 'Arial'; rl.font.size = Pt(10); rl.bold = True

    doc.add_paragraph()

    def add_section_header(title_text):
        p = doc.add_paragraph()
        set_paragraph_rtl(p, WD_ALIGN_PARAGRAPH.RIGHT)
        r = p.add_run(title_text)
        r.font.name = 'Arial'; r.font.size = Pt(12); r.bold = True

    # س1: القرآن الكريم
    add_section_header("القرآن الكريم : ( 20 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س1 : أجب عن أحد الفرعين :").bold = True
    
    q1_samples = random.sample(QUESTION_BANK["q1_quran"], 2)
    p_a = doc.add_paragraph(); set_paragraph_rtl(p_a)
    p_a.add_run(f"أ / اكتب ما تحفظه من {q1_samples[0]}").font.name = 'Arial'
    p_b = doc.add_paragraph(); set_paragraph_rtl(p_b)
    p_b.add_run(f"ب / اكتب ما تحفظه من {q1_samples[1]}").font.name = 'Arial'

    # س2: المعاني والتفسير
    add_section_header("المعاني والتفسير : ( 10 درجات )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س2 : أجب عن ما يلي :").bold = True

    p_w_title = doc.add_paragraph(); set_paragraph_rtl(p_w_title)
    p_w_title.add_run("أ / أعط معاني لخمس من الكلمات الآتية :").font.name = 'Arial'
    
    words = random.sample(QUESTION_BANK["q2_meanings"], 6)
    words_str = "   ".join([f"{i+1}- {w}" for i, w in enumerate(words)])
    p_w = doc.add_paragraph(); set_paragraph_rtl(p_w)
    p_w.add_run(f"\u200f( {words_str} )\u200f").font.name = 'Arial'

    tafseer = random.choice(QUESTION_BANK["q2_tafseer"])
    p_t = doc.add_paragraph(); set_paragraph_rtl(p_t)
    p_t.add_run(f"ب / {tafseer}").font.name = 'Arial'

    # س3: الحديث الشريف
    add_section_header("الحديث الشريف : ( 15 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س3 : الإجابة عن أحد الفرعين :").bold = True

    hadith_pair = random.choice(QUESTION_BANK["q3_hadith"])
    p_h = doc.add_paragraph(); set_paragraph_rtl(p_h)
    p_h.add_run(f"أ / {hadith_pair[0]}           ب / {hadith_pair[1]}").font.name = 'Arial'

    # س4: العقائد والعبادات
    add_section_header("العقائد والعبادات : ( 15 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س4 : أجب بكلمة ( صح ) عن العبارة الصحيحة وكلمة ( خطأ ) عن العبارة الخاطئة :").bold = True

    beliefs = random.sample(QUESTION_BANK["q4_beliefs"], 6)
    for idx, b in enumerate(beliefs, 1):
        pq = doc.add_paragraph(); set_paragraph_rtl(pq)
        pq.add_run(f"{idx}- {b}").font.name = 'Arial'

    # س5: السيرة النبوية
    add_section_header("السيرة النبوية والآداب الإسلامية : ( 20 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س5 : املأ الفراغات الآتية :").bold = True

    seerah = random.sample(QUESTION_BANK["q5_seerah"], 8)
    for idx, s in enumerate(seerah, 1):
        pq = doc.add_paragraph(); set_paragraph_rtl(pq)
        pq.add_run(f"{idx}- {s}").font.name = 'Arial'

    # التوقيع
    p_sig = doc.add_paragraph()
    set_paragraph_rtl(p_sig, WD_ALIGN_PARAGRAPH.LEFT)
    rs = p_sig.add_run("معلم المادة\nحيدر محمد عبد الكريم")
    rs.font.name = 'Arial'; rs.bold = True

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# 5. عرض الاختيارات وزر التوليد
# ---------------------------------------------------------
st.info(f"📌 النموذج الحالي: **{exam_type}** | العام الدراسي: **{academic_year}** | الزمن: **{time_limit}**")

if st.button("🎲 توليد امتحان جديد بالخيارات المحددة"):
    st.session_state['exam_file'] = generate_exam(exam_type, academic_year, time_limit)
    st.success("تم توليد النموذج وسحب الأسئلة بنجاح!")

if 'exam_file' in st.session_state:
    st.download_button(
        label="📝 تحميل ملف Word النهائي",
        data=st.session_state['exam_file'],
        file_name=f"Exam_{exam_type}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
