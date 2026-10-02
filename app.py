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
# بنك الأسئلة الشامل (يمكنك إضافة أي أسئلة جديدة هنا مستقبلاً)
# ---------------------------------------------------------
QUESTION_BANK = {
    "q1_quran": [
        "أ / اكتب ما تحفظه من سورة ( الملك ) من قوله تعالى ( تَبَارَكَ الَّذِي بِيَدِهِ الْمُلْكُ ) إلى قوله تعالى ( عَذَابَ جَهَنَّمَ وَبِئْسَ الْمَصِيرُ )",
        "ب / اكتب ما تحفظه من سورة ( البلد ) من قوله تعالى ( لَا أُقْسِمُ بِهَذَا الْبَلَدِ ) إلى قوله تعالى ( وَهَدَيْنَاهُ النَّجْدَيْنِ )",
        "جـ / اكتب ما تحفظه من سورة ( الأعلى ) من قوله تعالى ( سَبِّحِ اسْمَ رَبِّكَ الْأَعْلَى ) إلى قوله تعالى ( فَنَسَى )",
        "د / اكتب ما تحفظه من سورة ( المعارج ) من قوله تعالى ( سَأَلَ سَائِلٌ بِعَذَابٍ وَاقِعٍ ) إلى قوله تعالى ( لَّيْسَ لَهُ دَافِعٌ )"
    ],
    "q2_meanings": [
        "مشفقون", "وما يسطرون", "طباقا", "حل", "كرتين", "هلوعا", 
        "سوى", "قدر فهدى", "الغثاء", "أحوى", "فلا تنسى", "النجدين"
    ],
    "q2_tafseer": [
        "ما المعنى العام للآية القرآنية التالية : بسم الله الرحمن الرحيم ( الَّذِينَ هُمْ عَلَى صَلَاتِهِمْ دَائِمُونَ ) .",
        "ما المعنى العام للآية القرآنية التالية : بسم الله الرحمن الرحيم ( وَالَّذِينَ فِي أَمْوَالِهِمْ حَقٌّ مَّعْلُومٌ ) .",
        "ما المعنى العام للآية القرآنية التالية : بسم الله الرحمن الرحيم ( الَّذِي خَلَقَ فَسَوَّى ) ."
    ],
    "q3_hadith": [
        {"a": "اكتب حديثاً نبوياً شريفاً في ( التوبة ) ؟", "b": "اكتب حديثاً نبوياً شريفاً في ( حفظ اللسان ) ؟"},
        {"a": "اكتب حديثاً نبوياً شريفاً في ( إعادة العارية ) ؟", "b": "اكتب حديثاً نبوياً شريفاً في ( الحث على الصدق ) ؟"}
    ],
    "q4_beliefs": [
        "الإنجيل هو الكتاب المنزل على النبي يوسف (ع) .",
        "آمن علماء اليهود والنصاري بالنبي محمد (ص) .",
        "التوراة هو الكتاب المنزل على النبي إبراهيم (ع) .",
        "من أسماء الله الحسنى المنتقم والودود و الشكور .",
        "يتصف جميع الأنبياء بالصدق والحكمة والصبر ومكارم الأخلاق .",
        "حارب الأنبياء الطواغيت والحكام الظالمين لنصرة المستضعفين وتحريرهم من سيطرة الطواغيت .",
        "القرآن الكريم هو الكتاب الخاتم الذي نزل على النبي محمد (ص) ."
    ],
    "q5_seerah": [
        "مرت الدعوة الإسلامية بمرحلتين ________________ و ________________ .",
        "يرجع نسب النبي أيوب (ع) إلى النبي ________________ .",
        "أول الآيات التي نزلت على النبي محمد (ص) كانت من سورة ________________ .",
        "كان اسم المدينة المنورة قبل مجيء الرسول اليها تسمى ________________ .",
        "سمى القرآن الكريم يوم معركة بدر بيوم ________________ .",
        "من ابرز شهداء معركة احد مصعب بن عمير و ________________ .",
        "تبعد المدينة المنورة عن مكة ________________ .",
        "سميت السور التي نزلت بمكة بالسور ________________ التي نزلت بالمدينة بالسور ________________ ."
    ]
}

# ضبط اتجاه النص والأقواس من اليمين إلى اليسار
def set_paragraph_rtl(paragraph, align=WD_ALIGN_PARAGRAPH.RIGHT):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.alignment = align

def generate_exam_from_bank():
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # الترويسة الثلاثية المعتمدة
    table = doc.add_table(rows=1, cols=3)
    table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width = Inches(2.2)
    table.columns[1].width = Inches(3.0)
    table.columns[2].width = Inches(2.3)

    # الخلية اليمنى: معلومات المادة والصف
    p_r = table.cell(0, 0).paragraphs[0]
    set_paragraph_rtl(p_r, WD_ALIGN_PARAGRAPH.RIGHT)
    r = p_r.add_run("المادة : التربية الإسلامية\nالصف : الخامس الابتدائي\nالزمن : ساعتان")
    r.font.name = 'Arial'; r.font.size = Pt(10); r.bold = True

    # الخلية الوسطى: نوع الامتحان والعام الدراسي
    p_c = table.cell(0, 1).paragraphs[0]
    set_paragraph_rtl(p_c, WD_ALIGN_PARAGRAPH.CENTER)
    rc = p_c.add_run("بسم الله الرحمن الرحيم\nأسئلة امتحانات نهاية السنة\nللعام الدراسي 2026/2025 الدور الثاني")
    rc.font.name = 'Arial'; rc.font.size = Pt(11); rc.bold = True

    # الخلية اليسرى: إدارية المدرسة
    p_l = table.cell(0, 2).paragraphs[0]
    set_paragraph_rtl(p_l, WD_ALIGN_PARAGRAPH.LEFT)
    rl = p_l.add_run("ادارة\nمدرسة الذاريات\nالابتدائية المختلطة")
    rl.font.name = 'Arial'; rl.font.size = Pt(10); rl.bold = True

    doc.add_paragraph()

    def add_section_header(title_text):
        p = doc.add_paragraph()
        set_paragraph_rtl(p, WD_ALIGN_PARAGRAPH.RIGHT)
        r = p.add_run(title_text)
        r.font.name = 'Arial'; r.font.size = Pt(12); r.bold = True

    # 1. القرآن الكريم (سحب فرعين عشوائياً من البنك)
    add_section_header("القرآن الكريم : ( 20 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س1 : اجب عن احد الفرعين :").bold = True
    
    selected_quran = random.sample(QUESTION_BANK["q1_quran"], 2)
    p_a = doc.add_paragraph(); set_paragraph_rtl(p_a)
    p_a.add_run(f"أ / {selected_quran[0].split(' / ', 1)[-1]}").font.name = 'Arial'
    p_b = doc.add_paragraph(); set_paragraph_rtl(p_b)
    p_b.add_run(f"ب / {selected_quran[1].split(' / ', 1)[-1]}").font.name = 'Arial'

    # 2. المعاني والتفسير (سحب 6 كلمات وسؤال تفسير عشوائياً)
    add_section_header("المعاني والتفسير : ( 10 درجات )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س2 : اجب عن ما يلي :").bold = True

    p_w_title = doc.add_paragraph(); set_paragraph_rtl(p_w_title)
    p_w_title.add_run("أ / أعط معاني لخمس من الكلمات الآتيين :").font.name = 'Arial'
    
    selected_words = random.sample(QUESTION_BANK["q2_meanings"], 6)
    words_str = "   ".join([f"{i+1}- {w}" for i, w in enumerate(selected_words)])
    p_w = doc.add_paragraph(); set_paragraph_rtl(p_w)
    p_w.add_run(f"( {words_str} )").font.name = 'Arial'

    selected_tafseer = random.choice(QUESTION_BANK["q2_tafseer"])
    p_t = doc.add_paragraph(); set_paragraph_rtl(p_t)
    p_t.add_run(f"ب / {selected_tafseer}").font.name = 'Arial'

    # 3. الحديث الشريف
    add_section_header("الحديث الشريف : ( 15 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س3 : الإجابة عن احد الفرعين :").bold = True

    selected_hadith = random.choice(QUESTION_BANK["q3_hadith"])
    p_h = doc.add_paragraph(); set_paragraph_rtl(p_h)
    p_h.add_run(f"أ / {selected_hadith['a']}           ب / {selected_hadith['b']}").font.name = 'Arial'

    # 4. العقائد والعبادات (سحب 6 أسئلة عشوائياً)
    add_section_header("العقائد والعبادات : ( 15 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س4 : اجب عن الإجابة الصحيحة بكلمة ( صح ) وعن الإجابة الخاطئة بكلمة ( خطا ) :").bold = True

    selected_beliefs = random.sample(QUESTION_BANK["q4_beliefs"], 6)
    for idx, item in enumerate(selected_beliefs, 1):
        pq = doc.add_paragraph(); set_paragraph_rtl(pq)
        pq.add_run(f"{idx}- {item}").font.name = 'Arial'

    # 5. السيرة النبوية (سحب 8 فراغات عشوائياً)
    add_section_header("السيرة النبوية والآداب الإسلامية : ( 20 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س5 : املأ الفراغات الآتية :").bold = True

    selected_seerah = random.sample(QUESTION_BANK["q5_seerah"], 8)
    for idx, item in enumerate(selected_seerah, 1):
        pq = doc.add_paragraph(); set_paragraph_rtl(pq)
        pq.add_run(f"{idx}- {item}").font.name = 'Arial'

    # التوقيع
    p_sig = doc.add_paragraph()
    set_paragraph_rtl(p_sig, WD_ALIGN_PARAGRAPH.LEFT)
    rs = p_sig.add_run("معلم المادة\nحيدر محمد عبد الكريم")
    rs.font.name = 'Arial'; rs.bold = True

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# واجهة المستخدم
if st.button("🎲 توليد نموذج امتحان جديد عشوائياً من البنك"):
    st.session_state['exam_file'] = generate_exam_from_bank()
    st.success("تم سحب وتوليد أسئلة جديدة من بنك الأسئلة بنجاح!")

if 'exam_file' in st.session_state:
    st.download_button(
        label="📝 تحميل ملف Word النهائي",
        data=st.session_state['exam_file'],
        file_name="Islamic_Exam_Grade5_Generated.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
