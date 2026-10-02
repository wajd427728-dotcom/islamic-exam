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

st.title("منظومة توليد الامتحانات المباشرة - من بنك الأسئلة حصراً")
st.write("مدرسة الذاريات الابتدائية المختلطة | الصف الخامس الابتدائي")

# ---------------------------------------------------------
# بنك الأسئلة الخاص بك حصراً (لا يتم توليد أي سؤال خارجه)
# ---------------------------------------------------------
USER_QUESTION_BANK = {
    "q1_quran": [
        "أ / اكتب ما تحفظه من سورة ( الملك ) من قوله تعالى ( تَبَارَكَ الَّذِي بِيَدِهِ الْمُلْكُ ) إلى قوله تعالى ( عَذَابَ جَهَنَّمَ وَبِئْسَ الْمَصِيرُ )",
        "ب / اكتب ما تحفظه من سورة ( البلد ) من قوله تعالى ( لَا أُقْسِمُ بِهَذَا الْبَلَدِ ) إلى قوله تعالى ( وَهَدَيْنَاهُ النَّجْدَيْنِ )"
    ],
    "q2_meanings": ["مشفقون", "وما يسطرون", "طباقا", "حل", "كرتين", "هلوعا"],
    "q2_tafseer": [
        "ما المعنى العام للآية القرآنية التالية : بسم الله الرحمن الرحيم ( الَّذِينَ هُمْ عَلَى صَلَاتِهِمْ دَائِمُونَ ) ."
    ],
    "q3_hadith": [
        {"a": "اكتب حديثاً نبوياً شريفاً في ( التوبة ) ؟", "b": "اكتب حديثاً نبوياً شريفاً في ( حفظ اللسان ) ؟"}
    ],
    "q4_beliefs": [
        "الإنجيل هو الكتاب المنزل على النبي يوسف (ع) .",
        "آمن علماء اليهود والنصاري بالنبي محمد (ص) .",
        "التوراة هو الكتاب المنزل على النبي إبراهيم (ع) .",
        "من أسماء الله الحسنى المنتقم والودود و الشكور .",
        "يتصف جميع الأنبياء بالصدق والحكمة والصبر ومكارم الأخلاق .",
        "حارب الأنبياء الطواغيت والحكام الظالمين لنصرة المستضعفين وتحريرهم من سيطرة الطواغيت ."
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

def set_paragraph_rtl(paragraph, align=WD_ALIGN_PARAGRAPH.RIGHT):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.alignment = align

def generate_exam_from_strict_bank():
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # الترويسة الرسمية
    table = doc.add_table(rows=1, cols=3)
    table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table.columns[0].width = Inches(2.2)
    table.columns[1].width = Inches(3.0)
    table.columns[2].width = Inches(2.3)

    p_r = table.cell(0, 0).paragraphs[0]
    set_paragraph_rtl(p_r, WD_ALIGN_PARAGRAPH.RIGHT)
    r = p_r.add_run("المادة : التربية الإسلامية\nالصف : الخامس الابتدائي\nالزمن : ساعتان")
    r.font.name = 'Arial'; r.font.size = Pt(10); r.bold = True

    p_c = table.cell(0, 1).paragraphs[0]
    set_paragraph_rtl(p_c, WD_ALIGN_PARAGRAPH.CENTER)
    rc = p_c.add_run("بسم الله الرحمن الرحيم\nأسئلة امتحانات نهاية السنة\nللعام الدراسي 2026/2025 الدور الثاني")
    rc.font.name = 'Arial'; rc.font.size = Pt(11); rc.bold = True

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

    # السؤال الأول: القرآن الكريم
    add_section_header("القرآن الكريم : ( 20 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س1 : اجب عن احد الفرعين :").bold = True
    
    quran_q = USER_QUESTION_BANK["q1_quran"]
    for q in quran_q:
        pq = doc.add_paragraph(); set_paragraph_rtl(pq)
        pq.add_run(q).font.name = 'Arial'

    # السؤال الثاني: المعاني والتفسير
    add_section_header("المعاني والتفسير : ( 10 درجات )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س2 : اجب عن ما يلي :").bold = True

    p_w_title = doc.add_paragraph(); set_paragraph_rtl(p_w_title)
    p_w_title.add_run("أ / أعط معاني لخمس من الكلمات الآتيين :").font.name = 'Arial'
    
    meanings_list = USER_QUESTION_BANK["q2_meanings"]
    # يمكن ترتيب الكلمات عشوائياً فقط دون إضافة كلمات من الخارج
    shuffled_words = random.sample(meanings_list, len(meanings_list))
    words_str = "   ".join([f"{i+1}- {w}" for i, w in enumerate(shuffled_words)])
    
    p_w = doc.add_paragraph(); set_paragraph_rtl(p_w)
    p_w.add_run(f"( {words_str} )").font.name = 'Arial'

    tafseer_q = random.choice(USER_QUESTION_BANK["q2_tafseer"])
    p_t = doc.add_paragraph(); set_paragraph_rtl(p_t)
    p_t.add_run(f"ب / {tafseer_q}").font.name = 'Arial'

    # السؤال الثالث: الحديث الشريف
    add_section_header("الحديث الشريف : ( 15 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س3 : الإجابة عن احد الفرعين :").bold = True

    hadith_q = USER_QUESTION_BANK["q3_hadith"][0]
    p_h = doc.add_paragraph(); set_paragraph_rtl(p_h)
    p_h.add_run(f"أ / {hadith_q['a']}           ب / {hadith_q['b']}").font.name = 'Arial'

    # السؤال الرابع: العقائد والعبادات
    add_section_header("العقائد والعبادات : ( 15 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س4 : اجب عن الإجابة الصحيحة بكلمة ( صح ) وعن الإجابة الخاطئة بكلمة ( خطا ) :").bold = True

    beliefs_q = USER_QUESTION_BANK["q4_beliefs"]
    shuffled_beliefs = random.sample(beliefs_q, len(beliefs_q)) # إعادة ترتيب عشوائي لأسئلة البنك فقط
    for idx, item in enumerate(shuffled_beliefs, 1):
        pq = doc.add_paragraph(); set_paragraph_rtl(pq)
        pq.add_run(f"{idx}- {item}").font.name = 'Arial'

    # السؤال الخامس: السيرة النبوية
    add_section_header("السيرة النبوية والآداب الإسلامية : ( 20 درجة )")
    p = doc.add_paragraph(); set_paragraph_rtl(p)
    p.add_run("س5 : املأ الفراغات الآتية :").bold = True

    seerah_q = USER_QUESTION_BANK["q5_seerah"]
    shuffled_seerah = random.sample(seerah_q, len(seerah_q)) # إعادة ترتيب عشوائي لأسئلة البنك فقط
    for idx, item in enumerate(shuffled_seerah, 1):
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

st.subheader("توليد نموذج الامتحان")
if st.button("🎲 توليد نموذج من البنك"):
    st.session_state['strict_exam'] = generate_exam_from_strict_bank()
    st.success("تم توليد النموذج من أسئلة البنك المحددة حصراً!")

if 'strict_exam' in st.session_state:
    st.download_button(
        label="📝 تحميل ملف Word النهائي",
        data=st.session_state['strict_exam'],
        file_name="Islamic_Exam_Grade5_StrictBank.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
