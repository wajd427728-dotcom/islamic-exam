import streamlit as st
import pandas as pd
import docx
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

st.set_page_config(page_title="منظومة امتحانات التربية الإسلامية", page_icon="📝", layout="centered")

st.title("📝 منظومة توليد امتحانات التربية الإسلامية")
st.subheader("الصف الخامس الابتدائي - مدرسة الذاريات الابتدائية المختلطة")

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def generate_word_exam(q1_text, q2_a_text, q2_b_text, q3_text, q4_text, q5_text):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)

    header_table = doc.add_table(rows=1, cols=3)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths = [Inches(2.3), Inches(2.8), Inches(2.3)]
    row = header_table.rows[0]
    
    p_right = row.cells[2].paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_r = p_right.add_run("إدارة\nمدرسة الذاريات\nالابتدائية المختلطة")
    r_r.bold = True
    r_r.font.size = Pt(10.5)

    p_mid = row.cells[1].paragraphs[0]
    p_mid.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_m = p_mid.add_run("بسم الله الرحمن الرحيم\nأسئلة امتحان ( الفصل الأول - الشهر الأول )\nللعام الدراسي 2026 / 2027\nالمادة: التربية الإسلامية")
    r_m.bold = True
    r_m.font.size = Pt(11)

    p_left = row.cells[0].paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_l = p_left.add_run("الصف : الخامس الابتدائي\nالشعبة : (   )\nالوقت : ساعتان")
    r_l.bold = True
    r_l.font.size = Pt(10.5)

    for i, cell in enumerate(row.cells):
        cell.width = col_widths[i]

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    p_hr = doc.add_paragraph()
    p_hr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_hr = p_hr.add_run("―" * 55)
    r_hr.bold = True
    p_hr.paragraph_format.space_after = Pt(4)

    p_note = doc.add_paragraph()
    p_note.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_note = p_note.add_run("ملاحظة: أجب عن جميع الأسئلة الآتية:")
    r_note.bold = True
    p_note.paragraph_format.space_after = Pt(4)

    def add_q_banner(q_text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(7.4)
        set_cell_background(cell, "E6E6E6")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(q_text)
        run.bold = True
        run.font.size = Pt(10.5)
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    add_q_banner("س1 / القرآن الكريم (20 درجة)")
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p1.add_run(q1_text)
    p1.paragraph_format.space_after = Pt(6)

    add_q_banner("س2 / المعاني والتفسير (10 درجات)")
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run(f"أ) المعاني (5 درجات):\n{q2_a_text}\n\nب) المعنى العام (5 درجات):\n{q2_b_text}")
    p2.paragraph_format.space_after = Pt(6)

    add_q_banner("س3 / الحديث النبوي الشريف (15 درجة)")
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p3.add_run(q3_text)
    p3.paragraph_format.space_after = Pt(6)

    add_q_banner("س4 / العقيدة والعبادات (15 درجة)")
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p4.add_run(q4_text)
    p4.paragraph_format.space_after = Pt(6)

    add_q_banner("س5 / السيرة والآداب (20 درجة)")
    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p5.add_run(q5_text)
    p5.paragraph_format.space_after = Pt(12)

    p_ft = doc.add_paragraph()
    p_ft.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_ft = p_ft.add_run("معلم المادة:\nحيدر محمد عبد الكريم")
    r_ft.bold = True
    r_ft.font.size = Pt(10.5)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

st.success("تم الاتصال بالمنظومة بنجاح!")
st.markdown("### ⚙️ تخصيص محتوى الامتحان")

q1_input = st.text_area("س1 / القرآن الكريم (20 درجة):", "اكتب ما تحفظه من إحدى السورتين الآتيتين:\n1. سورة الأعلى من قوله تعالى: ((سَبِّحِ اسْمَ رَبِّكَ الأَعْلَى)) إلى قوله تعالى: ((فَهَدَى)).\n2. سورة البلد من قوله تعالى: ((لا أُقْسِمُ بِهَذَا الْبَلَدِ)) إلى قوله تعالى: ((أَصْحَابُ الْمَيْمَنَةِ)).")
q2_a_input = st.text_area("س2 - فرع أ (المعاني - 5 درجات):", "بيّن معاني خمس من الكلمات الآتية:\n(1. غُثَاءً أحْوَى ، 2. الَّذِي خَلَقَ فَسَوَّى ، 3. النَّجْدَيْنِ ، 4. فَلَكْتَحَمَ الْعَقَبَةَ ، 5. كَبَدٍ ، 6. قَدَّرَ فَهَدَى)")
q2_b_input = st.text_area("س2 - فرع ب (المعنى العام - 5 درجات):", "فسّر قوله تعالى: ((إِنَّ هَذَا لَفِي الصُّحُفِ الأُولَى)).")
q3_input = st.text_area("س3 / الحديث النبوي الشريف (15 درجة):", "اكتب حديثاً نبوياً شريفاً يحثّ على (التوبة) ضبطاً بالشكل، مع ذكر فائدة واحدة من التوبة.")
q4_input = st.text_area("س4 / العقيدة والعبادات (15 درجة):", "أجب عن الأسئلة الآتية:\n1. لماذا أرسل الله تعالى الأنبياء والرسل إلى البشر؟\n2. ضع كلمة (صح) أو (خطأ) أمام العبارة الآتية: صلاة الفجر تتكون من أربع ركعات. (   )")
q5_input = st.text_area("س5 / السيرة والآداب (20 درجة):", "إملأ الفراغات الآتية بالكلمات المناسبة:\n1. مرت الدعوة الإسلامية بمرحلتين هما: .................... و ....................\n2. ولد النبي محمد (صلى الله عليه وآله وسلم) في عام ....................\n3. من صفات المسلم الأدب مع .................... ووالديه.")

if st.button("🚀 توليد ورقة الامتحان (Word)"):
    docx_file = generate_word_exam(q1_input, q2_a_input, q2_b_input, q3_input, q4_input, q5_input)
    st.download_button(
        label="📥 تحميل مستند الامتحان (.docx)",
        data=docx_file,
        file_name="امتحان_التربية_الإسلامية_الصف_الخامس.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
