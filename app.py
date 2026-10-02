import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import weasyprint

# 1. إنشاء ملف Word متقن
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

table = doc.add_table(rows=1, cols=3)
table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
table.columns[0].width = Inches(2.2)
table.columns[1].width = Inches(3.0)
table.columns[2].width = Inches(2.3)

p_r = table.cell(0, 0).paragraphs[0]
set_p_rtl(p_r, WD_ALIGN_PARAGRAPH.RIGHT)
r = p_r.add_run("المادة : التربية الإسلامية\nالصف : الخامس الابتدائي\nالزمن : ساعة واحدة")
r.font.name = 'Arial'; r.font.size = Pt(10); r.bold = True

p_c = table.cell(0, 1).paragraphs[0]
set_p_rtl(p_c, WD_ALIGN_PARAGRAPH.CENTER)
rc = p_c.add_run("بسم الله الرحمن الرحيم\nأسئلة امتحانات الشهر الأول\nللعام الدراسي 2026/2025")
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
p_a.add_run("أ / اكتب ما تحفظه من سورة ( الملك ) من قوله تعالى ( تَبَارَكَ الَّذِي بِيَدِهِ الْمُلْكُ ) إلى قوله تعالى ( عَذَابَ جَهَنَّمَ وَبِئْسَ الْمَصِيرُ )").font.name = 'Arial'

p_b = doc.add_paragraph(); set_p_rtl(p_b)
p_b.add_run("ب / اكتب ما تحفظه من سورة ( البلد ) من قوله تعالى ( لَا أُقْسِمُ بِهَذَا الْبَلَدِ ) إلى قوله تعالى ( وَهَدَيْنَاهُ النَّجْدَيْنِ )").font.name = 'Arial'

add_section_header("المعاني والتفسير : ( 10 درجات )")
p = doc.add_paragraph(); set_p_rtl(p)
p.add_run("س2 : أجب عن ما يلي :").bold = True

p_w_title = doc.add_paragraph(); set_p_rtl(p_w_title)
p_w_title.add_run("أ / أعط معاني لخمس من الكلمات الآتية :").font.name = 'Arial'

words_str = "1- فلا تنسى    2- هلوعا    3- الغثاء    4- النجدين    5- أحوى    6- حل"
p_w = doc.add_paragraph(); set_p_rtl(p_w)
p_w.add_run(f"( {words_str} )").font.name = 'Arial'

p_t = doc.add_paragraph(); set_p_rtl(p_t)
p_t.add_run("ب / ما المعنى العام للآية الكريمة : ( الَّذِينَ هُمْ عَلَى صَلَاتِهِمْ دَائِمُونَ ) ؟").font.name = 'Arial'

add_section_header("الحديث الشريف : ( 15 درجة )")
p = doc.add_paragraph(); set_p_rtl(p)
p.add_run("س3 : الإجابة عن أحد الفرعين :").bold = True

p_h = doc.add_paragraph(); set_p_rtl(p_h)
p_h.add_run("أ / اكتب حديثاً نبوياً شريفاً في ( التوبة ) ؟           ب / اكتب حديثاً نبوياً شريفاً في ( حفظ اللسان ) ؟").font.name = 'Arial'

add_section_header("العقائد والعبادات : ( 15 درجة )")
p = doc.add_paragraph(); set_p_rtl(p)
p.add_run("س4 : أجب بكلمة ( صح ) عن العبارة الصحيحة وكلمة ( خطأ ) عن العبارة الخاطئة :").bold = True

beliefs = [
    "آمن علماء اليهود والنصارى بالنبي محمد ( ص ) .",
    "الإنجيل هو الكتاب المنزل على النبي يوسف ( ع ) .",
    "التوراة هو الكتاب المنزل على النبي إبراهيم ( ع ) .",
    "من أسماء الله الحسنى المنتقم والودود والشكور .",
    "يتصف جميع الأنبياء بالصدق والحكمة والصبر ومكارم الأخلاق .",
    "حارب الأنبياء الطواغيت والحكام الظالمين لنصرة المستضعفين وتحريرهم ."
]
for idx, b in enumerate(beliefs, 1):
    pq = doc.add_paragraph(); set_p_rtl(pq)
    pq.add_run(f"{idx}- {b}").font.name = 'Arial'

add_section_header("السيرة النبوية والآداب الإسلامية : ( 20 درجة )")
p = doc.add_paragraph(); set_p_rtl(p)
p.add_run("س5 : املأ الفراغات الآتية :").bold = True

seerah = [
    "كان اسم المدينة المنورة قبل مجيء الرسول إليها يسمى ______________ .",
    "سميت السور التي نزلت بمكة بالسور ______________ والتي نزلت بالمدينة بالسور ______________ .",
    "يرجع نسب النبي أيوب ( ع ) إلى النبي ______________ .",
    "أول الآيات التي نزلت على النبي محمد ( ص ) كانت من سورة ______________ .",
    "مرت الدعوة الإسلامية بمرحلتين ______________ و ______________ .",
    "من أبرز شهداء معركة أحد مصعب بن عمير و ______________ .",
    "تبعد المدينة المنورة عن مكة المكرمة مسافة ______________ .",
    "سمى القرآن الكريم يوم معركة بدر بيوم ______________ ."
]
for idx, s in enumerate(seerah, 1):
    pq = doc.add_paragraph(); set_p_rtl(pq)
    pq.add_run(f"{idx}- {s}").font.name = 'Arial'

p_sig = doc.add_paragraph()
set_p_rtl(p_sig, WD_ALIGN_PARAGRAPH.LEFT)
rs = p_sig.add_run("معلم المادة\nحيدر محمد عبد الكريم")
rs.font.name = 'Arial'; rs.bold = True

doc.save("Exam_Final_v5.docx")

# 2. إنشاء ملف PDF بتنسيق HTML مدمج
html_content = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<style>
    @page { size: A4; margin: 1cm 1.2cm; }
    body { font-family: 'DejaVu Sans', 'Arial', sans-serif; direction: rtl; text-align: right; font-size: 11pt; line-height: 1.4; color: #000; }
    .header-table { width: 100%; border-collapse: collapse; margin-bottom: 10px; border-bottom: 2px solid #000; padding-bottom: 5px; }
    .header-table td { vertical-align: top; font-weight: bold; }
    .right-header { text-align: right; width: 33%; font-size: 10pt; }
    .center-header { text-align: center; width: 34%; font-size: 11pt; }
    .left-header { text-align: left; width: 33%; font-size: 10pt; }
    .section-title { font-size: 11.5pt; font-weight: bold; margin-top: 8px; margin-bottom: 3px; background-color: #f0f0f0; padding: 2px 6px; border-radius: 3px; }
    .question-title { font-weight: bold; margin-top: 4px; margin-bottom: 2px; }
    .item { margin-right: 12px; margin-bottom: 2px; }
    .signature { margin-top: 15px; text-align: left; font-weight: bold; font-size: 10pt; }
</style>
</head>
<body>

<table class="header-table">
    <tr>
        <td class="right-header">المادة : التربية الإسلامية<br>الصف : الخامس الابتدائي<br>الزمن : ساعة واحدة</td>
        <td class="center-header">بسم الله الرحمن الرحيم<br>أسئلة امتحانات الشهر الأول<br>للعام الدراسي 2026/2025</td>
        <td class="left-header">إدارة<br>مدرسة الذاريات<br>الابتدائية المختلطة</td>
    </tr>
</table>

<div class="section-title">القرآن الكريم : ( 20 درجة )</div>
<div class="question-title">س1 : أجب عن أحد الفرعين :</div>
<div class="item">أ / اكتب ما تحفظه من سورة ( الملك ) من قوله تعالى ( تَبَارَكَ الَّذِي بِيَدِهِ الْمُلْكُ ) إلى قوله تعالى ( عَذَابَ جَهَنَّمَ وَبِئْسَ الْمَصِيرُ )</div>
<div class="item">ب / اكتب ما تحفظه من سورة ( البلد ) من قوله تعالى ( لَا أُقْسِمُ بِهَذَا الْبَلَدِ ) إلى قوله تعالى ( وَهَدَيْنَاهُ النَّجْدَيْنِ )</div>

<div class="section-title">المعاني والتفسير : ( 10 درجات )</div>
<div class="question-title">س2 : أجب عن ما يلي :</div>
<div class="item">أ / أعط معاني لخمس من الكلمات الآتية :</div>
<div class="item" style="text-align: center; font-weight: bold;">( 1- فلا تنسى &nbsp;&nbsp;&nbsp; 2- هلوعا &nbsp;&nbsp;&nbsp; 3- الغثاء &nbsp;&nbsp;&nbsp; 4- النجدين &nbsp;&nbsp;&nbsp; 5- أحوى &nbsp;&nbsp;&nbsp; 6- حل )</div>
<div class="item">ب / ما المعنى العام للآية الكريمة : ( الَّذِينَ هُمْ عَلَى صَلَاتِهِمْ دَائِمُونَ ) ؟</div>

<div class="section-title">الحديث الشريف : ( 15 درجة )</div>
<div class="question-title">س3 : الإجابة عن أحد الفرعين :</div>
<div class="item">أ / اكتب حديثاً نبوياً شريفاً في ( التوبة ) ؟ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ب / اكتب حديثاً نبوياً شريفاً في ( حفظ اللسان ) ؟</div>

<div class="section-title">العقائد والعبادات : ( 15 درجة )</div>
<div class="question-title">س4 : أجب بكلمة ( صح ) عن العبارة الصحيحة وكلمة ( خطأ ) عن العبارة الخاطئة :</div>
<div class="item">1- آمن علماء اليهود والنصارى بالنبي محمد ( ص ) .</div>
<div class="item">2- الإنجيل هو الكتاب المنزل على النبي يوسف ( ع ) .</div>
<div class="item">3- التوراة هو الكتاب المنزل على النبي إبراهيم ( ع ) .</div>
<div class="item">4- من أسماء الله الحسنى المنتقم والودود والشكور .</div>
<div class="item">5- يتصف جميع الأنبياء بالصدق والحكمة والصبر ومكارم الأخلاق .</div>
<div class="item">6- حارب الأنبياء الطواغيت والحكام الظالمين لنصرة المستضعفين وتحريرهم .</div>

<div class="section-title">السيرة النبوية والآداب الإسلامية : ( 20 درجة )</div>
<div class="question-title">س5 : املأ الفراغات الآتية :</div>
<div class="item">1- كان اسم المدينة المنورة قبل مجيء الرسول إليها يسمى ________________ .</div>
<div class="item">2- سميت السور التي نزلت بمكة بالسور ________________ والتي نزلت بالمدينة بالسور ________________ .</div>
<div class="item">3- يرجع نسب النبي أيوب ( ع ) إلى النبي ________________ .</div>
<div class="item">4- أول الآيات التي نزلت على النبي محمد ( ص ) كانت من سورة ________________ .</div>
<div class="item">5- مرت الدعوة الإسلامية بمرحلتين ________________ و ________________ .</div>
<div class="item">6- من أبرز شهداء معركة أحد مصعب بن عمير و ________________ .</div>
<div class="item">7- تبعد المدينة المنورة عن مكة المكرمة مسافة ________________ .</div>
<div class="item">8- سمى القرآن الكريم يوم معركة بدر بيوم ________________ .</div>

<div class="signature">
    معلم المادة<br>
    حيدر محمد عبد الكريم
</div>

</body>
</html>
"""

weasyprint.HTML(string=html_content).write_pdf("Exam_Final_v5.pdf")
