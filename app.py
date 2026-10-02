import streamlit as st
import pandas as pd
import random
import os

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="منظومة امتحانات التربية الإسلامية", layout="centered")

st.markdown("<h2 style='text-align: center;'>منظومة توليد الامتحانات المباشرة</h2>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: #555;'>مدرسة الذاريات الابتدائية المختلطة</h3>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>الصف: الخامس الابتدائي | العام الدراسي: 2026 / 2027 م | معلم المادة: حيدر محمد عبد الكريم</p>", unsafe_allow_html=True)
st.markdown("---")

# البحث التلقائي عن ملف الإكسل
excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~$')]

if excel_files:
    EXCEL_FILE = excel_files[0]
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        st.success(f"تم الاتصال ببنك الأسئلة بنجاح! 🟢")
    except Exception as e:
        st.error(f"خطأ في قراءة ملف الإكسل: {e}")
        xls = None
else:
    st.error("لم يتم العثور على ملف إكسل لبنك الأسئلة في المستودع.")
    xls = None

if st.button("📄 توليد ورقة الامتحان الرسمية (PDF - 80 درجة)", type="primary", use_container_width=True):
    if xls is None:
        st.error("الرجاء التأكد من توفر ملف بنك الأسئلة.")
    else:
        try:
            # دالة قراءة مرنة تطابق الاسم مع تنظيف المسافات الزائدة
            sheet_map = {s.strip(): s for s in xls.sheet_names}

            def read_sheet(target_name, default_index):
                if target_name in sheet_map:
                    return pd.read_excel(xls, sheet_map[target_name])
                # مطابقة جزئية في حال وجود اختلاف بسيط في الأحرف
                for clean_name, actual_name in sheet_map.items():
                    if target_name in clean_name or clean_name in target_name:
                        return pd.read_excel(xls, actual_name)
                # الرجوع للترتيب الرقمي كخيار أمان أخيرة
                return pd.read_excel(xls, xls.sheet_names[min(default_index, len(xls.sheet_names)-1)])

            # قراءة البيانات بالأسماء المطابقة لـ Google Sheets
            df_qur = read_sheet('القران الكريم', 0)
            df_mne = read_sheet('المعاني والتفسير', 1)
            df_had = read_sheet('الحديث الشريف', 2)
            df_aqd = read_sheet('العقائد والعبادات', 3)
            df_sir = read_sheet('السيرة النبوية والاداب الاسلامية', 4)

            # اختيار الأسئلة
            vocab_sample = df_mne.sample(n=min(3, len(df_mne)))
            interp_sample = df_mne.dropna(subset=[df_mne.columns[1]]).sample(n=1) if len(df_mne.columns) > 1 else df_mne.sample(n=1)
            hadith_sample = df_had.sample(n=1).iloc[0]
            aqd_sample = df_aqd.sample(n=min(2, len(df_aqd)))
            sir_sample = df_sir.sample(n=min(2, len(df_sir)))

            # إنشاء ملف PDF
            pdf_filename = "exam_output.pdf"
            doc = SimpleDocTemplate(pdf_filename, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            story = []

            styles = getSampleStyleSheet()
            arabic_style = ParagraphStyle(
                'ArabicStyle',
                parent=styles['Normal'],
                fontName='Helvetica',
                fontSize=11,
                leading=16,
                alignment=2
            )

            # ترويسة الامتحان الرسمية
            header_data = [
                [
                    Paragraph("<b>جمهورية العراق<br/>وزارة التربية<br/>مدرسة الذاريات الابتدائية المختلطة</b>", arabic_style),
                    Paragraph("<b>امتحان مادة التربية الإسلامية<br/>الصف الخامس الابتدائي<br/>العام الدراسي: 2026 / 2027 م</b>", arabic_style),
                    Paragraph("<b>الدرجة الكلية: 80<br/><br/>[ &nbsp;&nbsp;&nbsp;&nbsp; / 80 ]</b>", arabic_style)
                ]
            ]
            t_header = Table(header_data, colWidths=[150, 245, 110])
            t_header.setStyle(TableStyle([
                ('BOX', (0,0), (-1,-1), 1.5, colors.black),
                ('INNERGRID', (0,0), (-1,-1), 0.5, colors.black),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ]))
            story.append(t_header)
            story.append(Spacer(1, 10))

            # بيانات التلميذ
            info_data = [[
                Paragraph("<b>اسم التلميذ:</b> ...........................................................", arabic_style),
                Paragraph("<b>الشعبة:</b> ........", arabic_style),
                Paragraph("<b>الزمن:</b> ساعة ونصف", arabic_style)
            ]]
            t_info = Table(info_data, colWidths=[310, 100, 95])
            story.append(t_info)
            story.append(HRFlowable(width="100%", thickness=1, color=colors.black, spaceAfter=10))

            # استخراج النصوص للأسئلة
            vocab_col = vocab_sample.columns[0]
            vocab_text = " ، ".join(vocab_sample[vocab_col].astype(str).tolist())
            
            interp_col = interp_sample.columns[0]
            interp_text = str(interp_sample.iloc[0][interp_col])

            hadith_col = hadith_sample.index[0]
            hadith_text = str(hadith_sample[hadith_col])

            aqd_col = aqd_sample.columns[0]
            q4_text = "<br/>".join([f"{i+1}. {q}" for i, q in enumerate(aqd_sample[aqd_col].astype(str).tolist())])

            sir_col = sir_sample.columns[0]
            q5_text = "<br/>".join([f"{i+1}. {q}" for i, q in enumerate(sir_sample[sir_col].astype(str).tolist())])

            # جدول الأسئلة
            questions_content = [
                ("السؤال الأول: القرآن الكريم (20 درجة)", "أكتب من سورة (الملك) من قوله تعالى: ( تَبَارَكَ الَّذِي بِيَدِهِ الْمُلْكُ ... ) إلى قوله تعالى: ( ... وَهُوَ الْعَزِيزُ الْغَفُورُ )."),
                ("السؤال الثاني: المعاني والتفسير (10 درجات)", f"أ) [5 درجات] بين معاني الكلمات الآتية:<br/>( {vocab_text} )<br/><br/>ب) [5 درجات] ما المعنى العام للآية الكريمة التالية:<br/><b>( {interp_text} )</b>"),
                ("السؤال الثالث: الحديث الشريف (15 درجة)", f"أكتب حديثاً نبوياً شريفاً في: <b>( {hadith_text} )</b>."),
                ("السؤال الرابع: العقائد والعبادات (15 درجة)", f"أجب عن الأسئلة الآتية:<br/>{q4_text}"),
                ("السؤال الخامس: السيرة النبوية والآداب الإسلامية (20 درجة)", f"أكمل الفراغات الآتية بما يناسبها:<br/>{q5_text}")
            ]

            for q_title, q_body in questions_content:
                q_html = f"<b>{q_title}</b><br/>{q_body}"
                t_q = Table([[Paragraph(q_html, arabic_style)]], colWidths=[505])
                t_q.setStyle(TableStyle([
                    ('BOX', (0,0), (-1,-1), 1, colors.gray),
                    ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#fdfdfd")),
                    ('PADDING', (0,0), (-1,-1), 6),
                ]))
                story.append(t_q)
                story.append(Spacer(1, 8))

            # التوقيع والمدرس
            footer_data = [[
                Paragraph("<b>مدرس المادة: حيدر محمد عبد الكريم</b>", arabic_style),
                Paragraph("<b>توقيع اللجنة الامتحانية / الإدارة</b>", arabic_style)
            ]]
            t_footer = Table(footer_data, colWidths=[250, 255])
            story.append(Spacer(1, 15))
            story.append(t_footer)

            doc.build(story)

            with open(pdf_filename, "rb") as pdf_file:
                st.success("تم توليد ورقة الامتحان بنجاح تام! 🟢")
                st.download_button(
                    label="📥 تحميل ملف الامتحان (PDF)",
                    data=pdf_file,
                    file_name="نموذج_امتحان_التربية_الاسلامية_الخامس_الابتدائي.pdf",
                    mime="application/pdf"
                )
        except Exception as ex:
            st.error(f"حدث خطأ أثناء معالجة البيانات: {ex}")
