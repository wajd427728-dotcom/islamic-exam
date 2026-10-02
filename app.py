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

excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~$')]

if excel_files:
    EXCEL_FILE = excel_files[0]
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        st.success("تم الاتصال ببنك الأسئلة بنجاح! 🟢")
    except Exception as e:
        st.error(f"خطأ في قراءة الملف: {e}")
        xls = None
else:
    st.error("لم يتم العثور على ملف الإكسل في المستودع.")
    xls = None

if st.button("📄 توليد ورقة الامتحان الرسمية (PDF - 80 درجة)", type="primary", use_container_width=True):
    if xls is None:
        st.error("الرجاء التأكد من وجود ملف بنك الأسئلة.")
    else:
        try:
            sheets = xls.sheet_names

            def get_clean_df(sheet_idx):
                if sheet_idx < len(sheets):
                    df = pd.read_excel(xls, sheets[sheet_idx])
                    return df.dropna(how='all')
                return pd.DataFrame()

            df_qur = get_clean_df(0)
            df_mne = get_clean_df(1)
            df_had = get_clean_df(2)
            df_aqd = get_clean_df(3)
            df_sir = get_clean_df(4)

            # سحب العينات بأمان تام حسب المتوفر
            vocab_n = min(3, len(df_mne)) if len(df_mne) > 0 else 1
            vocab_sample = df_mne.sample(n=vocab_n) if vocab_n > 0 else pd.DataFrame()
            
            interp_sample = df_mne.sample(n=1) if len(df_mne) > 0 else pd.DataFrame()
            hadith_sample = df_had.sample(n=1).iloc[0] if len(df_had) > 0 else None
            
            aqd_n = min(2, len(df_aqd)) if len(df_aqd) > 0 else 1
            aqd_sample = df_aqd.sample(n=aqd_n) if aqd_n > 0 else pd.DataFrame()
            
            sir_n = min(2, len(df_sir)) if len(df_sir) > 0 else 1
            sir_sample = df_sir.sample(n=sir_n) if sir_n > 0 else pd.DataFrame()

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

            info_data = [[
                Paragraph("<b>اسم التلميذ:</b> ...........................................................", arabic_style),
                Paragraph("<b>الشعبة:</b> ........", arabic_style),
                Paragraph("<b>الزمن:</b> ساعة ونصف", arabic_style)
            ]]
            t_info = Table(info_data, colWidths=[310, 100, 95])
            story.append(t_info)
            story.append(HRFlowable(width="100%", thickness=1, color=colors.black, spaceAfter=10))

            vocab_text = " ، ".join(vocab_sample.iloc[:, 0].astype(str).tolist()) if not vocab_sample.empty else ""
            
            if not interp_sample.empty and len(interp_sample.columns) > 1:
                interp_text = f"{str(interp_sample.iloc[0, 0])} : {str(interp_sample.iloc[0, 1])}"
            elif not interp_sample.empty:
                interp_text = str(interp_sample.iloc[0, 0])
            else:
                interp_text = ""

            hadith_text = str(hadith_sample.iloc[0]) if hadith_sample is not None else ""
            q4_text = "<br/>".join([f"{i+1}. {str(q)}" for i, q in enumerate(aqd_sample.iloc[:, 0].tolist())]) if not aqd_sample.empty else ""
            q5_text = "<br/>".join([f"{i+1}. {str(q)}" for i, q in enumerate(sir_sample.iloc[:, 0].tolist())]) if not sir_sample.empty else ""

            questions_content = [
                ("السؤال الأول: القرآن الكريم (20 درجة)", "أكتب من سورة (الملك) من قوله تعالى: ( تَبَارَكَ الَّذِي بِيَدِهِ الْمُلْكُ ... ) إلى قوله تعالى: ( ... وَهُوَ الْعَزِيزُ الْغَفُورُ )."),
                ("السؤال الثاني: المعاني والتفسير (10 درجات)", f"أ) [5 درجات] بين معاني الكلمات الآتية:<br/>( {vocab_text} )<br/><br/>ب) [5 درجات] ما المعنى العام للآية الكريمة أو الكلمة التالية:<br/><b>( {interp_text} )</b>"),
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
