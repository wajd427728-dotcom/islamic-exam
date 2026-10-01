import streamlit as st
import pandas as pd
import random

# ضبط إعدادات الصفحة
st.set_page_config(page_title="منظومة امتحانات التربية الإسلامية", layout="centered")

st.markdown("<h2 style='text-align: center;'>منظومة توليد الامتحانات المباشرة</h2>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: #555;'>مدرسة الذاريات الابتدائية المختلطة</h3>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>الصف: الخامس الابتدائي | العام الدراسي: 2026 / 2027 م</p>", unsafe_allow_html=True)
st.markdown("---")

EXCEL_FILE = 'بنك_اسئلة_التربية_الاسلامية.xlsx'

# التحقق من وجود WeasyPrint
try:
    from weasyprint import HTML
    WEASYPRINT_AVAILABLE = True
except ImportError:
    WEASYPRINT_AVAILABLE = False

try:
    xls = pd.ExcelFile(EXCEL_FILE)
    sheet_names = xls.sheet_names
    st.success("تم الاتصال ببنك الأسئلة بنجاح! 🟢")
except Exception as e:
    st.error(f"تنبيه: لم يتم العثور على ملف Excel لبنك الأسئلة ({EXCEL_FILE}). الرجاء التأكد من رفعه على GitHub بنفس الاسم.")
    sheet_names = []

if st.button("📄 توليد ورقة الامتحان الرسمية (PDF - 80 درجة)", type="primary", use_container_width=True):
    if not sheet_names:
        st.error("الرجاء رفع ملف الـ Excel الخاص ببنك الأسئلة أولاً.")
    elif not WEASYPRINT_AVAILABLE:
        st.error("جاري تحميل مكتبة الـ PDF على السيرفر، يرجى إعادة تحديث الصفحة بعد دقيقة.")
    else:
        try:
            # قراءة الأسئلة عشوائياً
            df_mne = pd.read_excel(xls, 'المعاني والتفسير')
            vocab_sample = df_mne.sample(n=min(6, len(df_mne)))
            interp_sample = df_mne.dropna(subset=['المعنى العام']).sample(n=1) if 'المعنى العام' in df_mne.columns else df_mne.sample(n=1)

            df_had = pd.read_excel(xls, 'الحديث الشريف')
            hadith_sample = df_had.sample(n=1).iloc[0]

            df_aqd = pd.read_excel(xls, 'العقائد والعبادات')
            aqd_sample = df_aqd.sample(n=min(2, len(df_aqd)))

            df_sir = pd.read_excel(xls, 'السيرة النبوية والآداب الإسلامية')
            sir_sample = df_sir.sample(n=min(2, len(df_sir)))

            # بناء تصميم نموذج الامتحان
            html_template = f"""
            <!DOCTYPE html>
            <html lang="ar" dir="rtl">
            <head>
            <meta charset="UTF-8">
            <style>
                @page {{ size: A4; margin: 12mm; }}
                body {{ font-family: 'Amiri', 'Arial', sans-serif; direction: rtl; font-size: 13pt; line-height: 1.5; color: #000; }}
                .header-table {{ width: 100%; border-collapse: collapse; border: 2px solid #000; margin-bottom: 12px; }}
                .header-table td {{ padding: 6px; border: 1px solid #000; text-align: center; vertical-align: middle; }}
                .question {{ margin-bottom: 12px; border: 1px solid #444; padding: 8px; border-radius: 4px; }}
                .q-title {{ font-weight: bold; border-bottom: 1px dashed #000; padding-bottom: 4px; margin-bottom: 6px; }}
                .verse {{ font-weight: bold; text-align: center; margin: 6px 0; }}
                .footer-table {{ width: 100%; margin-top: 25px; text-align: center; font-weight: bold; }}
            </style>
            </head>
            <body>

            <table class="header-table">
                <tr>
                    <td style="width: 35%;">جمهورية العراق<br>وزارة التربية<br>مدرسة الذاريات الابتدائية المختلطة</td>
                    <td style="width: 43%; font-weight: bold; background-color: #f8f9fa;">
                        امتحان مادة التربية الإسلامية<br>الصف الخامس الابتدائي<br>العام الدراسي: 2026 / 2027 م
                    </td>
                    <td style="width: 22%; font-weight: bold;">
                        الدرجة الكلية: 80<br>
                        <div style="border: 1px solid #000; margin-top: 5px; height: 35px; line-height: 35px; font-size: 16pt;"> / 80</div>
                    </td>
                </tr>
            </table>

            <table style="width: 100%; margin-bottom: 10px; font-weight: bold;">
                <tr>
                    <td>اسم التلميذ: ...........................................................</td>
                    <td>الشعبة: ........</td>
                    <td>الزمن: ساعة ونصف</td>
                </tr>
            </table>

            <!-- س1 -->
            <div class="question">
                <div class="q-title">السؤال الأول: القرآن الكريم (20 درجة)</div>
                <div>أكتب من سورة (الملك) من قوله تعالى: <span class="verse">«تَبَارَكَ الَّذِي بِيَدِهِ الْمُلْكُ ...»</span> إلى قوله تعالى: <span class="verse">«... وَهُوَ الْعَزِيزُ الْغَفُورُ»</span>.</div>
            </div>

            <!-- س2 -->
            <div class="question">
                <div class="q-title">السؤال الثاني: المعاني والتفسير (10 درجات)</div>
                <div><strong>أ) [5 درجات]</strong> بين معاني الكلمات الآتية لـ (خمس) فقط:<br>
                    ( {' | '.join(vocab_sample['الكلمة أو الآية'].tolist())} )
                </div>
                <div style="margin-top: 8px;">
                    <strong>ب) [5 درجات]</strong> ما المعنى العام للآية الكريمة التالية:<br>
                    <div class="verse">«{interp_sample.iloc[0]['الكلمة أو الآية']}»</div>
                </div>
            </div>

            <!-- س3 -->
            <div class="question">
                <div class="q-title">السؤال الثالث: الحديث الشريف (15 درجة)</div>
                <div>أكتب حديثاً نبوياً شريفاً في: <strong>({hadith_sample['موضوع الحديث']})</strong>.</div>
            </div>

            <!-- س4 -->
            <div class="question">
                <div class="q-title">السؤال الرابع: العقائد والعبادات (15 درجة)</div>
                <div>أجب عن الأسئلة الآتية:<br>
                    {'<br>'.join([f"{i+1}. {q}" for i, q in enumerate(aqd_sample['موضوع السؤال'].tolist())])}
                </div>
            </div>

            <!-- س5 -->
            <div class="question">
                <div class="q-title">السؤال الخامس: السيرة النبوية والآداب الإسلامية (20 درجة)</div>
                <div>أكمل الفراغات الآتية بما يناسبها:<br>
                    {'<br>'.join([f"{i+1}. {q}" for i, q in enumerate(sir_sample['السؤال او الفراغ'].tolist())])}
                </div>
            </div>

            <table class="footer-table">
                <tr>
                    <td>مدرس المادة: حيدر محمد عبد الكريم</td>
                    <td>توقيع الدفتر الامتحان / اللجنة الامتحانية</td>
                </tr>
            </table>

            </body>
            </html>
            """

            pdf_filename = "exam_output.pdf"
            HTML(string=html_template).write_pdf(pdf_filename)

            with open(pdf_filename, "rb") as pdf_file:
                st.success("تم توليد ورقة الامتحان بنجاح!")
                st.download_button(
                    label="📥 تحميل ملف الامتحان (PDF)",
                    data=pdf_file,
                    file_name="امتحان_التربية_الاسلامية_الخامس_الابتدائي.pdf",
                    mime="application/pdf"
                )
        except Exception as ex:
            st.error(f"حدث خطأ أثناء معالجة الأسئلة: {ex}")
