import pandas as pd
import random
import os
from flask import Flask, render_template_string, request, send_file
from weasyprint import HTML

app = Flask(_name_)

# مسار ملف الاكسل الخاص ببنك الأسئلة
EXCEL_FILE = 'بنك_اسئلة_التربية_الاسلامية.xlsx'

@app.route('/')
def home():
    return '''
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <title>منظومة الامتحانات - مدرسة الذاريات الابتدائية</title>
        <style>
            body { font-family: Arial, sans-serif; text-align: center; margin-top: 50px; background-color: #f4f6f9; }
            .card { background: white; padding: 30px; border-radius: 10px; display: inline-block; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
            .btn { background-color: #27ae60; color: white; padding: 12px 24px; text-decoration: none; font-size: 18px; border-radius: 5px; font-weight: bold; }
            .btn:hover { background-color: #219150; }
        </style>
    </head>
    <body>
        <div class="card">
            <h2>منظومة توليد الامتحانات المباشرة</h2>
            <h3>مدرسة الذاريات الابتدائية المختلطة</h3>
            <p>الصف: الخامس الابتدائي | المادة: التربية الإسلامية</p>
            <br>
            <a href="/generate_exam" class="btn">📄 توليد ورقة امتحان PDF جديدة (80 درجة)</a>
        </div>
    </body>
    </html>
    '''

@app.route('/generate_exam')
def generate_exam():
    # 1. قراءة البيانات من شيتات Excel
    xls = pd.ExcelFile(EXCEL_FILE)
    
    # سحب سؤال القرآن الكريم (س1 - 20 درجة)
    df_qur = pd.read_excel(xls, 'القرآن الكريم') if 'القرآن الكريم' in xls.sheet_names else None
    
    # سحب المعاني والتفسير (س2 - 10 درجات)
    df_mne = pd.read_excel(xls, 'المعاني والتفسير')
    vocab_sample = df_mne.sample(n=min(6, len(df_mne)))  # 6 كلمات يختار منها 5
    interp_sample = df_mne.dropna(subset=['المعنى العام']).sample(n=1) if 'المعنى العام' in df_mne.columns else df_mne.sample(n=1)

    # سحب الحديث الشريف (س3 - 15 درجة)
    df_had = pd.read_excel(xls, 'الحديث الشريف')
    hadith_sample = df_had.sample(n=1).iloc[0]

    # سحب العقائد والعبادات (س4 - 15 درجة)
    df_aqd = pd.read_excel(xls, 'العقائد والعبادات')
    aqd_sample = df_aqd.sample(n=min(2, len(df_aqd)))

    # سحب السيرة والآداب (س5 - 20 درجة)
    df_sir = pd.read_excel(xls, 'السيرة النبوية والآداب الإسلامية')
    sir_sample = df_sir.sample(n=min(2, len(df_sir)))

    # 2. بناء قالب HTML للـ PDF
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

    # 3. حفظ الـ PDF وإرساله للمستخدم
    pdf_path = "exam_output.pdf"
    HTML(string=html_template).write_pdf(pdf_path)
    return send_file(pdf_path, as_attachment=True, download_name="امتحان_التربية_الاسلامية_الخامس_الابتدائي.pdf")

if _name_ == '_main_':
    app.run(debug=True)
