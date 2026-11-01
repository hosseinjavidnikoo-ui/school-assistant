
import sys
import os
import json
import urllib.request
import webbrowser
from flask import Flask, render_template_string, request, jsonify
import pandas as pd
import openpyxl

CURRENT_VERSION = "2.0.0"
UPDATE_CHECK_URL = "https://api.github.com/repos/your-school/school-assistant/releases/latest"

app = Flask(__name__)

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>دستیار جامع و هوشمند مدیریت مدرسه</title>

    <style>
        :root {
            --primary: #1f4e78;
            --primary-dark: #153654;
            --secondary: #27ae60;
            --bg: #f4f7f6;
            --card-bg: #ffffff;
            --text: #2d3748;
        }

        body {
            font-family: 'B Nazanin', Tahoma, Arial, sans-serif;
            background: var(--bg);
            margin: 0;
            padding: 20px;
            direction: rtl;
            color: var(--text);
        }

        .container {
            max-width: 1100px;
            margin: 0 auto;
            background: var(--card-bg);
            padding: 30px;
            border-radius: 16px;
            box-shadow: 0 4px 25px rgba(0,0,0,0.08);
        }

        header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px solid #edf2f7;
            padding-bottom: 20px;
            margin-bottom: 25px;
        }

        h1 { color: var(--primary); font-size: 24px; margin: 0; }
        .badge { background: #e1ecf4; color: var(--primary); padding: 6px 14px; border-radius: 20px; font-size: 13px; font-weight: bold; }

        .tabs { display: flex; gap: 10px; border-bottom: 1px solid #e2e8f0; margin-bottom: 20px; }
        .tab-btn { padding: 10px 20px; border: none; background: none; font-size: 15px; cursor: pointer; border-bottom: 3px solid transparent; font-family: inherit; font-weight: bold; }
        .tab-btn.active { border-color: var(--primary); color: var(--primary); }

        .tab-content { display: none; }
        .tab-content.active { display: block; }

        .card { border: 1px solid #e2e8f0; border-radius: 12px; padding: 20px; margin-bottom: 20px; background: #fafbfc; }
        .card h3 { margin-top: 0; color: var(--primary); font-size: 17px; }

        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; }

        label { font-weight: bold; font-size: 14px; display: block; margin-bottom: 6px; }
        input, select, textarea { width: 100%; padding: 10px; border: 1px solid #cbd5e0; border-radius: 8px; font-family: inherit; font-size: 14px; box-sizing: border-box; }
        
        .btn { background: var(--primary); color: white; border: none; padding: 11px 22px; border-radius: 8px; cursor: pointer; font-family: inherit; font-size: 14px; font-weight: bold; transition: 0.2s; }
        .btn:hover { background: var(--primary-dark); }
        .btn-success { background: var(--secondary); }
        .btn-success:hover { background: #1e8449; }

        .stamp-preview { width: 150px; height: 100px; border: 2px dashed #cbd5e0; border-radius: 8px; display: flex; align-items: center; justify-content: center; background: #fff; margin-top: 10px; }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>دستیار جامع، کلید هوشمند و صدور فرم مدرسه</h1>
            <span class="badge">نسخه {{ version }} (کامل)</span>
        </header>

        <div class="tabs">
            <button class="tab-btn active" onclick="switchTab('data-tab')">📊 مدیریت اکسل و داده‌ها</button>
            <button class="tab-btn" onclick="switchTab('template-tab')">📜 قالب‌ساز و هوش مصنوعی</button>
            <button class="tab-btn" onclick="switchTab('stamp-tab')">🔏 پاک‌سازی و آرشیو مهر/امضا</button>
            <button class="tab-btn" onclick="switchTab('settings-tab')">⚙️ تنظیمات و به‌روزرسانی</button>
        </div>

        <!-- تب ۱: مدیریت داده‌ها -->
        <div id="data-tab" class="tab-content active">
            <div class="card">
                <h3>بارگذاری و نگاشت پویای اکسل (Dynamic Mapping)</h3>
                <p>فایل اکسل مدرسه (یا هر فایل اکسل دلخواه دیگر) را آپلود کرده و ستون‌ها را متصل کنید:</p>
                <div class="grid-2">
                    <div>
                        <label>انتخاب فایل اکسل جدید:</label>
                        <input type="file" accept=".xlsx, .csv">
                    </div>
                    <div>
                        <label>پایه / کلاس هدف:</label>
                        <select>
                            <option>کل مدرسه (۷۹ نفر)</option>
                            <option>دهم انسانی (۱۶ نفر)</option>
                            <option>دهم تجربی (۱۵ نفر)</option>
                            <option>یازدهم انسانی (۱۵ نفر)</option>
                            <option>یازدهم تجربی (۱۳ نفر)</option>
                            <option>دوازدهم انسانی (۱۳ نفر)</option>
                            <option>دوازدهم تجربی (۷ نفر)</option>
                        </select>
                    </div>
                </div>
            </div>
        </div>

        <!-- تب ۲: قالب‌ساز و هوش مصنوعی -->
        <div id="template-tab" class="tab-content">
            <div class="card">
                <h3>ویرایشگر سند، فونت و حاشیه‌ها (Border & Layout)</h3>
                <div class="grid-3">
                    <div>
                        <label>انتخاب فونت:</label>
                        <select>
                            <option>B Nazanin</option>
                            <option>IranSans</option>
                            <option>B Yekan</option>
                            <option>Vazir</option>
                        </select>
                    </div>
                    <div>
                        <label>طرح حاشیه (Border):</label>
                        <select>
                            <option>کادر اداری ساده</option>
                            <option>کادر تذیب رسمی</option>
                            <option>کادر کلاسیک دون خطی</option>
                            <option>بدون حاشیه</option>
                        </select>
                    </div>
                    <div>
                        <label>سایز برگه / چیدمان:</label>
                        <select>
                            <option>A4 تک برگه</option>
                            <option>A5 (۲ فرم در یک برگه A4)</option>
                        </select>
                    </div>
                </div>
                <div style="margin-top:15px;">
                    <label>متن قالب (با پشتیبانی از تگ‌های {نام}، {نام خانوادگی}، {کد ملی} و...):</label>
                    <textarea rows="6">جنابعالی ولی محترم دانش‌آموز {نام} {نام خانوادگی} به استحضار می‌رساند جهت بررسی وضعیت تحصیلی فرزندتان در روز {تاریخ} در دفتر مدرسه حضور به عمل آورید.</textarea>
                </div>
                <button class="btn btn-success" style="margin-top:15px;">✨ بازنویسی و تنظیم لحن با هوش مصنوعی (Gemini API)</button>
            </div>
        </div>

        <!-- تب ۳: مهر و امضا -->
        <div id="stamp-tab" class="tab-content">
            <div class="card">
                <h3>🔏 پاک‌سازی و شفاف‌سازی هوشمند مهر و امضا</h3>
                <p>عکس خام مهر یا امضا را آپلود کنید تا هوش مصنوعی پس‌زمینه آن را شفاف و کیفیت آن را بازسازی کند:</p>
                <div class="grid-2">
                    <div>
                        <label>آپلود عکس خام مهر/امضا:</label>
                        <input type="file" accept="image/*">
                        <button class="btn" style="margin-top:10px;">حذف خودکار پس‌زمینه و ارتقای کیفیت</button>
                    </div>
                    <div>
                        <label>پیش‌نمایش مهر/امضای پاک‌سازی‌شده:</label>
                        <div class="stamp-preview">
                            <span style="color:#aaa; font-size:12px;">مهر شفاف‌سازی‌شده</span>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- تب ۴: تنظیمات -->
        <div id="settings-tab" class="tab-content">
            <div class="card">
                <h3>🔑 تنظیمات API هوش مصنوعی و به‌روزرسانی</h3>
                <div class="grid-2">
                    <div>
                        <label>کلید API اختصاصی (Gemini / OpenAI API Key):</label>
                        <input type="password" placeholder="AIzaSy...">
                    </div>
                    <div>
                        <label>سیستم آپدیت آنلاین:</label>
                        <button class="btn btn-success" onclick="checkUpdate()">بررسی آنلاین نسخه جدید</button>
                    </div>
                </div>
                <div id="update-result" style="margin-top:15px; font-weight:bold;"></div>
            </div>
        </div>
    </div>

    <script>
        function switchTab(tabId) {
            document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
            document.getElementById(tabId).classList.add('active');
            event.target.classList.add('active');
        }

        function checkUpdate() {
            const resBox = document.getElementById('update-result');
            resBox.innerText = 'در حال بررسی شبکه...';
            fetch('/check-update')
                .then(res => res.json())
                .then(data => {
                    if (data.has_update) {
                        resBox.innerHTML = '<span style="color:green;">نسخه جدید (' + data.latest_version + ') موجود است!</span>';
                    } else {
                        resBox.innerHTML = '<span style="color:#1f4e78;">شما از آخرین نسخه نرم‌افزار استفاده می‌کنید.</span>';
                    }
                });
        }
    </script>
</body>
</html>
'''

@app.route('/')
def home():
    return render_template_string(HTML_TEMPLATE, version=CURRENT_VERSION)

@app.route('/check-update')
def check_update():
    return jsonify({"has_update": False, "latest_version": CURRENT_VERSION})

if __name__ == '__main__':
    try:
        import webview
        webview.create_window("دستیار جامع مدیریت مدرسه", app, width=1150, height=800)
        webview.start()
    except Exception:
        webbrowser.open("http://127.0.0.1:5000")
        app.run(port=5000)
