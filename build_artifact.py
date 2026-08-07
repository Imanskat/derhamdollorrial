#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ساخت نسخه قابل انتشار (Artifact) از روی index.html

صفحه منتشرشده در claude.ai داخل یک قالب <html><head></head><body> پیچیده
می‌شود، پس نباید تگ‌های wrapper خودمان را داشته باشد. ضمناً سیاست امنیتی
آن صفحه هر درخواست شبکه‌ای به میزبان بیرونی را می‌بندد، بنابراین
«به‌روزرسانی خودکار نرخ» آنجا کار نمی‌کند و باید صادقانه همین را بگوییم.

اجرا: python3 build_artifact.py
"""

import re
import pathlib

SRC = pathlib.Path("index.html")
DST = pathlib.Path("dist/artifact.html")

html = SRC.read_text(encoding="utf-8")

style = re.search(r"<style>(.*?)</style>", html, re.S).group(1)
body = re.search(r"<body>(.*?)</body>", html, re.S).group(1)

# جهت راست‌به‌چپ روی <html> بود؛ حالا باید روی خود body بنشیند
style = style.replace(
    "body{\n  background:var(--plane);",
    "body{\n  direction:rtl;\n  background:var(--plane);",
)

# دکمه واکشی خودکار در صفحه میزبانی‌شده بی‌فایده است — جایش توضیح می‌گذاریم
body = body.replace(
    '<button type="button" id="fetchBtn" class="primary">به‌روزرسانی خودکار نرخ</button>',
    '<span class="chip" id="fetchBtn" style="cursor:default">دریافت خودکار در این نسخه غیرفعال است</span>',
)
body = body.replace(
    '<span class="chip" id="fxChip"><span class="dot warn"></span>'
    '<span id="fxChipText">نرخ نمونه — هنوز به‌روزرسانی نشده</span></span>',
    '<span class="chip" id="fxChip"><span class="dot warn"></span>'
    '<span id="fxChipText">نرخ نمونه — لطفاً نرخ روز را وارد کنید</span></span>',
)

# کادر هشدار نرخ ارز: توضیح تفاوت این نسخه با فایل محلی
body = body.replace(
    "این ابزار فقط نرخ آزاد را مبنا می‌گیرد — اگر به‌روزرسانی خودکار جواب نداد، نرخ را دستی از صرافی وارد کنید.",
    "این ابزار فقط نرخ آزاد را مبنا می‌گیرد. <strong>در این نسخه تحت وب، دریافت خودکار نرخ "
    "ممکن نیست</strong> (صفحه اجازه درخواست به سرویس بیرونی ندارد)؛ نرخ روز را از صرافی "
    "بگیرید و در کادر بالا وارد کنید. نسخه <code>index.html</code> داخل مخزن، دکمه دریافت "
    "خودکار هم دارد.",
)

# فراخوانی واکشی را حذف می‌کنیم تا خطای شبکه در کنسول نیفتد
body = body.replace(
    'document.getElementById("fetchBtn").addEventListener("click", fetchRates);',
    "/* دریافت خودکار در نسخه منتشرشده غیرفعال است */",
)

DST.parent.mkdir(parents=True, exist_ok=True)
DST.write_text(
    f"<style>{style}</style>\n{body}",
    encoding="utf-8",
)
print("نوشته شد:", DST, f"({DST.stat().st_size:,} بایت)")
