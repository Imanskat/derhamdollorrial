#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ساخت فایل اکسل «بهای تمام‌شده واردات لوازم آرایشی» (دبی ← بندرعباس ← تهران)

همان مدل محاسباتی داشبورد وب، این بار به‌صورت فرمول‌های زنده اکسل.
هیچ عددی در ستون‌های محاسباتی hardcode نشده — همه از سلول‌های ورودی می‌آیند.

اجرا:  python3 build_workbook.py
خروجی: بهای-تمام-شده-واردات.xlsx
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "بهای-تمام-شده-واردات.xlsx"

FONT = "Tahoma"          # فونت استاندارد و حرفه‌ای برای متن فارسی
N_ROWS = 25              # تعداد ردیف‌های آماده برای کالا

# ---------------------------------------------------------------- سبک‌ها
BLUE  = Font(name=FONT, size=10, color="0000FF")           # ورودی دستی
BLACK = Font(name=FONT, size=10, color="000000")           # فرمول
GREEN = Font(name=FONT, size=10, color="008000")           # ارجاع به شیت دیگر
BOLD  = Font(name=FONT, size=10, bold=True)
TITLE = Font(name=FONT, size=13, bold=True)
SECT  = Font(name=FONT, size=11, bold=True, color="1F3864")
SMALL = Font(name=FONT, size=8.5, color="595959")

YELLOW = PatternFill("solid", fgColor="FFFF99")            # سلول‌هایی که باید پر شوند
HEADFILL = PatternFill("solid", fgColor="1F3864")
HEADFONT = Font(name=FONT, size=9, bold=True, color="FFFFFF")
DERIVFILL = PatternFill("solid", fgColor="F2F2F2")
TOTFILL = PatternFill("solid", fgColor="E2EFDA")

thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

RIAL = '#,##0;(#,##0);-'
AED  = '#,##0.00;(#,##0.00);-'
PCT  = '0.0%;(0.0%);-'
NUM  = '#,##0;(#,##0);-'
NUM1 = '#,##0.0;(#,##0.0);-'

wb = Workbook()

# ======================================================================
#  شیت ۱ — پارامترها
# ======================================================================
p = wb.active
p.title = "پارامترها"
p.sheet_view.rightToLeft = True

def put(ws, cell, value, font=BLACK, fmt=None, fill=None, border=False, align=None):
    c = ws[cell]
    c.value = value
    c.font = font
    if fmt:  c.number_format = fmt
    if fill: c.fill = fill
    if border: c.border = BOX
    if align: c.alignment = Alignment(horizontal=align, vertical="center")
    return c

p.column_dimensions["A"].width = 42
p.column_dimensions["B"].width = 18
p.column_dimensions["C"].width = 52

put(p, "A1", "پارامترهای محاسبه بهای تمام‌شده واردات", TITLE)
put(p, "A2", "کالاها از دبی خریداری، به بندرعباس حمل و از آنجا به تهران منتقل می‌شوند.", SMALL)

put(p, "A4", "۱ — نرخ ارز (بازار آزاد)", SECT)
put(p, "A5", "دلار آمریکا → ریال")
put(p, "B5", 1_150_000, BLUE, RIAL, YELLOW, True)
put(p, "C5", "نرخ بازار آزاد. نرخ رسمی/نیما را وارد نکنید.", SMALL)

put(p, "A6", "درهم امارات → ریال")
put(p, "B6", 313_138, BLUE, RIAL, YELLOW, True)
put(p, "C6", "مبنای اصلی همه محاسبات خرید. از صرافی بگیرید.", SMALL)

put(p, "A7", "دلار → درهم (نرخ میخکوب)")
put(p, "B7", 3.6725, BLUE, '0.0000', YELLOW, True)
put(p, "C7", "درهم به دلار میخکوب است و تقریباً ثابت می‌ماند.", SMALL)

put(p, "A8", "درهم بر حسب دلار (کنترلی)")
put(p, "B8", "='پارامترها'!$B$5/'پارامترها'!$B$7", BLACK, RIAL, DERIVFILL, True)
put(p, "C8", "اگر با سلول B6 اختلاف زیاد دارد، یکی از نرخ‌ها اشتباه است.", SMALL)

put(p, "A9", "نرخ درهم در روز خرید محموله")
put(p, "B9", 313_138, BLUE, RIAL, YELLOW, True)
put(p, "C9", "برای سنجش ریسک ارزی در شیت «خلاصه» استفاده می‌شود.", SMALL)

put(p, "A10", "تاریخ و منبع نرخ")
put(p, "B10", "دستی", BLUE, None, YELLOW, True)
put(p, "C10", "مثلاً: ۱۴۰۵/۰۵/۱۶ — صرافی/tgju", SMALL)

put(p, "A12", "۲ — پارامترهای محموله", SECT)
put(p, "A13", "کرایه دبی ← بندرعباس (درهم بر کیلوگرم)")
put(p, "B13", 2.5, BLUE, AED, YELLOW, True)
put(p, "C13", "حمل دریایی گروپاژ. حمل هوایی چند برابر است.", SMALL)

put(p, "A14", "بیمه (٪ ارزش کالا)")
put(p, "B14", 0.005, BLUE, PCT, YELLOW, True)

put(p, "A15", "حقوق ورودی گمرک (٪ ارزش CIF)")
put(p, "B15", 0.40, BLUE, PCT, YELLOW, True)
put(p, "C15", "برای لوازم آرایشی معمولاً بالاست — از ترخیص‌کار خود بگیرید.", SMALL)

put(p, "A16", "حمل بندرعباس ← تهران (ریال بر کیلوگرم)")
put(p, "B16", 120_000, BLUE, RIAL, YELLOW, True)

put(p, "A17", "ترخیص، انبارداری و اسناد (ریال — کل محموله)")
put(p, "B17", 900_000_000, BLUE, RIAL, YELLOW, True)

put(p, "A18", "سایر هزینه‌های محموله (ریال — کل محموله)")
put(p, "B18", 450_000_000, BLUE, RIAL, YELLOW, True)
put(p, "C18", "مجوز، برچسب اصالت، آزمایشگاه، حق‌العمل", SMALL)

put(p, "A19", "جمع هزینه‌های مقطوع محموله")
put(p, "B19", "='پارامترها'!$B$17+'پارامترها'!$B$18", BLACK, RIAL, DERIVFILL, True)
put(p, "C19", "این مبلغ بین کالاها سرشکن می‌شود.", SMALL)

put(p, "A20", "مبنای سرشکن‌کردن هزینه‌های مقطوع")
put(p, "B20", "وزن", BLUE, None, YELLOW, True)
put(p, "C20", "«وزن» یا «ارزش» — از فهرست کشویی انتخاب کنید.", SMALL)

put(p, "A21", "ارزش افزوده (٪)")
put(p, "B21", 0.10, BLUE, PCT, YELLOW, True)

put(p, "A22", "ارزش افزوده گمرک جزو بهای تمام‌شده است؟")
put(p, "B22", "خیر", BLUE, None, YELLOW, True)
put(p, "C22", "اگر مؤدی ثبت‌نام‌شده هستید «خیر» (با VAT فروش تهاتر می‌شود).", SMALL)

dv_basis = DataValidation(type="list", formula1='"وزن,ارزش"', allow_blank=False)
p.add_data_validation(dv_basis); dv_basis.add(p["B20"])
dv_yn = DataValidation(type="list", formula1='"بله,خیر"', allow_blank=False)
p.add_data_validation(dv_yn); dv_yn.add(p["B22"])

put(p, "A24", "راهنمای رنگ‌ها", SECT)
put(p, "A25", "متن آبی روی زمینه زرد = سلول ورودی؛ فقط این‌ها را تغییر دهید.", BLUE)
put(p, "A26", "متن سیاه = فرمول؛ دست نزنید.", BLACK)
put(p, "A27", "متن سبز = ارجاع به شیت دیگر.", GREEN)

put(p, "A29", "هشدار درباره نرخ ارز", SECT)
for i, line in enumerate([
    "سرویس‌های عمومی ارز (Google، exchangerate و مانند آن) نرخ «رسمی» ریال را می‌دهند، نه نرخ «بازار آزاد».",
    "اگر بهای تمام‌شده را با نرخ رسمی حساب کنید، عدد چند برابر کمتر از واقعیت درمی‌آید و کالا را زیر قیمت می‌فروشید.",
    "همیشه نرخ آزاد روز را از صرافی یا بازار بگیرید و در سلول B6 وارد کنید.",
]):
    put(p, f"A{30+i}", line, SMALL)

# ======================================================================
#  شیت ۲ — کالاها
# ======================================================================
s = wb.create_sheet("کالاها")
s.sheet_view.rightToLeft = True

P = "'پارامترها'!"
R_USD, R_AED  = f"{P}$B$5", f"{P}$B$6"
R_BUY         = f"{P}$B$9"
FREIGHT       = f"{P}$B$13"
INS           = f"{P}$B$14"
DUTY          = f"{P}$B$15"
INLAND        = f"{P}$B$16"
LUMP          = f"{P}$B$19"
BASIS         = f"{P}$B$20"
VAT           = f"{P}$B$21"
VATCOST       = f"{P}$B$22"

COLS = [
    # (عنوان، عرض، قالب، نوع)   نوع: in=ورودی، fx=فرمول
    ("نام محصول",                        26, None, "in"),
    ("قیمت خرید هر واحد\n(درهم)",        13, AED,  "in"),
    ("تعداد",                             9, NUM,  "in"),
    ("وزن کل\n(کیلوگرم)",                11, NUM1, "in"),
    ("هزینه جانبی هر واحد\n(ریال)",      15, RIAL, "in"),
    ("قیمت روز بازار تهران\n(ریال)",     17, RIAL, "in"),
    ("قیمت فروش من\n(ریال، با VAT)",     17, RIAL, "in"),
    ("هزینه فروش\n(٪)",                  11, PCT,  "in"),
    ("روز خرید\nتا فروش",                10, NUM,  "in"),
    ("ارزش کالا FOB\n(ریال)",            16, RIAL, "fx"),
    ("کرایه دریایی\n(ریال)",             14, RIAL, "fx"),
    ("بیمه\n(ریال)",                     12, RIAL, "fx"),
    ("ارزش CIF\n(ریال)",                 16, RIAL, "fx"),
    ("حقوق ورودی گمرک\n(ریال)",          16, RIAL, "fx"),
    ("حمل داخلی\n(ریال)",                14, RIAL, "fx"),
    ("سهم از هزینه‌های مقطوع\n(ریال)",   17, RIAL, "fx"),
    ("هزینه جانبی کل\n(ریال)",           14, RIAL, "fx"),
    ("ارزش افزوده گمرک\n(ریال)",         16, RIAL, "fx"),
    ("بهای تمام‌شده کل\n(ریال)",         17, RIAL, "fx"),
    ("بهای تمام‌شده هر واحد\n(ریال)",    17, RIAL, "fx"),
    ("فروش خالص هر واحد\n(ریال)",        16, RIAL, "fx"),
    ("هزینه فروش هر واحد\n(ریال)",       15, RIAL, "fx"),
    ("سود هر واحد\n(ریال)",              14, RIAL, "fx"),
    ("حاشیه سود\n(٪)",                   11, PCT,  "fx"),
    ("سود کل ردیف\n(ریال)",              16, RIAL, "fx"),
    ("نقطه سربه‌سر\nقیمت فروش (ریال)",   16, RIAL, "fx"),
    ("نرخ درهم سربه‌سر\n(ریال)",         15, RIAL, "fx"),
    ("حاشیه امنیت ارزی\n(٪)",            14, PCT,  "fx"),
    ("بازده سالانه‌شده\n(٪)",            13, PCT,  "fx"),
    ("فاصله تا بازار تهران\n(٪)",        15, PCT,  "fx"),
    ("ضریب ارزی A\n(کمکی)",              14, NUM,  "fx"),
    ("ثابت ریالی B\n(کمکی)",             16, RIAL, "fx"),
    ("سود در دسترس\n(کمکی)",             16, RIAL, "fx"),
]

put(s, "A1", "بهای تمام‌شده و سود هر کالا", TITLE)
put(s, "A2", "فقط ستون‌های آبی/زرد (A تا I) را پر کنید؛ بقیه خودکار محاسبه می‌شوند.", SMALL)

HEAD_ROW = 4
FIRST = HEAD_ROW + 1
LAST = FIRST + N_ROWS - 1

for i, (title, width, fmt, kind) in enumerate(COLS, start=1):
    L = get_column_letter(i)
    s.column_dimensions[L].width = width
    c = s.cell(row=HEAD_ROW, column=i, value=title)
    c.font = HEADFONT
    c.fill = HEADFILL
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BOX
s.row_dimensions[HEAD_ROW].height = 42
s.freeze_panes = "B5"

SAMPLE = [
    ("کرم پودر (۳۰ میلی‌لیتر)",       32,  600, 42, 400_000, 40_600_000, 37_600_000, 0.12,  75),
    ("رژ لب مایع",                    14, 1200, 30, 250_000, 17_700_000, 16_400_000, 0.12,  60),
    ("پالت سایه چشم",                 48,  300, 39, 500_000, 61_500_000, 56_900_000, 0.15, 110),
    ("سرم ویتامین ث (۳۰ میلی‌لیتر)",  55,  400, 32, 450_000, 68_700_000, 63_600_000, 0.10,  90),
    ("ریمل حجم‌دهنده",                18,  900, 27, 300_000, 22_700_000, 21_000_000, 0.12,  65),
]

W_SUM = f"SUM($D${FIRST}:$D${LAST})"          # جمع وزن
V_SUM = f"SUMPRODUCT($B${FIRST}:$B${LAST},$C${FIRST}:$C${LAST})"  # جمع ارزش درهمی

for r in range(FIRST, LAST + 1):
    g = f'IF($A{r}="","",{{}})'      # ردیف خالی ⇒ سلول خالی
    n = r - FIRST

    if n < len(SAMPLE):
        vals = SAMPLE[n]
    else:
        vals = ("", None, None, None, None, None, None, None, None)

    for i in range(1, 10):
        c = s.cell(row=r, column=i, value=vals[i-1])
        c.font = BLUE
        c.fill = YELLOW
        c.border = BOX
        if COLS[i-1][2]:
            c.number_format = COLS[i-1][2]

    F = {
        10: f"$B{r}*$C{r}*{R_AED}",                                    # FOB
        11: f"$D{r}*{FREIGHT}*{R_AED}",                                # کرایه
        12: f"$J{r}*{INS}",                                            # بیمه
        13: f"$J{r}+$K{r}+$L{r}",                                      # CIF
        14: f"$M{r}*{DUTY}",                                           # حقوق ورودی
        15: f"$D{r}*{INLAND}",                                         # حمل داخلی
        16: (f'{LUMP}*IF({BASIS}="وزن",'
             f'IF({W_SUM}=0,0,$D{r}/{W_SUM}),'
             f'IF({V_SUM}=0,0,$B{r}*$C{r}/{V_SUM}))'),                 # سهم مقطوع
        17: f"$E{r}*$C{r}",                                            # جانبی
        18: f"($M{r}+$N{r})*{VAT}",                                    # VAT گمرک
        19: (f'$J{r}+$K{r}+$L{r}+$N{r}+$O{r}+$P{r}+$Q{r}'
             f'+IF({VATCOST}="بله",$R{r},0)'),                         # بهای تمام‌شده کل
        20: f"IF($C{r}=0,0,$S{r}/$C{r})",                              # هر واحد
        21: f"$G{r}/(1+{VAT})",                                        # فروش خالص
        22: f"$U{r}*$H{r}",                                            # هزینه فروش
        23: f"$U{r}-$T{r}-$V{r}",                                      # سود هر واحد
        24: f"IF($U{r}=0,0,$W{r}/$U{r})",                              # حاشیه سود
        25: f"$W{r}*$C{r}",                                            # سود کل
        26: f"IF($H{r}>=1,0,$T{r}/(1-$H{r})*(1+{VAT}))",               # سربه‌سر قیمت فروش
        27: f"IF($AE{r}=0,0,($AG{r}-$AF{r})/$AE{r})",                  # نرخ درهم سربه‌سر
        28: f"IF({R_AED}=0,0,($AA{r}-{R_AED})/{R_AED})",               # حاشیه امنیت ارزی
        29: f"IF(OR($S{r}=0,$I{r}=0),0,$Y{r}/$S{r}*365/$I{r})",        # بازده سالانه‌شده
        30: f"IF($F{r}=0,0,($G{r}-$F{r})/$F{r})",                      # فاصله تا بازار
        # ستون‌های کمکی برای حل نرخ سربه‌سر:  بهای تمام‌شده = A×نرخ + B
        31: (f"($B{r}*$C{r}*(1+{INS})+$D{r}*{FREIGHT})*(1+{DUTY})"
             f'*IF({VATCOST}="بله",1+{VAT},1)'),                       # ضریب A
        32: f"$O{r}+$P{r}+$Q{r}",                                      # ثابت B
        33: f"($U{r}-$V{r})*$C{r}",                                    # سود در دسترس
    }

    for i, body in F.items():
        c = s.cell(row=r, column=i, value="=" + g.format(body))
        c.font = BLACK
        c.fill = DERIVFILL
        c.border = BOX
        if COLS[i-1][2]:
            c.number_format = COLS[i-1][2]

# --- ردیف جمع
TOT = LAST + 1
put(s, f"A{TOT}", "جمع", BOLD, None, TOTFILL, True)
for i in range(2, len(COLS) + 1):
    L = get_column_letter(i)
    c = s.cell(row=TOT, column=i)
    c.font = BOLD; c.fill = TOTFILL; c.border = BOX
    if COLS[i-1][2]: c.number_format = COLS[i-1][2]
    # جمع‌زدن قیمت واحد یا میانگین‌گیری ساده از درصدها بی‌معناست ⇒ خالی می‌ماند.
    # «بازده سالانه‌شده» هم وزن‌دهی لازم دارد و در شیت «خلاصه» آمده است.
    if i in (2, 8, 9, 20, 21, 22, 23, 26, 27, 28, 29, 30):
        c.value = None
    elif i == 24:   # حاشیه سود کل = سود کل ÷ فروش خالص کل
        rev = f"SUMPRODUCT($U${FIRST}:$U${LAST},$C${FIRST}:$C${LAST})"
        c.value = f"=IF({rev}=0,0,$Y${TOT}/{rev})"
    else:
        c.value = f"=SUM({L}{FIRST}:{L}{LAST})"

put(s, f"A{TOT+2}",
    "ستون‌های «کمکی» (AE تا AG) برای حل معادله نرخ سربه‌سر لازم‌اند: بهای تمام‌شده = A×نرخ درهم + B. "
    "می‌توانید آن‌ها را پنهان کنید اما حذف نکنید.", SMALL)

# ======================================================================
#  شیت ۳ — خلاصه
# ======================================================================
q = wb.create_sheet("خلاصه")
q.sheet_view.rightToLeft = True
q.column_dimensions["A"].width = 40
q.column_dimensions["B"].width = 22
q.column_dimensions["C"].width = 20
q.column_dimensions["D"].width = 46

K = "'کالاها'!"

put(q, "A1", "خلاصه محموله", TITLE)

put(q, "A3", "سرمایه و حجم", SECT)
put(q, "B3", "ریال", BOLD, None, None, False, "center")
put(q, "C3", "تومان", BOLD, None, None, False, "center")

rows = [
    ("ارزش خرید از دبی (درهم)",        f"=SUMPRODUCT({K}$B${FIRST}:$B${LAST},{K}$C${FIRST}:$C${LAST})", AED, "قیمت خرید × تعداد، جمع همه کالاها"),
    ("وزن کل محموله (کیلوگرم)",        f"={K}$D${TOT}", NUM1, ""),
    ("بهای تمام‌شده کالا (بدون VAT)",  f"={K}$S${TOT}", RIAL, "کالا + کرایه + بیمه + حقوق ورودی + حمل داخلی + ترخیص + جانبی"),
    ("ارزش افزوده گمرک",               f"={K}$R${TOT}", RIAL, "اگر مؤدی ثبت‌نام‌شده باشید با VAT فروش تهاتر می‌شود"),
    ("کل وجه پرداختی (سرمایه لازم)",   f"={K}$S${TOT}+IF({VATCOST}=\"بله\",0,{K}$R${TOT})", RIAL, "نقدینگی‌ای که باید تأمین شود"),
]
r = 4
for label, formula, fmt, note in rows:
    put(q, f"A{r}", label, BLACK, None, None, True)
    c = put(q, f"B{r}", formula, GREEN, fmt, None, True)
    if fmt == RIAL:
        put(q, f"C{r}", f"=$B{r}/10", BLACK, RIAL, None, True)
    put(q, f"D{r}", note, SMALL)
    r += 1

r += 1
put(q, f"A{r}", "فروش و سود", SECT); put(q, f"B{r}", "ریال", BOLD, None, None, False, "center"); put(q, f"C{r}", "تومان", BOLD, None, None, False, "center"); r += 1
for label, formula, fmt, note in [
    ("فروش خالص (بدون VAT)",   f"=SUMPRODUCT({K}$U${FIRST}:$U${LAST},{K}$C${FIRST}:$C${LAST})", RIAL, ""),
    ("سود ناخالص",             f"={K}$Y${TOT}", RIAL, "پس از کسر بهای تمام‌شده و هزینه فروش"),
]:
    put(q, f"A{r}", label, BLACK, None, None, True)
    put(q, f"B{r}", formula, GREEN, fmt, None, True)
    put(q, f"C{r}", f"=$B{r}/10", BLACK, RIAL, None, True)
    put(q, f"D{r}", note, SMALL)
    r += 1

MARGIN_ROW = r
put(q, f"A{r}", "حاشیه سود کل", BLACK, None, None, True)
put(q, f"B{r}", f"=IF($B{r-2}=0,0,$B{r-1}/$B{r-2})", GREEN, PCT, None, True)
put(q, f"D{r}", "سود ÷ فروش خالص", SMALL); r += 1

put(q, f"A{r}", "بازده سرمایه (ROI)", BLACK, None, None, True)
put(q, f"B{r}", f"=IF({K}$S${TOT}=0,0,{K}$Y${TOT}/{K}$S${TOT})", GREEN, PCT, None, True)
put(q, f"D{r}", "سود ÷ بهای تمام‌شده", SMALL); r += 1

DAYS_ROW = r
put(q, f"A{r}", "میانگین روز درگیری سرمایه", BLACK, None, None, True)
put(q, f"B{r}", f"=IF({K}$S${TOT}=0,0,SUMPRODUCT({K}$I${FIRST}:$I${LAST},{K}$S${FIRST}:$S${LAST})/{K}$S${TOT})", GREEN, NUM1, None, True)
put(q, f"D{r}", "میانگین وزنی بر مبنای سرمایه هر کالا", SMALL); r += 1

put(q, f"A{r}", "بازده سالانه‌شده", BLACK, None, None, True)
put(q, f"B{r}", f"=IF($B{DAYS_ROW}=0,0,$B{r-2}*365/$B{DAYS_ROW})", GREEN, PCT, None, True)
put(q, f"D{r}", "بازده × ۳۶۵ ÷ روز. فرض می‌کند سرمایه بلافاصله دوباره به کار می‌افتد.", SMALL); r += 2

put(q, f"A{r}", "ریسک ارزی", SECT); r += 1
BE_ROW = r
put(q, f"A{r}", "نرخ درهم سربه‌سر کل سبد (ریال)", BLACK, None, None, True)
put(q, f"B{r}", f"=IF({K}$AE${TOT}=0,0,({K}$AG${TOT}-{K}$AF${TOT})/{K}$AE${TOT})", GREEN, RIAL, None, True)
put(q, f"D{r}", "بالاتر از این نرخ، کل محموله با قیمت‌های فروش فعلی زیان‌ده می‌شود.", SMALL); r += 1

put(q, f"A{r}", "نرخ درهم امروز (ریال)", BLACK, None, None, True)
put(q, f"B{r}", f"={R_AED}", GREEN, RIAL, None, True); r += 1

put(q, f"A{r}", "حاشیه امنیت ارزی", BLACK, None, None, True)
put(q, f"B{r}", f"=IF({R_AED}=0,0,($B{BE_ROW}-{R_AED})/{R_AED})", GREEN, PCT, None, True)
put(q, f"D{r}", "ریال تا این درصد می‌تواند بی‌ارزش شود و هنوز ضرر نکنید.", SMALL); r += 1

put(q, f"A{r}", "اثر تغییر نرخ از روز خرید تا امروز", BLACK, None, None, True)
put(q, f"B{r}", f"=({R_AED}-{R_BUY})*{K}$AE${TOT}", GREEN, RIAL, None, True)
put(q, f"C{r}", f"=$B{r}/10", BLACK, RIAL, None, True)
put(q, f"D{r}", "اگر همین محموله را امروز دوباره بخرید، این مبلغ گران‌تر تمام می‌شود.", SMALL); r += 1

# بهای تمام‌شده با نرخ امروز حساب می‌شود، پس سود ستون Y از ابتدا «سود جایگزینی» است.
# سود دفتری = همان سود + اختلاف نرخ (اگر ریال ضعیف شده باشد، بزرگ‌تر به نظر می‌رسد).
put(q, f"A{r}", "سود دفتری (بر مبنای نرخ روز خرید)", BLACK, None, None, True)
put(q, f"B{r}", f"={K}$Y${TOT}+({R_AED}-{R_BUY})*{K}$AE${TOT}", GREEN, RIAL, None, True)
put(q, f"C{r}", f"=$B{r}/10", BLACK, RIAL, None, True)
put(q, f"D{r}", "عددی که دفترها نشان می‌دهند — با نرخی که واقعاً پرداخت کرده‌اید.", SMALL); r += 1

put(q, f"A{r}", "سود واقعی (بر مبنای نرخ امروز)", BLACK, None, None, True)
put(q, f"B{r}", f"={K}$Y${TOT}", GREEN, RIAL, None, True)
put(q, f"C{r}", f"=$B{r}/10", BLACK, RIAL, None, True)
put(q, f"D{r}", "سودی که پس از خرید دوباره همان مقدار جنس واقعاً باقی می‌ماند.", SMALL); r += 2

# --- تحلیل حساسیت
put(q, f"A{r}", "حساسیت سود به نرخ درهم", SECT); r += 1
hdr = r
for j, t in enumerate(["تغییر نرخ", "نرخ درهم (ریال)", "بهای تمام‌شده کل (ریال)", "سود کل (ریال)", "حاشیه سود"]):
    c = q.cell(row=hdr, column=1 + j, value=t)
    c.font = HEADFONT; c.fill = HEADFILL; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
q.column_dimensions["E"].width = 16
r += 1
for k, f in enumerate([-0.2, -0.1, 0.0, 0.1, 0.2, 0.3, 0.4, 0.5]):
    rr = r + k
    put(q, f"A{rr}", f, BLUE, PCT, YELLOW, True)
    put(q, f"B{rr}", f"={R_AED}*(1+$A{rr})", BLACK, RIAL, None, True)
    put(q, f"C{rr}", f"={K}$AE${TOT}*$B{rr}+{K}$AF${TOT}", BLACK, RIAL, None, True)
    put(q, f"D{rr}", f"={K}$AG${TOT}-$C{rr}", BLACK, RIAL, None, True)
    put(q, f"E{rr}", f"=IF(SUMPRODUCT({K}$U${FIRST}:$U${LAST},{K}$C${FIRST}:$C${LAST})=0,0,"
                     f"$D{rr}/SUMPRODUCT({K}$U${FIRST}:$U${LAST},{K}$C${FIRST}:$C${LAST}))",
        BLACK, PCT, None, True)
r += 8

put(q, f"A{r+1}",
    "توجه: «بازده سالانه‌شده» فرض می‌کند سرمایه بلافاصله در محموله بعدی به کار می‌افتد. "
    "در عمل فاصله بین محموله‌ها این عدد را کم می‌کند.", SMALL)
put(q, f"A{r+2}",
    "همه اعداد بر پایه نرخ ارز شیت «پارامترها» است. پیش از هر تصمیم، نرخ روز و تعرفه گمرکی را تأیید کنید.", SMALL)

# ======================================================================
#  شیت ۴ — راهنما
# ======================================================================
h = wb.create_sheet("راهنما")
h.sheet_view.rightToLeft = True
h.column_dimensions["A"].width = 30
h.column_dimensions["B"].width = 95

put(h, "A1", "راهنمای استفاده", TITLE)

GUIDE = [
    ("۱ — نرخ ارز را وارد کنید",
     "در شیت «پارامترها» سلول B6 (درهم → ریال) را با نرخ بازار آزاد روز پر کنید. همه محاسبات از همین‌جا می‌آید."),
    ("۲ — پارامترهای محموله",
     "کرایه، بیمه، حقوق ورودی، حمل داخلی و هزینه‌های ترخیص را وارد کنید. این‌ها برای کل محموله مشترک‌اند."),
    ("۳ — کالاها را ثبت کنید",
     "در شیت «کالاها» ستون‌های A تا I (زمینه زرد) را پر کنید. تا ۲۵ کالا آماده است؛ ردیف‌های خالی نادیده گرفته می‌شوند."),
    ("۴ — نتیجه را بخوانید",
     "ستون T بهای تمام‌شده هر واحد، ستون X حاشیه سود و ستون AA نرخ درهمی است که در آن آن کالا سربه‌سر می‌شود."),
    ("۵ — خلاصه",
     "شیت «خلاصه» سرمایه لازم، سود، بازده و ریسک ارزی کل محموله را نشان می‌دهد."),
    ("", ""),
    ("مسیر هزینه", "دبی (خرید به درهم) ← حمل دریایی ← بندرعباس ← گمرک (حقوق ورودی + ارزش افزوده) ← حمل جاده‌ای ← تهران"),
    ("", ""),
    ("بهای تمام‌شده چیست؟",
     "کالا + کرایه دریایی + بیمه + حقوق ورودی گمرک + حمل داخلی + سهم ترخیص و هزینه‌های مقطوع + هزینه‌های جانبی."),
    ("سرشکن‌کردن",
     "هزینه‌های مقطوع (ترخیص، انبارداری، مجوز) بر مبنای وزن یا ارزش بین کالاها تقسیم می‌شود — در سلول B20 انتخاب کنید."),
    ("ارزش افزوده",
     "قیمت فروش را «با ارزش افزوده» وارد کنید؛ فایل خودش آن را کنار می‌گذارد تا سود واقعی به دست آید."),
    ("نرخ درهم سربه‌سر",
     "چون بخشی از هزینه‌ها ارزی و بخشی ریالی است، بهای تمام‌شده به‌صورت خطی به نرخ درهم وابسته است: "
     "بهای تمام‌شده = A×نرخ + B. حل این معادله نرخی را می‌دهد که در آن سود صفر می‌شود."),
    ("ریسک ارزی",
     "سود دفتری با نرخ روز خرید حساب می‌شود، اما محموله بعدی را با نرخ امروز می‌خرید. "
     "شیت «خلاصه» تفاوت این دو را نشان می‌دهد — همان چیزی که در تورم ارزی سرمایه را می‌خورد."),
    ("", ""),
    ("هشدار",
     "همه اعداد پیش‌فرض نمونه‌اند. نرخ ارز را از صرافی و درصد حقوق ورودی را از ترخیص‌کار خود بگیرید."),
]
rr = 3
for a, b in GUIDE:
    if a:
        put(h, f"A{rr}", a, BOLD)
        put(h, f"B{rr}", b, Font(name=FONT, size=10))
        h.cell(row=rr, column=2).alignment = Alignment(wrap_text=True, vertical="top")
        h.row_dimensions[rr].height = 30
    rr += 1

wb.save(OUT)
print("نوشته شد:", OUT)
