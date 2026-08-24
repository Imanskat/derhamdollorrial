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

put(p, "A23", "ضریب ارزش گمرکی (٪ بالاتر از فاکتور)")
put(p, "B23", 0.0, BLUE, PCT, YELLOW, True)
put(p, "C23", "اگر گمرک ارزش را بالاتر از فاکتور شما می‌بندد. صفر = فاکتور مبناست.", SMALL)

put(p, "A24", "حاشیه سود هدف (٪)")
put(p, "B24", 0.35, BLUE, PCT, YELLOW, True)
put(p, "C24", "مبنای ستون «قیمت پیشنهادی» در شیت کالاها.", SMALL)

dv_basis = DataValidation(type="list", formula1='"وزن,ارزش"', allow_blank=False)
p.add_data_validation(dv_basis); dv_basis.add(p["B20"])
dv_yn = DataValidation(type="list", formula1='"بله,خیر"', allow_blank=False)
p.add_data_validation(dv_yn); dv_yn.add(p["B22"])

put(p, "A26", "راهنمای رنگ‌ها", SECT)
put(p, "A27", "متن آبی روی زمینه زرد = سلول ورودی؛ فقط این‌ها را تغییر دهید.", BLUE)
put(p, "A28", "متن سیاه = فرمول؛ دست نزنید.", BLACK)
put(p, "A29", "متن سبز = ارجاع به شیت دیگر.", GREEN)

put(p, "A31", "هشدار درباره نرخ ارز", SECT)
for i, line in enumerate([
    "سرویس‌های عمومی ارز (Google، exchangerate و مانند آن) نرخ «رسمی» ریال را می‌دهند، نه نرخ «بازار آزاد».",
    "اگر بهای تمام‌شده را با نرخ رسمی حساب کنید، عدد چند برابر کمتر از واقعیت درمی‌آید و کالا را زیر قیمت می‌فروشید.",
    "همیشه نرخ آزاد روز را از صرافی یا بازار بگیرید و در سلول B6 وارد کنید.",
]):
    put(p, f"A{32+i}", line, SMALL)

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

UPLIFT        = f"{P}$B$23"
TARGET        = f"{P}$B$24"

COLS = [
    # (عنوان، عرض، قالب، نوع)   نوع: in=ورودی، fx=فرمول
    ("نام محصول",                        26, None, "in"),
    ("قیمت خرید هر واحد\n(درهم)",        13, AED,  "in"),
    ("تعداد",                             9, NUM,  "in"),
    ("وزن کل\n(کیلوگرم)",                11, NUM1, "in"),
    ("ارزش گمرکی هر واحد\n(درهم)",       15, AED,  "in"),
    ("تعرفه این کالا\n(٪)",              12, PCT,  "in"),
    ("ضایعات و مرجوعی\n(٪)",             13, PCT,  "in"),
    ("هزینه جانبی هر واحد\n(ریال)",      15, RIAL, "in"),
    ("قیمت روز بازار تهران\n(ریال)",     17, RIAL, "in"),
    ("قیمت فروش من\n(ریال، با VAT)",     17, RIAL, "in"),
    ("هزینه فروش\n(٪)",                  11, PCT,  "in"),
    ("روز خرید\nتا فروش",                10, NUM,  "in"),
    ("تعداد قابل فروش",                   13, NUM1, "fx"),
    ("تعرفه اعمال‌شده\n(٪)",             12, PCT,  "fx"),
    ("ارزش کالا FOB\n(ریال)",            16, RIAL, "fx"),
    ("کرایه دریایی\n(ریال)",             14, RIAL, "fx"),
    ("بیمه\n(ریال)",                     12, RIAL, "fx"),
    ("ارزش CIF فاکتور\n(ریال)",          16, RIAL, "fx"),
    ("پایه ارزش گمرکی\n(ریال)",          16, RIAL, "fx"),
    ("حقوق ورودی گمرک\n(ریال)",          16, RIAL, "fx"),
    ("حمل داخلی\n(ریال)",                14, RIAL, "fx"),
    ("سهم از هزینه‌های مقطوع\n(ریال)",   17, RIAL, "fx"),
    ("هزینه جانبی کل\n(ریال)",           14, RIAL, "fx"),
    ("ارزش افزوده گمرک\n(ریال)",         16, RIAL, "fx"),
    ("بهای تمام‌شده کل\n(ریال)",         17, RIAL, "fx"),
    ("بهای تمام‌شده هر واحد\n(ریال)",    17, RIAL, "fx"),
    ("فروش خالص هر واحد\n(ریال)",        16, RIAL, "fx"),
    ("هزینه فروش هر واحد\n(ریال)",       15, RIAL, "fx"),
    ("درآمد خالص ردیف\n(ریال)",          16, RIAL, "fx"),
    ("سود در دسترس\n(ریال)",             16, RIAL, "fx"),
    ("سود کل ردیف\n(ریال)",              16, RIAL, "fx"),
    ("سود هر واحد\n(ریال)",              14, RIAL, "fx"),
    ("حاشیه سود\n(٪)",                   11, PCT,  "fx"),
    ("نقطه سربه‌سر\nقیمت فروش (ریال)",   16, RIAL, "fx"),
    ("قیمت پیشنهادی\nحاشیه هدف (ریال)",  16, RIAL, "fx"),
    ("نرخ درهم سربه‌سر\n(ریال)",         15, RIAL, "fx"),
    ("حاشیه امنیت ارزی\n(٪)",            14, PCT,  "fx"),
    ("بازده سالانه‌شده\n(٪)",            13, PCT,  "fx"),
    ("فاصله تا بازار تهران\n(٪)",        15, PCT,  "fx"),
    ("ارزش CIF درهمی\n(کمکی)",           14, AED,  "fx"),
    ("پایه گمرکی درهمی\n(کمکی)",         14, AED,  "fx"),
    ("ضریب ارزی A\n(کمکی)",              14, NUM,  "fx"),
    ("ثابت ریالی B\n(کمکی)",             16, RIAL, "fx"),
]

put(s, "A1", "بهای تمام‌شده و سود هر کالا", TITLE)
put(s, "A2", "فقط ستون‌های آبی/زرد (A تا L) را پر کنید؛ بقیه خودکار محاسبه می‌شوند.", SMALL)

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
s.row_dimensions[HEAD_ROW].height = 46
s.freeze_panes = "B5"

#          نام، درهم، تعداد، وزن، ارزش گمرکی، تعرفه، ضایعات، جانبی، بازار، فروش، هزینه‌فروش، روز
SAMPLE = [
    ("کرم پودر (۳۰ میلی‌لیتر)",       32,  600, 42, None, None, 0.03, 400_000, 40_600_000, 37_600_000, 0.12,  75),
    ("رژ لب مایع",                    14, 1200, 30, None, None, 0.04, 250_000, 17_700_000, 16_400_000, 0.12,  60),
    ("پالت سایه چشم",                 48,  300, 39, None, None, 0.06, 500_000, 61_500_000, 56_900_000, 0.15, 110),
    ("سرم ویتامین ث (۳۰ میلی‌لیتر)",  55,  400, 32, None, None, 0.05, 450_000, 68_700_000, 63_600_000, 0.10,  90),
    ("ریمل حجم‌دهنده",                18,  900, 27, None, None, 0.04, 300_000, 22_700_000, 21_000_000, 0.12,  65),
]

W_SUM = f"SUM($D${FIRST}:$D${LAST})"
V_SUM = f"SUMPRODUCT($B${FIRST}:$B${LAST},$C${FIRST}:$C${LAST})"

N_IN = 12                      # تعداد ستون‌های ورودی

for r in range(FIRST, LAST + 1):
    g = 'IF($A{}="","",{{}})'.format(r)
    n = r - FIRST
    vals = SAMPLE[n] if n < len(SAMPLE) else ("",) + (None,)*11

    for i in range(1, N_IN + 1):
        c = s.cell(row=r, column=i, value=vals[i-1])
        c.font = BLUE; c.fill = YELLOW; c.border = BOX
        if COLS[i-1][2]: c.number_format = COLS[i-1][2]

    F = {
        13: f"$C{r}*(1-$G{r})",                                        # تعداد قابل فروش
        14: f'IF($F{r}="",{DUTY},$F{r})',                              # تعرفه اعمال‌شده
        15: f"$B{r}*$C{r}*{R_AED}",                                    # FOB
        16: f"$D{r}*{FREIGHT}*{R_AED}",                                # کرایه
        17: f"$O{r}*{INS}",                                            # بیمه
        18: f"$O{r}+$P{r}+$Q{r}",                                      # CIF فاکتور
        # گمرک لزوماً فاکتور را قبول نمی‌کند: یا ارزش اعلامی هر واحد، یا فاکتور با ضریب
        19: f"IF($E{r}>0,$E{r}*$C{r}*{R_AED},$R{r}*(1+{UPLIFT}))",     # پایه ارزش گمرکی
        20: f"$S{r}*$N{r}",                                            # حقوق ورودی
        21: f"$D{r}*{INLAND}",                                         # حمل داخلی
        22: (f'{LUMP}*IF({BASIS}="وزن",'
             f'IF({W_SUM}=0,0,$D{r}/{W_SUM}),'
             f'IF({V_SUM}=0,0,$B{r}*$C{r}/{V_SUM}))'),                 # سهم مقطوع
        23: f"$H{r}*$C{r}",                                            # جانبی
        24: f"($S{r}+$T{r})*{VAT}",                                    # VAT گمرک (روی ارزش گمرکی)
        25: (f'$R{r}+$T{r}+$U{r}+$V{r}+$W{r}'
             f'+IF({VATCOST}="بله",$X{r},0)'),                         # بهای تمام‌شده کل
        26: f"IF($C{r}=0,0,$Y{r}/$C{r})",                              # هر واحد
        27: f"$J{r}/(1+{VAT})",                                        # فروش خالص
        28: f"$AA{r}*$K{r}",                                           # هزینه فروش هر واحد
        29: f"$AA{r}*$M{r}",                                           # درآمد خالص (فقط اقلام سالم)
        30: f"($AA{r}-$AB{r})*$M{r}",                                  # سود در دسترس
        31: f"$AD{r}-$Y{r}",                                           # سود کل ردیف
        32: f"IF($C{r}=0,0,$AE{r}/$C{r})",                             # سود هر واحد خریداری‌شده
        33: f"IF($AC{r}=0,0,$AE{r}/$AC{r})",                           # حاشیه سود
        34: (f"IF(OR($K{r}>=1,$M{r}=0),0,"
             f"$Y{r}/((1-$K{r})*$M{r})*(1+{VAT}))"),                   # سربه‌سر قیمت فروش
        35: (f"IF(OR((1-$K{r}-{TARGET})<=0,$M{r}=0),0,"
             f"$Y{r}/($M{r}*(1-$K{r}-{TARGET}))*(1+{VAT}))"),          # قیمت حاشیه هدف
        36: f"IF($AP{r}=0,0,($AD{r}-$AQ{r})/$AP{r})",                  # نرخ درهم سربه‌سر
        37: f"IF({R_AED}=0,0,($AJ{r}-{R_AED})/{R_AED})",               # حاشیه امنیت ارزی
        38: f"IF(OR($Y{r}=0,$L{r}=0),0,$AE{r}/$Y{r}*365/$L{r})",       # بازده سالانه‌شده
        39: f"IF($I{r}=0,0,($J{r}-$I{r})/$I{r})",                      # فاصله تا بازار
        # ستون‌های کمکی: بهای تمام‌شده = A × نرخ درهم + B
        40: f"$B{r}*$C{r}*(1+{INS})+$D{r}*{FREIGHT}",                  # CIF درهمی
        41: f"IF($E{r}>0,$E{r}*$C{r},$AN{r}*(1+{UPLIFT}))",            # پایه گمرکی درهمی
        42: (f"$AN{r}+$AO{r}*$N{r}"
             f'+IF({VATCOST}="بله",$AO{r}*(1+$N{r})*{VAT},0)'),        # ضریب A
        43: f"$U{r}+$V{r}+$W{r}",                                      # ثابت B
    }

    for i, body in F.items():
        c = s.cell(row=r, column=i, value="=" + g.format(body))
        c.font = BLACK; c.fill = DERIVFILL; c.border = BOX
        if COLS[i-1][2]: c.number_format = COLS[i-1][2]

# --- ردیف جمع
TOT = LAST + 1
put(s, f"A{TOT}", "جمع", BOLD, None, TOTFILL, True)
SUMMABLE = {3,4,13,15,16,17,18,19,20,21,22,23,24,25,29,30,31,40,41,42,43}
for i in range(2, len(COLS) + 1):
    L = get_column_letter(i)
    c = s.cell(row=TOT, column=i)
    c.font = BOLD; c.fill = TOTFILL; c.border = BOX
    if COLS[i-1][2]: c.number_format = COLS[i-1][2]
    if i in SUMMABLE:
        c.value = f"=SUM({L}{FIRST}:{L}{LAST})"
    elif i == 33:                       # حاشیه سود کل = سود کل ÷ درآمد خالص کل
        c.value = f"=IF($AC${TOT}=0,0,$AE${TOT}/$AC${TOT})"
    else:
        c.value = None                  # جمع یا میانگین این ستون‌ها معنا ندارد

put(s, f"A{TOT+2}",
    "ستون‌های «کمکی» (AN تا AQ) برای حل معادله نرخ سربه‌سر لازم‌اند: بهای تمام‌شده = A×نرخ درهم + B. "
    "می‌توانید آن‌ها را پنهان کنید اما حذف نکنید.", SMALL)
put(s, f"A{TOT+3}",
    "ستون «تعرفه این کالا» را خالی بگذارید تا از درصد کلی محموله استفاده شود؛ صفر یعنی واقعاً تعرفه صفر.", SMALL)
put(s, f"A{TOT+4}",
    "ستون «ارزش گمرکی هر واحد» را خالی بگذارید تا ارزش فاکتور (با ضریب سلول B23) مبنا شود.", SMALL)

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
    ("تعداد خریداری‌شده",              f"={K}$C${TOT}", NUM, ""),
    ("تعداد قابل فروش",                f"={K}$M${TOT}", NUM1, "پس از کسر ضایعات و مرجوعی"),
    ("پایه ارزش گمرکی",                f"={K}$S${TOT}", RIAL, "عددی که گمرک روی آن حق ورودی می‌بندد"),
    ("بهای تمام‌شده کالا (بدون VAT)",  f"={K}$Y${TOT}", RIAL, "کالا + کرایه + بیمه + حقوق ورودی + حمل داخلی + ترخیص + جانبی"),
    ("ارزش افزوده گمرک",               f"={K}$X${TOT}", RIAL, "اگر مؤدی ثبت‌نام‌شده باشید با VAT فروش تهاتر می‌شود"),
    ("کل وجه پرداختی (سرمایه لازم)",   f"={K}$Y${TOT}+IF({VATCOST}=\"بله\",0,{K}$X${TOT})", RIAL, "نقدینگی‌ای که باید تأمین شود"),
]
r = 4
for label, formula, fmt, note in rows:
    put(q, f"A{r}", label, BLACK, None, None, True)
    put(q, f"B{r}", formula, GREEN, fmt, None, True)
    if fmt == RIAL:
        put(q, f"C{r}", f"=$B{r}/10", BLACK, RIAL, None, True)
    put(q, f"D{r}", note, SMALL)
    r += 1

r += 1
put(q, f"A{r}", "فروش و سود", SECT)
put(q, f"B{r}", "ریال", BOLD, None, None, False, "center")
put(q, f"C{r}", "تومان", BOLD, None, None, False, "center"); r += 1
REV_ROW = r
for label, formula, note in [
    ("فروش خالص (بدون VAT)",   f"={K}$AC${TOT}", "فقط اقلام سالم، بدون ارزش افزوده"),
    ("سود ناخالص",             f"={K}$AE${TOT}", "پس از کسر بهای تمام‌شده و هزینه فروش"),
]:
    put(q, f"A{r}", label, BLACK, None, None, True)
    put(q, f"B{r}", formula, GREEN, RIAL, None, True)
    put(q, f"C{r}", f"=$B{r}/10", BLACK, RIAL, None, True)
    put(q, f"D{r}", note, SMALL)
    r += 1

put(q, f"A{r}", "حاشیه سود کل", BLACK, None, None, True)
put(q, f"B{r}", f"=IF($B{REV_ROW}=0,0,$B{REV_ROW+1}/$B{REV_ROW})", GREEN, PCT, None, True)
put(q, f"D{r}", "سود ÷ فروش خالص", SMALL); r += 1

ROI_ROW = r
put(q, f"A{r}", "بازده سرمایه (ROI)", BLACK, None, None, True)
put(q, f"B{r}", f"=IF({K}$Y${TOT}=0,0,{K}$AE${TOT}/{K}$Y${TOT})", GREEN, PCT, None, True)
put(q, f"D{r}", "سود ÷ بهای تمام‌شده", SMALL); r += 1

DAYS_ROW = r
put(q, f"A{r}", "میانگین روز درگیری سرمایه", BLACK, None, None, True)
put(q, f"B{r}", f"=IF({K}$Y${TOT}=0,0,SUMPRODUCT({K}$L${FIRST}:$L${LAST},{K}$Y${FIRST}:$Y${LAST})/{K}$Y${TOT})", GREEN, NUM1, None, True)
put(q, f"D{r}", "میانگین وزنی بر مبنای سرمایه هر کالا", SMALL); r += 1

put(q, f"A{r}", "بازده سالانه‌شده", BLACK, None, None, True)
put(q, f"B{r}", f"=IF($B{DAYS_ROW}=0,0,$B{ROI_ROW}*365/$B{DAYS_ROW})", GREEN, PCT, None, True)
put(q, f"D{r}", "بازده × ۳۶۵ ÷ روز. فرض می‌کند سرمایه بلافاصله دوباره به کار می‌افتد.", SMALL); r += 2

put(q, f"A{r}", "ریسک ارزی", SECT); r += 1
BE_ROW = r
put(q, f"A{r}", "نرخ درهم سربه‌سر کل سبد (ریال)", BLACK, None, None, True)
put(q, f"B{r}", f"=IF({K}$AP${TOT}=0,0,({K}$AD${TOT}-{K}$AQ${TOT})/{K}$AP${TOT})", GREEN, RIAL, None, True)
put(q, f"D{r}", "بالاتر از این نرخ، کل محموله با قیمت‌های فروش فعلی زیان‌ده می‌شود.", SMALL); r += 1

put(q, f"A{r}", "نرخ درهم امروز (ریال)", BLACK, None, None, True)
put(q, f"B{r}", f"={R_AED}", GREEN, RIAL, None, True); r += 1

put(q, f"A{r}", "حاشیه امنیت ارزی", BLACK, None, None, True)
put(q, f"B{r}", f"=IF({R_AED}=0,0,($B{BE_ROW}-{R_AED})/{R_AED})", GREEN, PCT, None, True)
put(q, f"D{r}", "ریال تا این درصد می‌تواند بی‌ارزش شود و هنوز ضرر نکنید.", SMALL); r += 1

put(q, f"A{r}", "اثر تغییر نرخ از روز خرید تا امروز", BLACK, None, None, True)
put(q, f"B{r}", f"=({R_AED}-{R_BUY})*{K}$AP${TOT}", GREEN, RIAL, None, True)
put(q, f"C{r}", f"=$B{r}/10", BLACK, RIAL, None, True)
put(q, f"D{r}", "اگر همین محموله را امروز دوباره بخرید، این مبلغ گران‌تر تمام می‌شود.", SMALL); r += 1

put(q, f"A{r}", "سود دفتری (بر مبنای نرخ روز خرید)", BLACK, None, None, True)
put(q, f"B{r}", f"={K}$AE${TOT}+({R_AED}-{R_BUY})*{K}$AP${TOT}", GREEN, RIAL, None, True)
put(q, f"C{r}", f"=$B{r}/10", BLACK, RIAL, None, True)
put(q, f"D{r}", "عددی که دفترها نشان می‌دهند — با نرخی که واقعاً پرداخت کرده‌اید.", SMALL); r += 1

put(q, f"A{r}", "سود واقعی (بر مبنای نرخ امروز)", BLACK, None, None, True)
put(q, f"B{r}", f"={K}$AE${TOT}", GREEN, RIAL, None, True)
put(q, f"C{r}", f"=$B{r}/10", BLACK, RIAL, None, True)
put(q, f"D{r}", "سودی که پس از خرید دوباره همان مقدار جنس واقعاً باقی می‌ماند.", SMALL); r += 2

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
    put(q, f"C{rr}", f"={K}$AP${TOT}*$B{rr}+{K}$AQ${TOT}", BLACK, RIAL, None, True)
    put(q, f"D{rr}", f"={K}$AD${TOT}-$C{rr}", BLACK, RIAL, None, True)
    put(q, f"E{rr}", f"=IF({K}$AC${TOT}=0,0,$D{rr}/{K}$AC${TOT})", BLACK, PCT, None, True)
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
    ("ارزش گمرکی",
     "گمرک لزوماً فاکتور شما را مبنا قرار نمی‌دهد و برای بسیاری از کالاها ارزش‌گذاری خودش را دارد. "
     "اگر عدد دقیق هر کالا را می‌دانید در ستون E بنویسید؛ اگر فقط تخمین کلی دارید، ستون E را خالی "
     "بگذارید و ضریب را در سلول B23 شیت پارامترها وارد کنید. حق ورودی و ارزش افزوده هر دو روی این پایه بسته می‌شوند."),
    ("تعرفه هر کالا",
     "درصد حقوق ورودی به ردیف تعرفه خودِ کالا بستگی دارد نه به محموله. ستون F را خالی بگذارید تا از "
     "درصد کلی (سلول B15) استفاده شود. عدد صفر یعنی واقعاً تعرفه صفر — خالی و صفر یکی نیستند."),
    ("ضایعات و مرجوعی",
     "پول همه واحدها را داده‌اید ولی همه‌شان به قیمت کامل فروخته نمی‌شوند: شکستگی، نشتی، انقضای نزدیک، "
     "مرجوعی و تخفیف آخر فصل. ستون G تعداد قابل فروش را کم می‌کند ولی بهای تمام‌شده را دست نمی‌زند. "
     "برای آرایشی معمولاً ۳ تا ۸ درصد منطقی است."),
    ("قیمت پیشنهادی",
     "حاشیه سود هدف را در سلول B24 شیت پارامترها بگذارید؛ ستون AI قیمت قفسه لازم برای رسیدن به آن حاشیه "
     "را می‌دهد، با احتساب ضایعات، هزینه فروش و ارزش افزوده. اگر «حاشیه هدف + هزینه فروش» به ۱۰۰٪ برسد "
     "این ستون صفر می‌شود، یعنی آن حاشیه با این ساختار هزینه شدنی نیست."),
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
