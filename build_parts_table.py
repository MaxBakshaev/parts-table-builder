#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_parts_table.py  (v4)

Собирает Excel-таблицу деталей из Details.txt + PDF-чертежей.

Что нового в v4:
  - многострочные поля (Дополнительная информация / Additional Information)
    читаются целиком, включая списки и абзацы;
  - поддержаны Polymer Modification и Additional Information;
  - колонки, пустые у всех деталей, автоматически скрываются
    (у полимерных деталей нет термообработки, у стальных — модификации полимера).

Запуск:
    python3 build_parts_table.py --details Details.txt --pdf-dir . --out out.xlsx
Ключи:
    --currency CNY         валюта по умолчанию (USD или CNY)
    --keep-all-columns     не скрывать пустые колонки
"""

import argparse, os, re, subprocess, sys, tempfile, unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.formatting.rule import Rule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PImage, ImageChops

# ---------- колонки ----------
# Порядок параметров в таблице: (заголовок, ключ в [EN]-блоке)
PARAM_COLUMNS = [
    ("Part",                      "Part"),
    ("Quantity",                  "Quantity"),
    ("Material Group",            "Material Group"),
    ("Material Grade",            "Material Grade"),
    ("Steel Type",                "Steel Type"),
    ("Polymer Modification",      "Polymer Modification"),
    ("Heat Treatment",            "Heat Treatment"),
    ("Hardness (HRC)",            "Hardness (HRC)"),
    ("Chemical-thermal Coating",  "Chemical-thermal Coating"),
    ("Electrochemical Treatment", "Electrochemical Treatment"),
    ("Precision Grade",           "Precision Grade"),
    ("Additional Information",    "Additional Information"),
]
# Колонки, которые всегда остаются, даже если пустые
ALWAYS_KEEP = {"Part", "Quantity", "Material Group", "Material Grade"}

WIDTHS = {
    "No.": 4, "Drawing": 22, "Part": 26, "Quantity": 8,
    "Material Group": 14, "Material Grade": 22, "Steel Type": 14,
    "Polymer Modification": 15, "Heat Treatment": 14, "Hardness (HRC)": 13,
    "Chemical-thermal Coating": 16, "Electrochemical Treatment": 16,
    "Precision Grade": 16, "Additional Information": 42,
    "Price": 13, "Total": 14,
}

CURRENCIES = {
    "USD": '"$"#,##0.00',
    "CNY": '"\u00a5"#,##0.00',
}
DEFAULT_CURRENCY = "USD"

SETTINGS_ROW = 1
HEADER_ROW = 2
FIRST_DATA_ROW = 3


# ---------- разбор Details.txt ----------
def parse_blocks(txt, tag, name_kw):
    """
    Разбирает блоки [RU] или [EN] с поддержкой МНОГОСТРОЧНЫХ значений.
    Строка поля  : '   - Ключ: значение'  (дефис + ПРОБЕЛ)
    Продолжение  : любая другая строка    (например '-Deburr sharp edges')
    Возвращает {номер: {поле: значение}}.
    """
    if tag == "RU":
        chunks = re.findall(r"\[RU\](.*?)\[EN\]", txt, flags=re.S)
    else:
        chunks = re.findall(r"\[EN\](.*?)(?=\[RU\]|\Z)", txt, flags=re.S)

    out = {}
    for ch in chunks:
        num, d, cur_key = None, {}, None
        for raw in ch.strip("\n").splitlines():
            s = raw.strip()
            if not s:
                if cur_key and d.get(cur_key):
                    d[cur_key] += "\n"          # сохраняем разрыв абзаца
                continue
            head = re.match(r"(\d+)\.\s*" + re.escape(name_kw) + r":\s*(.+)", s)
            if head and num is None:
                num = int(head.group(1))
                d[name_kw] = head.group(2).strip()
                cur_key = None
                continue
            field = re.match(r"-\s+(.+?):\s*(.*)$", s)   # дефис + пробел = новое поле
            if field:
                cur_key = field.group(1).strip()
                d[cur_key] = field.group(2).strip()
            elif cur_key:                                 # продолжение поля
                d[cur_key] = (d[cur_key] + "\n" + s).strip("\n")
        if num is not None:
            out[num] = {k: v.strip() for k, v in d.items()}
    return out


def norm(s):
    s = unicodedata.normalize("NFKC", s).lower()
    return re.sub(r"[^a-z\u0430-\u044f0-9]+", "", s)


def index_pdfs(pdf_dir):
    by_num, by_name = {}, {}
    for fn in sorted(os.listdir(pdf_dir)):
        if not fn.lower().endswith(".pdf"):
            continue
        path = os.path.join(pdf_dir, fn)
        stem = fn[:-4]
        m = re.match(r"\s*(\d+)\s*[.\-_)]\s*(.+)", stem)
        if m:
            by_num[int(m.group(1))] = path
            by_name[norm(m.group(2))] = path
        else:
            by_name[norm(stem)] = path
    return by_num, by_name


def render_page1(pdf_path, out_png, dpi=150):
    prefix = out_png[:-4] if out_png.lower().endswith(".png") else out_png
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-f", "1", "-l", "1",
                    "-singlefile", pdf_path, prefix],
                   check=True, capture_output=True)
    produced = prefix + ".png"
    try:
        im = PImage.open(produced).convert("RGB")
        bg = PImage.new("RGB", im.size, (255, 255, 255))
        bbox = ImageChops.difference(im, bg).getbbox()
        if bbox:
            im.crop(bbox).save(produced)
    except Exception:
        pass
    return produced


# ---------- сборка ----------
def build(details_txt, pdf_dir, out_xlsx, box=(150, 200), dpi=150,
          currency=DEFAULT_CURRENCY, keep_all_columns=False):
    txt = open(details_txt, encoding="utf-8").read()
    ru = parse_blocks(txt, "RU", "\u0414\u0435\u0442\u0430\u043b\u044c")
    en = parse_blocks(txt, "EN", "Part")
    if not en:
        raise SystemExit("В Details.txt не найдено ни одной детали ([EN] блоки).")
    by_num, by_name = index_pdfs(pdf_dir)
    nums = sorted(en.keys())

    # какие параметры реально заполнены хотя бы у одной детали
    used = []
    for title, key in PARAM_COLUMNS:
        has = any(str(en[n].get(key, "")).strip() not in ("", "-") for n in nums)
        if keep_all_columns or has or title in ALWAYS_KEEP:
            used.append((title, key))
    dropped = [t for t, k in PARAM_COLUMNS if (t, k) not in used]

    headers = ["No.", "Drawing"] + [t for t, _ in used] + ["Price", "Total"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Parts"

    A = "Arial"
    hdr_font = Font(name=A, bold=True, color="FFFFFF", size=10)
    cell_font = Font(name=A, size=10)
    thin = Side(style="thin", color="B7C3D6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="305496")
    money_hdr = PatternFill("solid", fgColor="C55A11")
    price_fill = PatternFill("solid", fgColor="FFF2CC")
    total_fill = PatternFill("solid", fgColor="E2EFDA")
    warn_fill = PatternFill("solid", fgColor="F4B183")
    alt_fill = PatternFill("solid", fgColor="EAF0FA")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    topleft = Alignment(horizontal="left", vertical="top", wrap_text=True)

    price_col = headers.index("Price") + 1
    total_col = headers.index("Total") + 1
    qty_col = headers.index("Quantity") + 1
    PC, TC, QC = (get_column_letter(price_col), get_column_letter(total_col),
                  get_column_letter(qty_col))
    cur_cell = f"${PC}${SETTINGS_ROW}"

    # ---- выбор валюты ----
    lbl = ws.cell(row=SETTINGS_ROW, column=price_col - 1, value="Currency:")
    lbl.font = Font(name=A, bold=True, size=10)
    lbl.alignment = Alignment(horizontal="right", vertical="center")
    cur = ws.cell(row=SETTINGS_ROW, column=price_col, value=currency)
    cur.font = Font(name=A, bold=True, size=11, color="833C00")
    cur.alignment = center
    cur.fill = price_fill
    med = Side(style="medium", color="C55A11")
    cur.border = Border(left=med, right=med, top=med, bottom=med)
    dv = DataValidation(type="list", formula1='"' + ",".join(CURRENCIES) + '"',
                        allow_blank=False, showDropDown=False)
    dv.promptTitle, dv.prompt = "Currency", "Выберите валюту из списка"
    dv.errorTitle, dv.error = "Недопустимое значение", "Выберите валюту из списка"
    ws.add_data_validation(dv)
    dv.add(cur)
    ws.row_dimensions[SETTINGS_ROW].height = 20

    # ---- заголовки ----
    # ВАЖНО: заголовки Price/Total — обычный текст, а не формула.
    # Формула без пересчёта извлекается как пустая ячейка, и колонка «теряется».
    # Валюта берётся из ячейки со списком, а символ подставляется форматированием.
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=HEADER_ROW, column=c, value=h)
        cell.font = hdr_font
        cell.fill = money_hdr if h in ("Price", "Total") else hdr_fill
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[HEADER_ROW].height = 32

    # ---- строки ----
    BOX_W, BOX_H = box
    tmp = tempfile.mkdtemp(prefix="thumbs_")
    unmatched = []

    for idx, num in enumerate(nums):
        r = FIRST_DATA_ROW + idx
        p = en[num]
        ws.row_dimensions[r].height = BOX_H * 0.75

        ws.cell(row=r, column=1, value=num)
        for c, h in enumerate(headers, start=1):
            if h in ("No.", "Drawing"):
                continue
            if h == "Price":
                val = ""
            elif h == "Total":
                val = (f'=IF({PC}{r}="","",'
                       f'IF(ISNUMBER({PC}{r}),{QC}{r}*{PC}{r},"CHECK PRICE"))')
            else:
                key = dict(used)[h]
                val = p.get(key, "")
                if h == "Quantity":
                    try: val = int(val)
                    except Exception: pass
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = cell_font
            cell.border = border
            if h == "Additional Information":
                cell.alignment = topleft          # длинный текст — от верха
            elif c >= price_col or h in ("Quantity",):
                cell.alignment = center
            elif c >= 5:
                cell.alignment = left
            else:
                cell.alignment = center
        for c in (1, 2):
            ws.cell(row=r, column=c).font = cell_font
            ws.cell(row=r, column=c).alignment = center
            ws.cell(row=r, column=c).border = border
        if idx % 2 == 1:
            for c in range(1, len(headers) + 1):
                if c not in (2, price_col, total_col):
                    ws.cell(row=r, column=c).fill = alt_fill
        ws.cell(row=r, column=price_col).fill = price_fill
        ws.cell(row=r, column=total_col).fill = total_fill

        # чертёж: по номеру, затем по русскому имени
        pdf = by_num.get(num)
        if not pdf and num in ru:
            pdf = by_name.get(norm(ru[num].get("\u0414\u0435\u0442\u0430\u043b\u044c", "")))
        if not pdf:
            unmatched.append(num)
            ws.cell(row=r, column=2, value="PDF not found").fill = warn_fill
            continue
        try:
            png = render_page1(pdf, os.path.join(tmp, f"p{num}.png"), dpi=dpi)
            iw, ih = PImage.open(png).size
            sc = min(BOX_W / iw, BOX_H / ih)
            img = XLImage(png)
            img.width, img.height = int(iw * sc), int(ih * sc)
            img.anchor = f"B{r}"
            ws.add_image(img)
        except Exception:
            unmatched.append(num)
            ws.cell(row=r, column=2, value="render error").fill = warn_fill

    # ---- запрет вводить в Price текст ----
    dv_price = DataValidation(type="decimal", operator="greaterThanOrEqual",
                              formula1="0", allow_blank=True)
    dv_price.promptTitle = "Price"
    dv_price.prompt = ("Введите ТОЛЬКО число, например 25 или 25.50.\n"
                       "Не пишите $25/1PCS \u2014 валюта задаётся списком в шапке.")
    dv_price.errorTitle = "Только число"
    dv_price.error = ("Введите цену за единицу числом, без символа валюты и без /1PCS.\n"
                      "Валюта выбирается в выпадающем списке над колонкой.")
    ws.add_data_validation(dv_price)
    dv_price.add(f"{PC}{FIRST_DATA_ROW}:{PC}{FIRST_DATA_ROW + len(nums) - 1}")

    # ---- итоговая строка ----
    last = FIRST_DATA_ROW + len(nums) - 1
    tr = last + 1
    ws.row_dimensions[tr].height = 24
    tl = ws.cell(row=tr, column=price_col - 1, value="TOTAL")
    tl.font = Font(name=A, bold=True, size=11)
    tl.alignment = Alignment(horizontal="right", vertical="center")
    tl.border = border
    qt = ws.cell(row=tr, column=qty_col, value=f"=SUM({QC}{FIRST_DATA_ROW}:{QC}{last})")
    qt.font = Font(name=A, bold=True, size=10)
    qt.alignment = center
    qt.border = border
    gt = ws.cell(row=tr, column=total_col,
                 value=f"=SUM({TC}{FIRST_DATA_ROW}:{TC}{last})")
    gt.font = Font(name=A, bold=True, size=11)
    gt.alignment = center
    gt.border = border
    gt.fill = PatternFill("solid", fgColor="C6E0B4")

    # ---- предупреждение, если цену вписали текстом ----
    wr = tr + 1
    warn = ws.cell(row=wr, column=3,
                   value=(f'=IF(SUMPRODUCT(--ISTEXT({PC}{FIRST_DATA_ROW}:{PC}{last}))>0,'
                          f'"\u26a0 CHECK: "&SUMPRODUCT(--ISTEXT({PC}{FIRST_DATA_ROW}:{PC}{last}))'
                          f'&" price cell(s) contain text instead of a number '
                          f'(write 25, not $25/1PCS)","")'))
    warn.font = Font(name=A, bold=True, size=10, color="C00000")
    warn.alignment = Alignment(horizontal="left", vertical="center")

    # ---- символ валюты ----
    rng = f"{PC}{FIRST_DATA_ROW}:{TC}{tr}"
    for i, (code, fmt) in enumerate(CURRENCIES.items()):
        dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=1000 + i, formatCode=fmt))
        rule = Rule(type="expression", dxf=dxf, stopIfTrue=False)
        rule.formula = [f'{cur_cell}="{code}"']
        ws.conditional_formatting.add(rng, rule)

    for c, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(c)].width = WIDTHS.get(h, 15)
    ws.freeze_panes = f"C{FIRST_DATA_ROW}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{get_column_letter(len(headers))}{HEADER_ROW}"

    note = ws.cell(row=tr + 3, column=1,
                   value="Fill in the yellow Price cells \u2014 Total and the grand total "
                         f"are calculated automatically. Currency: dropdown in {PC}{SETTINGS_ROW}.")
    note.font = Font(name=A, italic=True, size=8, color="808080")

    wb.save(out_xlsx)
    return {"parts": len(nums), "unmatched": unmatched, "out": out_xlsx,
            "currency": currency, "columns": headers, "dropped_columns": dropped}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--details", default=os.getenv("DETAILS_TXT", "Details.txt"))
    ap.add_argument("--pdf-dir", default=os.getenv("PDF_DIR", "."))
    ap.add_argument("--out", default=os.getenv("OUTPUT_XLSX", "Parts_table_EN.xlsx"))
    ap.add_argument("--dpi", type=int, default=int(os.getenv("DPI", "150")))
    ap.add_argument("--currency", default=os.getenv("CURRENCY", DEFAULT_CURRENCY),
                    choices=list(CURRENCIES))
    ap.add_argument("--keep-all-columns", action="store_true")
    a = ap.parse_args()
    res = build(a.details, a.pdf_dir, a.out, dpi=a.dpi, currency=a.currency,
                keep_all_columns=a.keep_all_columns)
    print(res)
    if res["unmatched"]:
        print("WARNING: без чертежа остались детали:", res["unmatched"], file=sys.stderr)
