"""
Soorty Textiles (BD) Limited — Buyer: C&A (Trillora সিস্টেম)
'MASTER CARTON WORK ORDER' Excel ফরম্যাট।
  - একটা ফাইলে এক বা একাধিক শিট থাকতে পারে (প্রতিটা শিট একটা স্টাইল/
    অর্ডার) — সব visible শিট থেকে ডাটা নিয়ে একসাথে কম্বাইন করা হয়,
    হাইড শিট বাদ।
  - শিটের উপরে 'Style' লেবেলের পাশের পুরো টেক্সট (যেমন '2270102 SS27 EPP
    Slim Coloured - Vetiver') -> Style (/TRI যোগ হয় ca_rules.py-তে)।
  - টেবিল: SL NO | Item Name | CTN. MEAS. | ITEM/PLY | TOTAL CARTON | Remarks
      Item Name    -> কোড: 'C&A-28', 'C&A-10 Perf.' ইত্যাদি লিস্টের কোড হলে
                      Gmt. PO-তে কোড + Remarks-এ প্রাইস; 'C&A-MOC EB'-জাতীয় EB
                      বা ফাঁকা/অচেনা হলে EXCEPTIONAL MEASUREMENT (সব C&A
                      নিয়ম ca_rules.py/ca_calc.py-তে, এখানে না)
      CTN. MEAS.   -> Length/Width/Height। এই কলামের মাপ mm-এ ('430x330x190'),
                      টেমপ্লেটে cm-এ কনভার্ট হয়ে বসে।
      ITEM/PLY     -> Ply ('3 Ply' -> 3) — ফাইলে থাকলে এটাই (UI-র Ply না);
                      ফাঁকা থাকলে UI-র Ply।
      TOTAL CARTON -> Order Qty (রাউন্ড করে)। 'Remarks' কলামের সংখ্যা
                      (পিস-সংখ্যা) কখনো Qty হিসেবে ধরা হয় না।
  - Item Name (Master Carton ইত্যাদি) UI থেকে।
  - উপরের 'PO' সংখ্যা ব্যবহার হয় না (C&A নিয়মে Gmt. PO-তে কোড/EXCEPTIONAL
    বসে; হেডারের PO শুধু UI থেকে)। উপরের 'Date :' (ইস্যু-তারিখ) ব্যবহার হয় না।
  - ডেলিভারি ডেট: 'Delivery Date' লেবেলের পাশের তারিখ প্রতিটা লাইনে
    '_file_delivery_date'-এ বসে (USE_FILE_DELIVERY_DATE = False করলে বন্ধ)।
  - 'Grand Total' রো-তে থেমে যায়।
"""
import re
from openpyxl import load_workbook

from ..ca_rules import apply_ca_rules

USE_FILE_DELIVERY_DATE = False


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def _is_num(v):
    if v is None or isinstance(v, bool):
        return False
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


_MEAS_RE = re.compile(r'(\d+\.?\d*)\s*[xX×*]\s*(\d+\.?\d*)\s*[xX×*]\s*(\d+\.?\d*)')


def _parse_measurement(text):
    if not text:
        return None
    m = _MEAS_RE.search(str(text))
    if not m:
        return None
    return float(m.group(1)), float(m.group(2)), float(m.group(3))


def _parse_ply(v):
    m = re.search(r'(\d+)', str(v or ''))
    return m.group(1) if m else ''


def _date_format():
    try:
        from config import DELIVERY_DATE_FORMAT
        return DELIVERY_DATE_FORMAT
    except Exception:
        return '%d-%b-%Y'


_DATE_FORMATS = ('%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y', '%d-%b-%Y', '%d %b %Y', '%Y-%m-%d')


def _to_date_str(v):
    import datetime
    fmt = _date_format()
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime(fmt)
    if isinstance(v, str):
        t = _clean(v)
        for f in _DATE_FORMATS:
            try:
                return datetime.datetime.strptime(t, f).strftime(fmt)
            except ValueError:
                continue
    return None


def _right_value(ws, r, c, max_cols=12):
    for c2 in range(c + 1, min(ws.max_column, c + max_cols) + 1):
        v = ws.cell(row=r, column=c2).value
        if v is not None and _clean(v) != '':
            return v
    return None


def _find_header(ws, max_scan=40):
    """'SL NO' আর 'Item Name' একই রো-তে থাকলে সেটাই হেডার। রিটার্ন করে
    (row, {label: col})।"""
    for r in range(1, min(ws.max_row, max_scan) + 1):
        labels = {}
        for c in range(1, ws.max_column + 1):
            lbl = _norm(ws.cell(row=r, column=c).value)
            if lbl and lbl not in labels:
                labels[lbl] = c
        if 'slno' in labels and 'itemname' in labels:
            return r, labels
    return None, None


def _find_top_info(ws, header_row):
    """হেডারের উপরের অংশ থেকে Style টেক্সট আর ডেলিভারি ডেট বের করে।"""
    style = ''
    delivery = None
    for r in range(1, header_row):
        for c in range(1, min(ws.max_column, 6) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None or not isinstance(v, str):
                continue
            n = _norm(v)
            if n == 'style' and not style:
                style = _clean(_right_value(ws, r, c))
            elif n == 'deliverydate' and delivery is None:
                delivery = _to_date_str(_right_value(ws, r, c))
    return style, delivery


def read_soorty_textiles_excel(file_stream, filename='', lookup=None, item_name_override='', manual_ply=''):
    """মূল entry point। রিটার্ন করে (line_items, warnings)।"""
    wb = load_workbook(file_stream, data_only=True)
    items, warnings = [], []

    for sn in wb.sheetnames:
        ws = wb[sn]
        if ws.sheet_state != 'visible':
            continue
        header_row, labels = _find_header(ws)
        if header_row is None:
            continue

        item_col = labels['itemname']
        meas_col = next((c for l, c in labels.items() if l.startswith('ctnmeas') or l == 'measurement'), None)
        ply_col = next((c for l, c in labels.items() if 'ply' in l), None)
        qty_col = next((c for l, c in labels.items() if l.startswith('totalcarton')), None)
        if meas_col is None or qty_col is None:
            warnings.append(f"⚠️ '{filename}' শিট '{sn}': CTN. MEAS./TOTAL CARTON কলাম পাওয়া যায়নি — শিট স্কিপ।")
            continue

        style_text, file_date = _find_top_info(ws, header_row)
        if not style_text:
            warnings.append(f"⚠️ '{filename}' শিট '{sn}': Style পাওয়া যায়নি।")

        hidden_rows = {r for r, d in ws.row_dimensions.items() if d.hidden}
        blank_run = 0
        for r in range(header_row + 1, min(ws.max_row, 5000) + 1):
            if r in hidden_rows:
                continue
            first = _norm(ws.cell(row=r, column=1).value)
            if 'grandtotal' in first:
                break

            qty_val = ws.cell(row=r, column=qty_col).value
            meas = _parse_measurement(ws.cell(row=r, column=meas_col).value)
            code_text = _clean(ws.cell(row=r, column=item_col).value)
            if not code_text and meas is None and not _is_num(qty_val):
                blank_run += 1
                if blank_run >= 100:
                    break
                continue
            blank_run = 0
            if meas is None or not _is_num(qty_val) or float(qty_val) <= 0:
                if _is_num(qty_val) and float(qty_val) > 0:
                    warnings.append(
                        f"⚠️ '{filename}' শিট '{sn}' রো {r}: measurement পাওয়া যায়নি — রো স্কিপ।")
                continue

            l, w, h = meas
            file_ply = _parse_ply(ws.cell(row=r, column=ply_col).value) if ply_col else ''
            raw_item = {
                'raw_style': style_text,
                'raw_code_text': code_text,
                'length_cm': l, 'width_cm': w, 'height_cm': h,
                'qty': qty_val,
                'ply': file_ply,
                # এই কলামের মাপ mm-এ (সংখ্যা বড় হলে) — অনুমানের ওয়ার্নিং ছাড়াই
                'unit_hint': 'mm' if max(l, w, h) > 150 else None,
                '_sheet': sn, '_source_file': filename,
            }
            item, warns = apply_ca_rules(
                raw_item, lookup, item_name_override=item_name_override,
                # ফাইলে Ply থাকলে সেটাই; না থাকলে UI-র Ply
                manual_ply='' if file_ply else manual_ply)
            warnings.extend(warns)
            if item:
                if USE_FILE_DELIVERY_DATE and file_date:
                    item['_file_delivery_date'] = file_date
                items.append(item)

    return items, warnings
