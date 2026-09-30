"""
D&S Pretty Fashions Ltd. — Buyer: C&A
'Please arrange Carton as per below breakdown' .xls ফরম্যাট (পুরনো
বাইনারি, xlrd দিয়ে পড়া হয়)।
  - উপরের ইনফো-ব্লকে 'STYLE:' লেবেলের পাশে একটা সিঙ্গেল স্টাইল টেক্সট
    (যেমন 'USIM :84794-2265276') — /TRI যোগ হয় ca_rules.py-তে।
  - মূল টেবিলের হেডার: Style | Item Name | PACK | PCS/SIZE | Carton
    Measurement | Actual Carton Qty | With X % Total Carton | Remarks |
    Amount | Remarks।
      - Item Name কলাম ফাঁকা/C&A-প্যাটার্নের বাইরে -> এক্সেপশনাল
        (ca_rules.py নিজেই ধরে)।
      - Qty: 'With X % Total Carton' কলাম থাকলে (ওয়েস্টেজ-সহ) সেটাই
        নেওয়া হয়, না থাকলে 'Actual Carton Qty' (ইউজার-কনফার্মড)।
  - 'TOTAL Pcs' রো-তে থেমে যায়।
"""
import re

from ..ca_rules import apply_ca_rules


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


_MEAS_RE = re.compile(r'(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)')


def _parse_measurement_cm(text):
    if not text:
        return None, None, None
    m = _MEAS_RE.search(str(text))
    if not m:
        return None, None, None
    return float(m.group(1)), float(m.group(2)), float(m.group(3))


def _get_rows(file_stream, filename):
    """.xls/.xlsx দুটোই সাপোর্ট করে (এই কাস্টমারের ফাইল সাধারণত .xls,
    কিন্তু ভবিষ্যতে .xlsx আসলেও যেন কাজ করে) — একই outhouse_extractor.py-এর
    মাল্টি-ইঞ্জিন হেল্পার পুনরায় ব্যবহার করা হচ্ছে।"""
    from outhouse_extractor import _read_excel_rows
    return _read_excel_rows(file_stream, filename)


def read_dandspretty_excel(file_stream, filename='', lookup=None, item_name_override='', manual_ply=''):
    """মূল entry point। রিটার্ন করে (line_items, warnings)।"""
    rows = _get_rows(file_stream, filename)

    style_text = ''
    for row in rows[:15]:
        if not row:
            continue
        first = _clean(row[0])
        if _norm(first) == 'style' and len(row) > 1 and _clean(row[1]):
            style_text = _clean(row[1])
            break

    header_row_idx = None
    col_map = {}
    for i, row in enumerate(rows[:20]):
        labels = {}
        for c, v in enumerate(row):
            lbl = _norm(v)
            if lbl:
                labels[lbl] = c
        if 'itemname' in labels and 'cartonmeasurement' in labels:
            header_row_idx = i
            col_map = labels
            break
    if header_row_idx is None:
        return [], []

    code_col = col_map.get('itemname')
    meas_col = col_map.get('cartonmeasurement')
    actual_qty_col = col_map.get('actualcartonqty')
    # 'With 2 % Total Carton' -> normalize করলে 'with2totalcarton' — 'with'
    # দিয়ে শুরু হওয়া যেকোনো কলামকেই ওয়েস্টেজ-সহ কোয়ান্টিটি কলাম ধরা হয়
    pct_qty_col = None
    for lbl, c in col_map.items():
        if lbl.startswith('with') and 'total' in lbl:
            pct_qty_col = c
            break
    qty_col = pct_qty_col if pct_qty_col is not None else actual_qty_col
    if meas_col is None or qty_col is None:
        return [], [f"⚠️ '{filename}': প্রত্যাশিত কলাম (Item Name/Carton Measurement/Qty) পাওয়া যায়নি।"]

    line_items = []
    all_warnings = []
    for row in rows[header_row_idx + 1:]:
        if not row:
            continue
        first = _clean(row[0]) if row[0] is not None else ''
        if _norm(first) in ('totalpcs', 'total'):
            break
        meas_val = row[meas_col] if meas_col < len(row) else None
        qty_val = row[qty_col] if qty_col < len(row) else None
        l_cm, w_cm, h_cm = _parse_measurement_cm(meas_val)
        if l_cm is None or qty_val in (None, ''):
            continue

        raw_item = {
            'raw_style': style_text,
            'raw_code_text': _clean(row[code_col]) if code_col is not None and code_col < len(row) else '',
            'length_cm': l_cm, 'width_cm': w_cm, 'height_cm': h_cm,
            'qty': qty_val,
            '_source_file': filename,
        }
        item, warns = apply_ca_rules(raw_item, lookup, item_name_override=item_name_override, manual_ply=manual_ply)
        if item:
            line_items.append(item)
        all_warnings.extend(warns)

    return line_items, all_warnings
