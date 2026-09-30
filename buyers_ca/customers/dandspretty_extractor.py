"""
D&S Pretty Fashions Ltd. — Buyer: C&A
'Please arrange Carton as per below breakdown' .xls ফরম্যাট (পুরনো
বাইনারি, xlrd দিয়ে পড়া হয়)। দুই রকম সাব-ভ্যারিয়েন্ট দেখা গেছে, দুটোই
এই একই ফাংশন হ্যান্ডেল করে:
  - কিছু ফাইলে উপরে একটা 'STYLE:' লেবেল-সহ একটাই স্টাইল পুরো ফাইলের জন্য।
  - কিছু ফাইলে কোনো গ্লোবাল লেবেল নেই, বরং টেবিলের নিজস্ব 'Style' কলামে
    প্রতিটা ব্লকের প্রথম রো-তে স্টাইল নম্বর থাকে (একাধিক স্টাইল-ব্লক
    একই ফাইলে থাকতে পারে), পরের রো-গুলোতে ফাঁকা।
  উভয় ক্ষেত্রেই সমাধান একটাই: টেবিলের নিজস্ব 'Style' কলাম থেকে
  forward-fill করা (মান পেলে আপডেট, না পেলে আগের মানই বহাল) — এতে
  দুই ভ্যারিয়েন্টই সঠিকভাবে কাজ করে, আলাদা লজিক লাগে না।

  মূল টেবিলের হেডার: Style | Item Name | Carton Measurement | Actual
  Carton Qty | With X % Total Carton | Remarks | Amount | Remarks।
    - Item Name কলাম ফাঁকা/C&A-প্যাটার্নের বাইরে -> এক্সেপশনাল
      (ca_rules.py নিজেই ধরে)।
    - Qty: প্রতিটা রো-তে আলাদাভাবে চেক হয় — 'With X % Total Carton'
      কলামে সেই রো-তে ভ্যালু থাকলে সেটাই নেওয়া হয়, ফাঁকা থাকলে (কিছু
      ফাইলে পুরো কলামই ফাঁকা থাকে) সেই রো-র 'Actual Carton Qty' ব্যবহার
      হয় (ইউজার-কনফার্মড, প্রতি-রো ভিত্তিতে — পুরো ফাইলের জন্য একবার না)।
  'TOTAL Pcs'/'Total' লেখা রো-এর যেকোনো কলামে পাওয়া গেলেই থেমে যায়
  (কোন কলামে এই লেখা থাকবে তা ফাইল-ভেদে বদলাতে পারে)।
"""
import re

from ..ca_rules import apply_ca_rules


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def _clean_style(v):
    """Style কলামের ভ্যালু কখনো সংখ্যা (xlrd ফ্লোট হিসেবে পড়ে, যেমন
    2267658.0), কখনো টেক্সট (যেমন 'USIM :84794-2265276') — দুটোই
    পরিষ্কার স্ট্রিং-এ রূপান্তর করে।"""
    if v in (None, ''):
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return _clean(v)


_MEAS_RE = re.compile(r'(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)')


def _parse_measurement_cm(text):
    if not text:
        return None, None, None
    m = _MEAS_RE.search(str(text))
    if not m:
        return None, None, None
    return float(m.group(1)), float(m.group(2)), float(m.group(3))


def _is_total_row(row):
    return any(_norm(c) in ('totalpcs', 'total') for c in row if c is not None)


def _get_rows(file_stream, filename):
    """.xls/.xlsx দুটোই সাপোর্ট করে (এই কাস্টমারের ফাইল সাধারণত .xls,
    কিন্তু ভবিষ্যতে .xlsx আসলেও যেন কাজ করে) — একই outhouse_extractor.py-এর
    মাল্টি-ইঞ্জিন হেল্পার পুনরায় ব্যবহার করা হচ্ছে।"""
    from outhouse_extractor import _read_excel_rows
    return _read_excel_rows(file_stream, filename)


def read_dandspretty_excel(file_stream, filename='', lookup=None, item_name_override='', manual_ply=''):
    """মূল entry point। রিটার্ন করে (line_items, warnings)।"""
    rows = _get_rows(file_stream, filename)

    header_row_idx = None
    col_map = {}
    for i, row in enumerate(rows[:20]):
        labels = {}
        for c, v in enumerate(row):
            lbl = _norm(v)
            if lbl and lbl not in labels:
                labels[lbl] = c
        if 'itemname' in labels and 'cartonmeasurement' in labels:
            header_row_idx = i
            col_map = labels
            break
    if header_row_idx is None:
        return [], []

    style_col = col_map.get('style')
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
    if meas_col is None or (pct_qty_col is None and actual_qty_col is None):
        return [], [f"⚠️ '{filename}': প্রত্যাশিত কলাম (Item Name/Carton Measurement/Qty) পাওয়া যায়নি।"]

    line_items = []
    all_warnings = []
    current_style = ''
    for row in rows[header_row_idx + 1:]:
        if not row:
            continue
        if _is_total_row(row):
            break

        if style_col is not None and style_col < len(row):
            style_val = _clean_style(row[style_col])
            if style_val:
                current_style = style_val

        meas_val = row[meas_col] if meas_col < len(row) else None
        l_cm, w_cm, h_cm = _parse_measurement_cm(meas_val)
        if l_cm is None:
            continue

        # প্রতিটা রো-তে আলাদাভাবে: আগে % কলাম চেক, ফাঁকা হলে Actual Qty
        qty_val = None
        if pct_qty_col is not None and pct_qty_col < len(row):
            v = row[pct_qty_col]
            if v not in (None, ''):
                qty_val = v
        if qty_val is None and actual_qty_col is not None and actual_qty_col < len(row):
            v = row[actual_qty_col]
            if v not in (None, ''):
                qty_val = v
        if qty_val is None:
            continue

        raw_item = {
            'raw_style': current_style,
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