import re

# ---------------------------------------------------------------------------
# Majumder Garments Ltd. — Buyer: Original Marines
# 'CARTON ORDER FOR BUYER ORIGINAL MARINES' .xls ফরম্যাট। দুই রকম সাব-
# ভ্যারিয়েন্ট (দুটোই এই একই ফাংশন হ্যান্ডেল করে, হেডারে 'COLOR' কলাম
# আছে কিনা দেখে অটো-ডিটেক্ট হয়):
#
#   Solid (color+size ব্রেকডাউন): SL | STYLE | COLOR | MEASUREMENT |
#     SIZE RATIO | CONTENTS | QUALITY | WEIGHT | SHIPPING MARK | QTY |
#     UNIT | DELIVERY DATE। STYLE/COLOR/MEASUREMENT শুধু প্রতিটা ব্লকের
#     প্রথম রো-তে থাকে (forward-fill দরকার); MEASUREMENT ব্লকের মাঝেও
#     বদলাতে পারে।
#       COLOR      -> Reference/SKU Number
#       SIZE RATIO -> Pack Type
#
#   Assorted (সিঙ্গেল-রো, কোনো COLOR/SIZE কলাম নেই): SL | STYLE |
#     MEASUREMENT | SIZE RATIO | ... | QTY | ...। Reference/Pack Type
#     প্রযোজ্য না (ইউজার-কনফার্মড), ফাঁকা থাকে।
#
# উভয় ফরম্যাটেই:
#   STYLE        -> Gmt. Style No
#   MEASUREMENT  -> Length/Width/Height
#   QTY          -> Order Qty
#   PO           -> কোনো PO-নম্বর কলাম নেই এই ফরম্যাটে, তাই 'N/A'
#
# Top & Bottom: শিটের কোথাও 'Note:** Must be Top & Bottom' টেক্সট থাকলে —
#   প্রতিটা ইউনিক Master Carton measurement-গ্রুপের জন্য একটা করে Top
#   Bottom লাইন যোগ হয়: Length/Width মূল measurement থেকে ৫cm কম (Height
#   লাগে না), Qty = সেই গ্রুপের মোট Qty-এর দ্বিগুণ, Ply ফিক্সড ৩।
#   নোটটা না পেলে কোনো Top Bottom যোগ হয় না, বরং একটা ⚠️ ওয়ার্নিং রিটার্ন
#   হয় (ইউজার-কনফার্মড) — এই জন্যই এটা batch-style ফাংশন হিসেবে বানানো
#   (REGISTRY না, BATCH_REGISTRY-তে বসবে), কারণ শুধু batch-কলিং-কনভেনশনেই
#   (line_items, warnings) দুটো একসাথে রিটার্ন করা যায়।
# ---------------------------------------------------------------------------


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def _is_num(v):
    if v is None:
        return False
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


def _fmt_num(v):
    if v is None or v == '':
        return ''
    try:
        f = float(v)
    except (TypeError, ValueError):
        return ''
    return str(int(f)) if f == int(f) else str(round(f, 3))


_MEAS_RE = re.compile(r'L\s*(\d+\.?\d*)\s*CM.*?W\s*(\d+\.?\d*)\s*CM.*?H\s*(\d+\.?\d*)\s*CM', re.I | re.S)


def _parse_measurement(text):
    """'L 36 CM X W 27 CM X H 16 CM' -> ('36','27','16')।"""
    if not text:
        return '', '', ''
    m = _MEAS_RE.search(str(text))
    if not m:
        return '', '', ''
    return _fmt_num(m.group(1)), _fmt_num(m.group(2)), _fmt_num(m.group(3))


def _get_rows(file_stream, filename):
    """pandas-ভিত্তিক _read_excel_rows ব্যবহার করে (.xls/.xlsx দুটোই
    সাপোর্ট করতে) — কিন্তু pandas ফাঁকা সেলের জন্য NaN রিটার্ন করে
    (None না), যেটা str()-এ করলে ভুল করে লিটারেল 'nan' টেক্সট হয়ে যায়
    (D&S Pretty-তে একবার এই বাগ ধরা পড়েছিল) — তাই এখানেই সব NaN-কে
    None-এ বদলে দেওয়া হচ্ছে।"""
    from outhouse_extractor import _read_excel_rows
    rows = _read_excel_rows(file_stream, filename)

    def _clean_nan(v):
        if isinstance(v, float) and v != v:
            return None
        return v

    return [[_clean_nan(v) for v in row] for row in rows]


def _find_header_row(rows, max_scan=20):
    """কলাম A-তে 'SL' আর কলাম B-তে 'STYLE' — এই দুটো একসাথে থাকা রো-কে
    হেডার ধরা হয়।"""
    for i, row in enumerate(rows[:max_scan]):
        if not row:
            continue
        a = _norm(row[0]) if len(row) > 0 else ''
        b = _norm(row[1]) if len(row) > 1 else ''
        if a == 'sl' and b == 'style':
            return i
    return None


def _build_col_map(header_row):
    col_map = {}
    for c, v in enumerate(header_row):
        label = _norm(v)
        if not label:
            continue
        if label == 'style':
            col_map['style_no'] = c
        elif label == 'color':
            col_map['reference'] = c
        elif label == 'measurement':
            col_map['measurement'] = c
        elif 'sizeratio' in label:
            col_map['pack_type'] = c
        elif label == 'qty':
            col_map['qty'] = c
    return col_map


def _has_top_bottom_note(rows):
    for row in rows:
        for v in row:
            if v is not None and 'mustbetopbottom' in _norm(v):
                return True
    return False


def _read_one_file(file_stream, filename, item_name_override, manual_ply):
    """একটা ফাইল থেকে ডাটা বের করে। রিটার্ন করে (line_items, warnings)।
    এই ফরম্যাট না হলে ([], []) রিটার্ন করে।"""
    rows = _get_rows(file_stream, filename)

    header_idx = _find_header_row(rows)
    if header_idx is None:
        return [], []

    col_map = _build_col_map(rows[header_idx])
    required = ('style_no', 'measurement', 'qty')
    if not all(k in col_map for k in required):
        return [], []

    style_col = col_map['style_no']
    ref_col = col_map.get('reference')
    meas_col = col_map['measurement']
    pack_col = col_map.get('pack_type')
    qty_col = col_map['qty']

    default_item_name = item_name_override or 'Master Carton'
    ply_value = manual_ply.strip() if manual_ply else 'N/A'

    current_style = ''
    current_ref = ''
    current_meas_text = ''
    master_items = []
    group_order = []
    group_sums = {}

    for row in rows[header_idx + 1:]:
        if not row:
            continue
        first = _clean(row[0]) if len(row) > 0 else ''
        if _norm(first) == 'total':
            break

        if style_col < len(row) and row[style_col] not in (None, ''):
            current_style = _clean(row[style_col])
        if ref_col is not None and ref_col < len(row) and row[ref_col] not in (None, ''):
            current_ref = _clean(row[ref_col])
        if meas_col < len(row) and row[meas_col] not in (None, ''):
            current_meas_text = str(row[meas_col])

        qty_val = row[qty_col] if qty_col < len(row) else None
        if not _is_num(qty_val) or float(qty_val) <= 0:
            continue

        l, w, h = _parse_measurement(current_meas_text)
        if not l:
            continue

        pack_val = _clean(row[pack_col]) if pack_col is not None and pack_col < len(row) else ''

        master_items.append({
            'item_name': default_item_name,
            'ewo_no': 'N/A',
            'style_no': current_style or 'N/A',
            'po_no': 'N/A',
            'length': l, 'width': w, 'height': h,
            'ply': ply_value,
            'qty': round(float(qty_val)),
            'pack_type': pack_val or 'N/A',
            'reference': current_ref or 'N/A',
            'color': '', 'size': '',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '', 'delivery_address_pdf': '',
            '_source_file': filename,
        })

        key = (l, w, h)
        if key not in group_sums:
            group_sums[key] = 0
            group_order.append(key)
        group_sums[key] += round(float(qty_val))

    if not master_items:
        return [], []

    all_items = list(master_items)
    warnings = []

    if _has_top_bottom_note(rows):
        for key in group_order:
            l, w, h = key
            try:
                tb_l = _fmt_num(float(l) - 5)
                tb_w = _fmt_num(float(w) - 5)
            except (TypeError, ValueError):
                continue
            all_items.append({
                'item_name': 'Top Bottom',
                'ewo_no': 'N/A',
                'style_no': 'N/A',
                'po_no': 'N/A',
                'length': tb_l, 'width': tb_w, 'height': '',
                'ply': '3',
                'qty': group_sums[key] * 2,
                'pack_type': 'N/A',
                'reference': 'N/A',
                'color': '', 'size': '',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': '', 'delivery_address_pdf': '',
                '_source_file': filename,
            })
    else:
        warnings.append(
            f"⚠️ '{filename}': এই ফাইলে 'Note:** Must be Top & Bottom' টেক্সট পাওয়া যায়নি — "
            f"তাই Top Bottom লাইন যোগ করা হয়নি। ম্যানুয়ালি চেক করুন এটা দরকার কিনা।"
        )

    return all_items, warnings


def combine_majumder_booking_files(files, item_name_override='', manual_ply=''):
    """BATCH_REGISTRY-এর uniform কল-সিগনেচার। প্রতিটা ফাইল আলাদাভাবে
    প্রসেস হয় (Top Bottom গ্রুপিং প্রতিটা ফাইলের নিজস্ব measurement-
    গ্রুপের ভেতরেই সীমাবদ্ধ, ফাইল-ভেদে মেশানো হয় না — প্রতিটা ফাইল
    সাধারণত একটা আলাদা স্টাইলের বুকিং)। রিটার্ন করে (line_items, warnings)।
    """
    all_items = []
    all_warnings = []
    for file_stream, filename in files:
        items, warns = _read_one_file(file_stream, filename, item_name_override, manual_ply)
        if not items and not warns:
            all_warnings.append(f"⚠️ '{filename}': পরিচিত ফরম্যাট মেলেনি, স্কিপ করা হয়েছে।")
            continue
        all_items.extend(items)
        all_warnings.extend(warns)
    return all_items, all_warnings
