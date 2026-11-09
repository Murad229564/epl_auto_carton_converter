import re
import pandas as pd

# ---------------------------------------------------------------------------
# Fortis Garments Limited — Buyer: NA-KD
# 'Purchase Order' (শিট: 'Carton') এক্সেল ফরম্যাট। শুধু প্রথম শিট ('Carton')
# পড়া হয় — 'Shipping mark' ও hidden 'Sheet2/3' ব্যবহার হয় না।
#
# টেবিল (হেডার: Sl no / order no / PO Number / Item Description / ... /
# Measurement (Cm) / ... / Need Carton Quantity):
#   order no            -> Gmt. PO  (এবং টেমপ্লেটের উপরের PO ফিল্ড — _header_po)
#   Item Description    -> কার্টন রো না Top & Bottom রো — সেটা ঠিক করে:
#       'Carton (5 ply)'       -> Item Name = UI সিলেকশন, Ply = UI-এর manual_ply
#                                 (UI-তে না দিলে বন্ধনীর 'N ply' থেকে, তাও না
#                                 থাকলে 'N/A')
#       'Top & Bottom (3 ply)' -> Item Name = 'Top Bottom', Ply = 3 (অটো)
#   Measurement (Cm)    -> Length / Width / Height ('40x30x40 cm';
#                          '28 x38' হলে হাইট ফাঁকা)
#   Need Carton Quantity-> Order Qty (এক্সেলে '#,### pcs' ফরম্যাটে যেভাবে
#                          রাউন্ড হয়ে দেখায়, সেভাবেই রাউন্ড: 3.9 -> 4)
# অন্য কলাম (PO Number='All', Carton Colors, Size='All', Garments Order Qty
# (hidden), Con., Unit Price, Remarks) ব্যবহার হয় না — Style/Reference/Pack
# Type তাই ফাঁকা থাকবে। 'Total' রো পেলে টেবিল শেষ।
# ---------------------------------------------------------------------------

NA = 'N/A'

# True করলে উপরের PO ফিল্ডে ফাইলের নামের PO নম্বরও যোগ হবে
# (যেমন 'Jersey Headband (1100-011200) / 324937') — কারণ একাধিক ফাইলে
# 'order no' একই লেখা থাকে (যেমন PO-324937 / 324938 সবগুলোতে একই),
# ফলে আলাদা ফাইল চেনা কঠিন হয়। ডিফল্ট False = শুধু 'order no'-এর লেখা।
HEADER_PO_APPEND_FILE_PO = False


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    if isinstance(v, float) and pd.isna(v):
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def _num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    try:
        return float(str(v).replace(',', '').strip())
    except (TypeError, ValueError):
        return None


def _round_half_up(x):
    # Excel-এর মতো (2.5 -> 3); Python-এর round() banker's rounding করে
    return int(x + 0.5)


def _parse_measurement(m):
    """'40x30x40 cm' -> (40.0, 30.0, 40.0); '28 x38' -> (28.0, 38.0, '')"""
    nums = re.findall(r'\d+(?:\.\d+)?', str(m or ''))
    vals = [float(n) for n in nums[:3]]
    while len(vals) < 3:
        vals.append('')
    return vals[0], vals[1], vals[2]


def _load_rows(file_stream, filename):
    try:
        from outhouse_extractor import _read_excel_rows
        return _read_excel_rows(file_stream, filename)
    except ImportError:
        file_stream.seek(0)
        return pd.read_excel(file_stream, header=None).values.tolist()


def _find_header(rows, max_scan=30):
    for i, row in enumerate(rows[:max_scan]):
        normed = [_norm(c) for c in row]
        if ('orderno' in normed and 'itemdescription' in normed
                and 'needcartonquantity' in normed):
            heads = {}
            for ci, n in enumerate(normed):
                if n and n not in heads:
                    heads[n] = ci
            return i, heads
    return None, None


def _file_po(filename):
    m = re.search(r'PO[-_ ]?(\d{4,})', filename or '', re.I)
    return m.group(1) if m else ''


def read_fortis_nakd_style_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """মূল entry point। এই ফরম্যাট না হলে ValueError ছোঁড়ে।"""
    rows = _load_rows(file_stream, filename)

    header_idx, heads = _find_header(rows)
    if header_idx is None:
        raise ValueError("'order no / Item Description / Need Carton Quantity' হেডিং-সহ Fortis/NA-KD টেবিল পাওয়া যায়নি")

    c_po = heads['orderno']
    c_item = heads['itemdescription']
    c_meas = heads.get('measurementcm')
    c_qty = heads['needcartonquantity']

    def get(row, ci):
        if ci is None or ci >= len(row):
            return ''
        return row[ci]

    default_item_name = item_name_override or 'Master Carton'
    ui_ply = manual_ply.strip() if manual_ply else ''

    items = []
    last_po = ''
    for row in rows[header_idx + 1:]:
        if row is None:
            continue
        # 'Total' রো (যেকোনো কলামে) পেলে টেবিল শেষ
        if any(_norm(c) == 'total' for c in row):
            break

        p = _clean(get(row, c_po))
        if p:
            last_po = p
        desc = _clean(get(row, c_item))
        qty = _num(get(row, c_qty))
        if not desc or qty is None or qty <= 0:
            continue

        d = desc.lower()
        is_top_bottom = ('top' in d and 'bottom' in d)
        if is_top_bottom:
            item_name, ply = 'Top Bottom', '3'
        else:
            item_name = default_item_name
            m = re.search(r'(\d+)\s*ply', desc, re.I)
            ply = ui_ply or (m.group(1) if m else NA)

        length, width, height = _parse_measurement(get(row, c_meas))
        items.append({
            'item_name': item_name,
            'ewo_no': NA,
            'style_no': '',
            'po_no': last_po,
            'length': length,
            'width': width,
            'height': height,
            'ply': ply,
            'qty': _round_half_up(qty),
            'pack_type': '',
            'reference': '',
            'color': '',
            'size': '',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_source_file': filename,
        })

    if not items:
        raise ValueError("হেডার পাওয়া গেছে কিন্তু কোয়ান্টিটি-সহ কোনো ডাটা রো পাওয়া যায়নি")

    # উপরের PO ফিল্ড: ফাইলের 'order no'-এর লেখা (একাধিক ইউনিক হলে '/' দিয়ে জোড়া)
    uniq = []
    for it in items:
        if it['po_no'] and it['po_no'] not in uniq:
            uniq.append(it['po_no'])
    header_po = '/'.join(uniq)
    if HEADER_PO_APPEND_FILE_PO:
        fpo = _file_po(filename)
        if fpo:
            header_po = f"{header_po} / {fpo}" if header_po else fpo
    for it in items:
        it['_header_po'] = header_po

    return items
