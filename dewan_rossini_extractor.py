import re
import pandas as pd

# ---------------------------------------------------------------------------
# Dewan Fashion Wears Ltd — Buyer: Rossini
# 'PERCHASE ORDER (CARTON)' এক্সেল (.xls/.xlsx) ফরম্যাট।
#   - প্রতিটা ফাইল = একটা Style (একটা শিট)। ফাইলের উপরে:
#       'ORDER NUMBER : 1/2026/01182-HH196'  -> টেমপ্লেটের উপরের PO ফিল্ড
#                                               (কোলনের পরের পুরো লেখাটা)
#       'CTN : 5 PLY PRINTED ...'            -> Ply (5)
#   - টেবিল (হেডার: STYLE NO / ORDER NO / COLOR / SIZE / CARTON MEAS. /
#     CARTON QTY / TTL QTY / UNIT)। STYLE NO, ORDER NO, COLOR শুধু ব্লকের
#     প্রথম রো-তে লেখা থাকে, নিচের সাইজ-রো গুলোতে ফাঁকা — তাই forward-fill।
#       STYLE NO     -> Gmt. Style No
#       ORDER NO     -> Gmt. PO
#       COLOR        -> Reference/SKU Number
#       SIZE         -> Pack Type এবং GMT Size (দুটোতেই)
#       CARTON MEAS. -> Length / Width / Height ('58 X 38 X 38 CM')
#       CARTON QTY   -> Order Qty (ফাঁকা/0 হলে ওই রো বাদ)
#       TTL QTY      -> ব্যবহার হয় না (ব্লকের সামারি, শুধু যাচাইয়ের কাজে)
#       'PCS PER CTN' (L কলামের নোট) -> ব্যবহার হয় না
#   - 'TOTAL' রো পেলে টেবিল শেষ।
#   - Item Name: UI সিলেকশন (ডিফল্ট Master Carton)।
#   - Ply: CTN লাইনের 'N PLY' থেকে; না পেলে UI-এর manual_ply; তাও না থাকলে N/A।
#
# প্রতিটা আইটেমে '_header_po' (ফাইলের ORDER NUMBER লেখা) থাকে — টেমপ্লেটের
# উপরের PO ফিল্ডে বসানোর জন্য (derive_po_header দেখুন)।
# ---------------------------------------------------------------------------

NA = 'N/A'


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    if isinstance(v, float) and pd.isna(v):
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def _num(v):
    """সংখ্যা হলে float, নাহলে None।"""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    try:
        return float(str(v).replace(',', '').strip())
    except (TypeError, ValueError):
        return None


def _parse_measurement(m):
    """'58 X 38 X 38 CM' -> (58.0, 38.0, 38.0)। শেষের 'CM' ইত্যাদি উপেক্ষা।
    দুটো সংখ্যা থাকলে হাইট ফাঁকা।"""
    nums = re.findall(r'\d+(?:\.\d+)?', str(m or ''))
    vals = [float(n) for n in nums[:3]]
    while len(vals) < 3:
        vals.append('')
    return vals[0], vals[1], vals[2]


def _load_rows(file_stream, filename):
    # outhouse_extractor-এর মাল্টি-ইঞ্জিন রিডার (xlrd/calamine/LibreOffice
    # ফলব্যাক) ব্যবহার — lazy import, যাতে circular import না হয়।
    try:
        from outhouse_extractor import _read_excel_rows
        return _read_excel_rows(file_stream, filename)
    except ImportError:
        file_stream.seek(0)
        return pd.read_excel(file_stream, header=None).values.tolist()


def _find_header(rows, max_scan=30):
    for i, row in enumerate(rows[:max_scan]):
        normed = [_norm(c) for c in row]
        if ('styleno' in normed and 'orderno' in normed
                and 'size' in normed and 'cartonqty' in normed):
            return i, {n: ci for ci, n in enumerate(normed) if n and n not in ()}
    return None, None


def _header_text(rows, header_idx):
    """হেডার-এরিয়ার সব সেল এক লাইনে করে রিটার্ন (ORDER NUMBER / CTN খোঁজার জন্য)।"""
    cells = []
    for row in rows[:header_idx]:
        for c in row:
            t = _clean(c)
            if t:
                cells.append(t)
    return cells


def _extract_header_po(cells):
    """'ORDER NUMBER : 1/2026/01182- HH132' -> '1/2026/01182- HH132'
    (কোলনের পরের পুরো লেখাটা হুবহু, কোনো পরিবর্তন ছাড়া — শুধু দুই পাশের স্পেস ট্রিম)।"""
    for t in cells:
        m = re.search(r'order\s*number\s*:\s*(.+)$', t, re.I)
        if m:
            return m.group(1).strip()
    return ''


def _extract_ply(cells):
    for t in cells:
        if re.match(r'^\s*ctn\b', t, re.I) or 'ply' in t.lower():
            m = re.search(r'(\d+)\s*ply', t, re.I)
            if m:
                return m.group(1)
    return ''


def read_dewan_style_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """মূল entry point। এই ফরম্যাট না হলে ValueError ছোঁড়ে।"""
    rows = _load_rows(file_stream, filename)

    header_idx, heads = _find_header(rows)
    if header_idx is None:
        raise ValueError("'STYLE NO / ORDER NO / SIZE / CARTON QTY' হেডিং-সহ Dewan/Rossini টেবিল পাওয়া যায়নি")

    c_style = heads['styleno']
    c_po = heads['orderno']
    c_color = heads.get('color')
    c_size = heads['size']
    c_meas = heads.get('cartonmeas')
    c_qty = heads['cartonqty']

    cells = _header_text(rows, header_idx)
    header_po = _extract_header_po(cells)
    ply_found = _extract_ply(cells)
    ply_value = ply_found or (manual_ply.strip() if manual_ply else NA)
    item_name = item_name_override or 'Master Carton'

    def get(row, ci):
        if ci is None or ci >= len(row):
            return ''
        return row[ci]

    items = []
    last_style = last_po = last_color = ''
    for row in rows[header_idx + 1:]:
        if row is None:
            continue
        size_txt = _clean(get(row, c_size))
        # TOTAL রো (SIZE কলামে 'TOTAL') পেলে টেবিল শেষ
        if _norm(size_txt).startswith('total'):
            break

        # ব্লকের প্রথম রো-তে লেখা মানগুলো নিচের রো-তে forward-fill
        s = _clean(get(row, c_style))
        if s:
            last_style = s
        p = _clean(get(row, c_po))
        if p:
            last_po = p
        c = _clean(get(row, c_color))
        if c:
            last_color = c

        qty = _num(get(row, c_qty))
        if not size_txt or qty is None or qty <= 0:
            continue  # ফাঁকা কোয়ান্টিটির রো বসবে না

        length, width, height = _parse_measurement(get(row, c_meas))
        items.append({
            'item_name': item_name,
            'ewo_no': NA,
            'style_no': last_style,
            'po_no': last_po,
            'length': length,
            'width': width,
            'height': height,
            'ply': ply_value,
            'qty': round(qty),
            'pack_type': size_txt,
            'reference': last_color,
            'color': '',
            'size': size_txt,          # GMT Size কলাম
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_header_po': header_po,   # টেমপ্লেটের উপরের PO ফিল্ডের জন্য
            '_source_file': filename,
        })

    if not items:
        raise ValueError("হেডার পাওয়া গেছে কিন্তু কোয়ান্টিটি-সহ কোনো ডাটা রো পাওয়া যায়নি")

    return items