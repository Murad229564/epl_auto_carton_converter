import re
from openpyxl import load_workbook

# ---------------------------------------------------------------------------
# PRUDENT FASHION LTD. / Norp Knit Industries Ltd. — Buyer: Kohl's
# Carton বুকিং এক্সেল ফরম্যাট।
#   - একই গ্রুপের দুই কাস্টমারের (নাম আলাদা হতে পারে) একই ফরম্যাট — তাই
#     এই extractor customer name দেখে না, শুধু ফরম্যাট (হেডার সিগনেচার)
#     দেখে চেনে, যেকোনো কাস্টমারের আপলোডেই কাজ করবে।
#   - একাধিক শিট/একাধিক ফাইল থাকতে পারে (Simba-র মতোই), প্রতিটা শিট থেকেই
#     ডাটা নেওয়া হয়।
#   - Item Name টেবিলের বাইরে (উপরে বা নিচে) কোথাও 'ELASTIC' শব্দ-সহ একটা
#     নোট-লাইন থেকে বোঝা যায়:
#       - 'NO NEED ELASTIC' / 'NO ELASTIC' -> Master Carton
#       - 'NEED ELASTIC' / 'WITH ELASTIC' (কিন্তু 'NO' ছাড়া) -> Elastic
#         Hanger Carton
#       - কিছু না পাওয়া গেলে ডিফল্ট Master Carton
#   - PO Number টেবিলের উপরের ইনফো-ব্লকে 'PO NUMBER' লেবেলের পাশে থাকে,
#     এটাই পুরো শিটের জন্য PO No এবং EWO No দুটোতেই বসে (এই ফরম্যাটে আলাদা
#     কোনো EWO নম্বর নেই)।
#   - Style No টেবিলের 'STYLE' কলাম থেকে (উপরের ইনফো-ব্লকের STYLE NUMBER
#     থেকে না — কারণ টেবিলে একাধিক স্টাইল-ভ্যারিয়েন্ট (যেমন RS/RH/HS/HH)
#     মিশ্রিত থাকতে পারে একই শিটে)।
#   - Reference <- COLOR কলাম, Pack Type <- ITEM UPC NUMBER কলাম।
#   - Ply সবসময় ফিক্সড ৫ (ইউজারের নির্দেশ অনুযায়ী)।
#   - Measurement টেবিলে ইঞ্চিতে থাকে (দুই রকম টেক্সট-প্যাটার্ন দেখা গেছে:
#     '(L-20) × (W-12) × (H-6.5)' এবং '(L)18.5 × (W)13 × (H)5.5 Inc.') —
#     regex দুটো প্যাটার্নই হ্যান্ডেল করে। টেমপ্লেটের measurement_unit সবসময়
#     'Inch' বসবে (CM না)।
#   - সাইজ-কলামের সংখ্যা (S/M/L/XL বনাম S/S HUSKY/M/M HUSKY...) শিট-ভেদে
#     ভিন্ন হওয়ায় MEASUREMENT/CARTON QTY কলামের পজিশন শিফট হয় — তাই
#     header-label স্ক্যান করে ডাইনামিকভাবে বের করা হয়, fixed index ধরে
#     রাখা হয় না (norp/Simba-র মতোই)।
# ---------------------------------------------------------------------------


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


# দুই রকম প্যাটার্নই কভার করে: '(L-20) × (W-12) × (H-6.5)' এবং
# '(L)18.5 × (W)13 × (H)5.5 Inc.'
_MEASUREMENT_RE = re.compile(
    r'L[-)]?\s*(\d+(?:\.\d+)?).*?W[-)]?\s*(\d+(?:\.\d+)?).*?H[-)]?\s*(\d+(?:\.\d+)?)',
    re.I | re.S,
)


def _parse_measurement(text):
    if not text:
        return '', '', ''
    m = _MEASUREMENT_RE.search(str(text))
    if m:
        return m.group(1), m.group(2), m.group(3)
    return '', '', ''


def _find_header_row(ws, max_scan=60):
    """'STYLE' (কলাম A), 'ITEM UPC NUMBER' (কলাম B), 'COLOR' (কলাম C) —
    এই তিনটার কম্বিনেশন norp/Simba/AEO ফরম্যাটের সাথে গুলিয়ে যায় না।"""
    for r in range(1, max_scan + 1):
        a = _norm(ws.cell(row=r, column=1).value)
        b = _norm(ws.cell(row=r, column=2).value)
        c = _norm(ws.cell(row=r, column=3).value)
        if a == 'style' and b == 'itemupcnumber' and c == 'color':
            return r
    return None


def _build_dynamic_col_map(ws, header_row):
    col_map = {'style_no': 1, 'pack_type': 2, 'reference': 3}
    for c in range(1, ws.max_column + 1):
        label = _norm(ws.cell(row=header_row, column=c).value)
        if not label:
            continue
        if label == 'measurement':
            col_map['measurement'] = c
        elif 'cartonqty' in label:
            col_map['qty'] = c
    return col_map


def _extract_po_no(ws, max_scan=25):
    """'PO NUMBER' লেবেলের একই রো-তে প্রথম numeric ভ্যালু খুঁজে বের করে —
    এই লেবেল আর ভ্যালুর মাঝের কলাম-দূরত্ব ফাইল-ভেদে বদলাতে পারে, তাই fixed
    কলাম না ধরে scan করা হচ্ছে।"""
    for r in range(1, max_scan + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and _norm(v) == 'ponumber':
                for c2 in range(c + 1, ws.max_column + 1):
                    v2 = ws.cell(row=r, column=c2).value
                    if isinstance(v2, (int, float)):
                        return str(int(v2)) if float(v2).is_integer() else str(v2)
                    if v2 is not None and _clean(v2) and re.search(r'\d{4,}', str(v2)):
                        m = re.search(r'\d{4,}', str(v2))
                        return m.group(0)
    return ''


def _classify_item_name(ws):
    """পুরো শিটে (টেবিলের উপরে/নিচে) 'ELASTIC' শব্দ-সহ প্রথম লাইনটা খুঁজে
    Item Name ঠিক করে। কিছু না পেলে ডিফল্ট Master Carton।"""
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            t = str(v).lower()
            if 'elastic' not in t:
                continue
            if re.search(r'\bno\b[^.]*elastic', t):
                return 'Master Carton'
            return 'Elastic Hanger Carton'
    return 'Master Carton'


def read_pfl_style_excel(file_stream, filename=''):
    """মূল entry point। এই ফরম্যাট না হলে (হেডার না মিললে) খালি লিস্ট [] রিটার্ন
    করে, যাতে outhouse_extractor.py-এর auto-detect চেইনে পরের ফরম্যাটে
    silently fallback হতে পারে।"""
    wb = load_workbook(file_stream, data_only=True)
    all_items = []

    for sn in wb.sheetnames:
        ws = wb[sn]
        header_row = _find_header_row(ws)
        if header_row is None:
            continue

        col_map = _build_dynamic_col_map(ws, header_row)
        if 'measurement' not in col_map or 'qty' not in col_map:
            continue  # প্রত্যাশিত কলাম পাওয়া যায়নি — এই ফরম্যাট না

        po_no = _extract_po_no(ws)
        item_name = _classify_item_name(ws)

        style_col = col_map['style_no']
        upc_col = col_map['pack_type']
        color_col = col_map['reference']
        meas_col = col_map['measurement']
        qty_col = col_map['qty']

        r = header_row + 1
        max_row = ws.max_row
        while r <= max_row:
            row_values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
            if any(_norm(v) == 'gttl' for v in row_values):
                break  # টেবিলের শেষ — এর পরের সব রো ফুটার/নোট, ডাটা না

            qty_val = ws.cell(row=r, column=qty_col).value
            style_val = ws.cell(row=r, column=style_col).value

            # data-row চেনার উপায়: qty numeric আর style_no ফাঁকা না —
            # এতে G.TTL/টোটাল রো আর টেবিলের নিচের নোট-লাইনগুলো (যেখানে
            # style_no কলাম ফাঁকা থাকে) নিজে থেকেই বাদ পড়ে যায়।
            if not isinstance(qty_val, (int, float)) or not _clean(style_val):
                r += 1
                continue

            length, width, height = _parse_measurement(ws.cell(row=r, column=meas_col).value)
            if not length:
                r += 1
                continue

            all_items.append({
                'item_name': item_name,
                'ewo_no': po_no,
                'style_no': _clean(style_val),
                'po_no': po_no,
                'length': length,
                'width': width,
                'height': height,
                'ply': '5',  # ইউজারের নির্দেশ অনুযায়ী — ফিক্সড ৫ প্লাই
                'qty': qty_val,
                'pack_type': _clean(ws.cell(row=r, column=upc_col).value),
                'reference': _clean(ws.cell(row=r, column=color_col).value),
                'remarks': '',
                'color': '',
                'size': '',
                'delivery_date': '',
                'measurement_unit': 'Inch',
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': sn,
            })
            r += 1

    return all_items
