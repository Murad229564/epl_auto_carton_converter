import re
from openpyxl import load_workbook

# ---------------------------------------------------------------------------
# Bettex Industries Ltd — Buyer: C&A BUYING GMBH & CO. KG
# 'Carton shipping mark work order' Excel ফরম্যাট (Uniglory Packaging-কে
# পাঠানো রিকুয়েস্ট)।
#   - একটা ফাইলে একাধিক শিট থাকতে পারে (যেমন 'PO_112049 ROYAL BLUE',
#     'PO_112049 BLACK') — সব visible শিট থেকেই ডাটা নেওয়া হয়। শিটের নামের
#     ওপর কোনো নির্ভরতা নেই (কিছু শিটের নামের রঙ আর ভেতরের রঙ মেলে না)।
#   - শিটের উপরের অংশ থেকে:
#       'Requesting For Carton Shipping Mark of STYLE:2027-1-1293,PO:112043'
#           -> Gmt. Style No (2027-1-1293) এবং Gmt. PO (112043)
#           (এই লাইন না পেলে PO নেওয়া হয় 'ORDER NUMBER' ঘর থেকে)
#       'Box Measurement: 65 X 52 X 15 CM' -> Length / Width / Height (Cm)
#   - শিটে একাধিক 'ব্লক' থাকে (প্রতিটা সাইজের জন্য একটা ব্লক — PP/P/M/G/GG)।
#     প্রতিটা ব্লক শুরু হয় 'SUPPLIER :' লেবেল দিয়ে (কলাম A)। প্রতিটা ব্লক = ১টা
#     Master Carton লাইন:
#       PACK (কলোনের পরের নম্বর, যেমন 12327623) -> Reference/SKU Number
#       'SHIPPING MARK ORDER QTY' কলামের (M) বড় সংখ্যা -> Order Qty
#       Qty ফাঁকা/০ হলে ব্লক বাদ।
#     নিচের 'TOTAL' রো (SUM ফর্মুলা) ব্লক নয় — ব্যবহার হয় না।
#   - Item Name / Ply UI থেকে সিলেক্ট করা (Ply না দিলে 'N/A')।
#   - Pack Type: ডিফল্টে ফাঁকা (আপনি বলেননি)। ব্লকের SIZE (PP/P/M/G/GG) Pack
#     Type-এ বসাতে চাইলে নিচের USE_SIZE_AS_PACK_TYPE = True করুন।
# ---------------------------------------------------------------------------

NA = 'N/A'

USE_SIZE_AS_PACK_TYPE = False


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return re.sub(r'\s+', ' ', str(v)).strip()


def _to_num(v):
    if v is None or isinstance(v, bool):
        return None
    try:
        return float(str(v).replace(',', '').strip())
    except ValueError:
        return None


def _fmt_num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return ''
    return str(int(f)) if f == int(f) else str(round(f, 3))


_MEAS_RE = re.compile(r'(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)')
_STYLE_PO_RE = re.compile(r'STYLE\s*[:\-]\s*([^,;]+?)\s*[,;]\s*PO\s*[:\-]?\s*([A-Za-z0-9\-/]+)', re.I)


def _parse_measurement(text):
    m = _MEAS_RE.search(str(text or ''))
    if not m:
        return None
    unit = 'Inch' if re.search(r'\b(inch|in)\b|"', str(text), re.I) else 'Cm'
    return _fmt_num(m.group(1)), _fmt_num(m.group(2)), _fmt_num(m.group(3)), unit


def _right_value(ws, r, c, max_gap=12):
    """সেলের ডানদিকের প্রথম নন-এম্পটি সেলের মান।"""
    for c2 in range(c + 1, min(ws.max_column, c + max_gap) + 1):
        v = ws.cell(row=r, column=c2).value
        if v is not None and _clean(v) != '':
            return v
    return None


def _find_style_po(ws, max_row=15):
    for r in range(1, max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str):
                m = _STYLE_PO_RE.search(v)
                if m:
                    return _clean(m.group(1)), _clean(m.group(2))
    return '', ''


def _find_box_measurement(ws, max_row=20):
    for r in range(1, max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is None or not _norm(v).startswith('boxmeas'):
                continue
            text = str(v).split(':', 1)[1] if ':' in str(v) and _MEAS_RE.search(str(v)) else None
            if text is None:
                text = _right_value(ws, r, c)
            meas = _parse_measurement(text)
            if meas:
                return meas
    return None


def _find_qty_col(ws, max_row=20):
    """'SHIPPING MARK ORDER QTY' হেডারের কলাম।"""
    for r in range(1, max_row + 1):
        for c in range(1, ws.max_column + 1):
            if _norm(ws.cell(row=r, column=c).value).startswith('shippingmarkorderqty'):
                return c
    return None


def _find_blocks(ws):
    """কলাম A-তে 'SUPPLIER' লেবেলের রো-গুলো ব্লকের শুরু; 'TOTAL' রো-তে শেষ।"""
    starts = []
    end_limit = ws.max_row
    for r in range(1, ws.max_row + 1):
        a = _norm(ws.cell(row=r, column=1).value)
        if a.startswith('supplier'):
            starts.append(r)
        elif a == 'total' and starts:
            end_limit = r - 1
            break
    blocks = []
    for i, s in enumerate(starts):
        e = (starts[i + 1] - 1) if i + 1 < len(starts) else end_limit
        blocks.append((s, e))
    return blocks


def read_bettex_style_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """মূল entry point। এই ফরম্যাট না হলে খালি লিস্ট [] রিটার্ন করে।"""
    wb = load_workbook(file_stream, data_only=True)
    items = []
    default_item_name = item_name_override or 'Master Carton'
    ply_value = manual_ply.strip() if manual_ply else NA

    for sn in wb.sheetnames:
        ws = wb[sn]
        if ws.sheet_state != 'visible':
            continue

        qty_col = _find_qty_col(ws)
        blocks = _find_blocks(ws)
        meas = _find_box_measurement(ws)
        if qty_col is None or not blocks or meas is None:
            continue  # এই শিট এই ফরম্যাটের না

        style_no, po_no = _find_style_po(ws)
        hidden_rows = {r for r, dim in ws.row_dimensions.items() if dim.hidden}
        length, width, height, unit = meas

        for start, end in blocks:
            if start in hidden_rows:
                continue

            pack = size = order_no = ''
            qty = None
            for r in range(start, end + 1):
                if qty is None:
                    q = _to_num(ws.cell(row=r, column=qty_col).value)
                    if q is not None:
                        qty = q
                # PACK: ডানদিকের (কলাম F+) 'PACK :' লেবেল, পাশের মান
                for c in range(6, ws.max_column + 1):
                    if _norm(ws.cell(row=r, column=c).value) == 'pack' and not pack:
                        pack = _clean(_right_value(ws, r, c))
                a_norm = _norm(ws.cell(row=r, column=1).value)
                if a_norm == 'size' and not size:
                    size = _clean(ws.cell(row=r, column=2).value)
                if a_norm.startswith('ordernumber') and not order_no:
                    order_no = _clean(_right_value(ws, r, 1))

            if qty is None or qty <= 0:
                continue  # ফাঁকা কোয়ান্টিটি বসানো হয় না

            items.append({
                'item_name': default_item_name,
                'ewo_no': NA,
                'style_no': style_no or NA,
                'po_no': po_no or order_no or NA,
                'length': length,
                'width': width,
                'height': height,
                'ply': ply_value,
                'qty': int(round(qty)),
                'pack_type': (size if USE_SIZE_AS_PACK_TYPE else ''),
                'reference': pack or NA,
                'color': '',
                'size': '',
                'delivery_date': '',
                'measurement_unit': unit,
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': sn,
                '_source_file': filename,
            })

    return items
