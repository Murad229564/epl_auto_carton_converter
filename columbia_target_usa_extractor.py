import re
from openpyxl import load_workbook

# ---------------------------------------------------------------------------
# Columbia Apparels Limited / Columbia Garments Limited — Buyer: Target-USA
# 'Carton & Poly Order Sheet' এক্সেল ফরম্যাট (Non-Print-Carton-Order)।
#   - এই ফরম্যাটে বুকিং হয় Style + Wash Color + Measurement — এই তিনটার
#     কম্বিনেশন-ওয়াইজ। সোর্স ফাইলে প্রতিটা GMS Size (28x30, 28x32 ...)-এর
#     জন্য আলাদা রো থাকে, কিন্তু আমরা ব্রেকডাউন করি না — একই
#     (Style, Wash Color, Measurement) কম্বিনেশনের সব রো-র Quantity (Pcs)
#     যোগ করে একটাই লাইন-আইটেম বানানো হয় (ইউজার-কনফার্মড)।
#   - ফাইলে একাধিক শিট থাকে ('Carton & Poly Order Sheet', 'Measurement
#     sheet', 'Commit-PO') — শুধু 'Carton & Poly Order Sheet'-ই আসল বুকিং
#     ডাটা, বাকি দুইটা ব্যবহার হয় না।
#   - Excel-এ hidden করা রো (সাধারণত qty=0 বা ব্যবহার না-হওয়া Wash Color
#     ব্লক) সম্পূর্ণ বাদ — ইউজার-কনফার্মড নিয়ম: 'যেটুকু ভিজিবল ডাটা
#     পাবেন সেইটুকু নিয়েই কাজ করবেন, তার বেশি না'। qty<=0 রো-ও বাদ
#     (রিডানডেন্ট সেফটি-নেট, hidden-row স্কিপ সাধারণত এটাই কভার করে)।
#   - শেষে একটা 'Total Qty (Pcs)' সামারি-রো থাকে (Po No কলাম ফাঁকা) —
#     Po No ফাঁকা চেক দিয়েই এটা এমনিতে বাদ পড়ে যায়।
#   - ইউজার-কনফার্মড ম্যাপিং:
#       Po No        -> Gmt. PO
#       Style No     -> Gmt. Style No
#       Wash Color   -> Reference/SKU Number
#       Measurement  -> Length/Width/Height
#       Quantity (Pcs) [গ্রুপ-সামড] -> Order Qty
#       Type ('5 PLY'/'3 PLY') -> Ply (সংখ্যাটা বের করে নেওয়া হয়)
#   - Item Name: ডিফল্ট 'Master Carton', কিন্তু শিটের উপরের টাইটেল-এরিয়ায়
#     (রো ১-৩, যেমন 'Non-Print-Carton-Order for ...') 'ELASTIC' শব্দ
#     থাকলে পুরো ফাইলের জন্য Item Name 'Elastic Hanger Carton' হয়ে যায়
#     (ইউজার-কনফার্মড, non-print-এর জায়গাতেই সাধারণত এই শব্দ থাকে)।
#
#   নতুন ফরম্যাট (V2) — "5 PLY SOLID CARTON BOOKING ... STYLE ... PO ..."
#   ফাইল, শিটের নাম যেমন "CARTON COM 209126 C1'27" (আগের নামের শিট না
#   থাকলে তবেই V2 চেক হয়, তাই আগের ফরম্যাটের আচরণ একদম অপরিবর্তিত):
#     হেডার: PO | ITEM | STYLE | Description | GMTS QTY | GMTS SIZE |
#     VCP/SSP | CARTON PLY | CARTON MEASUREMENT CM | CARTON QTY (PCS)।
#       - PO শুধু প্রথম রো-তে (merged cell) — forward-fill করা হয়
#         (merged cell-এর আসল ভ্যালু hidden রো-তে থাকলেও ধরা হয়)।
#       - কোনো Wash Color কলাম নেই -> Reference ফাঁকা (কালার কলাম থাকলে
#         সেটাই Reference-এ যাবে, গ্রুপিং-এও ধরা হবে)।
#       - GMTS SIZE (29x30 ...) ব্রেকডাউন হয় না — (PO, Style, Measurement)
#         ধরে CARTON QTY (PCS) যোগ হয়ে একটাই লাইন (ইউজার-কনফার্মড:
#         কালার না থাকলে শুধু মেজারমেন্ট ধরে সাম)।
#       - CARTON PLY ('5 PLY') -> Ply; Total/ফাঁকা Style রো বাদ।
#
#   ডেলিভারি ডেট (দুই ফরম্যাটেই): শিটের নিচের দিকে 'Delivery Date' লেবেলের
#   কাছে (একই রো-র ডানে বা কয়েক রো নিচে) যে তারিখ থাকে সেটা প্রতিটা
#   লাইন-আইটেমে '_file_delivery_date' key-তে (টেমপ্লেটের ফরম্যাটে, যেমন
#   '21-Oct-2026') বসানো হয় — builder.py এটা থাকলে Delivery Date ঘরে এটাই
#   বসায়, না থাকলে আগের মতো UI/অটো ডেট।
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


_MEAS_RE = re.compile(
    r'L[\s\-]*(\d+\.?\d*)\s*[xX×]\s*W[\s\-]*(\d+\.?\d*)\s*[xX×]\s*H[\s\-]*(\d+\.?\d*)', re.I)


def _parse_measurement(text):
    """'L-58  x  W-34  x  H-31 CM' -> ('58','34','31')।"""
    if not text:
        return '', '', ''
    m = _MEAS_RE.search(str(text))
    if not m:
        return '', '', ''
    return _fmt_num(m.group(1)), _fmt_num(m.group(2)), _fmt_num(m.group(3))


_PLY_RE = re.compile(r'(\d+)\s*PLY', re.I)


def _parse_ply(text):
    m = _PLY_RE.search(str(text or ''))
    return m.group(1) if m else ''


def _get_hidden_rows_cols(ws):
    hidden_rows = {r for r, dim in ws.row_dimensions.items() if dim.hidden}
    from openpyxl.utils import column_index_from_string
    hidden_cols = set()
    for letter, dim in ws.column_dimensions.items():
        if dim.hidden:
            try:
                hidden_cols.add(column_index_from_string(letter))
            except Exception:
                pass
    return hidden_rows, hidden_cols


def _find_header_row(ws, max_scan=15):
    """কলাম A-তে 'Po No' আর কলাম C-তে 'Style No' — এই দুটো একসাথে থাকা
    রো-কেই হেডার ধরা হয়।"""
    for r in range(1, max_scan + 1):
        a = _norm(ws.cell(row=r, column=1).value)
        c = _norm(ws.cell(row=r, column=3).value)
        if a == 'pono' and c == 'styleno':
            return r
    return None


def _build_col_map(ws, header_row, hidden_cols):
    col_map = {}
    for c in range(1, ws.max_column + 1):
        if c in hidden_cols:
            continue
        label = _norm(ws.cell(row=header_row, column=c).value)
        if not label:
            continue
        if label == 'pono':
            col_map['po_no'] = c
        elif label == 'styleno':
            col_map['style_no'] = c
        elif label == 'washcolor':
            col_map['reference'] = c
        elif label == 'measurement':
            col_map['measurement'] = c
        elif 'quantitypcs' in label:
            col_map['qty'] = c
        elif label == 'type':
            col_map['ply'] = c
    return col_map


def _detect_elastic(ws, max_scan=5):
    """শিটের উপরের টাইটেল-এরিয়ায় (রো ১ থেকে max_scan) 'ELASTIC' বা
    'HANGER' শব্দ থাকলে True রিটার্ন করে (Item Name override করার জন্য)।
    এছাড়া হেডার রো-তে একটা 'Hanger Position' নামের কলাম থাকলেও (অতিরিক্ত
    সহায়ক সংকেত, কিছু ফাইলে টাইটেল-এ 'Hanger' শব্দ না থাকলেও এই কলাম
    দিয়েই বোঝা যায়) True রিটার্ন করে।"""
    for r in range(1, max_scan + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            t = str(v).lower()
            if 'elastic' in t or 'hanger' in t:
                return True
    return False


_DATE_FORMATS = ('%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y', '%d-%b-%Y', '%d %b %Y', '%Y-%m-%d')


def _delivery_date_format():
    try:
        from config import DELIVERY_DATE_FORMAT
        return DELIVERY_DATE_FORMAT
    except Exception:
        return '%d-%b-%Y'


def _to_date_str(v):
    """datetime/date বা পরিচিত ফরম্যাটের টেক্সট -> টেমপ্লেটের ডেট-ফরম্যাটে
    স্ট্রিং। বোঝা না গেলে None।"""
    import datetime
    fmt = _delivery_date_format()
    if isinstance(v, datetime.datetime):
        return v.strftime(fmt)
    if isinstance(v, datetime.date):
        return v.strftime(fmt)
    if isinstance(v, str):
        t = _clean(v)
        for f in _DATE_FORMATS:
            try:
                return datetime.datetime.strptime(t, f).strftime(fmt)
            except ValueError:
                continue
    return None


def _find_delivery_date(ws, max_rows=600, max_cols=30):
    """'Delivery Date' লেবেল (শুধু Date — 'Delivery Address' না) খুঁজে তার
    একই রো-র ডানে বা ঠিক নিচের কয়েক রো-তে প্রথম তারিখটা রিটার্ন করে।
    পুরনো ফরম্যাটে তারিখ লেবেলের ঠিক নিচের ঘরে, নতুন ফরম্যাটে লেবেলের
    দুই রো নিচে অন্য কলামে — দুটোই ধরা পড়ে। তারিখ না পেলে None।
    (উপরের Booking-date-জাতীয় তারিখ ধরা হয় না, কারণ শুধু লেবেলের
    পরের জায়গা দেখা হয়।)"""
    last_row = min(ws.max_row, max_rows)
    last_col = min(ws.max_column, max_cols)
    for r in range(1, last_row + 1):
        for c in range(1, last_col + 1):
            v = ws.cell(row=r, column=c).value
            if v is None or not isinstance(v, str):
                continue
            if _norm(v) != 'deliverydate':
                continue
            for rr in range(r, min(r + 4, last_row) + 1):
                for cc in range(c, min(c + 12, last_col) + 1):
                    if rr == r and cc == c:
                        continue
                    d = _to_date_str(ws.cell(row=rr, column=cc).value)
                    if d:
                        return d
    return None


def _attach_delivery_date(items, ws):
    d = _find_delivery_date(ws)
    if d:
        for it in items:
            it['_file_delivery_date'] = d
    return items


def _find_header_row_v2(ws, max_scan=25):
    """নতুন ফরম্যাটের হেডার: একই রো-তে 'PO', 'STYLE', 'CARTON MEASUREMENT...'
    আর 'CARTON QTY...' লেবেল।"""
    for r in range(1, max_scan + 1):
        labels = [_norm(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        if ('po' in labels and 'style' in labels
                and any('cartonmeasurement' in l for l in labels)
                and any('cartonqty' in l for l in labels)):
            return r
    return None


def _build_col_map_v2(ws, header_row, hidden_cols):
    col_map = {}
    for c in range(1, ws.max_column + 1):
        if c in hidden_cols:
            continue
        label = _norm(ws.cell(row=header_row, column=c).value)
        if not label:
            continue
        if label == 'po':
            col_map['po_no'] = c
        elif label == 'style':
            col_map['style_no'] = c
        elif 'cartonmeasurement' in label:
            col_map['measurement'] = c
        elif 'cartonqty' in label:
            col_map['qty'] = c
        elif label in ('cartonply', 'ply', 'type'):
            col_map['ply'] = c
        elif label in ('washcolor', 'color', 'colour'):
            col_map['reference'] = c
    return col_map


def _read_v2(ws, sheet_name, filename, item_name_override, manual_ply):
    """নতুন ফরম্যাট (V2) — উপরের মডিউল-কমেন্ট দেখুন।"""
    header_row = _find_header_row_v2(ws)
    if header_row is None:
        return []
    hidden_rows, hidden_cols = _get_hidden_rows_cols(ws)
    col_map = _build_col_map_v2(ws, header_row, hidden_cols)
    required = ('po_no', 'style_no', 'measurement', 'qty')
    if not all(k in col_map for k in required):
        return []

    default_item_name = item_name_override or 'Master Carton'
    if _detect_elastic(ws):
        default_item_name = 'Elastic Hanger Carton'

    po_col = col_map['po_no']
    style_col = col_map['style_no']
    meas_col = col_map['measurement']
    qty_col = col_map['qty']
    ply_col = col_map.get('ply')
    ref_col = col_map.get('reference')

    groups, group_order, group_ply = {}, [], {}
    current_po = ''
    r = header_row + 1
    max_row = ws.max_row
    consecutive_blank = 0
    BLANK_STOP_THRESHOLD = 100
    while r <= max_row:
        # PO merged cell-এর আসল ভ্যালু hidden রো-তে থাকলেও ধরা হয় (forward-fill)
        po_val = ws.cell(row=r, column=po_col).value
        if po_val is not None and _clean(po_val):
            current_po = _clean(po_val)

        style_val = _clean(ws.cell(row=r, column=style_col).value)
        if not style_val:
            # Total রো / ফাঁকা রো / টেবিলের নিচের লেখা — ডাটা না
            consecutive_blank += 1
            if consecutive_blank >= BLANK_STOP_THRESHOLD:
                break
            r += 1
            continue
        consecutive_blank = 0

        if r in hidden_rows:
            r += 1
            continue

        qty_val = ws.cell(row=r, column=qty_col).value
        if not _is_num(qty_val) or float(qty_val) <= 0:
            r += 1
            continue

        length, width, height = _parse_measurement(ws.cell(row=r, column=meas_col).value)
        if not length:
            r += 1
            continue

        ref_val = _clean(ws.cell(row=r, column=ref_col).value) if ref_col else ''
        ply_val = ''
        if ply_col:
            ply_val = _parse_ply(ws.cell(row=r, column=ply_col).value)
        ply_val = ply_val or (manual_ply.strip() if manual_ply else '')

        key = (current_po, style_val, ref_val, length, width, height)
        if key not in groups:
            groups[key] = 0
            group_order.append(key)
            group_ply[key] = ply_val
        groups[key] += round(float(qty_val))
        r += 1

    items = []
    for key in group_order:
        po_c, style_v, ref_v, length, width, height = key
        items.append({
            'item_name': default_item_name,
            'ewo_no': 'N/A',
            'style_no': style_v,
            'po_no': po_c,
            'length': length,
            'width': width,
            'height': height,
            'ply': group_ply[key],
            'qty': groups[key],
            'pack_type': '',
            'reference': ref_v,
            'color': '',
            'size': '',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_sheet': sheet_name,
            '_source_file': filename,
        })
    return _attach_delivery_date(items, ws)


def read_columbia_target_usa_style_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """মূল entry point। আগে পুরনো ফরম্যাট ('Carton & Poly Order Sheet' শিট)
    — সেই শিট থাকলে শুধু সেটাই (আগের আচরণ হুবহু)। না থাকলে নতুন ফরম্যাট
    (V2): প্রতিটা visible শিটে V2-হেডার খোঁজা হয়। কিছুই না মিললে []।"""
    wb = load_workbook(file_stream, data_only=True)

    has_v1_sheet = any(_norm(sn) == _norm('Carton & Poly Order Sheet') for sn in wb.sheetnames)
    if has_v1_sheet:
        return _read_v1(wb, filename, item_name_override, manual_ply)

    items = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        if ws.sheet_state != 'visible':
            continue  # হাইড শিট — ডাটা নেওয়া হবে না
        items.extend(_read_v2(ws, sn, filename, item_name_override, manual_ply))
    return items


def _read_v1(wb, filename, item_name_override, manual_ply):
    """পুরনো ফরম্যাট ('Carton & Poly Order Sheet') — আগের কোড অপরিবর্তিত;
    শুধু শেষে ডেলিভারি ডেট যোগ হয়েছে।"""

    target_sheet = None
    for sn in wb.sheetnames:
        if _norm(sn) == _norm('Carton & Poly Order Sheet'):
            target_sheet = sn
            break
    if target_sheet is None:
        return []

    ws = wb[target_sheet]
    header_row = _find_header_row(ws)
    if header_row is None:
        return []

    hidden_rows, hidden_cols = _get_hidden_rows_cols(ws)
    col_map = _build_col_map(ws, header_row, hidden_cols)
    required = ('po_no', 'style_no', 'reference', 'measurement', 'qty', 'ply')
    if not all(k in col_map for k in required):
        return []  # প্রত্যাশিত কলাম পাওয়া যায়নি (বা hidden) — এই ফরম্যাট না

    default_item_name = item_name_override or 'Master Carton'
    if _detect_elastic(ws):
        default_item_name = 'Elastic Hanger Carton'

    po_col = col_map['po_no']
    style_col = col_map['style_no']
    ref_col = col_map['reference']
    meas_col = col_map['measurement']
    qty_col = col_map['qty']
    ply_col = col_map['ply']

    # (PO, style, wash color, L, W, H) -> summed qty। PO-কে গ্রুপ-কী-তে
    # রাখা হয়েছে ইচ্ছাকৃতভাবে — একই ফাইলে একাধিক PO থাকলে (কম দেখা যায়,
    # কিন্তু সম্ভব) একই Style/Color/Measurement কম্বিনেশন দুই ভিন্ন PO-তে
    # থাকলেও তাদের Qty মিশে না যায়, প্রতিটা PO নিজের মতো আলাদা লাইনে থাকে
    # (ইউজার-কনফার্মড)। insertion-order আলাদা লিস্টে ট্র্যাক করা হয় (প্রথম
    # যেই কম্বিনেশন আসবে সেটাই আউটপুটে আগে)।
    groups = {}
    group_order = []
    group_ply = {}
    po_seen_order = []  # ফাইলে যতগুলো ইউনিক PO আছে, প্রথম-দেখা ক্রমে

    r = header_row + 1
    max_row = ws.max_row
    # কিছু সোর্স ফাইলে ws.max_row বাস্তব ডাটার তুলনায় বিশাল হয়ে যায় (যেমন
    # ১০+ লাখ রো, পুরো কলামজুড়ে ফরম্যাটিং প্রয়োগ করা থাকলে Excel পুরো
    # রেঞ্জটাকেই 'ব্যবহৃত এরিয়া' ধরে নেয়) — সেই পুরো রেঞ্জ সেল-বাই-সেল
    # স্ক্যান করলে রিকোয়েস্ট টাইমআউট হয়ে যায়। তাই আসল ডাটা-রো (Po No/qty/
    # measurement যেকোনো একটা থাকা রো) না পেয়ে একটানা অনেক রো ফাঁকা গেলে
    # ধরে নেওয়া হয় টেবিল শেষ, বাকি রো আর স্ক্যান করা হয় না।
    consecutive_blank = 0
    BLANK_STOP_THRESHOLD = 100
    while r <= max_row:
        if r in hidden_rows:
            r += 1
            continue

        po_val = ws.cell(row=r, column=po_col).value
        if po_val is None or not _clean(po_val):
            consecutive_blank += 1
            if consecutive_blank >= BLANK_STOP_THRESHOLD:
                break
            r += 1
            continue  # Po No ফাঁকা — 'Total Qty (Pcs)' সামারি-রো বা খালি রো
        consecutive_blank = 0

        qty_val = ws.cell(row=r, column=qty_col).value
        if not _is_num(qty_val) or float(qty_val) <= 0:
            r += 1
            continue  # qty না থাকলে/০ হলে বাদ

        length, width, height = _parse_measurement(ws.cell(row=r, column=meas_col).value)
        if not length:
            r += 1
            continue

        po_val_clean = _clean(po_val)
        if po_val_clean not in po_seen_order:
            po_seen_order.append(po_val_clean)

        style_val = _clean(ws.cell(row=r, column=style_col).value)
        ref_val = _clean(ws.cell(row=r, column=ref_col).value)
        ply_val = _parse_ply(ws.cell(row=r, column=ply_col).value) or (manual_ply.strip() if manual_ply else '')

        key = (po_val_clean, style_val, ref_val, length, width, height)
        if key not in groups:
            groups[key] = 0
            group_order.append(key)
            group_ply[key] = ply_val
        groups[key] += round(float(qty_val))
        r += 1

    all_items = []
    for key in group_order:
        po_val_clean, style_val, ref_val, length, width, height = key
        all_items.append({
            'item_name': default_item_name,
            'ewo_no': 'N/A',
            'style_no': style_val,
            'po_no': po_val_clean,
            'length': length,
            'width': width,
            'height': height,
            'ply': group_ply[key],
            'qty': groups[key],
            'pack_type': '',
            'reference': ref_val,
            'color': '',
            'size': '',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_sheet': target_sheet,
            '_source_file': filename,
        })

    return _attach_delivery_date(all_items, ws)