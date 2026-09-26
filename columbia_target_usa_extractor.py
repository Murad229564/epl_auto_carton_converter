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
    """শিটের উপরের টাইটেল-এরিয়ায় (রো ১ থেকে max_scan) 'ELASTIC' শব্দ
    থাকলে True রিটার্ন করে (Item Name override করার জন্য)।"""
    for r in range(1, max_scan + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and 'elastic' in str(v).lower():
                return True
    return False


def read_columbia_target_usa_style_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """মূল entry point। 'Carton & Poly Order Sheet' শিট থেকেই শুধু ডাটা
    নেওয়া হয় (অন্য শিট — 'Measurement sheet'/'Commit-PO' — ব্যবহার হয় না)।
    এই ফরম্যাট না হলে (শিট/হেডার না পেলে) খালি লিস্ট [] রিটার্ন করে।"""
    wb = load_workbook(file_stream, data_only=True)

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

    return all_items
