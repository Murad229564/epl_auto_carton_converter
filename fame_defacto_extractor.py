import re
from openpyxl import load_workbook

# ---------------------------------------------------------------------------
# Fame Apparels Limited — Buyer: Defacto
# 'Summary Table-English Format' Carton বুকিং এক্সেল ফরম্যাট। দুই রকম
# লে-আউট দেখা গেছে (দুটোই এই একই ফাংশন হ্যান্ডেল করে, প্রতিটা শিটে কোনটা
# ম্যাচ করে সেটা অটো-ডিটেক্ট হয়):
#
#   ফরম্যাট V1 (পুরনো, একটা স্টাইলের বুকিং): উপরের ইনফো-ব্লকে 'STYLE NAME'
#     লেবেলের পাশে একটা সিঙ্গেল স্টাইল নম্বর, টেবিলের হেডার শুরু হয়
#     'Order Number' দিয়ে (কোনো Style/Color কলাম প্রতি-রো-তে থাকে না)।
#     কলাম: Order Number / Prepack Code / [সাইজ কলাম] / TTL QTY /
#     CARTON QTY / CARTON MESURMENT।
#
#   ফরম্যাট V2 (নতুন): টেবিলের হেডার শুরু হয় 'Style Code' দিয়ে (প্রতি-রো-তে
#     স্টাইল থাকে, মাঝে মাঝে 'Season' কলামও থাকে), তারপর Order Number,
#     ColorCode-Name, Prepack Code, [DESCRIPTION/Set Content — নাই হতে
#     পারে], [সাইজ কলাম], Qty. In A Blister/TTL CARTON BOOKING (দরকার
#     নেই), Carton Qty-জাতীয় কলাম (হেডার স্পেলিং ভিন্ন হতে পারে: 'CARTON
#     QTY', 'TTL CARTON QTY', 'TTL C ARTON QTY' ইত্যাদি — তাই flexible
#     ম্যাচ করা হয়), CARTON MEASUREMENT।
#
#   ইউজার-কনফার্মড ম্যাপিং (দুই ফরম্যাটেই):
#       Style (V1: গ্লোবাল STYLE NAME, V2: প্রতি-রো Style Code) -> Gmt. Style No
#       Order Number          -> Gmt. PO
#       ColorCode-Name (V2-তে থাকলে) -> Reference/SKU Number (V1-এ এই
#         কলাম নেই, তাই ফাঁকা)
#       Prepack Code          -> Pack Type (⚠️ আগে V1-এ এটা Reference-এ
#         বসতো, ইউজার-কনফার্মড পরিবর্তনে এখন Pack Type-এ বসে — দুই
#         ফরম্যাটেই একই নিয়ম)
#       CARTON QTY/TTL CARTON QTY/TTL C ARTON QTY (flexible, স্পেলিং-ভেদে)
#         -> Order Qty (TTL CARTON BOOKING/TTL QTY-এর মতো অন্য qty-সদৃশ
#         কলাম ভুল করে ম্যাচ না হয় সেভাবে প্যাটার্ন রাখা হয়েছে)
#       CARTON MEASUREMENT/CARTON MESURMENT (merged cell, forward-fill)
#         -> Length/Width/Height
#       Item Name/Ply         -> UI থেকে সিলেক্ট করা ভ্যালু
#   Hidden রো এবং হাইড/ভেরি-হাইড শীট সম্পূর্ণ বাদ (ইউজার-কনফার্মড, অন্য
#   এক্সট্র্যাক্টরগুলোর মতোই কনভেনশন)।
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


def _po_str(v):
    if isinstance(v, (int, float)):
        return str(int(v)) if float(v).is_integer() else str(v)
    return _clean(v)


_MEAS_RE = re.compile(r'(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)')


def _parse_measurement(text):
    """'60X40X35CM' -> ('60','40','35')।"""
    if not text:
        return '', '', ''
    m = _MEAS_RE.search(str(text))
    if not m:
        return '', '', ''
    return _fmt_num(m.group(1)), _fmt_num(m.group(2)), _fmt_num(m.group(3))


def _extract_style_no(ws, max_scan=10):
    """V1 ফরম্যাটের জন্য: উপরের ইনফো-ব্লকে 'STYLE NAME' লেবেলের পাশে
    (একই রো-তে, ডানদিকের প্রথম নন-এম্পটি সেল) সিঙ্গেল স্টাইল নম্বরটা
    খুঁজে বের করে।"""
    for r in range(1, max_scan + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and _norm(v) == 'stylename':
                for c2 in range(c + 1, ws.max_column + 1):
                    v2 = ws.cell(row=r, column=c2).value
                    if v2 is not None and _clean(v2):
                        return _clean(v2)
    return ''


def _find_header_row_v2(ws, max_scan=25):
    """V2 (নতুন) ফরম্যাট: কলাম A-তে 'Style Code', আর তার ডানদিকের কয়েক
    কলামের মধ্যে (Season কলাম থাকলে/না থাকলে দুটো ক্ষেত্রেই) 'Order
    Number' — এই কম্বিনেশন থাকা রো-কে হেডার ধরা হয়।"""
    for r in range(1, max_scan + 1):
        if _norm(ws.cell(row=r, column=1).value) != 'stylecode':
            continue
        for c2 in range(2, min(6, ws.max_column) + 1):
            if _norm(ws.cell(row=r, column=c2).value) == 'ordernumber':
                return r
    return None


def _find_header_row_v1(ws, max_scan=25):
    """V1 (পুরনো) ফরম্যাট: কলাম A-তে 'Order Number' আর কলাম B-তে
    'Prepack Code' — এই দুটো একসাথে থাকা রো-কেই হেডার ধরা হয়।"""
    for r in range(1, max_scan + 1):
        a = _norm(ws.cell(row=r, column=1).value)
        b = _norm(ws.cell(row=r, column=2).value)
        if a == 'ordernumber' and b.startswith('prepack'):
            return r
    return None


def _build_col_map(ws, header_row):
    """হেডার-লেবেল স্ক্যান করে কলাম-পজিশন বের করে — V1/V2 দুই ফরম্যাটেই
    একই ফাংশন কাজ করে, কারণ শুধু যেসব লেবেল উপস্থিত থাকে সেগুলোই ম্যাপ
    হয় (V1-এ 'stylecode'/'colorcodename' পাওয়া যাবে না, col_map-এ
    সেগুলোর key-ই থাকবে না)। Carton Qty/Measurement কলামের হেডার-স্পেলিং
    ফাইল-ভেদে ভিন্ন হতে পারে (স্পেস/টাইপো), তাই সাব-স্ট্রিং ম্যাচ করা হয়,
    exact match না।"""
    col_map = {}
    for c in range(1, ws.max_column + 1):
        label = _norm(ws.cell(row=header_row, column=c).value)
        if not label:
            continue
        if label == 'stylecode':
            col_map['style_no'] = c
        elif label == 'ordernumber':
            col_map['po_no'] = c
        elif 'colorcode' in label:
            col_map['reference'] = c
        elif label.startswith('prepack'):
            col_map['pack_type'] = c
        elif ('carton' in label or 'ctn' in label) and ('measur' in label or 'mesur' in label):
            col_map['measurement'] = c
        elif ('carton' in label or 'ctn' in label) and 'booking' not in label:
            # Qty কলামের হেডার-স্পেলিং ফাইল-ভেদে অনেক আলাদা হতে পারে:
            # 'CARTON QTY', 'TTL CARTON QTY', 'TTL C ARTON QTY', এমনকি
            # শুধু 'TTL CARTON' (QTY শব্দটাই নেই) — তাই 'carton' থাকা আর
            # 'booking'/'measur' না থাকা কলামকেই qty ধরা হয় (booking-ওয়ালা
            # কলাম যেমন 'TTL CARTON BOOKING' আসলে টোটাল-পিস কাউন্ট, carton
            # কাউন্ট না — তাই বাদ)। প্রথম যেই কলাম মিলবে সেটাই রাখা হয়।
            if 'qty' not in col_map:
                col_map['qty'] = c
    return col_map


def read_fame_defacto_style_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """মূল entry point। প্রতিটা (visible) শিটে প্রথমে V2 (নতুন, বেশি
    স্পেসিফিক সিগনেচার) ফরম্যাট চেক করা হয়, না মিললে V1 (পুরনো) ফরম্যাট
    চেক করা হয়। কোনোটাই না মিললে সেই শিট স্কিপ। এই ফরম্যাট না হলে (কোনো
    শিটেই কিছু না মিললে) খালি লিস্ট [] রিটার্ন করে।"""
    wb = load_workbook(file_stream, data_only=True)
    all_items = []
    default_item_name = item_name_override or 'Master Carton'
    ply_value = manual_ply.strip() if manual_ply else 'N/A'

    for sn in wb.sheetnames:
        ws = wb[sn]
        if ws.sheet_state != 'visible':
            continue  # হাইড/ভেরি-হাইড শীট — ডাটা নেওয়া হবে না

        header_row = _find_header_row_v2(ws)
        is_v2 = header_row is not None
        if header_row is None:
            header_row = _find_header_row_v1(ws)
        if header_row is None:
            continue  # এই শিট কোনো পরিচিত ফরম্যাটের না

        col_map = _build_col_map(ws, header_row)
        required = ('po_no', 'pack_type', 'qty', 'measurement')
        if is_v2:
            required = required + ('style_no',)
        if not all(k in col_map for k in required):
            continue  # প্রত্যাশিত কলাম পাওয়া যায়নি — এই ফরম্যাট না

        hidden_rows = {r for r, dim in ws.row_dimensions.items() if dim.hidden}

        global_style_no = '' if is_v2 else (_extract_style_no(ws) or 'N/A')
        po_col = col_map['po_no']
        pack_col = col_map['pack_type']
        qty_col = col_map['qty']
        meas_col = col_map['measurement']
        style_col = col_map.get('style_no')
        ref_col = col_map.get('reference')

        r = header_row + 1
        max_row = ws.max_row
        last_measurement = ''
        consecutive_blank = 0
        BLANK_STOP_THRESHOLD = 100
        while r <= max_row:
            po_val = ws.cell(row=r, column=po_col).value
            if po_val is None or not _clean(po_val):
                consecutive_blank += 1
                if consecutive_blank >= BLANK_STOP_THRESHOLD:
                    break
                r += 1
                continue  # Order Number ফাঁকা — Total রো বা খালি রো, স্কিপ
            consecutive_blank = 0

            # measurement merged-cell-এর 'আসল' ভ্যালু কখনো কখনো একটা হাইড
            # রো-তেই থাকে (নিচের ভিজিবল রোগুলো শুধু সেটা ভিজ্যুয়ালি
            # ইনহেরিট করে দেখায়) — তাই forward-fill ক্যাপচার করা হয় হাইড/
            # ভিজিবল নির্বিশেষে সব রো থেকেই, hidden-row-skip চেকের আগেই।
            meas_raw = ws.cell(row=r, column=meas_col).value
            if meas_raw is not None and _clean(meas_raw):
                last_measurement = str(meas_raw)

            if r in hidden_rows:
                r += 1
                continue  # হাইড রো — এই রো নিজে কোনো লাইন-আইটেম হবে না

            qty_val = ws.cell(row=r, column=qty_col).value
            if not _is_num(qty_val):
                r += 1
                continue

            length, width, height = _parse_measurement(last_measurement)
            if not length:
                r += 1
                continue

            style_no = _clean(ws.cell(row=r, column=style_col).value) if style_col else global_style_no
            reference_val = _clean(ws.cell(row=r, column=ref_col).value) if ref_col else ''

            all_items.append({
                'item_name': default_item_name,
                'ewo_no': 'N/A',
                'style_no': style_no or 'N/A',
                'po_no': _po_str(po_val),
                'length': length,
                'width': width,
                'height': height,
                'ply': ply_value,
                'qty': round(float(qty_val)),
                'pack_type': _clean(ws.cell(row=r, column=pack_col).value),
                'reference': reference_val,
                'color': '',
                'size': '',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': sn,
                '_source_file': filename,
            })
            r += 1

    return all_items