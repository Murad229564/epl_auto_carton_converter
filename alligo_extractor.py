import re
from openpyxl import load_workbook

# ---------------------------------------------------------------------------
# Green Life Knit Composite Ltd. — Buyer: ALLIGO
# 'Revised Purchase Order' / Carton Booking এক্সেল ফরম্যাট।
#   - হেডার রো একটাই (S/L #, Buyer Name, ITEM, Style #, Order No., Artcle
#     No., [Color — অপশনাল, কিছু ফাইলে থাকে না], Garments Qty., Item,
#     Measurement, O/Qty, Unit Price $, Total Price $, Remarks)। 'Color'
#     কলাম থাকলে বাকি সব কলাম একঘর ডানে শিফট হয়ে যায় — তাই fixed index
#     ধরা হয় না, সব হেডার-লেবেল স্ক্যান করে ডাইনামিকভাবে বের করা হয়।
#   - সতর্কতা: হেডারে 'ITEM' নামের কলাম দুইবার আসে — একবার Style#-এর আগে
#     (গার্মেন্টস ক্যাটেগরি, যেমন 'TEE SHIRT' — এটা দরকার নেই), আরেকবার
#     Measurement-এর ঠিক আগে (আসল Item Name — 'ITEM'-এর শেষ/ডানদিকের
#     ওকারেন্সটাই আসল Item Name কলাম)।
#   - প্রতিটা S/L# ব্লকে ৩টা পর্যন্ত সাব-রো থাকে:
#       ১) মূল রো — Style#/Order No./Article No./Color সহ, Item =
#          '5 Ply CTN Best Quality' (বা কাছাকাছি বানান) -> Master Carton,
#          Measurement 'L X W X H CM' (৩ সংখ্যা), Ply = 5
#       ২) 'Top & Bottom' রো — Style/Order/Article/Color ফাঁকা, Measurement
#          শুধু 'L X W CM' (২ সংখ্যা, Height নেই), Ply = 3
#       ৩) 'Divider' রো — একই রকম, Ply = 3
#   - ইউজার-কনফার্মড বিজনেস রুল: Top & Bottom আর Divider ব্রেকডাউন হয় না —
#     পুরো ফাইল(গুলো) জুড়ে measurement (L x W) অনুযায়ী গ্রুপ করে Qty যোগ
#     করে একটাই সামারি-লাইন বসে (একই measurement একাধিক ব্লকে থাকলেও)।
#     Style/PO/Article/Pack Type — Top & Bottom/Divider-এর ক্ষেত্রে
#     সবসময় 'N/A' (এসব আইটেমের জন্য প্রযোজ্য না)।
#   - Master Carton-এর Pack Type: 'Color' কলাম থাকলে সেই ভ্যালু, না থাকলে
#     'N/A'।
#   - আউটপুট অর্ডার: সব ফাইলের সব Master Carton আগে (মূল ক্রমে), তারপর
#     সব Top Bottom সামারি-লাইন (measurement-ওয়াইজ), তারপর সব Divider
#     সামারি-লাইন — Sterling/Amigo কনভেনশনের সাথে সামঞ্জস্যপূর্ণ।
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


_MEAS_RE = re.compile(r'(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)(?:\s*[xX×]\s*(\d+\.?\d*))?')


def _parse_measurement(text):
    """'58 X 39 X 12 CM' -> ('58','39','12')। '55 X 36 CM' (Height ছাড়া) ->
    ('55','36','')।"""
    if not text:
        return '', '', ''
    m = _MEAS_RE.search(str(text))
    if not m:
        return '', '', ''
    l = _fmt_num(m.group(1))
    w = _fmt_num(m.group(2))
    h = _fmt_num(m.group(3)) if m.group(3) else ''
    return l, w, h


def _po_str(v):
    if isinstance(v, (int, float)):
        return str(int(v)) if float(v).is_integer() else str(v)
    return _clean(v)


def _find_header_row(ws, max_scan=30):
    """কলাম A-তে 'S/L #' আর কলাম B-তে 'Buyer Name' — এই দুটো একসাথে থাকা
    রো-কেই হেডার ধরা হয়।"""
    for r in range(1, max_scan + 1):
        a = _norm(ws.cell(row=r, column=1).value)
        b = _norm(ws.cell(row=r, column=2).value)
        if a == 'sl' and b.startswith('buyername'):
            return r
    return None


def _build_col_map(ws, header_row):
    """হেডার-লেবেল স্ক্যান করে কলাম-পজিশন বের করে। 'Item' লেবেল দুইবার
    আসতে পারে (ক্যাটেগরি + আসল Item Name) — সবচেয়ে ডানদিকের (max index)
    ওকারেন্সটাই আসল Item Name কলাম হিসেবে নেওয়া হয়।"""
    col_map = {}
    item_cols = []
    for c in range(1, ws.max_column + 1):
        label = _norm(ws.cell(row=header_row, column=c).value)
        if not label:
            continue
        if label == 'style':
            col_map['style_no'] = c
        elif label == 'orderno':
            col_map['po_no'] = c
        elif label in ('artcleno', 'articleno'):
            col_map['reference'] = c
        elif label == 'color':
            col_map['pack_type'] = c
        elif label == 'measurement':
            col_map['measurement'] = c
        elif label == 'oqty':
            col_map['qty'] = c
        elif label == 'item':
            item_cols.append(c)
    if item_cols:
        col_map['item_name'] = max(item_cols)  # ডানদিকেরটাই আসল Item Name
    return col_map


def _classify_item(item_text):
    """রিটার্ন করে (category, ply) — category: 'master'/'topbottom'/
    'divider'/None (অচেনা হলে)।"""
    n = _norm(item_text)
    if 'divider' in n:
        return 'divider', '3'
    if 'top' in n and 'bottom' in n:
        return 'topbottom', '3'
    if 'carton' in n or 'ply' in n:
        return 'master', '5'
    return None, ''


def read_alligo_booking_file(file_stream, filename=''):
    """একটা ফাইলের সব (visible) শিট থেকে ডাটা বের করে। রিটার্ন করে
    (master_items, topbottom_by_measurement, divider_by_measurement,
    warnings) — measurement-ওয়াইজ dict-গুলো {(L,W): summed_qty} আকারে,
    যাতে caller সব ফাইল/শিট জুড়ে এই dict-গুলো merge করতে পারে।"""
    wb = load_workbook(file_stream, data_only=True)
    master_items = []
    topbottom_sums = {}
    divider_sums = {}
    warnings = []

    for sn in wb.sheetnames:
        ws = wb[sn]
        header_row = _find_header_row(ws)
        if header_row is None:
            continue

        col_map = _build_col_map(ws, header_row)
        required = ('style_no', 'po_no', 'reference', 'measurement', 'qty', 'item_name')
        if not all(k in col_map for k in required):
            continue  # প্রত্যাশিত কলাম পাওয়া যায়নি — এই ফরম্যাট না

        item_col = col_map['item_name']
        meas_col = col_map['measurement']
        qty_col = col_map['qty']
        style_col = col_map['style_no']
        po_col = col_map['po_no']
        ref_col = col_map['reference']
        pack_col = col_map.get('pack_type')

        r = header_row + 1
        max_row = ws.max_row
        while r <= max_row:
            row_b = _clean(ws.cell(row=r, column=2).value)
            if _norm(row_b).startswith('gtotal'):
                break  # টেবিলের শেষ — G TOTAL রো-এর পর থেকে সব ফুটার/স্বাক্ষর সেকশন

            item_text = _clean(ws.cell(row=r, column=item_col).value)
            if not item_text:
                r += 1
                continue

            category, ply = _classify_item(item_text)
            if category is None:
                warnings.append(f"⚠️ শিট '{sn}' রো {r}: অচেনা Item টাইপ '{item_text}' — স্কিপ করা হয়েছে।")
                r += 1
                continue

            qty_val = ws.cell(row=r, column=qty_col).value
            if not _is_num(qty_val):
                r += 1
                continue
            qty = round(float(qty_val))
            if qty <= 0:
                r += 1
                continue

            length, width, height = _parse_measurement(ws.cell(row=r, column=meas_col).value)
            if not length or not width:
                r += 1
                continue

            if category == 'master':
                master_items.append({
                    'item_name': 'Master Carton',
                    'ewo_no': 'N/A',
                    'style_no': _clean(ws.cell(row=r, column=style_col).value) or 'N/A',
                    'po_no': _po_str(ws.cell(row=r, column=po_col).value) or 'N/A',
                    'length': length,
                    'width': width,
                    'height': height,
                    'ply': '5',
                    'qty': qty,
                    'pack_type': (_clean(ws.cell(row=r, column=pack_col).value) or 'N/A') if pack_col else 'N/A',
                    'reference': _clean(ws.cell(row=r, column=ref_col).value) or 'N/A',
                    'color': '',
                    'size': '',
                    'delivery_date': '',
                    'measurement_unit': 'Cm',
                    'delivery_place_pdf': '',
                    'delivery_address_pdf': '',
                    '_sheet': sn,
                    '_source_file': filename,
                })
            else:
                key = (length, width)
                target = topbottom_sums if category == 'topbottom' else divider_sums
                target[key] = target.get(key, 0) + qty

            r += 1

    return master_items, topbottom_sums, divider_sums, warnings


def combine_alligo_booking_files(files, item_name_override='', manual_ply=''):
    """files: [(BytesIO, filename), ...] — BATCH_REGISTRY-এর uniform কল-
    সিগনেচার। item_name_override/manual_ply ব্যবহার হয় না (Item Name/Ply
    সম্পূর্ণ ফাইলের নিজস্ব Item কলাম থেকেই ফিক্সড নিয়মে ঠিক হয়)।

    ইউজার-কনফার্মড অর্ডারিং: সব ফাইলের সব Master Carton আগে (মূল ক্রমে),
    তারপর সব ফাইল/শিট জুড়ে measurement-ওয়াইজ যোগ করা এক-একটা Top Bottom
    সামারি-লাইন, তারপর একইভাবে Divider সামারি-লাইন।
    """
    all_master = []
    all_warnings = []
    topbottom_sums = {}
    divider_sums = {}
    # ক্রম ঠিক রাখতে (প্রথম যেই measurement আসবে সেটাই আগে বসবে) আলাদা
    # লিস্টে insertion-order ট্র্যাক করা হচ্ছে
    topbottom_order = []
    divider_order = []

    for file_stream, filename in files:
        master_items, tb_sums, div_sums, warns = read_alligo_booking_file(file_stream, filename)
        all_master.extend(master_items)
        all_warnings.extend(warns)
        for key, qty in tb_sums.items():
            if key not in topbottom_sums:
                topbottom_order.append(key)
            topbottom_sums[key] = topbottom_sums.get(key, 0) + qty
        for key, qty in div_sums.items():
            if key not in divider_sums:
                divider_order.append(key)
            divider_sums[key] = divider_sums.get(key, 0) + qty

    combined = list(all_master)
    for key in topbottom_order:
        length, width = key
        combined.append({
            'item_name': 'Top Bottom',
            'ewo_no': 'N/A',
            'style_no': 'N/A',
            'po_no': 'N/A',
            'length': length,
            'width': width,
            'height': '',
            'ply': '3',
            'qty': topbottom_sums[key],
            'pack_type': 'N/A',
            'reference': 'N/A',
            'color': '',
            'size': '',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
        })
    for key in divider_order:
        length, width = key
        combined.append({
            'item_name': 'Divider',
            'ewo_no': 'N/A',
            'style_no': 'N/A',
            'po_no': 'N/A',
            'length': length,
            'width': width,
            'height': '',
            'ply': '3',
            'qty': divider_sums[key],
            'pack_type': 'N/A',
            'reference': 'N/A',
            'color': '',
            'size': '',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
        })

    return combined, all_warnings
