"""
Innovative Knitex Ltd. — Buyer: Biscana
'Carton Order Sheet Summery' PDF ফরম্যাট।

Barnali-স্টাইল Trims Booking PDF-এর মতোই আলাদা extractor+route দরকার
(Kenpark-এর কনভেনশন অনুযায়ী) — এই ফরম্যাট সম্পূর্ণ ভিন্ন, PO/Article/
Order Code উপরে, তারপর Size x Color-এর একটা ব্রেকডাউন টেবিল (measurement
প্রতিটা সাইজ-কলামের জন্য আলাদা, quantity প্রতিটা color x size সেলে)।

pdfplumber-এর extract_tables() এই PDF-এ পাতাভিত্তিক টেবিল দেয় (merged
header সেলগুলো None-ফিল হয়ে থাকে)।

মাল্টি-পেজ সাপোর্ট (নতুন): কালার-রো বেশি হলে টেবিলটা পরের পাতায় চলে যায়।
পরের পাতার টেবিলে টাইটেল/হেডার/সাইজ রো থাকে না — শুধু কালার-রো (এবং
সবশেষে TOTAL রো) থাকে। তাই এখন:
  * টাইটেল সহ প্রথম টেবিল থেকে PO/Article/Order Code/Size/Measurement নেওয়া হয়
  * তারপরের সব পাতার টেবিলের রো-গুলো একই কলাম-ম্যাপিং দিয়ে ক্রমানুসারে
    জোড়া লাগানো হয় (২, ৩, ৪ ... যত পাতাই হোক)
  * কোনো পাতায় আবার 'SIZE' হেডার রো রিপিট হলে সেটা থেকে সাইজ/measurement
    নতুন করে নেওয়া হয় (পাতা-ভেদে measurement বদলালেও ঠিক থাকবে)
  * কালার-কলাম ফাঁকা + সংখ্যা আছে এমন রো (= TOTAL রো) পেলেই থামে

ব্যবসায়িক নিয়ম (ইউজার-কনফার্মড):
- PO Number (টেমপ্লেটের একদম উপরে, শুধু UI-তে PO ফাঁকা রাখলে auto-fill
  হয়): 'IKL-BISCANA-24-2026' + '/' + Article Name, যেমন
  'IKL-BISCANA-24-2026/BERLIN'। এই কম্বিনেশন-ফরম্যাট শুধু header_info-তেই
  (টপ-লেভেল PO override), প্রতিটা লাইন-আইটেমের নিজস্ব po_no ফিল্ডে না।
- GMT PO (প্রতিটা লাইন-আইটেমে) <- শুধু PO নম্বরটাই (Article suffix ছাড়া),
  যেমন 'IKL-BISCANA-24-2026'।
- GMT Style <- Article Name (যেমন 'BERLIN')।
- Reference <- ORDER CODE (যেমন '373-IK-OP-2641'), সব রো-তে একই।
- Pack Type <- Color-এর সম্পূর্ণ লেবেল (দুই লাইনের টেক্সট, যেমন
  'BRANCO WHITE')।
- GMT Size <- সাইজের লেবেল (যেমন 'XS', 'S', 'M' ...)।
- Measurement <- সাইজ-কলামের নিচে থাকা measurement টেক্সট (যেমন
  '43X36X27 CM'), সেই একই সাইজ-কলামের Quantity সেলের সাথে মেলানো।
- Qty <- প্রতিটা color x size সেল থেকে, 0/ফাঁকা হলে সেই সেল বাদ।
- Item Name/Ply — UI থেকে যা সিলেক্ট করা হয় তাই বসে (item_name_override/
  manual_ply); ডিফল্ট Master Carton/5।
"""
import re
import pdfplumber


def _norm(v):
    return re.sub(r'[^a-z0-9]', '', str(v or '').lower())


def _clean(v):
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def _fmt_num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return ''
    return str(int(f)) if f == int(f) else str(round(f, 3))


def _parse_lwh(text):
    """'43X36X27 CM' -> ('43','36','27')।"""
    t = _clean(text)
    nums = re.findall(r'(\d+\.?\d*)', t)
    l = _fmt_num(nums[0]) if len(nums) > 0 else ''
    w = _fmt_num(nums[1]) if len(nums) > 1 else ''
    h = _fmt_num(nums[2]) if len(nums) > 2 else ''
    return l, w, h


def _to_num(v):
    """সেলের মান সংখ্যা হলে float, নাহলে None (ফাঁকা/টেক্সট -> None)।"""
    s = _clean(v).replace(',', '')
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _find_row(table, predicate, start=0):
    for i in range(start, len(table)):
        if predicate(table[i]):
            return i
    return None


def _find_col(row, *keywords):
    for c, v in enumerate(row):
        n = _norm(v)
        if n and all(k in n for k in keywords):
            return c
    return None


def _first_nonblank_row(table, start):
    for i in range(start, len(table)):
        row = table[i]
        if any(_clean(c) for c in row):
            return i
    return None


def _is_title_row(r):
    return 'cartonordersheetsummery' in _norm(r[0] if r else '')


def _collect_tables(pdf):
    """সব পাতা ক্রমানুসারে স্ক্যান করে রিটার্ন করে:
      main  = (টাইটেল-সহ প্রথম টেবিল, টাইটেল রো-র ইনডেক্স)  — না পেলে None
      conts = মেইন টেবিলের পরের সব পাতার (টাইটেল ছাড়া) টেবিল, পাতা-ক্রমে।
    """
    main = None
    conts = []
    for page in pdf.pages:
        for table in page.extract_tables():
            if not table:
                continue
            if main is None:
                title_row = _find_row(table, _is_title_row)
                if title_row is not None:
                    main = (table, title_row)
            else:
                # আরেকটা নতুন টাইটেল থাকলে সেটা আলাদা শিট — continuation না
                if _find_row(table, _is_title_row) is None:
                    conts.append(table)
    return main, conts


def _parse_size_cols(size_row_cells):
    cols = []  # [(col_idx, size_label), ...]
    for c in range(1, len(size_row_cells)):
        label = _clean(size_row_cells[c])
        if not label or _norm(label) == 'total':
            break
        cols.append((c, label))
    return cols


def _read_measurements(row, size_cols):
    m = {}
    for c, _label in size_cols:
        m[c] = _clean(row[c]) if c < len(row) else ''
    return m


def read_ikl_biscana_pdf(file_stream, filename='', item_name_override='', manual_ply=''):
    """মূল entry point। রিটার্ন করে (header_info, line_items)। header_info-তে
    'po_number'/'customer'/'buyer' থাকে (customer/buyer এই ফরম্যাটে PDF থেকে
    বের করা যায় না, তাই ফাঁকা — UI-এর সিলেকশনই ব্যবহার হবে)। এই ফরম্যাট না
    হলে (টেবিল/টাইটেল না মিললে) header_info-তে po_number ফাঁকা আর
    line_items=[] রিটার্ন করে।"""
    file_stream.seek(0)
    with pdfplumber.open(file_stream) as pdf:
        main, conts = _collect_tables(pdf)

    empty = ({'po_number': '', 'customer': '', 'buyer': ''}, [])
    if main is None:
        return empty
    table, title_row = main

    po_row = _first_nonblank_row(table, title_row + 1)
    po_number = _clean(table[po_row][0]) if po_row is not None else ''

    header_row = _find_row(table, lambda r: any('articlename' in _norm(c) for c in r))
    article_name = order_code = ''
    if header_row is not None:
        article_col = _find_col(table[header_row], 'articlename')
        code_col = _find_col(table[header_row], 'ordercode')
        data_row = _first_nonblank_row(table, header_row + 1)
        if data_row is not None:
            if article_col is not None and article_col < len(table[data_row]):
                article_name = _clean(table[data_row][article_col])
            if code_col is not None and code_col < len(table[data_row]):
                order_code = _clean(table[data_row][code_col])

    size_row = _find_row(table, lambda r: _norm(r[0] if r else '') == 'size')
    if size_row is None:
        return empty

    size_cols = _parse_size_cols(table[size_row])
    meas_row = size_row + 1  # COLOUR + measurement রো, সাইজ রো-এর ঠিক পরেই
    measurements = _read_measurements(table[meas_row], size_cols) if meas_row < len(table) else {}

    header_info = {
        'po_number': f"{po_number}/{article_name}" if po_number and article_name else (po_number or ''),
        'customer': '',
        'buyer': '',
    }

    # সব পাতার ডাটা-রো একটা ফ্ল্যাট লিস্টে: মেইন টেবিলের measurement-রো-র পরের
    # রো-গুলো, তারপর পরের পাতাগুলোর সব রো
    flat_rows = list(table[meas_row + 1:])
    for t in conts:
        flat_rows.extend(t)

    line_items = []
    expect_meas = False
    for row in flat_rows:
        if not row or not any(_clean(c) for c in row):
            continue  # সম্পূর্ণ ফাঁকা রো

        first = _clean(row[0])
        first_norm = _norm(first)

        # পরের পাতায় হেডার রিপিট হলে — সাইজ/measurement নতুন করে নাও
        if first_norm == 'size':
            new_cols = _parse_size_cols(row)
            if new_cols:
                size_cols = new_cols
                measurements = {}
                expect_meas = True
            continue
        if expect_meas and first_norm in ('colour', 'color'):
            measurements = _read_measurements(row, size_cols)
            expect_meas = False
            continue
        if first_norm in ('colour', 'color'):
            continue

        qty_cells = [_to_num(row[c]) if c < len(row) else None for c, _ in size_cols]

        if not first:
            # কালার-কলাম ফাঁকা + সংখ্যা আছে = TOTAL রো — সব পাতা শেষ, থামুন
            if any(q is not None for q in qty_cells):
                break
            continue

        if all(q is None for q in qty_cells):
            continue  # নোট/টেক্সট রো — কোনো সাইজ-কোয়ান্টিটি নেই

        color_label = first.replace('\n', ' ')
        for (c, size_label), qty_num in zip(size_cols, qty_cells):
            if qty_num is None or qty_num <= 0:
                continue  # ইউজার-কনফার্মড: 0/ফাঁকা কোয়ান্টিটি বাদ

            length, width, height = _parse_lwh(measurements.get(c, ''))
            line_items.append({
                'item_name': item_name_override or 'Master Carton',
                'ewo_no': 'N/A',
                'style_no': article_name or 'N/A',
                'po_no': po_number or 'N/A',
                'length': length,
                'width': width,
                'height': height,
                'ply': manual_ply.strip() if manual_ply else '5',
                'qty': int(qty_num) if qty_num == int(qty_num) else qty_num,
                'pack_type': color_label or 'N/A',
                'reference': order_code or 'N/A',
                'remarks': '',
                'color': 'N/A',
                'size': size_label or 'N/A',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
            })

    return header_info, line_items