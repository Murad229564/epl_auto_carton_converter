"""
Dhaka Garments And Washing Ltd. — Buyer: Kohl`s
Carton Booking Excel ফরম্যাট।

একটা ফাইলে একাধিক শিট থাকতে পারে (BULK/E-com/Pre-pack ইত্যাদি নামে,
কখনো একই নামের একাধিক কপি যেমন 'Bulk (2)', 'E-com (2)') — প্রতিটা শিট
থেকেই ডাটা নেওয়া হয়।

প্রতিটা শিটের গঠন:
- উপরের দিকে (rows ~1-24) generic ছাপা instruction টেক্সট — সব শিটে প্রায়
  একই রকম, ডাটার সাথে সম্পর্কিত না, উপেক্ষা করা হয়।
- একটা 'PO' মার্কার-রো (কলাম B/col2-এ ঠিক 'PO' লেখা) — এই রো-তেই PO নম্বর
  (তার ঠিক পাশের কলামে) আর Item Name-এর ইঙ্গিত টেক্সট থাকে ('Normal
  carton' / 'Elastic  Hanger Carton') — এই টেক্সটের কলাম-পজিশন শিট-ভেদে
  বদলায় (কখনো col5, কখনো col11), তাই পুরো রো স্ক্যান করে বের করা হয়।
- তার ঠিক পরের রো-তে আসল হেডার: STYLE / UPC NUMBER / COLOR / সাইজ-কলামগুলো
  (সংখ্যা ও অবস্থান শিট-ভেদে ভিন্ন) / CTN MEASUREMENT / CARTON BOOKING QTY
  / তারপর 'Do not follow' মার্ক করা কলামগুলো (GMT/QTY, ACTUAL, QTY/x%,
  RATIO) — এগুলো ব্যবহার হয় না।
- ডাটা রো-গুলোয় প্রতিটা Style+UPC+Color+Size কম্বিনেশনের জন্য একটা করে
  লাইন, CARTON BOOKING QTY কলামেই আসল কোয়ান্টিটি।
- শেষে 'G.TTL' রো (col 'CTN MEASUREMENT'-এর কলামেই লেখা থাকে, CARTON
  BOOKING QTY কলামে টোটাল) — ডাটা-শেষের মার্কার, cross-check-এর জন্যও
  ব্যবহার করা হচ্ছে।

ব্যবসায়িক নিয়ম (ইউজার-কনফার্মড):
- Item Name: PO-মার্কার-রো-তে টেক্সট স্ক্যান করে —
    'elastic' থাকলে -> 'Elastic Hanger Carton'
    'normal' থাকলে -> 'Master Carton'
  কিছু না পাওয়া গেলে শিটের নাম দেখে (trailing '(2)'/'(3)' বাদ দিয়ে): নাম
  'E-com'-এর সাথে মিললে Master Carton, নাহলে Elastic Hanger Carton।
- Ply: UI থেকে সিলেক্ট করা মান থাকলে সেটাই, না থাকলে ডিফল্ট 5।
- Style <- 'STYLE' কলাম। GMT PO <- PO-মার্কার-রো-এর PO নম্বর (পুরো শিটের
  জন্য একটাই, প্রতি রো-তে বসবে)। Reference <- 'COLOR' কলাম। Pack Type <-
  'UPC NUMBER' কলাম। Measurement <- 'CTN MEASUREMENT' কলাম (ফরম্যাট
  ফাইল-ভেদে বদলাতে পারে — L:/W:/H: প্রিফিক্স-সহ বা ছাড়া, দুটোই সামলানো
  হয়েছে)। Qty <- 'CARTON BOOKING QTY' কলাম (GMT/QTY, ACTUAL, RATIO ইত্যাদি
  'Do not follow' কলাম ব্যবহার হয় না)।
- Qty 0/ফাঁকা হলে সেই রো বাদ (আগের সব কাস্টমারের কনভেনশন অনুযায়ী)।
"""
import re
import pandas as pd


def _norm(v):
    return re.sub(r'[^a-z0-9]', '', str(v or '').lower())


def _clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def _is_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _fmt_num(v):
    if v is None:
        return ''
    try:
        f = float(v)
    except (TypeError, ValueError):
        return ''
    return str(int(f)) if f == int(f) else str(round(f, 3))


# 'L: 19.5 X W: 14.5 X H: 3 INCH' বা 'L:19.5 X W:14.5 X H:3 CM' বা প্রিফিক্স
# ছাড়া শুধু '19.5 X 14.5 X 3' — যেকোনো ফরম্যাটই সামলাতে শুধু সংখ্যাগুলোই
# ক্রম অনুযায়ী নেওয়া হচ্ছে; ইউনিট (INCH/CM) টেক্সট থেকে আলাদাভাবে বের করা।
def _parse_measurement(text):
    t = _clean(text)
    nums = re.findall(r'(\d+\.?\d*)', t)
    l = _fmt_num(nums[0]) if len(nums) > 0 else ''
    w = _fmt_num(nums[1]) if len(nums) > 1 else ''
    h = _fmt_num(nums[2]) if len(nums) > 2 else ''
    unit = 'Inch' if 'inch' in t.lower() else 'Cm'
    return l, w, h, unit


def _row_label_map(row):
    labels = {}
    for c, v in enumerate(row):
        t = _clean(v)
        if t:
            labels[_norm(t)] = c
    return labels


def _find_col(labels, *must_contain):
    for key, col in labels.items():
        if all(s in key for s in must_contain):
            return col
    return None


def _classify_item_name(po_row, sheet_name):
    """PO-মার্কার-রো-এর টেক্সট স্ক্যান করে Item Name ঠিক করে; না পেলে
    শিটের নামের ওপর ভিত্তি করে ('E-com' -> Master Carton, নাহলে Elastic
    Hanger Carton)।"""
    for v in po_row:
        t = _clean(v).lower()
        if not t:
            continue
        if 'elastic' in t:
            return 'Elastic Hanger Carton'
        if 'normal' in t:
            return 'Master Carton'

    base_name = re.sub(r'\s*\(\d+\)\s*$', '', sheet_name or '').strip().lower()
    if base_name.startswith('e-com') or base_name.startswith('ecom'):
        return 'Master Carton'
    return 'Elastic Hanger Carton'


def _find_po_marker_rows(rows):
    """কলাম B (0-indexed col1)-এ ঠিক 'PO' লেখা এমন সব রো খুঁজে বের করে —
    একটা শিটে একাধিক PO-ব্লকও থাকতে পারে (এই ফাইলে একটাই দেখা গেছে, তবে
    জেনারেল রাখা হচ্ছে)।"""
    marker_rows = []
    for i, row in enumerate(rows):
        if len(row) > 1 and _norm(row[1]) == 'po':
            marker_rows.append(i)
    return marker_rows


def read_dhaka_booking_sheet(rows, sheet_name, item_name_override='', manual_ply=''):
    """একটা শিট থেকে লাইন-আইটেম বের করে। রিটার্ন করে (line_items, warnings)।"""
    marker_rows = _find_po_marker_rows(rows)
    if not marker_rows:
        return [], []

    items = []
    warnings = []
    n_rows = len(rows)

    for bi, marker_row in enumerate(marker_rows):
        po_row = rows[marker_row]
        po_no = _clean(po_row[2]) if len(po_row) > 2 else ''
        item_name = item_name_override or _classify_item_name(po_row, sheet_name)
        ply = manual_ply.strip() if manual_ply else '5'

        header_row_idx = marker_row + 1
        if header_row_idx >= n_rows:
            continue
        labels = _row_label_map(rows[header_row_idx])

        style_col = labels.get('style')
        upc_col = _find_col(labels, 'upcnumber')
        color_col = labels.get('color')
        meas_col = _find_col(labels, 'ctnmeasurement')
        qty_col = _find_col(labels, 'cartonbookingqty')

        if style_col is None or meas_col is None or qty_col is None:
            warnings.append(
                f"⚠️ শিট '{sheet_name}' (PO ব্লক {bi + 1}): Style/CTN Measurement/"
                f"Carton Booking Qty কলাম পাওয়া যায়নি — এই ব্লক স্কিপ করা হয়েছে।"
            )
            continue

        end_row = marker_rows[bi + 1] if bi + 1 < len(marker_rows) else n_rows
        total_row_qty = None
        extracted_total = 0.0
        row_count = 0

        r = header_row_idx + 1
        while r < end_row:
            row = rows[r]
            style_val = _clean(row[style_col]) if style_col < len(row) else ''
            meas_val = _clean(row[meas_col]) if meas_col < len(row) else ''

            if _norm(meas_val) == 'gttl' or _norm(style_val) == 'gttl':
                qty_val = row[qty_col] if qty_col < len(row) else None
                if _is_num(qty_val):
                    total_row_qty = _num(qty_val)
                r += 1
                continue

            if not style_val:
                r += 1
                continue

            qty_val = row[qty_col] if qty_col < len(row) else None
            if not _is_num(qty_val) or (_num(qty_val) or 0) <= 0:
                r += 1
                continue  # qty 0/ফাঁকা — বাদ

            length, width, height, unit = _parse_measurement(meas_val)
            upc_val = _clean(row[upc_col]) if upc_col is not None and upc_col < len(row) else ''
            color_val = _clean(row[color_col]) if color_col is not None and color_col < len(row) else ''
            # ইউজার-কনফার্মড: Qty ফ্র্যাকশনে থাকলে (এই ফরম্যাটে প্রায়ই থাকে,
            # যেমন 6.15, 18.34) রাউন্ড করে বসবে।
            qty_num = round(_num(qty_val))
            extracted_total += qty_num
            row_count += 1

            items.append({
                'item_name': item_name,
                'ewo_no': 'N/A',
                'style_no': style_val,
                'po_no': po_no or 'N/A',
                'length': length,
                'width': width,
                'height': height,
                'ply': ply,
                'qty': qty_num,
                'pack_type': upc_val or 'N/A',
                'reference': color_val or 'N/A',
                'remarks': '',
                'color': 'N/A',
                'size': 'N/A',
                'delivery_date': '',
                'measurement_unit': unit,
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': sheet_name,
            })
            r += 1

        # প্রতিটা রো আলাদাভাবে রাউন্ড হওয়ায় G.TTL (যেটা নিজেই ফ্র্যাকশনাল)
        # -এর সাথে সামান্য পার্থক্য (প্রতি রো-তে সর্বোচ্চ ০.৫ পর্যন্ত) স্বাভাবিক —
        # তাই tolerance রো-সংখ্যা অনুযায়ী বাড়ানো হয়েছে, যাতে false-positive
        # warning না আসে।
        if total_row_qty is not None and abs(total_row_qty - extracted_total) > max(0.51, 0.51 * row_count):
            warnings.append(
                f"⚠️ শিট '{sheet_name}' (PO {po_no}): G.TTL রো-তে Carton Booking Qty মোট "
                f"{total_row_qty:g}, কিন্তু বের করা (রাউন্ড করা) লাইন-আইটেমগুলোর মোট Qty "
                f"{extracted_total:g} — পার্থক্য {total_row_qty - extracted_total:g} (রাউন্ডিং-এর "
                f"চেয়ে বেশি), ভালোভাবে চেক করে নিন।"
            )

    return items, warnings


def read_dhaka_booking_file(file_stream, filename, item_name_override='', manual_ply=''):
    """একটা .xls/.xlsx ফাইলের সব শিট থেকে ডাটা বের করে। রিটার্ন করে
    (line_items, warnings)।"""
    file_stream.seek(0)
    sheets = pd.read_excel(file_stream, sheet_name=None, header=None)

    all_items = []
    all_warnings = []
    for sheet_name, df in sheets.items():
        items, warns = read_dhaka_booking_sheet(
            df.values.tolist(), sheet_name,
            item_name_override=item_name_override, manual_ply=manual_ply)
        all_items.extend(items)
        all_warnings.extend(warns)

    if not all_items:
        all_warnings.append(f"⚠️ '{filename}': কোনো ভ্যালিড (qty>0) লাইন-আইটেম পাওয়া যায়নি।")

    return all_items, all_warnings


def combine_dhaka_booking_files(files, item_name_override='', manual_ply=''):
    """files: [(BytesIO, filename), ...] — BATCH_REGISTRY-এর uniform কল-
    সিগনেচার। item_name_override সাধারণত ব্যবহার হয় না (Item Name সম্পূর্ণ
    ফাইলের কনটেন্ট/শিট-নাম থেকেই ডিটেক্ট হয়), কিন্তু ভবিষ্যতে UI থেকে জোর
    করে override দিতে চাইলে সাপোর্ট করা আছে। manual_ply UI থেকে এলে
    ব্যবহার হয়, না দিলে ডিফল্ট 5। রিটার্ন করে (line_items, warnings)।"""
    combined = []
    all_warnings = []
    for file_stream, filename in files:
        items, warns = read_dhaka_booking_file(
            file_stream, filename,
            item_name_override=item_name_override, manual_ply=manual_ply)
        for it in items:
            it['_source_file'] = filename
        combined.extend(items)
        all_warnings.extend(warns)
    return combined, all_warnings
