"""
Sinha Knit and Denims Limited — Buyer: Tata Trent
Carton Booking Excel ফরম্যাট।

একটা ফাইলে একাধিক শিট থাকতে পারে (প্রতিটা সাধারণত এক-একটা স্টাইল/কালার
কম্বিনেশনের জন্য) — প্রতিটা শিট থেকেই ডাটা নেওয়া হয়। প্রতিটা শিটের গঠন:

- উপরে একটা মেইন টেবিল (হেডার: STYLE / COLOR/WASH / SIZE / PO NO /
  ARTICLE NO / GARMENTS QNTY / LENGTH / WIDTH / HEIGHT / BOOKING QNTY) —
  একেকটা PO-ব্লকে ৬টা সাইজ-রো থাকে (STYLE কলাম শুধু ব্লকের প্রথম রো-তে
  থাকে, ফরওয়ার্ড-ফিল দরকার)। Booking Qty 0/ফাঁকা হলে সেই রো বাদ।
- টেবিলের শেষে (মেইন ডাটার পরে) একটা "TOP BOTTOM ... CM = <qty>" বা
  "... DIVIDER ... CM = <qty>" জাতীয় লাইন থাকে — এটা আলাদা লাইন-আইটেম
  হিসেবে সবার শেষে (Amigo Bangladesh-এর কনভেনশন অনুযায়ী) বসবে, বাকি সব
  ফিল্ড N/A।

ব্যবসায়িক নিয়ম (ইউজার-কনফার্মড):
- Item Name ও Ply — UI থেকে যা সিলেক্ট করা হয় তাই বসে (item_name_override/
  manual_ply), শুধু Top Bottom/Divider লাইনের item name তার নিজের টেক্সট
  থেকেই ডিটেক্ট হয়।
- GMT Style = Style + '/' + Excel-এর সাইজ যেভাবে আছে হুবহু (আনমডিফাইড),
  যেমন 'ZIN01/1/1.5 YRS'।
- GMT Size = মডিফাইড সাইজ:
    - '1/1.5 YRS' -> '1/1.5Y' (স্পেস বাদ, 'YRS' -> 'Y')
    - '26/X' বা '10/XL' -> শুধু স্ল্যাশের পরের অংশ ('X'/'XL')
- GMT PO <- 'PO NO' কলাম। Reference <- 'COLOR/WASH' কলাম। Pack Type <-
  'ARTICLE NO' কলাম। EWO No সবসময় N/A (OUT-HOUSE-এ লাগে না)।
- Booking Qty 0/ফাঁকা হলে রো বাদ (Amigo-র মতোই একই নিয়ম)।
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


def _get_visible_sheet_names(file_stream, filename):
    """হাইড শিট বাদ দেয়। শনাক্ত করতে না পারলে None রিটার্ন করে (তখন সব
    শিটই বিবেচনা করা হবে)।"""
    try:
        file_stream.seek(0)
        if filename.lower().endswith('.xls'):
            import xlrd
            book = xlrd.open_workbook(file_contents=file_stream.read())
            return [
                name for name in book.sheet_names()
                if book.sheet_by_name(name).visibility == 0
            ]
        else:
            from openpyxl import load_workbook as _load_wb
            wb = _load_wb(file_stream, read_only=True, data_only=True)
            names = [ws.title for ws in wb.worksheets if ws.sheet_state == 'visible']
            wb.close()
            return names
    except Exception:
        return None
    finally:
        file_stream.seek(0)


def _row_label_map(df, r):
    labels = {}
    for c in range(df.shape[1]):
        v = _clean(df.iat[r, c])
        if v:
            labels[_norm(v)] = c
    return labels


def _find_header_row(df, max_scan=40):
    """'STYLE'/'SIZE'/'PO NO'/'BOOKING QNTY'/'LENGTH'/'WIDTH'/'HEIGHT' —
    এই লেবেলগুলো একই রো-তে থাকা খুঁজে বের করে। রিটার্ন করে
    (header_row, label_map) অথবা None।"""
    required = {'style', 'size', 'pono', 'length', 'width', 'height'}
    for r in range(min(df.shape[0], max_scan)):
        labels = _row_label_map(df, r)
        norm_keys = set(labels.keys())
        has_qty = any('bookingqnty' in k or k == 'bookingqty' for k in norm_keys)
        if required.issubset(norm_keys) and has_qty:
            qty_col = next((c for k, c in labels.items() if 'bookingqnty' in k or k == 'bookingqty'), None)
            labels['bookingqnty'] = qty_col
            return r, labels
    return None


def _transform_gmt_size(raw_size):
    """'1/1.5 YRS' -> '1/1.5Y' ; '26/X' -> 'X' ; '10/XL' -> 'XL'।"""
    s = (raw_size or '').strip()
    if re.search(r'yrs?\b', s, re.I):
        return re.sub(r'\s*yrs?\b', 'Y', s, flags=re.I).replace(' ', '')
    if '/' in s:
        last = s.split('/')[-1].strip()
        if last and not last.replace('.', '', 1).isdigit():
            return last
    return s


_TRAILER_MEAS_RE = re.compile(r'(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*cm', re.I)


def _find_trailer_items(df, header_row, qty_col):
    """মেইন টেবিলের নিচে 'TOP BOTTOM ... CM = <qty>' / '... DIVIDER ...
    CM = <qty>' জাতীয় লাইন খুঁজে বের করে। qty ওই একই রো-এর 'BOOKING QNTY'
    কলাম থেকে নেওয়া হয় (কলাম-পজিশন হেডার থেকেই ডাইনামিকভাবে জানা)।"""
    items = []
    n_rows = df.shape[0]
    for r in range(header_row + 1, n_rows):
        row_text = ' '.join(_clean(df.iat[r, c]) for c in range(df.shape[1]))
        low = row_text.lower()
        if 'divider' in low:
            item_name = 'Divider'
        elif 'top' in low and 'bottom' in low:
            item_name = 'Top Bottom'
        else:
            continue

        m = _TRAILER_MEAS_RE.search(low)
        if not m:
            continue
        length, width = m.group(1), m.group(2)

        qty_val = df.iat[r, qty_col] if qty_col is not None and qty_col < df.shape[1] else None
        if not _is_num(qty_val) or (_num(qty_val) or 0) <= 0:
            continue

        items.append({
            'item_name': item_name,
            'ewo_no': 'N/A',
            'style_no': 'N/A',
            'po_no': 'N/A',
            'length': length,
            'width': width,
            'height': '',  # Top Bottom/Divider-এ Height থাকে না
            'ply': '3',  # ইউজার-কনফার্মড: Top Bottom/Divider সবসময় 3 ply (UI সিলেকশন এখানে প্রযোজ্য না)
            'qty': round(_num(qty_val)),  # ইউজার-কনফার্মড: Carton Qty রাউন্ড করে বসবে
            'pack_type': 'N/A',
            'reference': 'N/A',
            'remarks': '',
            'color': 'N/A',
            'size': 'N/A',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
        })
    return items


def read_sinha_booking_sheet(df, sheet_name, item_name_override='', manual_ply=''):
    """একটা শিট থেকে (master_items, trailer_items) রিটার্ন করে। এই ফরম্যাট
    না হলে (হেডার না মিললে) দুটোই খালি লিস্ট।"""
    found = _find_header_row(df)
    if not found:
        return [], []
    header_row, labels = found

    style_col = labels['style']
    colorwash_col = labels.get('colorwash')
    size_col = labels['size']
    po_col = labels['pono']
    article_col = labels.get('articleno')
    length_col = labels['length']
    width_col = labels['width']
    height_col = labels['height']
    qty_col = labels['bookingqnty']

    master_items = []
    running_style = ''
    running_color = ''
    running_po = ''
    running_article = ''

    r = header_row + 1
    n_rows = df.shape[0]
    while r < n_rows:
        style_val = _clean(df.iat[r, style_col])
        if style_val:
            running_style = style_val
        if colorwash_col is not None:
            cw = _clean(df.iat[r, colorwash_col])
            if cw:
                running_color = cw
        po_val = _clean(df.iat[r, po_col])
        if po_val:
            running_po = po_val
        if article_col is not None:
            art = _clean(df.iat[r, article_col])
            if art:
                running_article = art

        size_val = _clean(df.iat[r, size_col])  # নিজের রো-তেই থাকা লাগবে — ফরওয়ার্ড-ফিল না
        qty_val = df.iat[r, qty_col]

        is_data_row = bool(running_style) and bool(size_val) and _is_num(qty_val)
        if not is_data_row:
            r += 1
            continue

        req_qty = _num(qty_val) or 0
        if req_qty <= 0:
            r += 1
            continue  # ইউজার-কনফার্মড: 0/ফাঁকা কোয়ান্টিটির রো বাদ
        req_qty = round(req_qty)  # ইউজার-কনফার্মড: Carton Qty রাউন্ড করে বসবে

        length = _fmt_num(df.iat[r, length_col])
        width = _fmt_num(df.iat[r, width_col])
        height = _fmt_num(df.iat[r, height_col])

        master_items.append({
            'item_name': item_name_override or 'Master Carton',
            'ewo_no': 'N/A',
            'style_no': f"{running_style}/{size_val}",
            'po_no': running_po or 'N/A',
            'length': length,
            'width': width,
            'height': height,
            'ply': manual_ply.strip() if manual_ply else '5',
            'qty': req_qty,
            'pack_type': running_article or 'N/A',
            'reference': running_color or 'N/A',
            'remarks': '',
            'color': 'N/A',
            'size': _transform_gmt_size(size_val),
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_sheet': sheet_name,
        })
        r += 1

    trailer_items = _find_trailer_items(df, header_row, qty_col)
    for it in trailer_items:
        # ply এখানে ওভাররাইড করা হচ্ছে না — Top Bottom/Divider সবসময় 3 ply,
        # সেটা _find_trailer_items()-এই হার্ডকোড করা আছে (UI-এর Ply
        # সিলেকশন শুধু Master Carton রো-তে প্রযোজ্য)।
        it['_sheet'] = sheet_name

    return master_items, trailer_items


def read_sinha_booking_file(file_stream, filename, item_name_override='', manual_ply=''):
    """একটা .xls/.xlsx ফাইলের সব (visible) শিট থেকে ডাটা বের করে।
    রিটার্ন করে (master_items, trailer_items, warnings)।"""
    file_stream.seek(0)
    sheets = pd.read_excel(file_stream, sheet_name=None, header=None)
    visible_names = _get_visible_sheet_names(file_stream, filename)

    all_master = []
    all_trailer = []
    warnings = []

    for sheet_name, df in sheets.items():
        if visible_names is not None and sheet_name not in visible_names:
            continue
        master_items, trailer_items = read_sinha_booking_sheet(
            df, sheet_name, item_name_override=item_name_override, manual_ply=manual_ply)
        all_master.extend(master_items)
        all_trailer.extend(trailer_items)

    if not all_master and not all_trailer:
        warnings.append(f"⚠️ '{filename}': কোনো লাইন-আইটেম পাওয়া যায়নি (পরিচিত ফরম্যাট না হতে পারে)।")

    return all_master, all_trailer, warnings


def combine_sinha_booking_files(file_tuples, item_name_override='', manual_ply=''):
    """একাধিক Sinha বুকিং ফাইল কম্বাইন করে। অর্ডারিং নিয়ম (Amigo Bangladesh
    কনভেনশন অনুযায়ী): সব ফাইলের সব শিটের মেইন (Master Carton) ডাটা আগে,
    তারপর সব ফাইলের Top Bottom/Divider লাইন একেবারে সবার শেষে।
    file_tuples: [(BytesIO, filename), ...]
    রিটার্ন করে (line_items, warnings)।"""
    all_master = []
    all_trailer = []
    all_warnings = []
    for file_stream, filename in file_tuples:
        master_items, trailer_items, warns = read_sinha_booking_file(
            file_stream, filename, item_name_override=item_name_override, manual_ply=manual_ply)
        for it in master_items:
            it['_source_file'] = filename
        for it in trailer_items:
            it['_source_file'] = filename
        all_master.extend(master_items)
        all_trailer.extend(trailer_items)
        all_warnings.extend(warns)

    combined = all_master + all_trailer
    return combined, all_warnings
