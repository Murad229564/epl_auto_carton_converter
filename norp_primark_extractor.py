import re
import pandas as pd

# ---------------------------------------------------------------------------
# Norp Knit Industries Ltd. — Buyer: Primark
# 'BULK CARTON BOOKING SHEET' এক্সেল (.xlsx/.xls) ফরম্যাট।
# ফাইলের নাম: norp_primark_extractor.py  (customer_buyer নামকরণ)
# ফাংশন: combine_norp_primark_booking_files (batch), read_norp_primark_style_excel (এক ফাইল)
#
# টেবিল (হেডার: STYLE / For Sizes / Carton Meas. / Carton Qty / Divider Meas.
# / Divider Qty / Garments Qty / Remark):
#   STYLE         -> Gmt. Style No  (এবং টেমপ্লেটের উপরের PO ফিল্ড — _header_po)
#   For Sizes     -> Reference/SKU Number
#   Carton Meas.  -> Master/Elastic Hanger Carton-এর L x W x H
#   Carton Qty    -> ওই কার্টনের Order Qty
#   Divider Meas. -> Divider-এর L x W (হাইট নেই)
#   Divider Qty   -> Divider-এর Order Qty
#   Garments Qty  -> কখনোই ব্যবহার হয় না
#   Remark        -> Item Name ঠিক করে (নিচে দেখুন)
#
# প্রতি সাইজ-রো থেকে দুটো লাইন হয়: একটা কার্টন, একটা Divider। ক্রম: সব
# ফাইলের সব কার্টন আগে, তারপর সব Divider।
#
# Remark-এর নিয়ম (STYLE-ব্লকের প্রথম রো-তে লেখা থাকে, মার্জড সেল — নিচের
# রো-গুলোতে একই Remark প্রযোজ্য):
#   ফাঁকা / 'No Need Elastic' ...      -> Master Carton (UI-র সিলেকশন নয়)
#   'Elastic position ...' / 'Elastic' -> Elastic Hanger Carton
#   ফাইলে Remark কলামই না থাকলে        -> Master Carton + Warning
#
# Ply: ফাইলের লেখা ('40 LBS 3 PLY AUTO CARTON...') থেকে; না পেলে UI-এর
# manual_ply; তাও না থাকলে 'N/A'। Divider-এর Ply সবসময় 3।
#
# Warning (রিটার্ন হওয়া দ্বিতীয় লিস্টে): Remark কলাম নেই; Carton Qty-র যোগফল
# ফাইলের 'Total Qty'-র সাথে না মিললে; মেজারমেন্ট ফাঁকা থাকলে।
# ---------------------------------------------------------------------------

NA = 'N/A'
DIVIDER_PLY = '3'


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    if isinstance(v, float) and pd.isna(v):
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def _num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    try:
        return float(str(v).replace(',', '').strip())
    except (TypeError, ValueError):
        return None


def _parse_measurement(m):
    """'49x38x13 cm' -> (49.0, 38.0, 13.0); '45x34 cm' -> (45.0, 34.0, '')"""
    nums = re.findall(r'\d+(?:\.\d+)?', str(m or ''))
    vals = [float(n) for n in nums[:3]]
    while len(vals) < 3:
        vals.append('')
    return vals[0], vals[1], vals[2]


def _load_rows(file_stream, filename):
    try:
        from outhouse_extractor import _read_excel_rows
        return _read_excel_rows(file_stream, filename)
    except ImportError:
        file_stream.seek(0)
        return pd.read_excel(file_stream, header=None).values.tolist()


def _find_header(rows, max_scan=40):
    for i, row in enumerate(rows[:max_scan]):
        normed = [_norm(c) for c in row]
        if ('style' in normed and 'cartonmeas' in normed and 'cartonqty' in normed):
            heads = {}
            for ci, n in enumerate(normed):
                if n and n not in heads:
                    heads[n] = ci
            return i, heads
    return None, None


_NEGATIVE = re.compile(r'\b(no|not|without)\b[^a-z]*(need|require|elastic)', re.I)


def _is_elastic(remark):
    r = _clean(remark)
    if not r:
        return False
    if _NEGATIVE.search(r):   # 'No Need Elastic' -> Master Carton
        return False
    return 'elastic' in r.lower()


def read_norp_primark_style_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """এক ফাইল -> (cartons, dividers, warnings)। ফরম্যাট না মিললে ValueError।"""
    rows = _load_rows(file_stream, filename)
    header_idx, heads = _find_header(rows)
    if header_idx is None:
        raise ValueError("'STYLE / Carton Meas. / Carton Qty' হেডিং-সহ Primark টেবিল পাওয়া যায়নি")

    c_style = heads['style']
    c_sizes = heads.get('forsizes')
    c_cm = heads['cartonmeas']
    c_cq = heads['cartonqty']
    c_dm = heads.get('dividermeas')
    c_dq = heads.get('dividerqty')
    c_rem = heads.get('remark') or heads.get('remarks')

    warnings = []
    if c_rem is None:
        warnings.append(f"{filename}: 'Remark' কলাম পাওয়া যায়নি — সব কার্টন Master Carton ধরা হয়েছে, যাচাই করুন")

    # ফাইলের Ply (হেডার-এরিয়ার লেখা থেকে)
    ply_found = ''
    for row in rows[:header_idx]:
        for c in row:
            m = re.search(r'(\d+)\s*ply', _clean(c), re.I)
            if m:
                ply_found = m.group(1)
                break
        if ply_found:
            break
    ui_ply = manual_ply.strip() if manual_ply else ''
    carton_ply = ply_found or ui_ply or NA

    def get(row, ci):
        if ci is None or ci >= len(row):
            return ''
        return row[ci]

    cartons, dividers = [], []
    last_style = ''
    last_remark = ''
    total_row_qty = None
    for row in rows[header_idx + 1:]:
        if row is None:
            continue
        # 'Total Qty' রো পেলে টেবিল শেষ
        if any(_norm(c) == 'totalqty' for c in row):
            total_row_qty = _num(get(row, c_cq))
            break

        s = _clean(get(row, c_style))
        if s:  # নতুন STYLE-ব্লক শুরু -> Remark রিসেট
            last_style = s
            last_remark = ''
        r = _clean(get(row, c_rem)) if c_rem is not None else ''
        if r:
            last_remark = r

        sizes = _clean(get(row, c_sizes))
        cqty = _num(get(row, c_cq))
        dqty = _num(get(row, c_dq))
        has_c = cqty is not None and cqty > 0
        has_d = dqty is not None and dqty > 0
        if not (has_c or has_d):
            continue

        base = {
            'ewo_no': NA,
            'style_no': last_style,
            'po_no': '',
            'color': '',
            'size': '',
            'pack_type': '',
            'reference': sizes,
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_source_file': filename,
        }

        if has_c:
            if _is_elastic(last_remark):
                item_name = 'Elastic Hanger Carton'
            elif item_name_override in ('', 'Elastic Hanger Carton'):
                item_name = 'Master Carton'
            else:
                item_name = item_name_override
            L, W, H = _parse_measurement(get(row, c_cm))
            if L == '' or W == '':
                warnings.append(f"{filename}: '{sizes}' সাইজের Carton Meas. পাওয়া যায়নি")
            cartons.append({**base, 'item_name': item_name, 'length': L, 'width': W,
                            'height': H, 'ply': carton_ply, 'qty': round(cqty)})

        if has_d:
            L, W, _h = _parse_measurement(get(row, c_dm))
            if L == '' or W == '':
                warnings.append(f"{filename}: '{sizes}' সাইজের Divider Meas. পাওয়া যায়নি")
            dividers.append({**base, 'item_name': 'Divider', 'length': L, 'width': W,
                             'height': '', 'ply': DIVIDER_PLY, 'qty': round(dqty)})

    if not cartons and not dividers:
        raise ValueError("হেডার পাওয়া গেছে কিন্তু কোয়ান্টিটি-সহ কোনো ডাটা রো পাওয়া যায়নি")

    if total_row_qty is not None:
        got = sum(i['qty'] for i in cartons)
        if got != round(total_row_qty):
            warnings.append(f"{filename}: Carton Qty-র যোগফল ({got}) ফাইলের Total Qty ({round(total_row_qty)})-এর সাথে মেলেনি")

    # উপরের PO ফিল্ড = STYLE (একাধিক ইউনিক হলে '/' দিয়ে জোড়া)
    styles = []
    for it in cartons + dividers:
        if it['style_no'] and it['style_no'] not in styles:
            styles.append(it['style_no'])
    header_po = '/'.join(styles)
    for it in cartons + dividers:
        it['_header_po'] = header_po

    return cartons, dividers, warnings


def combine_norp_primark_booking_files(files, item_name_override='', manual_ply=''):
    """files: [(file_stream, filename), ...] -> (line_items, warnings)
    সব ফাইলের কার্টন আগে, তারপর সব Divider। কোনো ফাইল ব্যর্থ হলে স্কিপ হয়ে
    মেসেজ warnings-এ যায়।"""
    all_c, all_d, warnings = [], [], []
    for file_stream, filename in files:
        try:
            file_stream.seek(0)
            c, d, w = read_norp_primark_style_excel(
                file_stream, filename,
                item_name_override=item_name_override, manual_ply=manual_ply)
            all_c.extend(c)
            all_d.extend(d)
            warnings.extend(w)
        except Exception as e:
            warnings.append(f"{filename}: {e}")
    return all_c + all_d, warnings
