import re
import xlrd

# ---------------------------------------------------------------------------
# RADICAL DESIGN LTD — Buyer: Original Marines
# 'PURCHASE ORDER "CARTON"' .xls ফরম্যাট। একটা ফাইলে অনেকগুলো hidden শিট
# থাকে (ড্রাফট/পুরনো সংস্করণ), কিন্তু শুধু একটামাত্র শিট visible থাকে —
# সেটাই আসল ডাটা, বাকি সব hidden শিট স্কিপ হয়ে যায়।
#
# দুই রকম সাব-ভ্যারিয়েন্ট দেখা গেছে, একই হেডার-সিগনেচার (SL.NO + BUYER)
# দিয়ে দুটোই চেনা যায়:
#   সাধারণ ভার্সন: SL.NO | BUYER | MEAS.(CM) [৬টা আলাদা সেলে ছড়ানো, যেমন
#     '35.0 | X | 26.0 | X | 13.0 | CM'] | PCS PER CARTON | ... |
#     QTY.(PCS) | CBM। কোনো COLOR/SIZE কলাম নেই।
#   ডিটেইলড/LOGISTIC ভার্সন: SL.NO | BUYER | SIZE | MEAS.(CM) [একটাই সেলে,
#     যেমন '29X20X10'] | PCS PER CARTON | ... | QTY.(PCS) | CBM | COLOR।
#     COLOR শুধু প্রতিটা কালার-ব্লকের প্রথম রো-তে থাকে (forward-fill)।
#
# ইউজার-কনফার্মড ম্যাপিং:
#   শিটের নাম  -> Gmt. PO
#   BUYER      -> Gmt. Style No
#   COLOR      (থাকলে) -> Reference/SKU Number
#   SIZE       (থাকলে) -> Pack Type
#   MEAS.(CM)  -> Length/Width/Height
#   QTY.(PCS)  -> Order Qty
#
# Top Bottom: শিটের নিচে 'NOTE : Please provide us with N pieces of tops
#   for each carton. ... should be X cm less than the carton
#   measurements.' — এই নোট থেকেই cm-কমানোর পরিমাণ (X) আর প্রতি-কার্টুন
#   কয়টা top (N) সরাসরি পড়ে নেওয়া হয় (হার্ডকোড করা হয়নি, কারণ বুকিং-ভেদে
#   সংখ্যা বদলাতে পারে — এই দুই ফাইলেই যেমন 3cm/2pcs, ভবিষ্যতে অন্য হতে
#   পারে)। প্রতিটা ইউনিক Master Carton measurement-গ্রুপের জন্য একটা করে
#   Top Bottom লাইন: Length/Width নোটে বলা cm কম (Height লাগে না), Qty =
#   গ্রুপের মোট Qty × নোটে বলা সংখ্যা, Ply ফিক্সড ৩। নোট না পেলে বা নোট
#   থেকে সংখ্যা বের করা না গেলে Top Bottom যোগ হয় না, বরং ⚠️ ওয়ার্নিং।
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


_MEAS_RE = re.compile(r'(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)')


def _parse_measurement(text):
    if not text:
        return '', '', ''
    m = _MEAS_RE.search(str(text))
    if not m:
        return '', '', ''
    return _fmt_num(m.group(1)), _fmt_num(m.group(2)), _fmt_num(m.group(3))


_CM_LESS_RE = re.compile(r'(\d+\.?\d*)\s*cm\s*less', re.I)
_PIECES_RE = re.compile(r'(\d+)\s*pieces?\s*of\s*tops?', re.I)


def _parse_top_bottom_note(all_text):
    """পুরো শিটের টেক্সট থেকে cm-কমানোর পরিমাণ আর প্রতি-কার্টুন top-এর
    সংখ্যা বের করে। রিটার্ন করে (cm_less, pieces_per_carton) — কিছু বের
    করতে না পারলে সেই অংশটা None।"""
    cm_m = _CM_LESS_RE.search(all_text)
    pcs_m = _PIECES_RE.search(all_text)
    cm_less = float(cm_m.group(1)) if cm_m else None
    pieces = int(pcs_m.group(1)) if pcs_m else None
    return cm_less, pieces


def _get_visible_sheet(book):
    for s in book.sheets():
        if s.visibility == 0:  # 0 = visible
            return s
    return None


def _find_header_row(sheet, max_scan=30):
    """'SL.NO' আর 'BUYER' — এই দুটো লেবেল একই রো-তে (যেকোনো কলামে,
    কলাম-পজিশন ফাইল-ভেদে বদলাতে পারে) থাকলে সেটাই হেডার।"""
    for r in range(min(max_scan, sheet.nrows)):
        labels = {}
        for c in range(sheet.ncols):
            lbl = _norm(sheet.cell_value(r, c))
            if lbl and lbl not in labels:
                labels[lbl] = c
        if 'slno' in labels and 'buyer' in labels:
            return r, labels
    return None, None


def _read_one_file(file_stream, filename, item_name_override, manual_ply):
    """একটা ফাইল থেকে ডাটা বের করে। রিটার্ন করে (line_items, warnings)।"""
    file_stream.seek(0)
    book = xlrd.open_workbook(file_contents=file_stream.read())
    sheet = _get_visible_sheet(book)
    if sheet is None:
        return [], []

    header_row, labels = _find_header_row(sheet)
    if header_row is None:
        return [], []

    style_col = labels.get('buyer')
    meas_col = labels.get('meascm')
    pcs_per_carton_col = labels.get('pcspercarton')
    qty_col = labels.get('qtypcs')
    ref_col = labels.get('color')
    pack_col = labels.get('size')
    if style_col is None or meas_col is None or qty_col is None:
        return [], []

    boundary_col = pcs_per_carton_col if (pcs_per_carton_col is not None and pcs_per_carton_col > meas_col) else meas_col + 1

    default_item_name = item_name_override or 'Master Carton'
    ply_value = manual_ply.strip() if manual_ply else 'N/A'
    po_no = _clean(sheet.name) or 'N/A'

    master_items = []
    group_order = []
    group_sums = {}
    current_ref = ''
    consecutive_blank = 0

    for r in range(header_row + 1, sheet.nrows):
        row = [sheet.cell_value(r, c) if sheet.cell_value(r, c) != '' else None for c in range(sheet.ncols)]

        first_text = _clean(row[0]) if len(row) > 0 else ''
        if _norm(first_text) == 'total':
            break

        style_val = _clean(row[style_col]) if style_col < len(row) else ''
        if not style_val:
            consecutive_blank += 1
            if consecutive_blank >= 100:
                break
            continue
        consecutive_blank = 0

        if ref_col is not None and ref_col < len(row) and row[ref_col] not in (None, ''):
            current_ref = _clean(row[ref_col])

        meas_text = ' '.join(
            str(row[c]) for c in range(meas_col, min(boundary_col, len(row)))
            if row[c] not in (None, '')
        )
        l, w, h = _parse_measurement(meas_text)
        if not l:
            continue

        qty_val = row[qty_col] if qty_col < len(row) else None
        if not _is_num(qty_val) or float(qty_val) <= 0:
            continue

        pack_val = _clean(row[pack_col]) if pack_col is not None and pack_col < len(row) else ''

        master_items.append({
            'item_name': default_item_name,
            'ewo_no': 'N/A',
            'style_no': style_val,
            'po_no': po_no,
            'length': l, 'width': w, 'height': h,
            'ply': ply_value,
            'qty': round(float(qty_val)),
            'pack_type': pack_val or 'N/A',
            'reference': current_ref or 'N/A',
            'color': '', 'size': '',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '', 'delivery_address_pdf': '',
            '_source_file': filename,
        })

        key = (l, w, h)
        if key not in group_sums:
            group_sums[key] = 0
            group_order.append(key)
        group_sums[key] += round(float(qty_val))

    if not master_items:
        return [], []

    all_items = list(master_items)
    warnings = []

    all_text = ' '.join(
        _clean(sheet.cell_value(r, c))
        for r in range(sheet.nrows) for c in range(sheet.ncols)
        if sheet.cell_value(r, c) not in ('', None)
    )
    cm_less, pieces = _parse_top_bottom_note(all_text)

    if cm_less is not None and pieces is not None:
        for key in group_order:
            l, w, h = key
            try:
                tb_l = _fmt_num(float(l) - cm_less)
                tb_w = _fmt_num(float(w) - cm_less)
            except (TypeError, ValueError):
                continue
            all_items.append({
                'item_name': 'Top Bottom',
                'ewo_no': 'N/A',
                'style_no': 'N/A',
                'po_no': po_no,
                'length': tb_l, 'width': tb_w, 'height': '',
                'ply': '3',
                'qty': group_sums[key] * pieces,
                'pack_type': 'N/A',
                'reference': 'N/A',
                'color': '', 'size': '',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': '', 'delivery_address_pdf': '',
                '_source_file': filename,
            })
    else:
        warnings.append(
            f"⚠️ '{filename}': Top Bottom-এর নোট পাওয়া যায়নি বা নোট থেকে cm-কমানোর "
            f"পরিমাণ/প্রতি-কার্টুন সংখ্যা বের করা যায়নি — তাই Top Bottom লাইন যোগ করা "
            f"হয়নি। ম্যানুয়ালি চেক করুন।"
        )

    return all_items, warnings


def combine_radicaldesign_booking_files(files, item_name_override='', manual_ply=''):
    """BATCH_REGISTRY-এর uniform কল-সিগনেচার। রিটার্ন করে
    (line_items, warnings)।"""
    all_items = []
    all_warnings = []
    for file_stream, filename in files:
        try:
            items, warns = _read_one_file(file_stream, filename, item_name_override, manual_ply)
        except Exception as e:
            all_warnings.append(f"⚠️ '{filename}' প্রসেস করতে সমস্যা হয়েছে: {e}")
            continue
        if not items and not warns:
            all_warnings.append(f"⚠️ '{filename}': পরিচিত ফরম্যাট মেলেনি (বা visible শিট পাওয়া যায়নি), স্কিপ করা হয়েছে।")
            continue
        all_items.extend(items)
        all_warnings.extend(warns)
    return all_items, all_warnings
