import re
import pandas as pd

# ---------------------------------------------------------------------------
# Columbia Apparels Limited — বায়ার: Target Australia
# 'PRINTED 3 & 5 PLY SOLID CARTON BOOKING FOR BUYER TARGET AUSTRALIA' এক্সেল
# ফরম্যাট। একই Columbia কাস্টমারের GU বায়ারের ফরম্যাট থেকে সম্পূর্ণ ভিন্ন
# (বায়ার আলাদা হওয়ায় লেআউট আলাদা) — তাই আলাদা এক্সট্র্যাক্টর।
#
# ম্যাপিং:
#   PO No column       -> po_no (Master Carton রো-তে সরাসরি এখান থেকে;
#                          Divider রো-তে এই কলামে 'Top / Bottom / Divider'
#                          টেক্সট থাকে বলে, ফাইলের Master Carton সেকশন থেকে
#                          পাওয়া PO No-ই Divider রো-তে reuse হয়)
#   Style (হেডার সেল)   -> style_no (পুরো ফাইলে একটাই স্টাইল, সব রো-তে বসে)
#   CPID                -> reference (ফাঁকা হলে N/A)
#   Measurement         -> length/width/height (L x W x H পার্স করে; কিছু
#                          Divider রো-তে 'x' এর বদলে ড্যাশ থাকে, সেটাও সামলানো হয়)
#   Qty Pcs             -> qty
#   PLY                 -> ply ('5 PLY'/'3 PLY' থেকে শুধু সংখ্যাটা, '5'/'3')
#
# যেসব রো-তে Measurement বা Qty ফাঁকা, সেগুলো বাদ যায় (ইউজারের নির্দেশ
# অনুযায়ী — ব্লাংক রো টেমপ্লেটে যাবে না)।
#
# 'Top / Bottom / Divider' লেবেলযুক্ত রো-গুলো Master Carton হিসেবে না গিয়ে
# 'Divider' নামে আলাদা আইটেম হিসেবে যায় — UI থেকে সিলেক্ট করা item_name_override
# (Master Carton/Elastic Hanger/ইত্যাদি) এই রো-গুলোতে প্রযোজ্য না।
#
# EWO No এই ফরম্যাটেও কখনো থাকে না -> সরাসরি 'N/A'।
# ---------------------------------------------------------------------------


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _is_blank(v):
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return str(v).strip() == ''


def _parse_measurement(m):
    """'72 x 44 x 27 CM' -> (72.0, 44.0, 27.0)। কিছু Divider রো-তে
    'TOP-60 CM-50 CM'-এর মতো 'x' ছাড়া ড্যাশ দিয়ে লেখা থাকে — সেক্ষেত্রে
    স্ট্রিং-এর সব সংখ্যা বের করে প্রথম দুই/তিনটাকে L/W/H হিসেবে ধরা হয়।"""
    if _is_blank(m):
        return '', '', ''
    s = str(m).replace('\xa0', ' ')
    match3 = re.search(r'(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)', s)
    if match3:
        return float(match3.group(1)), float(match3.group(2)), float(match3.group(3))
    match2 = re.search(r'(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)', s)
    if match2:
        return float(match2.group(1)), float(match2.group(2)), ''
    nums = re.findall(r'\d+(?:\.\d+)?', s)
    if len(nums) >= 3:
        return float(nums[0]), float(nums[1]), float(nums[2])
    if len(nums) == 2:
        return float(nums[0]), float(nums[1]), ''
    return '', '', ''


def _parse_ply(v):
    """'5 PLY' -> '5', '3 PLY' -> '3'।"""
    if _is_blank(v):
        return ''
    m = re.search(r'(\d+)', str(v))
    return m.group(1) if m else ''


_REQUIRED_HEADER_TOKENS = (
    _norm('Sl No'), _norm('PO No'), _norm('CPID'),
    _norm('Measurement'), _norm('Qty Pcs'), _norm('PLY'),
)


def _find_header_row(rows, max_scan=25):
    """হেডার রো খুঁজে বের করে — 'Sl No'/'PO No'/'CPID'/'Measurement'/
    'Qty Pcs'/'PLY' সবগুলো একসাথে যেই রো-তে পাওয়া যায় সেটাই হেডার।"""
    for i, row in enumerate(rows[:max_scan]):
        normed = [_norm(c) for c in row]
        if all(tok in normed for tok in _REQUIRED_HEADER_TOKENS):
            return i
    return None


def _find_style_no(rows, max_scan=25):
    """'Style' লেখা সেলের ঠিক পরের সেল থেকে Style No বের করে (হেডার রো-এর
    আগে থাকে, তাই আলাদাভাবে স্ক্যান করা হয়)।"""
    for row in rows[:max_scan]:
        for ci, cell in enumerate(row):
            if _norm(cell) == _norm('Style') and ci + 1 < len(row) and not _is_blank(row[ci + 1]):
                return str(row[ci + 1]).strip()
    return ''


def read_columbia_target_australia_style_excel(
        file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """Columbia Apparels Limited — Target Australia বুকিং এক্সেল থেকে
    canonical লাইন-আইটেম লিস্ট বের করে (Master Carton + Divider)।

    item_name_override: Master Carton রো-গুলোর জন্য UI থেকে সিলেক্ট করা
    আইটেম নেম (Master Carton/Elastic Hanger Carton/Both Side Hanger Carton)।
    Divider রো-গুলোতে এটা প্রযোজ্য না — সবসময় 'Divider' বসে।

    manual_ply: PLY কলামে মান না পাওয়া গেলে (খুবই বিরল এই ফরম্যাটে যেহেতু
    প্রতি রো-তেই PLY থাকে) UI থেকে দেওয়া ফলব্যাক মান ব্যবহার হবে।
    """
    from outhouse_extractor import _read_excel_rows  # শেয়ার্ড multi-engine reader পুনর্ব্যবহার
    rows = _read_excel_rows(file_stream, filename)

    header_idx = _find_header_row(rows)
    if header_idx is None:
        raise ValueError(
            "'Sl No'/'PO No'/'CPID'/'Measurement'/'Qty Pcs'/'PLY' হেডিং-সহ "
            "Target Australia বুকিং টেবিল ফরম্যাট পাওয়া যায়নি"
        )

    header_row = rows[header_idx]
    col_map = {}
    for ci, cell in enumerate(header_row):
        n = _norm(cell)
        if n == _norm('PO No'):
            col_map['po_no'] = ci
        elif n == _norm('CPID'):
            col_map['cpid'] = ci
        elif n == _norm('Measurement'):
            col_map['measurement'] = ci
        elif n == _norm('Qty Pcs'):
            col_map['qty'] = ci
        elif n == _norm('PLY'):
            col_map['ply'] = ci

    style_no = _find_style_no(rows, max_scan=header_idx + 1)

    def get(row, field):
        ci = col_map.get(field)
        if ci is None or ci >= len(row):
            return ''
        v = row[ci]
        return '' if _is_blank(v) else v

    items = []
    file_po_no = ''  # Master Carton সেকশন থেকে পাওয়া PO No — Divider রো-তে reuse হবে

    for row in rows[header_idx + 1:]:
        if row is None or all(_is_blank(c) for c in row):
            continue

        po_raw = get(row, 'po_no')
        po_str = str(po_raw).strip()
        is_divider_row = 'divider' in _norm(po_str)

        measurement_val = get(row, 'measurement')
        qty_val = get(row, 'qty')
        # ব্লাংক রো বাদ — measurement আর qty দুটোই থাকতে হবে (ইউজারের নির্দেশ)
        if _is_blank(measurement_val) or _is_blank(qty_val):
            continue

        length, width, height = _parse_measurement(measurement_val)
        ply = _parse_ply(get(row, 'ply')) or (manual_ply.strip() if manual_ply else '')

        if is_divider_row:
            po_no_final = file_po_no or 'N/A'
            item_name = 'Divider'
        else:
            # টেবিলের নিচে ফুটার/নোট টেক্সট এই কলামে পড়ে যেতে পারে — তাই
            # PO No আসলেই সংখ্যা কিনা যাচাই করে সেগুলো বাদ দেওয়া হচ্ছে
            try:
                float(po_str.replace(',', ''))
            except ValueError:
                continue
            po_no_final = po_str
            file_po_no = file_po_no or po_no_final
            item_name = item_name_override or 'Master Carton'

        cpid_val = get(row, 'cpid')
        reference = str(cpid_val).strip() if not _is_blank(cpid_val) else 'N/A'

        items.append({
            'item_name': item_name,
            'ewo_no': 'N/A',
            'style_no': style_no,
            'po_no': po_no_final,
            'length': length,
            'width': width,
            'height': height,
            'ply': ply or 'N/A',
            'qty': qty_val,
            'pack_type': '',
            'reference': reference,
            'color': '',
            'size': '',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_source_file': filename,
        })

    if not items:
        raise ValueError("হেডার পাওয়া গেছে কিন্তু বৈধ (measurement+qty সহ) কোনো ডাটা রো পাওয়া যায়নি")

    return items
