"""
Intimate Attire Limited — Buyer: Max-Dubai
'Purchase Order / Carton Booking' Excel ফরম্যাট।

একটা ফাইলে একটাই শিট, টেবিলে প্রতিটা "ব্লক" ৩টা (কখনো ২টা, যদি U Divider
না থাকে) রো নিয়ে গঠিত:
  1. 'Carton'         — Style/PO/Generic name/Specification/Quantity সব
     এই রো-তেই থাকে (Master Carton, ply 5)।
  2. '2 Leg Dividder' (বানান ফাইল-ভেদে বদলাতে পারে, যেমন '2 Leg Divider')
     — শুধু Specification/Quantity থাকে, Style/PO/Generic name ফাঁকা
     (আগের 'Carton' রো থেকে forward-fill করা হয়) — টেমপ্লেটে Item Name
     'U Divider', ply 3।
  3. 'Top Bottom'      — একই রকম, ply 3। Specification-এ শুধু L x W
     (height থাকে না)।

ব্যবসায়িক নিয়ম (ইউজার-কনফার্মড):
- Item Name ও Ply — সম্পূর্ণ Item কলামের টেক্সট থেকেই ডিটেক্ট হয় (UI
  সিলেকশন প্রযোজ্য না): Carton->Master Carton/5, 2 Leg Divider->U
  Divider/3, Top Bottom->Top Bottom/3।
- Style No <- 'Style NO.' কলাম, GMT PO <- 'PO No' কলাম, Reference <-
  'Generic name' কলাম — এই তিনটা শুধু 'Carton' রো-তেই থাকে, তাই একই
  ব্লকের U Divider/Top Bottom রো-তে forward-fill করে বসানো হয়।
- Measurement <- 'Specification' কলাম (L/W/H বা শুধু L/W টেক্সট পার্স)।
- Qty <- 'Quantity' কলাম। 0/ফাঁকা হলে সেই রো বাদ।
- হেডারের টেক্সট সামান্য ভিন্ন হতে পারে বলে (ইউজার-কনফার্মড), সব কলাম
  label-এর substring match দিয়ে ডাইনামিকভাবে বের করা হয়, fixed index
  ধরা হয় না।
- অর্ডারিং: সব ফাইলের সব Master Carton আগে, তারপর সব U Divider, তারপর
  সব Top Bottom — এই ক্রমে (ইউজার-কনফার্মড, Everbright-এর কনভেনশনের
  সম্প্রসারণ, ৩ গ্রুপে)।
"""
import re
import pandas as pd


def _norm(v):
    return re.sub(r'[^a-z0-9]', '', str(v or '').lower())


def _clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    # \xa0 (non-breaking space) সহ যেকোনো হোয়াইটস্পেস normalize করা হচ্ছে —
    # কিছু ফাইলে Specification টেক্সটে non-breaking space থাকে।
    return re.sub(r'[\s\xa0]+', ' ', str(v)).strip()


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


def _parse_measurement(text):
    """'L58 X W36 X H30 CM' -> ('58','36','30') ; 'L35 X W25 CM' ->
    ('35','25','') — height না থাকলে ফাঁকা।"""
    t = _clean(text)
    nums = re.findall(r'(\d+\.?\d*)', t)
    l = _fmt_num(nums[0]) if len(nums) > 0 else ''
    w = _fmt_num(nums[1]) if len(nums) > 1 else ''
    h = _fmt_num(nums[2]) if len(nums) > 2 else ''
    return l, w, h


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


def _find_header_row(rows, max_scan=25):
    for i, row in enumerate(rows[:max_scan]):
        labels = _row_label_map(row)
        if 'item' in labels and any('style' in k for k in labels) and any('specification' in k for k in labels):
            return i, labels
    return None, None


def _classify_item(item_text):
    """Item কলামের টেক্সট থেকে (item_name, ply) ঠিক করে। চেনা টাইপ না
    হলে None রিটার্ন করে (সেই রো স্কিপ হবে, যেমন 'Total…' রো)।"""
    n = _norm(item_text)
    if n == 'carton':
        return 'Master Carton', '5'
    if 'divid' in n:  # '2 Leg Dividder'/'2 Leg Divider'/'Divider' — বানান-ভিন্নতা সহনশীল
        return 'U Divider', '3'
    if 'topbottom' in n:
        return 'Top Bottom', '3'
    return None


def read_intimate_booking_sheet(rows, sheet_name):
    """একটা শিট থেকে (master_items, divider_items, topbottom_items,
    warnings) বের করে — তিনটা আলাদা গ্রুপ, caller এগুলো ইউজার-কনফার্মড
    ক্রমে (Master -> U Divider -> Top Bottom) সাজিয়ে বসাবে।"""
    header_row, labels = _find_header_row(rows)
    if header_row is None:
        return [], [], [], []

    item_col = labels.get('item')
    style_col = _find_col(labels, 'style')
    po_col = _find_col(labels, 'pono') or labels.get('pono')
    generic_col = _find_col(labels, 'genericname')
    spec_col = _find_col(labels, 'specification')
    qty_col = _find_col(labels, 'quantity')

    if item_col is None or spec_col is None or qty_col is None:
        return [], [], [], [f"⚠️ শিট '{sheet_name}': Item/Specification/Quantity কলাম পাওয়া যায়নি — স্কিপ করা হয়েছে।"]

    master_items, divider_items, topbottom_items = [], [], []
    running_style = ''
    running_po = ''
    running_ref = ''

    r = header_row + 1
    n_rows = len(rows)
    while r < n_rows:
        row = rows[r]
        item_val = _clean(row[item_col]) if item_col < len(row) else ''
        classified = _classify_item(item_val)
        if classified is None:
            r += 1
            continue  # 'Total…' রো বা অচেনা টেক্সট — স্কিপ
        item_name, ply = classified

        style_val = _clean(row[style_col]) if style_col is not None and style_col < len(row) else ''
        po_val = _clean(row[po_col]) if po_col is not None and po_col < len(row) else ''
        ref_val = _clean(row[generic_col]) if generic_col is not None and generic_col < len(row) else ''

        # Style/PO/Generic name শুধু 'Carton' রো-তেই থাকে — এই ব্লকের
        # পরের U Divider/Top Bottom রো-তে সেই একই ভ্যালু forward-fill।
        if item_name == 'Master Carton':
            running_style = style_val
            running_po = po_val
            running_ref = ref_val
        else:
            style_val = style_val or running_style
            po_val = po_val or running_po
            ref_val = ref_val or running_ref

        qty_val = row[qty_col] if qty_col < len(row) else None
        if not _is_num(qty_val) or (_num(qty_val) or 0) <= 0:
            r += 1
            continue  # qty 0/ফাঁকা — বাদ

        length, width, height = _parse_measurement(row[spec_col] if spec_col < len(row) else '')

        item = {
            'item_name': item_name,
            'ewo_no': 'N/A',
            'style_no': style_val or 'N/A',
            'po_no': po_val or 'N/A',
            'length': length,
            'width': width,
            'height': height,
            'ply': ply,
            'qty': round(_num(qty_val)),
            'pack_type': 'N/A',
            'reference': ref_val or 'N/A',
            'remarks': '',
            'color': 'N/A',
            'size': 'N/A',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_sheet': sheet_name,
        }
        if item_name == 'Master Carton':
            master_items.append(item)
        elif item_name == 'U Divider':
            divider_items.append(item)
        else:
            topbottom_items.append(item)

        r += 1

    warnings = []
    if not master_items and not divider_items and not topbottom_items:
        warnings.append(f"⚠️ শিট '{sheet_name}': কোনো ভ্যালিড (qty>0) লাইন-আইটেম পাওয়া যায়নি।")

    return master_items, divider_items, topbottom_items, warnings


def read_intimate_booking_file(file_stream, filename):
    """একটা .xls/.xlsx ফাইলের সব শিট থেকে ডাটা বের করে। রিটার্ন করে
    (master_items, divider_items, topbottom_items, warnings)।"""
    file_stream.seek(0)
    sheets = pd.read_excel(file_stream, sheet_name=None, header=None)

    all_master, all_divider, all_topbottom = [], [], []
    all_warnings = []
    for sheet_name, df in sheets.items():
        m, d, t, warns = read_intimate_booking_sheet(df.values.tolist(), sheet_name)
        all_master.extend(m)
        all_divider.extend(d)
        all_topbottom.extend(t)
        all_warnings.extend(warns)

    if not all_master and not all_divider and not all_topbottom:
        all_warnings.append(f"⚠️ '{filename}': কোনো ভ্যালিড (qty>0) লাইন-আইটেম পাওয়া যায়নি।")

    return all_master, all_divider, all_topbottom, all_warnings


def combine_intimate_booking_files(files, item_name_override='', manual_ply=''):
    """files: [(BytesIO, filename), ...] — BATCH_REGISTRY-এর uniform কল-
    সিগনেচার (item_name_override/manual_ply ইচ্ছাকৃতভাবে ব্যবহার হচ্ছে না,
    কারণ এই কাস্টমারের Item Name/Ply সম্পূর্ণ ফাইলের Item কলাম থেকেই
    ডিটেক্ট হয়)। রিটার্ন করে (line_items, warnings) — সব ফাইলের সব Master
    Carton আগে, তারপর সব U Divider, তারপর সব Top Bottom (ইউজার-কনফার্মড
    ক্রম)।"""
    all_master, all_divider, all_topbottom = [], [], []
    all_warnings = []
    for file_stream, filename in files:
        m, d, t, warns = read_intimate_booking_file(file_stream, filename)
        for it in m + d + t:
            it['_source_file'] = filename
        all_master.extend(m)
        all_divider.extend(d)
        all_topbottom.extend(t)
        all_warnings.extend(warns)

    combined = all_master + all_divider + all_topbottom
    return combined, all_warnings
