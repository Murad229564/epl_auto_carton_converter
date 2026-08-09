"""
Amigo Bangladesh Ltd — Buyer: Uniqlo
Carton Booking Excel ফরম্যাট (SOLID / ASSORT)।

একটা বুকিং-এ সাধারণত ১-৫টা .xls/.xlsx ফাইল আসতে পারে, প্রতিটাই SOLID বা
ASSORT টাইপের (ফাইলের ভেতরেই এই ইনফো পাওয়া যায়, filename-এর উপর ভরসা করা
হয় না)। প্রতিটা ফাইল থেকে দুই জায়গার ডাটা নেওয়া হয়:

1. 'SIZE BREAKDOWN' শিট (নাম-এ শেষে স্পেস/ভিন্নতা থাকতে পারে, এবং একই নামে
   একটা hidden ডুপ্লিকেট শিটও থাকতে পারে — শুধু visible শিটটাই পড়া হয়) —
   এখান থেকে Master Carton-এর লাইন-আইটেম আসে। SOLID টাইপে explicit
   Length/Width/Height কলাম থাকে (নাম্বার আকারেই), ASSORT টাইপে শুধু
   MEASUREMENT টেক্সট কলাম থাকে (পার্স করে নিতে হয়)।

2. 'Indent for Carton' শিট — এখান থেকে শুধু Top/Bottom আর Divider লাইন
   নেওয়া হয় (Master Carton রো এখানে স্কিপ করা হয়, কারণ সেটা Size
   Breakdown শিট থেকেই ইতিমধ্যে নেওয়া হয়ে গেছে)। Ply প্রতিটা রো-এর নিজের
   বর্ণনা-টেক্সট থেকে ডাইনামিকভাবে বের করা হয় (গ্রুপ-হেডারে ভুল/ভিন্ন ply
   লেখা থাকতে পারে বলে গ্রুপ-হেডারে ভরসা করা হয় না)।

ব্যবসায়িক নিয়ম (ইউজার-কনফার্মড):
- Required Carton Quantity কলামে যা লেখা আছে সেটাই হুবহু ব্যবহার হবে —
  অন্য কোনো কলাম দিয়ে হিসাব করে বসানো হবে না।
- এই কোয়ান্টিটি 0/ফাঁকা হলে সেই রো সম্পূর্ণ বাদ (both Size Breakdown ও
  Indent-এ, একই নিয়ম)।
- Z/U/W/Tray Divider (যেগুলোর height থাকতে পারে) — এখনো সাপোর্ট করা হয়নি,
  ভবিষ্যতে আলাদাভাবে যোগ হবে।
"""
import re
import pandas as pd


def _norm(v):
    """Label matching-এর জন্য: শুধু a-z0-9 রেখে lowercase — স্পেস/স্ল্যাশ/
    বিরাম চিহ্ন ভিন্ন হলেও (যেমন 'GROSSWT.' vs 'GROSS WT.', বা শিটের নামের
    শেষে বাড়তি স্পেস) মিলিয়ে ফেলার জন্য।"""
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


def _parse_two_or_three_nums(text):
    """'56CMX30CM' -> (56,30, None) ; '58CMX36CMX24CM(0.050M3)' -> (58,36,24).
    mm থাকলে (mm আছে কিনা লেখা থেকে বোঝা কঠিন এই ফরম্যাটে, তাই সবসময়
    Cm ধরা হয় — কাস্টমারের ফাইলে সবসময় CM-ই ব্যবহার হয়)।"""
    nums = re.findall(r'(\d+\.?\d*)', text or '')
    l = _fmt_num(nums[0]) if len(nums) > 0 else ''
    w = _fmt_num(nums[1]) if len(nums) > 1 else ''
    h = _fmt_num(nums[2]) if len(nums) > 2 else ''
    return l, w, h


def _get_visible_sheet_names(file_stream, filename):
    """হাইড (hidden/veryHidden) শিট বাদ দিতে দৃশ্যমান শিটের নামের লিস্ট
    বের করে। শনাক্ত করতে না পারলে None রিটার্ন করে (তখন নিরাপদ ডিফল্ট
    হিসেবে সব শিটই বিবেচনা করা হবে)।"""
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
    """একটা রো-এর সব সেল normalize করে {label: col_index} ডিকশনারি বানায়।
    ডুপ্লিকেট লেবেল থাকলে সবচেয়ে ডানপাশের কলামটাই থেকে যায় (dict overwrite) —
    এটা ইচ্ছাকৃত: SOLID শিটে ভুয়া 'LENGTH' কলাম (col 10, সবসময় '000') এর
    ডানপাশেই আসল 'Length/Weight(আসলে Width)/Height' ট্রিপলেট থাকে, তাই শেষেরটাই
    টিকে থাকবে।"""
    labels = {}
    for c in range(df.shape[1]):
        v = _clean(df.iat[r, c])
        if v:
            labels[_norm(v)] = c
    return labels


def _find_lwh_triplet(df, header_row):
    """একই রো-তে পাশাপাশি তিনটা কলাম 'length' + ('weight' বা 'width') +
    'height' খুঁজে বের করে (SOLID-টাইপ শিটে থাকে, ASSORT-টাইপে থাকে না)।"""
    n_cols = df.shape[1]
    for c in range(n_cols - 2):
        a = _norm(df.iat[header_row, c])
        b = _norm(df.iat[header_row, c + 1])
        d = _norm(df.iat[header_row, c + 2])
        if a == 'length' and b in ('width', 'weight') and d == 'height':
            return c, c + 1, c + 2
    return None


def _find_size_breakdown_header_row(df):
    """SIZE BREAKDOWN শিটে (SOLID বা ASSORT দুই টাইপেরই) হেডার রো খুঁজে বের
    করে এবং টাইপ (SOLID/ASSORT) শনাক্ত করে। এই শিটে হেডার সাধারণত দুই রো-তে
    ভাগ হয়ে থাকে (উপরের রো-তে গ্রুপ-লেবেল যেমন 'Required Carton'/'WH Code',
    নিচের রো-তে ডিটেইল-লেবেল যেমন 'PO NO#'/'COLOR'/'SIZE') — তাই দুই রো-এর
    লেবেলই মার্জ করে রিটার্ন করা হয়। রিটার্ন করে (header_row, sheet_type,
    label_map) অথবা কিছু না পেলে None।"""
    for r in range(min(df.shape[0], 10)):
        labels = _row_label_map(df, r)
        sheet_type = None
        if 'pono' in labels and 'color' in labels and 'size' in labels:
            sheet_type = 'SOLID'
        elif 'pono' in labels and 'set' in labels:
            sheet_type = 'ASSORT'
        if sheet_type:
            merged = dict(_row_label_map(df, r - 1)) if r >= 1 else {}
            merged.update(labels)
            return r, sheet_type, merged
    return None


def _find_col_containing(labels, *must_contain):
    """labels ডিকশনারিতে এমন একটা কলাম খুঁজে বের করে যার normalized label-এ
    দেওয়া সব সাবস্ট্রিং আছে। একাধিক কলাম মিললে প্রথমটা (labels dict-এর
    insertion order অনুযায়ী, অর্থাৎ বাম থেকে ডান) রিটার্ন করে। এটা exact-key
    match-এর চেয়ে বেশি নমনীয় — কাস্টমারের ফাইলে হেডার টেক্সট মাঝেমধ্যে
    সামান্য বদলায় (যেমন 'Required Carton' vs 'Required CTN' vs
    'Required Carton Quantity') — এই ফাংশন সবগুলো ভ্যারিয়েন্টই ধরতে পারবে।"""
    for key, col in labels.items():
        if all(s in key for s in must_contain):
            return col
    return None


def _extract_size_breakdown(df, sheet_name):
    """একটা 'SIZE BREAKDOWN' শিট থেকে Master Carton লাইন-আইটেম বের করে।
    রিটার্ন করে (line_items, warnings)।"""
    found = _find_size_breakdown_header_row(df)
    if not found:
        return [], []
    header_row, sheet_type, labels = found

    po_col = labels.get('pono')
    # 'Required Carton' / 'Required CTN' / 'Required Carton Quantity' —
    # কাস্টমারের ফাইলভেদে এই কলামের হেডার টেক্সট একটু একটু বদলায়, তাই
    # exact-key না ধরে substring-ভিত্তিক flexible ম্যাচ ব্যবহার করা হচ্ছে।
    qty_col = (
        _find_col_containing(labels, 'required', 'carton')
        or _find_col_containing(labels, 'required', 'ctn')
        or _find_col_containing(labels, 'required')
    )
    wh_col = labels.get('whcode')

    warnings = []

    if sheet_type == 'SOLID':
        color_col = labels.get('color')
        size_col = labels.get('size')
        triplet = _find_lwh_triplet(df, header_row)
        if not triplet or qty_col is None or po_col is None:
            return [], [f"⚠️ শিট '{sheet_name}': SOLID-টাইপ হেডার শনাক্ত হলেও প্রয়োজনীয় কলাম (Length/Width/Height বা Required Carton বা PO NO#) মেলেনি — এই শিট স্কিপ করা হয়েছে।"]
        l_col, w_col, h_col = triplet
    else:  # ASSORT
        set_col = labels.get('set')
        meas_col = labels.get('measurement')
        if set_col is None or meas_col is None or qty_col is None or po_col is None:
            return [], [f"⚠️ শিট '{sheet_name}': ASSORT-টাইপ হেডার শনাক্ত হলেও প্রয়োজনীয় কলাম (SET/MEASUREMENT/Required Carton quantity/PO NO#) মেলেনি — এই শিট স্কিপ করা হয়েছে।"]

    items = []
    total_row_qty = None
    running_po = ''
    running_color = ''
    running_wh = ''

    r = header_row + 1
    n_rows = df.shape[0]
    while r < n_rows:
        # PO/Color/WH Code মাঝেমধ্যে শুধু ব্লকের প্রথম রো-তে থাকে, পরের
        # রো-গুলোয় ফাঁকা থাকে — তাই ফরওয়ার্ড-ফিল করা হচ্ছে।
        po_val = _clean(df.iat[r, po_col]) if po_col is not None else ''
        if po_val:
            running_po = po_val
        if sheet_type == 'SOLID' and color_col is not None:
            color_val = _clean(df.iat[r, color_col])
            if color_val:
                running_color = color_val
        if wh_col is not None:
            wh_val = _clean(df.iat[r, wh_col])
            if wh_val:
                running_wh = wh_val

        qty_val = df.iat[r, qty_col] if qty_col is not None else None

        if sheet_type == 'SOLID':
            style_marker = _clean(df.iat[r, size_col]) if size_col is not None else ''
        else:
            style_marker = _clean(df.iat[r, set_col]) if set_col is not None else ''

        is_data_row = bool(style_marker) and _is_num(qty_val)

        if not style_marker and _is_num(qty_val) and not is_data_row:
            # সম্ভাব্য "Total" রো (style/size ফাঁকা কিন্তু qty কলামে একটা
            # সংখ্যা আছে) — এখান থেকেই টোটাল qty ধরে cross-check করা হবে।
            # তবে এটা যেন Assort-এর 'Quantity per carton#'-এর মতো অন্য কোনো
            # সংখ্যার সাথে গুলিয়ে না যায়, তাই শুধু qty_col-এর মান নেওয়া হচ্ছে।
            if total_row_qty is None:
                total_row_qty = _num(qty_val)

        if not is_data_row:
            r += 1
            continue

        req_qty = _num(qty_val) or 0
        if req_qty <= 0:
            # ইউজার-কনফার্মড নিয়ম: Required Carton Quantity 0/ফাঁকা হলে
            # এই রো একদম বাদ।
            r += 1
            continue

        if sheet_type == 'SOLID':
            length = _fmt_num(df.iat[r, l_col])
            width = _fmt_num(df.iat[r, w_col])
            height = _fmt_num(df.iat[r, h_col])
            gmt_color = running_color
            gmt_size = style_marker
            style_no = 'N/A'
            pack_type = 'SOLID'
        else:
            meas_text = _clean(df.iat[r, meas_col])
            length, width, height = _parse_two_or_three_nums(meas_text)
            gmt_color = 'N/A'
            gmt_size = 'N/A'
            style_no = style_marker  # SET কোড
            pack_type = 'ASSORT'

        # ইউজার-কনফার্মড: EWO No কলামে কিছু বসবে না, সবসময় N/A।
        ewo_no = 'N/A'

        items.append({
            'item_name': 'Master Carton',
            'ewo_no': ewo_no,
            'style_no': style_no,
            'po_no': running_po or 'N/A',
            'length': length,
            'width': width,
            'height': height,
            'ply': '5',
            'qty': req_qty,
            'pack_type': pack_type,
            'reference': running_wh or 'N/A',
            'remarks': '',
            'color': gmt_color,
            'size': gmt_size,
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_sheet': sheet_name,
        })
        r += 1

    extracted_total = sum(it['qty'] for it in items)
    if total_row_qty is not None and abs(total_row_qty - extracted_total) > 0.001:
        warnings.append(
            f"⚠️ শিট '{sheet_name}': Total রো-তে Required Carton Qty = {total_row_qty:g}, "
            f"কিন্তু বের করা লাইন-আইটেমগুলোর মোট Qty = {extracted_total:g} — পার্থক্য "
            f"{total_row_qty - extracted_total:g}, ভালোভাবে চেক করে নিন।"
        )

    return items, warnings


def _classify_indent_item(desc):
    d = (desc or '').lower()
    if 'divider' in d:
        return 'Divider'
    if 'top' in d and 'bottom' in d:
        return 'Top Bottom'
    return None  # Master Carton বা অচেনা — স্কিপ


def _extract_ply(desc):
    m = re.search(r'(\d+)\s*f?\s*ply', (desc or '').lower())
    return m.group(1) if m else ''


def _extract_indent(df, sheet_name):
    """'Indent for Carton' শিট থেকে শুধু Top Bottom / Divider লাইন-আইটেম
    বের করে (Master Carton রো স্কিপ, কারণ সেটা Size Breakdown শিট থেকেই
    নেওয়া হয়েছে)। রিটার্ন করে (line_items, warnings)।"""
    header_row = None
    labels = None
    for r in range(min(df.shape[0], 10)):
        lm = _row_label_map(df, r)
        if 'descriptionply' in lm and 'specification' in lm and 'requiredqty' in lm:
            header_row, labels = r, lm
            break
    if header_row is None:
        return [], []

    desc_col = labels['descriptionply']
    spec_col = labels['specification']
    qty_col = labels['requiredqty']

    # PO নম্বর: হেডার রো-এর ঠিক নিচের রো-তে, Description/Ply কলামেই থাকে
    po_no = _clean(df.iat[header_row + 1, desc_col]) if header_row + 1 < df.shape[0] else ''

    items = []
    r = header_row + 2  # +1 হেডারের পরের PO-নম্বর রো, তারপর থেকে আসল ডাটা
    n_rows = df.shape[0]
    while r < n_rows:
        desc = _clean(df.iat[r, desc_col])
        spec = _clean(df.iat[r, spec_col])
        qty_val = df.iat[r, qty_col]

        if not desc and not spec:
            r += 1
            continue

        item_name = _classify_indent_item(desc)
        if item_name is None:
            r += 1
            continue  # Master Carton বা অচেনা রো — স্কিপ

        if not _is_num(qty_val) or _num(qty_val) <= 0:
            r += 1
            continue  # qty 0/ফাঁকা — বাদ

        ply = _extract_ply(desc)
        length, width, _h = _parse_two_or_three_nums(spec)

        items.append({
            'item_name': item_name,
            'ewo_no': 'N/A',
            'style_no': 'N/A',
            'po_no': po_no or 'N/A',
            'length': length,
            'width': width,
            'height': '',  # Divider/Top Bottom-এ Height থাকে না (builder.py-এর height-exempt লজিক এটা হ্যান্ডেল করবে)
            'ply': ply,
            'qty': _num(qty_val),
            'pack_type': 'N/A',
            'reference': 'N/A',
            'remarks': '',
            'color': 'N/A',
            'size': 'N/A',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_sheet': sheet_name,
        })
        r += 1

    return items, []


def read_amigo_booking_file(file_stream, filename):
    """একটা .xls/.xlsx Amigo/Uniqlo বুকিং ফাইল থেকে সব লাইন-আইটেম বের করে।
    রিটার্ন করে (size_breakdown_items, indent_items, warnings) — দুইটা
    গ্রুপ আলাদা রাখা হচ্ছে যাতে combine_amigo_booking_files সব ফাইলের
    Size Breakdown (Master Carton) আগে বসিয়ে, সব ফাইলের Indent (Top
    Bottom/Divider) একেবারে শেষে বসাতে পারে।"""
    file_stream.seek(0)
    sheets = pd.read_excel(file_stream, sheet_name=None, header=None)
    visible_names = _get_visible_sheet_names(file_stream, filename)

    sb_items = []
    indent_items = []
    all_warnings = []

    for sheet_name, df in sheets.items():
        if visible_names is not None and sheet_name not in visible_names:
            continue
        norm_name = _norm(sheet_name)
        if norm_name == 'sizebreakdown':
            items, warns = _extract_size_breakdown(df, sheet_name)
            sb_items.extend(items)
            all_warnings.extend(warns)
        elif norm_name == 'indentforcarton':
            items, warns = _extract_indent(df, sheet_name)
            indent_items.extend(items)
            all_warnings.extend(warns)

    if not sb_items and not indent_items:
        all_warnings.append(f"⚠️ '{filename}': কোনো লাইন-আইটেম পাওয়া যায়নি (পরিচিত ফরম্যাট না হতে পারে)।")

    return sb_items, indent_items, all_warnings


def combine_amigo_booking_files(file_tuples):
    """একাধিক Amigo বুকিং ফাইল (SOLID/ASSORT, যতগুলোই আসুক — ২টা বা ৪টা)
    কম্বাইন করে। ইউজার-কনফার্মড অর্ডারিং নিয়ম: সব ফাইলের Size Breakdown
    (Master Carton) লাইন-আইটেমগুলো আগে, তারপর সব ফাইলের Indent (Top
    Bottom/Divider) লাইন-আইটেমগুলো একেবারে সবার শেষে।
    file_tuples: [(BytesIO, filename), ...]
    রিটার্ন করে (line_items, warnings)।"""
    all_sb_items = []
    all_indent_items = []
    all_warnings = []
    for file_stream, filename in file_tuples:
        sb_items, indent_items, warns = read_amigo_booking_file(file_stream, filename)
        for it in sb_items:
            it['_source_file'] = filename
        for it in indent_items:
            it['_source_file'] = filename
        all_sb_items.extend(sb_items)
        all_indent_items.extend(indent_items)
        all_warnings.extend(warns)

    combined = all_sb_items + all_indent_items
    return combined, all_warnings