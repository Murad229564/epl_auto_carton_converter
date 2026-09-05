"""
আউট হাউজ Carton — 'Multiple Job Wise Trims Booking' / 'Main Trims Booking'
স্টাইলের PDF ফরম্যাট (Barnali, Modele de Capital, URMI/Fakhruddin ইত্যাদি)।

প্রতিটা PDF-এ একাধিক "ব্লক" থাকে — প্রতিটা ব্লক একটা নির্দিষ্ট PO-এর জন্য:
  - একটা টাইটেল-রো (যেমন "As Per Garments Color (Job NO:...) Style NO:...
    Po Qty.: ... Po No: ... Shipment Date: ...", বা "Color & size sensitive
    (...)...", বা "NO sensitive (...)...")
  - একটা হেডার-রো (Sl, Item Group, Item Description, ... — ফরম্যাট ভেদে
    কলাম-সেট আলাদা হতে পারে)
  - এক বা একাধিক ডাটা-রো
  - একটা "Item Total" রো
  - একটা "Total" রো (ব্লক শেষ)

⚠️ গুরুত্বপূর্ণ ফিক্স (পেজ-ব্রেক বাগ): যখন কোনো ব্লকের ডাটা-রো (বিশেষ করে
'Item Group' কলামের লম্বা মাল্টি-লাইন টেক্সট, যেমন "Top Bottom\\nCzech\\n
Republic,Germany,...UAE") পাতার শেষে গিয়ে পরের পাতায় চলে যায়, তখন
pdfplumber প্রতিটা পাতা আলাদাভাবে পড়ে বলে সেই একটা রো **দুই টুকরায় ভেঙে
যায়** — প্রথম পাতায় Sl/Item Group ফাঁকা (None) হয়ে যায়, পরের পাতায়
Item Group-এর বাকি অংশ 'Item Total' রো-এর সাথে মিশে অদ্ভুত রো হয়ে যায়,
এবং কখনো কখনো continuation রো-তে একটা এক্সট্রা ফাঁকা কলামও ঢুকে গিয়ে
বাকি কলামগুলো ডানদিকে শিফট হয়ে যায়। আগের কোড পাতা-ভিত্তিক আলাদাভাবে
প্রসেস করত বলে এই ভাঙা রো-গুলোর আসল কোয়ান্টিটি ডাটা হারিয়ে যাচ্ছিল।

সমাধান:
  1. পাতা-ভিত্তিক আলাদা না করে, PDF-এর সব পাতার সব টেবিল একসাথে ফ্ল্যাট
     করে নেওয়া হয় প্রথমে।
  2. ব্লক-বাউন্ডারি টেবিল/পাতার সীমানা দিয়ে না, বরং নতুন টাইটেল-রো বা
     "Total" রো দিয়ে ঠিক হয় — তাই ব্লকের ডাটা টেবিল/পাতার সীমানায় ভেঙে
     গেলেও রো-স্ক্যান থামে না, পরের টেবিল/পাতা থেকেও রো টেনে আনতে থাকে।
  3. "Item Group" (Item Name-এর উৎস) কোনো রো-তে ফাঁকা পেলে প্রথমে একই
     ব্লকের আগের রো থেকে forward-fill, সেটাও না পেলে পুরো ডকুমেন্ট জুড়ে
     সবচেয়ে বেশিবার পাওয়া ভ্যালু (একটা ফাইলে সাধারণত একটাই আইটেম-টাইপ
     থাকে) ফলব্যাক হিসেবে ব্যবহার হয়।
  4. "Item Total" রো-এর টেক্সট মাঝেমধ্যে দুই সেলে ভেঙে যায় (যেমন
     'Item Tota' + 'l 88.0000') — সব সেল জোড়া লাগিয়ে সঠিকভাবে চেনা হয়।
  5. কলাম-ম্যাপিং রো-এর **শেষ থেকে দূরত্ব (distance-from-end)** দিয়ে করা
     হয়, শুরু থেকে ইনডেক্স দিয়ে না — কারণ কিছু ভাঙা continuation রো-তে
     শুরুর দিকে একটা এক্সট্রা ফাঁকা কলাম ঢুকে যায় (পুরো রো ডানে শিফট হয়ে
     যায়), কিন্তু qty/UOM/Rate/Amount কলামগুলো সবসময় রো-এর শেষের দিকে
     নিজেদের আপেক্ষিক অবস্থানে ঠিকই থাকে — তাই শেষ থেকে গুনলে শিফট হওয়া
     রো-তেও সঠিক কলাম মেলে।
  6. Measurement (L x W [x H] CM) আর Qty কলামের নাম/অবস্থান ফরম্যাট-ভেদে
     ভিন্ন — regex + label-ভিত্তিক flexible ম্যাচিং দিয়ে বের করা হয়।
"""
import re
import pdfplumber


def _norm(v):
    return re.sub(r'\s+', '', str(v or '')).lower()


def _clean(v):
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def _is_num(v):
    if v is None:
        return False
    try:
        float(str(v).replace(',', '').strip())
        return True
    except (TypeError, ValueError):
        return False


def _num(v):
    try:
        return float(str(v).replace(',', '').strip())
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


_MEAS_RE = re.compile(r'(\d+\.?\d*)\s*[xX]\s*(\d+\.?\d*)(?:\s*[xX]\s*(\d+\.?\d*))?\s*C?M?\b')


def _find_measurement(row_cells):
    """পুরো রো-এর সব সেলে measurement প্যাটার্ন (L x W [x H] [C]M) খোঁজে —
    ফরম্যাট-ভেদে এটা 'Item Description' বা 'Item Size' যেকোনো কলামে
    থাকতে পারে, তাই fixed কলাম ধরা হচ্ছে না।"""
    for cell in row_cells:
        text = _clean(cell)
        if not text:
            continue
        m = _MEAS_RE.search(text)
        if m:
            l = _fmt_num(m.group(1))
            w = _fmt_num(m.group(2))
            h = _fmt_num(m.group(3)) if m.group(3) else ''
            return l, w, h
    return '', '', ''


_TITLE_STYLE_RE = re.compile(r'Style\s*NO\s*:\s*([^\s]+)', re.I)
_TITLE_PONO_RE = re.compile(
    r'Po\s*No\s*:\s*(.+?)(?:\s+Shipment\s*Date\s*:|\s+LC\s*/\s*SC\s*:?\s*$|$)', re.I)
_TITLE_JOBNO_RE = re.compile(r'Job\s*NO\s*:\s*([^\)]+)\)', re.I)


def _is_title_row(row):
    first = _clean(row[0]) if row else ''
    n = _norm(first)
    return 'styleno' in n and 'pono' in n


def _parse_title_row(row):
    text = _clean(row[0])
    style_m = _TITLE_STYLE_RE.search(text)
    po_m = _TITLE_PONO_RE.search(text)
    return {
        'style_no': style_m.group(1).strip() if style_m else '',
        'po_no': po_m.group(1).strip().rstrip(',') if po_m else '',
    }


def _is_header_row(row):
    return bool(row) and _norm(row[0]) == 'sl'


_HEADER_LABELS = {
    'itemgroup': 'item_group',
    'itemdescription': 'item_description',
    'brandsupplierref.': 'brand_ref',
    'brandsupplierref': 'brand_ref',
    'itemcolor': 'item_color',
    'gmtscolor': 'gmts_color',
    'gmtssize': 'gmts_size',
    'itemsize': 'item_size',
    'woqty': 'qty',
    'woqty.': 'qty',
    'qnty': 'qty',
    'uom': 'uom',
    'rate': 'rate',
    'amount': 'amount',
    'remarks': 'remarks',
}


def _build_header_map(row):
    """হেডার-রো থেকে প্রতিটা কলামের 'শেষ থেকে দূরত্ব' হিসেব করে রাখে —
    দেখুন উপরের মডিউল-ডকস্ট্রিং পয়েন্ট ৫।"""
    n = len(row)
    col_map = {}
    for c, cell in enumerate(row):
        label = _norm(cell)
        if label in _HEADER_LABELS:
            key = _HEADER_LABELS[label]
            if key not in col_map:
                col_map[key] = n - 1 - c
    return col_map


def _resolve_col(header_map, key, row):
    if key not in header_map:
        return None
    idx = len(row) - 1 - header_map[key]
    if 0 <= idx < len(row):
        return idx
    return None


def _joined_row_text(row):
    return _norm(''.join(_clean(c) for c in row if c is not None))


def _is_item_total_row(row):
    return 'itemtotal' in _joined_row_text(row)


def _is_block_total_row(row):
    return bool(row) and _norm(row[0]) == 'total'


def _is_document_footer_start(row):
    """ডকুমেন্টের একদম নিচে একটা আলাদা সামারি টেবিল থাকে (Item Name |
    PO Number | PO Qty | WO Qty | Rate | Amount, আর তার আগে 'Total Booking
    Amount' লাইন) — এটা আসল ব্লক-ডাটা না, তাই এখান থেকে শুরু করে বাকি সব
    রো একদম বাদ (নাহলে একই qty দ্বিতীয়/তৃতীয়বার ভুল করে গোনা হয়ে যায়)।"""
    if not row:
        return False
    first = _norm(row[0])
    if 'totalbookingamount' in first:
        return True
    normed = [_norm(c) for c in row]
    if 'itemname' in normed and 'ponumber' in normed:
        return True
    return False


def _classify_item(item_group_text):
    t = (item_group_text or '').lower()
    if 'top' in t and 'bottom' in t:
        return 'Top Bottom', '3', False
    return 'Master Carton', '5', True


def process_trims_booking_pdf(file_stream, customer_list=None, buyer_list=None):
    """মূল entry point। রিটার্ন করে (header_info, line_items)। এই ফরম্যাট না
    হলে (কোনো টাইটেল-রো না পাওয়া গেলে) header_info-তে সব ফাঁকা আর
    line_items=[] রিটার্ন করে।"""
    file_stream.seek(0)
    all_rows = []
    with pdfplumber.open(file_stream) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                all_rows.extend(table)

    # ধাপ ১: ডকুমেন্ট-জুড়ে সবচেয়ে বেশি পাওয়া Item Group ভ্যালু বের করা —
    # page-break-এ কোনো রো তার নিজের Item Group সম্পূর্ণ হারিয়ে ফেললেও
    # এটাই শেষ ভরসা হিসেবে ব্যবহার হবে।
    item_group_counts = {}
    header_map_probe = None
    for row in all_rows:
        if _is_document_footer_start(row):
            break
        if _is_header_row(row):
            header_map_probe = _build_header_map(row)
            continue
        if not header_map_probe or 'item_group' not in header_map_probe:
            continue
        ig_col = _resolve_col(header_map_probe, 'item_group', row)
        if ig_col is not None:
            val = _clean(row[ig_col])
            first_line = val.split('\n')[0].strip() if val else ''
            if first_line:
                item_group_counts[first_line] = item_group_counts.get(first_line, 0) + 1
    dominant_item_group = max(item_group_counts, key=item_group_counts.get) if item_group_counts else ''

    # ধাপ ২: রো-স্ক্যান
    line_items = []
    header_map = {}
    current_style_no = ''
    current_po_no = ''
    current_item_group = ''
    doc_po_number = ''

    for row in all_rows:
        if not row or all(c is None or _clean(c) == '' for c in row):
            continue

        if _is_document_footer_start(row):
            break  # নিচের সামারি/স্বাক্ষর সেকশন — এখান থেকে আর কিছু প্রসেস করা হবে না

        if _is_title_row(row):
            parsed = _parse_title_row(row)
            current_style_no = parsed['style_no'] or current_style_no
            current_po_no = parsed['po_no']
            current_item_group = ''  # নতুন ব্লক শুরু — রিসেট
            if not doc_po_number and current_po_no:
                doc_po_number = current_po_no
            continue

        if _is_header_row(row):
            new_map = _build_header_map(row)
            if new_map:
                header_map = new_map
            continue

        if _is_item_total_row(row):
            continue

        if _is_block_total_row(row):
            current_item_group = ''
            continue

        if not header_map or 'qty' not in header_map:
            continue

        qty_col = _resolve_col(header_map, 'qty', row)
        qty_val = row[qty_col] if qty_col is not None else None
        if not _is_num(qty_val):
            continue

        qty = _num(qty_val)
        if not qty or qty <= 0:
            continue

        ig_col = _resolve_col(header_map, 'item_group', row)
        row_item_group = _clean(row[ig_col]).split('\n')[0].strip() if ig_col is not None else ''
        if row_item_group:
            current_item_group = row_item_group
        effective_item_group = current_item_group or dominant_item_group

        item_name, ply, has_height = _classify_item(effective_item_group)
        length, width, height = _find_measurement(row)
        if not has_height:
            height = ''

        brand_col = _resolve_col(header_map, 'brand_ref', row)
        color_col = _resolve_col(header_map, 'item_color', row)
        if color_col is None:
            color_col = _resolve_col(header_map, 'gmts_color', row)
        gmts_size_col = _resolve_col(header_map, 'gmts_size', row)

        reference_val = _clean(row[brand_col]) if brand_col is not None else ''
        pack_type_val = _clean(row[gmts_size_col]) if gmts_size_col is not None else ''
        color_val = _clean(row[color_col]) if color_col is not None else ''

        line_items.append({
            'item_name': item_name,
            'ewo_no': 'N/A',
            'style_no': current_style_no or 'N/A',
            'po_no': current_po_no or 'N/A',
            'length': length,
            'width': width,
            'height': height,
            'ply': ply,
            'qty': qty,
            'pack_type': pack_type_val or 'N/A',
            'reference': reference_val or 'N/A',
            'remarks': '',
            'color': color_val or 'N/A',
            'size': 'N/A',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
        })

    if not line_items:
        return {'po_number': '', 'customer': '', 'buyer': ''}, []

    file_stream.seek(0)
    try:
        with pdfplumber.open(file_stream) as pdf:
            first_page_text = pdf.pages[0].extract_text() or ''
    except Exception:
        first_page_text = ''

    doc_buyer = ''
    buyer_m = re.search(r'Buyer\.?\s*:\s*([^\n]+)', first_page_text)
    if buyer_m:
        doc_buyer = buyer_m.group(1).strip()
    doc_customer = first_page_text.split('\n')[0].strip() if first_page_text else ''

    header_info = {
        'po_number': doc_po_number or '',
        'customer': doc_customer or '',
        'buyer': doc_buyer or '',
    }
    return header_info, line_items