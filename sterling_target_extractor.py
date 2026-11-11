"""
Sterling Styles Limited — Buyer: Target
Carton Order Excel ফরম্যাট ('CARTON ORDER S# ...' / 'Carton booking for TA ...')।

একটা বুকিং-এ ১টা থেকে ৪-৫টা .xls/.xlsx ফাইল একসাথে আসতে পারে। প্রতিটা
ফাইল একটা নির্দিষ্ট Style-এর জন্য — সেই Style নম্বরটা পুরো ফাইলের সব
লাইন-আইটেমে (Master Carton এবং Divider/Top Bottom দুটোতেই) বসে।

--- Style নম্বর (দুই রকম লেখা সাপোর্ট করে) ---
  'Style: 2900330454'  (কোলন)   অথবা   'Style#2900330454 PO-2659395'  (হ্যাশ)
  হেডারের উপরের অংশে খোঁজা হয়; না পেলে টেবিলের 'STYLE NO#' কলামের প্রথম
  ভ্যালু; তাও না পেলে 'N/A'।

--- তিন ধরনের ডাটা-গ্রুপ (হেডিং দেখে চেনা হয়, কলামের অবস্থান নয়) ---
  1. Master Carton — 'Carton Meas. L x W x H cm' + 'Carton Order Qty (Pcs)'
     / 'Master Carton Order Qty (Pcs)' — প্রতিটা রো আলাদা লাইন-আইটেম, সবসময়
     5 Ply।
  2. Divider — হেডিংয়ে 'SUPPORTING BOARD (NORMAL 5 PLY)' / '5 PLY BROAD (BOARD)'
     / 'Divider' থাকলে। সবসময় 5 Ply।
  3. Top Bottom — হেডিংয়ে 'TOP BOTTOM (3 PLY)' / '3 PLY BROAD (BOARD)' / '3 ply
     board' থাকলে। সবসময় 3 Ply।

  Divider/Top Bottom-এর মেজারমেন্ট দুইভাবে আসতে পারে (দুটোই সামলানো হয়):
    (ক) আলাদা মেজারমেন্ট কলাম + তার পাশের Qty কলাম (রো-ভেদে মেজারমেন্ট
        আলাদা হতে পারে) — যেমন '5 PLY BROAD MMC' + '5 PLY BROAD QNTY'।
    (খ) মেজারমেন্ট হেডিংয়ের ভেতরেই লেখা, কলামে শুধু Qty — যেমন
        '3 ply board (58*50)' বা 'SUPPORTING BOARD ... qnty (58*50)'। তখন
        ওই কলামের সব Qty-র জন্য হেডিংয়ের মেজারমেন্টই ধরা হয়।
  Divider/Top Bottom breakdown-ওয়াইজ না — একই ফাইলের ভেতরে যতগুলো ইউনিক
  measurement থাকে, প্রতিটার জন্য একটা করে লাইন-আইটেম, আর সেই
  measurement-এর সব রো-এর qty যোগ করে (সামারি) বসে। এই লাইনগুলো মেইন
  ডাটার পরে, সব ফাইলের একেবারে শেষে বসে।
  Qty-তে সংখ্যা না থাকলে (যেমন 'N/A') বা মোট 0 হলে সেই আইটেম বাদ।

ব্যবসায়িক নিয়ম (ইউজার-কনফার্মড):
- Item Name: বুকিং হেডারের কাছাকাছি কোথাও 'ELASTIC' শব্দ পাওয়া গেলে সব
  Master Carton রো-তে Item Name 'Elastic Hanger Carton' বসে (আর একটা
  warning যোগ হয়) — না পাওয়া গেলে ডিফল্ট 'Master Carton'। UI-এর Item Name
  dropdown এখানে প্রযোজ্য না।
- Ply: Master Carton সবসময় 5, Divider সবসময় 5, Top Bottom সবসময় 3।
- Reference <- CP ID No কলাম। Pack Type <- PACK TYPE কলাম। Color/Size সবসময়
  N/A। PO No/CP ID No ফাঁকা থাকলে N/A বসবে, রো বাদ যাবে না (qty ভ্যালিড হলে)।
- PO/CP ID সংখ্যা হলে '2659395.0' না, '2659395' বসে।
- Master Carton-এ Qty 0/ফাঁকা হলে সেই রো বাদ।
- কোনো ফাইলে Divider/Top Bottom কলাম Excel-এ hidden থাকলে সেটা বাদ।
- 'Total' রো পেলে টেবিল শেষ; সেই রো-র যোগফলের সাথে মিলিয়ে Warning দেওয়া
  হয় (Master Carton, Divider, Top Bottom)।
"""
import re
import pandas as pd


def _norm(v):
    return re.sub(r'[^a-z0-9]', '', str(v or '').lower())


def _clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def _clean_id(v):
    """PO/CP ID/Style-এর জন্য — পূর্ণসংখ্যা-মানের float হলে '.0' বাদ।"""
    if isinstance(v, float) and not pd.isna(v) and v == int(v):
        return str(int(v))
    return _clean(v)


def _is_num(v):
    if v is None or isinstance(v, bool) or (isinstance(v, float) and pd.isna(v)):
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


def _parse_lw(text):
    """'57x37x20 CM' -> ('57','37','20') ; '36x32 CM' -> ('36','32','')।"""
    nums = re.findall(r'(\d+\.?\d*)', text or '')
    l = _fmt_num(nums[0]) if len(nums) > 0 else ''
    w = _fmt_num(nums[1]) if len(nums) > 1 else ''
    h = _fmt_num(nums[2]) if len(nums) > 2 else ''
    return l, w, h


def _load_rows(file_stream, filename):
    # প্রোজেক্টের মাল্টি-ইঞ্জিন রিডার (xlrd/calamine/LibreOffice ফলব্যাক)
    # থাকলে সেটাই; না থাকলে সরাসরি pandas।
    try:
        from outhouse_extractor import _read_excel_rows
        return _read_excel_rows(file_stream, filename)
    except ImportError:
        file_stream.seek(0)
        return pd.read_excel(file_stream, sheet_name=0, header=None).values.tolist()


def _get_hidden_columns(file_stream, filename):
    """Excel-এ column-hide করা কলামগুলোর (0-indexed) সেট রিটার্ন করে।
    শনাক্ত করতে না পারলে ফাঁকা সেট রিটার্ন করে (নিরাপদ ডিফল্ট)।"""
    try:
        file_stream.seek(0)
        if filename.lower().endswith('.xls'):
            import xlrd
            book = xlrd.open_workbook(file_contents=file_stream.read(), formatting_info=True)
            sheet = book.sheet_by_index(0)
            hidden = set()
            for c, info in getattr(sheet, 'colinfo_map', {}).items():
                if getattr(info, 'hidden', 0):
                    hidden.add(c)
            return hidden
        else:
            from openpyxl import load_workbook as _load_wb
            wb = _load_wb(file_stream, read_only=False, data_only=True)
            ws = wb.worksheets[0]
            hidden = set()
            for letter, dim in ws.column_dimensions.items():
                if dim.hidden:
                    from openpyxl.utils import column_index_from_string
                    hidden.add(column_index_from_string(letter) - 1)
            wb.close()
            return hidden
    except Exception:
        return set()
    finally:
        file_stream.seek(0)


# 'Style: 123' (কোলন) এবং 'Style#123' / 'Style # 123' (হ্যাশ) দুটোই
_STYLE_RE = re.compile(r'\bstyle\s*[:#]\s*([^\s][\w\-/]*)', re.I)


def _extract_style(rows, max_scan=10):
    for r in rows[:max_scan]:
        for cell in r:
            text = _clean(cell)
            if not text:
                continue
            m = _STYLE_RE.search(text)
            if m:
                return m.group(1).strip()
    return ''


def _detect_elastic(rows, header_row):
    for r in rows[:header_row]:
        for cell in r:
            if 'elastic' in str(cell or '').lower():
                return True
    return False


def _row_label_map(row):
    labels = {}
    for c, v in enumerate(row):
        t = _clean(v)
        if t:
            labels.setdefault(_norm(t), c)
    return labels


def _find_header_row(rows, max_scan=15):
    for i, row in enumerate(rows[:max_scan]):
        labels = _row_label_map(row)
        if any('pono' in k for k in labels) and any('cpid' in k for k in labels):
            return i, labels
    return None, None


def _find_col(labels, *must_contain):
    for key, col in labels.items():
        if all(s in key for s in must_contain):
            return col
    return None


# মেজারমেন্ট হেডিংয়ের ভেতরে লেখা থাকলে ('(58*50)', '58 x 50', '58*50*3')
_EMBEDDED_MEAS_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*[*xX×]\s*(\d+(?:\.\d+)?)(?:\s*[*xX×]\s*(\d+(?:\.\d+)?))?')


def _classify_headers(header_row_cells):
    """হেডার-রো-র প্রতিটা সেল দেখে গ্রুপ ঠিক করে। রিটার্ন:
    {'master': {'meas': col, 'qty': col},
     'divider': {'meas': col|None, 'qty': col|None, 'fixed': 'LxW'|None},
     'topbottom': {...}}"""
    groups = {
        'master': {'meas': None, 'qty': None},
        'divider': {'meas': None, 'qty': None, 'fixed': None},
        'topbottom': {'meas': None, 'qty': None, 'fixed': None},
    }
    for c, cell in enumerate(header_row_cells):
        text = _clean(cell)
        if not text:
            continue
        n = _norm(text)
        if 'cartonmeas' in n:
            groups['master']['meas'] = groups['master']['meas'] if groups['master']['meas'] is not None else c
            continue
        if 'cartonorderqty' in n or ('mastercarton' in n and ('qty' in n or 'qnty' in n)):
            if groups['master']['qty'] is None:
                groups['master']['qty'] = c
            continue
        if 'topbottom' in n or '3ply' in n:
            g = groups['topbottom']
        elif 'supportingboard' in n or '5ply' in n or 'divider' in n:
            g = groups['divider']
        else:
            continue

        emb = _EMBEDDED_MEAS_RE.search(text)
        has_qty_word = ('qty' in n) or ('qnty' in n) or ('quantity' in n)
        if emb:
            # (খ) মেজারমেন্ট হেডিংয়ে, এই কলামটাই Qty কলাম
            if g['qty'] is None:
                g['qty'] = c
                g['fixed'] = f"{emb.group(1)}x{emb.group(2)}"
        elif has_qty_word:
            if g['qty'] is None:
                g['qty'] = c
        else:
            # (ক) মেজারমেন্ট কলাম (Measurement / MSRT / MMC ...)
            if g['meas'] is None:
                g['meas'] = c
    return groups


def read_sterling_booking_file(file_stream, filename):
    """একটা .xls/.xlsx ফাইল থেকে (master_items, trailer_items, warnings)
    বের করে।"""
    rows = _load_rows(file_stream, filename)

    header_row, labels = _find_header_row(rows)
    if header_row is None:
        return [], [], [f"⚠️ '{filename}': পরিচিত 'PO No'/'CP ID No' হেডিং-সহ বুকিং টেবিল পাওয়া যায়নি — স্কিপ করা হয়েছে।"]

    po_col = _find_col(labels, 'pono')
    cpid_col = _find_col(labels, 'cpid')
    pack_col = _find_col(labels, 'packtype')
    style_col = _find_col(labels, 'styleno')

    groups = _classify_headers(rows[header_row])
    mc_meas_col = groups['master']['meas']
    if mc_meas_col is None:
        return [], [], [f"⚠️ '{filename}': Master Carton Measurement কলাম পাওয়া যায়নি — স্কিপ করা হয়েছে।"]
    mc_qty_col = groups['master']['qty'] if groups['master']['qty'] is not None else mc_meas_col + 1

    # Divider/Top Bottom: (meas_col|None, qty_col, fixed_meas|None)
    def resolve(g):
        if g['qty'] is not None and g['fixed']:
            return None, g['qty'], g['fixed']
        if g['meas'] is not None:
            qty = g['qty'] if g['qty'] is not None else g['meas'] + 1
            return g['meas'], qty, None
        return None

    div = resolve(groups['divider'])
    tb = resolve(groups['topbottom'])

    hidden_cols = _get_hidden_columns(file_stream, filename)
    warnings = []
    if div is not None and ((div[0] is not None and div[0] in hidden_cols) or div[1] in hidden_cols):
        div = None
    if tb is not None and ((tb[0] is not None and tb[0] in hidden_cols) or tb[1] in hidden_cols):
        tb = None

    style_no = _extract_style(rows)
    if not style_no and style_col is not None:
        for row in rows[header_row + 1:]:
            if style_col < len(row) and _clean_id(row[style_col]):
                style_no = _clean_id(row[style_col])
                break
    style_no = style_no or 'N/A'

    is_elastic = _detect_elastic(rows, header_row)
    item_name = 'Elastic Hanger Carton' if is_elastic else 'Master Carton'
    if is_elastic:
        warnings.append(
            f"⚠️ '{filename}': বুকিং হেডারের কাছাকাছি 'ELASTIC' শব্দ পাওয়া গেছে — "
            f"Item Name 'Elastic Hanger Carton' বসানো হয়েছে, একবার চেক করে নিন।"
        )

    def cell(row, c):
        return row[c] if c is not None and c < len(row) else None

    master_items = []
    divider_qty_by_meas = {}    # (l, w) -> qty  (ঢোকানোর ক্রম বজায় থাকে)
    topbottom_qty_by_meas = {}
    sum_master = 0
    totals = {}

    def accumulate(spec, row, bucket):
        if spec is None:
            return
        meas_col, qty_col, fixed = spec
        qty_val = cell(row, qty_col)
        if not _is_num(qty_val):
            return
        meas_text = fixed if fixed else _clean(cell(row, meas_col))
        if not meas_text:
            return
        l, w, _h = _parse_lw(meas_text)
        key = (l, w)
        bucket[key] = bucket.get(key, 0) + (_num(qty_val) or 0)

    r = header_row + 1
    n_rows = len(rows)
    while r < n_rows:
        row = rows[r]
        first_cell = _clean(row[0]) if row else ''
        if _norm(first_cell) == 'total':
            totals['master'] = _num(cell(row, mc_qty_col))
            if div is not None:
                totals['divider'] = _num(cell(row, div[1]))
            if tb is not None:
                totals['topbottom'] = _num(cell(row, tb[1]))
            break

        mc_qty_val = cell(row, mc_qty_col)
        if _is_num(mc_qty_val) and (_num(mc_qty_val) or 0) > 0:
            length, width, height = _parse_lw(_clean(cell(row, mc_meas_col)))
            po_val = _clean_id(cell(row, po_col)) if po_col is not None else ''
            cpid_val = _clean_id(cell(row, cpid_col)) if cpid_col is not None else ''
            pack_val = _clean(cell(row, pack_col)) if pack_col is not None else ''
            master_items.append({
                'item_name': item_name,
                'ewo_no': 'N/A',
                'style_no': style_no,
                'po_no': po_val or 'N/A',
                'length': length,
                'width': width,
                'height': height,
                'ply': '5',
                'qty': round(_num(mc_qty_val)),
                'pack_type': pack_val or 'N/A',
                'reference': cpid_val or 'N/A',
                'remarks': '',
                'color': 'N/A',
                'size': 'N/A',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': filename,
            })
            sum_master += round(_num(mc_qty_val))

        accumulate(div, row, divider_qty_by_meas)
        accumulate(tb, row, topbottom_qty_by_meas)
        r += 1

    def build_trailer(bucket, name, ply):
        out = []
        for (length, width), total_qty in bucket.items():
            if total_qty <= 0:
                continue
            out.append({
                'item_name': name,
                'ewo_no': 'N/A',
                'style_no': style_no,
                'po_no': 'N/A',
                'length': length,
                'width': width,
                'height': '',
                'ply': ply,
                'qty': round(total_qty),
                'pack_type': 'N/A',
                'reference': 'N/A',
                'remarks': '',
                'color': 'N/A',
                'size': 'N/A',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': filename,
            })
        return out

    trailer_items = (build_trailer(divider_qty_by_meas, 'Divider', '5')
                     + build_trailer(topbottom_qty_by_meas, 'Top Bottom', '3'))

    # 'Total' রো-র সাথে যাচাই
    def check(label, got, expected):
        if expected is not None and round(got) != round(expected):
            warnings.append(f"⚠️ '{filename}': {label} যোগফল ({round(got)}) ফাইলের Total ({round(expected)})-এর সাথে মেলেনি।")

    check('Master Carton', sum_master, totals.get('master'))
    check('Divider', sum(divider_qty_by_meas.values()), totals.get('divider'))
    check('Top Bottom', sum(topbottom_qty_by_meas.values()), totals.get('topbottom'))

    if not master_items and not trailer_items:
        warnings.append(f"⚠️ '{filename}': কোনো ভ্যালিড (qty>0) লাইন-আইটেম পাওয়া যায়নি।")

    return master_items, trailer_items, warnings


def combine_sterling_booking_files(files, item_name_override='', manual_ply=''):
    """files: [(BytesIO, filename), ...] — BATCH_REGISTRY-এর uniform কল-
    সিগনেচার (item_name_override/manual_ply এখানে ইচ্ছাকৃতভাবে ব্যবহার
    হচ্ছে না, কারণ এই কাস্টমারের Item Name/Ply সম্পূর্ণ ফাইলের নিজস্ব
    কনটেন্ট থেকেই বের হয়)। রিটার্ন করে (line_items, warnings) — সব
    ফাইলের Master Carton আগে, তারপর সব ফাইলের Divider/Top Bottom সবার
    শেষে (Amigo/Sinha কনভেনশন)।"""
    all_master = []
    all_trailer = []
    all_warnings = []
    for file_stream, filename in files:
        master_items, trailer_items, warns = read_sterling_booking_file(file_stream, filename)
        for it in master_items:
            it['_source_file'] = filename
        for it in trailer_items:
            it['_source_file'] = filename
        all_master.extend(master_items)
        all_trailer.extend(trailer_items)
        all_warnings.extend(warns)

    combined = all_master + all_trailer
    return combined, all_warnings