"""
Everbright Sweater Ltd. — Buyer: Dunnes Stores
'Carton (work order)' Excel ফরম্যাট।

একটা বুকিং-এ ১টা ফাইল, বা একসাথে একাধিক (৪-৫টা) ফাইল আসতে পারে। একটা
ফাইলে একাধিক শিট থাকতে পারে (প্রতিটা শিট থেকেই ডাটা নেওয়া হয়)।

⚠️ এই ফরম্যাট Amigo/Sinha/Sterling-এর মতো "সামারি করে ফাইলের শেষে" না —
বরং Columbia (GU buyer)-স্টাইল breakdown-ওয়াইজ: প্রতিটা রো-তে Master
Carton এবং (থাকলে) তার নিজস্ব Divider — দুটোই আলাদা লাইন-আইটেম হিসেবে,
সেই রো-এর নিজস্ব Style/PO(Item Description)/Reference(Color)/Pack Type
(Gmt Size) সহ, একই জায়গায় (একটার পরে একটা) বসে।

কলাম-লেআউট ফাইল-ভেদে একটু শিফট হয় (কখনো 'Pack Number' কলাম থাকে, কখনো
থাকে না; হেডার টেক্সট কখনো 'CARTON Order Qty', কখনো শুধু 'Order Qty') —
তাই সব কলাম header-label স্ক্যান করে ডাইনামিকভাবে বের করা হয়, fixed
index ধরা হয় না।

ব্যবসায়িক নিয়ম (ইউজার-কনফার্মড):
- Item Name (Master Carton-এর জন্য): UI থেকে সিলেক্ট করা ভ্যালু বসে
  (item_name_override)। Divider-এর Item Name সবসময় ফিক্সড 'Divider'
  (ফাইলে 'TOP BOTTOM BORD' লেখা থাকলেও)।
- Ply: Master Carton সবসময় 5, Divider সবসময় 3 — দুটোই হার্ডকোড, UI
  Ply-সিলেকশন এখানে প্রযোজ্য না।
- Style <- 'Style' কলাম। GMT PO <- 'Item Description' কলাম। Reference
  <- 'Colors' কলাম। Pack Type <- 'Gmt Size' কলাম। Color/Gmt Size (কালার/
  সাইজ) ফাঁকা হলে N/A।
- Qty (Master Carton বা Divider, যেটাই) 0/ফাঁকা হলে সেই লাইন বাদ।
- Excel-এ hidden করা রো বা কলাম (দৃশ্যমান না) — সেই ডাটা একদম বাদ, ভ্যালু
  থাকলেও নেওয়া হবে না।
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


_LWH_RE = re.compile(r'L\s*(\d+\.?\d*).*?W\s*(\d+\.?\d*).*?H\s*(\d+\.?\d*)', re.I | re.S)
_LW_RE = re.compile(r'(\d+\.?\d*)\s*CM?\s*[xX×]\s*(\d+\.?\d*)\s*CM?', re.I)


def _parse_lwh(text):
    """'L 60 X W 32 X H 13 CM' -> ('60','32','13')।"""
    t = text or ''
    m = _LWH_RE.search(t)
    if m:
        return _fmt_num(m.group(1)), _fmt_num(m.group(2)), _fmt_num(m.group(3))
    return '', '', ''


def _parse_lw(text):
    """'58 CM X 30 CM' -> ('58','30')।"""
    t = text or ''
    m = _LW_RE.search(t)
    if m:
        return _fmt_num(m.group(1)), _fmt_num(m.group(2))
    return '', ''


def _get_hidden_rows_cols(file_stream, filename, sheet_name=None, sheet_index=0):
    """Excel-এ hidden করা রো এবং কলামের (0-indexed) সেট রিটার্ন করে —
    দুটোই একসাথে (hidden_rows, hidden_cols)। শনাক্ত করতে না পারলে দুটোই
    ফাঁকা সেট (নিরাপদ ডিফল্ট — কিছুই hidden ধরা হবে না)।"""
    try:
        file_stream.seek(0)
        if filename.lower().endswith('.xls'):
            import xlrd
            book = xlrd.open_workbook(file_contents=file_stream.read(), formatting_info=True)
            sheet = book.sheet_by_name(sheet_name) if sheet_name else book.sheet_by_index(sheet_index)
            hidden_rows = {r for r, info in sheet.rowinfo_map.items() if getattr(info, 'hidden', 0)}
            hidden_cols = {c for c, info in sheet.colinfo_map.items() if getattr(info, 'hidden', 0)}
            return hidden_rows, hidden_cols
        else:
            from openpyxl import load_workbook as _load_wb
            wb = _load_wb(file_stream, read_only=False, data_only=True)
            ws = wb[sheet_name] if sheet_name else wb.worksheets[sheet_index]
            hidden_rows = {r - 1 for r, dim in ws.row_dimensions.items() if dim.hidden}
            from openpyxl.utils import column_index_from_string
            hidden_cols = {
                column_index_from_string(letter) - 1
                for letter, dim in ws.column_dimensions.items() if dim.hidden
            }
            wb.close()
            return hidden_rows, hidden_cols
    except Exception:
        return set(), set()
    finally:
        file_stream.seek(0)


def _get_visible_sheet_names(file_stream, filename):
    try:
        file_stream.seek(0)
        if filename.lower().endswith('.xls'):
            import xlrd
            book = xlrd.open_workbook(file_contents=file_stream.read())
            return [name for name in book.sheet_names() if book.sheet_by_name(name).visibility == 0]
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
        if 'style' in labels and any('ctnmmts' in k for k in labels):
            return i, labels
    return None, None


def read_everbright_booking_sheet(rows, sheet_name, hidden_rows, hidden_cols, item_name_override=''):
    """একটা শিটের rows (list-of-lists) থেকে লাইন-আইটেম বের করে। রিটার্ন
    করে (line_items, warnings)।"""
    header_row, labels = _find_header_row(rows)
    if header_row is None:
        return [], []

    style_col = labels.get('style')
    desc_col = _find_col(labels, 'itemdescription')
    color_col = _find_col(labels, 'colors') or _find_col(labels, 'color')
    size_col = _find_col(labels, 'gmtsize')
    mc_meas_col = _find_col(labels, 'ctnmmts')
    mc_qty_col = _find_col(labels, 'orderqty')
    tb_meas_col = None
    tb_qty_col = None
    for key, col in labels.items():
        if 'topbottombord' in key and 'order' not in key:
            tb_meas_col = col
        if 'topbottombord' in key and 'order' in key:
            tb_qty_col = col

    if style_col is None or mc_meas_col is None or mc_qty_col is None:
        return [], [f"⚠️ শিট '{sheet_name}': Style/Ctn Mmts/Order Qty কলাম পাওয়া যায়নি — স্কিপ করা হয়েছে।"]

    # Divider কলাম hidden থাকলে (measurement বা qty যেকোনো একটা), পুরো
    # ডিভাইডার ডাটাই বাদ — ইউজার-কনফার্মড নিয়ম।
    if tb_meas_col is not None and (tb_meas_col in hidden_cols or (tb_qty_col is not None and tb_qty_col in hidden_cols)):
        tb_meas_col = tb_qty_col = None

    items = []
    r = header_row + 1
    # হেডারের ঠিক পরের রো-তে (row19-এর মতো) মাঝেমধ্যে দ্বিতীয় সাব-লেবেল
    # রো ('Quantity'/'U.S. Dollars') থাকে, ওটায় কোনো ডাটা-রো লক্ষণ (Style
    # কলামে কিছু) থাকে না — তাই এটা এমনিতেই নিচের while-loop-এ স্কিপ
    # হয়ে যাবে (style_val ফাঁকা থাকবে)।
    n_rows = len(rows)
    while r < n_rows:
        if r in hidden_rows:
            r += 1
            continue

        row = rows[r]
        style_val = _clean(row[style_col]) if style_col < len(row) else ''
        if _norm(style_val) == 'gtotal' or (style_val and style_val.upper().startswith('G. TOTAL')):
            break
        if not style_val:
            r += 1
            continue

        desc_val = _clean(row[desc_col]) if desc_col is not None and desc_col < len(row) else ''
        color_val = _clean(row[color_col]) if color_col is not None and color_col < len(row) else ''
        size_val = _clean(row[size_col]) if size_col is not None and size_col < len(row) else ''

        mc_qty_val = row[mc_qty_col] if mc_qty_col < len(row) else None
        if _is_num(mc_qty_val) and (_num(mc_qty_val) or 0) > 0:
            length, width, height = _parse_lwh(_clean(row[mc_meas_col]) if mc_meas_col < len(row) else '')
            items.append({
                'item_name': item_name_override or 'Master Carton',
                'ewo_no': 'N/A',
                'style_no': style_val,
                'po_no': desc_val or 'N/A',
                'length': length,
                'width': width,
                'height': height,
                'ply': '5',
                'qty': round(_num(mc_qty_val)),
                'pack_type': size_val or 'N/A',
                'reference': color_val or 'N/A',
                'remarks': '',
                'color': 'N/A',
                'size': 'N/A',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': sheet_name,
            })

        if tb_meas_col is not None:
            tb_qty_val = row[tb_qty_col] if tb_qty_col < len(row) else None
            if _is_num(tb_qty_val) and (_num(tb_qty_val) or 0) > 0:
                length, width = _parse_lw(_clean(row[tb_meas_col]) if tb_meas_col < len(row) else '')
                items.append({
                    'item_name': 'Divider',
                    'ewo_no': 'N/A',
                    'style_no': style_val,
                    'po_no': desc_val or 'N/A',
                    'length': length,
                    'width': width,
                    'height': '',
                    'ply': '3',
                    'qty': round(_num(tb_qty_val)),
                    'pack_type': size_val or 'N/A',
                    'reference': color_val or 'N/A',
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


def read_everbright_booking_file(file_stream, filename, item_name_override=''):
    """একটা .xls/.xlsx ফাইলের সব (visible) শিট থেকে ডাটা বের করে।
    রিটার্ন করে (line_items, warnings)।"""
    file_stream.seek(0)
    sheets = pd.read_excel(file_stream, sheet_name=None, header=None)
    visible_names = _get_visible_sheet_names(file_stream, filename)

    all_items = []
    all_warnings = []
    for idx, (sheet_name, df) in enumerate(sheets.items()):
        if visible_names is not None and sheet_name not in visible_names:
            continue
        hidden_rows, hidden_cols = _get_hidden_rows_cols(file_stream, filename, sheet_name=sheet_name)
        items, warns = read_everbright_booking_sheet(
            df.values.tolist(), sheet_name, hidden_rows, hidden_cols,
            item_name_override=item_name_override)
        all_items.extend(items)
        all_warnings.extend(warns)

    if not all_items:
        all_warnings.append(f"⚠️ '{filename}': কোনো ভ্যালিড (qty>0) লাইন-আইটেম পাওয়া যায়নি।")

    return all_items, all_warnings


def combine_everbright_booking_files(files, item_name_override='', manual_ply=''):
    """files: [(BytesIO, filename), ...] — BATCH_REGISTRY-এর uniform কল-
    সিগনেচার। manual_ply ইচ্ছাকৃতভাবে ব্যবহার হচ্ছে না (Ply এখানে সবসময়
    ফিক্সড: Master Carton=5, Divider=3)। item_name_override ব্যবহার হয়
    শুধু Master Carton-এর Item Name-এর জন্য (Divider সবসময় 'Divider')।

    এই ফরম্যাট Amigo/Sinha/Sterling-এর মতো সামারি-করে-শেষে-বসানো না —
    প্রতিটা রো-এর Master Carton আর তার নিজস্ব Divider পাশাপাশি (একই
    ক্রমে) বসে, তাই এখানে আলাদা 'trailer' গ্রুপিং নেই।"""
    combined = []
    all_warnings = []
    for file_stream, filename in files:
        items, warns = read_everbright_booking_file(file_stream, filename, item_name_override=item_name_override)
        for it in items:
            it['_source_file'] = filename
        combined.extend(items)
        all_warnings.extend(warns)
    return combined, all_warnings
