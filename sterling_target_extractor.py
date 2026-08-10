"""
Sterling Styles Limited — Buyer: Target
Carton Order Excel ফরম্যাট ('CARTON ORDER S# ...')।

একটা বুকিং-এ ১টা থেকে ৪-৫টা .xls/.xlsx ফাইল একসাথে আসতে পারে। প্রতিটা
ফাইল একটা নির্দিষ্ট Style-এর জন্য (হেডারের কাছে 'Style: <নম্বর>' লেখা
থাকে) — সেই Style নম্বরটা পুরো ফাইলের সব লাইন-আইটেমে (Master Carton এবং
Divider/Top Bottom দুটোতেই) বসে।

প্রতিটা ফাইলে তিন ধরনের ডাটা-গ্রুপ থাকে (হেডারে পাশাপাশি Measurement +
Qty কলাম জোড়ায়):
  1. Master Carton  — 'Carton Meas. L x W x H cm' + 'Master Carton Order
     Qty (Pcs)' — এটাই একমাত্র breakdown-ওয়াইজ (প্রতিটা রো আলাদা লাইন-
     আইটেম) বসে।
  2. Divider (Supporting Board) — 'SUPPORTING BOARD (NORMAL 5 PLY)
     Measurement' + পাশের Qty কলাম, সবসময় 5 Ply।
  3. Top Bottom — 'TOP BOTTOM (3 PLY) Measurement' + পাশের Qty কলাম,
     সবসময় 3 Ply।
  Divider/Top Bottom breakdown-ওয়াইজ না — একই ফাইলের ভেতরে যতগুলো
  ইউনিক measurement থাকে, প্রতিটার জন্য একটা করে লাইন-আইটেম, আর সেই
  measurement-এর সব রো-এর qty যোগ করে (সামারি) বসে। এই লাইনগুলো Amigo/
  Sinha-র কনভেনশন অনুযায়ী মেইন ডাটার পরে, ফাইলের একেবারে শেষে বসে।

ব্যবসায়িক নিয়ম (ইউজার-কনফার্মড):
- Item Name: বুকিং হেডারের কাছাকাছি কোথাও 'ELASTIC' শব্দ পাওয়া গেলে সব
  Master Carton রো-তে Item Name 'Elastic Hanger Carton' বসে (আর একটা
  warning যোগ হয়, যাতে ম্যানুয়ালি চেক করা যায়) — না পাওয়া গেলে ডিফল্ট
  'Master Carton'। UI-এর Item Name dropdown এখানে প্রযোজ্য না।
- Ply: Master Carton সবসময় 5. Divider সবসময় 5, Top Bottom সবসময় 3 —
  এগুলো হেডারেই লেখা থাকে, তাই হার্ডকোড করা।
- Style No: ফাইলের হেডার থেকে বের করা Style নম্বর — Master Carton এবং
  Divider/Top Bottom সব লাইনেই বসে (Amigo/Sinha-র উল্টো — ওখানে trailer
  লাইনে Style N/A থাকত, এখানে অবশ্যই থাকতে হবে)।
- Reference <- CP ID No কলাম। Color/Size এই বায়ারের জন্য প্রযোজ্য না
  (সবসময় N/A)। PO No/CP ID No ফাঁকা থাকলে N/A বসবে, রো বাদ যাবে না
  (যদি qty ভ্যালিড থাকে)।
- Qty 0/ফাঁকা হলে সেই রো বাদ (Master Carton-এ)। Divider/Top Bottom-এর
  ক্ষেত্রে গ্রুপ-সামারি করার পর টোটাল qty 0 হলে সেই আইটেমই বাদ।
- কোনো ফাইলে Divider/Top Bottom কলাম Excel-এ hidden (column hide) করা
  থাকলে, ভ্যালু থাকলেও সেটা একদম বাদ (শুধু visible কলামের ডাটাই নেওয়া
  হবে)।
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


def _parse_lw(text):
    """'57x37x20 CM' -> ('57','37','20') ; '36x32 CM' -> ('36','32','')।"""
    nums = re.findall(r'(\d+\.?\d*)', text or '')
    l = _fmt_num(nums[0]) if len(nums) > 0 else ''
    w = _fmt_num(nums[1]) if len(nums) > 1 else ''
    h = _fmt_num(nums[2]) if len(nums) > 2 else ''
    return l, w, h


def _get_hidden_columns(file_stream, filename):
    """Excel-এ column-hide করা কলামগুলোর (0-indexed) সেট রিটার্ন করে।
    শনাক্ত করতে না পারলে ফাঁকা সেট রিটার্ন করে (তখন কোনো কলামই hidden
    ধরা হবে না, নিরাপদ ডিফল্ট)।"""
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


_STYLE_RE = re.compile(r'\bstyle\s*:\s*([^\s][\w\-/]*)', re.I)


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
            labels[_norm(t)] = c
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


def read_sterling_booking_file(file_stream, filename):
    """একটা .xls/.xlsx ফাইল থেকে (master_items, trailer_items, warnings)
    বের করে।"""
    file_stream.seek(0)
    df = pd.read_excel(file_stream, sheet_name=0, header=None)
    rows = df.values.tolist()

    header_row, labels = _find_header_row(rows)
    if header_row is None:
        return [], [], [f"⚠️ '{filename}': পরিচিত 'PO No'/'CP ID No' হেডিং-সহ বুকিং টেবিল পাওয়া যায়নি — স্কিপ করা হয়েছে।"]

    po_col = _find_col(labels, 'pono')
    cpid_col = _find_col(labels, 'cpid')
    pack_col = _find_col(labels, 'packtype')
    mc_meas_col = _find_col(labels, 'cartonmeas')
    div_meas_col = _find_col(labels, 'supportingboard')
    tb_meas_col = _find_col(labels, 'topbottom', 'measurement')

    if mc_meas_col is None:
        return [], [], [f"⚠️ '{filename}': Master Carton Measurement কলাম পাওয়া যায়নি — স্কিপ করা হয়েছে।"]

    mc_qty_col = mc_meas_col + 1
    div_qty_col = div_meas_col + 1 if div_meas_col is not None else None
    tb_qty_col = tb_meas_col + 1 if tb_meas_col is not None else None

    hidden_cols = _get_hidden_columns(file_stream, filename)
    warnings = []
    if div_meas_col is not None and (div_meas_col in hidden_cols or div_qty_col in hidden_cols):
        div_meas_col = div_qty_col = None
    if tb_meas_col is not None and (tb_meas_col in hidden_cols or tb_qty_col in hidden_cols):
        tb_meas_col = tb_qty_col = None

    style_no = _extract_style(rows) or 'N/A'
    is_elastic = _detect_elastic(rows, header_row)
    item_name = 'Elastic Hanger Carton' if is_elastic else 'Master Carton'
    if is_elastic:
        warnings.append(
            f"⚠️ '{filename}': বুকিং হেডারের কাছাকাছি 'ELASTIC' শব্দ পাওয়া গেছে — "
            f"Item Name 'Elastic Hanger Carton' বসানো হয়েছে, একবার চেক করে নিন।"
        )

    master_items = []
    divider_qty_by_meas = {}
    topbottom_qty_by_meas = {}

    r = header_row + 1
    n_rows = len(rows)
    while r < n_rows:
        row = rows[r]
        first_cell = _clean(row[0]) if row else ''
        if _norm(first_cell) == 'total':
            break

        mc_qty_val = row[mc_qty_col] if mc_qty_col < len(row) else None
        if _is_num(mc_qty_val) and (_num(mc_qty_val) or 0) > 0:
            length, width, height = _parse_lw(_clean(row[mc_meas_col]) if mc_meas_col < len(row) else '')
            po_val = _clean(row[po_col]) if po_col is not None and po_col < len(row) else ''
            cpid_val = _clean(row[cpid_col]) if cpid_col is not None and cpid_col < len(row) else ''
            pack_val = _clean(row[pack_col]) if pack_col is not None and pack_col < len(row) else ''
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

        if div_meas_col is not None:
            div_meas_val = _clean(row[div_meas_col]) if div_meas_col < len(row) else ''
            div_qty_val = row[div_qty_col] if div_qty_col < len(row) else None
            if div_meas_val and _is_num(div_qty_val):
                divider_qty_by_meas[div_meas_val] = divider_qty_by_meas.get(div_meas_val, 0) + (_num(div_qty_val) or 0)

        if tb_meas_col is not None:
            tb_meas_val = _clean(row[tb_meas_col]) if tb_meas_col < len(row) else ''
            tb_qty_val = row[tb_qty_col] if tb_qty_col < len(row) else None
            if tb_meas_val and _is_num(tb_qty_val):
                topbottom_qty_by_meas[tb_meas_val] = topbottom_qty_by_meas.get(tb_meas_val, 0) + (_num(tb_qty_val) or 0)

        r += 1

    trailer_items = []
    for meas_text, total_qty in divider_qty_by_meas.items():
        if total_qty <= 0:
            continue
        length, width, _h = _parse_lw(meas_text)
        trailer_items.append({
            'item_name': 'Divider',
            'ewo_no': 'N/A',
            'style_no': style_no,
            'po_no': 'N/A',
            'length': length,
            'width': width,
            'height': '',
            'ply': '5',
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
    for meas_text, total_qty in topbottom_qty_by_meas.items():
        if total_qty <= 0:
            continue
        length, width, _h = _parse_lw(meas_text)
        trailer_items.append({
            'item_name': 'Top Bottom',
            'ewo_no': 'N/A',
            'style_no': style_no,
            'po_no': 'N/A',
            'length': length,
            'width': width,
            'height': '',
            'ply': '3',
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
