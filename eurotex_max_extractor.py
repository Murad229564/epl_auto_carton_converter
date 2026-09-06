import re
import openpyxl
import pandas as pd

# ---------------------------------------------------------------------------
# Eurotex Knitwear Ltd. — Buyer: MAX
# Carton Booking Excel ফরম্যাট। batch-স্টাইল extractor (Amigo/Sinha/Sterling-
# এর মতোই uniform সিগনেচার: files -> (line_items, warnings)), কারণ
# ক্রস-ফাইল/ক্রস-শিট অর্ডারিং লাগে — সবগুলো ফাইল/শিটের Master Carton
# আইটেম আগে, তারপর সবগুলোর 'Top & Bottom' সামারি-আইটেম সবার শেষে।
#
# ফরম্যাটের বৈশিষ্ট্য (ইউজার-কনফার্মড):
#   - একটা ফাইলে এক বা একাধিক শিট থাকতে পারে (প্রতিটা শিটের একই লেআউট,
#     সাধারণত শিটের নামেই কোন কোন PO আছে বোঝা যায় — যেমন
#     'CKSS271039-CKSS271040')।
#   - হেডারের উপরে কোথাও 'Ship To : <নাম>' লেখা থাকে। নাম দেওয়া থাকলে
#     (যেমন 'Ship To : EUROTEX') সেই নামটাই পুরো শিটের সব রো-এর জন্য
#     Remarks-এ বসবে। নাম ফাঁকা থাকলে ('Ship To : ' খালি — যেমন প্রথম
#     স্যাম্পল ফাইলে), শিটের ভেতরেই একটা 'DELIVERY STATUS'-জাতীয় কলাম
#     থাকে যেখানে প্রতিটা PO-ব্লকের প্রথম রো-তেই শুধু একটা নাম লেখা থাকে
#     (merged সেলের কারণে বাকি রো ফাঁকা) — সেই কলাম থেকে forward-fill করে
#     প্রতিটা রো-এর নিজের ডেলিভারি-নাম বের করা হয়। যেভাবেই পাওয়া যাক, এই
#     নামটাই Remarks কলামে বসে (ইউজার পরে ফিল্টার করে আলাদা করতে পারবেন)।
#   - কলাম ম্যাপিং: PO NUMBER -> Gmt. PO, STYLE NO -> Gmt. Style No,
#     COLOR -> Reference/SKU Number, TYPE OF CARTON (SOLID/ASSORT) ->
#     Pack Type, CARTON MEASUREMENT -> Length/Width/Height, CARTON
#     QUANTITY -> Order Qty। Quantity-তে ফ্র্যাকশন থাকলে রাউন্ড করে নেওয়া
#     হয় (Master Carton আর Top & Bottom দুটোতেই)।
#   - Item Name ডিফল্ট 'Master Carton', কিন্তু রো-এর যেকোনো কলামে (Remarks-
#     জাতীয় কলাম-সহ) 'ELASTIC' বা 'HANGER' শব্দ পাওয়া গেলে 'Elastic Hanger
#     Carton' বসে।
#   - প্রতিটা শিটের একদম নিচে একটা 'TOP & BOTTOM' সামারি-রো থাকে (প্রতি-PO
#     আলাদা না, পুরো শিটের জন্য একটাই) — Measurement-এ শুধু Length x Width
#     (Height থাকে না), Quantity পুরো শিটের Master Carton টোটালের ডাবল
#     হওয়ার কথা। এগুলো সবগুলো ফাইল/শিটের Master Carton বসানোর পরে, সবার
#     শেষে একসাথে বসানো হয়।
#   - হাইড করা রো/কলাম সম্পূর্ণ বাদ (যেমন এক ফাইলের হাইড করা 'DIVIDER TOP'
#     রো) — শুধু visible ডাটা থেকেই লাইন-আইটেম তৈরি হয়।
#   - Ply এই ফরম্যাটে কোথাও থাকে না, তাই UI থেকে দেওয়া manual_ply সব
#     রো-তে (Master Carton আর Top & Bottom দুটোতেই) বসে।
#   - প্রতিটা শিটে Master Carton টোটাল Qty-এর ডাবল-এর সাথে Top & Bottom
#     Qty না মিললে, ঠিক কোন ফাইলের কোন শিটে কত পিসের মিসম্যাচ — সেটা
#     উল্লেখ করে warning দেওয়া হয়।
# ---------------------------------------------------------------------------


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


def _round_qty(v):
    n = _num(v)
    if n is None:
        return ''
    return int(round(n))


def _fmt_num(n):
    f = float(n)
    return str(int(f)) if f == int(f) else str(round(f, 3))


def _parse_two_or_three_nums(text):
    """'54X34X29 CM' -> ('54','34','29') ; '40 X 30   CM' -> ('40','30','')।"""
    nums = re.findall(r'(\d+\.?\d*)', text or '')
    l = _fmt_num(nums[0]) if len(nums) > 0 else ''
    w = _fmt_num(nums[1]) if len(nums) > 1 else ''
    h = _fmt_num(nums[2]) if len(nums) > 2 else ''
    return l, w, h


HEADER_TOKEN_MAP = {
    _norm('PO NUMBER'): 'po_no',
    _norm('STYLE NO'): 'style_no',
    _norm('COLOR'): 'color',
    _norm('TYPE OF CARTON'): 'pack_type',
    _norm('CARTON MEASUREMENT'): 'measurement',
    _norm('CARTON QUANTITY'): 'qty',
}


def _find_header_row(ws, max_scan=15):
    """প্রথম max_scan রো-র মধ্যে যেই রো-তে HEADER_TOKEN_MAP-এর সবগুলো
    টোকেনই একসাথে পাওয়া যাবে, সেটাকেই হেডার রো ধরা হয় — fixed row-position
    ধরা হয় না (ফাইল-ভেদে ৬ বা ৮ নম্বর রো-তে থাকতে পারে)। রিটার্ন করে
    (header_row_idx, col_map, delivery_status_col) অথবা কিছু না পেলে
    (None, {}, None)। delivery_status_col-এর হেডিং বানান ফাইল-ভেদে একটু
    ভিন্ন হতে পারে ('DELIVERY  STATUS' — বাড়তি স্পেস-সহ), তাই normalize
    করার পর শুধু 'delivery' আর 'status' দুটো শব্দই একসাথে থাকলেই ধরা হয়।"""
    required = set(HEADER_TOKEN_MAP.values())
    for r in range(1, max_scan + 1):
        col_map = {}
        delivery_status_col = None
        for c in range(1, ws.max_column + 1):
            v = _clean(ws.cell(row=r, column=c).value)
            if not v:
                continue
            key = HEADER_TOKEN_MAP.get(_norm(v))
            if key:
                col_map[key] = c
                continue
            nv = _norm(v)
            if 'delivery' in nv and 'status' in nv:
                delivery_status_col = c
        if required.issubset(col_map.keys()):
            return r, col_map, delivery_status_col
    return None, {}, None


def _find_sheet_ship_to(ws, header_row):
    """হেডার রো-এর উপরে 'Ship To :' লেখা সেল খুঁজে তার পরের নামটা বের করে।
    নাম ফাঁকা থাকলে (colon-এর পরে কিছু নেই) '' রিটার্ন করে — তখন কলার
    ডেলিভারি-স্ট্যাটাস কলাম থেকে per-row নেবে।"""
    for r in range(1, header_row):
        for c in range(1, ws.max_column + 1):
            v = _clean(ws.cell(row=r, column=c).value)
            if v.lower().startswith('ship to'):
                return v.split(':', 1)[1].strip() if ':' in v else ''
    return ''


def _get_hidden_rows_cols(ws):
    """এই শিটের ভেতরে হাইড করা রো/কলামের index (1-indexed, openpyxl-এর
    সাথে সামঞ্জস্যপূর্ণ) বের করে।"""
    from openpyxl.utils import column_index_from_string
    hidden_rows = {r for r, d in ws.row_dimensions.items() if d.hidden}
    hidden_cols = set()
    for col_letter, d in ws.column_dimensions.items():
        if d.hidden:
            try:
                hidden_cols.add(column_index_from_string(col_letter))
            except ValueError:
                continue
    return hidden_rows, hidden_cols


def _row_has_elastic_hanger(ws, row, n_cols, hidden_cols):
    for c in range(1, n_cols + 1):
        if c in hidden_cols:
            continue
        v = _clean(ws.cell(row=row, column=c).value).lower()
        if 'elastic' in v or 'hanger' in v:
            return True
    return False


def _process_sheet(ws, filename, item_name_default):
    """একটা visible শিট থেকে (master_items, tb_items, warnings) বের করে।
    tb_items শিট-প্রতি ০ বা ১টা 'Top & Bottom' সামারি-আইটেম থাকতে পারে।"""
    header_row, col_map, delivery_status_col = _find_header_row(ws)
    if header_row is None:
        return [], [], []

    hidden_rows, hidden_cols = _get_hidden_rows_cols(ws)
    sheet_ship_to = _find_sheet_ship_to(ws, header_row)

    master_items = []
    tb_items = []
    warnings = []
    running_status = ''  # ডেলিভারি-স্ট্যাটাস কলামের forward-fill ভ্যালু
    n_cols = ws.max_column

    for r in range(header_row + 1, ws.max_row + 1):
        if r in hidden_rows:
            continue

        po_val = _clean(ws.cell(row=r, column=col_map['po_no']).value)
        qty_val = ws.cell(row=r, column=col_map['qty']).value

        if delivery_status_col is not None:
            status_val = _clean(ws.cell(row=r, column=delivery_status_col).value)
            if status_val:
                running_status = status_val

        po_norm = _norm(po_val)

        if 'top' in po_norm and 'bottom' in po_norm:
            # শিটের একদম নিচের 'TOP & BOTTOM' সামারি-রো — প্রতি-PO আলাদা না,
            # পুরো শিটের জন্য একটাই।
            length, width, _h = _parse_two_or_three_nums(
                _clean(ws.cell(row=r, column=col_map['measurement']).value))
            tb_items.append({
                'item_name': 'Top Bottom',
                'ewo_no': 'N/A',
                'style_no': 'N/A',
                'po_no': 'N/A',
                'length': length,
                'width': width,
                'height': '',
                'ply': '',
                'qty': _round_qty(qty_val),
                'pack_type': 'N/A',
                'reference': 'N/A',
                'remarks': sheet_ship_to or '',
                'color': 'N/A',
                'size': 'N/A',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': ws.title,
                '_source_file': filename,
            })
            continue

        if po_norm == 'totalqty':
            continue  # শুধু cross-check রেফারেন্সের জন্য, লাইন-আইটেম না

        if not po_val or not _is_num(qty_val):
            continue  # ফাঁকা/অচেনা রো — স্কিপ

        style_no = _clean(ws.cell(row=r, column=col_map['style_no']).value)
        color_val = _clean(ws.cell(row=r, column=col_map['color']).value)
        pack_type_val = _clean(ws.cell(row=r, column=col_map['pack_type']).value)
        measurement_val = _clean(ws.cell(row=r, column=col_map['measurement']).value)
        length, width, height = _parse_two_or_three_nums(measurement_val)

        remarks_val = running_status if delivery_status_col is not None else sheet_ship_to

        item_name = 'Elastic Hanger Carton' if _row_has_elastic_hanger(
            ws, r, n_cols, hidden_cols) else (item_name_default or 'Master Carton')

        master_items.append({
            'item_name': item_name,
            'ewo_no': 'N/A',
            'style_no': style_no or 'N/A',
            'po_no': po_val or 'N/A',
            'length': length,
            'width': width,
            'height': height,
            'ply': '',
            'qty': _round_qty(qty_val),
            'pack_type': pack_type_val or 'N/A',
            'reference': color_val or 'N/A',
            'remarks': remarks_val,
            'color': 'N/A',
            'size': 'N/A',
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_sheet': ws.title,
            '_source_file': filename,
        })

    if tb_items:
        master_total = sum(it['qty'] for it in master_items if _is_num(it['qty']))
        tb_total = sum(it['qty'] for it in tb_items if _is_num(it['qty']))
        expected = master_total * 2
        diff = expected - tb_total
        if abs(diff) > 0.001:
            warnings.append(
                f"⚠️ ফাইল '{filename}', শিট '{ws.title}': Master Carton-এর মোট Qty "
                f"{master_total:g} (ডাবল = {expected:g}), কিন্তু Top & Bottom-এর Qty "
                f"পাওয়া গেছে {tb_total:g} — পার্থক্য {diff:g} পিস। ভালোভাবে চেক করে নিন।"
            )

    return master_items, tb_items, warnings


def combine_eurotex_max_booking_files(files, item_name_override='', manual_ply=''):
    """মূল entry point। files: [(file_stream, filename), ...]।

    প্রতিটা ফাইলের প্রতিটা visible শিট থেকে Master Carton আইটেম নেওয়া হয়
    (হাইড রো/কলাম বাদে), আর প্রতিটা শিটের 'Top & Bottom' সামারি-রো আলাদাভাবে
    জমা রাখা হয়। ফাইনাল লিস্টে সবগুলো ফাইল/শিটের Master Carton আইটেম আগে,
    তারপর সবগুলোর Top & Bottom আইটেম সবার শেষে — এই অর্ডারেই রিটার্ন হয়
    (ইউজার-কনফার্মড কনভেনশন)।

    item_name_override: 'ELASTIC'/'HANGER' শব্দ না পাওয়া গেলে ডিফল্ট আইটেম
    নাম হিসেবে ব্যবহার হয় (ফাঁকা থাকলে 'Master Carton')।
    manual_ply: এই ফরম্যাটে Ply কোথাও থাকে না, তাই UI থেকে দেওয়া মান সব
    রো-তে (Master Carton + Top & Bottom) বসে; ফাঁকা থাকলে 'N/A'।

    Returns (combined_line_items, warnings) — uniform batch-signature,
    outhouse_extractor.py-এর BATCH_REGISTRY-তে ব্যবহারের জন্য।
    """
    all_master, all_tb, all_warnings = [], [], []

    for file_stream, filename in files:
        try:
            file_stream.seek(0)
            wb = openpyxl.load_workbook(file_stream, data_only=True)
        except Exception as e:
            all_warnings.append(f"⚠️ '{filename}': ফাইল পড়তে সমস্যা হয়েছে — {e}")
            continue

        visible_sheets = [ws for ws in wb.worksheets if ws.sheet_state == 'visible']
        if not visible_sheets:
            all_warnings.append(f"⚠️ '{filename}': কোনো visible শিট পাওয়া যায়নি।")
            wb.close()
            continue

        found_any = False
        for ws in visible_sheets:
            master_items, tb_items, warns = _process_sheet(ws, filename, item_name_override)
            if master_items or tb_items:
                found_any = True
            all_master.extend(master_items)
            all_tb.extend(tb_items)
            all_warnings.extend(warns)
        wb.close()

        if not found_any:
            all_warnings.append(
                f"⚠️ '{filename}': পরিচিত ফরম্যাট (PO NUMBER/STYLE NO/COLOR/TYPE OF "
                f"CARTON/CARTON MEASUREMENT/CARTON QUANTITY হেডিং) পাওয়া যায়নি।"
            )

    ply_val = manual_ply.strip() if manual_ply else 'N/A'
    for it in all_master:
        it['ply'] = ply_val
    for it in all_tb:
        it['ply'] = ply_val

    return all_master + all_tb, all_warnings
