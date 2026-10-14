"""
Dhaka Garments And Washing Ltd. — Buyer: Kohl`s
Carton Booking Excel ফরম্যাট।

একটা ফাইলে একাধিক শিট থাকতে পারে (BULK/E-com/Pre-pack ইত্যাদি নামে, অথবা PO
নম্বর নামে, কখনো একই নামের একাধিক কপি যেমন 'Bulk (2)', 'E-com (2)') — প্রতিটা
শিট থেকেই ডাটা নেওয়া হয়।

দুইটা ফরম্যাট সাপোর্ট করা হয় (অটো-ডিটেক্ট):

────────── ফরম্যাট A (আগের ফরম্যাট) ──────────
- উপরের দিকে (rows ~1-24) generic ছাপা instruction টেক্সট — উপেক্ষা করা হয়।
- একটা 'PO' মার্কার-রো (কলাম B/col2-এ ঠিক 'PO' লেখা) — এই রো-তেই PO নম্বর
  (তার ঠিক পাশের কলামে) আর Item Name-এর ইঙ্গিত টেক্সট থাকে ('Normal carton' /
  'Elastic  Hanger Carton') — টেক্সটের কলাম-পজিশন শিট-ভেদে বদলায়, তাই পুরো রো
  স্ক্যান করে বের করা হয়।
- তার ঠিক পরের রো-তে আসল হেডার: STYLE / UPC NUMBER / COLOR / সাইজ-কলাম /
  CTN MEASUREMENT (এক কলামে 'L: x W: y H: z') / CARTON BOOKING QTY / তারপর
  'Do not follow' কলামগুলো (GMT/QTY, ACTUAL, QTY/x%, RATIO)।
- শেষে 'G.TTL' রো।

────────── ফরম্যাট B (নতুন ফরম্যাট — Style GJ51D103 E-COM ইত্যাদি) ──────────
- 'PO' মার্কার-রো নেই। উপরের দিকে 'PO NUMBER : 16838570' (লেবেল একটা কলামে, ভ্যালু
  ': '-এর পরের কলামে) — এখান থেকে PO নেওয়া হয়। না পেলে শিটের নাম যদি শুধু সংখ্যা
  হয় (এই ফাইলে শিটের নামই PO নম্বর) সেটাই PO।
- হেডার-রো খোঁজা হয় এভাবে: যে রো-তে 'STYLE' আর 'CARTON QTY' (বা 'CARTON BOOKING
  QTY') দুটোই আছে।
- হেডার: STYLE / ITEM UPC NUMBER / COLOR / সাইজ-কলাম / L / H / H (তিনটা আলাদা
  কলামে — প্রথম 'H' আসলে Width, দ্বিতীয় 'H' Height) / একটা সাইজ+কার্টন-কোড কলাম
  (যেমন '4 (40340)' — ব্যবহার হয় না) / CARTON QTY / GMT/QTY, ACTUAL, QTY/2%, RATIO
  (এগুলো 'Do not follow', ব্যবহার হয় না)। কলাম-পজিশন শিট-ভেদে এক কলাম সরে যেতে
  পারে (এই ফাইলে কিছু শিটে সরেছে) — তাই সবকিছু হেডার-লেবেল দেখে ধরা হয়।
- Measurement-এর ইউনিট 'MEASUREMENT In Inch' ক্যাপশন থেকে নেওয়া হয়।
- 'G.TTL' রো-তে ডাটা শেষ — এর পরের রো-গুলো (carton sticker নোট, মেজারমেন্ট
  অনুযায়ী সামারি টোটাল ইত্যাদি) ডাটা না, তাই 'G.TTL'-এ এসে থেমে যাওয়া হয়।

ব্যবসায়িক নিয়ম (ইউজার-কনফার্মড):
- Item Name (Elastic ডিটেকশন negation-aware):
    • 'elastic' লেখা আছে এবং আগে 'no' / 'no need' / 'not' / 'without' / 'non'
      ইত্যাদি নেই (যেমন 'Elastic Hanger Carton')  -> 'Elastic Hanger Carton'
    • 'NO ELASTIC' / '(NO NEED ELASTIC)' / 'without elastic' / 'non-elastic'
      অথবা 'normal' লেখা                           -> 'Master Carton'
  ধাপ ১: ফরম্যাট A-তে PO-মার্কার-রো, ফরম্যাট B-তে হেডারের উপরের/G.TTL-এর নিচের
          নোট-রো স্ক্যান।
  ধাপ ২: ধাপ ১-এ কিছু না পেলে পুরো শিটের টেক্সটে শুধু 'elastic' শব্দ খোঁজা
          (positive -> Elastic Hanger Carton, negated -> Master Carton)।
  ধাপ ৩: তাও না পেলে শিটের নাম দেখে (trailing '(2)'/'(3)' বাদ দিয়ে): নাম
          'E-com'-এর সাথে মিললে Master Carton, নাহলে Elastic Hanger Carton।
- Ply: UI থেকে সিলেক্ট করা মান থাকলে সেটাই, না থাকলে ডিফল্ট 5।
- Style <- 'STYLE' কলাম। GMT PO <- PO নম্বর (পুরো শিটের জন্য একটাই, প্রতি রো-তে
  বসবে)। Reference <- 'COLOR' কলাম। Pack Type <- 'UPC NUMBER' কলাম। Measurement <-
  ফরম্যাট A: 'CTN MEASUREMENT' কলাম (L:/W:/H: প্রিফিক্স-সহ বা ছাড়া, দুটোই
  সামলানো); ফরম্যাট B: L, W, H তিন কলাম। Qty <- 'CARTON BOOKING QTY' / 'CARTON
  QTY' কলাম (GMT/QTY, ACTUAL, RATIO ইত্যাদি 'Do not follow' কলাম ব্যবহার হয় না)।
- Qty 0/ফাঁকা হলে সেই রো বাদ (আগের সব কাস্টমারের কনভেনশন অনুযায়ী)।
"""
import re
import pandas as pd


def _norm(v):
    return re.sub(r'[^a-z0-9]', '', str(v or '').lower())


def _clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    # 16838570.0 -> '16838570' (Excel-এর সংখ্যা float হয়ে এলে '.0' যেন না বসে)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
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
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if pd.isna(f) else f


def _fmt_num(v):
    if v is None:
        return ''
    try:
        f = float(v)
    except (TypeError, ValueError):
        return ''
    if pd.isna(f):
        return ''
    return str(int(f)) if f == int(f) else str(round(f, 3))


# 'L: 19.5 X W: 14.5 X H: 3 INCH' বা 'L:19.5 X W:14.5 X H:3 CM' বা প্রিফিক্স
# ছাড়া শুধু '19.5 X 14.5 X 3' — যেকোনো ফরম্যাটই সামলাতে শুধু সংখ্যাগুলোই
# ক্রম অনুযায়ী নেওয়া হচ্ছে; ইউনিট (INCH/CM) টেক্সট থেকে আলাদাভাবে বের করা।
def _parse_measurement(text):
    t = _clean(text)
    nums = re.findall(r'(\d+\.?\d*)', t)
    l = _fmt_num(nums[0]) if len(nums) > 0 else ''
    w = _fmt_num(nums[1]) if len(nums) > 1 else ''
    h = _fmt_num(nums[2]) if len(nums) > 2 else ''
    unit = 'Inch' if 'inch' in t.lower() else 'Cm'
    return l, w, h, unit


# ─────────────────────── Item Name (Elastic / Master) ───────────────────────

_NEG_BEFORE = re.compile(r'\b(no|not|non|without|never|nai)\b|\bw/o\b')
_NEG_AFTER = re.compile(r'^\s*[:\-=]?\s*(is\s+|are\s+)?(not\b|no\b|nai\b|n/a\b)')


def _elastic_is_negated(text, match):
    """'elastic' শব্দটার আগে/পরে 'no / no need / not / without / non' আছে কিনা।
    আগের অংশ শুধু একই বাক্যাংশের (',' ';' '.' ':' নতুন লাইন পর্যন্ত) শেষ ৩০
    অক্ষর দেখা হয়, যাতে অন্য বাক্যের 'no' ভুল করে না ধরে।"""
    before = re.split(r'[,;.:\n]', text[:match.start()])[-1][-30:]
    if _NEG_BEFORE.search(before):
        return True
    after = text[match.end():match.end() + 25]
    return bool(_NEG_AFTER.match(after))


def _elastic_hint(cells, use_normal=True):
    """cells স্ক্যান করে রিটার্ন করে:
       'elastic' — non-negated 'elastic' পাওয়া গেছে
       'master'  — শুধু negated elastic ('NO ELASTIC') অথবা 'normal' পাওয়া গেছে
       None      — কোনো ইঙ্গিত নেই"""
    has_master_hint = False
    for v in cells:
        t = _clean(v).lower()
        if not t:
            continue
        for m in re.finditer('elastic', t):
            if _elastic_is_negated(t, m):
                has_master_hint = True
            else:
                return 'elastic'
        if use_normal and 'normal' in t:
            has_master_hint = True
    return 'master' if has_master_hint else None


def _flatten(rows):
    for row in rows:
        for v in row:
            yield v


def _classify_item_name(context_cells, sheet_name, all_rows=None):
    """ধাপ ১: context_cells (ফরম্যাট A-তে PO-মার্কার-রো, B-তে নোট-রো) স্ক্যান।
    ধাপ ২: না পেলে all_rows (পুরো শিট) থেকে শুধু 'elastic' শব্দ খোঁজা।
    ধাপ ৩: তাও না পেলে শিটের নাম ('E-com' -> Master Carton, নাহলে Elastic Hanger
    Carton)।"""
    hint = _elastic_hint(context_cells, use_normal=True)
    if hint is None and all_rows is not None:
        hint = _elastic_hint(_flatten(all_rows), use_normal=False)

    if hint == 'elastic':
        return 'Elastic Hanger Carton'
    if hint == 'master':
        return 'Master Carton'

    base_name = re.sub(r'\s*\(\d+\)\s*$', '', sheet_name or '').strip().lower()
    if base_name.startswith('e-com') or base_name.startswith('ecom'):
        return 'Master Carton'
    return 'Elastic Hanger Carton'


# ─────────────────────────── ব্লক/হেডার খোঁজা ───────────────────────────

def _row_has_gttl(row):
    return any(_norm(v) == 'gttl' for v in row)


def _find_gttl_row(rows, start, end):
    for i in range(start, end):
        if _row_has_gttl(rows[i]):
            return i
    return None


def _find_po_marker_rows(rows):
    """[ফরম্যাট A] কলাম B (0-indexed col1)-এ ঠিক 'PO' লেখা এমন সব রো খুঁজে বের
    করে — একটা শিটে একাধিক PO-ব্লকও থাকতে পারে।"""
    marker_rows = []
    for i, row in enumerate(rows):
        if len(row) > 1 and _norm(row[1]) == 'po':
            marker_rows.append(i)
    return marker_rows


def _find_header_rows(rows):
    """[ফরম্যাট B] যে রো-তে 'STYLE' আর 'CARTON QTY'/'CARTON BOOKING QTY' দুটোই
    আছে সেটাই হেডার-রো।"""
    found = []
    for i, row in enumerate(rows):
        norms = [_norm(v) for v in row]
        if 'style' in norms and any(
                n == 'cartonqty' or 'cartonbookingqty' in n for n in norms):
            found.append(i)
    return found


def _find_po_number(ctx_rows):
    """'PO NUMBER' লেবেল খুঁজে তার ডানের প্রথম অর্থপূর্ণ ভ্যালু (':' বাদ দিয়ে)
    রিটার্ন করে। হেডারের সবচেয়ে কাছেরটা (নিচ থেকে উপরে) নেওয়া হয়।"""
    for row in reversed(ctx_rows):
        for c, v in enumerate(row):
            if _norm(v) in ('ponumber', 'pono'):
                for v2 in row[c + 1:]:
                    t = _clean(v2).lstrip(':').strip()
                    if t and t not in ('-', '='):
                        return t
    return ''


def _detect_unit(ctx_rows):
    """'MEASUREMENT In Inch' ক্যাপশন থেকে ইউনিট। রিটার্ন: 'Inch' / 'Cm' / None।"""
    captions, others = [], []
    for v in _flatten(ctx_rows):
        t = _clean(v).lower()
        if not t:
            continue
        (captions if 'measurement' in t else others).append(t)
    for t in captions + others:
        if 'inch' in t:
            return 'Inch'
        if captions and t in captions and re.search(r'\bcms?\b|centi', t):
            return 'Cm'
    return None


def _resolve_columns(header_row):
    """হেডার-রো থেকে কলাম-ইনডেক্স বের করে (ডুপ্লিকেট লেবেল যেমন দুইটা 'H' ও
    সামলানো যায় — তাই লিস্ট ধরে এগোনো)।"""
    cells = [(c, _norm(v)) for c, v in enumerate(header_row) if _clean(v)]

    def first(pred):
        for c, n in cells:
            if pred(n):
                return c
        return None

    cols = {
        'style': first(lambda n: n == 'style'),
        'upc': first(lambda n: 'upcnumber' in n),
        'color': first(lambda n: n == 'color'),
        'meas': first(lambda n: 'ctnmeasurement' in n),
        'split_l': None,
        'qty': first(lambda n: 'cartonbookingqty' in n),
    }
    if cols['meas'] is None:
        cols['split_l'] = first(lambda n: n in ('l', 'length'))
    if cols['qty'] is None:
        cols['qty'] = first(lambda n: n == 'cartonqty')
    return cols


# ───────────────────────────── শিট রিডার ─────────────────────────────

def _build_blocks(rows, sheet_name, item_name_override):
    """শিটটাকে এক বা একাধিক ব্লকে ভাগ করে। প্রতিটা ব্লক:
    header_idx, end_row, po_no, item_name, unit (ফরম্যাট B-তে), stop_at_total,
    label।"""
    n_rows = len(rows)
    blocks = []

    # ── ফরম্যাট A: 'PO' মার্কার-রো ──
    marker_rows = _find_po_marker_rows(rows)
    if marker_rows:
        for bi, marker_row in enumerate(marker_rows):
            po_row = rows[marker_row]
            po_no = _clean(po_row[2]) if len(po_row) > 2 else ''
            item_name = item_name_override or _classify_item_name(
                po_row, sheet_name, all_rows=rows)
            end_row = marker_rows[bi + 1] if bi + 1 < len(marker_rows) else n_rows
            blocks.append({
                'header_idx': marker_row + 1,
                'end_row': end_row,
                'po_no': po_no,
                'item_name': item_name,
                'unit': None,
                'stop_at_total': False,   # আগের আচরণ অপরিবর্তিত
                'label': f'PO ব্লক {bi + 1}',
            })
        return blocks

    # ── ফরম্যাট B: হেডার-রো দেখে ──
    header_rows = _find_header_rows(rows)
    prev_gttl = None
    for bi, h in enumerate(header_rows):
        next_h = header_rows[bi + 1] if bi + 1 < len(header_rows) else n_rows
        gttl_idx = _find_gttl_row(rows, h + 1, next_h)

        ctx_start = min(prev_gttl + 7, h) if prev_gttl is not None else 0
        ctx_rows = rows[ctx_start:h]
        if gttl_idx is not None:
            # G.TTL-এর নিচের ফুটার-নোট (যেমন 'NO ELASTIC, ONLY NEED CARTON STICKER ...')
            ctx_rows = ctx_rows + rows[gttl_idx + 1:min(gttl_idx + 7, next_h)]
        prev_gttl = gttl_idx if gttl_idx is not None else h

        po_no = _find_po_number(rows[ctx_start:h])
        if not po_no and re.fullmatch(r'\d+', (sheet_name or '').strip()):
            po_no = sheet_name.strip()   # শিটের নামই PO নম্বর

        item_name = item_name_override or _classify_item_name(
            _flatten(ctx_rows), sheet_name, all_rows=rows)

        blocks.append({
            'header_idx': h,
            'end_row': next_h,
            'po_no': po_no,
            'item_name': item_name,
            'unit': _detect_unit(ctx_rows),
            'stop_at_total': True,
            'label': f'ব্লক {bi + 1}',
        })
    return blocks


def read_dhaka_booking_sheet(rows, sheet_name, item_name_override='', manual_ply=''):
    """একটা শিট থেকে লাইন-আইটেম বের করে। রিটার্ন করে (line_items, warnings)।"""
    blocks = _build_blocks(rows, sheet_name, item_name_override)
    if not blocks:
        return [], []

    items = []
    warnings = []
    n_rows = len(rows)
    ply = manual_ply.strip() if manual_ply else '5'

    for blk in blocks:
        header_row_idx = blk['header_idx']
        if header_row_idx >= n_rows:
            continue
        po_no = blk['po_no']
        item_name = blk['item_name']
        end_row = blk['end_row']

        cols = _resolve_columns(rows[header_row_idx])
        style_col, upc_col = cols['style'], cols['upc']
        color_col, qty_col = cols['color'], cols['qty']
        meas_col, split_l = cols['meas'], cols['split_l']

        if style_col is None or qty_col is None or (meas_col is None and split_l is None):
            warnings.append(
                f"⚠️ শিট '{sheet_name}' ({blk['label']}): Style/Measurement (CTN MEASUREMENT "
                f"বা L-H-H)/Carton Qty কলাম পাওয়া যায়নি — এই ব্লক স্কিপ করা হয়েছে।"
            )
            continue

        unit_split = blk['unit']
        if split_l is not None and unit_split is None:
            unit_split = 'Cm'
            warnings.append(
                f"⚠️ শিট '{sheet_name}' (PO {po_no}): Measurement-এর ইউনিট (Inch/Cm) "
                f"ক্যাপশন পাওয়া যায়নি — 'Cm' ধরা হয়েছে, চেক করে নিন।"
            )

        total_row_qty = None
        extracted_total = 0.0
        row_count = 0
        had_fraction = False

        r = header_row_idx + 1
        while r < end_row:
            row = rows[r]

            if _row_has_gttl(row):
                qty_val = row[qty_col] if qty_col < len(row) else None
                if _is_num(qty_val):
                    total_row_qty = _num(qty_val)
                if blk['stop_at_total']:
                    break   # G.TTL-এর পরের নোট/সামারি রো ডাটা না
                r += 1
                continue

            style_val = _clean(row[style_col]) if style_col < len(row) else ''
            if not style_val:
                r += 1
                continue

            qty_val = row[qty_col] if qty_col < len(row) else None
            if not _is_num(qty_val) or (_num(qty_val) or 0) <= 0:
                r += 1
                continue  # qty 0/ফাঁকা — বাদ

            if meas_col is not None:
                meas_val = _clean(row[meas_col]) if meas_col < len(row) else ''
                length, width, height, unit = _parse_measurement(meas_val)
            else:
                def _cell(off):
                    c = split_l + off
                    return _fmt_num(_num(row[c])) if c < len(row) else ''
                length, width, height, unit = _cell(0), _cell(1), _cell(2), unit_split

            upc_val = _clean(row[upc_col]) if upc_col is not None and upc_col < len(row) else ''
            color_val = _clean(row[color_col]) if color_col is not None and color_col < len(row) else ''

            # ইউজার-কনফার্মড: Qty ফ্র্যাকশনে থাকলে (ফরম্যাট A-তে প্রায়ই থাকে, যেমন
            # 6.15, 18.34) রাউন্ড করে বসবে।
            raw_qty = _num(qty_val)
            if raw_qty != int(raw_qty):
                had_fraction = True
            qty_num = round(raw_qty)
            extracted_total += qty_num
            row_count += 1

            items.append({
                'item_name': item_name,
                'ewo_no': 'N/A',
                'style_no': style_val,
                'po_no': po_no or 'N/A',
                'length': length,
                'width': width,
                'height': height,
                'ply': ply,
                'qty': qty_num,
                'pack_type': upc_val or 'N/A',
                'reference': color_val or 'N/A',
                'remarks': '',
                'color': 'N/A',
                'size': 'N/A',
                'delivery_date': '',
                'measurement_unit': unit,
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': sheet_name,
            })
            r += 1

        # ফ্র্যাকশনাল Qty থাকলে প্রতিটা রো আলাদাভাবে রাউন্ড হওয়ায় G.TTL-এর সাথে
        # সামান্য পার্থক্য (প্রতি রো-তে সর্বোচ্চ ০.৫ পর্যন্ত) স্বাভাবিক — তাই
        # tolerance রো-সংখ্যা অনুযায়ী বাড়ে। সব Qty পূর্ণসংখ্যা হলে (ফরম্যাট B)
        # টোটাল হুবহু মিলতে হবে।
        tol = max(0.51, 0.51 * row_count) if had_fraction else 0.51
        if total_row_qty is not None and abs(total_row_qty - extracted_total) > tol:
            warnings.append(
                f"⚠️ শিট '{sheet_name}' (PO {po_no}): G.TTL রো-তে Carton Qty মোট "
                f"{total_row_qty:g}, কিন্তু বের করা লাইন-আইটেমগুলোর মোট Qty "
                f"{extracted_total:g} — পার্থক্য {total_row_qty - extracted_total:g}, "
                f"ভালোভাবে চেক করে নিন।"
            )

    return items, warnings


def read_dhaka_booking_file(file_stream, filename, item_name_override='', manual_ply=''):
    """একটা .xls/.xlsx ফাইলের সব শিট থেকে ডাটা বের করে। রিটার্ন করে
    (line_items, warnings)।"""
    file_stream.seek(0)
    sheets = pd.read_excel(file_stream, sheet_name=None, header=None)

    all_items = []
    all_warnings = []
    for sheet_name, df in sheets.items():
        items, warns = read_dhaka_booking_sheet(
            df.values.tolist(), str(sheet_name),
            item_name_override=item_name_override, manual_ply=manual_ply)
        all_items.extend(items)
        all_warnings.extend(warns)

    if not all_items:
        all_warnings.append(f"⚠️ '{filename}': কোনো ভ্যালিড (qty>0) লাইন-আইটেম পাওয়া যায়নি।")

    return all_items, all_warnings


def combine_dhaka_booking_files(files, item_name_override='', manual_ply=''):
    """files: [(BytesIO, filename), ...] — BATCH_REGISTRY-এর uniform কল-
    সিগনেচার। item_name_override সাধারণত ব্যবহার হয় না (Item Name সম্পূর্ণ
    ফাইলের কনটেন্ট/শিট-নাম থেকেই ডিটেক্ট হয়), কিন্তু ভবিষ্যতে UI থেকে জোর
    করে override দিতে চাইলে সাপোর্ট করা আছে। manual_ply UI থেকে এলে
    ব্যবহার হয়, না দিলে ডিফল্ট 5। রিটার্ন করে (line_items, warnings)।"""
    combined = []
    all_warnings = []
    for file_stream, filename in files:
        items, warns = read_dhaka_booking_file(
            file_stream, filename,
            item_name_override=item_name_override, manual_ply=manual_ply)
        for it in items:
            it['_source_file'] = filename
        combined.extend(items)
        all_warnings.extend(warns)
    return combined, all_warnings