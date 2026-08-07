import re
import pdfplumber
import pandas as pd

# Header extraction (PO No/Customer/Buyer) is byte-for-byte identical to the
# Carton module's cover-page format, so it's reused as-is — same as Thermal.
from extractor import clean, extract_header_info

# Thermal-এর মতোই — page0-এর রেট-সামারি টেবিলের প্রথম কলামের নাম বায়ার/আইটেম
# ভেদে ভিন্ন হতে পারে।
SUMMARY_HEADER_MARKERS = {'ST Caow', 'PT Caow', 'Sticker Type', 'Item Caow'}

_ITEM_NAME_RE = re.compile(r'Item\s*Name\s*:\s*(.+)', re.I)


def extract_item_name(pdf):
    """PDF-এর কভার পেজ থেকে আসল Item Name বের করে আনে (যেমন 'P.S Tag',
    'Poly Sticker', বা ভবিষ্যতে অন্য যেকোনো ট্যাগ) — Thermal মডিউলে যেমন
    'Thermal Sticker' ফিক্সড স্ট্রিং হিসেবে বসানো হতো, এখানে সেটা করা যাবে না,
    কারণ এই একই এক্সট্রাক্টরের আন্ডারে ভবিষ্যতে ভিন্ন ভিন্ন আইটেম নামে অনেক
    ধরনের PO আসবে (Tag, PS Tag, Poly Sticker...) — প্রতিটার আসল নাম PDF থেকেই
    ডাইনামিকভাবে বের করে সেভাবেই আউটপুটে বসানো হয়, যাতে এলোমেলো না হয়।"""
    if not pdf.pages:
        return ''
    text = pdf.pages[0].extract_text() or ''
    for line in text.split('\n'):
        m = _ITEM_NAME_RE.search(line)
        if m:
            return clean(m.group(1))
    return ''


def extract_summary_table_pp(pdf):
    """page0-এর রেট-সামারি টেবিল — Thermal-এর extract_summary_table_thermal-এর
    সাথে যুক্তি হুবহু এক।"""
    for page in pdf.pages:
        for t in page.extract_tables():
            if t and t[0] and clean(t[0][0]) in SUMMARY_HEADER_MARKERS:
                header = [clean(h) for h in t[0]]
                rows = [[clean(c) for c in r] for r in t[1:]
                        if r and r[0] and clean(r[0]).lower() not in
                        ('pcs wise total', 'pcs wise total qty', 'total', 'total value')]
                return pd.DataFrame(rows, columns=header)
    return pd.DataFrame()


def _clean_numeric_cell(v):
    """সরু সাইজ-কলামে বড় সংখ্যা মাঝপথে লাইন-ব্রেক হয়ে ভেঙে গেলে
    ('132.00' -> '132.0\\n0') জোড়া লাগায় — Thermal মডিউলের একই ফিক্স,
    দেখুন thermal_extractor.py-এর docstring।"""
    return re.sub(r'\s+', '', clean(v))


def _to_float(v):
    v = _clean_numeric_cell(v).replace(',', '')
    if not v:
        return None
    try:
        return float(v)
    except ValueError:
        return None


def _norm(s):
    return re.sub(r'\s+', '', str(s or '')).lower()


# EWO No/Style No/PONo/Item Caow/Item Reference-এর মতো কোড-ফিল্ডে মাঝপথে
# ভাঙা সংখ্যা-সাফিক্স জোড়া লাগানোর জন্য — দেখুন thermal_extractor.py-এর
# _CODE_SPLIT_DIGIT_RE-এর ডকস্ট্রিং, যুক্তি হুবহু এক।
_CODE_SPLIT_DIGIT_RE = re.compile(r'(?<=[A-Za-z0-9])\s+(?=\d+(?:\s|$))')

_CODE_FIELDS = {'EWO No', 'Style No', 'PONo', 'Item Caow', 'Item Reference'}


def _clean_code_field(v):
    return _CODE_SPLIT_DIGIT_RE.sub('', clean(v))


_SIZE_PREFIX_RE = re.compile(r'^[A-Za-z]{1,5}\s*[:=]\s*', re.I)


def _clean_size_label(raw):
    """Thermal-এর _clean_size_label-এর সাথে হুবহু এক যুক্তি — prefix বাদ দেয়
    এবং লাইন-ব্রেক-জনিত ভেতরের স্পেস মুছে দেয়।"""
    s = clean(raw)
    s = _SIZE_PREFIX_RE.sub('', s).strip()
    s = re.sub(r'\s+', '', s)
    return s if s else 'N/A'


PP_FIELD_MAP = {
    _norm('EWO No'): 'EWO No',
    _norm('Style No'): 'Style No',
    _norm('Item Caow'): 'Item Caow',
    _norm('Item Reference'): 'Item Reference',
    _norm('PONo'): 'PONo',
    _norm('PO No'): 'PONo',
    _norm('PO QTY'): 'PO QTY',
    _norm('POQty'): 'PO QTY',
    _norm('Gmt Color'): 'Gmt. Color',
    _norm('Gmt. Color'): 'Gmt. Color',
    _norm('Color'): 'Color',
    _norm('Instruction'): 'Instruction',
    _norm('Country'): 'Country',
    _norm('Delivery Place'): 'Delivery Place',
    _norm('Delivery Address'): 'Delivery Address',
    _norm('Delivery Start Date'): 'Delivery Start Date',
    _norm('Delivery End Date'): 'Delivery End Date',
    _norm('UOM'): 'UOM',
}


def _canonical_field_names(raw_names):
    return [PP_FIELD_MAP.get(_norm(n), n) for n in raw_names]


# raw_wide_df-এর যে কলামগুলো মেটা-ফিল্ড — বাকি সব কলামকে "সাইজ/qty কলাম" ধরে
# নেওয়া হয় qty-mismatch ডায়াগনস্টিকের জন্য (Thermal-এর সাথে একই যুক্তি)।
_META_COLUMN_NAMES = set(PP_FIELD_MAP.values()) | {'Reference'}


def compute_qty_mismatch_warnings(line_items, raw_wide_df, summary_df=None):
    """Thermal মডিউলের compute_qty_mismatch_warnings-এর সাথে হুবহু এক যুক্তি —
    গ্র্যান্ড টোটাল + সেল-লেভেল ডায়াগনস্টিক, দেখুন thermal_extractor.py।"""
    warnings = []

    total_extracted = sum((_to_float(li.get('qty', '')) or 0) for li in line_items)
    if summary_df is not None and not summary_df.empty:
        qty_col = next((c for c in summary_df.columns
                         if _norm(c) in (_norm('POQty'), _norm('PO QTY'))), None)
        if qty_col:
            total_pdf = sum((_to_float(v) or 0) for v in summary_df[qty_col])
            diff = round(total_pdf - total_extracted, 2)
            if abs(diff) > 0.01:
                warnings.append(
                    f"⚠️ মোট Qty মিলছে না — PDF Summary-তে মোট {total_pdf:.2f}, "
                    f"কিন্তু Excel-এ বসেছে {total_extracted:.2f} (পার্থক্য {diff:+.2f})। "
                    f"নিচের row-ভিত্তিক বিস্তারিত (যদি থাকে) দেখুন, নাহলে PO Details শীট মিলিয়ে দেখুন।"
                )

    if raw_wide_df is not None and not raw_wide_df.empty:
        size_cols = [c for c in raw_wide_df.columns if c not in _META_COLUMN_NAMES]
        for _, row in raw_wide_df.iterrows():
            for col in size_cols:
                raw_val = row.get(col, '')
                if raw_val is None or str(raw_val).strip() == '':
                    continue
                if _to_float(raw_val) is not None:
                    continue
                ident = (f"EWO {row.get('EWO No', 'N/A')} / Style {row.get('Style No', 'N/A')} / "
                         f"Color {row.get('Gmt. Color') or row.get('Color') or 'N/A'}")
                warnings.append(
                    f"⚠️ {ident} — সাইজ '{col}'-এর ভ্যালু '{raw_val}' সংখ্যা হিসেবে পড়া যায়নি, "
                    f"এই qty-টা Excel-এ যোগ হয়নি। PO Details শীটে গিয়ে সরাসরি চেক করুন।"
                )

    return warnings


_SUMMARY_ROW_MARKERS = ('pcs wise total', 'pcs wise total qty', 'total', 'total value')


def _looks_like_size_continuation_header(row):
    """Thermal-এর মতোই — সাইজ-ওভারফ্লো continuation টেবিল চেনার সিগন্যাল:
    শেষ কলামের নাম হুবহু 'Total'।"""
    if not row or len(row) < 2:
        return False
    return clean(row[-1]).lower() == 'total'


def extract_detail_rows_pp(pdf):
    """Printing Press PO-এর 'Purchase Order Details' টেবিল বের করে আনে।
    গঠন Thermal মডিউলের সাথে হুবহু এক (WIDE/FLAT, multi-page continuation,
    সাইজ-ওভারফ্লো টেবিল, লাইন-ব্রেকে ভাঙা সংখ্যা) — শুধু ফিল্ড-নামগুলো ভিন্ন
    (Sticker Caow/Reference -> Item Caow/Reference)। বিস্তারিত যুক্তির জন্য
    thermal_extractor.py-এর extract_detail_rows_thermal-এর docstring দেখুন।

    Returns (line_items_df, raw_wide_df).
    """
    field_names = None
    size_labels = None
    split_idx = None
    is_wide = False
    raw_wide_rows = []
    melted = []
    primary_rows = []
    last_ewo, last_style = '', ''

    def build_meta(row):
        nonlocal last_ewo, last_style
        meta_vals = []
        for i, name in enumerate(field_names):
            raw_v = row[i] if i < len(row) else ''
            meta_vals.append(_clean_code_field(raw_v) if name in _CODE_FIELDS else clean(raw_v))
        meta = dict(zip(field_names, meta_vals))

        if meta.get('EWO No'):
            last_ewo = meta['EWO No']
        else:
            meta['EWO No'] = last_ewo
        if meta.get('Style No'):
            last_style = meta['Style No']
        else:
            meta['Style No'] = last_style

        # ব্যবহারকারীর নির্দেশ অনুযায়ী — 'Instruction' কলামই এখানে প্রধান
        # Reference/SKU সোর্স। 'Item Reference'/'Pre Pack' fallback হিসেবে,
        # কোনো বায়ারের PDF-এ Instruction না থাকলে।
        meta['Reference'] = meta.get('Instruction') or meta.get('Item Reference') or meta.get('Pre Pack') or ''
        return meta

    def process_data_row(row):
        if not row:
            return
        first_cell = clean(row[0]) if row[0] else ''
        if first_cell.lower() in _SUMMARY_ROW_MARKERS:
            return

        meta = build_meta(row)

        if is_wide and size_labels:
            qty_cells = row[split_idx:split_idx + len(size_labels)]
            sizes = {
                (size_labels[i] or f'Size{i+1}'): (_clean_numeric_cell(qty_cells[i]) if i < len(qty_cells) else '')
                for i in range(len(size_labels))
            }
            primary_rows.append({'meta': meta, 'sizes': sizes})
        else:
            if 'PO QTY' in meta:
                meta['PO QTY'] = _clean_numeric_cell(meta['PO QTY'])
            qty = _to_float(meta.get('PO QTY', ''))
            raw_wide_rows.append(dict(meta))
            melted.append({**meta, 'Size': 'N/A', 'Qty': qty})

    for page in pdf.pages:
        for t in page.extract_tables():
            if not t or not t[0]:
                continue
            if clean(t[0][0]) in SUMMARY_HEADER_MARKERS:
                continue
            rows = t
            if rows[0] and clean(rows[0][0]) == 'Purchase Order Details':
                rows = rows[1:]
            if not rows:
                continue

            if field_names is None:
                row0 = rows[0]
                c0 = clean(row0[0]) if row0 and row0[0] else ''
                if c0 == 'Size/Measurement':
                    size_row = row0
                    field_row = rows[1] if len(rows) > 1 else []
                    for i, v in enumerate(size_row):
                        if i == 0:
                            continue
                        if clean(v):
                            split_idx = i
                            break
                    if split_idx is None:
                        continue
                    raw_sizes = [clean(v) for v in size_row[split_idx:]]
                    raw_sizes = [s for s in raw_sizes if s and s.lower() != 'total']
                    size_labels = [_clean_size_label(s) for s in raw_sizes]
                    is_wide = True
                    field_names = _canonical_field_names([clean(v) for v in field_row[:split_idx]])
                    for r in rows[2:]:
                        process_data_row(r)
                else:
                    is_wide = False
                    field_names = _canonical_field_names([clean(v) for v in row0])
                    for r in rows[1:]:
                        process_data_row(r)
            elif is_wide and _looks_like_size_continuation_header(rows[0]):
                extra_raw = [clean(v) for v in rows[0][:-1]]
                extra_labels = [_clean_size_label(s) for s in extra_raw]
                data_rows = rows[1:]
                n_primary = len(primary_rows)
                for ridx in range(min(n_primary, len(data_rows))):
                    r = data_rows[ridx]
                    for ci, label in enumerate(extra_labels):
                        val = _clean_numeric_cell(r[ci]) if ci < len(r) else ''
                        primary_rows[ridx]['sizes'][label] = val
                size_labels.extend(extra_labels)
            else:
                for r in rows:
                    process_data_row(r)

    if is_wide:
        for pr in primary_rows:
            meta, sizes = pr['meta'], pr['sizes']
            raw_wide_rows.append({**meta, **sizes})
            for size_label, val in sizes.items():
                qv = _to_float(val)
                if qv is None:
                    continue
                melted.append({**meta, 'Size': _clean_size_label(size_label), 'Qty': qv})

    if field_names is None:
        raise ValueError("এই PDF-এ পরিচিত Printing Press 'Purchase Order Details' টেবিল ফরম্যাট পাওয়া যায়নি।")

    line_items_df = pd.DataFrame(melted)
    raw_wide_df = pd.DataFrame(raw_wide_rows)
    return line_items_df, raw_wide_df


def _rate_lookup(summary_df):
    """Thermal-এর _rate_lookup-এর সাথে হুবহু এক যুক্তি।"""
    if summary_df is None or summary_df.empty:
        return {}, None, []

    rate_col = next((c for c in summary_df.columns if _norm(c) == _norm('Rate')), None)
    if rate_col is None:
        return {}, None, []

    all_rates = [clean(v) for v in summary_df[rate_col].tolist()]
    non_blank_rates = [r for r in all_rates if r]
    if len(set(non_blank_rates)) <= 1:
        return {}, (non_blank_rates[0] if non_blank_rates else ''), []

    key_cols = [c for c in summary_df.columns if c != rate_col and _norm(c) != _norm('Total Value')]
    canon_key_cols = _canonical_field_names(key_cols)
    lookup = {}
    for _, r in summary_df.iterrows():
        key = tuple(clean(r.get(c, '')) for c in key_cols)
        lookup[key] = clean(r.get(rate_col, ''))
    return lookup, None, canon_key_cols


def to_canonical_pp(df, summary_df=None):
    """মেল্ট/ফ্ল্যাট করা DataFrame-কে canonical line-item স্কিমায় রূপান্তর করে,
    যেটা printing_press_builder.py ব্যবহার করবে।"""
    rate_lookup, single_rate, key_fields = _rate_lookup(summary_df)
    line_items = []
    for _, r in df.iterrows():
        if single_rate is not None:
            rate = single_rate
        else:
            key = tuple(clean(r.get(f, '')) for f in key_fields) if key_fields else ()
            rate = rate_lookup.get(key, '')
        line_items.append({
            'ewo_no': r.get('EWO No', ''),
            'style_no': r.get('Style No', ''),
            'po_no': r.get('PONo', ''),
            'item_caow': r.get('Item Caow', ''),
            'item_reference': r.get('Item Reference', ''),
            'color': r.get('Gmt. Color', '') or r.get('Color', ''),
            'reference': r.get('Reference', ''),
            'size': _clean_size_label(r.get('Size', '')),
            'qty': r.get('Qty', ''),
            'uom': r.get('UOM', 'Pcs') or 'Pcs',
            'rate': rate,
            'delivery_date_pdf': r.get('Delivery Start Date', ''),
            'delivery_place_pdf': r.get('Delivery Place', ''),
            'delivery_address_pdf': r.get('Delivery Address', ''),
        })
    return line_items


def get_unique_delivery_info_pp(raw_wide_df):
    """Thermal-এর get_unique_delivery_info_thermal-এর সাথে হুবহু এক।"""
    def uniques(col):
        if raw_wide_df is None or raw_wide_df.empty or col not in raw_wide_df.columns:
            return []
        seen = []
        for v in raw_wide_df[col]:
            v = clean(v)
            if v and v not in seen:
                seen.append(v)
        return seen

    return {
        'delivery_places': uniques('Delivery Place'),
        'delivery_addresses': uniques('Delivery Address'),
    }


def process_pdf_pp(file_stream):
    """Printing Press মডিউলের মেইন এন্ট্রি পয়েন্ট।
    Returns (header_info, canonical line_items, raw_wide_df, summary_df)।
    header_info-তে Carton-এর মতো po_number/customer/buyer-এর পাশাপাশি
    'item_name'-ও থাকে (PDF থেকে ডাইনামিকভাবে পড়া, ফিক্সড স্ট্রিং না)।"""
    with pdfplumber.open(file_stream) as pdf:
        header_info = extract_header_info(pdf)
        header_info['item_name'] = extract_item_name(pdf) or 'N/A'
        summary_df = extract_summary_table_pp(pdf)
        melted_df, raw_wide_df = extract_detail_rows_pp(pdf)
    line_items = to_canonical_pp(melted_df, summary_df)
    return header_info, line_items, raw_wide_df, summary_df