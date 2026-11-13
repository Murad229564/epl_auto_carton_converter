"""
Label মডিউল (Additional Label) — Epyllion 'Purchase Order (PO)' PDF extractor।

PDF ফরম্যাট (Carhartt-এর Label PO, Item Name: Label):
  * পাতা ১(-২): 'Label Type / Caow / USIM / Ref. No / ... / POQty / UOM / Rate /
    Total Value' সামারি টেবিল — এখান থেকে Rate (এবং মোট Qty-র যাচাই)।
  * তারপর 'Purchase Order Details' টেবিল। সাইজ (GS =XS ... GS =XXL) এত বেশি যে এক
    পাতায় ধরে না — PDF সাইজ কলামগুলো দুই ভাগ করে ফেলে: একটা পাতায় (XS..XL) আরেক
    পাতায় (XXL + 'Total')। দ্বিতীয় পাতাতেও EWO/Style/PO/Item Name/Color Code/USIM
    সব মেটা-কলাম রিপিট হয় — তাই (PO No + Item Name + Item No + Color Code + USIM)
    দিয়ে দুই ভাগের রো জোড়া লাগানো হয় (পাতা যত বেশিই হোক)।
  * টেবিল-এক্সট্র্যাকশনে সারিগুলো কখনো ২২ কলাম (EWO/Style সহ), কখনো ২০ কলাম
    (EWO/Style মার্জড সেল বলে বাদ) আসে — তাই কলাম ডান দিক থেকে ধরা হয় (সাইজ
    কলাম, তারপর UOM, ... PO No), আর EWO/Style না থাকলে উপরের রো থেকে carry-forward।
  * 'PO ... Wise Total' / 'PCS Wise Total' সামারি রো বাদ।

ম্যাপিং (ইউজার-কনফার্মড):
  EWO No      -> Gmt. EWO No
  Style No    -> Gmt. Style No
  PO No       -> Gmt. PO
  Reference/SKU = Item Name / Color Code / USIM-Ref.No  (যেমন 'Color Label/3ITGD/Cozy Blue')
  সাইজ-হেডার ('GS =XS') থেকে '=' এর পরের অংশ -> Gmt. Size
  Rate        -> সামারি টেবিল থেকে
  Qty         -> প্রতিটা সাইজ-সেল (ফাঁকা বাদ)
"""
import re
import pdfplumber
import pandas as pd

try:
    from config import CUSTOMERS, CUSTOMER_ALIASES
except Exception:  # extractor আলাদাভাবে টেস্ট করার সময়
    CUSTOMERS = {}
    CUSTOMER_ALIASES = {}


def clean(v):
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _to_float(v):
    s = re.sub(r'\s+', '', clean(v)).replace(',', '')
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _join_code(v):
    """'26TR007\\n274' -> '26TR007274'; কমা-সেপারেটেড লিস্টে ', ' অক্ষত থাকে।"""
    return re.sub(r'(?<=[A-Za-z0-9])\s+(?=[A-Za-z0-9])', '', clean(v))


def _join_lines(v):
    """সেল-ওয়্যাপে মাঝপথে ভাঙা শব্দ জোড়া লাগায় ('HOODE\\nD UNI' -> 'HOODED UNI')।
    আগের লাইন '-' দিয়ে শেষ হলে ('T-' + 'SHIRT'), অথবা পরের লাইন একটামাত্র বড় হাতের
    অক্ষর দিয়ে শুরু ('D ...') এবং আগের লাইনের শেষ শব্দ ৪+ অক্ষরের হলে স্পেস ছাড়া জোড়া
    লাগে; বাকি সব ক্ষেত্রে স্পেস দিয়ে।"""
    if v is None:
        return ''
    lines = [l.strip() for l in str(v).split('\n') if l.strip()]
    if not lines:
        return ''
    out = lines[0]
    for nxt in lines[1:]:
        last_tok = out.split(' ')[-1]
        if out.endswith('-'):
            out += nxt  # 'T-' + 'SHIRT' -> 'T-SHIRT' (হাইফেনে ভাঙা শব্দ)
        elif re.match(r'^[A-Z](\s|$)', nxt) and len(last_tok) >= 4 and last_tok.isalpha():
            out += nxt
        else:
            out += ' ' + nxt
    return re.sub(r'\s+', ' ', out).strip()


_FIELD_KEYS = {
    'ewono': 'EWO No', 'styleno': 'Style No', 'pono': 'PO No', 'itemname': 'Item Name',
    'itemno': 'Item No', 'programname': 'Program Name', 'colorcode': 'Color Code',
    'labelcolor': 'Label Color', 'usimrefno': 'USIM / Ref. No', 'composition': 'Composition',
    'careinstruction': 'Care Instruction', 'instruction': 'Instruction',
    'deliveryplace': 'Delivery Place', 'deliveryaddress': 'Delivery Address',
    'deliverystartdate': 'Delivery Start Date', 'deliveryenddate': 'Delivery End Date',
    'uom': 'UOM',
}


def _canon_field(raw):
    n = _norm(raw)
    if n.startswith('composition'):
        return 'Composition'
    return _FIELD_KEYS.get(n, clean(raw))


def _size_label(raw):
    """'GS =XS' / 'GS\\n=XXL' -> 'XS' / 'XXL'; '=' না থাকলে স্পেস মুছে যা আছে।"""
    s = clean(raw)
    if '=' in s:
        s = s.split('=', 1)[1]
    s = re.sub(r'\s+', '', s)
    return s or 'N/A'



def _cust_key(s):
    """Customer নাম মেলানোর কী: case/স্পেস/দাঁড়ি-কমা উপেক্ষা, আর 'Limited' = 'Ltd'।"""
    t = re.sub(r'\blimited\b', 'ltd', str(s or '').lower())
    return re.sub(r'[^a-z0-9]', '', t)


def _canonical_customer(pdf_name):
    """PDF-এর শিরোনামের নামকে config.py-এর হুবহু (case-sensitive) Customer নামে রূপান্তর করে —
    যাতে UI-তে Customer অটো-লক হয় এবং ভ্যালিডেশন মেলে। প্রথমে CUSTOMERS লিস্ট (Limited/Ltd
    ও কেস উপেক্ষা করে), তারপর CUSTOMER_ALIASES। কিছুই না মিললে PDF-এর নাম যেমন আছে তেমন।"""
    if not pdf_name:
        return ''
    key = _cust_key(pdf_name)
    for lst in CUSTOMERS.values():
        for c in lst:
            if _cust_key(c) == key:
                return c
    for alias, canonical in CUSTOMER_ALIASES.items():
        if _cust_key(alias) == key:
            return canonical
    return pdf_name.strip()

# --------------------------------------------------------------------------
# হেডার (PO No / Buyer / Customer)
# --------------------------------------------------------------------------
def extract_header_info(pdf):
    text = pdf.pages[0].extract_text() or ''
    po = buyer = customer = ''

    m = re.search(r'PO\s*No\.?\s*:\s*(\S+)', text)
    if m:
        po = m.group(1).strip()
    m = re.search(r'Garments\s*Buyer\s*:\s*(.+?)(?:\s+Merchandising|\n|$)', text)
    if m:
        buyer = m.group(1).strip()

    first_line = next((l.strip() for l in text.split('\n') if l.strip()), '')
    customer = _canonical_customer(first_line)

    return {'po_number': po, 'customer': customer, 'buyer': buyer}


# --------------------------------------------------------------------------
# সামারি টেবিল (পাতা ১-২)
# --------------------------------------------------------------------------
def extract_summary_table(pdf):
    header = None
    rows = []
    for page in pdf.pages:
        for t in page.extract_tables():
            if not t or not t[0]:
                continue
            first = clean(t[0][0])
            if first == 'Label Type':
                header = [clean(h) for h in t[0]]
                rows.extend([[clean(c) for c in r] for r in t[1:]])
            elif header and len(t[0]) == len(header) and first != 'Purchase Order Details' \
                    and not first.startswith('Size/Measurement'):
                rows.extend([[clean(c) for c in r] for r in t])
        if header and any(clean(t[0][0]) in ('Purchase Order Details',) or clean(t[0][0]).startswith('Size/')
                          for t in page.extract_tables() if t and t[0]):
            break
    if not header:
        return pd.DataFrame()
    keep = [r for r in rows if r and r[0] and clean(r[0]).lower() not in
            ('pcs wise total qty', 'pcs wise total', 'total value', 'total')]
    return pd.DataFrame(keep, columns=header)


def _summary_lookup(summary_df):
    """(single_rate_or_None, {(label_type, usim): rate})"""
    if summary_df is None or summary_df.empty:
        return None, {}
    rate_col = next((c for c in summary_df.columns if _norm(c) == 'rate'), None)
    if not rate_col:
        return None, {}
    rates = [clean(v) for v in summary_df[rate_col] if clean(v)]
    if len(set(rates)) <= 1:
        return (rates[0] if rates else ''), {}
    type_col = summary_df.columns[0]
    usim_col = next((c for c in summary_df.columns if _norm(c).startswith('usim')), None)
    lk = {}
    for _, r in summary_df.iterrows():
        lk[(_norm(r.get(type_col)), _norm(r.get(usim_col)) if usim_col else '')] = clean(r.get(rate_col))
    return None, lk


# --------------------------------------------------------------------------
# ডিটেইল টেবিল (সব পাতা, সাইজ-ওভারফ্লো মার্জ-সহ)
# --------------------------------------------------------------------------
def extract_detail_rows(pdf):
    """Returns (rows, raw_df): rows = [{'meta': {...}, 'sizes': {label: qty}}, ...]"""
    field_names = None
    cur_labels = None
    seen_header = False
    last = {'EWO No': '', 'Style No': ''}
    entries = []          # সব মার্জড এন্ট্রি (প্রথম-দেখা ক্রমে)
    by_key = {}           # key -> [entry, ...]
    unmatched_overflow = 0

    def handle_row(cells):
        nonlocal unmatched_overflow
        n = len(cur_labels)
        cells = [('' if c is None else c) for c in cells]
        if len(cells) <= n:
            return
        meta_cells, size_cells = cells[:-n], cells[-n:]
        L = len(meta_cells)
        names = field_names
        if L >= len(names):
            meta_cells = meta_cells[-len(names):]
            aligned = names
        else:
            aligned = names[len(names) - L:]
        meta = dict(zip(aligned, meta_cells))

        # EWO/Style carry-forward (wise-total রো-তেও আপডেট হয়)
        ewo = _join_code(meta.get('EWO No', ''))
        style = _join_lines(meta.get('Style No', ''))
        if ewo:
            last['EWO No'] = ewo
        if style:
            last['Style No'] = style

        joined = ' '.join(clean(c) for c in cells).lower()
        if 'wise total' in joined:
            return

        item_name = clean(meta.get('Item Name', ''))
        if not item_name:
            return  # ডাটা-রো না

        m = {
            'EWO No': last['EWO No'],
            'Style No': last['Style No'],
            'PO No': _join_code(meta.get('PO No', '')),
            'Item Name': item_name,
            'Item No': _join_code(meta.get('Item No', '')),
            'Color Code': _join_code(meta.get('Color Code', '')),
            'USIM / Ref. No': clean(meta.get('USIM / Ref. No', '')),
            'Label Color': clean(meta.get('Label Color', '')),
            'Composition': clean(meta.get('Composition', '')),
            'Care Instruction': clean(meta.get('Care Instruction', '')),
            'Delivery Place': clean(meta.get('Delivery Place', '')),
            'Delivery Address': clean(meta.get('Delivery Address', '')),
            'UOM': clean(meta.get('UOM', '')) or 'Pcs',
        }
        sizes = {}
        for lab, v in zip(cur_labels, size_cells):
            if _norm(lab) == 'total':
                continue
            sizes[lab] = _to_float(v)

        key = tuple(_norm(m[k]) for k in ('PO No', 'Item Name', 'Item No', 'Color Code', 'USIM / Ref. No'))
        target = None
        for e in by_key.get(key, []):
            if not (set(e['sizes']) & set(sizes)):
                target = e
                break
        if target is None:
            target = {'meta': m, 'sizes': {}}
            entries.append(target)
            by_key.setdefault(key, []).append(target)
        target['sizes'].update(sizes)

    for page in pdf.pages:
        for t in page.extract_tables():
            if not t or not t[0]:
                continue
            rows = list(t)
            if clean(rows[0][0]) == 'Purchase Order Details':
                rows = rows[1:]
            if not rows:
                continue

            if clean(rows[0][0]) == 'Size/Measurement':
                size_row = rows[0]
                field_row = rows[1] if len(rows) > 1 else []
                split_idx = next((i for i in range(1, len(size_row)) if clean(size_row[i])), None)
                if split_idx is None:
                    continue
                cur_labels = [_size_label(v) if _norm(v) != 'total' else 'Total'
                              for v in size_row[split_idx:]]
                if field_names is None:
                    field_names = [_canon_field(v) for v in field_row[:split_idx]]
                seen_header = True
                data_rows = rows[2:]
            else:
                if not seen_header:
                    continue  # সামারি টেবিল (হেডারের আগে)
                data_rows = rows

            for r in data_rows:
                handle_row(r)

    if not seen_header:
        raise ValueError("এই PDF-এ পরিচিত 'Purchase Order Details' (Size/Measurement) টেবিল পাওয়া যায়নি।")

    raw_rows = []
    for e in entries:
        raw_rows.append({**e['meta'], **{f'Size {k}': v for k, v in e['sizes'].items()}})
    return entries, pd.DataFrame(raw_rows)


def to_canonical(entries, summary_df):
    single_rate, lookup = _summary_lookup(summary_df)
    items = []
    for e in entries:
        m = e['meta']
        ref_parts = [m['Item Name'], m['Color Code'], m['USIM / Ref. No']]
        reference = '/'.join(p for p in ref_parts if p)
        if single_rate is not None:
            rate = single_rate
        else:
            rate = lookup.get((_norm(m['Item Name']), _norm(m['USIM / Ref. No'])), '')
        for size, qty in e['sizes'].items():
            if qty is None or qty <= 0:
                continue
            items.append({
                'ewo_no': m['EWO No'],
                'style_no': m['Style No'],
                'po_no': m['PO No'],
                'reference': reference,
                'size': size,
                'qty': int(qty) if qty == int(qty) else qty,
                'uom': m['UOM'] or 'Pcs',
                'rate': rate,
                'item_name_pdf': m['Item Name'],
                'usim_pdf': m['USIM / Ref. No'],
                'delivery_place_pdf': m['Delivery Place'],
                'delivery_address_pdf': m['Delivery Address'],
            })
    return items


def compute_qty_warnings(line_items, summary_df):
    """সামারি টেবিলের (Label Type, USIM) -> POQty এর সাথে এক্সট্র্যাক্ট করা qty মেলানো।"""
    warnings = []
    if summary_df is None or summary_df.empty:
        return warnings
    qty_col = next((c for c in summary_df.columns if _norm(c) in ('poqty', 'qty')), None)
    usim_col = next((c for c in summary_df.columns if _norm(c).startswith('usim')), None)
    if not qty_col:
        return warnings

    total_pdf = sum((_to_float(v) or 0) for v in summary_df[qty_col])
    total_ext = sum((_to_float(i['qty']) or 0) for i in line_items)
    if abs(total_pdf - total_ext) > 0.01:
        warnings.append(
            f"⚠️ মোট Qty মিলছে না — PDF সামারিতে {total_pdf:,.0f}, কিন্তু Excel-এ {total_ext:,.0f} "
            f"(পার্থক্য {total_pdf - total_ext:+,.0f})। নিচের আইটেম-ভিত্তিক বিস্তারিত দেখুন।"
        )

    if usim_col:
        ext = {}
        for i in line_items:
            k = (_norm(i['item_name_pdf']), _norm(i['usim_pdf']))
            ext[k] = ext.get(k, 0) + (_to_float(i['qty']) or 0)
        type_col = summary_df.columns[0]
        for _, r in summary_df.iterrows():
            k = (_norm(r[type_col]), _norm(r[usim_col]))
            pdf_q = _to_float(r[qty_col]) or 0
            got = ext.get(k, 0)
            if abs(pdf_q - got) > 0.01:
                warnings.append(
                    f"⚠️ '{clean(r[type_col])} / {clean(r[usim_col])}' — PDF সামারিতে {pdf_q:,.0f}, "
                    f"Excel-এ {got:,.0f} (পার্থক্য {pdf_q - got:+,.0f})"
                )
    return warnings


def get_unique_delivery_info(raw_df):
    def uniques(col):
        if raw_df is None or raw_df.empty or col not in raw_df.columns:
            return []
        seen = []
        for v in raw_df[col]:
            v = clean(v)
            if v and v not in seen:
                seen.append(v)
        return seen
    return {'delivery_places': uniques('Delivery Place'), 'delivery_addresses': uniques('Delivery Address')}


def process_pdf_label(file_stream):
    """Returns (header_info, line_items, raw_df, summary_df)."""
    with pdfplumber.open(file_stream) as pdf:
        header_info = extract_header_info(pdf)
        summary_df = extract_summary_table(pdf)
        entries, raw_df = extract_detail_rows(pdf)
    line_items = to_canonical(entries, summary_df)
    return header_info, line_items, raw_df, summary_df