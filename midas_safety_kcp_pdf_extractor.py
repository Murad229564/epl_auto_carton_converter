import re
import pandas as pd

# ---------------------------------------------------------------------------
# MIDAS SAFETY BANGLADESH LTD. — Buyer: KCP
# 'PURCHASE ORDER IMPORTS' PDF ফরম্যাট (SAP-জেনারেটেড PO)। এক PDF = একটা PO,
# একাধিক পেজে টেবিল চলতে পারে (পরের পেজে হেডার-রো আবার আসে না, শুধু ডাটা-রো)।
#
# টেবিল-কলাম: S.No. / PO Item No. / Item Code / Description / Planned ETD /
# ETA Factory / Quantity / Unit / Unit Price / Amount
#   Item Code    -> Gmt. Style No
#   Description  -> Gmt. PO  (পুরো লেখা, লাইন-ভাঙা অংশ জোড়া লাগিয়ে:
#                   '5Y CTN(39X28X56)CM-BROW-SHELBY-6300')
#   Description থেকে  -> Ply ('5Y' = 5 Ply, '3Y' = 3 Ply) এবং
#                   Length/Width/Height ('CTN(39X28X56)CM'; ব্র্যাকেট ছাড়া
#                   'CTN38X30X51CM' ও ছোট হাতের x-ও চলে), unit = Cm
#   Quantity     -> Order Qty
#   Unit Price   -> Remarks (পরে মেলানোর জন্য, যেমন '0.960')
#   Item Name    -> UI থেকে
#   'P.O. Number / Date / Version: 1000616681 / 05.11.2026 /' -> স্ল্যাশের
#                   আগের নম্বর (1000616681) = উপরের PO ফিল্ড (_header_po)
#
# ভুল এড়াতে ডাবল-চেক (সবকিছুর জন্য Warning আসে, রো বাদ যায় না):
#   - মেজারমেন্ট দুইটা আলাদা পদ্ধতিতে বের করে মেলানো হয় (CTN(...)CM প্যাটার্ন
#     বনাম Description-এর ভেতরের সব 'N x N x N' ট্রিপল); না মিললে, একাধিক
#     ট্রিপল থাকলে, বা 'CM' না থাকলে Warning।
#   - Ply দুইটা জায়গা থেকে ('5Y' এবং 'N PLY' লেখা); পাওয়া না গেলে, দুটো
#     আলাদা হলে, বা 3/5 ছাড়া অন্য কিছু হলে Warning। না পেলে Ply = 'N/A'।
#   - Description-এ WHITE / WH / WHT থাকলে Warning (সাধারণত এটা White
#     Carton — Brown Carton না; গ্রেড আলাদা হয়, তাই যাচাই করতে বলা হয়)।
#   - Quantity x Unit Price != Amount হলে, বা রো-গুলোর Amount-এর যোগফল
#     'Net Amount'-এর সাথে না মিললে Warning।
#   - কোনো রো পার্স করা না গেলে Warning (কোনো রো চুপচাপ হারিয়ে যায় না)।
# ---------------------------------------------------------------------------

NA = 'N/A'

_ROW_RE = re.compile(
    r'^\s*(\d+)\s+(\d{3,6})\s+(\d{6,})\s+(.*?)\s*'
    r'(\d{2}\.\d{2}\.\d{4})\s+(\d+(?:\.\d+)?)\s+([A-Za-z]{1,5})\s+'
    r'(\d+(?:,\d{3})*(?:\.\d+)?)\s+(\d+(?:,\d{3})*(?:\.\d+)?)\s*$')

# ডাটা-রো-র শুরুর মতো দেখতে কিন্তু রো-র শেষ ম্যাচ করেনি — হারানো রো ধরার জন্য
_ROW_START_RE = re.compile(r'^\s*\d+\s+\d{3,6}\s+\d{6,}\b')

_STOP_PREFIXES = ('net amount', 'grand total', 'remarks', 'terms', 'documents',
                  'delivery at', 'shipping marks', 'sap system')

_DIM_RE = re.compile(
    r'CTN\s*\(?\s*(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)\s*\)?\s*(CM|MM|INCH|IN)?',
    re.I)
_ANY_TRIPLE_RE = re.compile(
    r'(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)')
_PLY_Y_RE = re.compile(r'^\s*(\d)\s*Y\b', re.I)
_PLY_WORD_RE = re.compile(r'(\d)\s*-?\s*PLY', re.I)
_WHITE_TOKENS = {'WHITE', 'WH', 'WHT', 'WHITES'}


def _f(s):
    try:
        return float(str(s).replace(',', ''))
    except (TypeError, ValueError):
        return None


def _fmt(f):
    return int(f) if f == int(f) else f


def _join_description(parts):
    """লাইন-ভাঙা অংশ জোড়া: আগের অংশ '-' দিয়ে শেষ হলে সরাসরি (হাইফেন রেখে),
    নাহলে একটা স্পেস দিয়ে।"""
    out = ''
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if not out:
            out = p
        elif out.endswith('-'):
            out += p
        else:
            out += ' ' + p
    return out


def _page_texts(file_stream):
    import pdfplumber
    file_stream.seek(0)
    texts = []
    with pdfplumber.open(file_stream) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text() or '')
    return texts


def _header_po(texts):
    for t in texts:
        m = re.search(r'(\d{6,})\s*/\s*\d{2}\.\d{2}\.\d{4}', t)
        if m:
            return m.group(1)
    return ''


def _extract_measurement(desc):
    """(L, W, H, warnings_list) — দুই পদ্ধতিতে মিলিয়ে।"""
    warns = []
    m = _DIM_RE.search(desc)
    triples = _ANY_TRIPLE_RE.findall(desc)
    if not m:
        warns.append("মেজারমেন্ট (CTN(LxWxH)CM) পাওয়া যায়নি")
        return '', '', '', warns
    L, W, H = (_fmt(float(m.group(i))) for i in (1, 2, 3))
    unit = (m.group(4) or '').upper()
    if unit != 'CM':
        warns.append(f"মেজারমেন্টের পরে 'CM' পাওয়া যায়নি (পাওয়া গেছে: '{unit or 'কিছুই না'}') — ইউনিট যাচাই করুন")
    if len(triples) != 1:
        warns.append(f"Description-এ {len(triples)}টা 'LxWxH' ট্রিপল পাওয়া গেছে — মেজারমেন্ট হাতে যাচাই করুন")
    else:
        t = tuple(_fmt(float(x)) for x in triples[0])
        if t != (L, W, H):
            warns.append(f"মেজারমেন্ট দুই পদ্ধতিতে মেলেনি ({(L, W, H)} বনাম {t}) — যাচাই করুন")
    if any(not (1 <= v <= 300) for v in (L, W, H)):
        warns.append(f"মেজারমেন্ট অস্বাভাবিক ({L}x{W}x{H}) — যাচাই করুন")
    return L, W, H, warns


def _extract_ply(desc):
    warns = []
    y = _PLY_Y_RE.search(desc)
    w = _PLY_WORD_RE.search(desc)
    ply_y = y.group(1) if y else ''
    ply_w = w.group(1) if w else ''
    if ply_y and ply_w and ply_y != ply_w:
        warns.append(f"Ply দুই জায়গায় আলাদা ('{ply_y}Y' বনাম '{ply_w} PLY') — যাচাই করুন")
    ply = ply_y or ply_w
    if not ply:
        warns.append("Ply ('5Y'/'3Y' বা 'N PLY') পাওয়া যায়নি — N/A বসানো হয়েছে")
        return NA, warns
    if ply not in ('3', '5'):
        warns.append(f"Ply '{ply}' (সাধারণত 3 বা 5) — যাচাই করুন")
    return ply, warns


def _is_white(desc):
    up = desc.upper()
    if 'WHITE' in up:
        return True
    tokens = set(re.split(r'[^A-Z0-9]+', up))
    return bool(tokens & _WHITE_TOKENS)


def _extract_midas_kcp(file_stream, filename='', item_name_override='Master Carton'):
    """এক PDF -> (items, warnings, rows, header_po, customer)। টেবিল না পেলে ValueError।"""
    texts = _page_texts(file_stream)
    if not any('PURCHASE ORDER' in t.upper() for t in texts):
        raise ValueError("'PURCHASE ORDER IMPORTS' ফরম্যাট পাওয়া যায়নি")

    header_po = _header_po(texts)
    item_name = item_name_override or 'Master Carton'
    warnings = []
    if not header_po:
        warnings.append(f"{filename}: 'P.O. Number' (স্ল্যাশের আগের নম্বর) পাওয়া যায়নি")

    rows = []          # {'sno','poitem','code','desc_parts','qty','unit','price','amount'}
    net_amount = None
    for page_no, text in enumerate(texts, start=1):
        current = None
        for line in text.splitlines():
            s = line.strip()
            if not s:
                continue
            m = _ROW_RE.match(s)
            if m:
                current = {
                    'sno': m.group(1), 'poitem': m.group(2), 'code': m.group(3),
                    'desc_parts': [m.group(4)], 'qty': _f(m.group(6)),
                    'unit': m.group(7), 'price_txt': m.group(8),
                    'amount': _f(m.group(9)), 'page': page_no,
                    'eta': m.group(5), 'amount_txt': m.group(9),
                }
                rows.append(current)
                continue
            low = s.lower()
            if low.startswith('net amount'):
                mm = re.search(r'(\d+(?:,\d{3})*\.\d+)', s)
                net_amount = _f(mm.group(1)) if mm else net_amount
            if low.startswith(_STOP_PREFIXES):
                current = None
                continue
            if _ROW_START_RE.match(s):
                warnings.append(f"{filename}: পেজ {page_no}-এ একটা রো পার্স করা যায়নি ('{s[:60]}') — হাতে যাচাই করুন")
                current = None
                continue
            if current is not None and not re.match(r'^(Midas Safety|Plot No|Chittagong|Tel:)', s):
                current['desc_parts'].append(s)

    if not rows:
        raise ValueError("টেবিলে কোনো ডাটা-রো পাওয়া যায়নি")

    items = []
    amount_sum = 0.0
    for r in rows:
        desc = _join_description(r['desc_parts'])
        who = f"{filename} [S.No {r['sno']} / Item {r['code']}]"
        L, W, H, w1 = _extract_measurement(desc)
        ply, w2 = _extract_ply(desc)
        for w in w1 + w2:
            warnings.append(f"{who}: {w}")
        if _is_white(desc):
            warnings.append(f"{who}: Description-এ WHITE/WH পাওয়া গেছে — এটা সম্ভবত White Carton (Brown না), গ্রেড আলাদা হয়, যাচাই করুন → {desc}")

        qty = r['qty']
        if qty is None or qty <= 0:
            warnings.append(f"{who}: Quantity পাওয়া যায়নি/0 — রো বাদ দেওয়া হয়েছে")
            continue
        price = _f(r['price_txt'])
        if price is not None and r['amount'] is not None:
            if abs(qty * price - r['amount']) > 0.011:
                warnings.append(f"{who}: Qty x Unit Price ({qty:g} x {r['price_txt']}) ≠ Amount ({r['amount']:.3f}) — যাচাই করুন")
        amount_sum += r['amount'] or 0

        items.append({
            'item_name': item_name,
            'ewo_no': NA,
            'style_no': r['code'],
            'po_no': desc,
            'length': L,
            'width': W,
            'height': H,
            'ply': ply,
            'qty': round(qty),
            'pack_type': '',
            'reference': '',
            'color': '',
            'size': '',
            'remarks': r['price_txt'],
            'delivery_date': '',
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_header_po': header_po,
            '_source_file': filename,
        })

    if net_amount is not None and abs(amount_sum - net_amount) > 0.011:
        warnings.append(f"{filename}: রো-গুলোর Amount-এর যোগফল ({amount_sum:.3f}) ≠ Net Amount ({net_amount:.3f}) — কোনো রো মিসিং হতে পারে")

    customer = (texts[0].splitlines()[0].strip() if texts and texts[0].strip() else '')
    return items, warnings, rows, header_po, customer


def extract_midas_kcp_pdf(file_stream, filename='', item_name_override='Master Carton'):
    """এক PDF -> (items, warnings)।"""
    items, warnings, _rows, _po, _cust = _extract_midas_kcp(file_stream, filename, item_name_override)
    return items, warnings


def process_pdf_midas_kcp(file_stream, filename='', item_name_override='Master Carton'):
    """আপনার বিদ্যমান process_pdf_rule_based()-এর হুবহু একই আকারে রিটার্ন করে:
    (header_info, line_items, raw_original_df, summary_df) — তাই যে জায়গায়
    process_pdf_rule_based() কল হয়, সেখানে এটা সরাসরি বদলে বসানো যায়।
    header_info-তে বাড়তি 'warnings' (লিস্ট) আছে — Warnings শীটে দেখানোর জন্য।"""
    items, warnings, rows, po, customer = _extract_midas_kcp(file_stream, filename, item_name_override)
    raw_df = pd.DataFrame(
        [[r['sno'], r['poitem'], r['code'], _join_description(r['desc_parts']), r['eta'],
          ('%g' % r['qty']) if r['qty'] is not None else '', r['unit'],
          r['price_txt'], r['amount_txt']] for r in rows],
        columns=['S.No.', 'PO Item No.', 'Item Code', 'Description', 'ETA Factory',
                 'Quantity', 'Unit', 'Unit Price', 'Amount'])
    header_info = {'po_number': po, 'customer': customer, 'buyer': '', 'warnings': warnings}
    return header_info, items, raw_df, pd.DataFrame()


def combine_midas_kcp_pdf_files(files, item_name_override=''):
    """files: [(file_stream, filename), ...] -> (line_items, warnings)
    ফাইল-ক্রমে। ব্যর্থ ফাইল স্কিপ হয়ে মেসেজ warnings-এ যায়।"""
    combined, warnings = [], []
    for file_stream, filename in files:
        try:
            its, w = extract_midas_kcp_pdf(file_stream, filename, item_name_override=item_name_override)
            combined.extend(its)
            warnings.extend(w)
        except Exception as e:
            warnings.append(f"{filename}: {e}")
    return combined, warnings


def read_midas_kcp_pdf(file_stream, filename='', item_name_override='', manual_ply='', customer_list=None):
    """app.py-এর OUTHOUSE_PDF_REGISTRY-র জন্য এন্ট্রি-পয়েন্ট — read_ikl_biscana_pdf-এর
    মতোই (header_info, items) রিটার্ন করে। manual_ply এখানে ব্যবহার হয় না
    (Ply Description-এর '5Y'/'3Y' থেকেই আসে)।

    header_info: {'po_number', 'customer', 'buyer', 'warnings'}
      - customer: PDF-এর লেটারহেড নাম, customer_list-এর সাথে case-insensitive
        মিললে ক্যানোনিকাল নাম (যেমন 'MIDAS SAFETY BANGLADESH LTD.'), না মিললে ফাঁকা
      - buyer: PDF-এ বায়ার নেই, তাই ফাঁকা (UI থেকে আসে)
      - warnings: এই ফাইলের সব Warning (লিস্ট)
    ফরম্যাট না মিললে ValueError ছোঁড়ে (app.py-এর try/except ধরে নেয়)।"""
    header_info, items, _raw_df, _summary_df = process_pdf_midas_kcp(
        file_stream, filename, item_name_override or 'Master Carton')

    def _key(s):
        return re.sub(r'[^a-z0-9]', '', str(s or '').lower())

    canonical = ''
    for c in (customer_list or []):
        if _key(c) == _key(header_info.get('customer')):
            canonical = c
            break
    header_info['customer'] = canonical
    return header_info, items
