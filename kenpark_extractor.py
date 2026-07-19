import re
import pdfplumber

# ---------------------------------------------------------------------------
# Kenpark Bangladesh Apparel (Pvt.) Limited / Kenpark Bangladesh (Pvt.) Limited
# — Buyer: Ralph Lauren — "PURCHASE ORDER (LOCAL FE)" PDF ফরম্যাট
# (Supplier: EPYLLION LIMITED)।
#
#   - এই PDF-এ Poly Bag আর Carton/Divider — দুই ধরনের লাইনই একসাথে থাকে।
#     শুধু Carton (item_name='Master Carton') আর Divider/টপ-বাটন
#     (item_name='Divider') লাইন রাখা হয়, Poly Bag লাইন বাদ দেওয়া হয়।
#   - টেবিলটা বর্ডার-লাইন ছাড়া (pdfplumber-এর extract_tables() এটা ধরতে
#     পারে না), এবং প্রতিটা সেল (Item number/Description/Size ইত্যাদি)
#     একাধিক টেক্সট-লাইনে wrap হয়ে যায় — তাই word-position (x0 কলাম-ব্যান্ড)
#     ভিত্তিক row-reconstruction ব্যবহার করা হয়েছে, শুধু extract_text()
#     রিজেক্স না।
#   - Style No এবং PO No — দুটোই "Buyer : ... / Style : ... / Sales order :
#     ..." লাইন থেকে (এই লাইন যতক্ষণ না বদলায় ততক্ষণ constant থাকে):
#       - Style -> style_no
#       - Sales order -> po_no
#   - পাতার উপরে-ডান কোণের মূল PO নম্বর (যেমন BDKAPO0114947) — এটা
#     লাইন-আইটেমের কোনো ফিল্ডে বসে না, বরং header_info-তে আলাদাভাবে
#     রিটার্ন হয় (UI-এর PO Number ফিল্ডে বসানোর জন্য, extractor.py-এর
#     extract_header_info()-এর মতোই কনভেনশন)। এই ফরম্যাটে আলাদা কোনো
#     EWO No না থাকায় প্রতিটা লাইন-আইটেমের ewo_no = 'N/A'।
#   - Pack Type/Reference/Color — এই কাস্টমারের জন্য নির্দিষ্ট কোনো ম্যাপিং
#     বলা হয়নি, তাই গেস না করে 'N/A' বসানো হয়েছে (ফাঁকা স্ট্রিং না)।
#   - Ply -> Description টেক্সট থেকে ডাইনামিকভাবে বের করা ('.. 5 Ply ..' /
#     '.. 3 Ply ..') — হার্ডকোড করা হয়নি, কারণ ভবিষ্যতে অন্য Ply-ও আসতে পারে।
#   - Measurement সবসময় CM-এ (টেক্সটে 'CM' সাফিক্স স্পষ্ট থাকে); Carton-এর
#     ক্ষেত্রে L×W×H, Divider-এর ক্ষেত্রে শুধু L×W (কোনো Height নেই)।
#   - Delivery Date -> 'Delivery date : ...' লাইন থেকে, যতক্ষণ না বদলায়
#     constant থাকে।
#   - এই ফরম্যাট শুধু তখনই কাজ করে যখন PDF-টা ডিজিটালি জেনারেট করা (সিলেক্টেবল
#     টেক্সট থাকে)। স্ক্যান করা/ছবি-PDF-এ (কোনো টেক্সট লেয়ার নেই) এই ফাংশন
#     খালি রেজাল্ট দেবে — এই কেসের জন্য আলাদা OCR-বেসড হেল্পার আছে
#     (kenpark_ocr_extractor.py দেখুন)।
# ---------------------------------------------------------------------------


def clean(v):
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v)).strip()


# 'L44XW30XH15CM' (Carton, H আছে) এবং 'L40XW28CM' (Divider, H নেই) — দুটোই
# একটা রেজেক্সে হ্যান্ডেল করা হয়েছে।
_SIZE_RE = re.compile(r'L(\d+(?:\.\d+)?)XW(\d+(?:\.\d+)?)(?:XH(\d+(?:\.\d+)?))?CM', re.I)
_PLY_RE = re.compile(r'(\d+)\s*Ply', re.I)
_TOP_PO_RE = re.compile(r'\b(BDKAPO\d+)\b', re.I)
_BUYER_LINE_RE = re.compile(
    r'Buyer\s*:\s*(.+?)\s*/\s*Style\s*:\s*(\S+)\s*/\s*Sales\s*order\s*:\s*(\S+)', re.I
)
_DELIVERY_DATE_RE = re.compile(r'Delivery\s*date\s*:\s*(.+)', re.I)

# প্রতিটা পাতার উপরে (PO নম্বর/কাস্টমার নাম/টাইটেল/কলাম-হেডার পুনরাবৃত্তি) আর
# শেষ পাতার নিচে (লিগ্যাল বয়লারপ্লেট/সামারি টেবিল) যে টেক্সট থাকে, সেগুলো
# ভুলবশত চলতি রো-র সাথে মিশে যাওয়া ঠেকাতে এই প্যাটার্নগুলো দেখলে লাইনটা
# উপেক্ষা করা হয় (কাস্টমার-নাম নির্দিষ্ট কোনো টেক্সট এখানে ব্যবহার করা হয়নি,
# যাতে ভবিষ্যতে এই একই লেআউটের অন্য কাস্টমারেও কাজ করে)।
_NOISE_PATTERNS = [
    _TOP_PO_RE,  # পাতার উপরের PO নম্বর পুনরাবৃত্তি (e.g. BDKAPO0114947)
    re.compile(r'PURCHASE\s*ORDER', re.I),
    re.compile(r'Item\s*number\s*Description', re.I),  # কলাম-হেডার পুনরাবৃত্তি
    re.compile(r'NO\s*PARTIAL\s*SHIPMENT', re.I),
    re.compile(r'Structure\s*Id', re.I),
    re.compile(r'Submitted\s*user', re.I),
    re.compile(r'Amount\s*in\s*words', re.I),
    re.compile(r'Grand\s*total', re.I),
    re.compile(r'Tax\s*Group', re.I),
    re.compile(r'committed\s*to\s*eliminating', re.I),
    re.compile(r'DIGITALLY\s*APPROVED', re.I),
    re.compile(r'Please\s*indicate\s*the\s*purchase\s*order', re.I),
]


def _is_noise_line(text):
    return any(p.search(text) for p in _NOISE_PATTERNS)

# হেডার রো-র শব্দ-পজিশন থেকে বের করা কলাম-ব্যান্ড (x0 রেঞ্জ)। ভবিষ্যতে যদি
# এই একই ফ্যামিলির অন্য PDF-এ সামান্য পজিশন শিফট হয়, তাহলে এই ব্যান্ডগুলো
# একটু চওড়া করে দিলেই চলবে।
_COLS = {
    'no': (0, 45),
    'item_no': (45, 102),
    'desc': (102, 152),
    'color': (152, 202),
    'size': (202, 238),
    'unit': (238, 278),
    'qty': (278, 330),
}


def _col_for_x(x0):
    for name, (lo, hi) in _COLS.items():
        if lo <= x0 < hi:
            return name
    return None


def _group_rows_by_top(words, tol=3):
    """একই ভিজ্যুয়াল টেক্সট-লাইনের শব্দগুলো এক গ্রুপে জড়ো করে (top একে
    অপরের কাছাকাছি হলে), এবং প্রতিটা গ্রুপের ভেতরে বাম থেকে ডানে (x0
    অনুযায়ী) সাজানো (OCR-এর ক্ষেত্রে একই ভিজ্যুয়াল লাইনের শব্দগুলোর 'top'
    সামান্য এদিক-ওদিক হতে পারে, তাই শুধু sort key (top, x0)-এর ওপর ভরসা
    করলে group-এর ভেতরে বাম-ডান ক্রম ভেঙে যেতে পারে — তাই গ্রুপিং-এর পর
    আলাদাভাবে x0 দিয়ে re-sort করা হয়েছে)।"""
    words = sorted(words, key=lambda w: w['top'])
    lines = []
    for w in words:
        if lines and abs(w['top'] - lines[-1][0]['top']) <= tol:
            lines[-1].append(w)
        else:
            lines.append([w])
    for line in lines:
        line.sort(key=lambda w: w['x0'])
    return lines


def _extract_top_po(first_page_text):
    m = _TOP_PO_RE.search(first_page_text or '')
    return m.group(1) if m else ''


def _extract_customer(first_page_text):
    """PO নম্বরের ঠিক পরের নন-এম্পটি লাইনটাই কাস্টমার নাম (এই লেআউটে সবসময়
    এই পজিশনে থাকে, কাস্টমার-নাম হার্ডকোড না করে)।"""
    lines = [clean(l) for l in (first_page_text or '').split('\n')]
    for i, line in enumerate(lines):
        if _TOP_PO_RE.search(line):
            for nxt in lines[i + 1:]:
                if nxt:
                    return nxt
            break
    return ''


def read_kenpark_pdf(file_stream, filename=''):
    """মূল entry point। রিটার্ন করে (header_info, items) — header_info-তে
    UI-এর জন্য po_number/customer/buyer থাকে (extractor.py-এর
    extract_header_info()-এর কনভেনশন অনুসরণ করে); items হলো লাইন-আইটেমের
    canonical dict লিস্ট। এই ফরম্যাট না হলে (কোনো লাইন-আইটেম না পাওয়া গেলে)
    items খালি লিস্ট [] হবে।"""
    with pdfplumber.open(file_stream) as pdf:
        if not pdf.pages:
            return {'po_number': '', 'customer': '', 'buyer': ''}, []

        first_page_text = pdf.pages[0].extract_text()
        top_po = _extract_top_po(first_page_text)
        customer = _extract_customer(first_page_text)

        buyer = style_no = po_no = ''
        delivery_date = ''
        items = []

        current_row = None  # চলতি রো-র কলাম-বাই-কলাম শব্দ জমা রাখার জায়গা

        def flush_row():
            nonlocal current_row
            if current_row is None:
                return
            desc = clean(' '.join(current_row.get('desc', [])))
            size_text = ''.join(current_row.get('size', []))  # স্পেস ছাড়া জোড়া, কোড ভাঙা না যায়
            qty_text = clean(' '.join(current_row.get('qty', [])))

            desc_lower = desc.lower()
            if desc_lower.startswith('carton'):
                item_name = 'Master Carton'
            elif desc_lower.startswith('divider'):
                item_name = 'Divider'
            else:
                current_row = None
                return  # Poly Bag বা অন্য কিছু — বাদ

            m = _SIZE_RE.search(size_text)
            if not m:
                current_row = None
                return  # measurement পার্স না হলে নিরাপদে স্কিপ (ভুল ডাটা ঢোকানোর চেয়ে বাদ দেওয়া ভালো)
            length, width, height = m.group(1), m.group(2), m.group(3) or ''

            ply_m = _PLY_RE.search(desc)
            ply = ply_m.group(1) if ply_m else ''

            try:
                qty_val = float(qty_text.replace(',', ''))
            except ValueError:
                current_row = None
                return

            items.append({
                'item_name': item_name,
                'ewo_no': 'N/A',
                'style_no': style_no,
                'po_no': po_no,
                'length': length,
                'width': width,
                'height': height,
                'ply': ply,
                'qty': qty_val,
                'pack_type': 'N/A',
                'reference': 'N/A',
                'remarks': '',
                'color': 'N/A',
                'size': '',
                'delivery_date': delivery_date,
                'measurement_unit': 'Cm',
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
            })
            current_row = None

        for page in pdf.pages:
            words = page.extract_words()
            lines = _group_rows_by_top(words)
            for line_words in lines:
                line_text = clean(' '.join(w['text'] for w in line_words))

                if line_words[0]['top'] < 65:
                    continue  # পাতার একদম উপরের হেডার-ব্লক (PO নম্বর/কাস্টমার নাম/টাইটেল) — উপেক্ষা

                if _is_noise_line(line_text):
                    continue  # পাতার পুনরাবৃত্ত হেডার/ফুটার — উপেক্ষা করা হচ্ছে

                if line_text.startswith('Buyer') and 'Style' in line_text and 'Sales order' in line_text:
                    flush_row()
                    bm = _BUYER_LINE_RE.search(line_text)
                    if bm:
                        buyer, style_no, po_no = clean(bm.group(1)), clean(bm.group(2)), clean(bm.group(3))
                    continue

                dm = _DELIVERY_DATE_RE.match(line_text)
                if dm:
                    flush_row()
                    delivery_date = clean(dm.group(1))
                    continue

                # নতুন আইটেম-রো শুরু: 'no' কলাম-ব্যান্ডে একটা pure-digit শব্দ
                first_w = line_words[0]
                if _col_for_x(first_w['x0']) == 'no' and first_w['text'].isdigit():
                    flush_row()
                    current_row = {}

                if current_row is None:
                    continue  # টেবিলের বাইরের লাইন (হেডার/ফুটার/ডিসক্লেইমার) — উপেক্ষা

                for w in line_words:
                    col = _col_for_x(w['x0'])
                    if col in ('desc', 'size', 'qty'):
                        current_row.setdefault(col, []).append(w['text'])

        flush_row()

    header_info = {'po_number': top_po, 'customer': customer, 'buyer': buyer}
    return header_info, items
