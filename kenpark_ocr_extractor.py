import io
import re
import fitz
import pytesseract
from pytesseract import Output
from PIL import Image

from kenpark_extractor import (
    clean, _group_rows_by_top, _col_for_x, _is_noise_line,
    _SIZE_RE, _PLY_RE, _BUYER_LINE_RE, _DELIVERY_DATE_RE,
    _extract_top_po, _extract_customer, _COLS,
)

# ---------------------------------------------------------------------------
# স্ক্যান করা/ছবি-PDF (কোনো সিলেক্টেবল টেক্সট লেয়ার নেই) Kenpark বুকিং-এর
# জন্য OCR-বেসড ফলব্যাক। kenpark_extractor.py-এর সাথে যতটা সম্ভব একই লজিক
# (কলাম-ব্যান্ড, রো-গ্রুপিং, ক্লাসিফিকেশন) শেয়ার করে — শুধু শব্দগুলো
# pdfplumber-এর বদলে Tesseract OCR থেকে আসে।
#
# ⚠️ গুরুত্বপূর্ণ সতর্কতা — এটা kenpark_extractor.py-এর মতো নির্ভরযোগ্য না:
#   - OCR সংখ্যায় ভুল করতে পারে (যেমন 0↔8, 1↔7, 5↔6, 3↔8) — বিশেষ করে
#     Measurement আর Qty-র মতো সংখ্যাসূচক ফিল্ডে এই ভুল ধরা কঠিন, কারণ
#     ভুল সংখ্যাও দেখতে "স্বাভাবিক" লাগে।
#   - তাই এই ফাংশন থেকে পাওয়া প্রতিটা লাইন-আইটেম **অবশ্যই** মূল PDF পাতার
#     ছবির সাথে মিলিয়ে ম্যানুয়ালি যাচাই করতে হবে, বিশেষ করে Length/Width/
#     Height/Qty। এই কারণেই read_kenpark_pdf_ocr() পাতার রেন্ডার করা ছবিও
#     সাথে রিটার্ন করে, যাতে পাশাপাশি রেখে চোখে যাচাই করা যায়।
#   - সম্ভব হলে সবসময় ডিজিটালি-জেনারেট করা (স্ক্যান না) PDF ব্যবহার করাই
#     ভালো — সেক্ষেত্রে kenpark_extractor.py (pdfplumber-বেসড) ব্যবহার হবে,
#     যেটা অনেক বেশি নির্ভরযোগ্য।
# ---------------------------------------------------------------------------


def _detect_dynamic_cols(lines):
    """স্ক্যান কপিতে কলামগুলোর পজিশন ফিক্সড টেক্সট-PDF-এর সাথে নাও মিলতে
    পারে (স্ক্যানার/রেজোলিউশন-ভেদে সামান্য শিফট বা স্কেল পরিবর্তন হতে পারে) —
    তাই OCR পাথে প্রতিটা ফাইলের হেডার-রো থেকে নিজে থেকেই কলাম-ব্যান্ড বের
    করা হয়, fixed _COLS-এর ওপর ভরসা করা হয় না। হেডার-রো না পাওয়া গেলে
    (max_scan-এর মধ্যে) None রিটার্ন করে, তখন caller fixed _COLS-এ ফলব্যাক
    করবে (নিরাপদ ডিফল্ট হিসেবে)।"""
    anchors = ['description', 'color', 'size', 'unit', 'qty']
    for line in lines:
        found = {}
        for w in line:
            key = w['text'].strip('.:%').lower()
            if key in anchors and key not in found:
                found[key] = w['x0']
        if len(found) >= 4:  # অন্তত ৪টা anchor পাওয়া গেলে এটাই হেডার-রো ধরে নেওয়া হচ্ছে
            ordered = sorted(found.items(), key=lambda kv: kv[1])
            xs = [x for _, x in ordered]
            names = [k for k, _ in ordered]
            bounds = [xs[0] - 15] + [(xs[i] + xs[i + 1]) / 2 for i in range(len(xs) - 1)] + [xs[-1] + 60]
            cols = {}
            for i, name in enumerate(names):
                cols[name] = (bounds[i], bounds[i + 1])
            if 'description' in cols:
                cols['desc'] = cols.pop('description')
            return cols
    return None


def _col_for_x_dynamic(x0, cols):
    for name, (lo, hi) in cols.items():
        if lo <= x0 < hi:
            return name
    return None


def _page_to_words(page, dpi=300):
    """একটা PDF পাতাকে ছবিতে রেন্ডার করে Tesseract দিয়ে OCR করে, এবং
    pdfplumber-এর extract_words()-এর মতোই {'top','x0','text'} ফরম্যাটে
    শব্দ-লিস্ট রিটার্ন করে (পিক্সেল কোঅর্ডিনেট PDF পয়েন্টে কনভার্ট করে,
    যাতে kenpark_extractor.py-এর একই _COLS ব্যান্ড পুনর্ব্যবহার করা যায়)।"""
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
    data = pytesseract.image_to_data(img, output_type=Output.DICT)
    scale = 72.0 / dpi
    words = []
    for i in range(len(data['text'])):
        text = data['text'][i].strip()
        if not text:
            continue
        words.append({
            'top': data['top'][i] * scale,
            'x0': data['left'][i] * scale,
            'text': text,
        })
    return words, img


def read_kenpark_pdf_ocr(file_stream, filename='', dpi=300):
    """OCR-বেসড entry point। রিটার্ন করে (header_info, items, page_images) —
    page_images হলো প্রতিটা পাতার রেন্ডার করা PIL Image, যাতে ম্যানুয়াল
    যাচাইয়ের জন্য পাশাপাশি দেখানো যায়। items-এর canonical schema
    kenpark_extractor.py-এর সাথে হুবহু এক।"""
    doc = fitz.open(stream=file_stream.read(), filetype='pdf')
    if doc.page_count == 0:
        return {'po_number': '', 'customer': '', 'buyer': ''}, [], []

    page_images = []
    buyer = style_no = po_no = ''
    delivery_date = ''
    items = []
    top_po = ''
    customer = ''

    current_row = None

    def flush_row():
        nonlocal current_row
        if current_row is None:
            return
        desc = clean(' '.join(current_row.get('desc', [])))
        size_text = ''.join(current_row.get('size', []))
        # OCR-এ সাধারণ ভুল: 'O' (অক্ষর) বনাম '0' (সংখ্যা), এবং 'X'
        # ডাবল-ডিটেক্ট হয়ে যাওয়া — এই ফিল্ডে 'O' অক্ষর কখনো বৈধভাবে আসার
        # কথা না (শুধু L/W/H/X/C/M আর সংখ্যা), তাই নিরাপদে normalize করা হচ্ছে।
        size_text = size_text.upper().replace('O', '0')
        size_text = re.sub(r'X{2,}', 'X', size_text)
        qty_tokens = current_row.get('qty', [])
        qty_text = clean(qty_tokens[0]) if qty_tokens else ''

        desc_lower = desc.lower()
        if desc_lower.startswith('carton'):
            item_name = 'Master Carton'
        elif desc_lower.startswith('divider'):
            item_name = 'Divider'
        else:
            current_row = None
            return

        m = _SIZE_RE.search(size_text)
        if not m:
            current_row = None
            return
        length, width, height = m.group(1), m.group(2), m.group(3) or ''

        ply_m = _PLY_RE.search(desc)
        ply = ply_m.group(1) if ply_m else ''

        try:
            qty_val = float(qty_text.replace(',', '').replace('O', '0'))
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
            '_needs_verification': True,  # OCR-বেসড — ম্যানুয়াল চেক ছাড়া ফাইনাল না করার জন্য মার্ক করা
        })
        current_row = None

    for page_index in range(doc.page_count):
        page = doc[page_index]
        words, img = _page_to_words(page, dpi=dpi)
        page_images.append(img)

        if page_index == 0:
            full_text = '\n'.join(w['text'] for w in sorted(words, key=lambda w: (w['top'], w['x0'])))
            top_po = _extract_top_po(full_text)
            customer = _extract_customer(full_text)

        lines = _group_rows_by_top(words)

        if page_index == 0:
            cols = _detect_dynamic_cols(lines) or _COLS
            col_lookup = lambda x0: _col_for_x_dynamic(x0, cols)  # noqa: E731

        for line_words in lines:
            line_text = clean(' '.join(w['text'] for w in line_words))

            if line_words[0]['top'] < 65:
                continue

            if _is_noise_line(line_text):
                continue

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

            # নতুন আইটেম-রো শুরু: fixed কলাম-ব্যান্ডের বদলে সরাসরি "লেফটমোস্ট
            # শব্দটা কি pure digit" চেক করা হচ্ছে — OCR-এ কলাম-পজিশন সামান্য
            # শিফট হতে পারে, কিন্তু row-নম্বর সবসময় pure digit আর item
            # code/description কখনো pure digit হয় না, তাই এটাই সবচেয়ে
            # নির্ভরযোগ্য সিগনাল।
            first_w = line_words[0]
            if first_w['text'].isdigit():
                flush_row()
                current_row = {}

            if current_row is None:
                continue

            for w in line_words:
                col = col_lookup(w['x0'])
                if col == 'color':
                    current_row.setdefault('desc', []).append(w['text'])
                elif col in ('desc', 'size', 'qty'):
                    current_row.setdefault(col, []).append(w['text'])

    flush_row()

    header_info = {'po_number': top_po, 'customer': customer, 'buyer': buyer}
    return header_info, items, page_images
