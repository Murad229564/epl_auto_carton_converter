"""
C&A বায়ারের কমন নিয়ম — কাস্টমার-নির্বিশেষে একই। প্রতিটা কাস্টমার
extractor শুধু raw_item ডিক্ট বানায় (কাঁচা style/code-text/measurement/
qty), তারপর এই ফাংশন C&A-স্পেসিফিক সব নিয়ম (কোড-ম্যাচ, এক্সেপশনাল-প্রাইস,
/TRI, qty রাউন্ড) একবারই প্রয়োগ করে — নতুন কাস্টমার যোগ করলে এই ফাইলে
কিছু বদলাতে হবে না।

raw_item-এর প্রত্যাশিত key:
    raw_style          (str, আবশ্যক) — /TRI যোগ করার আগের স্টাইল টেক্সট
    raw_code_text      (str, না থাকলে '') — কোড-বহনকারী কলামের raw টেক্সট
    length_cm/width_cm/height_cm  (সংখ্যা, আবশ্যক) — ইতিমধ্যে cm-এ
    qty                (সংখ্যা, আবশ্যক)
    ewo_no             (str, ঐচ্ছিক)
    ply                (str, ঐচ্ছিক — না দিলে ca_config.DEFAULT_PLY)
    item_name          (str, ঐচ্ছিক — না দিলে 'Master Carton')
    pack_type/reference/delivery_date  (ঐচ্ছিক)
    _sheet/_source_file (ঐচ্ছিক, ট্রেসিং-এর জন্য)
"""
from . import ca_config as cfg
from .ca_calc import match_code, check_measurement_match, exceptional_price, find_codes_by_measurement


def _fmt_num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return ''
    return str(int(f)) if f == int(f) else str(round(f, 3))


def mm_from_cm(v):
    return round(float(v) * 10)


# বাস্তবসম্মত কার্টুনের কোনো দিকই সচরাচর এর বেশি সেন্টিমিটার হয় না
# (আমাদের C&A প্রাইস-লিস্টে সবচেয়ে বড় মাপ ~৭৫০mm = ৭৫cm) — তাই raw
# ইনপুটের সবচেয়ে বড় সংখ্যাটা এর বেশি হলে ধরে নেওয়া হয় ইনপুট আগে
# থেকেই mm-এ দেওয়া, কম হলে cm।
_MM_DETECT_THRESHOLD = 150


def _normalize_measurement(l, w, h):
    """raw L/W/H (এককের নিশ্চয়তা নেই — কিছু কাস্টমার cm দেয়, কিছু সরাসরি
    mm) থেকে (l_cm, w_cm, h_cm, l_mm, w_mm, h_mm, was_mm) রিটার্ন করে।
    was_mm True মানে ইনপুট mm হিসেবে ধরে নেওয়া হয়েছে (ওয়ার্নিং দেওয়ার
    জন্য কাজে লাগে, যাতে ইউজার চোখ বুলিয়ে নিশ্চিত হতে পারেন)।"""
    l, w, h = float(l), float(w), float(h)
    if max(l, w, h) > _MM_DETECT_THRESHOLD:
        l_mm, w_mm, h_mm = round(l), round(w), round(h)
        l_cm, w_cm, h_cm = l / 10, w / 10, h / 10
        return l_cm, w_cm, h_cm, l_mm, w_mm, h_mm, True
    l_mm, w_mm, h_mm = mm_from_cm(l), mm_from_cm(w), mm_from_cm(h)
    return l, w, h, l_mm, w_mm, h_mm, False


def apply_ca_rules(raw_item, lookup, item_name_override=None, manual_ply=None):
    """একটা raw_item-কে canonical লাইন-আইটেম ডিক্টে রূপান্তর করে।
    item_name_override/manual_ply দেওয়া থাকলে (UI থেকে) raw_item-এর
    নিজস্ব item_name/ply-এর চেয়ে অগ্রাধিকার পায় — ইউজার-কনফার্মড নিয়ম
    অনুযায়ী এই বায়ারের ক্ষেত্রে Item Name/Ply সবসময় UI থেকেই আসে।
    রিটার্ন করে (line_item_dict, warnings_list)।"""
    warnings = []

    raw_style = str(raw_item.get('raw_style') or '').strip()
    style_no = f"{raw_style}{cfg.STYLE_SUFFIX}" if raw_style else (cfg.EXCEPTIONAL_TEXT)

    l_raw = raw_item.get('length_cm')
    w_raw = raw_item.get('width_cm')
    h_raw = raw_item.get('height_cm')
    if l_raw in (None, '') or w_raw in (None, '') or h_raw in (None, ''):
        warnings.append(f"স্টাইল '{raw_style}': measurement পাওয়া যায়নি — এই রো স্কিপ করা হয়েছে।")
        return None, warnings

    l_cm, w_cm, h_cm, l_mm, w_mm, h_mm, was_mm = _normalize_measurement(l_raw, w_raw, h_raw)
    if was_mm:
        warnings.append(
            f"স্টাইল '{raw_style}': মাপ ({l_raw}x{w_raw}x{h_raw}) mm ধরে নিয়ে "
            f"{l_cm:g}x{w_cm:g}x{h_cm:g} cm-এ কনভার্ট করা হয়েছে (সংখ্যা বড় দেখে "
            f"স্বয়ংক্রিয়ভাবে বোঝা হয়েছে) — একবার মিলিয়ে নিন এটা ঠিক আছে কিনা।"
        )

    raw_code_text = raw_item.get('raw_code_text', '')
    entry, code_warn = match_code(raw_code_text, lookup)
    if code_warn:
        warnings.append(f"স্টাইল '{raw_style}': {code_warn}")

    if entry is not None:
        po_field = entry['code']
        remarks_val = entry['price']
        mism = check_measurement_match(entry, l_mm, w_mm, h_mm)
        if mism:
            warnings.append(f"স্টাইল '{raw_style}': {mism}")
    else:
        po_field = cfg.EXCEPTIONAL_TEXT
        remarks_val = str(exceptional_price(l_mm, w_mm, h_mm))
        # কোনো কোড ম্যাচ হয়নি (হয়তো ফাঁকা/Regular/Irregular ছিল) — কিন্তু
        # তার আগেও এই মাপটা আমাদের প্রাইস-লিস্টের কোনো পরিচিত কোডের সাথে
        # হুবহু মিলে যাচ্ছে কিনা চেক করা হয় (ইউজার-কনফার্মড: এমন হলে
        # অবশ্যই ওয়ার্ন করতে হবে, যাতে বুঝা যায় বুকিং-এ কোড লিখতে হয়তো
        # ভুলে গেছে)। কোড অটোমেটিক বসানো হয় না, শুধু সতর্ক করা হয়।
        possible_codes = find_codes_by_measurement(l_mm, w_mm, h_mm, lookup)
        if possible_codes:
            codes_str = ', '.join(possible_codes)
            warnings.append(
                f"⚠️ স্টাইল '{raw_style}': এই মাপ ({l_mm}x{w_mm}x{h_mm} mm) আমাদের "
                f"C&A প্রাইস লিস্টের '{codes_str}' কোডের সাথে হুবহু মিলে যাচ্ছে, "
                f"কিন্তু বুকিং-এ কোনো কোড লেখা ছিল না (বা যা ছিল তা লিস্টে পাওয়া "
                f"যায়নি) — তাই EXCEPTIONAL MEASUREMENT বসানো হয়েছে। ম্যানুয়ালি "
                f"চেক করে দেখুন সঠিক কোডটা বসানো উচিত কিনা।"
            )

    qty_val = raw_item.get('qty')
    try:
        qty_final = round(float(qty_val))
    except (TypeError, ValueError):
        warnings.append(f"স্টাইল '{raw_style}': Qty পাওয়া যায়নি/সংখ্যা না — এই রো স্কিপ করা হয়েছে।")
        return None, warnings

    line_item = {
        'item_name': item_name_override or raw_item.get('item_name') or 'Master Carton',
        'ewo_no': raw_item.get('ewo_no') or 'N/A',
        'style_no': style_no,
        'po_no': po_field,
        'length': _fmt_num(l_cm),
        'width': _fmt_num(w_cm),
        'height': _fmt_num(h_cm),
        'ply': manual_ply or raw_item.get('ply') or cfg.DEFAULT_PLY,
        'qty': qty_final,
        'pack_type': raw_item.get('pack_type', ''),
        'reference': raw_item.get('reference', ''),
        'remarks': remarks_val,
        'color': '',
        'size': '',
        'delivery_date': raw_item.get('delivery_date', ''),
        'measurement_unit': 'Cm',
        'delivery_place_pdf': '',
        'delivery_address_pdf': '',
        '_sheet': raw_item.get('_sheet', ''),
        '_source_file': raw_item.get('_source_file', ''),
    }
    return line_item, warnings


def check_price_list_validity(price_list):
    """প্রাইস-লিস্টের 'valid_until' তারিখ পার হয়ে গেলে একটা ওয়ার্নিং
    রিটার্ন করে (থাকলে), নাহলে None।"""
    import datetime
    valid_until = price_list.get('valid_until')
    if not valid_until:
        return None
    try:
        d = datetime.date.fromisoformat(valid_until)
    except ValueError:
        return None
    if datetime.date.today() > d:
        return (
            f"⚠️ C&A প্রাইস লিস্টের মেয়াদ ({valid_until}) পার হয়ে গেছে — "
            f"নতুন কোটেশন PDF দিয়ে buyers_ca/build_price_list.py আবার চালান।"
        )
    return None