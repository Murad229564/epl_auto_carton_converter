import re
from datetime import datetime
import pandas as pd

# ---------------------------------------------------------------------------
# Columbia Apparels Limited (buyer: GU) — "CARTON BOOKING FOR G.U" এক্সেল
# ফরম্যাট। প্রতিটা ফাইলে একটাই Style/PO/Ship-To ব্লক থাকে, নিচে একটা
# টেবিলে Set Name-ওয়াইজ সারি — প্রতিটা সারিতে Auto Carton আর Carton-Top
# দুটো আলাদা মেজারমেন্ট/কোয়ান্টিটি একসাথে থাকে।
#
# তাই প্রতিটা ডাটা-রো থেকে টেমপ্লেটে দুইটা আলাদা লাইন-আইটেম বসে:
#   ১. Carton   -> item_name UI থেকে সিলেক্ট করা (Master Carton ইত্যাদি),
#                  L/W/H = "Auto Carton" মেজারমেন্ট কলাম থেকে, Qty = Carton
#                  Qty (Pcs), Ply = Remarks কলাম থেকে (৫ PLY / ৩ PLY)
#   ২. Top      -> item_name সবসময় 'Top', শুধু L/W = Carton-Top মেজারমেন্ট
#                  কলাম থেকে (H নেই), Qty = Carton-Top Qty (Pcs), Ply
#                  সবসময় '3' (Carton-Top-এর ply Remarks কলামে থাকে না)
#
# Style No/PO No/Reference (Set Name) দুই লাইন-আইটেমেই একই থাকে।
# ---------------------------------------------------------------------------


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _cell(row, idx):
    if idx is None or idx >= len(row):
        return ''
    v = row[idx]
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    return v


def _is_blank_row(row):
    for c in row:
        if c is None:
            continue
        if isinstance(c, float) and pd.isna(c):
            continue
        if str(c).strip() == '':
            continue
        return False
    return True


def _parse_measurement_lwh(m):
    """'L-52 X W-32 X H-20 CM' -> (52, 32, 20). H না থাকলে ('L-50 X W-30 CM')
    height ফাঁকা রিটার্ন করে।"""
    if not m:
        return '', '', ''
    text = str(m).upper()

    def _num_after(letter):
        match = re.search(letter + r'\s*[-:]?\s*([0-9]+(?:\.[0-9]+)?)', text)
        if match:
            try:
                return float(match.group(1))
            except ValueError:
                return ''
        return ''

    length = _num_after('L')
    width = _num_after('W')
    height = _num_after('H')
    return length, width, height


def _parse_ply(remark):
    """'5 PLY' -> '5', '3 PLY' -> '3'। সংখ্যা না পেলে raw টেক্সট ফেরত দেয়
    (যাতে নীরবে ভুল ডাটা হারিয়ে না যায়)।"""
    if not remark:
        return 'N/A'
    match = re.search(r'([0-9]+)\s*PLY', str(remark).upper())
    if match:
        return match.group(1)
    return str(remark).strip() or 'N/A'


def _read_rows(file_stream, filename):
    is_xls = str(filename).lower().endswith('.xls')
    engines_to_try = ['xlrd', 'calamine'] if is_xls else ['openpyxl', 'calamine']
    errors = []
    for engine in engines_to_try:
        try:
            file_stream.seek(0)
            df_raw = pd.read_excel(file_stream, header=None, engine=engine)
            return df_raw.values.tolist()
        except Exception as e:
            errors.append(f"[{engine}] {type(e).__name__}: {e}")
    raise ValueError("Columbia (GU) এক্সেল পড়া যায়নি — " + " | ".join(errors))


def read_columbia_style_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """একটা Columbia Apparels (buyer: GU) 'CARTON BOOKING FOR G.U' এক্সেল
    ফাইল থেকে canonical লাইন-আইটেম লিস্ট বের করে (Carton + Top, প্রতি
    ডাটা-রো-তে দুইটা করে)।

    item_name_override: Carton-অংশের item name (UI থেকে সিলেক্ট করা,
    যেমন Master Carton) — Top-অংশে এটা প্রযোজ্য না, ওখানে সবসময় 'Top' বসে।
    manual_ply: ব্যবহার হয় না (Ply এই ফরম্যাটে Remarks কলাম থেকেই সরাসরি
    পাওয়া যায়) — শুধু বাকি extractor-দের সাথে uniform ইন্টারফেস রাখতে
    প্যারামিটারটা রাখা হয়েছে।
    """
    rows = _read_rows(file_stream, filename)

    ship_to = ''
    style_no = ''
    po_no = ''
    header_idx = None

    for i, row in enumerate(rows):
        c0 = str(_cell(row, 0)).strip()
        if _norm(c0) == _norm('Ship To :') or _norm(c0).startswith(_norm('Ship To')):
            ship_to = str(_cell(row, 1)).strip()
        elif _norm(c0) == _norm('Style:') or _norm(c0) == _norm('Style'):
            style_no = str(_cell(row, 2)).strip()
        elif _norm(c0) == _norm('PO NO'):
            po_no = str(_cell(row, 2)).strip()
        elif _norm(c0) == _norm('Set Name'):
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("'Set Name' হেডিং-সহ পরিচিত Columbia (GU) বুকিং টেবিল ফরম্যাট পাওয়া যায়নি")
    if not style_no or not po_no:
        raise ValueError("Style/PO NO হেডার তথ্য পাওয়া যায়নি — ফাইলটা পরিচিত ফরম্যাটের কিনা যাচাই করুন")

    # হেডারের ঠিক পরের রো-টা সাব-হেডার ('03 Ply Auto Carton'/'Carton Top
    # (Only Top)') — ডাটা তার পরের রো থেকে শুরু। কলাম-পজিশন ফাইল-ভেদে একই
    # (D=Auto Carton measurement, N=Carton-Top measurement, S/T=Qty, V=Remarks)
    # তবে row-position একটু নড়াচড়া করতে পারে বলে fixed offset না ধরে
    # 'Set Name' হেডারের ঠিক ২ রো পরে থেকে স্ক্যান শুরু করা হচ্ছে এবং
    # প্রথম blank/summary রো পেলেই থেমে যাওয়া হচ্ছে।
    combined_label = f"{po_no}/{ship_to}" if ship_to else po_no
    po_number_final = combined_label

    # Carton আর Top লাইন-আইটেম আলাদা লিস্টে জমা করা হচ্ছে, তারপর শেষে
    # carton_items + top_items জোড়া লাগানো হবে — যাতে ফাইনাল আউটপুটে
    # সব Carton রো আগে, তারপর সব Top রো (row-by-row ইন্টারলিভড না)।
    carton_items = []
    top_items = []
    for row in rows[header_idx + 2:]:
        if row is None or _is_blank_row(row):
            break
        set_name = str(_cell(row, 0)).strip()
        if not set_name:
            # টোটাল/সামারি রো — সেট নেম না থাকলে ডাটা রো না
            break
        set_name_2 = str(_cell(row, 1)).strip()
        reference = f"{set_name}/{set_name_2}" if set_name_2 else set_name

        auto_measure = _cell(row, 3)
        top_measure = _cell(row, 13)
        carton_qty = _cell(row, 18)
        top_qty = _cell(row, 19)
        remark = _cell(row, 21)

        length, width, height = _parse_measurement_lwh(auto_measure)
        ply = _parse_ply(remark)

        # --- Carton লাইন ---
        if auto_measure and carton_qty != '':
            carton_items.append({
                'item_name': item_name_override or 'Master Carton',
                'ewo_no': 'N/A',
                'style_no': style_no,
                'po_no': po_number_final,
                'length': length,
                'width': width,
                'height': height,
                'ply': ply,
                'qty': carton_qty,
                'pack_type': '',
                'reference': reference,
                'color': '',
                'size': '',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': ship_to,
                'delivery_address_pdf': '',
                '_source_file': filename,
            })

        # --- Top লাইন (শুধু L/W, Ply সবসময় 3) ---
        if top_measure and top_qty != '':
            top_length, top_width, _ = _parse_measurement_lwh(top_measure)
            top_items.append({
                'item_name': 'Top',
                'ewo_no': 'N/A',
                'style_no': style_no,
                'po_no': po_number_final,
                'length': top_length,
                'width': top_width,
                'height': '',
                'ply': '3',
                'qty': top_qty,
                'pack_type': '',
                'reference': reference,
                'color': '',
                'size': '',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': ship_to,
                'delivery_address_pdf': '',
                '_source_file': filename,
            })

    items = carton_items + top_items
    if not items:
        raise ValueError("হেডার পাওয়া গেছে কিন্তু কোনো ডাটা রো পাওয়া যায়নি")

    return items