import re
import os
import shutil
import subprocess
import tempfile
from datetime import datetime
import pandas as pd

from norp_extractor import read_norp_style_excel
from simba_extractor import read_simba_style_excel
from pfl_extractor import read_pfl_style_excel
from ventura_extractor import read_ventura_style_excel
from knitconcept_extractor import read_knitconcept_style_excel
from columbia_extractor import read_columbia_style_excel
from columbia_target_australia_extractor import read_columbia_target_australia_style_excel
from amigo_uniqlo_extractor import combine_amigo_booking_files
from sinha_tatatrent_extractor import combine_sinha_booking_files
from sterling_target_extractor import combine_sterling_booking_files

# ---------------------------------------------------------------------------
# আউট হাউজ Carton বুকিং এক্সেল (.xls/.xlsx) থেকে ডাটা বের করার মডিউল।
# একাধিক ফাইল একসাথে আপলোড করা হলে, প্রতিটা থেকে ঠিক একই নিয়মে ডাটা নিয়ে
# আপলোড-ক্রম অনুযায়ী (প্রথম ফাইলের ঠিক নিচে দ্বিতীয়টা) মিলিয়ে একটাই
# লিস্ট রিটার্ন করা হয়।
# ---------------------------------------------------------------------------


def _norm(s):
    """তুলনা করার সময় সব স্পেস/পাংচুয়েশন বাদ দিয়ে lowercase করে — যাতে
    'PO#' vs 'PO #' vs 'PO No' জাতীয় ছোটখাটো ভিন্নতা ম্যাচ করানো যায়।"""
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


# PDF-এর মতোই এখানেও normalize করা key দিয়ে ম্যাচ করা হয়, যাতে হেডিং-এর
# সামান্য বানান/স্পেসিং তফাত থাকলেও কাজ করে। নতুন কাস্টমারের এক্সেলে হেডিং
# ভিন্ন হলে এখানে একটা এন্ট্রি যোগ করলেই যথেষ্ট, বাকি কোড বদলাতে হবে না।
FIELD_MAP = {
    _norm('PO#'): 'po_no',
    _norm('PO No'): 'po_no',
    _norm('STYLE#'): 'style_no',
    _norm('STYLE No'): 'style_no',
    _norm('COLOR NAME/CODE#'): 'color_name',
    _norm('COLOR NAME'): 'color_name',
    _norm('GMT SIZE'): 'gmt_size',
    _norm('SIZE'): 'gmt_size',
    _norm('UNIT PER CARTON [PCS]'): 'unit_per_carton',
    _norm('SKU#'): 'sku',
    _norm('MEASUREMENT [CM]'): 'measurement',
    _norm('MEASUREMENT'): 'measurement',
    _norm('ACTUAL QTY'): 'actual_qty',
    _norm('BOOKING QTY'): 'booking_qty',
    _norm('DEL DATE'): 'del_date',
    _norm('REMARK'): 'remark',
    _norm('PLY'): 'ply',
}

_REQUIRED_HEADER_TOKENS = (_norm('PO#'), _norm('STYLE#'))


def _find_header_row(rows, max_scan=25):
    """হেডার রো সাধারণত ১০ নম্বরে থাকে, কিন্তু ফাইল-ভেদে অবস্থান একটু
    আগে-পিছে হতে পারে — তাই row-position ধরে না রেখে, প্রথম max_scan রো-র
    মধ্যে যেই রো-তে 'PO#' এবং 'STYLE#' দুটোই পাওয়া যাবে সেটাকেই হেডার ধরা হয়।"""
    for i, row in enumerate(rows[:max_scan]):
        normed = [_norm(c) for c in row]
        if all(tok in normed for tok in _REQUIRED_HEADER_TOKENS):
            return i
    return None


def _parse_measurement(m):
    """'56.7X31.3X30.5' -> (56.7, 31.3, 30.5)। L x W (হাইট ছাড়া) দেওয়া থাকলেও
    সামলাতে পারে। কোনো সংখ্যা পার্স করা না গেলে সেটা ফাঁকা রাখা হয় (পুরো
    রো বাদ দেওয়া হয় না)।"""
    if not m:
        return '', '', ''
    parts = re.split(r'[xX×]', str(m).strip())
    parts = [p.strip() for p in parts if p.strip()]
    vals = []
    for p in parts:
        try:
            vals.append(float(p))
        except ValueError:
            vals.append('')
    while len(vals) < 3:
        vals.append('')
    return vals[0], vals[1], vals[2]


def _format_date(v):
    if v is None or v == '':
        return ''
    if isinstance(v, datetime):
        return v.strftime('%d-%b-%Y')
    return str(v).strip()


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


def _try_xlrd_ignore_corruption(file_stream):
    """কিছু ERP/পোর্টাল থেকে এক্সপোর্ট করা .xls ফাইলে সামান্য অ-স্ট্যান্ডার্ড
    BIFF রেকর্ড থাকে (ফাইল-কনটেইনার হিসেবে বৈধ, কিন্তু xlrd-এর কড়া parser
    এতে AssertionError ছুঁড়ে দেয়)। xlrd-এর নিজস্ব
    ignore_workbook_corruption মোড এই ধরনের ছোটখাটো অসঙ্গতি উপেক্ষা করে
    পড়তে পারে — শেষ চেষ্টা হিসেবে এটা ব্যবহার করা হচ্ছে।

    এখানে pandas বাইপাস করে সরাসরি xlrd ব্যবহার করা হচ্ছে বলে, ডেট-টাইপ
    সেলগুলো ম্যানুয়ালি datetime-এ কনভার্ট করতে হচ্ছে (pandas সাধারণত এটা
    নিজে থেকেই করে দেয়, raw xlrd করে না — না করলে DEL DATE কলামে সংখ্যা
    (Excel serial date) দেখাবে, আসল তারিখ না)।
    """
    import xlrd
    from xlrd.xldate import xldate_as_datetime

    file_stream.seek(0)
    book = xlrd.open_workbook(file_contents=file_stream.read(), ignore_workbook_corruption=True)
    sheet = book.sheet_by_index(0)
    rows = []
    for r in range(sheet.nrows):
        row_vals = []
        for c in range(sheet.ncols):
            cell = sheet.cell(r, c)
            if cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    row_vals.append(xldate_as_datetime(cell.value, book.datemode))
                except Exception:
                    row_vals.append(cell.value)
            else:
                row_vals.append(cell.value)
        rows.append(row_vals)
    return rows


def _find_soffice():
    """সিস্টেমে LibreOffice ইনস্টল আছে কিনা খুঁজে বের করে (PATH-এ, অথবা
    Windows-এর কমন ইনস্টল লোকেশনে)। না থাকলে None রিটার্ন করে।"""
    path = shutil.which('soffice') or shutil.which('soffice.exe')
    if path:
        return path
    for candidate in (
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    ):
        if os.path.exists(candidate):
            return candidate
    return None


def _try_libreoffice_convert(file_bytes, filename):
    """শেষ চেষ্টা হিসেবে — অনেকটা ম্যানুয়ালি Excel-এ ফাইলটা খুলে
    Copy → Paste Special → Values Only করে নতুন একটা ফাইল বানানোর মতোই,
    কিন্তু পুরোপুরি অটোমেটিক। LibreOffice-এর নিজস্ব parser xlrd/calamine-এর
    চেয়ে অনেক বেশি সহনশীল (Excel-এর মতোই বিভিন্ন অসঙ্গতি সামলে নিতে পারে),
    তাই ফাইলটাকে আগে একটা পরিষ্কার .xlsx-এ কনভার্ট করে নিলে সেটা তখন
    openpyxl দিয়ে সহজেই পড়া যায়।

    এটা সম্পূর্ণ ঐচ্ছিক — সিস্টেমে LibreOffice (soffice) ইনস্টল করা না
    থাকলে এই ধাপ চুপচাপ স্কিপ হয়ে যাবে (এরর যোগ হবে, কিন্তু বাকি ফলব্যাক
    চেইন প্রভাবিত হবে না)।
    """
    soffice_path = _find_soffice()
    if not soffice_path:
        raise RuntimeError('LibreOffice (soffice) এই সিস্টেমে ইনস্টল করা পাওয়া যায়নি')

    with tempfile.TemporaryDirectory() as tmpdir:
        src_path = os.path.join(tmpdir, filename)
        with open(src_path, 'wb') as f:
            f.write(file_bytes)
        result = subprocess.run(
            [soffice_path, '--headless', '--convert-to', 'xlsx', '--outdir', tmpdir, src_path],
            capture_output=True, timeout=60,
        )
        converted_path = os.path.join(tmpdir, os.path.splitext(filename)[0] + '.xlsx')
        if not os.path.exists(converted_path):
            raise RuntimeError(f'LibreOffice কনভার্সন ব্যর্থ হয়েছে: {result.stderr.decode(errors="ignore")[:200]}')
        df_raw = pd.read_excel(converted_path, header=None, engine='openpyxl')
        return df_raw.values.tolist()


def _read_excel_rows(file_stream, filename):
    """ফাইলটা একাধিক Excel-reading ইঞ্জিন দিয়ে চেষ্টা করা হয় (নির্ভরযোগ্যতা
    বাড়ানোর জন্য) — কোনো একটা লাইব্রেরি (যেমন xlrd) ইনস্টল করা না থাকলে বা
    কোনো কারণে ব্যর্থ হলে, পরেরটা দিয়ে আবার চেষ্টা করা হয়। .xls-এর জন্য আগে
    'xlrd' (স্ট্যান্ডার্ড), তারপর 'calamine' (নতুন, দ্রুত, xls/xlsx/xlsb/ods
    সবকটাই সাপোর্ট করে), তারপর xlrd-এর corruption-tolerant মোড (শেষ চেষ্টা,
    ERP-এক্সপোর্টেড .xls-এ প্রায়ই কাজ করে) — .xlsx-এর জন্য উল্টো ক্রমে।
    সবগুলো ব্যর্থ হলে তবেই এরর দেখানো হয়, সবগুলোর মেসেজসহ।

    দুই ধরনের ব্যর্থতা আলাদাভাবে চেনা হয়:
    - লাইব্রেরিই ইনস্টল করা নেই (ImportError) — এক্ষেত্রে "লাইব্রেরি মিসিং" মার্কার থাকবে
    - লাইব্রেরি ইনস্টল আছে কিন্তু এই নির্দিষ্ট ফাইলটাই পড়া যাচ্ছে না (করাপ্টেড/অচেনা
      ফরম্যাট ইত্যাদি) — এক্ষেত্রে আসল এরর মেসেজটাই দেখানো হয়, যাতে বোঝা যায় এটা
      লাইব্রেরির সমস্যা না, এই ফাইলেরই কোনো সমস্যা।
    """
    is_xls = filename.lower().endswith('.xls')
    engines_to_try = ['xlrd', 'calamine'] if is_xls else ['openpyxl', 'calamine']
    errors = []
    all_import_errors = True
    for engine in engines_to_try:
        try:
            file_stream.seek(0)
            df_raw = pd.read_excel(file_stream, header=None, engine=engine)
            return df_raw.values.tolist()
        except ImportError as e:
            errors.append(f"[{engine}] {e}")
        except Exception as e:
            all_import_errors = False
            errors.append(f"[{engine}] {type(e).__name__}: {e}")

    if is_xls:
        try:
            return _try_xlrd_ignore_corruption(file_stream)
        except ImportError as e:
            errors.append(f"[xlrd-corruption-tolerant] {e}")
        except Exception as e:
            all_import_errors = False
            errors.append(f"[xlrd-corruption-tolerant] {type(e).__name__}: {e}")

    # সবশেষ চেষ্টা: LibreOffice দিয়ে পরিষ্কার .xlsx-এ কনভার্ট করে পড়া (values-only
    # কপি-পেস্টের মতোই, কিন্তু অটোমেটিক) — LibreOffice ইনস্টল না থাকলে স্কিপ হয়ে যাবে
    try:
        file_stream.seek(0)
        return _try_libreoffice_convert(file_stream.read(), filename)
    except Exception as e:
        errors.append(f"[libreoffice] {type(e).__name__}: {e}")

    if all_import_errors:
        raise ValueError("(লাইব্রেরি মিসিং) কোনো Excel engine দিয়েই পড়া যায়নি — " + " | ".join(errors))
    raise ValueError(
        "এই নির্দিষ্ট ফাইলটা পড়া যায়নি (লাইব্রেরির সমস্যা না, ফাইল-নির্দিষ্ট সমস্যা) — "
        + " | ".join(errors)
    )


def read_booking_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """একটা বায়িং-হাউজ বুকিং এক্সেল ফাইল (.xls/.xlsx) থেকে canonical
    লাইন-আইটেম লিস্ট বের করে। ম্যাপিং (ব্যবহারকারীর নির্দেশ অনুযায়ী):
      PO#              -> Gmt. PO
      STYLE#           -> Gmt. Style No
      COLOR NAME/CODE# -> Reference/SKU Number
      GMT SIZE         -> Pack Type
      MEASUREMENT [CM] -> Length/Width/Height (L x W x H পার্স করে)
      BOOKING QTY      -> Order Qty (ACTUAL QTY ব্যবহার হয় না, স্পষ্ট নির্দেশ অনুযায়ী)
    EWO No এই ফরম্যাটে কখনোই থাকে না — সরাসরি 'N/A'।

    item_name_override: এই এক্সেল ফরম্যাটে Item Name কলাম থাকে না, তাই UI
    থেকে ইউজার যেটা সিলেক্ট করেছেন (Master Carton/Elastic Hanger Carton/
    Both Side Hanger Carton) সেটাই সব রো-তে বসবে।

    manual_ply: Ply কখনো কখনো এক্সেলে (PLY কলামে) থাকে, কখনো থাকে না।
    ফাইলে PLY কলাম পাওয়া গেলে সেটাই ব্যবহার হবে (row-by-row); না পাওয়া
    গেলে UI থেকে ম্যানুয়ালি সিলেক্ট করা মান (বা ফাঁকা থাকলে 'N/A') বসবে।
    """
    rows = _read_excel_rows(file_stream, filename)

    header_idx = _find_header_row(rows)
    if header_idx is None:
        raise ValueError("'PO#'/'STYLE#' হেডিং-সহ পরিচিত বুকিং টেবিল ফরম্যাট পাওয়া যায়নি")

    header_row = rows[header_idx]
    col_map = {}
    for ci, cell in enumerate(header_row):
        key = FIELD_MAP.get(_norm(cell))
        if key and key not in col_map:
            col_map[key] = ci

    has_ply_column = 'ply' in col_map

    def get(row, field):
        ci = col_map.get(field)
        if ci is None or ci >= len(row):
            return ''
        v = row[ci]
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ''
        return v

    items = []
    for row in rows[header_idx + 1:]:
        if row is None or _is_blank_row(row):
            continue
        po_no = get(row, 'po_no')
        po_no_str = str(po_no).strip()
        if po_no_str == '':
            continue  # সামারি/ফাঁকা রো
        # টেবিলের নিচে প্রায়ই নোট/ফুটার টেক্সট থাকে (যেমন 'Ship To #...',
        # 'Gross Carton Weight...') — সেই কলামেও কিছু টেক্সট পড়ে যেতে পারে,
        # তাই PO# আসলেই সংখ্যা কিনা যাচাই করে সেগুলো বাদ দেওয়া হচ্ছে।
        try:
            float(po_no_str.replace(',', ''))
        except ValueError:
            continue

        if has_ply_column:
            row_ply = str(get(row, 'ply')).strip() or 'N/A'
        else:
            row_ply = manual_ply.strip() if manual_ply else 'N/A'

        length, width, height = _parse_measurement(get(row, 'measurement'))
        items.append({
            'item_name': item_name_override or 'Master Carton',
            # এই এক্সেল ফরম্যাটে EWO No কখনোই থাকে না — তাই ফাঁকা রেখে প্রতিটা
            # রো-তে Warning তৈরি করার বদলে সরাসরি 'N/A' বসানো হচ্ছে (এটা
            # জেনুইনভাবে প্রযোজ্য না, মিসিং ডাটা না)।
            'ewo_no': 'N/A',
            'style_no': str(get(row, 'style_no')).strip(),
            'po_no': str(po_no).strip(),
            'length': length,
            'width': width,
            'height': height,
            'ply': row_ply,
            'qty': get(row, 'booking_qty'),
            'pack_type': str(get(row, 'gmt_size')).strip(),
            'reference': str(get(row, 'color_name')).strip(),
            'color': '',
            'size': '',
            'delivery_date': _format_date(get(row, 'del_date')),
            'measurement_unit': 'Cm',
            'delivery_place_pdf': '',
            'delivery_address_pdf': '',
            '_source_file': filename,
        })

    if not items:
        raise ValueError("হেডার পাওয়া গেছে কিন্তু কোনো ডাটা রো পাওয়া যায়নি")

    return items


def _norm_key(s):
    """Customer/Buyer নাম মেলানোর জন্য normalize — স্পেস/পাংচুয়েশন/কেস
    বাদ দিয়ে, যাতে সামান্য বানান-ভিন্নতাতেও সঠিক এন্ট্রি মেলে।"""
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


# --- প্রতিটা extractor-কে একটা কমন ইন্টারফেসে wrap করা হয়েছে
# (file_stream, filename, item_name_override, manual_ply, buyer_name) ->
# items — যাতে REGISTRY-তে সব এন্ট্রি একইভাবে কল করা যায়, extractor-ভেদে
# আলাদা প্যারামিটার মনে রাখতে না হয়। মূল extractor ফাইলগুলোর কোনো কোড
# বদলানো হয়নি, শুধু এখানে uniform ভাবে কল করা হচ্ছে।

def _wrap_norp(file_stream, filename, item_name_override, manual_ply, buyer_name):
    return read_norp_style_excel(file_stream, filename)


def _wrap_simba(file_stream, filename, item_name_override, manual_ply, buyer_name):
    return read_simba_style_excel(file_stream, filename, manual_ply=manual_ply)


def _wrap_pfl(file_stream, filename, item_name_override, manual_ply, buyer_name):
    return read_pfl_style_excel(file_stream, filename)


def _wrap_ventura(file_stream, filename, item_name_override, manual_ply, buyer_name):
    return read_ventura_style_excel(
        file_stream, filename, buyer_key=buyer_name,
        item_name_override=item_name_override, manual_ply=manual_ply)


def _wrap_knitconcept(file_stream, filename, item_name_override, manual_ply, buyer_name):
    return read_knitconcept_style_excel(file_stream, filename, manual_ply=manual_ply)


def _wrap_columbia(file_stream, filename, item_name_override, manual_ply, buyer_name):
    return read_columbia_style_excel(
        file_stream, filename,
        item_name_override=item_name_override, manual_ply=manual_ply)


def _wrap_columbia_target_australia(file_stream, filename, item_name_override, manual_ply, buyer_name):
    return read_columbia_target_australia_style_excel(
        file_stream, filename,
        item_name_override=item_name_override, manual_ply=manual_ply)


def _wrap_aeo(file_stream, filename, item_name_override, manual_ply, buyer_name):
    return read_booking_excel(
        file_stream, filename,
        item_name_override=item_name_override, manual_ply=manual_ply)


# ---------------------------------------------------------------------------
# CUSTOMER + BUYER -> extractor রেজিস্ট্রি। এখানে key হলো
# (normalize(customer_name), normalize(buyer_name))। বায়ার-নির্বিশেষে একই
# ফরম্যাট হলে '*' (wildcard) ব্যবহার করা হয়েছে।
#
# ১০০% বায়ার-ওয়াইজ স্ট্রিক্ট: প্রতিটা (customer, buyer) কম্বিনেশনের জন্য
# ঠিক একটাই extractor বাঁধা থাকে, কোনো ফলব্যাক-চেইন/ট্রাই নেই — ঠিক
# Ventura-এর মতো। UI-তে যে buyer সিলেক্ট করা হবে, শুধু সেই buyer-এর
# নির্দিষ্ট extractor-টাই চলবে; অন্য কোনো buyer-এর ফরম্যাট কখনো ট্রাই
# হবে না, তাই সময়ও কম লাগে আর কনফ্লিক্টের কোনো ঝুঁকিও থাকে না।
#
# PFL আর Norp একই গ্রুপের দুই কাস্টমার (Prudent Fashion Ltd. / Norp Knit
# Industries Ltd.), দুজনেরই বায়ার লিস্ট same (Macy, Kohl`s) — কিন্তু
# বায়ার অনুযায়ী ফরম্যাট আলাদা: Macy -> norp-স্টাইল এক্সেল, Kohl`s ->
# pfl-স্টাইল এক্সেল। তাই দুই কাস্টমারের ক্ষেত্রেই একই বায়ার-ওয়াইজ ম্যাপিং।
#
# Columbia Apparels Limited-এর ক্ষেত্রেও একই কারণে দুইটা আলাদা এন্ট্রি:
# GU -> columbia-স্টাইল এক্সেল, Target Australia -> সম্পূর্ণ ভিন্ন লেআউটের
# columbia_target_australia-স্টাইল এক্সেল।
#
# নতুন কাস্টমার/বায়ার যোগ করতে হলে এখানে শুধু একটা লাইন যোগ করলেই হবে —
# বাকি কোনো কোড বদলানোর দরকার নেই।
#
# ব্যতিক্রম: Amigo Bangladesh Ltd (Uniqlo), Sinha Knit and Denims Limited
# (Tata Trent) আর Sterling Styles Limited (Target) এই REGISTRY-তে নেই —
# এরা BATCH_REGISTRY-তে আলাদাভাবে হ্যান্ডল হয় (নিচে দেখুন), কারণ এদের
# এক্সট্র্যাক্টর প্রতি-ফাইল wrapper প্যাটার্নে চলে না (একাধিক ফাইল জুড়ে
# ক্রস-ফাইল অর্ডারিং লজিক লাগে)।
# ---------------------------------------------------------------------------
REGISTRY = {
    (_norm_key('Simba Fashions Limited'), '*'): [_wrap_simba],
    (_norm_key('PRUDENT FASHION LTD.'), _norm_key('Macy')): [_wrap_norp],
    (_norm_key('PRUDENT FASHION LTD.'), _norm_key("Kohl`s")): [_wrap_pfl],
    (_norm_key('Norp Knit Industries Ltd.'), _norm_key('Macy')): [_wrap_norp],
    (_norm_key('Norp Knit Industries Ltd.'), _norm_key("Kohl`s")): [_wrap_pfl],
    (_norm_key('Ventura (HK) Trading Limited'), _norm_key('Kate Spade')): [_wrap_ventura],
    (_norm_key('Ventura (HK) Trading Limited'), _norm_key('Michael Kors')): [_wrap_ventura],
    (_norm_key('Ventura (HK) Trading Limited'), _norm_key('Coach')): [_wrap_ventura],
    (_norm_key('Ventura (HK) Trading Limited'), _norm_key('Le Sportsac')): [_wrap_ventura],
    (_norm_key('Ventura (HK) Trading Limited'), _norm_key('Vera Bradley')): [_wrap_ventura],
    (_norm_key('Knit Concept LTD.'), '*'): [_wrap_knitconcept],
    (_norm_key('Columbia Apparels Limited'), _norm_key('GU')): [_wrap_columbia],
    (_norm_key('Columbia Apparels Limited'), _norm_key('Target Australia')): [_wrap_columbia_target_australia],
}


# ---------------------------------------------------------------------------
# BATCH_REGISTRY — REGISTRY-এর মতোই (customer, buyer) key দিয়ে buyer-ওয়াইজ
# স্ট্রিক্ট লুকআপ, কিন্তু এখানে value একটা wrapper না — সরাসরি একটা ফাংশন
# যেটা সবগুলো আপলোড করা ফাইল (files: [(BytesIO, filename), ...]) একসাথে
# নিয়ে (line_items, warnings) রিটার্ন করে। এটা তখনই দরকার হয় যখন কোনো
# ফরম্যাটে একাধিক ফাইলের মধ্যে ক্রস-ফাইল অর্ডারিং/লজিক জরুরি হয় — যেমন:
#   - Amigo Bangladesh Ltd (Uniqlo): সব ফাইলের Size Breakdown (Master
#     Carton) লাইন আগে বসবে, তারপর সব ফাইলের Indent (Top Bottom/Divider)
#     লাইন সবার শেষে।
#   - Sinha Knit and Denims Limited (Tata Trent): একই নিয়মে, সব শিট/ফাইলের
#     মেইন বুকিং ডাটা আগে, Top Bottom/Divider ট্রেইলার লাইন সবার শেষে।
#   - Sterling Styles Limited (Target): একই নিয়মে, সব ফাইলের Master
#     Carton আগে, Divider/Top Bottom (measurement-ওয়াইজ গ্রুপড-সামারি)
#     সবার শেষে — Item Name/Ply এখানেও সম্পূর্ণ ফাইলের কনটেন্ট থেকেই
#     ডিটেক্ট হয় (ELASTIC নোট + হেডারের 5/3 Ply উল্লেখ), UI সিলেকশন
#     প্রযোজ্য না।
# প্রতিটা ফাইল আলাদাভাবে প্রসেস করে পরে জোড়া লাগালে এই অর্ডারিং ঠিক রাখা
# যায় না, তাই এই ফাংশনগুলোকে সবগুলো ফাইল একসাথেই দেওয়া হয়।
#
# uniform কল-সিগনেচার: (files, item_name_override, manual_ply) -> এই তিনটা
# আর্গুমেন্ট দিয়েই সব batch ফাংশন কল হয়, নিচের ছোট wrapper-গুলো যার যার
# আসল ফাংশনের সাথে মিলিয়ে নেয় (Amigo-র নিজস্ব ফাংশন item_name/ply নেয় না,
# তাই ওই wrapper-এ সেগুলো উপেক্ষা করা হয়)।
# ---------------------------------------------------------------------------
def _batch_amigo(files, item_name_override='', manual_ply=''):
    # Amigo-র নিজস্ব ফরম্যাটে Item Name/Ply সবসময় নির্দিষ্ট (Master Carton
    # / Top Bottom / Divider, extractor নিজেই ঠিক করে) — UI সিলেকশন এখানে
    # প্রযোজ্য না, তাই ইচ্ছাকৃতভাবে উপেক্ষা করা হচ্ছে।
    return combine_amigo_booking_files(files)


def _batch_sinha(files, item_name_override='', manual_ply=''):
    return combine_sinha_booking_files(files, item_name_override=item_name_override, manual_ply=manual_ply)


def _batch_sterling(files, item_name_override='', manual_ply=''):
    # Sterling-এর নিজস্ব ফরম্যাটে Item Name (ELASTIC নোট) আর Ply (5/3,
    # হেডারেই লেখা) সম্পূর্ণ ফাইলের কনটেন্ট থেকে ডিটেক্ট হয় — UI সিলেকশন
    # এখানে প্রযোজ্য না, তাই ইচ্ছাকৃতভাবে উপেক্ষা করা হচ্ছে।
    return combine_sterling_booking_files(files)


BATCH_REGISTRY = {
    (_norm_key('Amigo Bangladesh Ltd'), _norm_key('Uniqlo')): _batch_amigo,
    (_norm_key('Sinha Knit and Denims Limited'), _norm_key('Tata Trent')): _batch_sinha,
    (_norm_key('Sterling Styles Limited'), _norm_key('Target')): _batch_sterling,
}


def _get_extractor_chain(customer_name, buyer_name):
    """Customer+Buyer অনুযায়ী সঠিক extractor-চেইন বের করে। REGISTRY-তে
    নির্দিষ্ট (customer, buyer) এন্ট্রি না থাকলে ওই কাস্টমারের '*'
    (যেকোনো বায়ার) এন্ট্রি চেক করা হয়; সেটাও না থাকলে সবার শেষে AEO-স্টাইল
    জেনেরিক ফলব্যাক ব্যবহার হয় (নতুন/এখনো-রেজিস্টার-না-করা কাস্টমারের জন্য)।"""
    c = _norm_key(customer_name)
    b = _norm_key(buyer_name)
    if (c, b) in REGISTRY:
        return REGISTRY[(c, b)]
    if (c, '*') in REGISTRY:
        return REGISTRY[(c, '*')]
    return [_wrap_aeo]


def combine_booking_excels(files, item_name_override='Master Carton', manual_ply='',
                            buyer_name='', customer_name=''):
    """files: [(file_stream, filename), ...] — আপলোড হওয়া ক্রম অনুযায়ী।
    প্রতিটা ফাইল থেকে ডাটা নিয়ে ক্রমানুসারে (প্রথম ফাইলের ঠিক নিচেই পরের
    ফাইলের ডাটা) একটাই কম্বাইনড লিস্টে জোড়া লাগিয়ে দেয়। কোনো একটা ফাইলে
    সমস্যা হলে সেটা স্কিপ হয়ে যায় (বাকিগুলো প্রসেস চলতে থাকে), আর সেই
    এরর মেসেজ আলাদাভাবে রিটার্ন হয় যাতে ইউজারকে জানানো যায়।

    ফরম্যাট বেছে নেওয়া হয় customer_name + buyer_name দিয়ে সরাসরি REGISTRY
    লুকআপ করে (দেখুন REGISTRY ডিক্ট) — আগের মতো সব ফরম্যাট একটার পর একটা
    "গেস" করে try করা হয় না। এতে দুই কাস্টমারের ফরম্যাটের মধ্যে ভুলবশত
    কনফ্লিক্ট/ওভারল্যাপ হওয়ার কোনো সুযোগ থাকে না, আর নতুন কাস্টমার/বায়ার
    যোগ করাও অনেক নিরাপদ (শুধু REGISTRY-তে একটা এন্ট্রি যোগ করলেই হয়)।

    আগে BATCH_REGISTRY (দেখুন উপরের কমেন্ট) চেক করা হয় — ম্যাচ পেলে সেই
    ফাংশনটাই সরাসরি সবগুলো ফাইল দিয়ে কল হয়ে যায় (প্রতি-ফাইল লুপ এড়িয়ে)।
    এখানেও buyer-ওয়াইজ স্ট্রিক্ট নিয়ম বহাল থাকে — নির্দিষ্ট customer+buyer
    ম্যাচ না করলে batch পথে যাবে না।

    Returns (combined_line_items, file_errors).
    """
    c = _norm_key(customer_name)
    b = _norm_key(buyer_name)
    if (c, b) in BATCH_REGISTRY:
        return BATCH_REGISTRY[(c, b)](files, item_name_override, manual_ply)

    chain = _get_extractor_chain(customer_name, buyer_name)
    combined = []
    errors = []
    for file_stream, filename in files:
        items = None
        last_error = None
        for wrapper in chain:
            try:
                file_stream.seek(0)
                result = wrapper(file_stream, filename, item_name_override, manual_ply, buyer_name)
                if result:
                    items = result
                    break
            except Exception as e:
                last_error = e
        if items:
            combined.extend(items)
        else:
            msg = str(last_error) if last_error else "এই কাস্টমার/বায়ারের জন্য পরিচিত কোনো ফরম্যাট মেলেনি"
            errors.append(f"{filename}: {msg}")
    return combined, errors