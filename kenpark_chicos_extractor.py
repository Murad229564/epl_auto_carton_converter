import os
import re
import shutil
import subprocess
import tempfile
import pandas as pd

# ---------------------------------------------------------------------------
# Kenpark Bangladesh (Pvt.) Limited — Buyer: Chicos
# 'SALE ORDER — ITEM: Carton' এক্সেল (.xls/.xlsx) ফরম্যাট ('Carton_Booking_
# Loose-...')। একসাথে একাধিক ফাইল (এবং প্রতি ফাইলে একাধিক শিট) আসতে পারে —
# সবগুলো মিলে একটাই টেমপ্লেট হয় (zip/আলাদা-টেমপ্লেট এখানে প্রযোজ্য না)।
#
# টেবিল (হেডার: Dimension LXWXH INCH / STYLE / SO No / SHIP TO / DIVISION /
# PO# / SKU / Pre-Packed / DEPARTMENT / COLOR / Sizes / Size Ratio /
# CARTON ORDER Q'TY (PCS)):
#   Dimension ('L 23 W 16 H 12.5 INCH') -> Length / Width / Height, unit = Inch
#                                           (হেডারে CM লেখা থাকলে Cm)
#   STYLE                -> Gmt. Style No (রো-ভেদে আলাদা হতে পারে, যেমন
#                           '630-27W44234' ও '030-27W44234')
#   PO#                  -> Gmt. PO
#   SO No                -> Reference/SKU Number
#   SKU                  -> Pack Type
#   CARTON ORDER Q'TY    -> Order Qty
#   Item Name ও Ply      -> UI থেকে (UI-তে Ply না দিলে 'N/A')
# অন্য কলাম (SHIP TO, DIVISION, Pre-Packed, DEPARTMENT, COLOR, Sizes, Size
# Ratio) ব্যবহার হয় না।
#
# একই শিটে হেডার-রো বারবার আসতে পারে (প্রতি PO-ব্লকের আগে) — সেগুলো ও ফাঁকা
# রো বাদ যায়, কোয়ান্টিটি-ফাঁকা রোও বাদ। 'TOTAL' রো পেলে ওই শিটের টেবিল শেষ,
# আর সেই যোগফলের সাথে না মিললে Warning। ক্রম: ফাইল-ক্রম, শিট-ক্রম, রো-ক্রম।
# ---------------------------------------------------------------------------

NA = 'N/A'


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    if isinstance(v, float) and pd.isna(v):
        return ''
    return re.sub(r'\s+', ' ', str(v).replace('\xa0', ' ')).strip()


def _clean_id(v):
    """সংখ্যা-মানের float হলে '.0' বাদ (4502279186.0 -> 4502279186)।"""
    if isinstance(v, float) and not pd.isna(v) and v == int(v):
        return str(int(v))
    return _clean(v)


def _num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    try:
        return float(str(v).replace(',', '').strip())
    except (TypeError, ValueError):
        return None


def _fmt(f):
    return int(f) if f == int(f) else f


_DIM_LWH = re.compile(
    r'L\s*(\d+(?:\.\d+)?)\s*W\s*(\d+(?:\.\d+)?)(?:\s*H\s*(\d+(?:\.\d+)?))?', re.I)


def _parse_dimension(text):
    """'L 23 W 16 H 12.5 INCH' -> (23, 16, 12.5); H না থাকলে হাইট ফাঁকা।
    'L W H' লেবেল না থাকলে '23x16x12.5' ধরনের সংখ্যা-ক্রম।"""
    t = str(text or '')
    m = _DIM_LWH.search(t)
    if m:
        h = _fmt(float(m.group(3))) if m.group(3) else ''
        return _fmt(float(m.group(1))), _fmt(float(m.group(2))), h
    nums = re.findall(r'\d+(?:\.\d+)?', t)
    vals = [_fmt(float(n)) for n in nums[:3]]
    while len(vals) < 3:
        vals.append('')
    return vals[0], vals[1], vals[2]


# ---------------------------------------------------------------------------
# মাল্টি-শিট রিডার: [(sheet_name, rows), ...]
# ---------------------------------------------------------------------------
def _sheets_from_openpyxl(path_or_stream):
    from openpyxl import load_workbook
    wb = load_workbook(path_or_stream, data_only=True)
    out = []
    for ws in wb.worksheets:
        if ws.sheet_state != 'visible':
            continue
        out.append((ws.title, [list(r) for r in ws.iter_rows(values_only=True)]))
    return out


def _sheets_from_pandas(file_stream, engine):
    file_stream.seek(0)
    d = pd.read_excel(file_stream, header=None, sheet_name=None, engine=engine)
    return [(name, df.values.tolist()) for name, df in d.items()]


def _find_soffice():
    path = shutil.which('soffice') or shutil.which('soffice.exe')
    if path:
        return path
    for cand in (r'C:\Program Files\LibreOffice\program\soffice.exe',
                 r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'):
        if os.path.exists(cand):
            return cand
    return None


def _sheets_from_libreoffice(file_bytes, filename):
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError('LibreOffice (soffice) ইনস্টল করা পাওয়া যায়নি')
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, os.path.basename(filename) or 'in.xls')
        with open(src, 'wb') as f:
            f.write(file_bytes)
        subprocess.run([soffice, '--headless', '--convert-to', 'xlsx', '--outdir', tmp, src],
                       capture_output=True, timeout=90)
        conv = os.path.join(tmp, os.path.splitext(os.path.basename(src))[0] + '.xlsx')
        if not os.path.exists(conv):
            raise RuntimeError('LibreOffice কনভার্সন ব্যর্থ')
        return _sheets_from_openpyxl(conv)


def _load_sheets(file_stream, filename):
    is_xls = (filename or '').lower().endswith('.xls')
    errors = []
    attempts = ([lambda: _sheets_from_pandas(file_stream, 'xlrd'),
                 lambda: _sheets_from_pandas(file_stream, 'calamine')]
                if is_xls else
                [lambda: (file_stream.seek(0), _sheets_from_openpyxl(file_stream))[1],
                 lambda: _sheets_from_pandas(file_stream, 'calamine')])
    for fn in attempts:
        try:
            return fn()
        except Exception as e:
            errors.append(f"{type(e).__name__}: {e}")
    try:
        file_stream.seek(0)
        return _sheets_from_libreoffice(file_stream.read(), filename)
    except Exception as e:
        errors.append(f"libreoffice {type(e).__name__}: {e}")
    raise ValueError("ফাইলটা পড়া যায়নি — " + " | ".join(errors))


# ---------------------------------------------------------------------------
def _header_map(row):
    """হেডার-রো হলে কলাম-ম্যাপ রিটার্ন, নাহলে None। হেডার-রো = STYLE + SO No +
    'CARTON ORDER Q'TY' তিনটাই আছে।"""
    normed = [_norm(c) for c in row]
    if 'style' not in normed or 'sono' not in normed:
        return None
    qty = next((i for i, n in enumerate(normed) if n.startswith('cartonorder')), None)
    if qty is None:
        return None
    dim = next((i for i, n in enumerate(normed) if n.startswith('dimension')), None)
    cols = {n: i for i, n in reversed(list(enumerate(normed))) if n}
    unit_text = _clean(row[dim]) if dim is not None else ''
    unit = 'Cm' if re.search(r'\bcm\b', unit_text, re.I) else 'Inch'
    return {
        'dim': dim,
        'style': cols['style'],
        'so': cols['sono'],
        'po': cols.get('po'),
        'sku': cols.get('sku'),
        'qty': qty,
        'unit': unit,
    }


def read_kenpark_chicos_style_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """এক ফাইল -> (items, warnings)। কোনো শিটেই টেবিল না পেলে ValueError।"""
    sheets = _load_sheets(file_stream, filename)
    item_name = item_name_override or 'Master Carton'
    ply = manual_ply.strip() if manual_ply else NA

    items, warnings = [], []
    found_any = False
    for sheet_name, rows in sheets:
        h = None
        sheet_items = []
        total_expected = None
        for row in rows:
            if row is None:
                continue
            new_h = _header_map(row)
            if new_h:                 # প্রথম বা পুনরাবৃত্ত হেডার-রো
                h = new_h
                found_any = True
                continue
            if h is None:
                continue

            first = _norm(row[0]) if row else ''
            if first == 'total':
                total_expected = _num(row[h['qty']]) if h['qty'] < len(row) else None
                break

            def get(ci):
                return row[ci] if ci is not None and ci < len(row) else None

            qty = _num(get(h['qty']))
            if qty is None or qty <= 0:
                continue

            L, W, H = _parse_dimension(_clean(get(h['dim'])))
            if L == '' or W == '':
                warnings.append(f"{filename} [{sheet_name}]: মেজারমেন্ট পাওয়া যায়নি ('{_clean(get(h['dim']))}')")
            sheet_items.append({
                'item_name': item_name,
                'ewo_no': NA,
                'style_no': _clean_id(get(h['style'])),
                'po_no': _clean_id(get(h['po'])),
                'length': L,
                'width': W,
                'height': H,
                'ply': ply,
                'qty': round(qty),
                'pack_type': _clean_id(get(h['sku'])),
                'reference': _clean_id(get(h['so'])),
                'color': '',
                'size': '',
                'delivery_date': '',
                'measurement_unit': h['unit'],
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': sheet_name,
                '_source_file': filename,
            })

        if total_expected is not None:
            got = sum(i['qty'] for i in sheet_items)
            if got != round(total_expected):
                warnings.append(f"{filename} [{sheet_name}]: Carton Qty-র যোগফল ({got}) ফাইলের TOTAL ({round(total_expected)})-এর সাথে মেলেনি")
        items.extend(sheet_items)

    if not found_any:
        raise ValueError("'STYLE / SO No / CARTON ORDER Q'TY' হেডিং-সহ Chicos টেবিল পাওয়া যায়নি")
    if not items:
        raise ValueError("হেডার পাওয়া গেছে কিন্তু কোয়ান্টিটি-সহ কোনো ডাটা রো পাওয়া যায়নি")
    return items, warnings


def combine_kenpark_chicos_booking_files(files, item_name_override='', manual_ply=''):
    """files: [(file_stream, filename), ...] -> (line_items, warnings)
    সব ফাইল ক্রমানুসারে একটাই লিস্টে। ব্যর্থ ফাইল স্কিপ হয়ে মেসেজ warnings-এ যায়।"""
    combined, warnings = [], []
    for file_stream, filename in files:
        try:
            file_stream.seek(0)
            its, w = read_kenpark_chicos_style_excel(
                file_stream, filename,
                item_name_override=item_name_override, manual_ply=manual_ply)
            combined.extend(its)
            warnings.extend(w)
        except Exception as e:
            warnings.append(f"{filename}: {e}")
    return combined, warnings
