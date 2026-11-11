import os
import re
import shutil
import subprocess
import tempfile
import pandas as pd

# ---------------------------------------------------------------------------
# Bangladesh Textile & Chemical Fiber Industry Ltd. — Buyer: Aeon
# 'WORK ORDER FOR CARTON' এক্সেল (.xls/.xlsx)। এক ফাইলে অনেকগুলো শিট থাকে
# (Japan Solid, Assort-1 ... HK) — প্রতিটা শিটে একটাই টেবিল, নিজের কার্টন
# মেজারমেন্ট/কোয়ান্টিটি সহ। সব শিট পড়া হয় (hidden শিট বাদ, শুধু .xlsx-এ
# ধরা যায়)।
#
# টেবিল (হেডার: Style & PO / Sizes in CM LxWxH / 5 ply Master Carton qty /
# CUSTOMER PO NUMBER / Top-Bottom in CM Measurement (LxW) / 3 Ply Top-Bottom
# Qty / REMARKS):
#   Style & PO            -> Gmt. Style No এবং Gmt. PO (দুটোতেই, কলামের লেখা
#                            হুবহু) এবং টেমপ্লেটের উপরের PO ফিল্ড (_header_po)
#   শিটের নাম             -> Reference/SKU (কার্টন ও Top Bottom দুই লাইনেই)
#   Sizes in CM LxWxH     -> Master Carton-এর L x W x H
#   5 ply Master Carton qty -> Master Carton Order Qty
#   Top-Bottom (LxW)      -> Top Bottom-এর L x W (হাইট নেই)
#   3 Ply Top-Bottom Qty  -> Top Bottom Order Qty
#   CUSTOMER PO NUMBER ('N/A'), REMARKS -> ব্যবহার হয় না
# Item Name ও Ply (কার্টন লাইনের) আসে UI থেকে; UI-তে Ply না দিলে হেডারের
# '5 ply' থেকে, তাও না থাকলে 'N/A'। Top Bottom সবসময় Ply 3, Item Name
# 'Top Bottom'। 'Total Carton=' রো পেলে টেবিল শেষ।
# ক্রম: সব শিটের সব কার্টন আগে, তারপর সব Top Bottom (PAIR_TOP_BOTTOM = True
# করলে প্রতি শিটের কার্টনের ঠিক নিচে তার Top Bottom)।
# ---------------------------------------------------------------------------

NA = 'N/A'

# True করলে 'Style: ' অংশ কেটে শুধু স্টাইল কোড (TV2796A1037R) যাবে।
# ডিফল্ট False = কলামের লেখা যেমন আছে ('Style: TV2796A1037R')।
STRIP_STYLE_PREFIX = True

# False (ডিফল্ট): সব কার্টন আগে, সব Top Bottom শেষে।
# True: প্রতি শিটের কার্টনের পরেই সেই শিটের Top Bottom।
PAIR_TOP_BOTTOM = True


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    if isinstance(v, float) and pd.isna(v):
        return ''
    return re.sub(r'\s+', ' ', str(v).replace('\xa0', ' ')).strip()


def _num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    try:
        return float(str(v).replace(',', '').strip())
    except (TypeError, ValueError):
        return None


def _parse_measurement(m):
    """'56*38*37' -> (56.0, 38.0, 37.0); '53*35' -> (53.0, 35.0, '')"""
    nums = re.findall(r'\d+(?:\.\d+)?', str(m or ''))
    vals = [float(n) for n in nums[:3]]
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
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        out.append((ws.title, rows))
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
def _find_header(rows, max_scan=80):
    for i, row in enumerate(rows[:max_scan]):
        normed = [_norm(c) for c in row]
        if 'stylepo' in normed and any('mastercartonqty' in n for n in normed):
            heads = {}
            for ci, n in enumerate(normed):
                if n and n not in heads:
                    heads[n] = ci

            def find(sub):
                for n, ci in heads.items():
                    if sub in n:
                        return ci
                return None

            return i, {
                'style': heads['stylepo'],
                'meas': find('sizesincm'),
                'qty': find('mastercartonqty'),
                'tb_meas': find('topbottomincm'),
                'tb_qty': find('topbottomqty'),
                'header_text': ' '.join(n for n in normed if n),
            }
    return None, None


def read_bangladesh_textile_aeon_style_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """এক ফাইল -> (cartons, top_bottoms, warnings); প্রতিটা তালিকা শিট-ক্রমে।
    কোনো টেবিল-শিটই না পেলে ValueError।"""
    sheets = _load_sheets(file_stream, filename)
    default_item = item_name_override or 'Master Carton'
    ui_ply = manual_ply.strip() if manual_ply else ''

    sheet_groups = []   # [(cartons, top_bottoms)] প্রতি শিট
    warnings = []
    for sheet_name, rows in sheets:
        header_idx, h = _find_header(rows)
        if header_idx is None:
            warnings.append(f"{filename} [{sheet_name}]: 'Style & PO' টেবিল পাওয়া যায়নি — শিটটা স্কিপ হয়েছে")
            continue

        m = re.search(r'(\d+)ply', h['header_text'])
        carton_ply = ui_ply or (m.group(1) if m else NA)
        reference = _clean(sheet_name)

        def get(row, ci):
            if ci is None or ci >= len(row):
                return ''
            return row[ci]

        cartons, tbs = [], []
        last_style = ''
        sum_c = sum_t = 0
        total_c = total_t = None
        for row in rows[header_idx + 1:]:
            if row is None:
                continue
            if any(_norm(c).startswith('totalcarton') for c in row):
                total_c = _num(get(row, h['qty']))
                total_t = _num(get(row, h['tb_qty']))
                break

            st = _clean(get(row, h['style']))
            if st:
                last_style = st
            style_val = last_style
            if STRIP_STYLE_PREFIX:
                style_val = re.sub(r'^\s*style\s*[:\-]?\s*', '', style_val, flags=re.I)

            base = {
                'ewo_no': NA, 'style_no': style_val, 'po_no': style_val,
                'pack_type': '', 'reference': reference, 'color': '', 'size': '',
                'delivery_date': '', 'measurement_unit': 'Cm',
                'delivery_place_pdf': '', 'delivery_address_pdf': '',
                '_sheet': sheet_name, '_source_file': filename,
            }

            qty = _num(get(row, h['qty']))
            if qty is not None and qty > 0:
                L, W, H = _parse_measurement(get(row, h['meas']))
                if L == '' or W == '':
                    warnings.append(f"{filename} [{sheet_name}]: Master Carton মেজারমেন্ট পাওয়া যায়নি")
                cartons.append({**base, 'item_name': default_item, 'length': L, 'width': W,
                                'height': H, 'ply': carton_ply, 'qty': round(qty)})
                sum_c += round(qty)

            tq = _num(get(row, h['tb_qty']))
            if tq is not None and tq > 0:
                L, W, _x = _parse_measurement(get(row, h['tb_meas']))
                if L == '' or W == '':
                    warnings.append(f"{filename} [{sheet_name}]: Top Bottom মেজারমেন্ট পাওয়া যায়নি")
                tbs.append({**base, 'item_name': 'Top Bottom', 'length': L, 'width': W,
                            'height': '', 'ply': '3', 'qty': round(tq)})
                sum_t += round(tq)

        if not cartons and not tbs:
            warnings.append(f"{filename} [{sheet_name}]: কোয়ান্টিটি-সহ কোনো ডাটা রো পাওয়া যায়নি")
            continue
        if total_c is not None and sum_c != round(total_c):
            warnings.append(f"{filename} [{sheet_name}]: Master Carton যোগফল ({sum_c}) ≠ Total Carton ({round(total_c)})")
        if total_t is not None and sum_t != round(total_t):
            warnings.append(f"{filename} [{sheet_name}]: Top Bottom যোগফল ({sum_t}) ≠ Total ({round(total_t)})")
        sheet_groups.append((cartons, tbs))

    if not sheet_groups:
        raise ValueError("কোনো শিটেই 'Style & PO / Master Carton qty' টেবিল পাওয়া যায়নি")

    if PAIR_TOP_BOTTOM:
        ordered_c = [it for c, t in sheet_groups for it in c + t]
        ordered_t = []
    else:
        ordered_c = [it for c, _t in sheet_groups for it in c]
        ordered_t = [it for _c, t in sheet_groups for it in t]

    # উপরের PO ফিল্ড = Style & PO (ফাইলের ইউনিক মানগুলো '/' দিয়ে)
    uniq = []
    for it in ordered_c + ordered_t:
        if it['po_no'] and it['po_no'] not in uniq:
            uniq.append(it['po_no'])
    header_po = '/'.join(uniq)
    for it in ordered_c + ordered_t:
        it['_header_po'] = header_po

    return ordered_c, ordered_t, warnings


def combine_bangladesh_textile_aeon_booking_files(files, item_name_override='', manual_ply=''):
    """files: [(file_stream, filename), ...] -> (line_items, warnings)
    সব ফাইলের কার্টন আগে, তারপর সব Top Bottom। ব্যর্থ ফাইল স্কিপ হয়ে মেসেজ
    warnings-এ যায়।"""
    all_c, all_t, warnings = [], [], []
    for file_stream, filename in files:
        try:
            file_stream.seek(0)
            c, t, w = read_bangladesh_textile_aeon_style_excel(
                file_stream, filename,
                item_name_override=item_name_override, manual_ply=manual_ply)
            all_c.extend(c)
            all_t.extend(t)
            warnings.extend(w)
        except Exception as e:
            warnings.append(f"{filename}: {e}")
    return all_c + all_t, warnings
