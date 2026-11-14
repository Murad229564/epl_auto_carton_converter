import os
import re
import shutil
import subprocess
import tempfile
import pandas as pd

# ---------------------------------------------------------------------------
# Comfit Composite Knit Limited — Buyer: Original Marines
# 'COMFIT COMPOSITE KNIT LIMITED' কার্টন+পলি বুকিং এক্সেল (.xlsx/.xls)। এক ফাইলে
# একাধিক শিট থাকতে পারে (যেমন 'CARTON', 'Carton+Poly') — প্রতিটা শিট পড়া হয়।
#
# শিটের উপরের অংশ:
#   'CCKL :- 01193/26'  -> Gmt. PO (সব লাইনে) এবং টেমপ্লেটের উপরের PO ফিল্ড
#                           (_header_po) — ইউজার UI-তে PO ফাঁকা রাখলে এটাই বসে
#   'Style- :- B734FHDP' -> Gmt. Style No (Master Carton এবং Top Bottom সব লাইনে)
#
# শিটের ভেতরের ব্লকগুলো (ওপর থেকে নিচে, লেবেল দেখে চেনা হয়):
#   1. ASSORT   (A কলামে 'ASSORT')  — একটাই লাইন: L/W/H + Qty
#   2. LIFCO    (A কলামে 'Lifco')   — একটাই লাইন: L/W/H + Qty
#   3. SOLID    (A কলামে 'Logistic <কালার-কোড>') — প্রতি সাইজ (5-6, 6-7 ...)
#               একটা করে লাইন: L/W/H + Qty
#   4. Top Bottom ('Top Bottom' লেবেলের নিচে L/W/Unit/Qty) — প্রতি রো আলাদা লাইন,
#               Height নেই, Ply সবসময় 3
#   5. 'Poly Measurement' থেকে শুরু করে নিচের সবকিছু (পলি) পুরোপুরি বাদ।
#
# কোথায় কী বসে:
#   Length/Width/Height -> L / W / H কলাম (Unit 'CM'/'cm' থাকলে Cm)
#   Qty                 -> Qty কলাম
#   SOLID               -> Reference/SKU = কালার-কোড (যেমন '11-0601TCX'),
#                          Pack Type = সাইজ (যেমন '5-6')
#   ASSORT / LIFCO      -> Reference/SKU = 'ASSORT' / 'LIFCO' (ASSORT_REFERENCE_MODE
#                          = 'colors' করলে ব্লকের কালার-কোডগুলো), Pack Type ফাঁকা
#   'PER CTN 15 Pcs', '48 X 1 = 48', NET.W/W.T (ওজন) -> ব্যবহার হয় না
#   Item Name ও Ply (কার্টন লাইনের) -> UI থেকে (Ply না দিলে 'N/A'); Top Bottom
#                          সবসময় Item Name 'Top Bottom', Ply 3
# ক্রম: ফাইল-ক্রম, শিট-ক্রম; প্রতি শিটে ASSORT, LIFCO, SOLID..., তারপর Top Bottom।
# ---------------------------------------------------------------------------

NA = 'N/A'

# 'label' (ডিফল্ট): ASSORT/LIFCO লাইনের Reference = 'ASSORT'/'LIFCO'
# 'colors': Reference = ব্লকের কালার-কোডগুলো ('11-0601TCX / 17-1755TCX / 13-0117TCX')
ASSORT_REFERENCE_MODE = 'label'

# Top Bottom লাইনেও PO (CCKL নম্বর) বসবে কিনা। False করলে Top Bottom-এ PO = 'N/A'।
TB_INCLUDE_PO = True


def _norm(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean(v):
    if v is None:
        return ''
    if isinstance(v, float) and pd.isna(v):
        return ''
    return re.sub(r'\s+', ' ', str(v).replace('\xa0', ' ')).strip()


def _clean_id(v):
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
def _label_value(rows, label_norm, max_row=20):
    """উপরের অংশে লেবেল-সেল ('CCKL :-', 'Style- :-') খুঁজে, একই রো-তে ডানদিকের
    প্রথম নন-ফাঁকা সেলের মান রিটার্ন করে।"""
    for row in rows[:max_row]:
        for ci, c in enumerate(row):
            if _norm(c) == label_norm:
                for c2 in row[ci + 1:]:
                    v = _clean_id(c2)
                    if v:
                        return v
    return ''


def _unit_is_inch(text):
    return bool(re.search(r'inch|\bin\b', str(text or ''), re.I))


def _read_sheet(sheet_name, rows, filename, item_name, ply):
    """এক শিট -> (items, warnings)। এই ফরম্যাটের শিট না হলে (CCKL/Style নেই এবং
    কোনো ব্লকও নেই) ([], []) রিটার্ন করে।"""
    warnings = []
    cckl = _label_value(rows, 'cckl')
    style = _label_value(rows, 'style')

    items = []
    section = None            # None | 'carton' | 'tb'
    block_kind = ''           # 'ASSORT' | 'LIFCO' | 'SOLID'
    block_label = ''
    block_color = ''
    block_colors_text = ''
    carton_cols = {'l': None, 'w': None, 'h': None, 'qty': None, 'size': None}
    tb_cols = {'l': None, 'w': None, 'qty': None}
    unit_text = 'CM'
    seen_block = False

    for row in rows:
        if row is None:
            continue
        normed = [_norm(c) for c in row]
        a_norm = normed[0] if normed else ''

        # 'Poly Measurement' থেকে শুরু করে নিচের সবকিছু বাদ
        if any(n.startswith('polymeasurement') or n == 'polymeasurement' for n in normed if n):
            break

        # --- ব্লক-লেবেল ---
        if a_norm in ('assort', 'lifco') or a_norm.startswith('logistic'):
            seen_block = True
            section = 'carton'
            if a_norm.startswith('logistic'):
                block_kind = 'SOLID'
                color = re.sub(r'^\s*logistic\s*', '', _clean(row[0]), flags=re.I).strip()
                if not color:
                    color = next((_clean(c) for c in row[1:6] if _clean(c)), '')
                block_color = color
                block_label = color
            else:
                block_kind = a_norm.upper()
                block_label = block_kind
                # কালার-কোডের তালিকা (E কলামের মার্জড সেল) — 'colors' মোডের জন্য
                parts = []
                for c in row[1:8]:
                    t = _clean(c)
                    if t and re.search(r'\d', t) and re.search(r'[A-Za-z]', t) and t.lower() not in ('unit', 'qty'):
                        parts.extend(t.split())
                block_colors_text = ' / '.join(parts)
            # নতুন ব্লকে কলাম-ম্যাপ রিসেট (হেডার-রো আবার আসবে)
            carton_cols = {'l': None, 'w': None, 'h': None, 'qty': None, 'size': None}

        # --- Top Bottom লেবেল ---
        if any(n == 'topbottom' for n in normed if n):
            section = 'tb'
            tb_cols = {'l': None, 'w': None, 'qty': None}
            continue

        if section is None:
            continue

        # --- হেডার-রো ডিটেকশন (L / W / H / Qty / Unit) ---
        has_l = 'l' in normed
        has_w = 'w' in normed
        if has_l and has_w:
            if section == 'tb':
                tb_cols['l'] = normed.index('l')
                tb_cols['w'] = normed.index('w')
                if 'qty' in normed:
                    tb_cols['qty'] = normed.index('qty')
                continue
            carton_cols['l'] = normed.index('l')
            carton_cols['w'] = normed.index('w')
            if 'h' in normed:
                carton_cols['h'] = normed.index('h')
            if 'qty' in normed:
                carton_cols['qty'] = normed.index('qty')
            if 'unit' in normed:
                carton_cols['size'] = normed.index('unit') + 1   # Unit-এর পরের কলামে সাইজ
            continue
        if section == 'carton' and 'qty' in normed and carton_cols['qty'] is None:
            carton_cols['qty'] = normed.index('qty')
        if section == 'carton' and 'unit' in normed and carton_cols['size'] is None:
            carton_cols['size'] = normed.index('unit') + 1
        if section == 'carton' and 'qty' in normed and 'unit' in normed:
            carton_cols['qty'] = normed.index('qty')
            carton_cols['size'] = normed.index('unit') + 1

        def cell(ci):
            return row[ci] if ci is not None and ci < len(row) else None

        # --- ডাটা-রো ---
        if section == 'carton':
            L = _num(cell(carton_cols['l']))
            W = _num(cell(carton_cols['w']))
            qty = _num(cell(carton_cols['qty']))
            if L is None or W is None or qty is None or qty <= 0:
                continue
            H = _num(cell(carton_cols['h']))
            size = _clean(cell(carton_cols['size'])) if carton_cols['size'] is not None else ''
            if block_kind == 'SOLID':
                reference, pack_type = block_color, size
            else:
                reference = block_colors_text if (ASSORT_REFERENCE_MODE == 'colors' and block_colors_text) else block_label
                pack_type = ''
            items.append({
                'item_name': item_name,
                'ewo_no': NA,
                'style_no': style,
                'po_no': cckl,
                'length': _fmt(L),
                'width': _fmt(W),
                'height': _fmt(H) if H is not None else '',
                'ply': ply,
                'qty': round(qty),
                'pack_type': pack_type,
                'reference': reference,
                'color': '',
                'size': '',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': sheet_name,
                '_source_file': filename,
            })
        elif section == 'tb':
            L = _num(cell(tb_cols['l']))
            W = _num(cell(tb_cols['w']))
            qty = _num(cell(tb_cols['qty']))
            if L is None or W is None or qty is None or qty <= 0:
                continue
            items.append({
                'item_name': 'Top Bottom',
                'ewo_no': NA,
                'style_no': style,
                'po_no': cckl if TB_INCLUDE_PO else NA,
                'length': _fmt(L),
                'width': _fmt(W),
                'height': '',
                'ply': '3',
                'qty': round(qty),
                'pack_type': NA,
                'reference': NA,
                'color': '',
                'size': '',
                'delivery_date': '',
                'measurement_unit': 'Cm',
                'delivery_place_pdf': '',
                'delivery_address_pdf': '',
                '_sheet': sheet_name,
                '_source_file': filename,
            })

    if not seen_block and not items:
        return [], []

    if not cckl:
        warnings.append(f"{filename} [{sheet_name}]: 'CCKL' নম্বর পাওয়া যায়নি — PO ফাঁকা থাকবে")
    if not style:
        warnings.append(f"{filename} [{sheet_name}]: 'Style' পাওয়া যায়নি — Style ফাঁকা থাকবে")
    if not any(i['item_name'] != 'Top Bottom' for i in items):
        warnings.append(f"{filename} [{sheet_name}]: কোনো কার্টন লাইন পাওয়া যায়নি")
    if not any(i['item_name'] == 'Top Bottom' for i in items):
        warnings.append(f"{filename} [{sheet_name}]: Top Bottom লাইন পাওয়া যায়নি")

    header_po = cckl
    for it in items:
        it['_header_po'] = header_po
    return items, warnings


def read_comfit_original_marines_style_excel(file_stream, filename='', item_name_override='Master Carton', manual_ply=''):
    """এক ফাইল -> (items, warnings)। কোনো শিটেই এই ফরম্যাট না পেলে ValueError।"""
    sheets = _load_sheets(file_stream, filename)
    item_name = item_name_override or 'Master Carton'
    ply = manual_ply.strip() if manual_ply else NA

    items, warnings = [], []
    found = False
    for sheet_name, rows in sheets:
        sheet_items, sheet_warn = _read_sheet(sheet_name, rows, filename, item_name, ply)
        if sheet_items:
            found = True
        items.extend(sheet_items)
        warnings.extend(sheet_warn)
    if not found:
        raise ValueError("'CCKL' / ASSORT / Logistic / Top Bottom ব্লক-সহ Comfit টেবিল পাওয়া যায়নি")
    return items, warnings


def combine_comfit_original_marines_booking_files(files, item_name_override='', manual_ply=''):
    """files: [(file_stream, filename), ...] -> (line_items, warnings)
    ফাইল-ক্রমে একটাই লিস্ট। ব্যর্থ ফাইল স্কিপ হয়ে মেসেজ warnings-এ যায়।"""
    combined, warnings = [], []
    for file_stream, filename in files:
        try:
            file_stream.seek(0)
            its, w = read_comfit_original_marines_style_excel(
                file_stream, filename,
                item_name_override=item_name_override, manual_ply=manual_ply)
            combined.extend(its)
            warnings.extend(w)
        except Exception as e:
            warnings.append(f"{filename}: {e}")
    return combined, warnings
