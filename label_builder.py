import os
from copy import copy
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font

from label_extractor import compute_qty_warnings

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template_files', 'template_label.xlsx')

# টেমপ্লেটের 'Ribbon Type' কলাম (A) এবং 'Item Name' কলাম (B) — ইউজার-কনফার্মড
RIBBON_TYPE = 'Paper'
ITEM_NAME = 'Additional Label'

# 'Measurement Unit' কলাম (R)। Width/Length (mm) UI থেকে মিলিমিটারেই নেওয়া হয়,
# তাই এখানে 'mm' বসে। ERP অন্য বানান চাইলে (যেমন 'Mm') শুধু এই লাইন বদলান।
MEASUREMENT_UNIT = ''

REQUIRED_FIELDS = ['style_no', 'po_no', 'size', 'qty']


def to_num(v):
    try:
        return float(str(v).replace(',', ''))
    except (ValueError, TypeError):
        return v


def na_if_blank(v):
    v = str(v).strip() if v is not None else ''
    return v if v else 'N/A'


def validate_label_line_items(line_items):
    warnings = []
    for i, item in enumerate(line_items):
        missing = [f for f in REQUIRED_FIELDS if not str(item.get(f, '')).strip()]
        if missing:
            warnings.append(f"Row {i + 1}: {', '.join(missing)} খালি আছে — চেক করুন")
    return warnings


def _strip_external_links(wb):
    try:
        if getattr(wb, '_external_links', None):
            wb._external_links = []
    except Exception:
        pass
    try:
        if hasattr(wb, '_external_references'):
            wb._external_references = []
    except Exception:
        pass


def _write_df_sheet(wb, sheet_name, df):
    if df is None or df.empty:
        return
    ws = wb.create_sheet(sheet_name)
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
    for cell in ws[1]:
        cell.font = Font(bold=True)


def build_label_excel(line_items, header_info, out_path,
                      customer_override=None, buyer_override=None,
                      po_override=None, delivery_date='', delivery_address='',
                      width_mm='', length_mm='', raw_df=None, summary_df=None,
                      warnings=None, remark_place=False, remark_address=False):
    """Label মডিউলের জন্য একটাই .xlsx বানায়। শীট: Sheet1 (Mapped Template), Raw Data,
    PO Details, PO Summary, Warnings।

    width_mm / length_mm: UI-তে ইনপুট দেওয়া মেজারমেন্ট (mm) — সব রো-তে এক ভ্যালু।
    ফাঁকা থাকলে কলাম ফাঁকাই থাকে (warning তৈরি করে না)।
    """
    wb = load_workbook(TEMPLATE_PATH)
    _strip_external_links(wb)
    ws = wb['Sheet1']

    ws['B2'] = po_override or header_info.get('po_number', '') or 'N/A'
    ws['B3'] = customer_override or header_info.get('customer', '') or 'N/A'
    ws['B4'] = buyer_override or header_info.get('buyer', '') or 'N/A'
    # B5 (Item Group = Label), B6 (Business Line = Label) টেমপ্লেট থেকেই থাকে

    sample_row = 8
    n_cols = 24  # A থেকে X
    style_ref = [copy(ws.cell(row=sample_row, column=c).font) for c in range(1, n_cols + 1)]
    fill_ref = [copy(ws.cell(row=sample_row, column=c).fill) for c in range(1, n_cols + 1)]
    align_ref = [copy(ws.cell(row=sample_row, column=c).alignment) for c in range(1, n_cols + 1)]
    border_ref = [copy(ws.cell(row=sample_row, column=c).border) for c in range(1, n_cols + 1)]
    numfmt_ref = [ws.cell(row=sample_row, column=c).number_format for c in range(1, n_cols + 1)]

    # টেমপ্লেটের স্যাম্পল রো (৮ থেকে নিচে) মুছে ফেলা — নাহলে লাইন-আইটেম কম হলে
    # স্যাম্পল ডাটা (100 / 0.1 / 20 / 35 ইত্যাদি) থেকে যেত
    if ws.max_row >= sample_row:
        ws.delete_rows(sample_row, ws.max_row - sample_row + 1)

    start_row = 8
    for i, r in enumerate(line_items):
        row = start_row + i

        remark_parts = []
        if remark_place and r.get('delivery_place_pdf'):
            remark_parts.append(f"Delivery Place: {r.get('delivery_place_pdf')}")
        if remark_address and r.get('delivery_address_pdf'):
            remark_parts.append(f"Delivery Address: {r.get('delivery_address_pdf')}")
        remarks_val = ' | '.join(remark_parts)

        values = [
            RIBBON_TYPE,                                   # A Ribbon Type
            ITEM_NAME,                                      # B Item Name
            na_if_blank(r.get('ewo_no')),                   # C Gmt. EWO No
            na_if_blank(r.get('style_no')),                 # D Gmt. Style No
            na_if_blank(r.get('po_no')),                    # E Gmt. PO
            'All',                                          # F Gmt. Destination
            na_if_blank(r.get('reference')),                # G Reference/SKU (Item Name/Color Code/USIM)
            'N/A',                                          # H Pack Type
            'N/A',                                          # I Gmt. Color
            na_if_blank(r.get('size')),                     # J Gmt. Size
            to_num(width_mm) if str(width_mm).strip() else '',     # K Width (mm)
            to_num(length_mm) if str(length_mm).strip() else '',   # L Length (mm)
            'N/A',                                          # M Composition
            'N/A',                                          # N Color
            'N/A',                                          # O Code/Caow
            'N/A',                                          # P Reference/USIM
            'N/A',                                          # Q Care Instruction
            MEASUREMENT_UNIT,                               # R Measurement Unit
            r.get('uom') or 'Pcs',                          # S UOM
            to_num(r.get('qty', '')),                       # T Order Qty
            to_num(r.get('rate', '')) if str(r.get('rate', '')).strip() else '',  # U Rate($)
            delivery_date or '',                            # V Delivery Date
            delivery_address or '',                         # W Delivery Place
            remarks_val,                                    # X Remarks
        ]
        for c, val in enumerate(values, start=1):
            cell = ws.cell(row=row, column=c, value=val)
            s = c - 1
            cell.font = style_ref[s]
            cell.fill = fill_ref[s]
            cell.alignment = align_ref[s]
            cell.border = border_ref[s]
            cell.number_format = numfmt_ref[s]

    _write_df_sheet(wb, 'Raw Data', pd.DataFrame(line_items))
    _write_df_sheet(wb, 'PO Details', raw_df)
    _write_df_sheet(wb, 'PO Summary', summary_df)

    all_warnings = list(warnings) if warnings else []
    all_warnings += compute_qty_warnings(line_items, summary_df)
    if all_warnings:
        _write_df_sheet(wb, 'Warnings', pd.DataFrame({'Warning': all_warnings}))

    wb.active = 0
    wb.save(out_path)
