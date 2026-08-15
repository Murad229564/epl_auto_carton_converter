import io
import os
import re
import tempfile

from flask import Flask, request, render_template, send_file, jsonify

from extractor import process_pdf_rule_based, get_unique_delivery_info
from builder import build_combined_excel, validate_line_items, build_pdf_full_dump, build_excel_full_dump
from outhouse_extractor import combine_booking_excels
from outhouse_pdf_extractor import process_trims_booking_pdf
from ikl_biscana_extractor import read_ikl_biscana_pdf
from kenpark_extractor import read_kenpark_pdf
from thermal_extractor import process_pdf_thermal, get_unique_delivery_info_thermal
from thermal_builder import build_thermal_excel, validate_thermal_line_items
from thermal_config import THERMAL_BUYERS, THERMAL_BUYER_ALIASES, THERMAL_VERIFIED_BUYERS
from printing_press_extractor import process_pdf_pp, get_unique_delivery_info_pp
from printing_press_builder import build_pp_excel, validate_pp_line_items
from printing_press_config import (
    PRINTING_PRESS_BUYERS, PRINTING_PRESS_BUYER_ALIASES, PRINTING_PRESS_VERIFIED_BUYERS,
    PRINTING_PRESS_ITEM_NAME_ALIASES,
)
from validators import (
    validate_customer, validate_buyer, validate_buyer_in_list, validate_po_number,
    validate_delivery_address, validate_matches_pdf, values_match_ci,
)
from date_logic import get_default_delivery_date, validate_manual_delivery_date, format_delivery_date
from config import CUSTOMERS, BUYERS, DELIVERY_ADDRESSES, BUYER_ALIASES, CUSTOMER_ALIASES, CARTON_VERIFIED_BUYERS, CUSTOMER_BUYER_MAP, resolve_alias

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB


def _norm_key(s):
    """Customer/Buyer নাম মেলানোর জন্য normalize — স্পেস/পাংচুয়েশন/কেস
    বাদ দিয়ে, যাতে সামান্য বানান-ভিন্নতাতেও সঠিক এন্ট্রি মেলে।"""
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _safe_filename_part(s):
    """ফাইলনেমে ব্যবহারের অযোগ্য ক্যারেক্টার ('/', '\\' ইত্যাদি) '-' দিয়ে বদলে দেয়,
    যাতে tempfile.TemporaryDirectory-এর ভেতরে ভুল সাব-ফোল্ডার তৈরির চেষ্টা না হয়।"""
    return re.sub(r'[\\/:*?"<>|]', '-', str(s or ''))


# ---------------------------------------------------------------------------
# OUT-HOUSE PDF (Trims Booking রেডিও অপশন) — customer+buyer অনুযায়ী সঠিক
# extractor-এ ডাইনামিকভাবে রুট করার রেজিস্ট্রি। ডিফল্ট (রেজিস্ট্রিতে না
# থাকলে) সবসময় আগের মতোই Barnali/Modele-স্টাইল process_trims_booking_pdf
# ব্যবহার হবে — তাই এই ফরম্যাটগুলোর জন্য কিচ্ছু বদলায়নি।
#
# নতুন কোনো "টোটালি ডিফারেন্ট" PDF ফরম্যাটের কাস্টমার এলে (যেমন Innovative
# Knitex Ltd./Biscana), ভবিষ্যতে UI-তে নতুন কোনো radio বাড়াতে হবে না —
# এখানে শুধু একটা এন্ট্রি যোগ করলেই হবে।
#
# uniform wrapper সিগনেচার: (file_stream, filename, customer_list,
# buyer_list, item_name_override, manual_ply) -> (header_info, items)
# ---------------------------------------------------------------------------
def _wrap_barnali_pdf(file_stream, filename, customer_list, buyer_list, item_name_override, manual_ply):
    return process_trims_booking_pdf(file_stream, customer_list, buyer_list)


def _wrap_ikl_pdf(file_stream, filename, customer_list, buyer_list, item_name_override, manual_ply):
    return read_ikl_biscana_pdf(
        file_stream, filename,
        item_name_override=item_name_override, manual_ply=manual_ply)


OUTHOUSE_PDF_REGISTRY = {
    (_norm_key('Innovative Knitex Ltd.'), _norm_key('Biscana')): _wrap_ikl_pdf,
}


def _get_outhouse_pdf_handler(customer_name, buyer_name):
    """customer_name/buyer_name জানা থাকলে (অর্থাৎ ফর্ম সাবমিটের সময়,
    /autocarton/process_outhouse_trims_booking_pdf-এ) নির্ভরযোগ্যভাবে
    সঠিক extractor বেছে নেয়। রেজিস্ট্রিতে না থাকলে ডিফল্ট Barnali/Modele
    হ্যান্ডলার।"""
    return OUTHOUSE_PDF_REGISTRY.get(
        (_norm_key(customer_name), _norm_key(buyer_name)), _wrap_barnali_pdf)


def _extract_outhouse_pdf_header_auto(file_stream, filename):
    """হেডার-অটোফিল ধাপে (PDF আপলোডের সাথে সাথেই, ফর্ম সাবমিটের আগে)
    customer/buyer এখনো জানা নেই — তাই ফরম্যাট নিজে থেকেই বুঝে নিতে হয়।
    প্রথমে ডিফল্ট (Barnali/Modele) ফরম্যাট ট্রাই করা হয়; সেটা যদি লাইন-
    আইটেম/PO না দেয় (ফরম্যাট না মেলার লক্ষণ), তাহলে রেজিস্ট্রিতে থাকা
    বাকি ফরম্যাটগুলো (যেমন IKL/Biscana) একে একে ট্রাই করা হয়। প্রতিটা
    extractor নিজের ফরম্যাট না মিললে নিরাপদে খালি রেজাল্ট রিটার্ন করে
    (এক্সেপশন ছোড়ে না), তাই এই চেইন নিরাপদ।"""
    try:
        file_stream.seek(0)
        header_info, items = process_trims_booking_pdf(file_stream, CUSTOMERS.get('OUT-HOUSE', []), BUYERS)
        if items or header_info.get('po_number'):
            return header_info, items
    except Exception:
        pass

    for handler in OUTHOUSE_PDF_REGISTRY.values():
        try:
            file_stream.seek(0)
            header_info, items = handler(
                file_stream, filename, CUSTOMERS.get('OUT-HOUSE', []), BUYERS, '', '')
            if items or header_info.get('po_number'):
                return header_info, items
        except Exception:
            continue

    return {'po_number': '', 'customer': '', 'buyer': ''}, []


@app.route('/')
def modules_home():
    """সব মডিউলের লিস্ট — এখান থেকে ক্লিক করে ভেতরের মডিউলে ঢোকা যাবে।"""
    return render_template('modules.html')


@app.route('/autocarton')
def autocarton_index():
    return render_template(
        'index.html',
        customers=CUSTOMERS,
        buyers=BUYERS,
        delivery_addresses=DELIVERY_ADDRESSES,
        customer_buyer_map=CUSTOMER_BUYER_MAP,
    )


@app.route('/thermal')
def thermal_index():
    return render_template(
        'thermal.html',
        customers=CUSTOMERS,
        buyers=THERMAL_BUYERS,
        delivery_addresses=DELIVERY_ADDRESSES,
    )


@app.route('/printing_press')
def printing_press_index():
    return render_template(
        'printing_press.html',
        customers=CUSTOMERS,
        buyers=PRINTING_PRESS_BUYERS,
        delivery_addresses=DELIVERY_ADDRESSES,
    )


@app.route('/extract_header', methods=['POST'])
def extract_header():
    """PDF আপলোড হওয়ার সাথে সাথেই (ফর্ম সাবমিটের আগেই) শুধু হেডার তথ্য
    (PO Number/Customer/Buyer) বের করে ফেরত দেয়, যাতে ফ্রন্টএন্ড সাথে সাথে
    এগুলো ফিল্ডে বসিয়ে দিতে পারে। এখানে কোনো Excel বানানো হয় না, শুধু
    দ্রুত extract করে JSON রিটার্ন করে।

    একাধিক ফাইল সিলেক্ট করা থাকলেও (নিচের /process-এ একসাথে একাধিক PDF
    পাঠানো যায়) এই autofill endpoint শুধু প্রথম ফাইলটার হেডারই দেখে —
    বাকি ফাইলগুলো একই Customer/Buyer-এর অধীনে হবে এই ধরে নিয়ে।"""
    if 'pdf_file' not in request.files:
        return jsonify({'error': 'PDF ফাইল পাওয়া যায়নি'}), 400

    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return jsonify({'error': 'ফাইল সিলেক্ট করা হয়নি'}), 400

    pdf_bytes_raw = pdf_file.read()
    try:
        header_info, line_items, raw_df, summary_df = process_pdf_rule_based(io.BytesIO(pdf_bytes_raw))
    except Exception as e:
        return jsonify({'error': f'PDF থেকে তথ্য বের করতে সমস্যা হয়েছে: {str(e)}'}), 422

    # PDF-এ 'M&S', 'DEKKO KNITWEARS LTD.'-এর মতো সংক্ষিপ্ত/ভিন্ন নাম থাকলে এখানেই
    # config.py-এর BUYER_ALIASES/CUSTOMER_ALIASES দিয়ে ক্যানোনিকাল নামে বদলে দেওয়া
    # হচ্ছে — যাতে ফ্রন্টএন্ডের ফিল্ড ঠিকভাবে অটো-লক হয় এবং পরে /process-এ
    # মিসম্যাচ এরর না আসে।
    header_info['buyer'] = resolve_alias(header_info.get('buyer', ''), BUYER_ALIASES)
    header_info['customer'] = resolve_alias(header_info.get('customer', ''), CUSTOMER_ALIASES)

    delivery_info = get_unique_delivery_info(raw_df)

    return jsonify({
        'po_number': header_info.get('po_number', '') or '',
        'customer': header_info.get('customer', '') or '',
        'buyer': header_info.get('buyer', '') or '',
        'delivery_places_pdf': delivery_info['delivery_places'],
        'delivery_addresses_pdf': delivery_info['delivery_addresses'],
    })


@app.route('/process', methods=['POST'])
def process():
    """IN-HOUSE PDF প্রসেসিং। একাধিক PDF একসাথে আপলোড করা যায় ('files' ফিল্ড,
    একাধিক এন্ট্রি) — ডিফল্টভাবে সবগুলো ফাইলের লাইন-আইটেম একটাই কম্বাইনড
    Excel-এ বসে (আগে যেমন একটা ফাইলের জন্য হতো)। 'separate_output' চেকমার্ক
    করা থাকলে প্রতিটা PDF-এর জন্য আলাদা Excel বানিয়ে একটা Zip-এ দেওয়া হয়
    (OUT-HOUSE Excel ফ্লো-র 'প্রতিটা ফাইল আলাদা Excel হিসেবে' চেকবক্সের
    মতোই কনভেনশন)।

    Backward-compat: পুরনো ক্লায়েন্ট যদি এখনো single 'pdf_file' পাঠায়,
    সেটাও কাজ করবে (এক-ফাইল getlist-এর মতোই ট্রিট হবে)।
    """
    files = request.files.getlist('files')
    if not files or not any(f and f.filename for f in files):
        single = request.files.get('pdf_file')
        files = [single] if single and single.filename else []
    files = [f for f in files if f and f.filename]
    if not files:
        return jsonify({'error': 'PDF ফাইল পাওয়া যায়নি'}), 400

    customer_type = request.form.get('customer_type', 'IN-HOUSE').strip()
    customer_name = request.form.get('customer_name', '').strip()
    buyer_name = request.form.get('buyer_name', '').strip()
    po_number_override = request.form.get('po_number', '').strip()
    delivery_mode = request.form.get('delivery_mode', 'auto').strip()
    delivery_date_manual = request.form.get('delivery_date', '').strip()
    delivery_address = request.form.get('delivery_address', '').strip()
    method = request.form.get('method', 'rule_based')
    remark_place = request.form.get('remark_place', '').strip().lower() in ('1', 'true', 'on', 'yes')
    remark_address = request.form.get('remark_address', '').strip().lower() in ('1', 'true', 'on', 'yes')
    # ইমারজেন্সি ফোর্স ওভাররাইড: চেক করা থাকলে PDF-এর সাথে Customer/Buyer না
    # মিললেও এরর দেওয়া হবে না, ম্যানুয়ালি দেওয়া নামটাই ব্যবহার হবে (ERP লিস্টের
    # সাথে case-sensitive মেলার শর্ত অবশ্য তখনও বহাল থাকবে)।
    force_override = request.form.get('force_override', '').strip().lower() in ('1', 'true', 'on', 'yes')
    separate_output = request.form.get('separate_output', '').strip().lower() in ('1', 'true', 'on', 'yes')

    if method not in ('rule_based',):
        if method == 'ai_based':
            return jsonify({
                'error': 'AI-Based মেথড এখনো চালু করা হয়নি। শীঘ্রই আসছে — আপাতত Rule-Based ব্যবহার করুন।'
            }), 501
        return jsonify({'error': f'অজানা মেথড: {method}'}), 400

    # --- Buyer বাধ্যতামূলক ও case-sensitive লিস্ট-ম্যাচ ---
    buyer_error = validate_buyer(buyer_name)
    if buyer_error:
        return jsonify({'error': buyer_error}), 422

    # --- Customer আবশ্যক ও case-sensitive লিস্ট-ম্যাচ ---
    customer_error = validate_customer(customer_type, customer_name)
    if customer_error:
        return jsonify({'error': customer_error}), 422

    # --- Delivery Address: যে Customer-এর জন্য address লিস্ট configure করা আছে, তার জন্য আবশ্যক ---
    address_error = validate_delivery_address(customer_name, delivery_address)
    if address_error:
        return jsonify({'error': address_error}), 422

    # --- Delivery Date: manual হলে আগেই ভ্যালিডেট করে নেওয়া (PDF পড়ার আগে, সময় বাঁচাতে) ---
    if delivery_mode == 'manual':
        is_valid, err, parsed_date = validate_manual_delivery_date(delivery_date_manual)
        if not is_valid:
            return jsonify({'error': err}), 422
        delivery_date_final = format_delivery_date(parsed_date)
    else:
        delivery_date_final = format_delivery_date(get_default_delivery_date())

    # --- প্রতিটা PDF আলাদাভাবে পড়ে, নিজের হেডার-ম্যাচ (PO/Customer/Buyer)
    # ভ্যালিডেট করে — কোনো একটা ফাইলে মিসম্যাচ/সমস্যা হলে সেই ফাইলটাই বাদ
    # (Warnings-এ নোট থাকবে), বাকি ফাইলগুলো স্বাভাবিকভাবে প্রসেস চলতে থাকে
    # (OUT-HOUSE ফ্লো-র resilience কনভেনশনের সাথে মিলিয়ে) ---
    per_file_results = []  # [(filename, header_info, line_items, raw_df, summary_df, pdf_bytes), ...]
    file_errors = []

    for pdf_file in files:
        pdf_bytes_raw = pdf_file.read()
        try:
            header_info, line_items, raw_df, summary_df = process_pdf_rule_based(io.BytesIO(pdf_bytes_raw))
        except Exception as e:
            file_errors.append(f"{pdf_file.filename}: PDF পড়তে সমস্যা হয়েছে (rule-based): {str(e)}")
            continue

        if not line_items:
            file_errors.append(f"{pdf_file.filename}: কোনো লাইন-আইটেম পাওয়া যায়নি")
            continue

        header_info['buyer'] = resolve_alias(header_info.get('buyer', ''), BUYER_ALIASES)
        header_info['customer'] = resolve_alias(header_info.get('customer', ''), CUSTOMER_ALIASES)

        po_error = validate_po_number(po_number_override, header_info.get('po_number', ''))
        if po_error:
            file_errors.append(f"{pdf_file.filename}: {po_error}")
            continue

        if not force_override:
            customer_pdf_error = validate_matches_pdf('Customer', customer_name, header_info.get('customer', ''))
            if customer_pdf_error:
                file_errors.append(f"{pdf_file.filename}: {customer_pdf_error}")
                continue
            buyer_pdf_error = validate_matches_pdf('Buyer', buyer_name, header_info.get('buyer', ''))
            if buyer_pdf_error:
                file_errors.append(f"{pdf_file.filename}: {buyer_pdf_error}")
                continue

        per_file_results.append((pdf_file.filename, header_info, line_items, raw_df, summary_df, pdf_bytes_raw))

    if not per_file_results:
        msg = 'কোনো লাইন-আইটেম পাওয়া যায়নি।'
        if file_errors:
            msg += ' সমস্যা: ' + '; '.join(file_errors)
        return jsonify({'error': msg}), 422

    def _build_force_override_notes(header_info):
        notes = []
        if force_override:
            pdf_customer = header_info.get('customer', '')
            pdf_buyer = header_info.get('buyer', '')
            if pdf_customer and not values_match_ci(customer_name, pdf_customer):
                notes.append(
                    f"⚠️ FORCE OVERRIDE: Customer ম্যানুয়ালি '{customer_name}' বসানো হয়েছে, "
                    f"কিন্তু PDF-এ ছিল '{pdf_customer}' — দয়া করে যাচাই করুন।"
                )
            if pdf_buyer and not values_match_ci(buyer_name, pdf_buyer):
                notes.append(
                    f"⚠️ FORCE OVERRIDE: Buyer ম্যানুয়ালি '{buyer_name}' বসানো হয়েছে, "
                    f"কিন্তু PDF-এ ছিল '{pdf_buyer}' — দয়া করে যাচাই করুন।"
                )
        return notes

    verified_warning = None
    if buyer_name not in CARTON_VERIFIED_BUYERS:
        verified_warning = (
            f"⚠️ '{buyer_name}' buyer-এর Carton PDF ফরম্যাট এখনো নির্দিষ্টভাবে "
            f"যাচাই করা হয়নি — আউটপুট ভালোভাবে চেক করে নিন।"
        )

    # ============================================================
    # SEPARATE OUTPUT (ZIP) মোড — প্রতিটা PDF-এর জন্য আলাদা Excel
    # ============================================================
    if separate_output:
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                import zipfile
                zip_path = os.path.join(tmpdir, 'AutoCarton_Outputs.zip')
                total_warn_count = 0
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for filename, header_info, line_items, raw_df, summary_df, pdf_bytes_raw in per_file_results:
                        warnings = validate_line_items(line_items)
                        if verified_warning:
                            warnings.append(verified_warning)
                        warnings.extend(_build_force_override_notes(header_info))
                        total_warn_count += len(warnings)

                        base_name = _safe_filename_part(os.path.splitext(filename)[0])
                        out_name = f'{base_name}_Output.xlsx'
                        out_path = os.path.join(tmpdir, out_name)
                        build_combined_excel(
                            line_items, header_info, out_path,
                            profile=customer_type,
                            customer_override=customer_name or None,
                            buyer_override=buyer_name or None,
                            po_override=po_number_override or None,
                            delivery_date=delivery_date_final,
                            delivery_address=delivery_address,
                            raw_df=raw_df,
                            summary_df=summary_df,
                            warnings=warnings,
                            remark_place=remark_place,
                            remark_address=remark_address,
                            full_dump=[build_pdf_full_dump(io.BytesIO(pdf_bytes_raw), filename)],
                        )
                        zf.write(out_path, arcname=out_name)
                with open(zip_path, 'rb') as f:
                    zip_bytes = f.read()
        except Exception as e:
            return jsonify({'error': f'Zip ফাইল বানাতে সমস্যা হয়েছে: {str(e)}'}), 500

        buf = io.BytesIO(zip_bytes)
        response = send_file(
            buf, as_attachment=True, download_name='AutoCarton_Outputs.zip',
            mimetype='application/zip',
        )
        response.headers['Content-Length'] = str(len(zip_bytes))
        response.headers['X-Warning-Count'] = str(total_warn_count)
        response.headers['X-File-Count'] = str(len(per_file_results))
        return response

    # ============================================================
    # কম্বাইনড মোড (ডিফল্ট) — সব ফাইলের লাইন-আইটেম একটাই Excel-এ
    # ============================================================
    combined_line_items = []
    combined_full_dump = []
    combined_header_info = per_file_results[0][1]
    combined_raw_df = per_file_results[0][3]
    combined_summary_df = per_file_results[0][4]
    force_override_notes = []
    for filename, header_info, line_items, raw_df, summary_df, pdf_bytes_raw in per_file_results:
        combined_line_items.extend(line_items)
        combined_full_dump.append(build_pdf_full_dump(io.BytesIO(pdf_bytes_raw), filename))
        force_override_notes.extend(_build_force_override_notes(header_info))

    warnings = validate_line_items(combined_line_items)
    if verified_warning:
        warnings.append(verified_warning)
    warnings.extend(force_override_notes)
    for e in file_errors:
        warnings.append(f"⚠️ এই ফাইলটা স্কিপ হয়েছে: {e}")

    base_name = _safe_filename_part(
        os.path.splitext(per_file_results[0][0])[0]
        if len(per_file_results) == 1
        else f"{customer_name}_{buyer_name}_Combined".replace(' ', '_')
    )

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, f'{base_name}_Output.xlsx')
            build_combined_excel(
                combined_line_items, combined_header_info, out_path,
                profile=customer_type,
                customer_override=customer_name or None,
                buyer_override=buyer_name or None,
                po_override=po_number_override or None,
                delivery_date=delivery_date_final,
                delivery_address=delivery_address,
                raw_df=combined_raw_df,
                summary_df=combined_summary_df,
                warnings=warnings,
                remark_place=remark_place,
                remark_address=remark_address,
                full_dump=combined_full_dump,
            )
            with open(out_path, 'rb') as f:
                file_bytes = f.read()
    except Exception as e:
        # Excel বানাতে গিয়ে কোনো সমস্যা হলে যেন কখনোই ভাঙা/অসম্পূর্ণ ফাইল
        # ডাউনলোড না হয়ে যায় — বরং সাফ একটা এরর মেসেজ দেখানো হয়
        return jsonify({'error': f'Excel ফাইল বানাতে সমস্যা হয়েছে: {str(e)}'}), 500

    if not file_bytes:
        return jsonify({'error': 'Excel ফাইল খালি তৈরি হয়েছে — আবার চেষ্টা করুন'}), 500

    buf = io.BytesIO(file_bytes)
    response = send_file(
        buf,
        as_attachment=True,
        download_name=f'{base_name}_Output.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response.headers['Content-Length'] = str(len(file_bytes))
    response.headers['X-Warning-Count'] = str(len(warnings))
    response.headers['X-File-Count'] = str(len(per_file_results))
    return response


@app.route('/thermal/extract_header', methods=['POST'])
def thermal_extract_header():
    """Carton-এর /extract_header-এর মতোই — PDF আপলোড হওয়ার সাথে সাথেই
    PO Number/Customer/Buyer এবং Delivery Place/Address হিন্ট বের করে ফেরত দেয়।"""
    if 'pdf_file' not in request.files:
        return jsonify({'error': 'PDF ফাইল পাওয়া যায়নি'}), 400

    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return jsonify({'error': 'ফাইল সিলেক্ট করা হয়নি'}), 400

    pdf_bytes_raw = pdf_file.read()
    try:
        header_info, line_items, raw_df, summary_df = process_pdf_thermal(io.BytesIO(pdf_bytes_raw))
    except Exception as e:
        return jsonify({'error': f'PDF থেকে তথ্য বের করতে সমস্যা হয়েছে: {str(e)}'}), 422

    header_info['buyer'] = resolve_alias(header_info.get('buyer', ''), THERMAL_BUYER_ALIASES)
    header_info['customer'] = resolve_alias(header_info.get('customer', ''), CUSTOMER_ALIASES)

    delivery_info = get_unique_delivery_info_thermal(raw_df)

    return jsonify({
        'po_number': header_info.get('po_number', '') or '',
        'customer': header_info.get('customer', '') or '',
        'buyer': header_info.get('buyer', '') or '',
        'delivery_places_pdf': delivery_info['delivery_places'],
        'delivery_addresses_pdf': delivery_info['delivery_addresses'],
    })


@app.route('/thermal/process', methods=['POST'])
def thermal_process():
    if 'pdf_file' not in request.files:
        return jsonify({'error': 'PDF ফাইল পাওয়া যায়নি'}), 400

    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return jsonify({'error': 'ফাইল সিলেক্ট করা হয়নি'}), 400

    customer_type = 'IN-HOUSE'  # Thermal মডিউলে আপাতত শুধু IN-HOUSE সাপোর্ট করা হচ্ছে
    customer_name = request.form.get('customer_name', '').strip()
    buyer_name = request.form.get('buyer_name', '').strip()
    po_number_override = request.form.get('po_number', '').strip()
    delivery_mode = request.form.get('delivery_mode', 'auto').strip()
    delivery_date_manual = request.form.get('delivery_date', '').strip()
    delivery_address = request.form.get('delivery_address', '').strip()
    measurement = request.form.get('measurement', '').strip()
    remark_place = request.form.get('remark_place', '').strip().lower() in ('1', 'true', 'on', 'yes')
    remark_address = request.form.get('remark_address', '').strip().lower() in ('1', 'true', 'on', 'yes')
    force_override = request.form.get('force_override', '').strip().lower() in ('1', 'true', 'on', 'yes')

    buyer_error = validate_buyer_in_list(buyer_name, THERMAL_BUYERS)
    if buyer_error:
        return jsonify({'error': buyer_error}), 422

    customer_error = validate_customer(customer_type, customer_name)
    if customer_error:
        return jsonify({'error': customer_error}), 422

    address_error = validate_delivery_address(customer_name, delivery_address)
    if address_error:
        return jsonify({'error': address_error}), 422

    if delivery_mode == 'manual':
        is_valid, err, parsed_date = validate_manual_delivery_date(delivery_date_manual)
        if not is_valid:
            return jsonify({'error': err}), 422
        delivery_date_final = format_delivery_date(parsed_date)
    else:
        delivery_date_final = format_delivery_date(get_default_delivery_date())

    pdf_bytes_raw = pdf_file.read()

    try:
        header_info, line_items, raw_df, summary_df = process_pdf_thermal(io.BytesIO(pdf_bytes_raw))
    except Exception as e:
        return jsonify({'error': f'PDF পড়তে সমস্যা হয়েছে: {str(e)}'}), 422

    if not line_items:
        return jsonify({'error': 'কোনো লাইন-আইটেম পাওয়া যায়নি এই PDF থেকে'}), 422

    header_info['buyer'] = resolve_alias(header_info.get('buyer', ''), THERMAL_BUYER_ALIASES)
    header_info['customer'] = resolve_alias(header_info.get('customer', ''), CUSTOMER_ALIASES)

    po_error = validate_po_number(po_number_override, header_info.get('po_number', ''))
    if po_error:
        return jsonify({'error': po_error}), 422

    customer_pdf_error = None if force_override else validate_matches_pdf(
        'Customer', customer_name, header_info.get('customer', ''))
    if customer_pdf_error:
        return jsonify({'error': customer_pdf_error}), 422

    buyer_pdf_error = None if force_override else validate_matches_pdf(
        'Buyer', buyer_name, header_info.get('buyer', ''))
    if buyer_pdf_error:
        return jsonify({'error': buyer_pdf_error}), 422

    warnings = validate_thermal_line_items(line_items)

    # buyer সিস্টেমে (মাস্টার লিস্টে) থাকলেই যথেষ্ট প্রসেসিং চালানোর জন্য —
    # কিন্তু এই buyer-এর Thermal PDF ফরম্যাট এখনো নির্দিষ্টভাবে যাচাই করা না
    # থাকলে ব্লক না করে শুধু একটা সতর্কতা যোগ করা হচ্ছে, যাতে ইউজার আউটপুট
    # ভালোভাবে চেক করে নিতে পারেন।
    if buyer_name not in THERMAL_VERIFIED_BUYERS:
        warnings.append(
            f"⚠️ '{buyer_name}' buyer-এর Thermal PDF ফরম্যাট এখনো নির্দিষ্টভাবে "
            f"যাচাই করা হয়নি — আউটপুট (বিশেষ করে সাইজ/কোয়ান্টিটি/রেফারেন্স) "
            f"ভালোভাবে চেক করে নিন।"
        )

    if force_override:
        pdf_customer = header_info.get('customer', '')
        pdf_buyer = header_info.get('buyer', '')
        if pdf_customer and not values_match_ci(customer_name, pdf_customer):
            warnings.append(
                f"⚠️ FORCE OVERRIDE: Customer ম্যানুয়ালি '{customer_name}' বসানো হয়েছে, "
                f"কিন্তু PDF-এ ছিল '{pdf_customer}' — দয়া করে যাচাই করুন।"
            )
        if pdf_buyer and not values_match_ci(buyer_name, pdf_buyer):
            warnings.append(
                f"⚠️ FORCE OVERRIDE: Buyer ম্যানুয়ালি '{buyer_name}' বসানো হয়েছে, "
                f"কিন্তু PDF-এ ছিল '{pdf_buyer}' — দয়া করে যাচাই করুন।"
            )

    if not measurement:
        warnings.append("Measurement ফাঁকা রাখা হয়েছে (ইউজার confirm করেছেন) — পরে ম্যানুয়ালি বসাতে হবে।")

    base_name = os.path.splitext(pdf_file.filename)[0]

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, f'{base_name}_Output.xlsx')
            build_thermal_excel(
                line_items, header_info, out_path,
                customer_override=customer_name or None,
                buyer_override=buyer_name or None,
                po_override=po_number_override or None,
                delivery_date=delivery_date_final,
                delivery_address=delivery_address,
                measurement=measurement,
                raw_df=raw_df,
                summary_df=summary_df,
                warnings=warnings,
                remark_place=remark_place,
                remark_address=remark_address,
            )
            with open(out_path, 'rb') as f:
                file_bytes = f.read()
    except Exception as e:
        return jsonify({'error': f'Excel ফাইল বানাতে সমস্যা হয়েছে: {str(e)}'}), 500

    if not file_bytes:
        return jsonify({'error': 'Excel ফাইল খালি তৈরি হয়েছে — আবার চেষ্টা করুন'}), 500

    buf = io.BytesIO(file_bytes)
    response = send_file(
        buf,
        as_attachment=True,
        download_name=f'{base_name}_Output.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response.headers['Content-Length'] = str(len(file_bytes))
    response.headers['X-Warning-Count'] = str(len(warnings))
    return response


@app.route('/printing_press/extract_header', methods=['POST'])
def printing_press_extract_header():
    """Thermal-এর /thermal/extract_header-এর সাথে হুবহু এক প্যাটার্ন — শুধু
    এখানে অতিরিক্ত 'item_name'-ও ফেরত পাঠানো হয় (PDF-এর কভার পেজ থেকে
    ডাইনামিকভাবে পড়া, P.S Tag/Poly Sticker যা-ই হোক), যাতে ফ্রন্টএন্ড সেটাও
    Item Name ফিল্ডে অটো বসিয়ে দিতে পারে।"""
    if 'pdf_file' not in request.files:
        return jsonify({'error': 'PDF ফাইল পাওয়া যায়নি'}), 400

    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return jsonify({'error': 'ফাইল সিলেক্ট করা হয়নি'}), 400

    pdf_bytes_raw = pdf_file.read()
    try:
        header_info, line_items, raw_df, summary_df = process_pdf_pp(io.BytesIO(pdf_bytes_raw))
    except Exception as e:
        return jsonify({'error': f'PDF থেকে তথ্য বের করতে সমস্যা হয়েছে: {str(e)}'}), 422

    header_info['buyer'] = resolve_alias(header_info.get('buyer', ''), PRINTING_PRESS_BUYER_ALIASES)
    header_info['customer'] = resolve_alias(header_info.get('customer', ''), CUSTOMER_ALIASES)
    header_info['item_name'] = resolve_alias(header_info.get('item_name', ''), PRINTING_PRESS_ITEM_NAME_ALIASES)

    delivery_info = get_unique_delivery_info_pp(raw_df)

    return jsonify({
        'po_number': header_info.get('po_number', '') or '',
        'customer': header_info.get('customer', '') or '',
        'buyer': header_info.get('buyer', '') or '',
        'item_name': header_info.get('item_name', '') or '',
        'delivery_places_pdf': delivery_info['delivery_places'],
        'delivery_addresses_pdf': delivery_info['delivery_addresses'],
    })


@app.route('/printing_press/process', methods=['POST'])
def printing_press_process():
    """Thermal-এর /thermal/process-এর সাথে হুবহু এক প্যাটার্ন — পার্থক্য শুধু:
    - measurement-এর বদলে length/width/height/measurement_unit (আলাদা ফিল্ড)
    - Item Name PDF থেকে ডাইনামিক (ফিক্সড 'Thermal Sticker'-এর মতো না),
      item_name ফর্ম-ফিল্ড দিয়ে ওভাররাইড করা যায়।
    """
    if 'pdf_file' not in request.files:
        return jsonify({'error': 'PDF ফাইল পাওয়া যায়নি'}), 400

    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return jsonify({'error': 'ফাইল সিলেক্ট করা হয়নি'}), 400

    customer_type = 'IN-HOUSE'  # Printing Press মডিউলেও আপাতত শুধু IN-HOUSE সাপোর্ট করা হচ্ছে
    customer_name = request.form.get('customer_name', '').strip()
    buyer_name = request.form.get('buyer_name', '').strip()
    po_number_override = request.form.get('po_number', '').strip()
    item_name_override = request.form.get('item_name', '').strip()
    delivery_mode = request.form.get('delivery_mode', 'auto').strip()
    delivery_date_manual = request.form.get('delivery_date', '').strip()
    delivery_address = request.form.get('delivery_address', '').strip()
    length = request.form.get('length', '').strip()
    width = request.form.get('width', '').strip()
    height = request.form.get('height', '').strip()
    measurement_unit = request.form.get('measurement_unit', 'CM').strip() or 'CM'
    remark_place = request.form.get('remark_place', '').strip().lower() in ('1', 'true', 'on', 'yes')
    remark_address = request.form.get('remark_address', '').strip().lower() in ('1', 'true', 'on', 'yes')
    force_override = request.form.get('force_override', '').strip().lower() in ('1', 'true', 'on', 'yes')

    buyer_error = validate_buyer_in_list(buyer_name, PRINTING_PRESS_BUYERS)
    if buyer_error:
        return jsonify({'error': buyer_error}), 422

    customer_error = validate_customer(customer_type, customer_name)
    if customer_error:
        return jsonify({'error': customer_error}), 422

    address_error = validate_delivery_address(customer_name, delivery_address)
    if address_error:
        return jsonify({'error': address_error}), 422

    if delivery_mode == 'manual':
        is_valid, err, parsed_date = validate_manual_delivery_date(delivery_date_manual)
        if not is_valid:
            return jsonify({'error': err}), 422
        delivery_date_final = format_delivery_date(parsed_date)
    else:
        delivery_date_final = format_delivery_date(get_default_delivery_date())

    pdf_bytes_raw = pdf_file.read()

    try:
        header_info, line_items, raw_df, summary_df = process_pdf_pp(io.BytesIO(pdf_bytes_raw))
    except Exception as e:
        return jsonify({'error': f'PDF পড়তে সমস্যা হয়েছে: {str(e)}'}), 422

    if not line_items:
        return jsonify({'error': 'কোনো লাইন-আইটেম পাওয়া যায়নি এই PDF থেকে'}), 422

    header_info['buyer'] = resolve_alias(header_info.get('buyer', ''), PRINTING_PRESS_BUYER_ALIASES)
    header_info['customer'] = resolve_alias(header_info.get('customer', ''), CUSTOMER_ALIASES)
    header_info['item_name'] = resolve_alias(header_info.get('item_name', ''), PRINTING_PRESS_ITEM_NAME_ALIASES)
    if item_name_override:
        item_name_override = resolve_alias(item_name_override, PRINTING_PRESS_ITEM_NAME_ALIASES)

    po_error = validate_po_number(po_number_override, header_info.get('po_number', ''))
    if po_error:
        return jsonify({'error': po_error}), 422

    customer_pdf_error = None if force_override else validate_matches_pdf(
        'Customer', customer_name, header_info.get('customer', ''))
    if customer_pdf_error:
        return jsonify({'error': customer_pdf_error}), 422

    buyer_pdf_error = None if force_override else validate_matches_pdf(
        'Buyer', buyer_name, header_info.get('buyer', ''))
    if buyer_pdf_error:
        return jsonify({'error': buyer_pdf_error}), 422

    warnings = validate_pp_line_items(line_items)

    # buyer সিস্টেমে (মাস্টার লিস্টে) থাকলেই যথেষ্ট প্রসেসিং চালানোর জন্য —
    # কিন্তু এই buyer-এর Printing Press PDF ফরম্যাট এখনো নির্দিষ্টভাবে যাচাই
    # করা না থাকলে ব্লক না করে শুধু একটা সতর্কতা যোগ করা হচ্ছে।
    if buyer_name not in PRINTING_PRESS_VERIFIED_BUYERS:
        warnings.append(
            f"⚠️ '{buyer_name}' buyer-এর Printing Press PDF ফরম্যাট এখনো নির্দিষ্টভাবে "
            f"যাচাই করা হয়নি — আউটপুট (বিশেষ করে সাইজ/কোয়ান্টিটি/রেফারেন্স) "
            f"ভালোভাবে চেক করে নিন।"
        )

    if force_override:
        pdf_customer = header_info.get('customer', '')
        pdf_buyer = header_info.get('buyer', '')
        if pdf_customer and not values_match_ci(customer_name, pdf_customer):
            warnings.append(
                f"⚠️ FORCE OVERRIDE: Customer ম্যানুয়ালি '{customer_name}' বসানো হয়েছে, "
                f"কিন্তু PDF-এ ছিল '{pdf_customer}' — দয়া করে যাচাই করুন।"
            )
        if pdf_buyer and not values_match_ci(buyer_name, pdf_buyer):
            warnings.append(
                f"⚠️ FORCE OVERRIDE: Buyer ম্যানুয়ালি '{buyer_name}' বসানো হয়েছে, "
                f"কিন্তু PDF-এ ছিল '{pdf_buyer}' — দয়া করে যাচাই করুন।"
            )

    if not length or not width:
        warnings.append("Length/Width ফাঁকা রাখা হয়েছে (ইউজার confirm করেছেন) — পরে ম্যানুয়ালি বসাতে হবে।")

    if not item_name_override and not header_info.get('item_name'):
        warnings.append("Item Name PDF থেকে বের করা যায়নি এবং ম্যানুয়ালিও দেওয়া হয়নি — 'N/A' বসানো হয়েছে, চেক করুন।")

    base_name = os.path.splitext(pdf_file.filename)[0]

    def to_num_or_blank(v):
        try:
            return float(v) if v else ''
        except ValueError:
            return ''

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, f'{base_name}_Output.xlsx')
            build_pp_excel(
                line_items, header_info, out_path,
                customer_override=customer_name or None,
                buyer_override=buyer_name or None,
                po_override=po_number_override or None,
                item_name_override=item_name_override or None,
                delivery_date=delivery_date_final,
                delivery_address=delivery_address,
                length=to_num_or_blank(length),
                width=to_num_or_blank(width),
                height=to_num_or_blank(height) if height else 0,
                measurement_unit=measurement_unit,
                raw_df=raw_df,
                summary_df=summary_df,
                warnings=warnings,
                remark_place=remark_place,
                remark_address=remark_address,
            )
            with open(out_path, 'rb') as f:
                file_bytes = f.read()
    except Exception as e:
        return jsonify({'error': f'Excel ফাইল বানাতে সমস্যা হয়েছে: {str(e)}'}), 500

    if not file_bytes:
        return jsonify({'error': 'Excel ফাইল খালি তৈরি হয়েছে — আবার চেষ্টা করুন'}), 500

    buf = io.BytesIO(file_bytes)
    response = send_file(
        buf,
        as_attachment=True,
        download_name=f'{base_name}_Output.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response.headers['Content-Length'] = str(len(file_bytes))
    response.headers['X-Warning-Count'] = str(len(warnings))
    return response


@app.route('/autocarton/process_outhouse_excel', methods=['POST'])
def autocarton_process_outhouse_excel():
    """আউট হাউজ Carton — একাধিক বুকিং এক্সেল (.xls/.xlsx) একসাথে আপলোড করে
    একটাই কম্বাইনড Excel টেমপ্লেট বানায়। Customer/Buyer/PO এখানে ম্যানুয়ালি
    ইনপুট দিতে হয় (এক্সেলে এসব হেডার-লেভেল তথ্য PDF-এর মতো পরিষ্কারভাবে
    থাকে না), শুধু লাইন-আইটেমগুলো ফাইল থেকে বের করে কম্বাইন করা হয়।

    এই একই রুট Amigo Bangladesh Ltd (Uniqlo)-এর জন্যও ব্যবহার হয় — UI-তে
    আলাদা কোনো নতুন অপশন যোগ করা হয়নি, Customer/Buyer ড্রপডাউনে Amigo/Uniqlo
    সিলেক্ট করলেই combine_booking_excels ভেতরে ভেতরে সঠিক (batch-মোড)
    extractor-এ রুট করে দেয় (দেখুন outhouse_extractor.py-এর BATCH_REGISTRY)।
    """
    files = request.files.getlist('files')
    files = [f for f in files if f and f.filename]
    if not files:
        return jsonify({'error': 'অন্তত একটা এক্সেল ফাইল আপলোড করুন'}), 400

    customer_name = request.form.get('customer_name', '').strip()
    buyer_name = request.form.get('buyer_name', '').strip()
    po_number_override = request.form.get('po_number', '').strip()
    item_name_override = request.form.get('item_name', '').strip() or 'Master Carton'
    manual_ply = request.form.get('ply', '').strip()
    delivery_mode = request.form.get('delivery_mode', 'auto').strip()
    delivery_date_manual = request.form.get('delivery_date', '').strip()
    delivery_address = request.form.get('delivery_address', '').strip()

    customer_error = validate_customer('OUT-HOUSE', customer_name)
    if customer_error:
        return jsonify({'error': customer_error}), 422

    buyer_error = validate_buyer_in_list(buyer_name, BUYERS)
    if buyer_error:
        return jsonify({'error': buyer_error}), 422

    address_error = validate_delivery_address(customer_name, delivery_address)
    if address_error:
        return jsonify({'error': address_error}), 422

    if delivery_mode == 'manual':
        is_valid, err, parsed_date = validate_manual_delivery_date(delivery_date_manual)
        if not is_valid:
            return jsonify({'error': err}), 422
        delivery_date_final = format_delivery_date(parsed_date)
    else:
        delivery_date_final = format_delivery_date(get_default_delivery_date())

    file_tuples = [(io.BytesIO(f.read()), f.filename) for f in files]
    try:
        line_items, file_errors = combine_booking_excels(
            file_tuples, item_name_override=item_name_override, manual_ply=manual_ply,
            buyer_name=buyer_name, customer_name=customer_name)
    except Exception as e:
        return jsonify({'error': f'এক্সেল ফাইল পড়তে সমস্যা হয়েছে: {str(e)}'}), 422

    if not line_items and file_errors and all('(লাইব্রেরি মিসিং)' in e for e in file_errors):
        return jsonify({
            'error': 'সার্ভারে .xls/.xlsx পড়ার জন্য দরকারি লাইব্রেরি (xlrd/python-calamine) '
                     'ইনস্টল করা নেই। Terminal-এ গিয়ে "pip install -r requirements.txt" '
                     'চালিয়ে সার্ভার আবার রিস্টার্ট করুন।'
        }), 422

    if not line_items:
        msg = 'কোনো লাইন-আইটেম পাওয়া যায়নি।'
        if file_errors:
            msg += ' সমস্যা: ' + '; '.join(file_errors)
        return jsonify({'error': msg}), 422

    warnings = validate_line_items(line_items)
    for e in file_errors:
        if e.strip().startswith('⚠️'):
            warnings.append(e)
        else:
            warnings.append(f"⚠️ এই ফাইলটা স্কিপ হয়েছে: {e}")

    if buyer_name not in CARTON_VERIFIED_BUYERS:
        warnings.append(
            f"⚠️ '{buyer_name}' buyer-এর OUT-HOUSE Excel ফরম্যাট এখনো নির্দিষ্টভাবে "
            f"যাচাই করা হয়নি — আউটপুট ভালোভাবে চেক করে নিন।"
        )
        
    separate_output = request.form.get('separate_output', '').strip().lower() in ('1', 'true', 'on', 'yes')

    if not po_number_override and not separate_output:
        source_files = {it.get('_source_file') for it in line_items if it.get('_source_file')}
        if len(source_files) <= 1:
            extracted_po_numbers = sorted({
                str(it.get('po_no', '')).strip() for it in line_items
                if str(it.get('po_no', '')).strip()
            })
            if len(extracted_po_numbers) == 1:
                po_number_override = extracted_po_numbers[0]
        else:
            warnings.append(
                "⚠️ একাধিক ফাইল থেকে ভিন্ন ভিন্ন PO NO/Ship To পাওয়া গেছে — একটাই কম্বাইনড Excel-এর "
                "হেডারে এগুলো মেশানো ঠিক না, তাই PO Number ফাঁকা/N/A রাখা হয়েছে (প্রতিটা রো-তে অবশ্য "
                "নিজের সঠিক PO NO/Ship To ঠিকই আছে)। আলাদা আলাদা PO প্রতিটা ফাইলে চাইলে 'প্রতিটা ফাইল "
                "আলাদা Excel হিসেবে ডাউনলোড করুন' চেকবক্সটা ব্যবহার করুন।"
            )

    header_info = {
        'po_number': po_number_override or '',
        'customer': customer_name,
        'buyer': buyer_name,
    }
    

    if separate_output:
        groups, order = {}, []
        for item in line_items:
            key = item.get('_source_file') or 'Unknown'
            if key not in groups:
                groups[key] = []
                order.append(key)
            groups[key].append(item)

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                import zipfile
                zip_path = os.path.join(tmpdir, 'AutoCarton_Outputs.zip')
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for src_name in order:
                        group_items = groups[src_name]
                        group_warnings = validate_line_items(group_items)
                        if buyer_name not in CARTON_VERIFIED_BUYERS:
                            group_warnings.append(
                                f"⚠️ '{buyer_name}' buyer-এর OUT-HOUSE Excel ফরম্যাট এখনো "
                                f"নির্দিষ্টভাবে যাচাই করা হয়নি — আউটপুট ভালোভাবে চেক করে নিন।"
                            )

                        group_po = po_number_override
                        if not group_po:
                            group_po_numbers = sorted({
                                str(it.get('po_no', '')).strip() for it in group_items
                                if str(it.get('po_no', '')).strip()
                            })
                            if len(group_po_numbers) == 1:
                                group_po = group_po_numbers[0]

                        out_name = re.sub(
                            r'[\\/:*?"<>|]', '-',
                            f"{os.path.splitext(src_name)[0]}_Output.xlsx"
                        )
                        out_path = os.path.join(tmpdir, out_name)
                        build_combined_excel(
                            group_items, header_info, out_path, profile='OUT-HOUSE',
                            customer_override=customer_name or None,
                            buyer_override=buyer_name or None,
                            po_override=group_po or None,
                            delivery_date=delivery_date_final,
                            delivery_address=delivery_address,
                            warnings=group_warnings,
                        )
                        zf.write(out_path, arcname=out_name)
                with open(zip_path, 'rb') as f:
                    zip_bytes = f.read()
        except Exception as e:
            return jsonify({'error': f'Zip ফাইল বানাতে সমস্যা হয়েছে: {str(e)}'}), 500

        buf = io.BytesIO(zip_bytes)
        response = send_file(
            buf, as_attachment=True, download_name='AutoCarton_Outputs.zip',
            mimetype='application/zip',
        )
        response.headers['Content-Length'] = str(len(zip_bytes))
        response.headers['X-File-Count'] = str(len(files))
        return response

    combined_label = '_'.join(sorted({str(it.get('po_no', '')) for it in line_items if it.get('po_no')}))[:60]
    base_name = f"{customer_name}_{buyer_name}_{combined_label}_OUTHOUSE".replace(' ', '_')
    base_name = re.sub(r'[\\/:*?"<>|]', '-', base_name)

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, f'{base_name}_Output.xlsx')
            build_combined_excel(
                line_items, header_info, out_path, profile='OUT-HOUSE',
                customer_override=customer_name or None,
                buyer_override=buyer_name or None,
                po_override=po_number_override or None,
                delivery_date=delivery_date_final,
                delivery_address=delivery_address,
                warnings=warnings,
                full_dump=[build_excel_full_dump(fs, fn) for fs, fn in file_tuples],
            )
            with open(out_path, 'rb') as f:
                file_bytes = f.read()
    except Exception as e:
        return jsonify({'error': f'Excel ফাইল বানাতে সমস্যা হয়েছে: {str(e)}'}), 500

    if not file_bytes:
        return jsonify({'error': 'Excel ফাইল খালি তৈরি হয়েছে — আবার চেষ্টা করুন'}), 500

    buf = io.BytesIO(file_bytes)
    response = send_file(
        buf,
        as_attachment=True,
        download_name=f'{base_name}_Output.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response.headers['Content-Length'] = str(len(file_bytes))
    response.headers['X-Warning-Count'] = str(len(warnings))
    response.headers['X-File-Count'] = str(len(files))
    return response


@app.route('/autocarton/extract_header_outhouse_pdf', methods=['POST'])
def autocarton_extract_header_outhouse_pdf():
    """OUT-HOUSE PDF (Trims Booking ফরম্যাট) আপলোড হওয়ার সাথে সাথেই Booking No
    (-> PO Number), Buyer, Customer বের করে ফেরত দেয় — IN-HOUSE PDF ফ্লো-র
    /extract_header endpoint-এর মতোই, autofill-এর জন্য।

    এই মুহূর্তে customer/buyer এখনো UI থেকে জানা নেই (এটা upload-এর সাথে
    সাথেই, submit-এর আগে কল হয়) — তাই ফরম্যাট বুঝে নেওয়া হয় try-চেইন দিয়ে
    (_extract_outhouse_pdf_header_auto): আগে ডিফল্ট Barnali/Modele ফরম্যাট,
    না মিললে রেজিস্ট্রিতে থাকা অন্য ফরম্যাট (যেমন IKL/Biscana)।
    """
    if 'pdf_file' not in request.files:
        return jsonify({'error': 'PDF ফাইল পাওয়া যায়নি'}), 400

    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return jsonify({'error': 'ফাইল সিলেক্ট করা হয়নি'}), 400

    try:
        header_info, _items = _extract_outhouse_pdf_header_auto(
            io.BytesIO(pdf_file.read()), pdf_file.filename)
    except Exception as e:
        return jsonify({'error': f'PDF থেকে তথ্য বের করতে সমস্যা হয়েছে: {str(e)}'}), 422

    header_info['buyer'] = resolve_alias(header_info.get('buyer', ''), BUYER_ALIASES)
    header_info['customer'] = resolve_alias(header_info.get('customer', ''), CUSTOMER_ALIASES)

    return jsonify({
        'po_number': header_info.get('po_number', '') or '',
        'customer': header_info.get('customer', '') or '',
        'buyer': header_info.get('buyer', '') or '',
    })


@app.route('/autocarton/process_outhouse_trims_booking_pdf', methods=['POST'])
def autocarton_process_outhouse_trims_booking_pdf():
    """আউট হাউজ Carton — 'Multiple Job Wise Trims Booking' PDF ফরম্যাট
    (যেমন Barnali Textile-এর বুকিং শিট), এবং একই radio অপশনের ভেতরেই
    customer+buyer অনুযায়ী রেজিস্ট্রি-ডিসপ্যাচ হওয়া অন্য "টোটালি ডিফারেন্ট"
    PDF ফরম্যাটও (যেমন Innovative Knitex Ltd./Biscana — দেখুন
    OUTHOUSE_PDF_REGISTRY)। এই ধাপে customer_name/buyer_name ফর্ম থেকে
    নিশ্চিতভাবে জানা যায়, তাই এখানে সরাসরি রেজিস্ট্রি-লুকআপ নির্ভরযোগ্য।

    Item Name/Ply Barnali-স্টাইল ফরম্যাটে PDF থেকেই automatic ঠিক হয়ে যায়
    (তাই ফর্ম-ফিল্ড না দিলেও চলে), কিন্তু IKL/Biscana-স্টাইল ফরম্যাটে UI
    থেকে সিলেক্ট করা Item Name/Ply দরকার হয় — তাই এই দুটো ফর্ম-ফিল্ডও এখন
    (ঐচ্ছিকভাবে) পড়া হয়, আর handler নিজে বেছে নেয় সেগুলো ব্যবহার করবে কিনা।

    'separate_output' চেকমার্ক করা থাকলে প্রতিটা PDF-এর জন্য আলাদা Excel
    (Zip-এ), না থাকলে (ডিফল্ট) আগের মতোই একটাই কম্বাইনড Excel।
    """
    files = request.files.getlist('files')
    files = [f for f in files if f and f.filename]
    if not files:
        return jsonify({'error': 'অন্তত একটা PDF ফাইল আপলোড করুন'}), 400

    customer_name = request.form.get('customer_name', '').strip()
    buyer_name = request.form.get('buyer_name', '').strip()
    po_number_override = request.form.get('po_number', '').strip()
    delivery_mode = request.form.get('delivery_mode', 'auto').strip()
    delivery_date_manual = request.form.get('delivery_date', '').strip()
    delivery_address = request.form.get('delivery_address', '').strip()
    primark_weight_class = request.form.get('primark_weight_class', '').strip()
    item_name_override = request.form.get('item_name', '').strip()
    manual_ply = request.form.get('ply', '').strip()
    separate_output = request.form.get('separate_output', '').strip().lower() in ('1', 'true', 'on', 'yes')

    customer_error = validate_customer('OUT-HOUSE', customer_name)
    if customer_error:
        return jsonify({'error': customer_error}), 422

    buyer_error = validate_buyer_in_list(buyer_name, BUYERS)
    if buyer_error:
        return jsonify({'error': buyer_error}), 422

    is_primark = buyer_name.strip().lower() == 'primark'
    if is_primark and primark_weight_class not in ('ABOVE 10KG', 'BELOW 10KG'):
        return jsonify({'error': "Primark buyer-এর জন্য 'ABOVE 10KG' বা 'BELOW 10KG' সিলেক্ট করা আবশ্যক।"}), 422

    address_error = validate_delivery_address(customer_name, delivery_address)
    if address_error:
        return jsonify({'error': address_error}), 422

    if delivery_mode == 'manual':
        is_valid, err, parsed_date = validate_manual_delivery_date(delivery_date_manual)
        if not is_valid:
            return jsonify({'error': err}), 422
        delivery_date_final = format_delivery_date(parsed_date)
    else:
        delivery_date_final = format_delivery_date(get_default_delivery_date())

    handler = _get_outhouse_pdf_handler(customer_name, buyer_name)
    customer_list = CUSTOMERS.get('OUT-HOUSE', [])

    per_file_results = []  # [(filename, header_info, items, raw_bytes), ...]
    file_errors = []
    for f in files:
        raw_bytes = f.read()
        try:
            _hdr, items = handler(
                io.BytesIO(raw_bytes), f.filename, customer_list, BUYERS,
                item_name_override, manual_ply)
            if not items:
                file_errors.append(f"{f.filename}: কোনো লাইন-আইটেম পাওয়া যায়নি (পরিচিত ফরম্যাট না হতে পারে)")
                continue
            per_file_results.append((f.filename, _hdr, items, raw_bytes))
        except Exception as e:
            file_errors.append(f"{f.filename}: {str(e)}")

    if not per_file_results:
        msg = 'কোনো লাইন-আইটেম পাওয়া যায়নি।'
        if file_errors:
            msg += ' সমস্যা: ' + '; '.join(file_errors)
        return jsonify({'error': msg}), 422

    def _apply_primark(items):
        if is_primark:
            for item in items:
                item['ply'] = '3'
                if item.get('style_no'):
                    item['style_no'] = f"{item['style_no']}/{primark_weight_class}"

    verified_warning = None
    if buyer_name not in CARTON_VERIFIED_BUYERS:
        verified_warning = (
            f"⚠️ '{buyer_name}' buyer-এর এই OUT-HOUSE PDF ফরম্যাট এখনো নির্দিষ্টভাবে "
            f"যাচাই করা হয়নি — আউটপুট ভালোভাবে চেক করে নিন।"
        )

    header_info = {
        'po_number': po_number_override or '',
        'customer': customer_name,
        'buyer': buyer_name,
    }

    # ============================================================
    # SEPARATE OUTPUT (ZIP) মোড
    # ============================================================
    if separate_output:
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                import zipfile
                zip_path = os.path.join(tmpdir, 'AutoCarton_Outputs.zip')
                total_warn_count = 0
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for filename, hdr, items, raw_bytes in per_file_results:
                        _apply_primark(items)
                        group_warnings = validate_line_items(items)
                        if verified_warning:
                            group_warnings.append(verified_warning)
                        total_warn_count += len(group_warnings)

                        group_po = po_number_override or hdr.get('po_number') or ''
                        out_name = _safe_filename_part(
                            f"{os.path.splitext(filename)[0]}_Output.xlsx"
                        )
                        out_path = os.path.join(tmpdir, out_name)
                        group_header_info = {
                            'po_number': group_po,
                            'customer': customer_name,
                            'buyer': buyer_name,
                        }
                        build_combined_excel(
                            items, group_header_info, out_path, profile='OUT-HOUSE',
                            customer_override=customer_name or None,
                            buyer_override=buyer_name or None,
                            po_override=group_po or None,
                            delivery_date=delivery_date_final,
                            delivery_address=delivery_address,
                            warnings=group_warnings,
                            full_dump=[build_pdf_full_dump(io.BytesIO(raw_bytes), filename)],
                        )
                        zf.write(out_path, arcname=out_name)
                with open(zip_path, 'rb') as f:
                    zip_bytes = f.read()
        except Exception as e:
            return jsonify({'error': f'Zip ফাইল বানাতে সমস্যা হয়েছে: {str(e)}'}), 500

        buf = io.BytesIO(zip_bytes)
        response = send_file(
            buf, as_attachment=True, download_name='AutoCarton_Outputs.zip',
            mimetype='application/zip',
        )
        response.headers['Content-Length'] = str(len(zip_bytes))
        response.headers['X-Warning-Count'] = str(total_warn_count)
        response.headers['X-File-Count'] = str(len(files))
        return response

    # ============================================================
    # কম্বাইনড মোড (ডিফল্ট)
    # ============================================================
    line_items = []
    full_dump = []
    for filename, hdr, items, raw_bytes in per_file_results:
        line_items.extend(items)
        full_dump.append(build_pdf_full_dump(io.BytesIO(raw_bytes), filename))

    _apply_primark(line_items)

    warnings = validate_line_items(line_items)
    for e in file_errors:
        warnings.append(f"⚠️ এই ফাইলটা স্কিপ হয়েছে: {e}")
    if verified_warning:
        warnings.append(verified_warning)

    if not po_number_override:
        header_pos = [hdr.get('po_number', '') for _fn, hdr, _it, _rb in per_file_results if hdr.get('po_number')]
        if header_pos:
            po_number_override = header_pos[0]
    header_info['po_number'] = po_number_override or ''

    combined_label = '_'.join(sorted({str(it.get('style_no', '')) for it in line_items if it.get('style_no')}))[:60]
    base_name = f"{customer_name}_{buyer_name}_{combined_label}_OUTHOUSE".replace(' ', '_')
    base_name = _safe_filename_part(base_name)

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, f'{base_name}_Output.xlsx')
            build_combined_excel(
                line_items, header_info, out_path, profile='OUT-HOUSE',
                customer_override=customer_name or None,
                buyer_override=buyer_name or None,
                po_override=po_number_override or None,
                delivery_date=delivery_date_final,
                delivery_address=delivery_address,
                warnings=warnings,
                full_dump=full_dump,
            )
            with open(out_path, 'rb') as f:
                file_bytes = f.read()
    except Exception as e:
        return jsonify({'error': f'Excel ফাইল বানাতে সমস্যা হয়েছে: {str(e)}'}), 500

    if not file_bytes:
        return jsonify({'error': 'Excel ফাইল খালি তৈরি হয়েছে — আবার চেষ্টা করুন'}), 500

    buf = io.BytesIO(file_bytes)
    response = send_file(
        buf,
        as_attachment=True,
        download_name=f'{base_name}_Output.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response.headers['Content-Length'] = str(len(file_bytes))
    response.headers['X-Warning-Count'] = str(len(warnings))
    response.headers['X-File-Count'] = str(len(files))
    return response


@app.route('/autocarton/extract_header_kenpark_pdf', methods=['POST'])
def autocarton_extract_header_kenpark_pdf():
    """Kenpark Bangladesh Apparel (Pvt.) Limited / Kenpark Bangladesh (Pvt.)
    Limited-এর 'PURCHASE ORDER (LOCAL FE)' PDF (Buyer: Ralph Lauren) আপলোড
    হওয়ার সাথে সাথেই PO Number/Customer/Buyer বের করে autofill-এর জন্য
    ফেরত দেয়।

    ⚠️ এই ফরম্যাট শুধু ডিজিটালি-জেনারেট করা (সিলেক্টেবল টেক্সট আছে) PDF-এর
    জন্য কাজ করে।
    """
    if 'pdf_file' not in request.files:
        return jsonify({'error': 'PDF ফাইল পাওয়া যায়নি'}), 400

    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return jsonify({'error': 'ফাইল সিলেক্ট করা হয়নি'}), 400

    try:
        header_info, _items = read_kenpark_pdf(io.BytesIO(pdf_file.read()), pdf_file.filename)
    except Exception as e:
        return jsonify({'error': f'PDF থেকে তথ্য বের করতে সমস্যা হয়েছে: {str(e)}'}), 422

    header_info['buyer'] = resolve_alias(header_info.get('buyer', ''), BUYER_ALIASES)
    header_info['customer'] = resolve_alias(header_info.get('customer', ''), CUSTOMER_ALIASES)

    return jsonify({
        'po_number': header_info.get('po_number', '') or '',
        'customer': header_info.get('customer', '') or '',
        'buyer': header_info.get('buyer', '') or '',
    })


@app.route('/autocarton/process_kenpark_pdf', methods=['POST'])
def autocarton_process_kenpark_pdf():
    """Kenpark Bangladesh Apparel (Pvt.) Limited / Kenpark Bangladesh (Pvt.)
    Limited — Buyer: Ralph Lauren — 'PURCHASE ORDER (LOCAL FE)' PDF ফরম্যাট।
    'separate_output' চেকমার্ক করা থাকলে প্রতিটা PDF-এর জন্য আলাদা Excel
    (Zip-এ), না থাকলে (ডিফল্ট) আগের মতোই একটাই কম্বাইনড Excel।

    ⚠️ শুধু ডিজিটালি-জেনারেট করা (স্ক্যান না) PDF-এর জন্য কাজ করে।
    """
    files = request.files.getlist('files')
    files = [f for f in files if f and f.filename]
    if not files:
        return jsonify({'error': 'অন্তত একটা PDF ফাইল আপলোড করুন'}), 400

    customer_name = request.form.get('customer_name', '').strip()
    buyer_name = request.form.get('buyer_name', '').strip()
    po_number_override = request.form.get('po_number', '').strip()
    delivery_mode = request.form.get('delivery_mode', 'auto').strip()
    delivery_date_manual = request.form.get('delivery_date', '').strip()
    delivery_address = request.form.get('delivery_address', '').strip()
    separate_output = request.form.get('separate_output', '').strip().lower() in ('1', 'true', 'on', 'yes')

    customer_error = validate_customer('OUT-HOUSE', customer_name)
    if customer_error:
        return jsonify({'error': customer_error}), 422

    buyer_error = validate_buyer_in_list(buyer_name, BUYERS)
    if buyer_error:
        return jsonify({'error': buyer_error}), 422

    address_error = validate_delivery_address(customer_name, delivery_address)
    if address_error:
        return jsonify({'error': address_error}), 422

    if delivery_mode == 'manual':
        is_valid, err, parsed_date = validate_manual_delivery_date(delivery_date_manual)
        if not is_valid:
            return jsonify({'error': err}), 422
        delivery_date_final = format_delivery_date(parsed_date)
    else:
        delivery_date_final = format_delivery_date(get_default_delivery_date())

    per_file_results = []  # [(filename, items, raw_bytes), ...]
    file_errors = []
    for f in files:
        raw_bytes = f.read()
        try:
            _hdr, items = read_kenpark_pdf(io.BytesIO(raw_bytes), f.filename)
            if not items:
                file_errors.append(
                    f"{f.filename}: কোনো Carton/Divider লাইন-আইটেম পাওয়া যায়নি "
                    f"(স্ক্যান করা/ছবি-PDF হতে পারে, বা পরিচিত ফরম্যাট না)"
                )
                continue
            per_file_results.append((f.filename, items, raw_bytes))
        except Exception as e:
            file_errors.append(f"{f.filename}: {str(e)}")

    if not per_file_results:
        msg = 'কোনো লাইন-আইটেম পাওয়া যায়নি।'
        if file_errors:
            msg += ' সমস্যা: ' + '; '.join(file_errors)
        return jsonify({'error': msg}), 422

    verified_warning = None
    if buyer_name not in CARTON_VERIFIED_BUYERS:
        verified_warning = (
            f"⚠️ '{buyer_name}' buyer-এর এই Kenpark PDF ফরম্যাট এখনো নির্দিষ্টভাবে "
            f"যাচাই করা হয়নি — আউটপুট ভালোভাবে চেক করে নিন।"
        )

    header_info = {
        'po_number': po_number_override or '',
        'customer': customer_name,
        'buyer': buyer_name,
    }

    # ============================================================
    # SEPARATE OUTPUT (ZIP) মোড
    # ============================================================
    if separate_output:
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                import zipfile
                zip_path = os.path.join(tmpdir, 'AutoCarton_Outputs.zip')
                total_warn_count = 0
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for filename, items, raw_bytes in per_file_results:
                        group_warnings = validate_line_items(items)
                        if verified_warning:
                            group_warnings.append(verified_warning)
                        total_warn_count += len(group_warnings)

                        group_po = po_number_override
                        if not group_po:
                            group_po_numbers = sorted({str(it.get('po_no', '')).strip() for it in items if str(it.get('po_no', '')).strip()})
                            if len(group_po_numbers) == 1:
                                group_po = group_po_numbers[0]

                        out_name = _safe_filename_part(f"{os.path.splitext(filename)[0]}_Output.xlsx")
                        out_path = os.path.join(tmpdir, out_name)
                        group_header_info = {'po_number': group_po or '', 'customer': customer_name, 'buyer': buyer_name}
                        build_combined_excel(
                            items, group_header_info, out_path, profile='OUT-HOUSE',
                            customer_override=customer_name or None,
                            buyer_override=buyer_name or None,
                            po_override=group_po or None,
                            delivery_date=delivery_date_final,
                            delivery_address=delivery_address,
                            warnings=group_warnings,
                            full_dump=[build_pdf_full_dump(io.BytesIO(raw_bytes), filename)],
                        )
                        zf.write(out_path, arcname=out_name)
                with open(zip_path, 'rb') as f:
                    zip_bytes = f.read()
        except Exception as e:
            return jsonify({'error': f'Zip ফাইল বানাতে সমস্যা হয়েছে: {str(e)}'}), 500

        buf = io.BytesIO(zip_bytes)
        response = send_file(
            buf, as_attachment=True, download_name='AutoCarton_Outputs.zip',
            mimetype='application/zip',
        )
        response.headers['Content-Length'] = str(len(zip_bytes))
        response.headers['X-Warning-Count'] = str(total_warn_count)
        response.headers['X-File-Count'] = str(len(files))
        return response

    # ============================================================
    # কম্বাইনড মোড (ডিফল্ট)
    # ============================================================
    line_items = []
    file_bytes_list = []
    for filename, items, raw_bytes in per_file_results:
        line_items.extend(items)
        file_bytes_list.append((raw_bytes, filename))

    warnings = validate_line_items(line_items)
    for e in file_errors:
        warnings.append(f"⚠️ এই ফাইলটা স্কিপ হয়েছে: {e}")
    if verified_warning:
        warnings.append(verified_warning)

    combined_label = '_'.join(sorted({str(it.get('po_no', '')) for it in line_items if it.get('po_no')}))[:60]
    base_name = f"{customer_name}_{buyer_name}_{combined_label}_KENPARK".replace(' ', '_')
    base_name = _safe_filename_part(base_name)

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, f'{base_name}_Output.xlsx')
            build_combined_excel(
                line_items, header_info, out_path, profile='OUT-HOUSE',
                customer_override=customer_name or None,
                buyer_override=buyer_name or None,
                po_override=po_number_override or None,
                delivery_date=delivery_date_final,
                delivery_address=delivery_address,
                warnings=warnings,
                full_dump=[build_pdf_full_dump(io.BytesIO(b), fn) for b, fn in file_bytes_list],
            )
            with open(out_path, 'rb') as f:
                file_bytes = f.read()
    except Exception as e:
        return jsonify({'error': f'Excel ফাইল বানাতে সমস্যা হয়েছে: {str(e)}'}), 500

    if not file_bytes:
        return jsonify({'error': 'Excel ফাইল খালি তৈরি হয়েছে — আবার চেষ্টা করুন'}), 500

    buf = io.BytesIO(file_bytes)
    response = send_file(
        buf,
        as_attachment=True,
        download_name=f'{base_name}_Output.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response.headers['Content-Length'] = str(len(file_bytes))
    response.headers['X-Warning-Count'] = str(len(warnings))
    response.headers['X-File-Count'] = str(len(files))
    return response


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)