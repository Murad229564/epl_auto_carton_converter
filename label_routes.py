"""
Label মডিউলের সব Flask রুট — একটা Blueprint হিসেবে, যাতে app.py-তে শুধু ২ লাইন
লাগে (import + register) এবং রুটগুলো ভুল জায়গায় (যেমন app.run-এর নিচে) পড়ে
নিষ্ক্রিয় হয়ে যাওয়ার ঝুঁকি না থাকে।

app.py-তে যা করবেন:
  ১) উপরের import-গুলোর সাথে:      from label_routes import label_bp
  ২) app = Flask(__name__) এবং app.config[...] লাইনের ঠিক নিচে:   app.register_blueprint(label_bp)
  (আগে app.py-তে label-এর যেসব import/রুট হাতে বসিয়েছিলেন, সেগুলো মুছে দিন — নাহলে
   '/label' দুইবার রেজিস্টার হয়ে দ্বন্দ্ব হবে।)

ডায়াগনস্টিক: ব্রাউজারে /label/health খুললে দেখা যাবে রুট/টেমপ্লেট ঠিকমতো লোড হয়েছে কিনা।
"""
import io
import os
import tempfile
import traceback

from flask import Blueprint, request, render_template, send_file, jsonify, current_app

from label_extractor import process_pdf_label, get_unique_delivery_info as get_unique_delivery_info_label
from label_builder import build_label_excel, validate_label_line_items, TEMPLATE_PATH as LABEL_TEMPLATE_PATH
from label_config import LABEL_BUYERS, LABEL_BUYER_ALIASES, LABEL_VERIFIED_BUYERS
from validators import (
    validate_customer, validate_buyer_in_list, validate_po_number,
    validate_delivery_address, validate_matches_pdf, values_match_ci,
)
from date_logic import get_default_delivery_date, validate_manual_delivery_date, format_delivery_date
from config import CUSTOMERS, DELIVERY_ADDRESSES, CUSTOMER_ALIASES, resolve_alias

label_bp = Blueprint('label', __name__)


def _err(message, status=422):
    return jsonify({'error': message}), status


@label_bp.route('/label')
def label_index():
    return render_template(
        'label.html',
        customers=CUSTOMERS,
        buyers=LABEL_BUYERS,
        delivery_addresses=DELIVERY_ADDRESSES,
    )


@label_bp.route('/label/health')
def label_health():
    """রুট/টেমপ্লেট ঠিক আছে কিনা এক নজরে দেখার জন্য (ব্রাউজারে /label/health)।"""
    return jsonify({
        'ok': True,
        'template_label_xlsx_exists': os.path.exists(LABEL_TEMPLATE_PATH),
        'template_path': LABEL_TEMPLATE_PATH,
        'label_buyers_count': len(LABEL_BUYERS),
    })


@label_bp.route('/label/extract_header', methods=['POST'])
def label_extract_header():
    """PDF আপলোডের সাথে সাথেই PO Number/Customer/Buyer + Delivery Place/Address হিন্ট।"""
    try:
        if 'pdf_file' not in request.files:
            return _err('PDF ফাইল পাওয়া যায়নি', 400)
        pdf_file = request.files['pdf_file']
        if pdf_file.filename == '':
            return _err('ফাইল সিলেক্ট করা হয়নি', 400)

        try:
            header_info, line_items, raw_df, summary_df = process_pdf_label(io.BytesIO(pdf_file.read()))
        except Exception as e:
            current_app.logger.exception('label extract_header failed')
            return _err(f'PDF থেকে তথ্য বের করতে সমস্যা হয়েছে: {type(e).__name__}: {e}')

        header_info['buyer'] = resolve_alias(header_info.get('buyer', ''), LABEL_BUYER_ALIASES)
        header_info['customer'] = resolve_alias(header_info.get('customer', ''), CUSTOMER_ALIASES)
        delivery_info = get_unique_delivery_info_label(raw_df)

        return jsonify({
            'po_number': header_info.get('po_number', '') or '',
            'customer': header_info.get('customer', '') or '',
            'buyer': header_info.get('buyer', '') or '',
            'delivery_places_pdf': delivery_info['delivery_places'],
            'delivery_addresses_pdf': delivery_info['delivery_addresses'],
        })
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return _err(f'সার্ভারে অপ্রত্যাশিত সমস্যা: {type(e).__name__}: {e}', 500)


@label_bp.route('/label/process', methods=['POST'])
def label_process():
    try:
        return _label_process_impl()
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return _err(f'সার্ভারে অপ্রত্যাশিত সমস্যা: {type(e).__name__}: {e}', 500)


def _label_process_impl():
    if 'pdf_file' not in request.files:
        return _err('PDF ফাইল পাওয়া যায়নি', 400)
    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return _err('ফাইল সিলেক্ট করা হয়নি', 400)

    customer_type = 'IN-HOUSE'  # Label মডিউলে আপাতত শুধু IN-HOUSE
    customer_name = request.form.get('customer_name', '').strip()
    buyer_name = request.form.get('buyer_name', '').strip()
    po_number_override = request.form.get('po_number', '').strip()
    delivery_mode = request.form.get('delivery_mode', 'auto').strip()
    delivery_date_manual = request.form.get('delivery_date', '').strip()
    delivery_address = request.form.get('delivery_address', '').strip()
    width_mm = request.form.get('width_mm', '').strip()
    length_mm = request.form.get('length_mm', '').strip()
    remark_place = request.form.get('remark_place', '').strip().lower() in ('1', 'true', 'on', 'yes')
    remark_address = request.form.get('remark_address', '').strip().lower() in ('1', 'true', 'on', 'yes')
    force_override = request.form.get('force_override', '').strip().lower() in ('1', 'true', 'on', 'yes')

    buyer_error = validate_buyer_in_list(buyer_name, LABEL_BUYERS)
    if buyer_error:
        return _err(buyer_error)
    customer_error = validate_customer(customer_type, customer_name)
    if customer_error:
        return _err(customer_error)
    address_error = validate_delivery_address(customer_name, delivery_address)
    if address_error:
        return _err(address_error)

    if delivery_mode == 'manual':
        is_valid, err, parsed_date = validate_manual_delivery_date(delivery_date_manual)
        if not is_valid:
            return _err(err)
        delivery_date_final = format_delivery_date(parsed_date)
    else:
        delivery_date_final = format_delivery_date(get_default_delivery_date())

    pdf_bytes_raw = pdf_file.read()
    try:
        header_info, line_items, raw_df, summary_df = process_pdf_label(io.BytesIO(pdf_bytes_raw))
    except Exception as e:
        current_app.logger.exception('label process: PDF read failed')
        return _err(f'PDF পড়তে সমস্যা হয়েছে: {type(e).__name__}: {e}')

    if not line_items:
        return _err('কোনো লাইন-আইটেম পাওয়া যায়নি এই PDF থেকে (এটা Label PO ফরম্যাট কিনা দেখুন)')

    header_info['buyer'] = resolve_alias(header_info.get('buyer', ''), LABEL_BUYER_ALIASES)
    header_info['customer'] = resolve_alias(header_info.get('customer', ''), CUSTOMER_ALIASES)

    po_error = validate_po_number(po_number_override, header_info.get('po_number', ''))
    if po_error:
        return _err(po_error)
    if not force_override:
        customer_pdf_error = validate_matches_pdf('Customer', customer_name, header_info.get('customer', ''))
        if customer_pdf_error:
            return _err(customer_pdf_error)
        buyer_pdf_error = validate_matches_pdf('Buyer', buyer_name, header_info.get('buyer', ''))
        if buyer_pdf_error:
            return _err(buyer_pdf_error)

    warnings = validate_label_line_items(line_items)

    if buyer_name not in LABEL_VERIFIED_BUYERS:
        warnings.append(
            f"⚠️ '{buyer_name}' buyer-এর Label PO ফরম্যাট এখনো নির্দিষ্টভাবে যাচাই করা হয়নি — "
            f"আউটপুট (বিশেষ করে সাইজ/কোয়ান্টিটি/রেফারেন্স) ভালোভাবে চেক করে নিন।"
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
    if not width_mm or not length_mm:
        warnings.append("Width/Length (mm) ফাঁকা রাখা হয়েছে (ইউজার confirm করেছেন) — পরে ম্যানুয়ালি বসাতে হবে।")

    if not os.path.exists(LABEL_TEMPLATE_PATH):
        return _err(
            f"টেমপ্লেট ফাইল পাওয়া যায়নি: {LABEL_TEMPLATE_PATH} — "
            f"'template_label.xlsx' ফাইলটা template_files ফোল্ডারে রাখুন।", 500)

    base_name = os.path.splitext(pdf_file.filename)[0]
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, f'{base_name}_Output.xlsx')
            build_label_excel(
                line_items, header_info, out_path,
                customer_override=customer_name or None,
                buyer_override=buyer_name or None,
                po_override=po_number_override or None,
                delivery_date=delivery_date_final,
                delivery_address=delivery_address,
                width_mm=width_mm, length_mm=length_mm,
                raw_df=raw_df, summary_df=summary_df,
                warnings=warnings,
                remark_place=remark_place, remark_address=remark_address,
            )
            with open(out_path, 'rb') as f:
                file_bytes = f.read()
    except Exception as e:
        current_app.logger.exception('label process: build failed')
        return _err(f'Excel ফাইল বানাতে সমস্যা হয়েছে: {type(e).__name__}: {e}', 500)

    if not file_bytes:
        return _err('Excel ফাইল খালি তৈরি হয়েছে — আবার চেষ্টা করুন', 500)

    response = send_file(
        io.BytesIO(file_bytes), as_attachment=True,
        download_name=f'{base_name}_Output.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response.headers['Content-Length'] = str(len(file_bytes))
    response.headers['X-Warning-Count'] = str(len(warnings))
    return response
