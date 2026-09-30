"""
C&A বায়ারের সব কাস্টমারের জন্য একটাই এন্ট্রি পয়েন্ট। outhouse_extractor.py-এর
BATCH_REGISTRY থেকে (customer, buyer) দেখে এখানে কল আসে — এখানে customer_name
দিয়ে ঠিক extractor বেছে নেওয়া হয়, price-list একবার লোড করে সবার জন্য শেয়ার
করা হয়।

নতুন কাস্টমার যোগ করতে হলে:
  ১. customers/<নতুন>_extractor.py বানান (অন্যগুলোর প্যাটার্ন অনুসরণ করে —
     মূল entry point: read_xxx_excel(file_stream, filename, lookup=...,
     item_name_override=..., manual_ply=...) -> (line_items, warnings),
     আর প্রতিটা রো-র জন্য ca_rules.apply_ca_rules() কল করবে)
  ২. নিচের CUSTOMER_REGISTRY-তে একটা এন্ট্রি যোগ করুন
  ৩. outhouse_config.py + outhouse_extractor.py-এ (নিচে ইন্টিগ্রেশন
     ইনস্ট্রাকশন দেখুন) সাধারণ এন্ট্রি যোগ করুন
"""
import json
import os
import re

from .ca_calc import build_lookup
from .ca_rules import check_price_list_validity
from .customers.fakir_fashion_extractor import read_fakir_fashion_excel
from .customers.epyllion_knitwears_cna_extractor import read_epyllion_cna_excel
from .customers.dandspretty_extractor import read_dandspretty_excel

HERE = os.path.dirname(os.path.abspath(__file__))
_PRICE_LIST_PATH = os.path.join(HERE, 'ca_price_list.json')
_cache = {}  # {'lookup':..., 'price_list':...} — প্রসেসে একবার লোড হলেই যথেষ্ট


def _norm_key(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


CUSTOMER_REGISTRY = {
    _norm_key('Fakir Fashion Limited'): read_fakir_fashion_excel,
    _norm_key('Epyllion Knitwears Limited'): read_epyllion_cna_excel,
    _norm_key('D&S Pretty Fashions Ltd.'): read_dandspretty_excel,
}


def _get_price_data():
    if not _cache:
        with open(_PRICE_LIST_PATH, encoding='utf-8') as f:
            price_list = json.load(f)
        _cache['price_list'] = price_list
        _cache['lookup'] = build_lookup(price_list)
    return _cache['lookup'], _cache['price_list']


def combine_ca_booking_files(files, item_name_override='', manual_ply='',
                              buyer_name='', customer_name=''):
    """BATCH_REGISTRY-এর uniform কল-সিগনেচার। রিটার্ন করে
    (line_items, warnings)।"""
    lookup, price_list = _get_price_data()

    handler = CUSTOMER_REGISTRY.get(_norm_key(customer_name))
    if handler is None:
        return [], [
            f"⚠️ '{customer_name}'-এর জন্য এখনো C&A extractor বানানো হয়নি — "
            f"buyers_ca/dispatch.py-এর CUSTOMER_REGISTRY-তে এন্ট্রি নেই।"
        ]

    all_items = []
    all_warnings = []
    for file_stream, filename in files:
        try:
            file_stream.seek(0)
            items, warns = handler(
                file_stream, filename, lookup=lookup,
                item_name_override=item_name_override, manual_ply=manual_ply)
        except Exception as e:
            all_warnings.append(f"⚠️ '{filename}' প্রসেস করতে সমস্যা হয়েছে: {e}")
            continue
        all_items.extend(items)
        all_warnings.extend(warns)

    validity_warn = check_price_list_validity(price_list)
    if validity_warn:
        all_warnings.append(validity_warn)

    return all_items, all_warnings
