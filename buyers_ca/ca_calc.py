"""
C&A কোড-ম্যাচিং আর এক্সেপশনাল-প্রাইস ক্যালকুলেশন — Decimal দিয়ে, যাতে
বাইনারি ফ্লোটিং-পয়েন্ট রাউন্ডিং ভুল (যেমন Python-এর সাধারণ round() মাঝে
মাঝে .5 কেসে ভুল দিক রাউন্ড করে) না হয় এবং Excel/LibreOffice-এর ROUND()-এর
সাথে হুবহু মিলে।
"""
import re
from decimal import Decimal, ROUND_HALF_UP

from . import ca_config as cfg


def _extract_base_and_perf(text):
    """'C&A-10 Perf', 'C&A 10 Perf', 'Additional-C&A-10 Perf (OL)',
    'C&A-17B' — সব রকম লেখাকে (base_code, is_perf) জোড়ায় নরমালাইজ করে,
    যাতে স্পেসিং/হাইফেন/প্রিফিক্স/'(OL)'-এর তফাত থাকলেও প্রাইস-লিস্টের
    কোডের সাথে মিলে যায়। C&A কোড-প্যাটার্নের বাইরের যেকোনো টেক্সট
    (যেমন 'Regular', 'Irregular') base=='' রিটার্ন করবে।"""
    if not text:
        return '', False
    t = str(text).upper()
    perf = bool(re.search(r'PERF|PERFORATED|\(\s*OL\s*\)', t))
    t = re.sub(r'ADDITIONAL[\s\-:]*', '', t)
    t = re.sub(r'PERFORATED|PERF|\(\s*OL\s*\)', '', t)
    t = re.sub(r'C\s*&\s*A', 'C&A', t)
    if 'C&A' not in t:
        return '', perf
    t = re.sub(r'C&A[\s\-]+', 'C&A-', t)
    t = re.sub(r'\s*-\s*', '-', t)
    t = re.sub(r'\s+', ' ', t).strip()
    t = t.rstrip('-').strip()
    if not t.startswith('C&A'):
        return '', perf
    return t, perf


def build_lookup(price_list):
    """price_list (ca_price_list.json-এর ডিক্ট) থেকে (base_code, is_perf)
    -> item এর লুকআপ ডিক্ট বানায়।"""
    lookup = {}
    for item in price_list.get('items', []):
        base, perf = _extract_base_and_perf(item['code'])
        if base:
            lookup[(base, perf)] = item
    return lookup


def match_code(raw_code_text, lookup):
    """রিটার্ন করে (matched_item_or_None, warning_or_None)। raw_code_text
    ফাঁকা/C&A-প্যাটার্নের বাইরের কিছু (Regular/Irregular ইত্যাদি) হলে
    (None, None) — কোনো ওয়ার্নিং ছাড়াই, কারণ এটাই স্বাভাবিক (এক্সেপশনাল
    হবে)। C&A-প্যাটার্নের কোড দেখতে পেলে কিন্তু লিস্টে না থাকলে ওয়ার্নিং
    সহ (None, warning)।"""
    if not raw_code_text or not str(raw_code_text).strip():
        return None, None
    base, perf = _extract_base_and_perf(raw_code_text)
    if not base:
        return None, None
    entry = lookup.get((base, perf))
    if entry is not None:
        return entry, None
    alt = lookup.get((base, not perf))
    if alt is not None:
        return None, (
            f"কোড '{raw_code_text}' দেওয়া আছে কিন্তু Perf/non-Perf মিলছে না — "
            f"কাছাকাছি '{alt['code']}' পাওয়া গেছে, ম্যানুয়ালি চেক করুন।"
        )
    return None, f"কোড '{raw_code_text}' আমাদের C&A প্রাইস লিস্টে পাওয়া যায়নি — ম্যানুয়ালি চেক করুন।"


def check_measurement_match(entry, l_mm, w_mm, h_mm):
    """entry (matched price-list item)-এর mm-মাপ আর বুকিং-এর mm-মাপ না
    মিললে ওয়ার্নিং স্ট্রিং রিটার্ন করে, মিললে None।"""
    if entry is None:
        return None
    if (entry['l_mm'], entry['w_mm'], entry['h_mm']) != (l_mm, w_mm, h_mm):
        return (
            f"⚠️ কোড '{entry['code']}'-এর প্রাইস-লিস্ট মেজারমেন্ট "
            f"{entry['l_mm']}x{entry['w_mm']}x{entry['h_mm']} mm, কিন্তু বুকিং-এ "
            f"{l_mm}x{w_mm}x{h_mm} mm পাওয়া গেছে — মিসম্যাচ, চেক করুন।"
        )
    return None


def exceptional_price(l_mm, w_mm, h_mm):
    """OEM ক্যালকুলেটর Excel-এর ফর্মুলা হুবহু (Decimal দিয়ে, Excel-এর
    ROUND()-এর সাথে মেলার জন্য ROUND_HALF_UP):
        area (sqm) = 2*(L+W+60)*(W+H+40) / 1,000,000        [L/W/H mm]
        price = ROUND(area * SO_SQM_RATE, 3)
    রিটার্ন করে Decimal (৩ দশমিক পর্যন্ত)।"""
    L, W, H = Decimal(l_mm), Decimal(w_mm), Decimal(h_mm)
    area = (Decimal(2) * (L + W + Decimal(cfg.AREA_ADD_LW)) *
            (W + H + Decimal(cfg.AREA_ADD_WH))) / Decimal(1_000_000)
    price = area * cfg.SO_SQM_RATE
    return price.quantize(Decimal('1.' + '0' * cfg.PRICE_DECIMALS), rounding=ROUND_HALF_UP)
