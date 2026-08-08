"""
Printing Press মডিউলের কনফিগ। Thermal-এর thermal_config.py-এর সাথে হুবহু একই
প্যাটার্ন — Customer/Buyer/Delivery Address সবই config.py-এর শেয়ার্ড মাস্টার
লিস্ট থেকে re-export করা হচ্ছে।

"in-house customer/buyer" মানে এখানে CUSTOMERS["IN-HOUSE"] (config.py-এর
মাস্টার dict-এর IN-HOUSE key) — Thermal মডিউলও ঠিক এই একই ধরনের কাজ করে
(app.py-তে thermal_index() রুট দেখুন, যেখানে পুরো CUSTOMERS dict পাঠানো হয়
কিন্তু thermal.html-এ customer_type হার্ডকোড করে IN-HOUSE রাখা থাকে) — তাই
এখানেও সেই একই কনভেনশন অনুসরণ করা হচ্ছে, আলাদা কোনো IN_HOUSE_CUSTOMERS
ভ্যারিয়েবলের দরকার নেই।
"""
from config import CUSTOMERS, DELIVERY_ADDRESSES, CUSTOMER_ALIASES, BUYERS, resolve_alias  # noqa: F401  (re-export)

# Thermal-এর THERMAL_BUYERS-এর মতোই — ড্রপডাউনে পুরো শেয়ার্ড মাস্টার Buyer লিস্ট
PRINTING_PRESS_BUYERS = BUYERS

# ---------------------------------------------------------------------------
# যেসব buyer-এর Printing Press PDF ফরম্যাট আসলেই টেস্ট করে নিশ্চিত হওয়া গেছে —
# Thermal-এর VERIFIED_BUYERS প্যাটার্নের মতোই। নতুন buyer verify হওয়ার পর
# এখানে যোগ করুন (BUYERS লিস্টে যেভাবে বানান লেখা আছে হুবহু সেভাবেই)।
# ---------------------------------------------------------------------------
PRINTING_PRESS_VERIFIED_BUYERS = [
    "Stanley Stella",
]

# ---------------------------------------------------------------------------
# Buyer Aliases — PDF-এ ভিন্ন/সংক্ষিপ্ত নামে থাকলে এখানে যোগ করুন
# (Thermal-এর THERMAL_BUYER_ALIASES-এর মতোই)।
# ---------------------------------------------------------------------------
PRINTING_PRESS_BUYER_ALIASES = {}

# ---------------------------------------------------------------------------
# Item Name Aliases — config.py-এর ITEM_NAME_ALIASES-এর মতোই প্যাটার্ন।
# PDF-এর কভার পেজে 'Item Name'-এর জায়গায় বিভিন্ন সংক্ষিপ্ত/ভিন্ন বানান আসতে
# পারে (P.S Tag, P.S. Tag, PS Tag, P S Tag...) — কিন্তু আমাদের ERP/টেমপ্লেটে
# এগুলো সবই আসলে একই আইটেম 'Poly Sticker' হিসেবে গণ্য হয়। এখানে যা যা ম্যাপ
# করা আছে, সব variant-ই ফাইনাল আউটপুটে 'Poly Sticker' হয়ে বসবে (resolve_alias
# case-insensitive/বাড়তি-স্পেস উপেক্ষা করে মেলায়, কিন্তু পাংচুয়েশন উপেক্ষা করে
# না — তাই ডট-সহ/ডট-ছাড়া দুই ভ্যারিয়েন্টই আলাদা করে এখানে রাখা হয়েছে)।
# নতুন কোনো ভ্যারিয়েন্ট PDF-এ দেখলে এখানে আরেকটা লাইন যোগ করে দিন।
# ---------------------------------------------------------------------------
PRINTING_PRESS_ITEM_NAME_ALIASES = {
    "P.S Tag": "Poly Sticker",
    "P.S. Tag": "Poly Sticker",
    "PS Tag": "Poly Sticker",
    "P S Tag": "Poly Sticker",
    "PS. Tag": "Poly Sticker",
    "P.S": "Poly Sticker",
    "PS": "Poly Sticker",
}