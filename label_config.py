"""
Label মডিউলের কনফিগ — Customer/Delivery Address/Buyer সব config.py-এর শেয়ার্ড
মাস্টার লিস্ট থেকে আসে (Thermal/Printing Press মডিউলের মতোই)। LABEL_VERIFIED_BUYERS-এ
শুধু ওই buyer-গুলো আছে যাদের Label PO PDF আসলেই টেস্ট করা হয়েছে — অন্য buyer ব্লক
হয় না, শুধু Warnings শীটে একটা নোট যোগ হয়।
"""
from config import CUSTOMERS, DELIVERY_ADDRESSES, CUSTOMER_ALIASES, BUYERS, resolve_alias  # noqa: F401

LABEL_BUYERS = BUYERS

LABEL_VERIFIED_BUYERS = [
    "Carhartt",
]

LABEL_BUYER_ALIASES = {}
