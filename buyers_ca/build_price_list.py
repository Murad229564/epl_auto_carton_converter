"""
নতুন C&A প্রাইস-কোটেশন PDF এলে এই স্ক্রিপ্ট চালান — ca_price_list.json আবার বানাবে।

ব্যবহার (প্রজেক্টের রুট ফোল্ডার থেকে):
    python -m buyers_ca.build_price_list "path/to/new_quote.pdf"
PDF না দিলে reference/-এর PDF ব্যবহার হবে।
"""
import json
import os
import re
import sys

import pdfplumber

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(HERE, "ca_price_list.json")
DEFAULT_PDF = os.path.join(HERE, "reference", "C_A_Prices_-_Bangladesh_20260701.pdf")

_ROW_RE = re.compile(
    r"(C&A-(?:EB\s*T\d+|\d+[A-Z]?)(?:\s+Perf\s*\(OL\))?)\s+(\d+)\s+(\d+)\s+(\d+)\s+.*?\$\s*([0-9]+\.[0-9]+)\s*$"
)
_EFFECTIVE_RE = re.compile(r"Effective Date:\s*(\d{4}-\d{2}-\d{2})")
_VALID_RE = re.compile(r"Good\s+\w+:\s*(\d{4}-\d{2}-\d{2})")


def parse_pdf(pdf_path):
    items, effective, valid_until = [], "", ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").split("\n"):
                m = _EFFECTIVE_RE.search(line)
                if m:
                    effective = m.group(1)
                m = _VALID_RE.search(line)
                if m:
                    valid_until = m.group(1)
                m = _ROW_RE.search(line.strip())
                if m:
                    code = re.sub(r"\s+", " ", m.group(1)).strip()
                    items.append({
                        "code": code,
                        "l_mm": int(m.group(2)),
                        "w_mm": int(m.group(3)),
                        "h_mm": int(m.group(4)),
                        "price": m.group(5),   # স্ট্রিং, যেমন "0.960" — ট্রেইলিং জিরো অক্ষুণ্ণ
                    })
    return {
        "effective_date": effective,
        "valid_until": valid_until,
        "currency": "USD",
        "source": os.path.basename(pdf_path),
        "items": items,
    }


def main():
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PDF
    data = parse_pdf(pdf_path)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"{len(data['items'])}টা কোড পাওয়া গেছে -> {OUT_PATH}")
    print(f"effective: {data['effective_date']}, valid until: {data['valid_until']}")


if __name__ == "__main__":
    main()
