"""
C&A বায়ারের সব fixed কনফিগ এক জায়গায় — রেট, ফর্মুলার ধ্রুবক, লেখা।
রেট/ফর্মুলা বদলালে শুধু এই ফাইল বদলান।
"""
from decimal import Decimal

BUYER_NAME = "C&A BUYING GMBH & CO. KG"

# Excel ক্যালকুলেটরের (OEM_version_CA_EB_2H26_EPL.xlsx) ফর্মুলা:
#   Blank area (sqm) = 2 * (L + W + 60) * (W + H + 40) / 1,000,000     [L/W/H mm-এ]
#   Carton Unit price = ROUND(Blank area * SO SQM rate, 3)
# Remarks-এ 'Carton Unit price' বসে (ইউজার-কনফার্মড)। 'PO Unit Price' (PO রেট)
# শুধু রেফারেন্স — এখন ব্যবহার হয় না।
SO_SQM_RATE = Decimal("0.814")   # Carton Unit price-এর রেট (Excel-এর 'SO SQM rate')
PO_SQM_RATE = Decimal("0.74")    # শুধু রেফারেন্স ('PO SQM rate')
AREA_ADD_LW = 60                 # (L + W + 60)
AREA_ADD_WH = 40                 # (W + H + 40)
PRICE_DECIMALS = 3

EXCEPTIONAL_TEXT = "EXCEPTIONAL MEASUREMENT"   # Gmt. PO কলামে বসবে
STYLE_SUFFIX = "/TRI"                          # Style-এর শেষে
DEFAULT_PLY = "3"
PRICE_CHECK_TOLERANCE = Decimal("0.0015")      # PO-র Rate-এর সাথে মেলানোর সহনশীলতা
