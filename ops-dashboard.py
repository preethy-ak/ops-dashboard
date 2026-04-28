"""
Chat Analyzer Dashboard — Shopee & Lazada
==========================================
Graas.ai-themed Streamlit app for daily chat enquiry analysis.

Run:  streamlit run chat_analyzer_dashboard.py
Deps: pip install streamlit pandas openpyxl xlsxwriter
"""

import streamlit as st
import pandas as pd
import numpy as np
import re, io, warnings, gc
from datetime import datetime, timedelta
from collections import Counter

warnings.filterwarnings("ignore")

# ───────────────────A──────────────────────────────────────────────────────────
# PAGE CONFIG — Graas.ai theme
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Chat Analyzer Daashboard | Graas.ai",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM CSS — Graas.ai brand colours (#1B2A4A navy, #00C4B4 teal, #FF6B35 orange)
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Global ── */
html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', sans-serif; }
.main { background: #F4F6FB; }
.block-container { padding: 1.5rem 2rem; }

/* ── Top header bar ── */
.graas-header {
    background: linear-gradient(135deg, #1B2A4A 0%, #243554 100%);
    border-radius: 12px;
    padding: 1.2rem 1.8rem;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    gap: 1rem;
}
.graas-header h1 { color: #fff; margin: 0; font-size: 1.5rem; font-weight: 700; }
.graas-header p  { color: #A8C0D6; margin: 0; font-size: 0.85rem; }
.graas-logo { color: #00C4B4; font-size: 2rem; }

/* ── Metric cards ── */
.metric-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
.metric-card {
    background: #fff;
    border-radius: 10px;
    padding: 1rem 1.3rem;
    flex: 1;
    min-width: 150px;
    border-left: 4px solid #00C4B4;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.metric-card.orange { border-left-color: #FF6B35; }
.metric-card.red    { border-left-color: #E74C3C; }
.metric-card.navy   { border-left-color: #1B2A4A; }
.metric-card.green  { border-left-color: #27AE60; }
.metric-val { font-size: 1.9rem; font-weight: 800; color: #1B2A4A; }
.metric-label { font-size: 0.78rem; color: #7A8EA8; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; }
.metric-sub { font-size: 0.75rem; color: #A0AEC0; margin-top: 2px; }

/* ── Section titles ── */
.section-title {
    font-size: 1rem;
    font-weight: 700;
    color: #1B2A4A;
    border-bottom: 2px solid #00C4B4;
    padding-bottom: 0.4rem;
    margin: 1.5rem 0 1rem;
}

/* ── Priority badges ── */
.badge-high   { background:#FDECEA; color:#C0392B; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
.badge-medium { background:#FEF9E7; color:#D68910; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
.badge-low    { background:#EAF4FB; color:#2980B9; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }

/* ── Sentiment ── */
.sent-pos { color:#27AE60; font-weight:600; }
.sent-neu { color:#7F8C8D; font-weight:600; }
.sent-neg { color:#C0392B; font-weight:600; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] { background: #1B2A4A !important; }
section[data-testid="stSidebar"] .stMarkdown h2,
section[data-testid="stSidebar"] .stMarkdown h3 {
    color: #00C4B4 !important; font-size: 1rem !important; font-weight: 700 !important;
}
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stMarkdown span { color: #FFFFFF !important; }
/* Filter labels — bright white */
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stSelectbox > label,
section[data-testid="stSidebar"] .stMultiSelect > label,
section[data-testid="stSidebar"] .stDateInput > label,
section[data-testid="stSidebar"] .stTextInput > label {
    color: #FFFFFF !important;
    font-size: 0.85rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.3px !important;
}
/* Input boxes — white background for contrast */
section[data-testid="stSidebar"] .stSelectbox > div > div,
section[data-testid="stSidebar"] .stMultiSelect > div > div,
section[data-testid="stSidebar"] .stDateInput > div > div > input,
section[data-testid="stSidebar"] .stTextInput > div > div > input {
    background: #FFFFFF !important;
    color: #1B2A4A !important;
    border-radius: 6px !important;
    border: 1.5px solid #00C4B4 !important;
}
section[data-testid="stSidebar"] .stSelectbox svg,
section[data-testid="stSidebar"] .stMultiSelect svg { color: #1B2A4A !important; fill: #1B2A4A !important; }
/* Multiselect tag pills */
section[data-testid="stSidebar"] .stMultiSelect span[data-baseweb="tag"] {
    background: #00C4B4 !important; color: #fff !important;
}
/* Sidebar divider */
section[data-testid="stSidebar"] hr { border-color: #2E4A6A !important; }
/* Result count text */
section[data-testid="stSidebar"] strong { color: #00C4B4 !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] { background: #fff; border-radius:8px; padding:4px; gap:4px; }
.stTabs [data-baseweb="tab"] { border-radius:6px; padding:6px 18px; font-weight:600; color:#7A8EA8; }
.stTabs [aria-selected="true"] { background:#00C4B4 !important; color:#fff !important; }

/* ── Suggested reply box ── */
.reply-box {
    background: #F0FBF9;
    border: 1px solid #00C4B4;
    border-radius: 8px;
    padding: 0.9rem 1rem;
    font-size: 0.85rem;
    color: #1B2A4A;
    line-height: 1.6;
    margin-top: 0.5rem;
}
.reply-label { font-size:0.75rem; color:#00C4B4; font-weight:700; text-transform:uppercase; margin-bottom:4px; }

/* ── Upload area ── */
.upload-area {
    background: #fff;
    border: 2px dashed #00C4B4;
    border-radius: 12px;
    padding: 2rem;
    text-align: center;
    margin-bottom: 1rem;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

# Issue-type keyword mapping
ISSUE_KEYWORDS = {
    "Refund": [
        "refund", "คืนเงิน", "pengembalian dana", "dana kembali", "ibalik", "irefund",
        "bayar balik", "money back", "reimburse", "reimbursement",
    ],
    "Return": [
        "return", "คืนสินค้า", "retur", "rma", "send back", "ส่งคืน", "kembalikan",
        "return item", "product return",
    ],
    "Cancellation": [
        "cancel", "cancelled", "ยกเลิก", "batalkan", "batal", "cancellation",
        "cancel order", "ยกเลิกคำสั่งซื้อ",
    ],
    "Delay": [
        "delay", "late", "slow", "ช้า", "lambat", "belum sampai", "haven't received",
        "not arrived", "waiting", "รอนาน", "still waiting", "lama", "terlambat",
        "overdue", "not delivered yet", "ยังไม่ได้รับ", "belum diterima",
    ],
    "Damaged/Wrong Item": [
        "wrong item", "wrong product", "damaged", "broken", "defective",
        "สินค้าผิด", "ของเสีย", "ของแตก", "ของชำรุด", "rusak", "cacat",
        "salah barang", "salah produk", "not as described", "different item",
        "wrong size", "wrong colour", "wrong color", "different from picture",
    ],
    "Missing Item": [
        "missing", "not received", "didn't receive", "never received",
        "ไม่ได้รับ", "ของหาย", "hilang", "tidak diterima", "tidak ada", "kurang",
        "incomplete", "item missing", "package empty",
    ],
    "Payment Issue": [
        "payment", "ชำระเงิน", "bayar", "pembayaran", "charge", "double charge",
        "overcharged", "wrong charge", "billing", "invoice", "โอนเงิน", "จ่ายเงิน",
        "pay", "transfer", "deducted", "not paid",
    ],
    "Product Inquiry": [
        "how to", "how do", "วิธีใช้", "ราคา", "price", "size", "ขนาด",
        "สี", "colour", "color", "spec", "specification", "ingredient",
        "cara pakai", "ukuran", "warna", "harga", "stok", "stock", "available",
        "variant", "model", "version",
    ],
    "Promotion Issue": [
        "voucher", "promo", "discount", "coupon", "code", "sale", "offer",
        "โปรโมชั่น", "ส่วนลด", "โค้ด", "diskon", "kode promo", "cashback",
        "flash sale", "deal", "bundle",
    ],
    "Technical Issue": [
        "error", "bug", "cannot", "can't", "unable", "failed", "not working",
        "app issue", "website", "login", "checkout problem", "system",
        "ไม่สามารถ", "เกิดข้อผิดพลาด", "tidak bisa", "gagal", "eror",
    ],
    "Complaint": [
        "complain", "complaint", "terrible", "horrible", "awful", "worst",
        "ร้องเรียน", "ไม่พอใจ", "รำคาญ", "โกรธ", "disappointed",
        "frustrated", "unacceptable", "poor service", "bad service",
        "kecewa", "mengecewakan", "tidak puas", "buruk", "parah",
    ],
}

PRIORITY_MAP = {
    "High":   ["Refund", "Complaint", "Damaged/Wrong Item"],
    "Medium": ["Delay", "Missing Item", "Return", "Cancellation"],
    "Low":    ["Product Inquiry", "Promotion Issue", "Payment Issue", "Technical Issue"],
}

# ── Team Member → Store Code mapping (effective 30 March 2026) ────────────────
# GED = General Ecommerce Department (market label, not a store code)
TEAM_ASSIGNMENTS = {
    "Yeria":      ["AACMH", "FFH", "IKU",
                   "GED MY", "GEDMY", "GED_MY"],       # GED MY · all format variants
    "Syahira":    ["EWG", "HFC", "AAISS",
                   "GED SG", "GEDSG", "GED_SG"],       # GED SG · all format variants
    "Keerthana":  ["AABIY", "AABIW", "AAFTP",
                   "GED PH", "GEDPH", "GED_PH"],       # GED PH · all format variants
    "Alfian":     ["IGZ", "AADMJ", "AAEDD", "AADWP",
                   "IGZ ID", "IGZID", "IGZ_ID"],       # IGZ ID · all format variants
    "Jaye":       ["GSK", "DBC", "IEI", "FYW", "ILL"],
    "Ratchakorn": ["AABWU", "AAFHU", "AAFHB"],        # Full-time
}

# Reverse lookup: store_code → agent name
STORE_TO_AGENT = {
    store.upper(): agent
    for agent, stores in TEAM_ASSIGNMENTS.items()
    for store in stores
}

# Shift / market label per agent
AGENT_SHIFT = {
    "Yeria":      "GED MY · AACMH / FFH / IKU",
    "Syahira":    "GED SG · EWG / HFC / AAISS",
    "Keerthana":  "GED PH · AABIY / AABIW / AAFTP",
    "Alfian":     "IGZ ID · AADMJ / AAEDD / AADWP",
    "Jaye":       "GSK / DBC / IEI / FYW / ILL",
    "Ratchakorn": "Full-time · AABWU / AAFHU / AAFHB",
}

# Stalling phrases (seller still processing — NOT resolved)
STALLING_PATTERNS = [
    r"will (check|look|get back|follow up|investigate|verify|review|update)",
    r"let me (check|look into|verify|confirm|see)",
    r"(checking|looking into|investigating|following up|reviewing)",
    r"please (wait|hold on|allow us|bear with)",
    r"i will (check|get back|follow up|update)",
    r"we (are|will) (checking|looking|investigating|getting back|following up)",
    r"get back to you",
    r"bear with us",
    r"kindly (wait|allow|hold)",
    r"we'?ll? (check|look|get back|follow up)",
    r"akan (kami|segera) (cek|periksa|tindak lanjut|proses|hubungi)",
    r"mohon (tunggu|ditunggu|bersabar)",
    r"kami (sedang|akan) (cek|periksa|proses|tindak lanjut)",
    r"จะตรวจสอบ", r"กำลังตรวจสอบ", r"จะแจ้งกลับ",
    r"จะดำเนินการ", r"ขอตรวจสอบ", r"ขอเวลา",
    r"จะติดต่อกลับ", r"ติดตามให้", r"กำลังประสานงาน",
    r"escalat",
]

# Resolution phrases (conversation closed / solved)
RESOLUTION_PATTERNS = [
    r"refund (has been|was|is) (processed|completed|done|issued|approved)",
    r"(your|the) (order|item|package) (has been|was|is) (shipped|dispatched|replaced|delivered)",
    r"(issue|problem|case) (has been|was|is) (resolved|fixed|closed|sorted|handled)",
    r"(cancellation|cancel) (has been|was|is) (processed|done|completed|approved)",
    r"(we have|we've) (processed|completed|resolved|fixed|issued|sent)",
    r"please (expect|allow) (\d|few|some|a couple)",
    r"track.*link.*sent", r"tracking (number|id|code) (is|was|has been)",
    r"you (should|will) (receive|get) (it|your order|the item)",
    r"ดำเนินการเรียบร้อย", r"จัดการเรียบร้อย", r"แก้ไขเรียบร้อย",
    r"คืนเงินเรียบร้อย", r"ยกเลิกเรียบร้อย",
    r"sudah (diproses|selesai|dikirim|dikembalikan|dibatalkan)",
    r"telah (diproses|selesai|diselesaikan|dikirimkan)",
]

# Auto-reply detection
AUTO_REPLY_PATTERNS = [
    r"(thank you for contacting|thanks for reaching out).*auto",
    r"auto.?reply", r"automated (response|message|reply)",
    r"we'?ll? (get back|respond) (to you )?(within|in|shortly|soon)",
    r"our (team|agent).*(will|shall) (respond|reply|contact)",
    r"welcome to .*(official store|store).*\nhow (can|may) (we|i) help",
    r"สวัสดีค่ะ.*แอดมิน.*ยินดีให้บริการ",
    r"ยินดีต้อนรับ.*ร้าน",
    r"hi.{0,30}welcome to.{0,40}store",
]

# Positive / Negative sentiment keywords (multilingual)
POSITIVE_KWS = [
    "thank", "thanks", "great", "excellent", "awesome", "perfect", "love",
    "good", "nice", "happy", "satisfied", "wonderful", "amazing", "fantastic",
    "superb", "appreciate", "helpful", "fast", "quick", "well done", "recommend",
    "ขอบคุณ", "ดีมาก", "ประทับใจ", "พอใจ", "ยอดเยี่ยม", "ดีเลย", "ดีค่ะ", "ดีครับ",
    "terima kasih", "bagus", "mantap", "keren", "memuaskan", "puas", "oke baik",
    "salamat", "maganda", "ayos", "galing",
]

NEGATIVE_KWS = [
    "terrible", "worst", "angry", "disappointed", "frustrated", "cheated", "scam",
    "fraud", "fake", "broken", "damaged", "wrong item", "missing", "never received",
    "unacceptable", "horrible", "awful", "complain", "complaint", "refund",
    "ผิดหวัง", "โกรธ", "ไม่พอใจ", "แย่มาก", "แย่", "หลอกลวง", "ของเสีย",
    "ของปลอม", "ช้ามาก", "รอนาน", "สินค้าไม่ตรง", "ไม่ได้รับ", "ชำรุด",
    "tipu", "rusak", "cacat", "mengecewakan", "marah", "kecewa", "buruk", "parah",
    "salah", "tidak diterima", "hilang",
]

# Suggested replies per issue type
SUGGESTED_REPLIES = {
    "Refund": (
        "Thank you for reaching out, and we sincerely apologise for the inconvenience. "
        "We have reviewed your request and are pleased to confirm that your refund of [AMOUNT] "
        "has been initiated and will be reflected in your original payment method within 3–5 business days. "
        "Your order reference is [ORDER_ID]. We truly value your trust in us and hope to serve you better next time. "
        "If you have any further questions, please don't hesitate to reach out. 😊\n\n"
        "We'd love to hear your feedback — could you take a moment to rate your experience with us?"
    ),
    "Return": (
        "Thank you for contacting us about your return request. We're sorry to hear the product "
        "didn't meet your expectations. We've initiated the return process for order [ORDER_ID]. "
        "Please use the return label / return portal link we'll send to your registered email within 24 hours. "
        "Once we receive the item, the replacement or refund will be processed within 3–5 business days. "
        "We appreciate your patience and your continued support. 😊\n\n"
        "How would you rate your experience with us today?"
    ),
    "Cancellation": (
        "We've received your cancellation request for order [ORDER_ID]. We're sorry to see you go! "
        "Your order has been successfully cancelled and any payment made will be refunded within 3–5 business days. "
        "If you change your mind or need assistance with a future purchase, we're always here to help. 😊\n\n"
        "We'd appreciate your feedback — how was your experience with our team today?"
    ),
    "Delay": (
        "Thank you for your patience, and we sincerely apologise for the delay with your order [ORDER_ID]. "
        "We've checked with our logistics partner and your package is currently [STATUS]. "
        "Estimated delivery is [DATE]. We understand how frustrating delays can be and we truly appreciate your understanding. "
        "You can track your order in real time here: [TRACKING_LINK]. "
        "Please reach out if the delivery isn't received by [DATE+1] and we'll escalate immediately. 😊\n\n"
        "How was your experience with our support team today?"
    ),
    "Damaged/Wrong Item": (
        "We're truly sorry to hear that you received a damaged / incorrect item for order [ORDER_ID]. "
        "This is not the experience we want for you. To resolve this as quickly as possible, "
        "we've arranged a replacement to be dispatched within 1–2 business days. "
        "You do not need to return the incorrect / damaged item. "
        "We sincerely apologise for the inconvenience caused and will ensure this doesn't happen again. 😊\n\n"
        "Could you spare a moment to rate your support experience today?"
    ),
    "Missing Item": (
        "We're sorry to hear that your order [ORDER_ID] arrived with a missing item. "
        "We've raised an investigation with our fulfilment team and will have an update for you within 24 hours. "
        "In the meantime, we'll arrange a replacement or full refund, whichever you prefer. "
        "We apologise for this experience and truly appreciate your patience. 😊\n\n"
        "We'd love your feedback — how would you rate your experience with us today?"
    ),
    "Payment Issue": (
        "Thank you for flagging this payment concern. We've reviewed your account and order [ORDER_ID]. "
        "Our finance team has been notified and the discrepancy will be resolved within 2–3 business days. "
        "A confirmation will be sent to your registered email once completed. "
        "We apologise for any inconvenience and truly value your trust in us. 😊\n\n"
        "How was your experience with our support team today?"
    ),
    "Product Inquiry": (
        "Thank you for your interest in [PRODUCT_NAME]! "
        "Here are the details you requested: [DETAILS]. "
        "If you have more questions about specifications, sizing, or availability, "
        "please feel free to ask — we're happy to help you find the perfect product. 😊\n\n"
        "How can we assist you further today?"
    ),
    "Promotion Issue": (
        "Thank you for reaching out about the promotion. We're sorry for the confusion. "
        "We've reviewed your order [ORDER_ID] and confirmed that the discount of [AMOUNT] is applicable. "
        "The adjustment will be reflected within 24–48 hours. "
        "If the voucher code didn't apply correctly, please share it with us and we'll verify it right away. 😊\n\n"
        "How was your support experience today?"
    ),
    "Technical Issue": (
        "We apologise for the technical difficulty you're experiencing. "
        "Our team has been notified and is working on a resolution. "
        "In the meantime, please try [TROUBLESHOOTING STEP] and let us know if the issue persists. "
        "We aim to have this fully resolved within [TIMEFRAME]. "
        "Thank you for your patience — we appreciate it greatly. 😊\n\n"
        "How was your experience with our support today?"
    ),
    "Complaint": (
        "Thank you for taking the time to share your feedback, and we sincerely apologise for the experience you had. "
        "This is not the standard of service we strive for. We've escalated your case [CASE_ID] to our senior team "
        "for immediate review, and a dedicated agent will contact you within 4 hours. "
        "We take every concern seriously and are committed to making this right for you. 😊\n\n"
        "Your feedback helps us improve — how would you rate your support experience today?"
    ),
    "Other": (
        "Thank you for reaching out to us! We've reviewed your message and our team is addressing your concern. "
        "We aim to provide a resolution within 24 hours and will keep you updated throughout. "
        "We appreciate your patience and your trust in us. 😊\n\n"
        "How was your experience with our support team today?"
    ),
}

# Team tracking start date
TEAM_START_DATE = pd.Timestamp("2026-03-30")

# ── OOS / Lost-sale / Upsell keywords ────────────────────────────────────────
OOS_SELLER_PATTERNS = [
    r"(out of stock|sold out|no stock|not available|unavailable)",
    r"(habis|stok habis|tidak tersedia|kehabisan)",
    r"(หมดสต็อก|สินค้าหมด|ไม่มีสินค้า|หมดแล้ว|ไม่มีแล้ว|ยังไม่มี|ไม่มีจำหน่าย)",
    r"(wala na|wala nang stock|ubos na|hindi available)",
    r"(currently (not|out of|no) (stock|available|inventory))",
    r"(restock|back in stock|will be available soon)",
    r"(belum ada|belum tersedia|akan restock)",
    r"will not.*selling.*now",
]

OOS_BUYER_KEYWORDS = [
    "out of stock","no stock","sold out","not available","unavailable",
    "habis","stok habis","tidak tersedia","kehabisan",
    "หมดสต็อก","สินค้าหมด","ไม่มีสินค้า","หมดแล้ว","ไม่มีแล้ว",
    "wala na","wala nang stock","ubos na","hindi available",
]

LOST_SALE_PATTERNS = [
    r"(never mind|forget it|don.t want|not interested anymore|cancel it)",
    r"(found elsewhere|buying from another|going to another store|other shop)",
    r"(too expensive|price is high|cheaper elsewhere|other shop cheaper|mas mura)",
    r"(ไม่เป็นไร|ไม่ซื้อแล้ว|ซื้อที่อื่น|แพงไป|ที่อื่นถูกกว่า)",
    r"(tidak jadi|beli di tempat lain|terlalu mahal|lebih murah)",
    r"(hindi na|bibili na lang sa ibang|mas mura sa ibang)",
    r"(ok bye|okay bye|goodbye then|nevermind)\b",
]

UPSELL_BUYER_KEYWORDS = [
    "similar","other option","alternative","recommend","suggestion","bundle","combo",
    "what else","anything else","other product","related","go with","pair with",
    "ตัวอื่น","แนะนำ","ตัวไหนดี","อะไรดี","มีอะไรอีก","คล้ายกัน",
    "yang lain","alternatif","rekomendasi","produk lain","pilihan lain",
    "iba pa","ano pa","katulad","mas maganda",
]

ALTERNATIVE_SUGGESTED_PATTERNS = [
    r"(alternatively|similar product|you might like|we also have|how about)",
    r"(may i suggest|recommend|try this|check out|have you tried)",
    r"(นอกจากนี้|สินค้าอื่น|ลองดู|แนะนำ|มีตัวนี้ด้วย)",
    r"(kami juga ada|produk lain|alternatif|coba ini)",
    r"(meron din kami|pwede rin|katulad nito|subukan mo ito)",
]

SIZE_PATTERNS_SALES = [
    r"\b(XS|S|M|L|XL|XXL|XXXL)\b",
    r"\b(\d{1,3})\s*(cm|ml|mg|g|kg|oz)\b",
    r"\b(size|ukuran|ขนาด|sukat)\s*:?\s*([A-Z0-9\-\/]+)\b",
]

COLOR_PATTERNS_SALES = [
    r"\b(red|blue|green|black|white|pink|purple|yellow|orange|grey|gray|brown|navy|beige|cream|gold|silver)\b",
    r"\b(merah|biru|hijau|hitam|putih|ungu|kuning|oranye|abu|coklat|krem|emas|perak)\b",
    r"\b(แดง|น้ำเงิน|เขียว|ดำ|ขาว|ชมพู|ม่วง|เหลือง|ส้ม|เทา|น้ำตาล|ครีม|ทอง|เงิน)\b",
]

# Conversion / guided-order keywords (multilingual)
CONVERSION_KEYWORDS = [
    "i want to buy", "i'd like to buy", "i would like to buy", "how to buy",
    "how to order", "how do i order", "place an order", "can i order",
    "add to cart", "how to purchase", "i want to purchase", "proceed to checkout",
    "ready to buy", "i'll take it", "i want this", "i'll buy", "i want to get",
    "interested to buy", "interested in buying", "want to order",
    "อยากสั่ง", "สั่งซื้อ", "จะซื้อ", "ซื้อ", "สนใจซื้อ", "จะสั่ง",
    "mau beli", "mau order", "mau pesan", "ingin beli", "ingin order", "cara beli",
    "mag-order", "gusto kong bilhin", "bibilhin ko", "paano mag-order",
]

# Action steps per issue type (DKSH / GRAAS operational guide)
ACTION_STEPS = {
    "Refund": (
        "1. Verify order ID and payment method in Seller Centre.\n"
        "2. Check refund eligibility (within 15 days of purchase).\n"
        "3. Initiate refund via platform refund portal — select 'Approved by Seller'.\n"
        "4. Confirm refund amount matches original payment.\n"
        "5. Notify buyer with expected timeline (3–5 business days).\n"
        "6. Log in DKSH tracker under 'Refund Cases'."
    ),
    "Return": (
        "1. Verify product condition and return reason with buyer.\n"
        "2. Check return window (platform-specific: Lazada 7 days, Shopee 15 days).\n"
        "3. Approve return request in Seller Centre.\n"
        "4. Send return shipping label to buyer via platform chat.\n"
        "5. Once item received, inspect and process refund/replacement.\n"
        "6. Update DKSH tracker under 'Return Cases'."
    ),
    "Cancellation": (
        "1. Check order status — cancellable only before 'Ready to Ship'.\n"
        "2. Approve cancellation in Seller Centre if eligible.\n"
        "3. If already shipped, advise buyer to reject delivery.\n"
        "4. Refund will auto-process within 3–5 business days.\n"
        "5. Log in DKSH tracker under 'Cancellation Cases'."
    ),
    "Delay": (
        "1. Check logistics tracking in Seller Centre → Order Details.\n"
        "2. Contact logistics provider if package stalled > 3 days.\n"
        "3. Share tracking link with buyer immediately.\n"
        "4. If lost in transit, file a claim with logistics partner.\n"
        "5. Offer replacement or refund if delivery fails SLA.\n"
        "6. Escalate to platform CS if logistics partner unresponsive."
    ),
    "Damaged/Wrong Item": (
        "1. Request photo evidence from buyer (damaged/wrong item + packaging).\n"
        "2. Log dispute in Seller Centre under 'Return & Refund'.\n"
        "3. Approve replacement dispatch — do NOT ask buyer to return.\n"
        "4. Arrange courier pickup of damaged item (optional).\n"
        "5. Update DKSH tracker under 'Damaged/Wrong Item'.\n"
        "6. Report to warehouse for quality investigation."
    ),
    "Missing Item": (
        "1. Request unboxing video/photo from buyer as evidence.\n"
        "2. Check packing list vs order items in warehouse system.\n"
        "3. If confirmed missing, dispatch replacement within 24 hours.\n"
        "4. If uncertain, raise internal investigation with warehouse.\n"
        "5. Log in DKSH tracker under 'Missing Item'."
    ),
    "Payment Issue": (
        "1. Verify transaction details in platform payment dashboard.\n"
        "2. Check for double-charge or incorrect deduction.\n"
        "3. Raise dispute ticket with platform finance team.\n"
        "4. Provide buyer with case/ticket reference number.\n"
        "5. Follow up within 2 business days for resolution update."
    ),
    "Product Inquiry": (
        "1. Provide accurate product specs/details from official product sheet.\n"
        "2. If stock inquiry — check live inventory in Seller Centre.\n"
        "3. For sizing — share size guide image or chart.\n"
        "4. For availability — advise on restock ETA if applicable.\n"
        "5. Opportunity to upsell / cross-sell related products."
    ),
    "Promotion Issue": (
        "1. Verify voucher/promo code validity in Seller Centre → Promotions.\n"
        "2. Check eligibility criteria (min. spend, product category, date range).\n"
        "3. If code valid but not applied — advise buyer to re-checkout.\n"
        "4. If code expired — offer alternative discount if authorised.\n"
        "5. Escalate to marketing team for promo setup errors."
    ),
    "Technical Issue": (
        "1. Identify the platform and device buyer is using.\n"
        "2. Advise standard troubleshooting: clear cache, update app, reinstall.\n"
        "3. If platform-side issue — check platform status page.\n"
        "4. Raise support ticket with platform technical team.\n"
        "5. Keep buyer updated with ETA from platform team."
    ),
    "Complaint": (
        "1. Acknowledge and empathise — do NOT be defensive.\n"
        "2. Log complaint details in DKSH escalation tracker.\n"
        "3. Identify root cause (product/logistics/service failure).\n"
        "4. Offer concrete resolution: refund / replacement / discount.\n"
        "5. Escalate to senior manager if buyer threatens churn/review.\n"
        "6. Follow up within 4 hours with resolution update."
    ),
    "Other": (
        "1. Understand buyer's concern fully before responding.\n"
        "2. Route to appropriate team if issue is specialised.\n"
        "3. Aim to resolve within 24 hours.\n"
        "4. Log in DKSH tracker under 'General Enquiries'."
    ),
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def detect_sentiment(text: str) -> str:
    """Keyword-based multilingual sentiment detector."""
    if not isinstance(text, str) or not text.strip():
        return "Neutral"
    t = text.lower()
    neg = sum(1 for kw in NEGATIVE_KWS if kw in t)
    pos = sum(1 for kw in POSITIVE_KWS if kw in t)
    if neg > pos:
        return "Negative"
    if pos > neg:
        return "Positive"
    return "Neutral"


def detect_issue_type(text: str) -> str:
    """Classify message into one of 11 issue types using keyword matching."""
    if not isinstance(text, str) or not text.strip():
        return "Other"
    t = text.lower()
    scores = {}
    for issue, kws in ISSUE_KEYWORDS.items():
        score = sum(1 for kw in kws if kw.lower() in t)
        if score > 0:
            scores[issue] = score
    if not scores:
        return "Other"
    return max(scores, key=scores.get)


def get_priority(issue_type: str) -> str:
    """Map issue type to priority level."""
    for priority, issues in PRIORITY_MAP.items():
        if issue_type in issues:
            return priority
    return "Low"


def matches_any(text: str, patterns: list) -> bool:
    """Check if text matches any regex pattern (case-insensitive)."""
    if not isinstance(text, str):
        return False
    t = text.lower()
    return any(re.search(p, t, re.IGNORECASE) for p in patterns)


def is_auto_reply(text: str) -> bool:
    return matches_any(text, AUTO_REPLY_PATTERNS)


def conversation_is_unresolved(seller_msgs: list) -> bool:
    """
    Returns True if conversation has stalling phrases without a following resolution phrase.
    Strategy: scan chronologically. If stall found but no later resolution → unresolved.
    """
    stall_found = False
    for msg in seller_msgs:
        if matches_any(msg, STALLING_PATTERNS):
            stall_found = True
        if matches_any(msg, RESOLUTION_PATTERNS):
            stall_found = False   # Resolution found after stall → mark resolved
    return stall_found


def compute_csat(sentiment: str, is_resolved: bool) -> float:
    """Proxy CSAT score 1–5 based on sentiment + resolution."""
    matrix = {
        ("Positive", True):  5.0,
        ("Positive", False): 3.5,
        ("Neutral",  True):  4.0,
        ("Neutral",  False): 3.0,
        ("Negative", True):  2.5,
        ("Negative", False): 1.0,
    }
    return matrix.get((sentiment, is_resolved), 3.0)


def generate_summary(buyer_msgs: list, issue_type: str) -> str:
    """Rule-based buyer chat summary."""
    if not buyer_msgs:
        return "No buyer messages."
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)])[:400]
    return f"[{issue_type}] Buyer enquiry: {combined[:200]}{'...' if len(combined) > 200 else ''}"


def fmt_mins(mins) -> str:
    """Format float minutes → human-readable."""
    if pd.isna(mins) or mins < 0:
        return "—"
    if mins < 60:
        return f"{int(mins)}m"
    h = int(mins // 60)
    m = int(mins % 60)
    return f"{h}h {m}m" if m else f"{h}h"


def get_team_member(store_code: str) -> str:
    """Return agent name for a given store code.
    Unknown stores are grouped as 'Others' so they aggregate together
    in Team Performance. Individual store codes remain visible via STORE_CODE column.
    """
    code = str(store_code).strip().upper()
    if not code:
        return "Others"
    return STORE_TO_AGENT.get(code, "Others")


def detect_conversion(buyer_msgs: list) -> bool:
    """Detect if buyer expressed intent to buy / guided order."""
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)]).lower()
    return any(kw.lower() in combined for kw in CONVERSION_KEYWORDS)


def get_action_steps(issue_type: str) -> str:
    """Return DKSH/GRAAS operational action steps for the issue type."""
    return ACTION_STEPS.get(issue_type, ACTION_STEPS["Other"])



# ─────────────────────────────────────────────────────────────────────────────
# SALES INTELLIGENCE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def detect_oos_seller_confirmed(seller_msgs: list) -> bool:
    for msg in seller_msgs:
        if matches_any(msg, OOS_SELLER_PATTERNS):
            return True
    return False

def detect_oos_buyer_inquiry(buyer_msgs: list) -> bool:
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)]).lower()
    return any(kw.lower() in combined for kw in OOS_BUYER_KEYWORDS)

def detect_lost_sale(buyer_msgs: list) -> bool:
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)])
    return matches_any(combined, LOST_SALE_PATTERNS)

def detect_upsell_opportunity(buyer_msgs: list) -> bool:
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)]).lower()
    return any(kw.lower() in combined for kw in UPSELL_BUYER_KEYWORDS)

def detect_alternative_suggested(seller_msgs: list) -> bool:
    combined = " ".join([m for m in seller_msgs if isinstance(m, str)])
    return matches_any(combined, ALTERNATIVE_SUGGESTED_PATTERNS)

def extract_item_ids(msgs: list) -> list:
    """Extract item IDs from both item-type messages and text messages."""
    combined = " ".join([m for m in msgs if isinstance(m, str)])
    return list(dict.fromkeys(re.findall(r"item_id:(\d+)", combined, re.IGNORECASE)))

def extract_size_mentions(msgs: list) -> list:
    combined = " ".join([m for m in msgs if isinstance(m, str)])
    sizes = []
    for pat in SIZE_PATTERNS_SALES:
        for m in re.findall(pat, combined, re.IGNORECASE):
            sz = " ".join(m).strip() if isinstance(m, tuple) else str(m).strip()
            if sz: sizes.append(sz.upper())
    return list(set(sizes))

def extract_color_mentions(msgs: list) -> list:
    combined = " ".join([m for m in msgs if isinstance(m, str)])
    colors = []
    for pat in COLOR_PATTERNS_SALES:
        colors.extend(re.findall(pat, combined, re.IGNORECASE))
    return list(set([c.lower() for c in colors]))

def classify_sales_stage(buyer_msgs, is_conversion, is_oos_confirmed, is_lost_sale) -> str:
    if is_conversion: return "Converted"
    if is_lost_sale: return "Lost Sale"
    if is_oos_confirmed: return "OOS Demand"
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)]).lower()
    if any(kw in combined for kw in ["want to buy","interested","i want","อยากได้","mau beli","gusto ko"]): return "High Intent"
    if any(kw in combined for kw in ["price","ราคา","harga","how much","berapa","presyo","magkano"]): return "Price Check"
    if any(kw in combined for kw in ["size","available","stock","variant","color","colour","specification"]): return "Product Research"
    return "Awareness"

def compute_wow_mom(conv_df: pd.DataFrame) -> tuple:
    """Compute Week-on-Week and Month-on-Month performance comparison."""
    df = conv_df.copy()
    df = df[df["LAST_MSG_TIME"].notna()].copy()
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    # Add period columns (use start_time for consistent datetime grouping)
    df["WEEK"]  = df["LAST_MSG_TIME"].dt.to_period("W").apply(lambda r: r.start_time)
    df["MONTH"] = df["LAST_MSG_TIME"].dt.to_period("M").apply(lambda r: r.start_time)

    def agg_metrics(df_in, period_col):
        agg = (
            df_in.groupby(period_col)
            .agg(
                Conversations=("CONVERSATION_ID", "count"),
                Resolved=("IS_RESOLVED", "sum"),
                Unresolved=("IS_UNRESOLVED", "sum"),
                Avg_CSAT=("CSAT_PROXY", "mean"),
                Avg_CRT_mins=("AVG_CRT_MINS", "mean"),
                Negative=("SENTIMENT", lambda x: (x == "Negative").sum()),
                Positive=("SENTIMENT", lambda x: (x == "Positive").sum()),
                Conversions=("IS_CONVERSION", "sum"),
            )
            .reset_index()
            .sort_values(period_col)
        )
        agg["CRR_%"] = (agg["Resolved"] / agg["Conversations"] * 100).round(1)
        agg["Avg_CSAT"] = agg["Avg_CSAT"].round(2)
        agg["Avg_CRT_mins"] = agg["Avg_CRT_mins"].round(1)
        # Deltas (vs previous period)
        for col in ["Conversations", "Avg_CSAT", "CRR_%", "Avg_CRT_mins", "Conversions"]:
            agg[f"Δ {col}"] = agg[col].diff().round(2)
        return agg

    wow = agg_metrics(df, "WEEK")
    mom = agg_metrics(df, "MONTH")
    return wow, mom


def compute_team_performance(conv_df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate metrics per team member (from TEAM_START_DATE onwards)."""
    df = conv_df.copy()
    df = df[df["LAST_MSG_TIME"] >= TEAM_START_DATE].copy()
    if df.empty or "TEAM_MEMBER" not in df.columns:
        return pd.DataFrame()

    perf = (
        df.groupby("TEAM_MEMBER")
        .agg(
            Conversations=("CONVERSATION_ID", "count"),
            Resolved=("IS_RESOLVED", "sum"),
            Unresolved=("IS_UNRESOLVED", "sum"),
            Avg_CSAT=("CSAT_PROXY", "mean"),
            Avg_CRT_mins=("AVG_CRT_MINS", "mean"),
            Positive_Sent=("SENTIMENT", lambda x: (x == "Positive").sum()),
            Negative_Sent=("SENTIMENT", lambda x: (x == "Negative").sum()),
            Conversions=("IS_CONVERSION", "sum"),
            High_Priority=("PRIORITY", lambda x: (x == "High").sum()),
        )
        .reset_index()
    )
    perf["CRR_%"]    = (perf["Resolved"] / perf["Conversations"] * 100).round(1)
    perf["Avg_CSAT"] = perf["Avg_CSAT"].round(2)
    perf["Avg_CRT_mins"] = perf["Avg_CRT_mins"].round(1)
    perf["Shift"]    = perf["TEAM_MEMBER"].map(AGENT_SHIFT).fillna("Day")
    perf = perf.sort_values("Conversations", ascending=False).reset_index(drop=True)
    return perf


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes) -> pd.DataFrame:
    """Load Excel, combine both sheets, add PLATFORM column, parse dates."""
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets_found = xl.sheet_names

    dfs = []
    platform_map = {}
    for s in sheets_found:
        name_lower = s.lower()
        if "lazada" in name_lower:
            platform_map[s] = "Lazada"
        elif "shopee" in name_lower:
            platform_map[s] = "Shopee"
        else:
            platform_map[s] = "Unknown"
        df = xl.parse(s, dtype=str)
        df["PLATFORM"] = platform_map[s]
        dfs.append(df)

    combined = pd.concat(dfs, ignore_index=True)

    # Parse MESSAGE_TIME
    combined["MESSAGE_TIME"] = pd.to_datetime(combined["MESSAGE_TIME"], errors="coerce")

    # Normalise columns
    for col in ["STORE_CODE", "SITE_NICK_NAME_ID", "CHANNEL_NAME", "COUNTRY_CODE",
                "CONVERSATION_ID", "BUYER_NAME", "MESSAGE_PARSED",
                "MESSAGE_TYPE", "SENDER"]:
        if col in combined.columns:
            combined[col] = combined[col].fillna("").astype(str).str.strip()

    # Boolean flags — vectorised to avoid pandas 3.x lambda issues
    for flag in ["IS_READ", "IS_ANSWERED"]:
        if flag in combined.columns:
            combined[flag] = (
                combined[flag].astype(str).str.strip().str.lower()
                .isin(["true", "1", "yes"])
            )

    return combined


# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS ENGINE
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False, max_entries=1)
def analyse(df: pd.DataFrame) -> pd.DataFrame:
    """
    Analyse all conversations. Cached (max 1 entry) so re-renders never re-process.
    Memory optimised: categorical dtypes, truncated text, no large text duplicated per row.
    """
    df = df.copy()
    df["_sender_lower"] = df["SENDER"].str.lower().fillna("")
    df_sorted = df.sort_values(["CONVERSATION_ID", "MESSAGE_TIME"])

    buyer_mask  = df_sorted["_sender_lower"] == "buyer"
    seller_mask = df_sorted["_sender_lower"] == "seller"

    # ── Vectorize: aggregate buyer text per conversation ──────────────────────
    buyer_text_per_conv = (
        df_sorted[buyer_mask]
        .groupby("CONVERSATION_ID")["MESSAGE_PARSED"]
        .apply(lambda msgs: " ".join(m for m in msgs if isinstance(m, str)))
    )

    # ── Vectorize: sentiment & issue classification ───────────────────────────
    issue_map     = buyer_text_per_conv.apply(detect_issue_type)
    sentiment_map = buyer_text_per_conv.apply(detect_sentiment)

    # ── Metadata per conversation (first row) ─────────────────────────────────
    meta_cols = ["PLATFORM", "STORE_CODE", "SITE_NICK_NAME_ID", "CHANNEL_NAME",
                 "COUNTRY_CODE", "BUYER_NAME", "BUYER_ID", "IS_ANSWERED", "IS_READ"]
    meta_cols = [c for c in meta_cols if c in df_sorted.columns]
    meta_df = df_sorted.groupby("CONVERSATION_ID")[meta_cols].first()

    # ── Time bounds ───────────────────────────────────────────────────────────
    time_df = df_sorted.groupby("CONVERSATION_ID")["MESSAGE_TIME"].agg(
        FIRST_MSG_TIME="min", LAST_MSG_TIME="max"
    )

    # ── Message counts ────────────────────────────────────────────────────────
    total_msgs  = df_sorted.groupby("CONVERSATION_ID").size().rename("MSG_COUNT")
    buyer_msgs_count  = df_sorted[buyer_mask].groupby("CONVERSATION_ID").size().rename("BUYER_MSG_COUNT")
    seller_msgs_count = df_sorted[seller_mask].groupby("CONVERSATION_ID").size().rename("SELLER_MSG_COUNT")

    # ── Seller message lists for unresolved detection ─────────────────────────
    seller_msgs_per_conv = (
        df_sorted[seller_mask]
        .groupby("CONVERSATION_ID")["MESSAGE_PARSED"]
        .apply(list)
    )
    buyer_msgs_per_conv = (
        df_sorted[buyer_mask]
        .groupby("CONVERSATION_ID")["MESSAGE_PARSED"]
        .apply(list)
    )

    # ── Build result row by row (only CRT needs sequential access) ────────────
    rows = []
    for conv_id, grp in df_sorted.groupby("CONVERSATION_ID", sort=False):
        issue_type  = issue_map.get(conv_id, "Other")
        sentiment   = sentiment_map.get(conv_id, "Neutral")
        b_msgs      = buyer_msgs_per_conv.get(conv_id, [])
        s_msgs      = seller_msgs_per_conv.get(conv_id, [])
        meta        = meta_df.loc[conv_id] if conv_id in meta_df.index else {}

        is_unresolved = conversation_is_unresolved(s_msgs)
        is_resolved   = not is_unresolved
        priority      = get_priority(issue_type)
        csat          = compute_csat(sentiment, is_resolved)

        # ── Sales intelligence ────────────────────────────────────────────────
        is_conversion       = detect_conversion(b_msgs)
        is_oos_confirmed    = detect_oos_seller_confirmed(s_msgs)
        is_oos_buyer_inq    = detect_oos_buyer_inquiry(b_msgs)
        is_lost_sale        = detect_lost_sale(b_msgs)
        is_upsell_opp       = detect_upsell_opportunity(b_msgs)
        alt_suggested       = detect_alternative_suggested(s_msgs)
        # Extract item IDs from ALL messages (item-type rows store id in MESSAGE_PARSED)
        _all_msgs_list = list(grp["MESSAGE_PARSED"].fillna("").tolist())
        item_ids            = extract_item_ids(_all_msgs_list)
        size_mentions       = extract_size_mentions(b_msgs)
        color_mentions      = extract_color_mentions(b_msgs)
        sales_stage         = classify_sales_stage(b_msgs, is_conversion, is_oos_confirmed, is_lost_sale)

        # CRT: time between each buyer→seller pair
        crt_list = []
        last_buyer_time = None
        for sender, msg_time in zip(grp["_sender_lower"].tolist(), grp["MESSAGE_TIME"].tolist()):
            if sender == "buyer":
                last_buyer_time = msg_time
            elif sender == "seller" and last_buyer_time is not None:
                delta = (msg_time - last_buyer_time).total_seconds() / 60
                if 0 <= delta <= 1440:
                    crt_list.append(delta)
                last_buyer_time = None
        avg_crt = float(np.mean(crt_list)) if crt_list else np.nan

        def _get(field, default=""):
            try:
                return meta[field] if hasattr(meta, "__getitem__") else getattr(meta, field, default)
            except Exception:
                return default

        rows.append({
            "CONVERSATION_ID":   conv_id,
            "PLATFORM":          _get("PLATFORM"),
            "STORE_CODE":        _get("STORE_CODE"),
            "SITE_NICK_NAME_ID": _get("SITE_NICK_NAME_ID"),
            "CHANNEL_NAME":      _get("CHANNEL_NAME"),
            "COUNTRY_CODE":      _get("COUNTRY_CODE"),
            "BUYER_NAME":        _get("BUYER_NAME"),
            "BUYER_ID":          _get("BUYER_ID"),
            "FIRST_MSG_TIME":    time_df.loc[conv_id, "FIRST_MSG_TIME"] if conv_id in time_df.index else pd.NaT,
            "LAST_MSG_TIME":     time_df.loc[conv_id, "LAST_MSG_TIME"]  if conv_id in time_df.index else pd.NaT,
            "MSG_COUNT":         int(total_msgs.get(conv_id, 0)),
            "BUYER_MSG_COUNT":   int(buyer_msgs_count.get(conv_id, 0)),
            "SELLER_MSG_COUNT":  int(seller_msgs_count.get(conv_id, 0)),
            "ISSUE_TYPE":        issue_type,
            "PRIORITY":          priority,
            "SENTIMENT":         sentiment,
            "IS_UNRESOLVED":     is_unresolved,
            "IS_RESOLVED":       is_resolved,
            "CSAT_PROXY":        round(csat, 1),
            "AVG_CRT_MINS":      round(avg_crt, 1) if not np.isnan(avg_crt) else None,
            "BUYER_SUMMARY":     generate_summary(b_msgs, issue_type),
            "SUGGESTED_REPLY":   SUGGESTED_REPLIES.get(issue_type, SUGGESTED_REPLIES["Other"]),
            "ACTION_STEPS":      get_action_steps(issue_type),
            "IS_CONVERSION":     is_conversion,
            "IS_OOS_CONFIRMED":  is_oos_confirmed,
            "IS_OOS_BUYER_INQ":  is_oos_buyer_inq,
            "IS_LOST_SALE":      is_lost_sale,
            "IS_UPSELL_OPP":     is_upsell_opp,
            "ALT_SUGGESTED":     alt_suggested,
            "ITEM_IDS":          "|".join(item_ids) if item_ids else "",
            "SIZE_MENTIONS":     "|".join(size_mentions) if size_mentions else "",
            "COLOR_MENTIONS":    "|".join(color_mentions) if color_mentions else "",
            "SALES_STAGE":       sales_stage,
            "TEAM_MEMBER":       get_team_member(_get("STORE_CODE")),
            "IS_ANSWERED":       str(_get("IS_ANSWERED")).lower() == "true",
            "IS_READ":           str(_get("IS_READ")).lower() == "true",
        })

    result = pd.DataFrame(rows)

    # ── Memory optimisation: categorical dtypes for low-cardinality columns ───
    for col in ["PLATFORM", "ISSUE_TYPE", "PRIORITY", "SENTIMENT",
                "STORE_CODE", "CHANNEL_NAME", "COUNTRY_CODE", "TEAM_MEMBER", "SITE_NICK_NAME_ID", "SALES_STAGE"]:
        if col in result.columns:
            result[col] = result[col].astype("category")

    # Truncate long text columns to reduce RAM (full text not needed in-memory)
    for col in ["BUYER_SUMMARY"]:
        if col in result.columns:
            result[col] = result[col].str[:300]

    # Drop the heavy reply/action columns — looked up on-the-fly during display
    # They are re-attached only when building Excel export
    result.drop(columns=["SUGGESTED_REPLY", "ACTION_STEPS"], errors="ignore", inplace=True)

    gc.collect()
    return result


# ─────────────────────────────────────────────────────────────────────────────
# SALES / AM / MERCH AGGREGATIONS
# ─────────────────────────────────────────────────────────────────────────────

def build_sales_funnel(conv_df: pd.DataFrame) -> dict:
    total = len(conv_df)
    if total == 0:
        return {}
    df = conv_df.copy()
    for _bc in ["IS_CONVERSION","IS_OOS_CONFIRMED","IS_LOST_SALE","IS_UPSELL_OPP","ALT_SUGGESTED"]:
        if _bc in df.columns:
            df[_bc] = df[_bc].astype(bool).astype(int)
    prod_inq   = int((df["ISSUE_TYPE"].astype(str) == "Product Inquiry").sum())
    high_intent= int(df["SALES_STAGE"].astype(str).isin(["High Intent","Converted"]).sum())
    converted  = int(df["IS_CONVERSION"].sum())
    oos_total  = int(df["IS_OOS_CONFIRMED"].sum())
    lost       = int(df["IS_LOST_SALE"].sum())
    upsell_opp = int(df["IS_UPSELL_OPP"].sum())
    alt_acted  = int(df["ALT_SUGGESTED"].sum())
    return {
        "total": total, "prod_inq": prod_inq, "high_intent": high_intent,
        "converted": converted, "oos_total": oos_total, "lost": lost,
        "upsell_opp": upsell_opp, "alt_acted": alt_acted,
        "conv_rate":      round(converted / total * 100, 1),
        "lost_rate":      round(lost / total * 100, 1),
        "upsell_act_rate":round(alt_acted / upsell_opp * 100, 1) if upsell_opp else 0.0,
        "upsell_missed":  upsell_opp - alt_acted,
    }


def build_oos_tracker(conv_df: pd.DataFrame) -> pd.DataFrame:
    """OOS demand summary — derived entirely from conv_df (no raw_df needed)."""
    oos = conv_df[conv_df["IS_OOS_CONFIRMED"] == True].copy()
    if oos.empty:
        return pd.DataFrame()
    # Aggregate by store + item_ids to show restock priority
    rows_out = []
    for _, row in oos.iterrows():
        rows_out.append({
            "CONVERSATION_ID":  row["CONVERSATION_ID"],
            "STORE_CODE":       row.get("STORE_CODE", ""),
            "COUNTRY_CODE":     row.get("COUNTRY_CODE", ""),
            "PLATFORM":         str(row.get("PLATFORM", "")),
            "ITEM_IDS_INQUIRED":row.get("ITEM_IDS", "") or "Unknown",
            "SIZE_REQUESTED":   row.get("SIZE_MENTIONS", "") or "—",
            "COLOR_REQUESTED":  row.get("COLOR_MENTIONS", "") or "—",
            "ALT_SUGGESTED":    bool(row.get("ALT_SUGGESTED", False)),
            "LOST_SALE":        bool(row.get("IS_LOST_SALE", False)),
            "ISSUE_TYPE":       str(row.get("ISSUE_TYPE", "")),
            "BUYER_SUMMARY":    str(row.get("BUYER_SUMMARY", ""))[:120],
        })
    return pd.DataFrame(rows_out)


def build_product_demand(conv_df: pd.DataFrame):
    """Rich product demand — item IDs with full store/platform/channel/country context."""
    # All conversations that mention an item_id
    has_item = conv_df[conv_df["ITEM_IDS"].fillna("") != ""].copy()

    # Explode: one row per item_id per conversation
    rows_exp = []
    for _, row in has_item.iterrows():
        for iid in str(row["ITEM_IDS"]).split("|"):
            iid = iid.strip()
            if not iid:
                continue
            rows_exp.append({
                "Item_ID":       iid,
                "CONVERSATION_ID": row["CONVERSATION_ID"],
                "STORE_CODE":    str(row.get("STORE_CODE","")),
                "COUNTRY_CODE":  str(row.get("COUNTRY_CODE","")),
                "PLATFORM":      str(row.get("PLATFORM","")),
                "SITE_NICK_NAME_ID": str(row.get("SITE_NICK_NAME_ID","")),
                "CHANNEL_NAME":  str(row.get("CHANNEL_NAME","")),
                "TEAM_MEMBER":   str(row.get("TEAM_MEMBER","")),
                "ISSUE_TYPE":    str(row.get("ISSUE_TYPE","")),
                "IS_CONVERSION": bool(row.get("IS_CONVERSION", False)),
                "IS_OOS_CONFIRMED": bool(row.get("IS_OOS_CONFIRMED", False)),
                "IS_LOST_SALE":  bool(row.get("IS_LOST_SALE", False)),
                "IS_UPSELL_OPP": bool(row.get("IS_UPSELL_OPP", False)),
                "ALT_SUGGESTED": bool(row.get("ALT_SUGGESTED", False)),
                "SIZE_MENTIONS": str(row.get("SIZE_MENTIONS","")),
                "COLOR_MENTIONS":str(row.get("COLOR_MENTIONS","")),
                "SENTIMENT":     str(row.get("SENTIMENT","")),
                "BUYER_SUMMARY": str(row.get("BUYER_SUMMARY",""))[:120],
            })

    if not rows_exp:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    exp_df = pd.DataFrame(rows_exp)

    # ── Summary per Item ID ───────────────────────────────────────────────────
    item_summary = exp_df.groupby("Item_ID").agg(
        Total_Inquiries    = ("CONVERSATION_ID","count"),
        Unique_Convs       = ("CONVERSATION_ID","nunique"),
        Stores             = ("STORE_CODE",  lambda x: ", ".join(sorted(set(str(v) for v in x if v and v!="nan")))),
        Countries          = ("COUNTRY_CODE", lambda x: ", ".join(sorted(set(str(v) for v in x if v and v!="nan")))),
        Platforms          = ("PLATFORM",     lambda x: ", ".join(sorted(set(str(v) for v in x if v and v!="nan")))),
        Sites              = ("SITE_NICK_NAME_ID", lambda x: ", ".join(sorted(set(str(v) for v in x if v and v!="nan")))),
        Channels           = ("CHANNEL_NAME", lambda x: ", ".join(sorted(set(str(v) for v in x if v and v!="nan" and v)))),
        Conversions        = ("IS_CONVERSION",    lambda x: x.astype(bool).sum()),
        OOS_Confirmed      = ("IS_OOS_CONFIRMED", lambda x: x.astype(bool).sum()),
        Lost_Sales         = ("IS_LOST_SALE",     lambda x: x.astype(bool).sum()),
        Upsell_Opps        = ("IS_UPSELL_OPP",    lambda x: x.astype(bool).sum()),
        Sizes_Requested    = ("SIZE_MENTIONS",  lambda x: " | ".join(sorted(set(v for vals in x for v in str(vals).split("|") if v.strip())))),
        Colors_Requested   = ("COLOR_MENTIONS", lambda x: " | ".join(sorted(set(v for vals in x for v in str(vals).split("|") if v.strip())))),
    ).reset_index()
    item_summary["Conv_Rate_%"] = (item_summary["Conversions"]/item_summary["Total_Inquiries"]*100).round(1)
    item_summary["OOS_Rate_%"]  = (item_summary["OOS_Confirmed"]/item_summary["Total_Inquiries"]*100).round(1)
    item_summary = item_summary.sort_values("Total_Inquiries", ascending=False).reset_index(drop=True)

    # ── Variation demand ──────────────────────────────────────────────────────
    all_sizes  = [s for szs in conv_df["SIZE_MENTIONS"].fillna("").str.split("|") for s in szs if s.strip()]
    all_colors = [c for cls in conv_df["COLOR_MENTIONS"].fillna("").str.split("|") for c in cls if c.strip()]
    var_rows = (
        [{"Variation": k, "Type": "Size",  "Count": v} for k, v in Counter(all_sizes).most_common(20)] +
        [{"Variation": k.title(), "Type": "Color", "Count": v} for k, v in Counter(all_colors).most_common(20)]
    )
    var_df = pd.DataFrame(var_rows).sort_values("Count", ascending=False).reset_index(drop=True) if var_rows else pd.DataFrame(columns=["Variation","Type","Count"])

    return item_summary, var_df, exp_df


def build_am_scorecard(conv_df: pd.DataFrame) -> pd.DataFrame:
    """Per-store AM scorecard with sales + ops metrics."""
    df = conv_df.copy()
    for _bc in ["IS_CONVERSION","IS_LOST_SALE","IS_OOS_CONFIRMED","IS_UPSELL_OPP","ALT_SUGGESTED","IS_UNRESOLVED"]:
        if _bc in df.columns:
            df[_bc] = df[_bc].astype(bool).astype(int)
    grp = df.groupby(["STORE_CODE", "COUNTRY_CODE", "PLATFORM"]).agg(
        Total_Chats        = ("CONVERSATION_ID", "count"),
        Product_Inquiries  = ("ISSUE_TYPE",       lambda x: (x == "Product Inquiry").sum()),
        Conversions        = ("IS_CONVERSION",     "sum"),
        Lost_Sales         = ("IS_LOST_SALE",      "sum"),
        OOS_Hits           = ("IS_OOS_CONFIRMED",  "sum"),
        Upsell_Opps        = ("IS_UPSELL_OPP",     "sum"),
        Alt_Suggested      = ("ALT_SUGGESTED",     "sum"),
        Avg_CSAT           = ("CSAT_PROXY",        "mean"),
        Unresolved         = ("IS_UNRESOLVED",     "sum"),
        Avg_CRT_mins       = ("AVG_CRT_MINS",      "mean"),
    ).reset_index()
    grp["Conv_Rate_%"]   = (grp["Conversions"] / grp["Total_Chats"] * 100).round(1)
    grp["Lost_Rate_%"]   = (grp["Lost_Sales"]  / grp["Total_Chats"] * 100).round(1)
    grp["OOS_Rate_%"]    = (grp["OOS_Hits"]    / grp["Total_Chats"] * 100).round(1)
    grp["Upsell_Act_%"]  = (grp["Alt_Suggested"] / grp["Upsell_Opps"].replace(0, np.nan) * 100).round(1)
    grp["CRR_%"]         = ((grp["Total_Chats"] - grp["Unresolved"]) / grp["Total_Chats"] * 100).round(1)
    grp["Avg_CSAT"]      = grp["Avg_CSAT"].round(1)
    grp["Avg_CRT_mins"]  = grp["Avg_CRT_mins"].round(0)
    return grp.sort_values("Total_Chats", ascending=False).reset_index(drop=True)


def build_team_sales_perf(conv_df: pd.DataFrame) -> pd.DataFrame:
    """Extended team performance including sales KPIs."""
    df = conv_df[conv_df["LAST_MSG_TIME"] >= TEAM_START_DATE].copy()
    if df.empty or "TEAM_MEMBER" not in df.columns:
        return pd.DataFrame()
    # Cast boolean columns to int to avoid sum() issues with category dtype
    for _bc in ["IS_CONVERSION","IS_UPSELL_OPP","ALT_SUGGESTED","IS_LOST_SALE","IS_OOS_CONFIRMED","IS_RESOLVED","IS_UNRESOLVED"]:
        if _bc in df.columns:
            df[_bc] = df[_bc].astype(bool).astype(int)
    perf = df.groupby("TEAM_MEMBER").agg(
        Conversations    = ("CONVERSATION_ID", "count"),
        Resolved         = ("IS_RESOLVED",     "sum"),
        Unresolved       = ("IS_UNRESOLVED",   "sum"),
        Avg_CSAT         = ("CSAT_PROXY",       "mean"),
        Avg_CRT_mins     = ("AVG_CRT_MINS",     "mean"),
        Positive_Sent    = ("SENTIMENT",        lambda x: (x.astype(str) == "Positive").sum()),
        Negative_Sent    = ("SENTIMENT",        lambda x: (x.astype(str) == "Negative").sum()),
        Conversions      = ("IS_CONVERSION",    "sum"),
        High_Priority    = ("PRIORITY",         lambda x: (x.astype(str) == "High").sum()),
        Upsell_Opps      = ("IS_UPSELL_OPP",    "sum"),
        Alt_Suggested    = ("ALT_SUGGESTED",    "sum"),
        Lost_Sales       = ("IS_LOST_SALE",     "sum"),
        OOS_Handled      = ("IS_OOS_CONFIRMED", "sum"),
    ).reset_index()
    perf["CRR_%"]           = (perf["Resolved"] / perf["Conversations"] * 100).round(1)
    perf["Conv_Rate_%"]     = (perf["Conversions"] / perf["Conversations"] * 100).round(1)
    perf["Upsell_Act_%"]    = (perf["Alt_Suggested"] / perf["Upsell_Opps"].replace(0, np.nan) * 100).round(1)
    perf["Avg_CSAT"]        = perf["Avg_CSAT"].round(2)
    perf["Avg_CRT_mins"]    = perf["Avg_CRT_mins"].round(1)
    perf["Shift"]           = perf["TEAM_MEMBER"].map(AGENT_SHIFT).fillna("Day")
    return perf.sort_values("Conversations", ascending=False).reset_index(drop=True)


def generate_key_improvements(conv_df: pd.DataFrame, funnel: dict) -> list:
    """Return prioritised improvement recommendations from chat data."""
    recs = []
    total = funnel.get("total", 1)
    oos_pct   = funnel.get("oos_total", 0) / total * 100
    lost_pct  = funnel.get("lost_rate", funnel.get("lost", 0) / total * 100)
    upsell_act = funnel.get("upsell_act_rate", 0)
    conv_rate  = funnel.get("conv_rate", 0)
    neg_pct   = (conv_df["SENTIMENT"] == "Negative").sum() / total * 100 if total else 0
    unres_pct = conv_df["IS_UNRESOLVED"].mean() * 100 if total else 0
    avg_crt   = conv_df["AVG_CRT_MINS"].mean()

    if oos_pct > 5:
        recs.append(("🔴 HIGH", "Stock / OOS Management",
            f"{funnel.get('oos_total',0)} OOS inquiries ({oos_pct:.1f}%). "
            "Restock top-inquired items immediately. Share OOS tracker with buying team weekly."))
    if upsell_act < 40 and funnel.get("upsell_opp", 0) > 5:
        missed = funnel.get("upsell_missed", funnel.get("upsell_opp", 0) - funnel.get("alt_acted", 0))
        recs.append(("🔴 HIGH", "Upsell Execution Gap",
            f"{missed} of {funnel.get('upsell_opp',0)} upsell opps had no alternative suggested. "
            "Train agents: always offer a similar product when buyer asks 'any other options?' or item is OOS."))
    if lost_pct > 3:
        recs.append(("🔴 HIGH", "Lost Sales Recovery",
            f"{funnel.get('lost',0)} buyers ({lost_pct:.1f}%) disengaged without purchasing. "
            "Introduce a last-resort voucher or bundle offer to retain high-intent buyers."))
    if conv_rate < 10:
        recs.append(("🟡 MEDIUM", "Conversion Rate Improvement",
            f"Chat conversion rate is {conv_rate:.1f}%. Add CTA phrases and product links when buyer intent is high."))
    if unres_pct > 20:
        recs.append(("🟡 MEDIUM", "Chat Resolution Rate",
            f"{unres_pct:.1f}% of conversations are unresolved. Implement SLA reminders for stalled cases."))
    if not np.isnan(avg_crt) and avg_crt > 60:
        recs.append(("🟡 MEDIUM", "Response Time",
            f"Avg CRT is {fmt_mins(avg_crt)}. Target <30 min. Use quick-reply templates for common product enquiries."))
    if neg_pct > 15:
        recs.append(("🟡 MEDIUM", "Negative Sentiment",
            f"{neg_pct:.1f}% negative sentiment. Review top complaint themes and introduce proactive outreach."))
    recs.append(("🟢 OPPORTUNITY", "Variation Demand → Merch",
        "Top requested sizes and colours from chats = real demand. Share monthly with merch/buying team for stock planning."))
    recs.append(("🟢 OPPORTUNITY", "Product Inquiry → Conversion",
        "Warm leads from product inquiries are unconverted. Flag for follow-up. Add urgency cues (limited stock, flash sale)."))
    return recs


def build_excel(conv_df: pd.DataFrame, today_str: str) -> bytes:
    """Build full 10-sheet Excel — Ops + Sales + OOS + Products + AM + Team + Improvements."""
    df = conv_df.copy()
    # Re-attach reply/action columns
    if "SUGGESTED_REPLY" not in df.columns and "ISSUE_TYPE" in df.columns:
        df["SUGGESTED_REPLY"] = df["ISSUE_TYPE"].astype(str).map(lambda it: SUGGESTED_REPLIES.get(it, SUGGESTED_REPLIES["Other"]))
    if "ACTION_STEPS" not in df.columns and "ISSUE_TYPE" in df.columns:
        df["ACTION_STEPS"] = df["ISSUE_TYPE"].astype(str).map(get_action_steps)

    # Bool-safe copies for aggregation
    for _bc in ["IS_CONVERSION","IS_LOST_SALE","IS_OOS_CONFIRMED","IS_UPSELL_OPP","ALT_SUGGESTED","IS_RESOLVED","IS_UNRESOLVED"]:
        if _bc in df.columns:
            df[_bc] = df[_bc].astype(bool).astype(int)

    total = len(df)
    today_df = df[df["LAST_MSG_TIME"].dt.normalize() == pd.Timestamp(today_str)]
    funnel   = build_sales_funnel(conv_df)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book

        # ── Formats ──────────────────────────────────────────────────────────
        hdr_fmt  = wb.add_format({"bold":True,"bg_color":"#1B2A4A","font_color":"#FFFFFF","border":1,"font_size":10,"align":"center","valign":"vcenter"})
        ops_fmt  = wb.add_format({"bold":True,"bg_color":"#00C4B4","font_color":"#FFFFFF","border":1,"font_size":10})
        sal_fmt  = wb.add_format({"bold":True,"bg_color":"#FF6B35","font_color":"#FFFFFF","border":1,"font_size":10})
        mer_fmt  = wb.add_format({"bold":True,"bg_color":"#F59E0B","font_color":"#FFFFFF","border":1,"font_size":10})
        imp_fmt  = wb.add_format({"bold":True,"bg_color":"#8B5CF6","font_color":"#FFFFFF","border":1,"font_size":10})
        num_fmt  = wb.add_format({"num_format":"#,##0","border":1})
        pct_fmt  = wb.add_format({"num_format":"0.0%","border":1})
        cell_fmt = wb.add_format({"border":1,"font_size":9,"text_wrap":True,"valign":"top"})
        red_fmt  = wb.add_format({"border":1,"font_size":9,"bg_color":"#FDECEA","font_color":"#C0392B"})
        grn_fmt  = wb.add_format({"border":1,"font_size":9,"bg_color":"#E9F7EF","font_color":"#196F3D"})
        yel_fmt  = wb.add_format({"border":1,"font_size":9,"bg_color":"#FEF9E7","font_color":"#D68910"})
        title_fmt= wb.add_format({"bold":True,"font_size":14,"font_color":"#1B2A4A"})
        sub_title= wb.add_format({"italic":True,"font_color":"#7A8EA8","font_size":10})
        note_fmt = wb.add_format({"italic":True,"font_color":"#94A3B8","font_size":9,"text_wrap":True})

        def write_df_sheet(ws, df_in, hdr_f=hdr_fmt, start_row=0, col_widths=None):
            df_in = df_in.copy()
            # Format datetimes
            for col in df_in.select_dtypes(include=["datetime64[ns]","datetime64[ns, UTC]"]).columns:
                df_in[col] = df_in[col].dt.strftime("%Y-%m-%d %H:%M").fillna("")
            for c_idx, col in enumerate(df_in.columns):
                ws.write(start_row, c_idx, col, hdr_f)
            for r_idx, row in enumerate(df_in.itertuples(index=False), start=start_row+1):
                for c_idx, val in enumerate(row):
                    if val is None or (isinstance(val, float) and np.isnan(val)):
                        ws.write(r_idx, c_idx, "", cell_fmt)
                    elif isinstance(val, bool):
                        ws.write(r_idx, c_idx, "Yes" if val else "No", grn_fmt if val else cell_fmt)
                    elif isinstance(val, (int, float)):
                        ws.write_number(r_idx, c_idx, float(val), num_fmt)
                    else:
                        ws.write(r_idx, c_idx, str(val), cell_fmt)
            if col_widths:
                for c_idx, w in enumerate(col_widths):
                    ws.set_column(c_idx, c_idx, w)

        def kv_block(ws, data, row_start, label_fmt, val_fmt=cell_fmt):
            for i, (k, v) in enumerate(data):
                ws.write(row_start+i, 0, k, label_fmt)
                ws.write(row_start+i, 1, str(v), val_fmt)
            return row_start + len(data)

        # ════════════════════════════════════════════════════════════════════
        # SHEET 1 : README — What is this report?
        # ════════════════════════════════════════════════════════════════════
        ws_readme = wb.add_worksheet("📖 README")
        writer.sheets["📖 README"] = ws_readme
        ws_readme.set_column(0, 0, 30); ws_readme.set_column(1, 1, 80)
        ws_readme.write(0, 0, "Chat Analyzer Dashboard — Export Guide", title_fmt)
        ws_readme.write(1, 0, f"Generated: {today_str} | Graas.ai", sub_title)

        readme_rows = [
            ("WHAT IS THIS REPORT?", ""),
            ("Purpose", "This report turns Shopee/Lazada buyer chat conversations into 3 types of intelligence: (1) Service — how well are we handling buyers, (2) Sales — how much revenue are we missing, (3) Merch/Stock — what should we restock and list."),
            ("Data Source", "Shopee/Lazada chat export (.xlsx). Each row = one conversation (not one message). Fields are derived by AI analysis of the full message thread."),
            ("", ""),
            ("SHEET GUIDE", ""),
            ("📊 Ops Summary", "Overall performance: resolution rate, response time, CSAT, issue breakdown by store & country. For CS managers & team leads."),
            ("🔥 Today Priority", "Today's conversations sorted by priority (High first). Use for daily triage. Includes suggested reply per conversation."),
            ("🔴 Unresolved Chats", "All conversations detected as unresolved — buyer issue not yet closed. Action immediately."),
            ("💬 Suggested Replies", "All conversations with AI-generated reply templates and action steps. Use for agent training and response quality."),
            ("💰 Sales Intelligence", "Conversion funnel, lost sales, upsell opportunities, OOS demand. For sales managers and e-com growth leads."),
            ("💸 Lost Sales Detail", "Every conversation where a buyer disengaged without purchasing. Includes reason, store, agent, item IDs."),
            ("📦 OOS Restock List", "Items confirmed out-of-stock in chat, ranked by demand + lost sales. Share with buying team for restock prioritisation."),
            ("🛍️ Product Intelligence", "Every item_id mentioned in chat — inquiry count, platform, store, country, sizes/colours requested, OOS & conversion flags."),
            ("👥 Team Performance", "Per-agent scorecard: resolution rate, response time, CSAT, conversions, upsell action rate, lost sales."),
            ("🏪 AM Store Scorecard", "Per-store: conversion %, lost sale %, OOS %, upsell action %, CRR. For account managers."),
            ("🎯 Key Improvements", "Auto-generated prioritised recommendations from the data. High / Medium / Opportunity."),
            ("", ""),
            ("KEY METRICS EXPLAINED", ""),
            ("CRR (Chat Resolution Rate)", "% of conversations where the buyer's issue was fully resolved. Target: >80%"),
            ("CRT (Chat Response Time)", "Average minutes from buyer message to seller reply. Target: <30 minutes"),
            ("CSAT Proxy (1–5)", "Estimated satisfaction score based on sentiment + resolution status. Not from a survey."),
            ("Sales Stage", "Classified conversation intent: Awareness / Product Research / Price Check / High Intent / Converted / OOS Demand / Lost Sale"),
            ("Upsell Action Rate", "% of conversations where buyer asked for alternatives AND seller actually suggested one. Low rate = revenue leak."),
            ("OOS Priority Score", "Demand Count + (Lost Sales × 2). Higher = more urgent to restock."),
            ("Lost Sale", "Buyer showed purchase intent but disengaged — due to OOS, price, no response, or went elsewhere."),
        ]
        row = 3
        for k, v in readme_rows:
            if k in ("WHAT IS THIS REPORT?","SHEET GUIDE","KEY METRICS EXPLAINED"):
                ws_readme.write(row, 0, k, ops_fmt); ws_readme.write(row, 1, "", ops_fmt)
            elif k == "":
                row += 1; continue
            else:
                ws_readme.write(row, 0, k, hdr_fmt); ws_readme.write(row, 1, v, cell_fmt)
            row += 1

        # ════════════════════════════════════════════════════════════════════
        # SHEET 2 : Ops Summary
        # ════════════════════════════════════════════════════════════════════
        ws2 = wb.add_worksheet("📊 Ops Summary")
        writer.sheets["📊 Ops Summary"] = ws2
        ws2.set_column(0, 0, 32); ws2.set_column(1, 1, 22)
        ws2.write(0, 0, f"Operations Summary — {today_str}", title_fmt)
        ws2.write(1, 0, "Chat service performance across all stores and platforms", sub_title)

        crr      = round(df["IS_RESOLVED"].sum()/total*100,1) if total else 0
        avg_crt  = df["AVG_CRT_MINS"].mean()
        avg_csat = df["CSAT_PROXY"].mean()
        hi_today = len(today_df[today_df["PRIORITY"].astype(str)=="High"])
        neg_pct  = round((df["SENTIMENT"].astype(str)=="Negative").sum()/total*100,1) if total else 0

        ops_data = [
            ("── OPERATIONS KPIs ──", ""),
            ("Total Conversations", total),
            ("Today's Conversations", len(today_df)),
            ("Resolved", int(df["IS_RESOLVED"].sum())),
            ("Unresolved", int(df["IS_UNRESOLVED"].sum())),
            ("Chat Resolution Rate (CRR)", f"{crr}%"),
            ("Avg Response Time (CRT)", fmt_mins(avg_crt)),
            ("Avg CSAT Proxy (1–5)", round(avg_csat,2) if not np.isnan(avg_csat) else "—"),
            ("Today's High Priority Chats", hi_today),
            ("Negative Sentiment %", f"{neg_pct}%"),
        ]
        row = 3
        for k, v in ops_data:
            fmt = ops_fmt if "──" in k else cell_fmt
            ws2.write(row, 0, k, ops_fmt if "──" in k else hdr_fmt)
            ws2.write(row, 1, str(v), cell_fmt); row += 1

        row += 1
        ws2.write(row, 0, "Issue Type", ops_fmt); ws2.write(row, 1, "Count", hdr_fmt); ws2.write(row, 2, "% of Total", hdr_fmt); row += 1
        for issue, cnt in df["ISSUE_TYPE"].value_counts().items():
            ws2.write(row, 0, issue, cell_fmt); ws2.write_number(row, 1, int(cnt), num_fmt)
            ws2.write(row, 2, f"{cnt/total*100:.1f}%", cell_fmt); row += 1

        row += 1
        ws2.write(row, 0, "Store", ops_fmt); ws2.write(row, 1, "Convs", hdr_fmt); ws2.write(row, 2, "Unresolved", hdr_fmt); ws2.write(row, 3, "CRR%", hdr_fmt); ws2.write(row, 4, "Avg CSAT", hdr_fmt); row += 1
        store_agg = df.groupby("STORE_CODE").agg(Convs=("CONVERSATION_ID","count"), Unres=("IS_UNRESOLVED","sum"), CSAT=("CSAT_PROXY","mean")).reset_index().sort_values("Convs",ascending=False)
        for _, r2 in store_agg.iterrows():
            ws2.write(row, 0, str(r2["STORE_CODE"]), cell_fmt)
            ws2.write_number(row, 1, int(r2["Convs"]), num_fmt)
            ws2.write_number(row, 2, int(r2["Unres"]), num_fmt)
            ws2.write(row, 3, f"{(1-r2['Unres']/r2['Convs'])*100:.1f}%", cell_fmt)
            ws2.write(row, 4, f"{r2['CSAT']:.1f}" if not np.isnan(r2['CSAT']) else "—", cell_fmt); row += 1

        # ════════════════════════════════════════════════════════════════════
        # SHEET 3 : Today Priority Chats
        # ════════════════════════════════════════════════════════════════════
        p_cols = [c for c in ["CONVERSATION_ID","PLATFORM","STORE_CODE","CHANNEL_NAME","SITE_NICK_NAME_ID","COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME","LAST_MSG_TIME","ISSUE_TYPE","PRIORITY","SENTIMENT","IS_UNRESOLVED","CSAT_PROXY","AVG_CRT_MINS","IS_CONVERSION","BUYER_SUMMARY","SUGGESTED_REPLY","ACTION_STEPS"] if c in df.columns]
        today_pri = today_df.sort_values("PRIORITY", key=lambda s: s.map({"High":0,"Medium":1,"Low":2}).fillna(3))[p_cols]
        write_df_sheet(wb.add_worksheet("🔥 Today Priority"), today_pri, hdr_fmt, col_widths=[36,10,10,12,10,8,12,16,16,16,8,10,11,6,7,11,50,60,60])
        writer.sheets["🔥 Today Priority"] = wb.worksheets()[-1]

        # ════════════════════════════════════════════════════════════════════
        # SHEET 4 : Unresolved Chats
        # ════════════════════════════════════════════════════════════════════
        unres_df = df[df["IS_UNRESOLVED"]==1][p_cols].sort_values("PRIORITY", key=lambda s: s.map({"High":0,"Medium":1,"Low":2}).fillna(3))
        write_df_sheet(wb.add_worksheet("🔴 Unresolved Chats"), unres_df, hdr_fmt, col_widths=[36,10,10,12,10,8,12,16,16,16,8,10,11,6,7,11,50,60,60])
        writer.sheets["🔴 Unresolved Chats"] = wb.worksheets()[-1]

        # ════════════════════════════════════════════════════════════════════
        # SHEET 5 : Suggested Replies — All Conversations
        # ════════════════════════════════════════════════════════════════════
        reply_cols = [c for c in ["CONVERSATION_ID","STORE_CODE","COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME","ISSUE_TYPE","PRIORITY","SENTIMENT","BUYER_SUMMARY","SUGGESTED_REPLY","ACTION_STEPS"] if c in df.columns]
        write_df_sheet(wb.add_worksheet("💬 Suggested Replies"), df[reply_cols], hdr_fmt, col_widths=[36,10,8,12,16,16,8,10,50,70,60])
        writer.sheets["💬 Suggested Replies"] = wb.worksheets()[-1]

        # ════════════════════════════════════════════════════════════════════
        # SHEET 6 : Sales Intelligence
        # ════════════════════════════════════════════════════════════════════
        ws6 = wb.add_worksheet("💰 Sales Intelligence")
        writer.sheets["💰 Sales Intelligence"] = ws6
        ws6.set_column(0, 0, 32); ws6.set_column(1, 1, 20)
        ws6.write(0, 0, "Sales Intelligence — Revenue from Chat", title_fmt)
        ws6.write(1, 0, "How much money are we leaving on the table in chat?", sub_title)

        sal_kpis = [
            ("── SALES FUNNEL ──",""),
            ("Total Conversations", total),
            ("Product Inquiries (warm leads)", int((df["ISSUE_TYPE"].astype(str)=="Product Inquiry").sum())),
            ("High Intent Buyers", funnel.get("high_intent",0)),
            ("Conversions (purchase intent detected)", funnel.get("converted",0)),
            ("Conversion Rate", f"{funnel.get('conv_rate',0)}%"),
            ("Lost Sales (buyer disengaged)", funnel.get("lost",0)),
            ("Lost Sale Rate", f"{funnel.get('lost_rate',0):.1f}%"),
            ("",""),
            ("── UPSELL ──",""),
            ("Upsell Opportunities (buyer asked for alternatives)", funnel.get("upsell_opp",0)),
            ("Alternatives Actually Suggested by Agent", funnel.get("alt_acted",0)),
            ("Missed Upsells (no suggestion given)", funnel.get("upsell_missed",0)),
            ("Upsell Action Rate", f"{funnel.get('upsell_act_rate',0)}%"),
            ("",""),
            ("── OOS DEMAND ──",""),
            ("OOS Inquiries (seller confirmed no stock)", funnel.get("oos_total",0)),
            ("OOS → Lost Sale", int(df[(df["IS_OOS_CONFIRMED"]==1)&(df["IS_LOST_SALE"]==1)]["CONVERSATION_ID"].count())),
        ]
        row = 3
        for k, v in sal_kpis:
            if "──" in k:
                ws6.write(row, 0, k, sal_fmt); ws6.write(row, 1, "", sal_fmt)
            elif k == "":
                row += 1; continue
            else:
                ws6.write(row, 0, k, hdr_fmt); ws6.write(row, 1, str(v), cell_fmt)
            row += 1

        row += 1
        ws6.write(row, 0, "Sales Stage Distribution", sal_fmt); row += 1
        for stage, cnt in df["SALES_STAGE"].astype(str).value_counts().items():
            ws6.write(row, 0, stage, cell_fmt); ws6.write_number(row, 1, int(cnt), num_fmt)
            ws6.write(row, 2, f"{cnt/total*100:.1f}%", cell_fmt); row += 1

        row += 2
        ws6.write(row, 0, "Per-Store Sales Scorecard", sal_fmt); row += 1
        store_sal = df.groupby(["STORE_CODE","COUNTRY_CODE"]).agg(
            Convs=("CONVERSATION_ID","count"), Conversions=("IS_CONVERSION","sum"),
            Lost=("IS_LOST_SALE","sum"), OOS=("IS_OOS_CONFIRMED","sum"),
            Upsell_Opps=("IS_UPSELL_OPP","sum"), Alt_Suggested=("ALT_SUGGESTED","sum"),
        ).reset_index()
        store_sal["Conv%"] = (store_sal["Conversions"]/store_sal["Convs"]*100).round(1)
        store_sal["Lost%"] = (store_sal["Lost"]/store_sal["Convs"]*100).round(1)
        store_sal["OOS%"]  = (store_sal["OOS"]/store_sal["Convs"]*100).round(1)
        store_sal["Upsell_Act%"] = (store_sal["Alt_Suggested"]/store_sal["Upsell_Opps"].replace(0,np.nan)*100).round(1)
        ws6.write(row, 0, "Store", hdr_fmt); ws6.write(row,1,"Country",hdr_fmt); ws6.write(row,2,"Convs",hdr_fmt); ws6.write(row,3,"Conversions",hdr_fmt); ws6.write(row,4,"Conv%",hdr_fmt); ws6.write(row,5,"Lost Sales",hdr_fmt); ws6.write(row,6,"Lost%",hdr_fmt); ws6.write(row,7,"OOS",hdr_fmt); ws6.write(row,8,"OOS%",hdr_fmt); ws6.write(row,9,"Upsell Act%",hdr_fmt); row += 1
        for _, r2 in store_sal.sort_values("Convs",ascending=False).iterrows():
            for ci, val in enumerate([r2["STORE_CODE"],r2["COUNTRY_CODE"],int(r2["Convs"]),int(r2["Conversions"]),f"{r2['Conv%']:.1f}%",int(r2["Lost"]),f"{r2['Lost%']:.1f}%",int(r2["OOS"]),f"{r2['OOS%']:.1f}%",f"{r2['Upsell_Act%']:.0f}%" if not np.isnan(r2['Upsell_Act%']) else "—"]):
                ws6.write(row, ci, str(val) if isinstance(val,str) else val, cell_fmt)
            row += 1

        # Full detail tab
        sal_detail_cols = [c for c in ["CONVERSATION_ID","PLATFORM","STORE_CODE","CHANNEL_NAME","SITE_NICK_NAME_ID","COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME","LAST_MSG_TIME","ISSUE_TYPE","SALES_STAGE","IS_CONVERSION","IS_OOS_CONFIRMED","IS_LOST_SALE","IS_UPSELL_OPP","ALT_SUGGESTED","ITEM_IDS","SIZE_MENTIONS","COLOR_MENTIONS","SENTIMENT","BUYER_SUMMARY"] if c in df.columns]
        write_df_sheet(wb.add_worksheet("💰 Sales Detail"), df[sal_detail_cols], sal_fmt, col_widths=[36,10,10,12,10,8,12,16,16,16,14,11,11,11,11,13,20,14,14,10,50])
        writer.sheets["💰 Sales Detail"] = wb.worksheets()[-1]

        # ════════════════════════════════════════════════════════════════════
        # SHEET 7 : Lost Sales Detail
        # ════════════════════════════════════════════════════════════════════
        lost_df2 = df[df["IS_LOST_SALE"]==1].copy()
        lost_cols = [c for c in ["CONVERSATION_ID","PLATFORM","STORE_CODE","CHANNEL_NAME","SITE_NICK_NAME_ID","COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME","LAST_MSG_TIME","ISSUE_TYPE","IS_OOS_CONFIRMED","ALT_SUGGESTED","ITEM_IDS","SIZE_MENTIONS","COLOR_MENTIONS","SENTIMENT","BUYER_SUMMARY"] if c in lost_df2.columns]
        write_df_sheet(wb.add_worksheet("💸 Lost Sales Detail"), lost_df2[lost_cols] if not lost_df2.empty else pd.DataFrame(columns=lost_cols), sal_fmt, col_widths=[36,10,10,12,10,8,12,16,16,16,11,13,20,14,14,10,50])
        writer.sheets["💸 Lost Sales Detail"] = wb.worksheets()[-1]

        # ════════════════════════════════════════════════════════════════════
        # SHEET 8 : OOS Restock Priority
        # ════════════════════════════════════════════════════════════════════
        oos_df2 = df[df["IS_OOS_CONFIRMED"]==1].copy()
        if not oos_df2.empty:
            oos_agg = oos_df2.groupby(["STORE_CODE","ITEM_IDS"]).agg(
                Demand_Count   = ("CONVERSATION_ID","count"),
                Lost_Sales     = ("IS_LOST_SALE","sum"),
                Alt_Suggested  = ("ALT_SUGGESTED","sum"),
                Countries      = ("COUNTRY_CODE",  lambda x:", ".join(sorted(set(str(v) for v in x if v and str(v)!="nan")))),
                Platforms      = ("PLATFORM",       lambda x:", ".join(sorted(set(str(v) for v in x if v and str(v)!="nan")))),
                Size_Requested = ("SIZE_MENTIONS",  lambda x:" | ".join(sorted(set(v for vals in x for v in str(vals).split("|") if v.strip())))),
                Color_Requested= ("COLOR_MENTIONS", lambda x:" | ".join(sorted(set(v for vals in x for v in str(vals).split("|") if v.strip())))),
            ).reset_index()
            oos_agg["Priority_Score"] = oos_agg["Demand_Count"] + oos_agg["Lost_Sales"]*2
            oos_agg["Alt_Act_%"]      = (oos_agg["Alt_Suggested"]/oos_agg["Demand_Count"]*100).round(0)
            oos_agg = oos_agg.sort_values("Priority_Score", ascending=False).reset_index(drop=True)
            oos_agg.insert(0,"Rank",range(1,len(oos_agg)+1))
            write_df_sheet(wb.add_worksheet("📦 OOS Restock Priority"), oos_agg, mer_fmt, col_widths=[6,10,20,10,10,12,10,12,8,20,20,12,12])
        else:
            ws_oos = wb.add_worksheet("📦 OOS Restock Priority")
            ws_oos.write(0,0,"No OOS inquiries detected in this dataset.",note_fmt)
            writer.sheets["📦 OOS Restock Priority"] = ws_oos

        # ════════════════════════════════════════════════════════════════════
        # SHEET 9 : Product Intelligence
        # ════════════════════════════════════════════════════════════════════
        item_sum, var_df2, _ = build_product_demand(conv_df)
        if not item_sum.empty:
            pi_cols2 = [c for c in ["Item_ID","Total_Inquiries","Unique_Convs","Stores","Countries","Platforms","Sites","Conv_Rate_%","OOS_Confirmed","OOS_Rate_%","Lost_Sales","Upsell_Opps","Sizes_Requested","Colors_Requested"] if c in item_sum.columns]
            write_df_sheet(wb.add_worksheet("🛍️ Product Intelligence"), item_sum[pi_cols2], mer_fmt, col_widths=[18,10,10,14,12,12,16,8,10,8,10,10,30,30])
            writer.sheets["🛍️ Product Intelligence"] = wb.worksheets()[-1]

            # Variation demand sub-table
            if not var_df2.empty:
                ws_var = wb.add_worksheet("📐 Size & Colour Demand")
                writer.sheets["📐 Size & Colour Demand"] = ws_var
                ws_var.write(0,0,"Top Requested Sizes & Colours from Chat Conversations",title_fmt)
                ws_var.write(1,0,"Use for stock planning and listing optimisation — Merch/Buying team",sub_title)
                write_df_sheet(ws_var, var_df2, mer_fmt, start_row=3, col_widths=[20,12,8])
        else:
            ws_pi = wb.add_worksheet("🛍️ Product Intelligence")
            ws_pi.write(0,0,"No item IDs detected in this dataset.",note_fmt)
            writer.sheets["🛍️ Product Intelligence"] = ws_pi

        # ════════════════════════════════════════════════════════════════════
        # SHEET 10 : Team Performance
        # ════════════════════════════════════════════════════════════════════
        team_perf2 = build_team_sales_perf(conv_df)
        if not team_perf2.empty:
            tp_cols = [c for c in ["TEAM_MEMBER","Shift","Conversations","Resolved","Unresolved","CRR_%","Avg_CSAT","Avg_CRT_mins","Conversions","Conv_Rate_%","Upsell_Opps","Alt_Suggested","Upsell_Act_%","Lost_Sales","OOS_Handled","High_Priority","Positive_Sent","Negative_Sent"] if c in team_perf2.columns]
            write_df_sheet(wb.add_worksheet("👥 Team Performance"), team_perf2[tp_cols], hdr_fmt, col_widths=[14,28,10,10,10,8,8,8,10,8,10,12,10,10,10,10,12,12])
            writer.sheets["👥 Team Performance"] = wb.worksheets()[-1]

        # ════════════════════════════════════════════════════════════════════
        # SHEET 11 : AM Store Scorecard
        # ════════════════════════════════════════════════════════════════════
        am_sc = build_am_scorecard(conv_df)
        if not am_sc.empty:
            write_df_sheet(wb.add_worksheet("🏪 AM Store Scorecard"), am_sc, hdr_fmt, col_widths=[12,10,12,10,14,10,10,10,12,10,10,8,8,8,10])
            writer.sheets["🏪 AM Store Scorecard"] = wb.worksheets()[-1]

        # ════════════════════════════════════════════════════════════════════
        # SHEET 12 : Key Improvements
        # ════════════════════════════════════════════════════════════════════
        ws_imp = wb.add_worksheet("🎯 Key Improvements")
        writer.sheets["🎯 Key Improvements"] = ws_imp
        ws_imp.set_column(0,0,14); ws_imp.set_column(1,1,28); ws_imp.set_column(2,2,80); ws_imp.set_column(3,4,14)
        ws_imp.write(0,0,"Key Improvement Areas — Auto-generated from Chat Data",title_fmt)
        ws_imp.write(1,0,"Act on HIGH items this week. Share OPPORTUNITY items with AM/Merch team.",sub_title)
        ws_imp.write(3,0,"Priority",hdr_fmt); ws_imp.write(3,1,"Area",hdr_fmt); ws_imp.write(3,2,"Recommendation",hdr_fmt); ws_imp.write(3,3,"Status",hdr_fmt); ws_imp.write(3,4,"Owner",hdr_fmt)
        improvements = generate_key_improvements(conv_df, funnel)
        for i, (priority_level, area, rec) in enumerate(improvements, start=4):
            fmt_use = red_fmt if "HIGH" in priority_level else (yel_fmt if "MEDIUM" in priority_level else grn_fmt)
            ws_imp.write(i, 0, priority_level, fmt_use)
            ws_imp.write(i, 1, area, cell_fmt)
            ws_imp.write(i, 2, rec, cell_fmt)
            ws_imp.write(i, 3, "Open", cell_fmt)
            ws_imp.write(i, 4, "", cell_fmt)

        # ════════════════════════════════════════════════════════════════════
        # SHEET 13 : Full Conversation Detail
        # ════════════════════════════════════════════════════════════════════
        full_cols = [c for c in ["CONVERSATION_ID","PLATFORM","STORE_CODE","CHANNEL_NAME","SITE_NICK_NAME_ID","COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME","FIRST_MSG_TIME","LAST_MSG_TIME","MSG_COUNT","ISSUE_TYPE","PRIORITY","SENTIMENT","IS_RESOLVED","IS_UNRESOLVED","CSAT_PROXY","AVG_CRT_MINS","SALES_STAGE","IS_CONVERSION","IS_OOS_CONFIRMED","IS_LOST_SALE","IS_UPSELL_OPP","ALT_SUGGESTED","ITEM_IDS","SIZE_MENTIONS","COLOR_MENTIONS","BUYER_SUMMARY","SUGGESTED_REPLY"] if c in df.columns]
        write_df_sheet(wb.add_worksheet("📋 Full Conversation Detail"), df[full_cols], hdr_fmt, col_widths=[36,10,10,12,10,8,12,16,16,16,8,16,8,10,10,10,6,7,14,11,11,11,11,13,20,14,14,50,60])
        writer.sheets["📋 Full Conversation Detail"] = wb.worksheets()[-1]

    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# UI COMPONENTS
# ─────────────────────────────────────────────────────────────────────────────

def render_header():
    st.markdown("""
    <div class="graas-header">
        <div class="graas-logo">📊</div>
        <div>
            <h1>Chat Analyzer Dashboard</h1>
            <p>Turn every buyer chat into 3 signals: <b style="color:#00C4B4">Service</b> · <b style="color:#FF6B35">Sales</b> · <b style="color:#F59E0B">Stock & Merch</b> — Shopee & Lazada</p>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_metrics(conv_df: pd.DataFrame, today_ts: pd.Timestamp):
    total      = len(conv_df)
    resolved   = int(conv_df["IS_RESOLVED"].sum())
    unresolved = int(conv_df["IS_UNRESOLVED"].sum())
    crr        = round(resolved / total * 100, 1) if total else 0
    avg_crt    = conv_df["AVG_CRT_MINS"].mean()
    avg_csat   = conv_df["CSAT_PROXY"].mean()

    # Compare at day granularity using Timestamps (pandas 3.0 safe)
    today_conv  = conv_df[conv_df["LAST_MSG_TIME"].dt.normalize() == today_ts]
    hi_today    = len(today_conv[today_conv["PRIORITY"] == "High"])

    neg_pct = round(len(conv_df[conv_df["SENTIMENT"] == "Negative"]) / total * 100, 1) if total else 0

    cols = st.columns(8)
    metrics = [
        (cols[0], "🗣️ Total Convs", f"{total:,}",   "7-day window", ""),
        (cols[1], "📅 Today",       f"{len(today_conv):,}", "conversations", "navy"),
        (cols[2], "✅ Resolved",    f"{resolved:,}", f"CRR {crr}%", "green"),
        (cols[3], "🔴 Unresolved",  f"{unresolved:,}", "need action", "red"),
        (cols[4], "⚡ CRT",         fmt_mins(avg_crt), "avg response time", "orange"),
        (cols[5], "⭐ CSAT",        f"{avg_csat:.1f}/5" if not np.isnan(avg_csat) else "—", "proxy score", ""),
        (cols[6], "😠 Negative",    f"{neg_pct}%",  "sentiment", "red"),
        (cols[7], "🔥 High Pri",    f"{hi_today}",  "today's urgent", "orange"),
    ]
    for col, label, val, sub, cls in metrics:
        with col:
            st.markdown(f"""
            <div class="metric-card {cls}">
                <div class="metric-label">{label}</div>
                <div class="metric-val">{val}</div>
                <div class="metric-sub">{sub}</div>
            </div>
            """, unsafe_allow_html=True)


def priority_badge(p: str) -> str:
    cls = {"High": "badge-high", "Medium": "badge-medium", "Low": "badge-low"}.get(p, "badge-low")
    return f'<span class="{cls}">{p}</span>'


def sentiment_span(s: str) -> str:
    cls = {"Positive": "sent-pos", "Neutral": "sent-neu", "Negative": "sent-neg"}.get(s, "sent-neu")
    icon = {"Positive": "😊", "Neutral": "😐", "Negative": "😠"}.get(s, "")
    return f'<span class="{cls}">{icon} {s}</span>'


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR FILTERS
# ─────────────────────────────────────────────────────────────────────────────

def apply_filters(conv_df: pd.DataFrame, today_ts: pd.Timestamp) -> pd.DataFrame:
    """
    All filter OPTIONS are drawn from the full conv_df (source data) so they
    never disappear. Selections are then applied to produce the filtered result.
    """
    src = conv_df  # keep full source for populating option lists

    st.sidebar.markdown("## 🔍 Filters")
    st.sidebar.markdown("---")

    # ── Platform — options from full source ───────────────────────────────────
    platforms = ["All"] + sorted(src["PLATFORM"].dropna().unique().tolist())
    sel_platform = st.sidebar.selectbox("🌐 Platform", platforms)

    # ── Date Range — min/max from full source, default = last 7 days ──────────
    _ts_min = src["LAST_MSG_TIME"].dropna().min()
    _ts_max = src["LAST_MSG_TIME"].dropna().max()
    min_date = _ts_min.date() if pd.notna(_ts_min) else datetime.today().date()
    max_date = _ts_max.date() if pd.notna(_ts_max) else datetime.today().date()
    default_start = max(min_date, (today_ts - pd.Timedelta(days=6)).date())

    date_range = st.sidebar.date_input(
        "📅 Date Range",
        value=(default_start, max_date),
        min_value=min_date,
        max_value=max_date,
        help="Default: last 7 days. Expand to compare Jan vs Feb or any custom range.",
    )

    # ── Priority ──────────────────────────────────────────────────────────────
    sel_prio = st.sidebar.selectbox("🔴 Priority", ["All", "High", "Medium", "Low"])

    # ── Sentiment ─────────────────────────────────────────────────────────────
    sel_sent = st.sidebar.selectbox("😊 Sentiment", ["All", "Positive", "Neutral", "Negative"])

    # ── Resolution Status ─────────────────────────────────────────────────────
    sel_res = st.sidebar.selectbox("✅ Resolution Status", ["All", "Resolved", "Unresolved"])

    # ── Issue Type — options from full source ─────────────────────────────────
    issue_opts = ["All"] + sorted(src["ISSUE_TYPE"].dropna().unique().tolist())
    sel_issue = st.sidebar.selectbox("🏷️ Issue Type", issue_opts)

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔎 Search & Filter")

    # ── Team Member — options from full source ────────────────────────────────
    if "TEAM_MEMBER" in src.columns:
        all_agents = sorted(src["TEAM_MEMBER"].dropna().unique().tolist())
        sel_agents = st.sidebar.multiselect("👤 Team Member", all_agents)
    else:
        sel_agents = []

    # ── Store Code — options from full source ─────────────────────────────────
    all_stores = sorted(src["STORE_CODE"].dropna().unique().tolist())
    sel_stores = st.sidebar.multiselect("🏪 Store Code", all_stores)

    # ── Country — options from full source ────────────────────────────────────
    all_countries = sorted(src["COUNTRY_CODE"].dropna().unique().tolist())
    sel_countries = st.sidebar.multiselect("🌍 Country", all_countries)

    # ── Channel Name — options from full source ───────────────────────────────
    if "CHANNEL_NAME" in src.columns:
        all_channels = sorted(src["CHANNEL_NAME"].dropna().replace("", pd.NA).dropna().unique().tolist())
        sel_channels = st.sidebar.multiselect("📡 Channel Name", all_channels)
    else:
        sel_channels = []

    # ── Buyer Name free-text ──────────────────────────────────────────────────
    buyer_search = st.sidebar.text_input("🔍 Buyer Name")

    # ── Conversation ID free-text ─────────────────────────────────────────────
    conv_search = st.sidebar.text_input("🔍 Conversation ID")

    # ── Apply all filters to source ───────────────────────────────────────────
    result = src.copy()

    if sel_platform != "All":
        result = result[result["PLATFORM"] == sel_platform]

    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start_ts = pd.Timestamp(date_range[0])
        end_ts   = pd.Timestamp(date_range[1]) + pd.Timedelta(hours=23, minutes=59, seconds=59)
        result = result[
            (result["LAST_MSG_TIME"] >= start_ts) &
            (result["LAST_MSG_TIME"] <= end_ts)
        ]

    if sel_prio != "All":
        result = result[result["PRIORITY"] == sel_prio]

    if sel_sent != "All":
        result = result[result["SENTIMENT"] == sel_sent]

    if sel_res == "Resolved":
        result = result[result["IS_RESOLVED"]]
    elif sel_res == "Unresolved":
        result = result[result["IS_UNRESOLVED"]]

    if sel_issue != "All":
        result = result[result["ISSUE_TYPE"] == sel_issue]

    if sel_agents:
        result = result[result["TEAM_MEMBER"].isin(sel_agents)]

    if sel_stores:
        result = result[result["STORE_CODE"].isin(sel_stores)]

    if sel_countries:
        result = result[result["COUNTRY_CODE"].isin(sel_countries)]

    if sel_channels and "CHANNEL_NAME" in result.columns:
        result = result[result["CHANNEL_NAME"].isin(sel_channels)]

    if buyer_search:
        result = result[result["BUYER_NAME"].str.contains(buyer_search, case=False, na=False)]

    if conv_search:
        result = result[result["CONVERSATION_ID"].str.contains(conv_search, case=False, na=False)]

    st.sidebar.markdown("---")
    total = len(result)
    st.sidebar.markdown(f"**{total:,}** of **{len(src):,}** conversations")
    if total == 0:
        st.sidebar.warning("No results — try widening the date range or clearing filters.")

    # ── Cache / reload controls ───────────────────────────────────────────────
    st.sidebar.markdown("---")
    st.sidebar.markdown("### ⚙️ Controls")
    if st.sidebar.button("🔄 Clear Cache & Reload", use_container_width=True,
                          help="Clears all cached data. Use when uploading a new file or if the app seems stuck."):
        st.cache_data.clear()
        st.rerun()

    return result


# ─────────────────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────────────────

def main():
    render_header()

    # ── File Upload ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📂 Upload Chat Data</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Upload Excel file with sheets: lazada_chat_enquiries & shopee_chat_enquiries",
        type=["xlsx"],
        help="Single Excel file containing both Lazada and Shopee chat sheets.",
    )

    if not uploaded:
        st.info("👆 Upload your chat enquiries Excel file to get started.")
        st.markdown("""
        **Expected Excel format:**
        - Sheet 1: `lazada_chat_enquiries`
        - Sheet 2: `shopee_chat_enquiries`
        - Columns: `STORE_CODE`, `SITE_NICK_NAME_ID`, `COUNTRY_CODE`, `CONVERSATION_ID`,
          `IS_READ`, `IS_ANSWERED`, `MESSAGE_TIME`, `BUYER_NAME`, `MESSAGE_PARSED`,
          `MESSAGE_TYPE`, `MESSAGE_ID`, `SENDER`, `BUYER_ID`
        """)
        return

    # ── Load ALL Data ─────────────────────────────────────────────────────────
    with st.spinner("⏳ Loading chat data…"):
        raw_df = load_data(uploaded.read())

    # pandas 3.0 fix: use .dropna().max() on Timestamp series, never .dt.date.max()
    _max_ts = raw_df["MESSAGE_TIME"].dropna().max()
    _min_ts = raw_df["MESSAGE_TIME"].dropna().min()
    today_date = _max_ts.date() if pd.notna(_max_ts) else datetime.today().date()
    today_str  = today_date.strftime("%Y-%m-%d")
    today_ts   = pd.Timestamp(today_date)
    data_start = _min_ts.date() if pd.notna(_min_ts) else today_date

    st.success(
        f"✅ Loaded **{len(raw_df):,}** messages · "
        f"**{raw_df['CONVERSATION_ID'].nunique():,}** conversations · "
        f"**{raw_df['PLATFORM'].nunique()}** platforms · "
        f"Data range: **{data_start}** → **{today_date}**"
    )

    # ── Analyse ALL conversations (full dataset) ───────────────────────────────
    with st.spinner("🔍 Analysing conversations — this runs once and is cached…"):
        conv_df = analyse(raw_df)
    del raw_df; gc.collect()   # free raw messages from memory immediately after analysis

    # ── Sidebar Filters (default = last 7 days view) ───────────────────────────
    conv_filtered = apply_filters(conv_df, today_ts)

    if conv_filtered.empty:
        st.warning("No conversations match the current filters.")
        return

    # ── Metrics Row ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📈 Key Metrics</div>', unsafe_allow_html=True)
    render_metrics(conv_filtered, today_ts)

    # ── Charts Row ────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📊 Analytics</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)

    with c1:
        issue_counts = conv_filtered["ISSUE_TYPE"].value_counts().reset_index()
        issue_counts.columns = ["Issue Type", "Count"]
        st.markdown("**Issue Type Distribution**")
        st.bar_chart(issue_counts.set_index("Issue Type")["Count"], color="#00C4B4")

    with c2:
        sent_counts = conv_filtered["SENTIMENT"].value_counts().reset_index()
        sent_counts.columns = ["Sentiment", "Count"]
        st.markdown("**Sentiment Breakdown**")
        color_map = {"Positive": "#27AE60", "Neutral": "#7F8C8D", "Negative": "#E74C3C"}
        st.bar_chart(sent_counts.set_index("Sentiment")["Count"])

    with c3:
        # Use dt.normalize() (returns Timestamp at midnight) — pandas 3.0 safe
        daily = (
            conv_filtered
            .assign(DATE=conv_filtered["LAST_MSG_TIME"].dt.normalize())
            .groupby("DATE")
            .size()
            .reset_index(name="Conversations")
        )
        st.markdown("**Daily Conversation Volume**")
        st.line_chart(daily.set_index("DATE")["Conversations"], color="#FF6B35")

    # ── Tabs ─────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs([
        "🔥 Today's Priority Chats",
        "📋 All Conversations",
        "🔴 Unresolved Chats",
        "💬 Suggested Replies",
        "📈 WoW / MoM Performance",
        "👥 Team Performance",
        "💰 Sales Intelligence",
        "🛍️ AM & Merch Performance",
        "🎯 Key Improvements",
        "📦 Product Intelligence",
    ])

    display_cols = [
        "CONVERSATION_ID", "PLATFORM", "STORE_CODE", "CHANNEL_NAME",
        "SITE_NICK_NAME_ID", "COUNTRY_CODE", "BUYER_NAME",
        "ISSUE_TYPE", "PRIORITY", "SENTIMENT", "IS_UNRESOLVED",
        "CSAT_PROXY", "AVG_CRT_MINS", "BUYER_SUMMARY",
    ]
    # Only include cols that actually exist in the dataframe
    display_cols = [c for c in display_cols if c in conv_filtered.columns]

    with tab1:
        # Use normalize() for date comparison — pandas 3.0 safe
        today_df = conv_filtered[conv_filtered["LAST_MSG_TIME"].dt.normalize() == today_ts]
        today_sorted = today_df.sort_values(
            "PRIORITY",
            key=lambda s: s.map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
        )
        st.markdown(f"**{len(today_sorted)} conversations today** — sorted by priority")
        if today_sorted.empty:
            st.info("No conversations found for today.")
        else:
            st.dataframe(
                today_sorted[display_cols].reset_index(drop=True),
                use_container_width=True,
                height=450,
                column_config={
                    "CSAT_PROXY":   st.column_config.NumberColumn("CSAT (1-5)", format="%.1f"),
                    "AVG_CRT_MINS": st.column_config.NumberColumn("CRT (mins)", format="%.0f"),
                    "IS_UNRESOLVED": st.column_config.CheckboxColumn("Unresolved?"),
                    "BUYER_SUMMARY": st.column_config.TextColumn("Summary", width="large"),
                },
            )

    with tab2:
        all_sorted = conv_filtered.sort_values("LAST_MSG_TIME", ascending=False)
        st.markdown(f"**{len(all_sorted)} conversations** in filtered view")
        st.dataframe(
            all_sorted[display_cols].reset_index(drop=True),
            use_container_width=True,
            height=500,
            column_config={
                "CSAT_PROXY":   st.column_config.NumberColumn("CSAT (1-5)", format="%.1f"),
                "AVG_CRT_MINS": st.column_config.NumberColumn("CRT (mins)", format="%.0f"),
                "IS_UNRESOLVED": st.column_config.CheckboxColumn("Unresolved?"),
                "BUYER_SUMMARY": st.column_config.TextColumn("Summary", width="large"),
            },
        )

    with tab3:
        unres_df = conv_filtered[conv_filtered["IS_UNRESOLVED"]].sort_values(
            "PRIORITY",
            key=lambda s: s.map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
        )
        st.markdown(
            f"**{len(unres_df)} unresolved conversations** — contain stalling phrases without resolution"
        )
        if unres_df.empty:
            st.success("🎉 No unresolved conversations found!")
        else:
            st.dataframe(
                unres_df[display_cols].reset_index(drop=True),
                use_container_width=True,
                height=450,
                column_config={
                    "CSAT_PROXY":   st.column_config.NumberColumn("CSAT (1-5)", format="%.1f"),
                    "AVG_CRT_MINS": st.column_config.NumberColumn("CRT (mins)", format="%.0f"),
                    "IS_UNRESOLVED": st.column_config.CheckboxColumn("Unresolved?"),
                    "BUYER_SUMMARY": st.column_config.TextColumn("Summary", width="large"),
                },
            )

    with tab4:
        st.markdown("### 💬 Suggested Reply Templates by Issue Type")
        st.caption(
            "Empathetic, resolution-oriented replies — replace [PLACEHOLDERS] before sending."
        )
        for issue_type, reply_text in SUGGESTED_REPLIES.items():
            if issue_type == "Other":
                continue
            priority = get_priority(issue_type)
            badge_color = {"High": "🔴", "Medium": "🟡", "Low": "🔵"}.get(priority, "⚪")
            with st.expander(f"{badge_color} {issue_type}  ({priority} Priority)"):
                st.markdown(f"""
                <div class="reply-label">Suggested Reply</div>
                <div class="reply-box">{reply_text}</div>
                """, unsafe_allow_html=True)

        # Per-conversation reply lookup
        st.markdown("---")
        st.markdown("### 🔍 Look Up Reply for a Specific Conversation")
        conv_ids = conv_filtered["CONVERSATION_ID"].tolist()
        if conv_ids:
            sel_conv = st.selectbox("Select Conversation ID", conv_ids[:500])
            row = conv_filtered[conv_filtered["CONVERSATION_ID"] == sel_conv].iloc[0]
            st.markdown(f"""
            **Issue Type:** {row['ISSUE_TYPE']}  |
            **Priority:** {row['PRIORITY']}  |
            **Sentiment:** {row['SENTIMENT']}  |
            **CSAT Proxy:** {row['CSAT_PROXY']}
            """)
            st.markdown(f"""
            <div class="reply-label">Buyer Summary</div>
            <div class="reply-box">{row['BUYER_SUMMARY']}</div>
            """, unsafe_allow_html=True)
            suggested = SUGGESTED_REPLIES.get(str(row['ISSUE_TYPE']), SUGGESTED_REPLIES["Other"])
            st.markdown(f"""
            <div class="reply-label">Suggested Reply</div>
            <div class="reply-box">{suggested}</div>
            """, unsafe_allow_html=True)

    # ── Tab 5 : WoW / MoM Performance ────────────────────────────────────────
    with tab5:
        st.markdown("### 📈 Week-on-Week & Month-on-Month Performance")
        wow_df, mom_df = compute_wow_mom(conv_filtered)

        wow_tab, mom_tab = st.tabs(["📅 Week-on-Week", "🗓️ Month-on-Month"])

        with wow_tab:
            if wow_df.empty:
                st.info("Not enough data for weekly comparison.")
            else:
                st.markdown("**Weekly Conversation Trend**")
                wow_chart = wow_df.set_index("WEEK")[["Conversations"]].copy()
                st.bar_chart(wow_chart, color="#00C4B4")

                st.markdown("**Weekly Metrics Table**")
                disp_wow = wow_df.copy()
                disp_wow["WEEK"] = disp_wow["WEEK"].dt.strftime("%d %b %Y")
                disp_wow["Avg_CRT_mins"] = disp_wow["Avg_CRT_mins"].apply(
                    lambda x: fmt_mins(x) if pd.notna(x) else "—"
                )
                st.dataframe(
                    disp_wow[["WEEK","Conversations","CRR_%","Avg_CSAT",
                               "Avg_CRT_mins","Conversions",
                               "Δ Conversations","Δ CRR_%","Δ Avg_CSAT"]].reset_index(drop=True),
                    use_container_width=True,
                    column_config={
                        "WEEK":           st.column_config.TextColumn("Week Starting"),
                        "CRR_%":          st.column_config.NumberColumn("CRR %", format="%.1f%%"),
                        "Avg_CSAT":       st.column_config.NumberColumn("CSAT", format="%.2f"),
                        "Conversions":    st.column_config.NumberColumn("Conversions"),
                        "Δ Conversations":st.column_config.NumberColumn("Δ Conv", format="%+.0f"),
                        "Δ CRR_%":        st.column_config.NumberColumn("Δ CRR%", format="%+.1f"),
                        "Δ Avg_CSAT":     st.column_config.NumberColumn("Δ CSAT", format="%+.2f"),
                    },
                )

        with mom_tab:
            if mom_df.empty:
                st.info("Not enough data for monthly comparison.")
            else:
                st.markdown("**Monthly Conversation Trend**")
                mom_chart = mom_df.set_index("MONTH")[["Conversations"]].copy()
                st.bar_chart(mom_chart, color="#FF6B35")

                st.markdown("**Monthly Metrics Table**")
                disp_mom = mom_df.copy()
                disp_mom["MONTH"] = disp_mom["MONTH"].dt.strftime("%b %Y")
                disp_mom["Avg_CRT_mins"] = disp_mom["Avg_CRT_mins"].apply(
                    lambda x: fmt_mins(x) if pd.notna(x) else "—"
                )
                st.dataframe(
                    disp_mom[["MONTH","Conversations","CRR_%","Avg_CSAT",
                               "Avg_CRT_mins","Conversions",
                               "Δ Conversations","Δ CRR_%","Δ Avg_CSAT"]].reset_index(drop=True),
                    use_container_width=True,
                    column_config={
                        "MONTH":          st.column_config.TextColumn("Month"),
                        "CRR_%":          st.column_config.NumberColumn("CRR %", format="%.1f%%"),
                        "Avg_CSAT":       st.column_config.NumberColumn("CSAT", format="%.2f"),
                        "Conversions":    st.column_config.NumberColumn("Conversions"),
                        "Δ Conversations":st.column_config.NumberColumn("Δ Conv", format="%+.0f"),
                        "Δ CRR_%":        st.column_config.NumberColumn("Δ CRR%", format="%+.1f"),
                        "Δ Avg_CSAT":     st.column_config.NumberColumn("Δ CSAT", format="%+.2f"),
                    },
                )

    # ── Tab 6 : Team Member Performance ──────────────────────────────────────
    with tab6:
        st.markdown("### 👥 Team Member Performance")
        st.caption(
            f"Data from **{TEAM_START_DATE.strftime('%d %b %Y')}** onwards · "
            f"Store → Agent mapping as configured in constants"
        )

        team_perf = compute_team_performance(conv_filtered)

        if team_perf.empty:
            st.info(
                "No team performance data available. "
                "This may be because no conversations fall within the tracking period "
                f"(from {TEAM_START_DATE.strftime('%d %b %Y')}) or store codes don't match assignments."
            )
        else:
            # ── KPI scorecards per agent ──────────────────────────────────────
            agents = team_perf["TEAM_MEMBER"].tolist()
            agents_per_row = 3
            for i in range(0, len(agents), agents_per_row):
                cols = st.columns(agents_per_row)
                for j, agent in enumerate(agents[i:i+agents_per_row]):
                    row_a = team_perf[team_perf["TEAM_MEMBER"] == agent].iloc[0]
                    with cols[j]:
                        st.markdown(
                            f"""
                            <div style="background:#1B2A4A;border-radius:10px;padding:14px;color:white;margin-bottom:8px;">
                              <div style="font-size:16px;font-weight:700;color:#00C4B4;">👤 {agent}</div>
                              <div style="font-size:11px;color:#aaa;margin-bottom:8px;">{row_a['Shift']}</div>
                              <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;">
                                <div><span style="font-size:20px;font-weight:700">{int(row_a['Conversations'])}</span><br><span style="font-size:11px;color:#ccc;">Conversations</span></div>
                                <div><span style="font-size:20px;font-weight:700">{row_a['CRR_%']:.1f}%</span><br><span style="font-size:11px;color:#ccc;">CRR</span></div>
                                <div><span style="font-size:20px;font-weight:700">{row_a['Avg_CSAT']:.2f}</span><br><span style="font-size:11px;color:#ccc;">CSAT</span></div>
                                <div><span style="font-size:20px;font-weight:700">{int(row_a['Avg_CRT_mins']) if pd.notna(row_a['Avg_CRT_mins']) else '—'}m</span><br><span style="font-size:11px;color:#ccc;">Avg CRT</span></div>
                                <div><span style="font-size:20px;font-weight:700;color:#FF6B35">{int(row_a['Conversions'])}</span><br><span style="font-size:11px;color:#ccc;">Conversions</span></div>
                                <div><span style="font-size:20px;font-weight:700;color:#f87171">{int(row_a['High_Priority'])}</span><br><span style="font-size:11px;color:#ccc;">High Priority</span></div>
                              </div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )

            st.markdown("---")

            # ── Summary table ─────────────────────────────────────────────────
            st.markdown("**Team Summary Table**")
            summary_cols = [
                "TEAM_MEMBER", "Shift", "Conversations", "Resolved", "Unresolved",
                "CRR_%", "Avg_CSAT", "Avg_CRT_mins", "Positive_Sent",
                "Negative_Sent", "Conversions", "High_Priority",
            ]
            st.dataframe(
                team_perf[summary_cols].reset_index(drop=True),
                use_container_width=True,
                column_config={
                    "TEAM_MEMBER":   st.column_config.TextColumn("Agent"),
                    "Shift":         st.column_config.TextColumn("Shift / Market"),
                    "Conversations": st.column_config.NumberColumn("Conv"),
                    "Resolved":      st.column_config.NumberColumn("Resolved"),
                    "Unresolved":    st.column_config.NumberColumn("Unresolved"),
                    "CRR_%":         st.column_config.NumberColumn("CRR %", format="%.1f%%"),
                    "Avg_CSAT":      st.column_config.NumberColumn("CSAT", format="%.2f"),
                    "Avg_CRT_mins":  st.column_config.NumberColumn("CRT (min)", format="%.0f"),
                    "Positive_Sent": st.column_config.NumberColumn("Positive"),
                    "Negative_Sent": st.column_config.NumberColumn("Negative"),
                    "Conversions":   st.column_config.NumberColumn("Conversions"),
                    "High_Priority": st.column_config.NumberColumn("High Pri."),
                },
            )

            # ── Per-agent drilldown ───────────────────────────────────────────
            st.markdown("---")
            st.markdown("**Drill Down by Agent**")
            agent_sel = st.selectbox("Select Agent", ["(All)"] + agents)
            if agent_sel == "(All)":
                drilldown_df = conv_filtered[conv_filtered["LAST_MSG_TIME"] >= TEAM_START_DATE]
            else:
                drilldown_df = conv_filtered[
                    (conv_filtered["TEAM_MEMBER"] == agent_sel) &
                    (conv_filtered["LAST_MSG_TIME"] >= TEAM_START_DATE)
                ]

            # If "Others" selected, show a sub-breakdown by store code
            if agent_sel == "Others" and not drilldown_df.empty:
                st.markdown("**Others — Store Code Breakdown**")
                others_summary = (
                    drilldown_df.groupby("STORE_CODE")
                    .agg(
                        Conversations=("CONVERSATION_ID", "count"),
                        Unresolved=("IS_UNRESOLVED", "sum"),
                        Avg_CSAT=("CSAT_PROXY", "mean"),
                        Platform=("PLATFORM", lambda x: x.mode().iloc[0] if not x.empty else "—"),
                        Country=("COUNTRY_CODE", lambda x: x.mode().iloc[0] if not x.empty else "—"),
                    )
                    .reset_index()
                    .sort_values("Conversations", ascending=False)
                )
                others_summary["Avg_CSAT"] = others_summary["Avg_CSAT"].round(1)
                others_summary["CRR%"] = (
                    (others_summary["Conversations"] - others_summary["Unresolved"])
                    / others_summary["Conversations"] * 100
                ).round(1)
                st.dataframe(others_summary, use_container_width=True, hide_index=True,
                    column_config={
                        "Conversations": st.column_config.NumberColumn("Conv"),
                        "Unresolved":    st.column_config.NumberColumn("Unresolved"),
                        "Avg_CSAT":      st.column_config.NumberColumn("CSAT", format="%.1f"),
                        "CRR%":          st.column_config.NumberColumn("CRR %", format="%.1f%%"),
                    }
                )
                st.markdown("**All Conversations — Others**")

            drill_cols = [
                "CONVERSATION_ID", "STORE_CODE", "CHANNEL_NAME",
                "SITE_NICK_NAME_ID", "COUNTRY_CODE",
                "BUYER_NAME", "LAST_MSG_TIME", "ISSUE_TYPE", "PRIORITY",
                "SENTIMENT", "IS_RESOLVED", "CSAT_PROXY", "AVG_CRT_MINS",
                "IS_CONVERSION", "TEAM_MEMBER",
            ]
            available_drill = [c for c in drill_cols if c in drilldown_df.columns]
            st.dataframe(
                drilldown_df[available_drill].sort_values("LAST_MSG_TIME", ascending=False).reset_index(drop=True),
                use_container_width=True,
                height=400,
                column_config={
                    "CSAT_PROXY":    st.column_config.NumberColumn("CSAT", format="%.1f"),
                    "AVG_CRT_MINS":  st.column_config.NumberColumn("CRT(m)", format="%.0f"),
                    "IS_RESOLVED":   st.column_config.CheckboxColumn("Resolved?"),
                    "IS_CONVERSION": st.column_config.CheckboxColumn("Conversion?"),
                },
            )

            # ── Others — stores not assigned to any agent ─────────────────────
            st.markdown("---")
            st.markdown("**🔍 Unassigned / Other Stores in This Data**")
            st.caption("Based on current sidebar filters — change filters to see different stores.")
            all_known = set(STORE_TO_AGENT.keys())
            if "STORE_CODE" in conv_filtered.columns:
                others_stores = sorted(
                    s for s in conv_filtered["STORE_CODE"].dropna().unique()
                    if str(s).strip().upper() not in all_known and str(s).strip()
                )
                if others_stores:
                    others_rows = []
                    for sc in others_stores:
                        sc_df = conv_filtered[conv_filtered["STORE_CODE"] == sc]
                        # Most common site nickname, channel & country for this store
                        site = (
                            sc_df["SITE_NICK_NAME_ID"].mode().iloc[0]
                            if "SITE_NICK_NAME_ID" in sc_df.columns and not sc_df["SITE_NICK_NAME_ID"].dropna().empty
                            else "—"
                        )
                        channel = (
                            sc_df["CHANNEL_NAME"].mode().iloc[0]
                            if "CHANNEL_NAME" in sc_df.columns and not sc_df["CHANNEL_NAME"].replace("", pd.NA).dropna().empty
                            else "—"
                        )
                        country = (
                            sc_df["COUNTRY_CODE"].mode().iloc[0]
                            if "COUNTRY_CODE" in sc_df.columns and not sc_df["COUNTRY_CODE"].dropna().empty
                            else "—"
                        )
                        platform = (
                            sc_df["PLATFORM"].mode().iloc[0]
                            if "PLATFORM" in sc_df.columns and not sc_df["PLATFORM"].dropna().empty
                            else "—"
                        )
                        others_rows.append({
                            "Store Code":      sc,
                            "Channel Name":    channel,
                            "Site Nickname":   site,
                            "Platform":        platform,
                            "Country":         country,
                            "Conversations":   len(sc_df),
                            "Unresolved":      int(sc_df["IS_UNRESOLVED"].sum()) if "IS_UNRESOLVED" in sc_df.columns else 0,
                            "Avg CSAT":        round(sc_df["CSAT_PROXY"].mean(), 1) if "CSAT_PROXY" in sc_df.columns else "—",
                            "Assign To":       "⚠️ Not assigned",
                        })
                    others_df = pd.DataFrame(others_rows).sort_values("Conversations", ascending=False)
                    st.dataframe(others_df, use_container_width=True, hide_index=True,
                        column_config={
                            "Conversations": st.column_config.NumberColumn("Conv", format="%d"),
                            "Unresolved":    st.column_config.NumberColumn("Unresolved", format="%d"),
                            "Avg CSAT":      st.column_config.NumberColumn("CSAT", format="%.1f"),
                        }
                    )
                    st.warning(
                        f"⚠️ **{len(others_stores)} store(s)** found with no team member assigned. "
                        "Share these store codes to get them added to the team assignments."
                    )
                else:
                    st.success("✅ All stores in this dataset are assigned to team members.")

            # ── Store assignments reference ───────────────────────────────────
            with st.expander("📋 Store → Agent Assignment Reference"):
                assign_rows = []
                for agent_name, stores in TEAM_ASSIGNMENTS.items():
                    assign_rows.append({
                        "Agent":           agent_name,
                        "Shift":           AGENT_SHIFT.get(agent_name, "Day"),
                        "Assigned Stores": ", ".join(stores),
                    })
                st.dataframe(pd.DataFrame(assign_rows), use_container_width=True, hide_index=True)



    # ── shared drilldown columns ─────────────────────────────────────────────
    _DRILL_BASE = [c for c in [
        "CONVERSATION_ID","PLATFORM","STORE_CODE","CHANNEL_NAME",
        "SITE_NICK_NAME_ID","COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME",
        "LAST_MSG_TIME","ISSUE_TYPE","PRIORITY","SENTIMENT",
        "IS_UNRESOLVED","CSAT_PROXY","AVG_CRT_MINS",
        "SALES_STAGE","IS_CONVERSION","IS_OOS_CONFIRMED",
        "IS_LOST_SALE","IS_UPSELL_OPP","ALT_SUGGESTED",
        "ITEM_IDS","SIZE_MENTIONS","COLOR_MENTIONS","BUYER_SUMMARY",
    ] if c in conv_filtered.columns]
    _DRILL_CFG = {
        "CSAT_PROXY":      st.column_config.NumberColumn("CSAT",       format="%.1f"),
        "AVG_CRT_MINS":    st.column_config.NumberColumn("CRT(m)",     format="%.0f"),
        "IS_UNRESOLVED":   st.column_config.CheckboxColumn("Unresolved?"),
        "IS_CONVERSION":   st.column_config.CheckboxColumn("Converted?"),
        "IS_OOS_CONFIRMED":st.column_config.CheckboxColumn("OOS?"),
        "IS_LOST_SALE":    st.column_config.CheckboxColumn("Lost Sale?"),
        "IS_UPSELL_OPP":   st.column_config.CheckboxColumn("Upsell Opp?"),
        "ALT_SUGGESTED":   st.column_config.CheckboxColumn("Alt Suggested?"),
        "BUYER_SUMMARY":   st.column_config.TextColumn("Summary",      width="large"),
        "ITEM_IDS":        st.column_config.TextColumn("Item IDs"),
        "SIZE_MENTIONS":   st.column_config.TextColumn("Sizes"),
        "COLOR_MENTIONS":  st.column_config.TextColumn("Colours"),
        "LAST_MSG_TIME":   st.column_config.DatetimeColumn("Last Msg",  format="YYYY-MM-DD HH:mm"),
    }

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 7 : SALES INTELLIGENCE  — every KPI is drillable
    # ══════════════════════════════════════════════════════════════════════════
    with tab7:
        st.markdown("### 💰 Sales Intelligence")
        st.caption("Click **▶ View Details** under any metric to see the full supporting chat data.")

        funnel = build_sales_funnel(conv_filtered)

        # ── bool-safe copy for aggregations ──────────────────────────────────
        _cf = conv_filtered.copy()
        for _bc in ["IS_CONVERSION","IS_LOST_SALE","IS_OOS_CONFIRMED","IS_UPSELL_OPP","ALT_SUGGESTED","IS_UNRESOLVED"]:
            if _bc in _cf.columns:
                _cf[_bc] = _cf[_bc].astype(bool).astype(int)

        # ── Subsets ───────────────────────────────────────────────────────────
        _conv_rows   = conv_filtered[conv_filtered["IS_CONVERSION"].astype(bool)]
        _oos_rows    = conv_filtered[conv_filtered["IS_OOS_CONFIRMED"].astype(bool)]
        _lost_rows   = conv_filtered[conv_filtered["IS_LOST_SALE"].astype(bool)]
        _upsell_rows = conv_filtered[conv_filtered["IS_UPSELL_OPP"].astype(bool)]
        _upsell_missed = _upsell_rows[~_upsell_rows["ALT_SUGGESTED"].astype(bool)]
        _prod_rows   = conv_filtered[conv_filtered["ISSUE_TYPE"].astype(str) == "Product Inquiry"]

        # ── KPI row ───────────────────────────────────────────────────────────
        k1,k2,k3,k4,k5 = st.columns(5)
        with k1:
            st.markdown(f"""<div class="metric-card green">
                <div class="metric-label">💰 Conversions</div>
                <div class="metric-val">{funnel.get('converted',0)}</div>
                <div class="metric-sub">Conv rate: {funnel.get('conv_rate',0)}%</div>
            </div>""", unsafe_allow_html=True)
        with k2:
            st.markdown(f"""<div class="metric-card orange">
                <div class="metric-label">🔁 Upsell Opps</div>
                <div class="metric-val">{funnel.get('upsell_opp',0)}</div>
                <div class="metric-sub">Act rate: {funnel.get('upsell_act_rate',0)}%</div>
            </div>""", unsafe_allow_html=True)
        with k3:
            st.markdown(f"""<div class="metric-card red">
                <div class="metric-label">📦 OOS Inquiries</div>
                <div class="metric-val">{funnel.get('oos_total',0)}</div>
                <div class="metric-sub">Demand captured</div>
            </div>""", unsafe_allow_html=True)
        with k4:
            st.markdown(f"""<div class="metric-card red">
                <div class="metric-label">💸 Lost Sales</div>
                <div class="metric-val">{funnel.get('lost',0)}</div>
                <div class="metric-sub">Lost rate: {funnel.get('lost_rate',0):.1f}%</div>
            </div>""", unsafe_allow_html=True)
        with k5:
            st.markdown(f"""<div class="metric-card navy">
                <div class="metric-label">🛍️ Product Inquiries</div>
                <div class="metric-val">{funnel.get('prod_inq',0)}</div>
                <div class="metric-sub">Warm leads</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("---")

        # ── DRILLDOWN 1 : Conversions ─────────────────────────────────────────
        with st.expander(f"💰 View {len(_conv_rows)} Conversion Conversations", expanded=False):
            st.caption("Conversations where buyer expressed clear purchase intent or placed an order.")
            if not _conv_rows.empty:
                st.dataframe(_conv_rows[_DRILL_BASE].sort_values("LAST_MSG_TIME", ascending=False).reset_index(drop=True),
                    use_container_width=True, height=400, column_config=_DRILL_CFG)
                _c1,_c2,_c3 = st.columns(3)
                with _c1:
                    st.markdown("**By Store**")
                    st.dataframe(_cf[_cf["IS_CONVERSION"]==1].groupby("STORE_CODE").agg(Conversions=("CONVERSATION_ID","count"),Avg_CSAT=("CSAT_PROXY","mean")).reset_index().sort_values("Conversions",ascending=False), use_container_width=True, hide_index=True)
                with _c2:
                    st.markdown("**By Country**")
                    st.dataframe(_cf[_cf["IS_CONVERSION"]==1].groupby("COUNTRY_CODE").agg(Conversions=("CONVERSATION_ID","count")).reset_index().sort_values("Conversions",ascending=False), use_container_width=True, hide_index=True)
                with _c3:
                    st.markdown("**By Platform**")
                    st.dataframe(_cf[_cf["IS_CONVERSION"]==1].groupby("PLATFORM").agg(Conversions=("CONVERSATION_ID","count")).reset_index().sort_values("Conversions",ascending=False), use_container_width=True, hide_index=True)
            else:
                st.info("No conversions detected.")

        # ── DRILLDOWN 2 : OOS Inquiries ───────────────────────────────────────
        with st.expander(f"📦 View {len(_oos_rows)} OOS Inquiry Conversations", expanded=False):
            st.caption("Conversations where seller confirmed item is out of stock. Use to build restock priority list.")
            if not _oos_rows.empty:
                # Summary by item ID + store
                _oos_agg = _cf[_cf["IS_OOS_CONFIRMED"]==1].groupby(["STORE_CODE","COUNTRY_CODE","ITEM_IDS"]).agg(
                    Demand_Count   = ("CONVERSATION_ID","count"),
                    Lost_Sales     = ("IS_LOST_SALE","sum"),
                    Alt_Suggested  = ("ALT_SUGGESTED","sum"),
                    Size_Requested = ("SIZE_MENTIONS",  lambda x: " | ".join(sorted(set(v for vals in x for v in str(vals).split("|") if v.strip())))),
                    Color_Requested= ("COLOR_MENTIONS", lambda x: " | ".join(sorted(set(v for vals in x for v in str(vals).split("|") if v.strip())))),
                ).reset_index().rename(columns={"ITEM_IDS":"Item IDs"})
                _oos_agg["Priority Score"] = _oos_agg["Demand_Count"] + _oos_agg["Lost_Sales"]*2
                _oos_agg["Alt_Act_%"] = (_oos_agg["Alt_Suggested"] / _oos_agg["Demand_Count"]*100).round(0)
                _oos_agg = _oos_agg.sort_values("Priority Score", ascending=False).reset_index(drop=True)
                st.markdown("**🔢 OOS Restock Priority — by Item ID, Store & Demand**")
                st.caption("Priority Score = Demand Count + (Lost Sales × 2). Share with buying team.")
                max_ps = int(_oos_agg["Priority Score"].max()) if len(_oos_agg) else 1
                st.dataframe(_oos_agg, use_container_width=True, hide_index=True,
                    column_config={
                        "Priority Score": st.column_config.ProgressColumn(format="%d", min_value=0, max_value=max_ps),
                        "Lost_Sales":     st.column_config.NumberColumn("Lost Sales"),
                        "Alt_Act_%":      st.column_config.NumberColumn("Alt Suggested %", format="%.0f%%"),
                    })
                st.markdown("---")
                st.markdown("**📋 Full OOS Conversation Detail**")
                _oos_detail_cols = [c for c in [
                    "CONVERSATION_ID","STORE_CODE","CHANNEL_NAME","SITE_NICK_NAME_ID",
                    "COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME","LAST_MSG_TIME",
                    "ITEM_IDS","SIZE_MENTIONS","COLOR_MENTIONS",
                    "IS_LOST_SALE","ALT_SUGGESTED","SENTIMENT","BUYER_SUMMARY"
                ] if c in _oos_rows.columns]
                st.dataframe(_oos_rows[_oos_detail_cols].sort_values("LAST_MSG_TIME", ascending=False).reset_index(drop=True),
                    use_container_width=True, height=380, column_config={
                        "IS_LOST_SALE":  st.column_config.CheckboxColumn("Lost Sale?"),
                        "ALT_SUGGESTED": st.column_config.CheckboxColumn("Alt Suggested?"),
                        "LAST_MSG_TIME": st.column_config.DatetimeColumn("Last Msg", format="YYYY-MM-DD HH:mm"),
                        "BUYER_SUMMARY": st.column_config.TextColumn("Summary", width="large"),
                        "ITEM_IDS":      st.column_config.TextColumn("Item IDs"),
                        "SIZE_MENTIONS": st.column_config.TextColumn("Sizes"),
                        "COLOR_MENTIONS":st.column_config.TextColumn("Colours"),
                    })
            else:
                st.info("No OOS inquiries detected.")

        # ── DRILLDOWN 3 : Lost Sales ──────────────────────────────────────────
        with st.expander(f"💸 View {len(_lost_rows)} Lost Sale Conversations", expanded=False):
            st.caption("Buyers who disengaged without purchasing — OOS, price concern, or no response.")
            if not _lost_rows.empty:
                _lc1, _lc2, _lc3 = st.columns(3)
                _ld = _cf[_cf["IS_LOST_SALE"]==1].copy()
                with _lc1:
                    st.markdown("**By Store**")
                    st.dataframe(_ld.groupby("STORE_CODE").agg(Lost=("CONVERSATION_ID","count"),OOS_Related=("IS_OOS_CONFIRMED","sum")).reset_index().sort_values("Lost",ascending=False), use_container_width=True, hide_index=True)
                with _lc2:
                    st.markdown("**By Issue / Root Cause**")
                    st.dataframe(_ld.groupby("ISSUE_TYPE").agg(Lost=("CONVERSATION_ID","count")).reset_index().sort_values("Lost",ascending=False), use_container_width=True, hide_index=True)
                with _lc3:
                    st.markdown("**By Country**")
                    st.dataframe(_ld.groupby("COUNTRY_CODE").agg(Lost=("CONVERSATION_ID","count")).reset_index().sort_values("Lost",ascending=False), use_container_width=True, hide_index=True)
                st.markdown("---")
                st.markdown("**📋 Full Lost Sale Conversation Detail**")
                _lost_cols = [c for c in [
                    "CONVERSATION_ID","STORE_CODE","CHANNEL_NAME","SITE_NICK_NAME_ID",
                    "COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME","LAST_MSG_TIME",
                    "ISSUE_TYPE","PRIORITY","SENTIMENT","IS_OOS_CONFIRMED",
                    "ALT_SUGGESTED","ITEM_IDS","SIZE_MENTIONS","COLOR_MENTIONS","BUYER_SUMMARY"
                ] if c in _lost_rows.columns]
                st.dataframe(_lost_rows[_lost_cols].sort_values("LAST_MSG_TIME", ascending=False).reset_index(drop=True),
                    use_container_width=True, height=400, column_config={
                        "IS_OOS_CONFIRMED": st.column_config.CheckboxColumn("OOS?"),
                        "ALT_SUGGESTED":    st.column_config.CheckboxColumn("Alt Suggested?"),
                        "LAST_MSG_TIME":    st.column_config.DatetimeColumn("Last Msg", format="YYYY-MM-DD HH:mm"),
                        "BUYER_SUMMARY":    st.column_config.TextColumn("Summary", width="large"),
                        "ITEM_IDS":         st.column_config.TextColumn("Item IDs"),
                        "SIZE_MENTIONS":    st.column_config.TextColumn("Sizes"),
                        "COLOR_MENTIONS":   st.column_config.TextColumn("Colours"),
                    })
            else:
                st.success("✅ No lost sales detected.")

        # ── DRILLDOWN 4 : Upsell Opportunities ───────────────────────────────
        with st.expander(f"🔁 View {len(_upsell_rows)} Upsell Opportunity Conversations ({len(_upsell_missed)} missed)", expanded=False):
            st.caption("Buyers who asked for alternatives, recommendations, or bundling options.")
            if not _upsell_rows.empty:
                _u1, _u2 = st.columns(2)
                with _u1:
                    st.markdown(f"**✅ Acted ({len(_upsell_rows) - len(_upsell_missed)}) — Alternative was suggested**")
                    _acted = _upsell_rows[_upsell_rows["ALT_SUGGESTED"].astype(bool)]
                    if not _acted.empty:
                        st.dataframe(_acted[[c for c in ["CONVERSATION_ID","STORE_CODE","COUNTRY_CODE","TEAM_MEMBER","SENTIMENT","BUYER_SUMMARY"] if c in _acted.columns]].reset_index(drop=True),
                            use_container_width=True, height=280,
                            column_config={"BUYER_SUMMARY": st.column_config.TextColumn("Summary", width="large")})
                    else:
                        st.info("None acted on.")
                with _u2:
                    st.markdown(f"**❌ Missed ({len(_upsell_missed)}) — No alternative suggested**")
                    if not _upsell_missed.empty:
                        st.dataframe(_upsell_missed[[c for c in ["CONVERSATION_ID","STORE_CODE","COUNTRY_CODE","TEAM_MEMBER","SENTIMENT","BUYER_SUMMARY"] if c in _upsell_missed.columns]].reset_index(drop=True),
                            use_container_width=True, height=280,
                            column_config={"BUYER_SUMMARY": st.column_config.TextColumn("Summary", width="large")})
                    else:
                        st.success("All upsell opps were acted on!")
                st.markdown("---")
                st.markdown("**By Agent — Upsell Action Rate**")
                _uagg = _cf[_cf["IS_UPSELL_OPP"]==1].groupby("TEAM_MEMBER").agg(
                    Upsell_Opps    = ("CONVERSATION_ID","count"),
                    Alt_Suggested  = ("ALT_SUGGESTED","sum"),
                ).reset_index()
                _uagg["Action_Rate_%"] = (_uagg["Alt_Suggested"] / _uagg["Upsell_Opps"]*100).round(1)
                _uagg["Missed"] = _uagg["Upsell_Opps"] - _uagg["Alt_Suggested"]
                st.dataframe(_uagg.sort_values("Upsell_Opps", ascending=False), use_container_width=True, hide_index=True,
                    column_config={"Action_Rate_%": st.column_config.ProgressColumn("Action Rate%", format="%.1f%%", min_value=0, max_value=100)})
            else:
                st.info("No upsell opportunities detected.")

        # ── DRILLDOWN 5 : Product Inquiries ───────────────────────────────────
        with st.expander(f"🛍️ View {len(_prod_rows)} Product Inquiry Conversations", expanded=False):
            st.caption("Buyers asking about products — price, size, colour, availability. These are warm leads.")
            if not _prod_rows.empty:
                _pa1, _pa2, _pa3 = st.columns(3)
                _pd_int = _cf[_cf["ISSUE_TYPE"].astype(str)=="Product Inquiry"]
                with _pa1:
                    st.markdown("**By Store**")
                    st.dataframe(_pd_int.groupby(["STORE_CODE","COUNTRY_CODE"]).agg(
                        Inquiries=("CONVERSATION_ID","count"),
                        Converted=("IS_CONVERSION","sum"),
                        OOS=("IS_OOS_CONFIRMED","sum"),
                        Lost=("IS_LOST_SALE","sum"),
                    ).reset_index().sort_values("Inquiries",ascending=False), use_container_width=True, hide_index=True)
                with _pa2:
                    st.markdown("**By Channel / Platform**")
                    st.dataframe(_pd_int.groupby(["PLATFORM","CHANNEL_NAME"] if "CHANNEL_NAME" in _pd_int.columns else ["PLATFORM"]).agg(
                        Inquiries=("CONVERSATION_ID","count"),
                        Converted=("IS_CONVERSION","sum"),
                    ).reset_index().sort_values("Inquiries",ascending=False).head(15), use_container_width=True, hide_index=True)
                with _pa3:
                    # Top item IDs inquired
                    _all_items = [i for ids in _prod_rows["ITEM_IDS"].fillna("").str.split("|") for i in ids if i.strip()]
                    if _all_items:
                        st.markdown("**Top Item IDs Inquired**")
                        _item_ct = pd.DataFrame(Counter(_all_items).most_common(15), columns=["Item ID","Count"])
                        max_ic = int(_item_ct["Count"].max())
                        st.dataframe(_item_ct, use_container_width=True, hide_index=True,
                            column_config={"Count": st.column_config.ProgressColumn(format="%d", min_value=0, max_value=max_ic)})
                    else:
                        st.info("No item IDs found.")
                st.markdown("---")

                # Variation demand
                _v1, _v2 = st.columns(2)
                _all_sizes  = [s for szs in _prod_rows["SIZE_MENTIONS"].fillna("").str.split("|") for s in szs if s.strip()]
                _all_colors = [c for cls in _prod_rows["COLOR_MENTIONS"].fillna("").str.split("|") for c in cls if c.strip()]
                with _v1:
                    if _all_sizes:
                        st.markdown("**📐 Size Demand**")
                        _sdf = pd.DataFrame(Counter(_all_sizes).most_common(15), columns=["Size","Count"])
                        st.dataframe(_sdf, use_container_width=True, hide_index=True,
                            column_config={"Count": st.column_config.ProgressColumn(format="%d", min_value=0, max_value=int(_sdf["Count"].max()))})
                    else:
                        st.info("No size mentions.")
                with _v2:
                    if _all_colors:
                        st.markdown("**🎨 Colour Demand**")
                        _cdf = pd.DataFrame(Counter([c.title() for c in _all_colors]).most_common(15), columns=["Colour","Count"])
                        st.dataframe(_cdf, use_container_width=True, hide_index=True,
                            column_config={"Count": st.column_config.ProgressColumn(format="%d", min_value=0, max_value=int(_cdf["Count"].max()))})
                    else:
                        st.info("No colour mentions.")
                st.markdown("---")
                st.markdown("**📋 Full Product Inquiry Conversation Detail**")
                _pi_cols = [c for c in [
                    "CONVERSATION_ID","STORE_CODE","CHANNEL_NAME","SITE_NICK_NAME_ID",
                    "COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME","LAST_MSG_TIME",
                    "IS_CONVERSION","IS_OOS_CONFIRMED","IS_LOST_SALE","ALT_SUGGESTED",
                    "ITEM_IDS","SIZE_MENTIONS","COLOR_MENTIONS","SENTIMENT","BUYER_SUMMARY"
                ] if c in _prod_rows.columns]
                st.dataframe(_prod_rows[_pi_cols].sort_values("LAST_MSG_TIME", ascending=False).reset_index(drop=True),
                    use_container_width=True, height=420, column_config={
                        "IS_CONVERSION":   st.column_config.CheckboxColumn("Converted?"),
                        "IS_OOS_CONFIRMED":st.column_config.CheckboxColumn("OOS?"),
                        "IS_LOST_SALE":    st.column_config.CheckboxColumn("Lost?"),
                        "ALT_SUGGESTED":   st.column_config.CheckboxColumn("Alt Suggested?"),
                        "LAST_MSG_TIME":   st.column_config.DatetimeColumn("Last Msg", format="YYYY-MM-DD HH:mm"),
                        "BUYER_SUMMARY":   st.column_config.TextColumn("Summary", width="large"),
                        "ITEM_IDS":        st.column_config.TextColumn("Item IDs"),
                        "SIZE_MENTIONS":   st.column_config.TextColumn("Sizes"),
                        "COLOR_MENTIONS":  st.column_config.TextColumn("Colours"),
                    })
            else:
                st.info("No product inquiries detected.")

        st.markdown("---")

        # ── Sales Funnel + Stage Summary ──────────────────────────────────────
        st.markdown("#### 🔽 Sales Funnel Overview")
        _fc1, _fc2 = st.columns(2)
        with _fc1:
            funnel_df = pd.DataFrame([
                {"Stage":"All Conversations","Count":funnel.get("total",0)},
                {"Stage":"Product Inquiries","Count":funnel.get("prod_inq",0)},
                {"Stage":"High Intent",      "Count":funnel.get("high_intent",0)},
                {"Stage":"Converted",        "Count":funnel.get("converted",0)},
                {"Stage":"Lost Sales",       "Count":funnel.get("lost",0)},
            ])
            funnel_df["% of Total"] = (funnel_df["Count"]/max(funnel.get("total",1),1)*100).round(1)
            st.dataframe(funnel_df, use_container_width=True, hide_index=True,
                column_config={"% of Total": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=100)})
        with _fc2:
            st.markdown("**Sales Stage Distribution**")
            stage_ct = conv_filtered["SALES_STAGE"].astype(str).value_counts().reset_index()
            stage_ct.columns = ["Stage","Count"]
            stage_ct["% Share"] = (stage_ct["Count"]/len(conv_filtered)*100).round(1)
            st.dataframe(stage_ct, use_container_width=True, hide_index=True,
                column_config={"% Share": st.column_config.ProgressColumn(format="%.1f%%", min_value=0, max_value=100)})

        st.markdown("---")

        # ── Per-store sales summary ───────────────────────────────────────────
        st.markdown("#### 🏪 Sales Intelligence by Store")
        store_sales = _cf.groupby(["STORE_CODE","COUNTRY_CODE","PLATFORM"]).agg(
            Conversations    =("CONVERSATION_ID","count"),
            Conversions      =("IS_CONVERSION","sum"),
            Lost_Sales       =("IS_LOST_SALE","sum"),
            OOS_Hits         =("IS_OOS_CONFIRMED","sum"),
            Upsell_Opps      =("IS_UPSELL_OPP","sum"),
            Alt_Suggested    =("ALT_SUGGESTED","sum"),
            Product_Inquiries=("ISSUE_TYPE", lambda x: (x.astype(str)=="Product Inquiry").sum()),
        ).reset_index()
        store_sales["Conv_%"]       = (store_sales["Conversions"]/store_sales["Conversations"]*100).round(1)
        store_sales["Lost_%"]       = (store_sales["Lost_Sales"] /store_sales["Conversations"]*100).round(1)
        store_sales["OOS_%"]        = (store_sales["OOS_Hits"]   /store_sales["Conversations"]*100).round(1)
        store_sales["Upsell_Act_%"] = (store_sales["Alt_Suggested"]/store_sales["Upsell_Opps"].replace(0,np.nan)*100).round(1)
        st.dataframe(store_sales.sort_values("Conversations",ascending=False), use_container_width=True, hide_index=True,
            column_config={
                "Conv_%":       st.column_config.ProgressColumn("Conv%",      format="%.1f%%",min_value=0,max_value=100),
                "Lost_%":       st.column_config.ProgressColumn("Lost%",      format="%.1f%%",min_value=0,max_value=100),
                "OOS_%":        st.column_config.ProgressColumn("OOS%",       format="%.1f%%",min_value=0,max_value=100),
                "Upsell_Act_%": st.column_config.ProgressColumn("Upsell Act%",format="%.1f%%",min_value=0,max_value=100),
            })

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 8 : AM & MERCH PERFORMANCE — with drilldowns
    # ══════════════════════════════════════════════════════════════════════════
    with tab8:
        st.markdown("### 🛍️ Account Management & Merchandising Performance")
        st.caption("Chat data as a sales signal — for AM and Merch teams to action.")

        _cf8 = conv_filtered.copy()
        for _bc in ["IS_CONVERSION","IS_LOST_SALE","IS_OOS_CONFIRMED","IS_UPSELL_OPP","ALT_SUGGESTED","IS_UNRESOLVED"]:
            if _bc in _cf8.columns:
                _cf8[_bc] = _cf8[_bc].astype(bool).astype(int)

        # ── AM Scorecard ──────────────────────────────────────────────────────
        st.markdown("#### 🏪 Per-Store AM Scorecard")
        am_df = build_am_scorecard(conv_filtered)
        if not am_df.empty:
            st.dataframe(am_df, use_container_width=True, hide_index=True,
                column_config={
                    "Conv_Rate_%":  st.column_config.ProgressColumn("Conv%",      format="%.1f%%",min_value=0,max_value=100),
                    "Lost_Rate_%":  st.column_config.ProgressColumn("Lost%",      format="%.1f%%",min_value=0,max_value=100),
                    "OOS_Rate_%":   st.column_config.ProgressColumn("OOS%",       format="%.1f%%",min_value=0,max_value=100),
                    "Upsell_Act_%": st.column_config.ProgressColumn("Upsell Act%",format="%.1f%%",min_value=0,max_value=100),
                    "CRR_%":        st.column_config.ProgressColumn("CRR%",       format="%.1f%%",min_value=0,max_value=100),
                    "Avg_CSAT":     st.column_config.NumberColumn("CSAT",         format="%.1f"),
                    "Avg_CRT_mins": st.column_config.NumberColumn("CRT (min)",    format="%.0f"),
                })

            # ── Drilldown by Store ────────────────────────────────────────────
            _store_sel = st.selectbox("🔍 Drill into a Store", ["(All Stores)"] + sorted(conv_filtered["STORE_CODE"].astype(str).unique().tolist()), key="am_store_sel")
            _store_data = conv_filtered if _store_sel == "(All Stores)" else conv_filtered[conv_filtered["STORE_CODE"].astype(str) == _store_sel]
            with st.expander(f"📋 {_store_sel} — Full Conversation Detail ({len(_store_data)} convs)", expanded=False):
                _sd_cols = [c for c in [
                    "CONVERSATION_ID","PLATFORM","CHANNEL_NAME","SITE_NICK_NAME_ID","COUNTRY_CODE",
                    "TEAM_MEMBER","BUYER_NAME","LAST_MSG_TIME","ISSUE_TYPE","PRIORITY","SENTIMENT",
                    "IS_CONVERSION","IS_OOS_CONFIRMED","IS_LOST_SALE","IS_UPSELL_OPP","ALT_SUGGESTED",
                    "ITEM_IDS","SIZE_MENTIONS","COLOR_MENTIONS","CSAT_PROXY","AVG_CRT_MINS","BUYER_SUMMARY"
                ] if c in _store_data.columns]
                st.dataframe(_store_data[_sd_cols].sort_values("LAST_MSG_TIME", ascending=False).reset_index(drop=True),
                    use_container_width=True, height=420, column_config={
                        "IS_CONVERSION":   st.column_config.CheckboxColumn("Converted?"),
                        "IS_OOS_CONFIRMED":st.column_config.CheckboxColumn("OOS?"),
                        "IS_LOST_SALE":    st.column_config.CheckboxColumn("Lost?"),
                        "IS_UPSELL_OPP":   st.column_config.CheckboxColumn("Upsell Opp?"),
                        "ALT_SUGGESTED":   st.column_config.CheckboxColumn("Alt Suggested?"),
                        "CSAT_PROXY":      st.column_config.NumberColumn("CSAT",   format="%.1f"),
                        "AVG_CRT_MINS":    st.column_config.NumberColumn("CRT(m)", format="%.0f"),
                        "LAST_MSG_TIME":   st.column_config.DatetimeColumn("Last Msg", format="YYYY-MM-DD HH:mm"),
                        "BUYER_SUMMARY":   st.column_config.TextColumn("Summary", width="large"),
                        "ITEM_IDS":        st.column_config.TextColumn("Item IDs"),
                        "SIZE_MENTIONS":   st.column_config.TextColumn("Sizes"),
                        "COLOR_MENTIONS":  st.column_config.TextColumn("Colours"),
                    })

        st.markdown("---")

        # ── Extended Team Performance ─────────────────────────────────────────
        st.markdown("#### 👥 Team Performance — Ops + Sales")
        st.caption(f"From {TEAM_START_DATE.strftime('%d %b %Y')} onwards")
        team_sales = build_team_sales_perf(conv_filtered)
        if not team_sales.empty:
            team_cols = [c for c in [
                "TEAM_MEMBER","Shift","Conversations","Resolved","Unresolved",
                "CRR_%","Avg_CSAT","Avg_CRT_mins","Conversions","Conv_Rate_%",
                "Upsell_Opps","Alt_Suggested","Upsell_Act_%","Lost_Sales","OOS_Handled",
                "High_Priority","Positive_Sent","Negative_Sent"
            ] if c in team_sales.columns]
            st.dataframe(team_sales[team_cols].reset_index(drop=True), use_container_width=True,
                column_config={
                    "CRR_%":        st.column_config.ProgressColumn("CRR%",       format="%.1f%%",min_value=0,max_value=100),
                    "Conv_Rate_%":  st.column_config.ProgressColumn("Conv%",      format="%.1f%%",min_value=0,max_value=100),
                    "Upsell_Act_%": st.column_config.ProgressColumn("Upsell Act%",format="%.1f%%",min_value=0,max_value=100),
                    "Avg_CSAT":     st.column_config.NumberColumn("CSAT",         format="%.2f"),
                    "Avg_CRT_mins": st.column_config.NumberColumn("CRT (min)",    format="%.1f"),
                })
            if len(team_sales) > 1:
                st.markdown("---")
                h1,h2,h3 = st.columns(3)
                best_crt  = team_sales.dropna(subset=["Avg_CRT_mins"]).nsmallest(1,"Avg_CRT_mins")
                best_conv = team_sales.nlargest(1,"Conv_Rate_%")
                best_csat = team_sales.nlargest(1,"Avg_CSAT")
                with h1:
                    if not best_crt.empty: st.success(f"⚡ **Fastest:** {best_crt.iloc[0]['TEAM_MEMBER']} ({fmt_mins(best_crt.iloc[0]['Avg_CRT_mins'])})")
                with h2:
                    if not best_conv.empty: st.success(f"💰 **Top Converter:** {best_conv.iloc[0]['TEAM_MEMBER']} ({best_conv.iloc[0]['Conv_Rate_%']:.1f}%)")
                with h3:
                    if not best_csat.empty: st.success(f"⭐ **Top CSAT:** {best_csat.iloc[0]['TEAM_MEMBER']} ({best_csat.iloc[0]['Avg_CSAT']:.2f}/5)")

            # ── Per-agent drilldown ───────────────────────────────────────────
            st.markdown("---")
            _agent_sel = st.selectbox("🔍 Drill into an Agent", ["(All Agents)"] + team_sales["TEAM_MEMBER"].tolist(), key="am_agent_sel")
            _agent_data = (conv_filtered if _agent_sel == "(All Agents)"
                           else conv_filtered[conv_filtered["TEAM_MEMBER"].astype(str) == _agent_sel])
            _agent_data = _agent_data[_agent_data["LAST_MSG_TIME"] >= TEAM_START_DATE]
            with st.expander(f"📋 {_agent_sel} — All Conversations ({len(_agent_data)})", expanded=False):
                st.dataframe(_agent_data[_DRILL_BASE].sort_values("LAST_MSG_TIME", ascending=False).reset_index(drop=True),
                    use_container_width=True, height=420, column_config=_DRILL_CFG)
        else:
            st.info("No team data available for the selected period.")

        st.markdown("---")

        # ── Product & Variation Demand — Merch ────────────────────────────────
        st.markdown("#### 🔍 Product & Variation Demand — Merch Signals")
        st.caption("Top inquired items and requested sizes/colours — for listing optimisation and stock planning.")
        _item_sum, _var_df, _exp_df = build_product_demand(conv_filtered)
        _m1, _m2 = st.columns(2)
        with _m1:
            st.markdown("**📦 Top Inquired Item IDs**")
            if not _item_sum.empty:
                _disp = _item_sum[["Item_ID","Total_Inquiries","Stores","Countries","Platforms","OOS_Confirmed","Lost_Sales","Conv_Rate_%"]].head(20)
                max_i = int(_disp["Total_Inquiries"].max())
                st.dataframe(_disp, use_container_width=True, hide_index=True,
                    column_config={"Total_Inquiries": st.column_config.ProgressColumn("Inquiries", format="%d", min_value=0, max_value=max_i),
                                   "Conv_Rate_%": st.column_config.NumberColumn("Conv%", format="%.1f%%")})
            else:
                st.info("No item IDs detected.")
        with _m2:
            st.markdown("**📐 Top Requested Variations (All convs)**")
            if not _var_df.empty:
                max_v = int(_var_df["Count"].max())
                st.dataframe(_var_df, use_container_width=True, hide_index=True,
                    column_config={"Count": st.column_config.ProgressColumn(format="%d", min_value=0, max_value=max_v)})
            else:
                st.info("No variation mentions detected.")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 9 : KEY IMPROVEMENTS — with linked data
    # ══════════════════════════════════════════════════════════════════════════
    with tab9:
        st.markdown("### 🎯 Key Improvement Areas")
        st.caption("Auto-generated from your chat data. Each area links to the supporting conversations.")

        funnel9 = build_sales_funnel(conv_filtered)
        improvements = generate_key_improvements(conv_filtered, funnel9)

        for priority_level, area, rec in improvements:
            bg     = "#FDECEA" if "HIGH" in priority_level else ("#FFFBEB" if "MEDIUM" in priority_level else "#F0FDF4")
            border = "#E74C3C" if "HIGH" in priority_level else ("#F59E0B" if "MEDIUM" in priority_level else "#22C55E")
            st.markdown(f"""
            <div style="background:{bg};border-left:4px solid {border};border-radius:6px;
                        padding:0.9rem 1.1rem;margin-bottom:0.8rem">
              <b>{priority_level} · {area}</b><br>
              <span style="font-size:0.88rem;color:#374151">{rec}</span>
            </div>""", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("#### 📋 Improvement Action Tracker")
        impr_df = pd.DataFrame([
            {"Priority":pl,"Area":ar,"Recommendation":rc,"Status":"Open","Owner":""}
            for pl,ar,rc in improvements
        ])
        st.dataframe(impr_df, use_container_width=True, hide_index=True,
            column_config={"Recommendation": st.column_config.TextColumn(width="large")})

        st.markdown("---")

        # ── Linked data for each improvement area ─────────────────────────────
        st.markdown("#### 🔗 Supporting Data per Improvement Area")

        with st.expander("📦 OOS — Restock Priority List", expanded=False):
            _oos9 = conv_filtered[conv_filtered["IS_OOS_CONFIRMED"].astype(bool)]
            if not _oos9.empty:
                _cf9 = _oos9.copy()
                for _bc9 in ["IS_LOST_SALE","ALT_SUGGESTED"]:
                    if _bc9 in _cf9.columns: _cf9[_bc9] = _cf9[_bc9].astype(bool).astype(int)
                _oos9_agg = _cf9.groupby(["STORE_CODE","ITEM_IDS"]).agg(
                    Demand=("CONVERSATION_ID","count"),
                    Lost=("IS_LOST_SALE","sum"),
                    Alt_Suggested=("ALT_SUGGESTED","sum"),
                    Sizes=("SIZE_MENTIONS", lambda x:" | ".join(sorted(set(v for vals in x for v in str(vals).split("|") if v.strip())))),
                    Colors=("COLOR_MENTIONS",lambda x:" | ".join(sorted(set(v for vals in x for v in str(vals).split("|") if v.strip())))),
                ).reset_index()
                _oos9_agg["Priority"] = _oos9_agg["Demand"] + _oos9_agg["Lost"]*2
                st.dataframe(_oos9_agg.sort_values("Priority",ascending=False).reset_index(drop=True),
                    use_container_width=True, hide_index=True,
                    column_config={"Priority": st.column_config.ProgressColumn(format="%d",min_value=0,max_value=int(_oos9_agg["Priority"].max()))})
            else:
                st.info("No OOS data.")

        with st.expander("💸 Lost Sales — Full Detail", expanded=False):
            _ls9 = conv_filtered[conv_filtered["IS_LOST_SALE"].astype(bool)]
            if not _ls9.empty:
                _ls9_cols = [c for c in ["CONVERSATION_ID","STORE_CODE","CHANNEL_NAME","COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME","LAST_MSG_TIME","ISSUE_TYPE","IS_OOS_CONFIRMED","ALT_SUGGESTED","SENTIMENT","ITEM_IDS","BUYER_SUMMARY"] if c in _ls9.columns]
                st.dataframe(_ls9[_ls9_cols].sort_values("LAST_MSG_TIME",ascending=False).reset_index(drop=True),
                    use_container_width=True, height=350, column_config={
                        "IS_OOS_CONFIRMED":st.column_config.CheckboxColumn("OOS?"),
                        "ALT_SUGGESTED":   st.column_config.CheckboxColumn("Alt Suggested?"),
                        "LAST_MSG_TIME":   st.column_config.DatetimeColumn("Last Msg",format="YYYY-MM-DD HH:mm"),
                        "BUYER_SUMMARY":   st.column_config.TextColumn("Summary",width="large"),
                    })
            else:
                st.success("No lost sales.")

        with st.expander("🔁 Missed Upsells — Agent Action Needed", expanded=False):
            _mu9 = conv_filtered[conv_filtered["IS_UPSELL_OPP"].astype(bool) & ~conv_filtered["ALT_SUGGESTED"].astype(bool)]
            if not _mu9.empty:
                _mu9_cols = [c for c in ["CONVERSATION_ID","STORE_CODE","CHANNEL_NAME","COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME","LAST_MSG_TIME","ISSUE_TYPE","SENTIMENT","BUYER_SUMMARY"] if c in _mu9.columns]
                st.dataframe(_mu9[_mu9_cols].sort_values("LAST_MSG_TIME",ascending=False).reset_index(drop=True),
                    use_container_width=True, height=350, column_config={
                        "LAST_MSG_TIME": st.column_config.DatetimeColumn("Last Msg",format="YYYY-MM-DD HH:mm"),
                        "BUYER_SUMMARY": st.column_config.TextColumn("Summary",width="large"),
                    })
            else:
                st.success("All upsell opportunities were acted on!")

        with st.expander("🛍️ Unconverted Product Inquiries — Warm Leads", expanded=False):
            _up9 = conv_filtered[(conv_filtered["ISSUE_TYPE"].astype(str)=="Product Inquiry") & ~conv_filtered["IS_CONVERSION"].astype(bool)]
            if not _up9.empty:
                _up9_cols = [c for c in ["CONVERSATION_ID","STORE_CODE","CHANNEL_NAME","SITE_NICK_NAME_ID","COUNTRY_CODE","TEAM_MEMBER","BUYER_NAME","LAST_MSG_TIME","IS_OOS_CONFIRMED","ALT_SUGGESTED","ITEM_IDS","SIZE_MENTIONS","COLOR_MENTIONS","SENTIMENT","BUYER_SUMMARY"] if c in _up9.columns]
                st.dataframe(_up9[_up9_cols].sort_values("LAST_MSG_TIME",ascending=False).reset_index(drop=True),
                    use_container_width=True, height=380, column_config={
                        "IS_OOS_CONFIRMED":st.column_config.CheckboxColumn("OOS?"),
                        "ALT_SUGGESTED":   st.column_config.CheckboxColumn("Alt Suggested?"),
                        "LAST_MSG_TIME":   st.column_config.DatetimeColumn("Last Msg",format="YYYY-MM-DD HH:mm"),
                        "BUYER_SUMMARY":   st.column_config.TextColumn("Summary",width="large"),
                        "ITEM_IDS":        st.column_config.TextColumn("Item IDs"),
                        "SIZE_MENTIONS":   st.column_config.TextColumn("Sizes"),
                        "COLOR_MENTIONS":  st.column_config.TextColumn("Colours"),
                    })
            else:
                st.success("All product inquiries converted!")


    # ══════════════════════════════════════════════════════════════════════════
    # TAB 10 : PRODUCT INTELLIGENCE — item_id level analysis
    # ══════════════════════════════════════════════════════════════════════════
    with tab10:
        st.markdown("### 📦 Product Intelligence")
        st.caption("Every item_id mentioned in chat — inquiry count, platform, store, channel, OOS, lost sale, size & colour demand.")

        _item_sum, _var_df_pi, _exp_df = build_product_demand(conv_filtered)

        if _item_sum.empty:
            st.info("No item IDs detected in the current filtered data. Item IDs are extracted from messages where buyer shares a product link (item_id:XXXXXXXXX).")
        else:
            # ── Top KPI row ───────────────────────────────────────────────────
            _pi_k1, _pi_k2, _pi_k3, _pi_k4, _pi_k5 = st.columns(5)
            with _pi_k1:
                st.markdown(f"""<div class="metric-card navy">
                    <div class="metric-label">📦 Unique Items Inquired</div>
                    <div class="metric-val">{len(_item_sum)}</div>
                    <div class="metric-sub">Distinct item IDs</div>
                </div>""", unsafe_allow_html=True)
            with _pi_k2:
                st.markdown(f"""<div class="metric-card green">
                    <div class="metric-label">🔢 Total Item Mentions</div>
                    <div class="metric-val">{int(_item_sum["Total_Inquiries"].sum())}</div>
                    <div class="metric-sub">Across all conversations</div>
                </div>""", unsafe_allow_html=True)
            with _pi_k3:
                _oos_items = int((_item_sum["OOS_Confirmed"] > 0).sum())
                st.markdown(f"""<div class="metric-card red">
                    <div class="metric-label">📦 Items with OOS</div>
                    <div class="metric-val">{_oos_items}</div>
                    <div class="metric-sub">Need restock</div>
                </div>""", unsafe_allow_html=True)
            with _pi_k4:
                _lost_items = int((_item_sum["Lost_Sales"] > 0).sum())
                st.markdown(f"""<div class="metric-card red">
                    <div class="metric-label">💸 Items with Lost Sales</div>
                    <div class="metric-val">{_lost_items}</div>
                    <div class="metric-sub">Revenue at risk</div>
                </div>""", unsafe_allow_html=True)
            with _pi_k5:
                _conv_items = int((_item_sum["Conversions"] > 0).sum())
                st.markdown(f"""<div class="metric-card orange">
                    <div class="metric-label">💰 Items with Conversions</div>
                    <div class="metric-val">{_conv_items}</div>
                    <div class="metric-sub">Confirmed demand</div>
                </div>""", unsafe_allow_html=True)

            st.markdown("---")

            # ── Filter controls ───────────────────────────────────────────────
            st.markdown("#### 🔍 Filter & Search Products")
            _pf1, _pf2, _pf3, _pf4 = st.columns(4)
            with _pf1:
                _pi_store = st.selectbox("Store", ["All"] + sorted(_exp_df["STORE_CODE"].unique().tolist()), key="pi_store")
            with _pf2:
                _pi_country = st.selectbox("Country", ["All"] + sorted(_exp_df["COUNTRY_CODE"].unique().tolist()), key="pi_country")
            with _pf3:
                _pi_platform = st.selectbox("Platform", ["All"] + sorted(_exp_df["PLATFORM"].unique().tolist()), key="pi_platform")
            with _pf4:
                _pi_flag = st.selectbox("Flag", ["All","OOS Items","Lost Sale Items","Converted Items","High Demand (≥10)"], key="pi_flag")

            # Apply filters
            _item_filt = _item_sum.copy()
            if _pi_store != "All":
                _item_filt = _item_filt[_item_filt["Stores"].str.contains(_pi_store, na=False)]
            if _pi_country != "All":
                _item_filt = _item_filt[_item_filt["Countries"].str.contains(_pi_country, na=False)]
            if _pi_platform != "All":
                _item_filt = _item_filt[_item_filt["Platforms"].str.contains(_pi_platform, na=False)]
            if _pi_flag == "OOS Items":
                _item_filt = _item_filt[_item_filt["OOS_Confirmed"] > 0]
            elif _pi_flag == "Lost Sale Items":
                _item_filt = _item_filt[_item_filt["Lost_Sales"] > 0]
            elif _pi_flag == "Converted Items":
                _item_filt = _item_filt[_item_filt["Conversions"] > 0]
            elif _pi_flag == "High Demand (≥10)":
                _item_filt = _item_filt[_item_filt["Total_Inquiries"] >= 10]

            st.markdown(f"**{len(_item_filt)} items** match current filters")

            # ── Main product table ────────────────────────────────────────────
            st.markdown("#### 📊 Product Inquiry Summary Table")
            st.caption("One row per item ID — inquiry count, which store/platform/channel, OOS flags, sizes & colours requested.")
            _disp_cols = ["Item_ID","Total_Inquiries","Unique_Convs","Stores","Countries","Platforms","Sites",
                          "Conv_Rate_%","OOS_Confirmed","OOS_Rate_%","Lost_Sales","Upsell_Opps",
                          "Sizes_Requested","Colors_Requested"]
            _disp_cols = [c for c in _disp_cols if c in _item_filt.columns]
            if not _item_filt.empty:
                _max_inq = int(_item_filt["Total_Inquiries"].max()) if len(_item_filt) else 1
                st.dataframe(
                    _item_filt[_disp_cols].reset_index(drop=True),
                    use_container_width=True,
                    height=450,
                    column_config={
                        "Item_ID":         st.column_config.TextColumn("Item ID"),
                        "Total_Inquiries": st.column_config.ProgressColumn("Inquiries", format="%d", min_value=0, max_value=_max_inq),
                        "Unique_Convs":    st.column_config.NumberColumn("Unique Convs"),
                        "Stores":          st.column_config.TextColumn("Stores"),
                        "Countries":       st.column_config.TextColumn("Countries"),
                        "Platforms":       st.column_config.TextColumn("Platforms"),
                        "Sites":           st.column_config.TextColumn("Sites (Channels)"),
                        "Conv_Rate_%":     st.column_config.NumberColumn("Conv%", format="%.1f%%"),
                        "OOS_Confirmed":   st.column_config.NumberColumn("OOS Count"),
                        "OOS_Rate_%":      st.column_config.NumberColumn("OOS%", format="%.1f%%"),
                        "Lost_Sales":      st.column_config.NumberColumn("Lost Sales"),
                        "Upsell_Opps":     st.column_config.NumberColumn("Upsell Opps"),
                        "Sizes_Requested": st.column_config.TextColumn("Sizes Requested"),
                        "Colors_Requested":st.column_config.TextColumn("Colours Requested"),
                    }
                )

            st.markdown("---")

            # ── Drill into specific item ──────────────────────────────────────
            st.markdown("#### 🔎 Drill Into a Specific Item ID")
            _all_item_ids = _item_filt["Item_ID"].tolist() if not _item_filt.empty else _item_sum["Item_ID"].tolist()
            if _all_item_ids:
                _sel_item = st.selectbox("Select Item ID to drill into", _all_item_ids, key="pi_item_sel",
                    help="Shows all conversations mentioning this item — full context, store, channel, buyer summary")
                _item_convs = _exp_df[_exp_df["Item_ID"] == _sel_item]
                _item_conv_ids = _item_convs["CONVERSATION_ID"].unique().tolist()
                _item_full = conv_filtered[conv_filtered["CONVERSATION_ID"].isin(_item_conv_ids)]

                if not _item_full.empty:
                    # Item-level KPIs
                    _ik1, _ik2, _ik3, _ik4, _ik5, _ik6 = st.columns(6)
                    with _ik1: st.metric("Total Inquiries", len(_item_full))
                    with _ik2: st.metric("Conversions", int(_item_full["IS_CONVERSION"].astype(bool).sum()))
                    with _ik3: st.metric("OOS Confirmed", int(_item_full["IS_OOS_CONFIRMED"].astype(bool).sum()))
                    with _ik4: st.metric("Lost Sales", int(_item_full["IS_LOST_SALE"].astype(bool).sum()))
                    with _ik5: st.metric("Upsell Opps", int(_item_full["IS_UPSELL_OPP"].astype(bool).sum()))
                    with _ik6:
                        _avg_csat = _item_full["CSAT_PROXY"].mean()
                        st.metric("Avg CSAT", f"{_avg_csat:.1f}" if not pd.isna(_avg_csat) else "—")

                    _id1, _id2, _id3 = st.columns(3)
                    with _id1:
                        st.markdown("**By Store**")
                        st.dataframe(_item_full.groupby("STORE_CODE").size().reset_index(name="Count").sort_values("Count",ascending=False), use_container_width=True, hide_index=True)
                    with _id2:
                        st.markdown("**By Platform / Site**")
                        st.dataframe(_item_full.groupby(["PLATFORM","SITE_NICK_NAME_ID"] if "SITE_NICK_NAME_ID" in _item_full.columns else ["PLATFORM"]).size().reset_index(name="Count").sort_values("Count",ascending=False), use_container_width=True, hide_index=True)
                    with _id3:
                        st.markdown("**By Country**")
                        st.dataframe(_item_full.groupby("COUNTRY_CODE").size().reset_index(name="Count").sort_values("Count",ascending=False), use_container_width=True, hide_index=True)

                    st.markdown("**📋 All Conversations for this Item**")
                    _item_detail_cols = [c for c in [
                        "CONVERSATION_ID","STORE_CODE","CHANNEL_NAME","SITE_NICK_NAME_ID",
                        "COUNTRY_CODE","PLATFORM","TEAM_MEMBER","BUYER_NAME","LAST_MSG_TIME",
                        "ISSUE_TYPE","PRIORITY","SENTIMENT","IS_CONVERSION","IS_OOS_CONFIRMED",
                        "IS_LOST_SALE","ALT_SUGGESTED","SIZE_MENTIONS","COLOR_MENTIONS",
                        "CSAT_PROXY","AVG_CRT_MINS","BUYER_SUMMARY"
                    ] if c in _item_full.columns]
                    st.dataframe(
                        _item_full[_item_detail_cols].sort_values("LAST_MSG_TIME", ascending=False).reset_index(drop=True),
                        use_container_width=True,
                        height=420,
                        column_config={
                            "IS_CONVERSION":   st.column_config.CheckboxColumn("Converted?"),
                            "IS_OOS_CONFIRMED":st.column_config.CheckboxColumn("OOS?"),
                            "IS_LOST_SALE":    st.column_config.CheckboxColumn("Lost?"),
                            "ALT_SUGGESTED":   st.column_config.CheckboxColumn("Alt Suggested?"),
                            "LAST_MSG_TIME":   st.column_config.DatetimeColumn("Last Msg", format="YYYY-MM-DD HH:mm"),
                            "CSAT_PROXY":      st.column_config.NumberColumn("CSAT", format="%.1f"),
                            "AVG_CRT_MINS":    st.column_config.NumberColumn("CRT(m)", format="%.0f"),
                            "BUYER_SUMMARY":   st.column_config.TextColumn("Summary", width="large"),
                            "SIZE_MENTIONS":   st.column_config.TextColumn("Sizes"),
                            "COLOR_MENTIONS":  st.column_config.TextColumn("Colours"),
                        }
                    )

            st.markdown("---")

            # ── OOS Items Restock Priority ────────────────────────────────────
            with st.expander(f"📦 OOS Restock Priority — {int((_item_sum['OOS_Confirmed']>0).sum())} items confirmed OOS", expanded=False):
                st.caption("Items where seller confirmed OOS in chat. Priority Score = Demand + (Lost Sales × 2). Share with buying team.")
                _oos_items_df = _item_sum[_item_sum["OOS_Confirmed"] > 0].copy()
                _oos_items_df["Priority_Score"] = _oos_items_df["Total_Inquiries"] + _oos_items_df["Lost_Sales"] * 2
                _oos_items_df = _oos_items_df.sort_values("Priority_Score", ascending=False).reset_index(drop=True)
                _max_ps = int(_oos_items_df["Priority_Score"].max()) if len(_oos_items_df) else 1
                st.dataframe(
                    _oos_items_df[["Item_ID","Total_Inquiries","OOS_Confirmed","Lost_Sales","Priority_Score","Stores","Countries","Sizes_Requested","Colors_Requested"]],
                    use_container_width=True, hide_index=True,
                    column_config={
                        "Priority_Score":  st.column_config.ProgressColumn("Priority", format="%d", min_value=0, max_value=_max_ps),
                        "Total_Inquiries": st.column_config.NumberColumn("Demand"),
                        "OOS_Confirmed":   st.column_config.NumberColumn("OOS Count"),
                        "Lost_Sales":      st.column_config.NumberColumn("Lost Sales"),
                    }
                )

            # ── Size & Colour Demand ──────────────────────────────────────────
            st.markdown("---")
            st.markdown("#### 📐🎨 Size & Colour Demand (All Conversations)")
            st.caption("Most requested sizes and colours — use for stock planning and listing optimisation.")
            _sv1, _sv2 = st.columns(2)
            _size_data = [r for r in _var_df_pi.to_dict("records") if r["Type"] == "Size"]
            _color_data = [r for r in _var_df_pi.to_dict("records") if r["Type"] == "Color"]
            with _sv1:
                st.markdown("**📐 Size Demand**")
                if _size_data:
                    _sdf = pd.DataFrame(_size_data)
                    max_s = int(_sdf["Count"].max())
                    st.dataframe(_sdf[["Variation","Count"]].rename(columns={"Variation":"Size"}),
                        use_container_width=True, hide_index=True,
                        column_config={"Count": st.column_config.ProgressColumn(format="%d", min_value=0, max_value=max_s)})
                else:
                    st.info("No size mentions detected.")
            with _sv2:
                st.markdown("**🎨 Colour Demand**")
                if _color_data:
                    _cdf = pd.DataFrame(_color_data)
                    max_c = int(_cdf["Count"].max())
                    st.dataframe(_cdf[["Variation","Count"]].rename(columns={"Variation":"Colour"}),
                        use_container_width=True, hide_index=True,
                        column_config={"Count": st.column_config.ProgressColumn(format="%d", min_value=0, max_value=max_c)})
                else:
                    st.info("No colour mentions detected.")


    # ── Issue / Store Breakdown Tables (below all tabs — unchanged) ────────────

    # ── Issue Breakdown Table ─────────────────────────────────────────────────
    st.markdown('<div class="section-title">📂 Issue Type Breakdown</div>', unsafe_allow_html=True)
    ib = (
        conv_filtered
        .groupby(["ISSUE_TYPE", "PRIORITY"])
        .agg(
            Count=("CONVERSATION_ID", "count"),
            Unresolved=("IS_UNRESOLVED", "sum"),
            Avg_CSAT=("CSAT_PROXY", "mean"),
            Avg_CRT_mins=("AVG_CRT_MINS", "mean"),
        )
        .reset_index()
        .sort_values("Count", ascending=False)
    )
    ib["Avg_CSAT"] = ib["Avg_CSAT"].round(1)
    ib["Avg_CRT_mins"] = ib["Avg_CRT_mins"].round(0).fillna(0).astype(int)
    ib["Unresolved"] = ib["Unresolved"].astype(int)
    st.dataframe(ib, use_container_width=True, height=300)

    # ── Store Performance ─────────────────────────────────────────────────────
    st.markdown('<div class="section-title">🏪 Store Performance</div>', unsafe_allow_html=True)
    sp = (
        conv_filtered
        .groupby(["STORE_CODE", "PLATFORM"])
        .agg(
            Conversations=("CONVERSATION_ID", "count"),
            Unresolved=("IS_UNRESOLVED", "sum"),
            Avg_CSAT=("CSAT_PROXY", "mean"),
            Avg_CRT_mins=("AVG_CRT_MINS", "mean"),
            Negative_Sent=("SENTIMENT", lambda x: (x == "Negative").sum()),
        )
        .reset_index()
        .sort_values("Conversations", ascending=False)
    )
    sp["Avg_CSAT"] = sp["Avg_CSAT"].round(1)
    sp["Avg_CRT_mins"] = sp["Avg_CRT_mins"].round(0).fillna(0).astype(int)
    sp["Unresolved"] = sp["Unresolved"].astype(int)
    sp["CRR%"] = ((sp["Conversations"] - sp["Unresolved"]) / sp["Conversations"] * 100).round(1)
    st.dataframe(sp, use_container_width=True, height=350)

    # ── Excel Download ────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">⬇️ Download Report</div>', unsafe_allow_html=True)
    cutoff_7d = today_ts - pd.Timedelta(days=6)
    conv_7day = conv_df[conv_df["LAST_MSG_TIME"] >= cutoff_7d].copy()

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        if st.button(f"📊 Generate Last 7 Days Report", use_container_width=True):
            with st.spinner("Building report…"):
                excel_7day = build_excel(conv_7day, today_str)
            st.download_button(
                label=f"📥 Download ({cutoff_7d.date()} → {today_date})",
                data=excel_7day,
                file_name=f"Chat_Analysis_Last7Days_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    with dl_col2:
        if st.button("📊 Generate Filtered View Report", use_container_width=True):
            with st.spinner("Building report…"):
                excel_filtered = build_excel(conv_filtered, today_str)
            st.download_button(
                label="📥 Download Filtered View",
                data=excel_filtered,
                file_name=f"Chat_Analysis_Filtered_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    st.caption(
        "Click **Generate** first, then **Download**. "
        "**Last 7 Days** = default daily export · "
        "**Filtered View** = matches current sidebar selection."
    )


if __name__ == "__main__":
    main()
