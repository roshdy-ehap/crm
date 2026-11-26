import os
import re
import time
import random
import sqlite3
import smtplib
import threading
import mimetypes
from datetime import datetime
from urllib.parse import quote
from email.message import EmailMessage

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB_FILE = "crm_pro.db"
LINKEDIN_SESSION_DIR = "./session_linkedin"
WA_SESSION_DIR = "./session_whatsapp"

STAGES = ["جديد", "تم التواصل", "مهتم", "معاينة", "تم البيع", "ملغي", "عدم الإزعاج"]
STAGE_COLORS = {
    "جديد": "#3B82F6",
    "تم التواصل": "#8B5CF6",
    "مهتم": "#F59E0B",
    "معاينة": "#06B6D4",
    "تم البيع": "#10B981",
    "ملغي": "#64748B",
    "عدم الإزعاج": "#EF4444",
}

DEFAULT_TEMPLATES = [
    (
        "ترحيب B2B (مع تنويع تلقائي)",
        "{تعاون محتمل|فرصة شراكة|تواصل سريع} مع {name}",
        "{السلام عليكم|أهلاً بكم|تحياتي} فريق {name}،\n"
        "{أتمنى تكونوا بخير|يسعدني التواصل معكم}.\n"
        "اطلعنا على شغلكم المميز في مجال {category} في {location}، وحابين نعرض عليكم فكرة تعاون مفيدة.\n"
        "{لو وقتكم يسمح، ممكن نرتب مكالمة سريعة؟|هل يناسبكم نرسل التفاصيل الكاملة؟}\n\n"
        "تحياتي،\n{sender_name}",
        "الكل",
    ),
    (
        "عرض واتساب مختصر وسريع",
        "عرض خاص لـ {name}",
        "{السلام عليكم|مساء الخير|أهلاً بحضراتكم} فريق {name}،\n"
        "متابعين شغلكم في {category} ({location}) وحابين نتواصل مع الشخص المسؤول عن الشراكات أو المبيعات.\n"
        "ممكن أعرف أنسب شخص أتواصل معاه؟",
        "WhatsApp",
    ),
    (
        "متابعة بعد التواصل الأول",
        "متابعة سريعة بخصوص {category}",
        "{السلام عليكم|أهلاً} أ/ {name}،\n"
        "حبينا نطمن هل وصلكم عرضنا السابق بخصوص {category} وهل فيه أي استفسار نقدر نوضحه؟\n\n"
        "تحياتي،\n{sender_name}",
        "الكل",
    ),
]


# ======================= محرك النصوص والـ Spintax =======================

class SafeDict(dict):
    def __missing__(self, key):
        return "{" + key + "}"


def resolve_spintax(text: str) -> str:
    """يدعم تنويع النصوص مثل {السلام عليكم|أهلاً بكم|تحياتي} دون المساس بمتغيرات {name}"""
    if not text:
        return ""
    pattern = re.compile(r"\{([^{}]*\|[^{}]*)\}")
    result = str(text)
    for _ in range(10):
        match = pattern.search(result)
        if not match:
            break
        result = pattern.sub(lambda m: random.choice(m.group(1).split("|")), result)
    return result


def render_msg(template_text: str, lead: dict, sender_name: str = "", apply_spintax: bool = True) -> str:
    raw = str(template_text or "")
    if apply_spintax:
        raw = resolve_spintax(raw)
    return raw.format_map(SafeDict(
        name=lead.get("name") or "فريق العمل",
        company_name=lead.get("name") or "فريق العمل",
        category=lead.get("category") or "مجالكم",
        industry=lead.get("category") or "مجالكم",
        budget=lead.get("budget") or "الميزانية المحددة",
        location=lead.get("location") or "",
        website=lead.get("website") or "",
        source=lead.get("source") or "",
        sender_name=sender_name or "",
    ))


# ======================= تنظيف الأرقام والإيميلات =======================

def extract_mobile_number(raw_text) -> str | None:
    """استخراج رقم موبايل صالح لواتساب (يدعم مصر والدولي ويتجاهل الأرضي المصري مثل 050/02)"""
    if not raw_text:
        return None
    text = str(raw_text)

    # 1. الموبايل المصري: 010 / 011 / 012 / 015
    eg_matches = re.findall(r"(?:\+?20|0020|0)?[\s\-\.]?1[0125](?:[\s\-\.]?\d){8}", text)
    for m in eg_matches:
        digits = re.sub(r"\D", "", m)
        if digits.startswith("0020"):
            digits = digits[2:]
        elif digits.startswith("01") and len(digits) == 11:
            digits = "2" + digits
        elif digits.startswith("1") and len(digits) == 10:
            digits = "20" + digits
        if digits.startswith("201") and len(digits) == 12:
            return digits

    # 2. الأرقام الدولية (+966, +971, +1...)
    intl_matches = re.findall(r"(?:\+|00)\d[\d\s\-\.]{8,16}", text)
    for m in intl_matches:
        digits = re.sub(r"\D", "", m)
        if digits.startswith("00"):
            digits = digits[2:]
        if digits.startswith("20") and not digits.startswith("201"):
            continue  # تجاهل الأرضي المصري بصيغة +2050 أو +202
        if 10 <= len(digits) <= 15:
            return digits
    return None


def format_whatsapp_phone(raw_phone) -> str | None:
    if not raw_phone:
        return None
    s = str(raw_phone).strip()
    if any(bad in s for bad in ("غير متوفر", "—", "N/A")):
        return None
    mob = extract_mobile_number(s)
    if mob:
        return mob
    digits = re.sub(r"\D", "", s)
    if digits.startswith("0020"):
        digits = digits[2:]
    elif digits.startswith("01") and len(digits) == 11:
        digits = "2" + digits
    if digits.startswith("201") and len(digits) == 12:
        return digits
    if digits.startswith("20") and not digits.startswith("201"):
        return None
    if not digits.startswith("0") and 10 <= len(digits) <= 15:
        return digits
    return None


def extract_email(raw_text) -> str | None:
    if not raw_text:
        return None
    emails = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", str(raw_text))
    ignored = (
        "sentry", "wixpress", "example.com", "yourdomain", "domain.com", "email.com",
        "noreply", "no-reply", ".png", ".jpg", ".jpeg", ".webp", ".svg", ".gif", ".css", ".js",
    )
    for email in emails:
        low = email.lower()
        if not any(bad in low for bad in ignored) and len(email) <= 65:
            return email
    return None


def is_valid_email(val) -> bool:
    if not val:
        return False
    s = str(val).strip()
    if "غير متوفر" in s or " " in s:
        return False
    return bool(re.match(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$", s))


def enrich_from_website(website_url: str, timeout: int = 6):
    """فحص سريع لموقع الشركة (الرئيسية + صفحة Contact لو متاحة) لاستخراج الإيميل ورقم الواتساب"""
    found_email, found_wa = None, None
    if not website_url or website_url in ("N/A", "غير متوفر"):
        return found_email, found_wa
    if not website_url.startswith(("http://", "https://")):
        website_url = "https://" + website_url

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    try:
        resp = requests.get(website_url, headers=headers, timeout=timeout, allow_redirects=True)
        html = resp.text

        wa_link = re.search(r"(?:wa\.me/|api\.whatsapp\.com/send\?phone=)(\d{10,15})", html)
        if wa_link:
            found_wa = extract_mobile_number("+" + wa_link.group(1)) or wa_link.group(1)
        else:
            found_wa = extract_mobile_number(html)

        found_email = extract_email(html)

        # لو ملقيناش إيميل في الصفحة الرئيسية، نبحث عن رابط Contact ونفحصه
        if not found_email:
            m_contact = re.search(r'href=["\']([^"\']*(?:contact|اتصل|تواصل)[^"\']*)["\']', html, re.I)
            if m_contact:
                c_href = m_contact.group(1)
                if c_href.startswith("/"):
                    base = re.match(r"^(https?://[^/]+)", resp.url)
                    if base:
                        c_href = base.group(1) + c_href
                if c_href.startswith("http"):
                    r2 = requests.get(c_href, headers=headers, timeout=5)
                    found_email = extract_email(r2.text) or found_email
                    if not found_wa:
                        found_wa = extract_mobile_number(r2.text)
    except Exception:
        pass
    return found_email, found_wa


def parse_employee_count(size_text) -> int | None:
    if not size_text:
        return None
    nums = [int(n.replace(",", "")) for n in re.findall(r"\d[\d,]*", str(size_text))]
    return nums[0] if nums else None


def matches_icp_and_contact(lead: dict, cfg: dict) -> tuple[bool, str]:
    has_wa = bool(format_whatsapp_phone(lead.get("phone")))
    has_email = is_valid_email(lead.get("email"))
    req = cfg.get("contact_req", "الكل (أي بيانات متاحة)")

    if req == "أرقام واتساب فقط" and not has_wa:
        return False, "لا يوجد رقم واتساب"
    if req == "إيميلات فقط" and not has_email:
        return False, "لا يوجد إيميل"
    if req == "الاثنين معاً (واتساب + إيميل)" and not (has_wa and has_email):
        return False, "ينقصه إيميل أو واتساب"
    if req == "رقم أو إيميل على الأقل" and not (has_wa or has_email):
        return False, "لا توجد وسيلة تواصل"

    if cfg.get("must_have_web"):
        web = str(lead.get("website") or "").strip()
        if not web or web in ("N/A", "غير متوفر"):
            return False, "بدون موقع إلكتروني"

    ind_filter = (cfg.get("icp_industry") or "").strip()
    if ind_filter:
        tokens = [t.strip().lower() for t in re.split(r"[,،\n]+", ind_filter) if t.strip()]
        haystack = f"{lead.get('name', '')} {lead.get('category', '')}".lower()
        if tokens and not any(tok in haystack for tok in tokens):
            return False, "لا يطابق فلتر التخصص"

    size_pref = cfg.get("icp_size", "الكل")
    if size_pref != "الكل" and lead.get("source") == "LinkedIn":
        emp = parse_employee_count(lead.get("size_rating", ""))
        if emp is not None:
            if size_pref == "صغيرة (1-10)" and emp > 10:
                return False, "الحجم أكبر من المطلوب"
            elif size_pref == "متوسطة (11-50)" and not (11 <= emp <= 50):
                return False, "الحجم خارج النطاق"
            elif size_pref == "كبيرة (51-200)" and not (51 <= emp <= 200):
                return False, "الحجم خارج النطاق"
            elif size_pref == "مؤسسات (+200)" and emp <= 200:
                return False, "الحجم أصغر من المطلوب"

    try:
        min_rating = float(cfg.get("min_rating") or 0)
    except ValueError:
        min_rating = 0.0
    if min_rating > 0 and lead.get("source") == "Google Maps":
        m = re.search(r"(\d(?:\.\d)?)", str(lead.get("size_rating", "")))
        rating_val = float(m.group(1)) if m else 0.0
        if rating_val < min_rating:
            return False, f"التقييم ({rating_val}) أقل من {min_rating}"

    return True, "مطابق"


def session_is_ready(folder_path: str) -> bool:
    return os.path.isdir(folder_path) and len(os.listdir(folder_path)) > 0


# ======================= قاعدة البيانات الآمنة للتزامن (SQLite WAL) =======================

class DatabaseManager:
    """قاعدة بيانات SQLite بوضع WAL + Thread Lock لدعم السحب والإرسال في نفس الوقت بدون أي تعليق"""

    def __init__(self, db_path: str = DB_FILE):
        self.db_path = db_path
        self._lock = threading.RLock()
        self._init_db()

    def _connect(self):
        conn = sqlite3.connect(self.db_path, timeout=15.0, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
        conn.execute("PRAGMA busy_timeout=15000;")
        return conn

    def _init_db(self):
        with self._lock, self._connect() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS leads (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    source TEXT DEFAULT 'يدوي',
                    phone TEXT DEFAULT '',
                    email TEXT DEFAULT '',
                    website TEXT DEFAULT '',
                    profile_url TEXT DEFAULT '',
                    category TEXT DEFAULT '',
                    size_rating TEXT DEFAULT '',
                    budget TEXT DEFAULT '',
                    location TEXT DEFAULT '',
                    stage TEXT DEFAULT 'جديد',
                    next_followup TEXT DEFAULT '',
                    notes TEXT DEFAULT '',
                    email_status TEXT DEFAULT '',
                    wa_status TEXT DEFAULT '',
                    li_status TEXT DEFAULT '',
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS templates (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    title TEXT UNIQUE NOT NULL,
                    subject TEXT DEFAULT '',
                    body TEXT NOT NULL,
                    channel TEXT DEFAULT 'الكل'
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS activity_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    lead_id INTEGER,
                    channel TEXT NOT NULL,
                    action_date TEXT NOT NULL,
                    status TEXT NOT NULL,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)

            if conn.execute("SELECT COUNT(*) FROM templates").fetchone()[0] == 0:
                conn.executemany(
                    "INSERT OR IGNORE INTO templates (title, subject, body, channel) VALUES (?, ?, ?, ?)",
                    DEFAULT_TEMPLATES,
                )

            defaults = {
                "appearance": "Dark",
                "keywords": "digital marketing agency, software agency",
                "location": "Mansoura",
                "target_count": "50",
                "contact_req": "الكل (أي بيانات متاحة)",
                "icp_industry": "",
                "icp_size": "الكل",
                "min_rating": "0",
                "sender_email": "",
                "app_password": "",
                "sender_name": "",
                "attachment_path": "",
                "li_verified": "0",
                "wa_verified": "0",
                "gmail_verified": "0",
                # إعدادات درع الحماية من الحظر (Anti-Ban)
                "safety_profile": "متوازن (آمن)",
                "daily_email_cap": "35",
                "daily_wa_cap": "20",
                "daily_li_cap": "15",
                "email_min_delay": "35",
                "email_max_delay": "75",
                "wa_min_delay": "50",
                "wa_max_delay": "95",
                "li_min_delay": "45",
                "li_max_delay": "85",
                "batch_size": "5",
                "batch_break_sec": "180",
            }
            for k, v in defaults.items():
                conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (k, v))

    def get_settings(self) -> dict:
        with self._lock, self._connect() as conn:
            return {r["key"]: r["value"] for r in conn.execute("SELECT key, value FROM settings")}

    def save_settings(self, data: dict):
        with self._lock, self._connect() as conn:
            for k, v in data.items():
                conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (k, str(v)))

    # --- تتبع الكوتا اليومية (Anti-Ban Daily Quotas) ---
    def record_activity(self, lead_id: int | None, channel: str, status: str = "sent"):
        today = datetime.now().strftime("%Y-%m-%d")
        with self._lock, self._connect() as conn:
            conn.execute(
                "INSERT INTO activity_log (lead_id, channel, action_date, status) VALUES (?, ?, ?, ?)",
                (lead_id, channel, today, status),
            )

    def get_today_counts(self) -> dict:
        today = datetime.now().strftime("%Y-%m-%d")
        counts = {"email": 0, "whatsapp": 0, "linkedin": 0}
        with self._lock, self._connect() as conn:
            rows = conn.execute(
                "SELECT channel, COUNT(*) as c FROM activity_log WHERE action_date = ? AND status = 'sent' GROUP BY channel",
                (today,),
            ).fetchall()
            for r in rows:
                if r["channel"] in counts:
                    counts[r["channel"]] = r["c"]
        return counts

    # --- إدارة العملاء ---
    def get_leads(self, search: str = "", stage_filter: str = "الكل", quick_filter: str = "الكل") -> list[dict]:
        query = "SELECT * FROM leads WHERE 1=1"
        params = []
        if stage_filter and stage_filter != "الكل":
            query += " AND stage = ?"
            params.append(stage_filter)
        if quick_filter == "بها إيميل ولم تُرسل":
            query += " AND email != '' AND email_status NOT LIKE 'تم%' AND stage != 'عدم الإزعاج'"
        elif quick_filter == "بها واتساب ولم تُرسل":
            query += " AND phone != '' AND wa_status NOT LIKE 'تم%' AND wa_status NOT LIKE 'غير مسجل%' AND stage != 'عدم الإزعاج'"
        elif quick_filter == "عملاء LinkedIn":
            query += " AND profile_url LIKE '%linkedin.com/company%'"
        elif quick_filter == "المضاف اليوم":
            today = datetime.now().strftime("%Y-%m-%d")
            query += " AND created_at LIKE ?"
            params.append(f"{today}%")

        if search.strip():
            like = f"%{search.strip()}%"
            query += " AND (name LIKE ? OR phone LIKE ? OR email LIKE ? OR category LIKE ? OR location LIKE ? OR source LIKE ? OR notes LIKE ?)"
            params.extend([like] * 7)
        query += " ORDER BY id DESC"
        with self._lock, self._connect() as conn:
            return [dict(r) for r in conn.execute(query, params).fetchall()]

    def get_lead_by_id(self, lead_id: int) -> dict | None:
        with self._lock, self._connect() as conn:
            row = conn.execute("SELECT * FROM leads WHERE id = ?", (lead_id,)).fetchone()
            return dict(row) if row else None

    def lead_exists(self, name: str, profile_url: str = "", phone: str = "", email: str = "") -> bool:
        with self._lock, self._connect() as conn:
            if profile_url and profile_url not in ("N/A", ""):
                if conn.execute("SELECT 1 FROM leads WHERE LOWER(profile_url) = LOWER(?)", (profile_url.strip(),)).fetchone():
                    return True
            if email and is_valid_email(email):
                if conn.execute("SELECT 1 FROM leads WHERE LOWER(email) = LOWER(?)", (email.strip(),)).fetchone():
                    return True
            if phone:
                norm = format_whatsapp_phone(phone)
                if norm and conn.execute("SELECT 1 FROM leads WHERE phone = ?", (norm,)).fetchone():
                    return True
            if name and conn.execute("SELECT 1 FROM leads WHERE LOWER(name) = LOWER(?)", (name.strip(),)).fetchone():
                return True
            return False

    def add_lead(self, data: dict) -> int:
        with self._lock, self._connect() as conn:
            cur = conn.execute("""
                INSERT INTO leads (
                    name, source, phone, email, website, profile_url,
                    category, size_rating, budget, location, stage, next_followup, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                data["name"], data.get("source", "يدوي"), data.get("phone", ""),
                data.get("email", ""), data.get("website", ""), data.get("profile_url", ""),
                data.get("category", ""), data.get("size_rating", ""), data.get("budget", ""),
                data.get("location", ""), data.get("stage", "جديد"),
                data.get("next_followup", ""), data.get("notes", ""),
            ))
            return cur.lastrowid

    def update_lead(self, lead_id: int, data: dict):
        with self._lock, self._connect() as conn:
            conn.execute("""
                UPDATE leads
                SET name=?, source=?, phone=?, email=?, website=?, profile_url=?,
                    category=?, size_rating=?, budget=?, location=?, stage=?, next_followup=?, notes=?
                WHERE id=?
            """, (
                data["name"], data.get("source", "يدوي"), data.get("phone", ""),
                data.get("email", ""), data.get("website", ""), data.get("profile_url", ""),
                data.get("category", ""), data.get("size_rating", ""), data.get("budget", ""),
                data.get("location", ""), data.get("stage", "جديد"),
                data.get("next_followup", ""), data.get("notes", ""), lead_id,
            ))

    def update_stage(self, lead_id: int, stage: str):
        with self._lock, self._connect() as conn:
            conn.execute("UPDATE leads SET stage = ? WHERE id = ?", (stage, lead_id))

    def update_channel_status(self, lead_id: int, channel_col: str, status_text: str, promote_stage: bool = False):
        if channel_col not in ("email_status", "wa_status", "li_status"):
            return
        with self._lock, self._connect() as conn:
            if promote_stage:
                row = conn.execute("SELECT stage FROM leads WHERE id = ?", (lead_id,)).fetchone()
                if row and row["stage"] == "جديد":
                    conn.execute(f"UPDATE leads SET {channel_col} = ?, stage = 'تم التواصل' WHERE id = ?", (status_text, lead_id))
                    return
            conn.execute(f"UPDATE leads SET {channel_col} = ? WHERE id = ?", (status_text, lead_id))

    def delete_leads(self, lead_ids):
        if not lead_ids:
            return
        with self._lock, self._connect() as conn:
            placeholders = ",".join("?" for _ in lead_ids)
            conn.execute(f"DELETE FROM leads WHERE id IN ({placeholders})", list(lead_ids))

    def get_pipeline_stats(self):
        with self._lock, self._connect() as conn:
            total = conn.execute("SELECT COUNT(*) FROM leads").fetchone()[0]
            counts = {s: 0 for s in STAGES}
            for row in conn.execute("SELECT stage, COUNT(*) as c FROM leads GROUP BY stage"):
                if row["stage"] in counts:
                    counts[row["stage"]] = row["c"]
            today = datetime.now().strftime("%Y-%m-%d")
            due = conn.execute(
                "SELECT * FROM leads WHERE next_followup != '' AND next_followup <= ? AND stage NOT IN ('تم البيع', 'ملغي', 'عدم الإزعاج') ORDER BY next_followup ASC",
                (today,),
            ).fetchall()
            return total, counts, [dict(r) for r in due]

    # --- القوالب ---
    def get_templates(self) -> list[dict]:
        with self._lock, self._connect() as conn:
            return [dict(r) for r in conn.execute("SELECT * FROM templates ORDER BY id ASC").fetchall()]

    def save_template(self, title: str, subject: str, body: str, channel: str):
        with self._lock, self._connect() as conn:
            conn.execute("""
                INSERT INTO templates (title, subject, body, channel)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(title) DO UPDATE SET subject=excluded.subject, body=excluded.body, channel=excluded.channel
            """, (title, subject, body, channel))

    def delete_template(self, title: str):
        with self._lock, self._connect() as conn:
            conn.execute("DELETE FROM templates WHERE title = ?", (title,))


# ======================= تصدير واستيراد Excel =======================

def export_leads_to_excel(leads: list[dict], filepath: str, template_subject: str, template_body: str, sender_name: str = ""):
    wb = Workbook()
    ws = wb.active
    ws.title = "Smart CRM Leads"

    headers = [
        "ID", "اسم العميل / الشركة", "المصدر", "التخصص", "الحجم/التقييم",
        "المنطقة", "الموقع الإلكتروني", "واتساب / موبايل", "الإيميل", "رابط الحساب",
        "المرحلة", "حالة الإيميل", "حالة واتساب", "حالة LinkedIn",
        "فتح واتساب", "إرسال Gmail",
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    center = Alignment(horizontal="center", vertical="center")
    border = Border(
        left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"),
    )

    for col in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col)
        c.fill, c.font, c.alignment = header_fill, header_font, center
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(2, len(leads) + 1)}"

    wa_fill = PatternFill(start_color="16A34A", end_color="16A34A", fill_type="solid")
    mail_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    off_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    btn_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    off_font = Font(name="Segoe UI", size=10, color="64748B")

    for r_idx, lead in enumerate(leads, start=2):
        wa_num = format_whatsapp_phone(lead.get("phone"))
        email = lead.get("email", "")

        ws.append([
            lead["id"], lead["name"], lead.get("source") or "", lead.get("category") or "",
            lead.get("size_rating") or "", lead.get("location") or "", lead.get("website") or "",
            wa_num or lead.get("phone") or "غير متوفر", email or "غير متوفر",
            lead.get("profile_url") or "", lead.get("stage") or "جديد",
            lead.get("email_status") or "", lead.get("wa_status") or "", lead.get("li_status") or "",
            "", "",
        ])
        ws.row_dimensions[r_idx].height = 24

        wa_c = ws.cell(row=r_idx, column=15)
        wa_c.alignment, wa_c.border = center, border
        if wa_num:
            msg = quote(render_msg(template_body, lead, sender_name))
            wa_c.value = "💬 واتساب"
            wa_c.hyperlink = f"https://wa.me/{wa_num}?text={msg}"
            wa_c.fill, wa_c.font = wa_fill, btn_font
        else:
            wa_c.value = "لا يوجد رقم"
            wa_c.fill, wa_c.font = off_fill, off_font

        em_c = ws.cell(row=r_idx, column=16)
        em_c.alignment, em_c.border = center, border
        if is_valid_email(email):
            subj = quote(render_msg(template_subject, lead, sender_name))
            body = quote(render_msg(template_body, lead, sender_name))
            em_c.value = "✉️ Gmail"
            em_c.hyperlink = f"https://mail.google.com/mail/?view=cm&fs=1&to={email}&su={subj}&body={body}"
            em_c.fill, em_c.font = mail_fill, btn_font
        else:
            em_c.value = "لا يوجد إيميل"
            em_c.fill, em_c.font = off_fill, off_font

    widths = [7, 25, 14, 18, 14, 14, 24, 17, 25, 30, 14, 18, 18, 18, 14, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(filepath)


def import_leads_from_excel(db: DatabaseManager, filepath: str) -> int:
    wb = load_workbook(filepath, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if len(rows) < 2:
        return 0

    headers = [str(h or "").strip().lower() for h in rows[0]]

    def find_col(keywords):
        for idx, h in enumerate(headers):
            if any(k in h for k in keywords):
                return idx
        return None

    c_name = find_col(["name", "company", "اسم", "العميل", "الشركة"])
    if c_name is None:
        c_name = 0
    c_phone = find_col(["phone", "whatsapp", "mobile", "هاتف", "واتساب", "رقم"])
    c_email = find_col(["email", "mail", "إيميل", "ايميل"])
    c_cat = find_col(["industry", "category", "تخصص", "مجال", "طلب"])
    c_loc = find_col(["location", "city", "لوكيشن", "مدينة", "منطقة"])
    c_web = find_col(["website", "موقع"])
    c_url = find_col(["linkedin", "profile", "url", "رابط"])
    c_src = find_col(["source", "مصدر"])

    added = 0
    for row in rows[1:]:
        if not row or len(row) <= c_name or not row[c_name]:
            continue
        name = str(row[c_name]).strip()
        phone = str(row[c_phone]).strip() if c_phone is not None and len(row) > c_phone and row[c_phone] else ""
        email = str(row[c_email]).strip() if c_email is not None and len(row) > c_email and row[c_email] else ""
        url = str(row[c_url]).strip() if c_url is not None and len(row) > c_url and row[c_url] else ""

        phone = "" if "غير متوفر" in phone else phone
        email = "" if "غير متوفر" in email else email

        if db.lead_exists(name, url, phone, email):
            continue

        db.add_lead({
            "name": name,
            "source": str(row[c_src]).strip() if c_src is not None and len(row) > c_src and row[c_src] else "Excel",
            "phone": phone,
            "email": email,
            "website": str(row[c_web]).strip() if c_web is not None and len(row) > c_web and row[c_web] else "",
            "profile_url": url,
            "category": str(row[c_cat]).strip() if c_cat is not None and len(row) > c_cat and row[c_cat] else "",
            "location": str(row[c_loc]).strip() if c_loc is not None and len(row) > c_loc and row[c_loc] else "",
        })
        added += 1
    return added


# ======================= إرسال إيميل فردي عبر Gmail SMTP =======================

def send_single_smtp_email(sender_email: str, app_password: str, sender_name: str, to_email: str, subject: str, body: str, attachment_path: str = ""):
    msg = EmailMessage()
    msg["From"] = f"{sender_name} <{sender_email}>" if sender_name else sender_email
    msg["To"] = to_email.strip()
    msg["Subject"] = subject
    msg.set_content(body, charset="utf-8")

    if attachment_path and os.path.isfile(attachment_path):
        mt, _ = mimetypes.guess_type(attachment_path)
        main_t, sub_t = (mt or "application/octet-stream").split("/", 1)
        with open(attachment_path, "rb") as f:
            msg.add_attachment(f.read(), maintype=main_t, subtype=sub_t, filename=os.path.basename(attachment_path))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
        smtp.login(sender_email.strip(), app_password.replace(" ", ""))
        smtp.send_message(msg)
