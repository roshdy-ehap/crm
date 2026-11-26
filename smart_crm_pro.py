import os
import re
import time
import queue
import shutil
import random
import sqlite3
import smtplib
import platform
import threading
import mimetypes
import webbrowser
import subprocess
from datetime import datetime
from urllib.parse import quote
from email.message import EmailMessage

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk

import requests
from playwright.sync_api import sync_playwright
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB_FILE = "crm_pro.db"
LINKEDIN_SESSION_DIR = "./session_linkedin"
WA_SESSION_DIR = "./session_whatsapp"

STAGES = ["جديد", "تم التواصل", "مهتم", "معاينة", "تم البيع", "ملغي"]
STAGE_COLORS = {
    "جديد": "#3B82F6",
    "تم التواصل": "#8B5CF6",
    "مهتم": "#F59E0B",
    "معاينة": "#06B6D4",
    "تم البيع": "#10B981",
    "ملغي": "#EF4444",
}

DEFAULT_TEMPLATES = [
    (
        "ترحيب عام",
        "تعاون محتمل مع {name}",
        "السلام عليكم فريق {name}،\nأتمنى تكونوا بخير.\nسعداء بالتواصل معكم بخصوص مجال {category} في {location}.\nلو وقتكم يسمح، ممكن نرتب مكالمة سريعة؟\n\nتحياتي،\n{sender_name}",
        "الكل",
    ),
    (
        "عرض مبيعات مخصص",
        "فرصة تعاون مع {name}",
        "أهلاً فريق {name}،\nاطلعنا على شغلكم في {category} بمنطقة {location} وعندنا عرض مناسب جداً.\nتحب نبعتلك التفاصيل الكاملة؟",
        "WhatsApp",
    ),
    (
        "متابعة بعد التواصل",
        "متابعة بخصوص {category}",
        "السلام عليكم أ/ {name}،\nحبينا نتابع معاك بخصوص تواصلنا السابق وهل فيه أي تفاصيل إضافية محتاج نوضحها؟\n\nتحياتي،\n{sender_name}",
        "الكل",
    ),
]


# ======================= دوال مساعدة وتنظيف البيانات =======================

class SafeDict(dict):
    def __missing__(self, key):
        return "{" + key + "}"


def render_msg(template_text, lead, sender_name=""):
    return (template_text or "").format_map(SafeDict(
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


def extract_mobile_number(raw_text):
    """استخراج رقم موبايل صالح لواتساب (يدعم مصر والدولي ويتجاهل الأرضي)"""
    if not raw_text:
        return None
    text = str(raw_text)
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

    intl_matches = re.findall(r"(?:\+|00)\d[\d\s\-\.]{9,16}", text)
    for m in intl_matches:
        digits = re.sub(r"\D", "", m)
        if digits.startswith("00"):
            digits = digits[2:]
        if digits.startswith("20") and not digits.startswith("201"):
            continue
        if 10 <= len(digits) <= 15:
            return digits
    return None


def format_whatsapp_phone(raw_phone):
    if not raw_phone:
        return None
    s = str(raw_phone).strip()
    if "غير متوفر" in s or s in ("—", "N/A"):
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
    if not digits.startswith(("0", "20")) and 10 <= len(digits) <= 15:
        return digits
    return None


def extract_email(raw_text):
    if not raw_text:
        return None
    emails = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", str(raw_text))
    ignored = (
        "sentry", "wixpress", "example.com", "yourdomain", "domain.com",
        ".png", ".jpg", ".jpeg", ".webp", ".svg", ".gif", ".css", ".js",
    )
    for email in emails:
        low = email.lower()
        if not any(bad in low for bad in ignored) and len(email) <= 65:
            return email
    return None


def is_valid_email(val):
    if not val:
        return False
    s = str(val).strip()
    return "@" in s and "." in s and "غير متوفر" not in s


def enrich_from_website(website_url):
    """فحص سريع لموقع الشركة لاستخراج الإيميل ورقم الواتساب"""
    found_email, found_wa = None, None
    if not website_url or website_url in ("N/A", "غير متوفر"):
        return found_email, found_wa
    if not website_url.startswith(("http://", "https://")):
        website_url = "https://" + website_url
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        resp = requests.get(website_url, headers=headers, timeout=6)
        html = resp.text
        wa_link = re.search(r"(?:wa\.me/|api\.whatsapp\.com/send\?phone=)(\d{10,15})", html)
        if wa_link:
            found_wa = extract_mobile_number("+" + wa_link.group(1)) or wa_link.group(1)
        else:
            found_wa = extract_mobile_number(html)
        found_email = extract_email(html)
    except Exception:
        pass
    return found_email, found_wa


def parse_employee_count(size_text):
    if not size_text:
        return None
    nums = [int(n.replace(",", "")) for n in re.findall(r"\d[\d,]*", str(size_text))]
    return nums[0] if nums else None


def session_is_ready(folder_path):
    return os.path.isdir(folder_path) and len(os.listdir(folder_path)) > 0


def launch_persistent_browser(p, session_dir):
    """فتح متصفح ذكي يجرب Chrome ثم Edge ثم Chromium الافتراضي لضمان العمل داخل EXE"""
    for ch in ("chrome", "msedge", None):
        try:
            kwargs = {
                "user_data_dir": session_dir,
                "headless": False,
                "viewport": {"width": 1280, "height": 800},
            }
            if ch:
                kwargs["channel"] = ch
            return p.chromium.launch_persistent_context(**kwargs)
        except Exception:
            continue
    raise RuntimeError("لم يتم العثور على متصفح Chrome أو Edge أو Chromium. ثبت Google Chrome أو شغل playwright install chromium.")


def launch_temp_browser(p):
    for ch in ("chrome", "msedge", None):
        try:
            kwargs = {"headless": False}
            if ch:
                kwargs["channel"] = ch
            return p.chromium.launch(**kwargs)
        except Exception:
            continue
    raise RuntimeError("لم يتم العثور على متصفح Chrome أو Edge أو Chromium.")


def matches_icp_and_contact(lead, cfg):
    has_wa = bool(format_whatsapp_phone(lead.get("phone")))
    has_email = is_valid_email(lead.get("email"))
    req = cfg["contact_req"]

    if req == "أرقام واتساب فقط" and not has_wa:
        return False, "لا يوجد رقم واتساب"
    if req == "إيميلات فقط" and not has_email:
        return False, "لا يوجد إيميل"
    if req == "الاثنين معاً (واتساب + إيميل)" and not (has_wa and has_email):
        return False, "ينقصه إيميل أو واتساب"
    if req == "رقم أو إيميل على الأقل" and not (has_wa or has_email):
        return False, "لا توجد وسيلة تواصل"

    if cfg["must_have_web"]:
        web = str(lead.get("website") or "").strip()
        if not web or web in ("N/A", "غير متوفر"):
            return False, "بدون موقع إلكتروني"

    ind_filter = cfg["icp_industry"].strip()
    if ind_filter:
        tokens = [t.strip().lower() for t in re.split(r"[,،\n]+", ind_filter) if t.strip()]
        haystack = f"{lead.get('name', '')} {lead.get('category', '')}".lower()
        if tokens and not any(tok in haystack for tok in tokens):
            return False, "لا يطابق فلتر التخصص"

    size_pref = cfg["icp_size"]
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
        min_rating = float(cfg["min_rating"] or 0)
    except ValueError:
        min_rating = 0.0
    if min_rating > 0 and lead.get("source") == "Google Maps":
        m = re.search(r"(\d(?:\.\d)?)", str(lead.get("size_rating", "")))
        rating_val = float(m.group(1)) if m else 0.0
        if rating_val < min_rating:
            return False, f"التقييم ({rating_val}) أقل من {min_rating}"

    return True, "مطابق"


# ======================= طبقة قاعدة البيانات (SQLite) =======================

class DatabaseManager:
    def __init__(self, db_path=DB_FILE):
        self.db_path = db_path
        self._init_db()

    def _connect(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        with self._connect() as conn:
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
                "contact_req": "رقم أو إيميل على الأقل",
                "icp_industry": "",
                "icp_size": "الكل",
                "min_rating": "0",
                "sender_email": "",
                "app_password": "",
                "sender_name": "",
                "attachment_path": "",
                "max_emails": "35",
                "email_min": "35",
                "email_max": "70",
                "max_wa": "20",
                "wa_min": "45",
                "wa_max": "85",
                "max_li": "15",
                "li_min": "40",
                "li_max": "80",
            }
            for k, v in defaults.items():
                conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (k, v))

    def get_settings(self):
        with self._connect() as conn:
            return {r["key"]: r["value"] for r in conn.execute("SELECT key, value FROM settings")}

    def save_settings(self, data: dict):
        with self._connect() as conn:
            for k, v in data.items():
                conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (k, str(v)))

    def get_leads(self, search="", stage_filter="الكل", quick_filter="الكل"):
        query = "SELECT * FROM leads WHERE 1=1"
        params = []
        if stage_filter and stage_filter != "الكل":
            query += " AND stage = ?"
            params.append(stage_filter)
        if quick_filter == "بها إيميل ولم تُرسل":
            query += " AND email != '' AND email_status NOT LIKE 'تم%'"
        elif quick_filter == "بها واتساب ولم تُرسل":
            query += " AND phone != '' AND wa_status NOT LIKE 'تم%' AND wa_status NOT LIKE 'غير مسجل%'"
        elif quick_filter == "عملاء LinkedIn":
            query += " AND profile_url LIKE '%linkedin.com/company%'"

        if search.strip():
            like = f"%{search.strip()}%"
            query += " AND (name LIKE ? OR phone LIKE ? OR email LIKE ? OR category LIKE ? OR location LIKE ? OR source LIKE ? OR notes LIKE ?)"
            params.extend([like] * 7)
        query += " ORDER BY id DESC"
        with self._connect() as conn:
            return [dict(r) for r in conn.execute(query, params).fetchall()]

    def get_lead_by_id(self, lead_id):
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM leads WHERE id = ?", (lead_id,)).fetchone()
            return dict(row) if row else None

    def lead_exists(self, name, profile_url=""):
        with self._connect() as conn:
            if profile_url and profile_url not in ("N/A", ""):
                r = conn.execute("SELECT 1 FROM leads WHERE LOWER(profile_url) = LOWER(?)", (profile_url.strip(),)).fetchone()
                if r:
                    return True
            r = conn.execute("SELECT 1 FROM leads WHERE LOWER(name) = LOWER(?)", (name.strip(),)).fetchone()
            return bool(r)

    def add_lead(self, data: dict):
        with self._connect() as conn:
            conn.execute("""
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

    def update_lead(self, lead_id, data: dict):
        with self._connect() as conn:
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

    def update_stage(self, lead_id, stage):
        with self._connect() as conn:
            conn.execute("UPDATE leads SET stage = ? WHERE id = ?", (stage, lead_id))

    def update_channel_status(self, lead_id, channel_col, status_text, promote_stage=False):
        if channel_col not in ("email_status", "wa_status", "li_status"):
            return
        with self._connect() as conn:
            if promote_stage:
                row = conn.execute("SELECT stage FROM leads WHERE id = ?", (lead_id,)).fetchone()
                if row and row["stage"] == "جديد":
                    conn.execute(f"UPDATE leads SET {channel_col} = ?, stage = 'تم التواصل' WHERE id = ?", (status_text, lead_id))
                    return
            conn.execute(f"UPDATE leads SET {channel_col} = ? WHERE id = ?", (status_text, lead_id))

    def delete_leads(self, lead_ids):
        if not lead_ids:
            return
        with self._connect() as conn:
            placeholders = ",".join("?" for _ in lead_ids)
            conn.execute(f"DELETE FROM leads WHERE id IN ({placeholders})", list(lead_ids))

    def get_pipeline_stats(self):
        with self._connect() as conn:
            total = conn.execute("SELECT COUNT(*) FROM leads").fetchone()[0]
            counts = {s: 0 for s in STAGES}
            for row in conn.execute("SELECT stage, COUNT(*) as c FROM leads GROUP BY stage"):
                if row["stage"] in counts:
                    counts[row["stage"]] = row["c"]
            today = datetime.now().strftime("%Y-%m-%d")
            due = conn.execute(
                "SELECT * FROM leads WHERE next_followup != '' AND next_followup <= ? AND stage NOT IN ('تم البيع', 'ملغي') ORDER BY next_followup ASC",
                (today,),
            ).fetchall()
            return total, counts, [dict(r) for r in due]

    def get_templates(self):
        with self._connect() as conn:
            return [dict(r) for r in conn.execute("SELECT * FROM templates ORDER BY id ASC").fetchall()]

    def save_template(self, title, subject, body, channel):
        with self._connect() as conn:
            conn.execute("""
                INSERT INTO templates (title, subject, body, channel)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(title) DO UPDATE SET subject=excluded.subject, body=excluded.body, channel=excluded.channel
            """, (title, subject, body, channel))

    def delete_template(self, title):
        with self._connect() as conn:
            conn.execute("DELETE FROM templates WHERE title = ?", (title,))


# ======================= تصدير واستيراد Excel =======================

def export_leads_to_excel(leads, filepath, template_subject, template_body, sender_name=""):
    wb = Workbook()
    ws = wb.active
    ws.title = "Smart CRM Leads"

    headers = [
        "ID", "اسم العميل / الشركة", "المصدر", "التخصص", "الحجم/التقييم",
        "المنطقة", "الموقع الإلكتروني", "واتساب / موبايل", "الإيميل", "رابط الحساب",
        "المرحلة", "حالة الإيميل", "حالة واتساب", "حالة LinkedIn",
        "زر واتساب", "زر Gmail",
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
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

    wa_fill = PatternFill(start_color="25D366", end_color="25D366", fill_type="solid")
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

    widths = [7, 24, 13, 18, 14, 14, 24, 17, 25, 30, 13, 18, 18, 18, 14, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(filepath)


def import_leads_from_excel(db: DatabaseManager, filepath):
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

    c_name = find_col(["name", "company", "اسم", "العميل", "الشركة"]) or 0
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
        url = str(row[c_url]).strip() if c_url is not None and len(row) > c_url and row[c_url] else ""
        if db.lead_exists(name, url):
            continue

        phone = str(row[c_phone]).strip() if c_phone is not None and len(row) > c_phone and row[c_phone] else ""
        email = str(row[c_email]).strip() if c_email is not None and len(row) > c_email and row[c_email] else ""
        db.add_lead({
            "name": name,
            "source": str(row[c_src]).strip() if c_src is not None and len(row) > c_src and row[c_src] else "Excel",
            "phone": "" if "غير متوفر" in phone else phone,
            "email": "" if "غير متوفر" in email else email,
            "website": str(row[c_web]).strip() if c_web is not None and len(row) > c_web and row[c_web] else "",
            "profile_url": url,
            "category": str(row[c_cat]).strip() if c_cat is not None and len(row) > c_cat and row[c_cat] else "",
            "location": str(row[c_loc]).strip() if c_loc is not None and len(row) > c_loc and row[c_loc] else "",
        })
        added += 1
    return added


# ======================= نافذة إضافة / تعديل عميل =======================

class LeadDialog(ctk.CTkToplevel):
    def __init__(self, parent, on_save, lead=None):
        super().__init__(parent)
        self.on_save = on_save
        self.lead = lead
        self.title("تعديل بيانات العميل" if lead else "إضافة عميل جديد")
        self.geometry("540x640")
        self.resizable(False, False)
        self.grab_set()
        self._build_form()

    def _build_form(self):
        box = ctk.CTkScrollableFrame(self)
        box.pack(fill="both", expand=True, padx=16, pady=12)

        def add_field(lbl, val=""):
            ctk.CTkLabel(box, text=lbl, font=("Segoe UI", 12, "bold"), anchor="e").pack(fill="x", pady=(5, 2))
            e = ctk.CTkEntry(box, justify="right")
            e.pack(fill="x", pady=(0, 4))
            if val:
                e.insert(0, str(val))
            return e

        L = self.lead or {}
        self.e_name = add_field("اسم العميل أو الشركة *:", L.get("name", ""))
        self.e_phone = add_field("رقم الموبايل / واتساب:", L.get("phone", ""))
        self.e_email = add_field("البريد الإلكتروني:", L.get("email", ""))
        self.e_cat = add_field("التخصص / الطلب:", L.get("category", ""))
        self.e_budget = add_field("الميزانية:", L.get("budget", ""))
        self.e_loc = add_field("المدينة / المنطقة:", L.get("location", ""))
        self.e_web = add_field("الموقع الإلكتروني:", L.get("website", ""))
        self.e_url = add_field("رابط LinkedIn أو Google Maps:", L.get("profile_url", ""))
        self.e_src = add_field("المصدر:", L.get("source", "يدوي"))

        ctk.CTkLabel(box, text="مرحلة المبيعات:", font=("Segoe UI", 12, "bold"), anchor="e").pack(fill="x", pady=(5, 2))
        self.cmb_stage = ctk.CTkComboBox(box, values=STAGES, state="readonly")
        self.cmb_stage.set(L.get("stage", "جديد"))
        self.cmb_stage.pack(fill="x", pady=(0, 4))

        self.e_followup = add_field("موعد المتابعة القادم (YYYY-MM-DD):", L.get("next_followup", ""))

        ctk.CTkLabel(box, text="ملاحظات:", font=("Segoe UI", 12, "bold"), anchor="e").pack(fill="x", pady=(5, 2))
        self.txt_notes = ctk.CTkTextbox(box, height=70)
        self.txt_notes.pack(fill="x", pady=(0, 6))
        if L.get("notes"):
            self.txt_notes.insert("1.0", L["notes"])

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkButton(btns, text="حفظ", fg_color="#10B981", hover_color="#059669", command=self._submit).pack(side="right", padx=6, expand=True, fill="x")
        ctk.CTkButton(btns, text="إلغاء", fg_color="#64748B", command=self.destroy).pack(side="left", padx=6)

    def _submit(self):
        name = self.e_name.get().strip()
        if not name:
            messagebox.showwarning("حقل مطلوب", "اكتب اسم العميل أو الشركة.", parent=self)
            return
        data = {
            "name": name, "phone": self.e_phone.get().strip(), "email": self.e_email.get().strip(),
            "category": self.e_cat.get().strip(), "budget": self.e_budget.get().strip(),
            "location": self.e_loc.get().strip(), "website": self.e_web.get().strip(),
            "profile_url": self.e_url.get().strip(), "source": self.e_src.get().strip() or "يدوي",
            "size_rating": (self.lead or {}).get("size_rating", ""),
            "stage": self.cmb_stage.get(), "next_followup": self.e_followup.get().strip(),
            "notes": self.txt_notes.get("1.0", "end-1c").strip(),
        }
        self.on_save(data, self.lead["id"] if self.lead else None)
        self.destroy()


# ======================= التطبيق الرئيسي المتكامل =======================

class SmartCRMFullApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()
        self.settings = self.db.get_settings()

        ctk.set_appearance_mode(self.settings.get("appearance", "Dark"))
        ctk.set_default_color_theme("blue")

        self.title("Smart Sales CRM Pro — سحب العملاء والاستهداف متعدد القنوات")
        self.geometry("1240x780")
        self.minsize(1040, 660)

        self.selected_lead_ids = set()
        self.log_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.is_running = False

        self._configure_treeview_style()
        self._build_layout()
        self.refresh_account_badges()
        self._show_page("scraper")
        self._poll_logs()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_treeview_style(self):
        is_dark = ctk.get_appearance_mode() == "Dark"
        bg = "#1E293B" if is_dark else "#FFFFFF"
        fg = "#F8FAFC" if is_dark else "#0F172A"
        head_bg = "#0F172A" if is_dark else "#E2E8F0"

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("CRM.Treeview", background=bg, foreground=fg, fieldbackground=bg, rowheight=30, font=("Segoe UI", 9), borderwidth=0)
        style.configure("CRM.Treeview.Heading", background=head_bg, foreground=fg, font=("Segoe UI", 9, "bold"), relief="flat", padding=5)
        style.map("CRM.Treeview", background=[("selected", "#2563EB")], foreground=[("selected", "#FFFFFF")])

    def _build_layout(self):
        self.sidebar = ctk.CTkFrame(self, width=225, corner_radius=0)
        self.sidebar.pack(side="right", fill="y")
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(self.sidebar, text="Smart CRM Pro", font=("Segoe UI", 20, "bold"), text_color="#38BDF8").pack(pady=(20, 2))
        ctk.CTkLabel(self.sidebar, text="سحب عملاء + CRM + حملات", font=("Segoe UI", 11), text_color="#94A3B8").pack(pady=(0, 18))

        self.nav_buttons = {}
        nav_items = [
            ("dashboard", "📊  لوحة التحكم"),
            ("scraper", "🔍  سحب العملاء (Scraping)"),
            ("leads", "👥  العملاء ومراحل البيع"),
            ("templates", "📝  قوالب الرسائل"),
            ("campaigns", "🚀  الحملات وربط الحسابات"),
        ]
        for key, text in nav_items:
            btn = ctk.CTkButton(
                self.sidebar, text=text, anchor="e", height=42,
                font=("Segoe UI", 13, "bold"), corner_radius=8,
                fg_color="transparent", hover_color=("#CBD5E1", "#334155"),
                text_color=("#0F172A", "#F8FAFC"),
                command=lambda k=key: self._show_page(k),
            )
            btn.pack(fill="x", padx=12, pady=3)
            self.nav_buttons[key] = btn

        # مؤشرات حالة الربط في الشريط الجانبي
        badge_box = ctk.CTkFrame(self.sidebar, corner_radius=10)
        badge_box.pack(fill="x", padx=12, pady=(18, 8))
        ctk.CTkLabel(badge_box, text="حالة الحسابات:", font=("Segoe UI", 11, "bold")).pack(anchor="e", padx=10, pady=(8, 4))
        self.lbl_side_li = ctk.CTkLabel(badge_box, text="○ LinkedIn", font=("Segoe UI", 11), anchor="e")
        self.lbl_side_li.pack(fill="x", padx=10, pady=1)
        self.lbl_side_wa = ctk.CTkLabel(badge_box, text="○ WhatsApp", font=("Segoe UI", 11), anchor="e")
        self.lbl_side_wa.pack(fill="x", padx=10, pady=1)
        self.lbl_side_gm = ctk.CTkLabel(badge_box, text="○ Gmail", font=("Segoe UI", 11), anchor="e")
        self.lbl_side_gm.pack(fill="x", padx=10, pady=(1, 8))

        # إيقاف فوري
        self.btn_global_stop = ctk.CTkButton(
            self.sidebar, text="⛔ إيقاف العملية الحالية", height=38,
            font=("Segoe UI", 12, "bold"), fg_color="#DC2626", hover_color="#991B1B",
            state="disabled", command=self.stop_current_task,
        )
        self.btn_global_stop.pack(fill="x", padx=12, pady=(8, 4))

        ctk.CTkLabel(self.sidebar, text="المظهر:", font=("Segoe UI", 11)).pack(side="bottom", pady=(0, 2))
        self.mode_switch = ctk.CTkSegmentedButton(self.sidebar, values=["Dark", "Light"], command=self._change_theme)
        self.mode_switch.set(self.settings.get("appearance", "Dark"))
        self.mode_switch.pack(side="bottom", padx=12, pady=(0, 12), fill="x")

        self.main_area = ctk.CTkFrame(self, fg_color="transparent")
        self.main_area.pack(side="left", fill="both", expand=True, padx=14, pady=12)

        self.pages = {k: ctk.CTkFrame(self.main_area, fg_color="transparent") for k in ("dashboard", "scraper", "leads", "templates", "campaigns")}
        self._build_dashboard_page(self.pages["dashboard"])
        self._build_scraper_page(self.pages["scraper"])
        self._build_leads_page(self.pages["leads"])
        self._build_templates_page(self.pages["templates"])
        self._build_campaigns_page(self.pages["campaigns"])

    def _show_page(self, key):
        for k, frame in self.pages.items():
            frame.pack_forget()
            self.nav_buttons[k].configure(
                fg_color="#2563EB" if k == key else "transparent",
                text_color="#FFFFFF" if k == key else ("#0F172A", "#F8FAFC"),
            )
        self.pages[key].pack(fill="both", expand=True)
        if key == "dashboard":
            self.refresh_dashboard()
        elif key == "leads":
            self.refresh_leads_table()
        elif key == "templates":
            self.refresh_templates_ui()
        elif key == "campaigns":
            self.refresh_campaign_templates()

    def _change_theme(self, mode):
        ctk.set_appearance_mode(mode)
        self.db.save_settings({"appearance": mode})
        self._configure_treeview_style()
        self.refresh_leads_table()

    def refresh_account_badges(self):
        li = session_is_ready(LINKEDIN_SESSION_DIR)
        wa = session_is_ready(WA_SESSION_DIR)
        s = self.db.get_settings()
        gm = bool(s.get("sender_email") and s.get("app_password"))

        self.lbl_side_li.configure(text="● LinkedIn: مربوط" if li else "○ LinkedIn: غير مربوط", text_color="#22C55E" if li else "#F87171")
        self.lbl_side_wa.configure(text="● WhatsApp: مربوط" if wa else "○ WhatsApp: غير مربوط", text_color="#22C55E" if wa else "#F87171")
        self.lbl_side_gm.configure(text="● Gmail: مربوط" if gm else "○ Gmail: غير مربوط", text_color="#22C55E" if gm else "#F87171")

    # ======================= صفحة 1: لوحة التحكم =======================

    def _build_dashboard_page(self, parent):
        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(hdr, text="لوحة التحكم وملخص مراحل المبيعات", font=("Segoe UI", 21, "bold")).pack(side="right")
        ctk.CTkButton(hdr, text="تحديث", width=100, command=self.refresh_dashboard).pack(side="left")

        self.kpi_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self.kpi_frame.pack(fill="x", pady=(0, 12))
        self.kpi_labels = {}

        cards = [
            ("total", "إجمالي العملاء", "#0EA5E9"),
            ("جديد", "جدد", STAGE_COLORS["جديد"]),
            ("مهتم", "مهتمين", STAGE_COLORS["مهتم"]),
            ("معاينة", "معاينات", STAGE_COLORS["معاينة"]),
            ("تم البيع", "تم البيع", STAGE_COLORS["تم البيع"]),
            ("rate", "نسبة التحويل", "#8B5CF6"),
        ]
        for idx, (k, title, color) in enumerate(cards):
            c = ctk.CTkFrame(self.kpi_frame, corner_radius=12, border_width=2, border_color=color)
            c.grid(row=0, column=idx, padx=5, pady=4, sticky="nsew")
            self.kpi_frame.columnconfigure(idx, weight=1)
            ctk.CTkLabel(c, text=title, font=("Segoe UI", 12, "bold"), text_color="#94A3B8").pack(pady=(10, 2))
            v = ctk.CTkLabel(c, text="0", font=("Segoe UI", 24, "bold"), text_color=color)
            v.pack(pady=(0, 10))
            self.kpi_labels[k] = v

        due_box = ctk.CTkFrame(parent, corner_radius=12)
        due_box.pack(fill="both", expand=True)
        ctk.CTkLabel(due_box, text="⏰ متابعات مستحقة اليوم أو متأخرة", font=("Segoe UI", 15, "bold")).pack(anchor="e", padx=14, pady=(12, 6))
        self.due_list_frame = ctk.CTkScrollableFrame(due_box)
        self.due_list_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    def refresh_dashboard(self):
        total, counts, due = self.db.get_pipeline_stats()
        self.kpi_labels["total"].configure(text=str(total))
        for st in ("جديد", "مهتم", "معاينة", "تم البيع"):
            self.kpi_labels[st].configure(text=str(counts.get(st, 0)))
        rate = (counts.get("تم البيع", 0) / total * 100) if total else 0.0
        self.kpi_labels["rate"].configure(text=f"{rate:.1f}%")

        for w in self.due_list_frame.winfo_children():
            w.destroy()
        if not due:
            ctk.CTkLabel(self.due_list_frame, text="لا توجد متابعات متأخرة اليوم.", text_color="#94A3B8").pack(pady=30)
            return
        for lead in due:
            r = ctk.CTkFrame(self.due_list_frame, corner_radius=8)
            r.pack(fill="x", pady=4, padx=4)
            ctk.CTkLabel(r, text=f"{lead['name']}  |  {lead['stage']}  |  {lead['next_followup']}  |  {lead.get('phone') or ''}", font=("Segoe UI", 12, "bold")).pack(side="right", padx=12, pady=8)
            ctk.CTkButton(r, text="تعديل", width=70, command=lambda l=lead: self._open_edit_lead(l)).pack(side="left", padx=6, pady=6)

    # ======================= صفحة 2: محرك سحب العملاء (Scraping) =======================

    def _build_scraper_page(self, parent):
        s = self.settings
        ctk.CTkLabel(parent, text="🔍 محرك سحب العملاء ومواصفات العميل المستهدف (ICP)", font=("Segoe UI", 20, "bold")).pack(anchor="e", pady=(0, 8))

        # 1. اختيار المصادر
        src_box = ctk.CTkFrame(parent, corner_radius=12)
        src_box.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(src_box, text="1. اختر مواقع السحب:", font=("Segoe UI", 13, "bold")).pack(side="right", padx=12, pady=10)

        self.var_src_li = tk.BooleanVar(value=True)
        self.var_src_gm = tk.BooleanVar(value=True)
        self.var_src_web = tk.BooleanVar(value=False)

        ctk.CTkCheckBox(src_box, text="LinkedIn (شركات + About)", variable=self.var_src_li).pack(side="right", padx=10, pady=10)
        ctk.CTkCheckBox(src_box, text="Google Maps (تليفونات + تقييمات)", variable=self.var_src_gm).pack(side="right", padx=10, pady=10)
        ctk.CTkCheckBox(src_box, text="محركات البحث (Web)", variable=self.var_src_web).pack(side="right", padx=10, pady=10)

        # 2. الكلمات واللوكيشن والكمية والفلاتر
        cfg_box = ctk.CTkFrame(parent, corner_radius=12)
        cfg_box.pack(fill="x", pady=(0, 8))

        self.e_keywords = ctk.CTkEntry(cfg_box, justify="right", placeholder_text="marketing agency, software house")
        self.e_keywords.insert(0, s.get("keywords", "digital marketing agency, software agency"))
        self.e_location = ctk.CTkEntry(cfg_box, justify="right", placeholder_text="Mansoura / Cairo / Riyadh")
        self.e_location.insert(0, s.get("location", "Mansoura"))
        self.e_target = ctk.CTkEntry(cfg_box, width=90)
        self.e_target.insert(0, s.get("target_count", "50"))

        self.cmb_contact_req = ctk.CTkComboBox(
            cfg_box, state="readonly", width=210,
            values=[
                "الكل (أي بيانات متاحة)",
                "رقم أو إيميل على الأقل",
                "أرقام واتساب فقط",
                "إيميلات فقط",
                "الاثنين معاً (واتساب + إيميل)",
            ],
        )
        self.cmb_contact_req.set(s.get("contact_req", "رقم أو إيميل على الأقل"))

        self.e_icp_ind = ctk.CTkEntry(cfg_box, justify="right", placeholder_text="اختياري: marketing, software...")
        self.e_icp_ind.insert(0, s.get("icp_industry", ""))
        self.cmb_icp_size = ctk.CTkComboBox(cfg_box, state="readonly", width=150, values=["الكل", "صغيرة (1-10)", "متوسطة (11-50)", "كبيرة (51-200)", "مؤسسات (+200)"])
        self.cmb_icp_size.set(s.get("icp_size", "الكل"))
        self.e_min_rating = ctk.CTkEntry(cfg_box, width=70)
        self.e_min_rating.insert(0, s.get("min_rating", "0"))

        self.var_must_web = tk.BooleanVar(value=False)
        self.var_enrich_web = tk.BooleanVar(value=True)

        # ترتيب الشبكة
        ctk.CTkLabel(cfg_box, text="الكلمات المفتاحية (افصل بفاصلة):").grid(row=0, column=3, sticky="e", padx=10, pady=6)
        self.e_keywords.grid(row=0, column=0, columnspan=3, sticky="ew", padx=10, pady=6)

        ctk.CTkLabel(cfg_box, text="اللوكيشن / المدينة:").grid(row=1, column=3, sticky="e", padx=10, pady=5)
        self.e_location.grid(row=1, column=2, sticky="ew", padx=10, pady=5)
        ctk.CTkLabel(cfg_box, text="كمية الداتا المطلوبة:").grid(row=1, column=1, sticky="e", padx=10, pady=5)
        self.e_target.grid(row=1, column=0, sticky="e", padx=10, pady=5)

        ctk.CTkLabel(cfg_box, text="بيانات التواصل المطلوبة:").grid(row=2, column=3, sticky="e", padx=10, pady=5)
        self.cmb_contact_req.grid(row=2, column=2, sticky="ew", padx=10, pady=5)
        ctk.CTkLabel(cfg_box, text="فلتر التخصص (اختياري):").grid(row=2, column=1, sticky="e", padx=10, pady=5)
        self.e_icp_ind.grid(row=2, column=0, sticky="ew", padx=10, pady=5)

        ctk.CTkLabel(cfg_box, text="حجم الشركة (LinkedIn):").grid(row=3, column=3, sticky="e", padx=10, pady=5)
        self.cmb_icp_size.grid(row=3, column=2, sticky="e", padx=10, pady=5)
        ctk.CTkLabel(cfg_box, text="أقل تقييم (Maps):").grid(row=3, column=1, sticky="e", padx=10, pady=5)
        self.e_min_rating.grid(row=3, column=0, sticky="e", padx=10, pady=5)

        chk_row = ctk.CTkFrame(cfg_box, fg_color="transparent")
        chk_row.grid(row=4, column=0, columnspan=4, sticky="ew", padx=10, pady=(4, 10))
        ctk.CTkCheckBox(chk_row, text="فحص موقع كل شركة لاستخراج الإيميل والواتساب تلقائياً", variable=self.var_enrich_web).pack(side="right", padx=8)
        ctk.CTkCheckBox(chk_row, text="يشترط وجود موقع إلكتروني للعميل", variable=self.var_must_web).pack(side="right", padx=12)

        self.btn_start_scrape = ctk.CTkButton(
            chk_row, text="🚀 بدء سحب العملاء وحفظهم في القاعدة", height=38,
            font=("Segoe UI", 13, "bold"), fg_color="#0A66C2", hover_color="#004182",
            command=self._start_scraping,
        )
        self.btn_start_scrape.pack(side="left", padx=6)

        cfg_box.columnconfigure(0, weight=1)
        cfg_box.columnconfigure(2, weight=1)

        # 3. شاشة المتابعة المباشرة للسحب
        log_card = ctk.CTkFrame(parent, corner_radius=12)
        log_card.pack(fill="both", expand=True)
        top_l = ctk.CTkFrame(log_card, fg_color="transparent")
        top_l.pack(fill="x", padx=12, pady=(8, 4))
        ctk.CTkLabel(top_l, text="سجل عمليات السحب والحملات المباشر (Live Log)", font=("Segoe UI", 13, "bold")).pack(side="right")
        ctk.CTkButton(top_l, text="مسح السجل", width=90, fg_color="#475569", command=self._clear_logs).pack(side="left")

        self.scrape_log = ctk.CTkTextbox(log_card, font=("Consolas", 11))
        self.scrape_log.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.scrape_log.configure(state="disabled")

    # ======================= صفحة 3: قاعدة العملاء ومراحل البيع =======================

    def _build_leads_page(self, parent):
        top_bar = ctk.CTkFrame(parent, corner_radius=10)
        top_bar.pack(fill="x", pady=(0, 6))

        self.var_search = tk.StringVar()
        self.var_search.trace_add("write", lambda *_: self.refresh_leads_table())
        ctk.CTkEntry(top_bar, textvariable=self.var_search, placeholder_text="🔍 بحث في العملاء...", width=210, justify="right").pack(side="right", padx=6, pady=8)

        self.cmb_filter_stage = ctk.CTkComboBox(top_bar, values=["الكل"] + STAGES, state="readonly", width=115, command=lambda _: self.refresh_leads_table())
        self.cmb_filter_stage.set("الكل")
        self.cmb_filter_stage.pack(side="right", padx=4, pady=8)

        self.cmb_quick_filter = ctk.CTkComboBox(
            top_bar, state="readonly", width=165,
            values=["الكل", "بها إيميل ولم تُرسل", "بها واتساب ولم تُرسل", "عملاء LinkedIn"],
            command=lambda _: self.refresh_leads_table(),
        )
        self.cmb_quick_filter.set("الكل")
        self.cmb_quick_filter.pack(side="right", padx=4, pady=8)

        ctk.CTkButton(top_bar, text="+ عميل", width=80, fg_color="#10B981", command=self._open_add_lead).pack(side="left", padx=4, pady=8)
        ctk.CTkButton(top_bar, text="📥 استيراد Excel", width=105, fg_color="#475569", command=self._import_excel).pack(side="left", padx=4, pady=8)
        ctk.CTkButton(top_bar, text="📤 تصدير Excel", width=105, fg_color="#0284C7", command=self._export_excel).pack(side="left", padx=4, pady=8)

        pipe_bar = ctk.CTkFrame(parent, fg_color="transparent")
        pipe_bar.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(pipe_bar, text="نقل المحدد لمرحلة:", font=("Segoe UI", 12, "bold")).pack(side="right", padx=4)
        self.cmb_bulk_stage = ctk.CTkComboBox(pipe_bar, values=STAGES, state="readonly", width=115)
        self.cmb_bulk_stage.set("مهتم")
        self.cmb_bulk_stage.pack(side="right", padx=4)
        ctk.CTkButton(pipe_bar, text="تطبيق", width=70, command=self._apply_bulk_stage).pack(side="right", padx=4)

        ctk.CTkButton(pipe_bar, text="تحديد الظاهر", width=90, fg_color="#334155", command=self._select_all_visible).pack(side="right", padx=(12, 3))
        ctk.CTkButton(pipe_bar, text="إلغاء التحديد", width=90, fg_color="#334155", command=self._clear_selection).pack(side="right", padx=3)
        self.lbl_selected_count = ctk.CTkLabel(pipe_bar, text="المحدد: 0", font=("Segoe UI", 12, "bold"), text_color="#38BDF8")
        self.lbl_selected_count.pack(side="left", padx=8)

        table_box = ctk.CTkFrame(parent, corner_radius=10)
        table_box.pack(fill="both", expand=True, pady=(0, 8))

        cols = ("sel", "id", "src", "name", "phone", "email", "cat", "loc", "stage", "em_st", "wa_st", "li_st")
        self.tree = ttk.Treeview(table_box, columns=cols, show="headings", style="CRM.Treeview", selectmode="none")
        specs = [
            ("sel", "تحديد", 46, "center"), ("id", "#", 36, "center"), ("src", "المصدر", 88, "center"),
            ("name", "الاسم / الشركة", 160, "e"), ("phone", "واتساب / موبايل", 115, "center"),
            ("email", "الإيميل", 150, "w"), ("cat", "التخصص", 115, "center"),
            ("loc", "المنطقة", 85, "center"), ("stage", "المرحلة", 85, "center"),
            ("em_st", "حالة إيميل", 100, "center"), ("wa_st", "حالة واتساب", 100, "center"),
            ("li_st", "حالة LinkedIn", 100, "center"),
        ]
        for cid, h, w, anc in specs:
            self.tree.heading(cid, text=h)
            self.tree.column(cid, width=w, minwidth=36, anchor=anc)

        vsb = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(6, 0), pady=6)
        vsb.grid(row=0, column=1, sticky="ns", pady=6)
        hsb.grid(row=1, column=0, sticky="ew", padx=6)
        table_box.rowconfigure(0, weight=1)
        table_box.columnconfigure(0, weight=1)

        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        act_bar = ctk.CTkFrame(parent, fg_color="transparent")
        act_bar.pack(fill="x")

        self.btn_launch_sel_camp = ctk.CTkButton(
            act_bar, text="🚀 إطلاق الحملة للمحددين (Gmail / WhatsApp / LinkedIn)",
            fg_color="#16A34A", hover_color="#15803D", font=("Segoe UI", 12, "bold"),
            command=self._launch_campaign_for_selected,
        )
        self.btn_launch_sel_camp.pack(side="right", padx=4)

        ctk.CTkButton(act_bar, text="تعديل", width=80, command=self._edit_first_selected).pack(side="left", padx=4)
        ctk.CTkButton(act_bar, text="حذف المحدد", width=95, fg_color="#DC2626", hover_color="#B91C1C", command=self._delete_selected).pack(side="left", padx=4)

    def refresh_leads_table(self):
        leads = self.db.get_leads(
            search=self.var_search.get(),
            stage_filter=self.cmb_filter_stage.get(),
            quick_filter=self.cmb_quick_filter.get(),
        )
        self.tree.delete(*self.tree.get_children())
        is_dark = ctk.get_appearance_mode() == "Dark"
        self.tree.tag_configure("checked", background="#1D4ED8" if is_dark else "#DBEAFE", foreground="#FFFFFF" if is_dark else "#0F172A")

        for lead in leads:
            lid = lead["id"]
            is_sel = lid in self.selected_lead_ids
            self.tree.insert(
                "", "end", iid=str(lid),
                values=(
                    "☑" if is_sel else "☐", lid, lead.get("source") or "—", lead["name"],
                    lead.get("phone") or "—", lead.get("email") or "—", lead.get("category") or "—",
                    lead.get("location") or "—", lead.get("stage") or "جديد",
                    lead.get("email_status") or "—", lead.get("wa_status") or "—", lead.get("li_status") or "—",
                ),
                tags=("checked",) if is_sel else (),
            )
        self.lbl_selected_count.configure(text=f"المحدد: {len(self.selected_lead_ids)}")

    def _on_tree_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        lid = int(iid)
        vals = list(self.tree.item(iid, "values"))
        if lid in self.selected_lead_ids:
            self.selected_lead_ids.remove(lid)
            vals[0] = "☐"
            self.tree.item(iid, values=vals, tags=())
        else:
            self.selected_lead_ids.add(lid)
            vals[0] = "☑"
            self.tree.item(iid, values=vals, tags=("checked",))
        self.lbl_selected_count.configure(text=f"المحدد: {len(self.selected_lead_ids)}")
        return "break"

    def _on_tree_double_click(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            lead = self.db.get_lead_by_id(int(iid))
            if lead:
                self._open_edit_lead(lead)

    def _select_all_visible(self):
        for iid in self.tree.get_children():
            self.selected_lead_ids.add(int(iid))
        self.refresh_leads_table()

    def _clear_selection(self):
        self.selected_lead_ids.clear()
        self.refresh_leads_table()

    def _open_add_lead(self):
        LeadDialog(self, self._save_lead_callback, lead=None)

    def _open_edit_lead(self, lead):
        LeadDialog(self, self._save_lead_callback, lead=lead)

    def _edit_first_selected(self):
        if not self.selected_lead_ids:
            messagebox.showwarning("تحديد", "حدد عميلاً للتعديل (أو دبل كليك عليه).")
            return
        lead = self.db.get_lead_by_id(next(iter(self.selected_lead_ids)))
        if lead:
            self._open_edit_lead(lead)

    def _save_lead_callback(self, data, lead_id):
        if lead_id:
            self.db.update_lead(lead_id, data)
        else:
            self.db.add_lead(data)
        self.refresh_leads_table()
        self.refresh_dashboard()

    def _apply_bulk_stage(self):
        if not self.selected_lead_ids:
            return
        st = self.cmb_bulk_stage.get()
        for lid in self.selected_lead_ids:
            self.db.update_stage(lid, st)
        self.refresh_leads_table()
        self.refresh_dashboard()

    def _delete_selected(self):
        if not self.selected_lead_ids:
            return
        if messagebox.askyesno("تأكيد الحذف", f"حذف {len(self.selected_lead_ids)} عميل نهائياً؟"):
            self.db.delete_leads(self.selected_lead_ids)
            self.selected_lead_ids.clear()
            self.refresh_leads_table()
            self.refresh_dashboard()

    def _import_excel(self):
        path = filedialog.askopenfilename(title="استيراد Excel", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            n = import_leads_from_excel(self.db, path)
            self.refresh_leads_table()
            messagebox.showinfo("تم الاستيراد", f"تمت إضافة {n} عميل جديد.")
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def _export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile=f"leads_{datetime.now().strftime('%Y%m%d')}.xlsx")
        if not path:
            return
        leads = self.db.get_leads(self.var_search.get(), self.cmb_filter_stage.get(), self.cmb_quick_filter.get())
        if not leads:
            messagebox.showwarning("فارغ", "لا يوجد عملاء لتصديرهم.")
            return
        tpls = self.db.get_templates()
        tpl = tpls[0] if tpls else {"subject": "تعاون مع {name}", "body": "السلام عليكم {name}"}
        try:
            export_leads_to_excel(leads, path, tpl["subject"], tpl["body"], self.db.get_settings().get("sender_name", ""))
            if messagebox.askyesno("تم التصدير", f"تم تصدير {len(leads)} عميل.\nفتح الملف الآن؟"):
                if platform.system() == "Windows":
                    os.startfile(path)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", path])
                else:
                    subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    # ======================= صفحة 4: قوالب الرسائل =======================

    def _build_templates_page(self, parent):
        ctk.CTkLabel(parent, text="📝 إدارة قوالب الرسائل المخصصة", font=("Segoe UI", 20, "bold")).pack(anchor="e", pady=(0, 8))
        grid = ctk.CTkFrame(parent, fg_color="transparent")
        grid.pack(fill="both", expand=True)
        grid.columnconfigure(0, weight=2)
        grid.columnconfigure(1, weight=1)
        grid.rowconfigure(0, weight=1)

        ed = ctk.CTkFrame(grid, corner_radius=12)
        ed.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        ctk.CTkLabel(ed, text="اسم القالب:", font=("Segoe UI", 12, "bold")).pack(anchor="e", padx=12, pady=(10, 2))
        self.e_tpl_title = ctk.CTkEntry(ed, justify="right")
        self.e_tpl_title.pack(fill="x", padx=12, pady=(0, 6))

        ctk.CTkLabel(ed, text="عنوان الإيميل (Subject):", font=("Segoe UI", 12, "bold")).pack(anchor="e", padx=12, pady=(4, 2))
        self.e_tpl_subject = ctk.CTkEntry(ed, justify="right")
        self.e_tpl_subject.pack(fill="x", padx=12, pady=(0, 6))

        ctk.CTkLabel(ed, text="نص الرسالة — المتغيرات: {name} {category} {location} {budget} {sender_name}", text_color="#38BDF8").pack(anchor="e", padx=12, pady=(4, 2))
        self.txt_tpl_body = ctk.CTkTextbox(ed, font=("Segoe UI", 12))
        self.txt_tpl_body.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        btns = ctk.CTkFrame(ed, fg_color="transparent")
        btns.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(btns, text="💾 حفظ القالب", fg_color="#10B981", command=self._save_template).pack(side="right", padx=4)
        ctk.CTkButton(btns, text="👁️ معاينة", fg_color="#0284C7", command=self._preview_template).pack(side="right", padx=4)
        ctk.CTkButton(btns, text="جديد", fg_color="#475569", width=70, command=self._clear_tpl).pack(side="left", padx=4)
        ctk.CTkButton(btns, text="حذف", fg_color="#DC2626", width=70, command=self._del_tpl).pack(side="left", padx=4)

        sl = ctk.CTkFrame(grid, corner_radius=12)
        sl.grid(row=0, column=1, sticky="nsew")
        ctk.CTkLabel(sl, text="القوالب المحفوظة", font=("Segoe UI", 14, "bold")).pack(pady=(12, 6))
        self.tpl_scroll = ctk.CTkScrollableFrame(sl)
        self.tpl_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 10))

    def refresh_templates_ui(self):
        for w in self.tpl_scroll.winfo_children():
            w.destroy()
        tpls = self.db.get_templates()
        for t in tpls:
            ctk.CTkButton(
                self.tpl_scroll, text=t["title"], anchor="e",
                fg_color=("#E2E8F0", "#334155"), text_color=("#0F172A", "#F8FAFC"),
                command=lambda x=t: self._load_tpl(x),
            ).pack(fill="x", pady=3, padx=4)
        if tpls and not self.e_tpl_title.get().strip():
            self._load_tpl(tpls[0])

    def _load_tpl(self, t):
        self.e_tpl_title.delete(0, "end")
        self.e_tpl_title.insert(0, t["title"])
        self.e_tpl_subject.delete(0, "end")
        self.e_tpl_subject.insert(0, t["subject"])
        self.txt_tpl_body.delete("1.0", "end")
        self.txt_tpl_body.insert("1.0", t["body"])

    def _clear_tpl(self):
        self.e_tpl_title.delete(0, "end")
        self.e_tpl_subject.delete(0, "end")
        self.txt_tpl_body.delete("1.0", "end")

    def _save_template(self):
        title = self.e_tpl_title.get().strip()
        body = self.txt_tpl_body.get("1.0", "end-1c").strip()
        if not title or not body:
            messagebox.showwarning("ناقص بيانات", "اكتب اسم القالب ونص الرسالة.")
            return
        self.db.save_template(title, self.e_tpl_subject.get().strip(), body, "الكل")
        self.refresh_templates_ui()
        messagebox.showinfo("تم", f"تم حفظ القالب '{title}'.")

    def _del_tpl(self):
        title = self.e_tpl_title.get().strip()
        if title and len(self.db.get_templates()) > 1:
            self.db.delete_template(title)
            self._clear_tpl()
            self.refresh_templates_ui()

    def _preview_template(self):
        sample = {"name": "وكالة الإبداع", "category": "تسويق رقمي", "location": "المنصورة", "budget": "50,000 ج.م"}
        sn = self.db.get_settings().get("sender_name", "فريق المبيعات")
        messagebox.showinfo("معاينة", render_msg(self.txt_tpl_body.get("1.0", "end-1c"), sample, sn))

    # ======================= صفحة 5: الحملات وربط الحسابات =======================

    def _build_campaigns_page(self, parent):
        s = self.settings

        # 1. ربط حسابات LinkedIn + WhatsApp + Gmail
        acc_box = ctk.CTkFrame(parent, corner_radius=12)
        acc_box.pack(fill="x", pady=(0, 8))

        sess_row = ctk.CTkFrame(acc_box, fg_color="transparent")
        sess_row.pack(fill="x", padx=12, pady=(10, 6))
        ctk.CTkLabel(sess_row, text="ربط الحسابات:", font=("Segoe UI", 13, "bold")).pack(side="right", padx=6)

        self.btn_conn_li = ctk.CTkButton(sess_row, text="🔗 ربط LinkedIn", width=130, fg_color="#0A66C2", command=lambda: self._start_task(self._worker_connect_browser, "ربط LinkedIn", LINKEDIN_SESSION_DIR, "https://www.linkedin.com/login", "LinkedIn"))
        self.btn_conn_li.pack(side="right", padx=4)
        ctk.CTkButton(sess_row, text="فصل", width=50, fg_color="#475569", command=lambda: self._disconnect_session(LINKEDIN_SESSION_DIR, "LinkedIn")).pack(side="right", padx=(0, 12))

        self.btn_conn_wa = ctk.CTkButton(sess_row, text="🔗 ربط WhatsApp Web", width=155, fg_color="#25D366", hover_color="#16A34A", command=lambda: self._start_task(self._worker_connect_browser, "ربط WhatsApp", WA_SESSION_DIR, "https://web.whatsapp.com", "WhatsApp Web"))
        self.btn_conn_wa.pack(side="right", padx=4)
        ctk.CTkButton(sess_row, text="فصل", width=50, fg_color="#475569", command=lambda: self._disconnect_session(WA_SESSION_DIR, "WhatsApp")).pack(side="right", padx=4)

        gm_grid = ctk.CTkFrame(acc_box, fg_color="transparent")
        gm_grid.pack(fill="x", padx=12, pady=(4, 10))
        self.e_sender_email = ctk.CTkEntry(gm_grid, placeholder_text="your_email@gmail.com")
        self.e_sender_email.insert(0, s.get("sender_email", ""))
        self.e_app_pass = ctk.CTkEntry(gm_grid, placeholder_text="Gmail App Password (16 حرف)", show="*")
        self.e_app_pass.insert(0, s.get("app_password", ""))
        self.e_sender_name = ctk.CTkEntry(gm_grid, placeholder_text="اسمك (المرسل)", justify="right")
        self.e_sender_name.insert(0, s.get("sender_name", ""))
        self.e_attach = ctk.CTkEntry(gm_grid, placeholder_text="ملف مرفق اختياري")
        self.e_attach.insert(0, s.get("attachment_path", ""))

        self.e_sender_email.grid(row=0, column=3, sticky="ew", padx=4, pady=3)
        self.e_app_pass.grid(row=0, column=2, sticky="ew", padx=4, pady=3)
        self.e_sender_name.grid(row=0, column=1, sticky="ew", padx=4, pady=3)
        att_box = ctk.CTkFrame(gm_grid, fg_color="transparent")
        att_box.grid(row=0, column=0, sticky="ew", padx=4, pady=3)
        self.e_attach.pack(in_=att_box, side="left", fill="x", expand=True)
        ctk.CTkButton(att_box, text="مرفق", width=55, command=self._browse_attachment).pack(side="left", padx=(3, 0))

        gm_btns = ctk.CTkFrame(gm_grid, fg_color="transparent")
        gm_btns.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(4, 0))
        self.e_max_emails = ctk.CTkEntry(gm_btns, width=55)
        self.e_max_emails.insert(0, s.get("max_emails", "35"))
        self.e_max_wa = ctk.CTkEntry(gm_btns, width=55)
        self.e_max_wa.insert(0, s.get("max_wa", "20"))
        self.e_max_li = ctk.CTkEntry(gm_btns, width=55)
        self.e_max_li.insert(0, s.get("max_li", "15"))

        ctk.CTkLabel(gm_btns, text="حد Gmail:").pack(side="right", padx=3)
        self.e_max_emails.pack(side="right", padx=(0, 10))
        ctk.CTkLabel(gm_btns, text="حد WhatsApp:").pack(side="right", padx=3)
        self.e_max_wa.pack(side="right", padx=(0, 10))
        ctk.CTkLabel(gm_btns, text="حد LinkedIn:").pack(side="right", padx=3)
        self.e_max_li.pack(side="right", padx=(0, 10))

        ctk.CTkButton(gm_btns, text="💾 حفظ الإعدادات", width=115, fg_color="#10B981", command=self._save_campaign_settings).pack(side="left", padx=4)
        ctk.CTkButton(gm_btns, text="✓ اختبار Gmail", width=110, fg_color="#EA4335", command=self._test_gmail).pack(side="left", padx=4)

        for c in range(4):
            gm_grid.columnconfigure(c, weight=1)

        # 2. إطلاق الحملة متعددة القنوات
        camp_box = ctk.CTkFrame(parent, corner_radius=12)
        camp_box.pack(fill="both", expand=True)

        row_ch = ctk.CTkFrame(camp_box, fg_color="transparent")
        row_ch.pack(fill="x", padx=12, pady=(10, 6))

        self.var_ch_gmail = tk.BooleanVar(value=True)
        self.var_ch_wa = tk.BooleanVar(value=True)
        self.var_ch_li = tk.BooleanVar(value=False)
        self.var_dry_run = tk.BooleanVar(value=True)

        ctk.CTkLabel(row_ch, text="قنوات الحملة:", font=("Segoe UI", 12, "bold")).pack(side="right", padx=4)
        ctk.CTkCheckBox(row_ch, text="Gmail", variable=self.var_ch_gmail).pack(side="right", padx=8)
        ctk.CTkCheckBox(row_ch, text="WhatsApp Web", variable=self.var_ch_wa).pack(side="right", padx=8)
        ctk.CTkCheckBox(row_ch, text="LinkedIn", variable=self.var_ch_li).pack(side="right", padx=8)
        ctk.CTkCheckBox(row_ch, text="وضع التجربة (DRY RUN)", variable=self.var_dry_run).pack(side="left", padx=8)

        row_go = ctk.CTkFrame(camp_box, fg_color="transparent")
        row_go.pack(fill="x", padx=12, pady=(0, 8))

        ctk.CTkLabel(row_go, text="القالب:").pack(side="right", padx=4)
        self.cmb_camp_tpl = ctk.CTkComboBox(row_go, values=[], state="readonly", width=180)
        self.cmb_camp_tpl.pack(side="right", padx=4)

        ctk.CTkLabel(row_go, text="استهداف مرحلة:").pack(side="right", padx=(10, 4))
        self.cmb_camp_stage = ctk.CTkComboBox(row_go, values=["الكل"] + STAGES, state="readonly", width=120)
        self.cmb_camp_stage.set("جديد")
        self.cmb_camp_stage.pack(side="right", padx=4)

        self.btn_run_camp_all = ctk.CTkButton(
            row_go, text="🚀 إطلاق الحملة للمرحلة المختارة", fg_color="#16A34A", hover_color="#15803D",
            font=("Segoe UI", 12, "bold"), command=lambda: self._trigger_campaign(None),
        )
        self.btn_run_camp_all.pack(side="left", padx=4)

        self.camp_log = ctk.CTkTextbox(camp_box, font=("Consolas", 11))
        self.camp_log.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.camp_log.configure(state="disabled")

    def refresh_campaign_templates(self):
        names = [t["title"] for t in self.db.get_templates()]
        if names:
            cur = self.cmb_camp_tpl.get()
            self.cmb_camp_tpl.configure(values=names)
            if cur not in names:
                self.cmb_camp_tpl.set(names[0])

    def _browse_attachment(self):
        p = filedialog.askopenfilename(title="اختر ملفاً")
        if p:
            self.e_attach.delete(0, "end")
            self.e_attach.insert(0, p)

    def _save_campaign_settings(self, popup=True):
        data = {
            "sender_email": self.e_sender_email.get().strip(),
            "app_password": self.e_app_pass.get().strip(),
            "sender_name": self.e_sender_name.get().strip(),
            "attachment_path": self.e_attach.get().strip(),
            "max_emails": self.e_max_emails.get().strip() or "35",
            "max_wa": self.e_max_wa.get().strip() or "20",
            "max_li": self.e_max_li.get().strip() or "15",
        }
        self.db.save_settings(data)
        self.settings.update(data)
        self.refresh_account_badges()
        if popup:
            messagebox.showinfo("تم الحفظ", "تم حفظ إعدادات الحسابات والحملة بنجاح.")

    def _disconnect_session(self, folder, name):
        if self.is_running:
            messagebox.showwarning("مشغول", "أوقف العملية الحالية أولاً.")
            return
        if os.path.isdir(folder) and messagebox.askyesno("فصل", f"مسح جلسة {name} المحفوظة؟"):
            shutil.rmtree(folder, ignore_errors=True)
            self.log(f"تم فصل حساب {name}.")
        self.refresh_account_badges()

    def _test_gmail(self):
        self._save_campaign_settings(popup=False)
        em = self.settings.get("sender_email")
        pw = self.settings.get("app_password")
        if not em or not pw:
            messagebox.showwarning("ناقص", "اكتب إيميل Gmail والـ App Password الأول.")
            return

        def _w():
            try:
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=20) as smtp:
                    smtp.login(em, pw.replace(" ", ""))
                self.log("✓ نجح الاتصال بحساب Gmail.")
                self.after(0, lambda: messagebox.showinfo("نجاح", "تم الاتصال بحساب Gmail بنجاح!"))
            except Exception as e:
                self.log(f"✗ فشل اتصال Gmail: {e}")
                self.after(0, lambda err=str(e): messagebox.showerror("فشل الاتصال", err))

        threading.Thread(target=_w, daemon=True).start()

    # ======================= إدارة الـ Threads والـ Logs =======================

    def log(self, text):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{ts}] {text}")

    def _poll_logs(self):
        while not self.log_queue.empty():
            m = self.log_queue.get_nowait()
            for box in (self.scrape_log, self.camp_log):
                box.configure(state="normal")
                box.insert("end", m + "\n")
                box.see("end")
                box.configure(state="disabled")
        self.after(150, self._poll_logs)

    def _clear_logs(self):
        for box in (self.scrape_log, self.camp_log):
            box.configure(state="normal")
            box.delete("1.0", "end")
            box.configure(state="disabled")

    def ask_main_thread(self, title, msg):
        res = {"ok": False}
        ev = threading.Event()
        def _pop():
            res["ok"] = messagebox.askokcancel(title, msg, parent=self)
            ev.set()
        self.after(0, _pop)
        ev.wait()
        return res["ok"]

    def _set_running(self, running):
        self.is_running = running
        st = "disabled" if running else "normal"
        for btn in (self.btn_start_scrape, self.btn_launch_sel_camp, self.btn_run_camp_all, self.btn_conn_li, self.btn_conn_wa):
            btn.configure(state=st)
        self.btn_global_stop.configure(state="normal" if running else "disabled")

    def _start_task(self, fn, title, *args):
        if self.is_running:
            messagebox.showwarning("مشغول", "هناك عملية شغالة حالياً، انتظر انتهاءها أو أوقفها.")
            return
        self.stop_event.clear()
        self._set_running(True)
        self.log("=" * 52)
        self.log(f"بدء العملية: {title}")

        def _wrap():
            try:
                fn(*args)
            except Exception as e:
                self.log(f"✗ خطأ: {e}")
            finally:
                self.after(0, lambda: self._set_running(False))
                self.after(0, self.refresh_account_badges)
                self.after(0, self.refresh_leads_table)
                self.after(0, self.refresh_dashboard)

        threading.Thread(target=_wrap, daemon=True).start()

    def stop_current_task(self):
        self.stop_event.set()
        self.log("⛔ تم طلب الإيقاف الفوري... جاري الإنهاء بأمان.")

    def _on_close(self):
        if self.is_running:
            if not messagebox.askyesno("تأكيد الخروج", "فيه عملية شغالة حالياً. هل تريد الإيقاف والخروج؟"):
                return
            self.stop_event.set()
        self.destroy()

    def _worker_connect_browser(self, session_dir, url, label):
        self.log(f"فتح متصفح {label} لحفظ الجلسة...")
        with sync_playwright() as p:
            ctx = launch_persistent_browser(p, session_dir)
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto(url, wait_until="domcontentloaded")
            self.ask_main_thread(f"ربط {label}", f"سجل دخولك في نافذة {label} ثم اضغط OK لحفظ الجلسة.")
            ctx.close()
        self.log(f"✓ تم حفظ جلسة {label} بنجاح.")

    # ======================= محرك السحب الفعلي (LinkedIn + Google Maps + Web) =======================

    def _start_scraping(self):
        try:
            target = int(self.e_target.get().strip())
        except ValueError:
            messagebox.showerror("خطأ", "كمية الداتا لازم تكون رقم صحيح.")
            return

        cfg = {
            "src_li": self.var_src_li.get(),
            "src_gm": self.var_src_gm.get(),
            "src_web": self.var_src_web.get(),
            "keywords": self.e_keywords.get().strip(),
            "location": self.e_location.get().strip(),
            "target": target,
            "contact_req": self.cmb_contact_req.get(),
            "icp_industry": self.e_icp_ind.get().strip(),
            "icp_size": self.cmb_icp_size.get(),
            "min_rating": self.e_min_rating.get().strip() or "0",
            "must_have_web": self.var_must_web.get(),
            "enrich_web": self.var_enrich_web.get(),
        }
        self.db.save_settings({
            "keywords": cfg["keywords"], "location": cfg["location"],
            "target_count": str(target), "contact_req": cfg["contact_req"],
            "icp_industry": cfg["icp_industry"], "icp_size": cfg["icp_size"],
            "min_rating": cfg["min_rating"],
        })
        self._start_task(self._worker_scrape_all, "سحب العملاء وحفظهم في SQLite", cfg)

    def _worker_scrape_all(self, cfg):
        keywords = [k.strip() for k in re.split(r"[,،\n]+", cfg["keywords"]) if k.strip()]
        if not keywords:
            self.log("خطأ: اكتب كلمة مفتاحية واحدة على الأقل.")
            return
        if not (cfg["src_li"] or cfg["src_gm"] or cfg["src_web"]):
            self.log("خطأ: اختار موقع واحد على الأقل للسحب.")
            return

        target = cfg["target"]
        loc = cfg["location"]
        added_count = [0]

        def accept_lead(lead):
            if self.db.lead_exists(lead["name"], lead.get("profile_url", "")):
                self.log(f"  ↳ مكرر بالفعل: {lead['name']}")
                return False
            ok, reason = matches_icp_and_contact(lead, cfg)
            if not ok:
                self.log(f"  ↳ استبعاد ({lead['name']}): {reason}")
                return False
            self.db.add_lead(lead)
            added_count[0] += 1
            self.log(f"  ✓ [{added_count[0]}/{target}] حُفظ: {lead['name']} | WA: {lead.get('phone') or '-'} | Mail: {lead.get('email') or '-'}")
            return True

        if cfg["src_gm"] and added_count[0] < target and not self.stop_event.is_set():
            self._scrape_gmaps(cfg, keywords, loc, target, added_count, accept_lead)

        if cfg["src_li"] and added_count[0] < target and not self.stop_event.is_set():
            self._scrape_linkedin(cfg, keywords, loc, target, added_count, accept_lead)

        if cfg["src_web"] and added_count[0] < target and not self.stop_event.is_set():
            self._scrape_web(cfg, keywords, loc, target, added_count, accept_lead)

        self.log(f"انتهى السحب! إجمالي العملاء الجدد المحفوظين في قاعدة البيانات: {added_count[0]}")

    def _scrape_gmaps(self, cfg, keywords, loc, target, added_count, accept_fn):
        self.log("--- بدء السحب من Google Maps ---")
        with sync_playwright() as p:
            browser = launch_temp_browser(p)
            ctx = browser.new_context(viewport={"width": 1280, "height": 800}, locale="en-US")
            page = ctx.new_page()

            for kw in keywords:
                if added_count[0] >= target or self.stop_event.is_set():
                    break
                query = f"{kw} in {loc}" if loc else kw
                self.log(f"[Google Maps] بحث: {query}")
                page.goto(f"https://www.google.com/maps/search/{quote(query)}", wait_until="domcontentloaded")
                if self.stop_event.wait(4):
                    break

                place_urls, seen_u = [], set()
                feed = page.locator("div[role='feed']")
                for _ in range(12):
                    if self.stop_event.is_set() or len(place_urls) >= (target - added_count[0]) * 2:
                        break
                    links = page.locator("a.hfpxzc, a[href*='/maps/place/']")
                    for i in range(links.count()):
                        href = links.nth(i).get_attribute("href")
                        if href and "/maps/place/" in href and href not in seen_u:
                            seen_u.add(href)
                            place_urls.append(href)
                    if feed.count() > 0:
                        feed.first.evaluate("el => el.scrollBy(0, 2200)")
                    else:
                        page.mouse.wheel(0, 2000)
                    if self.stop_event.wait(2):
                        break

                self.log(f"[Google Maps] فحص {len(place_urls)} مكان...")
                for p_url in place_urls:
                    if added_count[0] >= target or self.stop_event.is_set():
                        break
                    try:
                        page.goto(p_url, wait_until="domcontentloaded", timeout=15000)
                        if self.stop_event.wait(random.uniform(2, 3.2)):
                            break
                        name = page.locator("h1").first.inner_text().strip() if page.locator("h1").count() else "Business"
                        cat = page.locator("button[jsaction*='category']").first.inner_text().strip() if page.locator("button[jsaction*='category']").count() else kw
                        rating = page.locator("div.F7nice span[aria-hidden='true']").first.inner_text().strip() if page.locator("div.F7nice span[aria-hidden='true']").count() else ""
                        phone = ""
                        if page.locator("button[data-item-id^='phone:']").count():
                            raw_ph = page.locator("button[data-item-id^='phone:']").first.get_attribute("aria-label") or ""
                            phone = extract_mobile_number(raw_ph) or re.sub(r"[^\d+]", "", raw_ph)
                        website = page.locator("a[data-item-id='authority']").first.get_attribute("href") if page.locator("a[data-item-id='authority']").count() else ""
                        email = ""
                        if cfg["enrich_web"] and website:
                            wm, wwa = enrich_from_website(website)
                            if wm:
                                email = wm
                            if not extract_mobile_number(phone) and wwa:
                                phone = wwa

                        accept_fn({
                            "name": name, "source": "Google Maps", "category": cat,
                            "size_rating": f"★ {rating}" if rating else "",
                            "location": loc, "website": website or "", "phone": phone,
                            "email": email, "profile_url": p_url,
                        })
                    except Exception:
                        continue
            browser.close()

    def _scrape_linkedin(self, cfg, keywords, loc, target, added_count, accept_fn):
        self.log("--- بدء السحب من LinkedIn ---")
        with sync_playwright() as p:
            ctx = launch_persistent_browser(p, LINKEDIN_SESSION_DIR)
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto("https://www.linkedin.com/feed/", wait_until="domcontentloaded")

            if not self.ask_main_thread("تأكيد دخول LinkedIn", "تأكد إن حساب LinkedIn مفتوح في المتصفح ثم اضغط OK للبدء."):
                ctx.close()
                return

            seen_urls = set()
            for kw in keywords:
                if added_count[0] >= target or self.stop_event.is_set():
                    break
                search_q = f"{kw} {loc}".strip()
                base_url = f"https://www.linkedin.com/search/results/companies/?keywords={quote(search_q)}"
                company_links, page_num = [], 1

                while len(company_links) < (target - added_count[0]) * 2 and page_num <= 10 and not self.stop_event.is_set():
                    self.log(f"[LinkedIn] بحث '{search_q}' - صفحة {page_num}")
                    page.goto(f"{base_url}&page={page_num}", wait_until="domcontentloaded")
                    if self.stop_event.wait(random.uniform(3.5, 5.5)):
                        break
                    page.mouse.wheel(0, 2000)
                    if self.stop_event.wait(1.5):
                        break
                    cards = page.locator("li.reusable-search__result-container, div[data-chameleon-result-urn]")
                    if cards.count() == 0:
                        break
                    for i in range(cards.count()):
                        try:
                            card = cards.nth(i)
                            href = card.locator("a[href*='/company/']").first.get_attribute("href")
                            if not href:
                                continue
                            clean = href.split("?")[0]
                            if clean in seen_urls:
                                continue
                            seen_urls.add(clean)
                            lines = [l.strip() for l in card.inner_text().split("\n") if l.strip()]
                            company_links.append((lines[0] if lines else "Agency", clean))
                        except Exception:
                            continue
                    page_num += 1

                for name, comp_url in company_links:
                    if added_count[0] >= target or self.stop_event.is_set():
                        break
                    try:
                        page.goto(comp_url.rstrip("/") + "/about/", wait_until="domcontentloaded", timeout=15000)
                        if self.stop_event.wait(random.uniform(2.5, 4)):
                            break
                        fields = page.evaluate("""() => {
                            const r = {};
                            document.querySelectorAll('dl dt').forEach(dt => {
                                const k = dt.innerText.trim().toLowerCase();
                                const dd = dt.nextElementSibling;
                                if (dd && dd.tagName === 'DD') r[k] = dd.innerText.trim();
                            });
                            return r;
                        }""")
                        website, industry, size_txt, phone, email = "", kw, "", "", ""
                        for k, v in fields.items():
                            cv = v.split("\n")[0].strip()
                            if "website" in k:
                                website = cv
                            elif "industry" in k:
                                industry = cv
                            elif "company size" in k:
                                size_txt = cv
                            elif "phone" in k:
                                phone = extract_mobile_number(cv) or cv

                        main_txt = page.locator("main").inner_text() if page.locator("main").count() else ""
                        if not email:
                            email = extract_email(main_txt) or ""
                        if not extract_mobile_number(phone):
                            mt = extract_mobile_number(main_txt)
                            if mt:
                                phone = mt

                        if cfg["enrich_web"] and website and (not email or not extract_mobile_number(phone)):
                            wm, wwa = enrich_from_website(website)
                            if not email and wm:
                                email = wm
                            if not extract_mobile_number(phone) and wwa:
                                phone = wwa

                        accept_fn({
                            "name": name, "source": "LinkedIn", "category": industry,
                            "size_rating": size_txt, "location": loc, "website": website,
                            "phone": phone, "email": email, "profile_url": comp_url,
                        })
                    except Exception:
                        continue
            ctx.close()

    def _scrape_web(self, cfg, keywords, loc, target, added_count, accept_fn):
        self.log("--- بدء السحب من محركات البحث (Web) ---")
        with sync_playwright() as p:
            browser = launch_temp_browser(p)
            page = browser.new_page(viewport={"width": 1280, "height": 800})
            seen_domains = set()

            for kw in keywords:
                if added_count[0] >= target or self.stop_event.is_set():
                    break
                query = f'"{kw}" "{loc}" (email OR whatsapp OR contact)' if loc else f'"{kw}" contact'
                self.log(f"[Web] بحث: {query}")
                for first_idx in (1, 11, 21):
                    if added_count[0] >= target or self.stop_event.is_set():
                        break
                    page.goto(f"https://www.bing.com/search?q={quote(query)}&first={first_idx}", wait_until="domcontentloaded")
                    if self.stop_event.wait(3):
                        break
                    items = page.locator("li.b_algo h2 a")
                    links = [(items.nth(i).inner_text().strip(), items.nth(i).get_attribute("href")) for i in range(items.count())]
                    for title, url in links:
                        if added_count[0] >= target or self.stop_event.is_set():
                            break
                        if not url or not url.startswith("http") or any(x in url.lower() for x in ("facebook.com", "youtube.com", "instagram.com", "wikipedia.org")):
                            continue
                        dom = re.sub(r"^https?://(www\.)?", "", url).split("/")[0].lower()
                        if dom in seen_domains:
                            continue
                        seen_domains.add(dom)
                        email, wa = enrich_from_website(url)
                        accept_fn({
                            "name": title[:60] or dom, "source": "Web Search", "category": kw,
                            "size_rating": "", "location": loc, "website": url,
                            "phone": wa or "", "email": email or "", "profile_url": url,
                        })
            browser.close()

    # ======================= محرك الحملات (Gmail + WhatsApp + LinkedIn) =======================

    def _launch_campaign_for_selected(self):
        if not self.selected_lead_ids:
            messagebox.showwarning("تحديد", "حدد عميلاً واحداً على الأقل من الجدول.")
            return
        chosen = set(self.selected_lead_ids)
        self._show_page("campaigns")
        self._trigger_campaign(chosen)

    def _trigger_campaign(self, selected_ids):
        self._save_campaign_settings(popup=False)
        tpl_title = self.cmb_camp_tpl.get()
        tpls = self.db.get_templates()
        tpl = next((t for t in tpls if t["title"] == tpl_title), tpls[0] if tpls else None)
        if not tpl:
            messagebox.showwarning("لا يوجد قالب", "أضف قالب رسالة أولاً.")
            return

        if selected_ids is not None:
            leads = [self.db.get_lead_by_id(i) for i in selected_ids]
            leads = [l for l in leads if l]
        else:
            leads = self.db.get_leads(stage_filter=self.cmb_camp_stage.get())

        if not leads:
            messagebox.showwarning("لا يوجد عملاء", "لا يوجد عملاء في النطاق المختار.")
            return

        camp_cfg = {
            "use_gmail": self.var_ch_gmail.get(),
            "use_wa": self.var_ch_wa.get(),
            "use_li": self.var_ch_li.get(),
            "dry_run": self.var_dry_run.get(),
            "tpl": tpl,
            "leads": leads,
            "settings": self.db.get_settings(),
        }
        if not (camp_cfg["use_gmail"] or camp_cfg["use_wa"] or camp_cfg["use_li"]):
            messagebox.showwarning("اختر قناة", "فعّل قناة واحدة على الأقل (Gmail أو WhatsApp أو LinkedIn).")
            return

        self._start_task(self._worker_campaign, f"حملة متعددة القنوات ({len(leads)} عميل)", camp_cfg)

    def _worker_campaign(self, c):
        leads = c["leads"]
        tpl = c["tpl"]
        s = c["settings"]
        dry = c["dry_run"]

        if c["use_gmail"] and not self.stop_event.is_set():
            self._phase_gmail(leads, tpl, s, dry)
        if c["use_wa"] and not self.stop_event.is_set():
            self._phase_whatsapp(leads, tpl, s, dry)
        if c["use_li"] and not self.stop_event.is_set():
            self._phase_linkedin(leads, tpl, s, dry)

        self.log("انتهت الحملة بالكامل.")

    def _phase_gmail(self, leads, tpl, s, dry):
        self.log("--- [Gmail] بدء الإرسال ---")
        em, pw, sn = s.get("sender_email", ""), s.get("app_password", "").replace(" ", ""), s.get("sender_name", "")
        att = s.get("attachment_path", "")
        max_e = int(s.get("max_emails") or 35)
        if not dry and (not em or not pw):
            self.log("تخطي Gmail: بيانات الحساب غير مكتملة.")
            return

        sent = 0
        for lead in leads:
            if self.stop_event.is_set() or sent >= max_e:
                break
            if (lead.get("email_status") or "").startswith("تم") or not is_valid_email(lead.get("email")):
                continue
            to_em = lead["email"].strip()
            if dry:
                self.log(f"[تجربة Gmail] {lead['name']} -> {to_em}")
                sent += 1
                continue
            try:
                msg = EmailMessage()
                msg["From"] = f"{sn} <{em}>" if sn else em
                msg["To"] = to_em
                msg["Subject"] = render_msg(tpl["subject"], lead, sn)
                msg.set_content(render_msg(tpl["body"], lead, sn), charset="utf-8")
                if att and os.path.isfile(att):
                    mt, _ = mimetypes.guess_type(att)
                    main_t, sub_t = (mt or "application/octet-stream").split("/", 1)
                    with open(att, "rb") as f:
                        msg.add_attachment(f.read(), maintype=main_t, subtype=sub_t, filename=os.path.basename(att))
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
                    smtp.login(em, pw)
                    smtp.send_message(msg)
                self.db.update_channel_status(lead["id"], "email_status", f"تم ({datetime.now().strftime('%m-%d %H:%M')})", promote_stage=True)
                sent += 1
                self.log(f"[Gmail {sent}/{max_e}] ✓ {lead['name']} ({to_em})")
                if sent < max_e:
                    wt = random.uniform(35, 65)
                    self.log(f"انتظار {wt:.1f} ثانية...")
                    if self.stop_event.wait(wt):
                        break
            except Exception as e:
                self.db.update_channel_status(lead["id"], "email_status", f"فشل: {str(e)[:20]}")
                self.log(f"[Gmail] ✗ {lead['name']}: {e}")

    def _phase_whatsapp(self, leads, tpl, s, dry):
        self.log("--- [WhatsApp Web] بدء الإرسال ---")
        max_w = int(s.get("max_wa") or 20)
        sn = s.get("sender_name", "")
        targets = []
        for lead in leads:
            if (lead.get("wa_status") or "").startswith(("تم", "غير مسجل")):
                continue
            ph = format_whatsapp_phone(lead.get("phone"))
            if ph:
                targets.append((lead, ph))
        if not targets:
            self.log("[WhatsApp] مفيش أرقام جاهزة.")
            return
        if dry:
            for lead, ph in targets[:max_w]:
                self.log(f"[تجربة WhatsApp] {lead['name']} -> +{ph}")
            return

        sent = 0
        with sync_playwright() as p:
            ctx = launch_persistent_browser(p, WA_SESSION_DIR)
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto("https://web.whatsapp.com", wait_until="domcontentloaded")
            if not self.ask_main_thread("تأكيد WhatsApp Web", "تأكد إن المحادثات ظهرت في WhatsApp Web ثم اضغط OK للبدء."):
                ctx.close()
                return

            send_sel = "button[aria-label='Send'], button[aria-label='إرسال'], span[data-icon='send'], span[data-icon='wds-ic-send-filled']"
            inp_sel = "footer div[contenteditable='true']"
            dlg_sel = "div[data-animate-modal-popup='true'], div[role='dialog']"

            for lead, ph in targets:
                if self.stop_event.is_set() or sent >= max_w:
                    break
                msg = render_msg(tpl["body"], lead, sn)
                self.log(f"[WhatsApp] إرسال إلى {lead['name']} (+{ph})...")
                page.goto(f"https://web.whatsapp.com/send?phone={ph}&text={quote(msg)}", wait_until="domcontentloaded")

                state, t0 = "timeout", time.time()
                while time.time() - t0 < 35 and not self.stop_event.is_set():
                    if page.locator(dlg_sel).count():
                        try:
                            if any(k in page.locator(dlg_sel).first.inner_text().lower() for k in ("invalid", "غير صالح", "غير مسجل")):
                                state = "invalid"
                                break
                        except Exception:
                            pass
                    if page.locator(send_sel).count():
                        state = "btn"
                        break
                    if page.locator(inp_sel).count():
                        try:
                            if page.locator(inp_sel).last.inner_text().strip():
                                state = "enter"
                                break
                        except Exception:
                            pass
                    time.sleep(1)

                if state == "invalid":
                    self.db.update_channel_status(lead["id"], "wa_status", "غير مسجل")
                    self.log("الرقم غير مسجل على واتساب.")
                    continue
                if state == "timeout" or self.stop_event.is_set():
                    self.db.update_channel_status(lead["id"], "wa_status", "فشل (Timeout)")
                    continue

                time.sleep(random.uniform(1.5, 2.8))
                if state == "btn":
                    page.locator(send_sel).first.click()
                else:
                    page.locator(inp_sel).last.click()
                    page.keyboard.press("Enter")
                time.sleep(4)
                self.db.update_channel_status(lead["id"], "wa_status", f"تم ({datetime.now().strftime('%m-%d %H:%M')})", promote_stage=True)
                sent += 1
                self.log(f"[WhatsApp {sent}/{max_w}] ✓ تم الإرسال.")
                if sent < max_w:
                    wt = random.uniform(45, 85)
                    self.log(f"انتظار {wt:.1f} ثانية...")
                    if self.stop_event.wait(wt):
                        break
            ctx.close()

    def _phase_linkedin(self, leads, tpl, s, dry):
        self.log("--- [LinkedIn] بدء الإرسال ---")
        max_l = int(s.get("max_li") or 15)
        sn = s.get("sender_name", "")
        targets = [l for l in leads if "linkedin.com/company" in (l.get("profile_url") or "").lower() and not (l.get("li_status") or "").startswith(("تم", "مغلق"))]
        if not targets:
            self.log("[LinkedIn] مفيش صفحات LinkedIn جاهزة.")
            return
        if dry:
            for l in targets[:max_l]:
                self.log(f"[تجربة LinkedIn] {l['name']} -> {l['profile_url']}")
            return

        sent = 0
        with sync_playwright() as p:
            ctx = launch_persistent_browser(p, LINKEDIN_SESSION_DIR)
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto("https://www.linkedin.com/feed/", wait_until="domcontentloaded")
            if not self.ask_main_thread("تأكيد LinkedIn", "تأكد إن حساب LinkedIn مفتوح ثم اضغط OK للبدء."):
                ctx.close()
                return

            for lead in targets:
                if self.stop_event.is_set() or sent >= max_l:
                    break
                msg = render_msg(tpl["body"], lead, sn)
                try:
                    page.goto(lead["profile_url"], wait_until="domcontentloaded", timeout=15000)
                    if self.stop_event.wait(random.uniform(3, 5)):
                        break
                    btn = page.locator("button:has-text('Message'), button:has-text('رسالة'), a:has-text('Message')").first
                    if btn.count() == 0 or not btn.is_visible():
                        self.db.update_channel_status(lead["id"], "li_status", "مغلق")
                        self.log(f"[LinkedIn] {lead['name']}: الرسائل مغلقة.")
                        continue
                    btn.click()
                    time.sleep(2.5)
                    box = page.locator("div.msg-form__contenteditable[contenteditable='true'], div[role='textbox']").last
                    if box.count() > 0:
                        box.click()
                        box.fill(msg)
                        time.sleep(1.5)
                        sub = page.locator("button.msg-form__send-button, button[type='submit']:has-text('Send'), button[type='submit']:has-text('إرسال')").last
                        if sub.count() > 0 and sub.is_enabled():
                            sub.click()
                            time.sleep(3)
                            self.db.update_channel_status(lead["id"], "li_status", f"تم ({datetime.now().strftime('%m-%d %H:%M')})", promote_stage=True)
                            sent += 1
                            self.log(f"[LinkedIn {sent}/{max_l}] ✓ {lead['name']}")
                            if sent < max_l:
                                wt = random.uniform(40, 80)
                                self.log(f"انتظار {wt:.1f} ثانية...")
                                if self.stop_event.wait(wt):
                                    break
                except Exception as e:
                    self.log(f"[LinkedIn] ✗ {lead['name']}: {e}")
            ctx.close()


if __name__ == "__main__":
    app = SmartCRMFullApp()
    app.mainloop()
