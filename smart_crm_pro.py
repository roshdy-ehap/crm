import os
import re
import time
import queue
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

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB_FILE = "crm_pro.db"

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
        "السلام عليكم {name}،\nأتمنى تكون بخير.\nسعداء بالتواصل معك بخصوص مجال {category} في {location}.\nلو وقتك يسمح، ممكن نرتب مكالمة سريعة؟\n\nتحياتي،\n{sender_name}",
        "الكل",
    ),
    (
        "عرض عقارات / سيارات",
        "عروض متاحة في حدود ميزانيتك ({budget})",
        "أهلاً أ/ {name}،\nمتاح عندنا حالياً خيارات مميزة في {category} بمنطقة {location} مناسبة لميزانية {budget}.\nتحب نبعتلك التفاصيل والصور أو نحدد موعد معاينة؟",
        "WhatsApp",
    ),
    (
        "متابعة بعد المعاينة",
        "متابعة بخصوص معاينة {category}",
        "السلام عليكم أ/ {name}،\nحبينا نطمن على رأيك بعد المعاينة وهل فيه أي تفاصيل إضافية محتاج نوضحها بخصوص العرض؟\n\nتحياتي،\n{sender_name}",
        "الكل",
    ),
]


# ======================= دوال مساعدة وتنسيق =======================

class SafeDict(dict):
    def __missing__(self, key):
        return "{" + key + "}"


def render_msg(template_text, lead, sender_name=""):
    return (template_text or "").format_map(SafeDict(
        name=lead.get("name") or "العميل",
        company_name=lead.get("name") or "العميل",
        category=lead.get("category") or "طلبكم",
        budget=lead.get("budget") or "الميزانية المحددة",
        location=lead.get("location") or "",
        source=lead.get("source") or "",
        sender_name=sender_name or "",
    ))


def format_whatsapp_phone(raw_phone):
    """تنظيف الرقم وتحويله لصيغة دولية صالحة لواتساب (يدعم مصر والدولي)"""
    if not raw_phone:
        return None
    digits = re.sub(r"\D", "", str(raw_phone))
    if not digits:
        return None
    if digits.startswith("0020"):
        digits = digits[2:]
    elif digits.startswith("01") and len(digits) == 11:
        digits = "2" + digits
    elif digits.startswith("1") and len(digits) == 10:
        digits = "20" + digits
    if digits.startswith("20") and not digits.startswith("201"):
        return None  # تجاهل الأرضي المصري
    if 10 <= len(digits) <= 15:
        return digits
    return None


def is_valid_email(val):
    if not val:
        return False
    s = str(val).strip()
    return "@" in s and "." in s and "غير متوفر" not in s


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
                    phone TEXT DEFAULT '',
                    email TEXT DEFAULT '',
                    category TEXT DEFAULT '',
                    budget TEXT DEFAULT '',
                    location TEXT DEFAULT '',
                    source TEXT DEFAULT 'يدوي',
                    stage TEXT DEFAULT 'جديد',
                    next_followup TEXT DEFAULT '',
                    notes TEXT DEFAULT '',
                    last_contact TEXT DEFAULT '',
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
                "sender_email": "",
                "app_password": "",
                "sender_name": "",
                "attachment_path": "",
                "max_emails": "35",
                "min_delay": "30",
                "max_delay": "60",
                "appearance": "Dark",
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

    def get_leads(self, search="", stage_filter="الكل"):
        query = "SELECT * FROM leads WHERE 1=1"
        params = []
        if stage_filter and stage_filter != "الكل":
            query += " AND stage = ?"
            params.append(stage_filter)
        if search.strip():
            like = f"%{search.strip()}%"
            query += " AND (name LIKE ? OR phone LIKE ? OR email LIKE ? OR category LIKE ? OR location LIKE ? OR notes LIKE ?)"
            params.extend([like] * 6)
        query += " ORDER BY id DESC"
        with self._connect() as conn:
            return [dict(r) for r in conn.execute(query, params).fetchall()]

    def get_lead_by_id(self, lead_id):
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM leads WHERE id = ?", (lead_id,)).fetchone()
            return dict(row) if row else None

    def add_lead(self, data: dict):
        with self._connect() as conn:
            conn.execute("""
                INSERT INTO leads (name, phone, email, category, budget, location, source, stage, next_followup, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                data["name"], data.get("phone", ""), data.get("email", ""),
                data.get("category", ""), data.get("budget", ""), data.get("location", ""),
                data.get("source", "يدوي"), data.get("stage", "جديد"),
                data.get("next_followup", ""), data.get("notes", ""),
            ))

    def update_lead(self, lead_id, data: dict):
        with self._connect() as conn:
            conn.execute("""
                UPDATE leads
                SET name=?, phone=?, email=?, category=?, budget=?, location=?,
                    source=?, stage=?, next_followup=?, notes=?
                WHERE id=?
            """, (
                data["name"], data.get("phone", ""), data.get("email", ""),
                data.get("category", ""), data.get("budget", ""), data.get("location", ""),
                data.get("source", "يدوي"), data.get("stage", "جديد"),
                data.get("next_followup", ""), data.get("notes", ""), lead_id,
            ))

    def update_stage(self, lead_id, stage):
        with self._connect() as conn:
            conn.execute("UPDATE leads SET stage = ? WHERE id = ?", (stage, lead_id))

    def mark_contacted(self, lead_id, note_text):
        with self._connect() as conn:
            lead = conn.execute("SELECT stage FROM leads WHERE id = ?", (lead_id,)).fetchone()
            new_stage = "تم التواصل" if (lead and lead["stage"] == "جديد") else (lead["stage"] if lead else "تم التواصل")
            conn.execute(
                "UPDATE leads SET last_contact = ?, stage = ? WHERE id = ?",
                (note_text, new_stage, lead_id),
            )

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
    ws.title = "CRM Pipeline"

    headers = [
        "ID", "اسم العميل / الشركة", "الهاتف / واتساب", "الإيميل", "التخصص / الطلب",
        "الميزانية", "اللوكيشن", "المصدر", "المرحلة", "موعد المتابعة",
        "آخر تواصل", "ملاحظات", "فتح واتساب", "إرسال إيميل",
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
            lead["id"], lead["name"], lead.get("phone") or "غير متوفر",
            email or "غير متوفر", lead.get("category") or "", lead.get("budget") or "",
            lead.get("location") or "", lead.get("source") or "", lead.get("stage") or "جديد",
            lead.get("next_followup") or "", lead.get("last_contact") or "",
            lead.get("notes") or "", "", "",
        ])
        ws.row_dimensions[r_idx].height = 24

        wa_c = ws.cell(row=r_idx, column=13)
        wa_c.alignment, wa_c.border = center, border
        if wa_num:
            msg = quote(render_msg(template_body, lead, sender_name))
            wa_c.value = "💬 واتساب"
            wa_c.hyperlink = f"https://wa.me/{wa_num}?text={msg}"
            wa_c.fill, wa_c.font = wa_fill, btn_font
        else:
            wa_c.value = "لا يوجد رقم"
            wa_c.fill, wa_c.font = off_fill, off_font

        em_c = ws.cell(row=r_idx, column=14)
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

    widths = [7, 24, 17, 25, 18, 14, 15, 13, 13, 14, 18, 28, 14, 14]
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

    c_name = find_col(["name", "company", "اسم", "العميل", "الشركة"])
    c_phone = find_col(["phone", "whatsapp", "mobile", "هاتف", "واتساب", "رقم"])
    c_email = find_col(["email", "mail", "إيميل", "ايميل"])
    c_cat = find_col(["industry", "category", "تخصص", "مجال", "طلب"])
    c_budget = find_col(["budget", "ميزانية", "سعر"])
    c_loc = find_col(["location", "city", "لوكيشن", "مدينة", "منطقة"])
    c_src = find_col(["source", "مصدر"])
    c_stage = find_col(["stage", "مرحلة"])
    c_notes = find_col(["notes", "ملاحظات", "website"])

    if c_name is None:
        c_name = 0

    existing_keys = {
        (l["name"].strip().lower(), re.sub(r"\D", "", l["phone"] or ""))
        for l in db.get_leads()
    }

    added = 0
    for row in rows[1:]:
        if not row or len(row) <= c_name or not row[c_name]:
            continue
        name = str(row[c_name]).strip()
        phone = str(row[c_phone]).strip() if c_phone is not None and len(row) > c_phone and row[c_phone] else ""
        if "غير متوفر" in phone:
            phone = ""
        email = str(row[c_email]).strip() if c_email is not None and len(row) > c_email and row[c_email] else ""
        if "غير متوفر" in email:
            email = ""

        key = (name.lower(), re.sub(r"\D", "", phone))
        if key in existing_keys:
            continue

        stage_val = str(row[c_stage]).strip() if c_stage is not None and len(row) > c_stage and row[c_stage] else "جديد"
        if stage_val not in STAGES:
            stage_val = "جديد"

        db.add_lead({
            "name": name,
            "phone": phone,
            "email": email,
            "category": str(row[c_cat]).strip() if c_cat is not None and len(row) > c_cat and row[c_cat] else "",
            "budget": str(row[c_budget]).strip() if c_budget is not None and len(row) > c_budget and row[c_budget] else "",
            "location": str(row[c_loc]).strip() if c_loc is not None and len(row) > c_loc and row[c_loc] else "",
            "source": str(row[c_src]).strip() if c_src is not None and len(row) > c_src and row[c_src] else "استيراد Excel",
            "stage": stage_val,
            "notes": str(row[c_notes]).strip() if c_notes is not None and len(row) > c_notes and row[c_notes] else "",
        })
        existing_keys.add(key)
        added += 1

    return added


# ======================= نافذة إضافة / تعديل عميل =======================

class LeadDialog(ctk.CTkToplevel):
    def __init__(self, parent, on_save, lead=None):
        super().__init__(parent)
        self.on_save = on_save
        self.lead = lead
        self.title("تعديل بيانات العميل" if lead else "إضافة عميل جديد")
        self.geometry("540x610")
        self.resizable(False, False)
        self.grab_set()

        self._build_form()

    def _build_form(self):
        container = ctk.CTkScrollableFrame(self)
        container.pack(fill="both", expand=True, padx=16, pady=12)

        def add_field(label_text, default_val=""):
            ctk.CTkLabel(container, text=label_text, font=("Segoe UI", 12, "bold"), anchor="e").pack(fill="x", pady=(6, 2))
            entry = ctk.CTkEntry(container, justify="right")
            entry.pack(fill="x", pady=(0, 4))
            if default_val:
                entry.insert(0, str(default_val))
            return entry

        L = self.lead or {}
        self.e_name = add_field("اسم العميل أو الشركة *:", L.get("name", ""))
        self.e_phone = add_field("رقم الهاتف / واتساب:", L.get("phone", ""))
        self.e_email = add_field("البريد الإلكتروني:", L.get("email", ""))
        self.e_category = add_field("الطلب / التخصص (مثلاً: شقة، سيارة، تسويق):", L.get("category", ""))
        self.e_budget = add_field("الميزانية المتوقعة (مثلاً: 1,000,000 ج.م):", L.get("budget", ""))
        self.e_location = add_field("المنطقة / المدينة:", L.get("location", ""))
        self.e_source = add_field("مصدر العميل (إعلان فيسبوك، LinkedIn، ترشيح...):", L.get("source", "يدوي"))

        ctk.CTkLabel(container, text="مرحلة المبيعات (Pipeline Stage):", font=("Segoe UI", 12, "bold"), anchor="e").pack(fill="x", pady=(6, 2))
        self.cmb_stage = ctk.CTkComboBox(container, values=STAGES, state="readonly")
        self.cmb_stage.set(L.get("stage", "جديد"))
        self.cmb_stage.pack(fill="x", pady=(0, 4))

        self.e_followup = add_field("تاريخ المتابعة القادم (YYYY-MM-DD):", L.get("next_followup", ""))

        ctk.CTkLabel(container, text="ملاحظات:", font=("Segoe UI", 12, "bold"), anchor="e").pack(fill="x", pady=(6, 2))
        self.txt_notes = ctk.CTkTextbox(container, height=80)
        self.txt_notes.pack(fill="x", pady=(0, 6))
        if L.get("notes"):
            self.txt_notes.insert("1.0", L["notes"])

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkButton(btn_row, text="حفظ البيانات", fg_color="#10B981", hover_color="#059669", command=self._submit).pack(side="right", padx=6, expand=True, fill="x")
        ctk.CTkButton(btn_row, text="إلغاء", fg_color="#64748B", hover_color="#475569", command=self.destroy).pack(side="left", padx=6)

    def _submit(self):
        name = self.e_name.get().strip()
        if not name:
            messagebox.showwarning("حقل مطلوب", "لازم تكتب اسم العميل أو الشركة.", parent=self)
            return

        data = {
            "name": name,
            "phone": self.e_phone.get().strip(),
            "email": self.e_email.get().strip(),
            "category": self.e_category.get().strip(),
            "budget": self.e_budget.get().strip(),
            "location": self.e_location.get().strip(),
            "source": self.e_source.get().strip() or "يدوي",
            "stage": self.cmb_stage.get(),
            "next_followup": self.e_followup.get().strip(),
            "notes": self.txt_notes.get("1.0", "end-1c").strip(),
        }
        self.on_save(data, self.lead["id"] if self.lead else None)
        self.destroy()


# ======================= التطبيق الرئيسي (CustomTkinter) =======================

class SmartCRMApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()
        self.settings = self.db.get_settings()

        ctk.set_appearance_mode(self.settings.get("appearance", "Dark"))
        ctk.set_default_color_theme("blue")

        self.title("Smart Sales CRM Pro — إدارة العملاء ومراحل المبيعات")
        self.geometry("1200x760")
        self.minsize(1020, 650)

        self.selected_lead_ids = set()
        self.log_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.is_sending = False

        self._configure_treeview_style()
        self._build_layout()
        self._show_page("dashboard")
        self._poll_logs()

    def _configure_treeview_style(self):
        is_dark = ctk.get_appearance_mode() == "Dark"
        bg = "#1E293B" if is_dark else "#FFFFFF"
        fg = "#F8FAFC" if is_dark else "#0F172A"
        head_bg = "#0F172A" if is_dark else "#E2E8F0"
        sel_bg = "#2563EB"

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "CRM.Treeview",
            background=bg, foreground=fg, fieldbackground=bg,
            rowheight=32, font=("Segoe UI", 10), borderwidth=0,
        )
        style.configure(
            "CRM.Treeview.Heading",
            background=head_bg, foreground=fg,
            font=("Segoe UI", 10, "bold"), relief="flat", padding=6,
        )
        style.map("CRM.Treeview", background=[("selected", sel_bg)], foreground=[("selected", "#FFFFFF")])

    def _build_layout(self):
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar.pack(side="right", fill="y")
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(
            self.sidebar, text="Smart CRM Pro", font=("Segoe UI", 20, "bold"), text_color="#38BDF8"
        ).pack(pady=(22, 4))
        ctk.CTkLabel(
            self.sidebar, text="نظام إدارة العملاء والمبيعات", font=("Segoe UI", 11), text_color="#94A3B8"
        ).pack(pady=(0, 20))

        self.nav_buttons = {}
        nav_items = [
            ("dashboard", "📊  لوحة التحكم"),
            ("leads", "👥  العملاء ومراحل البيع"),
            ("templates", "📝  قوالب الرسائل"),
            ("campaigns", "🚀  الحملات والإعدادات"),
        ]
        for key, text in nav_items:
            btn = ctk.CTkButton(
                self.sidebar, text=text, anchor="e", height=42,
                font=("Segoe UI", 13, "bold"), corner_radius=8,
                fg_color="transparent", hover_color=("#CBD5E1", "#334155"),
                text_color=("#0F172A", "#F8FAFC"),
                command=lambda k=key: self._show_page(k),
            )
            btn.pack(fill="x", padx=12, pady=4)
            self.nav_buttons[key] = btn

        ctk.CTkButton(
            self.sidebar, text="+ عميل جديد", height=40,
            font=("Segoe UI", 13, "bold"), fg_color="#10B981", hover_color="#059669",
            command=self._open_add_lead,
        ).pack(fill="x", padx=14, pady=(24, 8))

        ctk.CTkLabel(self.sidebar, text="مظهر البرنامج:", font=("Segoe UI", 11)).pack(side="bottom", pady=(0, 4))
        self.mode_switch = ctk.CTkSegmentedButton(
            self.sidebar, values=["Dark", "Light"], command=self._change_theme
        )
        self.mode_switch.set(self.settings.get("appearance", "Dark"))
        self.mode_switch.pack(side="bottom", padx=14, pady=(0, 14), fill="x")

        self.main_area = ctk.CTkFrame(self, fg_color="transparent")
        self.main_area.pack(side="left", fill="both", expand=True, padx=14, pady=12)

        self.pages = {
            "dashboard": ctk.CTkFrame(self.main_area, fg_color="transparent"),
            "leads": ctk.CTkFrame(self.main_area, fg_color="transparent"),
            "templates": ctk.CTkFrame(self.main_area, fg_color="transparent"),
            "campaigns": ctk.CTkFrame(self.main_area, fg_color="transparent"),
        }

        self._build_dashboard_page(self.pages["dashboard"])
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

    # ======================= صفحة 1: لوحة التحكم =======================

    def _build_dashboard_page(self, parent):
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(header, text="لوحة التحكم وملخص مراحل المبيعات", font=("Segoe UI", 22, "bold")).pack(side="right")
        ctk.CTkButton(header, text="تحديث الإحصائيات", width=130, command=self.refresh_dashboard).pack(side="left")

        self.kpi_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self.kpi_frame.pack(fill="x", pady=(0, 14))
        self.kpi_labels = {}

        cards_def = [
            ("total", "إجمالي العملاء", "#0EA5E9"),
            ("جديد", "عملاء جدد", STAGE_COLORS["جديد"]),
            ("مهتم", "مهتمين", STAGE_COLORS["مهتم"]),
            ("معاينة", "معاينات", STAGE_COLORS["معاينة"]),
            ("تم البيع", "تم البيع", STAGE_COLORS["تم البيع"]),
            ("rate", "نسبة التحويل", "#8B5CF6"),
        ]
        for idx, (k, title, color) in enumerate(cards_def):
            card = ctk.CTkFrame(self.kpi_frame, corner_radius=12, border_width=2, border_color=color)
            card.grid(row=0, column=idx, padx=6, pady=4, sticky="nsew")
            self.kpi_frame.columnconfigure(idx, weight=1)

            ctk.CTkLabel(card, text=title, font=("Segoe UI", 12, "bold"), text_color="#94A3B8").pack(pady=(12, 2))
            val_lbl = ctk.CTkLabel(card, text="0", font=("Segoe UI", 26, "bold"), text_color=color)
            val_lbl.pack(pady=(0, 12))
            self.kpi_labels[k] = val_lbl

        due_box = ctk.CTkFrame(parent, corner_radius=12)
        due_box.pack(fill="both", expand=True, pady=(4, 0))

        top_due = ctk.CTkFrame(due_box, fg_color="transparent")
        top_due.pack(fill="x", padx=14, pady=(12, 6))
        ctk.CTkLabel(top_due, text="⏰ متابعات مستحقة اليوم أو متأخرة", font=("Segoe UI", 15, "bold")).pack(side="right")
        ctk.CTkButton(
            top_due, text="الذهاب لجدول العملاء", width=150,
            command=lambda: self._show_page("leads"),
        ).pack(side="left")

        self.due_list_frame = ctk.CTkScrollableFrame(due_box)
        self.due_list_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    def refresh_dashboard(self):
        total, counts, due_leads = self.db.get_pipeline_stats()
        self.kpi_labels["total"].configure(text=str(total))
        for st in ("جديد", "مهتم", "معاينة", "تم البيع"):
            self.kpi_labels[st].configure(text=str(counts.get(st, 0)))
        conv = (counts.get("تم البيع", 0) / total * 100) if total else 0.0
        self.kpi_labels["rate"].configure(text=f"{conv:.1f}%")

        for w in self.due_list_frame.winfo_children():
            w.destroy()

        if not due_leads:
            ctk.CTkLabel(
                self.due_list_frame,
                text="لا توجد متابعات متأخرة حالياً. أضف تاريخ متابعة (YYYY-MM-DD) لأي عميل ليظهر هنا.",
                font=("Segoe UI", 12), text_color="#94A3B8",
            ).pack(pady=30)
            return

        for lead in due_leads:
            row = ctk.CTkFrame(self.due_list_frame, corner_radius=8)
            row.pack(fill="x", pady=4, padx=4)
            info = f"{lead['name']}   |   المرحلة: {lead['stage']}   |   الموعد: {lead['next_followup']}   |   {lead.get('phone') or ''}"
            ctk.CTkLabel(row, text=info, font=("Segoe UI", 12, "bold")).pack(side="right", padx=12, pady=10)
            ctk.CTkButton(
                row, text="💬 واتساب", width=90, fg_color="#25D366", hover_color="#16A34A",
                command=lambda l=lead: self._quick_open_whatsapp(l),
            ).pack(side="left", padx=6, pady=6)
            ctk.CTkButton(
                row, text="تعديل", width=70,
                command=lambda l=lead: self._open_edit_lead(l),
            ).pack(side="left", padx=4, pady=6)

    # ======================= صفحة 2: العملاء ومراحل المبيعات =======================

    def _build_leads_page(self, parent):
        top_bar = ctk.CTkFrame(parent, corner_radius=10)
        top_bar.pack(fill="x", pady=(0, 8))

        self.var_search = tk.StringVar()
        self.var_search.trace_add("write", lambda *_: self.refresh_leads_table())
        ctk.CTkEntry(
            top_bar, textvariable=self.var_search, placeholder_text="🔍 بحث بالاسم، الرقم، التخصص، المنطقة...",
            width=260, justify="right",
        ).pack(side="right", padx=8, pady=8)

        self.cmb_filter_stage = ctk.CTkComboBox(
            top_bar, values=["الكل"] + STAGES, state="readonly", width=130,
            command=lambda _: self.refresh_leads_table(),
        )
        self.cmb_filter_stage.set("الكل")
        self.cmb_filter_stage.pack(side="right", padx=6, pady=8)

        ctk.CTkButton(top_bar, text="+ إضافة عميل", width=105, fg_color="#10B981", hover_color="#059669", command=self._open_add_lead).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_bar, text="📥 استيراد Excel", width=115, fg_color="#475569", hover_color="#334155", command=self._import_excel).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_bar, text="📤 تصدير Excel", width=115, fg_color="#0284C7", hover_color="#0369A1", command=self._export_excel).pack(side="left", padx=5, pady=8)

        pipe_bar = ctk.CTkFrame(parent, fg_color="transparent")
        pipe_bar.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(pipe_bar, text="نقل المحدد لمرحلة:", font=("Segoe UI", 12, "bold")).pack(side="right", padx=(4, 6))
        self.cmb_bulk_stage = ctk.CTkComboBox(pipe_bar, values=STAGES, state="readonly", width=125)
        self.cmb_bulk_stage.set("مهتم")
        self.cmb_bulk_stage.pack(side="right", padx=4)
        ctk.CTkButton(pipe_bar, text="تطبيق المرحلة", width=100, command=self._apply_bulk_stage).pack(side="right", padx=4)

        ctk.CTkButton(pipe_bar, text="تحديد الكل", width=85, fg_color="#334155", command=self._select_all_visible).pack(side="right", padx=(14, 3))
        ctk.CTkButton(pipe_bar, text="إلغاء التحديد", width=90, fg_color="#334155", command=self._clear_selection).pack(side="right", padx=3)

        self.lbl_selected_count = ctk.CTkLabel(pipe_bar, text="المحدد: 0", font=("Segoe UI", 12, "bold"), text_color="#38BDF8")
        self.lbl_selected_count.pack(side="left", padx=8)

        table_frame = ctk.CTkFrame(parent, corner_radius=10)
        table_frame.pack(fill="both", expand=True, pady=(0, 8))

        cols = ("sel", "id", "name", "phone", "email", "category", "budget", "location", "stage", "followup", "last")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", style="CRM.Treeview", selectmode="none")
        specs = [
            ("sel", "تحديد", 50, "center"),
            ("id", "#", 40, "center"),
            ("name", "اسم العميل / الشركة", 165, "e"),
            ("phone", "واتساب / موبايل", 120, "center"),
            ("email", "الإيميل", 155, "w"),
            ("category", "الطلب / التخصص", 120, "center"),
            ("budget", "الميزانية", 100, "center"),
            ("location", "المنطقة", 95, "center"),
            ("stage", "المرحلة", 95, "center"),
            ("followup", "المتابعة", 95, "center"),
            ("last", "آخر تواصل", 125, "center"),
        ]
        for cid, h, w, anc in specs:
            self.tree.heading(cid, text=h)
            self.tree.column(cid, width=w, minwidth=40, anchor=anc)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew", padx=(6, 0), pady=6)
        vsb.grid(row=0, column=1, sticky="ns", pady=6)
        hsb.grid(row=1, column=0, sticky="ew", padx=6)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        act_bar = ctk.CTkFrame(parent, fg_color="transparent")
        act_bar.pack(fill="x")

        ctk.CTkLabel(act_bar, text="قالب الرسالة السريع:", font=("Segoe UI", 12, "bold")).pack(side="right", padx=(4, 6))
        self.cmb_quick_tpl = ctk.CTkComboBox(act_bar, values=[], state="readonly", width=180)
        self.cmb_quick_tpl.pack(side="right", padx=4)

        ctk.CTkButton(
            act_bar, text="💬 فتح واتساب للمحدد", fg_color="#25D366", hover_color="#16A34A",
            command=self._open_whatsapp_for_selected,
        ).pack(side="right", padx=5)

        ctk.CTkButton(
            act_bar, text="✉️ إرسال حملة Gmail للمحدد", fg_color="#EA4335", hover_color="#DC2626",
            command=self._send_gmail_to_selected,
        ).pack(side="right", padx=5)

        ctk.CTkButton(
            act_bar, text="تعديل المحدد", width=95, command=self._edit_first_selected
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            act_bar, text="حذف المحدد", width=95, fg_color="#DC2626", hover_color="#B91C1C",
            command=self._delete_selected,
        ).pack(side="left", padx=4)

    def refresh_leads_table(self):
        templates = self.db.get_templates()
        tpl_names = [t["title"] for t in templates]
        if tpl_names:
            current = self.cmb_quick_tpl.get()
            self.cmb_quick_tpl.configure(values=tpl_names)
            if current not in tpl_names:
                self.cmb_quick_tpl.set(tpl_names[0])

        leads = self.db.get_leads(
            search=self.var_search.get(),
            stage_filter=self.cmb_filter_stage.get(),
        )
        self.tree.delete(*self.tree.get_children())

        is_dark = ctk.get_appearance_mode() == "Dark"
        sel_bg = "#1D4ED8" if is_dark else "#DBEAFE"
        sel_fg = "#FFFFFF" if is_dark else "#0F172A"
        self.tree.tag_configure("checked", background=sel_bg, foreground=sel_fg)

        for lead in leads:
            lid = lead["id"]
            is_sel = lid in self.selected_lead_ids
            mark = "☑" if is_sel else "☐"
            tags = ("checked",) if is_sel else ()
            self.tree.insert(
                "", "end", iid=str(lid),
                values=(
                    mark, lid, lead["name"], lead.get("phone") or "—", lead.get("email") or "—",
                    lead.get("category") or "—", lead.get("budget") or "—", lead.get("location") or "—",
                    lead.get("stage") or "جديد", lead.get("next_followup") or "—", lead.get("last_contact") or "—",
                ),
                tags=tags,
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
            messagebox.showwarning("تحديد", "حدد عميلاً واحداً للتعديل (أو اضغط عليه دبل كليك).")
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
            messagebox.showwarning("تحديد", "حدد عميلاً واحداً على الأقل من الجدول.")
            return
        new_stage = self.cmb_bulk_stage.get()
        for lid in self.selected_lead_ids:
            self.db.update_stage(lid, new_stage)
        self.refresh_leads_table()
        self.refresh_dashboard()

    def _delete_selected(self):
        if not self.selected_lead_ids:
            return
        if messagebox.askyesno("تأكيد الحذف", f"هل تريد حذف {len(self.selected_lead_ids)} عميل نهائياً من قاعدة البيانات؟"):
            self.db.delete_leads(self.selected_lead_ids)
            self.selected_lead_ids.clear()
            self.refresh_leads_table()
            self.refresh_dashboard()

    def _get_selected_template_obj(self, title=None):
        templates = self.db.get_templates()
        target_title = title or self.cmb_quick_tpl.get()
        for t in templates:
            if t["title"] == target_title:
                return t
        return templates[0] if templates else {"title": "عام", "subject": "تواصل مع {name}", "body": "السلام عليكم {name}"}

    def _quick_open_whatsapp(self, lead):
        phone = format_whatsapp_phone(lead.get("phone"))
        if not phone:
            messagebox.showwarning("رقم غير صالح", f"العميل {lead['name']} ليس لديه رقم موبايل صالح لواتساب.")
            return
        tpl = self._get_selected_template_obj()
        sender_name = self.db.get_settings().get("sender_name", "")
        msg = render_msg(tpl["body"], lead, sender_name)
        webbrowser.open(f"https://wa.me/{phone}?text={quote(msg)}")
        self.db.mark_contacted(lead["id"], f"واتساب ({datetime.now().strftime('%m-%d %H:%M')})")
        self.refresh_leads_table()
        self.refresh_dashboard()

    def _open_whatsapp_for_selected(self):
        if not self.selected_lead_ids:
            messagebox.showwarning("تحديد", "حدد العملاء اللي عاوز تفتح شات واتساب ليهم.")
            return
        ids = list(self.selected_lead_ids)
        if len(ids) > 10:
            if not messagebox.askyesno("تنبيه", f"أنت محدد {len(ids)} عميل؛ هيتم فتح نوافذ واتساب ليهم بالتدريج. تكمل؟"):
                return

        tpl = self._get_selected_template_obj()
        sender_name = self.db.get_settings().get("sender_name", "")
        opened = 0
        for lid in ids:
            lead = self.db.get_lead_by_id(lid)
            if not lead:
                continue
            phone = format_whatsapp_phone(lead.get("phone"))
            if not phone:
                continue
            msg = render_msg(tpl["body"], lead, sender_name)
            webbrowser.open(f"https://wa.me/{phone}?text={quote(msg)}")
            self.db.mark_contacted(lid, f"واتساب ({datetime.now().strftime('%m-%d %H:%M')})")
            opened += 1
        self.refresh_leads_table()
        self.refresh_dashboard()
        if opened == 0:
            messagebox.showinfo("لا توجد أرقام", "العملاء المحددون ليس لديهم أرقام موبايل صالحة لواتساب.")

    def _import_excel(self):
        path = filedialog.askopenfilename(title="اختر ملف Excel للاستيراد", filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return
        try:
            added = import_leads_from_excel(self.db, path)
            self.refresh_leads_table()
            self.refresh_dashboard()
            messagebox.showinfo("تم الاستيراد", f"تمت إضافة {added} عميل جديد إلى قاعدة البيانات (مع تجاهل المكرر).")
        except Exception as e:
            messagebox.showerror("خطأ في الاستيراد", str(e))

    def _export_excel(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"crm_leads_{datetime.now().strftime('%Y%m%d')}.xlsx",
        )
        if not path:
            return
        leads = self.db.get_leads(search=self.var_search.get(), stage_filter=self.cmb_filter_stage.get())
        if not leads:
            messagebox.showwarning("لا توجد بيانات", "لا يوجد عملاء لتصديرهم.")
            return
        tpl = self._get_selected_template_obj()
        sender_name = self.db.get_settings().get("sender_name", "")
        try:
            export_leads_to_excel(leads, path, tpl["subject"], tpl["body"], sender_name)
            if messagebox.askyesno("تم التصدير بنجاح", f"تم تصدير {len(leads)} عميل إلى:\n{path}\n\nهل تريد فتح الملف الآن؟"):
                if platform.system() == "Windows":
                    os.startfile(path)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", path])
                else:
                    subprocess.Popen(["xdg-open", path])
        except PermissionError:
            messagebox.showerror("الملف مفتوح", "اقفل ملف الإكسل الأول وبعدين جرب التصدير تاني.")
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    # ======================= صفحة 3: إدارة قوالب الرسائل =======================

    def _build_templates_page(self, parent):
        ctk.CTkLabel(parent, text="إدارة قوالب الرسائل (للواتساب والإيميل)", font=("Segoe UI", 20, "bold")).pack(anchor="e", pady=(0, 10))

        grid = ctk.CTkFrame(parent, fg_color="transparent")
        grid.pack(fill="both", expand=True)
        grid.columnconfigure(0, weight=2)
        grid.columnconfigure(1, weight=1)
        grid.rowconfigure(0, weight=1)

        editor = ctk.CTkFrame(grid, corner_radius=12)
        editor.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        ctk.CTkLabel(editor, text="اسم القالب:", font=("Segoe UI", 12, "bold")).pack(anchor="e", padx=14, pady=(12, 2))
        self.e_tpl_title = ctk.CTkEntry(editor, justify="right")
        self.e_tpl_title.pack(fill="x", padx=14, pady=(0, 6))

        row_meta = ctk.CTkFrame(editor, fg_color="transparent")
        row_meta.pack(fill="x", padx=14, pady=4)
        ctk.CTkLabel(row_meta, text="القناة:", font=("Segoe UI", 12, "bold")).pack(side="right", padx=(0, 6))
        self.cmb_tpl_channel = ctk.CTkComboBox(row_meta, values=["الكل", "WhatsApp", "Email"], state="readonly", width=130)
        self.cmb_tpl_channel.set("الكل")
        self.cmb_tpl_channel.pack(side="right")

        ctk.CTkLabel(editor, text="عنوان الإيميل (Subject):", font=("Segoe UI", 12, "bold")).pack(anchor="e", padx=14, pady=(6, 2))
        self.e_tpl_subject = ctk.CTkEntry(editor, justify="right")
        self.e_tpl_subject.pack(fill="x", padx=14, pady=(0, 6))

        ctk.CTkLabel(
            editor,
            text="نص الرسالة — المتغيرات المتاحة: {name}  {category}  {budget}  {location}  {sender_name}",
            font=("Segoe UI", 11), text_color="#38BDF8",
        ).pack(anchor="e", padx=14, pady=(6, 2))

        self.txt_tpl_body = ctk.CTkTextbox(editor, font=("Segoe UI", 12))
        self.txt_tpl_body.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        btns = ctk.CTkFrame(editor, fg_color="transparent")
        btns.pack(fill="x", padx=14, pady=(0, 12))
        ctk.CTkButton(btns, text="💾 حفظ القالب", fg_color="#10B981", hover_color="#059669", command=self._save_template).pack(side="right", padx=4)
        ctk.CTkButton(btns, text="👁️ معاينة بمثال", fg_color="#0284C7", hover_color="#0369A1", command=self._preview_template).pack(side="right", padx=4)
        ctk.CTkButton(btns, text="قالب جديد", fg_color="#475569", command=self._clear_template_editor).pack(side="left", padx=4)
        ctk.CTkButton(btns, text="حذف", fg_color="#DC2626", hover_color="#B91C1C", width=80, command=self._delete_template).pack(side="left", padx=4)

        side_list = ctk.CTkFrame(grid, corner_radius=12)
        side_list.grid(row=0, column=1, sticky="nsew")
        ctk.CTkLabel(side_list, text="القوالب المحفوظة", font=("Segoe UI", 14, "bold")).pack(pady=(12, 6))
        self.tpl_list_scroll = ctk.CTkScrollableFrame(side_list)
        self.tpl_list_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 10))

    def refresh_templates_ui(self):
        for w in self.tpl_list_scroll.winfo_children():
            w.destroy()
        templates = self.db.get_templates()
        for t in templates:
            btn = ctk.CTkButton(
                self.tpl_list_scroll,
                text=f"{t['title']}  ({t['channel']})",
                anchor="e", fg_color=("#E2E8F0", "#334155"),
                text_color=("#0F172A", "#F8FAFC"),
                hover_color="#2563EB",
                command=lambda tpl=t: self._load_template_into_editor(tpl),
            )
            btn.pack(fill="x", pady=3, padx=4)
        if templates and not self.e_tpl_title.get().strip():
            self._load_template_into_editor(templates[0])

    def _load_template_into_editor(self, tpl):
        self.e_tpl_title.delete(0, "end")
        self.e_tpl_title.insert(0, tpl["title"])
        self.cmb_tpl_channel.set(tpl["channel"])
        self.e_tpl_subject.delete(0, "end")
        self.e_tpl_subject.insert(0, tpl["subject"])
        self.txt_tpl_body.delete("1.0", "end")
        self.txt_tpl_body.insert("1.0", tpl["body"])

    def _clear_template_editor(self):
        self.e_tpl_title.delete(0, "end")
        self.e_tpl_subject.delete(0, "end")
        self.txt_tpl_body.delete("1.0", "end")
        self.cmb_tpl_channel.set("الكل")

    def _save_template(self):
        title = self.e_tpl_title.get().strip()
        body = self.txt_tpl_body.get("1.0", "end-1c").strip()
        if not title or not body:
            messagebox.showwarning("بيانات ناقصة", "اكتب اسم القالب ونص الرسالة أولاً.")
            return
        self.db.save_template(title, self.e_tpl_subject.get().strip(), body, self.cmb_tpl_channel.get())
        self.refresh_templates_ui()
        messagebox.showinfo("تم الحفظ", f"تم حفظ القالب '{title}' بنجاح.")

    def _delete_template(self):
        title = self.e_tpl_title.get().strip()
        if not title:
            return
        if len(self.db.get_templates()) <= 1:
            messagebox.showwarning("تنبيه", "يجب الإبقاء على قالب واحد على الأقل.")
            return
        if messagebox.askyesno("حذف", f"حذف قالب '{title}'؟"):
            self.db.delete_template(title)
            self._clear_template_editor()
            self.refresh_templates_ui()

    def _preview_template(self):
        sample = {"name": "أحمد طارق", "category": "شقة سكنية", "budget": "1,000,000 ج.م", "location": "المنصورة"}
        sender = self.db.get_settings().get("sender_name") or "فريق المبيعات"
        subj = render_msg(self.e_tpl_subject.get().strip(), sample, sender)
        body = render_msg(self.txt_tpl_body.get("1.0", "end-1c"), sample, sender)
        messagebox.showinfo("معاينة الرسالة", f"العنوان:\n{subj}\n\nالرسالة:\n{body}")

    # ======================= صفحة 4: الحملات والإعدادات =======================

    def _build_campaigns_page(self, parent):
        cfg_box = ctk.CTkFrame(parent, corner_radius=12)
        cfg_box.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(cfg_box, text="إعدادات البريد الإلكتروني (Gmail SMTP)", font=("Segoe UI", 15, "bold")).grid(row=0, column=0, columnspan=4, sticky="e", padx=12, pady=(10, 6))

        s = self.settings
        self.e_sender_email = ctk.CTkEntry(cfg_box, placeholder_text="example@gmail.com")
        self.e_sender_email.insert(0, s.get("sender_email", ""))
        self.e_app_pass = ctk.CTkEntry(cfg_box, placeholder_text="16-char App Password", show="*")
        self.e_app_pass.insert(0, s.get("app_password", ""))
        self.e_sender_name = ctk.CTkEntry(cfg_box, placeholder_text="اسمك أو اسم شركتك", justify="right")
        self.e_sender_name.insert(0, s.get("sender_name", ""))
        self.e_attach = ctk.CTkEntry(cfg_box, placeholder_text="ملف مرفق اختياري (PDF / بروشور)")
        self.e_attach.insert(0, s.get("attachment_path", ""))

        ctk.CTkLabel(cfg_box, text="إيميل Gmail:").grid(row=1, column=3, sticky="e", padx=8, pady=4)
        self.e_sender_email.grid(row=1, column=2, sticky="ew", padx=6, pady=4)
        ctk.CTkLabel(cfg_box, text="App Password:").grid(row=1, column=1, sticky="e", padx=8, pady=4)
        self.e_app_pass.grid(row=1, column=0, sticky="ew", padx=(12, 6), pady=4)

        ctk.CTkLabel(cfg_box, text="اسم المرسل:").grid(row=2, column=3, sticky="e", padx=8, pady=4)
        self.e_sender_name.grid(row=2, column=2, sticky="ew", padx=6, pady=4)
        ctk.CTkLabel(cfg_box, text="مرفق:").grid(row=2, column=1, sticky="e", padx=8, pady=4)
        att_row = ctk.CTkFrame(cfg_box, fg_color="transparent")
        att_row.grid(row=2, column=0, sticky="ew", padx=(12, 6), pady=4)
        self.e_attach.pack(in_=att_row, side="left", fill="x", expand=True)
        ctk.CTkButton(att_row, text="اختر", width=55, command=self._browse_attachment).pack(side="left", padx=(4, 0))

        row_limits = ctk.CTkFrame(cfg_box, fg_color="transparent")
        row_limits.grid(row=3, column=0, columnspan=4, sticky="ew", padx=12, pady=(6, 12))

        self.e_max_emails = ctk.CTkEntry(row_limits, width=60)
        self.e_max_emails.insert(0, s.get("max_emails", "35"))
        self.e_min_delay = ctk.CTkEntry(row_limits, width=60)
        self.e_min_delay.insert(0, s.get("min_delay", "30"))
        self.e_max_delay = ctk.CTkEntry(row_limits, width=60)
        self.e_max_delay.insert(0, s.get("max_delay", "60"))

        ctk.CTkLabel(row_limits, text="حد/حملة:").pack(side="right", padx=4)
        self.e_max_emails.pack(side="right", padx=(0, 12))
        ctk.CTkLabel(row_limits, text="تأخير من (ث):").pack(side="right", padx=4)
        self.e_min_delay.pack(side="right", padx=(0, 8))
        ctk.CTkLabel(row_limits, text="إلى:").pack(side="right", padx=4)
        self.e_max_delay.pack(side="right", padx=(0, 12))

        ctk.CTkButton(row_limits, text="💾 حفظ الإعدادات", width=120, fg_color="#10B981", hover_color="#059669", command=self._save_smtp_settings).pack(side="left", padx=4)
        ctk.CTkButton(row_limits, text="✓ اختبار Gmail", width=110, fg_color="#0284C7", command=self._test_gmail).pack(side="left", padx=4)

        cfg_box.columnconfigure(0, weight=1)
        cfg_box.columnconfigure(2, weight=1)

        camp_box = ctk.CTkFrame(parent, corner_radius=12)
        camp_box.pack(fill="both", expand=True)

        top_c = ctk.CTkFrame(camp_box, fg_color="transparent")
        top_c.pack(fill="x", padx=12, pady=10)

        ctk.CTkLabel(top_c, text="القالب:").pack(side="right", padx=4)
        self.cmb_camp_tpl = ctk.CTkComboBox(top_c, values=[], state="readonly", width=170)
        self.cmb_camp_tpl.pack(side="right", padx=4)

        ctk.CTkLabel(top_c, text="استهداف مرحلة:").pack(side="right", padx=(10, 4))
        self.cmb_camp_stage = ctk.CTkComboBox(top_c, values=["الكل"] + STAGES, state="readonly", width=120)
        self.cmb_camp_stage.set("جديد")
        self.cmb_camp_stage.pack(side="right", padx=4)

        self.var_dry_run = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(top_c, text="وضع التجربة (DRY RUN)", variable=self.var_dry_run).pack(side="right", padx=12)

        self.btn_run_camp = ctk.CTkButton(
            top_c, text="🚀 بدء الحملة", width=115, fg_color="#16A34A", hover_color="#15803D",
            command=lambda: self._start_email_campaign(lead_ids=None),
        )
        self.btn_run_camp.pack(side="left", padx=4)

        self.btn_stop_camp = ctk.CTkButton(
            top_c, text="⛔ إيقاف", width=85, fg_color="#DC2626", hover_color="#B91C1C",
            state="disabled", command=self._stop_campaign,
        )
        self.btn_stop_camp.pack(side="left", padx=4)

        self.log_box = ctk.CTkTextbox(camp_box, font=("Consolas", 11))
        self.log_box.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.log_box.configure(state="disabled")

    def refresh_campaign_templates(self):
        names = [t["title"] for t in self.db.get_templates()]
        if names:
            cur = self.cmb_camp_tpl.get()
            self.cmb_camp_tpl.configure(values=names)
            if cur not in names:
                self.cmb_camp_tpl.set(names[0])

    def _browse_attachment(self):
        p = filedialog.askopenfilename(title="اختر ملفاً لإرفاقه")
        if p:
            self.e_attach.delete(0, "end")
            self.e_attach.insert(0, p)

    def _save_smtp_settings(self, popup=True):
        data = {
            "sender_email": self.e_sender_email.get().strip(),
            "app_password": self.e_app_pass.get().strip(),
            "sender_name": self.e_sender_name.get().strip(),
            "attachment_path": self.e_attach.get().strip(),
            "max_emails": self.e_max_emails.get().strip() or "35",
            "min_delay": self.e_min_delay.get().strip() or "30",
            "max_delay": self.e_max_delay.get().strip() or "60",
        }
        self.db.save_settings(data)
        self.settings.update(data)
        if popup:
            messagebox.showinfo("تم الحفظ", "تم حفظ إعدادات البريد بنجاح.")

    def _test_gmail(self):
        self._save_smtp_settings(popup=False)
        email = self.settings.get("sender_email")
        pwd = self.settings.get("app_password")
        if not email or not pwd:
            messagebox.showwarning("ناقص بيانات", "اكتب إيميل Gmail والـ App Password أولاً.")
            return

        def _work():
            try:
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=20) as smtp:
                    smtp.login(email, pwd.replace(" ", ""))
                self.log("✓ نجح اختبار الاتصال بحساب Gmail.")
                self.after(0, lambda: messagebox.showinfo("نجاح", "تم الاتصال بحساب Gmail بنجاح!"))
            except Exception as e:
                self.log(f"✗ فشل اتصال Gmail: {e}")
                self.after(0, lambda err=str(e): messagebox.showerror("فشل الاتصال", err))

        threading.Thread(target=_work, daemon=True).start()

    def _send_gmail_to_selected(self):
        if not self.selected_lead_ids:
            messagebox.showwarning("تحديد", "حدد عميلاً واحداً على الأقل من الجدول.")
            return
        ids = set(self.selected_lead_ids)
        tpl_title = self.cmb_quick_tpl.get()
        self._show_page("campaigns")
        self.cmb_camp_tpl.set(tpl_title)
        self._start_email_campaign(lead_ids=ids)

    def _start_email_campaign(self, lead_ids=None):
        if self.is_sending:
            return
        self._save_smtp_settings(popup=False)
        dry_run = self.var_dry_run.get()
        if not dry_run and (not self.settings["sender_email"] or not self.settings["app_password"]):
            messagebox.showwarning("إعدادات ناقصة", "اكتب إيميل Gmail والـ App Password أولاً.")
            return

        tpl = self._get_selected_template_obj(self.cmb_camp_tpl.get())
        stage_filter = self.cmb_camp_stage.get()

        if lead_ids is not None:
            leads = [self.db.get_lead_by_id(lid) for lid in lead_ids]
            leads = [l for l in leads if l and is_valid_email(l.get("email"))]
        else:
            leads = [l for l in self.db.get_leads(stage_filter=stage_filter) if is_valid_email(l.get("email"))]

        if not leads:
            self.log("لا يوجد عملاء لديهم إيميل صالح في النطاق المحدد.")
            return

        try:
            max_e = int(self.settings["max_emails"])
            d_min = int(self.settings["min_delay"])
            d_max = int(self.settings["max_delay"])
        except ValueError:
            messagebox.showerror("خطأ", "تأكد إن خانات الحد الأقصى والتأخير أرقام صحيحة.")
            return

        self.is_sending = True
        self.stop_event.clear()
        self.btn_run_camp.configure(state="disabled")
        self.btn_stop_camp.configure(state="normal")
        self.log("=" * 45)
        self.log(f"بدء الحملة | القالب: {tpl['title']} | العملاء: {len(leads)} | DRY RUN: {dry_run}")

        def _worker():
            sent = 0
            sender_email = self.settings["sender_email"]
            app_pass = self.settings["app_password"].replace(" ", "")
            sender_name = self.settings["sender_name"]
            attach = self.settings["attachment_path"]

            for lead in leads:
                if self.stop_event.is_set() or sent >= max_e:
                    break

                to_email = lead["email"].strip()
                if dry_run:
                    self.log(f"[تجربة] {lead['name']} -> {to_email}")
                    sent += 1
                    continue

                try:
                    msg = EmailMessage()
                    msg["From"] = f"{sender_name} <{sender_email}>" if sender_name else sender_email
                    msg["To"] = to_email
                    msg["Subject"] = render_msg(tpl["subject"], lead, sender_name)
                    msg.set_content(render_msg(tpl["body"], lead, sender_name), charset="utf-8")

                    if attach and os.path.isfile(attach):
                        mt, _ = mimetypes.guess_type(attach)
                        main_t, sub_t = (mt or "application/octet-stream").split("/", 1)
                        with open(attach, "rb") as f:
                            msg.add_attachment(f.read(), maintype=main_t, subtype=sub_t, filename=os.path.basename(attach))

                    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
                        smtp.login(sender_email, app_pass)
                        smtp.send_message(msg)

                    self.db.mark_contacted(lead["id"], f"إيميل ({datetime.now().strftime('%m-%d %H:%M')})")
                    sent += 1
                    self.log(f"[{sent}/{max_e}] ✓ تم الإرسال إلى {lead['name']} ({to_email})")

                    if sent < max_e and sent < len(leads):
                        wt = random.uniform(d_min, d_max)
                        self.log(f"انتظار {wt:.1f} ثانية...")
                        if self.stop_event.wait(wt):
                            break
                except Exception as e:
                    self.log(f"✗ فشل مع {lead['name']}: {e}")

            self.log(f"انتهت الحملة. الإجمالي المعالج: {sent}")
            self.after(0, self._finish_campaign_ui)

        threading.Thread(target=_worker, daemon=True).start()

    def _stop_campaign(self):
        self.stop_event.set()
        self.log("تم طلب الإيقاف... جاري إنهاء الخطوة الحالية.")

    def _finish_campaign_ui(self):
        self.is_sending = False
        self.btn_run_camp.configure(state="normal")
        self.btn_stop_camp.configure(state="disabled")
        self.refresh_dashboard()
        self.refresh_leads_table()

    def log(self, text):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{ts}] {text}")

    def _poll_logs(self):
        while not self.log_queue.empty():
            m = self.log_queue.get_nowait()
            self.log_box.configure(state="normal")
            self.log_box.insert("end", m + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(150, self._poll_logs)


if __name__ == "__main__":
    app = SmartCRMApp()
    app.mainloop()
