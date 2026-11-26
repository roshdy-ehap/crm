import os
import re
import time
import queue
import shutil
import random
import smtplib
import platform
import threading
import webbrowser
import subprocess
from datetime import datetime
from urllib.parse import quote

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk

from crm_core import (
    DB_FILE,
    LINKEDIN_SESSION_DIR,
    WA_SESSION_DIR,
    STAGES,
    STAGE_COLORS,
    DatabaseManager,
    render_msg,
    format_whatsapp_phone,
    is_valid_email,
    matches_icp_and_contact,
    session_is_ready,
    export_leads_to_excel,
    import_leads_from_excel,
    send_single_smtp_email,
)
from browser_engines import (
    connect_and_verify_account,
    scrape_google_maps,
    scrape_linkedin,
    scrape_web_search,
    run_email_campaign,
    run_whatsapp_campaign,
    run_linkedin_campaign,
)

SAFETY_PRESETS = {
    "حساب جديد (Warm-up)": {
        "daily_email_cap": "15", "daily_wa_cap": "10", "daily_li_cap": "8",
        "email_min_delay": "50", "email_max_delay": "95",
        "wa_min_delay": "65", "wa_max_delay": "120",
        "li_min_delay": "60", "li_max_delay": "110",
        "batch_size": "4", "batch_break_sec": "240",
    },
    "متوازن (آمن)": {
        "daily_email_cap": "35", "daily_wa_cap": "20", "daily_li_cap": "15",
        "email_min_delay": "35", "email_max_delay": "75",
        "wa_min_delay": "50", "wa_max_delay": "95",
        "li_min_delay": "45", "li_max_delay": "85",
        "batch_size": "5", "batch_break_sec": "180",
    },
    "حساب قديم (نشط)": {
        "daily_email_cap": "60", "daily_wa_cap": "35", "daily_li_cap": "25",
        "email_min_delay": "25", "email_max_delay": "55",
        "wa_min_delay": "35", "wa_max_delay": "70",
        "li_min_delay": "35", "li_max_delay": "65",
        "batch_size": "8", "batch_break_sec": "150",
    },
}


# ======================= نافذة إضافة / تعديل عميل =======================

class LeadDialog(ctk.CTkToplevel):
    def __init__(self, parent, on_save, lead=None):
        super().__init__(parent)
        self.on_save = on_save
        self.lead = lead
        self.title("تعديل بيانات العميل" if lead else "إضافة عميل جديد")
        self.geometry("540x650")
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
        ctk.CTkButton(btns, text="حفظ البيانات", fg_color="#10B981", hover_color="#059669", command=self._submit).pack(side="right", padx=6, expand=True, fill="x")
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


# ======================= نافذة الإرسال اليدوي المتتابع (آمن 100%) =======================

class ManualOutreachQueueDialog(ctk.CTkToplevel):
    """طابور إرسال يدوي سريع لعملاء متعددين مع تنويع تلقائي و0% مخاطرة بالحظر"""

    def __init__(self, parent, leads: list[dict], db: DatabaseManager, log_fn, on_update):
        super().__init__(parent)
        self.app = parent
        self.leads = leads
        self.db = db
        self.log_fn = log_fn
        self.on_update = on_update
        self.idx = 0

        self.title("مساعد الإرسال اليدوي الآمن (0% خطر حظر)")
        self.geometry("680x580")
        self.grab_set()
        self._build_ui()
        self._load_current()

    def _build_ui(self):
        top = ctk.CTkFrame(self, corner_radius=10)
        top.pack(fill="x", padx=14, pady=(12, 8))
        self.lbl_counter = ctk.CTkLabel(top, text="", font=("Segoe UI", 15, "bold"), text_color="#38BDF8")
        self.lbl_counter.pack(side="right", padx=12, pady=10)
        self.lbl_info = ctk.CTkLabel(top, text="", font=("Segoe UI", 12))
        self.lbl_info.pack(side="left", padx=12, pady=10)

        tpl_row = ctk.CTkFrame(self, fg_color="transparent")
        tpl_row.pack(fill="x", padx=14, pady=(0, 6))
        ctk.CTkLabel(tpl_row, text="القالب:", font=("Segoe UI", 12, "bold")).pack(side="right", padx=4)
        tpls = [t["title"] for t in self.db.get_templates()]
        self.cmb_tpl = ctk.CTkComboBox(tpl_row, values=tpls, state="readonly", width=220, command=lambda _: self._render_current())
        if tpls:
            self.cmb_tpl.set(tpls[0])
        self.cmb_tpl.pack(side="right", padx=4)
        ctk.CTkButton(tpl_row, text="🎲 تنويع الصياغة (Spintax)", width=165, fg_color="#6366F1", command=self._render_current).pack(side="left", padx=4)

        self.e_subj = ctk.CTkEntry(self, justify="right", placeholder_text="عنوان الإيميل")
        self.e_subj.pack(fill="x", padx=14, pady=(0, 6))

        self.txt_msg = ctk.CTkTextbox(self, font=("Segoe UI", 12), height=220)
        self.txt_msg.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        # أزرار القنوات اليدوية
        ch_row = ctk.CTkFrame(self, fg_color="transparent")
        ch_row.pack(fill="x", padx=14, pady=(0, 8))

        self.btn_wa = ctk.CTkButton(ch_row, text="💬 فتح واتساب مانوال", fg_color="#25D366", hover_color="#16A34A", command=self._send_wa_manual)
        self.btn_wa.pack(side="right", padx=4, expand=True, fill="x")

        self.btn_gm = ctk.CTkButton(ch_row, text="✉️ فتح Gmail مانوال", fg_color="#2563EB", hover_color="#1D4ED8", command=self._send_gmail_manual)
        self.btn_gm.pack(side="right", padx=4, expand=True, fill="x")

        self.btn_li = ctk.CTkButton(ch_row, text="💼 LinkedIn + نسخ", fg_color="#0A66C2", hover_color="#004182", command=self._send_li_manual)
        self.btn_li.pack(side="right", padx=4, expand=True, fill="x")

        ctk.CTkButton(ch_row, text="📋 نسخ النص", width=95, fg_color="#475569", command=self._copy_text).pack(side="left", padx=4)

        # التنقل
        nav = ctk.CTkFrame(self, corner_radius=10)
        nav.pack(fill="x", padx=14, pady=(0, 12))
        ctk.CTkButton(nav, text="⏭️ العميل التالي", fg_color="#10B981", hover_color="#059669", command=self._next).pack(side="right", padx=8, pady=8)
        ctk.CTkButton(nav, text="⏮️ السابق", width=90, fg_color="#475569", command=self._prev).pack(side="right", padx=4, pady=8)
        ctk.CTkButton(nav, text="إغلاق", width=80, fg_color="#64748B", command=self.destroy).pack(side="left", padx=8, pady=8)

    def _current_lead(self):
        return self.leads[self.idx]

    def _load_current(self):
        lead = self._current_lead()
        self.lbl_counter.configure(text=f"عميل {self.idx + 1} من {len(self.leads)}: {lead['name']}")
        ph = format_whatsapp_phone(lead.get("phone")) or "بدون رقم"
        em = lead.get("email") if is_valid_email(lead.get("email")) else "بدون إيميل"
        self.lbl_info.configure(text=f"📱 {ph}   |   ✉️ {em}")

        self.btn_wa.configure(state="normal" if format_whatsapp_phone(lead.get("phone")) else "disabled")
        self.btn_gm.configure(state="normal" if is_valid_email(lead.get("email")) else "disabled")
        self.btn_li.configure(state="normal" if (lead.get("profile_url") or "").startswith("http") else "disabled")
        self._render_current()

    def _render_current(self):
        lead = self._current_lead()
        tpls = self.db.get_templates()
        tpl = next((t for t in tpls if t["title"] == self.cmb_tpl.get()), tpls[0] if tpls else {"subject": "", "body": ""})
        sn = self.db.get_settings().get("sender_name", "")
        self.e_subj.delete(0, "end")
        self.e_subj.insert(0, render_msg(tpl["subject"], lead, sn, apply_spintax=True))
        self.txt_msg.delete("1.0", "end")
        self.txt_msg.insert("1.0", render_msg(tpl["body"], lead, sn, apply_spintax=True))

    def _copy_text(self):
        txt = self.txt_msg.get("1.0", "end-1c").strip()
        self.clipboard_clear()
        self.clipboard_append(txt)
        self.log_fn(f"📋 تم نسخ الرسالة المخصصة لـ {self._current_lead()['name']}.")

    def _send_wa_manual(self):
        lead = self._current_lead()
        ph = format_whatsapp_phone(lead.get("phone"))
        if not ph:
            return
        msg = self.txt_msg.get("1.0", "end-1c").strip()
        webbrowser.open(f"https://wa.me/{ph}?text={quote(msg)}")
        now_t = datetime.now().strftime("%m-%d %H:%M")
        self.db.update_channel_status(lead["id"], "wa_status", f"تم يدوي ({now_t})", promote_stage=True)
        self.db.record_activity(lead["id"], "whatsapp", "sent")
        self.log_fn(f"💬 [يدوي] فتح واتساب لـ {lead['name']} (+{ph})")
        self.on_update()

    def _send_gmail_manual(self):
        lead = self._current_lead()
        em = lead.get("email", "").strip()
        if not is_valid_email(em):
            return
        subj = self.e_subj.get().strip()
        body = self.txt_msg.get("1.0", "end-1c").strip()
        webbrowser.open(f"https://mail.google.com/mail/?view=cm&fs=1&to={em}&su={quote(subj)}&body={quote(body)}")
        now_t = datetime.now().strftime("%m-%d %H:%M")
        self.db.update_channel_status(lead["id"], "email_status", f"تم يدوي ({now_t})", promote_stage=True)
        self.db.record_activity(lead["id"], "email", "sent")
        self.log_fn(f"✉️ [يدوي] فتح Gmail لـ {lead['name']} ({em})")
        self.on_update()

    def _send_li_manual(self):
        lead = self._current_lead()
        url = (lead.get("profile_url") or "").strip()
        if not url.startswith("http"):
            return
        self._copy_text()
        webbrowser.open(url)
        now_t = datetime.now().strftime("%m-%d %H:%M")
        self.db.update_channel_status(lead["id"], "li_status", f"تم يدوي ({now_t})", promote_stage=True)
        self.db.record_activity(lead["id"], "linkedin", "sent")
        self.log_fn(f"💼 [يدوي] فتح صفحة {lead['name']} وتم نسخ الرسالة في الحافظة (اضغط Ctrl+V).")
        self.on_update()

    def _next(self):
        if self.idx < len(self.leads) - 1:
            self.idx += 1
            self._load_current()
        else:
            messagebox.showinfo("انتهى الطابور", "أتممت جميع العملاء في طابور الإرسال اليدوي!", parent=self)
            self.destroy()

    def _prev(self):
        if self.idx > 0:
            self.idx -= 1
            self._load_current()


# ======================= التطبيق الرئيسي العالمي (v2.0) =======================

class SmartCRMProApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()
        self.settings = self.db.get_settings()

        ctk.set_appearance_mode(self.settings.get("appearance", "Dark"))
        ctk.set_default_color_theme("blue")

        self.title("Smart Sales CRM Pro v2.0 — Enterprise Lead Generation & Outreach Suite")
        self.geometry("1320x820")
        self.minsize(1100, 680)

        # أقفال مستقلة لكل متصفح ومحركات متزامنة مستقلة (تقدر تسحب وتبعت في نفس الوقت!)
        self.li_browser_lock = threading.Lock()
        self.wa_browser_lock = threading.Lock()

        self.scraper_running = False
        self.scraper_stop = threading.Event()

        self.email_running = False
        self.email_stop = threading.Event()

        self.wa_running = False
        self.wa_stop = threading.Event()

        self.li_running = False
        self.li_stop = threading.Event()

        self.selected_lead_ids = set()
        self.active_lead_id = None
        self.event_queue = queue.Queue()
        self._pending_table_refresh = False

        self._configure_treeview_style()
        self._build_layout()
        self.refresh_account_badges()
        self._show_page("dashboard")
        self._poll_events()
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
        # شريط الحالة الحي في الأسفل (يظهر حالة السحب والإرسال المتزامن)
        self.status_dock = ctk.CTkFrame(self, height=40, corner_radius=0, fg_color=("#E2E8F0", "#0F172A"))
        self.status_dock.pack(side="bottom", fill="x")
        self.status_dock.pack_propagate(False)

        self.lbl_dock_scraper = ctk.CTkLabel(self.status_dock, text="🔍 السحب: متوقف", font=("Segoe UI", 11, "bold"), text_color="#94A3B8")
        self.lbl_dock_scraper.pack(side="right", padx=14)
        self.btn_dock_stop_scraper = ctk.CTkButton(self.status_dock, text="إيقاف السحب", width=85, height=24, fg_color="#DC2626", state="disabled", command=self.stop_scraper)
        self.btn_dock_stop_scraper.pack(side="right", padx=4)

        self.lbl_dock_camp = ctk.CTkLabel(self.status_dock, text="🚀 الحملات: جاهز", font=("Segoe UI", 11, "bold"), text_color="#94A3B8")
        self.lbl_dock_camp.pack(side="right", padx=(24, 8))
        self.btn_dock_stop_camp = ctk.CTkButton(self.status_dock, text="إيقاف الإرسال", width=90, height=24, fg_color="#B91C1C", state="disabled", command=self.stop_all_campaigns)
        self.btn_dock_stop_camp.pack(side="right", padx=4)

        self.lbl_dock_quota = ctk.CTkLabel(self.status_dock, text="🛡️ كوتا اليوم: جاري التحميل...", font=("Segoe UI", 11), text_color="#38BDF8")
        self.lbl_dock_quota.pack(side="left", padx=14)

        # الشريط الجانبي (Sidebar)
        self.sidebar = ctk.CTkFrame(self, width=230, corner_radius=0)
        self.sidebar.pack(side="right", fill="y")
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(self.sidebar, text="Smart CRM Pro", font=("Segoe UI", 21, "bold"), text_color="#38BDF8").pack(pady=(18, 0))
        ctk.CTkLabel(self.sidebar, text="Enterprise Edition v2.0", font=("Segoe UI", 10, "bold"), text_color="#64748B").pack(pady=(0, 16))

        self.nav_buttons = {}
        nav_items = [
            ("dashboard", "📊  لوحة التحكم والإحصائيات"),
            ("scraper", "🔍  محرك سحب العملاء"),
            ("leads", "👥  العملاء والإرسال اليدوي"),
            ("templates", "📝  قوالب الرسائل (Spintax)"),
            ("campaigns", "🚀  الحملات ودرع الحظر"),
            ("accounts", "🔗  ربط الحسابات"),
        ]
        for key, text in nav_items:
            btn = ctk.CTkButton(
                self.sidebar, text=text, anchor="e", height=40,
                font=("Segoe UI", 12, "bold"), corner_radius=8,
                fg_color="transparent", hover_color=("#CBD5E1", "#334155"),
                text_color=("#0F172A", "#F8FAFC"),
                command=lambda k=key: self._show_page(k),
            )
            btn.pack(fill="x", padx=12, pady=3)
            self.nav_buttons[key] = btn

        # حالة ربط الحسابات
        badge_box = ctk.CTkFrame(self.sidebar, corner_radius=10)
        badge_box.pack(fill="x", padx=12, pady=(16, 8))
        ctk.CTkLabel(badge_box, text="حالة الاتصال:", font=("Segoe UI", 11, "bold")).pack(anchor="e", padx=10, pady=(8, 4))
        self.lbl_side_li = ctk.CTkLabel(badge_box, text="○ LinkedIn", font=("Segoe UI", 11), anchor="e")
        self.lbl_side_li.pack(fill="x", padx=10, pady=1)
        self.lbl_side_wa = ctk.CTkLabel(badge_box, text="○ WhatsApp", font=("Segoe UI", 11), anchor="e")
        self.lbl_side_wa.pack(fill="x", padx=10, pady=1)
        self.lbl_side_gm = ctk.CTkLabel(badge_box, text="○ Gmail", font=("Segoe UI", 11), anchor="e")
        self.lbl_side_gm.pack(fill="x", padx=10, pady=(1, 8))

        ctk.CTkButton(self.sidebar, text="+ إضافة عميل يدوياً", height=36, fg_color="#10B981", hover_color="#059669", font=("Segoe UI", 12, "bold"), command=self._open_add_lead).pack(fill="x", padx=12, pady=(6, 4))

        ctk.CTkLabel(self.sidebar, text="المظهر:", font=("Segoe UI", 11)).pack(side="bottom", pady=(0, 2))
        self.mode_switch = ctk.CTkSegmentedButton(self.sidebar, values=["Dark", "Light"], command=self._change_theme)
        self.mode_switch.set(self.settings.get("appearance", "Dark"))
        self.mode_switch.pack(side="bottom", padx=12, pady=(0, 12), fill="x")

        # المحتوى الرئيسي
        self.main_area = ctk.CTkFrame(self, fg_color="transparent")
        self.main_area.pack(side="left", fill="both", expand=True, padx=14, pady=10)

        self.pages = {k: ctk.CTkFrame(self.main_area, fg_color="transparent") for k in ("dashboard", "scraper", "leads", "templates", "campaigns", "accounts")}
        self._build_dashboard_page(self.pages["dashboard"])
        self._build_scraper_page(self.pages["scraper"])
        self._build_leads_page(self.pages["leads"])
        self._build_templates_page(self.pages["templates"])
        self._build_campaigns_page(self.pages["campaigns"])
        self._build_accounts_page(self.pages["accounts"])

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
        elif key == "accounts":
            self.refresh_account_badges()

    def _change_theme(self, mode):
        ctk.set_appearance_mode(mode)
        self.db.save_settings({"appearance": mode})
        self._configure_treeview_style()
        self.refresh_leads_table()

    def refresh_account_badges(self):
        s = self.db.get_settings()
        li = session_is_ready(LINKEDIN_SESSION_DIR) and s.get("li_verified") == "1"
        wa = session_is_ready(WA_SESSION_DIR) and s.get("wa_verified") == "1"
        gm = bool(s.get("sender_email") and s.get("app_password"))
        gm_v = gm and s.get("gmail_verified") == "1"

        self.lbl_side_li.configure(text="● LinkedIn: مؤكد ✓" if li else "○ LinkedIn: غير مربوط", text_color="#22C55E" if li else "#F87171")
        self.lbl_side_wa.configure(text="● WhatsApp: مؤكد ✓" if wa else "○ WhatsApp: غير مربوط", text_color="#22C55E" if wa else "#F87171")
        self.lbl_side_gm.configure(
            text="● Gmail: مؤكد ✓" if gm_v else ("● Gmail: محفوظ" if gm else "○ Gmail: غير مربوط"),
            text_color="#22C55E" if gm_v else ("#F59E0B" if gm else "#F87171"),
        )
        if hasattr(self, "lbl_acc_li_st"):
            self.lbl_acc_li_st.configure(text="متصل ومؤكد ✓" if li else "غير مربوط بعد", text_color="#22C55E" if li else "#F87171")
            self.lbl_acc_wa_st.configure(text="متصل ومؤكد ✓" if wa else "غير مربوط بعد", text_color="#22C55E" if wa else "#F87171")
            self.lbl_acc_gm_st.configure(text="متصل ومؤكد ✓" if gm_v else ("بيانات محفوظة (اختبر الاتصال)" if gm else "غير مربوط"), text_color="#22C55E" if gm_v else "#F59E0B")
        self._refresh_quota_dock()

    def _refresh_quota_dock(self):
        s = self.db.get_settings()
        t = self.db.get_today_counts()
        self.lbl_dock_quota.configure(
            text=f"🛡️ كوتا اليوم الآمنة — Gmail: {t['email']}/{s.get('daily_email_cap', 35)} | WA: {t['whatsapp']}/{s.get('daily_wa_cap', 20)} | LI: {t['linkedin']}/{s.get('daily_li_cap', 15)}"
        )

    # ======================= 1. لوحة التحكم التنفيذية =======================

    def _build_dashboard_page(self, parent):
        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(hdr, text="لوحة التحكم ومؤشرات الأداء (Executive Dashboard)", font=("Segoe UI", 21, "bold")).pack(side="right")
        ctk.CTkButton(hdr, text="تحديث", width=95, command=self.refresh_dashboard).pack(side="left")

        self.kpi_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self.kpi_frame.pack(fill="x", pady=(0, 12))
        self.kpi_labels = {}
        cards = [
            ("total", "إجمالي العملاء", "#0EA5E9"),
            ("جديد", "جدد", STAGE_COLORS["جديد"]),
            ("تم التواصل", "تم التواصل", STAGE_COLORS["تم التواصل"]),
            ("مهتم", "مهتمين", STAGE_COLORS["مهتم"]),
            ("تم البيع", "تم البيع", STAGE_COLORS["تم البيع"]),
            ("rate", "نسبة التحويل", "#8B5CF6"),
        ]
        for idx, (k, title, color) in enumerate(cards):
            c = ctk.CTkFrame(self.kpi_frame, corner_radius=12, border_width=2, border_color=color)
            c.grid(row=0, column=idx, padx=5, pady=4, sticky="nsew")
            self.kpi_frame.columnconfigure(idx, weight=1)
            ctk.CTkLabel(c, text=title, font=("Segoe UI", 11, "bold"), text_color="#94A3B8").pack(pady=(10, 2))
            v = ctk.CTkLabel(c, text="0", font=("Segoe UI", 24, "bold"), text_color=color)
            v.pack(pady=(0, 10))
            self.kpi_labels[k] = v

        # شبكة وسطى: قمع المبيعات + درع الكوتا اليومي
        mid = ctk.CTkFrame(parent, fg_color="transparent")
        mid.pack(fill="x", pady=(0, 10))
        mid.columnconfigure(0, weight=1)
        mid.columnconfigure(1, weight=1)

        # درع الحماية اليومي (يسار)
        q_card = ctk.CTkFrame(mid, corner_radius=12)
        q_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        ctk.CTkLabel(q_card, text="🛡️ استهلاك الكوتا اليومية (درع تجنب الحظر)", font=("Segoe UI", 14, "bold")).pack(anchor="e", padx=14, pady=(10, 6))
        self.quota_bars = {}
        for ch_key, label, col in [("email", "Gmail SMTP", "#EA4335"), ("whatsapp", "WhatsApp", "#25D366"), ("linkedin", "LinkedIn", "#0A66C2")]:
            row = ctk.CTkFrame(q_card, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=5)
            lbl = ctk.CTkLabel(row, text=f"{label}: 0/0", font=("Segoe UI", 11, "bold"), width=140, anchor="e")
            lbl.pack(side="right", padx=(6, 0))
            bar = ctk.CTkProgressBar(row, progress_color=col, height=12)
            bar.set(0)
            bar.pack(side="left", fill="x", expand=True, padx=6)
            self.quota_bars[ch_key] = (lbl, bar)

        # قمع المبيعات (يمين)
        f_card = ctk.CTkFrame(mid, corner_radius=12)
        f_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        ctk.CTkLabel(f_card, text="📈 توزيع مراحل المبيعات (Pipeline Funnel)", font=("Segoe UI", 14, "bold")).pack(anchor="e", padx=14, pady=(10, 6))
        self.funnel_bars = {}
        for st in ["جديد", "تم التواصل", "مهتم", "معاينة", "تم البيع"]:
            row = ctk.CTkFrame(f_card, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=3)
            lbl = ctk.CTkLabel(row, text=f"{st}: 0", font=("Segoe UI", 11, "bold"), width=110, anchor="e")
            lbl.pack(side="right", padx=(6, 0))
            bar = ctk.CTkProgressBar(row, progress_color=STAGE_COLORS[st], height=10)
            bar.set(0)
            bar.pack(side="left", fill="x", expand=True, padx=6)
            self.funnel_bars[st] = (lbl, bar)

        # متابعات اليوم
        due_box = ctk.CTkFrame(parent, corner_radius=12)
        due_box.pack(fill="both", expand=True)
        ctk.CTkLabel(due_box, text="⏰ متابعات مستحقة اليوم أو متأخرة", font=("Segoe UI", 14, "bold")).pack(anchor="e", padx=14, pady=(10, 4))
        self.due_list_frame = ctk.CTkScrollableFrame(due_box)
        self.due_list_frame.pack(fill="both", expand=True, padx=12, pady=(0, 10))

    def refresh_dashboard(self):
        total, counts, due = self.db.get_pipeline_stats()
        self.kpi_labels["total"].configure(text=str(total))
        for st in ("جديد", "تم التواصل", "مهتم", "تم البيع"):
            self.kpi_labels[st].configure(text=str(counts.get(st, 0)))
        rate = (counts.get("تم البيع", 0) / total * 100) if total else 0.0
        self.kpi_labels["rate"].configure(text=f"{rate:.1f}%")

        for st, (lbl, bar) in self.funnel_bars.items():
            c = counts.get(st, 0)
            pct = (c / total) if total else 0.0
            lbl.configure(text=f"{st}: {c} ({pct*100:.0f}%)")
            bar.set(pct)

        s = self.db.get_settings()
        today_c = self.db.get_today_counts()
        caps = {"email": int(s.get("daily_email_cap") or 35), "whatsapp": int(s.get("daily_wa_cap") or 20), "linkedin": int(s.get("daily_li_cap") or 15)}
        labels_map = {"email": "Gmail", "whatsapp": "WhatsApp", "linkedin": "LinkedIn"}
        for k, (lbl, bar) in self.quota_bars.items():
            used, cap = today_c[k], max(1, caps[k])
            lbl.configure(text=f"{labels_map[k]}: {used} / {cap}")
            bar.set(min(1.0, used / cap))
        self._refresh_quota_dock()

        for w in self.due_list_frame.winfo_children():
            w.destroy()
        if not due:
            ctk.CTkLabel(self.due_list_frame, text="لا توجد متابعات متأخرة حالياً.", text_color="#94A3B8").pack(pady=24)
            return
        for lead in due:
            r = ctk.CTkFrame(self.due_list_frame, corner_radius=8)
            r.pack(fill="x", pady=3, padx=4)
            ctk.CTkLabel(r, text=f"{lead['name']}  |  {lead['stage']}  |  {lead['next_followup']}  |  {lead.get('phone') or ''}", font=("Segoe UI", 12, "bold")).pack(side="right", padx=12, pady=8)
            ctk.CTkButton(r, text="💬 واتساب مانوال", width=110, fg_color="#25D366", hover_color="#16A34A", command=lambda l=lead: self._quick_manual_wa(l)).pack(side="left", padx=4, pady=5)
            ctk.CTkButton(r, text="تعديل", width=70, command=lambda l=lead: self._open_edit_lead(l)).pack(side="left", padx=4, pady=5)

    # ======================= 2. محرك سحب العملاء (Scraping) =======================

    def _build_scraper_page(self, parent):
        s = self.settings
        hdr = ctk.CTkFrame(parent, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(hdr, text="🔍 محرك سحب العملاء الذكي (يعمل في الخلفية بالتوازي مع الإرسال)", font=("Segoe UI", 18, "bold")).pack(side="right")

        src_box = ctk.CTkFrame(parent, corner_radius=12)
        src_box.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(src_box, text="1. مصادر السحب:", font=("Segoe UI", 12, "bold")).pack(side="right", padx=12, pady=8)
        self.var_src_li = tk.BooleanVar(value=True)
        self.var_src_gm = tk.BooleanVar(value=True)
        self.var_src_web = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(src_box, text="LinkedIn Companies", variable=self.var_src_li).pack(side="right", padx=10, pady=8)
        ctk.CTkCheckBox(src_box, text="Google Maps", variable=self.var_src_gm).pack(side="right", padx=10, pady=8)
        ctk.CTkCheckBox(src_box, text="بحث الويب (Web)", variable=self.var_src_web).pack(side="right", padx=10, pady=8)

        cfg_box = ctk.CTkFrame(parent, corner_radius=12)
        cfg_box.pack(fill="x", pady=(0, 6))

        self.e_keywords = ctk.CTkEntry(cfg_box, justify="right", placeholder_text="marketing agency, software house")
        self.e_keywords.insert(0, s.get("keywords", "digital marketing agency, software agency"))
        self.e_location = ctk.CTkEntry(cfg_box, justify="right", placeholder_text="Mansoura / Cairo / Riyadh")
        self.e_location.insert(0, s.get("location", "Mansoura"))
        self.e_target = ctk.CTkEntry(cfg_box, width=85)
        self.e_target.insert(0, s.get("target_count", "50"))

        self.cmb_contact_req = ctk.CTkComboBox(
            cfg_box, state="readonly", width=210,
            values=["الكل (أي بيانات متاحة)", "رقم أو إيميل على الأقل", "أرقام واتساب فقط", "إيميلات فقط", "الاثنين معاً (واتساب + إيميل)"],
        )
        self.cmb_contact_req.set(s.get("contact_req", "الكل (أي بيانات متاحة)"))

        self.e_icp_ind = ctk.CTkEntry(cfg_box, justify="right", placeholder_text="اختياري: marketing, tech...")
        self.e_icp_ind.insert(0, s.get("icp_industry", ""))
        self.cmb_icp_size = ctk.CTkComboBox(cfg_box, state="readonly", width=150, values=["الكل", "صغيرة (1-10)", "متوسطة (11-50)", "كبيرة (51-200)", "مؤسسات (+200)"])
        self.cmb_icp_size.set(s.get("icp_size", "الكل"))
        self.e_min_rating = ctk.CTkEntry(cfg_box, width=70)
        self.e_min_rating.insert(0, s.get("min_rating", "0"))
        self.var_must_web = tk.BooleanVar(value=False)
        self.var_enrich_web = tk.BooleanVar(value=True)

        ctk.CTkLabel(cfg_box, text="الكلمات المفتاحية (بفاصلة):").grid(row=0, column=3, sticky="e", padx=10, pady=5)
        self.e_keywords.grid(row=0, column=0, columnspan=3, sticky="ew", padx=10, pady=5)

        ctk.CTkLabel(cfg_box, text="اللوكيشن / المدينة:").grid(row=1, column=3, sticky="e", padx=10, pady=4)
        self.e_location.grid(row=1, column=2, sticky="ew", padx=10, pady=4)
        ctk.CTkLabel(cfg_box, text="الكمية المطلوبة:").grid(row=1, column=1, sticky="e", padx=10, pady=4)
        self.e_target.grid(row=1, column=0, sticky="e", padx=10, pady=4)

        ctk.CTkLabel(cfg_box, text="شرط بيانات التواصل:").grid(row=2, column=3, sticky="e", padx=10, pady=4)
        self.cmb_contact_req.grid(row=2, column=2, sticky="ew", padx=10, pady=4)
        ctk.CTkLabel(cfg_box, text="فلتر التخصص (اختياري):").grid(row=2, column=1, sticky="e", padx=10, pady=4)
        self.e_icp_ind.grid(row=2, column=0, sticky="ew", padx=10, pady=4)

        ctk.CTkLabel(cfg_box, text="حجم الشركة (LinkedIn):").grid(row=3, column=3, sticky="e", padx=10, pady=4)
        self.cmb_icp_size.grid(row=3, column=2, sticky="e", padx=10, pady=4)
        ctk.CTkLabel(cfg_box, text="أقل تقييم (Maps):").grid(row=3, column=1, sticky="e", padx=10, pady=4)
        self.e_min_rating.grid(row=3, column=0, sticky="e", padx=10, pady=4)

        chk_row = ctk.CTkFrame(cfg_box, fg_color="transparent")
        chk_row.grid(row=4, column=0, columnspan=4, sticky="ew", padx=10, pady=(4, 10))
        ctk.CTkCheckBox(chk_row, text="فحص موقع كل شركة (الرئيسية + Contact) لسحب الإيميل والواتساب", variable=self.var_enrich_web).pack(side="right", padx=6)
        ctk.CTkCheckBox(chk_row, text="يشترط وجود موقع إلكتروني", variable=self.var_must_web).pack(side="right", padx=10)

        self.btn_start_scrape = ctk.CTkButton(chk_row, text="🚀 بدء السحب الآن", height=36, font=("Segoe UI", 12, "bold"), fg_color="#0A66C2", hover_color="#004182", command=self._start_scraping)
        self.btn_start_scrape.pack(side="left", padx=4)
        self.btn_stop_scrape = ctk.CTkButton(chk_row, text="⛔ إيقاف السحب", height=36, width=110, fg_color="#DC2626", state="disabled", command=self.stop_scraper)
        self.btn_stop_scrape.pack(side="left", padx=4)

        cfg_box.columnconfigure(0, weight=1)
        cfg_box.columnconfigure(2, weight=1)

        # شريط التقدم + Live Log
        log_card = ctk.CTkFrame(parent, corner_radius=12)
        log_card.pack(fill="both", expand=True)

        p_row = ctk.CTkFrame(log_card, fg_color="transparent")
        p_row.pack(fill="x", padx=12, pady=(8, 4))
        ctk.CTkLabel(p_row, text="📋 سجل السحب والعمليات المباشر (Live Log)", font=("Segoe UI", 12, "bold")).pack(side="right")
        self.scrape_progress = ctk.CTkProgressBar(p_row, width=220, height=12)
        self.scrape_progress.set(0)
        self.scrape_progress.pack(side="left", padx=8)
        ctk.CTkButton(p_row, text="مسح", width=65, fg_color="#475569", command=self._clear_logs).pack(side="left")

        self.scrape_log = ctk.CTkTextbox(log_card, font=("Consolas", 11))
        self.scrape_log.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        self.scrape_log.configure(state="disabled")

    # ======================= 3. قاعدة العملاء + لوحة الإرسال اليدوي =======================

    def _build_leads_page(self, parent):
        # شريط البحث والفلترة
        top_bar = ctk.CTkFrame(parent, corner_radius=10)
        top_bar.pack(fill="x", pady=(0, 6))

        self.var_search = tk.StringVar()
        self.var_search.trace_add("write", lambda *_: self.refresh_leads_table())
        ctk.CTkEntry(top_bar, textvariable=self.var_search, placeholder_text="🔍 بحث فوري...", width=200, justify="right").pack(side="right", padx=6, pady=7)

        self.cmb_filter_stage = ctk.CTkComboBox(top_bar, values=["الكل"] + STAGES, state="readonly", width=115, command=lambda _: self.refresh_leads_table())
        self.cmb_filter_stage.set("الكل")
        self.cmb_filter_stage.pack(side="right", padx=4, pady=7)

        self.cmb_quick_filter = ctk.CTkComboBox(
            top_bar, state="readonly", width=165,
            values=["الكل", "بها إيميل ولم تُرسل", "بها واتساب ولم تُرسل", "عملاء LinkedIn", "المضاف اليوم"],
            command=lambda _: self.refresh_leads_table(),
        )
        self.cmb_quick_filter.set("الكل")
        self.cmb_quick_filter.pack(side="right", padx=4, pady=7)

        ctk.CTkButton(top_bar, text="📥 استيراد Excel", width=105, fg_color="#475569", command=self._import_excel).pack(side="left", padx=4, pady=7)
        ctk.CTkButton(top_bar, text="📤 تصدير Excel", width=105, fg_color="#0284C7", command=self._export_excel).pack(side="left", padx=4, pady=7)
        ctk.CTkButton(top_bar, text="تحديث", width=65, command=self.refresh_leads_table).pack(side="left", padx=4, pady=7)

        # شريط العمليات الجماعية
        pipe_bar = ctk.CTkFrame(parent, fg_color="transparent")
        pipe_bar.pack(fill="x", pady=(0, 6))

        ctk.CTkButton(pipe_bar, text="تحديد الظاهر", width=85, fg_color="#334155", command=self._select_all_visible).pack(side="right", padx=2)
        ctk.CTkButton(pipe_bar, text="إلغاء التحديد", width=85, fg_color="#334155", command=self._clear_selection).pack(side="right", padx=2)

        self.cmb_bulk_stage = ctk.CTkComboBox(pipe_bar, values=STAGES, state="readonly", width=115)
        self.cmb_bulk_stage.set("مهتم")
        self.cmb_bulk_stage.pack(side="right", padx=(10, 3))
        ctk.CTkButton(pipe_bar, text="نقل المرحلة", width=85, command=self._apply_bulk_stage).pack(side="right", padx=2)

        self.btn_manual_queue = ctk.CTkButton(
            pipe_bar, text="🧑‍💻 إرسال مانوال بالتتابع (آمن 100%)", fg_color="#10B981", hover_color="#059669",
            font=("Segoe UI", 11, "bold"), command=self._open_manual_queue,
        )
        self.btn_manual_queue.pack(side="right", padx=(12, 3))

        self.btn_auto_sel = ctk.CTkButton(
            pipe_bar, text="🚀 حملة أوتوماتيك للمحدد", fg_color="#2563EB", hover_color="#1D4ED8",
            font=("Segoe UI", 11, "bold"), command=self._launch_campaign_for_selected,
        )
        self.btn_auto_sel.pack(side="right", padx=3)

        ctk.CTkButton(pipe_bar, text="حذف", width=55, fg_color="#DC2626", command=self._delete_selected).pack(side="left", padx=3)
        self.lbl_selected_count = ctk.CTkLabel(pipe_bar, text="المحدد: 0", font=("Segoe UI", 12, "bold"), text_color="#38BDF8")
        self.lbl_selected_count.pack(side="left", padx=8)

        # تقسيم الصفحة: جدول العملاء (يسار) + بطاقة المعاينة والإرسال اليدوي (يمين)
        split = ctk.CTkFrame(parent, fg_color="transparent")
        split.pack(fill="both", expand=True)
        split.columnconfigure(0, weight=3)
        split.columnconfigure(1, weight=1)
        split.rowconfigure(0, weight=1)

        # الجدول
        tbox = ctk.CTkFrame(split, corner_radius=10)
        tbox.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        cols = ("sel", "id", "src", "name", "phone", "email", "stage", "em_st", "wa_st", "li_st")
        self.tree = ttk.Treeview(tbox, columns=cols, show="headings", style="CRM.Treeview", selectmode="browse")
        specs = [
            ("sel", "☑", 38, "center"), ("id", "#", 34, "center"), ("src", "المصدر", 82, "center"),
            ("name", "الاسم / الشركة", 150, "e"), ("phone", "واتساب", 108, "center"),
            ("email", "الإيميل", 135, "w"), ("stage", "المرحلة", 82, "center"),
            ("em_st", "إيميل", 86, "center"), ("wa_st", "واتساب", 86, "center"), ("li_st", "LinkedIn", 82, "center"),
        ]
        for cid, h, w, anc in specs:
            self.tree.heading(cid, text=h)
            self.tree.column(cid, width=w, minwidth=34, anchor=anc)

        vsb = ttk.Scrollbar(tbox, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tbox, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(6, 0), pady=6)
        vsb.grid(row=0, column=1, sticky="ns", pady=6)
        hsb.grid(row=1, column=0, sticky="ew", padx=6)
        tbox.rowconfigure(0, weight=1)
        tbox.columnconfigure(0, weight=1)

        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # بطاقة الإرسال اليدوي السريع للعميل المختار (Side Inspector Drawer)
        drawer = ctk.CTkFrame(split, corner_radius=12, width=320)
        drawer.grid(row=0, column=1, sticky="nsew")
        drawer.pack_propagate(False)

        ctk.CTkLabel(drawer, text="⚡ إرسال مانوال وتفاصيل العميل", font=("Segoe UI", 13, "bold"), text_color="#38BDF8").pack(pady=(10, 4))
        self.lbl_insp_name = ctk.CTkLabel(drawer, text="اضغط على أي عميل من الجدول", font=("Segoe UI", 13, "bold"), wraplength=290)
        self.lbl_insp_name.pack(padx=10, pady=(0, 2))
        self.lbl_insp_meta = ctk.CTkLabel(drawer, text="—", font=("Segoe UI", 11), text_color="#94A3B8", wraplength=290)
        self.lbl_insp_meta.pack(padx=10, pady=(0, 8))

        tpl_r = ctk.CTkFrame(drawer, fg_color="transparent")
        tpl_r.pack(fill="x", padx=10, pady=(0, 4))
        self.cmb_insp_tpl = ctk.CTkComboBox(tpl_r, values=[], state="readonly", command=lambda _: self._render_inspector_msg())
        self.cmb_insp_tpl.pack(side="right", fill="x", expand=True)
        ctk.CTkButton(tpl_r, text="🎲", width=34, fg_color="#6366F1", command=self._render_inspector_msg).pack(side="left", padx=(4, 0))

        self.e_insp_subj = ctk.CTkEntry(drawer, justify="right", placeholder_text="عنوان الإيميل")
        self.e_insp_subj.pack(fill="x", padx=10, pady=(2, 4))

        self.txt_insp_msg = ctk.CTkTextbox(drawer, font=("Segoe UI", 11), height=150)
        self.txt_insp_msg.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        # أزرار الإرسال اليدوي الفردي
        ctk.CTkButton(drawer, text="💬 إرسال واتساب مانوال", fg_color="#25D366", hover_color="#16A34A", command=self._insp_send_wa).pack(fill="x", padx=10, pady=2)
        ctk.CTkButton(drawer, text="✉️ فتح Gmail بالمتصفح (مانوال)", fg_color="#2563EB", hover_color="#1D4ED8", command=self._insp_send_gmail_web).pack(fill="x", padx=10, pady=2)
        ctk.CTkButton(drawer, text="📨 إرسال إيميل فردي فوري (SMTP)", fg_color="#EA4335", hover_color="#DC2626", command=self._insp_send_smtp_one).pack(fill="x", padx=10, pady=2)
        ctk.CTkButton(drawer, text="💼 فتح LinkedIn + نسخ الرسالة", fg_color="#0A66C2", hover_color="#004182", command=self._insp_open_linkedin).pack(fill="x", padx=10, pady=2)

        sub_btns = ctk.CTkFrame(drawer, fg_color="transparent")
        sub_btns.pack(fill="x", padx=10, pady=(4, 10))
        ctk.CTkButton(sub_btns, text="📋 نسخ", width=75, fg_color="#475569", command=self._insp_copy).pack(side="right", padx=2)
        ctk.CTkButton(sub_btns, text="✏️ تعديل", width=75, fg_color="#334155", command=self._insp_edit).pack(side="right", padx=2)
        ctk.CTkButton(sub_btns, text="🚫 عدم إزعاج", width=95, fg_color="#991B1B", command=self._insp_mark_dnc).pack(side="left", padx=2)

    def refresh_leads_table(self):
        tpls = [t["title"] for t in self.db.get_templates()]
        if tpls:
            cur = self.cmb_insp_tpl.get()
            self.cmb_insp_tpl.configure(values=tpls)
            if cur not in tpls:
                self.cmb_insp_tpl.set(tpls[0])

        leads = self.db.get_leads(self.var_search.get(), self.cmb_filter_stage.get(), self.cmb_quick_filter.get())
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
                    lead.get("phone") or "—", lead.get("email") or "—", lead.get("stage") or "جديد",
                    lead.get("email_status") or "—", lead.get("wa_status") or "—", lead.get("li_status") or "—",
                ),
                tags=("checked",) if is_sel else (),
            )
        self.lbl_selected_count.configure(text=f"المحدد: {len(self.selected_lead_ids)}")
        if self.active_lead_id:
            self._load_inspector(self.active_lead_id)

    def _on_tree_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not iid:
            return
        lid = int(iid)
        self.active_lead_id = lid
        self._load_inspector(lid)

        # الضغط على عمود العلامة (#1) يبدّل التحديد دائماً، وكذلك الضغط العادي لو حابب تحديد سريع
        if col == "#1":
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

    def _load_inspector(self, lead_id):
        lead = self.db.get_lead_by_id(lead_id)
        if not lead:
            return
        self.lbl_insp_name.configure(text=lead["name"])
        ph = format_whatsapp_phone(lead.get("phone")) or "بدون واتساب"
        em = lead.get("email") if is_valid_email(lead.get("email")) else "بدون إيميل"
        self.lbl_insp_meta.configure(text=f"📱 {ph}\n✉️ {em}\n📍 {lead.get('location') or '—'} | {lead.get('stage')}")
        self._render_inspector_msg()

    def _render_inspector_msg(self):
        if not self.active_lead_id:
            return
        lead = self.db.get_lead_by_id(self.active_lead_id)
        if not lead:
            return
        tpls = self.db.get_templates()
        tpl = next((t for t in tpls if t["title"] == self.cmb_insp_tpl.get()), tpls[0] if tpls else {"subject": "", "body": ""})
        sn = self.db.get_settings().get("sender_name", "")
        self.e_insp_subj.delete(0, "end")
        self.e_insp_subj.insert(0, render_msg(tpl["subject"], lead, sn, apply_spintax=True))
        self.txt_insp_msg.delete("1.0", "end")
        self.txt_insp_msg.insert("1.0", render_msg(tpl["body"], lead, sn, apply_spintax=True))

    def _insp_copy(self):
        txt = self.txt_insp_msg.get("1.0", "end-1c").strip()
        if txt:
            self.clipboard_clear()
            self.clipboard_append(txt)
            self.log("📋 تم نسخ نص الرسالة إلى الحافظة.")

    def _insp_edit(self):
        if self.active_lead_id:
            lead = self.db.get_lead_by_id(self.active_lead_id)
            if lead:
                self._open_edit_lead(lead)

    def _insp_mark_dnc(self):
        if self.active_lead_id:
            self.db.update_stage(self.active_lead_id, "عدم الإزعاج")
            self.refresh_leads_table()
            self.log("🚫 تم نقل العميل لقائمة عدم الإزعاج (لن يتم مراسلته تلقائياً).")

    def _quick_manual_wa(self, lead):
        ph = format_whatsapp_phone(lead.get("phone"))
        if not ph:
            messagebox.showwarning("بدون رقم", "لا يوجد رقم موبايل صالح لواتساب لهذا العميل.")
            return
        tpls = self.db.get_templates()
        body = render_msg(tpls[0]["body"], lead, self.db.get_settings().get("sender_name", "")) if tpls else f"السلام عليكم {lead['name']}"
        webbrowser.open(f"https://wa.me/{ph}?text={quote(body)}")
        self.db.update_channel_status(lead["id"], "wa_status", f"تم يدوي ({datetime.now().strftime('%m-%d %H:%M')})", promote_stage=True)
        self.db.record_activity(lead["id"], "whatsapp", "sent")
        self.refresh_leads_table()
        self.refresh_dashboard()

    def _insp_send_wa(self):
        if not self.active_lead_id:
            return
        lead = self.db.get_lead_by_id(self.active_lead_id)
        ph = format_whatsapp_phone(lead.get("phone"))
        if not ph:
            messagebox.showwarning("لا يوجد رقم", "هذا العميل ليس لديه رقم موبايل صالح لواتساب.")
            return
        msg = self.txt_insp_msg.get("1.0", "end-1c").strip()
        webbrowser.open(f"https://wa.me/{ph}?text={quote(msg)}")
        self.db.update_channel_status(lead["id"], "wa_status", f"تم يدوي ({datetime.now().strftime('%m-%d %H:%M')})", promote_stage=True)
        self.db.record_activity(lead["id"], "whatsapp", "sent")
        self.log(f"💬 [مانوال] فتح واتساب لـ {lead['name']} (+{ph})")
        self.refresh_leads_table()
        self._refresh_quota_dock()

    def _insp_send_gmail_web(self):
        if not self.active_lead_id:
            return
        lead = self.db.get_lead_by_id(self.active_lead_id)
        em = (lead.get("email") or "").strip()
        if not is_valid_email(em):
            messagebox.showwarning("لا يوجد إيميل", "هذا العميل ليس لديه بريد إلكتروني صالح.")
            return
        subj = self.e_insp_subj.get().strip()
        body = self.txt_insp_msg.get("1.0", "end-1c").strip()
        webbrowser.open(f"https://mail.google.com/mail/?view=cm&fs=1&to={em}&su={quote(subj)}&body={quote(body)}")
        self.db.update_channel_status(lead["id"], "email_status", f"تم يدوي ({datetime.now().strftime('%m-%d %H:%M')})", promote_stage=True)
        self.db.record_activity(lead["id"], "email", "sent")
        self.log(f"✉️ [مانوال] فتح نافذة Gmail لـ {lead['name']} ({em})")
        self.refresh_leads_table()
        self._refresh_quota_dock()

    def _insp_send_smtp_one(self):
        if not self.active_lead_id:
            return
        lead = self.db.get_lead_by_id(self.active_lead_id)
        em = (lead.get("email") or "").strip()
        if not is_valid_email(em):
            messagebox.showwarning("لا يوجد إيميل", "هذا العميل ليس لديه بريد إلكتروني صالح.")
            return
        s = self.db.get_settings()
        if not s.get("sender_email") or not s.get("app_password"):
            messagebox.showwarning("ربط Gmail", "اربط حساب Gmail أولاً من صفحة (ربط الحسابات).")
            return
        subj = self.e_insp_subj.get().strip()
        body = self.txt_insp_msg.get("1.0", "end-1c").strip()

        def _w():
            try:
                send_single_smtp_email(s["sender_email"], s["app_password"], s.get("sender_name", ""), em, subj, body, s.get("attachment_path", ""))
                self.db.update_channel_status(lead["id"], "email_status", f"تم ({datetime.now().strftime('%m-%d %H:%M')})", promote_stage=True)
                self.db.record_activity(lead["id"], "email", "sent")
                self.log(f"✅ [إرسال فردي SMTP] تم الإرسال إلى {lead['name']} ({em})")
                self.event_queue.put(("refresh_views", None))
            except Exception as e:
                self.log(f"❌ [إرسال فردي SMTP] فشل مع {lead['name']}: {e}")

        threading.Thread(target=_w, daemon=True).start()

    def _insp_open_linkedin(self):
        if not self.active_lead_id:
            return
        lead = self.db.get_lead_by_id(self.active_lead_id)
        url = (lead.get("profile_url") or "").strip()
        if not url.startswith("http"):
            messagebox.showwarning("لا يوجد رابط", "لا يوجد رابط حساب محفوظ لهذا العميل.")
            return
        self._insp_copy()
        webbrowser.open(url)
        self.db.update_channel_status(lead["id"], "li_status", f"تم يدوي ({datetime.now().strftime('%m-%d %H:%M')})", promote_stage=True)
        self.db.record_activity(lead["id"], "linkedin", "sent")
        self.log(f"💼 [مانوال] فتح رابط {lead['name']} وتم نسخ الرسالة (Ctrl+V).")
        self.refresh_leads_table()

    def _open_manual_queue(self):
        ids = list(self.selected_lead_ids)
        if not ids:
            if self.active_lead_id:
                ids = [self.active_lead_id]
            else:
                messagebox.showwarning("اختر عملاء", "حدد عملاء من عمود العلامة ☑ أولاً (أو اضغط تحديد الظاهر).")
                return
        leads = [self.db.get_lead_by_id(i) for i in ids]
        leads = [l for l in leads if l and l.get("stage") != "عدم الإزعاج"]
        if not leads:
            return
        ManualOutreachQueueDialog(self, leads, self.db, self.log, lambda: (self.refresh_leads_table(), self._refresh_quota_dock()))

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
        p = filedialog.askopenfilename(title="استيراد Excel", filetypes=[("Excel", "*.xlsx")])
        if not p:
            return
        try:
            n = import_leads_from_excel(self.db, p)
            self.refresh_leads_table()
            messagebox.showinfo("تم الاستيراد", f"تمت إضافة {n} عميل جديد.")
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def _export_excel(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile=f"crm_leads_{datetime.now().strftime('%Y%m%d')}.xlsx")
        if not p:
            return
        leads = self.db.get_leads(self.var_search.get(), self.cmb_filter_stage.get(), self.cmb_quick_filter.get())
        if not leads:
            return
        tpls = self.db.get_templates()
        tpl = tpls[0] if tpls else {"subject": "تعاون مع {name}", "body": "السلام عليكم {name}"}
        try:
            export_leads_to_excel(leads, p, tpl["subject"], tpl["body"], self.db.get_settings().get("sender_name", ""))
            if messagebox.askyesno("تم التصدير", f"تم تصدير {len(leads)} عميل بنجاح.\nفتح الملف الآن؟"):
                if platform.system() == "Windows":
                    os.startfile(p)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", p])
                else:
                    subprocess.Popen(["xdg-open", p])
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    # ======================= 4. قوالب الرسائل (مع Spintax) =======================

    def _build_templates_page(self, parent):
        ctk.CTkLabel(parent, text="📝 إدارة قوالب الرسائل (يدعم التنويع التلقائي Spintax لتجنب التكرار والحظر)", font=("Segoe UI", 18, "bold")).pack(anchor="e", pady=(0, 8))
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

        ctk.CTkLabel(ed, text="عنوان الإيميل (يدعم Spintax مثل {تعاون|شراكة} مع {name}):", font=("Segoe UI", 12, "bold")).pack(anchor="e", padx=12, pady=(4, 2))
        self.e_tpl_subject = ctk.CTkEntry(ed, justify="right")
        self.e_tpl_subject.pack(fill="x", padx=12, pady=(0, 6))

        ctk.CTkLabel(ed, text="المتغيرات: {name} {category} {location} {budget} {sender_name}  |  التنويع: {السلام عليكم|أهلاً بكم|تحياتي}", text_color="#38BDF8", font=("Segoe UI", 11)).pack(anchor="e", padx=12, pady=(4, 2))
        self.txt_tpl_body = ctk.CTkTextbox(ed, font=("Segoe UI", 12))
        self.txt_tpl_body.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        btns = ctk.CTkFrame(ed, fg_color="transparent")
        btns.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(btns, text="💾 حفظ القالب", fg_color="#10B981", command=self._save_template).pack(side="right", padx=4)
        ctk.CTkButton(btns, text="🎲 معاينة عشوائية", fg_color="#6366F1", command=self._preview_template).pack(side="right", padx=4)
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
            return
        self.db.save_template(title, self.e_tpl_subject.get().strip(), body, "الكل")
        self.refresh_templates_ui()
        messagebox.showinfo("تم الحفظ", f"تم حفظ القالب '{title}'.")

    def _del_tpl(self):
        title = self.e_tpl_title.get().strip()
        if title and len(self.db.get_templates()) > 1:
            self.db.delete_template(title)
            self._clear_tpl()
            self.refresh_templates_ui()

    def _preview_template(self):
        sample = {"name": "وكالة الإبداع", "category": "التسويق الرقمي", "location": "المنصورة", "budget": "50,000 ج.م"}
        sn = self.db.get_settings().get("sender_name", "فريق المبيعات")
        subj = render_msg(self.e_tpl_subject.get().strip(), sample, sn, apply_spintax=True)
        body = render_msg(self.txt_tpl_body.get("1.0", "end-1c"), sample, sn, apply_spintax=True)
        messagebox.showinfo("معاينة (اضغط تاني لتشوف تنويع مختلف)", f"العنوان:\n{subj}\n\nالرسالة:\n{body}")

    # ======================= 5. الحملات ودرع الحماية من الحظر =======================

    def _build_campaigns_page(self, parent):
        s = self.settings

        # درع الحماية (Anti-Ban Guardrails)
        shield = ctk.CTkFrame(parent, corner_radius=12)
        shield.pack(fill="x", pady=(0, 8))

        p_hdr = ctk.CTkFrame(shield, fg_color="transparent")
        p_hdr.pack(fill="x", padx=12, pady=(10, 6))
        ctk.CTkLabel(p_hdr, text="🛡️ درع الحماية من الحظر (Anti-Ban Limits & Cooldowns)", font=("Segoe UI", 14, "bold"), text_color="#38BDF8").pack(side="right")

        self.cmb_preset = ctk.CTkComboBox(p_hdr, values=list(SAFETY_PRESETS.keys()), state="readonly", width=180, command=self._apply_safety_preset)
        self.cmb_preset.set(s.get("safety_profile", "متوازن (آمن)"))
        self.cmb_preset.pack(side="left", padx=4)
        ctk.CTkLabel(p_hdr, text="مستوى الأمان:").pack(side="left", padx=4)

        g = ctk.CTkFrame(shield, fg_color="transparent")
        g.pack(fill="x", padx=12, pady=(0, 10))

        def mk_entry(parent_f, key, default_v, width=55):
            e = ctk.CTkEntry(parent_f, width=width)
            e.insert(0, s.get(key, default_v))
            return e

        self.e_cap_em = mk_entry(g, "daily_email_cap", "35")
        self.e_cap_wa = mk_entry(g, "daily_wa_cap", "20")
        self.e_cap_li = mk_entry(g, "daily_li_cap", "15")
        self.e_batch_sz = mk_entry(g, "batch_size", "5")
        self.e_batch_brk = mk_entry(g, "batch_break_sec", "180")
        self.e_wa_min = mk_entry(g, "wa_min_delay", "50")
        self.e_wa_max = mk_entry(g, "wa_max_delay", "95")

        ctk.CTkLabel(g, text="حد يومي Gmail:").pack(side="right", padx=3)
        self.e_cap_em.pack(side="right", padx=(0, 10))
        ctk.CTkLabel(g, text="حد يومي WA:").pack(side="right", padx=3)
        self.e_cap_wa.pack(side="right", padx=(0, 10))
        ctk.CTkLabel(g, text="حد يومي LI:").pack(side="right", padx=3)
        self.e_cap_li.pack(side="right", padx=(0, 10))
        ctk.CTkLabel(g, text="استراحة كل (رسالة):").pack(side="right", padx=3)
        self.e_batch_sz.pack(side="right", padx=(0, 8))
        ctk.CTkLabel(g, text="مدة الاستراحة (ث):").pack(side="right", padx=3)
        self.e_batch_brk.pack(side="right", padx=(0, 10))
        ctk.CTkLabel(g, text="تأخير WA (ث):").pack(side="right", padx=3)
        self.e_wa_min.pack(side="right", padx=2)
        self.e_wa_max.pack(side="right", padx=(2, 8))

        ctk.CTkButton(g, text="💾 حفظ", width=70, fg_color="#10B981", command=self._save_shield_settings).pack(side="left")

        # إطلاق الحملة
        camp_box = ctk.CTkFrame(parent, corner_radius=12)
        camp_box.pack(fill="both", expand=True)

        row1 = ctk.CTkFrame(camp_box, fg_color="transparent")
        row1.pack(fill="x", padx=12, pady=(10, 6))
        self.var_ch_gmail = tk.BooleanVar(value=True)
        self.var_ch_wa = tk.BooleanVar(value=True)
        self.var_ch_li = tk.BooleanVar(value=False)
        self.var_dry_run = tk.BooleanVar(value=True)

        ctk.CTkLabel(row1, text="قنوات الإرسال:", font=("Segoe UI", 12, "bold")).pack(side="right", padx=4)
        ctk.CTkCheckBox(row1, text="Gmail SMTP", variable=self.var_ch_gmail).pack(side="right", padx=8)
        ctk.CTkCheckBox(row1, text="WhatsApp Web", variable=self.var_ch_wa).pack(side="right", padx=8)
        ctk.CTkCheckBox(row1, text="LinkedIn", variable=self.var_ch_li).pack(side="right", padx=8)
        ctk.CTkCheckBox(row1, text="وضع التجربة (DRY RUN)", variable=self.var_dry_run).pack(side="left", padx=8)

        row2 = ctk.CTkFrame(camp_box, fg_color="transparent")
        row2.pack(fill="x", padx=12, pady=(0, 8))
        ctk.CTkLabel(row2, text="القالب:").pack(side="right", padx=4)
        self.cmb_camp_tpl = ctk.CTkComboBox(row2, values=[], state="readonly", width=200)
        self.cmb_camp_tpl.pack(side="right", padx=4)

        ctk.CTkLabel(row2, text="المرحلة المستهدفة:").pack(side="right", padx=(10, 4))
        self.cmb_camp_stage = ctk.CTkComboBox(row2, values=["الكل"] + STAGES, state="readonly", width=120)
        self.cmb_camp_stage.set("جديد")
        self.cmb_camp_stage.pack(side="right", padx=4)

        self.btn_run_camp = ctk.CTkButton(row2, text="🚀 إطلاق الحملة (بالتوازي مع السحب)", fg_color="#16A34A", hover_color="#15803D", font=("Segoe UI", 12, "bold"), command=lambda: self._trigger_campaign(None))
        self.btn_run_camp.pack(side="left", padx=4)
        self.btn_stop_camp = ctk.CTkButton(row2, text="⛔ إيقاف الحملات", width=115, fg_color="#DC2626", state="disabled", command=self.stop_all_campaigns)
        self.btn_stop_camp.pack(side="left", padx=4)

        self.camp_progress = ctk.CTkProgressBar(camp_box, height=10)
        self.camp_progress.set(0)
        self.camp_progress.pack(fill="x", padx=12, pady=(0, 6))

        self.camp_log = ctk.CTkTextbox(camp_box, font=("Consolas", 11))
        self.camp_log.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.camp_log.configure(state="disabled")

    def _apply_safety_preset(self, preset_name):
        p = SAFETY_PRESETS.get(preset_name)
        if not p:
            return
        for entry, k in [
            (self.e_cap_em, "daily_email_cap"), (self.e_cap_wa, "daily_wa_cap"), (self.e_cap_li, "daily_li_cap"),
            (self.e_batch_sz, "batch_size"), (self.e_batch_brk, "batch_break_sec"),
            (self.e_wa_min, "wa_min_delay"), (self.e_wa_max, "wa_max_delay"),
        ]:
            entry.delete(0, "end")
            entry.insert(0, p[k])
        self.db.save_settings({"safety_profile": preset_name, **p})
        self._refresh_quota_dock()
        self.log(f"🛡️ تم تفعيل مستوى الأمان: {preset_name}")

    def _save_shield_settings(self, popup=True):
        data = {
            "safety_profile": self.cmb_preset.get(),
            "daily_email_cap": self.e_cap_em.get().strip() or "35",
            "daily_wa_cap": self.e_cap_wa.get().strip() or "20",
            "daily_li_cap": self.e_cap_li.get().strip() or "15",
            "batch_size": self.e_batch_sz.get().strip() or "5",
            "batch_break_sec": self.e_batch_brk.get().strip() or "180",
            "wa_min_delay": self.e_wa_min.get().strip() or "50",
            "wa_max_delay": self.e_wa_max.get().strip() or "95",
        }
        self.db.save_settings(data)
        self._refresh_quota_dock()
        if popup:
            messagebox.showinfo("تم الحفظ", "تم حفظ إعدادات درع الحماية من الحظر.")

    def refresh_campaign_templates(self):
        names = [t["title"] for t in self.db.get_templates()]
        if names:
            cur = self.cmb_camp_tpl.get()
            self.cmb_camp_tpl.configure(values=names)
            if cur not in names:
                self.cmb_camp_tpl.set(names[0])

    # ======================= 6. صفحة ربط الحسابات =======================

    def _build_accounts_page(self, parent):
        s = self.settings
        ctk.CTkLabel(parent, text="🔗 ربط والتحقق من حسابات المنصات (LinkedIn • WhatsApp • Gmail)", font=("Segoe UI", 19, "bold")).pack(anchor="e", pady=(0, 12))

        # بطاقة LinkedIn
        li_c = ctk.CTkFrame(parent, corner_radius=12)
        li_c.pack(fill="x", pady=6)
        ctk.CTkLabel(li_c, text="💼 حساب LinkedIn", font=("Segoe UI", 15, "bold")).pack(side="right", padx=14, pady=14)
        self.lbl_acc_li_st = ctk.CTkLabel(li_c, text="...", font=("Segoe UI", 12, "bold"))
        self.lbl_acc_li_st.pack(side="right", padx=10)
        ctk.CTkButton(li_c, text="فصل الجلسة", width=95, fg_color="#475569", command=lambda: self._disconnect_session(LINKEDIN_SESSION_DIR, "linkedin")).pack(side="left", padx=8, pady=10)
        ctk.CTkButton(li_c, text="🔗 ربط والتحقق الآن", width=155, fg_color="#0A66C2", command=lambda: self._connect_browser_account("linkedin")).pack(side="left", padx=4, pady=10)

        # بطاقة WhatsApp Web
        wa_c = ctk.CTkFrame(parent, corner_radius=12)
        wa_c.pack(fill="x", pady=6)
        ctk.CTkLabel(wa_c, text="💬 حساب WhatsApp Web", font=("Segoe UI", 15, "bold")).pack(side="right", padx=14, pady=14)
        self.lbl_acc_wa_st = ctk.CTkLabel(wa_c, text="...", font=("Segoe UI", 12, "bold"))
        self.lbl_acc_wa_st.pack(side="right", padx=10)
        ctk.CTkButton(wa_c, text="فصل الجلسة", width=95, fg_color="#475569", command=lambda: self._disconnect_session(WA_SESSION_DIR, "whatsapp")).pack(side="left", padx=8, pady=10)
        ctk.CTkButton(wa_c, text="🔗 ربط WhatsApp (QR)", width=155, fg_color="#25D366", hover_color="#16A34A", command=lambda: self._connect_browser_account("whatsapp")).pack(side="left", padx=4, pady=10)

        # بطاقة Gmail SMTP
        gm_c = ctk.CTkFrame(parent, corner_radius=12)
        gm_c.pack(fill="x", pady=6)
        top_gm = ctk.CTkFrame(gm_c, fg_color="transparent")
        top_gm.pack(fill="x", padx=14, pady=(12, 6))
        ctk.CTkLabel(top_gm, text="✉️ حساب Gmail (SMTP + App Password)", font=("Segoe UI", 15, "bold")).pack(side="right")
        self.lbl_acc_gm_st = ctk.CTkLabel(top_gm, text="...", font=("Segoe UI", 12, "bold"))
        self.lbl_acc_gm_st.pack(side="left")

        gm_f = ctk.CTkFrame(gm_c, fg_color="transparent")
        gm_f.pack(fill="x", padx=14, pady=(4, 12))
        self.e_sender_email = ctk.CTkEntry(gm_f, placeholder_text="example@gmail.com")
        self.e_sender_email.insert(0, s.get("sender_email", ""))
        self.e_app_pass = ctk.CTkEntry(gm_f, placeholder_text="16-char Google App Password", show="*")
        self.e_app_pass.insert(0, s.get("app_password", ""))
        self.e_sender_name = ctk.CTkEntry(gm_f, placeholder_text="اسمك أو اسم شركتك", justify="right")
        self.e_sender_name.insert(0, s.get("sender_name", ""))
        self.e_attach = ctk.CTkEntry(gm_f, placeholder_text="ملف مرفق اختياري (PDF)")
        self.e_attach.insert(0, s.get("attachment_path", ""))

        ctk.CTkLabel(gm_f, text="إيميل Gmail:").grid(row=0, column=3, sticky="e", padx=6, pady=4)
        self.e_sender_email.grid(row=0, column=2, sticky="ew", padx=6, pady=4)
        ctk.CTkLabel(gm_f, text="App Password:").grid(row=0, column=1, sticky="e", padx=6, pady=4)
        self.e_app_pass.grid(row=0, column=0, sticky="ew", padx=6, pady=4)

        ctk.CTkLabel(gm_f, text="اسم المرسل:").grid(row=1, column=3, sticky="e", padx=6, pady=4)
        self.e_sender_name.grid(row=1, column=2, sticky="ew", padx=6, pady=4)
        ctk.CTkLabel(gm_f, text="مرفق اختياري:").grid(row=1, column=1, sticky="e", padx=6, pady=4)
        att_box = ctk.CTkFrame(gm_f, fg_color="transparent")
        att_box.grid(row=1, column=0, sticky="ew", padx=6, pady=4)
        self.e_attach.pack(in_=att_box, side="left", fill="x", expand=True)
        ctk.CTkButton(att_box, text="اختر", width=55, command=self._browse_attachment).pack(side="left", padx=(4, 0))

        btn_r = ctk.CTkFrame(gm_f, fg_color="transparent")
        btn_r.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        ctk.CTkButton(btn_r, text="✓ اختبار وتأكيد اتصال Gmail", fg_color="#EA4335", hover_color="#DC2626", command=self._test_gmail).pack(side="left", padx=4)
        ctk.CTkButton(btn_r, text="💾 حفظ بيانات Gmail", fg_color="#10B981", hover_color="#059669", command=self._save_gmail_settings).pack(side="left", padx=4)

        gm_f.columnconfigure(0, weight=1)
        gm_f.columnconfigure(2, weight=1)

        # نصيحة استقرار الربط
        tip = ctk.CTkFrame(parent, corner_radius=12)
        tip.pack(fill="x", pady=(10, 0))
        ctk.CTkLabel(
            tip,
            text=(
                "💡 نصيحة مهمة لاستقرار الربط ومنع الأخطاء:\n"
                "• لـ Gmail: استخدم App Password (16 حرف) من إعدادات أمان جوجل بعد تفعيل التحقق بخطوتين (مش كلمة سر حسابك العادية).\n"
                "• لـ LinkedIn و WhatsApp: اضغط زر (ربط والتحقق الآن)، سجل دخولك في المتصفح، ولا تقفل المتصفح يدوياً — اضغط OK في الرسالة والبرنامج هيحفظ الجلسة ويتحقق منها تلقائياً."
            ),
            justify="right", anchor="e", font=("Segoe UI", 11), text_color="#94A3B8",
        ).pack(fill="x", padx=14, pady=12)

    def _browse_attachment(self):
        p = filedialog.askopenfilename(title="اختر ملفاً")
        if p:
            self.e_attach.delete(0, "end")
            self.e_attach.insert(0, p)

    def _save_gmail_settings(self, popup=True):
        data = {
            "sender_email": self.e_sender_email.get().strip(),
            "app_password": self.e_app_pass.get().strip(),
            "sender_name": self.e_sender_name.get().strip(),
            "attachment_path": self.e_attach.get().strip(),
        }
        self.db.save_settings(data)
        self.refresh_account_badges()
        if popup:
            messagebox.showinfo("تم الحفظ", "تم حفظ بيانات Gmail بنجاح.")

    def _test_gmail(self):
        self._save_gmail_settings(popup=False)
        s = self.db.get_settings()
        em, pw = s.get("sender_email", ""), s.get("app_password", "")
        if not em or not pw:
            messagebox.showwarning("ناقص", "اكتب إيميل Gmail والـ App Password أولاً.")
            return

        def _w():
            try:
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=20) as smtp:
                    smtp.login(em, pw.replace(" ", ""))
                self.db.save_settings({"gmail_verified": "1"})
                self.log("✅ نجح الاتصال وتأكيد حساب Gmail!")
                self.event_queue.put(("badges", None))
                self.after(0, lambda: messagebox.showinfo("نجاح الربط", "تم التحقق من حساب Gmail بنجاح!"))
            except Exception as e:
                self.db.save_settings({"gmail_verified": "0"})
                self.log(f"❌ فشل اتصال Gmail: {e}")
                self.event_queue.put(("badges", None))
                self.after(0, lambda err=str(e): messagebox.showerror("فشل الاتصال", f"تأكد من App Password (16 حرف):\n{err}"))

        threading.Thread(target=_w, daemon=True).start()

    def _connect_browser_account(self, acc_type):
        lock = self.li_browser_lock if acc_type == "linkedin" else self.wa_browser_lock
        session_dir = LINKEDIN_SESSION_DIR if acc_type == "linkedin" else WA_SESSION_DIR
        if not lock.acquire(blocking=False):
            messagebox.showwarning("المتصفح مشغول", "هذا الحساب مستخدم حالياً في عملية سحب أو إرسال. أوقفها أولاً.")
            return

        def _w():
            try:
                connect_and_verify_account(session_dir, acc_type, self.log, self.ask_main_thread, self.db)
            except Exception as e:
                self.log(f"❌ خطأ أثناء ربط {acc_type}: {e}")
                self.after(0, lambda err=str(e): messagebox.showerror("خطأ في الربط", err))
            finally:
                lock.release()
                self.event_queue.put(("badges", None))

        threading.Thread(target=_w, daemon=True).start()

    def _disconnect_session(self, folder, acc_type):
        lock = self.li_browser_lock if acc_type == "linkedin" else self.wa_browser_lock
        if not lock.acquire(blocking=False):
            messagebox.showwarning("مشغول", "أوقف العملية التي تستخدم هذا المتصفح أولاً.")
            return
        try:
            if os.path.isdir(folder) and messagebox.askyesno("تأكيد الفصل", f"مسح جلسة {acc_type} المحفوظة؟"):
                shutil.rmtree(folder, ignore_errors=True)
                self.db.save_settings({"li_verified" if acc_type == "linkedin" else "wa_verified": "0"})
                self.log(f"تم فصل حساب {acc_type}.")
        finally:
            lock.release()
            self.refresh_account_badges()

    # ======================= إدارة الأحداث والعمليات المتزامنة =======================

    def log(self, text: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.event_queue.put(("log", f"[{ts}] {text}"))

    def _clear_logs(self):
        for box in (self.scrape_log, self.camp_log):
            box.configure(state="normal")
            box.delete("1.0", "end")
            box.configure(state="disabled")

    def ask_main_thread(self, title: str, msg: str) -> bool:
        res = {"ok": False}
        ev = threading.Event()
        def _pop():
            res["ok"] = messagebox.askokcancel(title, msg, parent=self)
            ev.set()
        self.after(0, _pop)
        ev.wait()
        return res["ok"]

    def _poll_events(self):
        needs_refresh = False
        while not self.event_queue.empty():
            kind, payload = self.event_queue.get_nowait()
            if kind == "log":
                for box in (self.scrape_log, self.camp_log):
                    box.configure(state="normal")
                    box.insert("end", payload + "\n")
                    box.see("end")
                    box.configure(state="disabled")
            elif kind == "scrape_progress":
                cur, tot = payload
                self.scrape_progress.set(min(1.0, cur / max(1, tot)))
                self.lbl_dock_scraper.configure(text=f"🔍 السحب: {cur}/{tot}", text_color="#38BDF8")
            elif kind == "camp_progress":
                cur, tot = payload
                self.camp_progress.set(min(1.0, cur / max(1, tot)))
                self.lbl_dock_camp.configure(text=f"🚀 الإرسال: {cur}/{tot}", text_color="#22C55E")
            elif kind == "refresh_views":
                needs_refresh = True
            elif kind == "badges":
                self.refresh_account_badges()
            elif kind == "scraper_done":
                self.scraper_running = False
                self.btn_start_scrape.configure(state="normal")
                self.btn_stop_scrape.configure(state="disabled")
                self.btn_dock_stop_scraper.configure(state="disabled")
                self.lbl_dock_scraper.configure(text="🔍 السحب: مكتمل", text_color="#22C55E")
                needs_refresh = True
            elif kind == "camp_done":
                self._update_campaign_buttons()
                needs_refresh = True

        if needs_refresh:
            self.refresh_leads_table()
            self._refresh_quota_dock()
            if self.pages["dashboard"].winfo_ismapped():
                self.refresh_dashboard()

        self.after(150, self._poll_events)

    # --- تشغيل السحب (مستقل عن الحملات!) ---
    def _start_scraping(self):
        if self.scraper_running:
            messagebox.showinfo("جاري السحب", "محرك السحب يعمل بالفعل حالياً.")
            return
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
        if not (cfg["src_li"] or cfg["src_gm"] or cfg["src_web"]):
            messagebox.showwarning("اختر مصدر", "اختر موقعاً واحداً على الأقل للسحب.")
            return
        if not cfg["keywords"]:
            messagebox.showwarning("كلمات مفتاحية", "اكتب كلمة مفتاحية واحدة على الأقل.")
            return

        self.db.save_settings({
            "keywords": cfg["keywords"], "location": cfg["location"], "target_count": str(target),
            "contact_req": cfg["contact_req"], "icp_industry": cfg["icp_industry"],
            "icp_size": cfg["icp_size"], "min_rating": cfg["min_rating"],
        })

        self.scraper_running = True
        self.scraper_stop.clear()
        self.scrape_progress.set(0)
        self.btn_start_scrape.configure(state="disabled")
        self.btn_stop_scrape.configure(state="normal")
        self.btn_dock_stop_scraper.configure(state="normal")
        self.lbl_dock_scraper.configure(text=f"🔍 السحب: 0/{target}", text_color="#38BDF8")

        self.log("=" * 55)
        self.log(f"🚀 بدء السحب الموازي | المطلوب: {target} | يمكنك بدء الإرسال فوراً مع ظهور أول عميل!")

        def _worker():
            try:
                keywords = [k.strip() for k in re.split(r"[,،\n]+", cfg["keywords"]) if k.strip()]
                loc = cfg["location"]
                counter = [0]

                def progress_cb(c, t):
                    self.event_queue.put(("scrape_progress", (c, t)))

                def accept_lead(lead):
                    if self.db.lead_exists(lead["name"], lead.get("profile_url", ""), lead.get("phone", ""), lead.get("email", "")):
                        self.log(f"   ↳ تخطي (مكرر في القاعدة): {lead['name']}")
                        return False
                    ok, reason = matches_icp_and_contact(lead, cfg)
                    if not ok:
                        self.log(f"   ↳ استبعاد ({lead['name']}): {reason}")
                        return False
                    self.db.add_lead(lead)
                    counter[0] += 1
                    self.log(f"   ✅ [{counter[0]}/{target}] أُضيف للقاعدة: {lead['name']} | WA: {lead.get('phone') or '-'} | Mail: {lead.get('email') or '-'}")
                    self.event_queue.put(("refresh_views", None))
                    return True

                if cfg["src_gm"] and counter[0] < target and not self.scraper_stop.is_set():
                    scrape_google_maps(cfg, keywords, loc, target, counter, accept_lead, self.log, self.scraper_stop, progress_cb)

                if cfg["src_li"] and counter[0] < target and not self.scraper_stop.is_set():
                    scrape_linkedin(cfg, keywords, loc, target, counter, accept_lead, self.log, self.scraper_stop, progress_cb, self.ask_main_thread, self.li_browser_lock, self.db)

                if cfg["src_web"] and counter[0] < target and not self.scraper_stop.is_set():
                    scrape_web_search(cfg, keywords, loc, target, counter, accept_lead, self.log, self.scraper_stop, progress_cb)

                self.log(f"🏁 انتهى محرك السحب. إجمالي العملاء الجدد المحفوظين: {counter[0]}")
            except Exception as e:
                self.log(f"❌ خطأ في محرك السحب: {e}")
            finally:
                self.event_queue.put(("scraper_done", None))

        threading.Thread(target=_worker, daemon=True).start()

    def stop_scraper(self):
        if self.scraper_running:
            self.scraper_stop.set()
            self.log("⛔ تم طلب إيقاف محرك السحب...")

    # --- تشغيل الحملات (مستقلة وتعمل حتى أثناء السحب!) ---
    def _launch_campaign_for_selected(self):
        if not self.selected_lead_ids:
            messagebox.showwarning("تحديد", "حدد عملاء من عمود العلامة ☑ أولاً.")
            return
        chosen = set(self.selected_lead_ids)
        self._show_page("campaigns")
        self._trigger_campaign(chosen)

    def _update_campaign_buttons(self):
        any_running = self.email_running or self.wa_running or self.li_running
        self.btn_stop_camp.configure(state="normal" if any_running else "disabled")
        self.btn_dock_stop_camp.configure(state="normal" if any_running else "disabled")
        if not any_running:
            self.lbl_dock_camp.configure(text="🚀 الحملات: جاهز", text_color="#94A3B8")

    def _trigger_campaign(self, selected_ids):
        self._save_shield_settings(popup=False)
        tpls = self.db.get_templates()
        tpl = next((t for t in tpls if t["title"] == self.cmb_camp_tpl.get()), tpls[0] if tpls else None)
        if not tpl:
            messagebox.showwarning("لا يوجد قالب", "أضف قالب رسالة أولاً.")
            return

        if selected_ids is not None:
            leads = [self.db.get_lead_by_id(i) for i in selected_ids]
            leads = [l for l in leads if l]
        else:
            leads = self.db.get_leads(stage_filter=self.cmb_camp_stage.get())

        if not leads:
            messagebox.showwarning("لا يوجد عملاء", "لا يوجد عملاء جاهزون في النطاق المختار بعد.")
            return

        use_em = self.var_ch_gmail.get()
        use_wa = self.var_ch_wa.get()
        use_li = self.var_ch_li.get()
        dry = self.var_dry_run.get()
        if not (use_em or use_wa or use_li):
            messagebox.showwarning("اختر قناة", "فعّل قناة واحدة على الأقل.")
            return

        s = self.db.get_settings()
        self.camp_progress.set(0)

        def prog_cb(c, t):
            self.event_queue.put(("camp_progress", (c, t)))
            self.event_queue.put(("refresh_views", None))

        # تشغيل كل قناة في Thread مستقل عشان الإيميل والواتساب يشتغلوا بالتوازي مع السحب!
        if use_em and not self.email_running:
            self.email_running = True
            self.email_stop.clear()
            def _em_w():
                try:
                    run_email_campaign(leads, tpl, s, dry, self.db, self.log, self.email_stop, prog_cb)
                finally:
                    self.email_running = False
                    self.event_queue.put(("camp_done", None))
            threading.Thread(target=_em_w, daemon=True).start()

        if use_wa and not self.wa_running:
            self.wa_running = True
            self.wa_stop.clear()
            def _wa_w():
                try:
                    run_whatsapp_campaign(leads, tpl, s, dry, self.db, self.log, self.wa_stop, prog_cb, self.ask_main_thread, self.wa_browser_lock)
                finally:
                    self.wa_running = False
                    self.event_queue.put(("camp_done", None))
            threading.Thread(target=_wa_w, daemon=True).start()

        if use_li and not self.li_running:
            self.li_running = True
            self.li_stop.clear()
            def _li_w():
                try:
                    run_linkedin_campaign(leads, tpl, s, dry, self.db, self.log, self.li_stop, prog_cb, self.ask_main_thread, self.li_browser_lock)
                finally:
                    self.li_running = False
                    self.event_queue.put(("camp_done", None))
            threading.Thread(target=_li_w, daemon=True).start()

        self._update_campaign_buttons()

    def stop_all_campaigns(self):
        self.email_stop.set()
        self.wa_stop.set()
        self.li_stop.set()
        self.log("⛔ تم طلب إيقاف جميع حملات الإرسال...")

    def _on_close(self):
        if self.scraper_running or self.email_running or self.wa_running or self.li_running:
            if not messagebox.askyesno("تأكيد الخروج", "توجد عمليات شغالة حالياً. هل تريد إيقافها والخروج؟"):
                return
            self.scraper_stop.set()
            self.stop_all_campaigns()
        self.destroy()


if __name__ == "__main__":
    app = SmartCRMProApp()
    app.mainloop()
