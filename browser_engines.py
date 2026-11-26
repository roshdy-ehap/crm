import os
import re
import time
import random
from datetime import datetime
from urllib.parse import quote, unquote

import requests
from playwright.sync_api import sync_playwright

from crm_core import (
    LINKEDIN_SESSION_DIR,
    WA_SESSION_DIR,
    extract_mobile_number,
    format_whatsapp_phone,
    extract_email,
    is_valid_email,
    enrich_from_website,
    matches_icp_and_contact,
    render_msg,
    send_single_smtp_email,
)


def launch_persistent_browser(p, session_dir: str, log_fn=None):
    """فتح متصفح مستمر مع تجربة Chrome ثم Edge ثم Chromium وتفاصيل خطأ واضحة"""
    os.makedirs(session_dir, exist_ok=True)
    errors = []
    launch_args = [
        "--disable-blink-features=AutomationControlled",
        "--no-first-run",
        "--no-default-browser-check",
    ]
    for ch_label, ch in [("Google Chrome", "chrome"), ("Microsoft Edge", "msedge"), ("Chromium", None)]:
        try:
            kwargs = {
                "user_data_dir": os.path.abspath(session_dir),
                "headless": False,
                "viewport": {"width": 1280, "height": 820},
                "args": launch_args,
            }
            if ch:
                kwargs["channel"] = ch
            ctx = p.chromium.launch_persistent_context(**kwargs)
            if log_fn:
                log_fn(f"🌐 تم فتح المتصفح بنجاح عبر: {ch_label}")
            return ctx
        except Exception as e:
            err_msg = str(e).splitlines()[0]
            errors.append(f"{ch_label}: {err_msg}")
            continue
    raise RuntimeError(
        "تعذر فتح أي متصفح (Chrome / Edge / Chromium).\n"
        "تأكد من إغلاق أي نافذة متصفح قديمة تابعة للبرنامج أو تثبيت Google Chrome.\n"
        "التفاصيل: " + " | ".join(errors)
    )


def launch_temp_browser(p, log_fn=None):
    errors = []
    launch_args = [
        "--disable-blink-features=AutomationControlled",
        "--no-first-run",
        "--no-default-browser-check",
    ]
    for ch_label, ch in [("Google Chrome", "chrome"), ("Microsoft Edge", "msedge"), ("Chromium", None)]:
        try:
            kwargs = {"headless": False, "args": launch_args}
            if ch:
                kwargs["channel"] = ch
            browser = p.chromium.launch(**kwargs)
            if log_fn:
                log_fn(f"🌐 تم فتح المتصفح عبر: {ch_label}")
            return browser
        except Exception as e:
            errors.append(f"{ch_label}: {str(e).splitlines()[0]}")
            continue
    raise RuntimeError("تعذر فتح المتصفح: " + " | ".join(errors))


def connect_and_verify_account(session_dir: str, account_type: str, log_fn, ask_confirm_fn, db):
    """
    ربط الحساب والتحقق الفعلي من تسجيل الدخول (فحص كوكيز LinkedIn أو قائمة محادثات WhatsApp)
    """
    is_li = account_type == "linkedin"
    label = "LinkedIn" if is_li else "WhatsApp Web"
    url = "https://www.linkedin.com/feed/" if is_li else "https://web.whatsapp.com"

    log_fn(f"🔗 بدء ربط والتحقق من حساب {label}...")
    verified = False

    with sync_playwright() as p:
        ctx = launch_persistent_browser(p, session_dir, log_fn)
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            log_fn(f"➡️ جاري فتح {url}...")
            page.goto(url, wait_until="domcontentloaded", timeout=30000)

            ok = ask_confirm_fn(
                f"ربط حساب {label}",
                (
                    f"1. قم بتسجيل الدخول في نافذة {label} التي فتحت الآن.\n"
                    f"2. انتظر حتى تظهر صفحتك الرئيسية {'(Feed)' if is_li else '(قائمة المحادثات Chats)'} بالكامل.\n"
                    f"3. لا تغلق المتصفح يدوياً — اضغط OK هنا ليقوم البرنامج بالتحقق وحفظ الجلسة."
                ),
            )
            if ok:
                try:
                    if is_li:
                        cookies = ctx.cookies()
                        has_li_at = any(c.get("name") == "li_at" and c.get("value") for c in cookies)
                        cur_url = page.url.lower()
                        if has_li_at or "/feed" in cur_url or "/mynetwork" in cur_url:
                            verified = True
                            log_fn("✅ تم التحقق من جلسة LinkedIn بنجاح (كوكي الدخول فعّال).")
                        else:
                            log_fn("⚠️ لم يتم رصد تسجيل دخول فعلي في LinkedIn. تأكد من إدخال بياناتك ثم أعد المحاولة.")
                    else:
                        wa_logged = (
                            page.locator("#pane-side, div[aria-label='Chat list'], div[data-testid='chat-list']").count() > 0
                            or page.locator("canvas[aria-label*='Scan'], div[data-ref]").count() == 0
                        )
                        if wa_logged:
                            verified = True
                            log_fn("✅ تم التحقق من جلسة WhatsApp Web بنجاح وحفظها.")
                        else:
                            log_fn("⚠️ يبدو أن كود QR لم يُمسح بعد. أعد المحاولة وامسح الكود من موبايلك.")
                except Exception as check_err:
                    log_fn(f"⚠️ تعذر التحقق التلقائي (ربما أغلقت النافذة مبكراً): {check_err}")
                    verified = os.path.isdir(session_dir) and len(os.listdir(session_dir)) > 2
            else:
                log_fn(f"تم إلغاء ربط {label}.")
        finally:
            try:
                ctx.close()
            except Exception:
                pass

    db.save_settings({"li_verified" if is_li else "wa_verified": "1" if verified else "0"})
    return verified


# ======================= محرك السحب (Scraper Engine) =======================

def _dismiss_google_consent(page, log_fn):
    try:
        for sel in (
            "button:has-text('Accept all')",
            "button:has-text('I agree')",
            "button:has-text('قبول الكل')",
            "button:has-text('أوافق')",
            "form[action*='consent'] button",
        ):
            btn = page.locator(sel).first
            if btn.count() > 0 and btn.is_visible():
                btn.click()
                log_fn("✓ تم تجاوز نافذة موافقة Google تلقائياً.")
                time.sleep(2)
                return
    except Exception:
        pass


def scrape_google_maps(cfg: dict, keywords: list[str], loc: str, target: int, counter: list[int], accept_fn, log_fn, stop_event, progress_fn):
    log_fn("🗺️ === بدء السحب من Google Maps ===")
    with sync_playwright() as p:
        browser = launch_temp_browser(p, log_fn)
        try:
            ctx = browser.new_context(viewport={"width": 1280, "height": 820}, locale="en-US")
            page = ctx.new_page()

            for kw in keywords:
                if counter[0] >= target or stop_event.is_set():
                    break
                query = f"{kw} in {loc}" if loc else kw
                search_url = f"https://www.google.com/maps/search/{quote(query)}"
                log_fn(f"🔎 [Google Maps] فتح البحث: {query}")
                try:
                    page.goto(search_url, wait_until="domcontentloaded", timeout=25000)
                except Exception as e:
                    log_fn(f"⚠️ [Google Maps] تعذر فتح صفحة البحث: {e}")
                    continue

                if stop_event.wait(3.5):
                    break
                _dismiss_google_consent(page, log_fn)

                place_urls = []
                seen_u = set()
                feed = page.locator("div[role='feed']")
                max_needed = max(20, (target - counter[0]) * 3)

                for scroll_step in range(14):
                    if stop_event.is_set() or len(place_urls) >= max_needed:
                        break
                    links = page.locator("a.hfpxzc, a[href*='/maps/place/']")
                    cnt = links.count()
                    for i in range(cnt):
                        try:
                            href = links.nth(i).get_attribute("href")
                            if href and "/maps/place/" in href:
                                clean_h = href.split("&authuser=")[0]
                                if clean_h not in seen_u:
                                    seen_u.add(clean_h)
                                    place_urls.append(clean_h)
                        except Exception:
                            continue

                    if scroll_step == 0:
                        log_fn(f"📋 [Google Maps] جاري التمرير لتجميع الروابط (مبدئياً: {len(place_urls)} مكان)...")

                    try:
                        if feed.count() > 0:
                            feed.first.evaluate("el => el.scrollBy(0, 2500)")
                        else:
                            page.mouse.wheel(0, 2200)
                    except Exception:
                        pass
                    if stop_event.wait(1.8):
                        break

                log_fn(f"📌 [Google Maps] تم تجميع {len(place_urls)} رابط مكان. بدء الاستخراج المباشر...")
                if not place_urls:
                    log_fn("⚠️ [Google Maps] لم تظهر نتائج في القائمة. تأكد من كتابة اللوكيشن بشكل صحيح أو جرب كلمة مفتاحية أوسع.")
                    continue

                for idx, p_url in enumerate(place_urls, start=1):
                    if counter[0] >= target or stop_event.is_set():
                        break
                    try:
                        log_fn(f"➡️ [Maps {idx}/{len(place_urls)}] جاري فتح الصفحة واستخراج البيانات...")
                        page.goto(p_url, wait_until="domcontentloaded", timeout=18000)
                        if stop_event.wait(random.uniform(2.0, 3.2)):
                            break

                        # استخراج اسم المكان بعدة بدائل
                        name = ""
                        for h_sel in ("h1.DUwDvf", "h1[class*='fontHeadlineLarge']", "div[role='main'] h1"):
                            if page.locator(h_sel).count() > 0:
                                txt = page.locator(h_sel).first.inner_text().strip()
                                if txt and txt.lower() != "google maps":
                                    name = txt
                                    break
                        if not name:
                            pt = page.title().replace(" - Google Maps", "").strip()
                            name = pt if pt and pt.lower() != "google maps" else f"Business #{idx}"

                        cat = kw
                        cat_loc = page.locator("button[jsaction*='category']")
                        if cat_loc.count() > 0:
                            cat = cat_loc.first.inner_text().strip() or kw

                        rating = ""
                        r_loc = page.locator("div.F7nice span[aria-hidden='true']")
                        if r_loc.count() > 0:
                            rating = r_loc.first.inner_text().strip()

                        phone = ""
                        ph_loc = page.locator("button[data-item-id^='phone:'], a[href^='tel:']")
                        if ph_loc.count() > 0:
                            raw_ph = (ph_loc.first.get_attribute("aria-label") or ph_loc.first.get_attribute("href") or ph_loc.first.inner_text() or "")
                            phone = extract_mobile_number(raw_ph) or re.sub(r"[^\d+]", "", raw_ph)

                        website = ""
                        web_loc = page.locator("a[data-item-id='authority'], a[aria-label*='Website']")
                        if web_loc.count() > 0:
                            website = web_loc.first.get_attribute("href") or ""

                        email = ""
                        if cfg.get("enrich_web") and website:
                            log_fn(f"   🌐 فحص موقع {name} لاستخراج الإيميل والواتساب...")
                            wm, wwa = enrich_from_website(website)
                            if wm:
                                email = wm
                            if not format_whatsapp_phone(phone) and wwa:
                                phone = wwa

                        accept_fn({
                            "name": name,
                            "source": "Google Maps",
                            "category": cat,
                            "size_rating": f"★ {rating}" if rating else "",
                            "location": loc,
                            "website": website,
                            "phone": phone,
                            "email": email,
                            "profile_url": p_url,
                        })
                        progress_fn(counter[0], target)
                    except Exception as e:
                        log_fn(f"⚠️ [Google Maps] خطأ أثناء قراءة المكان رقم {idx}: {str(e)[:90]}")
                        continue
        finally:
            try:
                browser.close()
            except Exception:
                pass


def scrape_linkedin(cfg: dict, keywords: list[str], loc: str, target: int, counter: list[int], accept_fn, log_fn, stop_event, progress_fn, ask_confirm_fn, li_lock, db):
    log_fn("💼 === بدء السحب من LinkedIn ===")
    if not li_lock.acquire(blocking=False):
        log_fn("⚠️ متصفح LinkedIn مستخدم حالياً في عملية أخرى (إرسال رسائل أو ربط). انتظر انتهاءها أو أوقفها.")
        return

    try:
        with sync_playwright() as p:
            ctx = launch_persistent_browser(p, LINKEDIN_SESSION_DIR, log_fn)
            try:
                page = ctx.pages[0] if ctx.pages else ctx.new_page()
                page.goto("https://www.linkedin.com/feed/", wait_until="domcontentloaded", timeout=25000)
                time.sleep(2.5)

                # التحقق التلقائي من وجود كوكي الدخول قبل إزعاج المستخدم بنافذة تأكيد
                cookies = ctx.cookies()
                logged_in = any(c.get("name") == "li_at" and c.get("value") for c in cookies) and ("/login" not in page.url.lower())
                if not logged_in:
                    log_fn("🔑 LinkedIn يطلب تسجيل الدخول...")
                    if not ask_confirm_fn(
                        "تسجيل دخول LinkedIn",
                        "سجّل دخولك في نافذة LinkedIn الآن، وبعد ظهور الصفحة الرئيسية (Feed) اضغط OK لبدء السحب.",
                    ):
                        log_fn("تم إلغاء سحب LinkedIn.")
                        return
                    cookies = ctx.cookies()
                    if any(c.get("name") == "li_at" and c.get("value") for c in cookies):
                        db.save_settings({"li_verified": "1"})
                else:
                    log_fn("✅ تم الدخول تلقائياً إلى LinkedIn باستخدام الجلسة المحفوظة!")
                    db.save_settings({"li_verified": "1"})

                seen_urls = set()
                for kw in keywords:
                    if counter[0] >= target or stop_event.is_set():
                        break
                    search_q = f"{kw} {loc}".strip()
                    base_url = f"https://www.linkedin.com/search/results/companies/?keywords={quote(search_q)}"
                    company_links = []
                    page_num = 1
                    needed_links = max(15, (target - counter[0]) * 2)

                    while len(company_links) < needed_links and page_num <= 10 and not stop_event.is_set():
                        url = f"{base_url}&page={page_num}"
                        log_fn(f"🔎 [LinkedIn] فتح صفحة البحث {page_num} لـ '{search_q}'...")
                        try:
                            page.goto(url, wait_until="domcontentloaded", timeout=25000)
                        except Exception as e:
                            log_fn(f"⚠️ [LinkedIn] تعذر تحميل الصفحة {page_num}: {e}")
                            break

                        if stop_event.wait(random.uniform(3.5, 5.5)):
                            break
                        page.mouse.wheel(0, 2200)
                        if stop_event.wait(1.8):
                            break

                        # استخراج مرن جداً يشتغل حتى لو LinkedIn غيّر أسماء الـ CSS Classes
                        extracted = page.evaluate("""() => {
                            const out = [];
                            const seen = new Set();
                            const anchors = document.querySelectorAll("main a[href*='/company/']");
                            anchors.forEach(a => {
                                const rawHref = a.getAttribute("href") || "";
                                const m = rawHref.match(/https?:\\/\\/[^/]*linkedin\\.com\\/company\\/[^/?#]+|\\/company\\/[^/?#]+/i);
                                if (!m) return;
                                let full = m[0];
                                if (full.startsWith("/")) full = "https://www.linkedin.com" + full;
                                full = full.replace(/\\/$/, "");
                                if (seen.has(full)) return;
                                const txt = (a.innerText || "").trim().split("\\n")[0].trim();
                                if (!txt || txt.length < 2 || txt.toLowerCase().includes("followers") || txt.toLowerCase().includes("متابع")) return;
                                seen.add(full);
                                out.push({name: txt, url: full});
                            });
                            return out;
                        }""")

                        log_fn(f"📋 [LinkedIn] صفحة {page_num}: تم رصد {len(extracted)} شركة.")
                        if not extracted:
                            break

                        new_on_page = 0
                        for item in extracted:
                            if item["url"] not in seen_urls:
                                seen_urls.add(item["url"])
                                company_links.append((item["name"], item["url"]))
                                new_on_page += 1
                        if new_on_page == 0:
                            break
                        page_num += 1

                    log_fn(f"📌 [LinkedIn] إجمالي الشركات المرصودة لكلمة '{kw}': {len(company_links)}. بدء فحص صفحات About...")

                    for idx, (name, comp_url) in enumerate(company_links, start=1):
                        if counter[0] >= target or stop_event.is_set():
                            break
                        about_url = comp_url.rstrip("/") + "/about/"
                        try:
                            log_fn(f"➡️ [LinkedIn {idx}/{len(company_links)}] فتح: {name} ({about_url})")
                            page.goto(about_url, wait_until="domcontentloaded", timeout=20000)
                            if stop_event.wait(random.uniform(2.5, 4.2)):
                                break

                            fields = page.evaluate("""() => {
                                const r = {};
                                document.querySelectorAll('dl dt').forEach(dt => {
                                    const k = (dt.innerText || '').trim().toLowerCase();
                                    const dd = dt.nextElementSibling;
                                    if (dd && dd.tagName === 'DD') r[k] = (dd.innerText || '').trim();
                                });
                                // بديل لاستخراج الموقع الخارجي لو لم يوجد في dl
                                const ext = Array.from(document.querySelectorAll("main a[href^='http']"))
                                    .map(a => a.getAttribute("href"))
                                    .find(h => h && !h.includes("linkedin.com") && !h.includes("bing.com"));
                                if (ext) r["__fallback_website"] = ext;
                                return r;
                            }""")

                            website, industry, size_txt, phone, email = "", kw, "", "", ""
                            for k, v in fields.items():
                                if k == "__fallback_website":
                                    continue
                                cv = v.split("\n")[0].strip()
                                if any(w in k for w in ("website", "الموقع")):
                                    website = cv
                                elif any(w in k for w in ("industry", "المجال", "الصناعة")):
                                    industry = cv
                                elif any(w in k for w in ("company size", "حجم الشركة")):
                                    size_txt = cv
                                elif any(w in k for w in ("phone", "الهاتف")):
                                    phone = extract_mobile_number(cv) or cv

                            if not website and fields.get("__fallback_website"):
                                website = fields["__fallback_website"]

                            main_txt = page.locator("main").inner_text() if page.locator("main").count() else ""
                            if not email:
                                email = extract_email(main_txt) or ""
                            if not format_whatsapp_phone(phone):
                                mt = extract_mobile_number(main_txt)
                                if mt:
                                    phone = mt

                            if cfg.get("enrich_web") and website and (not email or not format_whatsapp_phone(phone)):
                                log_fn(f"   🌐 فحص موقع {website} لاستخراج الإيميل والواتساب...")
                                wm, wwa = enrich_from_website(website)
                                if not email and wm:
                                    email = wm
                                if not format_whatsapp_phone(phone) and wwa:
                                    phone = wwa

                            accept_fn({
                                "name": name,
                                "source": "LinkedIn",
                                "category": industry,
                                "size_rating": size_txt,
                                "location": loc,
                                "website": website,
                                "phone": phone,
                                "email": email,
                                "profile_url": comp_url,
                            })
                            progress_fn(counter[0], target)
                        except Exception as e:
                            log_fn(f"⚠️ [LinkedIn] تعذر قراءة {name}: {str(e)[:90]}")
                            continue
            finally:
                try:
                    ctx.close()
                except Exception:
                    pass
    finally:
        li_lock.release()


def scrape_web_search(cfg: dict, keywords: list[str], loc: str, target: int, counter: list[int], accept_fn, log_fn, stop_event, progress_fn):
    log_fn("🌐 === بدء السحب من محركات البحث (Web) ===")
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"}
    seen_domains = set()
    skip_domains = ("facebook.com", "youtube.com", "instagram.com", "tiktok.com", "wikipedia.org", "duckduckgo.com", "google.com")

    for kw in keywords:
        if counter[0] >= target or stop_event.is_set():
            break
        query = f'"{kw}" "{loc}" (contact OR email OR whatsapp)' if loc else f'"{kw}" contact'
        log_fn(f"🔎 [Web] جاري البحث عن: {query}")

        found_links = []
        try:
            resp = requests.post("https://html.duckduckgo.com/html/", data={"q": query}, headers=headers, timeout=12)
            matches = re.findall(r'<a[^>]+class="result__a"[^>]+href="([^"]+)"[^>]*>(.*?)</a>', resp.text, re.S)
            for raw_href, raw_title in matches:
                title = re.sub(r"<[^>]+>", "", raw_title).strip()
                m_uddg = re.search(r"uddg=([^&]+)", raw_href)
                url = unquote(m_uddg.group(1)) if m_uddg else raw_href
                if url.startswith("http"):
                    found_links.append((title, url))
        except Exception as e:
            log_fn(f"⚠️ [Web] تعذر البحث عبر DuckDuckGo: {e}")

        log_fn(f"📋 [Web] تم العثور على {len(found_links)} نتيجة. جاري فحص المواقع...")
        for idx, (title, url) in enumerate(found_links, start=1):
            if counter[0] >= target or stop_event.is_set():
                break
            if any(s in url.lower() for s in skip_domains):
                continue
            dom = re.sub(r"^https?://(www\.)?", "", url).split("/")[0].lower()
            if dom in seen_domains:
                continue
            seen_domains.add(dom)

            log_fn(f"➡️ [Web {idx}/{len(found_links)}] فحص موقع: {dom}")
            email, wa = enrich_from_website(url)
            accept_fn({
                "name": title[:60] or dom,
                "source": "Web Search",
                "category": kw,
                "size_rating": "",
                "location": loc,
                "website": url,
                "phone": wa or "",
                "email": email or "",
                "profile_url": url,
            })
            progress_fn(counter[0], target)


# ======================= محرك الحملات المحمي ضد الحظر =======================

def _smart_wait(seconds: float, stop_event, log_fn, label: str = ""):
    if seconds <= 0:
        return False
    if label:
        log_fn(label)
    return stop_event.wait(seconds)


def run_email_campaign(leads: list[dict], tpl: dict, s: dict, dry: bool, db, log_fn, stop_event, progress_fn):
    log_fn("✉️ === بدء حملة Gmail SMTP ===")
    em = s.get("sender_email", "").strip()
    pw = s.get("app_password", "").replace(" ", "")
    sn = s.get("sender_name", "").strip()
    att = s.get("attachment_path", "").strip()

    daily_cap = int(s.get("daily_email_cap") or 35)
    min_d = int(s.get("email_min_delay") or 35)
    max_d = int(s.get("email_max_delay") or 75)
    batch_sz = max(1, int(s.get("batch_size") or 5))
    batch_brk = int(s.get("batch_break_sec") or 180)

    if not dry and (not em or not pw):
        log_fn("❌ خطأ: لم يتم ربط حساب Gmail. أدخل الإيميل والـ App Password في صفحة ربط الحسابات.")
        return

    today_used = db.get_today_counts()["email"]
    remaining_quota = max(0, daily_cap - today_used)
    if not dry and remaining_quota == 0:
        log_fn(f"🛡️ درع الحماية: وصلت للحد اليومي الآمن لإيميلات Gmail ({today_used}/{daily_cap}). أوقفنا الإرسال تلقائياً لحماية حسابك.")
        return

    eligible = [
        l for l in leads
        if l.get("stage") != "عدم الإزعاج"
        and not (l.get("email_status") or "").startswith("تم")
        and is_valid_email(l.get("email"))
    ]
    if not eligible:
        log_fn("ℹ️ لا يوجد عملاء لديهم إيميل صالح (أو غير مرسل مسبقاً) في القائمة المختارة.")
        return

    log_fn(f"📊 عملاء جاهزون للإيميل: {len(eligible)} | المتبقي من الحد اليومي: {remaining_quota}/{daily_cap} | وضع التجربة: {dry}")
    sent = 0
    limit = len(eligible) if dry else min(len(eligible), remaining_quota)

    for idx, lead in enumerate(eligible[:limit], start=1):
        if stop_event.is_set():
            log_fn("⛔ تم إيقاف حملة Gmail.")
            break
        to_em = lead["email"].strip()
        subj = render_msg(tpl["subject"], lead, sn, apply_spintax=True)
        body = render_msg(tpl["body"], lead, sn, apply_spintax=True)

        if dry:
            log_fn(f"🧪 [تجربة Gmail {idx}/{limit}] {lead['name']} -> {to_em} | العنوان: {subj}")
            sent += 1
            progress_fn(sent, limit)
            continue

        try:
            send_single_smtp_email(em, pw, sn, to_em, subj, body, att)
            now_t = datetime.now().strftime("%m-%d %H:%M")
            db.update_channel_status(lead["id"], "email_status", f"تم ({now_t})", promote_stage=True)
            db.record_activity(lead["id"], "email", "sent")
            sent += 1
            progress_fn(sent, limit)
            log_fn(f"✅ [Gmail {sent}/{limit}] أُرسل إلى {lead['name']} ({to_em})")

            if sent < limit:
                if sent % batch_sz == 0:
                    brk = random.uniform(batch_brk * 0.85, batch_brk * 1.15)
                    if _smart_wait(brk, stop_event, log_fn, f"🛡️ استراحة حماية بين الدفعات: انتظار {brk/60:.1f} دقيقة..."):
                        break
                else:
                    wt = random.uniform(min_d, max_d)
                    if _smart_wait(wt, stop_event, log_fn, f"⏳ انتظار عشوائي آمن {wt:.1f} ثانية..."):
                        break
        except Exception as e:
            db.update_channel_status(lead["id"], "email_status", f"فشل: {str(e)[:22]}")
            log_fn(f"❌ [Gmail] فشل مع {lead['name']}: {e}")

    log_fn(f"🏁 انتهت حملة Gmail. الإجمالي المعالج: {sent}")


def run_whatsapp_campaign(leads: list[dict], tpl: dict, s: dict, dry: bool, db, log_fn, stop_event, progress_fn, ask_confirm_fn, wa_lock):
    log_fn("💬 === بدء حملة WhatsApp Web ===")
    daily_cap = int(s.get("daily_wa_cap") or 20)
    min_d = int(s.get("wa_min_delay") or 50)
    max_d = int(s.get("wa_max_delay") or 95)
    batch_sz = max(1, int(s.get("batch_size") or 5))
    batch_brk = int(s.get("batch_break_sec") or 180)
    sn = s.get("sender_name", "").strip()

    today_used = db.get_today_counts()["whatsapp"]
    remaining_quota = max(0, daily_cap - today_used)
    if not dry and remaining_quota == 0:
        log_fn(f"🛡️ درع الحماية: وصلت للحد اليومي الآمن لواتساب ({today_used}/{daily_cap}). أوقفنا الإرسال تلقائياً لتجنب الحظر.")
        return

    targets = []
    for l in leads:
        if l.get("stage") == "عدم الإزعاج":
            continue
        if (l.get("wa_status") or "").startswith(("تم", "غير مسجل")):
            continue
        ph = format_whatsapp_phone(l.get("phone"))
        if ph:
            targets.append((l, ph))

    if not targets:
        log_fn("ℹ️ لا يوجد أرقام واتساب جديدة صالحة في القائمة المختارة.")
        return

    limit = len(targets) if dry else min(len(targets), remaining_quota)
    log_fn(f"📊 أرقام جاهزة: {len(targets)} | المتبقي من الحد اليومي: {remaining_quota}/{daily_cap} | وضع التجربة: {dry}")

    if dry:
        for i, (l, ph) in enumerate(targets[:limit], start=1):
            log_fn(f"🧪 [تجربة WhatsApp {i}/{limit}] {l['name']} -> +{ph}")
            progress_fn(i, limit)
        log_fn("🏁 انتهت معاينة حملة WhatsApp.")
        return

    if not wa_lock.acquire(blocking=False):
        log_fn("⚠️ متصفح WhatsApp Web مفتوح حالياً في عملية أخرى. انتظر إغلاقه أولاً.")
        return

    sent = 0
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            ctx = launch_persistent_browser(p, WA_SESSION_DIR, log_fn)
            try:
                page = ctx.pages[0] if ctx.pages else ctx.new_page()
                page.goto("https://web.whatsapp.com", wait_until="domcontentloaded", timeout=30000)
                time.sleep(4)

                is_ready = page.locator("#pane-side, div[aria-label='Chat list'], div[data-testid='chat-list']").count() > 0
                if not is_ready:
                    if not ask_confirm_fn("تأكيد WhatsApp Web", "امسح QR Code لو طُلب منك، وبعد ظهور قائمة المحادثات اضغط OK لبدء الإرسال."):
                        log_fn("تم إلغاء حملة WhatsApp.")
                        return
                else:
                    log_fn("✅ جلسة WhatsApp Web جاهزة ومتصلة تلقائياً!")
                    db.save_settings({"wa_verified": "1"})

                send_sel = "button[aria-label='Send'], button[aria-label='إرسال'], span[data-icon='send'], span[data-icon='wds-ic-send-filled']"
                inp_sel = "footer div[contenteditable='true']"
                dlg_sel = "div[data-animate-modal-popup='true'], div[role='dialog']"

                for lead, ph in targets[:limit]:
                    if stop_event.is_set():
                        log_fn("⛔ تم إيقاف حملة WhatsApp.")
                        break

                    msg = render_msg(tpl["body"], lead, sn, apply_spintax=True)
                    log_fn(f"➡️ [WhatsApp] فتح محادثة {lead['name']} (+{ph})...")
                    try:
                        # نفتح المحادثة بدون تمرير النص في الرابط ونكتبه بمحاكاة بشرية أهدأ، أو نمرره مع فحص
                        page.goto(f"https://web.whatsapp.com/send?phone={ph}", wait_until="domcontentloaded", timeout=25000)

                        state, t0 = "timeout", time.time()
                        while time.time() - t0 < 35 and not stop_event.is_set():
                            if page.locator(dlg_sel).count():
                                try:
                                    d_txt = page.locator(dlg_sel).first.inner_text().lower()
                                    if any(k in d_txt for k in ("invalid", "غير صالح", "غير مسجل", "not on whatsapp")):
                                        state = "invalid"
                                        break
                                except Exception:
                                    pass
                            if page.locator(inp_sel).count() > 0:
                                state = "ready"
                                break
                            time.sleep(1)

                        if state == "invalid":
                            db.update_channel_status(lead["id"], "wa_status", "غير مسجل")
                            log_fn(f"⚠️ الرقم +{ph} غير مسجل على واتساب — تم التخطي.")
                            continue
                        if state != "ready" or stop_event.is_set():
                            db.update_channel_status(lead["id"], "wa_status", "فشل (Timeout)")
                            log_fn(f"⚠️ انتهى وقت انتظار تحميل الشات لـ {lead['name']}.")
                            continue

                        # كتابة الرسالة بسرعة بشرية متغيرة
                        box = page.locator(inp_sel).last
                        box.click()
                        time.sleep(random.uniform(0.8, 1.8))
                        lines = msg.split("\n")
                        for li_idx, line in enumerate(lines):
                            if line:
                                page.keyboard.type(line, delay=random.randint(25, 65))
                            if li_idx < len(lines) - 1:
                                page.keyboard.down("Shift")
                                page.keyboard.press("Enter")
                                page.keyboard.up("Shift")
                        time.sleep(random.uniform(1.0, 2.2))

                        if page.locator(send_sel).count() > 0:
                            page.locator(send_sel).first.click()
                        else:
                            page.keyboard.press("Enter")
                        time.sleep(3.5)

                        now_t = datetime.now().strftime("%m-%d %H:%M")
                        db.update_channel_status(lead["id"], "wa_status", f"تم ({now_t})", promote_stage=True)
                        db.record_activity(lead["id"], "whatsapp", "sent")
                        sent += 1
                        progress_fn(sent, limit)
                        log_fn(f"✅ [WhatsApp {sent}/{limit}] تم الإرسال بنجاح إلى {lead['name']}.")

                        if sent < limit:
                            if sent % batch_sz == 0:
                                brk = random.uniform(batch_brk * 0.85, batch_brk * 1.15)
                                if _smart_wait(brk, stop_event, log_fn, f"🛡️ استراحة حماية لواتساب: انتظار {brk/60:.1f} دقيقة..."):
                                    break
                            else:
                                wt = random.uniform(min_d, max_d)
                                if _smart_wait(wt, stop_event, log_fn, f"⏳ انتظار عشوائي آمن {wt:.1f} ثانية..."):
                                    break
                    except Exception as e:
                        db.update_channel_status(lead["id"], "wa_status", f"خطأ: {str(e)[:20]}")
                        log_fn(f"❌ [WhatsApp] خطأ مع {lead['name']}: {e}")
            finally:
                try:
                    ctx.close()
                except Exception:
                    pass
    finally:
        wa_lock.release()
    log_fn(f"🏁 انتهت حملة WhatsApp. الإجمالي المرسل: {sent}")


def run_linkedin_campaign(leads: list[dict], tpl: dict, s: dict, dry: bool, db, log_fn, stop_event, progress_fn, ask_confirm_fn, li_lock):
    log_fn("💼 === بدء حملة رسائل LinkedIn ===")
    daily_cap = int(s.get("daily_li_cap") or 15)
    min_d = int(s.get("li_min_delay") or 45)
    max_d = int(s.get("li_max_delay") or 85)
    sn = s.get("sender_name", "").strip()

    today_used = db.get_today_counts()["linkedin"]
    remaining_quota = max(0, daily_cap - today_used)
    if not dry and remaining_quota == 0:
        log_fn(f"🛡️ درع الحماية: وصلت للحد اليومي الآمن لـ LinkedIn ({today_used}/{daily_cap}).")
        return

    targets = [
        l for l in leads
        if l.get("stage") != "عدم الإزعاج"
        and "linkedin.com/company" in (l.get("profile_url") or "").lower()
        and not (l.get("li_status") or "").startswith(("تم", "مغلق"))
    ]
    if not targets:
        log_fn("ℹ️ لا يوجد صفحات LinkedIn جديدة في القائمة المختارة.")
        return

    limit = len(targets) if dry else min(len(targets), remaining_quota)
    if dry:
        for i, l in enumerate(targets[:limit], start=1):
            log_fn(f"🧪 [تجربة LinkedIn {i}/{limit}] {l['name']} -> {l['profile_url']}")
            progress_fn(i, limit)
        return

    if not li_lock.acquire(blocking=False):
        log_fn("⚠️ متصفح LinkedIn مشغول حالياً بعملية السحب. يمكنك إرسال الإيميلات والواتساب الآن، أو انتظار انتهاء سحب LinkedIn.")
        return

    sent = 0
    try:
        with sync_playwright() as p:
            ctx = launch_persistent_browser(p, LINKEDIN_SESSION_DIR, log_fn)
            try:
                page = ctx.pages[0] if ctx.pages else ctx.new_page()
                page.goto("https://www.linkedin.com/feed/", wait_until="domcontentloaded", timeout=25000)
                time.sleep(2.5)
                cookies = ctx.cookies()
                if not any(c.get("name") == "li_at" and c.get("value") for c in cookies):
                    if not ask_confirm_fn("تسجيل دخول LinkedIn", "سجل دخولك في LinkedIn ثم اضغط OK للبدء."):
                        return

                for lead in targets[:limit]:
                    if stop_event.is_set():
                        break
                    msg = render_msg(tpl["body"], lead, sn, apply_spintax=True)
                    log_fn(f"➡️ [LinkedIn] فتح صفحة {lead['name']}...")
                    try:
                        page.goto(lead["profile_url"], wait_until="domcontentloaded", timeout=20000)
                        if stop_event.wait(random.uniform(3, 5)):
                            break
                        btn = page.locator("button:has-text('Message'), button:has-text('رسالة'), a:has-text('Message')").first
                        if btn.count() == 0 or not btn.is_visible():
                            db.update_channel_status(lead["id"], "li_status", "مغلق")
                            log_fn(f"ℹ️ [LinkedIn] {lead['name']}: زر الرسائل غير متاح على الصفحة.")
                            continue
                        btn.click()
                        time.sleep(2.5)
                        box = page.locator("div.msg-form__contenteditable[contenteditable='true'], div[role='textbox']").last
                        if box.count() > 0:
                            box.click()
                            page.keyboard.type(msg, delay=random.randint(25, 55))
                            time.sleep(1.5)
                            sub = page.locator("button.msg-form__send-button, button[type='submit']:has-text('Send'), button[type='submit']:has-text('إرسال')").last
                            if sub.count() > 0 and sub.is_enabled():
                                sub.click()
                                time.sleep(3)
                                now_t = datetime.now().strftime("%m-%d %H:%M")
                                db.update_channel_status(lead["id"], "li_status", f"تم ({now_t})", promote_stage=True)
                                db.record_activity(lead["id"], "linkedin", "sent")
                                sent += 1
                                progress_fn(sent, limit)
                                log_fn(f"✅ [LinkedIn {sent}/{limit}] تم الإرسال لـ {lead['name']}.")
                                if sent < limit:
                                    wt = random.uniform(min_d, max_d)
                                    if _smart_wait(wt, stop_event, log_fn, f"⏳ انتظار عشوائي {wt:.1f} ثانية..."):
                                        break
                    except Exception as e:
                        log_fn(f"❌ [LinkedIn] خطأ مع {lead['name']}: {e}")
            finally:
                try:
                    ctx.close()
                except Exception:
                    pass
    finally:
        li_lock.release()
    log_fn(f"🏁 انتهت حملة LinkedIn. الإجمالي المرسل: {sent}")
