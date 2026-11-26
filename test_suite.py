import os
import tempfile
import threading
import unittest

from crm_core import (
    DatabaseManager,
    resolve_spintax,
    render_msg,
    extract_mobile_number,
    format_whatsapp_phone,
    extract_email,
    is_valid_email,
    matches_icp_and_contact,
    export_leads_to_excel,
    import_leads_from_excel,
)


class TestSmartCRMCore(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.tmp_dir.name, "test_crm.db")
        self.db = DatabaseManager(self.db_path)

    def tearDown(self):
        self.tmp_dir.cleanup()

    def test_spintax_and_template_rendering(self):
        tpl = "{السلام عليكم|أهلاً} {name} في {location} — {شراكة|تعاون} ({unknown_var})"
        lead = {"name": "وكالة المنصورة", "location": "المنصورة"}
        outputs = {render_msg(tpl, lead, "أحمد", apply_spintax=True) for _ in range(30)}
        # Should vary due to spintax and preserve safe unknown placeholders
        self.assertGreater(len(outputs), 1)
        for out in outputs:
            self.assertIn("وكالة المنصورة", out)
            self.assertIn("المنصورة", out)
            self.assertIn("{unknown_var}", out)
            self.assertNotIn("|", out)

    def test_phone_normalization_and_landline_rejection(self):
        self.assertEqual(format_whatsapp_phone("01012345678"), "201012345678")
        self.assertEqual(format_whatsapp_phone("+20 11 2345-6789"), "201123456789")
        self.assertEqual(format_whatsapp_phone("00201555555555"), "201555555555")
        self.assertEqual(format_whatsapp_phone("+966 50 123 4567"), "966501234567")
        # Egyptian landlines (Mansoura 050, Cairo 02) must be rejected
        self.assertIsNone(format_whatsapp_phone("0502233445"))
        self.assertIsNone(format_whatsapp_phone("+20 50 2233445"))
        self.assertIsNone(format_whatsapp_phone("0223456789"))
        self.assertIsNone(format_whatsapp_phone("غير متوفر"))

    def test_email_extraction_and_validation(self):
        html = "Contact us at sales@mansoura-agency.com or ignore logo@2x.png and user@sentry.io"
        self.assertEqual(extract_email(html), "sales@mansoura-agency.com")
        self.assertTrue(is_valid_email("hello@company.co"))
        self.assertFalse(is_valid_email("غير متوفر"))
        self.assertFalse(is_valid_email("bad email@test.com"))

    def test_icp_and_contact_filters(self):
        lead = {
            "name": "Delta Digital Agency",
            "source": "Google Maps",
            "category": "Marketing Agency",
            "size_rating": "★ 4.6",
            "website": "https://delta.eg",
            "phone": "201012345678",
            "email": "info@delta.eg",
        }
        cfg_ok = {
            "contact_req": "الاثنين معاً (واتساب + إيميل)",
            "must_have_web": True,
            "icp_industry": "marketing, software",
            "icp_size": "الكل",
            "min_rating": "4.0",
        }
        ok, _ = matches_icp_and_contact(lead, cfg_ok)
        self.assertTrue(ok)

        # Reject if rating below threshold
        cfg_high_rating = {**cfg_ok, "min_rating": "4.9"}
        ok2, reason = matches_icp_and_contact(lead, cfg_high_rating)
        self.assertFalse(ok2)
        self.assertIn("التقييم", reason)

        # Reject if missing email when email required
        lead_no_email = {**lead, "email": ""}
        ok3, _ = matches_icp_and_contact(lead_no_email, cfg_ok)
        self.assertFalse(ok3)

    def test_concurrent_scraping_and_sending_db_wal(self):
        """اختبار التزامن الحقيقي: ثريد يسحب ويضيف عملاء بينما ثريد آخر يرسل ويحدث الحالات بدون قفل DB"""
        errors = []

        def scraper_sim():
            try:
                for i in range(40):
                    self.db.add_lead({
                        "name": f"Agency {i}",
                        "source": "Google Maps",
                        "phone": f"2010000000{i:02d}",
                        "email": f"agency{i}@test.com",
                        "location": "Mansoura",
                    })
            except Exception as e:
                errors.append(e)

        def sender_sim():
            try:
                for _ in range(30):
                    leads = self.db.get_leads()
                    for l in leads[:5]:
                        self.db.update_channel_status(l["id"], "email_status", "تم (12:00)", promote_stage=True)
                        self.db.record_activity(l["id"], "email", "sent")
            except Exception as e:
                errors.append(e)

        t1 = threading.Thread(target=scraper_sim)
        t2 = threading.Thread(target=sender_sim)
        t1.start()
        t2.start()
        t1.join()
        t2.join()

        self.assertEqual(errors, [], f"Concurrency errors occurred: {errors}")
        self.assertEqual(len(self.db.get_leads()), 40)

    def test_daily_quota_tracking_and_dnc_filter(self):
        lid1 = self.db.add_lead({"name": "Active Co", "email": "a@co.com", "phone": "201011111111"})
        lid2 = self.db.add_lead({"name": "DNC Co", "email": "b@co.com", "phone": "201022222222", "stage": "عدم الإزعاج"})
        self.db.record_activity(lid1, "email", "sent")
        self.db.record_activity(lid1, "whatsapp", "sent")
        counts = self.db.get_today_counts()
        self.assertEqual(counts["email"], 1)
        self.assertEqual(counts["whatsapp"], 1)
        self.assertEqual(counts["linkedin"], 0)

        eligible_em = self.db.get_leads(quick_filter="بها إيميل ولم تُرسل")
        names = [l["name"] for l in eligible_em]
        self.assertIn("Active Co", names)
        self.assertNotIn("DNC Co", names)

    def test_excel_export_and_reimport_roundtrip(self):
        self.db.add_lead({
            "name": "Nile Tech",
            "source": "LinkedIn",
            "phone": "201234567890",
            "email": "contact@niletech.eg",
            "category": "Software",
            "location": "Mansoura",
            "profile_url": "https://www.linkedin.com/company/nile-tech",
        })
        xlsx_path = os.path.join(self.tmp_dir.name, "export.xlsx")
        leads = self.db.get_leads()
        export_leads_to_excel(leads, xlsx_path, "مرحباً {name}", "السلام عليكم {name}", "فريق المبيعات")
        self.assertTrue(os.path.isfile(xlsx_path))

        # Re-import into a fresh DB
        db2 = DatabaseManager(os.path.join(self.tmp_dir.name, "db2.db"))
        added = import_leads_from_excel(db2, xlsx_path)
        self.assertEqual(added, 1)
        # Re-importing again must skip duplicates
        added_again = import_leads_from_excel(db2, xlsx_path)
        self.assertEqual(added_again, 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
