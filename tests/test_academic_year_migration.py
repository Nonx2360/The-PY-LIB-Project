# tests/test_academic_year_migration.py — academic-year preview + promotion
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _harness as harness  # noqa: E402

_env = harness.TempEnv()

import db.schema as schema_mod  # noqa: E402
from db.repositories.base_repo import BaseRepository  # noqa: E402
from db.repositories.member_repo import MemberRepository  # noqa: E402
from services.academic_year_service import AcademicYearService  # noqa: E402
from config import settings as settings_mod  # noqa: E402

ORIG_SETTINGS_FILE = settings_mod.SETTINGS_FILE


class TestAcademicYearMigration(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        _env.apply()
        _env.patch()
        schema_mod.init_db()
        settings_mod.SETTINGS_FILE = os.path.join(_env.dir, "settings.json")
        settings_mod.load()
        cls.base = BaseRepository()
        cls.members = MemberRepository()
        cls.svc = AcademicYearService()

    @classmethod
    def tearDownClass(cls):
        settings_mod.SETTINGS_FILE = ORIG_SETTINGS_FILE
        settings_mod.load()
        _env.restore()
        harness.cleanup(_env.dir)

    def _add_member(self, name, grade, number, status="active"):
        mid = self.members.add(name, grade, number, "2026-01-01", "2027-01-01", "qr", None)
        if status != "active":
            self.base.execute("UPDATE members SET status = ? WHERE id = ?", (status, mid))
        # ensure placement row exists for this fresh member
        self.svc.repo.ensure_current_placements()
        return mid

    def test_preview_is_read_only(self):
        mid = self._add_member("เด็ก ตัวอย่าง", "ม.1/1", "101")
        before = self.base.fetchone(
            "SELECT grade, COALESCE(status,'active') FROM members WHERE id = ?", (mid,))
        runs_before = self.base.fetchone(
            "SELECT COUNT(*) FROM academic_year_migrations")[0]
        preview = self.svc.build_preview("2568", "2569")
        self.assertTrue(any(it["member_id"] == mid for it in preview["items"]))
        after = self.base.fetchone(
            "SELECT grade, COALESCE(status,'active') FROM members WHERE id = ?", (mid,))
        runs_after = self.base.fetchone(
            "SELECT COUNT(*) FROM academic_year_migrations")[0]
        self.assertEqual(before, after)
        self.assertEqual(runs_before, runs_after)

    def test_execute_promotes_and_preserves_history(self):
        mid = self._add_member("เลื่อนชั้น ดี", "ม.2/3", "202")
        bid = self.base.execute(
            "INSERT INTO books (code, title, status) VALUES ('AYM-1','B','ว่าง')")
        self.base.execute(
            "INSERT INTO borrow_log (member_id, book_id, borrow_date, return_due, returned) "
            "VALUES (?, ?, '2026-05-01', '2026-05-08', 1)", (mid, bid))
        preview = self.svc.build_preview("2568", "2569")
        reviewed = []
        for it in preview["items"]:
            row = dict(it)
            row["selected"] = (it["member_id"] == mid)
            reviewed.append(row)
        result = self.svc.apply_preview(
            "2568", "2569", "tester", reviewed, force=False, do_backup=False)
        self.assertGreater(result["run_id"], 0)
        self.assertEqual(
            self.base.fetchone("SELECT grade FROM members WHERE id = ?", (mid,))[0],
            "ม.3/3")
        # stable member id and historical loan row untouched
        loan = self.base.fetchone(
            "SELECT member_id, book_id, borrow_date, returned FROM borrow_log "
            "WHERE member_id = ?", (mid,))
        self.assertEqual(loan, (mid, bid, "2026-05-01", 1))
        placement = self.base.fetchone(
            "SELECT grade, status FROM member_placements "
            "WHERE member_id = ? AND status = 'current'", (mid,))
        self.assertEqual(placement, ("ม.3/3", "current"))

    def test_graduate_kept_and_manual(self):
        m_grad = self._add_member("จบ แล้ว", "ม.6/1", "301")
        m_left = self._add_member("ลา ออก", "ม.4/2", "302", status="left")
        m_man = self._add_member("ตรวจ เอง", "ห้องพิเศษ", "303")
        preview = self.svc.build_preview("2572", "2573")
        reviewed = []
        for it in preview["items"]:
            row = dict(it)
            if it["member_id"] == m_grad:
                row["action"] = "graduate"
                row["selected"] = True
            elif it["member_id"] == m_left:
                row["action"] = "keep"
                row["selected"] = True
            elif it["member_id"] == m_man:
                row["action"] = "manual"
                row["selected"] = True
            else:
                row["selected"] = False
            reviewed.append(row)
        res = self.svc.apply_preview(
            "2572", "2573", "tester", reviewed, force=False, do_backup=False)
        self.assertEqual(
            self.base.fetchone("SELECT status FROM members WHERE id = ?", (m_grad,))[0],
            "graduated")
        # left member must not be promoted; kept as-is
        self.assertEqual(
            self.base.fetchone("SELECT grade, status FROM members WHERE id = ?", (m_left,)),
            ("ม.4/2", "left"))
        self.assertEqual(
            self.base.fetchone("SELECT grade FROM members WHERE id = ?", (m_man,))[0],
            "ห้องพิเศษ")
        self.assertGreaterEqual(res["summary"].get("graduate", 0), 1)

    def test_duplicate_run_detected(self):
        mid = self._add_member("ซ้ำ ชุด", "ม.1/2", "401")
        preview = self.svc.build_preview("2570", "2571")
        reviewed = [dict(it, selected=(it["member_id"] == mid)) for it in preview["items"]]
        self.svc.apply_preview("2570", "2571", "tester", reviewed, do_backup=False)
        with self.assertRaises(ValueError):
            self.svc.apply_preview("2570", "2571", "tester", reviewed, do_backup=False)
        # forced repeat allowed (intentional) and recorded
        res = self.svc.apply_preview("2570", "2571", "tester", reviewed, force=True, do_backup=False)
        self.assertGreater(res["run_id"], 0)


if __name__ == "__main__":
    unittest.main()
