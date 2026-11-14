# db/repositories/academic_year_repo.py
import json
from datetime import datetime

from .base_repo import BaseRepository


class AcademicYearRepository(BaseRepository):
    def now(self):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ---------------- academic years ----------------
    def list_years(self):
        return self.fetchall(
            "SELECT id, label, start_date, end_date, is_active, created_at "
            "FROM academic_years ORDER BY id DESC")

    def get_year(self, year_id):
        return self.fetchone(
            "SELECT id, label, start_date, end_date, is_active, created_at "
            "FROM academic_years WHERE id = ?", (year_id,))

    def get_year_by_label(self, label):
        return self.fetchone(
            "SELECT id, label, start_date, end_date, is_active, created_at "
            "FROM academic_years WHERE label = ?", (label,))

    def ensure_year(self, label, start_date=None, end_date=None, active=False):
        label = (label or "").strip()
        if not label:
            raise ValueError("ต้องระบุชื่อปีการศึกษา")
        row = self.get_year_by_label(label)
        if row:
            return row[0]
        return self.execute(
            "INSERT INTO academic_years (label, start_date, end_date, is_active, created_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (label, start_date or None, end_date or None, 1 if active else 0, self.now()))

    # ---------------- placements / members snapshot ----------------
    def members_for_migration(self):
        """Current member rows with their current placement (if any).

        Stable member.id is preserved; grade/number are the mutable current
        badge fields and are the source of truth for the preview.
        """
        return self.fetchall(
            "SELECT m.id, m.name, m.grade, m.number, m.register_date, "
            "m.expire_date, COALESCE(m.status,'active'), "
            "mp.grade, mp.number, ay.label "
            "FROM members m "
            "LEFT JOIN member_placements mp "
            "  ON mp.member_id = m.id AND mp.status = 'current' "
            "LEFT JOIN academic_years ay ON ay.id = mp.academic_year_id "
            "ORDER BY m.grade, m.number, m.name, m.id")

    def ensure_current_placements(self):
        """Idempotent safety net: every member must have a current placement."""
        baseline = self.ensure_year("ปัจจุบัน (ก่อนระบบปีการศึกษา)", active=True)
        now = self.now()

        def _fix(conn):
            rows = conn.execute(
                "SELECT m.id, m.grade, m.number FROM members m "
                "WHERE NOT EXISTS (SELECT 1 FROM member_placements mp "
                "WHERE mp.member_id = m.id AND mp.status = 'current')").fetchall()
            for mid, grade, number in rows:
                conn.execute(
                    "INSERT INTO member_placements "
                    "(member_id, academic_year_id, grade, number, status, source, created_at) "
                    "VALUES (?, ?, ?, ?, 'current', 'ensure_current', ?)",
                    (mid, baseline, grade, number, now))
            return len(rows)
        return self.transact(_fix)

    # ---------------- migration runs ----------------
    def find_successful_run(self, source_label, target_label):
        return self.fetchone(
            "SELECT id, status, created_at, finished_at, summary_json "
            "FROM academic_year_migrations "
            "WHERE source_label = ? AND target_label = ? AND status = 'success' "
            "ORDER BY id DESC LIMIT 1",
            (source_label, target_label))

    def recent_runs(self, limit=20):
        rows = self.fetchall(
            "SELECT id, source_label, target_label, actor, status, summary_json, "
            "created_at, finished_at, notes FROM academic_year_migrations "
            "ORDER BY id DESC LIMIT ?", (limit,))
        out = []
        for r in rows:
            summary = {}
            try:
                summary = json.loads(r[5] or "{}")
            except Exception:
                summary = {}
            out.append({
                "id": r[0], "source_label": r[1], "target_label": r[2],
                "actor": r[3], "status": r[4], "summary": summary,
                "created_at": r[6], "finished_at": r[7], "notes": r[8],
            })
        return out

    def get_run(self, run_id):
        row = self.fetchone(
            "SELECT id, source_label, target_label, actor, status, summary_json, "
            "created_at, finished_at, notes FROM academic_year_migrations WHERE id = ?",
            (run_id,))
        if not row:
            return None
        try:
            summary = json.loads(row[5] or "{}")
        except Exception:
            summary = {}
        return {
            "id": row[0], "source_label": row[1], "target_label": row[2],
            "actor": row[3], "status": row[4], "summary": summary,
            "created_at": row[6], "finished_at": row[7], "notes": row[8],
        }

    def get_run_items(self, run_id):
        return self.fetchall(
            "SELECT member_id, action, reason, before_json, after_json, outcome, created_at "
            "FROM academic_year_migration_items WHERE migration_id = ? ORDER BY id",
            (run_id,))

    def execute_run(self, source_label, target_label, actor, plans, force=False):
        """Apply a reviewed plan atomically. `plans` is a list of dicts:
        {member_id, action, reason, before, after} where action is one of
        promote|keep|graduate|left|manual. Returns (run_id, summary).

        `before`/`after` are snapshots; for promote/graduate/left/keep the
        AFTER snapshot drives the member/placement write (except manual).
        """
        source_label = (source_label or "").strip()
        target_label = (target_label or "").strip()
        actor = (actor or "system").strip()
        if not source_label or not target_label:
            raise ValueError("ต้องระบุปีการศึกษาต้นทางและปลายทางให้ชัดเจน")
        if source_label == target_label:
            raise ValueError("ปีการศึกษาต้นทางและปลายทางต้องไม่ซ้ำกัน")
        if not force and self.find_successful_run(source_label, target_label):
            raise ValueError(
                "มีการย้ายปีการศึกษาชุดนี้สำเร็จไปแล้ว — หากต้องการทำซ้ำ "
                "ให้ยืนยันการทำซ้ำอย่างชัดเจน")

        now = self.now()

        def _apply(conn):
            source_year_id = self._ensure_year_conn(conn, source_label, now)
            target_year_id = self._ensure_year_conn(conn, target_label, now, active=True)
            cur = conn.execute(
                "INSERT INTO academic_year_migrations "
                "(source_year_id, target_year_id, source_label, target_label, actor, status, "
                "summary_json, created_at, notes) VALUES (?,?,?,?,?, 'running', NULL, ?, NULL)",
                (source_year_id, target_year_id, source_label, target_label, actor, now))
            run_id = cur.lastrowid

            summary = {"promoted": 0, "graduate": 0, "left": 0,
                       "kept": 0, "manual": 0, "skipped": 0, "failed": 0}
            for plan in plans:
                mid = int(plan["member_id"])
                action = (plan.get("action") or "manual").strip()
                reason = (plan.get("reason") or "").strip() or None
                before = plan.get("before") or {}
                after = plan.get("after") or {}
                outcome = "applied"
                try:
                    row = conn.execute(
                        "SELECT id, name, grade, number, expire_date, "
                        "COALESCE(status,'active') FROM members WHERE id = ?",
                        (mid,)).fetchone()
                    if not row:
                        outcome = "skipped_missing"
                        summary["skipped"] += 1
                    elif action == "manual":
                        outcome = "manual_review"
                        summary["manual"] += 1
                    elif action == "promote":
                        new_grade = str(after.get("grade") or row[2] or "").strip()
                        new_number = str(after.get("number") if after.get("number") is not None else row[3] or "").strip()
                        if not new_grade:
                            raise ValueError("ชั้น/ห้องปลายทางว่าง")
                        conn.execute(
                            "UPDATE members SET grade = ?, number = ?, status = 'active' WHERE id = ?",
                            (new_grade, new_number, mid))
                        self._close_current_placement(conn, mid, 'promoted', now)
                        conn.execute(
                            "INSERT INTO member_placements "
                            "(member_id, academic_year_id, grade, number, status, source, created_at) "
                            "VALUES (?, ?, ?, ?, 'current', 'migration_promote', ?)",
                            (mid, target_year_id, new_grade, new_number, now))
                        summary["promoted"] += 1
                    elif action in ("graduate", "left"):
                        status_val = 'graduated' if action == 'graduate' else 'left'
                        conn.execute(
                            "UPDATE members SET status = ? WHERE id = ?",
                            (status_val, mid))
                        self._close_current_placement(conn, mid, status_val, now)
                        conn.execute(
                            "INSERT INTO member_placements "
                            "(member_id, academic_year_id, grade, number, status, source, created_at) "
                            "VALUES (?, ?, ?, ?, 'current', ?, ?)",
                            (mid, target_year_id, row[2], row[3],
                             f"migration_{status_val}", now))
                        summary[action] += 1
                    elif action == "keep":
                        # keep member record as-is but record placement in target year
                        self._close_current_placement(conn, mid, 'kept', now)
                        conn.execute(
                            "INSERT INTO member_placements "
                            "(member_id, academic_year_id, grade, number, status, source, created_at) "
                            "VALUES (?, ?, ?, ?, 'current', 'migration_keep', ?)",
                            (mid, target_year_id, row[2], row[3], now))
                        summary["kept"] += 1
                    else:
                        outcome = "skipped_unknown"
                        summary["skipped"] += 1
                except Exception as exc:
                    outcome = f"failed:{exc}"
                    summary["failed"] += 1
                    # fail the whole run on unexpected per-row write errors to
                    # avoid partial promotions; caller gets rollback
                    raise

                conn.execute(
                    "INSERT INTO academic_year_migration_items "
                    "(migration_id, member_id, action, reason, before_json, after_json, outcome, created_at) "
                    "VALUES (?,?,?,?,?,?,?,?)",
                    (run_id, mid, action, reason,
                     json.dumps(before, ensure_ascii=False),
                     json.dumps(after, ensure_ascii=False),
                     outcome, now))

            conn.execute(
                "UPDATE academic_year_migrations SET status = 'success', "
                "summary_json = ?, finished_at = ? WHERE id = ?",
                (json.dumps(summary, ensure_ascii=False), self.now(), run_id))
            return run_id, summary

        return self.transact(_apply)

    @staticmethod
    def _ensure_year_conn(conn, label, now, active=False):
        row = conn.execute(
            "SELECT id FROM academic_years WHERE label = ?", (label,)).fetchone()
        if row:
            if active:
                conn.execute("UPDATE academic_years SET is_active = 0")
                conn.execute("UPDATE academic_years SET is_active = 1 WHERE id = ?", (row[0],))
            return row[0]
        if active:
            conn.execute("UPDATE academic_years SET is_active = 0")
        cur = conn.execute(
            "INSERT INTO academic_years (label, start_date, end_date, is_active, created_at) "
            "VALUES (?, NULL, NULL, ?, ?)",
            (label, 1 if active else 0, now))
        return cur.lastrowid

    @staticmethod
    def _close_current_placement(conn, member_id, to_status, now):
        conn.execute(
            "UPDATE member_placements SET status = ? "
            "WHERE member_id = ? AND status = 'current'",
            (to_status, member_id))  # partial unique index allows only one 'current'
