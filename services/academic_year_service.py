# services/academic_year_service.py
import csv
import io
import re
from datetime import datetime

from db.repositories.academic_year_repo import AcademicYearRepository
from services.backup_service import create_backup, verify_backup


class AcademicYearService:
    """Preview/build/execute academic-year promotions using shared repo logic."""

    def __init__(self, repo=None):
        self.repo = repo or AcademicYearRepository()

    # ---------------- preview (read-only) ----------------
    @staticmethod
    def _suggest_grade(grade):
        """Best-effort Thai/common grade progression; admin can override."""
        g = (grade or "").strip()
        m = re.match(r"^(ม\.?\s*)([1-6])([/\-].+)?$", g)
        if m:
            n = int(m.group(2))
            if n < 6:
                return f"ม.{n + 1}{m.group(3) or ''}"
            return None  # graduating class by default
        m = re.match(r"^(ป\.?\s*)([1-6])([/\-].+)?$", g)
        if m:
            n = int(m.group(2))
            if n < 6:
                return f"ป.{n + 1}{m.group(3) or ''}"
            return "ม.1"  # common progression; admin may override
        m = re.match(r"^([1-6])([/\-].+)$", g)
        if m:
            n = int(m.group(1))
            if n < 6:
                return f"{n + 1}{m.group(2)}"
            return None
        return None

    def build_preview(self, source_label, target_label):
        """Return a read-only preview plan. No writes, no side effects."""
        self.repo.ensure_current_placements()
        rows = self.repo.members_for_migration()
        items = []
        counts = {"promote": 0, "graduate": 0, "left": 0, "keep": 0, "manual": 0}
        for (mid, name, grade, number, _reg, exp, mstatus,
             _pgrade, _pnumber, pyear) in rows:
            cur_grade = grade or ""
            cur_number = number or ""
            suggested = self._suggest_grade(cur_grade)
            reason = ""
            if mstatus in ("graduated", "left"):
                action = "keep"
                proposed_grade = cur_grade
                reason = f"สถานะปัจจุบันคือ {mstatus} — ไม่เลื่อนชั้นอัตโนมัติ"
            elif mstatus != "active":
                action = "manual"
                proposed_grade = cur_grade
                reason = f"สถานะ '{mstatus}' ต้องตรวจสอบ"
            elif not cur_grade:
                action = "manual"
                proposed_grade = cur_grade
                reason = "ไม่มีข้อมูลชั้น/ห้อง"
            elif suggested is None:
                action = "graduate"
                proposed_grade = cur_grade
                reason = "ระดับชั้นสูงสุด/จบ — เสนอเป็นจบการศึกษา (ตรวจสอบได้)"
            else:
                action = "promote"
                proposed_grade = suggested
                reason = "เสนอเลื่อนชั้นอัตโนมัติ"

            before = {"id": mid, "name": name, "grade": cur_grade,
                      "number": cur_number, "status": mstatus,
                      "expire_date": exp, "placement_year": pyear}
            after = dict(before)
            after["grade"] = proposed_grade
            if action == "graduate":
                after["status"] = "graduated"
            elif action == "left":
                after["status"] = "left"
            elif action == "promote":
                after["status"] = "active"
            # keep/manual leave status/grade as-is unless overridden

            counts[action] = counts.get(action, 0) + 1
            items.append({
                "member_id": mid,
                "name": name or "",
                "current_grade": cur_grade,
                "current_number": cur_number,
                "member_status": mstatus,
                "proposed_grade": proposed_grade,
                "proposed_number": cur_number,
                "action": action,
                "reason": reason,
                "before": before,
                "after": after,
                "selected": action != "manual",
            })

        duplicate = bool(self.repo.find_successful_run(source_label, target_label)) \
            if source_label and target_label else False
        return {
            "source_label": source_label or "",
            "target_label": target_label or "",
            "items": items,
            "counts": counts,
            "total": len(items),
            "duplicate_success": duplicate,
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

    @staticmethod
    def preview_to_csv(preview):
        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow([
            "member_id", "name", "current_grade", "current_number",
            "member_status", "proposed_grade", "proposed_action", "reason", "selected"])
        for it in preview.get("items", []):
            writer.writerow([
                it["member_id"], it["name"], it["current_grade"], it["current_number"],
                it["member_status"], it["proposed_grade"], it["action"],
                it["reason"], "yes" if it.get("selected") else "no"])
        return "\ufeff" + buf.getvalue()  # BOM for Excel/Thai

    # ---------------- execute (writes) ----------------
    def apply_preview(self, source_label, target_label, actor, items,
                      force=False, do_backup=True):
        """Execute reviewed items. Items come from the client as reviewed rows:
        {member_id, action, proposed_grade, proposed_number, selected, reason}.
        A verified backup is taken before mutation when do_backup=True.
        """
        plans = []
        selected = 0
        for it in (items or []):
            mid = int(it.get("member_id"))
            action = (it.get("action") or "manual").strip()
            if not it.get("selected", False) and action != "manual":
                action = "keep"
                reason = (it.get("reason") or "") + " (ผู้ดูแลไม่เลือกดำเนินการ)"
            else:
                reason = it.get("reason") or ""
                selected += 1 if action != "manual" else 0
            proposed_grade = str(it.get("proposed_grade") or it.get("current_grade") or "")
            proposed_number = str(it.get("proposed_number") or it.get("current_number") or "")
            before = it.get("before") or {
                "id": mid, "name": it.get("name"), "grade": it.get("current_grade"),
                "number": it.get("current_number"), "status": it.get("member_status"),
            }
            after = dict(before)
            after["grade"] = proposed_grade
            after["number"] = proposed_number
            if action == "graduate":
                after["status"] = "graduated"
            elif action == "left":
                after["status"] = "left"
            elif action == "promote":
                after["status"] = "active"
            plans.append({
                "member_id": mid, "action": action, "reason": reason.strip(),
                "before": before, "after": after,
            })

        if not plans:
            raise ValueError("ไม่มีรายการให้ดำเนินการ")

        backup_path = None
        if do_backup:
            backup_path = create_backup()
            report = verify_backup(backup_path)
            if not report.get("ok"):
                raise ValueError("สำรองฐานข้อมูลก่อนย้ายปีการศึกษาไม่ผ่านการตรวจสอบ")

        run_id, summary = self.repo.execute_run(
            source_label, target_label, actor, plans, force=force)
        return {
            "run_id": run_id,
            "summary": summary,
            "backup_path": backup_path,
            "total_planned": len(plans),
        }
