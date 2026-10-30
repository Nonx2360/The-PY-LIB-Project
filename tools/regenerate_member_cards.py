# tools/regenerate_member_cards.py — rebuild member card PDFs from the live DB
#
# Card PDFs are fully derivable from member rows (name, number, sealed QR),
# so this tool can repair lost/overwritten/low-quality card files without
# touching any member data. Also optionaly prunes assets/qrcodes of files
# no member row references (test artifacts, stale regenerations).
#
# Usage:
#     python tools/regenerate_member_cards.py            # cards only
#     python tools/regenerate_member_cards.py --clean-orphans --yes
#
# NOTE: paths are CWD-relative, matching the app (run from the project root).
import argparse
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

from config.constants import DB_PATH  # noqa: E402
from db.repositories.member_repo import MemberRepository  # noqa: E402
from services.pdf_service import PDFService  # noqa: E402


def _norm(path):
    return os.path.normpath(path).replace("\\", "/")


def regenerate_cards():
    repo = MemberRepository()
    rows = repo.fetchall(
        "SELECT id, name, number, qrcode_data FROM members ORDER BY id")
    if not rows:
        print("db has no members - nothing to regenerate")
        return 0
    done, skipped = 0, []
    for mid, _name, number, sealed in rows:
        if not sealed:
            skipped.append(mid)
            continue
        target = f"assets/cards/member_{mid}.pdf"
        os.makedirs("assets/cards", exist_ok=True)
        # same production call shape as web/desktop card generation; the card
        # is recomputed from the DB, never from the old (possibly bad) file
        PDFService.generate_member_card_pdf(
            _name, number, "assets/logos/school_logo.png", target, sealed)
        print(f"member id={mid}: card -> {target}")
        done += 1
    if skipped:
        print(f"skipped members without sealed QR data: {skipped}")
    print(f"regenerated {done} card(s); db={DB_PATH}")
    return done


def clean_orphans():
    repo = MemberRepository()
    rows = repo.fetchall(
        "SELECT qrcode_path FROM members WHERE qrcode_path IS NOT NULL")
    referenced = {_norm(r[0]) for r in rows}
    qr_dir = "assets/qrcodes"
    if not os.path.isdir(qr_dir):
        return
    removed = 0
    for entry in sorted(os.listdir(qr_dir)):
        path = os.path.join(qr_dir, entry)
        if not os.path.isfile(path):
            continue  # keep subfolders (none expected)
        if _norm(path) in referenced:
            continue
        os.remove(path)
        print(f"removed orphan: {_norm(path)}")
        removed += 1
    print(f"orphan cleanup done: {removed} removed, {len(referenced)} kept (referenced)")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--clean-orphans", action="store_true",
                    help="also delete assets/qrcodes files no member references")
    ap.add_argument("--yes", action="store_true",
                    help="confirm the actions without an interactive prompt")
    args = ap.parse_args()

    if not os.path.exists(DB_PATH):
        print(f"database not found: {DB_PATH}")
        return sys.exit(1)

    if not args.yes:
        answer = input(
            f"Regenerate card PDFs for all members in {DB_PATH}"
            + (" and delete unreferenced QR files" if args.clean_orphans else "")
            + "? type yes to continue: ").strip().lower()
        if answer != "yes":
            print("aborted - nothing changed")
            return sys.exit(2)

    done = regenerate_cards()
    if args.clean_orphans:
        clean_orphans()
    print(f"summary: {done} card(s) regenerated from db "
          f"({os.path.basename(DB_PATH)})")


if __name__ == "__main__":
    main()
