# db/repositories/member_repo.py
from .base_repo import BaseRepository

class MemberRepository(BaseRepository):
    def get_all(self):
        return self.fetchall("SELECT * FROM members")

    def get_by_id(self, member_id):
        return self.fetchone("SELECT * FROM members WHERE id = ?", (member_id,))

    def find_by_badge(self, grade, number):
        """Row registered with the same class/number badge, else None.

        เลขที่ (class attendance number) is unique within one grade but can
        repeat across grades, so a member's badge key is (grade, number).
        QR badge matching relies on that key being unique — see
        web/app.py scan_resolve().
        """
        return self.fetchone(
            "SELECT id FROM members WHERE grade = ? AND number = ? LIMIT 1",
            (grade, number))

    def add(self, name, grade, number, reg_date, exp_date, qr_data, qr_path):
        query = '''
            INSERT INTO members (name, grade, number, register_date, expire_date, qrcode_data, qrcode_path)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        '''
        return self.execute(query, (name, grade, number, reg_date, exp_date, qr_data, qr_path))

    def update(self, member_id, name, grade, number, exp_date, qr_data, qr_path):
        query = '''
            UPDATE members
            SET name=?, grade=?, number=?, expire_date=?, qrcode_data=?, qrcode_path=?
            WHERE id=?
        '''
        self.execute(query, (name, grade, number, exp_date, qr_data, qr_path, member_id))

    def delete(self, member_id):
        self.execute("DELETE FROM members WHERE id = ?", (member_id,))
