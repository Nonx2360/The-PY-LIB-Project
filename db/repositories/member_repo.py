# db/repositories/member_repo.py
from .base_repo import BaseRepository

class MemberRepository(BaseRepository):
    def get_all(self):
        return self.fetchall("SELECT * FROM members")
        
    def get_by_id(self, member_id):
        return self.fetchone("SELECT * FROM members WHERE id = ?", (member_id,))
        
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
