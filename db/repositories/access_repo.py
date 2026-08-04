# db/repositories/access_repo.py
from .base_repo import BaseRepository

class AccessRepository(BaseRepository):
    def log_access(self, member_id, access_time, action):
        query = "INSERT INTO access_log (member_id, access_time, action) VALUES (?, ?, ?)"
        return self.execute(query, (member_id, access_time, action))
        
    def get_recent_access(self, limit=50):
        query = '''
            SELECT al.access_time, m.name, m.grade, al.action
            FROM access_log al
            JOIN members m ON al.member_id = m.id
            ORDER BY al.id DESC LIMIT ?
        '''
        return self.fetchall(query, (limit,))
