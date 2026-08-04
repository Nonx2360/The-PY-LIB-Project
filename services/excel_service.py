# services/excel_service.py
import pandas as pd

class ExcelService:
    @staticmethod
    def generate_template_df() -> pd.DataFrame:
        sample_data = {
            'รหัสหนังสือ': ['001', '002'],
            'ชื่อเรื่อง': ['ตัวอย่างหนังสือ 1', 'ตัวอย่างหนังสือ 2']
        }
        return pd.DataFrame(sample_data)
        
    @staticmethod
    def read_books_from_excel(filename: str) -> list[dict]:
        df = pd.read_excel(filename)
        required_columns = ['รหัสหนังสือ', 'ชื่อเรื่อง']
        if not all(col in df.columns for col in required_columns):
            raise ValueError("รูปแบบไฟล์ Excel ไม่ถูกต้อง กรุณาใช้แม่แบบที่กำหนด")
            
        books = []
        for _, row in df.iterrows():
            books.append({
                'code': str(row['รหัสหนังสือ']).strip(),
                'title': str(row['ชื่อเรื่อง']).strip()
            })
        return books
