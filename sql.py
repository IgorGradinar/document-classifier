import sqlite3
from typing import List, Dict

class EmailDatabase:
    def __init__(self, db_name: str = "emails.db"):
        self.conn = sqlite3.connect(db_name)
        self.create_table()
    
    def create_table(self):
        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sender TEXT,
                recipient TEXT,
                subject TEXT,
                date TEXT,
                body TEXT,
                attachments TEXT
            )
        """)
        self.conn.commit()
    
    def insert_email(self, email_data: Dict):
        cursor = self.conn.cursor()
        cursor.execute("""
            INSERT INTO emails (sender, recipient, subject, date, body, attachments)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            email_data["from"],
            email_data["to"],
            email_data["subject"],
            email_data["date"],
            email_data["body"],
            ", ".join(email_data["attachments"]) if email_data["attachments"] else ""
        ))
        self.conn.commit()
    
    def get_all_emails(self) -> List[Dict]:
        cursor = self.conn.cursor()
        cursor.execute("SELECT sender, recipient, subject, date, body, attachments FROM emails")
        emails = []
        for row in cursor.fetchall():
            emails.append({
                "from": row[0],
                "to": row[1],
                "subject": row[2],
                "date": row[3],
                "body": row[4],
                "attachments": row[5].split(", ") if row[5] else []
            })
        return emails
    
    def close(self):
        self.conn.close()
