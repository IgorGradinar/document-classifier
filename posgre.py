import psycopg2
from psycopg2.extras import RealDictCursor
from typing import List, Dict


class EmailDatabaseManager:
    def __init__(self, db_name: str = "Email", user: str = "postgres", password: str = "postgres", host: str = "localhost", port: int = 5432):

        self.conn = psycopg2.connect(
            dbname="Email",
            user="postgres",
            password="postgres",
            host="localhost",
            port=5432
        )
        self.create_table()

    def create_table(self):
        cursor = self.conn.cursor()
        # Создаём таблицу emails
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS emails (
                id SERIAL PRIMARY KEY,
                uid TEXT UNIQUE,  -- UID сообщения
                sender TEXT,
                recipient TEXT,
                subject TEXT,
                date TEXT,
                body TEXT,
                attachments TEXT
            )
        """)
        # Создаём таблицу attachments
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS attachments (
                id SERIAL PRIMARY KEY,
                email_id INTEGER REFERENCES emails(id) ON DELETE CASCADE,
                filename TEXT,
                path TEXT,
                content BYTEA,
                text TEXT,
                category INTEGER
            )
        """)
        self.conn.commit()
        cursor.close()

    def get_processed_uids(self) -> List[str]:
        cursor = self.conn.cursor()
        cursor.execute("SELECT uid FROM emails")
        uids = [row[0] for row in cursor.fetchall()]
        cursor.close()

        return uids

    def insert_email(self, email_data: Dict) -> int:
        cursor = self.conn.cursor()
        cursor.execute("""
            INSERT INTO emails (uid, sender, recipient, subject, date, body)
            VALUES (%s, %s, %s, %s, %s, %s)
            ON CONFLICT (uid) DO NOTHING
            RETURNING id
        """, (
            email_data["uid"].strip(),
            email_data["from"],
            email_data["to"],
            email_data["subject"],
            email_data["date"],
            email_data["body"]
        ))
        email_id = cursor.fetchone()
        self.conn.commit()
        cursor.close()
        return email_id[0] if email_id else None

    def insert_attachment(self, email_id: int, attachment_data: Dict):
        cursor = self.conn.cursor()
        cursor.execute("""
            INSERT INTO attachments (email_id, filename, path, content, text, category)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (
            email_id,
            attachment_data["filename"],
            attachment_data["path"],
            attachment_data["content"],
            attachment_data["text"],
            attachment_data["category"]
        ))
        self.conn.commit()
        cursor.close()

    def get_all_emails(self) -> List[Dict]:
        cursor = self.conn.cursor(cursor_factory=RealDictCursor)
        cursor.execute("SELECT sender, recipient, subject, date, body, attachments FROM emails")
        emails = cursor.fetchall()
        cursor.close()
        for email in emails:
            email["from"] = email.pop("sender") 
            email["to"] = email.pop("recipient")
            email["attachments"] = email["attachments"].split(", ") if email["attachments"] else []
        return emails

    def get_attachments_by_email_id(self, email_id: int) -> List[Dict]:
        cursor = self.conn.cursor(cursor_factory=RealDictCursor)
        cursor.execute("SELECT filename, path, text FROM attachments WHERE email_id = %s", (email_id,))
        attachments = cursor.fetchall()
        cursor.close()
        return attachments

    def update_email_attachments(self, email_id: int, attachments: List[str]):
        cursor = self.conn.cursor()
        cursor.execute("""
            UPDATE emails
            SET attachments = %s
            WHERE id = %s
        """, (attachments, email_id))
        self.conn.commit()
        cursor.close()

    def close(self):
        self.conn.close()