import tkinter as tk
from tkinter import ttk
from mail import EmailFetcherService
from posgre import EmailDatabaseManager
from tkinter import messagebox
import os
from PyPDF2 import PdfReader
from typing import List, Dict  # Добавьте этот импорт
from psycopg2.extras import RealDictCursor  # Добавьте этот импорт
import mammoth
ATTACHMENTS_DIR = "attachments"
if not os.path.exists(ATTACHMENTS_DIR):
    os.makedirs(ATTACHMENTS_DIR)

class EmailMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Monitor")
        self.emails = []
        self.db = EmailDatabaseManager(
            db_name="Email",
            user="postgres",
            password="postgres",
            host="localhost",
            port=5432
        )
        self.setup_ui()
        self.load_emails_from_db()
        
        # Запускаем автоматическую проверку новых сообщений
        self.auto_fetch_emails()
    
    def setup_ui(self):

        # Connection frame
        connection_frame = ttk.Frame(self.root, padding="10")
        connection_frame.grid(row=0, column=0, sticky="ew")
        
        ttk.Label(connection_frame, text="IMAP Server:").grid(row=0, column=0)
        self.imap_server = ttk.Entry(connection_frame)
        self.imap_server.grid(row=0, column=1, padx=5)
        self.imap_server.insert(0, "imap.mail.ru")  # Значение по умолчанию для IMAP Server
        
        ttk.Label(connection_frame, text="Username:").grid(row=1, column=0)
        self.username = ttk.Entry(connection_frame)
        self.username.grid(row=1, column=1, padx=5)
        self.username.insert(0, "testpy221@mail.ru")  # Значение по умолчанию для Username
        
        ttk.Label(connection_frame, text="Password:").grid(row=2, column=0)
        self.password = ttk.Entry(connection_frame, show="*")
        self.password.grid(row=2, column=1, padx=5)
        self.password.insert(0, "82WSPym0sdPvSqkgQsTn")  # Значение по умолчанию для Password
        
        self.connect_button = ttk.Button(
            connection_frame, 
            text="Connect", 
            command=self.connect_to_email
        )
        self.connect_button.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Email list frame
        list_frame = ttk.Frame(self.root, padding="10")
        list_frame.grid(row=1, column=0, sticky="nsew")
        
        self.tree = ttk.Treeview(list_frame, columns=("from", "subject", "date"), show="headings")
        self.tree.heading("from", text="From")
        self.tree.heading("subject", text="Subject")
        self.tree.heading("date", text="Date")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self.show_email)
        
        # Email content frame
        content_frame = ttk.Frame(self.root, padding="10")
        content_frame.grid(row=2, column=0, sticky="nsew")
        
        self.text_view = tk.Text(content_frame, wrap="word", state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(content_frame, command=self.text_view.yview)
        self.text_view.configure(yscrollcommand=scrollbar.set)
        
        self.text_view.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Configure grid weights
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
    
    def connect_to_email(self):
        self.connect_button.config(state=tk.DISABLED)
        self.root.config(cursor="watch")
        self.root.update()
        
        try:
            fetcher = EmailFetcherService(
                self.imap_server.get(),
                self.username.get(),
                self.password.get()
            )
            fetcher.connect()
            
            # Получаем уже обработанные UID из базы данных
            processed_uids = self.db.get_processed_uids()
            
            # Загружаем только новые сообщения
            self.emails = fetcher.fetch_emails(processed_uids)
            
            for email in self.emails:
                # Сохраняем вложения
                attachments = email["attachments"]
                saved_paths = []
                for attachment in attachments:
                    file_path = os.path.join(ATTACHMENTS_DIR, attachment["filename"])
                    with open(file_path, "wb") as f:
                        f.write(attachment["content"])
                    saved_paths.append(file_path)
                
                # Обновляем путь к вложениям в email
                email["attachments"] = saved_paths
                
                # Сохраняем письмо в базу данных
                self.db.insert_email(email)
            
            self.update_treeview()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to connect: {str(e)}")
        finally:
            self.root.after(0, lambda: self.connect_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.root.config(cursor=""))

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for email_data in self.emails:
            self.tree.insert("", tk.END, values=(
                email_data["from"],  # Отправитель
                email_data["subject"],  # Тема
                email_data["date"]  # Дата
            ))
    
    def show_email(self, event):
        selected_item = self.tree.focus()
        if not selected_item:
            return
            
        item_data = self.tree.item(selected_item)
        selected_email = next(
            (email for email in self.emails 
            if email["from"] == item_data["values"][0]
            and email["subject"] == item_data["values"][1]
            and email["date"] == item_data["values"][2]),
            None
        )
        
        if selected_email:
            self.text_view.config(state=tk.NORMAL)
            self.text_view.delete(1.0, tk.END)
            
            text = f"От: {selected_email['from']}\n"
            text += f"Кому: {selected_email['to']}\n"
            text += f"Дата: {selected_email['date']}\n"
            text += f"Тема: {selected_email['subject']}\n"
            
            if selected_email['attachments']:
                text += f"\nВложения:\n" + "\n".join(selected_email['attachments']) + "\n"
            
            text += "\n" + "-"*50 + "\n\n"
            text += selected_email['body']
            
            self.text_view.insert(tk.END, text)
            self.text_view.config(state=tk.DISABLED)
    
    def save_emails_to_db(self):
        for email_data in self.emails:
            self.db.insert_email(email_data)

    def load_emails_from_db(self):
        self.emails = self.db.get_all_emails()
        for email in self.emails:
            # Убедитесь, что поле "id" существует
            if "id" in email:
                email["attachments"] = self.db.get_attachments_by_email_id(email["id"])
        self.update_treeview()
    
    def __del__(self):
        self.db.close()


    
    def auto_fetch_emails(self):
        try:
            fetcher = EmailFetcherService(
                self.imap_server.get(),
                self.username.get(),
                self.password.get()
            )
            fetcher.connect()
            
            # Получаем уже обработанные UID из базы данных
            processed_uids = self.db.get_processed_uids()
            
            # Загружаем только непрочитанные сообщения
            new_emails = fetcher.fetch_emails(processed_uids)
            
            for email in new_emails:
                # Сохраняем письмо в базу данных
                email_id = self.db.insert_email(email)
                if not email_id:
                    continue
                
                # Сохраняем вложения
                saved_paths = []
                for attachment in email["attachments"]:
                    file_path = os.path.join(ATTACHMENTS_DIR, attachment["filename"])
                    with open(file_path, "wb") as f:
                        f.write(attachment["content"])
                    
                    attachment_data = {
                        "filename": attachment["filename"],
                        "path": file_path,
                        "content": attachment["content"],
                        "text": self.extract_text_from_attachment(file_path)
                    }
                    self.db.insert_attachment(email_id, attachment_data)
                    saved_paths.append(file_path)
                
                # Обновляем поле attachments в таблице emails
                self.db.update_email_attachments(email_id, saved_paths)
            
            # Обновляем интерфейс
            if new_emails:
                self.load_emails_from_db()
        except Exception as e:
            print(f"Ошибка при автоматической проверке писем: {e}")
        finally:
            # Запускаем таймер для следующей проверки
            self.root.after(10000, self.auto_fetch_emails)  # Проверяем каждые 60 секунд

    def extract_text_from_attachment(self, file_path: str) -> str:
        try:
            if file_path.endswith(".pdf"):               
                reader = PdfReader(file_path)
                return " ".join(page.extract_text() for page in reader.pages)
            elif file_path.endswith(".docx"): 
                with open(file_path, "rb") as docx_file:
                    result = mammoth.extract_raw_text(docx_file)
                    return result.value  # Возвращает текст из .docx
            else:
                return ""
        except Exception as e:
            print(f"Ошибка при извлечении текста из вложения {file_path}: {e}")
            return ""
    
    def get_all_emails(self) -> List[Dict]:
        cursor = self.conn.cursor(cursor_factory=RealDictCursor)
        cursor.execute("SELECT id, sender, recipient, subject, date, body FROM emails")
        emails = cursor.fetchall()
        cursor.close()
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
                body TEXT
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
                text TEXT
            )
        """)
        self.conn.commit()
        cursor.close()

    def remove_attachments_column(self):
        cursor = self.conn.cursor()
        cursor.execute("""
            ALTER TABLE emails
            DROP COLUMN IF EXISTS attachments
        """)
        self.conn.commit()
        cursor.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailMonitorApp(root)
    root.mainloop()