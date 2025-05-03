import tkinter as tk
from tkinter import ttk
from mail import EmailFetcherService
from posgre import EmailDatabaseManager
from tkinter import messagebox
import os
from typing import List, Dict  # Добавьте этот импорт
from psycopg2.extras import RealDictCursor  # Добавьте этот импорт
import mammoth
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import uuid
import io  # Добавлено для работы с изображениями в .docx
from langdetect import detect
import cv2
from docx import Document
import NeuroDocumentSorter
import rarfile
import zipfile

rarfile.UNRAR_TOOL = r"C:\Users\zbujh\OneDrive\Рабочий стол\UnRAR.exe"

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
ATTACHMENTS_DIR = "attachments"
if not os.path.exists(ATTACHMENTS_DIR):
    os.makedirs(ATTACHMENTS_DIR)

def detect_language(text):
    try:
        return detect(text)
    except:
        return "unknown"

def preprocess_image(image_path):
    # Загружаем изображение
    image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    
    # Увеличиваем контрастность
    image = cv2.equalizeHist(image)
    
    # Применяем бинаризацию (чёрно-белое изображение)
    _, image = cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    # Сохраняем обработанное изображение во временный файл
    processed_path = "processed_image.png"
    cv2.imwrite(processed_path, image)
    return processed_path

def extract_files_from_archive(archive_path: str, extract_to: str) -> List[str]:
    extracted_files = []
    try:
        if archive_path.endswith(".rar"):
            with rarfile.RarFile(archive_path, 'r') as rar_ref:
                rar_ref.extractall(extract_to)
                extracted_files = rar_ref.namelist()
        elif archive_path.endswith(".zip"):
            with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                zip_ref.extractall(extract_to)
                extracted_files = zip_ref.namelist()
    except Exception as e:
        print(f"Ошибка при распаковке архива {archive_path}: {e}")
    return [os.path.join(extract_to, file) for file in extracted_files]

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
                # Сохраняем письмо в базу данных и получаем email_id
                email_id = self.db.insert_email(email)
                if not email_id:
                    continue
                
                # Сохраняем вложения
                attachments = email["attachments"]
                saved_paths = []
                for attachment in attachments:
                    original_filename = attachment["filename"]
                    file_path = os.path.join(ATTACHMENTS_DIR, original_filename)

                    # Проверяем, существует ли файл, и создаём уникальное имя
                    if os.path.exists(file_path):
                        unique_id = str(uuid.uuid4())[:8]
                        file_name, file_extension = os.path.splitext(original_filename)
                        file_path = os.path.join(ATTACHMENTS_DIR, f"{file_name}_{unique_id}{file_extension}")

                    # Сохраняем файл с уникальным именем
                    with open(file_path, "wb") as f:
                        f.write(attachment["content"])

                    # Добавляем путь к файлу в saved_paths
                    saved_paths.append(file_path)

                    # Если файл является архивом, обрабатываем его
                    if file_path.endswith((".rar", ".zip")):
                        archive_data = {
                            "filename": os.path.basename(file_path),
                            "path": file_path,
                            "content": attachment["content"],
                            "text": None,
                            "category": None
                        }
                        self.db.insert_attachment(email_id, archive_data)

                        # Получаем ID архива
                        cursor = self.db.conn.cursor()
                        cursor.execute("SELECT id FROM attachments WHERE path = %s", (file_path,))
                        archive_id = cursor.fetchone()[0]
                        print(archive_id)
                        cursor.close()

                        # Извлекаем файлы из архива
                        extracted_files = extract_files_from_archive(file_path, ATTACHMENTS_DIR)
                        for extracted_file in extracted_files:
                            # Читаем бинарное содержимое файла
                            with open(extracted_file, "rb") as f:
                                file_content = f.read()

                            # Извлекаем текст из файла
                            document_text = self.extract_text_from_attachment(extracted_file, lang='rus+eng')

                            # Добавляем данные извлечённых файлов в таблицу attachments
                            attachment_data = {
                                "filename": os.path.basename(extracted_file),
                                "path": extracted_file,
                                "content": file_content,
                                "text": document_text,
                                "category": NeuroDocumentSorter.sort_document(document_text),
                                "owner": archive_id
                            }
                            self.db.insert_attachment(email_id, attachment_data)

                            # Добавляем путь извлечённого файла в saved_paths
                            saved_paths.append(extracted_file)
                    else:
                        # Обработка обычных вложений
                        document_text = self.extract_text_from_attachment(file_path, lang='rus+eng')

                        attachment_data = {
                            "filename": os.path.basename(file_path),
                            "path": file_path,
                            "content": attachment["content"],
                            "text": document_text,
                            "category": NeuroDocumentSorter.sort_document(document_text)
                        }
                        self.db.insert_attachment(email_id, attachment_data)
                
                # Обновляем путь к вложениям в email
                email["attachments"] = saved_paths
                # Обновляем поле attachments в таблице emails
                self.db.update_email_attachments(email_id, saved_paths)
                
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
                    original_filename = attachment["filename"]
                    file_path = os.path.join(ATTACHMENTS_DIR, original_filename)

                    # Проверяем, существует ли файл, и создаём уникальное имя
                    if os.path.exists(file_path):
                        unique_id = str(uuid.uuid4())[:8]
                        file_name, file_extension = os.path.splitext(original_filename)
                        file_path = os.path.join(ATTACHMENTS_DIR, f"{file_name}_{unique_id}{file_extension}")

                    # Сохраняем файл с уникальным именем
                    with open(file_path, "wb") as f:
                        f.write(attachment["content"])

                    # Добавляем путь к файлу в saved_paths
                    saved_paths.append(file_path)

                    # Если файл является архивом, обрабатываем его
                    if file_path.endswith((".rar", ".zip")):
                        archive_data = {
                            "filename": os.path.basename(file_path),
                            "path": file_path,
                            "content": attachment["content"],
                            "text": None,
                            "category": None
                        }
                        self.db.insert_attachment(email_id, archive_data)

                        # Получаем ID архива
                        cursor = self.db.conn.cursor()
                        cursor.execute("SELECT id FROM attachments WHERE path = %s", (file_path,))
                        archive_id = cursor.fetchone()[0]
                        cursor.close()

                        # Извлекаем файлы из архива
                        extracted_files = extract_files_from_archive(file_path, ATTACHMENTS_DIR)
                        for extracted_file in extracted_files:
                            # Читаем бинарное содержимое файла
                            with open(extracted_file, "rb") as f:
                                file_content = f.read()

                            # Извлекаем текст из файла
                            document_text = self.extract_text_from_attachment(extracted_file, lang='rus+eng')

                            # Добавляем данные извлечённых файлов в таблицу attachments
                            attachment_data = {
                                "filename": os.path.basename(extracted_file),
                                "path": extracted_file,
                                "content": file_content,
                                "text": document_text,
                                "category": NeuroDocumentSorter.sort_document(document_text),
                                "owner": archive_id
                            }
                            self.db.insert_attachment(email_id, attachment_data)

                            # Добавляем путь извлечённого файла в saved_paths
                            saved_paths.append(extracted_file)
                    else:
                        # Обработка обычных вложений
                        document_text = self.extract_text_from_attachment(file_path, lang='rus+eng')

                        attachment_data = {
                            "filename": os.path.basename(file_path),
                            "path": file_path,
                            "content": attachment["content"],
                            "text": document_text,
                            "category": NeuroDocumentSorter.sort_document(document_text)
                        }
                        self.db.insert_attachment(email_id, attachment_data)
                
                # Обновляем поле attachments в таблице emails
                self.db.update_email_attachments(email_id, saved_paths)
            
            # Обновляем интерфейс
            if new_emails:
                self.load_emails_from_db()
        except Exception as e:
            print(f"Ошибка при автоматической проверке писем: {e}")
        finally:
            # Запускаем таймер для следующей проверки
            self.root.after(10000, self.auto_fetch_emails)

    def extract_text_from_attachment(self, file_path: str, lang: str = 'rus') -> str:
        try:
            text = ""
            if file_path.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
                # Предварительная обработка изображения
                processed_path = preprocess_image(file_path)
                image = Image.open(processed_path)
                text = pytesseract.image_to_string(image, lang=lang)
            
            elif file_path.endswith(".pdf"):
                # Извлечение текста из PDF
                images = convert_from_path(file_path)
                for image in images:
                    text += " " + pytesseract.image_to_string(image, lang=lang)
            
            elif file_path.endswith(".docx"):
                # Извлечение текста из .docx
                with open(file_path, "rb") as docx_file:
                    result = mammoth.extract_raw_text(docx_file)
                    text = result.value.strip()
                
                # Если текст не извлечён, пробуем извлечь текст из изображений
                if not text.strip():

                    doc = Document(file_path)
                    for rel in doc.part.rels.values():
                        if "image" in rel.target_ref:
                            image_path = rel.target_part.blob
                            image = Image.open(io.BytesIO(image_path))
                            text += " " + pytesseract.image_to_string(image, lang=lang)
            
            # Постобработка текста
            return text.strip()
        except Exception as e:
            print(f"Ошибка при извлечении текста из вложения {file_path}: {e}")
            return ""
    
    def get_all_emails(self) -> List[Dict]:
        cursor = self.conn.cursor(cursor_factory=RealDictCursor)
        cursor.execute("SELECT id, sender, recipient, subject, date, body FROM emails ORDER BY date")
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
        # Преобразуем список в строку, разделённую запятыми
        attachments_str = ", ".join(attachments) if attachments else None
        cursor.execute("""
            UPDATE emails
            SET attachments = %s
            WHERE id = %s
        """, (attachments_str, email_id))
        self.conn.commit()
        cursor.close()

    def create_table(self):
        cursor = self.conn.cursor()
        # Создаём таблицу attachments
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS attachments (
                id SERIAL PRIMARY KEY,
                email_id INTEGER REFERENCES emails(id) ON DELETE CASCADE,
                filename TEXT,
                path TEXT,
                content BYTEA,
                text TEXT,
                category INTEGER,
                owner INTEGER  
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

    def insert_attachment(self, email_id: int, attachment_data: Dict):
        cursor = self.conn.cursor()
        cursor.execute("""
            INSERT INTO attachments (email_id, filename, path, content, text, category, owner)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, (
            email_id,
            attachment_data["filename"],
            attachment_data["path"],
            attachment_data["content"],
            attachment_data["text"],
            attachment_data["category"],
            attachment_data.get("owner")  # Указываем владельца (ID архива)
        ))
        self.conn.commit()
        cursor.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailMonitorApp(root)
    root.mainloop()
