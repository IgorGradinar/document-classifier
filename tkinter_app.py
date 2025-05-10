import tkinter as tk
from tkinter import ttk, filedialog
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
import subprocess
import shutil

rarfile.UNRAR_TOOL = r"C:\Users\zbujh\OneDrive\Рабочий стол\UnRAR.exe"

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
ATTACHMENTS_DIR = "attachments"
CATEGORY_DIR = "categorized_attachments"

if not os.path.exists(ATTACHMENTS_DIR):
    os.makedirs(ATTACHMENTS_DIR)

if not os.path.exists(CATEGORY_DIR):
    os.makedirs(CATEGORY_DIR)

def move_file_to_category(file_path, category):
    """Перемещает файл в папку категории."""
    if not category:
        return file_path  # Не перемещаем, если категория не определена
    category_folder = os.path.join(CATEGORY_DIR, str(category))
    if not os.path.exists(category_folder):
        os.makedirs(category_folder)
    new_path = os.path.join(category_folder, os.path.basename(file_path))
    try:
        shutil.move(file_path, new_path)
        return new_path
    except Exception as e:
        print(f"Ошибка при перемещении файла: {e}")
        return file_path

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
        # Search frame
        search_frame = ttk.Frame(self.root, padding="10")
        search_frame.grid(row=0, column=0, sticky="ew")
        
        ttk.Label(search_frame, text="Поиск:").grid(row=0, column=0)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.grid(row=0, column=1, padx=5)
        self.search_entry.bind('<KeyRelease>', self.search_emails)

        # --- Добавляем выпадающий список для категорий ---
        ttk.Label(search_frame, text="Категория:").grid(row=0, column=2, padx=(10, 0))
        self.category_var = tk.StringVar()
        self.category_combobox = ttk.Combobox(search_frame, textvariable=self.category_var, state="readonly")
        self.category_combobox.grid(row=0, column=3, padx=5)
        self.category_combobox.bind('<<ComboboxSelected>>', self.search_emails)
        self.category_combobox.bind('<Button-1>', lambda e: self.update_category_combobox())
        self.update_category_combobox()
        
        # Connection frame
        connection_frame = ttk.Frame(self.root, padding="10")
        connection_frame.grid(row=1, column=0, sticky="ew")
        
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
        
        # Email list frame with buttons
        list_frame = ttk.Frame(self.root, padding="10")
        list_frame.grid(row=2, column=0, sticky="nsew")
        
        # Buttons frame
        buttons_frame = ttk.Frame(list_frame)
        buttons_frame.pack(fill="x", pady=(0, 5))
        
        self.delete_button = ttk.Button(
            buttons_frame,
            text="Удалить",
            command=self.delete_selected_email,
            state=tk.DISABLED
        )
        self.delete_button.pack(side="left", padx=5)
        
        self.tree = ttk.Treeview(list_frame, columns=("from", "subject", "date"), show="headings")
        self.tree.heading("from", text="From")
        self.tree.heading("subject", text="Subject")
        self.tree.heading("date", text="Date")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self.show_email)
        self.tree.bind("<<TreeviewSelect>>", self.on_email_select)
        
        # Email content frame
        content_frame = ttk.Frame(self.root, padding="10")
        content_frame.grid(row=3, column=0, sticky="nsew")
        
        self.text_view = tk.Text(content_frame, wrap="word", state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(content_frame, command=self.text_view.yview)
        self.text_view.configure(yscrollcommand=scrollbar.set)
        
        self.text_view.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Configure grid weights
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_rowconfigure(3, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
    
    def update_category_combobox(self):
        # Собираем уникальные категории из всех вложений
        categories = set()
        for email in self.emails:
            for att in email.get('attachments', []):
                if isinstance(att, dict):
                    cat = att.get('category')
                    if cat is not None:
                        categories.add(str(cat))
        categories = sorted(categories)
        self.category_combobox['values'] = ["Все"] + categories
        self.category_combobox.set("Все")

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
            
            # Подготавливаем данные для пакетной вставки
            email_batch = []
            attachment_batch = []
            
            for email in self.emails:
                # Сохраняем письмо в базу данных и получаем email_id
                email_id = self.db.insert_email(email)
                if not email_id:
                    continue
                
                # Сохраняем вложения
                attachments = email["attachments"]
                saved_paths = []
                
                for attachment in attachments:
                    try:
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
                            attachment_batch.append((email_id, archive_data))

                            # Извлекаем файлы из архива
                            extracted_files = extract_files_from_archive(file_path, ATTACHMENTS_DIR)
                            for extracted_file in extracted_files:
                                try:
                                    # Читаем бинарное содержимое файла
                                    with open(extracted_file, "rb") as f:
                                        file_content = f.read()

                                    # Извлекаем текст из файла только если это текстовый документ
                                    document_text = ""
                                    if extracted_file.lower().endswith(('.txt', '.docx', '.pdf')):
                                        document_text = self.extract_text_from_attachment(extracted_file, lang='rus+eng')

                                    category = NeuroDocumentSorter.sort_document(document_text) if document_text else None
                                    extracted_file = move_file_to_category(extracted_file, category)

                                    # Добавляем данные извлечённых файлов в пакет
                                    attachment_data = {
                                        "filename": os.path.basename(extracted_file),
                                        "path": extracted_file,
                                        "content": file_content,
                                        "text": document_text,
                                        "category": category,
                                    }
                                    attachment_batch.append((email_id, attachment_data))
                                    saved_paths.append(extracted_file)
                                except Exception as e:
                                    print(f"Ошибка при обработке извлечённого файла {extracted_file}: {e}")
                                    continue
                        else:
                            # Обработка обычных вложений
                            document_text = ""
                            if file_path.lower().endswith(('.txt', '.docx', '.pdf')):
                                document_text = self.extract_text_from_attachment(file_path, lang='rus+eng')

                            category = NeuroDocumentSorter.sort_document(document_text) if document_text else None
                            file_path = move_file_to_category(file_path, category)

                            attachment_data = {
                                "filename": os.path.basename(file_path),
                                "path": file_path,
                                "content": attachment["content"],
                                "text": document_text,
                                "category": category
                            }
                            attachment_batch.append((email_id, attachment_data))
                    except Exception as e:
                        print(f"Ошибка при обработке вложения {attachment.get('filename', 'unknown')}: {e}")
                        continue
                
                # Обновляем путь к вложениям в email
                email["attachments"] = saved_paths
                email_batch.append((email_id, saved_paths))
            
            # Пакетное обновление базы данных
            if email_batch:
                self.db.batch_update_email_attachments(email_batch)
            if attachment_batch:
                self.db.batch_insert_attachments(attachment_batch)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to connect: {str(e)}")
        finally:
            self.root.after(0, lambda: self.connect_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.root.config(cursor=""))

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for email_data in self.emails:
            try:
                self.tree.insert("", tk.END, values=(
                    email_data.get("from", ""),
                    email_data.get("subject", ""),
                    email_data.get("date", "")
                ))
            except Exception as e:
                print(f"Ошибка при добавлении письма в список: {e}")
                continue
    
    def show_email(self, event):
        selected_item = self.tree.focus()
        if not selected_item:
            return
        
        item_data = self.tree.item(selected_item)
        values = item_data["values"]
        if not values or len(values) < 3:
            return
        
        selected_email = next(
            (email for email in self.emails 
            if email.get("from", "") == values[0]
            and email.get("subject", "") == values[1]
            and email.get("date", "") == values[2]),
            None
        )
        
        if selected_email:
            self.text_view.config(state=tk.NORMAL)
            self.text_view.delete(1.0, tk.END)
            
            # Добавляем основную информацию о письме
            self.text_view.insert(tk.END, f"От: {selected_email.get('from', '')}\n")
            self.text_view.insert(tk.END, f"Кому: {selected_email.get('to', '')}\n")
            self.text_view.insert(tk.END, f"Дата: {selected_email.get('date', '')}\n")
            self.text_view.insert(tk.END, f"Тема: {selected_email.get('subject', '')}\n")
            
            # Добавляем вложения
            if selected_email.get('attachments'):
                self.text_view.insert(tk.END, "\nВложения:\n")
                
                for idx, attachment in enumerate(selected_email['attachments']):
                    if isinstance(attachment, str):
                        filename = os.path.basename(attachment)
                        path = attachment
                    else:
                        filename = attachment.get('filename', '')
                        path = attachment.get('path', '')
                    
                    if filename and path:
                        self.text_view.insert(tk.END, "- ")
                        file_start = self.text_view.index("end-1c")
                        self.text_view.insert(tk.END, filename)
                        file_end = self.text_view.index("end-1c")
                        
                        # Уникальный тег для каждой ссылки
                        link_tag = f"link_{idx}"
                        self.text_view.tag_add(link_tag, file_start, file_end)
                        self.text_view.tag_configure(link_tag, foreground="blue", underline=1)
                        self.text_view.tag_bind(link_tag, "<Button-1>", lambda e, p=path: self.open_attachment(p))
                        self.text_view.tag_bind(link_tag, "<Enter>", lambda e, t=link_tag: self.text_view.tag_configure(t, foreground="purple"))
                        self.text_view.tag_bind(link_tag, "<Leave>", lambda e, t=link_tag: self.text_view.tag_configure(t, foreground="blue"))
                        
                        # Ссылка на папку
                        self.text_view.insert(tk.END, "\n")
                        folder_start = self.text_view.index("end-1c linestart")
                        folder_end = self.text_view.index("end-1c")
                        folder_tag = f"folder_link_{idx}"
                        self.text_view.tag_add(folder_tag, folder_start, folder_end)
                        self.text_view.tag_configure(folder_tag, foreground="green", underline=1)
                        self.text_view.tag_bind(folder_tag, "<Button-1>", lambda e, p=path: self.open_attachment_folder(p))
                        self.text_view.tag_bind(folder_tag, "<Enter>", lambda e, t=folder_tag: self.text_view.tag_configure(t, foreground="dark green"))
                        self.text_view.tag_bind(folder_tag, "<Leave>", lambda e, t=folder_tag: self.text_view.tag_configure(t, foreground="green"))
                        self.text_view.insert(tk.END, "\n")
            
            # Добавляем разделитель и тело письма
            self.text_view.insert(tk.END, "\n" + "-"*50 + "\n\n")
            self.text_view.insert(tk.END, selected_email.get('body', ''))
            
            self.text_view.config(state=tk.DISABLED)

    def save_emails_to_db(self):
        for email_data in self.emails:
            self.db.insert_email(email_data)

    def load_emails_from_db(self):
        try:
            cursor = self.db.conn.cursor(cursor_factory=RealDictCursor)
            cursor.execute("""
                SELECT e.id, 
                       e.sender as "from",
                       e.recipient as "to",
                       e.subject,
                       e.date,
                       e.body,
                       array_agg(json_build_object(
                           'filename', a.filename,
                           'path', a.path,
                           'text', a.text,
                           'category', a.category
                       )) as attachments
                FROM emails e
                LEFT JOIN attachments a ON e.id = a.email_id
                GROUP BY e.id, e.sender, e.recipient, e.subject, e.date, e.body
                ORDER BY e.date DESC
            """)
            emails = cursor.fetchall()
            cursor.close()

            self.emails = []
            for email in emails:
                # Преобразуем attachments из строки в список, если это строка
                if isinstance(email['attachments'], str):
                    email['attachments'] = []
                elif email['attachments'] and email['attachments'][0] is None:
                    email['attachments'] = []
                
                self.emails.append(dict(email))

            self.update_treeview()
        except Exception as e:
            print(f"Ошибка при загрузке писем из БД: {e}")
            messagebox.showerror("Ошибка", f"Не удалось загрузить письма: {str(e)}")
    
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
                try:
                    # Сохраняем письмо в базу данных
                    email_id = self.db.insert_email(email)
                    if not email_id:
                        continue
                    
                    # Сохраняем вложения
                    saved_paths = []
                    for attachment in email["attachments"]:
                        try:
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
                                result = cursor.fetchone()
                                cursor.close()

                                if result:
                                    archive_id = result[0]
                                else:
                                    print(f"Не удалось найти архив с path: {file_path}")
                                    archive_id = None

                                # Извлекаем файлы из архива
                                extracted_files = extract_files_from_archive(file_path, ATTACHMENTS_DIR)
                                for extracted_file in extracted_files:
                                    try:
                                        # Читаем бинарное содержимое файла
                                        with open(extracted_file, "rb") as f:
                                            file_content = f.read()

                                        # Извлекаем текст из файла
                                        document_text = self.extract_text_from_attachment(extracted_file, lang='rus+eng')
                                        print(f"Извлеченный текст из {extracted_file}: {len(document_text)} символов")

                                        category = NeuroDocumentSorter.sort_document(document_text) if document_text else None
                                        extracted_file = move_file_to_category(extracted_file, category)

                                        # Добавляем данные извлечённых файлов в таблицу attachments
                                        attachment_data = {
                                            "filename": os.path.basename(extracted_file),
                                            "path": extracted_file,
                                            "content": file_content,
                                            "text": document_text,
                                            "category": category,
                                        }
                                        self.db.insert_attachment(email_id, attachment_data)
                                        print(f"Вложение сохранено в БД: {extracted_file}")

                                        # Добавляем путь извлечённого файла в saved_paths
                                        saved_paths.append(extracted_file)
                                    except Exception as e:
                                        print(f"Ошибка при обработке извлечённого файла {extracted_file}: {e}")
                                        continue
                            else:
                                # Обработка обычных вложений
                                document_text = self.extract_text_from_attachment(file_path, lang='rus+eng')
                                print(f"Извлеченный текст из {file_path}: {len(document_text)} символов")

                                category = NeuroDocumentSorter.sort_document(document_text) if document_text else None
                                file_path = move_file_to_category(file_path, category)

                                attachment_data = {
                                    "filename": os.path.basename(file_path),
                                    "path": file_path,
                                    "content": attachment["content"],
                                    "text": document_text,
                                    "category": category
                                }
                                self.db.insert_attachment(email_id, attachment_data)
                        except Exception as e:
                            print(f"Ошибка при обработке вложения {attachment.get('filename', 'unknown')}: {e}")
                            continue
                    
                    # Обновляем поле attachments в таблице emails
                    self.db.update_email_attachments(email_id, saved_paths)
                except Exception as e:
                    print(f"Ошибка при обработке письма: {e}")
                    continue
            
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
            # Кэшируем результаты извлечения текста
            cache_key = f"{file_path}_{lang}"
            if hasattr(self, '_text_cache') and cache_key in self._text_cache:
                return self._text_cache[cache_key]
            
            text = ""
            
            # Обработка текстовых файлов
            if file_path.lower().endswith('.txt'):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        text = f.read()
                except UnicodeDecodeError:
                    try:
                        with open(file_path, 'r', encoding='cp1251') as f:
                            text = f.read()
                    except:
                        print(f"Не удалось прочитать текстовый файл: {file_path}")
                        return ""
            
            elif file_path.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
                try:
                    processed_path = preprocess_image(file_path)
                    image = Image.open(processed_path)
                    text = pytesseract.image_to_string(image, lang=lang)
                except Exception as e:
                    print(f"Ошибка при обработке изображения: {e}")
                    return ""
            
            elif file_path.endswith(".pdf"):
                try:
                    images = convert_from_path(file_path, poppler_path=r"C:\poppler-24.08.0\Library\bin")
                    text_parts = []
                    for img in images:
                        try:
                            page_text = pytesseract.image_to_string(img, lang=lang)
                            text_parts.append(page_text)
                        except Exception as e:
                            print(f"Ошибка при обработке страницы PDF: {e}")
                    text = "\n".join(text_parts)
                except Exception as e:
                    print(f"Ошибка при обработке PDF: {e}")
                    return ""

            elif file_path.endswith(".docx"):
                try:
                    with open(file_path, "rb") as docx_file:
                        result = mammoth.extract_raw_text(docx_file)
                        text = result.value.strip()
                except Exception as e:
                    print(f"Ошибка при обработке DOCX: {e}")
                    return ""

            # Очистка текста от лишних пробелов и переносов строк
            if text:
                text = ' '.join(text.split())
            
            # Сохраняем результат в кэш
            if not hasattr(self, '_text_cache'):
                self._text_cache = {}
            self._text_cache[cache_key] = text.strip()
            
            return text.strip()
        except Exception as e:
            print(f"Ошибка при извлечении текста из файла {file_path}: {e}")
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
                category INTEGER
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
            INSERT INTO attachments (email_id, filename, path, content, text, category)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (
            email_id,
            attachment_data["filename"],
            attachment_data["path"],
            attachment_data["content"],
            attachment_data["text"],
            attachment_data["category"],
        ))
        self.conn.commit()
        cursor.close()

    def search_emails(self, event=None):
        search_term = self.search_var.get().lower()
        selected_category = self.category_var.get()

        self.tree.delete(*self.tree.get_children())
        
        for email_data in self.emails:
            # Проверяем категории вложений
            if selected_category and selected_category != "Все":
                has_matching_category = False
                if "attachments" in email_data:
                    for attachment in email_data["attachments"]:
                        if isinstance(attachment, dict):
                            category = str(attachment.get("category")) if attachment.get("category") is not None else None
                            if category == selected_category:
                                has_matching_category = True
                                break
                if not has_matching_category:
                    continue

            # Если нет поискового запроса, просто добавляем письмо
            if not search_term:
                self.tree.insert("", tk.END, values=(
                    email_data.get("from", ""),
                    email_data.get("subject", ""),
                    email_data.get("date", "")
                ))
                continue

            # Поиск в основных полях письма
            found = False
            highlight_info = {
                "body_highlights": [],
                "attachment_highlights": []
            }

            if (search_term in email_data.get("from", "").lower() or
                search_term in email_data.get("subject", "").lower() or
                search_term in email_data.get("body", "").lower()):
                found = True
                body = email_data.get("body", "").lower()
                start = 0
                while True:
                    start = body.find(search_term, start)
                    if start == -1:
                        break
                    highlight_info["body_highlights"].append(start)
                    start += len(search_term)

            # Поиск в вложениях
            if "attachments" in email_data:
                for attachment in email_data["attachments"]:
                    if isinstance(attachment, str):
                        filename = os.path.basename(attachment)
                        path = attachment
                        text = ""
                    else:
                        filename = attachment.get("filename", "")
                        path = attachment.get("path", "")
                        text = attachment.get("text", "") or ""

                    # Поиск по имени файла
                    if search_term in filename.lower():
                        found = True
                        highlight_info["attachment_highlights"].append({
                            "filename": filename,
                            "path": path,
                            "text_highlights": []
                        })
                        continue

                    # Поиск по тексту вложения
                    if text:
                        text_lower = text.lower()
                        start = 0
                        text_highlights = []
                        while True:
                            start = text_lower.find(search_term, start)
                            if start == -1:
                                break
                            text_highlights.append(start)
                            start += len(search_term)
                        if text_highlights:
                            found = True
                            highlight_info["attachment_highlights"].append({
                                "filename": filename,
                                "path": path,
                                "text_highlights": text_highlights
                            })

            # Добавляем письмо в список если найдены совпадения
            if found:
                item = self.tree.insert("", tk.END, values=(
                    email_data.get("from", ""),
                    email_data.get("subject", ""),
                    email_data.get("date", "")
                ))
                self.tree.item(item, tags=(str(highlight_info),))

        self.tree.bind('<<TreeviewSelect>>', self.on_search_select)

    def on_search_select(self, event):
        selected_item = self.tree.focus()
        if not selected_item:
            return
        
        # Получаем информацию о подсветке
        tags = self.tree.item(selected_item)['tags']
        if not tags:
            return
        
        highlight_info = eval(tags[0])
        
        # Показываем письмо с подсветкой
        self.show_email_with_highlights(selected_item, highlight_info)

    def show_email_with_highlights(self, selected_item, highlight_info):
        item_data = self.tree.item(selected_item)
        values = item_data["values"]
        if not values or len(values) < 3:
            return
        
        selected_email = next(
            (email for email in self.emails 
            if email.get("from", "") == values[0]
            and email.get("subject", "") == values[1]
            and email.get("date", "") == values[2]),
            None
        )
        
        if selected_email:
            self.text_view.config(state=tk.NORMAL)
            self.text_view.delete(1.0, tk.END)
            
            # Настраиваем тег для подсветки
            self.text_view.tag_configure("highlight", background="yellow")
            
            # Добавляем основную информацию о письме
            self.text_view.insert(tk.END, f"От: {selected_email.get('from', '')}\n")
            self.text_view.insert(tk.END, f"Кому: {selected_email.get('to', '')}\n")
            self.text_view.insert(tk.END, f"Дата: {selected_email.get('date', '')}\n")
            self.text_view.insert(tk.END, f"Тема: {selected_email.get('subject', '')}\n")
            
            # Добавляем вложения
            if selected_email.get('attachments'):
                self.text_view.insert(tk.END, "\nВложения:\n")
                
                for idx, attachment in enumerate(selected_email['attachments']):
                    if isinstance(attachment, str):
                        filename = os.path.basename(attachment)
                        path = attachment
                    else:
                        filename = attachment.get('filename', '')
                        path = attachment.get('path', '')
                    
                    if filename and path:
                        should_highlight = any(
                            h.get('filename') == filename 
                            for h in highlight_info.get('attachment_highlights', [])
                        )
                        self.text_view.insert(tk.END, "- ")
                        file_start = self.text_view.index("end-1c")
                        self.text_view.insert(tk.END, filename)
                        file_end = self.text_view.index("end-1c")
                        
                        # Уникальный тег для каждой ссылки
                        link_tag = f"link_{idx}"
                        if should_highlight:
                            self.text_view.tag_add("highlight", file_start, file_end)
                        self.text_view.tag_add(link_tag, file_start, file_end)
                        self.text_view.tag_configure(link_tag, foreground="blue", underline=1)
                        self.text_view.tag_bind(link_tag, "<Button-1>", lambda e, p=path: self.open_attachment(p))
                        self.text_view.tag_bind(link_tag, "<Enter>", lambda e, t=link_tag: self.text_view.tag_configure(t, foreground="purple"))
                        self.text_view.tag_bind(link_tag, "<Leave>", lambda e, t=link_tag: self.text_view.tag_configure(t, foreground="blue"))
                        
                        # Ссылка на папку
                        self.text_view.insert(tk.END, "\n")
                        folder_start = self.text_view.index("end-1c linestart")
                        folder_end = self.text_view.index("end-1c")
                        folder_tag = f"folder_link_{idx}"
                        self.text_view.tag_add(folder_tag, folder_start, folder_end)
                        self.text_view.tag_configure(folder_tag, foreground="green", underline=1)
                        self.text_view.tag_bind(folder_tag, "<Button-1>", lambda e, p=path: self.open_attachment_folder(p))
                        self.text_view.tag_bind(folder_tag, "<Enter>", lambda e, t=folder_tag: self.text_view.tag_configure(t, foreground="dark green"))
                        self.text_view.tag_bind(folder_tag, "<Leave>", lambda e, t=folder_tag: self.text_view.tag_configure(t, foreground="green"))
                        self.text_view.insert(tk.END, "\n")
            
            # Добавляем разделитель и тело письма
            self.text_view.insert(tk.END, "\n" + "-"*50 + "\n\n")
            
            # Добавляем тело письма с подсветкой
            body = selected_email.get('body', '')
            if highlight_info.get('body_highlights'):
                # Подсвечиваем найденные слова в теле письма
                last_pos = 0
                for pos in sorted(highlight_info['body_highlights']):
                    # Добавляем текст до найденного слова
                    self.text_view.insert(tk.END, body[last_pos:pos])
                    # Добавляем найденное слово с подсветкой
                    highlight_start = self.text_view.index("end-1c")
                    self.text_view.insert(tk.END, body[pos:pos + len(self.search_var.get())])
                    highlight_end = self.text_view.index("end-1c")
                    self.text_view.tag_add("highlight", highlight_start, highlight_end)
                    last_pos = pos + len(self.search_var.get())
                # Добавляем оставшийся текст
                self.text_view.insert(tk.END, body[last_pos:])
            else:
                self.text_view.insert(tk.END, body)
            
            self.text_view.config(state=tk.DISABLED)

    def on_email_select(self, event):
        selected_items = self.tree.selection()
        self.delete_button.config(state=tk.NORMAL if selected_items else tk.DISABLED)

    def delete_selected_email(self):
        selected_item = self.tree.focus()
        if not selected_item:
            return
        
        item_data = self.tree.item(selected_item)
        values = item_data["values"]
        if not values or len(values) < 3:
            return
        
        selected_email = next(
            (email for email in self.emails 
            if email.get("from", "") == values[0]
            and email.get("subject", "") == values[1]
            and email.get("date", "") == values[2]),
            None
        )
        
        if selected_email and "id" in selected_email:
            try:
                # Удаляем вложения из файловой системы
                if "attachments" in selected_email:
                    for attachment in selected_email["attachments"]:
                        try:
                            if isinstance(attachment, str):
                                file_path = attachment
                            else:
                                file_path = attachment.get("path")
                            
                            if file_path and os.path.exists(file_path):
                                os.remove(file_path)
                                print(f"Удален файл: {file_path}")
                        except Exception as e:
                            print(f"Ошибка при удалении файла {file_path}: {e}")
                
                # Удаляем письмо из базы данных
                cursor = self.db.conn.cursor()
                try:
                    # Сначала удаляем все вложения, связанные с письмом
                    cursor.execute("DELETE FROM attachments WHERE email_id = %s", (selected_email["id"],))
                    # Затем удаляем само письмо
                    cursor.execute("DELETE FROM emails WHERE id = %s", (selected_email["id"],))
                    self.db.conn.commit()
                    print(f"Удалено письмо из БД с ID: {selected_email['id']}")
                except Exception as e:
                    self.db.conn.rollback()
                    raise e
                finally:
                    cursor.close()
                
                # Удаляем из интерфейса
                self.emails.remove(selected_email)
                self.tree.delete(selected_item)
                self.text_view.config(state=tk.NORMAL)
                self.text_view.delete(1.0, tk.END)
                self.text_view.config(state=tk.DISABLED)
                self.delete_button.config(state=tk.DISABLED)
                messagebox.showinfo("Успех", "Письмо успешно удалено")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось удалить письмо: {str(e)}")
                print(f"Ошибка при удалении письма: {e}")

    def open_attachment(self, file_path):
        if os.path.exists(file_path):
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(file_path)
                else:  # Linux/Mac
                    subprocess.run(['xdg-open', file_path])
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось открыть файл: {str(e)}")
        else:
            messagebox.showerror("Ошибка", "Файл не найден")

    def open_attachment_folder(self, file_path):
        if os.path.exists(file_path):
            try:
                folder_path = os.path.dirname(file_path)
                if os.name == 'nt':  # Windows
                    subprocess.run(['explorer', '/select,', file_path])
                else:  # Linux/Mac
                    subprocess.run(['xdg-open', folder_path])
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось открыть папку: {str(e)}")
        else:
            messagebox.showerror("Ошибка", "Файл не найден")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailMonitorApp(root)
    root.mainloop()
