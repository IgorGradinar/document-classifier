import imaplib
import email
from email.header import decode_header
from typing import List, Dict
import re
class EmailFetcherService:  # Renamed from EmailFetcher
    def __init__(self, imap_server: str, username: str, password: str):
        self.imap_server = imap_server
        self.username = username
        self.password = password
        self.mail = None
    
    def connect(self):
        self.mail = imaplib.IMAP4_SSL(self.imap_server)
        self.mail.login(self.username, self.password)
        self.mail.select("inbox")
    
    def fetch_emails(self, processed_uids: List[str], limit: int = 10) -> List[Dict]:
        if not self.mail:
            raise ConnectionError("Not connected to IMAP server")
        
        # Используем фильтр SINCE для загрузки сообщений с определённой даты
        _, messages = self.mail.search(None, "UNSEEN")
        email_ids = messages[0].split()
        emails = []
        
        for email_id in email_ids:
            # Получаем UID сообщения
            _, uid_data = self.mail.fetch(email_id, "(UID)")
            match = re.search(r"UID (\d+)", uid_data[0].decode())
            if match:
                uid = match.group(1)
            else:
                continue  # Пропускаем, если UID не найден
            
            # Пропускаем уже обработанные сообщения
            if uid in processed_uids:
                continue
            
            _, msg_data = self.mail.fetch(email_id, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            email_data = {
                "uid": uid,
                "from": self._decode_header(msg["From"]),
                "to": self._decode_header(msg["To"]),
                "subject": self._decode_header(msg["Subject"]),
                "date": self._decode_header(msg["Date"]),
                "body": self._get_email_body(msg),
                "attachments": self._get_attachments(msg)
            }
            emails.append(email_data)
            
            # Ограничиваем количество загружаемых сообщений
            if len(emails) >= limit:
                break
        
        return emails
    
    def _decode_header(self, header):
        decoded_parts = decode_header(header)
        return "".join(
            part.decode(encoding or "utf-8") if isinstance(part, bytes) else part
            for part, encoding in decoded_parts
        )
    
    def _get_email_body(self, msg) -> str:
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    body += part.get_payload(decode=True).decode("utf-8", errors="ignore")
        else:
         body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")
        return body
    
    def _get_attachments(self, msg) -> List[Dict]:
        attachments = []
        if msg.is_multipart():
            for part in msg.walk():
                content_disposition = part.get("Content-Disposition", "")
                if "attachment" in content_disposition or part.get_filename():
                    filename = part.get_filename()
                    if filename:
                        decoded_filename = self._decode_header(filename)
                        try:
                            content = part.get_payload(decode=True)
                            attachments.append({
                                "filename": decoded_filename,
                                "content": content
                            })
                        except Exception as e:
                            print(f"Ошибка при декодировании вложения {decoded_filename}: {e}")
        return attachments
    
    def disconnect(self):
        if self.mail:
            self.mail.close()
            self.mail.logout()