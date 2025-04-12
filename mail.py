import imaplib
import email
from email.header import decode_header
from typing import List, Dict

class EmailFetcher:
    def __init__(self, imap_server: str, username: str, password: str):
        self.imap_server = imap_server
        self.username = username
        self.password = password
        self.mail = None
    
    def connect(self):
        self.mail = imaplib.IMAP4_SSL(self.imap_server)
        self.mail.login(self.username, self.password)
        self.mail.select("inbox")
    
    def fetch_emails(self, limit: int = 10) -> List[Dict]:
        if not self.mail:
            raise ConnectionError("Not connected to IMAP server")
        
        _, messages = self.mail.search(None, "ALL")
        email_ids = messages[0].split()[-limit:]
        emails = []
        
        for email_id in email_ids:
            _, msg_data = self.mail.fetch(email_id, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            email_data = {
                "from": self._decode_header(msg["From"]),
                "to": self._decode_header(msg["To"]),
                "subject": self._decode_header(msg["Subject"]),
                "date": self._decode_header(msg["Date"]),
                "body": self._get_email_body(msg),
                "attachments": self._get_attachments(msg)
            }
            emails.append(email_data)
        
        return emails
    
    def _decode_header(self, header):
        if header is None:
            return ""
        decoded = decode_header(header)
        return "".join(
            part[0].decode(part[1] or "utf-8") if isinstance(part[0], bytes) else part[0]
            for part in decoded
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
    
    def _get_attachments(self, msg) -> List[str]:
        attachments = []
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_disposition() == "attachment":
                    filename = part.get_filename()
                    if filename:
                        attachments.append(self._decode_header(filename))
        return attachments
    
    def disconnect(self):
        if self.mail:
            self.mail.close()
            self.mail.logout()