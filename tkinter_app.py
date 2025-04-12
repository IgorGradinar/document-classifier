import tkinter as tk
from tkinter import ttk
from mail import EmailFetcher
from sql import EmailDatabase

class EmailMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Monitor")
        self.emails = []
        self.db = EmailDatabase()
        self.setup_ui()
        self.load_emails_from_db()
    
    def setup_ui(self):
        # Connection frame
        connection_frame = ttk.Frame(self.root, padding="10")
        connection_frame.grid(row=0, column=0, sticky="ew")
        
        ttk.Label(connection_frame, text="IMAP Server:").grid(row=0, column=0)
        self.imap_server = ttk.Entry(connection_frame)
        self.imap_server.grid(row=0, column=1, padx=5)
        
        ttk.Label(connection_frame, text="Username:").grid(row=1, column=0)
        self.username = ttk.Entry(connection_frame)
        self.username.grid(row=1, column=1, padx=5)
        
        ttk.Label(connection_frame, text="Password:").grid(row=2, column=0)
        self.password = ttk.Entry(connection_frame, show="*")
        self.password.grid(row=2, column=1, padx=5)
        
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
            fetcher = EmailFetcher(
                self.imap_server.get(),
                self.username.get(),
                self.password.get()
            )
            fetcher.connect()
            self.emails = fetcher.fetch_emails()
            self.update_treeview()
            self.save_emails_to_db()
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to connect: {str(e)}")
        finally:
            self.root.after(0, lambda: self.connect_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.root.config(cursor=""))
    
    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for email_data in self.emails:
            self.tree.insert("", tk.END, values=(
                email_data["from"],
                email_data["subject"],
                email_data["date"]
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
                text += f"\nВложения: {', '.join(selected_email['attachments'])}\n"
            
            text += "\n" + "-"*50 + "\n\n"
            text += selected_email['body']
            
            self.text_view.insert(tk.END, text)
            self.text_view.config(state=tk.DISABLED)
    
    def save_emails_to_db(self):
        for email_data in self.emails:
            self.db.insert_email(email_data)
    
    def load_emails_from_db(self):
        self.emails = self.db.get_all_emails()
        self.update_treeview()
    
    def __del__(self):
        self.db.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailMonitorApp(root)
    root.mainloop()
