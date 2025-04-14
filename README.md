# Document Classifier

This project is a document classification system that fetches emails, processes attachments, and classifies documents into predefined categories.

## Features
- Fetch emails from an IMAP server.
- Extract and process email attachments (PDF, DOCX, images, etc.).
- Classify documents using keywords or neural network models.
- Store emails and attachments in a PostgreSQL database.
- Provide a GUI for monitoring emails and viewing their content.

## Installation

### Prerequisites
- Python 3.10 or higher
- PostgreSQL database
- Tesseract OCR (for text extraction from images)
- Ollama (for neural network model integration)

### Dependencies
Install the required Python packages using `pip`:

```bash
pip install -r requirements.txt
```

### Required Python Packages
- `tkinter` (built-in with Python)
- `psycopg2` (PostgreSQL database connection)
- `PyPDF2` (PDF text extraction)
- `pytesseract` (OCR for images)
- `Pillow` (Image processing)
- `pdf2image` (Convert PDF pages to images)
- `mammoth` (Extract text from DOCX files)
- `pymorphy2` (Russian text lemmatization)
- `langdetect` (Language detection)
- `opencv-python` (Image preprocessing)
- `numpy` (Numerical operations)
- `langchain-ollama` (Neural network model integration)

### Additional Setup
1. **Tesseract OCR**: Install Tesseract OCR and configure its path in the code:
   ```python
   pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
   ```

2. **PostgreSQL Database**: Create a PostgreSQL database and configure the connection in `posgre.py`:
   ```python
   db_name="emails"
   user="postgres"
   password="password"
   host="localhost"
   port=5432
   ```

3. **IMAP Server**: Configure the IMAP server, username, and password in the GUI (`tkinter_app.py`).

4. **Ollama Setup**: Install Ollama and pull the required model:
   ```bash
   ollama pull gemma2:27b
   ```

## Usage

1. Run the GUI application:
   ```bash
   python tkinter_app.py
   ```

2. Connect to the email server using the GUI.

3. View and classify emails and their attachments.

## File Structure
- `tkinter_app.py`: Main GUI application.
- `mail.py`: Email fetching service.
- `posgre.py`: PostgreSQL database manager.
- `DocumentSorter.py`: Document classification using keywords.
- `NeuroDocumentSorter.py`: Neural network-based document classification.
- `categories_keywords.json`: JSON file containing categories and keywords.

## License
This project is licensed under the MIT License.