import os
import pytesseract
from docx import Document
from PIL import Image

doc = Document("example.docx")
output_dir = "extracted_images"

os.makedirs(output_dir, exist_ok=True)

text = ""

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# Обрабатываем содержимое документа последовательно
for index, element in enumerate(doc.element.body):
    if element.tag.endswith("p"):  # Абзац текста
        text += element.text + "\n"

for rel in doc.part.rels:
    if "image" in doc.part.rels[rel].target_ref:
        image_part = doc.part.rels[rel].target_part  # Используем target_part
        img_path = os.path.join(output_dir, f"image_{rel}.jpg")

        with open(img_path, "wb") as img_file:
            img_file.write(image_part.blob)  # .blob содержит бинарные данные

        # Распознаем текст на изображении через OCR
        img = Image.open(img_path)
        text += pytesseract.image_to_string(img, lang="rus") + "\n"

print(text)
