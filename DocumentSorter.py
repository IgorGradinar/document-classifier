'''
Запускается на Python 3.10
'''
import pymorphy2
import re
import json

morph = pymorphy2.MorphAnalyzer()

def remove_spaces_between_letters(text):
    """
    Убирает пробелы между всеми заглавными буквами, например, 'П Р И К А З' -> 'ПРИКАЗ'.
    """
    while True:
        new_text = re.sub(r'(?<=\b[А-Яа-я])\s(?=[А-Яа-я])', '', text)
        if new_text == text:
            break
        text = new_text
    return text

def normalize_text_and_lemmatize(text):
    """
    Нормализует текст (удаляет пробелы между буквами, приводит к нижнему регистру)
    и выполняет лемматизацию слов.
    """
    text = remove_spaces_between_letters(text)  # Убираем пробелы между буквами
    text = re.sub(r'\s+', ' ', text)  # Убираем лишние пробелы
    words = re.findall(r'\w+', text.lower())  # Разделяем текст на слова
    lemmatized_words = [morph.parse(word)[0].normal_form for word in words]  # Лемматизируем слова
    return lemmatized_words

def classify_document_with_lemmatization(raw_text, categories_keywords):
    """
    Функция лемматизирует текст документа и ключевые слова категорий,
    чтобы учитывать разные формы слов при поиске.
    """
    lemmatized_text = normalize_text_and_lemmatize(raw_text)
    print(f"Лемматизированный текст:\n{lemmatized_text}")
    category_scores = {}
    detailed_explanation = {}

    for category, keywords in categories_keywords.items():
        score = 0
        explanation_lines = []
        for keyword, weight in keywords.items():
            # Лемматизируем ключевое слово
            lemmatized_keyword = morph.parse(keyword)[0].normal_form
            count = lemmatized_text.count(lemmatized_keyword)
            keyword_score = count * weight
            score += keyword_score
            explanation_lines.append(
                f"Ключевое слово '{keyword}' (лемма: '{lemmatized_keyword}') найдено {count} раз, вес {weight} дает вклад {keyword_score:.2f}."
            )
        # Бонус за упоминание категории
        if morph.parse(category.split()[0])[0].normal_form in lemmatized_text:
            score += 0.5
            explanation_lines.append(
                f"Найдено прямое упоминание категории '{category}', добавляем бонус 0.50."
            )
        category_scores[category] = score
        detailed_explanation[category] = explanation_lines

    total_score = sum(category_scores.values())
    probabilities = {cat: (score / total_score if total_score > 0 else 0) for cat, score in category_scores.items()}
    most_likely_category = max(probabilities, key=probabilities.get)
    return most_likely_category, probabilities, detailed_explanation

# Пример: загрузка JSON-файла с ключевыми словами категорий
with open('categories_keywords.json', 'r', encoding='utf-8') as f:
    categories_keywords = json.load(f)

# Текст документа
raw_text = """
МИНОБРНАУКИ РОССИИ
Федеральное государственное бюджетное образовательное учреждение
высшего образования
«Комсомольский-на-Амуре государственный университет»
(ФГБОУ ВО «КнАГУ»)
  П Р И К А З  
   №   
  г. Комсомольск-на-Амуре  
О создании

С целью развития проектной деятельности в университете, развития научных инициатив и повышения качества подготовки студентов
ПРИКАЗЫВАЮ:
1  Создать на факультете СКБ/СПБ/СНО    
"""

# Запуск классификации
most_likely_category, probabilities, detailed_explanation = classify_document_with_lemmatization(raw_text, categories_keywords)

# Вывод краткого результата
sorted_probs = sorted(probabilities.items(), key=lambda x: x[1], reverse=True)
print(f"Наиболее вероятная категория: {most_likely_category}\n")
print("Вероятности для всех категорий:")
for cat, prob in sorted_probs:
    print(f"- {cat}: {prob * 100:.2f}%")

# Сохранение подробностей в файл
with open("detailed_explanation.txt", "w", encoding="utf-8") as f:
    for cat, details in detailed_explanation.items():
        f.write(f"\nКатегория: {cat}\n")
        for line in details:
            f.write(line + "\n")

print("\nПодробное пояснение сохранено в файл 'detailed_explanation.txt'.")
