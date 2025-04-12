import re
import json

def normalize_for_search(text):
    """
    Убирает все пробелы из текста, чтобы при поиске не учитывать их.
    """
    return re.sub(r'\s+', '', text)

def classify_document_manual_no_spaces(raw_text, categories_keywords):
    """
    Функция вычисляет суммарный балл для каждой категории,
    используя поиск ключевых слов с игнорированием пробелов.
    Для этого исходный текст и ключевые слова предварительно нормализуются
    (удаляются все пробелы и переводятся в нижний регистр).
    Если название категории явно встречается в тексте, добавляется бонус.
    """
    category_scores = {}
    detailed_explanation = {}

    # Нормализуем весь текст, удаляя все пробелы, и переводим в нижний регистр
    normalized_text = normalize_for_search(raw_text).lower()

    for category, keywords in categories_keywords.items():
        score = 0
        explanation_lines = []
        for keyword, weight in keywords.items():
            # Нормализуем ключевое слово (удаляются все пробелы и перевод в нижний регистр)
            normalized_keyword = normalize_for_search(keyword).lower()
            # Используем метод count, так как последовательное совпадение без пробелов
            count = normalized_text.count(normalized_keyword)
            keyword_score = count * weight
            score += keyword_score
            explanation_lines.append(
                f"Ключевое слово '{keyword}' (без пробелов: '{normalized_keyword}') найдено {count} раз, вес {weight} дает вклад {keyword_score:.2f}."
            )
        # Если название категории содержится в нормализованном тексте, добавляем бонус
        if normalize_for_search(category).lower() in normalized_text:
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

# Пример: загрузка словаря категорий с ключевыми словами из JSON-файла
with open('categories_keywords.json', 'r', encoding='utf-8') as f:
    categories_keywords = json.load(f)

# Пример заглушки для документа (замените этот текст на ваш документ)
raw_text = """
МИНОБРНАУКИ РОССИИ
Федеральное государственное бюджетное образовательное учреждение
высшего образования
«Комсомольский-на-Амуре государственный университет»
(ФГБОУ ВО «КнАГУ»)
  П Р И К А З  
   №   
  г. Комсомольск-на-Амуре  
  О создании

(название СКБ/СПБ/СНО) 
С целью развития проектной деятельности в университете, развития научных инициатив и повышения качества подготовки студентов
ПРИКАЗЫВАЮ:
1  Создать на факультете СКБ/СПБ/СНО    
(название факультета)
    (  )
(название СКБ/СПБ/СНО)
2  Декану факультета организовать подготовку к утверждению:
-  положение о;
(название СКБ/СПБ/СНО)
-  план работ на 20/уч. год.
(название СКБ/СПБ/СНО)
Срок исполнения не позднее.

Ректор университета  Э.А. Дмитриев
Проект приказа вносит декан      
И.О. Фамилия
СОГЛАСОВАНО  
Проректор по НР  А.В. Космынин
Начальник УНИД
Начальник ПУ  А.В. Ахметова
А.В. Ременников
"""

most_likely_category, probabilities, detailed_explanation = classify_document_manual_no_spaces(raw_text, categories_keywords)

# Вывод краткого резюме в консоль
sorted_probs = sorted([(cat, prob) for cat, prob in probabilities.items() if prob > 0], key=lambda x: x[1], reverse=True)
print(f"Наиболее вероятная категория: {most_likely_category}\n")
print("Вероятности для всех категорий (отсортировано):")
for cat, prob in sorted_probs:
    print(f"- {cat}: {prob * 100:.2f}%")

# Сохранение подробного пояснения в файл для детального анализа
with open("detailed_explanation.txt", "w", encoding="utf-8") as f:
    for cat, details in detailed_explanation.items():
        f.write(f"\nКатегория: {cat}\n")
        for line in details:
            f.write(line + "\n")

print("\nПодробное пояснение сохранено в файле 'detailed_explanation.txt'.")
