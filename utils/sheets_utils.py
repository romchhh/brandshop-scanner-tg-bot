import re
import csv
from config import data


def get_category_by_prefix(art):
    prefixes = {
        "Дж": ["jeans"],
        "Фут": ["tshirts_polo"],
        "Пл": ["shorts_swim"],
        "П": ["tshirts_polo"],
        "Кос": ["costumes_fleece", "costumes"], 
        "Ко": ["sweaters"],
        "Дш": ["shorts_jeans"],
        "Шор": ["shorts_textile"],
        "Шо": ["shorts_swim"],
        "Ку": ["jackets"],
        "Ж": ["waistcoats"],
        "Бр": ["trousers"],
        "Шт": ["sport_trousers"],
        "Ру": ["shirts"],
        "Н": ["socks"],
        "Тр": ["underwear"],
        "Об": ["shoes", "wintershoes"],
        "Ке": ["caps"],
        "Шап": ["hats"],
        "Ре": ["belts"],
        "Сум": ["bags"],
        "Клч": ["purses"],
        "Кл": ["costumes_summer"],
        "Т": ["tapki"],
        "Оч": ["glasses"],
    }

    # Приводимо артикул до верхнього регістру для порівняння
    art_upper = art.upper() if art else ""
    
    for prefix, categories in prefixes.items():
        # Приводимо префікс до верхнього регістру для порівняння
        prefix_upper = prefix.upper()
        if art_upper.startswith(prefix_upper):
            return categories 
    return None


def has_sizes(category):
    """
    Перевіряє, чи категорія має розміри
    Категорії без розмірів: hats, belts, bags, purses, glasses (окуляри)
    """
    categories_without_sizes = {'hats', 'belts', 'bags', 'purses', 'glasses'}
    return category not in categories_without_sizes


def parse_amount_from_cell(cell_value):
    """
    Витягує кількість з комірки Google таблиці.
    "2" -> 2
    "2, (,1,-склад)" -> 2+1=3
    "2, (,2,-склад)" -> 2+2=4
    Сумує всі числа в рядку.
    """
    if not cell_value or not str(cell_value).strip():
        return 0
    numbers = re.findall(r'\d+', str(cell_value).strip())
    return sum(int(n) for n in numbers) if numbers else 0



def get_size_mapping():
    """Повертає маппінг числових розмірів на буквені"""
    return {
        '46': 'S',
        '48': 'M',
        '50': 'L',
        '52': 'XL',
        '54': '2XL',
        '56': '3XL',
        '58': '4XL',
        '60': '5XL',
    }


def map_numeric_to_letter_size(size):
    """Конвертує числовий розмір в буквений за маппінгом"""
    size_mapping = get_size_mapping()
    return size_mapping.get(size, size)  # Повертаємо оригінальний розмір, якщо маппінгу немає


def normalize_size(size):
    """Нормалізує розмір, конвертуючи кирилицю в латиницю та числові розміри в буквені"""
    import logging
    import re
    
    if not size:
        return ""
    
    original_size = size
    size = size.strip()
    
    # Обробка розмірів типу "5(2ХL)" - витягуємо розмір з дужок
    # Шукаємо патерн: число(розмір), наприклад "5(2ХL)" -> "2ХL"
    bracket_match = re.search(r'\(([^)]+)\)', size)
    if bracket_match:
        size_in_brackets = bracket_match.group(1).strip()
        # Нормалізуємо розмір з дужок
        normalized_bracket = normalize_size(size_in_brackets)
        if normalized_bracket:
            logging.debug(f"[normalize_size] '{original_size}' -> витягнуто з дужок '{size_in_brackets}' -> '{normalized_bracket}'")
            return normalized_bracket
    
    # Спочатку перевіряємо маппінг числових розмірів (46, 48, 50, 52, 54, 56, 58, 60)
    size_mapping = get_size_mapping()
    if size in size_mapping:
        return size_mapping[size]
    
    # Конвертуємо до верхнього регістру для обробки кирилиці
    size_upper = size.upper()
    size_lower = size.lower()
    
    # Конвертація кирилиці в латиницю для розмірів (підтримка великих та малих букв)
    cyrillic_to_latin = {
        'С': 'S',
        'М': 'M',
        'Л': 'L',
        'ХС': 'XS',
        'ХЛ': 'XL',
        '2ХЛ': '2XL',
        '3ХЛ': '3XL',
        '4ХЛ': '4XL',
        '5ХЛ': '5XL',
        '6ХЛ': '6XL',
        '7ХЛ': '7XL',
        '8ХЛ': '8XL',
    }
    
    # Обробка малих букв кирилиці: с, м, л -> S, M, L
    if size_lower in ['м', 'л', 'с']:
        if size_lower == 'м':
            return 'M'
        elif size_lower == 'л':
            return 'L'
        elif size_lower == 'с':
            return 'S'
    
    # Обробка "хл", "2хл", "3хл", "4хл" тощо (малі букви)
    if 'хл' in size_lower:
        # Замінюємо "хл" на "XL" з урахуванням числа перед ним
        # Перевіряємо точне співпадіння або початок рядка
        if size_lower == '4хл' or size_lower.startswith('4хл'):
            return '4XL'
        elif size_lower == '3хл' or size_lower.startswith('3хл'):
            return '3XL'
        elif size_lower == '2хл' or size_lower.startswith('2хл'):
            return '2XL'
        elif size_lower == '5хл' or size_lower.startswith('5хл'):
            return '5XL'
        elif size_lower == '6хл' or size_lower.startswith('6хл'):
            return '6XL'
        elif size_lower == '7хл' or size_lower.startswith('7хл'):
            return '7XL'
        elif size_lower == '8хл' or size_lower.startswith('8хл'):
            return '8XL'
        elif size_lower == 'хл' or size_lower.startswith('хл'):
            return 'XL'
    
    # Обробка "хс" (кирилиця) -> XS
    if size_lower == 'хс':
        return 'XS'
    
    # Перевіряємо, чи це повний розмір у кирилиці (великі букви)
    if size_upper in cyrillic_to_latin:
        return cyrillic_to_latin[size_upper]
    
    # Виправлення випадків сканера (латиниця замість кирилиці): xc -> XS, c -> S
    if size_lower == 'xc':
        return 'XS'
    if size_lower == 'c' and len(size) == 1:
        return 'S'
    
    # Якщо розмір вже в латиниці (XS, S, M, L, XL, 2XL, 3XL, 4XL, 5XL, 6XL, 7XL, 8XL), повертаємо як є
    valid_latin_sizes = {'XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', '5XL', '6XL', '7XL', '8XL'}
    if size_upper in valid_latin_sizes:
        return size_upper
    
    # Замінюємо кириличні символи в рядку (великі букви)
    result = size_upper
    for cyr, lat in cyrillic_to_latin.items():
        result = result.replace(cyr, lat)
    
    logging.debug(f"[normalize_size] '{original_size}' -> '{result}'")
    return result

def split_sizes(size_string):
    sizes = re.split(r'[,\s\-]+', size_string)
    return [normalize_size(size) for size in sizes if size.strip()]


def split_allsizes(size_string):
    sizes = re.split(r'\s*,\s*', size_string)
    return [normalize_size(size) for size in sizes if size.strip()]


def adjust_size(original_size, filter_value):
    # Handle the '/' symbol by converting it to a positive adjustment
    if filter_value.startswith('/'):
        filter_value = '+' + filter_value[1:]

    # Check if the filter_value is a plain number
    if filter_value.isdigit():
        filter_value = '+' + filter_value

    if filter_value.startswith('-') or filter_value.startswith('+'):
        try:
            adjustment = int(filter_value)
        except ValueError:
            print(f"Invalid filter_value: {filter_value}")
            return original_size

        print(f"Adjustment: {adjustment}")

        if original_size in {"28", "29", "30", "31", "32", "33", "34", "35", "36", "38", "39", "40", "41", "42", "43", "44", "45", "46"}:
            size_order = ["28", "29", "30", "31", "32", "33", "34", "35", "36", "38", "39", "40", "41", "42", "43", "44", "45", "46"]
            current_index = size_order.index(original_size)
            new_index = max(0, min(current_index + adjustment, len(size_order) - 1))
            print(f"Original size: {original_size}, Current index: {current_index}, New index: {new_index}, New size: {size_order[new_index]}")
            return size_order[new_index]
        elif original_size in {"XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL", "8XL"}:
            size_order = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL", "8XL"]
            current_index = size_order.index(original_size)
            new_index = max(0, min(current_index + adjustment, len(size_order) - 1))
            print(f"Original size: {original_size}, Current index: {current_index}, New index: {new_index}, New size: {size_order[new_index]}")
            return size_order[new_index]
        
    return original_size




def find_item_by_size_in_category(client, size, categories):
    result = []
    normalized_size = normalize_size(size)

    allsize_values = {
        "28", "29", "30", "31", "32", "33", "34", "35", "36", "38", "40", "42", "44",
        "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL", "8XL",
        "39", "40", "41", "42", "43", "44", "45", "46",
        # Числові розміри (також додаємо їх буквені еквіваленти)
        "48", "50", "52", "54", "56", "58", "60"
    }

    for category in categories:
        if category not in data:
            continue

        details = data[category]
        for link in details["link"]:
            sheet_numbers = details["sheet"]

            if isinstance(sheet_numbers, list):
                for sheet_number in sheet_numbers:
                    sheet = client.open_by_url(link).get_worksheet(sheet_number)
                    all_data = sheet.get_all_values()

                    art_column = [row[details["art"] - 1] for row in all_data]
                    price_column = [row[details["price"] - 1] for row in all_data]
                    size_column = [
                        row[details["size"] - 1] if row[details["size"] - 1].strip() else "-"
                        for row in all_data
                    ]
                    photo_column = [row[details["photo"] - 1] for row in all_data]
                    amount_column = [row[details["amount"] - 1] for row in all_data]

                    for i, row_size in enumerate(size_column):
                        row_sizes = split_sizes(row_size)
                        valid_sizes = {size for size in row_sizes if size in allsize_values}

                        if size == "allsize":
                            if len(valid_sizes) >= 3:
                                result.append({
                                    'category': category,
                                    'price': price_column[i],
                                    'art': art_column[i],
                                    'size': size_column[i],
                                    'photo': photo_column[i]
                                })
                        elif size == "nosize":
                            if amount_column[i].strip():
                                result.append({
                                    'category': category,
                                    'price': price_column[i],
                                    'art': art_column[i],
                                    'size': size_column[i],
                                    'photo': photo_column[i]
                                })
                        elif normalized_size in valid_sizes:
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': size_column[i],
                                'photo': photo_column[i]
                            })
            else:
                sheet = client.open_by_url(link).get_worksheet(sheet_numbers)
                all_data = sheet.get_all_values()
                art_column = [row[details["art"] - 1] for row in all_data]
                price_column = [row[details["price"] - 1] for row in all_data]
                size_column = [
                        row[details["size"] - 1] if row[details["size"] - 1].strip() else "-"
                        for row in all_data
                    ]
                photo_column = [row[details["photo"] - 1] for row in all_data]
                amount_column = [row[details["amount"] - 1] for row in all_data]

                for i, row_size in enumerate(size_column):
                    row_sizes = split_sizes(row_size)
                    valid_sizes = {size for size in row_sizes if size in allsize_values}

                    if size == "allsize":
                        if len(valid_sizes) >= 3:
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': size_column[i],
                                'photo': photo_column[i]
                            })
                    elif size == "nosize":
                        if amount_column[i].strip():
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': size_column[i],
                                'photo': photo_column[i]
                            })
                    elif normalized_size in valid_sizes:
                        result.append({
                            'category': category,
                            'price': price_column[i],
                            'art': art_column[i],
                            'size': size_column[i],
                            'photo': photo_column[i]
                        })
    return result



def find_item_by_size_in_category_drop(client, size, categories):
    result = []
    normalized_size = normalize_size(size)

    allsize_values = {
        "28", "29", "30", "31", "32", "33", "34", "35", "36", "38", "40", "42", "44",
        "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL", "8XL",
        "39", "40", "41", "42", "43", "44", "45", "46",
        # Числові розміри (також додаємо їх буквені еквіваленти)
        "48", "50", "52", "54", "56", "58", "60"
    }

    for category in categories:
        if category not in data:
            continue

        details = data[category]
        for link in details["link"]:
            sheet_numbers = details["sheet"]

            if isinstance(sheet_numbers, list):
                for sheet_number in sheet_numbers:
                    sheet = client.open_by_url(link).get_worksheet(sheet_number)
                    all_data = sheet.get_all_values()

                    art_column = [row[details["art"] - 1] for row in all_data]
                    price_column = [
                        row[details["dropprice"] - 1] if row[details["dropprice"] - 1].strip() else "-"
                        for row in all_data
                    ]
                    size_column = [
                        row[details["size"] - 1] if row[details["size"] - 1].strip() else "-"
                        for row in all_data
                    ]
                    photo_column = [row[details["photo"] - 1] for row in all_data]
                    amount_column = [row[details["amount"] - 1] for row in all_data]

                    for i, row_size in enumerate(size_column):
                        row_sizes = split_sizes(row_size)
                        valid_sizes = {size for size in row_sizes if size in allsize_values}

                        if size == "allsize":
                            if len(valid_sizes) >= 3:
                                result.append({
                                    'category': category,
                                    'price': price_column[i],
                                    'art': art_column[i],
                                    'size': size_column[i],
                                    'photo': photo_column[i]
                                })
                        elif size == "nosize":
                            if amount_column[i].strip():
                                result.append({
                                    'category': category,
                                    'price': price_column[i],
                                    'art': art_column[i],
                                    'size': size_column[i],
                                    'photo': photo_column[i]
                                })
                        elif normalized_size in valid_sizes:
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': size_column[i],
                                'photo': photo_column[i]
                            })
            else:
                sheet = client.open_by_url(link).get_worksheet(sheet_numbers)
                all_data = sheet.get_all_values()
                art_column = [row[details["art"] - 1] for row in all_data]
                price_column = [
                        row[details["dropprice"] - 1] if row[details["dropprice"] - 1].strip() else "0"
                        for row in all_data
                    ]
                size_column = [
                        row[details["size"] - 1] if row[details["size"] - 1].strip() else "-"
                        for row in all_data
                    ]
                photo_column = [row[details["photo"] - 1] for row in all_data]
                amount_column = [row[details["amount"] - 1] for row in all_data]

                for i, row_size in enumerate(size_column):
                    row_sizes = split_sizes(row_size)
                    valid_sizes = {size for size in row_sizes if size in allsize_values}

                    if size == "allsize":
                        if len(valid_sizes) >= 3:
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': size_column[i],
                                'photo': photo_column[i]
                            })
                    elif size == "nosize":
                        if amount_column[i].strip():
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': size_column[i],
                                'photo': photo_column[i]
                            })
                    elif normalized_size in valid_sizes:
                        result.append({
                            'category': category,
                            'price': price_column[i],
                            'art': art_column[i],
                            'size': size_column[i],
                            'photo': photo_column[i]
                        })
    return result


def normalize_art(art):
    """Нормалізує артикул, видаляючи пробіли, дефіси та інші символи для порівняння"""
    if not art:
        return ""
    # Видаляємо всі пробіли, дефіси та приводимо до верхнього регістру
    normalized = re.sub(r'[\s\-]+', '', art.upper())
    return normalized


# Значення заголовків — пропускаємо при читанні з таблиць
HEADER_ART_VALUES = {'артикул', 'арт', 'код', 'назва', 'розмір', 'кількість', 'цена', 'ціна', 'фото'}


def _is_header_row(art):
    """Перевіряє, чи артикул є заголовком таблиці (пропускаємо при читанні)"""
    if not art or not str(art).strip():
        return True
    return str(art).strip().lower() in HEADER_ART_VALUES


def extract_base_art(art):
    """Витягує базовий артикул, видаляючи розмір з кінця"""
    if not art:
        return art
    
    # Видаляємо пробіли та дефіси для нормалізації
    art = art.strip()
    
    # Видаляємо тільки .число в кінці (.130, .38). НЕ -200 — це частина артикулу Ре-200
    base_art = re.sub(r'\.\d+$', '', art)
    
    return base_art.strip()


def parse_csv_file(file_path):
    """
    Парсить CSV файл з переобліком
    Повертає словник: {артикул: {розміри: set, кількість: int}}
    """
    inventory_data = {}
    
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            # Визначаємо роздільник (може бути ; або ,)
            first_line = f.readline()
            f.seek(0)
            
            delimiter = ';' if ';' in first_line else ','
            
            reader = csv.reader(f, delimiter=delimiter)
            next(reader)  # Пропускаємо заголовок
            
            for row in reader:
                if len(row) < 2:
                    continue
                
                # Артикул з другого стовпця (індекс 1)
                art = row[1].strip() if len(row) > 1 else ""
                # Розмір з четвертого стовпця (індекс 3) - "Інформація"
                size_info = row[3].strip() if len(row) > 3 else ""
                # Кількість: F (індекс 5) "Відскановано", fallback на E (індекс 4) для 5-колонкового CSV
                if len(row) > 5 and row[5].strip():
                    amount = row[5].strip()
                elif len(row) > 4 and row[4].strip():
                    amount = row[4].strip()
                else:
                    amount = "1"
                
                if not art or _is_header_row(art):
                    continue
                
                # Визначаємо категорію для перевірки, чи має товар розміри
                categories = get_category_by_prefix(art)
                category_has_sizes = True
                if categories:
                    category_has_sizes = has_sizes(categories[0])
                
                # Завжди витягуємо базовий артикул для пошуку (Ре-210.130 -> Ре-210)
                # Частина після крапки (.130) — це варіант/розмір, не використовуємо при пошуку
                base_art = extract_base_art(art)
                
                # Нормалізуємо артикул для порівняння
                normalized_art = normalize_art(base_art)
                
                if normalized_art not in inventory_data:
                    inventory_data[normalized_art] = {
                        'sizes': {},  # {нормалізований_розмір: кількість} - для порівняння
                        'original_sizes': {},  # {нормалізований_розмір: оригінальний_розмір} - для виведення
                        'amount': 0,
                        'original_art': base_art  # Зберігаємо базовий або повний артикул
                    }
                
                # Додаємо розмір з колонки "Інформація" з кількістю
                if size_info and size_info.strip() and category_has_sizes:
                    # Товар з розмірами
                    original_size = size_info.strip()  # Зберігаємо оригінальний розмір з CSV
                    normalized_size = normalize_size(size_info)
                    # Отримуємо кількість для цього розміру
                    try:
                        size_amount = int(amount) if amount else 1
                    except ValueError:
                        size_amount = 1
                    
                    # Додаємо або оновлюємо кількість для розміру
                    if normalized_size in inventory_data[normalized_art]['sizes']:
                        inventory_data[normalized_art]['sizes'][normalized_size] += size_amount
                    else:
                        inventory_data[normalized_art]['sizes'][normalized_size] = size_amount
                    
                    # Зберігаємо оригінальний розмір (якщо для цього нормалізованого розміру ще немає)
                    if normalized_size not in inventory_data[normalized_art]['original_sizes']:
                        inventory_data[normalized_art]['original_sizes'][normalized_size] = original_size
                else:
                    # Для товарів без розмірів - зберігаємо тільки загальну кількість в amount
                    # sizes залишаємо порожнім для таких товарів
                    pass
                
                # Додаємо загальну кількість
                try:
                    inventory_data[normalized_art]['amount'] += int(amount) if amount else 1
                except ValueError:
                    inventory_data[normalized_art]['amount'] += 1
                    
    except Exception as e:
        print(f"Помилка при парсингу CSV: {e}")
        return {}
    
    return inventory_data


def get_art_sizes_from_sheets(client, art, categories):
    """
    Отримує розміри артикулу з Google таблиць
    Повертає set розмірів для даного артикулу
    """
    all_sizes = set()
    normalized_art = normalize_art(art)
    
    allsize_values = {
        "28", "29", "30", "31", "32", "33", "34", "35", "36", "38", "40", "42", "44",
        "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL", "8XL",
        "39", "40", "41", "42", "43", "44", "45", "46",
        # Числові розміри (також додаємо їх буквені еквіваленти)
        "48", "50", "52", "54", "56", "58", "60"
    }
    
    for category in categories:
        if category not in data:
            continue
        
        details = data[category]
        for link in details["link"]:
            sheet_numbers = details["sheet"]
            
            if isinstance(sheet_numbers, list):
                for sheet_number in sheet_numbers:
                    try:
                        sheet = client.open_by_url(link).get_worksheet(sheet_number)
                        all_data = sheet.get_all_values()
                        
                        art_column = [row[details["art"] - 1] for row in all_data]
                        size_column = [
                            row[details["size"] - 1] if len(row) > details["size"] - 1 and row[details["size"] - 1].strip() else "-"
                            for row in all_data
                        ]
                        
                        for i, row_art in enumerate(art_column):
                            if normalize_art(row_art) == normalized_art:
                                row_size = size_column[i] if i < len(size_column) else "-"
                                
                                row_sizes = split_sizes(row_size)
                                valid_sizes = {size for size in row_sizes if size in allsize_values}
                                
                                all_sizes.update(valid_sizes)
                    except Exception as e:
                        print(f"Помилка при читанні таблиці {link}, sheet {sheet_number}: {e}")
                        continue
            else:
                try:
                    sheet = client.open_by_url(link).get_worksheet(sheet_numbers)
                    all_data = sheet.get_all_values()
                    
                    art_column = [row[details["art"] - 1] for row in all_data]
                    size_column = [
                        row[details["size"] - 1] if len(row) > details["size"] - 1 and row[details["size"] - 1].strip() else "-"
                        for row in all_data
                    ]
                    
                    for i, row_art in enumerate(art_column):
                        if normalize_art(row_art) == normalized_art:
                            row_size = size_column[i] if i < len(size_column) else "-"
                            
                            row_sizes = split_sizes(row_size)
                            valid_sizes = {size for size in row_sizes if size in allsize_values}
                            
                            all_sizes.update(valid_sizes)
                except Exception as e:
                    print(f"Помилка при читанні таблиці {link}, sheet {sheet_numbers}: {e}")
                    continue
    
    return all_sizes


def load_all_arts_from_category(client, category):
    """
    Зчитує всі артикули та їх розміри з таблиць категорії одним запитом
    Повертає словник: {нормалізований_артикул: {'sizes': {розмір: кількість}, 'amount': кількість, 'original_art': артикул}}
    Для товарів без розмірів (шапки, сумки, ремні, кошельки) зберігаємо тільки amount
    """
    all_arts_data = {}
    
    if category not in data:
        return all_arts_data
    
    # Перевіряємо, чи категорія має розміри
    category_has_sizes = has_sizes(category)
    
    allsize_values = {
        "28", "29", "30", "31", "32", "33", "34", "35", "36", "38", "40", "42", "44",
        "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL", "8XL",
        "39", "40", "41", "42", "43", "44", "45", "46",
        # Числові розміри (також додаємо їх буквені еквіваленти)
        "48", "50", "52", "54", "56", "58", "60"
    }
    
    details = data[category]
    for link in details["link"]:
        sheet_numbers = details["sheet"]
        
        if isinstance(sheet_numbers, list):
            for sheet_number in sheet_numbers:
                try:
                    sheet = client.open_by_url(link).get_worksheet(sheet_number)
                    all_data = sheet.get_all_values()
                    
                    art_column = [row[details["art"] - 1] if len(row) > details["art"] - 1 else "" for row in all_data]
                    size_column = [
                        row[details["size"] - 1] if len(row) > details["size"] - 1 and row[details["size"] - 1].strip() else "-"
                        for row in all_data
                    ]
                    amount_column = [
                        row[details["amount"] - 1] if len(row) > details["amount"] - 1 else ""
                        for row in all_data
                    ]
                    
                    for i, row_art in enumerate(art_column):
                        if not row_art or not row_art.strip() or _is_header_row(row_art):
                            continue
                        
                        base_art = extract_base_art(row_art)
                        normalized_art = normalize_art(base_art)
                        row_size = size_column[i] if i < len(size_column) else "-"
                        row_amount = amount_column[i] if i < len(amount_column) else ""
                        
                        # Логування для діагностики
                        import logging
                        logging.debug(f"[load_all_arts] Читаємо артикул: {row_art}, розміри: '{row_size}', кількість: '{row_amount}'")
                        
                        # Ініціалізуємо dict для артикулу, якщо його ще немає
                        # Структура: {'sizes': {розмір: кількість}, 'amount': кількість, 'original_art': оригінальний_артикул}
                        if normalized_art not in all_arts_data:
                            all_arts_data[normalized_art] = {
                                'sizes': {},
                                'amount': 0,
                                'original_art': base_art  # Базовий артикул для відображення
                            }
                        
                        # Для товарів без розмірів читаємо amount (колонка amount, не size)
                        # Обробляємо формат "2, (,1,-склад)" — сумуємо всі числа
                        if not category_has_sizes:
                            all_arts_data[normalized_art]['amount'] += parse_amount_from_cell(row_amount)
                        # Обробляємо розміри тільки якщо вони є та категорія має розміри
                        elif row_size and row_size != "-" and row_size.strip():
                            # Використовуємо split_allsizes для правильного розбиття по комах (зберігає дефіси)
                            row_sizes_list = split_allsizes(row_size)
                            
                            # Обробляємо розміри з кількістю
                            # Формат: "M,-2, L," де -2 - це кількість для попереднього розміру M
                            current_size = None
                            for size_item in row_sizes_list:
                                size_item = size_item.strip()
                                # Пропускаємо порожні
                                if not size_item:
                                    continue
                                
                                # Перевіряємо, чи це кількість (починається з - і містить тільки цифри, наприклад "-2")
                                if size_item.startswith('-') and size_item[1:].isdigit():
                                    # Це кількість для попереднього розміру
                                    if current_size is not None:
                                        quantity = int(size_item[1:])  # Видаляємо мінус
                                        sizes_dict = all_arts_data[normalized_art]['sizes']
                                        if current_size in sizes_dict:
                                            sizes_dict[current_size] += quantity
                                        else:
                                            sizes_dict[current_size] = quantity
                                    current_size = None
                                    continue
                                
                                # Пропускаємо фільтри (починаються з / і містить тільки цифри)
                                if size_item.startswith('/') and size_item[1:].isdigit():
                                    continue
                                
                                # Це розмір
                                # Якщо був попередній розмір без кількості, додаємо його з кількістю 1
                                if current_size is not None:
                                    sizes_dict = all_arts_data[normalized_art]['sizes']
                                    if current_size in sizes_dict:
                                        sizes_dict[current_size] += 1
                                    else:
                                        sizes_dict[current_size] = 1
                                
                                adjusted_size = size_item
                                
                                # Нормалізуємо розмір (конвертуємо кирилицю в латиницю та числові в буквені)
                                normalized_adjusted_size = normalize_size(adjusted_size)
                                
                                # Перевіряємо, чи це валідний розмір
                                # Також перевіряємо оригінальний розмір (для числових розмірів)
                                if normalized_adjusted_size in allsize_values or adjusted_size in allsize_values:
                                    # Використовуємо нормалізований розмір (буквений еквівалент для числових)
                                    current_size = normalized_adjusted_size
                                else:
                                    current_size = None
                            
                            # Додаємо останній розмір, якщо він не мав кількості
                            if current_size is not None:
                                if current_size in all_arts_data[normalized_art]['sizes']:
                                    all_arts_data[normalized_art]['sizes'][current_size] += 1
                                else:
                                    all_arts_data[normalized_art]['sizes'][current_size] = 1
                        # Якщо розмірів немає (порожній або "-"), артикул все одно додається з порожнім dict()
                        # Це означає, що товар існує, але не має розмірів (наприклад, шапки)
                except Exception as e:
                    print(f"Помилка при читанні таблиці {link}, sheet {sheet_number}: {e}")
                    continue
        else:
            try:
                sheet = client.open_by_url(link).get_worksheet(sheet_numbers)
                all_data = sheet.get_all_values()
                
                art_column = [row[details["art"] - 1] if len(row) > details["art"] - 1 else "" for row in all_data]
                size_column = [
                    row[details["size"] - 1] if len(row) > details["size"] - 1 and row[details["size"] - 1].strip() else "-"
                    for row in all_data
                ]
                amount_column = [
                    row[details["amount"] - 1] if len(row) > details["amount"] - 1 else ""
                    for row in all_data
                ]
                
                for i, row_art in enumerate(art_column):
                    if not row_art or not row_art.strip() or _is_header_row(row_art):
                        continue
                    
                    base_art = extract_base_art(row_art)
                    normalized_art = normalize_art(base_art)
                    row_size = size_column[i] if i < len(size_column) else "-"
                    row_amount = amount_column[i] if i < len(amount_column) else ""
                    
                    # Логування для діагностики
                    import logging
                    logging.debug(f"[load_all_arts] Читаємо артикул: {row_art}, base: {base_art}, розміри: '{row_size}', кількість: '{row_amount}'")
                    
                    # Ініціалізуємо dict для артикулу, якщо його ще немає
                    if normalized_art not in all_arts_data:
                        all_arts_data[normalized_art] = {
                            'sizes': {},
                            'amount': 0,
                            'original_art': base_art
                        }
                    
                    # Для товарів без розмірів читаємо amount (колонка amount, не size)
                    if not category_has_sizes:
                        all_arts_data[normalized_art]['amount'] += parse_amount_from_cell(row_amount)
                    elif row_size and row_size != "-" and row_size.strip():
                        # Використовуємо split_allsizes для правильного розбиття по комах (зберігає дефіси)
                        row_sizes_list = split_allsizes(row_size)
                        
                        # Обробляємо розміри з кількістю
                        # Формат: "M,-2, L," де -2 - це кількість для попереднього розміру M
                        current_size = None
                        for size_item in row_sizes_list:
                            size_item = size_item.strip()
                            # Пропускаємо порожні
                            if not size_item:
                                continue
                            
                            # Перевіряємо, чи це кількість (починається з - і містить тільки цифри, наприклад "-2")
                            if size_item.startswith('-') and size_item[1:].isdigit():
                                # Це кількість для попереднього розміру
                                if current_size is not None:
                                    quantity = int(size_item[1:])  # Видаляємо мінус
                                    sizes_dict = all_arts_data[normalized_art]['sizes']
                                    if current_size in sizes_dict:
                                        sizes_dict[current_size] += quantity
                                    else:
                                        sizes_dict[current_size] = quantity
                                current_size = None
                                continue
                            
                            # Пропускаємо фільтри (починаються з / і містять тільки цифри)
                            if size_item.startswith('/') and size_item[1:].isdigit():
                                continue
                            
                            # Це розмір
                            # Якщо був попередній розмір без кількості, додаємо його з кількістю 1
                            if current_size is not None:
                                sizes_dict = all_arts_data[normalized_art]['sizes']
                                if current_size in sizes_dict:
                                    sizes_dict[current_size] += 1
                                else:
                                    sizes_dict[current_size] = 1
                            
                            adjusted_size = size_item
                            
                            # Нормалізуємо розмір (конвертуємо кирилицю в латиницю та числові в буквені)
                            normalized_adjusted_size = normalize_size(adjusted_size)
                            
                            # Перевіряємо, чи це валідний розмір
                            # Також перевіряємо оригінальний розмір (для числових розмірів)
                            if normalized_adjusted_size in allsize_values or adjusted_size in allsize_values:
                                # Використовуємо нормалізований розмір (буквений еквівалент для числових)
                                current_size = normalized_adjusted_size
                            else:
                                current_size = None
                        
                        # Додаємо останній розмір, якщо він не мав кількості
                        if current_size is not None:
                            if current_size in all_arts_data[normalized_art]['sizes']:
                                all_arts_data[normalized_art]['sizes'][current_size] += 1
                            else:
                                all_arts_data[normalized_art]['sizes'][current_size] = 1
                    # Якщо розмірів немає (порожній або "-"), артикул все одно додається з порожнім dict()
                    # Це означає, що товар існує, але не має розмірів (наприклад, шапки)
            except Exception as e:
                print(f"Помилка при читанні таблиці {link}, sheet {sheet_numbers}: {e}")
                continue
    
    return all_arts_data


def compare_inventory_with_sheets(client, inventory_data):
    """
    Порівнює дані з файлу з даними в Google таблицях
    Зчитує всю таблицю одним запитом для оптимізації
    Повертає словник з результатами порівняння
    """
    results = {
        'missing_sizes': {},  # {артикул: [розміри яких немає]}
        'extra_sizes': {},    # {артикул: [розміри яких більше]}
        'not_found': [],      # Артикули які не знайдені в таблицях
        'not_scanned': {},    # {артикул: {розміри: кількість}} - є в таблицях, але немає в файлі
        'matched': []         # Артикули які повністю співпадають
    }
    
    # Групуємо артикули по категоріях
    # Для артикулів з кількома категоріями (як "Об" -> ["shoes", "wintershoes"])
    # додаємо їх до всіх відповідних категорій
    arts_by_category = {}
    art_to_categories = {}  # Зберігаємо відповідність артикул -> список категорій
    for normalized_art, data_info in inventory_data.items():
        original_art = data_info['original_art']
        categories = get_category_by_prefix(original_art)
        
        if not categories:
            results['not_found'].append(original_art)
            continue
        
        # Для артикулів з кількома категоріями додаємо їх до всіх категорій
        art_to_categories[normalized_art] = categories
        for category in categories:
            if category not in arts_by_category:
                arts_by_category[category] = []
            arts_by_category[category].append((normalized_art, data_info))
    
    # Зчитуємо дані кожної категорії ОДИН РАЗ (без зайвих API-запитів)
    category_sheet_data = {}
    category_has_sizes_map = {}
    for category in arts_by_category:
        category_sheet_data[category] = load_all_arts_from_category(client, category)
        category_has_sizes_map[category] = has_sizes(category)
    
    # Відстежуємо оброблені артикули, щоб уникнути дублювання для артикулів з кількома категоріями
    processed_arts = set()
    
    # Для кожної категорії порівнюємо артикули з вже завантаженими даними
    for category, arts_list in arts_by_category.items():
        all_sheet_arts = category_sheet_data[category]
        category_has_sizes = category_has_sizes_map[category]
        
        # Створюємо set нормалізованих артикулів з файлу для швидкого пошуку
        file_arts_set = {normalized_art for normalized_art, _ in arts_list}
        
        # Порівнюємо кожен артикул з файлу
        for normalized_art, data_info in arts_list:
            original_art = data_info['original_art']
            
            # Пропускаємо артикули, які вже оброблені (для артикулів з кількома категоріями)
            if normalized_art in processed_arts:
                continue
            
            file_sizes = data_info['sizes']  # {розмір: кількість}
            file_amount = data_info.get('amount', 0)  # Загальна кількість з файлу
            
            # Шукаємо артикул в завантажених даних поточної категорії
            sheet_art_data = all_sheet_arts.get(normalized_art, None)
            
            # Якщо артикул не знайдено в цій категорії, шукаємо в інших (вже завантажених)
            if sheet_art_data is None:
                art_categories = art_to_categories.get(normalized_art, [category])
                for other_category in art_categories:
                    if other_category == category:
                        continue
                    other_sheet_arts = category_sheet_data.get(other_category, {})
                    if normalized_art in other_sheet_arts:
                        sheet_art_data = other_sheet_arts[normalized_art]
                        category_has_sizes = category_has_sizes_map[other_category]
                        break
                
                if sheet_art_data is None:
                    results['not_found'].append(original_art)
                    processed_arts.add(normalized_art)
                    continue
            
            # Позначаємо артикул як оброблений
            processed_arts.add(normalized_art)
            
            # Для товарів без розмірів порівнюємо amount
            if not category_has_sizes:
                sheet_amount = sheet_art_data.get('amount', 0) if isinstance(sheet_art_data, dict) else 0
                
                import logging
                logging.info(f"[DEBUG] Артикул без розмірів: {original_art}")
                logging.info(f"[DEBUG] Кількість з файлу: {file_amount}, з таблиці: {sheet_amount}")
                
                if file_amount == sheet_amount:
                    results['matched'].append(original_art)
                elif file_amount < sheet_amount:
                    diff = sheet_amount - file_amount
                    results['missing_sizes'][original_art] = [f"Недостача {diff} шт"]
                else:
                    diff = file_amount - sheet_amount
                    results['extra_sizes'][original_art] = [f"Надлишок {diff} шт"]
                continue
            
            # Отримуємо розміри з даних таблиці
            if isinstance(sheet_art_data, dict) and 'sizes' in sheet_art_data:
                sheet_sizes_raw = sheet_art_data['sizes']
            else:
                # Старий формат (тільки sizes)
                sheet_sizes_raw = sheet_art_data
            
            # Нормалізуємо ключі розмірів для правильного порівняння
            # Створюємо новий dict з нормалізованими ключами
            sheet_sizes = {}
            for size_key, qty in sheet_sizes_raw.items():
                if size_key == '':
                    # Пропускаємо порожній розмір
                    continue
                # Очищаємо та нормалізуємо розмір
                size_key_clean = str(size_key).strip()
                normalized_key = normalize_size(size_key_clean)
                if normalized_key:  # Перевіряємо, що нормалізація дала результат
                    if normalized_key in sheet_sizes:
                        sheet_sizes[normalized_key] += qty
                    else:
                        sheet_sizes[normalized_key] = qty
            
            # Також нормалізуємо ключі в file_sizes для порівняння
            file_sizes_normalized = {}
            for size_key, qty in file_sizes.items():
                if size_key == '':
                    # Пропускаємо порожній розмір
                    continue
                # Очищаємо та нормалізуємо розмір
                size_key_clean = str(size_key).strip()
                normalized_key = normalize_size(size_key_clean)
                if normalized_key:  # Перевіряємо, що нормалізація дала результат
                    if normalized_key in file_sizes_normalized:
                        file_sizes_normalized[normalized_key] += qty
                    else:
                        file_sizes_normalized[normalized_key] = qty
            
            # ДІАГНОСТИКА: Логуємо для діагностики проблеми
            import logging
            logging.info(f"[DEBUG] Артикул: {original_art}")
            logging.info(f"[DEBUG] Розміри з файлу (оригінальні): {file_sizes}")
            logging.info(f"[DEBUG] Розміри з файлу (нормалізовані): {file_sizes_normalized}")
            logging.info(f"[DEBUG] Розміри з таблиці (оригінальні): {sheet_sizes_raw}")
            logging.info(f"[DEBUG] Розміри з таблиці (нормалізовані): {sheet_sizes}")
            
            # Для товарів без розмірів (як шапки) - sheet_sizes може бути порожнім dict()
            # Це означає, що артикул існує, але не має розмірів
            
            # Порівнюємо розміри з кількістю
            missing_sizes_list = []  # Розміри, яких немає або не вистачає
            extra_sizes_list = []    # Розміри, яких більше ніж потрібно
            
            # Перевіряємо всі розміри з таблиці
            for size, sheet_qty in sheet_sizes.items():
                file_qty = file_sizes_normalized.get(size, 0)
                if file_qty < sheet_qty:
                    # Недостача: в таблиці більше, ніж в файлі
                    diff = sheet_qty - file_qty
                    if diff > 0:
                        if diff == 1:
                            missing_sizes_list.append(size)
                        else:
                            missing_sizes_list.append(f"{size} (потрібно ще {diff})")
            
            # Перевіряємо всі розміри з файлу
            for size, file_qty in file_sizes_normalized.items():
                sheet_qty = sheet_sizes.get(size, 0)
                logging.info(f"[DEBUG] Порівняння розміру '{size}': файл={file_qty}, таблиця={sheet_qty}")
                if sheet_qty == 0:
                    # Надлишок: розмір є в файлі, але відсутній в таблиці
                    logging.info(f"[DEBUG] Розмір '{size}' відсутній в таблиці - додаємо до extra")
                    if file_qty == 1:
                        extra_sizes_list.append(size)
                    else:
                        extra_sizes_list.append(f"{size} ({file_qty} шт)")
                elif file_qty > sheet_qty:
                    # Надлишок: в файлі більше, ніж в таблиці
                    diff = file_qty - sheet_qty
                    logging.info(f"[DEBUG] Розмір '{size}': в файлі більше на {diff} - додаємо до extra")
                    if diff > 0:
                        if diff == 1:
                            extra_sizes_list.append(size)
                        else:
                            extra_sizes_list.append(f"{size} (більше на {diff})")
                elif file_qty == sheet_qty:
                    logging.info(f"[DEBUG] Розмір '{size}': співпадає (файл={file_qty}, таблиця={sheet_qty})")
                else:
                    logging.info(f"[DEBUG] Розмір '{size}': в таблиці більше (файл={file_qty}, таблиця={sheet_qty})")
            
            # Перевіряємо, чи всі розміри співпадають
            all_match = True
            # Перевіряємо, чи всі розміри з таблиці є в файлі з правильною кількістю
            for size, sheet_qty in sheet_sizes.items():
                file_qty = file_sizes_normalized.get(size, 0)
                if file_qty != sheet_qty:
                    all_match = False
                    logging.info(f"[DEBUG] all_match=False: розмір '{size}' не співпадає (файл={file_qty}, таблиця={sheet_qty})")
                    break
            # Перевіряємо, чи всі розміри з файлу є в таблиці з правильною кількістю
            if all_match:
                for size, file_qty in file_sizes_normalized.items():
                    sheet_qty = sheet_sizes.get(size, 0)
                    if file_qty != sheet_qty:
                        all_match = False
                        logging.info(f"[DEBUG] all_match=False: розмір '{size}' не співпадає (файл={file_qty}, таблиця={sheet_qty})")
                        break
            
            logging.info(f"[DEBUG] all_match={all_match}, missing_sizes_list={missing_sizes_list}, extra_sizes_list={extra_sizes_list}")
            
            # Для товарів без розмірів: якщо обидва порожні - це співпадіння
            if not file_sizes_normalized and not sheet_sizes:
                results['matched'].append(original_art)
            elif all_match and not missing_sizes_list and not extra_sizes_list:
                # Всі розміри співпадають
                results['matched'].append(original_art)
            else:
                # Якщо є недостача або надлишок - додаємо до відповідних категорій
                if missing_sizes_list:
                    results['missing_sizes'][original_art] = missing_sizes_list
                if extra_sizes_list:
                    results['extra_sizes'][original_art] = extra_sizes_list
                # Якщо немає ні недостачі, ні надлишку, але all_match = False
                # (наприклад, через порожні розміри), додаємо до matched
                if not missing_sizes_list and not extra_sizes_list:
                    results['matched'].append(original_art)
        
        # Знаходимо артикули, які є в таблицях, але немає в файлі
        for sheet_normalized_art, sheet_art_data in all_sheet_arts.items():
            if sheet_normalized_art not in file_arts_set:
                # Отримуємо оригінальний артикул та розміри
                if isinstance(sheet_art_data, dict) and 'original_art' in sheet_art_data:
                    original_sheet_art = sheet_art_data['original_art']
                    sheet_sizes_data = sheet_art_data.get('sizes', {})
                    sheet_amount = sheet_art_data.get('amount', 0)
                else:
                    # Старий формат - використовуємо normalized як original
                    original_sheet_art = sheet_normalized_art
                    sheet_sizes_data = sheet_art_data
                    sheet_amount = 0
                
                # Для товарів без розмірів використовуємо amount
                if not category_has_sizes:
                    if sheet_amount > 0:
                        # Створюємо структуру для товарів без розмірів
                        results['not_scanned'][original_sheet_art] = {'': sheet_amount}
                    continue
                
                # Пропускаємо артикули без розмірів (якщо sizes порожній dict або містить тільки порожній ключ)
                if not sheet_sizes_data or (len(sheet_sizes_data) == 1 and '' in sheet_sizes_data):
                    # Артикул без розмірів - пропускаємо
                    continue
                
                results['not_scanned'][original_sheet_art] = sheet_sizes_data
    
    return results
