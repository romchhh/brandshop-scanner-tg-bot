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



def normalize_size(size):
    """Нормалізує розмір, конвертуючи кирилицю в латиницю"""
    if not size:
        return ""
    
    size = size.upper().strip()
    
    # Конвертація кирилиці в латиницю для розмірів
    cyrillic_to_latin = {
        'М': 'M',
        'Л': 'L',
        'ХЛ': 'XL',
        '2ХЛ': '2XL',
        '3ХЛ': '3XL',
        '4ХЛ': '4XL',
        '5ХЛ': '5XL',
        '6ХЛ': '6XL',
        '7ХЛ': '7XL',
        'С': 'S',
    }
    
    # Перевіряємо, чи це повний розмір у кирилиці
    if size in cyrillic_to_latin:
        return cyrillic_to_latin[size]
    
    # Замінюємо кириличні символи в рядку
    result = size
    for cyr, lat in cyrillic_to_latin.items():
        result = result.replace(cyr, lat)
    
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
        elif original_size in {"S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL"}:
            size_order = ["S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL"]
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
        "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL",
        "39", "40", "41", "42", "43", "44", "45", "46"
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
                    filter_column = [row[details["filter"] - 1] for row in all_data]

                    for i, row_size in enumerate(size_column):
                        filter_value = filter_column[i].strip()
                        if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')):
                            row_sizes = split_sizes(row_size)
                            adjusted_sizes = {adjust_size(size, filter_value) for size in row_sizes}
                            valid_sizes = {size for size in adjusted_sizes if size in allsize_values}
                        else:
                            row_sizes = split_sizes(row_size)
                            valid_sizes = {size for size in row_sizes if size in allsize_values}

                        if size == "allsize":
                            if len(valid_sizes) >= 3:
                                result.append({
                                    'category': category,
                                    'price': price_column[i],
                                    'art': art_column[i],
                                    'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
                                    'photo': photo_column[i]
                                })
                        elif size == "nosize":
                            if amount_column[i].strip():
                                result.append({
                                    'category': category,
                                    'price': price_column[i],
                                    'art': art_column[i],
                                    'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
                                    'photo': photo_column[i]
                                })
                        elif normalized_size in valid_sizes:
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
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
                filter_column = [row[details["filter"] - 1] for row in all_data]

                for i, row_size in enumerate(size_column):
                    filter_value = filter_column[i].strip()
                    if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')):
                        row_sizes = split_sizes(row_size)
                        adjusted_sizes = {adjust_size(size, filter_value) for size in row_sizes}
                        valid_sizes = {size for size in adjusted_sizes if size in allsize_values}
                    else:
                        row_sizes = split_sizes(row_size)
                        valid_sizes = {size for size in row_sizes if size in allsize_values}

                    if size == "allsize":
                        if len(valid_sizes) >= 3:
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
                                'photo': photo_column[i]
                            })
                    elif size == "nosize":
                        if amount_column[i].strip():
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
                                'photo': photo_column[i]
                            })
                    elif normalized_size in valid_sizes:
                        result.append({
                            'category': category,
                            'price': price_column[i],
                            'art': art_column[i],
                            'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
                            'photo': photo_column[i]
                        })
    return result



def find_item_by_size_in_category_drop(client, size, categories):
    result = []
    normalized_size = normalize_size(size)

    allsize_values = {
        "28", "29", "30", "31", "32", "33", "34", "35", "36", "38", "40", "42", "44",
        "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL",
        "39", "40", "41", "42", "43", "44", "45", "46"
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
                    filter_column = [row[details["filter"] - 1] for row in all_data]

                    for i, row_size in enumerate(size_column):
                        filter_value = filter_column[i].strip()
                        if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')):
                            row_sizes = split_sizes(row_size)
                            adjusted_sizes = {adjust_size(size, filter_value) for size in row_sizes}
                            valid_sizes = {size for size in adjusted_sizes if size in allsize_values}
                        else:
                            row_sizes = split_sizes(row_size)
                            valid_sizes = {size for size in row_sizes if size in allsize_values}

                        if size == "allsize":
                            if len(valid_sizes) >= 3:
                                result.append({
                                    'category': category,
                                    'price': price_column[i],
                                    'art': art_column[i],
                                    'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
                                    'photo': photo_column[i]
                                })
                        elif size == "nosize":
                            if amount_column[i].strip():
                                result.append({
                                    'category': category,
                                    'price': price_column[i],
                                    'art': art_column[i],
                                    'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
                                    'photo': photo_column[i]
                                })
                        elif normalized_size in valid_sizes:
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
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
                filter_column = [row[details["filter"] - 1] for row in all_data]

                for i, row_size in enumerate(size_column):
                    filter_value = filter_column[i].strip()
                    if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')):
                        row_sizes = split_sizes(row_size)
                        adjusted_sizes = {adjust_size(size, filter_value) for size in row_sizes}
                        valid_sizes = {size for size in adjusted_sizes if size in allsize_values}
                    else:
                        row_sizes = split_sizes(row_size)
                        valid_sizes = {size for size in row_sizes if size in allsize_values}

                    if size == "allsize":
                        if len(valid_sizes) >= 3:
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
                                'photo': photo_column[i]
                            })
                    elif size == "nosize":
                        if amount_column[i].strip():
                            result.append({
                                'category': category,
                                'price': price_column[i],
                                'art': art_column[i],
                                'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
                                'photo': photo_column[i]
                            })
                    elif normalized_size in valid_sizes:
                        result.append({
                            'category': category,
                            'price': price_column[i],
                            'art': art_column[i],
                            'size': ', '.join(adjusted_sizes) if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')) else size_column[i],
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


def extract_base_art(art):
    """Витягує базовий артикул, видаляючи розмір з кінця"""
    if not art:
        return art
    
    # Видаляємо пробіли та дефіси для нормалізації
    art = art.strip()
    
    # Спроба видалити розмір з кінця (формат типу "Дж-553.38" або "Ко-450.38")
    # Шукаємо останню крапку або дефіс перед числом
    import re
    # Видаляємо розмір після останньої крапки або дефісу, якщо це число
    base_art = re.sub(r'[.\-]\d+$', '', art)
    
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
                # Кількість з шостого стовпця (індекс 5) - "Відскановано"
                amount = row[5].strip() if len(row) > 5 else "1"
                
                if not art:
                    continue
                
                # Для товарів БЕЗ розмірів (як шапки) - зберігаємо повний артикул
                # Для товарів З розмірами - витягуємо базовий артикул (групуємо по артикулу)
                if size_info and size_info.strip():
                    # Є розмір - витягуємо базовий артикул для групування
                    base_art = extract_base_art(art)
                else:
                    # Немає розміру - зберігаємо повний артикул
                    base_art = art
                
                # Нормалізуємо артикул для порівняння
                normalized_art = normalize_art(base_art)
                
                if normalized_art not in inventory_data:
                    inventory_data[normalized_art] = {
                        'sizes': {},  # {розмір: кількість}
                        'amount': 0,
                        'original_art': base_art  # Зберігаємо базовий або повний артикул
                    }
                
                # Додаємо розмір з колонки "Інформація" з кількістю
                if size_info and size_info.strip():
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
                else:
                    # Для товарів без розмірів - додаємо загальну кількість
                    try:
                        size_amount = int(amount) if amount else 1
                    except ValueError:
                        size_amount = 1
                    # Для товарів без розмірів використовуємо порожній рядок як ключ
                    if '' in inventory_data[normalized_art]['sizes']:
                        inventory_data[normalized_art]['sizes'][''] += size_amount
                    else:
                        inventory_data[normalized_art]['sizes'][''] = size_amount
                
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
        "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL",
        "39", "40", "41", "42", "43", "44", "45", "46"
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
                        filter_column = [
                            row[details["filter"] - 1] if len(row) > details["filter"] - 1 else ""
                            for row in all_data
                        ]
                        
                        for i, row_art in enumerate(art_column):
                            if normalize_art(row_art) == normalized_art:
                                row_size = size_column[i] if i < len(size_column) else "-"
                                filter_value = filter_column[i].strip() if i < len(filter_column) else ""
                                
                                if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')):
                                    row_sizes = split_sizes(row_size)
                                    adjusted_sizes = {adjust_size(size, filter_value) for size in row_sizes}
                                    valid_sizes = {size for size in adjusted_sizes if size in allsize_values}
                                else:
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
                    filter_column = [
                        row[details["filter"] - 1] if len(row) > details["filter"] - 1 else ""
                        for row in all_data
                    ]
                    
                    for i, row_art in enumerate(art_column):
                        if normalize_art(row_art) == normalized_art:
                            row_size = size_column[i] if i < len(size_column) else "-"
                            filter_value = filter_column[i].strip() if i < len(filter_column) else ""
                            
                            if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')):
                                row_sizes = split_sizes(row_size)
                                adjusted_sizes = {adjust_size(size, filter_value) for size in row_sizes}
                                valid_sizes = {size for size in adjusted_sizes if size in allsize_values}
                            else:
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
    Повертає словник: {нормалізований_артикул: set(розміри)}
    """
    all_arts_data = {}
    
    if category not in data:
        return all_arts_data
    
    allsize_values = {
        "28", "29", "30", "31", "32", "33", "34", "35", "36", "38", "40", "42", "44",
        "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL",
        "39", "40", "41", "42", "43", "44", "45", "46"
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
                    filter_column = [
                        row[details["filter"] - 1] if len(row) > details["filter"] - 1 else ""
                        for row in all_data
                    ]
                    
                    for i, row_art in enumerate(art_column):
                        if not row_art or not row_art.strip():
                            continue
                        
                        normalized_art = normalize_art(row_art)
                        row_size = size_column[i] if i < len(size_column) else "-"
                        filter_value = filter_column[i].strip() if i < len(filter_column) else ""
                        
                        # Ініціалізуємо set для артикулу, якщо його ще немає
                        if normalized_art not in all_arts_data:
                            all_arts_data[normalized_art] = set()
                        
                    # Ініціалізуємо dict для артикулу, якщо його ще немає
                    # Структура: {розмір: кількість}
                    if normalized_art not in all_arts_data:
                        all_arts_data[normalized_art] = {}
                    
                    # Обробляємо розміри тільки якщо вони є
                    if row_size and row_size != "-" and row_size.strip():
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
                                    if current_size in all_arts_data[normalized_art]:
                                        all_arts_data[normalized_art][current_size] += quantity
                                    else:
                                        all_arts_data[normalized_art][current_size] = quantity
                                current_size = None
                                continue
                            
                            # Пропускаємо фільтри (починаються з / і містять тільки цифри)
                            if size_item.startswith('/') and size_item[1:].isdigit():
                                continue
                            
                            # Це розмір
                            # Якщо був попередній розмір без кількості, додаємо його з кількістю 1
                            if current_size is not None:
                                if current_size in all_arts_data[normalized_art]:
                                    all_arts_data[normalized_art][current_size] += 1
                                else:
                                    all_arts_data[normalized_art][current_size] = 1
                            
                            # Застосовуємо фільтр, якщо є
                            if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')):
                                adjusted_size = adjust_size(size_item, filter_value)
                            else:
                                adjusted_size = size_item
                            
                            # Перевіряємо, чи це валідний розмір
                            if adjusted_size in allsize_values:
                                current_size = adjusted_size
                            else:
                                current_size = None
                        
                        # Додаємо останній розмір, якщо він не мав кількості
                        if current_size is not None:
                            if current_size in all_arts_data[normalized_art]:
                                all_arts_data[normalized_art][current_size] += 1
                            else:
                                all_arts_data[normalized_art][current_size] = 1
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
                filter_column = [
                    row[details["filter"] - 1] if len(row) > details["filter"] - 1 else ""
                    for row in all_data
                ]
                
                for i, row_art in enumerate(art_column):
                    if not row_art or not row_art.strip():
                        continue
                    
                    normalized_art = normalize_art(row_art)
                    row_size = size_column[i] if i < len(size_column) else "-"
                    filter_value = filter_column[i].strip() if i < len(filter_column) else ""
                    
                    # Ініціалізуємо dict для артикулу, якщо його ще немає
                    # Структура: {розмір: кількість}
                    if normalized_art not in all_arts_data:
                        all_arts_data[normalized_art] = {}
                    
                    # Обробляємо розміри тільки якщо вони є
                    if row_size and row_size != "-" and row_size.strip():
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
                                    if current_size in all_arts_data[normalized_art]:
                                        all_arts_data[normalized_art][current_size] += quantity
                                    else:
                                        all_arts_data[normalized_art][current_size] = quantity
                                current_size = None
                                continue
                            
                            # Пропускаємо фільтри (починаються з / і містять тільки цифри)
                            if size_item.startswith('/') and size_item[1:].isdigit():
                                continue
                            
                            # Це розмір
                            # Якщо був попередній розмір без кількості, додаємо його з кількістю 1
                            if current_size is not None:
                                if current_size in all_arts_data[normalized_art]:
                                    all_arts_data[normalized_art][current_size] += 1
                                else:
                                    all_arts_data[normalized_art][current_size] = 1
                            
                            # Застосовуємо фільтр, якщо є
                            if filter_value and (filter_value.startswith('-') or filter_value.startswith('/')):
                                adjusted_size = adjust_size(size_item, filter_value)
                            else:
                                adjusted_size = size_item
                            
                            # Перевіряємо, чи це валідний розмір
                            if adjusted_size in allsize_values:
                                current_size = adjusted_size
                            else:
                                current_size = None
                        
                        # Додаємо останній розмір, якщо він не мав кількості
                        if current_size is not None:
                            if current_size in all_arts_data[normalized_art]:
                                all_arts_data[normalized_art][current_size] += 1
                            else:
                                all_arts_data[normalized_art][current_size] = 1
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
        'matched': []         # Артикули які повністю співпадають
    }
    
    # Групуємо артикули по категоріях
    arts_by_category = {}
    for normalized_art, data_info in inventory_data.items():
        original_art = data_info['original_art']
        categories = get_category_by_prefix(original_art)
        
        if not categories:
            results['not_found'].append(original_art)
            continue
        
        # Використовуємо першу категорію
        category = categories[0]
        if category not in arts_by_category:
            arts_by_category[category] = []
        arts_by_category[category].append((normalized_art, data_info))
    
    # Для кожної категорії зчитуємо всі дані один раз
    for category, arts_list in arts_by_category.items():
        # Зчитуємо всі артикули з категорії одним запитом
        all_sheet_arts = load_all_arts_from_category(client, category)
        
        # Порівнюємо кожен артикул з файлу
        for normalized_art, data_info in arts_list:
            original_art = data_info['original_art']
            file_sizes = data_info['sizes']  # {розмір: кількість}
            
            # Шукаємо артикул в завантажених даних
            # Використовуємо None як маркер, що артикул не знайдено
            sheet_sizes = all_sheet_arts.get(normalized_art, None)
            
            if sheet_sizes is None:
                results['not_found'].append(original_art)
                continue
            
            # Для товарів без розмірів (як шапки) - sheet_sizes може бути порожнім dict()
            # Це означає, що артикул існує, але не має розмірів
            
            # Порівнюємо розміри з кількістю
            missing_sizes_list = []  # Розміри, яких немає або не вистачає
            extra_sizes_list = []    # Розміри, яких більше ніж потрібно
            
            # Перевіряємо всі розміри з таблиці
            for size, sheet_qty in sheet_sizes.items():
                # Пропускаємо порожній розмір для товарів без розмірів
                if size == '':
                    continue
                file_qty = file_sizes.get(size, 0)
                if file_qty < sheet_qty:
                    # Недостача: в таблиці більше, ніж в файлі
                    diff = sheet_qty - file_qty
                    if diff > 0:
                        if diff == 1:
                            missing_sizes_list.append(size)
                        else:
                            missing_sizes_list.append(f"{size} (потрібно ще {diff})")
            
            # Перевіряємо всі розміри з файлу
            for size, file_qty in file_sizes.items():
                # Пропускаємо порожній розмір для товарів без розмірів
                if size == '':
                    continue
                sheet_qty = sheet_sizes.get(size, 0)
                if file_qty > sheet_qty:
                    # Надлишок: в файлі більше, ніж в таблиці
                    diff = file_qty - sheet_qty
                    if diff > 0:
                        if diff == 1:
                            extra_sizes_list.append(size)
                        else:
                            extra_sizes_list.append(f"{size} (більше на {diff})")
            
            # Для товарів без розмірів: якщо обидва порожні - це співпадіння
            if not file_sizes and not sheet_sizes:
                results['matched'].append(original_art)
            elif missing_sizes_list:
                results['missing_sizes'][original_art] = missing_sizes_list
            elif extra_sizes_list:
                results['extra_sizes'][original_art] = extra_sizes_list
            else:
                results['matched'].append(original_art)
    
    return results
