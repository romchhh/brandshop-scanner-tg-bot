import tempfile
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


def generate_inventory_excel(results, inventory_data, category_ua):
    """
    Генерує Excel файл з результатами перевірки
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Переоблік"
    
    # Заголовки
    headers = ["Арт", "Розміри", "Недостача", "Товар +"]
    ws.append(headers)
    
    # Стилі для заголовків
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Збираємо дані для файлу
    file_rows = []
    
    # Створюємо мапу для швидкого пошуку по original_art (бо в results ключі - це original_art)
    art_map = {}
    for norm_art, data_info in inventory_data.items():
        original_art = data_info['original_art']
        # sizes тепер це dict {розмір: кількість}
        art_map[original_art] = {
            'original_art': original_art,
            'sizes': data_info['sizes']  # {розмір: кількість}
        }
    
    # 1. Рядки з надлишком (зелені) - на початку
    for original_art, sizes in results.get('extra_sizes', {}).items():
        if original_art in art_map:
            data = art_map[original_art]
            # sizes тепер це dict {розмір: кількість}
            if data['sizes']:
                sizes_list = []
                for size, qty in sorted(data['sizes'].items()):
                    if size:  # Пропускаємо порожній розмір
                        sizes_list.append(f"{size} ({qty})" if qty > 1 else size)
                sizes_str = ', '.join(sizes_list) if sizes_list else 'Без розмірів'
            else:
                sizes_str = 'Без розмірів'
            # sizes - це список рядків з описом надлишку
            if isinstance(sizes, list):
                extra_str = ', '.join(sizes) if sizes else ''
            else:
                extra_str = ', '.join(sorted(sizes)) if sizes else ''
            file_rows.append({
                'art': data['original_art'],
                'sizes': sizes_str,
                'missing': '',
                'extra': extra_str,
                'type': 'extra'  # Зелений
            })
    
    # 2. Рядки зі співпадінням (білі) - посередині
    for original_art in results.get('matched', []):
        if original_art in art_map:
            data = art_map[original_art]
            # sizes тепер це dict {розмір: кількість}
            if data['sizes']:
                sizes_list = []
                for size, qty in sorted(data['sizes'].items()):
                    if size:  # Пропускаємо порожній розмір
                        sizes_list.append(f"{size} ({qty})" if qty > 1 else size)
                sizes_str = ', '.join(sizes_list) if sizes_list else 'Без розмірів'
            else:
                sizes_str = 'Без розмірів'  # Для товарів без розмірів
            file_rows.append({
                'art': data['original_art'],
                'sizes': sizes_str,
                'missing': '',
                'extra': '',
                'type': 'matched'  # Білий
            })
    
    # 3. Рядки з недостачею (червоні) - в кінці
    for original_art, sizes in results.get('missing_sizes', {}).items():
        if original_art in art_map:
            data = art_map[original_art]
            # sizes тепер це dict {розмір: кількість}
            if data['sizes']:
                sizes_list = []
                for size, qty in sorted(data['sizes'].items()):
                    if size:  # Пропускаємо порожній розмір
                        sizes_list.append(f"{size} ({qty})" if qty > 1 else size)
                sizes_str = ', '.join(sizes_list) if sizes_list else 'Без розмірів'
            else:
                sizes_str = 'Без розмірів'
            # sizes - це список рядків з описом недостачі
            if isinstance(sizes, list):
                missing_str = ', '.join(sizes) if sizes else ''
            else:
                missing_str = ', '.join(sorted(sizes)) if sizes else ''
            file_rows.append({
                'art': data['original_art'],
                'sizes': sizes_str,
                'missing': missing_str,
                'extra': '',
                'type': 'missing'  # Червоний
            })
    
    # 4. Рядки не знайдені (червоні) - в кінці
    for art in results.get('not_found', []):
        # Шукаємо в inventory_data за оригінальним артикулом
        sizes_str = ''
        for norm_art, data_info in inventory_data.items():
            if data_info['original_art'] == art:
                # sizes тепер це dict {розмір: кількість}
                if data_info['sizes']:
                    sizes_list = []
                    for size, qty in sorted(data_info['sizes'].items()):
                        if size:  # Пропускаємо порожній розмір
                            sizes_list.append(f"{size} ({qty})" if qty > 1 else size)
                    sizes_str = ', '.join(sizes_list) if sizes_list else 'Без розмірів'
                else:
                    sizes_str = 'Без розмірів'
                break
        
        file_rows.append({
            'art': art,
            'sizes': sizes_str,
            'missing': 'Не знайдено в таблицях',
            'extra': '',
            'type': 'not_found'  # Червоний
        })
    
    # Перевіряємо, чи всі артикули з файлу додані до Excel
    # Збираємо всі артикули, які вже додані
    added_arts = set()
    for row in file_rows:
        added_arts.add(row['art'])
    
    # Додаємо артикули, які не потрапили в жодну категорію (на всяк випадок)
    for norm_art, data_info in inventory_data.items():
        original_art = data_info['original_art']
        if original_art not in added_arts:
            # Якщо артикул не доданий, додаємо його як співпадіння (білий)
            # sizes тепер це dict {розмір: кількість}
            if data_info['sizes']:
                sizes_list = []
                for size, qty in sorted(data_info['sizes'].items()):
                    if size:  # Пропускаємо порожній розмір
                        sizes_list.append(f"{size} ({qty})" if qty > 1 else size)
                sizes_str = ', '.join(sizes_list) if sizes_list else 'Без розмірів'
            else:
                sizes_str = 'Без розмірів'
            file_rows.append({
                'art': original_art,
                'sizes': sizes_str,
                'missing': '',
                'extra': '',
                'type': 'matched'  # Білий - все нормально
            })
    
    # Додаємо рядки до таблиці
    for row_data in file_rows:
        ws.append([
            row_data['art'],
            row_data['sizes'],
            row_data['missing'],
            row_data.get('extra', '')  # Товар + (надлишок)
        ])
    
    # Застосовуємо форматування
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Світло-зелений
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Світло-червоний
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # Білий
    
    # Застосовуємо кольори до рядків
    for idx, row_data in enumerate(file_rows):
        row_num = idx + 2  # Починаємо з 2, бо 1 - заголовок
        row_type = row_data['type']
        
        if row_type == 'extra':
            fill = green_fill
        elif row_type in ['missing', 'not_found']:
            fill = red_fill
        elif row_type == 'matched':
            fill = white_fill  # Явно встановлюємо білий для співпадінь
        else:
            fill = white_fill  # За замовчуванням білий
        
        # Застосовуємо кольори до всіх комірок рядка
        for col_num in range(1, 5):  # 4 стовпці
            cell = ws.cell(row=row_num, column=col_num)
            cell.fill = fill
            cell.alignment = Alignment(horizontal="left", vertical="center")
    
    # Налаштування ширини стовпців
    ws.column_dimensions['A'].width = 20  # Арт
    ws.column_dimensions['B'].width = 30  # Розміри
    ws.column_dimensions['C'].width = 30  # Недостача
    ws.column_dimensions['D'].width = 30  # Товар +
    
    # Зберігаємо файл
    file_path = tempfile.mktemp(suffix='.xlsx')
    wb.save(file_path)
    
    return file_path
