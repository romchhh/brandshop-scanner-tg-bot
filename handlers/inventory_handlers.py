import os
import tempfile
from aiogram import Router, F
from aiogram.types import Message, CallbackQuery, FSInputFile, InlineKeyboardButton, InlineKeyboardMarkup
from aiogram.filters import Command, StateFilter
from aiogram.fsm.context import FSMContext
from gspread import service_account, service_account_from_dict
from config import data
from keyboards.inventory_keyboards import get_inventory_keyboard
from states.inventory_states import InventoryStates
from utils.admin_utils import check_admin
from utils.category_translations import get_category_ua
from utils.excel_generator import generate_inventory_excel, calculate_statistics
from utils.sheets_utils import (
    parse_csv_file, 
    compare_inventory_with_sheets, 
    get_category_by_prefix,
    get_art_sizes_from_sheets
)

router = Router()

from config import google_credentials

# Ініціалізація Google Sheets клієнта
# Використовуємо credentials зі змінних середовища (dict) або fallback на файл
if google_credentials:
    client = service_account_from_dict(google_credentials)
else:
    # Fallback: якщо credentials не знайдено в .env, спробуємо файл
    import os
    credentials_path = os.getenv('CREDENTIALS_PATH', 'credentials.json')
    if os.path.exists(credentials_path):
        client = service_account(filename=credentials_path)
    else:
        raise ValueError("Google credentials не знайдено! Перевірте GOOGLE_CREDENTIALS в .env або credentials.json файл.")


@router.message(Command("start"))
async def cmd_start(message: Message):
    """Обробник команди /start"""
    if not check_admin(message.from_user.id):
        await message.answer("Ви не маєте доступу до цього бота.")
        return
    
    await message.answer(
        "Оберіть тип перевірки:",
        reply_markup=get_inventory_keyboard()
    )


@router.message(F.text == "Перевірка по артах з файлу")
async def check_file_arts(message: Message, state: FSMContext):
    """Обробник кнопки 'Перевірка по артах з файлу'"""
    if not check_admin(message.from_user.id):
        await message.answer("Ви не маєте доступу до цього бота.")
        return
    
    await state.set_state(InventoryStates.waiting_file)
    await message.answer(
        "Будь ласка, надішліть CSV файл з переобліком.\n\n"
        "Файл повинен містити:\n"
        "- Артикул (2-й стовпець)\n"
        "- Розміри (4-й стовпець)\n"
        "- Кількість (6-й стовпець)"
    )


@router.message(F.text == "Перевірка одного арту")
async def check_single_art(message: Message, state: FSMContext):
    """Обробник кнопки 'Перевірка одного арту'"""
    if not check_admin(message.from_user.id):
        await message.answer("Ви не маєте доступу до цього бота.")
        return
    
    await state.set_state(InventoryStates.waiting_single_art)
    await message.answer("Введіть артикул для перевірки:")


@router.message(F.text == "Перевірка всієї категорії")
async def check_category(message: Message, state: FSMContext):
    """Обробник кнопки 'Перевірка всієї категорії'"""
    if not check_admin(message.from_user.id):
        await message.answer("Ви не маєте доступу до цього бота.")
        return
    
    await state.set_state(InventoryStates.waiting_category)
    await message.answer(
        "Будь ласка, надішліть CSV файл з переобліком.\n\n"
        "Бот автоматично визначить категорію з артикулів у файлі та звірить всі товари з таблицями.\n\n"
        "Файл повинен містити:\n"
        "- Артикул (2-й стовпець)\n"
        "- Розміри (4-й стовпець)\n"
        "- Кількість (6-й стовпець)"
    )


@router.message(StateFilter(InventoryStates.waiting_file), F.document)
async def handle_document(message: Message, state: FSMContext):
    """Обробник завантаження файлів"""
    if not check_admin(message.from_user.id):
        await message.answer("Ви не маєте доступу до цього бота.")
        return
    
    document = message.document
    
    if not document:
        await message.answer("Будь ласка, надішліть файл.")
        return
    
    # Завантажуємо файл
    file_info = await message.bot.get_file(document.file_id)
    
    # Створюємо тимчасовий файл
    file_path = tempfile.mktemp(suffix='.csv')
    
    try:
        await message.bot.download_file(file_info.file_path, file_path)
        
        # Парсимо файл
        await message.answer("Обробляю файл...")
        inventory_data = parse_csv_file(file_path)
        
        if not inventory_data:
            await message.answer("Помилка: не вдалося прочитати файл або файл порожній.")
            await state.clear()
            await message.answer(
                "Оберіть тип перевірки:",
                reply_markup=get_inventory_keyboard()
            )
            return
        
        # Порівнюємо з таблицями
        await message.answer("Порівнюю з Google таблицями...")
        results = compare_inventory_with_sheets(client, inventory_data)
        
        # Визначаємо всі категорії з файлу (можливо кілька: взуття + зимове взуття тощо)
        # Для артикулів з кількома категоріями (як "Об") враховуємо всі категорії
        category_count = {}
        for data_info in inventory_data.values():
            original_art = data_info['original_art']
            categories = get_category_by_prefix(original_art)
            if categories:
                # Додаємо всі категорії зі списку (не тільки першу)
                for category in categories:
                    category_count[category] = category_count.get(category, 0) + 1
        
        if category_count:
            # Список усіх категорій (від найчастішої до рідкісної)
            categories_sorted = sorted(category_count, key=category_count.get, reverse=True)
            categories_ua = [get_category_ua(cat) for cat in categories_sorted]
            category = categories_sorted[0]
            category_ua = categories_ua[0]
        else:
            # Якщо не вдалося визначити, беремо з першого артикулу
            first_art = list(inventory_data.values())[0]['original_art']
            categories = get_category_by_prefix(first_art)
            category = categories[0] if categories else "unknown"
            category_ua = get_category_ua(category)
            categories_ua = [category_ua]
        
        # Зберігаємо результати в стані для генерації файлу (включаючи список категорій)
        await state.update_data(
            results=results,
            inventory_data=inventory_data,
            category=category,
            category_ua=category_ua,
            categories_ua=categories_ua
        )
        
        # Створюємо мапу для підрахунку статистики
        art_map = {}
        for norm_art, data_info in inventory_data.items():
            original_art = data_info['original_art']
            art_map[original_art] = {
                'original_art': original_art,
                'sizes': data_info['sizes'],
                'original_sizes': data_info.get('original_sizes', {}),
                'amount': data_info.get('amount', 0)
            }
        
        # Підраховуємо статистику
        stats = calculate_statistics(results, inventory_data, art_map)
        
        # Формуємо результат (показуємо всі категорії з файлу)
        categories_display = ", ".join(categories_ua)
        message_parts = []
        message_parts.append(f"📋 Категорії: {categories_display}\n")
        message_parts.append(f"📊 Всього артикулів у файлі: {len(inventory_data)}\n\n")
        
        # Додаємо статистику
        message_parts.append("📈 СТАТИСТИКА:\n")
        message_parts.append(f"Розмірів: {stats['total_sizes']}\n")
        message_parts.append(f"Сошлося: {stats['matched_sizes']}\n")
        message_parts.append(f"Недостача: {stats['missing_sizes']}\n")
        message_parts.append(f"Не відскановано: {stats['not_scanned_sizes']}\n\n")
        
        if results['missing_sizes']:
            message_parts.append("❌ НЕДОСТАЧА РОЗМІРІВ:")
            for art, sizes in results['missing_sizes'].items():
                message_parts.append(f"\n{art}: {', '.join(sizes)}")
        
        if results['extra_sizes']:
            message_parts.append("\n\n✅ НАДЛИШОК РОЗМІРІВ:")
            for art, sizes in results['extra_sizes'].items():
                message_parts.append(f"\n{art}: {', '.join(sizes)}")
        
        if results['not_found']:
            message_parts.append(f"\n\n⚠️ НЕ ЗНАЙДЕНО В ТАБЛИЦЯХ (є в файлі скану) ({len(results['not_found'])}):")
            message_parts.append(f"{', '.join(results['not_found'][:10])}")
            if len(results['not_found']) > 10:
                message_parts.append(f"\n... та ще {len(results['not_found']) - 10} артикулів")
        
        if results['matched']:
            message_parts.append(f"\n\n✓ СПІВПАДАЮТЬ ({len(results['matched'])} артикулів)")
        
        result_message = ''.join(message_parts) if message_parts else f"Всі артикули ({categories_display}) співпадають!"
        
        # Створюємо інлайн кнопку
        keyboard = [
            [InlineKeyboardButton(text="📥 Отримати файл", callback_data="get_excel_file")]
        ]
        reply_markup = InlineKeyboardMarkup(inline_keyboard=keyboard)
        
        # Розбиваємо повідомлення на частини, якщо воно занадто довге
        max_length = 4000
        if len(result_message) > max_length:
            parts = [result_message[i:i+max_length] for i in range(0, len(result_message), max_length)]
            for i, part in enumerate(parts):
                if i == len(parts) - 1:
                    await message.answer(part, reply_markup=reply_markup)
                else:
                    await message.answer(part)
        else:
            await message.answer(result_message, reply_markup=reply_markup)
        
        # Не очищаємо стан, щоб можна було згенерувати файл
        
    except Exception as e:
        await message.answer(f"Помилка при обробці файлу: {str(e)}")
        await state.clear()
        await message.answer(
            "Оберіть тип перевірки:",
            reply_markup=get_inventory_keyboard()
        )
    finally:
        # Видаляємо тимчасовий файл
        if os.path.exists(file_path):
            os.remove(file_path)


@router.message(StateFilter(InventoryStates.waiting_single_art), F.text)
async def handle_single_art(message: Message, state: FSMContext):
    """Обробник введення артикулу для перевірки"""
    if not check_admin(message.from_user.id):
        await message.answer("Ви не маєте доступу до цього бота.")
        return
    
    art = message.text.strip()
    
    # Визначаємо категорію
    categories = get_category_by_prefix(art)
    
    if not categories:
        await message.answer(f"Не вдалося визначити категорію для артикулу: {art}")
        await state.clear()
        await message.answer(
            "Оберіть тип перевірки:",
            reply_markup=get_inventory_keyboard()
        )
        return
    
    await message.answer("Перевіряю артикул в таблицях...")
    
    # Отримуємо розміри з таблиць
    sheet_sizes = get_art_sizes_from_sheets(client, art, categories)
    
    # Перекладаємо категорії на українську
    category = categories[0]
    category_ua = get_category_ua(category)
    categories_ua = [get_category_ua(cat) for cat in categories]
    
    if not sheet_sizes:
        await message.answer(
            f"📋 Артикул: {art}\n"
            f"📂 Категорія: {category_ua}\n"
            f"❌ Артикул не знайдено в таблицях."
        )
    else:
        sizes_list = sorted(list(sheet_sizes))
        if sizes_list:
            sizes_str = ', '.join(sizes_list)
        else:
            sizes_str = "Без розмірів (товар без розмірів)"
        
        await message.answer(
            f"📋 Артикул: {art}\n"
            f"📂 Категорія: {category_ua}\n"
            f"📏 Розміри в таблицях: {sizes_str}"
        )
    
    await state.clear()
    
    # Повертаємо до головного меню
    await message.answer(
        "Оберіть тип перевірки:",
        reply_markup=get_inventory_keyboard()
    )


@router.callback_query(F.data == "get_excel_file")
async def get_excel_file(callback: CallbackQuery, state: FSMContext):
    """Обробник кнопки 'Отримати файл'"""
    if not check_admin(callback.from_user.id):
        await callback.answer("Ви не маєте доступу до цього бота.", show_alert=True)
        return
    
    # Отримуємо дані зі стану
    state_data = await state.get_data()
    
    if not state_data or 'results' not in state_data:
        await callback.answer("Дані не знайдено. Будь ласка, виконайте перевірку спочатку.", show_alert=True)
        return
    
    results = state_data['results']
    inventory_data = state_data['inventory_data']
    # Підтримка кількох категорій: categories_ua — список, інакше fallback на одну категорію
    categories_ua = state_data.get('categories_ua')
    if not categories_ua:
        categories_ua = [state_data.get('category_ua', 'Невідома категорія')]
    if isinstance(categories_ua, str):
        categories_ua = [categories_ua]
    categories_display = ", ".join(categories_ua)
    filename_safe = "_".join(c.replace(" ", "_") for c in categories_ua)
    
    try:
        await callback.answer("Генерую файл...")
        
        # Генеруємо Excel файл (передаємо рядок з усіма категоріями для підпису)
        excel_path = generate_inventory_excel(results, inventory_data, categories_display)
        
        # Відправляємо файл з назвою за категоріями
        file = FSInputFile(excel_path, filename=f"переоблік_{filename_safe}.xlsx")
        await callback.message.answer_document(file, caption=f"📊 Результати переобліку: {categories_display}")
        
        # Видаляємо тимчасовий файл
        if os.path.exists(excel_path):
            os.remove(excel_path)
        
        # Очищаємо стан після відправки файлу
        await state.clear()
        
    except Exception as e:
        await callback.answer(f"Помилка при генерації файлу: {str(e)}", show_alert=True)


@router.message(StateFilter(InventoryStates.waiting_category), F.document)
async def handle_category_file(message: Message, state: FSMContext):
    """Обробник завантаження файлу для перевірки всієї категорії"""
    if not check_admin(message.from_user.id):
        await message.answer("Ви не маєте доступу до цього бота.")
        return
    
    document = message.document
    
    if not document:
        await message.answer("Будь ласка, надішліть файл.")
        return
    
    # Завантажуємо файл
    file_info = await message.bot.get_file(document.file_id)
    
    # Створюємо тимчасовий файл
    file_path = tempfile.mktemp(suffix='.csv')
    
    try:
        await message.bot.download_file(file_info.file_path, file_path)
        
        # Парсимо файл
        await message.answer("Обробляю файл...")
        inventory_data = parse_csv_file(file_path)
        
        if not inventory_data:
            await message.answer("Помилка: не вдалося прочитати файл або файл порожній.")
            await state.clear()
            await message.answer(
                "Оберіть тип перевірки:",
                reply_markup=get_inventory_keyboard()
            )
            return
        
        # Визначаємо всі категорії з файлу (можливо кілька)
        # Для артикулів з кількома категоріями (як "Об") враховуємо всі категорії
        category_count = {}
        for data_info in inventory_data.values():
            original_art = data_info['original_art']
            categories = get_category_by_prefix(original_art)
            if categories:
                # Додаємо всі категорії зі списку (не тільки першу)
                for category in categories:
                    category_count[category] = category_count.get(category, 0) + 1
        
        if not category_count:
            first_art = list(inventory_data.values())[0]['original_art']
            await message.answer(
                f"Не вдалося визначити категорію з артикулів у файлі.\n"
                f"Перший артикул: {first_art}\n"
                f"Перевірте, чи правильно заповнений файл."
            )
            await state.clear()
            await message.answer(
                "Оберіть тип перевірки:",
                reply_markup=get_inventory_keyboard()
            )
            return
        
        # Список усіх категорій (від найчастішої до рідкісної)
        categories_sorted = sorted(category_count, key=category_count.get, reverse=True)
        categories_ua = [get_category_ua(cat) for cat in categories_sorted]
        category = categories_sorted[0]
        category_ua = categories_ua[0]
        categories_display = ", ".join(categories_ua)
        
        await message.answer(f"Визначено категорії: {categories_display}\nПорівнюю з Google таблицями...")
        
        # Порівнюємо з таблицями
        results = compare_inventory_with_sheets(client, inventory_data)
        
        # Зберігаємо результати в стані для генерації файлу (включаючи список категорій)
        await state.update_data(
            results=results,
            inventory_data=inventory_data,
            category=category,
            category_ua=category_ua,
            categories_ua=categories_ua
        )
        
        # Створюємо мапу для підрахунку статистики
        art_map = {}
        for norm_art, data_info in inventory_data.items():
            original_art = data_info['original_art']
            art_map[original_art] = {
                'original_art': original_art,
                'sizes': data_info['sizes'],
                'original_sizes': data_info.get('original_sizes', {}),
                'amount': data_info.get('amount', 0)
            }
        
        # Підраховуємо статистику
        stats = calculate_statistics(results, inventory_data, art_map)
        
        # Формуємо результат (показуємо всі категорії)
        message_parts = []
        message_parts.append(f"📋 Категорії: {categories_display}\n")
        message_parts.append(f"📊 Всього артикулів у файлі: {len(inventory_data)}\n\n")
        
        # Додаємо статистику
        message_parts.append("📈 СТАТИСТИКА:\n")
        message_parts.append(f"Розмірів: {stats['total_sizes']}\n")
        message_parts.append(f"Сошлося: {stats['matched_sizes']}\n")
        message_parts.append(f"Недостача: {stats['missing_sizes']}\n")
        message_parts.append(f"Не відскановано: {stats['not_scanned_sizes']}\n\n")
        
        if results['missing_sizes']:
            message_parts.append("❌ НЕДОСТАЧА РОЗМІРІВ:")
            for art, sizes in results['missing_sizes'].items():
                message_parts.append(f"\n{art}: {', '.join(sizes)}")
        
        if results['extra_sizes']:
            message_parts.append("\n\n✅ НАДЛИШОК РОЗМІРІВ:")
            for art, sizes in results['extra_sizes'].items():
                message_parts.append(f"\n{art}: {', '.join(sizes)}")
        
        if results['not_found']:
            message_parts.append(f"\n\n⚠️ НЕ ЗНАЙДЕНО В ТАБЛИЦЯХ (є в файлі скану) ({len(results['not_found'])}):")
            message_parts.append(f"{', '.join(results['not_found'][:10])}")
            if len(results['not_found']) > 10:
                message_parts.append(f"\n... та ще {len(results['not_found']) - 10} артикулів")
        
        if results['matched']:
            message_parts.append(f"\n\n✓ СПІВПАДАЮТЬ ({len(results['matched'])} артикулів)")
        
        result_message = ''.join(message_parts) if message_parts else f"Всі артикули ({categories_display}) співпадають!"
        
        # Створюємо інлайн кнопку
        keyboard = [
            [InlineKeyboardButton(text="📥 Отримати файл", callback_data="get_excel_file")]
        ]
        reply_markup = InlineKeyboardMarkup(inline_keyboard=keyboard)
        
        # Розбиваємо повідомлення на частини, якщо воно занадто довге
        max_length = 4000
        if len(result_message) > max_length:
            parts = [result_message[i:i+max_length] for i in range(0, len(result_message), max_length)]
            for i, part in enumerate(parts):
                if i == len(parts) - 1:
                    await message.answer(part, reply_markup=reply_markup)
                else:
                    await message.answer(part)
        else:
            await message.answer(result_message, reply_markup=reply_markup)
        
        # Не очищаємо стан, щоб можна було згенерувати файл
        
    except Exception as e:
        await message.answer(f"Помилка при обробці файлу: {str(e)}")
        await state.clear()
        await message.answer(
            "Оберіть тип перевірки:",
            reply_markup=get_inventory_keyboard()
        )
    finally:
        # Видаляємо тимчасовий файл
        if os.path.exists(file_path):
            os.remove(file_path)
