from aiogram.types import ReplyKeyboardMarkup, KeyboardButton


def get_inventory_keyboard():
    keyboard = [
        [KeyboardButton(text="Перевірка всієї категорії")],
        [KeyboardButton(text="Перевірка по артах з файлу")],
        [KeyboardButton(text="Перевірка одного арту")]
    ]
    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)
