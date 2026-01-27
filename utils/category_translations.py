# Переклад категорій на українську мову
CATEGORY_TRANSLATIONS = {
    "jeans": "Джинси",
    "sweaters": "Кофти",
    "shoes": "Взуття",
    "wintershoes": "Зимове взуття",
    "tapki": "Тапки",
    "costumes": "Костюми",
    "costumes_fleece": "Костюми флісові",
    "costumes_summer": "Костюми літні",
    "jackets": "Куртки",
    "waistcoats": "Жилети",
    "trousers": "Штани",
    "sport_trousers": "Штани спортивні",
    "tshirts_polo": "Футболки/Поло",
    "shirts": "Сорочки",
    "underwear": "Білизна",
    "shorts_jeans": "Шорти джинсові",
    "shorts_textile": "Шорти текстильні",
    "shorts_swim": "Шорти для плавання",
    "bags": "Сумки",
    "purses": "Клатчі",
    "belts": "Ремені",
    "caps": "Кепки",
    "hats": "Шапки",
    "socks": "Шкарпетки",
    "glasses": "Окуляри"
}


def get_category_ua(category_en: str) -> str:
    """Повертає українську назву категорії"""
    return CATEGORY_TRANSLATIONS.get(category_en, category_en)
