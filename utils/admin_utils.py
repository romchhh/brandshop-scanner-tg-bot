from config import administrators


def check_admin(user_id: int) -> bool:
    """Перевіряє, чи є користувач адміністратором"""
    return user_id in administrators
