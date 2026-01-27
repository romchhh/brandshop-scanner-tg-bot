from aiogram.fsm.state import State, StatesGroup


class InventoryStates(StatesGroup):
    waiting_file = State()
    waiting_single_art = State()
    waiting_category = State()
