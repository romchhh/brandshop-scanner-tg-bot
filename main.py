import asyncio
import logging
from aiogram import Bot, Dispatcher
from aiogram.fsm.storage.memory import MemoryStorage
from config import token


logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)


bot = Bot(token=token)
storage = MemoryStorage()
dp = Dispatcher(bot=bot, storage=storage)


async def on_startup():
    """Функція, яка виконується при запуску бота"""
    logging.info("Бот запущено!")


async def on_shutdown():
    """Функція, яка виконується при зупинці бота"""
    logging.info("Бот зупинено!")


async def main():
    from handlers.inventory_handlers import router as inventory_router
    
    dp.include_router(inventory_router)
    
    dp.startup.register(on_startup)
    dp.shutdown.register(on_shutdown)

    while True:
        try:
            await dp.start_polling(bot, skip_updates=True)
        except Exception as e:
            logging.error(f"Bot crashed with error: {e}", exc_info=True)
            await asyncio.sleep(5)


if __name__ == '__main__':
    import time
    while True:
        try:
            asyncio.run(main())
        except Exception as e:
            logging.error(f"Critical error: {e}", exc_info=True)
            time.sleep(5)