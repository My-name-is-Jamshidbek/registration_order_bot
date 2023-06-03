from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Command
from aiogram.dispatcher.filters.state import State, StatesGroup
import sqlite3
import pandas as pd

def export_to_excel():
    # Ma'lumotlar bazasiga ulanish
    conn = sqlite3.connect('malumotlar.db')

    # Ma'lumotlarni olish
    query = "SELECT * FROM malumotlar"
    df = pd.read_sql_query(query, conn)

    # Excel fayliga saqlash
    df.to_excel('malumotlar.xlsx', index=False)

    # Ulanishni yopish
    conn.close()


# Bot tokeni
token = '5346182572:AAHc-WGROVX-bzCkF9b2Qw8v9BgABcUVMxE'

# Ma'lumotlar bazasiga ulanish
conn = sqlite3.connect('malumotlar.db')
cursor = conn.cursor()

# Telegram botni yaratish
bot = Bot(token=token)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)

# Statelar uchun FSM
class states(StatesGroup):
    ISM = State()
    FAMILIYA = State()
    UY_MANZILI = State()
    TELEFON = State()
    QIZIQISH = State()

# Start komandasi uchun handler
@dp.message_handler(Command('start'))
async def start(message: types.Message, state: FSMContext):
    await message.reply('Assalomu alaykum! Iltimos, ismni yuboring.')
    await states.ISM.set()

# Malumotlarni saqlash uchun handlerlar
@dp.message_handler(state=states.ISM)
async def process_ism(message: types.Message, state: FSMContext):
    await state.update_data(ism=message.text)
    await message.reply('Familiyangizni yuboring.')
    await states.FAMILIYA.set()

@dp.message_handler(state=states.FAMILIYA)
async def process_familiya(message: types.Message, state: FSMContext):
    await state.update_data(familiya=message.text)
    await message.reply('Qiziqishingizni yuboring.')
    await states.QIZIQISH.set()

@dp.message_handler(state=states.QIZIQISH)
async def process_qiziqish(message: types.Message, state: FSMContext):
    await state.update_data(qiziqish=message.text)
    await message.reply('Uy manzilingizni yuboring.')
    await states.UY_MANZILI.set()

@dp.message_handler(state=states.UY_MANZILI)
async def process_uy_manzili(message: types.Message, state: FSMContext):
    await state.update_data(uy_manzili=message.text)
    await message.reply('Telefon raqamingizni yuboring.')
    await states.TELEFON.set()

@dp.message_handler(state=states.TELEFON)
async def process_telefon(message: types.Message, state: FSMContext):
    data = await state.get_data()
    ism = data['ism']
    familiya = data['familiya']
    uy_manzili = data['uy_manzili']
    qiziqish = data['qiziqish']
    telefon_raqami = message.text

    # Malumotlarni SQLite bazasiga saqlash
    cursor.execute("INSERT INTO malumotlar (ism, familiya, uy_manzili, telefon_raqami, qiziqish) VALUES (?, ?, ?, ?, ?)",
                   (ism, familiya, uy_manzili, telefon_raqami, qiziqish))
    conn.commit()

    await message.reply('Malumotlar saqlandi!')
    await state.finish()

# Admin bo'limi uchun handler
@dp.message_handler(Command('admin'))
async def admin_panel(message: types.Message):
    # Admin ID-sini olish
    admin_id = 2081653869  # Adminning Telegram ID-sini kiritng

    # Foydalanuvchi ID-sini tekshirish
    if message.from_user.id == admin_id:
        # Malumotlarni Excelga yuklash tugmasi
        keyboard = InlineKeyboardMarkup()
        btn_export = InlineKeyboardButton('Malumotlarni Yuklash', callback_data='export')
        keyboard.row(btn_export)

        await message.reply('Admin paneliga xush kelibsiz!', reply_markup=keyboard)
    else:
        await message.reply('Siz admin emassiz!')

# Export tugmasi uchun callback handler
@dp.callback_query_handler(lambda callback_query: callback_query.data == 'export')
async def export_callback_handler(callback_query: types.CallbackQuery):
    # Excelga yuklash
    export_to_excel()

    # Excel faylini foydalanuvchiga yuborish
    with open('malumotlar.xlsx', 'rb') as file:
        await bot.send_document(callback_query.from_user.id, file)

# Malumotlarni SQLite bazasiga yozish
def create_database():
    # Ma'lumotlar bazasini yaratish
    cursor.execute('''CREATE TABLE IF NOT EXISTS malumotlar (
                      id INTEGER PRIMARY KEY AUTOINCREMENT,
                      ism TEXT,
                      familiya TEXT,
                      qiziqish TEXT,
                      uy_manzili TEXT,
                      telefon_raqami TEXT)''')

    # Indekslar uchun jadvallar
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_ism ON malumotlar (ism)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_familiya ON malumotlar (familiya)")

    # Ma'lumotlar bazasini saqlash
    conn.commit()

# Malumotlar bazasini yaratishni tekshirish va ishga tushirish
create_database()

# Telegram botni ishga tushirish
if __name__ == '__main__':
    from aiogram import executor
    executor.start_polling(dp)
