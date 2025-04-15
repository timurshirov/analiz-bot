import pandas as pd
import logging
import os
from aiogram import Bot, Dispatcher, types, F
from aiogram.utils.keyboard import ReplyKeyboardBuilder
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
import asyncio

TOKEN = "8082358030:AAE9YXX-zzcC9fvwjme12HJsVMxghKDX2zY"
bot = Bot(TOKEN)
dp = Dispatcher()

EXCEL_PATH = r"C:\\Users\\timur\\OneDrive\\Рабочий стол\\1122233.xlsx"
def load_data():
    if os.path.exists(EXCEL_PATH):
        return pd.read_excel(EXCEL_PATH)
    else:
        return pd.DataFrame()

df = load_data()
categories = df['Категория'].dropna().unique().tolist()
user_data = {}

@dp.message(F.text == "/start")
async def welcome(message: types.Message):
    kb = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="🚀 Начать"), KeyboardButton(text="ℹ️ Помощь")]],
        resize_keyboard=True
    )
    await message.answer(
        "👋 Добро пожаловать!\n\nНажмите кнопку, чтобы начать или получить помощь.",
        reply_markup=kb
    )

@dp.message(F.text == "ℹ️ Помощь")
async def show_help(message: types.Message):
    help_text = (
        "ℹ️ *Как пользоваться ботом:*\n\n"
        "1. Нажмите кнопку 🚀 Начать\n"
        "2. Введите возраст и выберите пол\n"
        "3. Выберите категорию анализов\n"
        "4. Отметьте нужные анализы и нажмите ✅ Готов\n"
        "5. Введите значения анализов\n"
        "6. Получите интерпретацию с нормами, причинами и примечаниями\n\n"
        "⚠️ Бот не заменяет консультацию врача."
    )
    await message.answer(help_text, parse_mode="Markdown")

@dp.message(F.text == "🚀 Начать")
async def start_analysis(message: types.Message):
    user_data[message.from_user.id] = {}
    user_data[message.from_user.id]["step"] = "awaiting_age"
    await message.answer("Введите возраст (только число):", reply_markup=types.ReplyKeyboardRemove())

@dp.message()
async def router(message: types.Message):
    global df
    df = load_data()

    user_id = message.from_user.id
    data = user_data.get(user_id, {})
    step = data.get("step")

    if step == "awaiting_age":
        if not message.text.isdigit():
            await message.answer("Пожалуйста, введите возраст числом.")
            return
        age = int(message.text)
        if age < 18:
            await message.answer("⚠️ Бот не предназначен для интерпретации анализов у детей (до 18 лет). Рекомендуем обратиться к педиатру.")
            return
        data["age"] = age
        data["step"] = "awaiting_sex"
        kb = ReplyKeyboardMarkup(
            keyboard=[[KeyboardButton(text="Мужской"), KeyboardButton(text="Женский")]],
            resize_keyboard=True
        )
        await message.answer("Выберите пол:", reply_markup=kb)
        return

    if step == "awaiting_sex":
        if message.text not in ["Мужской", "Женский"]:
            await message.answer("Пожалуйста, выберите пол кнопкой.")
            return
        data["sex"] = message.text
        data["step"] = "choosing_category"
        kb = ReplyKeyboardBuilder()
        for cat in categories:
            kb.add(KeyboardButton(text=cat))
        kb.adjust(2)
        await message.answer("Выберите категорию анализов:", reply_markup=kb.as_markup(resize_keyboard=True))
        return

    if step == "choosing_category":
        if message.text not in categories:
            await message.answer("Пожалуйста, выберите категорию из списка.")
            return
        data["category"] = message.text
        data["step"] = "choosing_analyses"
        data["analyses"] = []
        data["values"] = {}

        analyses = df[df['Категория'] == message.text]['Название анализа'].tolist()
        kb = ReplyKeyboardBuilder()
        kb.add(KeyboardButton(text="✅ Готов"), KeyboardButton(text="🔙 Назад к категории"))
        for name in analyses:
            if name != "Общий анализ крови (ОАК)":
                kb.add(KeyboardButton(text=name))
        kb.adjust(2)
        await message.answer("Выберите нужные анализы из этой категории (по одному, затем нажмите «✅ Готов»):",
                             reply_markup=kb.as_markup(resize_keyboard=True))
        return

    if step == "choosing_analyses":
        category = data.get("category")
        valid_analyses = df[df['Категория'] == category]['Название анализа'].tolist()

        if message.text == "🔙 Назад к категории":
            data["step"] = "choosing_category"
            kb = ReplyKeyboardBuilder()
            for cat in categories:
                kb.add(KeyboardButton(text=cat))
            kb.adjust(2)
            await message.answer("Выберите категорию анализов:", reply_markup=kb.as_markup(resize_keyboard=True))
            return

        if message.text == "✅ Готов":
            if not data["analyses"]:
                await message.answer("Вы ещё не выбрали ни одного анализа.")
                return
            data["step"] = "entering_values"
            data["current_input"] = data["analyses"][0]
            analysis = data["current_input"]
            row = df[(df['Категория'] == category) & (df['Название анализа'] == analysis)].iloc[0]
            norm = row['Норма М (от-до)'] if data["sex"] == "Мужской" else row['Норма Ж (от-до)']
            await message.answer(f"Введите значение для анализа «{analysis}» (пример: {norm}):", reply_markup=types.ReplyKeyboardRemove())
            return

        if message.text in valid_analyses and message.text not in data["analyses"]:
            data["analyses"].append(message.text)
            await message.answer(f"➕ Добавлен: {message.text}")
        elif message.text in data["analyses"]:
            await message.answer("Этот анализ уже выбран.")
        else:
            await message.answer("Пожалуйста, выберите анализ из списка или нажмите «✅ Готов».")
        return

    if step == "entering_values":
        try:
            val = float(message.text.replace(",", "."))
        except:
            await message.answer("Введите корректное числовое значение, например: 123.4")
            return

        current = data["current_input"]
        data["values"][current] = val
        remaining = [a for a in data["analyses"] if a not in data["values"]]
        if remaining:
            data["current_input"] = remaining[0]
            analysis = remaining[0]
            category = data["category"]
            row = df[(df['Категория'] == category) & (df['Название анализа'] == analysis)].iloc[0]
            norm = row['Норма М (от-до)'] if data["sex"] == "Мужской" else row['Норма Ж (от-до)']
            await message.answer(f"Введите значение для анализа «{analysis}» (пример: {norm}):")
        else:
            await send_summary(message, data)
            kb = ReplyKeyboardMarkup(
                keyboard=[[KeyboardButton(text="🔁 Новое исследование")]],
                resize_keyboard=True
            )
            await message.answer("Готово! Нажмите «🔁 Новое исследование», чтобы начать заново.", reply_markup=kb)
            data["step"] = "finished"
        return

    if step == "finished":
        if message.text == "🔁 Новое исследование":
            user_data.pop(user_id)
            await welcome(message)
        else:
            await message.answer("Нажмите «🔁 Новое исследование», чтобы начать заново.")
        return

    await message.answer("Пожалуйста, начните с команды /start.")

async def send_summary(message: types.Message, data):
    sex = data["sex"]
    category = data["category"]
    reply = "📊 *Результаты интерпретации:*\n\n"

    for analysis, value in data["values"].items():
        row = df[(df["Категория"] == category) & (df['Название анализа'] == analysis)]
        if row.empty:
            result = "⛔ Данные не найдены"
        else:
            row = row.iloc[0]
            norm_str = str(row['Норма М (от-до)'] if sex == "Мужской" else row['Норма Ж (от-до)'])
            try:
                min_val, max_val = map(float, norm_str.split("-"))
                if value < min_val:
                    status = "⬇️ Понижено"
                elif value > max_val:
                    status = "⬆️ Повышено"
                else:
                    status = "✅ В пределах нормы"
                result = f"{status} (норма: {norm_str})"
            except:
                result = f"⚠️ Не удалось обработать норму: {norm_str}"

            reasons_up = str(row.get("Повышено (причины)", "")).strip()
            reasons_down = str(row.get("Понижено (причины)", "")).strip()
            notes = str(row.get("Примечания", "")).strip()

            if status.startswith("⬆️") and reasons_up:
                result += f"\n🧾 Причины: {reasons_up}"
            if status.startswith("⬇️") and reasons_down:
                result += f"\n🧾 Причины: {reasons_down}"
            if notes and notes.lower() != "nan":
                result += f"\n📌 Примечание: {notes}"

        reply += f"• {analysis}: {value} → {result}\n\n"

    await message.answer(reply, parse_mode="Markdown")

async def main():
    logging.basicConfig(level=logging.INFO)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
