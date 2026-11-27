import asyncio
import csv
import io
import json
import logging
import os
from aiogram import Bot, Dispatcher, F, types
from aiogram.enums import ChatAction
from aiogram.filters import Command
from aiogram.types import FSInputFile
from gigachat import GigaChat
from gigachat.models import Chat, Function, FunctionParameters, Messages, MessagesRole
import openpyxl

TELEGRAM_TOKEN = "9000252271:AAG7SV3Z2KzBrA4KTlLTO_KsiZvCJfyztxg"
GIGACHAT_CREDENTIALS = "MDFhMjE5OTEtNzFmYi03YjdjLWE4MDUtZGM4NmZjNmE5YWU5OmVhMDM0NWU1LWE3ZjMtNDdmMS05ZmY5LWQ2OGI4ZmVkYzBiMA=="

SYSTEM_PROMPT = (
    "Ты умный ИИ-ассистент в Telegram на базе российской нейросети GigaChat. "
    "Отвечай на русском языке четко и по делу. Ты умеешь читать таблицы Excel "
    "и создавать новые Excel-файлы через функцию create_excel."
)

CREATE_EXCEL_FUNC = Function(
    name="create_excel",
    description="Создает Excel файл (.xlsx) с таблицей данных и отправляет его пользователю. Вызывай эту функцию, когда пользователь просит сделать, создать или скачать таблицу/файл Excel.",
    parameters=FunctionParameters(
        type="object",
        properties={
            "filename": {
                "type": "string",
                "description": "Имя файла латиницей с расширением .xlsx (например, budget.xlsx)",
            },
            "columns": {
                "type": "array",
                "items": {"type": "string"},
                "description": "Список названий столбцов таблицы",
            },
            "rows": {
                "type": "array",
                "items": {"type": "array", "items": {"type": "string"}},
                "description": "Список строк таблицы. Количество элементов в каждой строке должно совпадать с количеством столбцов.",
            },
        },
        required=["filename", "columns", "rows"],
    ),
)

user_histories = {}
bot = Bot(token=TELEGRAM_TOKEN)
dp = Dispatcher()


def get_history(user_id: int) -> list:
    if user_id not in user_histories:
        user_histories[user_id] = [
            Messages(role=MessagesRole.SYSTEM, content=SYSTEM_PROMPT)
        ]
    return user_histories[user_id]


@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    user_histories[message.from_user.id] = [
        Messages(role=MessagesRole.SYSTEM, content=SYSTEM_PROMPT)
    ]
    await message.answer(
        "Привет! Я твой ИИ-ассистент на базе российского GigaChat. Вот что я умею:\n"
        "1. Отвечать на любые вопросы и помогать с текстами.\n"
        "2. Читать Excel-файлы — отправь мне .xlsx файл и напиши, что в нём нужно найти или посчитать.\n"
        "3. Создавать Excel-таблицы — попроси сделать таблицу, и я пришлю готовый .xlsx файл.\n\n"
        "Команда /clear очистит память диалога."
    )


@dp.message(Command("clear"))
async def cmd_clear(message: types.Message):
    user_histories[message.from_user.id] = [
        Messages(role=MessagesRole.SYSTEM, content=SYSTEM_PROMPT)
    ]
    await message.answer("Память очищена! Начнём с чистого листа.")


async def process_gigachat_request(
    user_id: int, chat_id: int, user_text: str
):
    history = get_history(user_id)
    history.append(Messages(role=MessagesRole.USER, content=user_text))

    if len(history) > 16:
        history = [history[0]] + history[-15:]
        user_histories[user_id] = history

    await bot.send_chat_action(chat_id=chat_id, action=ChatAction.TYPING)

    try:
        async with GigaChat(
            credentials=GIGACHAT_CREDENTIALS,
            verify_ssl_certs=False,
            model="GigaChat",
        ) as giga:
            payload = Chat(
                messages=history,
                functions=[CREATE_EXCEL_FUNC],
                function_call="auto",
                temperature=0.5,
            )
            response = await giga.achat(payload)

        msg = response.choices[0].message

        if msg.function_call and msg.function_call.name == "create_excel":
            args = msg.function_call.arguments
            if isinstance(args, str):
                args = json.loads(args)

            filename = args.get("filename", "table.xlsx")
            if not filename.endswith(".xlsx"):
                filename += ".xlsx"

            columns = args.get("columns", [])
            rows = args.get("rows", [])

            wb = openpyxl.Workbook()
            ws = wb.active
            if columns:
                ws.append(columns)
            for row in rows:
                ws.append(row)

            temp_path = f"temp_{user_id}_{filename}"
            wb.save(temp_path)

            doc = FSInputFile(temp_path, filename=filename)
            await bot.send_document(
                chat_id=chat_id,
                document=doc,
                caption=f"Готово! Вот ваша таблица {filename}.",
            )

            if os.path.exists(temp_path):
                os.remove(temp_path)

            history.append(
                Messages(
                    role=MessagesRole.ASSISTANT,
                    content=f"Я создал и отправил файл {filename}.",
                )
            )
            return

        reply_text = msg.content or "Готово!"
        history.append(
            Messages(role=MessagesRole.ASSISTANT, content=reply_text)
        )
        await bot.send_message(chat_id=chat_id, text=reply_text)

    except Exception as e:
        await bot.send_message(
            chat_id=chat_id, text=f"Ошибка при обращении к GigaChat: {e}"
        )


@dp.message(F.document)
async def handle_document(message: types.Message):
    doc = message.document
    if not doc.file_name.endswith(".xlsx"):
        await message.answer(
            "Пока я умею читать только таблицы Excel в формате .xlsx"
        )
        return

    user_id = message.from_user.id
    await bot.send_chat_action(
        chat_id=message.chat.id, action=ChatAction.TYPING
    )

    doc_path = f"doc_{user_id}_{doc.file_name}"
    try:
        file_info = await bot.get_file(doc.file_id)
        await bot.download_file(file_info.file_path, doc_path)

        wb = openpyxl.load_workbook(doc_path, data_only=True)
        ws = wb.active

        output = io.StringIO()
        writer = csv.writer(output)
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= 100:
                break
            writer.writerow([cell if cell is not None else "" for cell in row])

        table_preview = output.getvalue()
        user_caption = (
            message.caption
            or "Проанализируй эту таблицу и кратко расскажи, что в ней находится."
        )

        prompt_with_table = (
            f"Пользователь прислал Excel-файл '{doc.file_name}'. Вот его содержимое (в формате CSV):\n"
            f"```csv\n{table_preview}\n```\n"
            f"Запрос пользователя: {user_caption}"
        )

        await process_gigachat_request(
            user_id, message.chat.id, prompt_with_table
        )

    except Exception as e:
        await message.answer(f"Ошибка при чтении файла Excel: {e}")
    finally:
        if os.path.exists(doc_path):
            os.remove(doc_path)


@dp.message(F.text)
async def handle_text(message: types.Message):
    await process_gigachat_request(
        message.from_user.id, message.chat.id, message.text
    )


async def main():
    logging.basicConfig(level=logging.INFO)
    print("Бот на базе GigaChat успешно запущен!")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
