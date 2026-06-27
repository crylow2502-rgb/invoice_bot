RAM
Telegram-бот: PDF накладная - Акт приёма-передачи (Excel)
Требования: pip install python-telegram-bot anthropic openpyxl
"""

import os
import json
import base64
import logging
from datetime import datetime
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

import anthropic
from excel_generator import generate_act

# ── Настройки ──────────────────────────────────────────────────────────────
 _TOKEN = "ВАШ_TELEGRAM_BOT_TOKEN"   # от @BotFather
ANTHROPIC_API_KEY = "ВАШ_ANTHROPIC_API_KEY"  # с console.anthropic.com
TEMPLATE_PATH = "template.xlsm"              # шаблон акта рядом с bot.py

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


# ── Извлечение данных из скана через Claude Vision ─────────────────────────
def extract_invoice_data(pdf_bytes: bytes) -> dict:
    """Отправляет PDF-скан в Claude и получает структурированные данные."""

    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    prompt = """Ты - эксперт по распознаванию документов. Перед тобой скан накладной (товарная накладная / УПД).

Извлеки следующие данные и верни ТОЛЬКО валидный JSON без пояснений:

{
  "doc_number": "номер документа",
  "doc_date": "дата документа в формате ДД.ММ.ГГГГ",
  "supplier": "наименование поставщика",
  "recipient": "наименование получателя",
  "items": [
    {
      "name": "наименование материала/товара",
      "article": "артикул или код (если есть, иначе пустая строка)",
      "quantity": число (только число, без единиц),
      "unit": "единица измерения (шт, м, кг и т.д.)"
    }
  ]
}

Если какое-то поле не удаётся прочитать — оставь пустую строку или 0.
Верни только JSON, без markdown, без пояснений."""

    response = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=2000,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_b64,
                        },
                    },
                    {"type": "text", "text": prompt},
                ],
            }
        ],
    )

    raw = response.content[0].text.strip()
    # Убираем возможные markdown-обёртки
    if raw.startswith("```"):
        raw = raw.split("```")[1]
TEMPLATE_PATH = "template.xlsm"              # шаблон акта рядом с bot.py

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


# ── Извлечение данных из скана через Claude Vision ─────────────────────────
def extract_invoice_data(pdf_bytes: bytes) -> dict:
    """Отправляет PDF-скан в Claude и получает структурированные данные."""

    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    prompt = """Ты — эксперт по распознаванию документов. Перед тобой скан накладной (товарная накладная / УПД).

Извлеки следующие данные и верни ТОЛЬКО валидный JSON без пояснений:

{
  "doc_number": "номер документа",
  "doc_date": "дата документа в формате ДД.ММ.ГГГГ",
  "supplier": "наименование поставщика",
  "recipient": "наименование получателя",
  "items": [
    {
      "name": "наименование материала/товара",
      "article": "артикул или код (если есть, иначе пустая строка)",
      "quantity": число (только число, без единиц),
      "unit": "единица измерения (шт, м, кг и т.д.)"
    }
  ]
}

Если какое-то поле не удаётся прочитать — оставь пустую строку или 0.
Верни только JSON, без markdown, без пояснений."""

    response = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=2000,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_b64,
                        },
                    },
                    {"type": "text", "text": prompt},
                ],
            }
        ],
    )

    raw = response.content[0].text.strip()
    # Убираем возможные markdown-обёртки
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


# ── Telegram-хэндлеры ──────────────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привет!\n\n"
        "Отправь мне PDF-скан накладной, и я сформирую акт приёма-передачи материалов в формате Excel.\n\n"
        "📎 Просто прикрепи PDF-файл к сообщению."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document

    # Проверяем, что это PDF
    if doc.mime_type != "application/pdf" and not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("⚠️ Пожалуйста, отправь файл в формате PDF.")
        return

    msg = await update.message.reply_text("⏳ Читаю накладную, подожди немного...")

    try:
        # Скачиваем PDF
        file = await context.bot.get_file(doc.file_id)
        pdf_bytes = await file.download_as_bytearray()

        await msg.edit_text("🔍 Распознаю данные через AI...")

        # Извлекаем данные через Claude
        data = extract_invoice_data(bytes(pdf_bytes))

        logger.info(f"Извлечено: {data}")

        await msg.edit_text("📊 Формирую акт Excel...")

        # Генерируем Excel
        output_path = f"act_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        generate_act(data, TEMPLATE_PATH, output_path)

        # Отправляем файл
        caption = (
            f"✅ Акт сформирован\n\n"
            f"📄 Накладная №{data.get('doc_number', '—')} "
            f"от {data.get('doc_date', '—')}\n"
            f"🏭 Поставщик: {data.get('supplier', '—')}\n"
            f"📦 Позиций: {len(data.get('items', []))}"
        )

        with open(output_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"Акт_{data.get('doc_number', 'без_номера')}.xlsx",
                caption=caption,
            )

        await msg.delete()

        # Удаляем временный файл
        os.remove(output_path)

    except json.JSONDecodeError:
        await msg.edit_text("❌ Не удалось распознать структуру накладной. Попробуй другой файл.")
        logger.exception("JSON parse error")
    except Exception as e:
        await msg.edit_text(f"❌ Ошибка: {e}")
        logger.exception("Unexpected error")


async def handle_other(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📎 Отправь PDF-файл накладной.")


# ── Запуск ─────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_other))

    logger.info("Бот запущен...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
