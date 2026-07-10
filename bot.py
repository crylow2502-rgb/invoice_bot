"""
Telegram-bot: PDF invoice to Excel act
Requirements: pip install python-telegram-bot anthropic openpyxl
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

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
TEMPLATE_PATH = "template.xlsm"

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def extract_invoice_data(pdf_bytes: bytes) -> dict:
    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    prompt = (
        "You are an expert at reading Russian invoice documents (nakladnaya/UPD). "
        "Extract the following fields and return ONLY valid JSON with no explanation:\n\n"
        "{\n"
        '  "doc_number": "document number",\n'
        '  "doc_date": "document date in format DD.MM.YYYY",\n'
        '  "supplier": "supplier company name",\n'
        '  "recipient": "recipient company name",\n'
        '  "items": [\n'
        '    {\n'
        '      "name": "material or product name",\n'
        '      "article": "article or code (empty string if not found)",\n'
        '      "quantity": 0,\n'
        '      "unit": "unit of measurement"\n'
        '    }\n'
        '  ]\n'
        "}\n\n"
        "If a field cannot be read, use empty string or 0. "
        "Return only JSON, no markdown, no explanation."
    )

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
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Privet!\n\n"
        "Otprav mne PDF-skan nakladnoj, i ya sformiruu akt priema-peredachi materialov v formate Excel.\n\n"
        "Prosto prikrepi PDF-fajl k soobshcheniyu."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document

    if doc.mime_type != "application/pdf" and not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("Pozhalujsta, otprav fajl v formate PDF.")
        return

    msg = await update.message.reply_text("Chitau nakladnuyu, podozhdite...")

    try:
        file = await context.bot.get_file(doc.file_id)
        pdf_bytes = await file.download_as_bytearray()

        await msg.edit_text("Raspoznayu dannye cherez AI...")

        data = extract_invoice_data(bytes(pdf_bytes))
        logger.info(f"Extracted: {data}")

        await msg.edit_text("Formiruyu akt Excel...")

        output_path = f"act_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        generate_act(data, TEMPLATE_PATH, output_path)

        caption = (
            f"Akt sformirovan\n\n"
            f"Nakladnaya No{data.get('doc_number', '-')} "
            f"ot {data.get('doc_date', '-')}\n"
            f"Postavshchik: {data.get('supplier', '-')}\n"
            f"Pozicij: {len(data.get('items', []))}"
        )

        with open(output_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"Akt_{data.get('doc_number', 'bez_nomera')}.xlsx",
                caption=caption,
            )

        await msg.delete()
        os.remove(output_path)

    except json.JSONDecodeError:
        await msg.edit_text("Ne udalos raspoznat strukturu nakladnoj. Poprobuj drugoj fajl.")
        logger.exception("JSON parse error")
    except Exception as e:
        await msg.edit_text(f"Oshibka: {e}")
        logger.exception("Unexpected error")


async def handle_other(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Otprav PDF-fajl nakladnoj.")


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_other))
    logger.info("Bot started...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
