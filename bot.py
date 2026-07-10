# -*- coding: utf-8 -*-
"""
Telegram-bot: PDF invoice to Excel act
"""

import os
import sys
import json
import base64
import logging
from datetime import datetime
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

import anthropic
from excel_generator import generate_act

# Force UTF-8 encoding
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

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
        "You are an expert at reading Russian invoice documents.\n"
        "Extract data and return ONLY valid JSON:\n"
        '{"doc_number":"","doc_date":"DD.MM.YYYY","supplier":"","recipient":"",'
        '"items":[{"name":"","article":"","quantity":0,"unit":""}]}\n'
        "Return only JSON, no markdown."
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
        "Bot zapuschen!\n\nOtprav PDF-skan nakladnoj i poluchi akt Excel."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document

    if doc.mime_type != "application/pdf" and not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("Otprav fajl v formate PDF.")
        return

    msg = await update.message.reply_text("Chitau nakladnuyu...")

    try:
        file = await context.bot.get_file(doc.file_id)
        pdf_bytes = await file.download_as_bytearray()

        await msg.edit_text("Raspoznayu...")

        data = extract_invoice_data(bytes(pdf_bytes))
        logger.info("Extracted: %s", str(data)[:200])

        await msg.edit_text("Formiruyu Excel...")

        output_path = f"act_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        generate_act(data, TEMPLATE_PATH, output_path)

        num = str(data.get("doc_number", "-"))
        dt = str(data.get("doc_date", "-"))
        sup = str(data.get("supplier", "-"))
        cnt = str(len(data.get("items", [])))
        caption = f"Akt gotov. No{num} ot {dt}. Postavshchik: {sup}. Pozicij: {cnt}"

        with open(output_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"Akt_{num}.xlsx",
                caption=caption,
            )

        await msg.delete()
        os.remove(output_path)

    except json.JSONDecodeError:
        await msg.edit_text("Ne udalos raspoznat nakladnuyu.")
        logger.exception("JSON parse error")
    except Exception as e:
        await msg.edit_text(f"Oshibka: {type(e).__name__}")
        logger.exception("Error: %s", e)


async def handle_other(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Otprav PDF nakladnuyu.")


def main():
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_other))
    logger.info("Bot started.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
