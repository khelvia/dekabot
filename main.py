import os
import google.generativeai as genai
import pandas as pd
import yfinance as yf
from docx import Document
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    ContextTypes,
    filters,
)

# Environment variables
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")

# ---------------- TEXT CHAT ----------------
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    response = model.generate_content(user_text)
    await update.message.reply_text(response.text)

# ---------------- WORD REWRITE ----------------
async def rewrite_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.document.mime_type != \
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        await update.message.reply_text("Upload a .docx file only.")
        return

    file = await update.message.document.get_file()
    input_path = "input.docx"
    output_path = "rewritten.docx"

    await file.download_to_drive(input_path)

    doc = Document(input_path)
    text = "\n".join([p.text for p in doc.paragraphs])

    response = model.generate_content(
        f"Rewrite this professionally:\n\n{text}"
    )

    new_doc = Document()
    new_doc.add_paragraph(response.text)
    new_doc.save(output_path)

    await update.message.reply_document(open(output_path, "rb"))

# ---------------- EXCEL ANALYSIS ----------------
async def handle_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = await update.message.document.get_file()
    input_path = "input.xlsx"
    output_path = "analysis.xlsx"

    await file.download_to_drive(input_path)

    df = pd.read_excel(input_path)

    summary = df.describe()

    writer = pd.ExcelWriter(output_path, engine="openpyxl")
    df.to_excel(writer, sheet_name="Original")
    summary.to_excel(writer, sheet_name="Summary")
    writer.close()

    await update.message.reply_document(open(output_path, "rb"))

# ---------------- STOCK ANALYSIS ----------------
async def stock(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if len(context.args) == 0:
        await update.message.reply_text("Use: /stock TCS.NS")
        return

    ticker = context.args[0]
    data = yf.download(ticker, period="6mo")

    if data.empty:
        await update.message.reply_text("Invalid ticker.")
        return

    returns = data["Close"].pct_change().mean()
    volatility = data["Close"].pct_change().std()

    analysis = model.generate_content(
        f"""
        Stock: {ticker}
        Avg Daily Return: {returns}
        Volatility: {volatility}

        Give interpretation and risk summary.
        """
    )

    await update.message.reply_text(analysis.text)

# ---------------- MAIN ----------------
app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()

app.add_handler(CommandHandler("stock", stock))
app.add_handler(MessageHandler(filters.Document.EXCEL, handle_excel))
app.add_handler(MessageHandler(filters.Document.ALL, rewrite_doc))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

app.run_polling()