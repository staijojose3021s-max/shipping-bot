import os
import re
import pandas as pd
from PIL import Image
import pytesseract
from datetime import datetime, time
from telegram import Update, ReplyKeyboardRemove
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes,
    ConversationHandler,
)

# ---------------------------------------------------------
# CONFIGURATION & STATES
# ---------------------------------------------------------
WAITING_FOR_EXCEL, WAITING_FOR_PHOTO = range(2)

# Shift timings
SHIFT_A_START = time(7, 0)
SHIFT_A_END = time(15, 0)
SHIFT_B_START = time(15, 0)
SHIFT_B_END = time(23, 0)


def get_shift_from_time(dt_time: time) -> str:
    """Determine shift (A, B, C) based on order request time."""
    if SHIFT_A_START <= dt_time < SHIFT_A_END:
        return 'A'
    elif SHIFT_B_START <= dt_time < SHIFT_B_END:
        return 'B'
    else:
        return 'C'


def parse_time(time_str: str) -> time:
    """Parse time string in HH:MM or HH:MM:SS format."""
    time_str = time_str.strip()
    for fmt in ('%H:%M:%S', '%H:%M'):
        try:
            return datetime.strptime(time_str, fmt).time()
        except ValueError:
            pass
    raise ValueError(f"Time format not supported: {time_str}")


def calculate_time_difference_minutes(t_start: time, t_end: time) -> float:
    """Calculate difference in minutes between two time objects (handles cross-midnight)."""
    dummy_date = datetime(2000, 1, 1)
    dt_start = datetime.combine(dummy_date, t_start)
    dt_end = datetime.combine(dummy_date, t_end)

    if dt_end < dt_start:
        dt_end += pd.Timedelta(days=1)

    diff = (dt_end - dt_start).total_seconds() / 60.0
    return round(diff, 2)


# ---------------------------------------------------------
# BOT HANDLERS
# ---------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Starts the conversation and requests the Excel record."""
    context.user_data.clear()
    await update.message.reply_text(
        "👋 Welcome to Shipping Delay Tracker Bot!\n\n"
        "Please upload the Excel file (.xlsx or .xls) containing previous shift records to begin."
    )
    return WAITING_FOR_EXCEL


async def restart(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Restarts the session and asks for a new Excel file."""
    context.user_data.clear()
    await update.message.reply_text(
        "🔄 Session reset successfully!\n\n"
        "Please upload a new Excel file (.xlsx or .xls) to start over."
    )
    return WAITING_FOR_EXCEL


async def handle_excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Processes the uploaded Excel document."""
    document = update.message.document

    if not (document.file_name.endswith('.xlsx') or document.file_name.endswith('.xls')):
        await update.message.reply_text("❌ Invalid file type. Please upload a valid Excel file (.xlsx or .xls).")
        return WAITING_FOR_EXCEL

    file = await context.bot.get_file(document.file_id)
    file_path = f"temp_{document.file_name}"
    await file.download_to_drive(file_path)

    try:
        # Load Excel file
        df = pd.read_excel(file_path)
        
        # Keep file in user_data for cross-shift comparisons
        context.user_data['excel_df'] = df
        context.user_data['dock_count'] = 0

        await update.message.reply_text(
            "✅ Excel file uploaded and parsed successfully!\n\n"
            "Now, please send an image of the HMI screen for Dock 1."
        )
        return WAITING_FOR_PHOTO

    except Exception as e:
        await update.message.reply_text(f"❌ Error processing Excel file: {str(e)}\nPlease try uploading again.")
        return WAITING_FOR_EXCEL
    finally:
        if os.path.exists(file_path):
            os.remove(file_path)


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Processes the uploaded HMI screen image using OCR."""
    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    photo_path = f"temp_{photo.file_id}.jpg"
    await file.download_to_drive(photo_path)

    await update.message.reply_text("🔍 Processing image with OCR...")

    try:
        image = Image.open(photo_path)
        ocr_text = pytesseract.image_to_string(image)

        # Extraction logic using Regex (Extract times like HH:MM or HH:MM:SS)
        times_found = re.findall(r'\b(?:[01]?\d|2[0-3]):[0-5]\d(?::[0-5]\d)?\b', ocr_text)

        if len(times_found) < 2:
            await update.message.reply_text(
                "⚠️ Could not clearly read Request Time and Dispatch Time from the image.\n"
                "Please upload a clearer HMI screen photo."
            )
            return WAITING_FOR_PHOTO

        req_time = parse_time(times_found[0])
        disp_time = parse_time(times_found[1])
        current_shift = get_shift_from_time(req_time)

        # Increment dock count
        context.user_data['dock_count'] = context.user_data.get('dock_count', 0) + 1
        current_dock = context.user_data['dock_count']

        # Calculate delay within current dock photo
        delay_minutes = calculate_time_difference_minutes(req_time, disp_time)

        response = (
            f"📊 **Dock {current_dock} Processing Summary**\n"
            f"----------------------------------------\n"
            f"• **Request Time:** {req_time.strftime('%H:%M:%S')}\n"
            f"• **Dispatch Time:** {disp_time.strftime('%H:%M:%S')}\n"
            f"• **Shift:** Shift {current_shift}\n"
            f"• **Dock Shipping Delay:** {delay_minutes} minutes\n"
        )

        # Cross-shift check logic if it is the first dock order
        if current_dock == 1 and 'excel_df' in context.user_data:
            df = context.user_data['excel_df']
            if not df.empty and 'Last Out Time' in df.columns:
                last_out_str = str(df.iloc[-1]['Last Out Time'])
                try:
                    last_out_time = parse_time(last_out_str)
                    cross_delay = calculate_time_difference_minutes(last_out_time, req_time)
                    response += (
                        f"\n🔄 **Cross-Shift Transition Delay:**\n"
                        f"• **Previous Shift Last Out Time:** {last_out_time.strftime('%H:%M:%S')}\n"
                        f"• **Transition Delay:** {cross_delay} minutes\n"
                    )
                except Exception:
                    pass

        if current_dock >= 6:
            response += (
                "\n🎉 **All 6 Docks Completed!**\n"
                "Type /restart or send a new Excel file to start a new batch."
            )
            await update.message.reply_text(response, parse_mode='Markdown')
            return WAITING_FOR_PHOTO
        else:
            response += f"\n📸 Please send the photo for **Dock {current_dock + 1}**."
            await update.message.reply_text(response, parse_mode='Markdown')
            return WAITING_FOR_PHOTO

    except Exception as e:
        await update.message.reply_text(f"❌ Error processing image: {str(e)}\nPlease try again.")
        return WAITING_FOR_PHOTO
    finally:
        if os.path.exists(photo_path):
            os.remove(photo_path)


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Cancels and ends the conversation."""
    context.user_data.clear()
    await update.message.reply_text(
        "❌ Operation cancelled. Send /start to begin again.",
        reply_markup=ReplyKeyboardRemove()
    )
    return ConversationHandler.END


# ---------------------------------------------------------
# MAIN FUNCTION
# ---------------------------------------------------------
def main():
    # Fetch token from Railway Environment Variables or replace with your string token
    BOT_TOKEN = os.environ.get("BOT_TOKEN", "YOUR_TELEGRAM_BOT_TOKEN_HERE")

    application = Application.builder().token("8927787436:AAHMMvz2e6qcMSklrKO7cDoSI1TaGKaPU3g").build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            WAITING_FOR_EXCEL: [
                MessageHandler(filters.Document.ALL, handle_excel)
            ],
            WAITING_FOR_PHOTO: [
                MessageHandler(filters.PHOTO, handle_photo)
            ],
        },
        fallbacks=[
            CommandHandler("restart", restart),
            CommandHandler("cancel", cancel)
        ],
    )

    application.add_handler(conv_handler)
    application.add_handler(CommandHandler("restart", restart))

    print("Bot is running...")
    application.run_polling()


if __name__ == "__main__":
    main()
