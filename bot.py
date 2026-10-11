import os
import re
import pandas as pd
from PIL import Image
import pytesseract
from datetime import datetime, time
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    filters, ContextTypes, ConversationHandler
)

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
# Un-comment the line below if testing locally on Windows:
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

BOT_TOKEN = os.environ.get("8927787436:AAH5Bq7JoziPGtlXH3ii9ZOP7aOsOTSZleI")
EXCEL_FILE_PATH = "dispatch_data.xlsx"

# Conversation States
WAITING_FOR_EXCEL, WAITING_FOR_SHIFT, WAITING_FOR_PHOTO = range(3)

# -----------------------------------------------------------------------------
# EXCEL & OCR HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def parse_excel_file(file_path):
    """Reads Excel file and creates a Last Out Time dictionary based on Order ID."""
    df = pd.read_excel(file_path)
    df.columns = df.columns.astype(str).str.strip()
    
    # Flexible Header Search for Order ID
    order_col = None
    for col in df.columns:
        if col.lower() in ['ship order no', 'order id', 'shipping order', 'ship_order_no', 'order_id']:
            order_col = col
            break
    if not order_col:
        order_col = df.columns[0]
        
    # Flexible Header Search for Last Out Time
    time_col = None
    for col in df.columns:
        if col.lower() in ['last out time', 'req date', 'req time', 'out time', 'request time', 'last_out_time']:
            time_col = col
            break
    if not time_col:
        time_col = df.columns[-1]

    excel_records = {}
    for _, row in df.iterrows():
        order_id = str(row[order_col]).strip()
        try:
            parsed_time = pd.to_datetime(str(row[time_col]), dayfirst=True).to_pydatetime()
            excel_records[order_id.lower()] = parsed_time
        except Exception:
            continue
            
    return excel_records

def extract_photo_details(image_path):
    """Extracts Order ID, Req Date, and Req Time from photo, supporting flexible table layouts."""
    img = Image.open(image_path).convert('L')
    text = pytesseract.image_to_string(img)
    
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    extracted_items = []

    for i, line in enumerate(lines):
        order_match = re.search(r'\b[A-Za-z0-9_-]{5,}\b', line)
        if not order_match:
            continue

        photo_order_id = order_match.group(0).strip()
        if photo_order_id.lower() in ['order', 'status', 'vehicle', 'dock', 'shipping']:
            continue

        date_match = re.search(r'\d{2}-\d{2}-\d{4}|\d{4}-\d{2}-\d{2}', line)
        time_match = re.search(r'\d{2}[:\.]\d{2}[:\.]\d{2}', line)

        if not (date_match and time_match):
            search_window = lines[max(0, i-2):min(len(lines), i+3)]
            combined_window = " ".join(search_window)
            if not date_match:
                date_match = re.search(r'\d{2}-\d{2}-\d{4}|\d{4}-\d{2}-\d{2}', combined_window)
            if not time_match:
                time_match = re.search(r'\d{2}[:\.]\d{2}[:\.]\d{2}', combined_window)

        if date_match and time_match:
            time_str_clean = time_match.group(0).replace('.', ':')
            datetime_str = f"{date_match.group(0)} {time_str_clean}"
            
            try:
                photo_time = datetime.strptime(datetime_str, "%d-%m-%Y %H:%M:%S")
            except ValueError:
                try:
                    photo_time = datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    continue

            if not any(item['order_id'] == photo_order_id for item in extracted_items):
                extracted_items.append({
                    'order_id': photo_order_id,
                    'req_time': photo_time
                })

    return extracted_items

def filter_by_shift(items, selected_shift):
    """Filters orders belonging only to the selected shift."""
    filtered = []
    for item in items:
        t = item['req_time'].time()
        
        # A Shift: 07:00 AM - 03:00 PM (15:00)
        if selected_shift == 'A' and time(7, 0) <= t < time(15, 0):
            filtered.append(item)
        # B Shift: 03:00 PM (15:00) - 11:00 PM (23:00)
        elif selected_shift == 'B' and time(15, 0) <= t < time(23, 0):
            filtered.append(item)
        # C Shift: 11:00 PM (23:00) - 07:00 AM (Next Day)
        elif selected_shift == 'C' and (t >= time(23, 0) or t < time(7, 0)):
            filtered.append(item)
            
    return filtered

# -----------------------------------------------------------------------------
# TELEGRAM BOT HANDLERS
# -----------------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "👋 Hello! Welcome to Shipping Delay Tracking Bot.\n\n"
        "To get started, please upload your Excel file (.xlsx) here."
    )
    return WAITING_FOR_EXCEL

async def restart(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Resets session data and asks for a fresh Excel upload."""
    context.user_data.clear()
    await update.message.reply_text(
        "🔄 Session reset successfully!\n\n"
        "Please send a new Excel file (.xlsx) to start over."
    )
    return WAITING_FOR_EXCEL

async def handle_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.file_name.endswith(('.xlsx', '.xls')):
        file = await doc.get_file()
        await file.download_to_drive(EXCEL_FILE_PATH)
        
        # Shift selection buttons
        keyboard = [
            [InlineKeyboardButton("A Shift (07:00 AM - 03:00 PM)", callback_data='A')],
            [InlineKeyboardButton("B Shift (03:00 PM - 11:00 PM)", callback_data='B')],
            [InlineKeyboardButton("C Shift (11:00 PM - 07:00 AM)", callback_data='C')]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        
        await update.message.reply_text("✅ Excel file saved!\n\nPlease select the Shift:", reply_markup=reply_markup)
        return WAITING_FOR_SHIFT
    else:
        await update.message.reply_text("⚠️ Please upload an Excel file (.xlsx) only.")
        return WAITING_FOR_EXCEL

async def handle_shift_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    selected_shift = query.data
    context.user_data['selected_shift'] = selected_shift
    
    await query.edit_message_text(
        f"✅ **Shift {selected_shift}** selected.\n\n"
        "Now send the HMI screen photo for Dock 1."
    )
    return WAITING_FOR_PHOTO

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not os.path.exists(EXCEL_FILE_PATH):
        await update.message.reply_text("⚠️ Excel file missing! Please type /start or /restart to upload Excel file first.")
        return WAITING_FOR_EXCEL

    photo_file = await update.message.photo[-1].get_file()
    image_path = "temp_dock_image.jpg"
    await photo_file.download_to_drive(image_path)
    
    await update.message.reply_text("🔎 Inspecting photo and calculating delay...")
    
    all_photo_items = extract_photo_details(image_path)
    
    if not all_photo_items:
        await update.message.reply_text(
            "⚠️ Could not detect orders from photo.\n\n"
            "➡️ Please send the photo for the next dock."
        )
        return WAITING_FOR_PHOTO

    selected_shift = context.user_data.get('selected_shift', 'C')
    excel_records = parse_excel_file(EXCEL_FILE_PATH)
    
    report_lines = []
    skipped_orders = []

    # -------------------------------------------------------------------------
    # CORE DELAY LOGIC:
    # From all orders in photo, if an order belongs to selected shift,
    # locate its preceding order and calculate:
    # Prev Order Excel Last Out Time - Current Order Photo Req Time = Delay
    # -------------------------------------------------------------------------
    for i, item in enumerate(all_photo_items):
        req_t = item['req_time'].time()
        
        # Check if this item belongs to current shift
        is_in_shift = False
        if selected_shift == 'A' and time(7, 0) <= req_t < time(15, 0):
            is_in_shift = True
        elif selected_shift == 'B' and time(15, 0) <= req_t < time(23, 0):
            is_in_shift = True
        elif selected_shift == 'C' and (req_t >= time(23, 0) or req_t < time(7, 0)):
            is_in_shift = True

        if is_in_shift:
            current_order_id = item['order_id']
            current_req_time = item['req_time']

            # Locate previous order item
            if i + 1 < len(all_photo_items):
                prev_order_item = all_photo_items[i + 1]
                prev_order_id = prev_order_item['order_id']
                
                # Check Last Out Time in Excel for previous order
                prev_out_time = excel_records.get(prev_order_id.lower())

                if prev_out_time:
                    time_diff = current_req_time - prev_out_time
                    delay_minutes = int(time_diff.total_seconds() / 60)

                    report_lines.append(
                        f"🔹 **Order ID**: {current_order_id}\n"
                        f"   ├ Previous Order ({prev_order_id}) Excel Out Time: {prev_out_time.strftime('%d-%m-%Y %H:%M:%S')}\n"
                        f"   ├ Current Req Time (Photo): {current_req_time.strftime('%d-%m-%Y %H:%M:%S')}\n"
                        f"   └ ⚠️ **Delay**: **{delay_minutes} Minutes**"
                    )
                else:
                    skipped_orders.append(f"{current_order_id} (Prev order {prev_order_id} not in Excel)")
            else:
                skipped_orders.append(f"{current_order_id} (Previous order not available in photo)")

    # -------------------------------------------------------------------------
    # REPORT GENERATION & SEQUENTIAL REQUEST
    # -------------------------------------------------------------------------
    shift_name = f"Shift {selected_shift}"
    summary_msg = f"📊 **{shift_name} Delay Summary Report**\n\n"

    if report_lines:
        summary_msg += "\n\n".join(report_lines)
    else:
        summary_msg += "No eligible orders found for delay calculation in this shift."

    if skipped_orders:
        summary_msg += f"\n\n⚠️ **Skipped**: {', '.join(skipped_orders)}"

    # 1. Send report
    await update.message.reply_text(summary_msg, parse_mode='Markdown')

    # 2. Prompt for next photo
    await update.message.reply_text("➡️ **Send Next Dock Photo (or type /restart to reset).**")

    # Clean up local image
    if os.path.exists(image_path):
        os.remove(image_path)

    return WAITING_FOR_PHOTO

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Processing stopped. Send /start to begin again.")
    return ConversationHandler.END

# -----------------------------------------------------------------------------
# MAIN APPLICATION
# -----------------------------------------------------------------------------
if __name__ == '__main__':
    app = Application.builder().token("8927787436:AAH5Bq7JoziPGtlXH3ii9ZOP7aOsOTSZleI").build()

    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler('start', start),
            CommandHandler('restart', restart)
        ],
        states={
            WAITING_FOR_EXCEL: [MessageHandler(filters.Document.ALL, handle_excel)],
            WAITING_FOR_SHIFT: [CallbackQueryHandler(handle_shift_selection)],
            WAITING_FOR_PHOTO: [MessageHandler(filters.PHOTO, handle_photo)],
        },
        fallbacks=[
            CommandHandler('restart', restart),
            CommandHandler('cancel', cancel)
        ],
    )

    app.add_handler(conv_handler)
    app.add_handler(CommandHandler('restart', restart))

    print("Bot is running successfully...")
    app.run_polling()


