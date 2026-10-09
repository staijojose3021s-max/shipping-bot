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
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

BOT_TOKEN = "8927787436:AAHMMvz2e6qcMSklrKO7cDoSI1TaGKaPU3g"
EXCEL_FILE_PATH = "dispatch_data.xlsx"

# Conversation States
WAITING_FOR_EXCEL, WAITING_FOR_SHIFT, WAITING_FOR_PHOTO = range(3)

# -----------------------------------------------------------------------------
# EXCEL & OCR HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def parse_excel_file(file_path):
    """Excel ഫയൽ വായിച്ച് Order ID അടിസ്ഥാനമാക്കി Last Out Time നിഘണ്ടു (Dictionary) ഉണ്ടാക്കുന്നു."""
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
    """ഫോട്ടോയിൽ നിന്ന് Order ID, Req Date, Req Time എന്നിവ എടുത്ത് ലിസ്റ്റ് ആയി നൽകുന്നു (മേൽ നിന്ന് താഴേക്ക്)."""
    img = Image.open(image_path).convert('L')
    text = pytesseract.image_to_string(img)
    
    lines = text.split('\n')
    extracted_items = []

    for line in lines:
        order_match = re.search(r'MOS\d+', line, re.IGNORECASE)
        date_match = re.search(r'\d{2}-\d{2}-\d{4}|\d{4}-\d{2}-\d{2}', line)
        time_match = re.search(r'\d{2}:\d{2}:\d{2}', line)

        if order_match and date_match and time_match:
            photo_order_id = order_match.group(0).strip()
            datetime_str = f"{date_match.group(0)} {time_match.group(0)}"
            try:
                photo_time = datetime.strptime(datetime_str, "%d-%m-%Y %H:%M:%S")
            except ValueError:
                photo_time = datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S")

            extracted_items.append({
                'order_id': photo_order_id,
                'req_time': photo_time
            })

    return extracted_items

def filter_by_shift(items, selected_shift):
    """തിരഞ്ഞെടുത്ത ഷിഫ്റ്റിന് അനുയോജ്യമായ ഓർഡറുകൾ മാത്രം ഫിൽട്ടർ ചെയ്യുന്നു."""
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
    await update.message.reply_text(
        "👋 ഹലോ! ഷിപ്പിംഗ് ഡിലേ ട്രാക്കിംഗ് ബോട്ടിലേക്ക് സ്വാഗതം.\n\n"
        "തുടങ്ങാനായി എക്സൽ ഫയൽ (.xlsx) ഇവിടെ അയക്കുക."
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
        
        await update.message.reply_text("✅ എക്സൽ ഫയൽ സേവ് ചെയ്തു!\n\nദയവായി ഷിഫ്റ്റ് (Shift) തിരഞ്ഞെടുക്കുക:", reply_markup=reply_markup)
        return WAITING_FOR_SHIFT
    else:
        await update.message.reply_text("⚠️ ദയവായി ഒരു എക്സൽ ഫയൽ (.xlsx) മാത്രം അയക്കുക.")
        return WAITING_FOR_EXCEL

async def handle_shift_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    selected_shift = query.data
    context.user_data['selected_shift'] = selected_shift
    
    await query.edit_message_text(
        f"✅ **Shift {selected_shift}** തിരഞ്ഞെടുത്തു.\n\n"
        "ഇനി ആദ്യത്തെ ഡോക്കിന്റെ (Dock 1) HMI സ്ക്രീൻ ഫോട്ടോ അയച്ചു തരൂ."
    )
    return WAITING_FOR_PHOTO

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not os.path.exists(EXCEL_FILE_PATH):
        await update.message.reply_text("⚠️ എക്സൽ ഫയൽ കാണാനില്ല! ദയവായി ആദ്യമേ /start കൊടുത്ത് എക്സൽ അയക്കൂ.")
        return WAITING_FOR_EXCEL

    photo_file = await update.message.photo[-1].get_file()
    image_path = "temp_dock_image.jpg"
    await photo_file.download_to_drive(image_path)
    
    await update.message.reply_text("🔎 ഫോട്ടോ പരിശോധിച്ചു ഡിലേ കണക്കാക്കുന്നു...")
    
    all_photo_items = extract_photo_details(image_path)
    
    if not all_photo_items:
        await update.message.reply_text(
            "⚠️ ഫോട്ടോയിൽ നിന്ന് ഓർഡറുകൾ കണ്ടെത്താനായില്ല.\n\n"
            "➡️ ദയവായി അടുത്ത ഡോക്കിന്റെ ഫോട്ടോ അയക്കുക."
        )
        return WAITING_FOR_PHOTO

    selected_shift = context.user_data.get('selected_shift', 'C')
    excel_records = parse_excel_file(EXCEL_FILE_PATH)
    
    report_lines = []
    skipped_orders = []

    # -------------------------------------------------------------------------
    # CORE DELAY LOGIC:
    # ഫോട്ടോയിലുള്ള എല്ലാ ഓർഡറുകളുടെയും ലിസ്റ്റിൽ നിന്ന്
    # തിരഞ്ഞെടുത്ത് വച്ച ഷിഫ്റ്റിലെ ഓർഡർ ആണെങ്കിൽ, അതിന് തൊട്ടുമുമ്പത്തെ ഓർഡർ ഏതെന്ന് കണ്ടെത്തി
    # ആ മുൻ ഓർഡറിന്റെ Excel Last Out Time - പുതിയ ഓർഡറിന്റെ Photo Req Time = Delay
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

            # തൊട്ടുമുമ്പത്തെ ഓർഡർ കണ്ടെത്തുന്നു (മുമ്പത്തെ ഷിഫ്റ്റിലെ റണ്ണിംഗ് ഓർഡർ ആയിരിക്കാം)
            if i + 1 < len(all_photo_items):
                prev_order_item = all_photo_items[i + 1]
                prev_order_id = prev_order_item['order_id']
                
                # എക്സലിൽ തൊട്ടുമുമ്പത്തെ ഓർഡറിന്റെ Last Out Time ഉണ്ടോ എന്ന് നോക്കുന്നു
                prev_out_time = excel_records.get(prev_order_id.lower())

                if prev_out_time:
                    time_diff = current_req_time - prev_out_time
                    delay_minutes = int(time_diff.total_seconds() / 60)

                    report_lines.append(
                        f"🔹 **Order ID**: {current_order_id}\n"
                        f"   ├ Previous Order ({prev_order_id}) Excel Out Time: {prev_out_time.strftime('%d-%m-%Y %H:%M:%S')}\n"
                        f"   ├ Current Req Time (Photo): {current_req_time.strftime('%d-%m-%Y %H:%M:%S')}\n"
                        f"   └ ⚠️ **Delay**: **{delay_minutes} മിനിറ്റ്**"
                    )
                else:
                    skipped_orders.append(f"{current_order_id} (മുൻ ഓർഡർ {prev_order_id} എക്സലിൽ ഇല്ല)")
            else:
                skipped_orders.append(f"{current_order_id} (തൊട്ടുമുമ്പത്തെ ഓർഡർ ഫോട്ടോയിൽ ലഭ്യമല്ല)")

    # -------------------------------------------------------------------------
    # REPORT GENERATION & SEQUENTIAL REQUEST
    # -------------------------------------------------------------------------
    shift_name = f"Shift {selected_shift}"
    summary_msg = f"📊 **{shift_name} ഡിലേ റിപ്പോർട്ട് (Delay Summary)**\n\n"

    if report_lines:
        summary_msg += "\n\n".join(report_lines)
    else:
        summary_msg += "ഈ ഷിഫ്റ്റിൽ കണക്കാക്കാൻ സാധിക്കുന്ന ഓർഡറുകൾ ഒന്നും കണ്ടെത്തിയില്ല."

    if skipped_orders:
        summary_msg += f"\n\n⚠️ **സ്കിപ്പ് ചെയ്തവ**: {', '.join(skipped_orders)}"

    # 1. റിപ്പോർട്ട് അയക്കുന്നു
    await update.message.reply_text(summary_msg, parse_mode='Markdown')

    # 2. അടുത്ത ഫോട്ടോ അയക്കാനുള്ള മെസ്സേജ്
    await update.message.reply_text("➡️ **അടുത്ത ഡോക്കിന്റെ ഫോട്ടോ അയക്കുക (Send Next Dock Photo).**")

    return WAITING_FOR_PHOTO

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("പ്രോസസ്സിംഗ് നിർത്തിവച്ചു. വീണ്ടും തുടങ്ങാൻ /start നൽകുക.")
    return ConversationHandler.END

# -----------------------------------------------------------------------------
# MAIN APPLICATION
# -----------------------------------------------------------------------------
if __name__ == '__main__':
    app = Application.builder().token("8927787436:AAHMMvz2e6qcMSklrKO7cDoSI1TaGKaPU3g").build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            WAITING_FOR_EXCEL: [MessageHandler(filters.Document.ALL, handle_excel)],
            WAITING_FOR_SHIFT: [CallbackQueryHandler(handle_shift_selection)],
            WAITING_FOR_PHOTO: [MessageHandler(filters.PHOTO, handle_photo)],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    app.add_handler(conv_handler)

    print("Bot is running successfully...")
    app.run_polling()
