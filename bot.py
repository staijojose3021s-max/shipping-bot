import os
import re
import pandas as pd
import numpy as np
from PIL import Image
import easyocr
from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
)

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
BOT_TOKEN = os.environ.get("8927787436:AAFOKgAY_iVc0jEn_dt2TQTl7Qw_W9nn47E")
TEMP_IMAGE_PATH = "temp_dock_image.jpg"
OUTPUT_EXCEL_PATH = "dock_extracted_data.xlsx"

WAITING_FOR_PHOTO = 1

# Initialize EasyOCR Reader
print("Initializing EasyOCR Model...")
reader = easyocr.Reader(['en'], gpu=False)

# -----------------------------------------------------------------------------
# EASYOCR PROCESSING FUNCTION (FIXED)
# -----------------------------------------------------------------------------
def process_dock_photo_easyocr(image_path, output_excel_path):
    results = reader.readtext(image_path)
    if not results:
        raise ValueError("ഫോട്ടോയിൽ നിന്ന് വിവരങ്ങളൊന്നും വായിക്കാൻ കഴിഞ്ഞില്ല. കൂടുതൽ വ്യക്തതയുള്ള ഫോട്ടോ അയക്കുക.")

    # 1. Safe parsing of bounding boxes
    parsed_boxes = []
    for bbox, text, prob in results:
        # bbox contains 4 points: [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
        top_y = min(pt[1] for pt in bbox)
        bottom_y = max(pt[1] for pt in bbox)
        center_y = (top_y + bottom_y) / 2
        min_x = min(pt[0] for pt in bbox)
        
        parsed_boxes.append({
            'center_y': center_y,
            'min_x': min_x,
            'text': text.strip(),
            'prob': prob
        })

    # Sort boxes top-to-bottom
    parsed_boxes = sorted(parsed_boxes, key=lambda b: b['center_y'])

    # 2. Group items into distinct horizontal rows (threshold = 15 pixels)
    rows = []
    current_row = []
    last_y = None

    for box in parsed_boxes:
        if last_y is None or abs(box['center_y'] - last_y) < 18:
            current_row.append(box)
            last_y = box['center_y']
        else:
            rows.append(sorted(current_row, key=lambda b: b['min_x']))
            current_row = [box]
            last_y = box['center_y']

    if current_row:
        rows.append(sorted(current_row, key=lambda b: b['min_x']))

    extracted_records = []

    # 3. Process every table row
    for row_boxes in rows:
        row_text = " ".join([b['text'] for b in row_boxes])

        # Match 'MOS' Order Number (supporting variations)
        order_match = re.search(r'\b(M[O0][S5]\d+)\b', row_text, re.IGNORECASE)
        if not order_match:
            alt_match = re.search(r'\b(\d{10,14})\b', row_text)
            if alt_match:
                ship_order_no = "MOS" + alt_match.group(1)
            else:
                continue
        else:
            raw_id = order_match.group(1).upper()
            ship_order_no = re.sub(r'^M[O0][S5]', 'MOS', raw_id)

        # Extract Date
        date_match = re.search(r'\b\d{2}[-/\.]\d{2}[-/\.]\d{4}\b|\b\d{4}[-/\.]\d{2}[-/\.]\d{2}\b', row_text)
        req_date = date_match.group(0).replace('/', '-').replace('.', '-') if date_match else "N/A"

        # Extract Time
        time_match = re.search(r'\b\d{2}[\.:\s]\d{2}[\.:\s]\d{2}\b', row_text)
        req_time = re.sub(r'[\.\s]', ':', time_match.group(0)) if time_match else "N/A"

        # Determine Row-Specific Dock No safely
        dock_no = "1"
        mos_index = -1
        for idx, box in enumerate(row_boxes):
            if re.search(r'M[O0][S5]\d+|\d{10,14}', box['text'], re.IGNORECASE):
                mos_index = idx
                break

        if mos_index != -1 and mos_index + 1 < len(row_boxes):
            next_token = row_boxes[mos_index + 1]['text']
            dock_match = re.search(r'\b\d{1,2}\b', next_token)
            if dock_match:
                dock_no = dock_match.group(0)

        # Append row details (deduplicating by Ship Order No)
        if not any(rec['Ship Order No'] == ship_order_no for rec in extracted_records):
            extracted_records.append({
                'Ship Order No': ship_order_no,
                'Dock No': dock_no,
                'Req Date': req_date,
                'Req Time': req_time
            })

    if not extracted_records:
        raise ValueError("ഫോട്ടോയിൽ നിന്ന് ഓർഡർ നമ്പറുകൾ കണ്ടെത്താനായില്ല. ദയവായി സ്ക്രീൻ കുറച്ചുകൂടി വ്യക്തമായി എടുത്ത് അയക്കുക.")

    # Save all captured rows to Excel
    df = pd.DataFrame(extracted_records, columns=['Ship Order No', 'Dock No', 'Req Date', 'Req Time'])
    df.to_excel(output_excel_path, index=False, engine='openpyxl')
    return len(extracted_records)

# -----------------------------------------------------------------------------
# TELEGRAM BOT HANDLERS
# -----------------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "👋 സ്വാഗതം!\n\n"
        "ഡോക്കിന്റെ HMI ഫോട്ടോ അയച്ചു തരൂ. എല്ലാ ഓർഡറുകളും അതാത് ഡോക്ക് നമ്പറും കൃത്യമായി എക്സൽ ഫയലാക്കി ഞാൻ തിരികെ തരാം."
    )
    return WAITING_FOR_PHOTO

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("🔎 ഫോട്ടോ പ്രോസസ്സ് ചെയ്ത് വിവരങ്ങൾ എക്സ്ട്രാക്ട് ചെയ്യുന്നു...")
    photo_file = await update.message.photo[-1].get_file()
    await photo_file.download_to_drive(TEMP_IMAGE_PATH)
    
    try:
        total_rows = process_dock_photo_easyocr(TEMP_IMAGE_PATH, OUTPUT_EXCEL_PATH)
        with open(OUTPUT_EXCEL_PATH, 'rb') as excel_file:
            await update.message.reply_document(
                document=excel_file,
                filename="dock_extracted_data.xlsx",
                caption=f"✅ പ്രോസസ്സിംഗ് പൂർത്തിയായി!\n\n📊 ആകെ കണ്ടെത്തിയ ഓർഡറുകൾ: {total_rows}"
            )
        await update.message.reply_text("➡️ അടുത്ത ഫോട്ടോ അയക്കൂ (അല്ലെങ്കിൽ /restart ചെയ്യുക).")
    except Exception as e:
        await update.message.reply_text(f"⚠️ എറർ: {str(e)}")
        
    for path in [TEMP_IMAGE_PATH, OUTPUT_EXCEL_PATH]:
        if os.path.exists(path):
            os.remove(path)
    return WAITING_FOR_PHOTO

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("പ്രോസസ്സിംഗ് റദ്ദാക്കി. വീണ്ടും ആരംഭിക്കാൻ /start ചെയ്യുക.")
    return ConversationHandler.END

# -----------------------------------------------------------------------------
# MAIN APPLICATION
# -----------------------------------------------------------------------------
if __name__ == '__main__':
    app = Application.builder().token("8927787436:AAFOKgAY_iVc0jEn_dt2TQTl7Qw_W9nn47E").build()
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('start', start), CommandHandler('restart', start)],
        states={WAITING_FOR_PHOTO: [MessageHandler(filters.PHOTO, handle_photo)]},
        fallbacks=[CommandHandler('cancel', cancel), CommandHandler('restart', start)],
    )
    app.add_handler(conv_handler)
    print("Telegram Bot running successfully...")
    app.run_polling()