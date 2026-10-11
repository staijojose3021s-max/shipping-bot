


import os
import pandas as pd
import datetime as dt
import zipfile
import shutil
from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
)

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
BOT_TOKEN = os.environ.get("8927787436:AAFOKgAY_iVc0jEn_dt2TQTl7Qw_W9nn47E")
INPUT_EXCEL_PATH = "input_dispatch.xlsx"
OUTPUT_EXCEL_PATH = "cleaned_dispatch_data.xlsx"

# Conversation State
WAITING_FOR_EXCEL = range(1)

# -----------------------------------------------------------------------------
# EXCEL HELPER FUNCTION
# -----------------------------------------------------------------------------
def process_and_clean_excel(file_path, output_path):
    """
    Cleans metadata to prevent openpyxl crashes, dynamically extracts 'Order ID' 
    and 'Last Out Time' columns, groups by 'Order ID' to keep only the latest 
    (maximum) 'Last Out Time', and saves to a new Excel file.
    """
    # 1. Clean metadata to prevent openpyxl core.xml crashes
    try:
        temp_path = file_path + ".tmp"
        with zipfile.ZipFile(file_path, 'r') as zin:
            with zipfile.ZipFile(temp_path, 'w') as zout:
                for item in zin.infolist():
                    if item.filename == 'docProps/core.xml':
                        continue
                    zout.writestr(item, zin.read(item.filename))
        shutil.move(temp_path, file_path)
    except Exception as e:
        print(f"Sanitization warning: {e}")

    # 2. Read workbook data
    import openpyxl
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active
    
    data = []
    headers = []
    for i, row in enumerate(sheet.iter_rows(values_only=True)):
        if i == 0:
            headers = [str(cell).strip() if cell is not None else f"Unnamed_{j}" for j, cell in enumerate(row)]
            continue
        if any(row):
            data.append(row)
            
    df = pd.DataFrame(data, columns=headers)
    df.columns = df.columns.astype(str).str.strip()

    # 3. Dynamic Header Search for Order ID
    order_col = None
    target_order_headers = [
        'ship order id', 'shipping order id', 'shipping order number',
        'ship order no', 'shipping order', 'ship_order_no', 'order id', 'order_id', 'order no', 'order number'
    ]
    for col in df.columns:
        if col.lower().strip() in target_order_headers:
            order_col = col
            break
    if not order_col:
        order_col = df.columns[0]

    # 4. Dynamic Header Search for Last Out Time
    time_col = None
    target_time_headers = ['last out time', 'req date', 'req time', 'out time', 'request time', 'last_out_time', 'gnt out date']
    for col in df.columns:
        if col.lower().strip() in target_time_headers:
            time_col = col
            break
    if not time_col:
        time_col = df.columns[-1]

    # 5. Extract and parse datetime objects safely
    parsed_records = []
    for _, row in df.iterrows():
        order_id = str(row[order_col]).strip()
        if not order_id or order_id.lower() in ['none', 'nan', '']:
            continue
            
        val = row[time_col]
        try:
            if isinstance(val, (dt.datetime, dt.date)):
                parsed_time = val if isinstance(val, dt.datetime) else dt.datetime.combine(val, dt.time.min)
            else:
                parsed_time = pd.to_datetime(str(val), dayfirst=True, errors='coerce')
                if pd.isna(parsed_time):
                    continue
                parsed_time = parsed_time.to_pydatetime()
        except Exception as e:
            print(f"Error parsing date for order {order_id}: {e}")
            continue

        parsed_records.append({
            'Order ID': order_id,
            'Last Out Time': parsed_time
        })

    if not parsed_records:
        raise ValueError("No valid records found in the uploaded Excel file.")

    # 6. Convert to DataFrame and Group By 'Order ID' to get the latest (maximum) time
    temp_df = pd.DataFrame(parsed_records)
    grouped_df = temp_df.groupby('Order ID', as_index=False)['Last Out Time'].max()

    # 7. Format datetime safely back to standard string format
    grouped_df['Last Out Time'] = grouped_df['Last Out Time'].apply(
        lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) and isinstance(x, (dt.datetime, dt.date)) else str(x)
    )

    # 8. Save extracted unique data to a new Excel file
    grouped_df.to_excel(output_path, index=False, engine='openpyxl')
    return len(grouped_df)

# -----------------------------------------------------------------------------
# TELEGRAM BOT HANDLERS
# -----------------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "👋 Welcome to the Excel Processing Bot!\n\n"
        "Please upload your Excel file (.xlsx) containing dispatch details."
    )
    return WAITING_FOR_EXCEL

async def handle_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("📥 Excel file received! Filtering latest Last Out Time per Order ID...")
        
        file = await doc.get_file()
        await file.download_to_drive(INPUT_EXCEL_PATH)
        
        try:
            total_rows = process_and_clean_excel(INPUT_EXCEL_PATH, OUTPUT_EXCEL_PATH)
            
            # Send the newly created Excel file back to user
            with open(OUTPUT_EXCEL_PATH, 'rb') as excel_file:
                await update.message.reply_document(
                    document=excel_file,
                    filename="cleaned_dispatch_data.xlsx",
                    caption=f"✅ Processing complete!\n\n📊 Unique Orders Extracted: {total_rows}"
                )
        except Exception as e:
            await update.message.reply_text(f"⚠️ Error processing file: {str(e)}")
            
        # Clean up local files
        for path in [INPUT_EXCEL_PATH, OUTPUT_EXCEL_PATH]:
            if os.path.exists(path):
                os.remove(path)
                
        return WAITING_FOR_EXCEL
    else:
        await update.message.reply_text("⚠️ Please upload a valid Excel file (.xlsx or .xls) only.")
        return WAITING_FOR_EXCEL

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Processing cancelled. Send /start to begin again.")
    return ConversationHandler.END

# -----------------------------------------------------------------------------
# MAIN APPLICATION
# -----------------------------------------------------------------------------
if __name__ == '__main__':
    app = Application.builder().token("8927787436:AAFOKgAY_iVc0jEn_dt2TQTl7Qw_W9nn47E").build()

    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler('start', start)
        ],
        states={
            WAITING_FOR_EXCEL: [MessageHandler(filters.Document.ALL, handle_excel)],
        },
        fallbacks=[
            CommandHandler('cancel', cancel)
        ],
    )

    app.add_handler(conv_handler)

    print("Bot is running successfully...")
    app.run_polling()