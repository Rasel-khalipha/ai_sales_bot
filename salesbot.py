import os
import re
import openpyxl
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler

EXCEL_FILE = "salesbot.xlsx"
TELEGRAM_TOKEN = "9104364470:AAFQxiPZuYBRvMgtpBDPQeSkzxpUZ31o_pU"  # Put your Telegram Bot Token here

def initialize_excel():
    """Create the Excel file with headers if it doesn't exist"""
    if os.path.exists(EXCEL_FILE):
        return
    
    wb = openpyxl.Workbook()
    ws_stock = wb.active
    ws_stock.title = "Daily_Inventory"
    
    headers = ["Product ID", "Product Name", "Initial Stock", "Daily Sold", "Remaining Stock", "Last Updated Date & Time"]
    ws_stock.append(headers)
    
    ws_report = wb.create_sheet(title="Reports")
    ws_report.append(["Report Type", "Total Products Sold", "Report Date & Time"])
    
    wb.save(EXCEL_FILE)

def extract_number(text):
    """Extract numeric value from text like '2 box' or ' 03 ' -> 3"""
    numbers = re.findall(r'\d+', str(text))
    return int(numbers[0]) if numbers else 0

def format_2digit(val):
    """Format single digits as 2 digits (e.g., 3 -> '03', 0 -> '00')"""
    try:
        num = int(val)
        return f"{num:02d}"
    except (ValueError, TypeError):
        return str(val)

# --- Telegram Command Handlers ---

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        "🤖 **Welcome to Smart Inventory AI Assistant!**\n\n"
        "🆕 **Register a single product:**\n"
        "`/new_product [ID] [Name] [Stock]`\n"
        "Example: `/new_product 01 Polo T-Shirt 50 box`\n\n"
        "📦 **Register MULTIPLE products at once:**\n"
        "`/new_multiple ID1,Name1,Stock1; ID2,Name2,Stock2`\n"
        "Example:\n"
        "`/new_multiple 01, Polo T-Shirt, 50 box; 02, Denim Jeans, 30 box`\n\n"
        "✏️ **Edit product:** `/edit 01 Cotton Shirt 60 box`\n"
        "🗑️ **Delete product:** `/delete 01`\n"
        "📥 **Add Stock:** `/add 01 05 box`\n"
        "🛒 **Record Sale:** `/sell 01 02 box`\n"
        "📊 **View Report:** `/report`"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")

async def register_new_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Register single product supporting multi-word names"""
    try:
        args = context.args
        if len(args) < 3:
            await update.message.reply_text("❌ Format: `/new_product [ID] [Name] [Stock]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        product_name = " ".join(args[1:-1]).strip()
        initial_stock = extract_number(args[-1])
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Daily_Inventory"]
        
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == product_id:
                await update.message.reply_text(f"⚠️ Product ID ({format_2digit(product_id)}) already exists!")
                return
        
        new_row_idx = ws.max_row + 1
        ws.cell(row=new_row_idx, column=1, value=product_id)
        ws.cell(row=new_row_idx, column=2, value=product_name)
        ws.cell(row=new_row_idx, column=3, value=initial_stock)
        ws.cell(row=new_row_idx, column=4, value=0)
        ws.cell(row=new_row_idx, column=5, value=f'=C{new_row_idx}-D{new_row_idx}')
        ws.cell(row=new_row_idx, column=6, value=current_time)
        
        wb.save(EXCEL_FILE)
        
        reply_msg = (
            f"✅ **New Product Added Successfully!**\n"
            f"🏷️ **Name:** {product_name}\n"
            f"🆔 **ID:** {format_2digit(product_id)}\n"
            f"📦 **Initial Stock:** {format_2digit(initial_stock)} boxes\n"
            f"⏰ **Time:** {current_time}"
        )
        await update.message.reply_text(reply_msg, parse_mode="Markdown")
        
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error occurred: {str(e)}")

async def register_multiple_products(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Register multiple products separated by semicolon (;), supporting spaces and 2-digit format"""
    try:
        full_text = update.message.text
        data_part = full_text.replace("/new_multiple", "").strip()
        
        if not data_part:
            await update.message.reply_text(
                "❌ Correct Format Example:\n"
                "`/new_multiple 01, Polo T-Shirt, 50 box; 02, Jeans, 30 box`",
                parse_mode="Markdown"
            )
            return

        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Daily_Inventory"]
        
        items = data_part.split(";")
        success_count = 0
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        for item in items:
            item = item.strip()
            if not item:
                continue
            
            parts = item.split(",")
            if len(parts) < 3:
                continue
                
            product_id = extract_number(parts[0].strip())
            # মাঝের অংশগুলোর স্পেস ঠিক রেখে নাম তৈরি করা
            product_name = ",".join(parts[1:-1]).strip()
            initial_stock = extract_number(parts[-1].strip())
            
            # ডুপ্লিকেট আইডি চেক করা
            exists = False
            for row in range(2, ws.max_row + 1):
                if ws.cell(row=row, column=1).value == product_id:
                    exists = True
                    break
            
            if not exists:
                new_row_idx = ws.max_row + 1
                ws.cell(row=new_row_idx, column=1, value=product_id)
                ws.cell(row=new_row_idx, column=2, value=product_name)
                ws.cell(row=new_row_idx, column=3, value=initial_stock)
                ws.cell(row=new_row_idx, column=4, value=0)
                ws.cell(row=new_row_idx, column=5, value=f'=C{new_row_idx}-D{new_row_idx}')
                ws.cell(row=new_row_idx, column=6, value=current_time)
                success_count += 1

        wb.save(EXCEL_FILE)
        await update.message.reply_text(f"✅ Successfully added **{format_2digit(success_count)}** products together at {current_time}!")
        
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error occurred: {str(e)}")

async def edit_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Edit existing product name or initial stock"""
    try:
        args = context.args
        if len(args) < 3:
            await update.message.reply_text("❌ Format: `/edit [ID] [New_Name] [New_Stock]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        new_name = " ".join(args[1:-1]).strip()
        new_stock = extract_number(args[-1])
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Daily_Inventory"]
        
        found = False
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == product_id:
                ws.cell(row=row, column=2, value=new_name)
                ws.cell(row=row, column=3, value=new_stock)
                ws.cell(row=row, column=6, value=current_time)
                found = True
                break
                
        if found:
            wb.save(EXCEL_FILE)
            reply_msg = (
                f"✏️ **Product Updated Successfully!**\n"
                f"🆔 **ID:** {format_2digit(product_id)}\n"
                f"🏷️ **New Name:** {new_name}\n"
                f"📦 **New Initial Stock:** {format_2digit(new_stock)} boxes\n"
                f"⏰ **Time:** {current_time}"
            )
            await update.message.reply_text(reply_msg, parse_mode="Markdown")
        else:
            await update.message.reply_text(f"❌ Product ID {format_2digit(product_id)} not found.")
            
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error occurred: {str(e)}")

async def delete_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Delete a product row from inventory Excel sheet"""
    try:
        args = context.args
        if len(args) < 1:
            await update.message.reply_text("❌ Format: `/delete [ID]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Daily_Inventory"]
        
        found = False
        deleted_name = ""
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == product_id:
                deleted_name = ws.cell(row=row, column=2).value
                ws.delete_rows(row)
                found = True
                break
                
        if found:
            wb.save(EXCEL_FILE)
            reply_msg = (
                f"🗑️ **Product Deleted Successfully!**\n"
                f"🏷️ **Name:** {deleted_name}\n"
                f"🆔 **ID:** {format_2digit(product_id)}"
            )
            await update.message.reply_text(reply_msg, parse_mode="Markdown")
        else:
            await update.message.reply_text(f"❌ Product ID {format_2digit(product_id)} not found.")
            
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error occurred: {str(e)}")

async def add_stock(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Add stock to an existing product"""
    try:
        args = context.args
        if len(args) < 2:
            await update.message.reply_text("❌ Format: `/add [ID] [Quantity]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        add_qty = extract_number(" ".join(args[1:]))
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Daily_Inventory"]
        
        found = False
        product_name = ""
        new_initial = 0
        remaining = 0
        
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == product_id:
                product_name = ws.cell(row=row, column=2).value
                current_initial = ws.cell(row=row, column=3).value or 0
                current_sold = ws.cell(row=row, column=4).value or 0
                
                new_initial = current_initial + add_qty
                ws.cell(row=row, column=3).value = new_initial
                
                remaining = new_initial - current_sold
                ws.cell(row=row, column=6).value = current_time
                found = True
                break
                
        if found:
            wb.save(EXCEL_FILE)
            reply_msg = (
                f"📥 **Stock Updated (Stock In)!**\n"
                f"🏷️ **Product:** {product_name}\n"
                f"🆔 **ID:** {format_2digit(product_id)}\n"
                f"➕ **Added:** {format_2digit(add_qty)} boxes\n"
                f"📈 **Total Stock:** {format_2digit(new_initial)} boxes\n"
                f"📉 **Remaining:** {format_2digit(remaining)} boxes\n"
                f"⏰ **Time:** {current_time}"
            )
            await update.message.reply_text(reply_msg, parse_mode="Markdown")
        else:
            await update.message.reply_text(f"❌ Product ID {format_2digit(product_id)} not found.")
            
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error occurred: {str(e)}")

async def sell_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Record a sale"""
    try:
        args = context.args
        if len(args) < 2:
            await update.message.reply_text("❌ Format: `/sell [ID] [Quantity]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        sold_qty = extract_number(" ".join(args[1:]))
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Daily_Inventory"]
        
        found = False
        product_name = ""
        remaining = 0
        initial = 0
        
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == product_id:
                product_name = ws.cell(row=row, column=2).value
                current_sold = ws.cell(row=row, column=4).value or 0
                initial = ws.cell(row=row, column=3).value or 0
                
                new_sold = current_sold + sold_qty
                
                if new_sold > initial:
                    await update.message.reply_text(f"⚠️ Insufficient stock! '{product_name}' has {format_2digit(initial)} boxes left.")
                    return
                
                ws.cell(row=row, column=4).value = new_sold
                ws.cell(row=row, column=6).value = current_time
                
                remaining = initial - new_sold
                found = True
                break
                
        if found:
            wb.save(EXCEL_FILE)
            reply_msg = (
                f"🛒 **Sale Recorded Successfully!**\n"
                f"🏷️ **Product:** {product_name}\n"
                f"🆔 **ID:** {format_2digit(product_id)}\n"
                f"➖ **Sold:** {format_2digit(sold_qty)} boxes\n"
                f"📉 **Remaining:** {format_2digit(remaining)} boxes\n"
                f"⏰ **Time:** {current_time}"
            )
            await update.message.reply_text(reply_msg, parse_mode="Markdown")
        else:
            await update.message.reply_text(f"❌ Product ID {format_2digit(product_id)} not found.")
            
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error occurred: {str(e)}")

async def get_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """View inventory & sales report in 2-digit format"""
    try:
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_stock = wb["Daily_Inventory"]
        
        if ws_stock.max_row <= 1:
            await update.message.reply_text("📭 Inventory is empty!")
            return
            
        total_items_sold = 0
        report_summary = "📊 **Current Inventory & Sales Report:**\n\n"
        
        for row in range(2, ws_stock.max_row + 1):
            p_id = ws_stock.cell(row=row, column=1).value
            p_name = ws_stock.cell(row=row, column=2).value
            initial = ws_stock.cell(row=row, column=3).value or 0
            sold = ws_stock.cell(row=row, column=4).value or 0
            rem = initial - sold
            last_time = ws_stock.cell(row=row, column=6).value or "N/A"
            
            total_items_sold += sold
            report_summary += (
                f"🔹 **{p_name}** (ID: {format_2digit(p_id)})\n"
                f"   - Initial: {format_2digit(initial)} | Sold: {format_2digit(sold)} | Remaining: {format_2digit(rem)} boxes\n"
                f"   - 🕒 Last Updated: {last_time}\n\n"
            )
            
        report_summary += f"🔥 **Total Products Sold:** {format_2digit(total_items_sold)} boxes"
        await update.message.reply_text(report_summary, parse_mode="Markdown")
        
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error occurred: {str(e)}")

def main():
    initialize_excel()
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("new_product", register_new_product))
    app.add_handler(CommandHandler("new_multiple", register_multiple_products))
    app.add_handler(CommandHandler("edit", edit_product))
    app.add_handler(CommandHandler("delete", delete_product))
    app.add_handler(CommandHandler("add", add_stock))
    app.add_handler(CommandHandler("sell", sell_product))
    app.add_handler(CommandHandler("report", get_report))
    
    print("🤖 Telegram bot running with multiple product support, spaces, edit, delete & 2-digit formatting...")
    app.run_polling()

if __name__ == "__main__":
    main()