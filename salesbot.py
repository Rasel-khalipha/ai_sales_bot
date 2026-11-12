import os
import re
import openpyxl
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler

EXCEL_FILE = "salesbot.xlsx"
WEEKLY_REPORT_FILE = "Weekly_Sales_Report.xlsx"
MONTHLY_REPORT_FILE = "Monthly_Sales_Report.xlsx"
TELEGRAM_TOKEN = "9104364470:AAFQxiPZuYBRvMgtpBDPQeSkzxpUZ31o_pU"  # Put your Telegram Bot Token here

def initialize_excel():
    """Create the Excel file with necessary sheets and headers if it doesn't exist"""
    if os.path.exists(EXCEL_FILE):
        return
    
    wb = openpyxl.Workbook()
    
    # Sheet 1: Inventory with Value
    ws_stock = wb.active
    ws_stock.title = "Inventory"
    ws_stock.append(["Product ID", "Product Name", "Initial Stock", "Daily Sold", "Remaining Stock", "Unit Price", "Total Value", "Last Updated"])
    
    # Sheet 2: Sales History
    ws_history = wb.create_sheet(title="Sales_History")
    ws_history.append(["Product ID", "Product Name", "Sold Quantity", "Unit Price", "Total Sold Value", "Date & Time"])
    
    wb.save(EXCEL_FILE)

def extract_number(text):
    """Extract numeric value from text like '2 box' or ' 500 ' -> number"""
    numbers = re.findall(r'\d+', str(text))
    return int(numbers[0]) if numbers else 0

def format_2digit(val):
    """Format single digits as 2 digits (e.g., 3 -> '03', 0 -> '00')"""
    try:
        num = int(val)
        return f"{num:02d}"
    except (ValueError, TypeError):
        return str(val)

def sort_inventory_by_id(ws):
    """Sort inventory rows based on Product ID in ascending order"""
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    if not rows:
        return
    
    rows.sort(key=lambda x: int(x[0]) if x[0] is not None and str(x[0]).isdigit() else 0)
    
    for row in range(ws.max_row, 1, -1):
        ws.delete_rows(row)
        
    for r_idx, row_data in enumerate(rows, start=2):
        for c_idx, val in enumerate(row_data, start=1):
            if c_idx == 5:
                ws.cell(row=r_idx, column=c_idx, value=f'=C{r_idx}-D{r_idx}')
            elif c_idx == 7:
                ws.cell(row=r_idx, column=c_idx, value=f'=E{r_idx}*F{r_idx}')
            else:
                ws.cell(row=r_idx, column=c_idx, value=val)

def generate_weekly_report_file(wb_main):
    """সাপ্তাহিক সেলস রিপোর্ট আলাদা এক্সেল ফাইলে তৈরি করা"""
    ws_history = wb_main["Sales_History"]
    now = datetime.now()
    current_week = now.isocalendar()[1]
    current_year = now.year
    
    wb_weekly = openpyxl.Workbook()
    ws_weekly = wb_weekly.active
    ws_weekly.title = "Weekly_Summary"
    ws_weekly.append(["Product ID", "Product Name", "Total Sold Qty", "Total Sales Value", "Week & Year"])
    
    product_summary = {}
    
    for row in range(2, ws_history.max_row + 1):
        p_id = ws_history.cell(row=row, column=1).value
        p_name = ws_history.cell(row=row, column=2).value
        s_qty = ws_history.cell(row=row, column=3).value or 0
        u_price = ws_history.cell(row=row, column=4).value or 0
        s_date = ws_history.cell(row=row, column=6).value
        
        if isinstance(s_date, str):
            try:
                s_date = datetime.strptime(s_date.split()[0], "%Y-%m-%d")
            except:
                continue
                
        if isinstance(s_date, datetime) and s_date.isocalendar()[1] == current_week and s_date.year == current_year:
            if p_id not in product_summary:
                product_summary[p_id] = {"name": p_name, "qty": 0, "value": 0}
            product_summary[p_id]["qty"] += s_qty
            product_summary[p_id]["value"] += (s_qty * u_price)
            
    total_weekly_sales_val = 0
    for p_id, data in product_summary.items():
        ws_weekly.append([p_id, data["name"], data["qty"], data["value"], f"Week-{current_week}/{current_year}"])
        total_weekly_sales_val += data["value"]
        
    ws_weekly.append([])
    ws_weekly.append(["Grand Total Weekly Sales Value:", "", "", total_weekly_sales_val, f"Week-{current_week}/{current_year}"])
    
    wb_weekly.save(WEEKLY_REPORT_FILE)
    return total_weekly_sales_val

def generate_monthly_report_file(wb_main):
    """মাসিক সেলস রিপোর্ট আলাদা এক্সেল ফাইলে তৈরি করা"""
    ws_history = wb_main["Sales_History"]
    now = datetime.now()
    current_month = now.month
    current_year = now.year
    
    wb_monthly = openpyxl.Workbook()
    ws_monthly = wb_monthly.active
    ws_monthly.title = "Monthly_Summary"
    ws_monthly.append(["Product ID", "Product Name", "Total Sold Qty", "Total Sales Value", "Month & Year"])
    
    product_summary = {}
    
    for row in range(2, ws_history.max_row + 1):
        p_id = ws_history.cell(row=row, column=1).value
        p_name = ws_history.cell(row=row, column=2).value
        s_qty = ws_history.cell(row=row, column=3).value or 0
        u_price = ws_history.cell(row=row, column=4).value or 0
        s_date = ws_history.cell(row=row, column=6).value
        
        if isinstance(s_date, str):
            try:
                s_date = datetime.strptime(s_date.split()[0], "%Y-%m-%d")
            except:
                continue
                
        if isinstance(s_date, datetime) and s_date.month == current_month and s_date.year == current_year:
            if p_id not in product_summary:
                product_summary[p_id] = {"name": p_name, "qty": 0, "value": 0}
            product_summary[p_id]["qty"] += s_qty
            product_summary[p_id]["value"] += (s_qty * u_price)
            
    total_monthly_sales_val = 0
    for p_id, data in product_summary.items():
        ws_monthly.append([p_id, data["name"], data["qty"], data["value"], f"{current_month}/{current_year}"])
        total_monthly_sales_val += data["value"]
        
    ws_monthly.append([])
    ws_monthly.append(["Grand Total Monthly Sales Value:", "", "", total_monthly_sales_val, f"{current_month}/{current_year}"])
    
    wb_monthly.save(MONTHLY_REPORT_FILE)
    return total_monthly_sales_val

# --- Telegram Command Handlers ---

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        "🤖 **Smart Inventory & Multi-Report Assistant**\n\n"
        "🆕 **Register Product:** `/new_product [ID] [Name] [Stock] [Price]`\n"
        "📦 **Multiple Products:** `/new_multiple 01, Shirt, 50, 500; 02, Pant, 30, 800`\n"
        "📥 **Add Stock:** `/add 01 10 box`\n"
        "🛒 **Sell Product:** `/sell 01 02 box`\n"
        "📊 **View Report & Value:** `/report`\n"
        "📁 **Export Weekly File:** `/weekly_export`\n"
        "📁 **Export Monthly File:** `/monthly_export`"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")

async def register_new_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        args = context.args
        if len(args) < 4:
            await update.message.reply_text("❌ Format: `/new_product [ID] [Name] [Stock] [Price]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        product_name = " ".join(args[1:-2]).strip()
        initial_stock = extract_number(args[-2])
        unit_price = extract_number(args[-1])
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Inventory"]
        
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
        ws.cell(row=new_row_idx, column=6, value=unit_price)
        ws.cell(row=new_row_idx, column=7, value=f'=E{new_row_idx}*F{new_row_idx}')
        ws.cell(row=new_row_idx, column=8, value=current_time)
        
        sort_inventory_by_id(ws)
        wb.save(EXCEL_FILE)
        
        await update.message.reply_text(
            f"✅ **Product Added Successfully!**\n"
            f"🆔 ID: {format_2digit(product_id)} | 🏷️ Name: {product_name}\n"
            f"📦 Stock: {format_2digit(initial_stock)} | 💵 Unit Price: {unit_price} BDT",
            parse_mode="Markdown"
        )
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

async def register_multiple_products(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        full_text = update.message.text
        data_part = full_text.replace("/new_multiple", "").strip()
        
        if not data_part:
            await update.message.reply_text("❌ Format: `/new_multiple 01, Shirt, 50, 500; 02, Pant, 30, 800`", parse_mode="Markdown")
            return

        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Inventory"]
        
        items = data_part.split(";")
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        success_count = 0
        
        for item in items:
            item = item.strip()
            if not item:
                continue
            parts = item.split(",")
            if len(parts) < 4:
                continue
                
            product_id = extract_number(parts[0].strip())
            product_name = ",".join(parts[1:-2]).strip()
            initial_stock = extract_number(parts[-2].strip())
            unit_price = extract_number(parts[-1].strip())
            
            exists = any(ws.cell(row=r, column=1).value == product_id for r in range(2, ws.max_row + 1))
            if not exists:
                new_row_idx = ws.max_row + 1
                ws.cell(row=new_row_idx, column=1, value=product_id)
                ws.cell(row=new_row_idx, column=2, value=product_name)
                ws.cell(row=new_row_idx, column=3, value=initial_stock)
                ws.cell(row=new_row_idx, column=4, value=0)
                ws.cell(row=new_row_idx, column=5, value=f'=C{new_row_idx}-D{new_row_idx}')
                ws.cell(row=new_row_idx, column=6, value=unit_price)
                ws.cell(row=new_row_idx, column=7, value=f'=E{new_row_idx}*F{new_row_idx}')
                ws.cell(row=new_row_idx, column=8, value=current_time)
                success_count += 1

        sort_inventory_by_id(ws)
        wb.save(EXCEL_FILE)
        await update.message.reply_text(f"✅ Successfully added **{format_2digit(success_count)}** products with pricing!")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

async def add_stock(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
        ws = wb["Inventory"]
        
        found = False
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == product_id:
                p_name = ws.cell(row=row, column=2).value
                curr_init = ws.cell(row=row, column=3).value or 0
                new_init = curr_init + add_qty
                ws.cell(row=row, column=3, value=new_init)
                ws.cell(row=row, column=8, value=current_time)
                found = True
                break
                
        if found:
            wb.save(EXCEL_FILE)
            await update.message.reply_text(f"📥 **Stock Added!**\n🆔 ID: {format_2digit(product_id)} | 🏷️ Name: {p_name}\n➕ Added: {format_2digit(add_qty)} | 📈 New Initial: {format_2digit(new_init)}")
        else:
            await update.message.reply_text(f"❌ Product ID {format_2digit(product_id)} not found.")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

async def sell_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        args = context.args
        if len(args) < 2:
            await update.message.reply_text("❌ Format: `/sell [ID] [Quantity]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        sold_qty = extract_number(" ".join(args[1:]))
        current_time_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        current_date_obj = datetime.now()
        
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Inventory"]
        ws_history = wb["Sales_History"]
        
        found = False
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == product_id:
                p_name = ws.cell(row=row, column=2).value
                initial = ws.cell(row=row, column=3).value or 0
                curr_sold = ws.cell(row=row, column=4).value or 0
                u_price = ws.cell(row=row, column=6).value or 0
                
                new_sold = curr_sold + sold_qty
                remaining = initial - new_sold
                if remaining < 0:
                    await update.message.reply_text(f"⚠️ Insufficient stock! '{p_name}' has max {format_2digit(initial)} boxes available.")
                    return
                
                ws.cell(row=row, column=4, value=new_sold)
                ws.cell(row=row, column=8, value=current_time_str)
                
                total_sold_value = sold_qty * u_price
                ws_history.append([product_id, p_name, sold_qty, u_price, total_sold_value, current_date_obj])
                found = True
                break
                
        if found:
            wb.save(EXCEL_FILE)
            await update.message.reply_text(f"🛒 **Sale Recorded!**\n🆔 ID: {format_2digit(product_id)} | 🏷️ Name: {p_name}\n➖ Sold: {format_2digit(sold_qty)} boxes | 💵 Total Value Deducted: {sold_qty * u_price} BDT")
        else:
            await update.message.reply_text(f"❌ Product ID {format_2digit(product_id)} not found.")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

async def get_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_stock = wb["Inventory"]
        ws_history = wb["Sales_History"]
        
        if ws_stock.max_row <= 1:
            await update.message.reply_text("📭 Inventory is empty!")
            return
            
        now = datetime.now()
        today_str = now.strftime("%Y-%m-%d")
        current_week = now.isocalendar()[1]
        current_month = now.month
        current_year = now.year
        
        daily_sales_qty, daily_sales_val = 0, 0
        weekly_sales_qty, weekly_sales_val = 0, 0
        monthly_sales_qty, monthly_sales_val = 0, 0
        
        for row in range(2, ws_history.max_row + 1):
            s_qty = ws_history.cell(row=row, column=3).value or 0
            u_price = ws_history.cell(row=row, column=4).value or 0
            s_val = ws_history.cell(row=row, column=5).value or (s_qty * u_price)
            s_date = ws_history.cell(row=row, column=6).value
            
            if isinstance(s_date, str):
                try:
                    s_date = datetime.strptime(s_date.split()[0], "%Y-%m-%d")
                except:
                    continue
            
            if isinstance(s_date, datetime):
                if s_date.strftime("%Y-%m-%d") == today_str:
                    daily_sales_qty += s_qty
                    daily_sales_val += s_val
                if s_date.isocalendar()[1] == current_week and s_date.year == current_year:
                    weekly_sales_qty += s_qty
                    weekly_sales_val += s_val
                if s_date.month == current_month and s_date.year == current_year:
                    monthly_sales_qty += s_qty
                    monthly_sales_val += s_val

        grand_total_inventory_value = 0
        report_msg = "📊 **Inventory & Sales Report Summary** 📊\n\n"
        report_msg += f"☀️ **Today's Sales:** {format_2digit(daily_sales_qty)} boxes (Value: {daily_sales_val} BDT)\n"
        report_msg += f"📅 **This Week's Sales:** {format_2digit(weekly_sales_qty)} boxes (Value: {weekly_sales_val} BDT)\n"
        report_msg += f"🗓️ **This Month's Sales:** {format_2digit(monthly_sales_qty)} boxes (Value: {monthly_sales_val} BDT)\n\n"
        report_msg += "📦 **Current Stock & Value Status:**\n"
        
        for row in range(2, ws_stock.max_row + 1):
            p_id = ws_stock.cell(row=row, column=1).value
            p_name = ws_stock.cell(row=row, column=2).value
            initial = ws_stock.cell(row=row, column=3).value or 0
            sold = ws_stock.cell(row=row, column=4).value or 0
            rem = initial - sold
            price = ws_stock.cell(row=row, column=6).value or 0
            tot_val = rem * price
            grand_total_inventory_value += tot_val
            
            report_msg += f"• ID {format_2digit(p_id)} | **{p_name}**: Rem {format_2digit(rem)} box | Price: {price} | Val: {tot_val} BDT\n"
            
        report_msg += f"\n💎 **Total Inventory Remaining Value:** **{grand_total_inventory_value} BDT**"
        await update.message.reply_text(report_msg, parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

async def export_weekly_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """সাপ্তাহিক এক্সেল ফাইল জেনারেট করার কমান্ড"""
    try:
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        total_val = generate_weekly_report_file(wb)
        await update.message.reply_text(f"📁 **Weekly Report File Generated!**\nFile Name: `{WEEKLY_REPORT_FILE}`\nTotal Weekly Sales Value: {total_val} BDT", parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

async def export_monthly_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """মাসিক এক্সেল ফাইল জেনারেট করার কমান্ড"""
    try:
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        total_val = generate_monthly_report_file(wb)
        await update.message.reply_text(f"📁 **Monthly Report File Generated!**\nFile Name: `{MONTHLY_REPORT_FILE}`\nTotal Monthly Sales Value: {total_val} BDT", parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

def main():
    initialize_excel()
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("new_product", register_new_product))
    app.add_handler(CommandHandler("new_multiple", register_multiple_products))
    app.add_handler(CommandHandler("add", add_stock))
    app.add_handler(CommandHandler("sell", sell_product))
    app.add_handler(CommandHandler("report", get_report))
    app.add_handler(CommandHandler("weekly_export", export_weekly_file))
    app.add_handler(CommandHandler("monthly_export", export_monthly_file))
    
    print("🤖 Telegram Bot running with Weekly & Monthly Export Files, Value Tracking & Sorting...")
    app.run_polling()

if __name__ == "__main__":
    main()