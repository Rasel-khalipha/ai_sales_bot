import os
import openpyxl
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler

EXCEL_FILE = "inventory_sales.xlsx"
TELEGRAM_TOKEN = "9104364470:AAFQxiPZuYBRvMgtpBDPQeSkzxpUZ31o_pU"  # আপনার বটের টোকেন দিন

def initialize_excel():
    """এক্সেল ফাইল না থাকলে নতুন তৈরি করবে"""
    if os.path.exists(EXCEL_FILE):
        return
    
    wb = openpyxl.Workbook()
    ws_stock = wb.active
    ws_stock.title = "Daily_Inventory"
    
    # হেডার তৈরি (Date & Time কলাম সহ)
    headers = ["Product ID", "Product Name", "Initial Stock", "Daily Sold", "Remaining Stock", "Last Updated Date & Time"]
    ws_stock.append(headers)
    
    ws_report = wb.create_sheet(title="Reports")
    ws_report.append(["Report Type", "Total Products Sold", "Report Date & Time"])
    
    wb.save(EXCEL_FILE)

# --- টেলিগ্রাম কমান্ড হ্যান্ডলার ---

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        "🤖 **স্মার্ট ইনভেন্টরি এআই অ্যাসিস্ট্যান্টে স্বাগতম!**\n\n"
        "🆕 **একটি নতুন প্রোডাক্ট রেজিস্টার করতে:**\n"
        "`/new_product [ID] [Name] [Initial_Stock]`\n"
        "উদাহরণ: `/new_product 101 T-Shirt 50`\n\n"
        "📦 **একসাথে একাধিক প্রোডাক্ট রেজিস্টার করতে:**\n"
        "`/new_multiple 101,T-Shirt,50; 102,Jeans,30`\n\n"
        "📥 **বিদ্যমান প্রোডাক্টে স্টক যোগ করতে:**\n"
        "`/add [Product_ID] [Quantity]`\n"
        "উদাহরণ: `/add 101 20`\n\n"
        "🛒 **পণ্য বিক্রি হলে (Sales):**\n"
        "`/sell [Product_ID] [Quantity]`\n"
        "উদাহরণ: `/sell 101 5`\n\n"
        "📊 **স্টক ও সেলস রিপোর্ট দেখতে:**\n"
        "`/report`"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")

async def register_new_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """নতুন প্রোডাক্ট রেজিস্টার (তারিখ ও সময়সহ)"""
    try:
        args = context.args
        if len(args) < 3:
            await update.message.reply_text(
                "❌ ভুল ফরম্যাট! সঠিক নিয়মে লিখুন:\n`/new_product [ID] [Name] [Stock]`", 
                parse_mode="Markdown"
            )
            return
            
        product_id = int(args[0])
        product_name = args[1]
        initial_stock = int(args[2])
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Daily_Inventory"]
        
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == product_id:
                await update.message.reply_text(f"⚠️ এই প্রোডাক্ট আইডি ({product_id}) অলরেডি তালিকায় আছে!")
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
            f"✅ **নতুন প্রোডাক্ট সফলভাবে যুক্ত হয়েছে!**\n"
            f"🏷️ **নাম:** {product_name}\n"
            f"🆔 **আইডি:** {product_id}\n"
            f"📦 **প্রাথমিক স্টক:** {initial_stock} পিস\n"
            f"⏰ **সময়:** {current_time}"
        )
        await update.message.reply_text(reply_msg, parse_mode="Markdown")
        
    except Exception as e:
        await update.message.reply_text(f"⚠️ ত্রুটি ঘটেছে: {str(e)}")

async def register_multiple_products(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """একসাথে একাধিক প্রোডাক্ট রেজিস্টার (তারিখ ও সময়সহ)"""
    try:
        full_text = update.message.text
        data_part = full_text.replace("/new_multiple", "").strip()
        
        if not data_part:
            await update.message.reply_text("❌ উদাহরণ: `/new_multiple 101,T-Shirt,50; 102,Jeans,30`", parse_mode="Markdown")
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
                
            product_id = int(parts[0].strip())
            product_name = parts[1].strip()
            initial_stock = int(parts[2].strip())
            
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
        await update.message.reply_text(f"✅ **{success_count}টি** প্রোডাক্ট সময়সহ ({current_time}) ইনভেন্টরিতে যুক্ত করা হয়েছে!")
        
    except Exception as e:
        await update.message.reply_text(f"⚠️ ত্রুটি ঘটেছে: {str(e)}")

async def add_stock(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """স্টক আপডেট (তারিখ ও সময়সহ)"""
    try:
        args = context.args
        if len(args) < 2:
            await update.message.reply_text("❌ ভুল ফরম্যাট! সঠিক নিয়মে লিখুন: `/add 101 20`", parse_mode="Markdown")
            return
            
        product_id = int(args[0])
        add_qty = int(args[1])
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
                f"📥 **স্টক আপডেট সফল (Stock In)!**\n"
                f"🏷️ **প্রোডাক্টের নাম:** {product_name}\n"
                f"🆔 **আইডি:** {product_id}\n"
                f"➕ **যোগ হয়েছে:** {add_qty} পিস\n"
                f"📈 **মোট স্টক:** {new_initial} পিস\n"
                f"📉 **অবশিষ্ট স্টক:** {remaining} পিস\n"
                f"⏰ **আপডেটের সময়:** {current_time}"
            )
            await update.message.reply_text(reply_msg, parse_mode="Markdown")
        else:
            await update.message.reply_text("❌ এই প্রোডাক্ট আইডি পাওয়া যায়নি।")
            
    except Exception as e:
        await update.message.reply_text(f"⚠️ ত্রুটি ঘটেছে: {str(e)}")

async def sell_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """বিক্রি এন্ট্রি (তারিখ ও সময়সহ)"""
    try:
        args = context.args
        if len(args) < 2:
            await update.message.reply_text("❌ ভুল ফরম্যাট! সঠিক নিয়মে লিখুন: `/sell 101 5`", parse_mode="Markdown")
            return
            
        product_id = int(args[0])
        sold_qty = int(args[1])
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
                    await update.message.reply_text(f"⚠️ অপর্যাপ্ত স্টক! '{product_name}' এর মোট স্টক আছে {initial} পিস, কিন্তু বিক্রি দেখাতে চাচ্ছেন {new_sold} পিস।")
                    return
                
                ws.cell(row=row, column=4).value = new_sold
                ws.cell(row=row, column=6).value = current_time
                
                remaining = initial - new_sold
                found = True
                break
                
        if found:
            wb.save(EXCEL_FILE)
            reply_msg = (
                f"🛒 **বিক্রি সফলভাবে এন্ট্রি হয়েছে (Sales)!**\n"
                f"🏷️ **প্রোডাক্টের নাম:** {product_name}\n"
                f"🆔 **আইডি:** {product_id}\n"
                f"➖ **বিক্রীত পরিমাণ:** {sold_qty} পিস\n"
                f"📉 **অবশিষ্ট স্টক:** {remaining} পিস\n"
                f"⏰ **বিক্রির সময়:** {current_time}"
            )
            await update.message.reply_text(reply_msg, parse_mode="Markdown")
        else:
            await update.message.reply_text("❌ এই প্রোডাক্ট আইডি পাওয়া যায়নি।")
            
    except Exception as e:
        await update.message.reply_text(f"⚠️ ত্রুটি ঘটেছে: {str(e)}")

async def get_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """স্টক ও সেলস রিপোর্ট (সর্বশেষ সময়ের উল্লেখ সহ)"""
    try:
        initialize_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws_stock = wb["Daily_Inventory"]
        
        if ws_stock.max_row <= 1:
            await update.message.reply_text("📭 ইনভেন্টরি একদম খালি!")
            return
            
        total_items_sold = 0
        report_summary = "📊 **বর্তমান ইনভেন্টরি ও সেলস রিপোর্ট:**\n\n"
        
        for row in range(2, ws_stock.max_row + 1):
            p_id = ws_stock.cell(row=row, column=1).value
            p_name = ws_stock.cell(row=row, column=2).value
            initial = ws_stock.cell(row=row, column=3).value or 0
            sold = ws_stock.cell(row=row, column=4).value or 0
            rem = initial - sold
            last_time = ws_stock.cell(row=row, column=6).value or "N/A"
            
            total_items_sold += sold
            report_summary += f"🔹 **{p_name}** (ID: {p_id})\n   - মোট স্টক: {initial} | বিক্রি: {sold}পিস | বাকি: {rem}পিস\n   - 🕒 শেষ আপডেট: {last_time}\n\n"
            
        report_summary += f"🔥 **সর্বমোট বিক্রীত পণ্য:** {total_items_sold} পিস"
        await update.message.reply_text(report_summary, parse_mode="Markdown")
        
    except Exception as e:
        await update.message.reply_text(f"⚠️ ত্রুটি ঘটেছে: {str(e)}")

def main():
    initialize_excel()
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("new_product", register_new_product))
    app.add_handler(CommandHandler("new_multiple", register_multiple_products))
    app.add_handler(CommandHandler("add", add_stock))
    app.add_handler(CommandHandler("sell", sell_product))
    app.add_handler(CommandHandler("report", get_report))
    
    print("🤖 টেলিগ্রাম বট চালু হয়েছে (তারিখ ও সময় ফিচারসহ)...")
    app.run_polling()

if __name__ == "__main__":
    main()