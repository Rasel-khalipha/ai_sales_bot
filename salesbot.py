import os
import re
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler
from pymongo import MongoClient

# --- MongoDB Atlas কানেকশন স্ট্রিং ---
MONGO_URI = os.environ.get("MONGO_URI", "mongodb+srv://mraselkhalifa_db_user:nUtzQOjFn0lXma7P@cluster0.3zchxnh.mongodb.net/salesbot_database?retryWrites=true&w=majority")
client = MongoClient(MONGO_URI)

db = client["salesbot_database"]
inventory_collection = db["inventory"]
sales_collection = db["sales_history"]

TELEGRAM_TOKEN = "9104364470:AAFQxiPZuYBRvMgtpBDPQeSkzxpUZ31o_pU"  # আপনার টেলিগ্রাম বট টোকেন এখানে দিন

# --- ডামি সার্ভার (রেন্ডারের জন্য) ---
class SimpleHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Bot is alive!")

def run_dummy_server():
    port = int(os.environ.get("PORT", 10000))
    server = HTTPServer(('0.0.0.0', port), SimpleHandler)
    server.serve_forever()

# --- হেল্পার ও লজিক ফাংশনসমূহ ---

def extract_number_and_unit(text):
    text_str = str(text).strip().lower()
    numbers = re.findall(r'\d+', text_str)
    qty = int(numbers[0]) if numbers else 0
    
    unit = "box"
    if "ph" in text_str:
        unit = "ph"
    elif "pc" in text_str or "pcs" in text_str:
        unit = "pcs"
    elif "box" in text_str:
        unit = "box"
        
    return qty, unit

def extract_number(text):
    numbers = re.findall(r'\d+', str(text))
    return int(numbers[0]) if numbers else 0

def format_2digit(val):
    try:
        return f"{int(val):02d}"
    except (ValueError, TypeError):
        return str(val)

def parse_product_args(args):
    unit_price = extract_number(args[-1])
    if args[-2].lower() in ["box", "ph", "pc", "pcs"] and len(args) >= 5:
        qty_unit_str = f"{args[-3]} {args[-2]}"
        name_end_idx = -3
    else:
        qty_unit_str = args[-2]
        name_end_idx = -2
        
    product_name = " ".join(args[1:name_end_idx]).strip()
    stock_qty, unit_type = extract_number_and_unit(qty_unit_str)
    return product_name, stock_qty, unit_type, unit_price

# --- Telegram Command Handlers ---

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        "🤖 **Smart Inventory & MongoDB Assistant**\n\n"
        "🆕 **Register Product:** `/new_product [ID] [Name] [Stock+Unit] [Price]`\n"
        "📦 **Multiple Products:** `/new_multiple 01, Shirt, 50 pcs, 500; 02, Pant, 30 box, 800`\n"
        "📥 **Add Stock:** `/add [ID] [Qty+Unit]`\n"
        "🛒 **Sell Product:** `/sell [ID] [Qty+Unit]`\n"
        "✏️ **Edit Product:** `/edit [ID] [Name] [Stock+Unit] [Price]`\n"
        "🗑️ **Delete Product:** `/delete [ID]`\n"
        "📊 **View Report:** `/report`"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")

async def register_new_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        args = context.args
        if len(args) < 4:
            await update.message.reply_text("❌ Format: `/new_product [ID] [Name] [Stock+Unit] [Price]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        product_name, initial_stock, unit_type, unit_price = parse_product_args(args)
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        existing = inventory_collection.find_one({"product_id": product_id})
        if existing:
            await update.message.reply_text(f"⚠️ Product ID ({format_2digit(product_id)}) already exists!")
            return
        
        product_data = {
            "product_id": product_id,
            "product_name": product_name,
            "initial_stock": initial_stock,
            "daily_sold": 0,
            "unit_type": unit_type.upper(),
            "unit_price": unit_price,
            "last_updated": current_time
        }
        inventory_collection.insert_one(product_data)
        
        await update.message.reply_text(
            f"✅ **Product Added to Cloud Successfully!**\n"
            f"🆔 ID: {format_2digit(product_id)} | 🏷️ Name: {product_name}\n"
            f"📦 Stock: {format_2digit(initial_stock)} {unit_type.upper()} | 💵 Unit Price: {unit_price} BDT",
            parse_mode="Markdown"
        )
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

# --- নতুন যুক্ত হওয়া মাল্টিপল প্রোডাক্ট হ্যান্ডলার ---
async def register_multiple_products(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        full_text = update.message.text
        # /new_multiple অংশটি বাদ দিয়ে বাকি টেক্সট নেওয়া
        if " " not in full_text:
            await update.message.reply_text("❌ Format: `/new_multiple ID, Name, Stock Unit, Price; ID, Name, Stock Unit, Price`", parse_mode="Markdown")
            return
            
        data_string = full_text.split(" ", 1)[1]
        items = data_string.split(";")
        
        success_count = 0
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        for item in items:
            item = item.strip()
            if not item:
                continue
            parts = [p.strip() for p in item.split(",")]
            if len(parts) < 4:
                continue # ফরম্যাট ভুল থাকলে স্কিপ করবে
                
            product_id = extract_number(parts[0])
            product_name = parts[1]
            stock_qty, unit_type = extract_number_and_unit(parts[2])
            unit_price = extract_number(parts[3])
            
            # ডাটাবেসে ইতিমধ্যে থাকলে আপডেট বা স্কিপ, না থাকলে ইনসার্ট
            inventory_collection.update_one(
                {"product_id": product_id},
                {
                    "$set": {
                        "product_name": product_name,
                        "initial_stock": stock_qty,
                        "unit_type": unit_type.upper(),
                        "unit_price": unit_price,
                        "last_updated": current_time
                    },
                    "$setOnInsert": {
                        "daily_sold": 0
                    }
                },
                upsert=True
            )
            success_count += 1
            
        await update.message.reply_text(f"✅ Successfully added/updated {success_count} products in MongoDB Atlas!")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error in multiple insert: {str(e)}")

async def edit_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        args = context.args
        if len(args) < 4:
            await update.message.reply_text("❌ Format: `/edit [ID] [New_Name] [New_Stock+Unit] [New_Price]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        new_name, new_stock, unit_type, new_price = parse_product_args(args)
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        result = inventory_collection.update_one(
            {"product_id": product_id},
            {
                "$set": {
                    "product_name": new_name,
                    "initial_stock": new_stock,
                    "unit_type": unit_type.upper(),
                    "unit_price": new_price,
                    "last_updated": current_time
                }
            }
        )
        
        if result.matched_count > 0:
            await update.message.reply_text(
                f"✏️ **Product Updated Successfully!**\n"
                f"🆔 ID: {format_2digit(product_id)} | 🏷️ Name: {new_name}\n"
                f"📦 New Stock: {format_2digit(new_stock)} {unit_type.upper()} | 💵 Price: {new_price} BDT",
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text(f"❌ Product ID {format_2digit(product_id)} not found.")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

async def delete_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        args = context.args
        if not args:
            await update.message.reply_text("❌ Format: `/delete [ID]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        result = inventory_collection.find_one_and_delete({"product_id": product_id})
        
        if result:
            await update.message.reply_text(
                f"🗑️ **Product Deleted Successfully!**\n"
                f"🆔 ID: {format_2digit(product_id)} | 🏷️ Name: {result['product_name']}",
                parse_mode="Markdown"
            )
        else:
            await update.message.reply_text(f"❌ Product ID {format_2digit(product_id)} not found.")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

async def add_stock(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        args = context.args
        if len(args) < 2:
            await update.message.reply_text("❌ Format: `/add [ID] [Quantity+Unit]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        add_qty, _ = extract_number_and_unit(" ".join(args[1:]))
        current_time = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        product = inventory_collection.find_one({"product_id": product_id})
        if product:
            new_init = product.get("initial_stock", 0) + add_qty
            inventory_collection.update_one(
                {"product_id": product_id},
                {"$set": {"initial_stock": new_init, "last_updated": current_time}}
            )
            await update.message.reply_text(
                f"📥 **Stock Added!**\n"
                f"🆔 ID: {format_2digit(product_id)} | 🏷️ Name: {product['product_name']}\n"
                f"➕ Added: {format_2digit(add_qty)} {product['unit_type']} | 📈 New Initial: {format_2digit(new_init)} {product['unit_type']}"
            )
        else:
            await update.message.reply_text(f"❌ Product ID {format_2digit(product_id)} not found.")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

async def sell_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        args = context.args
        if len(args) < 2:
            await update.message.reply_text("❌ Format: `/sell [ID] [Quantity+Unit]`", parse_mode="Markdown")
            return
            
        product_id = extract_number(args[0])
        sold_qty, _ = extract_number_and_unit(" ".join(args[1:]))
        current_time_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        product = inventory_collection.find_one({"product_id": product_id})
        if product:
            initial = product.get("initial_stock", 0)
            curr_sold = product.get("daily_sold", 0)
            unit_type = product.get("unit_type", "BOX")
            u_price = product.get("unit_price", 0)
            
            new_sold = curr_sold + sold_qty
            remaining = initial - new_sold
            if remaining < 0:
                await update.message.reply_text(f"⚠️ Insufficient stock! '{product['product_name']}' has max {format_2digit(initial - curr_sold)} {unit_type} available.")
                return
            
            inventory_collection.update_one(
                {"product_id": product_id},
                {"$set": {"daily_sold": new_sold, "last_updated": current_time_str}}
            )
            
            total_sold_value = sold_qty * u_price
            sales_record = {
                "product_id": product_id,
                "product_name": product['product_name'],
                "sold_quantity": sold_qty,
                "unit_type": unit_type,
                "unit_price": u_price,
                "total_sold_value": total_sold_value,
                "date_time": datetime.now()
            }
            sales_collection.insert_one(sales_record)
            
            await update.message.reply_text(
                f"🛒 **Sale Recorded!**\n"
                f"🆔 ID: {format_2digit(product_id)} | 🏷️ Name: {product['product_name']}\n"
                f"➖ Sold: {format_2digit(sold_qty)} {unit_type} | 💵 Total Value: {total_sold_value} BDT"
            )
        else:
            await update.message.reply_text(f"❌ Product ID {format_2digit(product_id)} not found.")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

async def get_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        products = list(inventory_collection.find().sort("product_id", 1))
        if not products:
            await update.message.reply_text("📭 Inventory is empty!")
            return
            
        now = datetime.now()
        today_str = now.strftime("%Y-%m-%d")
        
        daily_sales_qty, daily_sales_val = 0, 0
        all_sales = sales_collection.find()
        for sale in all_sales:
            s_date = sale.get("date_time")
            if isinstance(s_date, datetime) and s_date.strftime("%Y-%m-%d") == today_str:
                daily_sales_qty += sale.get("sold_quantity", 0)
                daily_sales_val += sale.get("total_sold_value", 0)

        grand_total_inventory_value = 0
        report_msg = "📊 **Inventory & Sales Report Summary (Cloud)** 📊\n\n"
        report_msg += f"☀️ **Today's Sales:** {format_2digit(daily_sales_qty)} units (Value: {daily_sales_val} BDT)\n\n"
        report_msg += "📦 **Current Stock & Value Status:**\n"
        
        for p in products:
            p_id = p["product_id"]
            p_name = p["product_name"]
            initial = p["initial_stock"]
            sold = p["daily_sold"]
            rem = initial - sold
            unit_type = p["unit_type"]
            price = p["unit_price"]
            tot_val = rem * price
            grand_total_inventory_value += tot_val
            
            report_msg += f"• ID {format_2digit(p_id)} | **{p_name}**: Rem {format_2digit(rem)} {unit_type} | Price: {price} | Val: {tot_val} BDT\n"
            
        report_msg += f"\n💎 **Total Inventory Remaining Value:** **{grand_total_inventory_value} BDT**"
        await update.message.reply_text(report_msg, parse_mode="Markdown")
    except Exception as e:
        await update.message.reply_text(f"⚠️ Error: {str(e)}")

def main():
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("new_product", register_new_product))
    app.add_handler(CommandHandler("new_multiple", register_multiple_products)) # নতুন হ্যান্ডলার যোগ করা হলো
    app.add_handler(CommandHandler("edit", edit_product))
    app.add_handler(CommandHandler("delete", delete_product))
    app.add_handler(CommandHandler("add", add_stock))
    app.add_handler(CommandHandler("sell", sell_product))
    app.add_handler(CommandHandler("report", get_report))
    
    print("🤖 Telegram Bot running with MongoDB Atlas & Multiple Insert Support...")
    app.run_polling()

if __name__ == "__main__":
    server_thread = threading.Thread(target=run_dummy_server)
    server_thread.daemon = True
    server_thread.start()
    
    main()