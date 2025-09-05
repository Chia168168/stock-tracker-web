from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify, make_response
import pandas as pd
import twstock
import os
from datetime import datetime
import re
import io
import logging

app = Flask(__name__)
app.secret_key = os.urandom(24).hex()  # 使用隨機密鑰

# 設置日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# File to store transactions and stock names
TRANSACTION_FILE = "stock_transactions.csv"
STOCK_NAMES_FILE = "stock_names.csv"

# Initialize CSV file for transactions if it doesn't exist
def initialize_csv():
    if not os.path.exists(TRANSACTION_FILE):
        df = pd.DataFrame(columns=[
            "Date", "Stock_Code", "Stock_Name", "Type", 
            "Quantity", "Price", "Fee", "Tax"
        ])
        df.to_csv(TRANSACTION_FILE, index=False, encoding='utf-8-sig')

# Load stock names from CSV with encoding fallback
from functools import lru_cache

@lru_cache(maxsize=1)
def load_stock_names():
    try:
        df = pd.read_csv(STOCK_NAMES_FILE, encoding='utf-8-sig')
        expected_columns = ["Code", "Name", "Market"]
        if list(df.columns) != expected_columns:
            logger.error(f"{STOCK_NAMES_FILE} 格式錯誤，應包含欄位: {expected_columns}")
            return {}
        stock_names = {(str(row["Code"]), row["Market"]): row["Name"] for _, row in df.iterrows()}
        logger.info(f"成功載入 {len(stock_names)} 個股票名稱")
        return stock_names
    except Exception as e:
        logger.error(f"載入 {STOCK_NAMES_FILE} 失敗: {e}")
        return {}

# Fetch stock info (price and name) using twstock
def fetch_stock_info(code, is_otc=False, retries=3):
    stock_names = load_stock_names()
    ticker = f"{code}.TWO" if is_otc else f"{code}.TW"
    name_key = (str(code), "TWO" if is_otc else "TWSE")
    name = stock_names.get(name_key, "未知名稱")
    try:
        import twstock  # 動態載入 twstock
        stock = twstock.Stock(code)
        price = stock.fetch(2025, 9)[-1].close
        logger.info(f"股票 {ticker} 股價: {price}")
        return {"price": round(price, 2), "name": name}
    except Exception as e:
        logger.error(f"抓取股票 {ticker} 的資訊失敗: {e}")
        return {"price": 0, "name": name}

# Calculate portfolio summary
def get_portfolio_summary():
    df = pd.read_csv(TRANSACTION_FILE, encoding='utf-8-sig')
    if df.empty:
        return [], 0, 0, 0, 0

    summary = {}
    for _, row in df.iterrows():
        code = row["Stock_Code"]
        if code not in summary:
            summary[code] = {
                "name": row["Stock_Name"],
                "quantity": 0,
                "total_cost": 0,
                "buy_quantity": 0,
                "realized_profit": 0,
                "is_otc": row["Stock_Code"].endswith(".TWO")
            }

        if row["Type"] == "Buy":
            summary[code]["quantity"] += row["Quantity"]
            summary[code]["total_cost"] += row["Quantity"] * row["Price"] + row["Fee"] + row["Tax"]
            summary[code]["buy_quantity"] += row["Quantity"]
        else:  # Sell
            summary[code]["quantity"] -= row["Quantity"]
            avg_buy_price = summary[code]["total_cost"] / summary[code]["buy_quantity"] if summary[code]["buy_quantity"] > 0 else 0
            summary[code]["realized_profit"] += (row["Price"] - avg_buy_price) * row["Quantity"] - row["Fee"] - row["Tax"]

    result = []
    total_cost = 0
    total_market_value = 0
    total_unrealized_profit = 0
    total_realized_profit = 0

    for code, data in summary.items():
        if data["quantity"] <= 0:
            continue
        stock_info = fetch_stock_info(code.split(".")[0], data["is_otc"])
        current_price = stock_info["price"]
        market_value = data["quantity"] * current_price
        avg_buy_price = data["total_cost"] / data["buy_quantity"] if data["buy_quantity"] > 0 else 0
        unrealized_profit = (current_price - avg_buy_price) * data["quantity"] if data["quantity"] > 0 else 0
        
        total_cost += data["total_cost"]
        total_market_value += market_value
        total_unrealized_profit += unrealized_profit
        total_realized_profit += data["realized_profit"]

        result.append({
            "Stock_Code": code,
            "Stock_Name": data["name"],
            "Quantity": int(data["quantity"]),
            "Avg_Buy_Price": round(avg_buy_price, 2),
            "Current_Price": round(current_price, 2),
            "Total_Cost": int(data["total_cost"]),
            "Market_Value": int(market_value),
            "Unrealized_Profit": int(unrealized_profit),
            "Realized_Profit": int(data["realized_profit"])
        })

    return result, int(total_cost), int(total_market_value), int(total_unrealized_profit), int(total_realized_profit)

@app.route("/", methods=["GET", "POST"])
def index():
    initialize_csv()
    error = None
    stock_name = None
    default_date = datetime.now().strftime("%Y-%m-%d")

    if request.method == "POST":
        action = request.form.get("action")
        
        if action == "add_transaction":
            try:
                date = request.form.get("date", default_date)
                code = request.form.get("code", "").strip()
                name = request.form.get("name", "").strip() or "未知股票"
                market = request.form.get("market", "TWSE")
                trans_type = request.form.get("type", "Buy")
                quantity = request.form.get("quantity")
                price = request.form.get("price")
                
                # Validate inputs
                if not code:
                    error = "股票代碼不能為空"
                elif not quantity or float(quantity) <= 0:
                    error = "股數必須為正數"
                elif float(quantity) % 1000 != 0:
                    error = "股數必須為1000的倍數"
                elif not price or float(price) <= 0:
                    error = "每股價格必須為正數"
                else:
                    quantity = float(quantity)
                    price = float(price)
                    fee = max(20, price * quantity * 0.001425)
                    tax = price * quantity * 0.003 if trans_type == "Sell" else 0

                    code_with_suffix = f"{code}.TWO" if market == "TWO" else f"{code}.TW"
                    new_transaction = {
                        "Date": date,
                        "Stock_Code": code_with_suffix,
                        "Stock_Name": name,
                        "Type": trans_type,
                        "Quantity": quantity,
                        "Price": price,
                        "Fee": fee,
                        "Tax": tax
                    }
                    df = pd.read_csv(TRANSACTION_FILE, encoding='utf-8-sig')
                    df = pd.concat([df, pd.DataFrame([new_transaction])], ignore_index=True)
                    df.to_csv(TRANSACTION_FILE, index=False, encoding='utf-8-sig')
                    flash("交易已新增！", "success")
                    return redirect(url_for("index"))
            except ValueError as e:
                error = f"輸入無效: {str(e)}。請確保股數和價格為有效數字"

        elif action == "import_transactions":
            file = request.files.get("import_file")
            if file and file.filename.endswith(".csv"):
                try:
                    imported_df = pd.read_csv(file, encoding='utf-8-sig')
                    expected_columns = ["Date", "Stock_Code", "Stock_Name", "Type", "Quantity", "Price", "Fee", "Tax"]
                    if list(imported_df.columns) != expected_columns:
                        error = "匯入檔案格式不正確，需包含正確欄位"
                    else:
                        if not all(imported_df["Quantity"].apply(lambda x: x % 1000 == 0)):
                            error = "匯入檔案中的股數必須為1000的倍數"
                        else:
                            overwrite = request.form.get("overwrite") == "on"
                            if overwrite:
                                imported_df.to_csv(TRANSACTION_FILE, index=False, encoding='utf-8-sig')
                            else:
                                current_df = pd.read_csv(TRANSACTION_FILE, encoding='utf-8-sig')
                                combined_df = pd.concat([current_df, imported_df], ignore_index=True)
                                combined_df.to_csv(TRANSACTION_FILE, index=False, encoding='utf-8-sig')
                            flash("交易紀錄已匯入！", "success")
                except Exception as e:
                    error = f"匯入失敗: {e}"
            else:
                error = "請選擇有效的 CSV 檔案"

    transactions = pd.read_csv(TRANSACTION_FILE, encoding='utf-8-sig').to_dict("records")
    summary, total_cost, total_market_value, total_unrealized_profit, total_realized_profit = get_portfolio_summary()
    
    return render_template(
        "index.html",
        transactions=transactions,
        summary=summary,
        total_cost=total_cost,
        total_market_value=total_market_value,
        total_unrealized_profit=total_unrealized_profit,
        total_realized_profit=total_realized_profit,
        error=error,
        stock_name=stock_name,
        form_data=request.form,
        default_date=default_date
    )

@app.route("/fetch_stock_name", methods=["POST"])
def fetch_stock_name():
    code = request.form.get("code", "").strip()
    market = request.form.get("market", "TWSE")
    logger.info(f"收到查詢請求: 代碼={code}, 市場={market}")
    if not code:
        response = jsonify({"error": "請輸入股票代碼"})
        response.headers["Content-Type"] = "application/json; charset=utf-8"
        return response
    stock_names = load_stock_names()
    name_key = (str(code), market)
    name = stock_names.get(name_key)
    if name:
        logger.info(f"從 stock_names.csv 找到名稱: {name}")
        response = jsonify({"name": name, "is_english": False})
        response.headers["Content-Type"] = "application/json; charset=utf-8"
        return response
    is_otc = market == "TWO"
    stock_info = fetch_stock_info(code, is_otc)
    if not stock_info["name"]:
        logger.error(f"無法抓取股票 {code} 的名稱")
        response = jsonify({"error": f"無法抓取股票 {code} 的名稱，請檢查代碼或市場選擇，或手動輸入名稱"})
        response.headers["Content-Type"] = "application/json; charset=utf-8"
        return response
    logger.info(f"返回股票名稱: {stock_info['name']}")
    response = jsonify({"name": stock_info["name"], "is_english": not re.search(r'[\u4e00-\u9fff]', stock_info["name"])})
    response.headers["Content-Type"] = "application/json; charset=utf-8"
    return response

@app.route("/export_transactions")
def export_transactions():
    try:
        df = pd.read_csv(TRANSACTION_FILE, encoding='utf-8-sig')
        output = io.StringIO()
        df.to_csv(output, index=False, encoding='utf-8-sig')
        output.seek(0)
        return send_file(
            io.BytesIO(output.getvalue().encode("utf-8-sig")),
            mimetype="text/csv; charset=utf-8",
            as_attachment=True,
            download_name=f"exported_transactions_{datetime.now().strftime('%Y%m%d')}.csv"
        )
    except Exception as e:
        flash(f"匯出失敗: {e}", "error")
        return redirect(url_for("index"))

if __name__ == "__main__":
    app.run()