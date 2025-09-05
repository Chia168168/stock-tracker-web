from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify, make_response
import pandas as pd
import yfinance as yf
import os
from datetime import datetime
import re
import io
import logging

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev_secret_key")  # 使用環境變數

# 設置日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 使用絕對路徑來存儲文件
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TRANSACTION_FILE = os.path.join(BASE_DIR, "stock_transactions.csv")
STOCK_NAMES_FILE = os.path.join(BASE_DIR, "stock_names.csv")

# Initialize CSV file for transactions if it doesn't exist
def initialize_csv():
    if not os.path.exists(TRANSACTION_FILE):
        df = pd.DataFrame(columns=[
            "Date", "Stock_Code", "Stock_Name", "Type", 
            "Quantity", "Price", "Fee", "Tax"
        ])
        df.to_csv(TRANSACTION_FILE, index=False, encoding='utf-8-sig')

# Load stock names from CSV with encoding fallback
def load_stock_names():
    try:
        if not os.path.exists(STOCK_NAMES_FILE):
            logger.warning(f"{STOCK_NAMES_FILE} 不存在，使用空映射")
            return {}
        try:
            df = pd.read_csv(STOCK_NAMES_FILE, encoding='utf-8-sig')
        except UnicodeDecodeError:
            logger.warning("無法以 utf-8-sig 編碼讀取 stock_names.csv，嘗試 big5")
            try:
                df = pd.read_csv(STOCK_NAMES_FILE, encoding='big5')
            except UnicodeDecodeError:
                logger.error("無法以 utf-8-sig 或 big5 編碼讀取 stock_names.csv，請檢查檔案編碼")
                return {}
        expected_columns = ["Code", "Name", "Market"]
        if list(df.columns) != expected_columns:
            logger.error(f"{STOCK_NAMES_FILE} 格式錯誤，應包含欄位: {expected_columns}")
            return {}
        stock_names = {}
        for _, row in df.iterrows():
            try:
                code = str(row["Code"])
                market = row["Market"]
                name = row["Name"]
                stock_names[(code, market)] = name
                logger.debug(f"股票映射: 代碼={code}, 市場={market}, 名稱={name}")
            except Exception as e:
                logger.warning(f"跳過無效行: {row.to_dict()}, 錯誤: {e}")
        logger.info(f"成功載入 {len(stock_names)} 個股票名稱")
        return stock_names
    except Exception as e:
        logger.error(f"載入 {STOCK_NAMES_FILE} 失敗: {e}")
        return {}

# Fetch stock info (price and name) using yfinance with retry
# Fetch stock info (price and name) using yfinance with retry and caching
# Fetch stock info (price and name) using yfinance with retry and caching
import requests
from datetime import datetime, time

# Fetch stock info using Taiwan Stock Exchange and OTC APIs
def fetch_stock_info(code, is_otc=False):
    # 使用緩存來減少 API 請求
    cache_key = f"{code}_{'TWO' if is_otc else 'TW'}"
    current_time = datetime.now().timestamp()
    
    # 檢查緩存是否存在且未過期（5分鐘）
    if hasattr(fetch_stock_info, 'cache'):
        cached_data = fetch_stock_info.cache.get(cache_key)
        if cached_data and current_time - cached_data['timestamp'] < 300:  # 5分鐘緩存
            logger.info(f"使用緩存的股票數據: {cache_key}")
            return cached_data['data']
    
    # 從本地 CSV 獲取股票名稱
    stock_names = load_stock_names()
    name_key = (str(code), "TWO" if is_otc else "TWSE")
    name = stock_names.get(name_key, "未知名稱")
    
    price = 0
    
    # 檢查是否在交易時間內（台灣時間 9:00-13:30）
    try:
        now_utc = datetime.utcnow()
        # 轉換為台灣時間 (UTC+8)
        now_tw = now_utc.replace(hour=now_utc.hour + 8)
        if now_tw.hour > 24:
            now_tw = now_tw.replace(day=now_tw.day + 1, hour=now_tw.hour - 24)
        
        # 檢查是否為工作日（週一至週五）
        is_weekday = now_tw.weekday() < 5
        
        # 檢查是否在交易時間內（9:00-13:30）
        market_open = time(9, 0)
        market_close = time(13, 30)
        is_market_hours = market_open <= now_tw.time() <= market_close
        
        if is_weekday and is_market_hours:
            # 在交易時間內，嘗試從台灣官方 API 獲取實時價格
            if is_otc:
                # 上櫃股票
                price = get_otc_stock_price(code)
            else:
                # 上市股票
                price = get_twse_stock_price(code)
        else:
            # 非交易時間，嘗試從 Yahoo Finance 獲取收盤價
            try:
                ticker = f"{code}.TWO" if is_otc else f"{code}.TW"
                stock = yf.Ticker(ticker)
                history = stock.history(period="1d")
                if not history.empty:
                    price = history["Close"].iloc[-1]
            except:
                # 如果 Yahoo Finance 也失敗，使用 0
                price = 0
    except Exception as e:
        logger.error(f"獲取股票價格時出錯: {e}")
        price = 0
    
    result = {"price": round(price, 2), "name": name}
    
    # 更新緩存
    if not hasattr(fetch_stock_info, 'cache'):
        fetch_stock_info.cache = {}
    fetch_stock_info.cache[cache_key] = {
        'timestamp': current_time,
        'data': result
    }
    
    return result

# 從台灣證交所獲取上市股票價格
def get_twse_stock_price(code):
    try:
        url = f"https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_{code}.tw&json=1&delay=0"
        response = requests.get(url, timeout=5)
        data = response.json()
        
        if data['msgArray']:
            stock_info = data['msgArray'][0]
            if 'z' in stock_info and stock_info['z']:
                return float(stock_info['z'])
    except Exception as e:
        logger.error(f"從證交所獲取股票 {code} 價格失敗: {e}")
    
    return 0

# 從櫃買中心獲取上櫃股票價格
def get_otc_stock_price(code):
    try:
        url = f"https://www.tpex.org.tw/web/stock/aftertrading/daily_trading_info/st43_result.php?l=zh-tw&d=20250905&stkno={code}"
        response = requests.get(url, timeout=5)
        data = response.json()
        
        if data['aaData']:
            # 獲取最新交易日的數據
            latest_data = data['aaData'][0]
            # 收盤價通常在索引 6
            if len(latest_data) > 6 and latest_data[6]:
                return float(latest_data[6].replace(',', ''))
    except Exception as e:
        logger.error(f"從櫃買中心獲取股票 {code} 價格失敗: {e}")
    
    return 0

# Calculate portfolio summary
# Calculate portfolio summary
def get_portfolio_summary():
    try:
        df = pd.read_csv(TRANSACTION_FILE, encoding='utf-8-sig')
    except:
        return [], 0, 0, 0, 0
        
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

    # 批量獲取股票信息，減少 API 調用
    stock_codes_to_fetch = []
    for code, data in summary.items():
        if data["quantity"] > 0:
            stock_codes_to_fetch.append((code.split(".")[0], data["is_otc"]))
    
    # 預先獲取所有需要的股票信息
    stock_info_cache = {}
    for code, is_otc in stock_codes_to_fetch:
        stock_info_cache[(code, is_otc)] = fetch_stock_info(code, is_otc)

    for code, data in summary.items():
        if data["quantity"] <= 0:
            continue
        
        base_code = code.split(".")[0]
        stock_info = stock_info_cache.get((base_code, data["is_otc"]), {"price": 0, "name": data["name"]})
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
            "Stock_Name": stock_info["name"] or data["name"],
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
                    # 自動計算手續費和交易稅
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

    try:
        transactions = pd.read_csv(TRANSACTION_FILE, encoding='utf-8-sig').to_dict("records")
    except:
        transactions = []
    
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
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)