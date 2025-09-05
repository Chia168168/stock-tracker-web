from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify, make_response
import pandas as pd
import yfinance as yf
import os
from datetime import datetime, time
import re
import io
import logging
import json
import gspread
from google.oauth2.service_account import Credentials
import threading
import time

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev_secret_key")  # 使用環境變數

# 設置日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 使用絕對路徑來存儲文件
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TRANSACTION_FILE = os.path.join(BASE_DIR, "stock_transactions.csv")
STOCK_NAMES_FILE = os.path.join(BASE_DIR, "stock_names.csv")

# Google Sheets 設置
def setup_google_sheets():
    try:
        # 從環境變量獲取憑證（在 Render 上設置）
        creds_json = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')
        if not creds_json:
            logger.error("未找到 Google Sheets 憑證環境變量")
            return None
        
        # 解析憑證
        scope = ['https://spreadsheets.google.com/feeds', 
                'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(
            json.loads(creds_json), scopes=scope)
        
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        logger.error(f"設置 Google Sheets 時出錯: {e}")
        return None

# 從 Google Sheets 讀取股票價格
def get_prices_from_google_sheet(client, sheet_name, worksheet_name="Sheet1"):
    try:
        # 打開試算表
        sheet = client.open(sheet_name).worksheet(worksheet_name)
        
        # 讀取所有數據
        data = sheet.get_all_records()
        
        # 轉換為字典格式：{股票代碼: 價格}
        prices = {}
        for row in data:
            if 'code' in row and 'price' in row:
                prices[str(row['code'])] = float(row['price'])
        
        return prices
    except Exception as e:
        logger.error(f"從 Google Sheets 讀取數據時出錯: {e}")
        return {}

# 在應用啟動時初始化 Google Sheets 連接
def init_google_sheets():
    try:
        client = setup_google_sheets()
        if client:
            # 從環境變量獲取試算表名稱
            sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票價格')
            worksheet_name = os.environ.get('GOOGLE_WORKSHEET_NAME', 'Sheet1')
            
            # 讀取價格數據
            prices = get_prices_from_google_sheet(client, sheet_name, worksheet_name)
            
            # 將價格數據存儲在函數屬性中
            fetch_stock_info.google_sheets_prices = prices
            logger.info(f"從 Google Sheets 成功讀取 {len(prices)} 個股票價格")
        else:
            logger.warning("無法初始化 Google Sheets 連接")
            fetch_stock_info.google_sheets_prices = {}
    except Exception as e:
        logger.error(f"初始化 Google Sheets 時出錯: {e}")
        fetch_stock_info.google_sheets_prices = {}

# 定期更新 Google Sheets 數據
def schedule_google_sheets_update(interval_minutes=15):
    def update():
        while True:
            try:
                time.sleep(interval_minutes * 60)
                init_google_sheets()
            except Exception as e:
                logger.error(f"定期更新 Google Sheets 數據時出錯: {e}")
    
    # 啟動後台線程
    thread = threading.Thread(target=update)
    thread.daemon = True
    thread.start()

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
            logger.error(f"實際欄位: {list(df.columns)}")
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

# Fetch stock info - 使用 Google Sheets 數據
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
    
    # 嘗試從 Google Sheets 獲取價格
    price = 0
    if hasattr(fetch_stock_info, 'google_sheets_prices'):
        price = fetch_stock_info.google_sheets_prices.get(str(code), 0)
    
    # 如果 Google Sheets 沒有數據，嘗試從 Yahoo Finance 獲取
    if price == 0:
        try:
            ticker = f"{code}.TWO" if is_otc else f"{code}.TW"
            stock = yf.Ticker(ticker)
            history = stock.history(period="1d")
            if not history.empty:
                price = history["Close"].iloc[-1]
        except Exception as e:
            logger.error(f"從 Yahoo Finance 獲取股票 {ticker} 價格失敗: {e}")
    
    result = {"price": round(price, 2), "name": name}
    
    # 更新緩存
    if not hasattr(fetch_stock_info, 'cache'):
        fetch_stock_info.cache = {}
    fetch_stock_info.cache[cache_key] = {
        'timestamp': current_time,
        'data': result
    }
    
    return result

# 添加函數來讀取和寫入交易數據到 Google Sheets
def get_transactions_from_google_sheet(client, sheet_name, worksheet_name="交易紀錄"):
    try:
        # 打開試算表
        sheet = client.open(sheet_name).worksheet(worksheet_name)
        
        # 讀取所有數據
        data = sheet.get_all_records()
        
        # 轉換為與原來相同的格式
        transactions = []
        for row in data:
            transactions.append({
                "Date": row.get("Date", ""),
                "Stock_Code": row.get("Stock_Code", ""),
                "Stock_Name": row.get("Stock_Name", ""),
                "Type": row.get("Type", ""),
                "Quantity": float(row.get("Quantity", 0)),
                "Price": float(row.get("Price", 0)),
                "Fee": float(row.get("Fee", 0)),
                "Tax": float(row.get("Tax", 0))
            })
        
        return transactions
    except Exception as e:
        logger.error(f"從 Google Sheets 讀取交易數據時出錯: {e}")
        return []

def add_transaction_to_google_sheet(client, sheet_name, worksheet_name, transaction):
    try:
        # 打開試算表
        sheet = client.open(sheet_name).worksheet(worksheet_name)
        
        # 獲取現有數據以確定新行的位置
        existing_data = sheet.get_all_values()
        next_row = len(existing_data) + 1 if existing_data else 2  # 標題行佔用第1行
        
        # 添加新交易
        sheet.append_row([
            transaction["Date"],
            transaction["Stock_Code"],
            transaction["Stock_Name"],
            transaction["Type"],
            transaction["Quantity"],
            transaction["Price"],
            transaction["Fee"],
            transaction["Tax"]
        ])
        
        return True
    except Exception as e:
        logger.error(f"添加交易到 Google Sheets 時出錯: {e}")
        return False

# 修改 initialize_csv 函數以初始化 Google Sheets
def initialize_google_sheets():
    try:
        client = setup_google_sheets()
        if client:
            # 從環境變量獲取試算表名稱
            sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
            
            # 檢查交易紀錄工作表是否存在，如果不存在則創建
            try:
                sheet = client.open(sheet_name).worksheet("交易紀錄")
            except gspread.exceptions.WorksheetNotFound:
                # 創建交易紀錄工作表
                sheet = client.open(sheet_name).add_worksheet(title="交易紀錄", rows=1000, cols=20)
                # 添加標題行
                sheet.append_row(["Date", "Stock_Code", "Stock_Name", "Type", "Quantity", "Price", "Fee", "Tax"])
            
            return True
        return False
    except Exception as e:
        logger.error(f"初始化 Google Sheets 時出錯: {e}")
        return False

# 修改 get_portfolio_summary 函數以使用 Google Sheets 數據
def get_portfolio_summary():
    try:
        client = setup_google_sheets()
        if client:
            sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
            transactions = get_transactions_from_google_sheet(client, sheet_name, "交易紀錄")
        else:
            transactions = []
    except:
        transactions = []
        
    if not transactions:
        return [], 0, 0, 0, 0

    summary = {}
    for row in transactions:
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
    initialize_google_sheets()
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
                    
                    # 添加到 Google Sheets
                    client = setup_google_sheets()
                    if client:
                        sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
                        if add_transaction_to_google_sheet(client, sheet_name, "交易紀錄", new_transaction):
                            flash("交易已新增！", "success")
                        else:
                            error = "無法將交易添加到 Google Sheets"
                    else:
                        error = "無法連接到 Google Sheets"
                    
                    return redirect(url_for("index"))
            except ValueError as e:
                error = f"輸入無效: {str(e)}。請確保股數和價格為有效數字"

        elif action == "import_transactions":
            # 這裡可以實現從 CSV 導入到 Google Sheets 的功能
            flash("導入功能暫不可用，請直接使用 Google Sheets 管理數據", "warning")
            
        elif action == "update_prices":
            # 手動更新價格
            for key, value in request.form.items():
                if key.startswith("price_"):
                    stock_code = key.replace("price_", "")
                    try:
                        new_price = float(value)
                        if new_price > 0:
                            # 更新緩存中的價格
                            if hasattr(fetch_stock_info, 'cache'):
                                for cache_key in list(fetch_stock_info.cache.keys()):
                                    if stock_code in cache_key:
                                        fetch_stock_info.cache[cache_key]['data']['price'] = new_price
                                        fetch_stock_info.cache[cache_key]['timestamp'] = datetime.now().timestamp()
                                        flash(f"已更新 {stock_code} 的價格為 {new_price}", "success")
                    except ValueError:
                        pass

    # 獲取交易數據
    try:
        client = setup_google_sheets()
        if client:
            sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
            transactions = get_transactions_from_google_sheet(client, sheet_name, "交易紀錄")
        else:
            transactions = []
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
    
    # 直接從本地 CSV 獲取股票名稱
    stock_names = load_stock_names()
    
    # 根據市場選擇正確的鍵
    if market == "TWO":
        market_key = "TWO"
    else:
        market_key = "TWSE"
    
    name_key = (str(code), market_key)
    logger.info(f"查找的鍵: {name_key}")
    
    # 記錄所有可用的鍵以便調試
    available_keys = list(stock_names.keys())
    logger.info(f"可用的鍵數量: {len(available_keys)}")
    if available_keys:
        logger.info(f"前幾個可用鍵: {available_keys[:5]}")
    
    name = stock_names.get(name_key, "")
    
    if not name:
        # 嘗試不區分市場查找
        for key, value in stock_names.items():
            if key[0] == str(code):
                name = value
                logger.info(f"找到不區分市場的名稱: {name}")
                break
        
        if not name:
            logger.error(f"無法找到股票 {code} 的名稱，查找的鍵: {name_key}")
            # 嘗試從 Google Sheets 獲取名稱
            try:
                if hasattr(fetch_stock_info, 'google_sheets_prices'):
                    # 假設 Google Sheets 中有名稱數據
                    # 這裡需要根據您的實際數據結構進行調整
                    pass
            except:
                pass
            
            response = jsonify({"error": f"無法找到股票 {code} 的名稱，請手動輸入名稱"})
            response.headers["Content-Type"] = "application/json; charset=utf-8"
            return response
    
    logger.info(f"返回股票名稱: {name}")
    response = jsonify({"name": name, "is_english": not re.search(r'[\u4e00-\u9fff]', name)})
    response.headers["Content-Type"] = "application/json; charset=utf-8"
    return response
@app.route("/export_transactions")
def export_transactions():
    try:
        client = setup_google_sheets()
        if client:
            sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
            transactions = get_transactions_from_google_sheet(client, sheet_name, "交易紀錄")
            
            # 轉換為 DataFrame 並導出為 CSV
            df = pd.DataFrame(transactions)
            output = io.StringIO()
            df.to_csv(output, index=False, encoding='utf-8-sig')
            output.seek(0)
            
            return send_file(
                io.BytesIO(output.getvalue().encode("utf-8-sig")),
                mimetype="text/csv; charset=utf-8",
                as_attachment=True,
                download_name=f"exported_transactions_{datetime.now().strftime('%Y%m%d')}.csv"
            )
        else:
            flash("無法連接到 Google Sheets", "error")
            return redirect(url_for("index"))
    except Exception as e:
        flash(f"匯出失敗: {e}", "error")
        return redirect(url_for("index"))
# 初始化 Google Sheets
init_google_sheets()
# 啟動定期更新
schedule_google_sheets_update(15)  # 每15分鐘更新一次

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
