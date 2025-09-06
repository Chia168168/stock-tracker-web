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

# 全局緩存變量
TRANSACTIONS_CACHE = None
TRANSACTIONS_CACHE_TIME = None
TRANSACTIONS_CACHE_DURATION = 300  # 5分鐘緩存

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
def get_prices_from_google_sheet(client, sheet_name, worksheet_name="stock_names"):
    try:
        # 打開試算表
        sheet = client.open(sheet_name).worksheet(worksheet_name)
        
        # 讀取所有數據
        data = sheet.get_all_records()
        
        # 轉換為字典格式：{股票代碼: 價格}
        prices = {}
        for row in data:
            if 'code' in row and 'price' in row:
                try:
                    # 嘗試轉換為浮點數，如果失敗則跳過
                    price_value = float(row['price'])
                    prices[str(row['code'])] = price_value
                except (ValueError, TypeError):
                    continue
        
        # 設置全局緩存
        if hasattr(fetch_stock_info, 'google_sheets_prices'):
            fetch_stock_info.google_sheets_prices = prices
        else:
            fetch_stock_info.google_sheets_prices = prices
            
        return prices
    except Exception as e:
        logger.error(f"從 Google Sheets 讀取數據時出錯: {e}")
        return {}

# 從 Google Sheets 讀取交易數據
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

# 添加交易到 Google Sheets
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
        
        # 如果是買入交易，檢查並添加股票到 stock_names 工作表
        if transaction["Type"] == "Buy":
            # 提取股票代碼（去掉 .TW 或 .TWO 後綴）
            stock_code = transaction["Stock_Code"].split('.')[0]
            
            # 檢查股票是否已存在於 stock_names 工作表
            if not check_stock_exists_in_names(client, sheet_name, stock_code):
                # 添加新股票到 stock_names 工作表
                add_stock_to_names_sheet(client, sheet_name, stock_code, transaction["Stock_Name"])
        
        return True
    except Exception as e:
        logger.error(f"添加交易到 Google Sheets 時出錯: {e}")
        return False
        
# 在應用啟動時初始化 Google Sheets 連接
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
            
            # 檢查 stock_names 工作表是否存在，如果不存在則創建
            try:
                stock_names_sheet = client.open(sheet_name).worksheet("stock_names")
            except gspread.exceptions.WorksheetNotFound:
                # 創建 stock_names 工作表
                stock_names_sheet = client.open(sheet_name).add_worksheet(title="stock_names", rows=1000, cols=10)
                # 添加標題行
                stock_names_sheet.append_row(["code", "name", "price", "pricenow"])
            
            return True
        return False
    except Exception as e:
        logger.error(f"初始化 Google Sheets 時出錯: {e}")
        return False

# 定期更新 Google Sheets 數據
def schedule_google_sheets_update(interval_minutes=30):  # 改為每30分鐘更新一次
    def update():
        while True:
            try:
                time.sleep(interval_minutes * 60)
                initialize_google_sheets()  # 修正函數名稱
            except Exception as e:
                logger.error(f"定期更新 Google Sheets 數據時出錯: {e}")
    
    # 啟動後台線程
    thread = threading.Thread(target=update)
    thread.daemon = True
    thread.start()

# 獲取交易數據（使用緩存）
def get_transactions():
    global TRANSACTIONS_CACHE, TRANSACTIONS_CACHE_TIME
    
    current_time = time.time()
    if (TRANSACTIONS_CACHE is not None and 
        TRANSACTIONS_CACHE_TIME is not None and
        current_time - TRANSACTIONS_CACHE_TIME < TRANSACTIONS_CACHE_DURATION):
        logger.info("使用緩存的交易數據")
        return TRANSACTIONS_CACHE
    
    try:
        client = setup_google_sheets()
        if client:
            sheet_name = os.environ.get('GOOGLE_SHEET_NAME', '股票投資管理')
            transactions = get_transactions_from_google_sheet(client, sheet_name, "交易紀錄")
            TRANSACTIONS_CACHE = transactions
            TRANSACTIONS_CACHE_TIME = current_time
            logger.info(f"從 Google Sheets 讀取 {len(transactions)} 筆交易數據")
            return transactions
        else:
            return []
    except Exception as e:
        logger.error(f"獲取交易數據時出錯: {e}")
        return []

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

# Fetch stock info - 使用 Google Sheets 數據
def fetch_stock_info(code, is_otc=False):
    # 使用緩存來減少 API 請求
    cache_key = f"{code}_{'TWO' if is_otc else 'TW'}"
    current_time = time.time()
    
    # 檢查緩存是否存在且未過期（30分鐘）
    if hasattr(fetch_stock_info, 'cache'):
        cached_data = fetch_stock_info.cache.get(cache_key)
        if cached_data and current_time - cached_data['timestamp'] < 1800:  # 30分鐘緩存
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
    
    # 如果 Google Sheets 沒有數據，使用默認值而不是嘗試 Yahoo Finance
    # 因為 Yahoo Finance 在 Render 環境中不可靠
    
    result = {"price": round(price, 2), "name": name}
    
    # 更新緩存
    if not hasattr(fetch_stock_info, 'cache'):
        fetch_stock_info.cache = {}
    fetch_stock_info.cache[cache_key] = {
        'timestamp': current_time,
        'data': result
    }
    
    return result

# Calculate portfolio summary
def get_portfolio_summary(transactions=None):
    if transactions is None:
        transactions = get_transactions()
        
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
    add_transaction_message = None

    # 獲取交易數據和投資組合摘要
    transactions = get_transactions()
    summary, total_cost, total_market_value, total_unrealized_profit, total_realized_profit = get_portfolio_summary(transactions)

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
                            # 清除交易緩存
                            global TRANSACTIONS_CACHE
                            TRANSACTIONS_CACHE = None
                            add_transaction_message = "交易已新增！"
                            
                            # 如果是買入交易，顯示額外訊息
                            if trans_type == "Buy":
                                add_transaction_message += " 已檢查並更新股票列表。"
                            
                            # 重新獲取交易數據
                            transactions = get_transactions()
                            summary, total_cost, total_market_value, total_unrealized_profit, total_realized_profit = get_portfolio_summary(transactions)
                        else:
                            error = "無法將交易添加到 Google Sheets"
                    else:
                        error = "無法連接到 Google Sheets"
            except ValueError as e:
                error = f"輸入無效: {str(e)}。請確保股數和價格為有效數字"

    # 渲染模板（適用於 GET 和 POST 請求）
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
        default_date=default_date,
        add_transaction_message=add_transaction_message
    )
        # 其餘的 POST 處理邏輯保持不變...
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

# 檢查股票是否存在於 stock_names 工作表
def check_stock_exists_in_names(client, sheet_name, code):
    try:
        stock_names_sheet = client.open(sheet_name).worksheet("stock_names")
        records = stock_names_sheet.get_all_values()
        
        # 检查所有记录，查找匹配的代码
        for row in records:
            if len(row) > 0 and row[0] == f"{code}.TW":  # 第一列是代码
                return True
        return False
    except gspread.exceptions.WorksheetNotFound:
        return False
    except Exception as e:
        logger.error(f"检查股票是否存在时出错: {e}")
        return False

# 添加新股票到 stock_names 工作表
def add_stock_to_names_sheet(client, sheet_name, code, name):
    try:
        # 尝试获取 stock_names 工作表，如果不存在则创建
        try:
            stock_names_sheet = client.open(sheet_name).worksheet("stock_names")
        except gspread.exceptions.WorksheetNotFound:
            # 创建更大的工作表（1000 行，10 列）
            stock_names_sheet = client.open(sheet_name).add_worksheet(title="stock_names", rows=1000, cols=10)
            # 添加标题行，注意顺序：code, price, name, pricenow
            stock_names_sheet.append_row(["code", "price", "name", "pricenow"])
        
        # 获取所有记录
        records = stock_names_sheet.get_all_values()
        
        # 找到第一个空行
        next_row = len(records) + 1
        
        # 检查是否超出网格限制
        if next_row > stock_names_sheet.row_count:
            # 增加行数
            stock_names_sheet.add_rows(100)
        
        # 构建您指定的公式
        formula = f'=IMPORTXML("https://tw.stock.yahoo.com/quote/"&A{next_row}&"","//*[@id=\'main-0-QuoteHeader-Proxy\']/div/div[2]/div[1]/div/span[1]")'
        
        # 使用批量更新并指定 value_input_option 为 USER_ENTERED 来避免单引号问题
        batch_data = [
            {
                'range': f'A{next_row}',
                'values': [[f"{code}.TW"]]
            },
            {
                'range': f'B{next_row}',
                'values': [[f'=D{next_row}']]
            },
            {
                'range': f'C{next_row}',
                'values': [[name]]
            },
            {
                'range': f'D{next_row}',
                'values': [[formula]]
            }
        ]
        
        # 执行批量更新，使用 USER_ENTERED 选项
        stock_names_sheet.batch_update(batch_data, value_input_option='USER_ENTERED')
        
        logger.info(f"已将股票 {code}.TW {name} 添加到 stock_names 工作表，行号: {next_row}")
        return True
    except Exception as e:
        logger.error(f"添加股票到 stock_names 工作表时出错: {e}")
        return False
        
# 初始化 Google Sheets 並啟動定期更新
initialize_google_sheets()
schedule_google_sheets_update(30)  # 每30分鐘更新一次

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
