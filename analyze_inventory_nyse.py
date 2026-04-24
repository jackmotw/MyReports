import pandas as pd
import yfinance as yf
import mplfinance as mpf
import os
from datetime import datetime
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
from google import genai
from google.genai import types
from PIL import Image
import json
import time
import argparse
import re
import warnings
import contextlib
import io

# 設定支援中文的字體 (適用於 Windows)
mpl.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
mpl.rcParams['axes.unicode_minus'] = False

def get_ai_diagnosis(client, stock_info, exit_criteria, latest_data, signals, chart_filename, extra_data):
    prompt = f"""
請根據以下「出場條件教學筆記（ExitCriteria.txt）」的內容，為以下這檔股票的最新技術面數值、基本面、籌碼面與新聞資訊進行綜合診斷。
你需要給予持有這檔股票的投資者一個明確、簡潔的動作建議（例如：續抱、減碼、嚴格停損等），並解釋原因。

--- 出場條件教學筆記（規則）---
{exit_criteria}

--- 基本股票資訊 ---
股票名稱：{stock_info['name']} ({stock_info['symbol']})
成本價：${stock_info['cost']}
目前收盤價：${latest_data['Close']:.2f} (漲跌幅比昨日：{((latest_data['Close'] / latest_data['Prev_Close']) - 1) * 100:.2f}%)
短期均線 (5MA)：${latest_data['MA5']:.2f}
波段均線 (10MA)：${latest_data['MA10']:.2f}
月線 (20MA)：${latest_data['MA20']:.2f}
季線 (60MA)：${latest_data['MA60']:.2f} (季線方向：{'向下' if latest_data['MA60_slope'] < 0 else '向上'})
年線 (200MA)：${latest_data['MA200']:.2f}
成交量：{latest_data['Volume']} (20日均量：{latest_data['Vol_MA20']:.2f})

--- 程式自動計算的技術面量化訊號 ---
{chr(10).join(signals) if signals else '無特別訊號'}

--- 籌碼與基本面資訊 ---
總市值：{extra_data.get('market_cap', '未知')}
本益比 (PE)：{extra_data.get('pe_ratio', '未知')}
股價淨值比 (PB)：{extra_data.get('pb_ratio', '未知')}
營收季成長率：{extra_data.get('revenue_growth', '未知')}
主要法人籌碼摘要：{extra_data.get('inst_holders', '無特別資訊')}

--- 你的任務 ---
1. 結合上述所有數據，並觀察附上的 K 線圖 (請判斷 K 線型態、均線排列、壓力支撐)。
2. 參考「出場條件教學筆記」的邏輯。
3. 請務必自行搜尋並分析該公司近期的重要新聞、產業趨勢或轉型近況 (例如：產品是否朝向 AI 發展、是否有特定利多/利空消息)，以作為輔助判斷。
4. 綜合評估這檔股票目前的狀態 (包含基本面、籌碼面的確認、技術面的強弱以及市場消息面的影響)。
5. 給出最終的操作建議與具體原因。

請使用繁體中文回答，使用 Markdown 格式排版，並讓重點清晰易懂。
"""
    for attempt in range(3):
        try:
            image = Image.open(chart_filename)
            # Use gemini-2.5-flash
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[prompt, image],
                config=types.GenerateContentConfig(
                    tools=[{"google_search": {}}]
                )
            )
            return response.text
        except Exception as e:
            err_str = str(e)
            if '503' in err_str or attempt < 2:
                print(f"  [AI] 發生錯誤 ({e})，等待 30 秒後進行第 {attempt + 2} 次重試...")
                time.sleep(30)
            else:
                return f"⚠️ 呼叫 Gemini API 發生錯誤且已達重試上限：{e}"

    time.sleep(30)

def generate_report(target_stock=None, force=False, sheet_name="Inventory2", excel_path="Watchlist.xlsx"):
    criteria_path = "ExitCriteria.txt"
    cache_path = "diagnosis_cache_nyse.json"
    
    output_dir = "docs"
    charts_dir = os.path.join(output_dir, "charts")
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    if not os.path.exists(charts_dir):
        os.makedirs(charts_dir)
        
    if not os.path.exists(excel_path):
        print(f"Error: {excel_path} not found.")
        return
        
    exit_criteria_text = ""
    if os.path.exists(criteria_path):
        with open(criteria_path, "r", encoding="utf-8") as f:
            exit_criteria_text = f.read()
    else:
        print(f"Warning: {criteria_path} not found. AI will not have the rules.")

    # Initialize Gemini Client
    api_key = os.environ.get("GEMINI_STOCK_API_KEY")
    ai_client = None
    if api_key:
        ai_client = genai.Client(api_key=api_key)
        print("Gemini API Client initialized successfully.")
    else:
        print("Warning: GEMINI_STOCK_API_KEY environment variable not set. AI diagnosis will be skipped.")
    
    # Load Cache
    cache = {}
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding="utf-8") as f:
                cache = json.load(f)
        except Exception as e:
            print(f"Warning: Could not read cache - {e}")

    # Read Excel File
    try:
        df_inventory = pd.read_excel(excel_path, sheet_name=sheet_name, engine='openpyxl')
    except Exception as e:
        print(f"讀取 Excel 發生錯誤: {e}")
        return

    # =====================================================================
    # 📝 Excel 欄位對應設定區 (強烈建議：若程式自動抓錯，請直接在此設定)
    # =====================================================================
    MANUAL_CONFIG = {
        'symbol': None,  # 股票代號 (你的表為 '股票代號')
        'name': None,    # 股票名稱 (你的表沒有獨立欄位，維持 None 即可，程式會自動從代號欄位切割)
        'amount': None,  # 股數/數量 (你的表為 '庫存股數')
        'cost': None     # 成本/買進單價 (你的表為 '買進均價')
    }
    
    col_keys = {
        'symbol': MANUAL_CONFIG['symbol'], 
        'name': MANUAL_CONFIG['name'], 
        'amount': MANUAL_CONFIG['amount'], 
        'cost': MANUAL_CONFIG['cost']
    }

    print("-" * 50)
    print(f"📊 Excel 讀取到的欄位清單:\n{df_inventory.columns.tolist()}")
    print("-" * 50)

    # 針對未手動設定 (None) 的欄位進行自動偵測
    for i, col in enumerate(df_inventory.columns):
        col_str = str(col).strip()
        
        if col_keys['symbol'] is None and any(kw in col_str for kw in ['股票代號', '代號', '代碼', '股票', 'Symbol', 'Ticker', 'Description']):
            col_keys['symbol'] = col
        if col_keys['name'] is None and any(kw in col_str for kw in ['股票名稱', '名稱', 'Name', 'Stock Name']):
            col_keys['name'] = col
        if col_keys['amount'] is None and any(kw in col_str for kw in ['庫存股數', '股數', '數量', '庫存', 'Qty', 'Position']):
            col_keys['amount'] = col
        if col_keys['cost'] is None and any(kw in col_str for kw in ['買進均價', '單價', '成本', 'Price', 'Cost']):
            # 排除掉總投資成本或帳面價值的關鍵字，精準鎖定「買進均價」
            if not any(exclude in col_str for exclude in ['總', '市值', '市價', '現價', '帳面']):
                col_keys['cost'] = col

    # 最終防呆：如果連自動偵測都找不到，依照剩餘欄位數量給予盲猜預設值
    if col_keys['symbol'] is None and len(df_inventory.columns) > 2: col_keys['symbol'] = 2
    if col_keys['amount'] is None and len(df_inventory.columns) > 3: col_keys['amount'] = 3
    if col_keys['cost'] is None and len(df_inventory.columns) > 5: col_keys['cost'] = 5

    print(f"🎯 實際對應使用的欄位 (若錯誤請至程式碼 MANUAL_CONFIG 修改):")
    print(f"  - 股票代號: [{col_keys['symbol']}]")
    print(f"  - 股票名稱: [{'包含於股票代號內' if col_keys['name'] is None else col_keys['name']}]")
    print(f"  - 庫存數量: [{col_keys['amount']}]")
    print(f"  - 平均成本: [{col_keys['cost']}]")
    print("-" * 50)

    # =====================================================================
    # 🧹 資料預先清理 (處理台幣雙數列/空白列)
    # =====================================================================
    print("🔍 正在檢查並過濾台幣計價附屬列...")
    original_len = len(df_inventory)
    
    if col_keys['symbol'] is not None:
        if isinstance(col_keys['symbol'], int):
            symbol_series = df_inventory.iloc[:, col_keys['symbol']]
        else:
            symbol_series = df_inventory[col_keys['symbol']]
            
        def is_valid_row(val):
            # 如果欄位為空白、NaN，則視為無效列(如台幣列)將其剃除
            s = str(val).strip().lower()
            return s not in ['', 'nan', 'none']
            
        mask = symbol_series.apply(is_valid_row)
        df_inventory = df_inventory[mask].reset_index(drop=True)
        
        filtered_len = len(df_inventory)
        if original_len != filtered_len:
            print(f"🧹 已自動過濾掉 {original_len - filtered_len} 列附屬資料 (例如：台幣計價列)。")
        else:
            print("🧹 原始資料已無附屬列，無需額外過濾。")
    print("-" * 50)

    # 輔助函數：用來安全地取得儲存格數值
    def get_val(row, key, default=""):
        if key is None: return default
        try:
            if isinstance(key, int):
                return row.iloc[key]
            else:
                return row[key]
        except:
            return default

    stocks = []
    
    for idx, row in df_inventory.iterrows():
        symbol_raw = str(get_val(row, col_keys['symbol'])).strip()
        if not symbol_raw or symbol_raw.lower() in ['nan', 'none']:
            continue
            
        # 1. 識別 Symbol 與 Name (針對 "名稱(代號)" 格式進行精準解析)
        match = re.search(r'\((.*?)\)', symbol_raw)
        if match:
            # 例如: "ADVANCED MICRO DEVICES(AMD)" -> symbol="AMD", name="ADVANCED MICRO DEVICES"
            symbol = match.group(1).strip()
            name_from_symbol = symbol_raw.split('(')[0].strip()
        else:
            symbol = ''.join(c for c in symbol_raw if c.isalnum() or c == '.')
            name_from_symbol = symbol_raw.replace(symbol, '').strip()
            
        # 2. 取得名稱
        if match:
            # 如果已經從「股票代號」成功切分出名稱，就強制優先使用該名稱
            name_raw = name_from_symbol
        else:
            # 否則嘗試從獨立的名稱欄位抓取
            name_raw = str(get_val(row, col_keys['name'])).strip()
            if not name_raw or name_raw.lower() in ['nan', 'none']:
                name_raw = name_from_symbol
                
        if not name_raw:
            name_raw = symbol
            
        # 排除掉標題列重複的情況
        if symbol in ["Ticker", "證券代號", "Symbol", "股票代號"]:
            continue
        
        # 3. 取得數量與成本
        try:
            amount_val = get_val(row, col_keys['amount'], 0)
            amount = float(amount_val) if pd.notna(amount_val) else 0
        except:
            amount = 0
            
        try:
            cost_val = get_val(row, col_keys['cost'], 0)
            cost = float(cost_val) if pd.notna(cost_val) else 0
        except:
            cost = 0
            
        # Debug 觀察前幾筆資料
        if len(stocks) < 5:
            print(f"  [資料讀取確認] {len(stocks)+1}: 代號={symbol}, 名稱={name_raw}, 股數={amount}, 成本={cost}")
            
        stocks.append({
            'symbol': symbol,
            'name': name_raw,
            'amount': amount,
            'cost': cost
        })
        
    print(f"✅ 成功從 Excel 整理出 {len(stocks)} 檔股票資料。")
    print("-" * 50)
    
    # 判斷快取是否過期 (12小時)
    def is_cache_valid(stock_cache):
        if 'timestamp' not in stock_cache:
            return False
        from datetime import timedelta
        cache_time = datetime.fromisoformat(stock_cache['timestamp'])
        return (datetime.now() - cache_time) < timedelta(hours=12)

    is_intraday = False # Default to post-market for analysis focus
    
    # Sync cache with current inventory
    inventory_symbols = {s['symbol'] for s in stocks}
    if not target_stock:
        stale_symbols = [s for s in cache if s not in inventory_symbols]
        if stale_symbols:
            print(f"  [Cache] 移除 {len(stale_symbols)} 檔已不在庫存的股票快取: {', '.join(stale_symbols)}")
            for s in stale_symbols:
                if s in cache:
                    del cache[s]
    
    if not os.path.exists('charts'):
        os.makedirs('charts')
        
    for stock in stocks:
        symbol = stock['symbol']
        is_targeted = (target_stock == symbol)
        
        if target_stock and not is_targeted:
            continue
            
        if not is_targeted and not force and symbol in cache:
            if is_cache_valid(cache[symbol]):
                print(f"Skipping {symbol} ({stock['name']}), valid cache found.")
                continue

        yf_symbol = symbol
        print(f"Fetching data and generating AI diagnosis for {stock['name']} ({yf_symbol})...")
        
        # Define filename relative to root for saving, but we will store relative to docs for HTML
        chart_filename_local = os.path.join(charts_dir, f"{symbol}_nyse_chart.png")
        chart_filename_html = f"charts/{symbol}_nyse_chart.png"
        
        try:
            with warnings.catch_warnings(), contextlib.redirect_stderr(io.StringIO()), contextlib.redirect_stdout(io.StringIO()):
                warnings.simplefilter("ignore")
                ticker = yf.Ticker(yf_symbol)
                df = ticker.history(period="1y")
            
            if df.empty:
                print(f"Warning: Could not fetch market data for {symbol}.")
                continue
                
            # Calculate MAs
            df['MA5'] = df['Close'].rolling(window=5).mean()
            df['MA10'] = df['Close'].rolling(window=10).mean()
            df['MA20'] = df['Close'].rolling(window=20).mean()
            df['MA60'] = df['Close'].rolling(window=60).mean()
            df['MA200'] = df['Close'].rolling(window=200).mean()
            df['Vol_MA20'] = df['Volume'].rolling(window=20).mean()
            
            latest = df.iloc[-1]
            prev = df.iloc[-2]
            
            # 判斷信號
            signals = []
            if latest['Close'] < latest['MA5']:
                signals.append("⚠️ 【Break below 5MA】Short-term weakness.")
            if latest['Close'] < latest['MA10']:
                signals.append("🚨 【Break below 10MA】Trend broken, consider exit.")
            
            ma60_slope = (latest['MA60'] - df['MA60'].iloc[-5]) / 5 if not pd.isna(df['MA60'].iloc[-5]) else 0
            if latest['Close'] < latest['MA60'] and ma60_slope < 0:
                signals.append("💀 【Death Cross/MA60 Down】Strong sell signal!")
                
            if not signals:
                signals.append("✅ No major sell signals triggered.")

            # Plotting chart
            df_plot = df.iloc[-60:].copy() 
            mc = mpf.make_marketcolors(up='g', down='r', edge='inherit', wick='inherit', volume='inherit')
            
            # 建立專屬的中文字體設定，強制傳給 mplfinance 使用
            my_rc = {
                'font.sans-serif': ['Microsoft JhengHei'],
                'font.family': 'sans-serif',
                'axes.unicode_minus': False
            }
            s  = mpf.make_mpf_style(marketcolors=mc, gridstyle=':', y_on_right=True, rc=my_rc)
            
            chart_filename = chart_filename_local
            apdict = [
                mpf.make_addplot(df_plot['MA5'], color='blue', width=1.5, label='MA5'),
                mpf.make_addplot(df_plot['MA10'], color='orange', width=1.5, label='MA10'),
                mpf.make_addplot(df_plot['MA20'], color='green', width=1.5, label='MA20'),
                mpf.make_addplot(df_plot['MA60'], color='red', width=2, label='MA60'),
                mpf.make_addplot(df_plot['MA200'], color='gray', width=2, label='MA200')
            ]
            
            mpf.plot(df_plot, type='candle', volume=True, style=s, addplot=apdict, 
                     title=f"{stock['name']} ({symbol})", savefig=chart_filename, 
                     figratio=(16,9), figscale=1.2)
            
            # Fetch Extra Data
            info = {}
            try: info = ticker.info
            except: pass
                
            market_cap = info.get('marketCap', 'Unknown')
            if isinstance(market_cap, (int, float)):
                if market_cap > 1e12: market_cap = f"${market_cap / 1e12:.2f} T"
                elif market_cap > 1e9: market_cap = f"${market_cap / 1e9:.2f} B"
                else: market_cap = f"${market_cap / 1e6:.2f} M"
                
            extra_data = {
                'market_cap': market_cap,
                'pe_ratio': info.get('trailingPE', 'Unknown'),
                'pb_ratio': info.get('priceToBook', 'Unknown'),
                'revenue_growth': f"{info.get('revenueGrowth', 0) * 100:.2f}%" if info.get('revenueGrowth') else "Unknown",
                'inst_holders': "Detailed data available via search."
            }

            # Call Gemini
            ai_diagnosis = ""
            if ai_client and not is_intraday:
                print(f"  [AI] 正在生成 {stock['name']} 的 AI 診斷報告...")
                latest_data = {
                    'Close': latest['Close'],
                    'Prev_Close': prev['Close'],
                    'MA5': latest['MA5'],
                    'MA10': latest['MA10'],
                    'MA20': latest['MA20'],
                    'MA60': latest['MA60'],
                    'MA200': latest['MA200'],
                    'MA60_slope': ma60_slope,
                    'Volume': latest['Volume'],
                    'Vol_MA20': latest['Vol_MA20']
                }
                ai_diagnosis = get_ai_diagnosis(ai_client, stock, exit_criteria_text, latest_data, signals, chart_filename_local, extra_data)
                
            cache[symbol] = {
                'stock': stock,
                'latest_price': latest['Close'],
                'signals': signals,
                'chart_file': chart_filename_html,
                'ai_diagnosis': ai_diagnosis,
                'timestamp': datetime.now().isoformat()
            }
            
        except Exception as e:
            print(f"Error processing {symbol}: {e}")
            
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=4)
            
    results = list(cache.values())
    results.sort(key=lambda x: x['stock']['symbol'])
    
    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>NYSE Inventory AI Analysis Report</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; margin: 0; padding: 20px; color: #333; display: flex; }}
            .sidebar {{ width: 220px; position: fixed; top: 20px; left: 20px; background: white; border-radius: 10px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); max-height: 90vh; overflow-y: auto; }}
            .sidebar h2 {{ font-size: 18px; color: #2c3e50; margin-top: 0; border-bottom: 2px solid #eee; padding-bottom: 10px; }}
            .sidebar ul {{ list-style: none; padding: 0; margin: 0; }}
            .sidebar li {{ margin-bottom: 8px; }}
            .sidebar a {{ text-decoration: none; color: #34495e; font-weight: 500; display: block; padding: 5px; border-radius: 4px; transition: background 0.2s; }}
            .sidebar a:hover {{ background: #eef2f5; color: #0056b3; }}
            .main-content {{ margin-left: 260px; flex: 1; max-width: 1200px; }}
            h1 {{ text-align: center; color: #2c3e50; }}
            .card {{ background: white; border-radius: 10px; padding: 20px; margin-bottom: 30px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
            .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #eee; padding-bottom: 10px; margin-bottom: 15px; }}
            .stock-title {{ font-size: 24px; font-weight: bold; color: #34495e; }}
            .price {{ font-size: 20px; font-weight: bold; color: #2ecc71; }}
            .content-wrapper {{ display: flex; gap: 20px; flex-wrap: wrap; }}
            .col-left {{ flex: 1; min-width: 300px; }}
            .col-right {{ flex: 1.5; min-width: 400px; }}
            .signal-item {{ margin: 5px 0; padding: 10px; background: #fdf2f2; border-left: 4px solid #e74c3c; border-radius: 4px; }}
            .signal-item.safe {{ background: #f2f9f2; border-left: 4px solid #2ecc71; }}
            .ai-box {{ background: #f0f7ff; border: 1px solid #cce3f6; border-radius: 8px; padding: 15px; font-size: 15px; line-height: 1.6; }}
            .chart-img {{ width: 100%; height: auto; border: 1px solid #ddd; border-radius: 4px; display: block; }}
            .footer {{ text-align: center; margin-top: 40px; color: #7f8c8d; font-size: 14px; margin-bottom: 40px; }}
        </style>
        <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
    </head>
    <body>
        <div class="sidebar">
            <h2>Portfolio Navigator</h2>
            <ul>
    """
    
    for res in results:
        html_content += f"<li><a href='#stock-{res['stock']['symbol']}'>{res['stock']['name']} ({res['stock']['symbol']})</a></li>"
        
    html_content += f"""
            </ul>
        </div>
        <div class="main-content">
            <h1>📊 NYSE/US Stock Inventory AI Diagnosis</h1>
            <p style="text-align:center; color:#555;">Generated at: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
    """
    
    for res in results:
        sig_html = "".join([f'<div class="signal-item {"safe" if "✅" in sig else ""}">{sig}</div>' for sig in res['signals']])
        ai_escaped = (res.get('ai_diagnosis') or "").replace('`', '\\`').replace('$', '\\$')
        
        cost, amount, latest_price = res['stock'].get('cost', 0), res['stock'].get('amount', 0), res.get('latest_price', 0)
        pnl_html = ""
        if cost > 0 and amount > 0:
            pnl_amount = (latest_price - cost) * amount
            pnl_rate = (latest_price - cost) / cost * 100
            pnl_color = "#2ecc71" if pnl_amount >= 0 else "#e74c3c"
            
            pnl_html = f"""
            <div style="background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px; padding: 15px; margin-bottom: 20px;">
                <h3 style="margin-top:0; color: #2c3e50; font-size: 16px;">💼 Portfolio Status</h3>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; font-size: 15px;">
                    <div>Avg Cost: <b>${cost:.2f}</b></div>
                    <div>Quantity: <b>{int(amount)}</b></div>
                    <div>Unrealized P/L: <b style="color: {pnl_color};">${pnl_amount:,.2f}</b></div>
                    <div>Return Rate: <b style="color: {pnl_color};">{pnl_rate:.2f}%</b></div>
                </div>
            </div>
            """
        
        html_content += f"""
        <div class="card" id="stock-{res['stock']['symbol']}">
            <div class="header">
                <div class="stock-title">{res['stock']['name']} ({res['stock']['symbol']})</div>
                <div class="price" style="color: #333;">Current Price: ${res['latest_price']:.2f}</div>
            </div>
            <div class="content-wrapper" style="flex-direction: column;">
                <div class="top-section" style="display: flex; gap: 20px; flex-wrap: wrap; width: 100%; margin-bottom: 20px;">
                    <div class="col-left">
                        {pnl_html}
                        <div class="signals">
                            <h3 style="margin-top:0;">📋 Analysis Signals</h3>
                            {sig_html}
                        </div>
                    </div>
                    <div class="col-right">
                        <img class="chart-img" src="{res['chart_file']}">
                    </div>
                </div>
                """
        if ai_escaped:
            html_content += f"""
                <div class="ai-box">
                    <h3>✨ AI Comprehensive Diagnosis</h3>
                    <div id="ai-content-{res['stock']['symbol']}"></div>
                    <script>
                        document.getElementById('ai-content-{res['stock']['symbol']}').innerHTML = marked.parse(`{ai_escaped}`);
                    </script>
                </div>
            """
        html_content += "</div></div>"
        
    html_content += """
            <div class="footer">
                <p>⚠️ Disclaimer: This AI-generated report is for reference only and does not constitute investment advice.</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    report_path = os.path.join(output_dir, "ExitDiagnosisReportNYSE.html")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"Report generated: {report_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="NYSE Stock Inventory AI Analysis")
    parser.add_argument("--stock", type=str, help="Specific stock ticker to update")
    parser.add_argument("--force", action="store_true", help="Force update")
    parser.add_argument("--sheet", type=str, default="Inventory2", help="Sheet name")
    parser.add_argument("--excel", type=str, default="Watchlist.xlsx", help="Excel file")
    args = parser.parse_args()
    generate_report(target_stock=args.stock, force=args.force, sheet_name=args.sheet, excel_path=args.excel)