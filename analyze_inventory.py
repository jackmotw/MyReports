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
成本價：{stock_info['cost']}
目前收盤價：{latest_data['Close']:.2f} (漲跌幅比昨日：{((latest_data['Close'] / latest_data['Prev_Close']) - 1) * 100:.2f}%)
短期均線 (5MA)：{latest_data['MA5']:.2f}
波段均線 (10MA)：{latest_data['MA10']:.2f}
月線 (20MA)：{latest_data['MA20']:.2f}
季線 (60MA)：{latest_data['MA60']:.2f} (季線方向：{'向下' if latest_data['MA60_slope'] < 0 else '向上'})
年線 (200MA)：{latest_data['MA200']:.2f}
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
            # Use gemini-2.5-flash, treating chart and text as multimodal, and enable search grounding
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

def generate_report(target_stock=None, force=False, sheet_name="Inventory", excel_path="Watchlist.xlsx"):
    criteria_path = "ExitCriteria.txt"
    cache_path = "diagnosis_cache.json"
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

    now = datetime.now()
    today_open = now.replace(hour=9, minute=0, second=0, microsecond=0)
    today_close = now.replace(hour=13, minute=30, second=0, microsecond=0)
    
    # Monday is 0, Sunday is 6. Trading is Mon-Fri (0-4)
    is_weekday = now.weekday() < 5
    is_trading_hours = is_weekday and (today_open <= now < today_close)
    
    def is_cache_valid(stock_cache):
        if 'timestamp' not in stock_cache:
            return False
        from datetime import timedelta
        cache_time = datetime.fromisoformat(stock_cache['timestamp'])
        
        # Define the target as the most recent trading day's close (13:30)
        target_close = today_close
        
        # 1. If we are currently before today's target time, the reference must be at least one day earlier
        if now < target_close:
            target_close -= timedelta(days=1)
            
        # 2. Continually roll back if the target falls on a weekend (Saturday=5, Sunday=6)
        while target_close.weekday() >= 5:
            target_close -= timedelta(days=1)
            
        return cache_time > target_close

    if is_trading_hours:
        print(f"Intraday mode: Current time ({now.strftime('%H:%M')}) is during trading hours (Mon-Fri 09:00-13:30). AI diagnosis will be skipped to prioritize speed.")
    
    # In generate_report, replaces is_intraday with is_trading_hours
    is_intraday = is_trading_hours
    
    df_inventory = pd.read_excel(excel_path, sheet_name=sheet_name)
    
    # 找尋證券代號欄位
    symbol_col = None
    for col in df_inventory.columns:
        if '代號' in str(col) or '代碼' in str(col) or '股票' in str(col):
            symbol_col = col
            break
            
    if symbol_col is None:
        print("Error: Could not find symbol column in excel. Available columns:", df_inventory.columns.tolist())
        return
        
    stocks = []
    
    for idx, row in df_inventory.iterrows():
        symbol_raw = str(row[symbol_col]).strip()
        # Extract digits
        symbol = ''.join(filter(str.isdigit, symbol_raw))
        name_raw = symbol_raw.replace(symbol, '').strip()  # remove digits from name
        
        if not symbol:
            continue
            
        stocks.append({
            'symbol': symbol,
            'name': name_raw,
            'amount': row.get('股數', row.get('數量', 0)),
            'cost': row.get('成交均價', row.get('單價', 0))
        })
        
    print(f"Found {len(stocks)} stocks in inventory.")
    
    # Sync cache with current inventory (if not targeting a specific stock)
    inventory_symbols = {s['symbol'] for s in stocks}
    if not target_stock:
        stale_symbols = [s for s in cache if s not in inventory_symbols]
        if stale_symbols:
            print(f"  [Cache] Removing {len(stale_symbols)} stale stocks from cache: {', '.join(stale_symbols)}")
            for s in stale_symbols:
                if s in cache:
                    del cache[s]
    
    if not os.path.exists('charts'):
        os.makedirs('charts')
        
    for stock in stocks:
        symbol = stock['symbol']
        
        # Determine if we should process or skip
        is_targeted = (target_stock == symbol)
        
        if target_stock and not is_targeted:
            # We are targeting a specific stock, and this is not it. Skip processing, keep existing cache.
            continue
            
        # Determine if we should reuse cached AI diagnosis
        use_cached_ai = False
        if not is_targeted and not force and symbol in cache:
            if is_cache_valid(cache[symbol]):
                if not is_trading_hours:
                    # After trading hours and we have valid cache for today's close -> Full skip
                    print(f"Skipping {symbol} ({stock['name']}), valid cache found for today (Post-market).")
                    continue
                else:
                    # During trading hours -> Update numbers, but skip Gemini API
                    use_cached_ai = True
                    print(f"Skipping Gemini API for {symbol} ({stock['name']}), valid cache found for today (Intraday).")

        # Define filename relative to root for saving, but we will store relative to docs for HTML
        chart_filename_local = os.path.join(charts_dir, f"{symbol}_chart.png")
        chart_filename_html = f"charts/{symbol}_chart.png"
        import io
        import warnings
        
        yf_symbol = f"{symbol}.TW"
        print(f"Fetching data and generating AI diagnosis for {stock['name']} ({yf_symbol})...")
        
        try:
            with warnings.catch_warnings(), contextlib.redirect_stderr(io.StringIO()), contextlib.redirect_stdout(io.StringIO()):
                warnings.simplefilter("ignore")
                ticker = yf.Ticker(yf_symbol)
                df = ticker.history(period="1y")
            
            if df.empty:
                yf_symbol = f"{symbol}.TWO" # try OTC
                with warnings.catch_warnings(), contextlib.redirect_stderr(io.StringIO()), contextlib.redirect_stdout(io.StringIO()):
                    warnings.simplefilter("ignore")
                    ticker = yf.Ticker(yf_symbol)
                    df = ticker.history(period="1y")
                
            if df.empty:
                print(f"Warning: Could not fetch data for {symbol}.")
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
            prev2 = df.iloc[-3]
            
            # 判斷信號
            signals = []
            
            # 1. 跌破 5 日線 (短線轉弱)
            if latest['Close'] < latest['MA5']:
                signals.append("⚠️ 【跌破 5日線】短線轉弱，留意停損/停利。")
                
            # 2. 跌破 10 日線 (波段轉弱)
            if latest['Close'] < latest['MA10']:
                signals.append("🚨 【跌破 10日線】波段趨勢破壞，強烈建議波段出場。")
                
            # 3. 跌破 60 日線且季線向下 (生命線斷裂)
            ma60_slope = (latest['MA60'] - df['MA60'].iloc[-5]) / 5 if not pd.isna(df['MA60'].iloc[-5]) else 0
            if latest['Close'] < latest['MA60'] and ma60_slope < 0:
                signals.append("💀 【跌破季線且季線下彎】生命線斷裂，絕對停損訊號！")
                
            # 4. 爆量長黑
            is_black = (latest['Open'] - latest['Close']) / latest['Close'] > 0.02 # 實體黑K > 2%
            is_high_vol = latest['Volume'] > latest['Vol_MA20'] * 1.5
            if is_black and is_high_vol:
                signals.append("💣 【高檔爆量長黑】主力可能出貨，符合停利/停損要件。")
                
            # 5. 下跌量增
            if latest['Close'] < prev['Close'] and latest['Volume'] > prev['Volume']:
                signals.append("📉 【下跌量增】賣壓沉重，跌勢可能持續。")
                
            if not signals:
                signals.append("✅ 目前未觸發明顯技術面出場條件，可依紀律續抱。")

            # Plotting chart
            df_plot = df.iloc[-60:].copy() # last 3 months
            
            mc = mpf.make_marketcolors(up='r', down='g', edge='inherit', wick='inherit', volume='inherit')
            s  = mpf.make_mpf_style(marketcolors=mc, gridstyle=':', y_on_right=True, rc={'font.family': 'Microsoft JhengHei'})
            
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
            
            # Fetch Extra Data (Fundamentals, News, Institutional)
            print(f"  Fetching fundamental and news data for {symbol}...")
            info = {}
            try:
                info = ticker.info
            except Exception:
                pass
                
            market_cap = info.get('marketCap', '未知')
            if isinstance(market_cap, (int, float)):
                market_cap = f"{market_cap / 100000000:.2f} 億"
                
            pe_ratio = info.get('trailingPE', '未知')
            pb_ratio = info.get('priceToBook', '未知')
            revenue_growth = info.get('revenueGrowth', '未知')
            if isinstance(revenue_growth, (int, float)):
                revenue_growth = f"{revenue_growth * 100:.2f}%"
                
            inst_str = "無法取得完整籌碼資料"
            try:
                inst = ticker.institutional_holders
                if inst is not None and not inst.empty:
                    # just mention the top holders if available
                    if 'Holder' in inst.columns:
                        holders = inst['Holder'].head(3).tolist()
                        inst_str = f"主要法人持有者包含: {', '.join(holders)} 等"
            except Exception:
                pass
                
            extra_data = {
                'market_cap': market_cap,
                'pe_ratio': pe_ratio,
                'pb_ratio': pb_ratio,
                'revenue_growth': revenue_growth,
                'inst_holders': inst_str
            }

            # Call Gemini
            ai_diagnosis = cache.get(symbol, {}).get('ai_diagnosis', "")
            
            if use_cached_ai:
                # Reuse from cache if valid
                ai_diagnosis = cache.get(symbol, {}).get('ai_diagnosis', "")
            elif ai_client and not is_intraday:
                print(f"  [AI] Generating AI diagnosis for {stock['name']} with gemini-2.5-flash...")
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
            elif is_intraday:
                # Keep existing diagnosis if available in cache, else empty
                ai_diagnosis = cache.get(symbol, {}).get('ai_diagnosis', "⚠️ 盤中快速模式：AI 診斷暫停，收盤後 (13:30) 再自動恢復。")
                
            # Update cache directly
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
            
    # Save cache once after all processing is complete
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=4)
        print(f"Cache saved to {cache_path}")
            
    # Generate HTML from the consolidated cache
    results = list(cache.values())
    # Sort results by stock symbol numerically
    results.sort(key=lambda x: int(x['stock']['symbol']) if x['stock']['symbol'].isdigit() else x['stock']['symbol'])
    
    html_content = f"""
    <!DOCTYPE html>
    <html lang="zh-TW">
    <head>
        <meta charset="UTF-8">
        <title>股票庫存 出場與 AI 綜合診斷報告</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; margin: 0; padding: 20px; color: #333; display: flex; }}
            .sidebar {{ width: 200px; position: fixed; top: 20px; left: 20px; background: white; border-radius: 10px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); max-height: 90vh; overflow-y: auto; }}
            .sidebar h2 {{ font-size: 18px; color: #2c3e50; margin-top: 0; border-bottom: 2px solid #eee; padding-bottom: 10px; }}
            .sidebar ul {{ list-style: none; padding: 0; margin: 0; }}
            .sidebar li {{ margin-bottom: 8px; }}
            .sidebar a {{ text-decoration: none; color: #34495e; font-weight: 500; display: block; padding: 5px; border-radius: 4px; transition: background 0.2s; }}
            .sidebar a:hover {{ background: #eef2f5; color: #0056b3; }}
            .main-content {{ margin-left: 240px; flex: 1; max-width: 1200px; }}
            h1 {{ text-align: center; color: #2c3e50; }}
            .card {{ background: white; border-radius: 10px; padding: 20px; margin-bottom: 30px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
            .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #eee; padding-bottom: 10px; margin-bottom: 15px; }}
            .stock-title {{ font-size: 24px; font-weight: bold; color: #34495e; }}
            .price {{ font-size: 20px; font-weight: bold; color: #e74c3c; }}
            .content-wrapper {{ display: flex; gap: 20px; flex-wrap: wrap; }}
            .col-left {{ flex: 1; min-width: 300px; }}
            .col-right {{ flex: 1.5; min-width: 400px; }}
            .signals {{ margin-bottom: 20px; }}
            .signal-item {{ margin: 5px 0; padding: 10px; background: #fdf2f2; border-left: 4px solid #e74c3c; border-radius: 4px; }}
            .signal-item.safe {{ background: #f2f9f2; border-left: 4px solid #2ecc71; }}
            .ai-box {{ background: #f0f7ff; border: 1px solid #cce3f6; border-radius: 8px; padding: 15px; font-size: 15px; line-height: 1.6; }}
            .ai-box h3 {{ color: #0056b3; margin-top: 0; margin-bottom: 10px; display: flex; align-items: center; }}
            .ai-box h3::before {{ content: '✨'; margin-right: 8px; }}
            .chart-img {{ width: 100%; height: auto; border: 1px solid #ddd; border-radius: 4px; display: block; }}
            .footer {{ text-align: center; margin-top: 40px; color: #7f8c8d; font-size: 14px; margin-bottom: 40px; }}
            /* Markdown rendering styles */
            .ai-box p {{ margin-top: 0; }}
            .ai-box strong {{ color: #111; }}
            .ai-box ul {{ margin-top: 5px; padding-left: 20px; }}
        </style>
        <!-- Simple markdown parser for the browser -->
        <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
    </head>
    <body>
        <div class="sidebar">
            <h2>庫存清單導覽</h2>
            <ul>
    """
    
    for res in results:
        html_content += f"<li><a href='#stock-{res['stock']['symbol']}'>{res['stock']['name']} ({res['stock']['symbol']})</a></li>"
        
    html_content += """
            </ul>
        </div>
        <div class="main-content">
            <h1>📊 股票庫存 出場與 AI 綜合診斷報告</h1>
            <p style="text-align:center; color:#555;">分析時間: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
            <p style="text-align:center; color:#555; margin-bottom: 40px;">結合量化條件與 Gemini AI 綜合技術分析</p>
    """
    
    for res in results:
        sig_html = ""
        for sig in res['signals']:
            safe_class = "safe" if "✅" in sig else ""
            sig_html += f'<div class="signal-item {safe_class}">{sig}</div>'
            
        ai_raw = res.get('ai_diagnosis', '') or ""
        # Escape backticks for JS string injection
        ai_escaped = ai_raw.replace('`', '\\`').replace('$', '\\$')
        
        # Calculate Inventory PnL
        cost = res['stock'].get('cost', 0)
        amount = res['stock'].get('amount', 0)
        latest_price = res.get('latest_price', 0)
        
        inventory_html = ""
        if cost > 0 and amount > 0:
            pnl_amount = (latest_price - cost) * amount
            pnl_rate = (latest_price - cost) / cost * 100
            # Red for profit, Green for loss in Taiwan
            pnl_color = "#e74c3c" if pnl_amount >= 0 else "#2ecc71" 
            
            inventory_html = f"""
            <div style="background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px; padding: 15px; margin-bottom: 20px;">
                <h3 style="margin-top:0; color: #2c3e50; font-size: 16px;">💼 庫存狀態</h3>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; font-size: 15px;">
                    <div>買進均價：<b>{cost:.2f}</b></div>
                    <div>庫存股數：<b>{int(amount)}</b></div>
                    <div>預估損益：<b style="color: {pnl_color};">{pnl_amount:,.0f}</b></div>
                    <div>報酬率：<b style="color: {pnl_color};">{pnl_rate:.2f}%</b></div>
                </div>
            </div>
            """
        
        html_content += f"""
        <div class="card" id="stock-{res['stock']['symbol']}">
            <div class="header">
                <div class="stock-title">{res['stock']['name']} ({res['stock']['symbol']})</div>
                <div class="price">最新收盤價: {res['latest_price']:.2f} <span style="font-size: 14px; color: #7f8c8d; font-weight: normal; margin-left: 10px;">(更新時間: {res.get('timestamp', '未知')[:16].replace('T', ' ')})</span></div>
            </div>
            
            <div class="content-wrapper" style="flex-direction: column;">
                <div class="top-section" style="display: flex; gap: 20px; flex-wrap: wrap; width: 100%; margin-bottom: 20px;">
                    <div class="col-left">
                        {inventory_html}
                        <div class="signals">
                            <h3 style="margin-top:0;">📋 量化診斷結果</h3>
                            {sig_html}
                        </div>
                    </div>
                    <div class="col-right" style="display: flex; flex-direction: column; justify-content: center;">
                        <img class="chart-img" src="{res['chart_file']}" alt="K線圖">
                        <p style="text-align:center; font-size:12px; color:#666; margin-top:5px;">藍線: 5MA, 橘線: 10MA, 綠線: 20MA, 紅線: 60MA, 灰線: 200MA</p>
                    </div>
                </div>
                """
        
        if ai_raw:
            html_content += f"""
                <div class="bottom-section" style="width: 100%;">
                    <div class="ai-box">
                        <h3>Gemini 綜合診斷</h3>
                        <div id="ai-content-{res['stock']['symbol']}"></div>
                        <script>
                            document.getElementById('ai-content-{res['stock']['symbol']}').innerHTML = marked.parse(`{ai_escaped}`);
                        </script>
                    </div>
                </div>
            """
            
        html_content += f"""
            </div>
        </div>
        """
        
    html_content += """
            <div class="footer">
                <p>⚠️ 免責聲明：本報告由 AI 輔助產生，僅供學習與參考，不構成任何投資建議，投資人應自負盈虧。</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    report_path = os.path.join(output_dir, "ExitDiagnosisReport.html")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(html_content)
        
    print(f"Report generated: {report_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Stock Inventory AI Analysis")
    parser.add_argument("--stock", type=str, help="Specific stock symbol to single update (e.g. 2481)")
    parser.add_argument("--force", action="store_true", help="Force update all stocks regardless of timestamp")
    parser.add_argument("--sheet", type=str, default="Inventory", help="Excel sheet name (default: Inventory)")
    parser.add_argument("--excel", type=str, default="Watchlist.xlsx", help="Excel file path (default: Watchlist.xlsx)")
    args = parser.parse_args()
    
    generate_report(target_stock=args.stock, force=args.force, sheet_name=args.sheet, excel_path=args.excel)
