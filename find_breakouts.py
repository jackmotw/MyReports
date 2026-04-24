import pandas as pd
import yfinance as yf
import mplfinance as mpf
import os
import requests
from datetime import datetime, timedelta
import numpy as np
from google import genai
from google.genai import types
from PIL import Image
import json
import time
import re
import argparse
import subprocess
from http.server import HTTPServer, BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs
import threading
import sys

# --- 設定 ---
OUTPUT_DIR = "docs"
BREAKOUT_CHARTS_DIR = os.path.join(OUTPUT_DIR, "breakout_charts")
os.makedirs(BREAKOUT_CHARTS_DIR, exist_ok=True)
GEMINI_API_KEY = os.environ.get("GEMINI_STOCK_API_KEY")
CACHE_FILE = "breakout_cache.json"
REPORT_FILE = os.path.join(OUTPUT_DIR, "BreakoutReport.html")
PYTHON_PATH = r"C:\ProgramData\anaconda3\envs\antigravity_project\python.exe"

def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=4)

def fetch_macro_retail_ratio():
    """計算微台指散戶多空比"""
    try:
        # 示範回傳 (實際上可爬 TAIFEX)
        return -0.02 
    except Exception as e:
        print(f"獲取總經數據錯誤: {e}")
        return 0

def fetch_tw_stock_list():
    """動態獲取股票代號、中文名稱及加入原因 (合併 Excel 觀察清單)"""
    url = "https://jackmotw.github.io/MyReports/"
    excel_file = "Watchlist.xlsx"
    sheet_name = "Watchlist"
    stock_info_map = {} # {symbol: {"name": str, "reason": str}}

    # 1. 從網頁獲取清單
    print(f"從 {url} 獲取動態清單...", flush=True)
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        content = response.text
        
        # 1.1 匹配 📈 6005 群益證 格式 (優先獲取基本清單)
        # 使用 \b 確保不匹配到 ID 內部的數字
        matches = re.findall(r'📈\s*\b(\d{4,6})\b\s*([^\s\n<]{1,20})', content)
        for sym, name in matches:
            stock_info_map[sym] = {"name": name.strip(), "reason": "MyReports 轉強標的"}
        print(f"  初步獲取 {len(matches)} 檔標的...", flush=True)

        # 1.2 改用更安全的非回溯匹配：逐行 (或逐 tr) 解析
        for row_html in re.findall(r'<tr[^>]*>(.*?)</tr>', content, re.DOTALL | re.IGNORECASE):
            m_sym = re.search(r'\b(\d{4,6})\b', row_html)
            m_star = re.search(r'(⭐+[^<]*)', row_html)
            if m_sym and m_star:
                sym = m_sym.group(1)
                rate = m_star.group(1).strip()
                if sym in stock_info_map:
                    stock_info_map[sym]["reason"] = rate
                else:
                    stock_info_map[sym] = {"name": sym, "reason": rate}
        
        print("  動態清單解析完成。", flush=True)
    except Exception as e:
        print(f"網頁獲取失敗: {e}")

    # 2. 從 Excel 獲取清單
    if os.path.exists(excel_file):
        print(f"讀取自定義觀察清單: {excel_file}...")
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            if 'stock number' in df.columns:
                for _, row in df.iterrows():
                    # 處理代號可能被讀成 float (2330.0) 的問題
                    sym = str(row['stock number']).split('.')[0].strip()
                    if sym and sym != 'nan':
                        # 優先捕捉 Excel 中的詳細原因 (處理 NaN 分支)
                        details_val = row.get('details')
                        if pd.isna(details_val) or str(details_val).strip() == '' or str(details_val) == 'nan':
                            reason = 'Excel 自定義觀察'
                        else:
                            reason = str(details_val).strip()
                        date_added = str(row.get('date added', '')).split(' ')[0] if pd.notna(row.get('date added')) else ''
                        if sym in stock_info_map:
                            # 優先以 Excel 的詳細內容作為顯示原因
                            stock_info_map[sym]["reason"] = reason
                            stock_info_map[sym]["date_added"] = date_added
                        else:
                            stock_info_map[sym] = {"name": sym, "reason": reason, "date_added": date_added}
            else:
                print("Excel 格式不符，請確保包含 'stock number' 欄位。")
        except Exception as e:
            print(f"讀取 Excel 錯誤: {e}")
    else:
        # 建立 Excel 範本
        print(f"建立自定義清單範本: {excel_file}...")
        df_template = pd.DataFrame(columns=['stock number', 'date added', 'details'])
        df_template.to_excel(excel_file, index=False, sheet_name=sheet_name)

    # 3. 從 PreliminaryWatchList.json 獲取清單
    json_file = "PreliminaryWatchList.json"
    if os.path.exists(json_file):
        print(f"讀取初步觀察清單: {json_file}...")
        try:
            with open(json_file, "r", encoding="utf-8") as f:
                json_data = json.load(f)
                for item in json_data:
                    sym = str(item.get('stockNo', '')).strip()
                    name = item.get('stockName', sym)
                    if sym:
                        if sym not in stock_info_map:
                            stock_info_map[sym] = {"name": name, "reason": "初步觀察名單 (Preliminary)"}
                        else:
                            # 如果已經存在，可以考慮在原因中追加標記
                            if "Preliminary" not in stock_info_map[sym]["reason"]:
                                stock_info_map[sym]["reason"] += " + Preliminary"
        except Exception as e:
            print(f"讀取 JSON 錯誤: {e}")

    if not stock_info_map:
        return {
            "2330": {"name": "台積電", "reason": "預設清單"}, 
            "2317": {"name": "鴻海", "reason": "預設清單"},
            "2454": {"name": "聯發科", "reason": "預設清單"},
            "2303": {"name": "聯電", "reason": "預設清單"},
            "2382": {"name": "廣達", "reason": "預設清單"}
        }
        
    print(f"最終合併清單共 {len(stock_info_map)} 個股票標的 (已移除重複)。")
    return stock_info_map

def update_watchlist_excel(symbol, date_str, details):
    excel_file = "Watchlist.xlsx"
    sheet_name = "Watchlist"

    try:
        if os.path.exists(excel_file):
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
        else:
            df = pd.DataFrame(columns=['stock number', 'date added', 'details'])
        
        # 統一處理代號格式，移除 .0 並轉由字串處理，確保比對精準
        symbol = str(symbol).split('.')[0].strip()
        df['stock number'] = df['stock number'].astype(str).str.split('.').str[0].str.strip()
        
        # 尋找是否存在該代號
        existing_mask = df['stock number'] == symbol
        
        if existing_mask.any():
            # 更新已存在的行 (只更新日期與詳情)
            idx = df.index[existing_mask].tolist()[0]
            df.loc[idx, 'date added'] = date_str
            df.loc[idx, 'details'] = details
            print(f"  [Excel] 更新成功 {symbol}")
        else:
            # 新增行
            new_row = {'stock number': symbol, 'date added': date_str, 'details': details}
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            print(f"  [Excel] 新增成功 {symbol}")
            
        # 嘗試儲存，若遇到檔案開啟中則提示，多試幾次
        for attempt in range(3):
            try:
                df.to_excel(excel_file, index=False, sheet_name=sheet_name)
                return True
            except PermissionError:
                if attempt < 2:
                    print(f"  [警告] {excel_file} 被開啟中，請關閉檔案，稍後重試 ({attempt+1}/3)...")
                    time.sleep(1.5)
                else:
                    raise
        return False
    except Exception as e:
        print(f"  [Excel 寫入錯誤] {e}")
        return False

def get_shareholder_trend(symbol):
    """模擬股權分散數據"""
    return {"shareholder_count_decreasing": True, "big_holder_increasing": True}

def get_tw_stock_name(symbol, default_name):
    """嘗試獲取台灣股票中文名稱"""
    # 這裡可以加入一些常見的硬編碼映射
    common_names = {
        "2330": "台積電", "2317": "鴻海", "2454": "聯發科", "2303": "聯電", "2382": "廣達",
        "2010": "春源", "2881": "富邦金", "2882": "國泰金", "2308": "台達電"
    }
    if symbol in common_names:
        return common_names[symbol]
        
    # 動態從 Yahoo 股市抓取中文名稱
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        res = requests.get(f"https://tw.stock.yahoo.com/quote/{symbol}", headers=headers, timeout=5)
        match = re.search(r'<title>(.*?)\(', res.text)
        if match:
            fetched_name = match.group(1).strip()
            if fetched_name and "Yahoo" not in fetched_name:
                return fetched_name
    except Exception as e:
        print(f"  [名稱解析] 獲取 {symbol} 中文名稱失敗: {e}")
        
    # 如果都失敗，回傳原本的 default_name
    return default_name

def quantitative_filter(symbol, scraped_info=None):
    """第一階段：量化數據過濾器 (Data Engine)"""
    yf_symbol = f"{symbol}.TW"
    try:
        import warnings
        import contextlib
        import io
        
        with warnings.catch_warnings(), contextlib.redirect_stderr(io.StringIO()), contextlib.redirect_stdout(io.StringIO()):
            warnings.simplefilter("ignore")
            ticker = yf.Ticker(yf_symbol)
            info = ticker.info
        
        if not info or 'regularMarketPrice' not in info:
            yf_symbol = f"{symbol}.TWO"
            with warnings.catch_warnings(), contextlib.redirect_stderr(io.StringIO()), contextlib.redirect_stdout(io.StringIO()):
                warnings.simplefilter("ignore")
                ticker = yf.Ticker(yf_symbol)
                info = ticker.info
        
        # --- 數據抓取 ---
        eps = info.get('trailingEps', 0)
        insider_pct = info.get('heldPercentInsiders', 0.1)
        rev_growth = info.get('revenueGrowth', 0)
        margins = info.get('operatingMargins', 0)
        trends = get_shareholder_trend(symbol)
        avg_vol = info.get('averageVolume', 1)
        curr_vol = info.get('volume', 1)
        
        # 優先使用抓取到的名稱
        name = scraped_info.get('name', symbol)
        if name == symbol or re.match(r'^[A-Za-z0-0\s]+$', name):
             # 如果名稱是純英文或代號，嘗試用我們的映射表
             name = get_tw_stock_name(symbol, name)
             # 如果還是沒變，嘗試 yfinance 的 shortName
             if name == symbol:
                 name = info.get('shortName', symbol)
        
        # 狀態矩陣
        status = {
            "eps_positive": eps > 0,
            "insider_ok": insider_pct > 0,
            "rev_dual_growth": rev_growth > 0,
            "operating_growth": margins > 0.05,
            "shareholders_decreasing": trends["shareholder_count_decreasing"],
            "big_holders_increasing": trends["big_holder_increasing"],
            "volume_support": curr_vol > avg_vol
        }

        # 計算評分 (Score) - 這裡補上原本遺失的變數定義
        score = sum(1 for v in status.values() if v)

        return score, yf_symbol, status, name, info.get('regularMarketPrice', 0)
        
    except Exception:
        return 0, None, {}, symbol, 0

def generate_kline_chart(yf_symbol, name):
    """第二階段：產生技術面 K 線圖 (含 Legend)"""
    ticker = yf.Ticker(yf_symbol)
    df = ticker.history(period="2y") # 增加歷史長度以計算 200MA
    if df.empty: return None
    
    symbol_code = yf_symbol.split('.')[0]
    df['5日均線'] = df['Close'].rolling(window=5).mean()
    df['10日均線'] = df['Close'].rolling(window=10).mean()
    df['20日均線'] = df['Close'].rolling(window=20).mean()
    df['60日均線'] = df['Close'].rolling(window=60).mean()
    df['200日均線'] = df['Close'].rolling(window=200).mean()
    df_plot = df.iloc[-120:]
    
    mc = mpf.make_marketcolors(up='r', down='g', edge='inherit', wick='inherit', volume='inherit')
    s  = mpf.make_mpf_style(marketcolors=mc, gridstyle=':', y_on_right=True, rc={'font.family': 'Microsoft JhengHei'})
    
    chart_path = os.path.join(BREAKOUT_CHARTS_DIR, f"{symbol_code}_kline.png")
    
    # 動態建立 addplot 與 Legend 清單，避免因數據不足出錯
    apdict = []
    legend_labels = []
    ma_map = [
        ('5日均線', '5MA(藍)', 'blue', 0.8),
        ('10日均線', '10MA(橘)', 'orange', 0.8),
        ('20日均線', '20MA(綠)', 'green', 1),
        ('60日均線', '60MA(紅)', 'red', 1.2),
        ('200日均線', '200MA(紫)', 'purple', 1.5)
    ]
    
    for col, label, color, width in ma_map:
        if not df_plot[col].dropna().empty:
            apdict.append(mpf.make_addplot(df_plot[col], color=color, width=width))
            legend_labels.append(label)
    
    # 建立圖表並增加 Legend
    fig, axlist = mpf.plot(df_plot, type='candle', volume=True, style=s, addplot=apdict,
                           title=f"{name} ({symbol_code}) 起漲點分析", savefig=chart_path, 
                           figratio=(16,9), figscale=1.2, returnfig=True)
    
    if legend_labels:
        axlist[0].legend(legend_labels, loc='upper left', fontsize=8)
    fig.savefig(chart_path)
    return chart_path

def vision_engine_analysis(client, symbol, chart_path):
    """Vision Engine: 呼叫 Gemini 進行圖片判定 (繁體中文提示詞)"""
    prompt = """
你是一位頂尖的技術分析師。請觀察這張包含 K 線、成交量以及 5MA(藍)、10MA(橘)、20MA(綠)、60MA(紅)、200MA(紫) 的股票走勢圖。
請根據以下視覺特徵進行嚴格判定，並回傳 JSON 格式：
{
  "均線糾結": bool, 
  "長期盤整": bool, 
  "帶量突破": bool, 
  "站上均線": bool,
  "信心程度": "0-100 之間的數值 (數字)",
  "投資建議": "一段簡短評語，說明為何建議或不建議買入"
}

判定標準：
1. 均線糾結： 圖中最右側近期的 5日、10日、60日均線是否在視覺上極度靠近、聚攏在一起？
2. 長期盤整： 在最右側發動點的左方，是否有一段明顯且漫長的橫盤整理區間（價格上下震盪幅度小）？
3. 帶量突破： 圖中最右側最新的一根 K 線，其下方的成交量柱狀圖是否出現『明顯的視覺突起』（顯著高於過去一段時間的平均低量）？
4. 站上均線： 最右側最新的一根 K 線是否為實體紅 K 棒，且其收盤價明確突破並站上 5 日均線（5MA）？

請「僅回傳 JSON」，不要有其他解釋文字。
"""
    try:
        img = Image.open(chart_path)
        response = client.models.generate_content(
            model='gemini-2.5-flash', 
            contents=[prompt, img],
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        return json.loads(response.text)
    except Exception as e:
        print(f"  [AI 錯誤] {symbol}: {e}")
        return None

def generate_html_report(cache):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sorted_items = sorted(cache.values(), key=lambda x: x.get('score', 0), reverse=True)
    
    # 防止 Windows 路徑在 JS 中被轉義
    escaped_python_path = PYTHON_PATH.replace('\\', '\\\\')
    
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>股票起漲點監控儀表板</title>
        <style>
            body {{ font-family: "Microsoft JhengHei", sans-serif; background: #f4f7f6; margin: 20px; }}
            .container {{ width: 100%; max-width: 1700px; margin: auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }}
            h1 {{ color: #2c3e50; text-align: center; border-bottom: 2px solid #3498db; padding-bottom: 10px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; table-layout: fixed; word-wrap: break-word; }}
            th, td {{ padding: 10px; text-align: left; border-bottom: 1px solid #eee; vertical-align: middle; }}
            th {{ background: #f8f9fa; color: #7f8c8d; text-transform: uppercase; font-size: 0.8em; white-space: nowrap; }}
            
            /* 具體欄位比例，加起來 100% */
            th:nth-child(1) {{ width: 4%; }}  /*勾選*/
            th:nth-child(2) {{ width: 6%; }}  /*代號*/
            th:nth-child(3) {{ width: 10%; }} /*名稱*/
            th:nth-child(4) {{ width: 8%; }}  /*日期*/
            th:nth-child(5) {{ width: 12%; }} /*原因*/
            th:nth-child(6) {{ width: 5%; }}  /*現價*/
            th:nth-child(7) {{ width: 5%; }}  /*評分*/
            th:nth-child(8) {{ width: 10%; }} /*量化*/
            th:nth-child(9) {{ width: 20%; }} /*AI 視覺鑑定*/
            th:nth-child(10) {{ width: 20%; }} /*圖表預覽*/
            
            .score {{ font-weight: bold; color: #e67e22; font-size: 1.1em; }}
            .status-icon {{ font-size: 1.1em; filter: grayscale(100%); opacity: 0.3; }}
            .status-icon.pass {{ filter: grayscale(0%); opacity: 1; }}
            .ai-pass {{ background: #d4edda; color: #155724; font-weight: bold; padding: 4px 8px; border-radius: 20px; display: inline-block; font-size: 0.8em; }}
            .chart-preview {{ width: 100%; border-radius: 4px; border: 1px solid #ddd; transition: 0.3s; }}
            .chart-preview:hover {{ transform: scale(1.02); }}
            .btn-group {{ margin: 20px 0; display: flex; gap: 15px; align-items: center; position: sticky; top: 0; background: white; padding: 15px 0; z-index: 1000; border-bottom: 1px solid #eee; }}
            .btn-action {{ padding: 10px 15px; background: #3498db; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 0.85em; font-weight: bold; width: 300px; }}
            .btn-action:disabled {{ background: #bdc3c7; cursor: not-allowed; }}
            #status-msg {{ font-weight: bold; color: #2980b9; }}
            .legend-inline {{ display: flex; gap: 8px; font-size: 0.8em; background: #fff; padding: 8px 12px; border-radius: 30px; border: 1px solid #e0e6ed; margin-left: auto; }}
            .legend-item {{ display: flex; align-items: center; gap: 4px; color: #555; white-space: nowrap; }}
            .command-box {{ margin: 15px 0; padding: 15px; background: #2c3e50; color: #ecf0f1; border-radius: 6px; font-family: 'Consolas', monospace; font-size: 0.9em; display: none; position: relative; }}
            .command-box b {{ color: #3498db; }}
            .copy-btn {{ position: absolute; right: 10px; top: 10px; padding: 5px 10px; background: #34495e; color: white; border: 1px solid #555; border-radius: 4px; cursor: pointer; font-size: 11px; }}
            .copy-btn:hover {{ background: #3498db; }}
            /* 新增格式優化 */
            td.reason-col {{ white-space: pre-wrap; font-size: 0.85em; line-height: 1.4; }}
            td.date-col {{ white-space: nowrap; font-family: 'Consolas', monospace; font-size: 0.85em; }}
        </style>
        <script>
            function selectAll(master) {{
                const checkboxes = document.querySelectorAll('input[name="select-stock"]');
                checkboxes.forEach(cb => cb.checked = master.checked);
                updateCommand();
            }}

            function updateCommand() {{
                const checked = Array.from(document.querySelectorAll('input[name="select-stock"]:checked')).map(el => el.value);
                const box = document.getElementById('cmd-box');
                const text = document.getElementById('cmd-text');
                if (checked.length > 0) {{
                    box.style.display = 'block';
                    text.innerText = `& "{escaped_python_path}" find_breakouts.py --analyze ${{checked.join(',')}}`;
                }} else {{
                    box.style.display = 'none';
                }}
            }}

            function copyCmd() {{
                const text = document.getElementById('cmd-text').innerText;
                navigator.clipboard.writeText(text);
                alert('指令已複製！請貼到終端機執行。');
            }}

            async function runAnalysis() {{
                const checked = Array.from(document.querySelectorAll('input[name="select-stock"]:checked')).map(el => el.value);
                if (checked.length === 0) {{ alert('請先勾選標的'); return; }}
                
                const btn = document.getElementById('btn-run');
                const status = document.getElementById('status-msg');
                btn.disabled = true;
                status.innerText = "⏳ AI 診斷執行中，請稍候...";
                
                try {{
                    const resp = await fetch('/run_ai', {{
                        method: 'POST',
                        headers: {{ 'Content-Type': 'application/json' }},
                        body: JSON.stringify({{ symbols: checked.join(',') }})
                    }});
                    const res = await resp.json();
                    if (res.success) {{
                        status.innerText = "✅ 執行完成！正在重新整理...";
                        setTimeout(() => location.reload(), 1500);
                    }} else {{
                        alert('執行錯誤: ' + res.error);
                        btn.disabled = false;
                        status.innerText = "";
                    }}
                }} catch (e) {{
                    alert('後端連線失敗，請確保程式已運行 --serve 模式');
                    btn.disabled = false;
                    status.innerText = "";
                }}
            }}

            async function addToWatchlist(symbol, details, btn) {{
                const today = new Date().toISOString().split('T')[0];
                try {{
                    const resp = await fetch('/add_to_watchlist', {{
                        method: 'POST',
                        headers: {{ 'Content-Type': 'application/json' }},
                        body: JSON.stringify({{ symbol, date: today, details }})
                    }});
                    const res = await resp.json();
                    if (res.success) {{
                        btn.innerText = '✅ 已加入';
                        btn.style.background = '#27ae60';
                        btn.disabled = true;
                    }} else {{
                        alert('加入失敗');
                    }}
                }} catch (e) {{
                    alert('連線失敗，請確保程式已運行 --serve 模式');
                }}
            }}
        </script>
    </head>
    <body>
        <div class="container">
            <h1>🚀 雙引擎起漲點偵測儀表板</h1>
            <p style="text-align:center; color: #7f8c8d;">最後更新時間: {now_str}</p>
            
            <div class="btn-group">
                <button id="btn-run" class="btn-action" onclick="runAnalysis()">一鍵 AI 診斷</button>
                <span id="status-msg"></span>
                
                <div class="legend-inline">
                    <div class="legend-item">💰 獲利</div>
                    <div class="legend-item">👤 內部持股</div>
                    <div class="legend-item">📈 營收雙增</div>
                    <div class="legend-item">⚙️ 營益成長</div>
                    <div class="legend-item">📉 散戶退場</div>
                    <div class="legend-item">🐳 大戶卡位</div>
                </div>
            </div>

            <div id="cmd-box" class="command-box">
                <button class="copy-btn" onclick="copyCmd()">點此複製</button>
                <b>[手動模式指令]</b><br>
                <span id="cmd-text"></span>
            </div>

            <table>
                <thead>
                    <tr>
                        <th><input type="checkbox" onclick="selectAll(this)"> 全選</th>
                        <th>代號</th>
                        <th>股票名稱</th>
                        <th>加入日期</th>
                        <th>加入原因 / 評級</th>
                        <th>現價</th>
                        <th>評分</th>
                        <th>量化指標狀態</th>
                        <th>AI 視覺鑑定</th>
                        <th>圖表預覽</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    for item in sorted_items:
        s = item.get('status', {})
        ai = item.get('ai_analysis')
        ai_html = ""
        excel_details = item.get('reason', '無系統標記')
        
        if ai:
            # 判斷是否為「強勢起漲」：四個視覺指標皆為 True 且信心度 > 70
            visual_indicators = [v for k, v in ai.items() if isinstance(v, bool)]
            is_all = all(visual_indicators) if visual_indicators else False
            confidence = ai.get('信心程度', 0)
            advice = ai.get('投資建議', '無建議內容')
            
            res_text = "⭐ 強勢起漲" if (is_all and int(confidence) >= 70) else "診斷中 (條件未全過)"
            bg_color = '#d4edda' if (is_all and int(confidence) >= 70) else '#fff3cd'
            
            # 建立 structured details 給 Excel 使用
            v1 = ai.get("均線糾結") is True
            v2 = ai.get("長期盤整") is True
            v3 = ai.get("帶量突破") is True
            
            if v1 and v2 and v3:
                excel_details = "均線糾結 & 長期盤整 & 帶量突破"
            elif v1 and v2:
                excel_details = "均線糾結 & 長期盤整"
            elif v1:
                excel_details = "均線糾結"
            else:
                labels = ["均線糾結", "長期盤整", "帶量突破", "站上均線"]
                indicators = []
                for label in labels:
                    val = ai.get(label)
                    icon = "✅" if val is True else "❌"
                    indicators.append(f"{label}: {icon}")
                excel_details = "\n".join(indicators)
                excel_details += f"\n💡 專家建議: {advice}"

            analysis_time = ai.get('analysis_time', 'N/A')
            ai_html = f"""
                <div class='ai-pass' style='background: {bg_color}'>{res_text}</div>
                <div style='font-size: 13px; margin-top: 8px;'>
                    <span style='font-size: 11px; color: #7f8c8d;'>🕒 {analysis_time}</span><br>
                    <b>🎯 信心程度: {confidence}%</b><br>
                    <b>💡 專家建議:</b> {advice}<br>
                    <hr style='border: 0.5px solid #eee; margin: 8px 0;'>
                    <div style='font-size: 11px; color: #666;'>
                        { "<br>".join([f"{k}: {'✅' if v else '❌'}" for k, v in ai.items() if isinstance(v, bool)]) }
                    </div>
                </div>
            """
        
        status_html = f"""
            <span class="status-icon {'pass' if s.get('eps_positive') else ''}" title="獲利 (EPS > 0)">💰</span>
            <span class="status-icon {'pass' if s.get('insider_ok') else ''}" title="內部人持股">👤</span>
            <span class="status-icon {'pass' if s.get('rev_dual_growth') else ''}" title="營收雙增 (YoY/MoM)">📈</span>
            <span class="status-icon {'pass' if s.get('operating_growth') else ''}" title="營益率成長">⚙️</span>
            <span class="status-icon {'pass' if s.get('shareholders_decreasing') else ''}" title="股東人數減少 (散戶退場)">📉</span>
            <span class="status-icon {'pass' if s.get('big_holders_increasing') else ''}" title="大戶持股增加">🐳</span>
        """
        
        # 處理單引號轉義，避免 JS 語法錯誤
        safe_excel_details = excel_details.replace("'", "\\'").replace("\n", "\\n")
        
        html += f"""
            <tr>
                <td><input type="checkbox" name="select-stock" value="{item['symbol']}" onchange="updateCommand()"></td>
                <td>{item['symbol']}</td>
                <td style="font-weight:bold; white-space:nowrap;">{item['name']}</td>
                <td class="date-col">{item.get('date_added', '-')}</td>
                <td class="reason-col">{item.get('reason', '')}</td>
                <td>{item.get('price', 0):.2f}</td>
                <td class="score">{item.get('score', 0)}</td>
                <td>{status_html}</td>
                <td>{ai_html}</td>
                <td>
                    <button class="btn-action" style="padding: 8px 12px; font-size: 0.8em; margin-bottom: 5px; width: 100%;" 
                            onclick="addToWatchlist('{item['symbol']}', '{safe_excel_details}', this)">
                        加入觀察清單
                    </button>
                    <a href="breakout_charts/{item['symbol']}_kline.png" target="_blank">
                        <img class="chart-preview" src="breakout_charts/{item['symbol']}_kline.png?t={int(time.time())}" style="width: 100%;">
                    </a>
                </td>
            </tr>
        """
        
    html += """
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """
    with open(REPORT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"報告已更新: {REPORT_FILE}")

# --- Server 部份 ---
class BreakoutHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        if self.path == '/run_ai':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            params = json.loads(post_data)
            symbols = params.get('symbols', '')
            
            print(f"收到分析請求: {symbols}")
            try:
                # 優先使用腳本偵測到的 Python 路徑
                python_exe = sys.executable if sys.executable else PYTHON_PATH
                print(f"  [Server] 啟動子程序分析: {symbols}", flush=True)
                cmd = [python_exe, __file__, "--analyze", symbols]
                # 確保子程序輸出到主控制台，讓使用者看到進度
                subprocess.run(cmd, check=True)
                
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"success": True}).encode())
            except Exception as e:
                print(f"  [Server 發生錯誤] {e}")
                self.send_response(500)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"success": False, "error": str(e)}).encode())
        elif self.path == '/add_to_watchlist':
            try:
                content_length = int(self.headers['Content-Length'])
                post_data = self.rfile.read(content_length)
                params = json.loads(post_data)
                symbol = params.get('symbol')
                date_str = params.get('date')
                details = params.get('details')
                
                print(f"收到加入觀察清單請求: {symbol}")
                success = update_watchlist_excel(symbol, date_str, details)
                
                self.send_response(200 if success else 500)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"success": success}).encode())
            except Exception as e:
                print(f"  [Server 發生錯誤 - Watchlist] {e}")
                import traceback
                traceback.print_exc()
                self.send_response(500)
                self.end_headers()
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        try:
            msg = format % args
            # 避免 UnicodeEncodeError: 編碼為 utf-8 再輸出，或過濾掉無法編碼的字元
            sys.stderr.write(f"{self.address_string()} - - {msg.encode('utf-8', 'replace').decode('ascii', 'ignore')}\n")
        except:
            pass

    def do_GET(self):
        # 靜態檔案處理
        if self.path == '/favicon.ico':
            self.send_response(204)
            self.end_headers()
            return

        if self.path == '/' or self.path == '/index.html':
            if not os.path.exists(REPORT_FILE):
                self.send_error(404, "Report file not found")
                return
            with open(REPORT_FILE, 'rb') as f:
                content = f.read()
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(content)))
            self.send_header('Connection', 'close')
            self.end_headers()
            self.wfile.write(content)
        elif self.path.startswith('/breakout_charts/'):
            filename = os.path.join(OUTPUT_DIR, self.path[1:].split('?')[0]) # 移除 timestamp query
            if os.path.exists(filename):
                with open(filename, 'rb') as f:
                    content = f.read()
                self.send_response(200)
                if filename.endswith('.png'):
                    self.send_header('Content-type', 'image/png')
                self.send_header('Content-Length', str(len(content)))
                self.send_header('Connection', 'close')
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_error(404, "Image file not found")
        else:
            self.send_error(404, "Invalid path")

def run_server(port=8080):
    server_address = ('127.0.0.1', port)
    try:
        httpd = ThreadingHTTPServer(server_address, BreakoutHandler)
        print(f"\n本地伺服器已啟動: http://127.0.0.1:{port}")
        print("請在瀏覽器輸入上述網址。")
        print("您可以在網頁中勾選標的並點擊分析按鈕。\n")
        httpd.serve_forever()
    except OSError as e:
        if e.errno == 98 or e.errno == 10048:
            print(f"錯誤：連接埠 {port} 已被佔用。請關閉其他執行中的腳本或更換連接埠。")
        else:
            print(f"伺服器啟動失敗: {e}")

# --- Main ---
def main():
    parser = argparse.ArgumentParser(description="起漲點雙引擎 Agent 2.0 (繁體中文版)")
    parser.add_argument("--scan", action="store_true", help="掃描市場並更新評分與圖表")
    parser.add_argument("--analyze", type=str, help="欲跑 AI 視覺分析的股票代號（逗號隔開）")
    parser.add_argument("--serve", action="store_true", help="啟動本地互動式伺服器")
    parser.add_argument("--force", action="store_true", help="強制更新（忽略快取）")
    args = parser.parse_args()

    if not any([args.scan, args.analyze, args.serve]):
        parser.print_help()
        return

    client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None
    cache = load_cache()
    today_str = datetime.now().strftime("%Y-%m-%d")

    if args.scan:
        print("開始市場掃描...", flush=True)
        macro_ratio = fetch_macro_retail_ratio()
        if macro_ratio >= 0 and not args.force:
            print(f"總經預警：多空比 {macro_ratio} 過高。停止。", flush=True)
            return
        
        stock_info_map = fetch_tw_stock_list()
        # 清理快取：只保留目前清單中有的股票 (解決 2330 殘留問題)
        keys_to_remove = [k for k in cache if k not in stock_info_map]
        for k in keys_to_remove:
            del cache[k]
        print(f"  已同步快取清單 (移除 {len(keys_to_remove)} 檔過期標的)。", flush=True)

        for symbol, s_info in stock_info_map.items():
            if symbol in cache and cache[symbol].get('last_scan') == today_str and not args.force:
                print(f"  [快取] 跳過 {symbol} (今日已完成)", flush=True)
                # 即使快取也強制更新原因與日期，確保最新的 Excel 或網頁變更會反應到報表並儲存
                cache[symbol]['reason'] = s_info.get('reason', cache[symbol].get('reason', ''))
                cache[symbol]['date_added'] = s_info.get('date_added', cache[symbol].get('date_added', ''))
                save_cache(cache)
                continue
            
            print(f"  正在分析量化指標: {symbol}...", flush=True)
            score, yf_symbol, status, final_name, price = quantitative_filter(symbol, s_info)
            chart_path = generate_kline_chart(yf_symbol, final_name) if yf_symbol else None
            
            cache[symbol] = {
                "symbol": symbol,
                "yf_symbol": yf_symbol,
                "name": final_name,
                "price": price,
                "score": score,
                "status": status,
                "reason": s_info.get('reason', '無系統標記'),
                "date_added": s_info.get('date_added', ''),
                "chart_path": chart_path,
                "last_scan": today_str,
                "ai_analysis": cache.get(symbol, {}).get("ai_analysis")
            }
            save_cache(cache)
            time.sleep(1) # 間隔防鎖
        
        generate_html_report(cache)
        print("掃描完成。請打開 BreakoutReport.html 查看結果。", flush=True)

    if args.analyze:
        if not client:
            print("錯誤：找不到 GEMINI_STOCK_API_KEY 環境變數。")
            return
        
        targets = args.analyze.split(',')
        print(f"開始 AI 視覺診斷 ({len(targets)} 檔)...")
        for symbol in targets:
            item = cache.get(symbol)
            if not item: continue
            
            print(f"  正在診斷 {item['name']} ({symbol})...")
            ai_data = vision_engine_analysis(client, symbol, item['chart_path'])
            if ai_data:
                ai_data['analysis_time'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                item['ai_analysis'] = ai_data
                save_cache(cache)
                time.sleep(1)
        
        generate_html_report(cache)
        print("AI 診斷完成。")

    if args.serve:
        run_server()

if __name__ == "__main__":
    main()
