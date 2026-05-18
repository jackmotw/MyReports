import time
import os
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import urllib.parse
import requests
import logging
import json

# 設定 Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class GoodinfoScraper:
    def __init__(self):
        self.user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': self.user_agent})

    def acquire_session_via_selenium(self):
        logging.info("正在透過 Selenium 獲取 Session Cookies...")
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument(f'user-agent={self.user_agent}')
        
        # Stealth options
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument('--disable-blink-features=AutomationControlled')
        
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        try:
            # 訪問指定頁面以建立 Cookies (使用者要求的加權指數頁面)
            target_url = "https://goodinfo.tw/tw/StockIdxDetail.asp?STOCK_ID=%E5%8A%A0%E6%AC%8A%E6%8C%87%E6%95%B8"
            driver.get(target_url)
            time.sleep(3)
            
            # 將 Selenium cookies 轉移到 requests session
            for cookie in driver.get_cookies():
                self.session.cookies.set(cookie['name'], cookie['value'], domain=cookie.get('domain'))
            
            logging.info("Session Cookies 獲取成功，正在關閉瀏覽器。")
        finally:
            driver.quit()

    def fetch_page(self, url, wait_seconds=5, retries=2):
        params = {}
        # 將原始網址轉為資料請求網址
        if "StockList.asp" in url and "STEP=DATA" not in url:
            decoded_url = urllib.parse.unquote(url)
            sheet = ""
            sheet2 = ""
            path = "StockList.asp"
            
            if "EPS" in decoded_url:
                sheet = "季獲利能力"
                sheet2 = "獲利能力 (季增減統計)"
            elif "單月營收年增率" in decoded_url:
                sheet = "營收狀況"
                sheet2 = "月營收狀況"
            elif "三大法人連買" in decoded_url:
                sheet = "法人買賣_三大"
                sheet2 = "法人連買連賣統計(日)"
                path = "StockList.asp"
            elif "週成交日均張" in decoded_url or "成交張數" in decoded_url:
                sheet = "漲跌及成交統計"
                sheet2 = "歷史排名及創近期新高新低(週)"
            
            # 建立 AJAX 參數
            query = urllib.parse.urlparse(url).query
            params = dict(urllib.parse.parse_qsl(query))
            params.update({
                "STEP": "DATA",
                "RANK_RANGE": "300",
                "SHEET": sheet,
                "SHEET2": sheet2,
                "RPT_TIME": "最新資料",
                "IS_RELOAD_REPORT": "T"
            })
            url = f"https://goodinfo.tw/tw/{path}"

        for i in range(retries + 1):
            try:
                logging.info(f"正在請求 (嘗試 {i+1}): {url}")
                if params:
                    logging.info(f"參數: {params}")
                
                headers = {
                    'Referer': 'https://goodinfo.tw/tw/StockList.asp',
                    'X-Requested-With': 'XMLHttpRequest'
                }
                response = self.session.get(url, params=params, headers=headers, timeout=30)
                response.encoding = 'utf-8'
                
                # AJAX 回傳通常是部分的 HTML
                if "divStockList" not in response.text and "<table" not in response.text:
                    if "稍後再試" in response.text or "異常訪問" in response.text:
                        logging.warning("偵測到異常訪問限制")
                    else:
                        logging.warning("頁面內容不完全 (找不到表格內容)")
                    raise Exception("Invalid page content")
                
                logging.info("資料內容抓取成功")
                break
            except Exception as e:
                logging.warning(f"請求失敗 (嘗試 {i+1}): {e}")
                if i == retries:
                    logging.error("已達最大重試次數")
                    return ""
                else:
                    time.sleep(10)
            
        logging.info(f"等待額外的 {wait_seconds} 秒...")
        time.sleep(wait_seconds)
        return response.text

    def parse_table(self, html_source):
        soup = BeautifulSoup(html_source, 'html.parser')
        
        # Goodinfo 的資料表格在 id="divStockList" 內
        table = soup.select_one('#divStockList table')
        if not table:
            # 備用選擇器
            table = soup.select_one('table.solid_1_padding_4_0_tbl')
        
        if not table:
            logging.error("找不到資料表格")
            return None

        # 處理 Header
        headers = []
        header_row = table.find('tr', bgcolor="#ebf3fb")
        if not header_row:
            header_row = table.find('tr')

        for cell in header_row.find_all(['th', 'td']):
            for br in cell.find_all('br'):
                br.replace_with("")
            text = cell.get_text().strip()
            headers.append(text)
        
        logging.info(f"解析到欄位: {headers}")
        
        # 尋找 '代號' 的位置
        try:
            id_idx = headers.index('代號')
        except ValueError:
            logging.error("找不到 '代號' 欄位")
            return pd.DataFrame(columns=headers)

        # 處理資料行
        data_rows = []
        rows = table.find_all('tr')
        logging.info(f"表格總列數 (含 header): {len(rows)}")
        
        for i, row in enumerate(rows):
            if row == header_row: continue
            
            cells = row.find_all(['td', 'th'])
            if len(cells) < len(headers): continue
            
            row_content = [cell.get_text().strip() for cell in cells]
            
            # 偵錯：印出第一筆可能的資料
            if i < 5:
                logging.info(f"第 {i} 列內容: {row_content}")
            
            stock_id = row_content[id_idx]
            if not stock_id or not re.match(r'^\d+$', stock_id):
                 continue
            
            data_rows.append(row_content[:len(headers)])

        logging.info(f"成功解析資料筆數: {len(data_rows)}")
        df = pd.DataFrame(data_rows, columns=headers)
        
        # 確保 '代號' 欄位在第一欄，方便後續使用
        if '代號' in df.columns:
            cols = ['代號'] + [c for c in df.columns if c != '代號']
            df = df[cols]
            
        return df

    def close(self):
        self.driver.quit()

def filter_eps(df):
    logging.info(f"開始 EPS 篩選, 原始數量: {len(df)}")
    # 尋找包含 EPS 且不包含 '年增' 或 '季增' 的欄位 (通常是 EPS(元))
    col = next((c for c in df.columns if 'EPS' in c and '增' not in c), None)
    if not col:
        logging.warning(f"找不到 EPS 欄位, 現有欄位: {df.columns.tolist()}")
        return df
    
    logging.info(f"使用 EPS 欄位: {col}, 範例值: {df[col].head(3).tolist()}")
    df[col] = pd.to_numeric(df[col], errors='coerce')
    filtered = df[df[col] > 0].copy()
    logging.info(f"EPS 篩選後數量: {len(filtered)}")
    return filtered

def filter_revenue(df):
    logging.info(f"開始營收篩選, 原始數量: {len(df)}")
    # 尋找年增與月增欄位
    col_yoy = next((c for c in df.columns if '營收' in c and '年增' in c), None)
    col_mom = next((c for c in df.columns if '營收' in c and '月增' in c), None)
    
    for c in [col_yoy, col_mom]:
        if c:
            logging.info(f"使用營收欄位: {c}, 範例值: {df[c].head(3).tolist()}")
            df[c] = pd.to_numeric(df[c], errors='coerce')
            df = df[df[c] > 0]
        else:
            logging.warning(f"缺少營收篩選欄位 (YoY={col_yoy}, MoM={col_mom}), 現有欄位: {df.columns.tolist()}")
            
    logging.info(f"營收篩選後數量: {len(df)}")
    return df.copy()

def filter_institutional(df, foreign_req=True):
    logging.info(f"開始法人篩選, 原始數量: {len(df)}")
    # 尋找外資與投信連續買賣日數
    col_f = next((c for c in df.columns if '外資' in c and '日數' in c), None)
    col_i = next((c for c in df.columns if '投信' in c and '日數' in c), None)
    
    if col_f:
        logging.info(f"使用外資欄位: {col_f}, 範例值: {df[col_f].head(3).tolist()}")
        df[col_f] = pd.to_numeric(df[col_f], errors='coerce').fillna(0)
        if foreign_req:
            df = df[df[col_f] >= 2]
    else:
        logging.warning(f"找不到外資日數欄位, 現有欄位: {df.columns.tolist()}")

    if col_i:
        logging.info(f"使用投信欄位: {col_i}")
        df[col_i] = pd.to_numeric(df[col_i], errors='coerce').fillna(0)
    
    logging.info(f"法人篩選後數量: {len(df)}")
    return df.copy()

def filter_volume(df):
    logging.info(f"開始成交量篩選, 原始數量: {len(df)}")
    # 尋找成交量相關欄位
    col = next((c for c in df.columns if '成交' in c and '張' in c), None)
    if col:
        logging.info(f"使用成交量欄位: {col}, 範例值: {df[col].head(3).tolist()}")
        df[col] = df[col].astype(str).str.replace(',', '')
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        df = df[df[col] > 500]
    else:
        logging.warning(f"找不到成交量欄位, 現有欄位: {df.columns.tolist()}")
        
    logging.info(f"成交量篩選後數量: {len(df)}")
    return df.copy()

def generate_report(results):
    """
    results: dict { 'EPS': df, 'Revenue': df, 'Institutional': df, 'Volume': df, 'Final': df }
    """
    now_str = time.strftime("%Y-%m-%d %H:%M:%S")
    
    html_template = f"""
    <!DOCTYPE html>
    <html lang="zh-Hant">
    <head>
        <meta charset="UTF-8">
        <title>初步觀察名單篩選結果</title>
        <style>
            :root {{
                --primary: #2563eb;
                --secondary: #64748b;
                --success: #22c55e;
                --bg: #f8fafc;
                --card-bg: #ffffff;
            }}
            body {{ font-family: 'Inter', 'Noto Sans TC', sans-serif; background: var(--bg); margin: 0; color: #1e293b; }}
            .container {{ max-width: 1200px; margin: 40px auto; padding: 0 20px; }}
            h1 {{ text-align: center; color: #0f172a; font-size: 2.5rem; margin-bottom: 10px; }}
            .update-time {{ text-align: center; color: var(--secondary); margin-bottom: 40px; }}
            
            /* Tabs */
            .tabs {{ 
                display: flex; gap: 10px; border-bottom: 2px solid #e2e8f0; 
                margin-bottom: 20px; padding-bottom: 5px; overflow-x: auto;
                position: sticky; top: 0; background: var(--bg); z-index: 1000;
                padding-top: 15px;
            }}
            .tab-btn {{ 
                padding: 12px 24px; border: none; background: none; cursor: pointer; 
                font-weight: 600; color: var(--secondary); transition: 0.3s;
                border-radius: 8px 8px 0 0;
            }}
            .tab-btn:hover {{ background: #f1f5f9; color: var(--primary); }}
            .tab-btn.active {{ color: var(--primary); border-bottom: 3px solid var(--primary); background: #eff6ff; }}
            
            .tab-content {{ 
                display: none; background: var(--card-bg); border-radius: 12px; 
                box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1); padding: 0 20px 20px 20px; 
                max-height: calc(100vh - 280px); overflow: auto;
            }}
            .tab-content.active {{ display: block; animation: fadeIn 0.4s; }}
            
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            th, td {{ padding: 14px; text-align: left; border-bottom: 1px solid #e2e8f0; font-size: 0.9rem; }}
            th {{ background: #f8fafc; color: #475569; position: sticky; top: 0; z-index: 900; }}
            tr:hover {{ background: #f1f5f9; }}
            
            .badge {{ padding: 4px 8px; border-radius: 6px; font-size: 0.75rem; font-weight: 700; }}
            .badge-primary {{ background: #dbeafe; color: #1e40af; }}
            .badge-success {{ background: #dcfce7; color: #166534; }}
            .star-marker {{ color: #eab308; font-size: 1.2rem; margin-right: 5px; }}
            
            @keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(10px); }} to {{ opacity: 1; transform: translateY(0); }} }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 Preliminary Watch List</h1>
            <p class="update-time">最後更新日期：{now_str}</p>
            
            <div class="tabs">
                <button class="tab-btn active" onclick="showTab(event, 'tab-eps')">1. EPS > 0</button>
                <button class="tab-btn" onclick="showTab(event, 'tab-rev')">2. 營收(YoY/MoM)</button>
                <button class="tab-btn" onclick="showTab(event, 'tab-inst')">3. 法人(外資/投信)</button>
                <button class="tab-btn" onclick="showTab(event, 'tab-vol')">4. 成交張數</button>
                <button class="tab-btn" onclick="showTab(event, 'tab-final')">🎯 最終篩選</button>
            </div>
            
            <div id="tab-eps" class="tab-content active">{results['EPS'].to_html(classes='table', index=False, escape=False)}</div>
            <div id="tab-rev" class="tab-content">{results['Revenue'].to_html(classes='table', index=False, escape=False)}</div>
            <div id="tab-inst" class="tab-content">{results['Institutional'].to_html(classes='table', index=False, escape=False)}</div>
            <div id="tab-vol" class="tab-content">{results['Volume'].to_html(classes='table', index=False, escape=False)}</div>
            <div id="tab-final" class="tab-content">{results['Final'].to_html(classes='table', index=False, escape=False)}</div>
        </div>
        
        <script>
            function showTab(evt, tabId) {{
                let i, contents, btns;
                contents = document.getElementsByClassName("tab-content");
                for (i = 0; i < contents.length; i++) contents[i].className = contents[i].className.replace(" active", "");
                btns = document.getElementsByClassName("tab-btn");
                for (i = 0; i < btns.length; i++) btns[i].className = btns[i].className.replace(" active", "");
                document.getElementById(tabId).className += " active";
                evt.currentTarget.className += " active";
            }}
            
            // 處理最終分頁的標記 (標記同時符合外資與投信的標的)
            document.addEventListener("DOMContentLoaded", function() {{
                const finalTable = document.querySelector("#tab-final table");
                if (!finalTable) return;
                
                const rows = finalTable.querySelectorAll("tr");
                const headers = Array.from(rows[0].querySelectorAll("th")).map(th => th.innerText.trim());
                const fIdx = headers.indexOf("外資連續買賣日數");
                const iIdx = headers.indexOf("投信連續買賣日數");
                
                if (fIdx === -1 || iIdx === -1) return;
                
                for (let j = 1; j < rows.length; j++) {{
                    const tds = rows[j].querySelectorAll("td");
                    const fVal = parseFloat(tds[fIdx].innerText);
                    const iVal = parseFloat(tds[iIdx].innerText);
                    
                    if (fVal >= 2 && iVal >= 1) {{
                        tds[0].innerHTML = '<span class="star-marker">⭐</span>' + tds[0].innerHTML;
                        rows[j].style.background = "#fffbeb";
                    }}
                }}
            }});
        </script>
    </body>
    </html>
    """
    output_dir = "docs"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    report_path = os.path.join(output_dir, "PreliminaryWatchList.html")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(html_template)
    logging.info(f"報表 {report_path} 產生完成。")

def generate_json_report(df):
    """
    產生 JSON 格式的觀察名單，僅包含 stockNo, stockName
    """
    if df.empty:
        data = []
    else:
        # 提取 代號 與 名稱
        cols_map = {}
        if '代號' in df.columns: cols_map['代號'] = 'stockNo'
        if '名稱' in df.columns: cols_map['名稱'] = 'stockName'
        
        if not cols_map:
            logging.warning("JSON 報告內容為空：找不到 '代號' 或 '名稱' 欄位")
            data = []
        else:
            # 僅選取存在的欄位並重新命名
            selected_cols = [c for c in ['代號', '名稱'] if c in df.columns]
            data = df[selected_cols].rename(columns=cols_map).to_dict(orient='records')
            
    output_dir = "docs"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    report_path = os.path.join(output_dir, "PreliminaryWatchList.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    logging.info(f"JSON 報表 {report_path} 產生完成。")

def main():
    urls = {
        'EPS': "https://goodinfo.tw/tw/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E5%96%AE%E5%AD%A3EPS%E6%9C%80%E9%AB%98%40%40%E6%AF%8F%E8%82%A1%E7%A8%85%E5%BE%8C%E7%9B%88%E9%A4%98+%28EPS%29%40%40%E5%96%AE%E5%AD%A3EPS%E6%9C%80%E9%AB%98",
        'Revenue': "https://goodinfo.tw/tw/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E5%96%AE%E6%9C%88%E7%87%9F%E6%94%B6%E5%B9%B4%E5%A2%9E%E7%8E%87%28%E6%9C%AC%E6%9C%88%E4%BB%BD%29%40%40%E5%96%AE%E6%9C%88%E7%87%9F%E6%94%B6%E5%B9%B4%E5%A2%9E%E6%B8%9B%E7%8E%87%40%40%E6%9C%AC%E6%9C%88%E4%BB%BD%E5%B9%B4%E5%A2%9E%E7%8E%87",
        'Institutional': "https://goodinfo.tw/tw/StockList.asp?MARKET_CAT=%E6%99%BA%E6%85%A7%E9%81%B8%E8%82%A1&INDUSTRY_CAT=%E4%B8%89%E5%A4%A7%E6%B3%95%E4%BA%BA%E9%80%A3%E8%B2%B7+%E2%80%93+%E6%97%A5%40%40%E4%B8%89%E5%A4%A7%E6%B3%95%E4%BA%BA%E9%80%A3%E7%BA%8C%E8%B2%B7%E8%B6%85%40%40%E4%B8%89%E5%A4%A7%E6%B3%95%E4%BA%BA%E9%80%A3%E7%BA%8C%E8%B2%B7%E8%B6%85+%E2%80%93+%E6%97%A5",
        'Volume': "https://goodinfo.tw/tw/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E9%80%B1%E6%88%90%E4%BA%A4%E6%97%A5%E5%9D%87%E5%BC%B5%E5%89%B5%E8%BF%91%E6%9C%9F%E6%96%B0%E9%AB%98%E9%80%B1%E6%95%B8%40%40%E6%88%90%E4%BA%A4%E5%BC%B5%E6%95%B8%40%40%E9%80%B1%E6%88%90%E4%BA%A4%E6%97%A5%E5%9D%87%E5%BC%B5%E5%89%B5%E8%BF%91%E6%9C%9F%E6%96%B0%E9%AB%98%E9%80%B1%E6%95%B8"
    }
    
    scraper = GoodinfoScraper()
    scraper.acquire_session_via_selenium()
    
    raw_dfs = {}
    filtered_dfs = {}
    
    try:
        # 抓取所有頁面
        for name, url in urls.items():
            html = scraper.fetch_page(url)
            df = scraper.parse_table(html)
            if df is not None:
                raw_dfs[name] = df
            else:
                raw_dfs[name] = pd.DataFrame()

        # 分別篩選
        filtered_dfs['EPS'] = filter_eps(raw_dfs['EPS'])
        filtered_dfs['Revenue'] = filter_revenue(raw_dfs['Revenue'])
        filtered_dfs['Institutional'] = filter_institutional(raw_dfs['Institutional'], foreign_req=True)
        filtered_dfs['Volume'] = filter_volume(raw_dfs['Volume'])
        
        # 5. 最終結果 (1 ∩ 2 ∩ 3 ∩ 4)
        common_ids = set(filtered_dfs['EPS'].iloc[:, 0]) & \
                     set(filtered_dfs['Revenue'].iloc[:, 0]) & \
                     set(filtered_dfs['Institutional'].iloc[:, 0]) & \
                     set(filtered_dfs['Volume'].iloc[:, 0])
        
        df_final = filtered_dfs['Institutional'][filtered_dfs['Institutional'].iloc[:, 0].isin(common_ids)].copy()
        
        # 合併資訊
        try:
            for name in ['EPS', 'Revenue', 'Volume']:
                other_df = filtered_dfs[name]
                cols_to_use = [other_df.columns[0]] + [c for c in other_df.columns if c not in df_final.columns]
                df_final = pd.merge(df_final, other_df[cols_to_use], on=other_df.columns[0], how='left')
        except Exception as e:
            logging.warning(f"合併欄位時發生錯誤: {e}")

        filtered_dfs['Final'] = df_final
        generate_report(filtered_dfs)
        generate_json_report(df_final)
        
    except Exception as e:
        logging.error(f"程式執行發生錯誤: {e}")

if __name__ == "__main__":
    main()
