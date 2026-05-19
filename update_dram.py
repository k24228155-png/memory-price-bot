import pandas as pd
import time
import os
import re
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from datetime import datetime
import matplotlib.pyplot as plt

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

EXCEL_FILE = "Memory_Daily_Report.xlsx"

TARGETS = {
    "DRAM_Spot": "https://www.trendforce.com.tw/price/dram/dram_spot",
    "Flash_Spot": "https://www.trendforce.com.tw/price/flash/flash_spot",
    "Wafer_Spot": "https://www.trendforce.com.tw/price/flash/wafer_spot",
    "DRAM_Contract": "https://www.trendforce.com.tw/price/dram/dram_contract",
    "Flash_Contract": "https://www.trendforce.com.tw/price/flash/flash_contract"
}

def clean_change_value(text):
    text = text.replace('▲', '').replace('▼', '').replace('%', '').strip()
    try:
        return float(text)
    except ValueError:
        return 0.0

def format_header_name(name, type_str):
    """🌟 終極排版：過濾贅字，並強制分為三行顯示，讓欄位極度縮窄"""
    # 1. 移除時脈數字 (如 4800/5600, 3200 等)
    name = re.sub(r'\s*\d{4}/\d{4}|\s*\d{4}', '', name).strip()
    # 2. 在括號前加上換行符號 (讓規格掉到第二行)
    name = name.replace(' (', '\n(')
    # 3. 加上均價/漲跌幅 (掉到第三行)
    return f"{name}\n{type_str}"

def get_page_data(driver, url, sheet_name):
    print(f"\n⏳ 正在抓取 [{sheet_name}] ...")
    try:
        driver.get(url)
        time.sleep(6) 
        
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        table = soup.find('table')
        
        if not table:
            print(f"❌ [{sheet_name}] 找不到表格。")
            return None
            
        today_str = datetime.now().strftime("%Y-%m-%d")
        new_row = {"Date": today_str}
        
        headers = table.find_all('th')
        avg_idx, change_idx = 5, 6
        for i, th in enumerate(headers):
            text = th.text.lower()
            if 'avg' in text or 'average' in text: avg_idx = i
            if 'change' in text or '%' in text: change_idx = i
            
        rows = table.find_all('tr')
        parsed_count = 0
        
        for tr in rows:
            tds = tr.find_all('td')
            if len(tds) > max(avg_idx, change_idx):
                item_name = tds[0].text.strip().replace('\n', '')
                
                if item_name and "Item" not in item_name:
                    try:
                        avg_str = tds[avg_idx].text.strip()
                        change_str = tds[change_idx].text.strip()
                        
                        avg_val = float(avg_str)
                        change_val = clean_change_value(change_str)
                        
                        # 🌟 使用全新的三層式表頭名稱
                        col_avg = format_header_name(item_name, "均價")
                        col_change = format_header_name(item_name, "漲跌幅")
                        
                        new_row[col_avg] = avg_val
                        new_row[col_change] = change_val
                        parsed_count += 1
                    except Exception:
                        pass
                        
        if parsed_count > 0:
            print(f"✅ [{sheet_name}] 成功抓取 {parsed_count} 筆資料！")
            return pd.DataFrame([new_row])
        else:
            return None
            
    except Exception as e:
        return None

def draw_trend_chart(df, sheet_name):
    if len(df) < 2: return 
    plt.figure(figsize=(14, 7))
    avg_cols = [col for col in df.columns if '均價' in str(col)]
    if not avg_cols: return
        
    for col in avg_cols:
        display_name = str(col).replace('\n均價', '').replace('\n', ' ') 
        plt.plot(df['Date'], df[col], marker='o', linewidth=2, label=display_name)
    
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei'] 
    plt.rcParams['axes.unicode_minus'] = False
    
    plt.title(f'{sheet_name} - 均價趨勢圖', fontsize=14, fontweight='bold')
    plt.xlabel('日期')
    plt.ylabel('平均價格 (USD)')
    plt.xticks(rotation=45)
    plt.grid(True, linestyle='--', alpha=0.5)
    plt.legend(loc='center left', bbox_to_anchor=(1, 0.5), fontsize=9)
    plt.tight_layout()
    plt.savefig(f"{sheet_name}_Chart.png")
    plt.close()

def format_excel_style(filename):
    """🌟 讀取 Excel：設定極致窄欄位與三層換行，達成一眼看清"""
    if not os.path.exists(filename): return
        
    wb = load_workbook(filename)
    date_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")   
    avg_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")    
    change_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") 
    header_font = Font(bold=True)
    center_wrap_align = Alignment(horizontal="center", vertical="center", wrap_text=True) 

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # 1. 第一列高度拉高，容納三行文字
        ws.row_dimensions[1].height = 60
        
        # 2. 欄位極度縮窄
        for col in ws.columns:
            col_letter = col[0].column_letter 
            if col_letter == 'A':
                ws.column_dimensions[col_letter].width = 13
            else:
                ws.column_dimensions[col_letter].width = 11.5 # 🌟 縮窄至 11.5，畫面極度緊湊
                
        # 3. 置中與顏色套用
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = center_wrap_align
                if cell.row == 1:
                    cell.font = header_font
                    if cell.value == "Date": cell.fill = date_fill
                    elif cell.value and "均價" in str(cell.value): cell.fill = avg_fill
                    elif cell.value and "漲跌幅" in str(cell.value): cell.fill = change_fill
                
    wb.save(filename)
    print("🎨 Excel 排版優化 (自動清除冗餘欄位、極窄版面、三行換行) 完成！")

def update_excel_and_draw_charts(all_new_data):
    all_history = {}
    if os.path.exists(EXCEL_FILE):
        try:
            all_history = pd.read_excel(EXCEL_FILE, sheet_name=None)
        except Exception:
            pass
            
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        for sheet_name, new_df in all_new_data.items():
            if sheet_name in all_history:
                # 🌟 防呆機制：如果發現舊檔案的欄位跟我們新設計的對不上，直接洗掉重來！
                old_cols = set(all_history[sheet_name].columns)
                new_cols = set(new_df.columns)
                if len(old_cols.intersection(new_cols)) < 2:
                    print(f"⚠️ 偵測到 [{sheet_name}] 舊版多餘欄位，已自動為您重置版面！")
                    combined_df = new_df
                else:
                    combined_df = pd.concat([all_history[sheet_name], new_df], ignore_index=True)
                    combined_df.drop_duplicates(subset=['Date'], keep='last', inplace=True)
                    # 強制剔除舊的冗長欄位，只保留這次抓到的精簡欄位
                    combined_df = combined_df[new_df.columns] 
            else:
                combined_df = new_df
                
            combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
            draw_trend_chart(combined_df, sheet_name)
            
    format_excel_style(EXCEL_FILE)
    print(f"\n🎉 所有資料已成功統整儲存至: {EXCEL_FILE}")

def main():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    all_new_data = {}
    
    for sheet_name, url in TARGETS.items():
        df = get_page_data(driver, url, sheet_name)
        if df is not None:
            all_new_data[sheet_name] = df
            
    driver.quit()
    
    if all_new_data:
        update_excel_and_draw_charts(all_new_data)
    else:
        print("\n❌ 今日所有網頁皆無抓取到有效資料。")

if __name__ == "__main__":
    main()