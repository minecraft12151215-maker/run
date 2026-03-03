import discord
from discord.ext import commands
import pandas as pd
import requests
import yfinance as yf
import datetime
import asyncio
import re
import os
import io
import base64
import numpy as np
from bs4 import BeautifulSoup
import logging
from dotenv import load_dotenv

# 載入 .env 檔案 (在 Railway 上執行時，這行不會報錯，會自動去抓 Railway 後台的變數)
load_dotenv()

# ✅ 強制設定 Agg 模式防止無形介面報錯
try:
    import matplotlib.pyplot as plt
    import matplotlib
    matplotlib.use('Agg') 
except ImportError:
    pass

try:
    import mplfinance as mpf
    HAS_MPF = True
except ImportError:
    HAS_MPF = False

try:
    from html2image import Html2Image
    HAS_H2I = True
except ImportError:
    HAS_H2I = False

logging.getLogger('yfinance').setLevel(logging.CRITICAL)

# ================= 機器人設定區 =================
# 1. 透過環境變數安全讀取 Token (請在 .env 或 Railway 後台設定 DISCORD_TOKEN)
TOKEN = os.getenv('DISCORD_TOKEN')

# 安全機制：如果抓不到 Token，直接停止執行並報錯
if not TOKEN:
    raise ValueError("❌ 找不到 DISCORD_TOKEN！請確認 .env 檔案或 Railway 環境變數是否已設定。")

intents = discord.Intents.all()
bot = commands.Bot(command_prefix='!', intents=intents)

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8'
}

# ================= 工具函數區 =================
def get_stock_name(stock_id):
    try:
        r = requests.get(f"https://tw.stock.yahoo.com/quote/{stock_id}", headers=HEADERS, timeout=3)
        if r.status_code == 200:
            return BeautifulSoup(r.text, 'html.parser').title.string.split('(')[0].strip()
    except: pass
    return "目標股票"

def parse_val(val_str):
    if not val_str or val_str == '-': return 0.0
    try: return float(val_str.replace(',', '').replace('%', '').replace('+', '').strip())
    except: return 0.0

def format_html_color(val_str):
    if not val_str or val_str == '-': return val_str
    val = parse_val(val_str)
    disp = val_str if val_str.startswith('+') or val <= 0 else f"+{val_str}"
    if val > 0: return f'<span class="text-red">{disp}</span>'
    elif val < 0: return f'<span class="text-green">{val_str}</span>'
    return val_str

def calc_pct_diff(curr_str, prev_str):
    if not prev_str or not curr_str: return ""
    diff = parse_val(curr_str) - parse_val(prev_str)
    if diff > 0: return f"+{diff:.2f}%"
    elif diff < 0: return f"{diff:.2f}%"
    return ""

# ================= 爬蟲區 =================
async def scrape_yahoo(stock_id, market):
    suffix = "TW" if market == "TWSE" else "TWO"
    margin_data, holder_data = {}, {}
    try:
        res_m = requests.get(f"https://tw.stock.yahoo.com/quote/{stock_id}.{suffix}/margin", headers=HEADERS, timeout=5)
        for li in BeautifulSoup(res_m.text, 'html.parser').find_all('li'):
            texts = list(li.stripped_strings)
            if len(texts) >= 9 and re.match(r'^(\d{4}/)?\d{2}/\d{2}$', texts[0]):
                margin_data = dict(zip(['date', 'm_diff', 'm_bal', 'm_use', 's_diff', 's_bal', 's_use', 'ratio', 'day_trade'], texts[:9]))
                break 
        await asyncio.sleep(0.1)
        res_h = requests.get(f"https://tw.stock.yahoo.com/quote/{stock_id}.{suffix}/major-holders", headers=HEADERS, timeout=5)
        rows = [list(li.stripped_strings) for li in BeautifulSoup(res_h.text, 'html.parser').find_all('li') if len(list(li.stripped_strings)) >= 4 and re.match(r'^(\d{4}/)?\d{2}/\d{2}$', list(li.stripped_strings)[0])]
        if rows:
            holder_data['current'] = rows[0]
            if len(rows) > 1: holder_data['prev'] = rows[1]
    except: pass
    return margin_data, holder_data

# ✅ 終極雙引擎財報爬蟲 (保證抓到 PE/PB/Yield 與正確 YoY)
async def scrape_yahoo_fundamentals(stock_id, market):
    suffix = "TW" if market == "TWSE" else "TWO"
    fund = {
        'pe': '-', 'pb': '-', 'yield': '-',
        'rev_m': '-', 'rev': '-', 'mom': '-', 'yoy': '-',
        'eps_q': '-', 'eps': '-'
    }
    
    # [主引擎] 1. 抓取估值 (暴力文本尋找，無視前端標籤變動)
    try:
        r1 = requests.get(f"https://tw.stock.yahoo.com/quote/{stock_id}.{suffix}", headers=HEADERS, timeout=5)
        texts = list(BeautifulSoup(r1.text, 'html.parser').stripped_strings)
        for i, t in enumerate(texts):
            if t == '本益比' and fund['pe'] == '-':
                if i+1 < len(texts) and re.match(r'^[\d\.]+$', texts[i+1]): fund['pe'] = texts[i+1]
            elif t == '股價淨值比' and fund['pb'] == '-':
                if i+1 < len(texts) and re.match(r'^[\d\.]+$', texts[i+1]): fund['pb'] = texts[i+1]
            elif t == '殖利率' and fund['yield'] == '-':
                if i+1 < len(texts) and '%' in texts[i+1]: fund['yield'] = texts[i+1]
    except: pass

    # [備用引擎] 呼叫 yfinance API 補齊空缺的 PE/PB/Yield
    try:
        if fund['pe'] == '-' or fund['pb'] == '-' or fund['yield'] == '-':
            yf_info = yf.Ticker(f"{stock_id}.{suffix}").info
            if fund['pe'] == '-' and yf_info.get('trailingPE'): fund['pe'] = f"{yf_info['trailingPE']:.2f}"
            if fund['pb'] == '-' and yf_info.get('priceToBook'): fund['pb'] = f"{yf_info['priceToBook']:.2f}"
            if fund['yield'] == '-' and yf_info.get('dividendYield'): fund['yield'] = f"{yf_info['dividendYield']*100:.2f}%"
    except: pass
    
    await asyncio.sleep(0.1)
    
    # 2. 抓取最新營收 (嚴格百分比過濾，修復 YoY 變天價的 Bug)
    try:
        r2 = requests.get(f"https://tw.stock.yahoo.com/quote/{stock_id}.{suffix}/revenue", headers=HEADERS, timeout=5)
        for li in BeautifulSoup(r2.text, 'html.parser').find_all('li'):
            texts = list(li.stripped_strings)
            if len(texts) >= 4 and re.match(r'^\d{4}/\d{2}$', texts[0]):
                fund['rev_m'] = texts[0]
                fund['rev'] = texts[1]
                # 嚴格只提取帶有 % 符號的字串，過濾掉任何絕對金額
                pcts = [t for t in texts if re.match(r'^[\+\-]?[\d\.]+%$', t.strip())]
                if len(pcts) >= 2:
                    fund['mom'] = pcts[0]
                    fund['yoy'] = pcts[1]
                break
    except: pass

    await asyncio.sleep(0.1)
    
    # 3. 抓取 EPS
    try:
        r3 = requests.get(f"https://tw.stock.yahoo.com/quote/{stock_id}.{suffix}/eps", headers=HEADERS, timeout=5)
        for li in BeautifulSoup(r3.text, 'html.parser').find_all('li'):
            texts = list(li.stripped_strings)
            if len(texts) >= 2 and ('Q' in texts[0] or '季' in texts[0]):
                fund['eps_q'] = texts[0]
                fund['eps'] = texts[1]
                break
    except: pass

    return fund

# ================= 核心技術運算與畫圖區 =================
def analyze_tech_data(df):
    try:
        if len(df) < 60: return None 
        
        df['MA5'] = df['Close'].rolling(5).mean()
        df['MA10'] = df['Close'].rolling(10).mean()
        df['MA20'] = df['Close'].rolling(20).mean()
        df['MA60'] = df['Close'].rolling(60).mean()
        df['VolMA5'] = df['Volume'].rolling(5).mean()
        df['STD'] = df['Close'].rolling(20).std()
        df['UB'] = df['MA20'] + (2 * df['STD'])
        df['LB'] = df['MA20'] - (2 * df['STD'])
        
        low_min, high_max = df['Low'].rolling(9).min(), df['High'].rolling(9).max()
        df['RSV'] = (df['Close'] - low_min) / (high_max - low_min) * 100
        df['K'] = df['RSV'].ewm(com=2).mean()
        df['D'] = df['K'].ewm(com=2).mean()

        exp1, exp2 = df['Close'].ewm(span=12, adjust=False).mean(), df['Close'].ewm(span=26, adjust=False).mean()
        df['MACD'] = exp1 - exp2
        df['Signal'] = df['MACD'].ewm(span=9, adjust=False).mean()
        df['Histogram'] = df['MACD'] - df['Signal']

        delta = df['Close'].diff()
        rs = delta.clip(lower=0).ewm(com=5, adjust=False).mean() / (-1 * delta.clip(upper=0)).ewm(com=5, adjust=False).mean()
        df['RSI'] = 100 - (100 / (1 + rs))
        
        df['Force'] = ((df['Close'] - df['Open']) / (df['High'] - df['Low']).replace(0, 0.01)) * df['Volume']
        conc_5 = (df['Force'].rolling(5).sum() / df['Volume'].rolling(5).sum()).fillna(0) * 100

        df_120 = df.tail(120).dropna(subset=['High', 'Low', 'Close', 'Volume'])
        poc_price = 0
        if len(df_120) > 0:
            tp = (df_120['High'] + df_120['Low'] + df_120['Close']) / 3
            hist, bins = np.histogram(tp, bins=50, weights=df_120['Volume'])
            poc_price = ((bins[:-1] + bins[1:]) / 2)[np.argmax(hist)]

        today, prev = df.iloc[-1], df.iloc[-2]

        info = {
            'trend': "🔥 強力多頭" if today['Close'] > today['MA20'] and today['MA20'] > prev['MA20'] else ("📈 多頭格局" if today['Close'] > today['MA20'] else "📉 弱勢整理"),
            'kd': "⚠️ 過熱區" if today['K'] > 80 else ("❄️ 超賣區" if today['K'] < 20 else ("✨ 黃金交叉" if today['K'] > today['D'] and prev['K'] <= prev['D'] else ("💀 死亡交叉" if today['K'] < today['D'] and prev['K'] >= prev['D'] else "持穩"))),
            'vol': "🚨 爆大量" if today['Volume'] > today['VolMA5'] * 2 else ("🔺 溫和增量" if today['Volume'] > today['VolMA5'] else "💤 量縮整理"),
            'bb': "🚀 突破上軌" if pd.notna(today['UB']) and today['Close'] > today['UB'] else ("🌊 跌破下軌" if pd.notna(today['LB']) and today['Close'] < today['LB'] else "區間震盪")
        }
        
        score = 60 + sum([
            10 if pd.notna(today['MA20']) and today['Close'] > today['MA20'] else 0,
            10 if pd.notna(today['MA20']) and today['MA20'] > prev['MA20'] else 0,
            5 if pd.notna(today['K']) and pd.notna(today['D']) and today['K'] > today['D'] else 0,
            5 if pd.notna(today['RSI']) and today['RSI'] > 50 else 0,
            -10 if pd.notna(today['Close']) and today['Close'] < today['MA20'] else 0,
            -5 if pd.notna(today['K']) and pd.notna(today['D']) and today['K'] < today['D'] else 0
        ])

        return {'data': today, 'info': info, 'score': max(0, min(100, score)), 'chip_val_pct': float(conc_5.iloc[-1]), 'poc_price': float(poc_price)}
    except Exception as e:
        print(f"技術運算失敗: {e}")
        return None

def draw_professional_chart(df, stock_id):
    if not HAS_MPF or len(df) < 20: return None
    try:
        df_plot = df.tail(120).copy() 
        mc = mpf.make_marketcolors(up='#f04747', down='#43b581', edge='inherit', wick='inherit', volume='inherit')
        s = mpf.make_mpf_style(marketcolors=mc, gridstyle='--', gridcolor='#e0e0e0', facecolor='white', edgecolor='#4f545c', figcolor='white', rc={'text.color': '#333', 'axes.labelcolor': '#333', 'font.size': 20})
        macd_colors = ['#f04747' if val > 0 else '#43b581' for val in df_plot['Histogram']]
        
        apds = [
            mpf.make_addplot(df_plot['MA5'], color='#ff6b6b', width=2.5, panel=0),
            mpf.make_addplot(df_plot['MA10'], color='#3498db', width=2.5, panel=0),
            mpf.make_addplot(df_plot['MA20'], color='#fbc531', width=3, panel=0),
            mpf.make_addplot(df_plot['MA60'], color='#4cd137', width=3, panel=0),
            mpf.make_addplot(df_plot['UB'], color='#dcdde1', width=1.5, panel=0, linestyle='dotted'),
            mpf.make_addplot(df_plot['LB'], color='#dcdde1', width=1.5, panel=0, linestyle='dotted'),
            mpf.make_addplot(df_plot['K'], color='#ff9f43', width=2, panel=2, ylabel='KD'),
            mpf.make_addplot(df_plot['D'], color='#0abde3', width=2, panel=2),
            mpf.make_addplot(df_plot['MACD'], color='#3498db', width=2, panel=3, ylabel='MACD'),
            mpf.make_addplot(df_plot['Signal'], color='#e1b12c', width=2, panel=3),
            mpf.make_addplot(df_plot['Histogram'], type='bar', color=macd_colors, panel=3),
            mpf.make_addplot(df_plot['RSI'], color='#9b59b6', width=2, panel=4, ylabel='RSI')
        ]
        
        fig, axlist = mpf.plot(df_plot, type='candle', style=s, addplot=apds, volume=True, 
                               figsize=(22, 10.5), panel_ratios=(6, 2.5, 1.5, 1.5, 1.5), 
                               tight_layout=True, returnfig=True)
        
        tp = (df_plot['High'] + df_plot['Low'] + df_plot['Close']) / 3
        hist, bins = np.histogram(tp, bins=50, weights=df_plot['Volume'])
        bin_centers = (bins[:-1] + bins[1:]) / 2
        ax_vp = axlist[0].twiny()
        ax_vp.barh(bin_centers, hist, height=(bins[1]-bins[0])*0.8, color='#3498db', alpha=0.25, align='center', edgecolor='none')
        ax_vp.set_axis_off()
        
        poc_price = bin_centers[np.argmax(hist)]
        axlist[0].axhline(poc_price, color='#e67e22', linestyle='-', linewidth=3, alpha=0.8)
        
        price_range = df_plot['High'].max() - df_plot['Low'].min()
        axlist[0].text(2, poc_price + (price_range * 0.01), f' POC: {poc_price:.2f} ', 
                       color='white', fontsize=18, fontweight='bold', va='bottom', ha='left',
                       bbox=dict(facecolor='#e67e22', edgecolor='none', boxstyle='round,pad=0.3', alpha=0.9))

        axlist[0].text(len(df_plot)/2, (df_plot['High'].max()+df_plot['Low'].min())/2, f'{stock_id} AI REPORT', fontsize=90, color='grey', alpha=0.1, ha='center', va='center', fontweight='bold', rotation=15)

        last_price = df_plot['Close'].iloc[-1]
        last_color = '#f04747' if last_price >= df_plot['Close'].iloc[-2] else '#43b581'
        axlist[0].text(len(df_plot) - 0.5, last_price, f' ◀ {last_price:.2f} ', color='white', fontsize=26, fontweight='bold', va='center', ha='left', bbox=dict(facecolor=last_color, edgecolor='none', boxstyle='round,pad=0.3'))

        buf = io.BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight', pad_inches=0.1)
        plt.close(fig)
        return base64.b64encode(buf.getvalue()).decode('utf-8')
    except Exception as e:
        print(f"畫圖錯誤: {e}")
        return None

# ================= 指令區 =================
@bot.command()
async def check(ctx, stock_id: str):
    loading = await ctx.send(f"📄 正在編制 `{stock_id}` 終極無敵版財報 (修復擠壓與雙引擎啟動)...")
    
    df = yf.Ticker(f"{stock_id}.TW").history(period="1y") 
    market = 'TWSE'
    if df.empty:
        df = yf.Ticker(f"{stock_id}.TWO").history(period="1y")
        market = 'TPEx'
    if df.empty: return await loading.edit(content="❌ 查無報價資料。")

    name = get_stock_name(stock_id)
    tech = analyze_tech_data(df)
    if not tech: return await loading.edit(content="❌ 技術指標運算失敗。")

    margin_data, holder_data = await scrape_yahoo(stock_id, market)
    fund_data = await scrape_yahoo_fundamentals(stock_id, market)
    
    d, info, score = tech['data'], tech['info'], tech['score']
    date_str = d.name.strftime('%Y/%m/%d')
    poc_str = f"{tech['poc_price']:.2f}"
    
    ub_str = f"{d['UB']:.2f}" if pd.notna(d['UB']) else "計算中"
    ma20_str = f"{d['MA20']:.2f}" if pd.notna(d['MA20']) else "計算中"
    lb_str = f"{d['LB']:.2f}" if pd.notna(d['LB']) else "計算中"
    k_str = f"{int(d['K'])}" if pd.notna(d['K']) else "-"
    d_str = f"{int(d['D'])}" if pd.notna(d['D']) else "-"
    rsi_str = f"{int(d['RSI'])}" if pd.notna(d['RSI']) else "-"

    tp_price, sl_price = d['UB'] * 1.03, d['MA20'] * 0.98
    color = "#f04747" if score >= 80 else ("#faa61a" if score >= 60 else ("#43b581" if score >= 40 else "#3498db"))
    sug = "強力買進" if score >= 80 else ("持股續抱" if score >= 60 else ("觀望減碼" if score >= 40 else "建議賣出"))

    curr, prev = holder_data.get('current'), holder_data.get('prev')
    f_diff = calc_pct_diff(curr[1], prev[1]) if curr and prev else ""
    l_diff = calc_pct_diff(curr[2], prev[2]) if curr and prev else ""

    yoy_val = parse_val(fund_data['yoy'])
    eps_val = parse_val(fund_data['eps'])
    pe_val = parse_val(fund_data['pe'])

    fund_ai = ""
    if fund_data['yoy'] == '-':
        fund_ai = "⏳ 暫無最新財務數據"
    else:
        if yoy_val > 10 and eps_val > 0: fund_ai = "🔥 營收強勁雙增，具爆發力"
        elif yoy_val > 0 and eps_val > 0: fund_ai = "📈 營收穩定成長，獲利健康"
        elif yoy_val < 0 and eps_val > 0: fund_ai = "⚠️ 營收面臨衰退，仍維持獲利"
        elif yoy_val < 0 and eps_val <= 0: fund_ai = "🧊 營收衰退且虧損，風險高"
        else: fund_ai = "⚖️ 基本面平穩，未見明顯爆發"
        
        if pe_val > 0 and pe_val < 12: fund_ai += "<br><span style='color:#e74c3c; font-size:26px;'>估值偏低具優勢</span>"
        elif pe_val > 30: fund_ai += "<br><span style='color:#2ecc71; font-size:26px;'>估值偏高需留意風險</span>"

    if HAS_H2I:
        try:
            b64_chart = draw_professional_chart(df, stock_id)
            c_val_pct = tech['chip_val_pct']
            boxes = max(0, min(10, int((c_val_pct + 40) / 80 * 10)))
            boxes_html = ('<div class="box bg-red"></div>' * boxes + '<div class="box bg-gray"></div>' * (10 - boxes)) if c_val_pct > 0 else ('<div class="box bg-gray"></div>' * boxes + '<div class="box bg-green"></div>' * (10 - boxes))

            html = f"""
            <!DOCTYPE html><html><head><meta charset="UTF-8">
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;700;900&display=swap');
                body {{ background: #f0f2f5; margin: 0; padding: 0; font-family: 'Noto Sans TC', sans-serif; display: flex; justify-content: center; }}
                
                .a4-container {{ width: 2480px; height: 3508px; background: white; padding: 60px 100px 160px 100px; box-sizing: border-box; display: flex; flex-direction: column; position: relative; box-shadow: 0 0 50px rgba(0,0,0,0.1); overflow: hidden; border-top: 35px solid {color}; }}
                
                .header {{ display: flex; justify-content: space-between; align-items: flex-end; padding-bottom: 20px; margin-bottom: 30px; border-bottom: 3px solid #eee; }}
                .header-left h1 {{ margin: 0; font-size: 110px; color: #222; font-weight: 900; letter-spacing: 3px; }}
                .header-left p {{ margin: 15px 0 0 0; font-size: 40px; color: #7f8c8d; font-weight: normal; }}
                .header-right {{ text-align: right; font-size: 38px; color: #95a5a6; line-height: 1.4; }}
                
                /* ✅ 修復擠壓：移除 height 限制，讓排版自由呼吸 */
                .metrics-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 50px; margin-bottom: 30px; }}
                .metric-card {{ background: #fff; border-radius: 12px; border: 1px solid #eee; border-top: 15px solid {color}; display: flex; flex-direction: column; justify-content: center; align-items: center; box-shadow: 0 4px 15px rgba(0,0,0,0.02); padding: 35px 20px; min-height: 200px; }}
                .m-title {{ font-size: 38px; color: #7f8c8d; margin: 0 0 15px 0; font-weight: 700; }}
                /* ✅ 修正字體大小與行高，確保 "強力買進" 四個字方正飽滿不擠壓 */
                .m-value {{ font-size: 90px; font-weight: 900; color: {color}; margin: 0; line-height: 1.2; letter-spacing: 2px; text-align: center; }}
                
                .chart-wrapper {{ margin-bottom: 30px; border: 1px solid #eee; border-radius: 12px; padding: 25px; background: #fff; box-shadow: 0 4px 15px rgba(0,0,0,0.02); }}
                .legend-container {{ display: flex; justify-content: center; align-items: center; gap: 40px; font-size: 32px; color: #555; font-weight: 700; margin-bottom: 20px; }}
                .legend-item {{ display: flex; align-items: center; gap: 12px; }}
                .l-dot {{ font-size: 40px; }}
                
                .chart-container {{ height: 860px; display: flex; justify-content: center; align-items: center; border-top: 15px solid {color}; padding-top: 20px; }}
                .chart-container img {{ max-width: 100%; max-height: 100%; object-fit: contain; }}

                .bottom-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 40px; margin-bottom: 30px; height: auto; }}
                
                .info-card {{ background: #fff; border-radius: 12px; border: 1px solid #eee; padding: 40px; display: flex; flex-direction: column; box-shadow: 0 4px 15px rgba(0,0,0,0.02); }}
                .info-card h2 {{ font-size: 48px; margin: 0 0 30px 0; color: #2c3e50; border-bottom: 2px solid #f0f0f0; padding-bottom: 20px; display: flex; align-items: center; gap: 15px; }}
                
                .data-table {{ width: 100%; border-collapse: collapse; font-size: 38px; color: #34495e; }}
                .data-table tr:nth-child(even) {{ background-color: #fcfcfc; }}
                .data-table td {{ padding: 18px 10px; border-bottom: 1px solid #f5f5f5; line-height: 1.3; }}
                .data-table td.label {{ color: #7f8c8d; width: 48%; font-weight: 700; }}
                .data-table td.value {{ font-weight: 900; text-align: right; }}
                
                .code-box {{ background: #fdfdfd; border-radius: 8px; padding: 25px; font-family: 'Consolas', monospace; font-size: 36px; margin-bottom: 25px; line-height: 1.6; border: 1px solid #f0f0f0; }}
                .box-container {{ display: flex; gap: 8px; margin-top: 15px; }} .box {{ width: 60px; height: 30px; border-radius: 4px; }}
                .bg-red {{ background: #e74c3c; }} .bg-green {{ background: #2ecc71; }} .bg-gray {{ background: #ecf0f1; }}
                .text-red {{ color: #e74c3c; font-weight: bold; }} .text-green {{ color: #2ecc71; font-weight: bold; }} .text-yellow {{ color: #f39c12; font-weight: bold; }}
                
                .fund-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; }}
                .fund-box {{ background: #fdfdfd; border: 1px solid #eee; border-radius: 12px; padding: 25px 10px; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; box-shadow: 0 2px 10px rgba(0,0,0,0.01); height: 160px; }}
                .fund-label {{ font-size: 28px; color: #7f8c8d; font-weight: 700; margin-bottom: 15px; }}
                .fund-val {{ font-size: 40px; font-weight: 900; color: #2c3e50; }}

                .footer {{ 
                    position: absolute; 
                    bottom: 40px; 
                    left: 100px; 
                    right: 100px; 
                    text-align: center; 
                    font-size: 32px; 
                    color: #95a5a6; 
                    padding-top: 25px; 
                    border-top: 2px solid #eee; 
                    line-height: 1.5; 
                }}
            </style></head><body>
            <div class="a4-container">
                
                <div class="header">
                    <div class="header-left">
                        <h1>{stock_id} {name}</h1>
                        <p>個股綜合健檢報告 / Equity Health Report</p>
                    </div>
                    <div class="header-right">
                        <p>報表日期: {date_str}</p>
                        <p>市場: {market} | 幣別: TWD</p>
                    </div>
                </div>
                
                <div class="metrics-grid">
                    <div class="metric-card">
                        <p class="m-title">綜合評分 / Score</p>
                        <p class="m-value">{score}</p>
                    </div>
                    <div class="metric-card">
                        <p class="m-title">戰略建議 / Strategy</p>
                        <p class="m-value" style="color: {color};">{sug}</p>
                    </div>
                </div>
                
                <div class="chart-wrapper">
                    <div class="legend-container">
                        <div class="legend-item"><span class="l-dot" style="color:#ff6b6b;">●</span> 5MA</div>
                        <div class="legend-item"><span class="l-dot" style="color:#3498db;">●</span> 10MA</div>
                        <div class="legend-item"><span class="l-dot" style="color:#fbc531;">●</span> 20MA</div>
                        <div class="legend-item"><span class="l-dot" style="color:#4cd137;">●</span> 60MA</div>
                        <div class="legend-item"><span class="l-dot" style="color:#e67e22;">●</span> POC鐵板</div>
                    </div>
                    <div class="chart-container">
                        {'<img src="data:image/png;base64,'+b64_chart+'">' if b64_chart else '<div style="font-size:40px; color:#bdc3c7;">K線圖生成失敗</div>'}
                    </div>
                </div>

                <div class="bottom-grid">
                    <div class="info-card" style="border-top: 15px solid #3498db;">
                        <h2>📊 技術指標詳情 / Technicals</h2>
                        <table class="data-table">
                            <tr><td class="label">昨日收盤價</td><td class="value">{d['Close']:.2f} 元</td></tr>
                            <tr><td class="label">🎯 實戰操作區間</td><td class="value">停利 <span class="text-red">{tp_price:.2f}</span> / 停損 <span class="text-green">{sl_price:.2f}</span></td></tr>
                            <tr><td class="label">均線趨勢格局</td><td class="value">{info['trend']}</td></tr>
                            <tr><td class="label">最強鐵板區 (POC)</td><td class="value"><span style="color: #e67e22;">{poc_str}</span></td></tr>
                            <tr><td class="label">壓力(上軌) / 支撐(月線)</td><td class="value">{ub_str} / {ma20_str}</td></tr>
                            <tr><td class="label">RSI 強弱指標 (14)</td><td class="value">{rsi_str} <span style="font-size:30px; color:#95a5a6; font-weight:normal;">(相對強弱)</span></td></tr>
                            <tr><td class="label">KD 指標 (K / D)</td><td class="value">{k_str} / {d_str} <span style="font-size:30px; color:#95a5a6; font-weight:normal;">({info['kd'].replace('⚠️ ','').replace('❄️ ','').replace('✨ ','').replace('💀 ','')})</span></td></tr>
                            <tr><td class="label">昨日成交量能</td><td class="value">{int(d['Volume']/1000):,} 張 <span style="font-size:30px; color:#95a5a6; font-weight:normal;">({info['vol'].replace('🚨 ','').replace('🔺 ','').replace('💤 ','')})</span></td></tr>
                        </table>
                    </div>
                    
                    <div class="info-card" style="border-top: 15px solid #e67e22;">
                        <h2>🏦 籌碼與資券分析 / Chips & Margin</h2>
                        <div class="code-box">
                            <div style="color:#7f8c8d; margin-bottom:10px; font-size:30px;">籌碼集中動能: <span style="font-size:36px;">{format_html_color(f"{c_val_pct:.2f}%")}</span></div>
                            <div class="box-container">{boxes_html}</div>
                        </div>
                        <div class="code-box">
                            <div style="color:#7f8c8d; margin-bottom:10px; font-size:30px;">股權分散 (週增減) - 更新: {curr[0] if curr else '-'}</div>
                            <div>外資籌碼: {curr[1] if curr else '-'} {f"({format_html_color(f_diff)})" if f_diff else ""}</div>
                            <div>大戶籌碼: {curr[2] if curr else '-'} {f"({format_html_color(l_diff)})" if l_diff else ""}</div>
                        </div>
                        <div class="code-box">
                            <div style="color:#7f8c8d; margin-bottom:10px; font-size:30px;">信用交易 - 更新: {margin_data.get('date', '-')}</div>
                            <div>融資增減: {format_html_color(margin_data.get('m_diff','-'))} | 券資比: <span class="text-yellow">{margin_data.get('ratio','-')}</span></div>
                        </div>
                    </div>
                </div>
                
                <div class="info-card" style="border-top: 15px solid #9b59b6; margin-bottom: 0px; padding: 30px;">
                    <h2 style="margin-bottom: 20px;">💰 財務與基本面健檢 / Fundamentals</h2>
                    <div class="fund-grid">
                        <div class="fund-box">
                            <div class="fund-label">最新本益比 (P/E)</div>
                            <div class="fund-val">{fund_data['pe']}</div>
                        </div>
                        <div class="fund-box">
                            <div class="fund-label">股價淨值比 (P/B)</div>
                            <div class="fund-val">{fund_data['pb']}</div>
                        </div>
                        <div class="fund-box">
                            <div class="fund-label">預估殖利率</div>
                            <div class="fund-val">{fund_data['yield']}</div>
                        </div>
                        <div class="fund-box">
                            <div class="fund-label">最新單季 EPS ({fund_data['eps_q']})</div>
                            <div class="fund-val">{format_html_color(fund_data['eps'])}</div>
                        </div>
                        <div class="fund-box">
                            <div class="fund-label">最新營收 ({fund_data['rev_m']})</div>
                            <div class="fund-val" style="font-size: 34px;">{fund_data['rev']} <span style="font-size:22px;color:#aaa;">千元</span></div>
                        </div>
                        <div class="fund-box">
                            <div class="fund-label">月營收月增 (MoM)</div>
                            <div class="fund-val">{format_html_color(fund_data['mom'])}</div>
                        </div>
                        <div class="fund-box">
                            <div class="fund-label">月營收年增 (YoY)</div>
                            <div class="fund-val">{format_html_color(fund_data['yoy'])}</div>
                        </div>
                        <div class="fund-box" style="background: #faf9fc; border-color: #eadaf5;">
                            <div class="fund-label" style="color: #9b59b6;">💡 基本面 AI 評估</div>
                            <div class="fund-val" style="font-size: 28px; font-weight: 700; line-height: 1.4;">{fund_ai}</div>
                        </div>
                    </div>
                </div>

                <div class="footer">
                   <strong>⚠️ 免責聲明 (Disclaimer)：</strong>本報告由 AI 自動程式抓取與生成，所有技術指標與財務數據僅供學術研究與參考，不構成任何投資買賣建議。<br>
                    金融市場瞬息萬變，投資一定有風險，投資人應獨立判斷並自負盈虧。資料來源：Yahoo Finance Taiwan
                </div>
            </div>
            </body></html>
            """
            file_name = f"report_{stock_id}.png"
            # 必須加上這三個 custom_flags，Railway 上的 Chrome 才不會崩潰
            hti = Html2Image(output_path='.', custom_flags=['--no-sandbox', '--disable-gpu', '--disable-dev-shm-usage'])
            await asyncio.to_thread(hti.screenshot, html_str=html, save_as=file_name, size=(2480, 3508))
            await loading.edit(content="✅ 不朽金剛版 (雙引擎爬蟲 + 絕對排版) 生成完畢！")
            await ctx.send(file=discord.File(file_name))
            os.remove(file_name) 
        except Exception as e:
            await loading.edit(content=f"❌ 報告生成失敗: {e}")
    else:
        await loading.edit(content="❌ 伺服器未安裝圖片生成模組。")

@bot.event
async def on_ready():
    print(f'🔥 A4 終極無敵雙引擎版 已上線！')


bot.run(TOKEN)
