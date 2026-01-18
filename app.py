import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime, timedelta

# 1. 網頁基本設定 (黑底風格)
st.set_page_config(page_title="HK報關文件轉換器", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #0E1117; }
    .big-title { font-size: 30px !important; font-weight: bold; color: #FFFFFF !important; }
    .stFileUploader section { background-color: #FFFFFF !important; border-radius: 10px; }
    div.stButton > button { background-color: #FFFFFF !important; color: #000000 !important; border: 2px solid #000000 !important; height: 50px; font-weight: bold; width: 100%; }
    .stMarkdown p, label { color: #FFFFFF !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<p class="big-title">🇭🇰 HK 報關文件轉換器</p>', unsafe_allow_html=True)

# 2. 檔案上傳區
st.write("### 📤 請上傳 4 個必要檔案")
c1, c2 = st.columns(2)
with c1:
    inv_f = st.file_uploader("1. 上傳 Invoice 原始檔", type=['xls', 'xlsx'])
    # 特別標註北方文件的名稱與頁籤
    north_f = st.file_uploader("2. 上傳 北方文件 (北方_XXXX_HK_Manifest)", type=['xls', 'xlsx'])
with c2:
    pac_f = st.file_uploader("3. 上傳 Packing 原始檔", type=['xls', 'xlsx'])
    order_f = st.file_uploader("4. 上傳 Order List 檔案", type=['xls', 'xlsx'])

# 3. 執行邏輯
if inv_f and north_f and pac_f and order_f:
    if st.button("🚀 執行轉換並產出 HK 報關文件", use_container_width=True):
        try:
            with st.spinner('正在讀取北方文件頁籤並計算明細...'):
                # 台灣日期修正
                tw_now = datetime.utcnow() + timedelta(hours=8)
                t_str = tw_now.strftime("%Y%m%d")

                # A. 讀取資料
                # Order List
                df_order = pd.read_excel(order_f, dtype=str).fillna('')
                
                # --- 北方文件：修正頁籤名稱 ---
                df_n_export = pd.read_excel(north_f, sheet_name='出口明細', dtype=str).fillna('')
                df_n_bag = pd.read_excel(north_f, sheet_name='袋數編號', dtype=str).fillna('')
                
                # B. 建立比對字典 (VLOOKUP 核心)
                # 1. 好馬吉袋號: 從 B 欄 (HAWB) 找 G 欄 (BAG_N)
                # 根據截圖，G 欄索引為 6
                bag_dict = df_n_export.set_index(df_n_export.columns[1])[df_n_export.columns[6]].to_dict()
                
                # 2. 條碼: 從 A 欄 (BAG_NO) 找 B 欄 (REF_BAG_NO)
                barcode_dict = df_n_bag.set_index(df_n_bag.columns[0])[df_n_bag.columns[1]].to_dict()

                # C. 建立最終 Excel
                wb = Workbook()
                ws = wb.active
                ws.title = "HK最終報關檔"

                # D. 搬運 Invoice 表頭 (1-10行)
                src_wb = load_workbook(inv_f)
                src_ws = src_wb.active
                for r in range(1, 11):
                    for c in range(1, 11): 
                        val = src_ws.cell(row=r, column=c).value
                        ws.cell(row=r, column=c, value=val)
                        ws.cell(row=r, column=c).font = Font(name='Arial', size=10)

                # E. 寫入第11行 FOB (黃底)
                ws['A11'] = "FOB"
                ws['A11'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                ws['A11'].font = Font(bold=True)

                # F. 寫入第13行 綠底標題 (B13-Q13)
                headers = ["提單編號", "訂單編號", "好馬吉袋號", "條碼", "單箱重量(GW)", "品項淨重", 
                           "品項英文名稱", "品項中文名稱", "品項備註", "品項品牌", "品項產地", 
                           "品項數量", "單位", "品項單價", "品項小計", "幣別"]
                
                green_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
                for i, title in enumerate(headers, 2): 
                    cell = ws.cell(row=13, column=i, value=title)
                    cell.fill = green_fill
                    cell.font = Font(bold=True, name='Arial', size=10)
                    cell.alignment = Alignment(horizontal='center')

                # G. 寫入明細資料
                prev_hawb = None
                curr_row = 14

                for _, r in df_order.iterrows():
                    hawb = str(r.iloc[1]).strip() # Order List B欄
                    oid = str(r.iloc[3]).strip()  # Order List D欄
                    
                    # 串聯邏輯：提單號碼 -> 袋號 -> 條碼
                    bag_no = bag_dict.get(hawb, "")
                    barcode = barcode_dict.get(bag_no, "")

                    # GW 與 NW
                    gw = ""
                    if hawb != prev_hawb:
                        gw = r.iloc[29] # AE欄 (第30欄)
                    
                    nw = ""
                    try:
                        if gw != "": nw = "{:.2f}".format(float(gw) - 0.2)
                    except: nw = ""

                    # 組合 B 到 Q 欄資料
                    data = [
                        hawb, oid, bag_no, barcode, gw, nw,
                        "COSMETICS", r.iloc[33], r.iloc[34], # AH, AI
                        "TRUU+TRUE YOU", r.iloc[36], r.iloc[37], # AK, AL
                        "SET", r.iloc[39], r.iloc[40], "TWD" # AN, AO
                    ]

                    for col_idx, val in enumerate(data, 2):
                        cell = ws.cell(row=curr_row, column=col_idx, value=val)
                        cell.font = Font(name='Arial', size=10)
                        cell.alignment = Alignment(horizontal='left')
                    
                    prev_hawb = hawb
                    curr_row += 1

                # H. 下載產出
                output = BytesIO()
                wb.save(output)
                
                st.balloons()
                st.success("🎉 HK 文件轉換成功！")
                st.download_button(
                    label="📥 下載 HK 報關最終文件",
                    data=output.getvalue(),
                    file_name=f"{t_str}_HK_Customs_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"轉換出錯！錯誤訊息: {e}")
            st.info("💡 請確認上傳的檔案順序與頁籤是否符合規範。")
