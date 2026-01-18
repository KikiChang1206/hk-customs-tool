import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime, timedelta

# 1. 網頁設定
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

# 2. 整合式上傳區
uploaded_files = st.file_uploader("請拖入所有 4 個必要檔案", type=['xls', 'xlsx'], accept_multiple_files=True)

files_dict = {"Invoice": None, "Packing": None, "北方文件": None, "OrderList": None}

if uploaded_files:
    for f in uploaded_files:
        fname = f.name.lower()
        if "invoice" in fname: files_dict["Invoice"] = f
        elif "packing" in fname: files_dict["Packing"] = f
        elif "manifest" in fname or "北方" in fname: files_dict["北方文件"] = f
        elif "order" in fname: files_dict["OrderList"] = f

# 3. 狀態顯示
st.write("### 📋 檔案讀取狀態")
c1, c2 = st.columns(2)
with c1:
    st.markdown(f"{'✅' if files_dict['Invoice'] else '❌'} **Invoice**")
    st.markdown(f"{'✅' if files_dict['Packing'] else '❌'} **Packing**")
with c2:
    st.markdown(f"{'✅' if files_dict['北方文件'] else '❌'} **北方文件**")
    st.markdown(f"{'✅' if files_dict['OrderList'] else '❌'} **Order List**")

# 4. 轉換邏輯
if all(files_dict.values()):
    st.write("---")
    if st.button("🚀 執行修正版轉換", use_container_width=True):
        try:
            with st.spinner('正在調整格式與修正重量公式...'):
                tw_now = datetime.utcnow() + timedelta(hours=8)
                t_str = tw_now.strftime("%Y%m%d")

                def smart_read_excel(file_obj, **kwargs):
                    if file_obj.name.endswith('.xls'): return pd.read_excel(file_obj, engine='xlrd', **kwargs)
                    else: return pd.read_excel(file_obj, engine='openpyxl', **kwargs)

                # 讀取數據
                df_order = smart_read_excel(files_dict["OrderList"], dtype=str).fillna('')
                df_n_export = smart_read_excel(files_dict["北方文件"], sheet_name='出口明細', dtype=str).fillna('')
                df_n_bag = smart_read_excel(files_dict["北方文件"], sheet_name='袋數編號', dtype=str).fillna('')
                
                # 建立字典
                bag_dict = df_n_export.set_index(df_n_export.columns[1])[df_n_export.columns[6]].to_dict()
                barcode_dict = df_n_bag.set_index(df_n_bag.columns[0])[df_n_bag.columns[1]].to_dict()

                wb = Workbook()
                ws = wb.active
                ws.title = "HK最終報關檔"

                # A. 搬運並處理合併單元格 (1-10行)
                df_inv_head = smart_read_excel(files_dict["Invoice"], header=None, nrows=10, dtype=str).fillna('')
                for r_idx, row_data in enumerate(df_inv_head.values, 1):
                    for c_idx, value in enumerate(row_data, 1):
                        ws.cell(row=r_idx, column=c_idx, value=value).font = Font(name='Arial', size=10)
                
                # 執行指定的合併需求
                merge_list = [
                    "B1:D1", "B2:I2", "B3:E3", "F3:I3", "B4:I4", "B5:E5", "F5:I5", 
                    "B6:I6", "B7:E7", "F7:I7", "B8:E8", "F8:I8", "B9:D9", "E9:G9", 
                    "H9:I9", "B10:E10", "F10:I10"
                ]
                for area in merge_list:
                    ws.merge_cells(area)

                # B. 寫入 FOB (黃底)
                ws['A11'] = "FOB"
                ws['A11'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                ws['A11'].font = Font(bold=True)

                # C. 寫入項次與標題 (A13-Q13)
                ws['A13'] = "項次"
                headers = ["提單編號", "訂單編號", "好馬吉袋號", "條碼", "單箱重量(GW)", "品項淨重", 
                           "品項英文名稱", "品項中文名稱", "品項備註", "品項品牌", "品項產地", 
                           "品項數量", "單位", "品項單價", "品項小計", "幣別"]
                
                green_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                # 設定 A13 格式
                ws['A13'].fill = green_fill
                ws['A13'].font = Font(bold=True, name='Arial', size=10)
                ws['A13'].alignment = Alignment(horizontal='center')
                ws['A13'].border = thin_border

                for i, title in enumerate(headers, 2): 
                    cell = ws.cell(row=13, column=i, value=title)
                    cell.fill = green_fill
                    cell.font = Font(bold=True, name='Arial', size=10)
                    cell.alignment = Alignment(horizontal='center')
                    cell.border = thin_border

                # D. 明細處理 (14行起)
                prev_hawb = None
                curr_row = 14
                item_no = 1  # 項次計數

                for index, r in df_order.iterrows():
                    # 1. 寫入項次 (A欄)
                    ws.cell(row=curr_row, column=1, value=item_no).border = thin_border
                    
                    hawb = str(r.iloc[1]).strip() # 提單編號 (B)
                    oid = str(r.iloc[3]).strip()  # 訂單編號 (D)
                    bag_no = bag_dict.get(hawb, "")
                    barcode = barcode_dict.get(oid, "") # 依公式 D14 抓條碼

                    # 2. 單箱重量修正 (FOB邏輯)
                    gw_raw = r.iloc[29] # AE欄
                    gw_display = ""
                    # 邏輯：如果當前 HAWB 與前一個相同，則顯示空值
                    if hawb != prev_hawb:
                        gw_display = gw_raw
                    
                    # 3. 品項淨重修正 (NW = GW - 0.2, 最小 0.01)
                    nw_display = ""
                    if gw_display != "":
                        try:
                            calc_nw = float(gw_display) - 0.2
                            nw_display = calc_nw if calc_nw > 0 else 0.01
                            nw_display = "{:.2f}".format(nw_display)
                        except:
                            nw_display = ""

                    data = [
                        hawb, oid, bag_no, barcode, gw_display, nw_display,
                        "COSMETICS", r.iloc[33], r.iloc[34], "TRUU+TRUE YOU", 
                        r.iloc[36], r.iloc[37], "SET", r.iloc[39], r.iloc[40], "TWD"
                    ]

                    for col_idx, val in enumerate(data, 2):
                        cell = ws.cell(row=curr_row, column=col_idx, value=val)
                        cell.font = Font(name='Arial', size=10)
                        cell.border = thin_border
                    
                    prev_hawb = hawb
                    curr_row += 1
                    item_no += 1

                output = BytesIO()
                wb.save(output)
                st.balloons()
                st.success("✅ 修正版轉換成功！")
                st.download_button(label="📥 下載修正版 HK 報關文件", data=output.getvalue(), file_name=f"{t_str}_HK_GM_Final_Fixed.xlsx", use_container_width=True)

        except Exception as e:
            st.error(f"錯誤：{e}")
