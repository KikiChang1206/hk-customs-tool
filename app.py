import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
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

# 2. 整合式上傳區
st.write("### 📤 請一次拖入 4 個必要檔案")
uploaded_files = st.file_uploader("可多選或一次拖入所有檔案", type=['xls', 'xlsx'], accept_multiple_files=True)

files_dict = {"Invoice": None, "Packing": None, "北方文件": None, "OrderList": None}

# 3. 智能辨識邏輯 (包含 OrderList 修正)
if uploaded_files:
    for f in uploaded_files:
        name = f.name
        if "MergeInvoice" in name: files_dict["Invoice"] = f
        elif "MergePackingList" in name: files_dict["Packing"] = f
        elif "Manifest" in name or "北方" in name: files_dict["北方文件"] = f
        elif "OrderList" in name or "Order List" in name: files_dict["OrderList"] = f

# 4. 狀態打勾顯示
st.write("---")
st.write("### 📋 檔案讀取狀態")
c1, c2 = st.columns(2)
with c1:
    st.markdown(f"{'✅' if files_dict['Invoice'] else '❌'} **Invoice** (含 MergeInvoice)")
    st.markdown(f"{'✅' if files_dict['Packing'] else '❌'} **Packing** (含 MergePackingList)")
with c2:
    st.markdown(f"{'✅' if files_dict['北方文件'] else '❌'} **北方文件** (含 北方/Manifest)")
    st.markdown(f"{'✅' if files_dict['OrderList'] else '❌'} **Order List** (含 OrderList)")

# 5. 轉換邏輯
if all(files_dict.values()):
    st.write("---")
    # Session State 保持按鈕狀態
    if 'hk_processed' not in st.session_state: st.session_state.hk_processed = False

    if st.button("🚀 開始執行轉換", use_container_width=True) or st.session_state.hk_processed:
        try:
            with st.spinner('正在比對北方文件頁籤與 OrderList 資料...'):
                tw_now = datetime.utcnow() + timedelta(hours=8)
                t_str = tw_now.strftime("%Y%m%d")

                # A. 讀取數據
                df_order = pd.read_excel(files_dict["OrderList"], dtype=str).fillna('')
                df_n_export = pd.read_excel(files_dict["北方文件"], sheet_name='出口明細', dtype=str).fillna('')
                df_n_bag = pd.read_excel(files_dict["北方文件"], sheet_name='袋數編號', dtype=str).fillna('')
                
                # B. VLOOKUP 字典
                bag_dict = df_n_export.set_index(df_n_export.columns[1])[df_n_export.columns[6]].to_dict()
                barcode_dict = df_n_bag.set_index(df_n_bag.columns[0])[df_n_bag.columns[1]].to_dict()

                # C. 建立活頁簿
                wb = Workbook()
                ws = wb.active
                ws.title = "HK最終報關檔"

                # D. 搬運 Invoice 表頭
                df_inv_head = pd.read_excel(files_dict["Invoice"], header=None, nrows=10, dtype=str).fillna('')
                for r_idx, row_data in enumerate(df_inv_head.values, 1):
                    for c_idx, value in enumerate(row_data, 1):
                        ws.cell(row=r_idx, column=c_idx, value=value).font = Font(name='Arial', size=10)

                # E. 寫入 FOB (黃底)
                ws['A11'] = "FOB"
                ws['A11'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                ws['A11'].font = Font(bold=True)

                # F. 寫入綠底標題
                headers = ["提單編號", "訂單編號", "好馬吉袋號", "條碼", "單箱重量(GW)", "品項淨重", 
                           "品項英文名稱", "品項中文名稱", "品項備註", "品項品牌", "品項產地", 
                           "品項數量", "單位", "品項單價", "品項小計", "幣別"]
                
                green_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
                for i, title in enumerate(headers, 2): 
                    cell = ws.cell(row=13, column=i, value=title)
                    cell.fill = green_fill
                    cell.font = Font(bold=True, name='Arial', size=10)
                    cell.alignment = Alignment(horizontal='center')

                # G. 明細處理
                prev_hawb = None
                curr_row = 14
                for _, r in df_order.iterrows():
                    hawb = str(r.iloc[1]).strip()
                    oid = str(r.iloc[3]).strip()
                    bag_no = bag_dict.get(hawb, "")
                    barcode = barcode_dict.get(bag_no, "")

                    gw = r.iloc[29] if hawb != prev_hawb else ""
                    nw = "{:.2f}".format(float(gw) - 0.2) if gw != "" else ""

                    data = [
                        hawb, oid, bag_no, barcode, gw, nw,
                        "COSMETICS", r.iloc[33], r.iloc[34], 
                        "TRUU+TRUE YOU", r.iloc[36], r.iloc[37], 
                        "SET", r.iloc[39], r.iloc[40], "TWD"
                    ]

                    for col_idx, val in enumerate(data, 2):
                        ws.cell(row=curr_row, column=col_idx, value=val).font = Font(name='Arial', size=10)
                    
                    prev_hawb = hawb
                    curr_row += 1

                # H. 檔案產出
                output = BytesIO()
                wb.save(output)
                
                if not st.session_state.hk_processed:
                    st.balloons()
                    st.session_state.hk_processed = True
                    
                st.success("✅ 轉換成功！")
                st.download_button(
                    label="📥 下載 HK 報關最終文件",
                    data=output.getvalue(),
                    file_name=f"{t_str}_HK_GM_Final.xlsx",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"錯誤：{e}")
else:
    if uploaded_files:
        st.warning("請確認 4 個檔案皆已上傳 (目前仍有項目顯示 ❌)")
