import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

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

# 2. 檔案上傳
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
    if st.button("🚀 執行最終格式化轉換", use_container_width=True):
        try:
            from datetime import datetime, timedelta
            tw_now = datetime.utcnow() + timedelta(hours=8)
            t_str = tw_now.strftime("%Y%m%d")

            def smart_read_excel(file_obj, **kwargs):
                if file_obj.name.endswith('.xls'): return pd.read_excel(file_obj, engine='xlrd', **kwargs)
                else: return pd.read_excel(file_obj, engine='openpyxl', **kwargs)

            df_order = smart_read_excel(files_dict["OrderList"], dtype=str).fillna('')
            df_n_export = smart_read_excel(files_dict["北方文件"], sheet_name='出口明細', dtype=str).fillna('')
            df_n_bag = smart_read_excel(files_dict["北方文件"], sheet_name='袋數編號', dtype=str).fillna('')
            df_inv_raw = smart_read_excel(files_dict["Invoice"], header=None, dtype=str).fillna('')

            def get_inv(cell_ref):
                col_map = {'A':0, 'B':1, 'C':2, 'D':3, 'E':4, 'F':5, 'G':6, 'H':7, 'I':8}
                c = col_map[cell_ref[0]]
                r = int(cell_ref[1:]) - 1
                try: return df_inv_raw.iloc[r, c]
                except: return ""

            wb = Workbook()
            ws = wb.active
            ws.title = "HK最終報關檔"

            # A. 欄寬設定
            column_widths = {
                'B': 18.64, 'C': 17.27, 'D': 12.64, 'E': 12.09, 'F': 12.64,
                'G': 8.09, 'H': 11.64, 'I': 51.82, 'J': 30, 'K': 15.82,
                'L': 8.09, 'M': 8.09, 'N': 8.09, 'O': 10.91, 'P': 7.91, 'Q': 8.09
            }
            for col, width in column_widths.items():
                ws.column_dimensions[col].width = width

            # B. 表頭填充與合併
            head_configs = [
                ("B1", "INVOICE/PACKING", "B1:D1", 28, True, False),
                ("B2", get_inv("A2"), "B2:I2", 10, False, True), # 自動換行
                ("B3", get_inv("A3"), "B3:E3", 10, False, False),
                ("F3", get_inv("E3"), "F3:I3", 10, False, False),
                ("B4", get_inv("A4"), "B4:I4", 10, False, False),
                ("B5", get_inv("A5"), "B5:E5", 10, False, False),
                ("F5", get_inv("E5"), "F5:I5", 10, False, False),
                ("B6", get_inv("A6"), "B6:I6", 10, False, False),
                ("B7", get_inv("A7"), "B7:E7", 10, False, False),
                ("F7", get_inv("E7"), "F7:I7", 10, False, True),  # 自動換行
                ("B8", get_inv("A8"), "B8:E8", 10, False, True),  # 自動換行
                ("F8", get_inv("E8"), "F8:I8", 10, False, False),
                ("B9", get_inv("A9"), "B9:D9", 10, False, False),
                ("E9", get_inv("D9"), "E9:G9", 10, False, False),
                ("H9", get_inv("G9"), "H9:I9", 10, False, False),
                ("B10", get_inv("A10"), "B10:E10", 10, False, False),
                ("F10", get_inv("E10"), "F10:I10", 10, False, False)
            ]

            for cell_id, content, merge_range, size, is_bold, is_wrap in head_configs:
                cell = ws[cell_id]
                cell.value = content
                cell.font = Font(name='Arial', size=size, bold=is_bold)
                cell.alignment = Alignment(wrap_text=is_wrap, vertical='center')
                ws.merge_cells(merge_range)

            # C. FOB 移至 B11
            ws['B11'] = "FOB"
            ws['B11'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            ws['B11'].font = Font(name='Arial', size=10, bold=True)

            # D. 標題列 (A13-Q13)
            ws['A13'] = "項次"
            headers = ["提單編號", "訂單編號", "好馬吉袋號", "條碼", "單箱重量(GW)", "品項淨重", 
                       "品項英文名稱", "品項中文名稱", "品項備註", "品項品牌", "品項產地", 
                       "品項數量", "單位", "品項單價", "品項小計", "幣別"]
            green_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            ws['A13'].fill = green_fill
            ws['A13'].font = Font(name='Arial', size=10, bold=True)
            ws['A13'].border = thin_border
            ws['A13'].alignment = Alignment(horizontal='center', vertical='center')

            for i, title in enumerate(headers, 2): 
                cell = ws.cell(row=13, column=i, value=title)
                cell.fill = green_fill
                cell.font = Font(name='Arial', size=10, bold=True)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # E. 明細處理
            bag_dict = df_n_export.set_index(df_n_export.columns[1])[df_n_export.columns[6]].to_dict()
            barcode_dict = df_n_bag.set_index(df_n_bag.columns[0])[df_n_bag.columns[1]].to_dict()

            prev_hawb, curr_row, item_no = None, 14, 1
            for _, r in df_order.iterrows():
                ws.cell(row=curr_row, column=1, value=item_no).border = thin_border
                ws.cell(row=curr_row, column=1).font = Font(name='Arial', size=10)
                
                hawb, oid = str(r.iloc[1]).strip(), str(r.iloc[3]).strip()
                bag_no, barcode = bag_dict.get(hawb, ""), barcode_dict.get(oid, "")

                gw_raw, gw_display = r.iloc[29], ""
                if hawb != prev_hawb: gw_display = gw_raw
                
                nw_display = ""
                if gw_display != "":
                    try:
                        calc_nw = float(gw_display) - 0.2
                        nw_display = "{:.2f}".format(calc_nw if calc_nw > 0 else 0.01)
                    except: nw_display = ""

                row_content = [
                    hawb, oid, bag_no, barcode, gw_display, nw_display,
                    "COSMETICS", r.iloc[33], r.iloc[34], "TRUU+TRUE YOU", 
                    r.iloc[36], r.iloc[37], "SET", r.iloc[39], r.iloc[40], "TWD"
                ]

                for col_idx, val in enumerate(row_content, 2):
                    c = ws.cell(row=curr_row, column=col_idx, value=val)
                    c.font = Font(name='Arial', size=10)
                    c.border = thin_border
                    # I 欄與 J 欄自動換行 (對應索引 9 與 10)
                    if col_idx in [9, 10]:
                        c.alignment = Alignment(wrap_text=True, vertical='center')
                    else:
                        c.alignment = Alignment(vertical='center')

                prev_hawb, curr_row, item_no = hawb, curr_row + 1, item_no + 1

            output = BytesIO()
            wb.save(output)
            st.balloons()
            st.success("✅ 最終格式文件已完成！")
            st.download_button(label="📥 下載最終 HK 報關文件", data=output.getvalue(), file_name=f"{t_str}_HK_Customs_Final.xlsx", use_container_width=True)

        except Exception as e:
            st.error(f"錯誤：{e}")
