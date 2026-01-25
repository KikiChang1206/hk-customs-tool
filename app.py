import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime, timedelta

# 1. 網頁基本設定 (修正國旗文字)
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

# 國旗圖標已改為文字
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

# 3. 執行轉換邏輯
if all(files_dict.values()):
    if st.button("🚀 執行排序與格式優化轉換", use_container_width=True):
        try:
            tw_now = datetime.utcnow() + timedelta(hours=8)
            t_str = tw_now.strftime("%Y%m%d")

            def smart_read_excel(file_obj, **kwargs):
                if file_obj.name.endswith('.xls'): return pd.read_excel(file_obj, engine='xlrd', **kwargs)
                else: return pd.read_excel(file_obj, engine='openpyxl', **kwargs)

            # A. 讀取數據
            df_order = smart_read_excel(files_dict["OrderList"], dtype=str).fillna('')
            df_n_export = smart_read_excel(files_dict["北方文件"], sheet_name='出口明細', dtype=str).fillna('')
            df_n_bag_raw = smart_read_excel(files_dict["北方文件"], sheet_name='袋數編號', dtype=str)
            bag_count = len(df_n_bag_raw[df_n_bag_raw.iloc[:, 1].str.strip() != ""])
            df_n_bag = df_n_bag_raw.fillna('')
            df_inv_raw = smart_read_excel(files_dict["Invoice"], header=None, dtype=str).fillna('')
            df_packing_raw = smart_read_excel(files_dict["Packing"], header=None, dtype=str).fillna('')

            def get_inv(cell_ref):
                col_map = {'A':0, 'B':1, 'C':2, 'D':3, 'E':4, 'F':5, 'G':6, 'H':7, 'I':8}
                c = col_map[cell_ref[0]]
                r = int(cell_ref[1:]) - 1
                try: return df_inv_raw.iloc[r, c]
                except: return ""

            wb = Workbook()
            ws = wb.active
            ws.title = "HK最終報關檔"

            # --- 1. 欄寬與行高設定 (修正 E 欄寬為 14) ---
            col_widths = {'B': 20.8, 'C': 19.2, 'D': 14.7, 'E': 14, 'F': 14, 'G': 8.7, 'H': 13, 'I': 51.82, 'J': 30, 'K': 17.9, 'L': 8.7, 'M': 8.7, 'N': 8.09, 'O': 10.91, 'P': 9, 'Q': 8.09}
            for col, width in col_widths.items(): ws.column_dimensions[col].width = width
            ws.row_dimensions[1].height = 77
            ws.row_dimensions[2].height = 25.2
            for r in range(3, 7): ws.row_dimensions[r].height = 12.5
            ws.row_dimensions[7].height = 49.5
            ws.row_dimensions[8].height = 25.2
            for r in range(9, 13): ws.row_dimensions[r].height = 12.5

            # --- 2. 表頭與 FOB ---
            ws["B1"] = "INVOICE/PACKING"
            ws["B1"].font = Font(name='Arial', size=28, bold=True)
            ws["B1"].alignment = Alignment(horizontal='left', vertical='center')
            ws.merge_cells("B1:E1")
            
            head_configs = [("B2", get_inv("A2"), "B2:I2", 10, False, True), ("B3", get_inv("A3"), "B3:E3", 10, False, False), ("F3", get_inv("E3"), "F3:I3", 10, False, False), ("B4", get_inv("A4"), "B4:I4", 10, False, False), ("B5", get_inv("A5"), "B5:E5", 10, False, False), ("F5", get_inv("E5"), "F5:I5", 10, False, False), ("B6", get_inv("A6"), "B6:I6", 10, False, False), ("B7", get_inv("A7"), "B7:E7", 10, False, False), ("F7", get_inv("E7"), "F7:I7", 10, False, True), ("B8", get_inv("A8"), "B8:E8", 10, False, True), ("F8", get_inv("E8"), "F8:I8", 10, False, False), ("B9", get_inv("A9"), "B9:D9", 10, False, False), ("E9", get_inv("D9"), "E9:G9", 10, False, False), ("H9", get_inv("G9"), "H9:I9", 10, False, False), ("B10", get_inv("A10"), "B10:E10", 10, False, False), ("F10", get_inv("E10"), "F10:I10", 10, False, False)]
            for cell_id, content, merge_range, size, is_bold, is_wrap in head_configs:
                ws[cell_id] = content
                ws[cell_id].font = Font(name='Arial', size=size, bold=is_bold)
                ws[cell_id].alignment = Alignment(wrap_text=is_wrap, vertical='center')
                ws.merge_cells(merge_range)

            ws['B11'] = "FOB"
            ws['B11'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            ws['B11'].font = Font(name='Arial', size=10, bold=True)

            # --- 3. 標題列 (A13項次空白) ---
            ws['A13'] = ""
            headers = ["提單編號", "訂單編號", "好馬吉袋號", "條碼", "單箱重量(GW)", "品項淨重", "品項英文名稱", "品項中文名稱", "品項備註", "品項品牌", "品項產地", "品項數量", "單位", "品項單價", "品項小計", "幣別"]
            green_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            for i, title in enumerate(headers, 2):
                cell = ws.cell(row=13, column=i, value=title)
                cell.fill = green_fill
                cell.font = Font(name='Arial', size=10, bold=True)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # --- 4. 準備排序數據 ---
            barcode_lookup_dict = df_n_bag.set_index(df_n_bag.columns[0])[df_n_bag.columns[1]].to_dict()
            bag_dict = df_n_export.set_index(df_n_export.columns[1])[df_n_export.columns[6]].to_dict()

            all_data_rows = []
            for _, r in df_order.iterrows():
                hawb, oid = str(r.iloc[1]).strip(), str(r.iloc[3]).strip()
                bag_no = bag_dict.get(hawb, "")
                barcode = barcode_lookup_dict.get(bag_no, "")
                gw_raw = r.iloc[30] # AE 欄
                
                # 用於計算與排序的原始數值
                try: gw_num = float(gw_raw)
                except: gw_num = 0.0

                all_data_rows.append({
                    "hawb": hawb,
                    "oid": oid,
                    "bag_no": bag_no,
                    "barcode": barcode,
                    "gw_raw": gw_raw,
                    "gw_num": gw_num,
                    "r_data": r
                })

            # --- 5. 排序邏輯 ---
            # 優先按 條碼(E欄)、次按 提單編號(B欄)、最後按 單箱重量(GW, F欄)
            all_data_rows.sort(key=lambda x: (x["barcode"], x["hawb"], x["gw_num"]))

            # --- 6. 資料處理與填寫 ---
            prev_hawb, curr_row, item_no = None, 14, 1
            sum_gw = sum_nw = sum_qty = sum_amount = 0.0

            for item in all_data_rows:
                ws.cell(row=curr_row, column=1, value=item_no).font = Font(name='Arial', size=10)
                ws.cell(row=curr_row, column=1).alignment = Alignment(horizontal='center', vertical='center')
                
                hawb, oid, bag_no, barcode, gw_raw = item["hawb"], item["oid"], item["bag_no"], item["barcode"], item["gw_raw"]
                r = item["r_data"]

                gw_display = nw_display = ""
                if hawb != prev_hawb:
                    gw_display = gw_raw
                    try:
                        val_gw = float(gw_raw)
                        sum_gw += val_gw
                        calc_nw = val_gw - 0.2
                        final_nw = calc_nw if calc_nw >= 0.01 else 0.01
                        nw_display = "{:.2f}".format(final_nw)
                        sum_nw += final_nw
                    except: pass
                
                try: sum_qty += float(r.iloc[37]) 
                except: pass
                try: sum_amount += float(r.iloc[40]) 
                except: pass

                row_content = [hawb, oid, bag_no, barcode, gw_display, nw_display, "COSMETICS", r.iloc[33], r.iloc[34], "TRUU+TRUE YOU", r.iloc[36], r.iloc[37], "SET", r.iloc[39], r.iloc[40], "TWD"]
                for col_idx, val in enumerate(row_content, 2):
                    c = ws.cell(row=curr_row, column=col_idx, value=val)
                    c.font = Font(name='Arial', size=10)
                    c.border = thin_border
                    c.alignment = Alignment(wrap_text=(col_idx in [9, 10]), vertical='center')

                prev_hawb, curr_row, item_no = hawb, curr_row + 1, item_no + 1

            # --- 7. 最後補充欄位 (總結列，不參與排序) ---
            packing_last_val = df_packing_raw.iloc[-1, 0] if not df_packing_raw.empty else ""
            ws.cell(row=curr_row, column=2, value=packing_last_val).font = Font(name='Arial', size=10, bold=True)
            ws.cell(row=curr_row, column=5, value=f"包{bag_count}袋").font = Font(name='Arial', size=10, bold=True)
            ws.cell(row=curr_row, column=6, value="{:.2f}".format(sum_gw)).font = Font(name='Arial', size=10, bold=True)
            ws.cell(row=curr_row, column=7, value="{:.2f}".format(sum_nw)).font = Font(name='Arial', size=10, bold=True)
            ws.cell(row=curr_row, column=13, value=sum_qty).font = Font(name='Arial', size=10, bold=True)
            ws.cell(row=curr_row, column=14, value="SET").font = Font(name='Arial', size=10, bold=True)
            ws.cell(row=curr_row, column=16, value="{:.2f}".format(sum_amount)).font = Font(name='Arial', size=10, bold=True)
            ws.cell(row=curr_row, column=17, value="TWD").font = Font(name='Arial', size=10, bold=True)

            output = BytesIO()
            wb.save(output)
            st.balloons()
            st.success("✅ 資料排序與彙整完成！")
            st.download_button(label="📥 下載最終 HK 報關文件", data=output.getvalue(), file_name=f"{t_str}_HK_GM_Final.xlsx", use_container_width=True)

        except Exception as e:
            st.error(f"錯誤：{e}")
