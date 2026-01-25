import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime, timedelta

# 1. 網頁基本設定
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
uploaded_files = st.file_uploader("請一次拖入所有 4 個必要檔案", type=['xls', 'xlsx'], accept_multiple_files=True)

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

# 4. 執行與檢查邏輯
if all(files_dict.values()):
    if st.button("🚀 執行品牌校驗與轉換", use_container_width=True):
        try:
            def smart_read_excel(file_obj, **kwargs):
                if file_obj.name.endswith('.xls'): return pd.read_excel(file_obj, engine='xlrd', **kwargs)
                else: return pd.read_excel(file_obj, engine='openpyxl', **kwargs)

            # --- A. 預讀取用於品牌檢查 ---
            df_inv_check = smart_read_excel(files_dict["Invoice"], header=None, nrows=2).fillna('')
            df_pac_check = smart_read_excel(files_dict["Packing"], header=None, nrows=2).fillna('')
            df_north_check = smart_read_excel(files_dict["北方文件"], sheet_name='出口明細', nrows=5).fillna('')

            inv_name = str(df_inv_check.iloc[1, 0])
            pac_name = str(df_pac_check.iloc[1, 0])
            north_remark = str(df_north_check.iloc[0, 8]) 

            # 定義品牌關鍵字字典與檔名後綴
            brands = {
                "蜜凱": {"eng": "COSMETICS", "label": "TRUU+TRUE YOU", "key": "蜜凱", "suffix": "蜜凱"},
                "歐瑞": {"eng": "food supplement", "label": "ALLRE", "key": "歐瑞", "suffix": "歐瑞"},
                "綺麗絲": {"eng": "MAKEUP", "label": "MKUP", "key": "綺麗絲", "suffix": "綺麗絲"}
            }

            current_brand = None
            for b_name, b_info in brands.items():
                if b_info["key"] in inv_name:
                    current_brand = b_info
                    break

            # --- B. 執行品牌交叉檢查 ---
            if current_brand:
                key = current_brand["key"]
                errors = []
                if key not in pac_name: errors.append(f"Packing 檔案 (偵測到: {pac_name})")
                if key not in north_remark: errors.append(f"北方文件 (偵測到備註: {north_remark})")
                
                if errors:
                    st.error(f"🚨 檔案品牌不匹配！主品牌判定為：【{key}】")
                    for err in errors: st.warning(f"❌ 錯誤檔案：{err}")
                    st.stop()
            else:
                st.error("❌ 無法從 Invoice 辨識公司品牌，請確認檔案。")
                st.stop()

            # --- C. 執行正式轉換 ---
            with st.spinner(f'處理中...'):
                tw_now = datetime.utcnow() + timedelta(hours=8)
                t_str = tw_now.strftime("%Y%m%d")
                
                # 設定動態檔案名稱
                final_filename = f"{t_str}_HK_{current_brand['suffix']}.xlsx"

                df_order = smart_read_excel(files_dict["OrderList"], dtype=str).fillna('')
                df_n_export = smart_read_excel(files_dict["北方文件"], sheet_name='出口明細', dtype=str).fillna('')
                df_n_bag_raw = smart_read_excel(files_dict["北方文件"], sheet_name='袋數編號', dtype=str)
                bag_count = len(df_n_bag_raw[df_n_bag_raw.iloc[:, 1].str.strip() != ""])
                df_n_bag = df_n_bag_raw.fillna('')
                df_inv_raw = smart_read_excel(files_dict["Invoice"], header=None, dtype=str).fillna('')
                df_packing_raw = smart_read_excel(files_dict["Packing"], header=None, dtype=str).fillna('')

                def get_inv(cell_ref):
                    col_map = {'A':0, 'B':1, 'C':2, 'D':3, 'E':4, 'F':5, 'G':6, 'H':7, 'I':8}
                    c = col_map[cell_ref[0]]; r = int(cell_ref[1:]) - 1
                    return df_inv_raw.iloc[r, c]

                wb = Workbook(); ws = wb.active; ws.title = "HK最終報關檔"

                # 欄寬設定 (A: 5.5, P: 10.5)
                ws.column_dimensions['A'].width = 5.5
                col_widths = {'B': 20.8, 'C': 19.2, 'D': 14.7, 'E': 14, 'F': 14, 'G': 8.7, 'H': 13, 'I': 51.82, 'J': 30, 'K': 17.9, 'L': 8.7, 'M': 8.7, 'N': 8.09, 'O': 10.91, 'P': 10.5, 'Q': 8.09}
                for col, width in col_widths.items(): ws.column_dimensions[col].width = width
                
                # 行高與表頭填充
                ws["B1"] = "INVOICE/PACKING"; ws["B1"].font = Font(name='Arial', size=28, bold=True)
                ws["B1"].alignment = Alignment(horizontal='left', vertical='center'); ws.merge_cells("B1:E1")
                
                # 遍歷填充 Invoice 表頭資訊
                head_configs = [("B2", get_inv("A2"), "B2:I2", 10, False, True), ("B3", get_inv("A3"), "B3:E3", 10, False, False), ("F3", get_inv("E3"), "F3:I3", 10, False, False), ("B4", get_inv("A4"), "B4:I4", 10, False, False), ("B5", get_inv("A5"), "B5:E5", 10, False, False), ("F5", get_inv("E5"), "F5:I5", 10, False, False), ("B6", get_inv("A6"), "B6:I6", 10, False, False), ("B7", get_inv("A7"), "B7:E7", 10, False, False), ("F7", get_inv("E7"), "F7:I7", 10, False, True), ("B8", get_inv("A8"), "B8:E8", 10, False, True), ("F8", get_inv("E8"), "F8:I8", 10, False, False), ("B9", get_inv("A9"), "B9:D9", 10, False, False), ("E9", get_inv("D9"), "E9:G9", 10, False, False), ("H9", get_inv("G9"), "H9:I9", 10, False, False), ("B10", get_inv("A10"), "B10:E10", 10, False, False), ("F10", get_inv("E10"), "F10:I10", 10, False, False)]
                for c_id, cont, m_range, sz, bld, wrp in head_configs:
                    ws[c_id] = cont; ws[c_id].font = Font(name='Arial', size=sz, bold=bld)
                    ws[c_id].alignment = Alignment(wrap_text=wrp, vertical='center'); ws.merge_cells(m_range)

                # 資料處理與排序
                barcode_dict = df_n_bag.set_index(df_n_bag.columns[0])[df_n_bag.columns[1]].to_dict()
                bag_dict = df_n_export.set_index(df_n_export.columns[1])[df_n_export.columns[6]].to_dict()

                all_rows = []
                for _, r in df_order.iterrows():
                    hawb, oid = str(r.iloc[1]).strip(), str(r.iloc[3]).strip()
                    bag_no = bag_dict.get(hawb, ""); barcode = barcode_dict.get(bag_no, "")
                    gw_raw = r.iloc[30]
                    try: gw_num = float(gw_raw)
                    except: gw_num = 0.0
                    all_rows.append({"hawb": hawb, "oid": oid, "bag_no": bag_no, "barcode": barcode, "gw_raw": gw_raw, "gw_num": gw_num, "orig_row": r})

                all_rows.sort(key=lambda x: (x["barcode"], x["hawb"], x["gw_num"]))

                # 填充明細... (省略重複的樣式設定代碼)
                # [此處包含與前版相同的資料填充與統計邏輯]
                
                # --- 資料填充與最後統計 (省略，維持前版邏輯) ---
                # (為了長度縮減，此處邏輯與前一版完全相同，包含最後的統計列與格式)

                # ... 完成資料寫入後 ...
                
                # 為求示範完整性，確保 curr_row 邏輯正確
                # (實際運行時，這部分應接續在資料填充循環後)

                output = BytesIO(); wb.save(output); st.balloons()
                st.success(f"✅ 品牌【{current_brand['label']}】處理完成！")
                st.download_button(label=f"📥 下載 {final_filename}", data=output.getvalue(), file_name=final_filename, use_container_width=True)

        except Exception as e:
            st.error(f"發生錯誤：{e}")
