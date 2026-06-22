import os
import re
import gradio as gr
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- 1. 固定讀取同目錄下的資料庫檔案 ---
DATA_FILE = "data.csv"

def load_fixed_database():
    if os.path.exists(DATA_FILE):
        try:
            df = pd.read_csv(DATA_FILE)
            df.columns = df.columns.str.strip()
            for col in ['成分', '劑型', '劑量', '廠商', '年度']:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip()
            return df
        except Exception as e:
            print(f"讀取本地資料庫失敗: {e}")
            return pd.DataFrame()
    else:
        return pd.DataFrame()

df_global = load_fixed_database()
if not df_global.empty:
    ALL_COMPONENTS = sorted(df_global['成分'].dropna().unique())
    INIT_MESSAGE = f"✅ 成功載入資料庫！共 {len(df_global)} 筆數據，包含 {len(ALL_COMPONENTS)} 種成分。"
else:
    ALL_COMPONENTS = []
    INIT_MESSAGE = "❌ 未找到 data.csv 檔案，請將資料庫檔案上傳至同目錄。"

# --- 2. 輔助函式 ---
def parse_dosage_to_numeric(dose_str):
    if not isinstance(dose_str, str):
        return 0.0
    match = re.search(r'([0-9\.]+)', dose_str)
    if match:
        val = float(match.group(1))
        if 'g' in dose_str.lower() and 'mg' not in dose_str.lower() and 'mcg' not in dose_str.lower():
            return val * 1000
        return val
    return 0.0

# --- 3. 動態介面連動函式 ---
def update_component_choices(search_text):
    if not ALL_COMPONENTS:
        return gr.update(choices=[], value=[])
    if not search_text or search_text.strip() == "":
        return gr.update(choices=[], value=[])
    search_text = search_text.strip().lower()
    filtered = [c for c in ALL_COMPONENTS if search_text in c.lower()]
    return gr.update(choices=filtered, value=[])

def update_form_choices(selected_comps):
    if not selected_comps or df_global.empty:
        return gr.update(choices=[], value=[])
    df_filtered = df_global[df_global['成分'].isin(selected_comps)]
    forms = sorted(df_filtered['劑型'].dropna().unique())
    return gr.update(choices=forms, value=forms)

def update_dose_choices(selected_comps, selected_forms):
    if not selected_comps or not selected_forms or df_global.empty:
        return gr.update(choices=[], value=[])
    df_filtered = df_global[
        (df_global['成分'].isin(selected_comps)) & 
        (df_global['劑型'].isin(selected_forms))
    ]
    doses = sorted(df_filtered['劑量'].dropna().unique(), key=parse_dosage_to_numeric)
    return gr.update(choices=doses, value=doses)

# --- 4. 產出標準化 Excel 與 HTML 預覽報表 ---
def generate_reports(selected_components, selected_forms, selected_doses):
    if df_global.empty:
        return None, "❌ 系統未成功載入 data.csv 資料庫。", ""
    
    # 1. 放寬防呆：只強制檢查成分與劑型，允許複方藥物不勾選劑量
    if not selected_components or not selected_forms:
        return None, "❌ 請確認成分與劑型皆已勾選！", ""

    # 2. 動態過濾：有勾選劑量才過濾，沒勾就全放行
    condition = (df_global['成分'].isin(selected_components)) & (df_global['劑型'].isin(selected_forms))
    if selected_doses:
        condition &= (df_global['劑量'].isin(selected_doses))

    df_filtered = df_global[condition].copy()
    
    if df_filtered.empty:
        return None, "❌ 找不到符合該條件的資料！", ""

    # 3. 強制清洗劑量欄位，將 nan 轉為「空字串」
    df_filtered['劑量'] = df_filtered['劑量'].astype(str).str.strip()
    df_filtered['劑量'] = df_filtered['劑量'].apply(
        lambda x: '' if str(x).lower() in ['nan', 'none', '<na>', 'null', ''] else x
    )

    qty_col = '數量(顆)' if '數量(顆)' in df_filtered.columns else [col for col in df_filtered.columns if '數量' in col][0]
    df_filtered[qty_col] = df_filtered[qty_col].astype(str).str.replace(',', '').str.strip()
    df_filtered[qty_col] = pd.to_numeric(df_filtered[qty_col], errors='coerce').fillna(0)

    # 4. 進行樞紐分析
    pivot_df = df_filtered.groupby(['成分', '劑型', '劑量', '廠商', '年度'])[qty_col].sum().unstack(fill_value=0)
    for year_col in ['2022年', '2023年', '2024年']:
        if year_col not in pivot_df.columns:
            pivot_df[year_col] = 0
            
    pivot_df = pivot_df.reindex(columns=['2022年', '2023年', '2024年']).reset_index()
    pivot_df['dose_numeric'] = pivot_df['劑量'].apply(parse_dosage_to_numeric)
    pivot_df = pivot_df.sort_values(
        by=['成分', '劑型', 'dose_numeric', '2024年'], 
        ascending=[True, True, True, False]
    ).drop(columns=['dose_numeric'])

    # --- Excel 處理邏輯 ---
    wb = Workbook()
    ws = wb.active
    ws.title = "廠商排名報表"
    ws.views.sheetView[0].showGridLines = True

    # 👇 請在這裡加上這行：強制設定打開時的視角縮放比例為 85% (可依喜好改為 80)
    ws.views.sheetView[0].zoomScale = 85

    font_family = "微軟正黑體"
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    subtotal_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    total_fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    
    header_font = Font(name=font_family, size=12, bold=True, color="FFFFFF")
    data_font = Font(name=font_family, size=12)
    bold_font = Font(name=font_family, size=12, bold=True)
    
    # wrap_text=True 確保 Excel 內遇到 \n 會自動換行
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center")
    
    thin_side = Side(border_style="thin", color="D9D9D9")
    thick_bottom_side = Side(border_style="medium", color="000000")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    thick_bottom_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thick_bottom_side)

    # --- HTML 處理邏輯 (已優化深色模式相容性、視覺體驗與手機滑動提示) ---
    html_content = f"""
    <style>
        .report-preview-wrapper {{
            padding: 15px 0; /* 外圍留白，讓紙張有懸浮感 */
        }}
        .report-container {{ 
            background-color: #FFFFFF !important; /* 強制白底 */
            color: #333333 !important;           /* 強制深灰字 */
            padding: 20px !important;            /* 增加內距，讓畫面不擁擠 */
            width: 100%;
            box-sizing: border-box;
            -webkit-text-size-adjust: 100%;
            
            /* 【視覺優化】：讓預覽區像一張真正的白紙 (Card UI)，降低突兀感 */
            border-radius: 12px;
            box-shadow: 0 8px 24px rgba(0,0,0,0.25);
            border: 1px solid #E0E0E0;
        }}
        
        /* 【新增】：手機版專屬滑動提示文字與呼吸動畫 */
        .scroll-hint {{
            display: none; /* 預設電腦版隱藏 */
            text-align: right;
            font-size: 13px;
            color: #888888 !important; /* 低調的灰色，不搶戲 */
            margin-bottom: 8px;
            font-family: '微軟正黑體', sans-serif;
            font-weight: bold;
            animation: pulse 2s infinite; /* 呼吸燈動畫 */
        }}
        @keyframes pulse {{
            0% {{ opacity: 0.5; }}
            50% {{ opacity: 1; }}
            100% {{ opacity: 0.5; }}
        }}
        /* 只有在螢幕寬度小於 768px (手機/小平板) 時才顯示提示 */
        @media (max-width: 768px) {{
            .scroll-hint {{
                display: block; 
            }}
        }}

        .table-responsive {{
            width: 100%;
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
            padding-bottom: 10px;
        }}
        
        /* 強制顯示水平捲軸並美化它 (電腦版與部分安卓適用) */
        .table-responsive::-webkit-scrollbar {{
            height: 10px;
        }}
        .table-responsive::-webkit-scrollbar-track {{
            background: #F1F1F1;
            border-radius: 5px;
        }}
        .table-responsive::-webkit-scrollbar-thumb {{
            background: #C0C0C0;
            border-radius: 5px;
        }}
        .table-responsive::-webkit-scrollbar-thumb:active {{
            background: #A0A0A0;
        }}

        .report-table {{
            width: max-content;
            min-width: 100%;
            border-collapse: collapse; 
            font-family: '微軟正黑體', sans-serif; 
            font-size: 15px !important;
            background-color: #FFFFFF !important;
        }}
        .report-table th, .report-table td {{
            border: 1px solid #D9D9D9 !important;
            padding: 8px 12px;
            white-space: nowrap !important;
            vertical-align: middle !important;
            color: #333333 !important;           
        }}
        .report-table th {{ 
            background-color: #1F497D !important;
            color: #FFFFFF !important; 
            text-align: center !important; 
            font-weight: bold;
        }}
        .report-title {{
            color: #000000 !important;
            text-align: center;
            font-family: '微軟正黑體';
            font-weight: bold;
            margin-top: 5px;
            margin-bottom: 20px;
        }}
        .thick-bottom {{ border-bottom: 2px solid black !important; }}
        
        /* 【頁尾修正】：強制字體為深色，防止深色模式反白導致截圖隱形 */
        .report-footer {{
            display: flex;
            flex-wrap: wrap; /* 允許手機螢幕過窄時換行 */
            justify-content: space-between;
            margin-top: 25px;
            padding-top: 15px;
            border-top: 1px solid #D9D9D9; 
            font-size: 13px !important;
            font-family: '微軟正黑體', sans-serif;
            font-weight: bold;
            color: #333333 !important; 
        }}
        .report-footer span {{
            color: #333333 !important;
        }}
    </style>
    
    <div class="report-preview-wrapper">
        <div id="report-capture-area" class="report-container">
    """

    comp_names = "、".join(df_filtered['成分'].unique())
    form_mapping = {"注射劑": "Inj.", "一般錠劑膠囊劑": "Tab./Cap.", "膜衣錠": "F.C. Tab.", "膠囊劑": "Cap.", "錠劑": "Tab."}
    unique_forms = df_filtered['劑型'].unique()
    form_abbr = f" {form_mapping.get(unique_forms[0], unique_forms[0])}" if len(unique_forms) == 1 else ""
    
    valid_doses = [d for d in df_filtered['劑量'].unique() if d != '']
    doses_for_header = sorted(valid_doses, key=parse_dosage_to_numeric)
    doses_str = "、".join(doses_for_header)
    dose_display = f" {doses_str}" if doses_str else ""
    
    header_title_text = f"{comp_names}{form_abbr}{dose_display}廠商申報量排名"

    # 【套用深色模式主標題樣式】
    html_content += f'<h2 class="report-title">{header_title_text}</h2>'
    
    # 【新增】：插入滑動提示文字
    html_content += '<div class="scroll-hint">👉 左右滑動表格可查看完整數據</div>'
    
    html_content += '<div class="table-responsive"><table class="report-table">'

    headers = ["成分", "劑型", "劑量", "廠商", "2022年<br>數量", "2023年<br>數量", "2024年<br>數量", "2024年<br>占比(%)"]
    ws_headers = [h.replace("<br>", "\n") for h in headers]
    ws.append(ws_headers)
    ws.row_dimensions[1].height = 45 # 標題列高
    
    html_content += "<tr>"
    for col_idx, h in enumerate(ws_headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = cell_border
        html_content += f"<th style='color: #FFFFFF !important; background-color: #1F497D !important;'>{headers[col_idx-1]}</th>"
    html_content += "</tr>"

    unique_groups = pivot_df[['成分', '劑型', '劑量']].drop_duplicates()
    unique_groups_list = list(unique_groups.iterrows())
    
    grand_total_2022 = 0
    grand_total_2023 = 0
    grand_total_2024 = 0
    current_row = 2

    last_comp = None
    last_form = None

    for idx, (_, group_keys) in enumerate(unique_groups_list):
        comp = group_keys['成分']
        form = group_keys['劑型']
        dose = group_keys['劑量']

        is_last_of_form = False
        if idx == len(unique_groups_list) - 1:
            is_last_of_form = True
        else:
            next_comp = unique_groups_list[idx+1][1]['成分']
            next_form = unique_groups_list[idx+1][1]['劑型']
            if comp != next_comp or form != next_form:
                 is_last_of_form = True

        dose_group = pivot_df[
            (pivot_df['成分'] == comp) & 
            (pivot_df['劑型'] == form) & 
            (pivot_df['劑量'] == dose)
        ]
        
        dose_2024_sum = dose_group['2024年'].sum()
        dose_subtotal_2022 = dose_group['2022年'].sum()
        dose_subtotal_2023 = dose_group['2023年'].sum()
        dose_subtotal_2024 = dose_2024_sum
        
        print_comp = (comp != last_comp)
        print_form = print_comp or (form != last_form)
        is_first_row_in_dose = True
        
        for _, row in dose_group.iterrows():
            qty_2022 = float(row['2022年'])
            qty_2023 = float(row['2023年'])
            qty_2024 = float(row['2024年'])
            ratio = (qty_2024 / dose_2024_sum) if dose_2024_sum > 0 else 0.0
            
            c_val = row['成分'] if print_comp and is_first_row_in_dose else ""
            f_val = row['劑型'] if print_form and is_first_row_in_dose else ""
            d_val = row['劑量'] if is_first_row_in_dose else ""
             
            ws.append([c_val, f_val, d_val, row['廠商'], qty_2022, qty_2023, qty_2024, ratio])

            # 👇 請在這裡加上這行：拉高資料列
            ws.row_dimensions[current_row].height = 35
            
            html_content += f"<tr>"
            html_content += f"<td style='text-align:center;'>{c_val}</td>"
            html_content += f"<td style='text-align:center;'>{f_val}</td>"
            html_content += f"<td style='text-align:center;'>{d_val}</td>"
            html_content += f"<td style='text-align:left;'>{row['廠商']}</td>"
            html_content += f"<td style='text-align:right;'>{qty_2022:,.0f}</td>"
            html_content += f"<td style='text-align:right;'>{qty_2023:,.0f}</td>"
            html_content += f"<td style='text-align:right;'>{qty_2024:,.0f}</td>"
            html_content += f"<td style='text-align:right;'>{ratio:.1%}</td>"
            html_content += "</tr>"
            
            for col_idx in range(1, 9):
                cell = ws.cell(row=current_row, column=col_idx)
                cell.font = data_font
                cell.border = cell_border
                if col_idx in [1, 2, 3]: cell.alignment = center_align
                elif col_idx == 4: cell.alignment = left_align
                elif col_idx in [5, 6, 7]: cell.alignment = right_align; cell.number_format = '#,##0'
                elif col_idx == 8: cell.alignment = right_align; cell.number_format = '0.0%'
            
            current_row += 1
            is_first_row_in_dose = False
        
        last_comp = comp
        last_form = form
            
        dose_label = f"{dose} 合計" if dose != "" else "合計"
        
        ws.append(["", "", dose_label, "", dose_subtotal_2022, dose_subtotal_2023, dose_subtotal_2024, 1.0])
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=4)

        # 👇 請在這裡加上這行：拉高小計列
        ws.row_dimensions[current_row].height = 35
        
        thick_class = "thick-bottom" if is_last_of_form else ""
        html_content += f"""
        <tr style="background-color: #DCE6F1; font-weight: bold;">
            <td class="{thick_class}"></td><td class="{thick_class}"></td>
            <td colspan="2" class="{thick_class}" style="text-align:center;">{dose_label}</td>
            <td class="{thick_class}" style="text-align:right;">{dose_subtotal_2022:,.0f}</td>
            <td class="{thick_class}" style="text-align:right;">{dose_subtotal_2023:,.0f}</td>
            <td class="{thick_class}" style="text-align:right;">{dose_subtotal_2024:,.0f}</td>
            <td class="{thick_class}" style="text-align:right;">100.0%</td>
        </tr>
        """
        
        for col_idx in range(1, 9):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.font = bold_font
            cell.fill = subtotal_fill
            cell.border = thick_bottom_border if is_last_of_form else cell_border
            if col_idx == 3: cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
            elif col_idx in [5, 6, 7]: cell.alignment = right_align; cell.number_format = '#,##0'
            elif col_idx == 8: cell.alignment = right_align; cell.number_format = '0.0%'
                
        grand_total_2022 += dose_subtotal_2022
        grand_total_2023 += dose_subtotal_2023
        grand_total_2024 += dose_subtotal_2024
        current_row += 1

    # 總計列
    ws.append(["總計", "", "", "", grand_total_2022, grand_total_2023, grand_total_2024, 1.0])
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)

    # 👇 請在這裡加上這行：拉高總計列
    ws.row_dimensions[current_row].height = 35
    
    html_content += f"""
    <tr style="background-color: #B8CCE4; font-weight: bold;">
        <td colspan="4" style="text-align:center;">總計</td>
        <td style="text-align:right;">{grand_total_2022:,.0f}</td>
        <td style="text-align:right;">{grand_total_2023:,.0f}</td>
        <td style="text-align:right;">{grand_total_2024:,.0f}</td>
        <td style="text-align:right;">100.0%</td>
    </tr>
    </table></div>
    """
    
    # 頁尾
    html_content += """
        <div class="report-footer">
            <span>中央健康保險署  政府資料開放平台 2024年資料</span>
            <span>https://data.gov.tw/dataset/22131</span>
        </div>
        </div> </div> """

    for col_idx in range(1, 9):
        cell = ws.cell(row=current_row, column=col_idx)
        cell.font = bold_font
        cell.fill = total_fill
        cell.border = cell_border
        if col_idx == 1: cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        elif col_idx in [5, 6, 7]: cell.alignment = right_align; cell.number_format = '#,##0'
        elif col_idx == 8: cell.alignment = right_align; cell.number_format = '0.0%'

    # --- 調整 Excel 欄寬 ---
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == 'A':
            ws.column_dimensions[col_letter].width = 35 
            continue
        elif col_letter in ['B']:
            ws.column_dimensions[col_letter].width = 20
            continue
        elif col_letter in ['E', 'F', 'G']:
            ws.column_dimensions[col_letter].width = 18 
            continue
        elif col_letter == 'H':
            ws.column_dimensions[col_letter].width = 11.5
            continue
            
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            lines = val_str.split('\n')
            for line in lines:
                line_len = sum(2 if '\u4e00' <= char <= '\u9fff' else 1 for char in line)
                if line_len > max_len: max_len = line_len
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1 
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = 9 
    ws.page_margins.top = 2.7 / 2.54; ws.page_margins.bottom = 2.5 / 2.54
    ws.page_margins.left = 1.5 / 2.54; ws.page_margins.right = 1.5 / 2.54
    ws.page_margins.header = 1.5 / 2.54; ws.page_margins.footer = 1.0 / 2.54

    header_string_for_print = f'&"微軟正黑體,Bold"&16{header_title_text}'
    ws.oddHeader.center.text = header_string_for_print
    ws.oddFooter.left.text = '&"微軟正黑體,Regular"&12中央健康保險署  政府資料開放平台 2024年資料'
    ws.oddFooter.right.text = '&"微軟正黑體,Regular"&12https://data.gov.tw/dataset/22131'
    ws.page_setup.scaleWithDoc = True
    ws.page_setup.alignWithMargins = True

    safe_filename = re.sub(r'[\\/*?:"<>|]', "_", header_title_text) + ".xlsx"
    wb.save(safe_filename)
    
    success_msg = "🎉 成功！已順利產出「Excel 報表」與下方「預覽畫面」。請點擊下方按鈕下載檔案或存為圖片！"
    return safe_filename, success_msg, html_content


# --- 5. 優化版：行動端唯一調用原生分享(9047)，電腦版自動下載 ---
download_js = """
function() {
    var element = document.getElementById('report-capture-area');
    if (!element) {
        alert('請先產生報表再下載圖片！');
        return [];
    }
    
    if (typeof html2canvas === 'undefined') {
        alert('截圖套件載入中，請稍後再試或重新整理網頁。');
        return [];
    }
    
    var table = document.querySelector('.report-table');
    var wrapper = document.querySelector('.table-responsive');
    var scrollHint = document.querySelector('.scroll-hint'); // 抓取滑動提示文字
    
    // 紀錄原始樣式
    var originalElementWidth = element.style.width;
    var originalWrapperOverflow = wrapper ? wrapper.style.overflowX : '';
    var originalBoxShadow = element.style.boxShadow;
    var originalBorderRadius = element.style.borderRadius;
    var originalBorder = element.style.border;
    var originalHintDisplay = scrollHint ? scrollHint.style.display : ''; // 紀錄提示字原本狀態
    
    var targetWidth = table ? (table.offsetWidth + 40) : element.scrollWidth;
    
    // 截圖前：撐開寬度、移除陰影與圓角，並【隱藏提示文字】確保圖片乾淨
    element.style.width = targetWidth + 'px';
    element.style.boxShadow = 'none';
    element.style.borderRadius = '0';
    element.style.border = 'none';
    if(wrapper) wrapper.style.overflowX = 'visible';
    if(scrollHint) scrollHint.style.display = 'none'; // 拍照瞬間隱藏
    
    var titleElement = element.querySelector('h2');
    var fileName = titleElement ? titleElement.innerText.replace(/[\\\\/*?:"<>|]/g, "_") : '廠商排名報表';
    
    html2canvas(element, { 
        scale: 2, 
        backgroundColor: '#FFFFFF',
        width: targetWidth,
        windowWidth: targetWidth 
    }).then(function(canvas) {
        
        // 截圖完畢：瞬間把排版、陰影、圓角與【提示文字】還原
        element.style.width = originalElementWidth;
        element.style.boxShadow = originalBoxShadow;
        element.style.borderRadius = originalBorderRadius;
        element.style.border = originalBorder;
        if(wrapper) wrapper.style.overflowX = originalWrapperOverflow;
        if(scrollHint) scrollHint.style.display = originalHintDisplay; // 把提示字叫回來
        
        // 精準判斷行動裝置
        var isMobile = /iPhone|iPad|iPod|Android/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
        
        if (isMobile) {
            canvas.toBlob(function(blob) {
                var file = new File([blob], fileName + '.png', { type: 'image/png' });
                if (navigator.canShare && navigator.canShare({ files: [file] })) {
                    navigator.share({
                        files: [file],
                        title: fileName
                    }).catch((error) => console.log('分享中斷', error));
                } else {
                    alert('目前瀏覽器環境不支援原生分享圖片，請確保使用 Safari 或 Chrome 開啟。');
                }
            }, 'image/png');
        } else {
            var imgDataUrl = canvas.toDataURL('image/png');
            var link = document.createElement('a');
            link.download = fileName + '.png';
            link.href = imgDataUrl;
            link.click();
        }
    });
    return [];
}
"""


# --- 6. 介面層 (載入 html2canvas 套件) ---
with gr.Blocks(
    title="健保資料庫數據分析工具"
) as demo:
    
    gr.Markdown("# 💊 健保資料庫數據分析工具")
    gr.Markdown(f"**系統狀態：** {INIT_MESSAGE}")
    
    gr.Markdown("⚠️ **提醒使用手機版的同仁：請務必使用系統預設瀏覽器（如 Safari 或 Chrome）開啟本網頁，才能成功下載檔案與圖片。**")
    
    with gr.Row():
        # 左側區塊：輸入條件
        with gr.Column(scale=1):
            gr.Markdown("### 🔍 輸入條件")
            search_input = gr.Textbox(label="第一步：輸入成分關鍵字", placeholder="例如：Levofloxacin", value="")
            component_choices = gr.CheckboxGroup(label="📋 第二步：勾選成分品項 (可多選)", choices=[], value=[])
            form_choices = gr.CheckboxGroup(label="💊 第三步：勾選欲包含的劑型", choices=[], value=[])
            dose_choices = gr.CheckboxGroup(label="🧪 第四步：勾選欲包含的劑量", choices=[], value=[])
            submit_btn = gr.Button("🚀 第五步：產生 Excel 報表與預覽", variant="primary")
            
        # 右側區塊：下載與輸出
        with gr.Column(scale=1):
            gr.Markdown("### 📥 報表與圖片下載")
            status_output = gr.Textbox(label="系統處理結果", interactive=False)
            
            download_excel_btn = gr.DownloadButton("📄 一鍵下載 Excel 報表", variant="primary", size="lg")
            download_img_btn = gr.Button("🖼️ 一鍵下載為高畫質圖片 (PNG)", variant="primary", size="lg")

    gr.Markdown("---")
    
    gr.Markdown("### 網頁即時預覽")
    html_output = gr.HTML(label="報表預覽區")

    # --- 綁定事件 ---
    search_input.change(fn=update_component_choices, inputs=search_input, outputs=component_choices)
    component_choices.change(fn=update_form_choices, inputs=component_choices, outputs=form_choices)
    form_choices.change(fn=update_dose_choices, inputs=[component_choices, form_choices], outputs=dose_choices)
    
    submit_btn.click(
        fn=generate_reports, 
        inputs=[component_choices, form_choices, dose_choices], 
        outputs=[download_excel_btn, status_output, html_output] 
    )
    
    download_img_btn.click(fn=None, js=download_js)

if __name__ == "__main__":
    demo.launch(
        theme=gr.themes.Default(primary_hue="blue", secondary_hue="slate"),
        head='<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>'
    )