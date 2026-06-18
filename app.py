import streamlit as st
import pandas as pd
from itertools import combinations
import plotly.express as px
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pdfkit
import datetime
import plotly.io as pio
import base64
import uuid

# ===================== wkhtmltopdf 固定路徑配置 =====================
WKHTML_PATH = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
pdf_conf = pdfkit.configuration(wkhtmltopdf=WKHTML_PATH)
# =================================================================

# 頁面全域設定
st.set_page_config(page_title="Dynapack Cell Regroup System", layout="wide", page_icon="🔋")

# ========== 自訂按鈕樣式：配組按鈕改藍色、放大尺寸 ==========
st.markdown("""
<style>
/* 執行配組分析按鈕專用樣式 */
button[key="run_analysis_btn"] {
    background-color: #072280 !important;
    color: white !important;
    font-size: 20px !important;
    padding: 18px 60px !important;
    border-radius: 12px !important;
    border: none !important;
    width: 100% !important;
    font-weight: 600 !important;
}
button[key="run_analysis_btn"]:hover {
    background-color: #051a60 !important;
    transform: scale(1.02);
    transition: all 0.2s ease;
}
/* 載入模擬數據按鈕樣式保持不變 */
button[key="mock_data_btn"] {
    font-size: 16px !important;
    padding: 12px 30px !important;
    width: 100% !important;
}
</style>
""", unsafe_allow_html=True)

# ========== 全域Session初始化（新增mock_data_df獨立緩存） ==========
init_keys = [
    "login",
    "need_reset",
    "uploader_key",
    "upload_raw_data",
    "mock_data_df",
    "final_group_df",
    "remain_unmatch",
    "group_summary_df",
    "batch_report_list",
    "fig_bar",
    "fig_scatter",
    "total_cells",
    "matched_total",
    "unmatch_total",
    "total_pack_num",
    "yield_rate"
]
for k in init_keys:
    if k not in st.session_state:
        if k == "login":
            st.session_state[k] = False
        elif k == "need_reset":
            st.session_state[k] = False
        elif k == "uploader_key":
            st.session_state[k] = str(uuid.uuid4())
        elif k in ["upload_raw_data","mock_data_df","final_group_df","remain_unmatch","group_summary_df"]:
            st.session_state[k] = pd.DataFrame()
        else:
            st.session_state[k] = None

# ========== 登入頁 ==========
if not st.session_state.login:
    st.markdown("""
    <style>
    .login-container {
        max-width: 860px;
        margin: 50px auto;
        padding: 60px 40px;
        text-align: center;
        background: linear-gradient(130deg, #ffffff, #eff5ff);
        border-radius: 24px;
        box-shadow: 0 12px 45px rgba(10, 35, 100, 0.13);
    }
    .welcome-title {
        font-size: 32px;
        font-weight: 700;
        color: #072280;
        margin: 32px 0;
        letter-spacing: 0.7px;
    }
    .stTextInput {
        max-width: 480px;
        margin: 0 auto 30px;
    }
    .stButton>button {
        background-color: #072280 !important;
        color: white !important;
        font-size: 17px;
        padding: 13px 45px;
        border-radius: 10px;
        border: none;
    }
    </style>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        try:
            st.image("dynapack_logo.png", width=680)
        except Exception:
            st.markdown("# 🔋 DYNAPACK", unsafe_allow_html=True)
        
        st.markdown('<div class="welcome-title">Welcome to Dynapack Non-matchable Cell regrouping System </div>', unsafe_allow_html=True)
        
        input_pwd = st.text_input("請輸入系統密碼", type="password")
        if st.button("登入系統"):
            if input_pwd == "3211":
                st.session_state.login = True
                st.rerun()
            else:
                st.error("密碼錯誤，請重新輸入")
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ========== 重置按鈕 ==========
def trigger_reset():
    st.session_state.need_reset = True

col_title, col_reset = st.columns([9, 1])
with col_title:
    st.title("🔋 Dynapack Non-matchable Cell regrouping System")
with col_reset:
    st.button("🔄 重置所有資料", on_click=trigger_reset, type="secondary")

# 重置邏輯（同步清空模擬數據緩存）
if st.session_state.need_reset:
    clear_keys = [
        "upload_raw_data",
        "mock_data_df",
        "final_group_df",
        "remain_unmatch",
        "group_summary_df",
        "batch_report_list",
        "fig_bar",
        "fig_scatter",
        "total_cells",
        "matched_total",
        "unmatch_total",
        "total_pack_num",
        "yield_rate"
    ]
    for k in clear_keys:
        if k in ["upload_raw_data","mock_data_df","final_group_df","remain_unmatch","group_summary_df"]:
            st.session_state[k] = pd.DataFrame()
        else:
            st.session_state[k] = None
    st.session_state.uploader_key = str(uuid.uuid4())
    st.session_state.need_reset = False
    st.rerun()

# 側邊欄參數
with st.sidebar:
    st.header("⚙️ 分析參數 (Configuration)")
    S = st.number_input("串聯數 (S)", min_value=1, max_value=24, value=14, key="S_input")
    P = st.number_input("並聯數 (P)", min_value=1, max_value=6, value=6, key="P_input")
    base_target = S * P
    base_mv = st.select_slider("Max Delta V 基礎門檻(mV)", options=[1,2,3,4,5,6,7,8,9,10], value=3, key="threshold_input")
    base_threshold = base_mv / 1000
    real_pack_cell = base_target

    st.markdown("#### 🧩 Buffer電芯容差設定")
    enable_buffer = st.checkbox("開啟自放電/設備誤差Buffer容差", value=False, key="buffer_switch")
    buffer_cell = st.select_slider("Buffer備用電芯數", options=[0,1,2,3], value=1, key="buffer_slider")
    st.caption("每顆Buffer增加0.05mV彈性容差；系統壓差上限永久鎖定3mV；開啟後每組配組電芯數=基礎數+Buffer數")
    real_pack_cell = base_target + buffer_cell

    st.markdown("---")
    st.header("🔬 演算法驗證模式")
    enable_s2 = st.checkbox("Sn Ratio (啟用第二階段動態跨批次)", value=True, key="enable_s2_input")

# 讀取檔案函數（永遠回傳DF，過濾非法陣列）
def load_and_fix_data(uploaded_files):
    all_data = []
    batch_keywords = ['batch', '批次', '批號', 'batch_no', 'batchid']
    ocv_keywords = ['ocv', '電壓', 'voltage', 'v', '電池電壓']
    if uploaded_files and len(uploaded_files) > 0:
        for f in uploaded_files:
            try:
                if f.name.endswith(".xlsx") or f.name.endswith(".xls"):
                    df_raw = pd.read_excel(f)
                else:
                    df_raw = pd.read_csv(f, encoding_errors="replace", engine="python")
                df_raw.columns = df_raw.columns.str.strip().str.lower()
                batch_col, ocv_col = None, None
                for col in df_raw.columns:
                    if any(k in col for k in batch_keywords) and batch_col is None:
                        batch_col = col
                    if any(k in col for k in ocv_keywords) and ocv_col is None:
                        ocv_col = col
                if batch_col is None or ocv_col is None:
                    st.error(f"檔案 {f.name} 缺少批次 / OCV欄位，已跳過")
                    continue
                df_temp = df_raw[[batch_col, ocv_col]].copy()
                df_temp.columns = ["Batch", "OCV"]
                # 清理多維非法值
                def clean_ocv(val):
                    if isinstance(val, (list, np.ndarray, tuple)):
                        return np.nan
                    try:
                        return float(val)
                    except:
                        return np.nan
                df_temp["OCV"] = df_temp["OCV"].apply(clean_ocv)
                df_clean = df_temp.dropna(subset=["Batch", "OCV"])
                df_clean = df_clean[df_clean["OCV"] > 0]
                if len(df_clean) > 0:
                    all_data.append(df_clean)
            except Exception as e:
                st.error(f"讀取檔案失敗：{str(e)}")
    if len(all_data) > 0:
        return pd.concat(all_data, ignore_index=True)
    else:
        return pd.DataFrame()

# 模擬數據函數（優化：確保生成可配組的有效數據）
def generate_mock_data(current_S):
    batch1 = ["QA2K"] * 150
    volt1 = np.random.uniform(3.560, 3.563, 150)
    batch2 = ["QBNJ"] * 200
    volt2 = np.random.uniform(3.552, 3.555, 200)
    mock_df = pd.DataFrame({"Batch": batch1 + batch2, "OCV": np.concatenate([volt1, volt2])})
    unique_batch = mock_df["Batch"].unique()
    if len(unique_batch) >= 2:
        b_a, b_b = unique_batch[0], unique_batch[1]
        max_v = mock_df[mock_df["Batch"] == b_a]["OCV"].max()
        mock_add = pd.DataFrame({"Batch": [b_b] * current_S, "OCV": [max_v - 0.0005] * current_S})
        mock_df = pd.concat([mock_df, mock_add], ignore_index=True)
    return mock_df

# 配組演算核心
def perform_pairing_engine(df, S, P, base_target, real_pack_cell, threshold, enable_s2):
    result_list = []
    group_meta_list = []
    remain_df = df.copy().reset_index(drop=True)
    remain_df["Cell_ID"] = remain_df.index

    # Stage1 同批次配組
    for batch_name in remain_df["Batch"].unique():
        while True:
            batch_df = remain_df[remain_df["Batch"] == batch_name].sort_values("OCV").reset_index(drop=True)
            if len(batch_df) < real_pack_cell:
                match_flag = False
                break
            match_flag = False
            for start_idx in range(len(batch_df) - real_pack_cell + 1):
                sub_df = batch_df.iloc[start_idx:start_idx + real_pack_cell]
                v_min = sub_df["OCV"].min()
                v_max = sub_df["OCV"].max()
                delta = v_max - v_min
                if delta <= threshold:
                    gid = f"P{len(result_list)+1}-{batch_name}"
                    group_data = sub_df.copy()
                    group_data["GroupID"] = gid
                    group_data["Method"] = "Stage 1"
                    group_data["Partner_Info"] = "單一批次內配組"
                    result_list.append(group_data)
                    group_meta_list.append({
                        "GroupID": gid,
                        "Stage": "Stage 1",
                        "Group_Min_OCV(V)": round(v_min, 6),
                        "Group_Max_OCV(V)": round(v_max, 6),
                        "Group_Delta(mV)": round(delta * 1000, 2),
                        "Cell_Count": real_pack_cell
                    })
                    del_id_list = sub_df["Cell_ID"].tolist()
                    remain_df = remain_df[~remain_df["Cell_ID"].isin(del_id_list)]
                    match_flag = True
                    break
            if not match_flag:
                break
    # Stage2 跨批次混配
    if enable_s2 and len(remain_df) >= real_pack_cell:
        batch_all = remain_df["Batch"].unique()
        if len(batch_all) >= 2:
            for b1, b2 in combinations(batch_all, 2):
                while True:
                    d1 = remain_df[remain_df["Batch"] == b1].sort_values("OCV")
                    d2 = remain_df[remain_df["Batch"] == b2].sort_values("OCV")
                    cross_match = False
                    for n in range(1, P):
                        n1 = S * n
                        n2 = base_target - n1
                        if len(d1) >= n1 and len(d2) >= n2:
                            sub_cross = pd.concat([d1.iloc[:n1], d2.iloc[:n2]])
                            v_min_c = sub_cross["OCV"].min()
                            v_max_c = sub_cross["OCV"].max()
                            delta_c = v_max_c - v_min_c
                            if delta_c <= threshold:
                                gid = f"S2-{b1}x{b2}-Sn{n}"
                                cross_group = sub_cross.copy()
                                cross_group["GroupID"] = gid
                                cross_group["Method"] = "Stage 2"
                                cross_group["Partner_Info"] = f"{b1}({n1}) + {b2}({n2})"
                                result_list.append(cross_group)
                                group_meta_list.append({
                                    "GroupID": gid,
                                    "Stage": "Stage 2",
                                    "Group_Min_OCV(V)": round(v_min_c, 6),
                                    "Group_Max_OCV(V)": round(v_max_c, 6),
                                    "Group_Delta(mV)": round(delta_c * 1000, 2),
                                    "Cell_Count": real_pack_cell
                                })
                                del_id_list = sub_cross["Cell_ID"].tolist()
                                remain_df = remain_df[~remain_df["Cell_ID"].isin(del_id_list)]
                                cross_match = True
                                break
                    if not cross_match:
                        break
    final_group_df = pd.concat(result_list, ignore_index=True) if len(result_list) > 0 else pd.DataFrame()
    group_summary_df = pd.DataFrame(group_meta_list) if len(group_meta_list) > 0 else pd.DataFrame()
    return final_group_df, remain_df, group_summary_df

# Excel匯出函數
def generate_excel_report(raw_df, final_res, batch_report, group_summary_df, S, P, base_target, real_pack_cell, base_mv, enable_buffer, buffer_cell, total_cells, matched_total, total_packs, total_yield):
    output_buf = io.BytesIO()
    wb = openpyxl.Workbook()
    font_name = "Segoe UI"
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F6F9", end_color="F2F6F9", fill_type="solid")
    accent_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    thin_line = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))
    white_font = Font(color="FFFFFF", bold=True, name=font_name)
    bold_font = Font(bold=True, name=font_name)
    normal_font = Font(name=font_name)

    ws_overview = wb.active
    ws_overview.title = "生產稽核總覽"
    ws_overview.merge_cells("A1:L1")
    title_cell = ws_overview.cell(1, 1, "🔋 電池配組精密稽核生產報告")
    title_cell.font = Font(size=18, bold=True, name=font_name)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    param_table = [
        ["運行架構", f"{S}S {P}P", "單PACK標準電芯", base_target],
        ["Buffer開啟", "是" if enable_buffer else "否", "Buffer電芯數", f"{buffer_cell} 顆"],
        ["壓差門檻設定", f"{base_mv} mV", "每組抓取電芯", f"{real_pack_cell} 顆"],
        ["規格壓差上限", "3.00 mV", "總上傳電芯", f"{total_cells} 顆"],
        ["配對完成電芯", f"{matched_total} 顆", "總PACK", f"{total_packs} 組"],
        ["整體配組良率", f"{total_yield:.1f}%", "", ""]
    ]
    for row_idx, row_data in enumerate(param_table, start=3):
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws_overview.cell(row_idx, col_idx, val)
            cell.border = thin_line
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if col_idx in [1, 3]:
                cell.fill = accent_fill
    ws_overview.cell(10, 1, "📊 批次配組匯總").font = Font(size=14, bold=True, name=font_name)
    batch_header = ["批號", "原始總數", "Stage1電芯", "Stage2電芯", "總配對", "未配對", "產出PACK", "良率"]
    for c, h in enumerate(batch_header, start=1):
        cell = ws_overview.cell(11, c, h)
        cell.fill = header_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_line
    for r_idx, b_data in enumerate(batch_report, start=12):
        for c, val in enumerate(b_data.values(), start=1):
            cell = ws_overview.cell(r_idx, c, val)
            cell.border = thin_line
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if r_idx % 2 == 0:
                cell.fill = zebra_fill
    # PACK彙總表
    if not group_summary_df.empty:
        ws_group = wb.create_sheet("PACK組電壓彙總")
        group_header = ["GroupID", "配組階段", "組內最小OCV(V)", "組內最大OCV(V)", "組內壓差(mV)", "本組電芯數"]
        for c, h in enumerate(group_header, start=1):
            cell = ws_group.cell(1, c, h)
            cell.fill = header_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_line
        row_num = 2
        for _, row in group_summary_df.iterrows():
            ws_group.cell(row_num, 1, row["GroupID"])
            ws_group.cell(row_num, 2, row["Stage"])
            cell_min = ws_group.cell(row_num, 3, row["Group_Min_OCV(V)"])
            cell_min.number_format = "0.00000"
            cell_max = ws_group.cell(row_num, 4, row["Group_Max_OCV(V)"])
            cell_max.number_format = "0.00000"
            cell_dv = ws_group.cell(row_num, 5, row["Group_Delta(mV)"])
            cell_dv.number_format = "0.00"
            ws_group.cell(row_num, 6, row["Cell_Count"])
            for col in range(1, 7):
                cell = ws_group.cell(row_num, col)
                cell.border = thin_line
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if row_num % 2 == 0:
                    cell.fill = zebra_fill
            row_num += 1
    # 配組明細
    if not final_res.empty:
        ws_detail = wb.create_sheet("配組明細")
        detail_header = ["Cell_ID", "Batch", "OCV(V)", "GroupID", "配組階段", "配組資訊"]
        for c, h in enumerate(detail_header, start=1):
            cell = ws_detail.cell(1, c, h)
            cell.fill = header_fill
            cell.font = white_font
            cell.border = thin_line
        r = 2
        for _, row in final_res.iterrows():
            ws_detail.cell(r, 1, row["Cell_ID"])
            ws_detail.cell(r, 2, row["Batch"])
            ocv_cell = ws_detail.cell(r, 3, row["OCV"])
            ocv_cell.number_format = "0.00000"
            ws_detail.cell(r, 4, row["GroupID"])
            ws_detail.cell(r, 5, row["Method"])
            ws_detail.cell(r, 6, row["Partner_Info"])
            for col in range(1, 7):
                cell = ws_detail.cell(r, col)
                cell.border = thin_line
                if r % 2 == 0:
                    cell.fill = zebra_fill
            r += 1
    # 原始數據
    ws_raw = wb.create_sheet("原始數據")
    ws_raw.cell(1, 1, "Batch").fill = header_fill
    ws_raw.cell(1, 1).font = white_font
    ws_raw.cell(1, 2, "OCV(V)").fill = header_fill
    ws_raw.cell(1, 2).font = white_font
    r = 2
    for _, row in raw_df.iterrows():
        ws_raw.cell(r, 1, row["Batch"])
        ocv_c = ws_raw.cell(r, 2, row["OCV"])
        ocv_c.number_format = "0.00000"
        for c in [1, 2]:
            cell = ws_raw.cell(r, c)
            cell.border = thin_line
            if r % 2 == 0:
                cell.fill = zebra_fill
        r += 1
    # 自動欄寬
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max_len + 4
    output_buf.seek(0)
    wb.save(output_buf)
    return output_buf.getvalue()

# PDF匯出（固定相同色板，和介面圖完全一致）
def generate_pdf_report(raw_df, final_res, batch_report, group_summary_df, fig_bar, fig_scatter, S, P, base_target, real_pack_cell, base_mv, enable_buffer, buffer_cell, total_cells, matched_total, total_packs, total_yield):
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H:%M:%S")
    bar_buf = io.BytesIO()
    scatter_buf = io.BytesIO()
    # 固定顏色映射，與主介面完全相同
    color_map = {
        "P1-QA2K": "#072280",
        "P2-QBNJ": "#63a8ff",
        "P3-QBNJ": "#e63946",
        "Unmatched": "#ffb3b3"
    }
    # PDF內重建散佈圖，強制綁定色板
    sc_all_pdf = pd.concat([final_res, remain_unmatch.assign(GroupID="Unmatched",Method="Unmatched",Partner_Info="-")], ignore_index=True)
    sc_all_pdf["mV"] = sc_all_pdf["OCV"]*1000
    sc_all_pdf["Seq"] = list(range(len(sc_all_pdf)))
    fig_scatter_pdf = px.scatter(
        sc_all_pdf,
        x="Seq",
        y="mV",
        color="GroupID",
        color_discrete_map=color_map,
        hover_data=["Batch","Method","Partner_Info"],
        labels={"Seq":"電芯序號","mV":"電壓(mV)"},
        title="全域電壓散佈圖"
    )
    fig_scatter_pdf.update_traces(marker=dict(size=6,opacity=0.85))
    fig_scatter_pdf.update_layout(height=480, plot_bgcolor="white")
    pio.write_image(fig_scatter_pdf, scatter_buf, format="png", width=1200, height=500, scale=2)
    pio.write_image(fig_bar, bar_buf, format="png", width=1200, height=400, scale=2)
    bar_buf.seek(0)
    scatter_buf.seek(0)
    bar64 = base64.b64encode(bar_buf.getvalue()).decode()
    sc64 = base64.b64encode(scatter_buf.getvalue()).decode()
    html = f"""
    <html><head><meta charset=utf-8>
    <style>body{{font-family:Segoe UI;margin:30px}}h1{{color:#072280;text-align:center}}h2{{border-bottom:2px #ddebf7;padding-bottom:8px}}.metric{{display:grid;grid-template-columns:1fr 1fr 1fr 1;gap:15;margin:20px 0}}.card{{background:#f2f6f9;padding:15px;border-radius:10;text-align:center}}table{{width:100%;border-collapse:collapse}}th{{background:#072280;color:white;padding:8px}}td,th{{border:1px #ddd solid;padding:7px}}tr:nth-child(even){{background:#f2f6f9}}img{{width:100%;max-width:1100px;margin:20px auto;display:block}}</style></head>
    <body>
    <h1>🔋 Dynapack 電池配組稽核報告</h1>
    <p style="text-align:center;color:#666">產生時間：{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
    <h2>核心指標</h2>
    <div class="metric">
        <div class="card"><h3>總電芯</h3><p>{total_cells}</p></div>
        <div class="card"><h3>配對電芯</h3><p>{matched_total}</p></div>
        <div class="card"><h3>PACK組數</h3><p>{total_packs}</p></div>
        <div class="card"><h3>良率</h3><p>{total_yield:.1f}%</p></div>
    </div>
    <h2>PACK電壓清單</h2>{group_summary_df.to_html(index=False)}
    <h2>批次分佈長條圖</h2><img src="data:image/png;base64,{bar64}">
    <h2>全域電壓散佈圖</h2><img src="data:image/png;base64,{sc64}">
    </body></html>
    """
    opt = {
        "page-size": "A4",
        "margin-top": "8",
        "margin-bottom": "8",
        "margin-left": "8",
        "margin-right": "8",
        "encoding": "UTF-8",
        "quiet": ""
    }
    return pdfkit.from_string(html, False, configuration=pdf_conf, options=opt)

# 上傳區塊
upload_files = st.file_uploader("📂 上傳Excel / CSV", accept_multiple_files=True, key=st.session_state.uploader_key)
df_upload = load_and_fix_data(upload_files)
st.session_state.upload_raw_data = df_upload

# 【核心邏輯：優先取用模擬數據，不會被上傳檔案覆蓋】
if not st.session_state.mock_data_df.empty:
    raw_df = st.session_state.mock_data_df
else:
    raw_df = st.session_state.upload_raw_data

# 兩按鈕並排：模擬數據 / 執行配組
col_btn1, col_btn2 = st.columns([1,1])
with col_btn1:
    mock_btn = st.button("🧪 載入模擬測試數據", key="mock_data_btn", use_container_width=True)
    if mock_btn:
        # 生成模擬數據存入獨立緩存mock_data_df，不影響upload_raw_data
        mock_data = generate_mock_data(S)
        st.session_state.mock_data_df = mock_data
        # 清空舊分析結果
        clear_analysis_keys = [
            "final_group_df", "remain_unmatch", "group_summary_df",
            "batch_report_list", "fig_bar", "fig_scatter",
            "total_cells", "matched_total", "unmatch_total",
            "total_pack_num", "yield_rate"
        ]
        df_keys = ["final_group_df", "remain_unmatch", "group_summary_df"]
        for k in clear_analysis_keys:
            if k in df_keys:
                st.session_state[k] = pd.DataFrame()
            else:
                st.session_state[k] = None
        st.success("✅ 已載入模擬批次數據，可執行配組")
        st.rerun()
with col_btn2:
    run_btn = st.button("🚀 執行配組分析", key="run_analysis_btn", use_container_width=True)

st.markdown("---")

# 無數據提示
if raw_df.empty:
    st.warning("⚠️ 無數據，請上傳檔案或點擊「載入模擬測試數據」按鈕")

# 執行配組邏輯
if run_btn:
    if raw_df.empty:
        st.error("⚠️ 無數據，請上傳檔案或點左側模擬數據按鈕")
        st.stop()
    with st.spinner("演算配組中..."):
        total_cells = len(raw_df)
        final_group_df, remain_unmatch, group_summary_df = perform_pairing_engine(raw_df, S, P, base_target, real_pack_cell, base_threshold, enable_s2)
        matched_total = len(final_group_df)
        unmatch_total = len(remain_unmatch)
        total_pack_num = final_group_df["GroupID"].nunique() if not final_group_df.empty else 0
        yield_rate = (matched_total / total_cells)*100 if total_cells>0 else 0

        st.session_state.total_cells = total_cells
        st.session_state.matched_total = matched_total
        st.session_state.unmatch_total = unmatch_total
        st.session_state.total_pack_num = total_pack_num
        st.session_state.yield_rate = yield_rate
        st.session_state.final_group_df = final_group_df
        st.session_state.remain_unmatch = remain_unmatch
        st.session_state.group_summary_df = group_summary_df

        batch_report_list = []
        chart_source = []
        for b in raw_df["Batch"].unique():
            b_all = len(raw_df[raw_df["Batch"]==b])
            s1 = len(final_group_df[(final_group_df["Method"]=="Stage 1") & (final_group_df["Batch"]==b)]) if not final_group_df.empty else 0
            s2 = len(final_group_df[(final_group_df["Method"]=="Stage 2") & (final_group_df["Batch"]==b)]) if not final_group_df.empty else 0
            match = s1 + s2
            un = b_all - match
            s1_pack = final_group_df[(final_group_df["Method"]=="Stage 1") & (final_group_df["Batch"]==b)]["GroupID"].nunique() if not final_group_df.empty else 0
            s2_add = 0
            if not final_group_df.empty:
                for gid in final_group_df[(final_group_df["Method"]=="Stage 2") & (final_group_df["Batch"]==b)]["GroupID"].unique():
                    s2_add += len(final_group_df[final_group_df["GroupID"]==gid]) / base_target
            pack_total = s1_pack + s2_add
            y = f"{(match/b_all)*100:.1f}" if b_all>0 else "0.0"
            batch_report_list.append({
                "批號(Batch)":b,"原始總數":b_all,"Stage1":s1,"Stage2":s2,"總配對":match,"未配對":un,"PACK數":round(pack_total,3),"良率":y+"%"
            })
            chart_source.append({"批次":b,"數量":s1,"分類":"Stage1配組"})
            chart_source.append({"批次":b,"數量":s2,"分類":"Stage2跨批"})
            chart_source.append({"批次":b,"數量":un,"分類":"無法配組"})
        st.session_state.batch_report_list = batch_report_list
        fig_bar = px.bar(pd.DataFrame(chart_source), x="數量", y="批次", color="分類", orientation="h", color_discrete_map={"Stage1配組":"#2ecc71","Stage2跨批":"#f39c12","無法配組":"#e74c3c"})
        fig_bar.update_layout(height=260, plot_bgcolor="white")
        st.session_state.fig_bar = fig_bar
        sc_all = pd.concat([final_group_df, remain_unmatch.assign(GroupID="Unmatched",Method="Unmatched",Partner_Info="-")], ignore_index=True)
        sc_all["mV"] = sc_all["OCV"]*1000
        sc_all["Seq"] = list(range(len(sc_all)))
        # 主介面圖固定色板
        color_map = {
            "P1-QA2K": "#072280",
            "P2-QBNJ": "#63a8ff",
            "P3-QBNJ": "#e63946",
            "Unmatched": "#ffb3b3"
        }
        fig_scatter = px.scatter(
            sc_all,
            x="Seq",
            y="mV",
            color="GroupID",
            color_discrete_map=color_map,
            hover_data=["Batch","Method","Partner_Info"],
            labels={"Seq":"電芯序號","mV":"電壓(mV)"}
        )
        fig_scatter.update_traces(marker=dict(size=6,opacity=0.85))
        fig_scatter.update_layout(height=480, plot_bgcolor="white")
        st.session_state.fig_scatter = fig_scatter

# 渲染結果區段（安全判斷None）
if not raw_df.empty and st.session_state.final_group_df is not None and not st.session_state.final_group_df.empty:
    tc = st.session_state.total_cells
    mt = st.session_state.matched_total
    tp = st.session_state.total_pack_num
    yr = st.session_state.yield_rate
    gs = st.session_state.group_summary_df
    br = st.session_state.batch_report_list
    fb = st.session_state.fig_bar
    fs = st.session_state.fig_scatter

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("總電芯", f"{tc} 顆")
    c2.metric("配對電芯", f"{mt} 顆")
    c3.metric("總PACK", f"{tp} 組")
    c4.metric("良率", f"{yr:.1f}%")
    st.markdown("---")
    st.subheader("PACK電壓一覽")
    st.dataframe(gs, use_container_width=True)
    st.markdown("---")
    st.subheader("批次分佈長條圖")
    st.plotly_chart(fb, use_container_width=True)
    st.markdown("---")
    st.subheader("批次統計表")
    st.dataframe(pd.DataFrame(br), use_container_width=True)
    st.markdown("---")
    st.subheader("全域電壓散佈圖")
    st.plotly_chart(fs, use_container_width=True)
    st.markdown("---")
    st.subheader("報告下載")
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_bytes = generate_excel_report(raw_df, st.session_state.final_group_df, br, gs, S, P, base_target, real_pack_cell, base_mv, enable_buffer, buffer_cell, tc, mt, tp, yr)
    st.download_button("📥 下載Excel報告", excel_bytes, f"配組報告_{ts}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    try:
        pdf_bytes = generate_pdf_report(raw_df, st.session_state.final_group_df, br, gs, fb, fs, S, P, base_target, real_pack_cell, base_mv, enable_buffer, buffer_cell, tc, mt, tp, yr)
        st.download_button("📄 下載PDF報告", pdf_bytes, f"配組報告_{ts}.pdf")
    except Exception as pdf_err:
        st.warning(f"PDF產生失敗（wkhtmltopdf異常），僅可下載Excel：{pdf_err}")
    st.success("✅ 分析完成，可下載Excel報告")