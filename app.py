import streamlit as st
import pandas as pd
from itertools import combinations
import plotly.express as px
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime
import plotly.io as pio
import base64
import uuid

# 全域固定顏色對照
FIX_COLOR_MAP = {
    "Unmatched": "#ffb3b3"
}
color_list = [
    "#072280", "#3366cc", "#28a745", "#ffc107", "#fd7e14", "#6f42c1", "#20c997", "#e83e8c"
]

# 頁面全域設定
st.set_page_config(page_title="Dynapack Cell Regroup System", layout="wide", page_icon="🔋")

# 自訂按鈕樣式
st.markdown("""
<style>
button[key="run_analysis_btn"] {
    background-color: #072280 !important;
    color: white !important;
    font-size: 20px !important;
    padding: 18px 60px !important;
    border-radius: 12px !important;
    border: none !important;
    width: 100% !important;
    font-weight: 600;
}
button[key="run_analysis_btn"]:hover {
    background-color: #051a60 !important;
    transform: scale(1.02);
    transition: 0.2s;
}
button[key="mock_data_btn"] {
    font-size: 16px;
    padding: 12px;
    width: 100%;
}
</style>
""", unsafe_allow_html=True)

# Session狀態初始化
init_keys = [
    "login", "need_reset", "uploader_key", "upload_raw_data", "mock_data_df",
    "final_group_df", "remain_unmatch", "group_summary_df", "batch_report_list",
    "fig_bar", "fig_scatter", "total_cells", "matched_total", "unmatch_total",
    "total_pack_num", "yield_rate"
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

# 登入頁
if not st.session_state.login:
    st.markdown("""
    <style>
    .login-container {max-width:860px;margin:50px auto;padding:60px 40px;text-align:center;background:linear-gradient(130deg,#fff,#eff5ff);border-radius:24px;box-shadow:0 12px 45px rgba(10,35,100,0.13);}
    .welcome-title {font-size:32px;font-weight:700;color:#072280;margin:32px 0;letter-spacing:0.7px;}
    .stTextInput {max-width:480px;margin:0 auto 30px;}
    .stButton>button {background:#072280;color:white;font-size:17px;padding:13px 45px;border-radius:10px;border:none;}
    </style>
    """, unsafe_allow_html=True)
    with st.container():
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        try:
            st.image("dynapack_logo.png", width=680)
        except:
            st.markdown("# 🔋 DYNAPACK", unsafe_allow_html=True)
        st.markdown('<div class="welcome-title">Welcome to Dynapack Non-matchable Cell regrouping System </div>', unsafe_allow_html=True)
        pwd = st.text_input("請輸入系統密碼", type="password")
        if st.button("登入系統"):
            if pwd == "3211":
                st.session_state.login = True
                st.rerun()
            else:
                st.error("密碼錯誤")
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# 重置按鈕
def trigger_reset():
    st.session_state.need_reset = True
col_title, col_reset = st.columns([9,1])
with col_title:
    st.title("🔋 Dynapack Non-matchable Cell regrouping System")
with col_reset:
    st.button("🔄 重置所有資料", on_click=trigger_reset, type="secondary")
if st.session_state.need_reset:
    clear_keys = [
        "upload_raw_data","mock_data_df","final_group_df","remain_unmatch","group_summary_df",
        "batch_report_list","fig_bar","fig_scatter","total_cells","matched_total",
        "unmatch_total","total_pack_num","yield_rate"
    ]
    for k in clear_keys:
        if k in ["upload_raw_data","mock_data_df","final_group_df","remain_unmatch","group_summary_df"]:
            st.session_state[k] = pd.DataFrame()
        else:
            st.session_state[k] = None
    st.session_state.uploader_key = str(uuid.uuid4())
    st.session_state.need_reset = False
    st.rerun()

# 側邊欄參數設定
with st.sidebar:
    st.header("⚙️ 分析參數 (Configuration)")
    S = st.number_input("串聯數 (S)", min_value=1, max_value=24, value=14, key="S_input")
    P = st.number_input("並聯數 (P)", min_value=1, max_value=6, value=6, key="P_input")
    base_target = S * P
    base_mv = st.select_slider("Max Delta V 基礎門檻(mV)", options=[1,2,3,4,5,6,7,8,9,10], value=3)
    base_threshold = base_mv / 1000

    st.markdown("#### 🧩 Buffer電芯容差設定")
    enable_buffer = st.checkbox("開啟自放電/設備誤差Buffer容差", value=False)
    buffer_cell = st.select_slider("Buffer備用電芯數", options=[0,1,2,3], value=1)

    # Buffer配組數邏輯：未開啟不追加電芯
    if enable_buffer:
        real_pack_cell = base_target + buffer_cell
        st.caption(f"已開啟Buffer，單組配組電芯數：{real_pack_cell}（基礎{S}×{P} + Buffer{buffer_cell}）")
    else:
        real_pack_cell = base_target
        st.caption(f"未開啟Buffer，單組配組電芯數：{real_pack_cell}（基礎{S}×{P}，無額外電芯）")

    st.markdown("---")
    st.header("🔬 演算法驗證模式")
    enable_s2 = st.checkbox("Sn Ratio (啟用第二階段動態跨批次)", value=True)

# 讀取檔案函數
def load_and_fix_data(uploaded_files):
    all_data = []
    batch_keys = ['batch','批次','批號','batch_no','batchid']
    ocv_keys = ['ocv','電壓','voltage','v','電池電壓']
    if uploaded_files:
        for f in uploaded_files:
            try:
                if f.name.endswith((".xlsx",".xls")):
                    df_raw = pd.read_excel(f)
                else:
                    df_raw = pd.read_csv(f, encoding_errors="replace")
                df_raw.columns = df_raw.columns.str.strip().str.lower()
                b_col, o_col = None, None
                for c in df_raw.columns:
                    if not b_col and any(k in c for k in batch_keys):
                        b_col = c
                    if not o_col and any(k in c for k in ocv_keys):
                        o_col = c
                if not b_col or not o_col:
                    st.error(f"{f.name} 缺少批次/OCV欄位，跳過")
                    continue
                temp = df_raw[[b_col, o_col]].copy()
                temp.columns = ["Batch","OCV"]
                def clean(v):
                    if isinstance(v,(list,np.ndarray,tuple)):
                        return np.nan
                    try: return float(v)
                    except: return np.nan
                temp["OCV"] = temp["OCV"].apply(clean)
                temp = temp.dropna(subset=["Batch","OCV"])
                temp = temp[temp["OCV"]>0]
                if len(temp):
                    all_data.append(temp)
            except Exception as e:
                st.error(f"讀取失敗：{e}")
    return pd.concat(all_data) if all_data else pd.DataFrame()

# 模擬測試數據
def generate_mock_data(current_S):
    b1 = ["QA2K"]*150
    v1 = np.random.uniform(3.560,3.563,150)
    b2 = ["QBNJ"]*200
    v2 = np.random.uniform(3.552,3.555,200)
    df = pd.DataFrame({"Batch":b1+b2, "OCV":np.concatenate([v1,v2])})
    uni = df["Batch"].unique()
    if len(uni)>=2:
        a,b = uni[0],uni[1]
        maxv = df[df.Batch==a]["OCV"].max()
        add = pd.DataFrame({"Batch":[b]*current_S, "OCV":[maxv-0.0005]*current_S})
        df = pd.concat([df,add], ignore_index=True)
    return df

# 【徹底重寫配組演算法，嚴格限制組內壓差≤門檻】
def perform_pairing_engine(df, S, P, base_target, real_pack_cell, threshold, enable_s2):
    res_list = []
    meta_list = []
    remain = df.copy().reset_index(drop=True)
    remain["Cell_ID"] = remain.index

    # Stage1 同批次配組 改良滑動視窗
    for b_name in remain["Batch"].unique():
        batch_all = remain[remain.Batch == b_name].sort_values("OCV").reset_index(drop=True)
        i = 0
        total_len = len(batch_all)
        while i <= total_len - real_pack_cell:
            window = batch_all.iloc[i:i+real_pack_cell]
            delta_v = window["OCV"].max() - window["OCV"].min()
            # 強制判斷：只有壓差完全小於等於門檻才會建立PACK
            if delta_v <= threshold:
                gid = f"P{len(res_list)+1}-{b_name}"
                window["GroupID"] = gid
                window["Method"] = "Stage 1"
                window["Partner_Info"] = "單一批次內配組"
                res_list.append(window)
                meta_list.append({
                    "GroupID":gid,"Stage":"Stage 1",
                    "Group_Min_OC(V)":round(window["OCV"].min(),6),
                    "Group_Max_OC(V)":round(window["OCV"].max(),6),
                    "Group_Delta(mV)":round(delta_v*1000,2),
                    "Cell_Count":real_pack_cell
                })
                # 移除已使用電芯，重新從頭滑動
                used_idx = window.index.tolist()
                batch_all = batch_all.drop(used_idx).reset_index(drop=True)
                total_len = len(batch_all)
                i = 0
            else:
                # 壓差超標，向後移一格繼續檢查
                i += 1
        # 更新全域剩餘未配組電芯
        used_ids = []
        for g_df in res_list:
            used_ids.extend(g_df["Cell_ID"].tolist())
        remain = remain[~remain.Cell_ID.isin(used_ids)]

    # Stage2 跨批次配組 同樣嚴格壓差判斷
    if enable_s2 and len(remain)>=real_pack_cell:
        batches = remain.Batch.unique()
        for b1,b2 in combinations(batches,2):
            d1_all = remain[remain.Batch==b1].sort_values("OCV").reset_index(drop=True)
            d2_all = remain[remain.Batch==b2].sort_values("OCV").reset_index(drop=True)
            for n in range(1,P):
                n1 = S*n
                n2 = base_target - n1
                if len(d1_all)>=n1 and len(d2_all)>=n2:
                    i = 0
                    while i <= len(d1_all)-n1:
                        j = 0
                        while j <= len(d2_all)-n2:
                            sub1 = d1_all.iloc[i:i+n1]
                            sub2 = d2_all.iloc[j:j+n2]
                            combine = pd.concat([sub1,sub2])
                            dv = combine["OCV"].max() - combine["OCV"].min()
                            if dv <= threshold:
                                gid = f"S2-{b1}x{b2}-Sn{n}"
                                combine["GroupID"] = gid
                                combine["Method"] = "Stage 2"
                                combine["Partner_Info"] = f"{b1}({n1}) + {b2}"
                                res_list.append(combine)
                                meta_list.append({
                                    "GroupID":gid,"Stage":"Stage 2",
                                    "Group_Min_OC(V)":round(combine["OCV"].min(),6),
                                    "Group_Max_OC(V)":round(combine["OCV"].max(),6),
                                    "Group_Delta(mV)":round(dv*1000,2),
                                    "Cell_Count":real_pack_cell
                                })
                                used_cid = combine["Cell_ID"].tolist()
                                remain = remain[~remain.Cell_ID.isin(used_cid)]
                                d1_all = d1_all[~d1_all.Cell_ID.isin(used_cid)].reset_index(drop=True)
                                d2_all = d2_all[~d2_all.Cell_ID.isin(used_cid)].reset_index(drop=True)
                                i = 0
                                j = 0
                            else:
                                j += 1
                        i += 1

    final = pd.concat(res_list, ignore_index=True) if res_list else pd.DataFrame()
    meta_df = pd.DataFrame(meta_list) if meta_list else pd.DataFrame()
    return final, remain, meta_df

# Excel匯出函數（修正色碼#ffffff，消除顏色報錯）
def generate_excel_report(raw_df, final_res, batch_report, group_summary, S, P, base_target, real_pack_cell, base_mv, enable_buffer, buffer_cell, tc, mt, tp, yr):
    buf = io.BytesIO()
    wb = openpyxl.Workbook()
    font = "Segoe UI"
    h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    z_fill = PatternFill(start_color="F2F6F9", end_color="F2F6F9", fill_type="solid")
    a_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    thin = Border(left=Side("thin"),right=Side("thin"),top=Side("thin"),bottom=Side("thin"))
    # 修正：openpyxl要求6位完整hex色碼 #ffffff，不再使用#fff簡寫
    w_font = Font(color="#ffffff", bold=True, name=font)
    b_font = Font(bold=True, name=font)
    # 總覽頁
    ws1 = wb.active
    ws1.title = "生產稽核總覽"
    ws1.merge_cells("A1:L1")
    ws1.cell(1,1,"🔋 電池配組精密稽核生產報告").font = Font(size=18,bold=True,name=font)
    table = [
        ["運行架構",f"{S}S{P}P","單PACK標準",base_target],
        ["Buffer開啟","是" if enable_buffer else "否","Buffer數",buffer_cell],
        ["壓差門檻",f"{base_mv}mV","每組電芯",real_pack_cell],
        ["規格上限","3.00mV","總電芯",tc],
        ["配對電芯",mt,"總PACK",tp],
        ["良率",f"{yr:.1f}%","",""]
    ]
    for r,row in enumerate(table,3):
        for c,val in enumerate(row,1):
            cell = ws1.cell(r,c,val)
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if c in [1,3]:
                cell.fill = a_fill
    ws1.cell(10,1,"📊 批次配組匯總").font = Font(size=14,bold=True,name=font)
    batch_header = ["批號","原始總數","Stage1電芯","Stage2電芯","總配對","未配對","PACK數","良率"]
    for c,h in enumerate(batch_header,1):
        cell = ws1.cell(11,c,h)
        cell.fill = h_fill
        cell.font = w_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin
    for r_idx,b_data in enumerate(batch_report,12):
        for c,val in enumerate(b_data.values(),1):
            cell = ws1.cell(r_idx,c,val)
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if r_idx % 2 == 0:
                cell.fill = z_fill
    # PACK彙總頁
    if not group_summary.empty:
        ws_group = wb.create_sheet("PACK組電壓彙總")
        g_header = ["GroupID","配組階段","最小OC(V)","最大OC(V)","壓差(mV)","電芯數量"]
        for c,h in enumerate(g_header,1):
            cell = ws_group.cell(1,c,h)
            cell.fill = h_fill
            cell.font = w_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin
        row_num = 2
        for _,row in group_summary.iterrows():
            ws_group.cell(row_num,1,row["GroupID"])
            ws_group.cell(row_num,2,row["Stage"])
            ws_group.cell(row_num,3,row["Group_Min_OC(V)"]).number_format = "0.00000"
            ws_group.cell(row_num,4,row["Group_Max_OC(V)"]).number_format = "0.00000"
            ws_group.cell(row_num,5,row["Group_Delta(mV)"]).number_format = "0.00"
            ws_group.cell(row_num,6,row["Cell_Count"])
            for col in range(1,7):
                cell = ws_group.cell(row_num,col)
                cell.border = thin
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if row_num % 2 == 0:
                    cell.fill = z_fill
            row_num += 1
    # 配組明細頁
    if not final_res.empty:
        ws_detail = wb.create_sheet("配組明細")
        d_header = ["Cell_ID","Batch","OC(V)","GroupID","配組階段","同組搭配批次"]
        for c,h in enumerate(d_header,1):
            cell = ws_detail.cell(1,c,h)
            cell.fill = h_fill
            cell.font = w_font
            cell.border = thin
        r = 2
        for _,row in final_res.iterrows():
            ws_detail.cell(r,1,row["Cell_ID"])
            ws_detail.cell(r,2,row["Batch"])
            ocv_cell = ws_detail.cell(r,3,row["OCV"])
            ocv_cell.number_format = "0.00000"
            ws_detail.cell(r,4,row["GroupID"])
            ws_detail.cell(r,5,row["Method"])
            ws_detail.cell(r,6,row["Partner_Info"])
            for col in range(1,7):
                cell = ws_detail.cell(r,col)
                cell.border = thin
                if r % 2 == 0:
                    cell.fill = z_fill
            r += 1
    # 原始數據頁
    ws_raw = wb.create_sheet("原始數據")
    ws_raw.cell(1,1,"Batch").fill = h_fill
    ws_raw.cell(1,1).font = w_font
    ws_raw.cell(1,2,"OC(V)").fill = h_fill
    ws_raw.cell(1,2).font = w_font
    r = 2
    for _,row in raw_df.iterrows():
        ws_raw.cell(r,1,row["Batch"])
        ocv_c = ws_raw.cell(r,2,row["OCV"])
        ocv_c.number_format = "0.00000"
        for c in [1,2]:
            cell = ws_raw.cell(r,c)
            cell.border = thin
            if r % 2 == 0:
                cell.fill = z_fill
        r += 1
    # 自動欄寬
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max_len + 4
    buf.seek(0)
    wb.save(buf)
    return buf.getvalue()

# 上傳檔案區塊
upload_files = st.file_uploader("📂 上傳Excel / CSV", accept_multiple_files=True, key=st.session_state.uploader_key)
df_upload = load_and_fix_data(upload_files)
st.session_state.upload_raw_data = df_upload

# 優先取用模擬數據
if not st.session_state.mock_data_df.empty:
    raw_df = st.session_state.mock_data_df
else:
    raw_df = st.session_state.upload_raw_data

# 功能按鈕
col_btn1, col_btn2 = st.columns([1,1])
with col_btn1:
    mock_btn = st.button("🧪 載入模擬測試數據", key="mock_data_btn", use_container_width=True)
    if mock_btn:
        mock_data = generate_mock_data(S)
        st.session_state.mock_data_df = mock_data
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

# 執行配組演算
if run_btn:
    if raw_df.empty:
        st.error("⚠️ 無數據，請上傳檔案或載入模擬數據")
        st.stop()
    with st.spinner("演算配組中..."):
        total_cells = len(raw_df)
        final_group_df, remain_unmatch, group_summary_df = perform_pairing_engine(raw_df, S, P, base_target, real_pack_cell, base_threshold, enable_buffer)
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
        color_map = {"Unmatched": "#ffb3b3"}
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

# 結果頁面渲染
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
    # 函數名稱完全匹配定義generate_excel_report，消除NameError
    excel_bytes = generate_excel_report(raw_df, st.session_state.final_group_df, br, gs, S, P, base_target, real_pack_cell, base_mv, enable_buffer, buffer_cell, tc, mt, tp, yr)
    st.download_button("📥 下載Excel報告", excel_bytes, f"配組報告_{ts}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.success("✅ 分析完成，僅支援Excel下載")