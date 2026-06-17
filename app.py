import streamlit as st
import pandas as pd
from itertools import combinations
import plotly.express as px
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 頁面配置
st.set_page_config(page_title="電池配組與稽核系統", layout="wide")

# 標題與說明
st.title("🔋 Dynapack Non-matchable Cell regrouping System")
st.markdown("本版本已整合電芯去向堆疊圖、全域電壓散佈圖、Excel一鍵匯出，新增Buffer電芯容差功能。")

# --- 側邊欄參數設定 ---
with st.sidebar:
    st.header("⚙️ 分析參數 (Configuration)")
    S = st.number_input("串聯數 (S)", 1, 24, 14, key="S_input")
    P = st.number_input("並聯數 (P)", 1, 20, 6, key="P_input")
    target = S * P  # 單Pack標準基礎電芯數
    base_mv = st.select_slider("Max Delta V 基礎門檻(mV)", options=[1,2,3,4,5,6,7,8,9,10], value=3, key="threshold_input")
    base_threshold = base_mv / 1000

    st.markdown("#### 🧩 Buffer電芯容差設定")
    enable_buffer = st.checkbox("開啟自放電/設備誤差Buffer容差", value=False, key="buffer_switch")
    buffer_cell = st.select_slider("Buffer備用電芯數", options=[0,1,2,3], value=1, key="buffer_slider")
    st.caption("每顆Buffer增加0.05mV容差；整組壓差永久上限3mV；開啟後每組配組電芯數=基礎數+Buffer數")

    if enable_buffer:
        add_mv = buffer_cell * 0.05
        real_mv = min(base_mv + add_mv, 3.0)
        real_threshold = real_mv / 1000
        real_pack_cell = target + buffer_cell
        st.info(f"當前有效配組門檻：{real_mv:.2f} mV（基礎{base_mv}mV + Buffer{add_mv:.2f}mV，上限3mV）\n每組抓取電芯：{real_pack_cell} 顆（基礎{target}+Buffer{buffer_cell}）")
    else:
        real_mv = base_mv
        real_threshold = base_threshold
        real_pack_cell = target
        st.info(f"未開啟Buffer，使用基礎門檻：{real_mv} mV，每組抓取電芯：{real_pack_cell} 顆")

    st.markdown("---")
    st.header("🔬 演算法驗證模式")
    enable_s2 = st.checkbox("Sn Ratio (啟用第二階段動態跨批次)", value=True, key="enable_s2_input")
    inject_mock_data = st.checkbox("🔥 強制注入測試電芯 (用於驗證第二階段)", value=False, key="inject_mock_input")
    st.caption(f"開啟後，無上傳檔案時自動生成測試數據，方便測試Stage2跨批次配組。")

# --- 資料讀取引擎 ---
def load_and_fix_data(uploaded_files, inject_test=False, current_S=16):
    all_data = []
    batch_keywords = ['batch', '批次', '批號', 'batch_no', 'batchid']
    ocv_keywords = ['ocv', '電壓', 'voltage', 'v', '電池電壓']
    
    if uploaded_files:
        for f in uploaded_files:
            try:
                if f.name.endswith('.xlsx') or f.name.endswith('.xls'):
                    df = pd.read_excel(f)
                else:
                    df = pd.read_csv(f, encoding_errors='replace', engine='python')
                
                df.columns = df.columns.str.strip().str.lower()
                batch_col = None
                ocv_col = None
                
                for col in df.columns:
                    if any(keyword in col for keyword in batch_keywords) and batch_col is None:
                        batch_col = col
                    if any(keyword in col for keyword in ocv_keywords) and ocv_col is None:
                        ocv_col = col
                
                if batch_col is None or ocv_col is None:
                    st.error(f"檔案 {f.name} 找不到批次/電壓欄位")
                    continue
                
                df_clean = df[[batch_col, ocv_col]].copy()
                df_clean.columns = ['Batch', 'OCV']
                df_clean['OCV'] = pd.to_numeric(df_clean['OCV'], errors='coerce')
                df_clean = df_clean.dropna(subset=['Batch', 'OCV'])
                df_clean = df_clean[df_clean['OCV'] > 0]
                
                if len(df_clean) > 0:
                    all_data.append(df_clean)
                else:
                    st.warning(f"檔案 {f.name} 電壓欄無有效數值，已跳過")
            except Exception as e:
                st.error(f"讀取 {f.name} 失敗: {e}")
    
    main_df = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
    
    if inject_test and main_df.empty:
        b1 = ['QA2K'] * 121
        v1 = np.random.uniform(3.560, 3.563, 121)
        b2 = ['QBNJ'] * 188
        v2 = np.random.uniform(3.552, 3.555, 188)
        main_df = pd.DataFrame({'Batch': b1 + b2, 'OCV': np.concatenate([v1, v2])})
        unique_batches = main_df['Batch'].unique()
        if len(unique_batches) >= 2:
            batch_a = unique_batches[0]
            batch_b = unique_batches[1]
            base_v = main_df[main_df['Batch'] == batch_a]['OCV'].max()
            mock_cells = pd.DataFrame({'Batch': [batch_b]*current_S, 'OCV': [base_v - 0.0005]*current_S})
            main_df = pd.concat([main_df, mock_cells], ignore_index=True)
    elif inject_test and not main_df.empty:
        st.info("已上傳真實數據，測試數據不注入")
    return main_df

# --- 配組核心引擎（修復drop index錯誤，改用Cell_ID過濾）
def perform_pairing_engine(df, S, P, base_target, real_pack_cell, threshold, enable_s2):
    results = []
    group_meta = []
    remaining = df.copy().reset_index(drop=True)
    remaining['Cell_ID'] = remaining.index

    # Stage1 同批次滑動視窗
    for batch in remaining['Batch'].unique():
        while True:
            b_df = remaining[remaining['Batch'] == batch].sort_values('OCV').reset_index(drop=True)
            if len(b_df) < real_pack_cell:
                matched = False
                break
            matched = False
            for start_idx in range(len(b_df) - real_pack_cell + 1):
                subset = b_df.iloc[start_idx:start_idx + real_pack_cell]
                v_min = subset['OCV'].min()
                v_max = subset['OCV'].max()
                delta_v = v_max - v_min
                if delta_v <= threshold:
                    gid = f"P{len(results)+1}-{batch}"
                    group = subset.copy()
                    group['GroupID'] = gid
                    group['Method'] = "Stage 1"
                    group['Partner_Info'] = "單一批次內配組"
                    results.append(group)
                    group_meta.append({
                        "GroupID": gid,
                        "Stage": "Stage 1",
                        "Group_Min_OCV(V)": round(v_min,6),
                        "Group_Max_OCV(V)": round(v_max,6),
                        "Group_Delta(mV)": round(delta_v*1000,2),
                        "Cell_Count": real_pack_cell
                    })
                    # 修正：透過Cell_ID刪除，不再使用index
                    del_ids = subset["Cell_ID"].tolist()
                    remaining = remaining[~remaining["Cell_ID"].isin(del_ids)]
                    matched = True
                    break
            if not matched:
                break
    # Stage2 跨批次Sn混配
    if enable_s2 and len(remaining) >= real_pack_cell:
        batches = remaining['Batch'].unique()
        if len(batches) >= 2:
            for b1, b2 in combinations(batches, 2):
                while True:
                    d1 = remaining[remaining['Batch'] == b1].sort_values('OCV')
                    d2 = remaining[remaining['Batch'] == b2].sort_values('OCV')
                    pair_found = False
                    for n in range(1, P):
                        n1 = S * n
                        n2 = base_target - n1
                        if len(d1) >= n1 and len(d2) >= n2:
                            subset = pd.concat([d1.iloc[:n1], d2.iloc[:n2]])
                            v_min = subset['OCV'].min()
                            v_max = subset['OCV'].max()
                            delta_v = v_max - v_min
                            if delta_v <= threshold:
                                gid = f"S2-{b1}x{b2}-Sn{n}"
                                group = subset.copy()
                                group['GroupID'] = gid
                                group['Method'] = "Stage 2"
                                group['Partner_Info'] = f"{b1}({n1}) + {b2}({n2})"
                                results.append(group)
                                group_meta.append({
                                    "GroupID": gid,
                                    "Stage": "Stage 2",
                                    "Group_Min_OCV(V)": round(v_min,6),
                                    "Group_Max_OCV(V)": round(v_max,6),
                                    "Group_Delta(mV)": round(delta_v*1000,2),
                                    "Cell_Count": real_pack_cell
                                })
                                # 修正：透過Cell_ID刪除
                                del_ids = subset["Cell_ID"].tolist()
                                remaining = remaining[~remaining["Cell_ID"].isin(del_ids)]
                                pair_found = True
                                break
                    if not pair_found:
                        break
    final_res = pd.concat(results, ignore_index=True) if results else pd.DataFrame()
    group_summary_df = pd.DataFrame(group_meta) if group_meta else pd.DataFrame()
    return final_res, remaining, group_summary_df

# --- Excel匯出函數
def generate_excel_report(raw_df, final_res, batch_report, group_summary_df, S, P, base_target, real_pack_cell, base_mv, enable_buffer, buffer_cell, real_mv, total_cells, matched_count, total_packs, total_yield, threshold_mv):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    font_family = "Segoe UI"
    navy_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F6F9", end_color="F2F6F9", fill_type="solid")
    accent_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    thin_border = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    white_font = Font(color="FFFFFF", bold=True, name=font_family)
    bold_font = Font(bold=True, name=font_family)
    normal_font = Font(name=font_family)

    ws1 = wb.active
    ws1.title = "生產稽核總覽"
    ws1.merge_cells('A1:L1')
    ws1.cell(1, 1, "🔋 電池配組與精密稽核生產報告").font = Font(size=18, bold=True)
    ws1.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")
    param_data = [
        ["運行架構", f"{S}S {P}P", "單Pack標準電芯", base_target],
        ["Buffer開啟", "是" if enable_buffer else "否", "Buffer電芯數", f"{buffer_cell} 顆"],
        ["每組實際抓取", f"{real_pack_cell} 顆", "單顆Buffer容差", "0.05 mV"],
        ["基礎壓差門檻", f"{base_mv} mV", "實際配組門檻", f"{real_mv:.2f} mV"],
        ["規格壓差上限", "3.00 mV", "總上傳電芯", f"{total_cells} 顆"],
        ["配對完成電芯", f"{matched_count} 顆", "總PACK", f"{total_packs} 組"],
        ["整體配組良率", f"{total_yield:.1f}%", "", ""]
    ]
    for r, row in enumerate(param_data, 3):
        for c, val in enumerate(row, 1):
            cell = ws1.cell(r, c, val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if c in [1,3]:
                cell.fill = accent_fill
    ws1.cell(10, 1, "📊 批次配組匯總").font = Font(size=14, bold=True)
    headers_batch = ['批號','原始總數','Stage1電芯','Stage2電芯','總配對','未配對','產出PACK','良率']
    for c, h in enumerate(headers_batch,1):
        cell = ws1.cell(11,c,h)
        cell.fill = navy_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    for r_idx, b_data in enumerate(batch_report, 12):
        for c, val in enumerate(b_data.values(),1):
            cell = ws1.cell(r_idx, c, val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if r_idx %2 ==0:
                cell.fill = zebra_fill

    if not group_summary_df.empty:
        ws_group = wb.create_sheet("PACK組電壓彙總")
        headers_group = ["GroupID","配組階段","組內最小OCV(V)","組內最大OCV(V)","組內壓差(mV)","本組電芯數"]
        for c, h in enumerate(headers_group,1):
            cell = ws_group.cell(1,c,h)
            cell.fill = navy_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
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
            for col in range(1,7):
                cell = ws_group.cell(row_num, col)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if row_num %2 ==0:
                    cell.fill = zebra_fill
            row_num +=1

    if not final_res.empty:
        ws2 = wb.create_sheet("配組明細")
        headers_detail = ["Cell_ID","Batch","OCV(V)","GroupID","配組階段","配組資訊"]
        for c,h in enumerate(headers_detail,1):
            cell = ws2.cell(1,c,h)
            cell.fill = navy_fill
            cell.font = white_font
            cell.border = thin_border
        r =2
        for _,row in final_res.iterrows():
            ws2.cell(r,1,row['Cell_ID'])
            ws2.cell(r,2,row['Batch'])
            ocv_cell = ws2.cell(r,3,row['OCV'])
            ocv_cell.number_format = "0.00000"
            ws2.cell(r,4,row['GroupID'])
            ws2.cell(r,5,row['Method'])
            ws2.cell(r,6,row['Partner_Info'])
            for col in range(1,7):
                cell = ws2.cell(r,col)
                cell.border = thin_border
                if r%2==0:
                    cell.fill = zebra_fill
            r +=1
    if not raw_df.empty:
        ws3 = wb.create_sheet("原始數據")
        ws3.cell(1,1,"Batch").fill=navy_fill
        ws3.cell(1,1).font=white_font
        ws3.cell(1,2,"OCV(V)").fill=navy_fill
        ws3.cell(1,2).font=white_font
        r=2
        for _,row in raw_df.iterrows():
            ws3.cell(r,1,row['Batch'])
            ocv_c = ws3.cell(r,2,row['OCV'])
            ocv_c.number_format="0.00000"
            for c in [1,2]:
                cell = ws3.cell(r,c)
                cell.border=thin_border
                if r%2==0:
                    cell.fill=zebra_fill
            r +=1
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max_len +4
    output.seek(0)
    wb.save(output)
    return output.getvalue()

# ---------------- UI主流程 ----------------
files = st.file_uploader("📂 請上傳電池數據 (CSV/Excel)", accept_multiple_files=True, key="file_uploader")
raw_df = load_and_fix_data(files, inject_test=inject_mock_data, current_S=S)
st.markdown("---")
run_btn = st.button("🚀 點擊執行配組分析", type="primary")

if run_btn:
    if raw_df.empty:
        st.error("無有效數據，請上傳檔案或開啟測試數據")
        st.stop()
    with st.spinner("計算配組中..."):
        total_cells = len(raw_df)
        final_res, remaining, group_summary_df = perform_pairing_engine(raw_df, S, P, target, real_pack_cell, real_threshold, enable_s2)
        matched_count = len(final_res)
        unmatched_count = len(remaining)
        total_packs = final_res['GroupID'].nunique() if not final_res.empty else 0
        total_yield = (matched_count / total_cells)*100 if total_cells>0 else 0
        threshold_mv = real_mv

        info_text = f"{S}S{P}P | 標準單Pack:{target}顆 | 每組抓取:{real_pack_cell}顆 | 基礎門檻:{base_mv}mV"
        if enable_buffer:
            info_text += f" | Buffer{buffer_cell}顆，實際容差{real_mv:.2f}mV(上限3mV)"
        st.info(info_text)
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("總電芯", f"{total_cells} 顆")
        c2.metric("配對電芯", f"{matched_count} 顆")
        c3.metric("總PACK", f"{total_packs} 組")
        c4.metric("良率", f"{total_yield:.1f}%")

        if not group_summary_df.empty:
            st.subheader("📌 各PACK電壓上下限與組內壓差一覽")
            st.dataframe(group_summary_df, use_container_width=True)
        st.markdown("---")

        chart_rows = []
        batch_report = []
        for b in raw_df['Batch'].unique():
            total = len(raw_df[raw_df['Batch']==b])
            s1 = len(final_res[(final_res['Method']=="Stage 1") & (final_res['Batch']==b)]) if not final_res.empty else 0
            s2 = len(final_res[(final_res['Method']=="Stage 2") & (final_res['Batch']==b)]) if not final_res.empty else 0
            match_total = s1 + s2
            unmatch = total - match_total
            s1_pack = final_res[(final_res['Method']=="Stage 1") & (final_res['Batch']==b)]['GroupID'].nunique() if not final_res.empty else 0
            s2_contrib = 0.0
            if not final_res.empty:
                for gid in final_res[(final_res['Method']=="Stage 2") & (final_res['Batch']==b)]['GroupID'].unique():
                    cnt = len(final_res[final_res['GroupID']==gid])
                    s2_contrib += cnt / target
            total_pack_b = s1_pack + s2_contrib
            yield_b = f"{(match_total/total)*100:.1f}" if total>0 else "0.0%"
            # 修正字典重複key語法錯誤
            batch_report.append({
                "批號(Batch)":b,
                "原始總數":total,
                "第一階配組電芯數(Stage1)":s1,
                "第二階配組電芯數(Stage2)":s2,
                "總成功配對電芯數":match_total,
                "未配對電芯數":unmatch,
                "貢獻產出組數(Packs)":round(total_pack_b,3),
                "良率":f"{yield_b}%"
            })
            chart_rows.append({"批次(Batch)":b,"電芯數量(顆)":s1,"去向結果":"第一階段配組 (Stage 1)"})
            chart_rows.append({"批次(Batch)":b,"電芯數量(顆)":s2,"去向結果":"第二階段混配 (Stage 2)"})
            chart_rows.append({"批次(Batch)":b,"電芯數量(顆)":unmatch,"去向結果":"無法配組 (Unmatched)"})
        st.subheader("📊 各批次電芯分佈堆疊圖")
        df_chart = pd.DataFrame(chart_rows)
        fig_bar = px.bar(df_chart, x="電芯數量(顆)", y="批次(Batch)", color="去向結果", orientation="h",
            color_discrete_map={"第一階段配組 (Stage 1)":"#2ecc71","第二階段混配 (Stage 2)":"#f39c12","無法配組 (Unmatched)":"#e74c3c"})
        fig_bar.update_layout(height=260, plot_bgcolor="white")
        st.plotly_chart(fig_bar, use_container_width=True)
        st.markdown("---")
        st.subheader("📊 批次匯總表格")
        st.dataframe(pd.DataFrame(batch_report), use_container_width=True)
        st.markdown("---")
        st.subheader(f"📊 全域電壓散佈圖 {S}S{P}P")
        plot_df = pd.concat([final_res, remaining.assign(GroupID="Unmatched",Method="Unmatched",Partner_Info="-")], ignore_index=True)
        plot_df['Voltage_mV'] = plot_df['OCV']*1000
        plot_df['Sequence'] = list(range(len(plot_df)))
        color_map = {}
        pal = px.colors.qualitative.Safe
        idx=0
        for g in plot_df['GroupID'].unique():
            if g=="Unmatched": color_map[g]="#e74c3c"
            elif g.startswith("S2"): color_map[g]="#f39c12"
            else:
                color_map[g] = pal[idx%len(pal)]
                idx +=1
        fig_scatter = px.scatter(plot_df, x="Sequence", y="Voltage_mV", color="GroupID",
            hover_data=["Batch","Method","GroupID","Partner_Info"],
            labels={"Sequence":"電芯序號","Voltage_mV":"電壓(mV)"}, color_discrete_map=color_map)
        fig_scatter.update_traces(marker=dict(size=6, opacity=0.85))
        fig_scatter.update_layout(height=480, plot_bgcolor="white")
        st.plotly_chart(fig_scatter, use_container_width=True)
        st.markdown("---")
        st.subheader("📥 匯出生產報告")
        excel_bytes = generate_excel_report(
            raw_df, final_res, batch_report, group_summary_df,
            S, P, target, real_pack_cell, base_mv, enable_buffer, buffer_cell, real_mv,
            total_cells, matched_count, total_packs, total_yield, real_mv
        )
        st.download_button("📥 下載稽核Excel報告", data=excel_bytes, file_name="🔋_電池配組稽核報告.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.success("Excel已納入Buffer參數紀錄！")