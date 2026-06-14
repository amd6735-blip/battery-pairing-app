import streamlit as st
import pandas as pd
from itertools import combinations
import plotly.express as px
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="電池配組與稽核系統", layout="wide")

st.title("🔋 Dynapack Non-matchable Cell regrouping System")
st.markdown("本版本已全面整合**電芯去向堆疊圖**、**全域電壓分佈散佈圖**，並支援 **Excel 生產報告一鍵導出功能**。")

# --- 側邊欄參數設定 ---
with st.sidebar:
    st.header("⚙️ 分析參數 (Configuration)")
    S = st.number_input("串聯數 (S)", 1, 24, 16)
    P = st.number_input("並聯數 (P)", 1, 20, 6)
    target = S * P  # Pack Size
    threshold = st.select_slider("Max Delta V (壓差門檻 mV)", options=[1, 2, 3, 4, 5, 10], value=3) / 1000
    enable_s2 = st.checkbox("Sn Ratio (啟用第二階段動態跨批次)", value=True)
    
    st.markdown("---")
    st.header("🔬 演算法驗證模式")
    inject_mock_data = st.checkbox("🔥 強制注入測試電芯 (用於驗證第二階段)", value=True)
    st.caption(f"開啟後，系統會自動在資料中注入符合 {S}S 架構的重疊電芯，確保 100% 觸發 Stage 2 跨批次配組以供檢視。")

# --- 防禦性資料讀取引擎 ---
def load_and_fix_data(uploaded_files, inject_test=False, current_S=16):
    all_data = []
    if uploaded_files:
        for f in uploaded_files:
            try:
                if f.name.endswith('.xlsx') or f.name.endswith('.xls'):
                    df = pd.read_excel(f)
                else:
                    df = pd.read_csv(f, encoding_errors='replace', engine='python')
                
                df = df.iloc[:, [0, 1]]
                df.columns = ['Batch', 'OCV']
                df['OCV'] = pd.to_numeric(df['OCV'], errors='coerce')
                all_data.append(df.dropna(subset=['Batch', 'OCV']))
            except Exception as e:
                st.error(f"檔案 {f.name} 讀取失敗: {e}")
                
    main_df = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
    
    if inject_test:
        if main_df.empty:
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
            
            mock_cells = pd.DataFrame({
                'Batch': [batch_b] * current_S,
                'OCV': [base_v - 0.0005] * current_S
            })
            main_df = pd.concat([main_df, mock_cells], ignore_index=True)
            
    return main_df

# --- 動態 Sn 雙階段配組核心引擎 ---
def perform_pairing_engine(df, S, P, target, threshold, enable_s2):
    results = []
    remaining = df.copy().reset_index(drop=True)
    remaining['Cell_ID'] = remaining.index

    # 【Stage 1】批次內配組
    for batch in remaining['Batch'].unique():
        while True:
            b_df = remaining[remaining['Batch'] == batch].sort_values('OCV')
            if len(b_df) < target:
                break
            
            subset = b_df.iloc[:target]
            if (subset['OCV'].max() - subset['OCV'].min()) <= threshold:
                group = subset.copy()
                group['GroupID'] = f"P{len(results)+1}-{batch}"
                group['Method'] = "Stage 1"
                group['Partner_Info'] = "單一批次內配組"
                results.append(group)
                remaining = remaining.drop(subset.index)
            else:
                remaining = remaining.drop(b_df.index[0])

    # 【Stage 2】動態 Sn 跨批次混搭配組
    if enable_s2 and len(remaining) >= target:
        batches = remaining['Batch'].unique()
        if len(batches) >= 2:
            for b1, b2 in combinations(batches, 2):
                while True:
                    d1 = remaining[remaining['Batch'] == b1].sort_values('OCV')
                    d2 = remaining[remaining['Batch'] == b2].sort_values('OCV')
                    
                    pair_found_in_this_loop = False
                    
                    for n in range(1, P):
                        n1, n2 = S * n, target - (S * n)
                        if len(d1) >= n1 and len(d2) >= n2:
                            subset = pd.concat([d1.iloc[:n1], d2.iloc[:n2]])
                            if (subset['OCV'].max() - subset['OCV'].min()) <= threshold:
                                group = subset.copy()
                                group['GroupID'] = f"S2-{b1}x{b2}-Sn{n}"
                                group['Method'] = "Stage 2"
                                group['Partner_Info'] = f"{b1}({n1}顆) + {b2}({n2}顆)"
                                results.append(group)
                                remaining = remaining.drop(subset.index)
                                pair_found_in_this_loop = True
                                break 
                    
                    if not pair_found_in_this_loop:
                        break

    final_res = pd.concat(results, ignore_index=True) if results else pd.DataFrame()
    return final_res, remaining

# --- 高級 Excel 報告生成器 ---
def generate_excel_report(raw_df, final_res, batch_report, S, P, total_cells, matched_count, total_packs, total_yield):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    font_family = "Segoe UI"
    navy_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F6F9", end_color="F2F6F9", fill_type="solid")
    accent_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    
    # 工作表 1：生產總覽
    ws1 = wb.active
    ws1.title = "生產稽核總覽"
    ws1.views.sheetView[0].showGridLines = True
    
    ws1["A1"] = "🔋 鋰電池精密配組稽核生產報告"
    ws1["A1"].font = Font(name=font_family, size=16, bold=True, color="1F4E78")
    ws1["A2"] = f"報告運行架構：{S}S {P}P (每組 {S*P} 顆)"
    ws1["A2"].font = Font(name=font_family, size=10, italic=True, color="595959")
    
    metrics = [
        ("全案總處理電芯數", total_cells, "顆"),
        ("成功配對電芯總數", matched_count, "顆"),
        ("最終成功產出組數 (Packs)", total_packs, "組"),
        ("整體配組綜合良率", f"{total_yield:.2f}%", "")
    ]
    
    for idx, (label, val, unit) in enumerate(metrics, start=4):
        ws1.cell(row=idx, column=1, value=label).font = Font(name=font_family)
        c_val = ws1.cell(row=idx, column=2, value=val)
        c_val.font = Font(name=font_family, bold=True)
        ws1.cell(row=idx, column=3, value=unit).font = Font(name=font_family, size=9, color="595959")
    
    start_r = 10
    ws1.cell(row=start_r-1, column=1, value="📋 各生產批次精確消耗拆分表").font = Font(name=font_family, size=11, bold=True)
    
    headers_b = list(batch_report[0].keys())
    for col_idx, text in enumerate(headers_b, start=1):
        cell = ws1.cell(row=start_r, column=col_idx, value=text)
        cell.font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        cell.fill = navy_fill
        cell.alignment = Alignment(horizontal="center")
    
    current_r = start_r + 1
    for r_idx, row_data in enumerate(batch_report):
        for col_idx, value in enumerate(row_data.values(), start=1):
            cell = ws1.cell(row=current_r, column=col_idx, value=value)
            cell.font = Font(name=font_family, size=10)
            cell.border = thin_border
            if r_idx % 2 == 1: cell.fill = zebra_fill
        current_r += 1
        
    # 工作表 2：詳細清單
    ws2 = wb.create_sheet(title="封包配組明細清單")
    ws2.views.sheetView[0].showGridLines = True
    if not final_res.empty:
        export_columns = ['Cell_ID', 'Batch', 'OCV', 'GroupID', 'Method', 'Partner_Info']
        df_export = final_res[export_columns].copy()
        df_export['OCV_mV'] = df_export['OCV'] * 1000
        df_export = df_export[['Cell_ID', 'Batch', 'OCV', 'OCV_mV', 'GroupID', 'Method', 'Partner_Info']]
        df_export.columns = ['電芯原始ID', '所屬批次', '電壓 (V)', '電壓 (mV)', '分配組包ID (GroupID)', '配組階段', '動態混配說明']
        
        for col_idx, text in enumerate(df_export.columns, start=1):
            cell = ws2.cell(row=1, column=col_idx, value=text)
            cell.font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
            cell.fill = navy_fill
            
        for r_idx, row_values in enumerate(df_export.values, start=2):
            for col_idx, val in enumerate(row_values, start=1):
                cell = ws2.cell(row=r_idx, column=col_idx, value=val)
                cell.font = Font(name=font_family, size=10)
                cell.border = thin_border
                if r_idx % 2 == 1: cell.fill = zebra_fill

    for ws in [ws1, ws2]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    wb.save(output)
    return output.getvalue()

# --- UI 主流程區塊 ---
files = st.file_uploader("📂 請上傳電池數據 (CSV/Excel)", accept_multiple_files=True)

if files or inject_mock_data:
    if st.button("🚀 執行動態參數稽核與計算"):
        with st.spinner("正在進行精密稽核計算..."):
            raw_df = load_and_fix_data(files, inject_test=inject_mock_data, current_S=S)
            
            total_cells = len(raw_df)
            final_res, remaining_df = perform_pairing_engine(raw_df, S, P, target, threshold, enable_s2)
            
            matched_count = len(final_res) if not final_res.empty else 0
            unmatched_count = len(remaining_df)
            total_packs = final_res['GroupID'].nunique() if not final_res.empty else 0
            total_yield = (matched_count / total_cells) * 100 if total_cells > 0 else 0
            
            # 頂部儀表板
            st.info(f"📋 當前運行架構：{S}S {P}P | 單個電池組(Pack)所需電芯數：{target} 顆")
            m_col1, m_col2, m_col3, m_col4 = st.columns(4)
            m_col1.metric("總處理電芯數", f"{total_cells} 顆")
            m_col2.metric("已成功配對電芯", f"{matched_count} 顆")
            m_col3.metric("總配組成功數 (Packs)", f"{total_packs} 組")
            m_col4.metric("整體配組良率", f"{total_yield:.1f}%")
            st.markdown("---")
            
            # 建立圖表與統計數據資料
            chart_rows = []
            batch_report = []
            
            for b in raw_df['Batch'].unique():
                total = len(raw_df[raw_df['Batch'] == b])
                s1_cells = len(final_res[(final_res['Batch'] == b) & (final_res['Method'] == 'Stage 1')]) if not final_res.empty else 0
                s2_cells = len(final_res[(final_res['Batch'] == b) & (final_res['Method'] == 'Stage 2')]) if not final_res.empty else 0
                matched = s1_cells + s2_cells
                unmatched = total - matched
                
                s1_packs = final_res[(final_res['Batch'] == b) & (final_res['Method'] == 'Stage 1')]['GroupID'].nunique() if not final_res.empty else 0
                s2_packs_contribution = 0.0
                if not final_res.empty:
                    s2_groups = final_res[(final_res['Batch'] == b) & (final_res['Method'] == 'Stage 2')]['GroupID'].unique()
                    for g_id in s2_groups:
                        g_data = final_res[final_res['GroupID'] == g_id]
                        b_contrib = len(g_data[g_data['Batch'] == b])
                        s2_packs_contribution += (b_contrib / target)
                
                total_packs_b = s1_packs + s2_packs_contribution
                yield_rate = f"{(matched / total) * 100:.1f}%" if total > 0 else "0.0%"
                
                batch_report.append({
                    '批號(Batch)': b, '原始總數': total, '第一階配組電芯數(Stage1)': s1_cells, 
                    '第二階配組電芯數(Stage2)': s2_cells, '總成功配對電芯數': matched, '未配對電芯數': unmatched, 
                    '貢獻產出組數(Packs)': round(total_packs_b, 3), '良率': yield_rate
                })
                
                chart_rows.append({'批次(Batch)': b, '電芯數量(顆)': s1_cells, '去向結果': '第一階段配組 (Stage 1)'})
                chart_rows.append({'批次(Batch)': b, '電芯數量(顆)': s2_cells, '去向結果': '第二階段混配 (Stage 2)'})
                chart_rows.append({'批次(Batch)': b, '電芯數量(顆)': unmatched, '去向結果': '無法配組殘料 (Unmatched)'})
            
            # =========================================================
            # 📊 圖表一：各批次電芯去向堆疊圖
            # =========================================================
            st.subheader("📊 各批次電芯配組結果與去向分佈報告 (Visual Distribution Chart)")
            df_chart = pd.DataFrame(chart_rows)
            fig_bar = px.bar(df_chart, x="電芯數量(顆)", y="批次(Batch)", color="去向結果", orientation='h',
                             color_discrete_map={'第一階段配組 (Stage 1)': '#2ecc71','第二階段混配 (Stage 2)': '#f39c12','無法配組殘料 (Unmatched)': '#e74c3c'})
            fig_bar.update_layout(barmode='stack', height=260, plot_bgcolor='white', margin=dict(t=10, b=10))
            st.plotly_chart(fig_bar, use_container_width=True)
            st.markdown("---")

            # 數據表格展示
            st.subheader("📊 批次數據匯總明細表格 (Batch Summary)")
            batch_df = pd.DataFrame(batch_report)
            st.dataframe(batch_df, use_container_width=True)
            st.markdown("---")

            # =========================================================
            # 🎯 圖表二：全域電壓分佈視覺化散佈圖（把漏掉的圖加回來！）
            # =========================================================
            st.subheader(f"📊 全域電壓分佈視覺化散佈圖 ({S}S {P}P 架構)")
            
            # 整理散佈圖資料
            if not remaining_df.empty:
                remaining_df['GroupID'] = 'Unmatched'
                remaining_df['Method'] = 'Unmatched'
                remaining_df['Partner_Info'] = '-'
            plot_df = pd.concat([final_res, remaining_df], ignore_index=True).sort_values('Cell_ID')
            plot_df['Voltage_mV'] = plot_df['OCV'] * 1000
            plot_df['Sequence'] = range(1, len(plot_df) + 1)
            
            # 定義散佈圖顏色分配 (殘料紅色、Stage 2 橘黃色、其餘自動分配)
            unique_groups = plot_df['GroupID'].unique()
            color_map = {}
            colors_palette = px.colors.qualitative.Safe
            color_idx = 0
            for g in unique_groups:
                if g == 'Unmatched': color_map[g] = '#e74c3c'
                elif 'S2' in str(g): color_map[g] = '#f39c12'
                else:
                    color_map[g] = colors_palette[color_idx % len(colors_palette)]
                    color_idx += 1

            fig_scatter = px.scatter(plot_df, x='Sequence', y='Voltage_mV', color='GroupID',
                             title="黃橘色點即為 Stage 2 跨批次精準混配組，點擊右側圖例可過列單一組別",
                             labels={'Sequence': 'Cell Sequence (電芯排序)', 'Voltage_mV': 'Voltage (電壓 mV)'},
                             color_discrete_map=color_map,
                             hover_data=['Batch', 'Method', 'GroupID', 'Partner_Info'])
            fig_scatter.update_traces(marker=dict(size=6, opacity=0.85, line=dict(width=0.5, color='White')))
            fig_scatter.update_layout(height=480, plot_bgcolor='white', yaxis=dict(gridcolor="#E5E7EB"), xaxis=dict(gridcolor="#E5E7EB"))
            st.plotly_chart(fig_scatter, use_container_width=True)
            
            # ==========================================
            # 📥 匯出 Excel 報告下載區塊
            # ==========================================
            st.markdown("---")
            st.subheader("📥 匯出生產報告 (Export Report)")
            excel_data = generate_excel_report(
                raw_df, final_res, batch_report, S, P, 
                total_cells, matched_count, total_packs, total_yield
            )
            st.download_button(
                label="📥 點擊下載官方工業級 Excel 生產稽核報告",
                data=excel_data,
                file_name="🔋_電池配組與精密稽核報告.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Excel 報表與所有圖表已同步編譯完畢！")
