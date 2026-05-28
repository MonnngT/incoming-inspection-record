"""
来料检验记录系统（跳批检验逻辑）
IQC Incoming Inspection Record System with Skip-Lot Logic
部署：GitHub + Streamlit Cloud ; 数据存储：Google Sheets
"""

import streamlit as st
import pandas as pd
import json
import io
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="来料检验记录系统", page_icon="📋", layout="wide")

# ---------- 自定义样式（深蓝表头，接近Excel） ----------
st.markdown("""
<style>
.block-container { padding-top: 2rem; max-width: 1400px; }
.tbl-header {
    background-color: #1F4E79; color: white; font-weight: bold;
    text-align: center; padding: 8px 4px; border: 1px solid #FFFFFF;
    font-size: 13px; line-height: 1.3; border-radius: 2px;
}
.tbl-cell {
    text-align: center; padding: 8px 4px; border: 1px solid #D0D0D0;
    font-size: 13px; background-color: #F8FBFF; min-height: 38px; line-height: 1.4;
}
.action-normal { color: #BF8F00; font-weight: bold; }
.action-skip   { color: #1E8449; font-weight: bold; }
.action-skip-inspect { color: #2E74B5; font-weight: bold; }
div[data-testid="stHorizontalBlock"] { gap: 0.4rem; }
</style>
""", unsafe_allow_html=True)

COLUMNS = ["到货日期","供应商","零件料号","零件名称","生产日期","累计批次数",
           "执行动作","开始时间","结束时间","检验用时(分钟)","结果","检验员","记录时间"]
INSPECTORS = ["杨明","田志高","其他"]
RESULTS = ["OK","NG"]
CONSECUTIVE_PASS_TO_START = 3
SKIP_PATTERN_SKIP = 2
SKIP_PATTERN_INSPECT = 1

@st.cache_data
def load_parts_data():
    with open("parts_data.json","r",encoding="utf-8") as f:
        return json.load(f)
PARTS_DATA = load_parts_data()
SUPPLIERS = list(PARTS_DATA.keys())

@st.cache_resource
def get_gsheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    client = gspread.authorize(creds)
    sh = client.open_by_key(st.secrets["sheet"]["key"])
    ws = sh.sheet1
    if not ws.get_all_values():
        ws.append_row(COLUMNS)
    return ws

def load_history(ws):
    records = ws.get_all_records()
    return pd.DataFrame(records) if records else pd.DataFrame(columns=COLUMNS)

def append_record(ws, row_dict):
    ws.append_row([str(row_dict.get(c,"")) for c in COLUMNS])

def compute_action_and_cumulative(history_df, supplier, part_number, production_date):
    if history_df.empty:
        return 1, "正常检验"
    mask = ((history_df["供应商"].astype(str)==str(supplier)) &
            (history_df["零件料号"].astype(str)==str(part_number)) &
            (history_df["生产日期"].astype(str)==str(production_date)))
    seq = history_df[mask].copy()
    cumulative = len(seq) + 1
    if seq.empty:
        return cumulative, "正常检验"
    state="正常检验"; cp=0; sk=0
    for _, rec in seq.iterrows():
        is_pass = str(rec.get("结果","")).strip().upper()=="OK"
        if state=="正常检验":
            if is_pass:
                cp += 1
                if cp >= CONSECUTIVE_PASS_TO_START:
                    state="跳批"; sk=0
            else:
                cp = 0
        else:
            if is_pass:
                sk = (sk+1) % (SKIP_PATTERN_SKIP + SKIP_PATTERN_INSPECT)
            else:
                state="正常检验"; cp=0; sk=0
    if state=="正常检验":
        action="正常检验"
    else:
        action = "跳批（不检验尺寸）" if sk < SKIP_PATTERN_SKIP else "跳批检验（全项目）"
    return cumulative, action

def get_production_dates(history_df, supplier, part_number):
    if history_df.empty:
        return []
    mask = ((history_df["供应商"].astype(str)==str(supplier)) &
            (history_df["零件料号"].astype(str)==str(part_number)))
    dates = history_df[mask]["生产日期"].astype(str).unique().tolist()
    return sorted([d for d in dates if d and d!="nan"], reverse=True)


def make_excel(df):
    """把DataFrame导出为带格式的Excel（深蓝表头，接近Excel原表样式）"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "来料检验记录"

    header_fill = PatternFill("solid", start_color="1F4E79")
    header_font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    cell_font = Font(name="微软雅黑", size=10)
    thin = Side(border_style="thin", color="BBBBBB")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    cols = list(df.columns)
    # 表头
    for ci, col in enumerate(cols, 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center
        c.border = border
    ws.row_dimensions[1].height = 26

    # 数据
    for ri, (_, row) in enumerate(df.iterrows(), 2):
        for ci, col in enumerate(cols, 1):
            c = ws.cell(row=ri, column=ci, value=str(row[col]))
            c.font = cell_font
            c.alignment = center
            c.border = border

    # 列宽自适应
    for ci, col in enumerate(cols, 1):
        max_len = max([len(str(col))] + [len(str(v)) for v in df[col].astype(str)])
        ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len * 1.6, 10), 40)

    # 冻结表头
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ---------- 界面 ----------
st.title("📋 来料检验记录系统")
st.caption("IQC Incoming Inspection Record · 跳批检验逻辑 · 数据存储于 Google Sheets")

try:
    ws = get_gsheet()
    history_df = load_history(ws)
    gsheet_ok = True
except Exception as e:
    st.error(f"⚠️ Google Sheets 连接失败：{e}")
    st.info("请检查 Streamlit secrets 中的 gcp_service_account 和 sheet.key 配置。")
    history_df = pd.DataFrame(columns=COLUMNS)
    gsheet_ok = False

tab1, tab2 = st.tabs(["✍️ 新增检验记录", "📜 历史记录"])

with tab1:
    st.markdown("##### 录入数据")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        arrival_date = st.date_input("到货日期", value=date.today())
    with c2:
        supplier = st.selectbox("供应商", SUPPLIERS, index=0)
    with c3:
        parts_list = PARTS_DATA.get(supplier, [])
        part_options, part_display_map = [], {}
        for p in parts_list:
            pn = p["part_number"]; pname = p.get("part_name","")
            disp = f"{pn} | {pname}" if pname else pn
            part_options.append(disp); part_display_map[disp] = p
        if part_options:
            part_disp = st.selectbox("零件料号", part_options, index=0)
            selected_part = part_display_map[part_disp]
            part_number = selected_part["part_number"]
            part_name = selected_part.get("part_name","")
        else:
            st.warning("该供应商暂无料号"); part_number=""; part_name=""
    with c4:
        hist_prod_dates = get_production_dates(history_df, supplier, part_number)
        if hist_prod_dates:
            mode = st.radio("生产日期来源", ["新日期","历史"], horizontal=True, label_visibility="collapsed")
        else:
            mode = "新日期"
        if mode=="历史" and hist_prod_dates:
            production_date = st.selectbox("生产日期", hist_prod_dates)
        else:
            production_date = str(st.date_input("生产日期", value=date.today(), key="prod_cal"))

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        start_time = st.time_input("开始时间", value=datetime.now().time())
    with c6:
        end_time = st.time_input("结束时间", value=datetime.now().time())
    with c7:
        result = st.selectbox("结果", RESULTS)
    with c8:
        inspector_sel = st.selectbox("检验员", INSPECTORS)
        if inspector_sel=="其他":
            inspector = st.text_input("输入姓名", "", label_visibility="collapsed", placeholder="请输入检验员姓名")
        else:
            inspector = inspector_sel

    cumulative, action = compute_action_and_cumulative(history_df, supplier, part_number, production_date)
    today = date.today()
    dt_start = datetime.combine(today, start_time)
    dt_end = datetime.combine(today, end_time)
    diff_min = (dt_end - dt_start).total_seconds()/60
    if diff_min < 0: diff_min += 24*60

    if "跳批（不检验尺寸）" in action:
        action_html = f'<span class="action-skip">{action}</span>'
    elif "跳批检验" in action:
        action_html = f'<span class="action-skip-inspect">{action}</span>'
    else:
        action_html = f'<span class="action-normal">{action}</span>'

    st.divider()
    st.markdown("##### 当前记录预览")

    header_cols = ["到货日期","供应商","零件料号","生产日期","累计批次数",
                   "执行动作","开始时间","结束时间","检验用时(分钟)","结果","检验员"]
    header_en = ["Arrival Date","Supplier","Part number","Production Date",
                 "Cumulative Lots","Action","Start","End","Time(min)","Result","Inspector"]
    weights = [1.1,0.9,1.5,1.1,0.8,1.4,0.7,0.7,0.9,0.7,0.9]

    hcols = st.columns(weights)
    for i, hc in enumerate(hcols):
        hc.markdown(f'<div class="tbl-header">{header_cols[i]}<br>'
                    f'<span style="font-size:10px;font-weight:normal;">{header_en[i]}</span></div>',
                    unsafe_allow_html=True)

    pn_disp = part_number if part_number else "—"
    values = [str(arrival_date), supplier, pn_disp, str(production_date), str(cumulative),
              action_html, start_time.strftime("%H:%M"), end_time.strftime("%H:%M"),
              f"{diff_min:.0f}", result, inspector if inspector else "—"]
    dcols = st.columns(weights)
    for i, dc in enumerate(dcols):
        dc.markdown(f'<div class="tbl-cell">{values[i]}</div>', unsafe_allow_html=True)

    if "跳批（不检验尺寸）" in action:
        st.success("✅ 本批跳过尺寸检验，只做外观 + 包装数量")
    elif "跳批检验" in action:
        st.info("🔍 跳批序列的检验批，需做全项目检验（外观+尺寸+包装数量）")
    else:
        st.warning("📋 正常检验：外观 + 尺寸 + 包装数量全检")

    st.divider()
    if st.button("💾 保存记录到 Google Sheets", type="primary", use_container_width=True):
        if not gsheet_ok:
            st.error("Google Sheets 未连接，无法保存。")
        elif not part_number:
            st.error("请选择有效料号。")
        elif inspector_sel=="其他" and not inspector.strip():
            st.error("请填写检验员姓名。")
        else:
            row = {"到货日期":str(arrival_date),"供应商":supplier,"零件料号":part_number,
                   "零件名称":part_name,"生产日期":str(production_date),"累计批次数":cumulative,
                   "执行动作":action,"开始时间":start_time.strftime("%H:%M"),
                   "结束时间":end_time.strftime("%H:%M"),"检验用时(分钟)":f"{diff_min:.0f}",
                   "结果":result,"检验员":inspector,
                   "记录时间":datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            try:
                append_record(ws, row)
                st.success("✅ 记录已保存！")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"保存失败：{e}")

with tab2:
    st.subheader("历史检验记录")
    if history_df.empty:
        st.info("暂无历史记录。")
    else:
        colF1, colF2, colF3 = st.columns(3)
        with colF1:
            f_supplier = st.multiselect("按供应商筛选", SUPPLIERS)
        with colF2:
            all_parts = history_df["零件料号"].astype(str).unique().tolist()
            f_part = st.multiselect("按料号筛选", all_parts)
        with colF3:
            f_result = st.multiselect("按结果筛选", RESULTS)
        view = history_df.copy()
        if f_supplier: view = view[view["供应商"].isin(f_supplier)]
        if f_part: view = view[view["零件料号"].astype(str).isin(f_part)]
        if f_result: view = view[view["结果"].isin(f_result)]
        show_cols = ["到货日期","供应商","零件料号","生产日期","累计批次数",
                     "执行动作","开始时间","结束时间","检验用时(分钟)","结果","检验员"]
        show_cols = [c for c in show_cols if c in view.columns]
        st.dataframe(view[show_cols], use_container_width=True, height=400)
        st.divider()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("总记录数", len(view))
        if "执行动作" in view.columns and len(view):
            skip_cnt = view["执行动作"].astype(str).str.contains("跳批（不检验尺寸）").sum()
            c2.metric("跳过尺寸检验批数", int(skip_cnt))
        if "结果" in view.columns and len(view):
            c3.metric("NG批数", int((view["结果"]=="NG").sum()))
        if "检验用时(分钟)" in view.columns and len(view):
            try:
                total_min = pd.to_numeric(view["检验用时(分钟)"], errors="coerce").sum()
                c4.metric("累计检验用时(小时)", f"{total_min/60:.1f}")
            except Exception:
                pass
        exp1, exp2 = st.columns(2)
        with exp1:
            csv = view.to_csv(index=False).encode("utf-8-sig")
            st.download_button("📥 导出 CSV", csv, "inspection_records.csv",
                               "text/csv", use_container_width=True)
        with exp2:
            try:
                xlsx_bytes = make_excel(view[show_cols] if show_cols else view)
                ts = datetime.now().strftime("%Y%m%d_%H%M")
                st.download_button(
                    "📊 导出 Excel",
                    xlsx_bytes,
                    f"来料检验记录_{ts}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"生成Excel失败：{e}")
