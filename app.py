"""
来料检验记录系统（跳批检验逻辑）
IQC Incoming Inspection Record System with Skip-Lot Logic

部署：GitHub + Streamlit Cloud
数据存储：Google Sheets
"""

import streamlit as st
import pandas as pd
import json
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(
    page_title="来料检验记录系统",
    page_icon="📋",
    layout="wide",
)

# ============================================================
# 常量
# ============================================================
# Google Sheet 表头（与你的图片一致 + 拆分检验用时）
COLUMNS = [
    "到货日期",          # Arrival Date
    "供应商",            # Supplier
    "零件料号",          # Part number
    "零件名称",          # Part name（自动带出，便于核对）
    "生产日期",          # Production Date
    "累计批次数",        # Cumulative Lots
    "执行动作",          # Action（自动判定：正常检验 / 跳批（不检验尺寸））
    "开始时间",          # Inspection Start Time
    "结束时间",          # Inspection End Time
    "检验用时(分钟)",    # Inspection Time（结束-开始，自动计算）
    "结果",              # Result (Pass/Fail -> OK/NG)
    "检验员",            # Inspector
    "记录时间",          # 系统记录时间戳
]

INSPECTORS = ["杨明", "田志高", "其他"]
RESULTS = ["OK", "NG"]

# 跳批规则参数
CONSECUTIVE_PASS_TO_START = 3   # 连续N批合格后启动跳批
SKIP_PATTERN_SKIP = 2          # 跳过批数
SKIP_PATTERN_INSPECT = 1       # 检验批数


# ============================================================
# 加载料号数据
# ============================================================
@st.cache_data
def load_parts_data():
    with open("parts_data.json", "r", encoding="utf-8") as f:
        return json.load(f)

PARTS_DATA = load_parts_data()
SUPPLIERS = list(PARTS_DATA.keys())


# ============================================================
# Google Sheets 连接
# ============================================================
@st.cache_resource
def get_gsheet():
    """连接 Google Sheet，返回 worksheet 对象"""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    client = gspread.authorize(creds)
    # 用 secrets 中的 sheet_url 或 sheet_key 打开
    sheet_key = st.secrets["sheet"]["key"]
    sh = client.open_by_key(sheet_key)
    ws = sh.sheet1
    # 确保有表头
    existing = ws.get_all_values()
    if not existing:
        ws.append_row(COLUMNS)
    return ws


def load_history(ws):
    """读取历史记录为 DataFrame"""
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=COLUMNS)
    return pd.DataFrame(records)


def append_record(ws, row_dict):
    """追加一条记录"""
    row = [str(row_dict.get(c, "")) for c in COLUMNS]
    ws.append_row(row)


# ============================================================
# 跳批逻辑核心
# ============================================================
def compute_action_and_cumulative(history_df, supplier, part_number, production_date):
    """
    根据历史记录，计算当前这批（同一料号+同一生产日期）的:
      - 累计批次数
      - 执行动作（正常检验 / 跳批（不检验尺寸））

    规则：
      - 跳批序列以 "料号 + 生产日期" 为 key（同一生产日期才能跳批累计）
      - 连续3批合格后启动跳批，进入"跳2检1"循环
      - 任一不合格 -> 退回正常检验，连续合格计数清零
    """
    if history_df.empty:
        # 第一批
        return 1, "正常检验"

    # 同一料号 + 同一生产日期 的历史记录（按记录时间排序）
    prod_date_str = str(production_date)
    mask = (
        (history_df["供应商"].astype(str) == str(supplier)) &
        (history_df["零件料号"].astype(str) == str(part_number)) &
        (history_df["生产日期"].astype(str) == prod_date_str)
    )
    seq = history_df[mask].copy()

    cumulative = len(seq) + 1  # 当前是第几批

    if seq.empty:
        # 此料号+此生产日期的第一批
        return cumulative, "正常检验"

    # 重建状态机：遍历历史，跟踪状态
    state = "正常检验"          # 当前检验状态
    consecutive_pass = 0        # 连续合格计数（正常检验阶段用）
    skip_counter = 0            # 跳批循环内的计数（0,1=跳过, 2=检验）

    # 按记录顺序处理每一条历史
    for _, rec in seq.iterrows():
        result = str(rec.get("结果", "")).strip().upper()
        is_pass = result == "OK"

        if state == "正常检验":
            if is_pass:
                consecutive_pass += 1
                if consecutive_pass >= CONSECUTIVE_PASS_TO_START:
                    # 达到连续3批合格 -> 下一批进入跳批
                    state = "跳批"
                    skip_counter = 0
            else:
                consecutive_pass = 0  # 不合格清零
        else:  # 跳批状态
            if is_pass:
                # 推进跳批循环
                skip_counter = (skip_counter + 1) % (SKIP_PATTERN_SKIP + SKIP_PATTERN_INSPECT)
            else:
                # 跳批中不合格 -> 退回正常检验，清零
                state = "正常检验"
                consecutive_pass = 0
                skip_counter = 0

    # 现在决定"当前这批"的动作
    if state == "正常检验":
        action = "正常检验"
    else:
        # 跳批状态：skip_counter 0 或 1 -> 跳过（不检验尺寸）；2 -> 检验
        if skip_counter < SKIP_PATTERN_SKIP:
            action = "跳批（不检验尺寸）"
        else:
            action = "跳批检验（全项目）"

    return cumulative, action


def get_production_dates(history_df, supplier, part_number):
    """返回该料号历史出现过的生产日期列表（去重，倒序）"""
    if history_df.empty:
        return []
    mask = (
        (history_df["供应商"].astype(str) == str(supplier)) &
        (history_df["零件料号"].astype(str) == str(part_number))
    )
    dates = history_df[mask]["生产日期"].astype(str).unique().tolist()
    return sorted([d for d in dates if d and d != "nan"], reverse=True)


# ============================================================
# 界面
# ============================================================
st.title("📋 来料检验记录系统")
st.caption("IQC Incoming Inspection Record System · 跳批检验逻辑 · 数据存储于 Google Sheets")

# 连接 Google Sheet
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

# ------------------------------------------------------------
# Tab 1: 新增记录
# ------------------------------------------------------------
with tab1:
    st.subheader("新增一条来料检验记录")

    colA, colB = st.columns(2)

    with colA:
        # 到货日期 - 日历下拉
        arrival_date = st.date_input("到货日期 / Arrival Date", value=date.today())

        # 供应商 - 下拉
        supplier = st.selectbox("供应商 / Supplier", SUPPLIERS, index=0)

        # 零件料号 - 根据供应商联动下拉
        parts_list = PARTS_DATA.get(supplier, [])
        part_options = []
        part_display_map = {}
        for p in parts_list:
            pn = p["part_number"]
            pname = p.get("part_name", "")
            sub = p.get("sub_supplier", "")
            # 压铸件显示子供应商，便于区分
            if sub:
                disp = f"{pn} | {pname} [{sub}]"
            else:
                disp = f"{pn} | {pname}"
            part_options.append(disp)
            part_display_map[disp] = p

        if part_options:
            part_disp = st.selectbox("零件料号 / Part number", part_options, index=0)
            selected_part = part_display_map[part_disp]
            part_number = selected_part["part_number"]
            part_name = selected_part.get("part_name", "")
        else:
            st.warning("该供应商暂无料号数据")
            part_number = ""
            part_name = ""

    with colB:
        # 生产日期 - 下拉（历史日期）+ 可新增
        hist_prod_dates = get_production_dates(history_df, supplier, part_number)
        prod_date_mode = st.radio(
            "生产日期 / Production Date",
            ["选择新日期", "从历史选择"] if hist_prod_dates else ["选择新日期"],
            horizontal=True,
        )
        if prod_date_mode == "从历史选择" and hist_prod_dates:
            production_date = st.selectbox("历史生产日期", hist_prod_dates)
        else:
            production_date = str(st.date_input(
                "生产日期(日历选择)", value=date.today(), key="prod_date_cal"
            ))

        # 检验员 - 下拉，"其他"则手动输入
        inspector_sel = st.selectbox("检验员 / Inspector", INSPECTORS)
        if inspector_sel == "其他":
            inspector = st.text_input("请输入检验员姓名", "")
        else:
            inspector = inspector_sel

        # 结果 - 下拉
        result = st.selectbox("结果 / Result", RESULTS)

    st.divider()

    # 自动计算：累计批次 + 执行动作
    cumulative, action = compute_action_and_cumulative(
        history_df, supplier, part_number, production_date
    )

    col1, col2, col3 = st.columns(3)
    col1.metric("累计批次数 / Cumulative Lots", cumulative)
    col2.metric("执行动作 / Action", action)
    # 动作配色提示
    if "跳批（不检验尺寸）" in action:
        col3.success("✅ 本批跳过尺寸检验，只做外观+包装数量")
    elif "跳批检验" in action:
        col3.info("🔍 跳批序列的检验批，需做全项目检验")
    else:
        col3.warning("📋 正常检验：外观+尺寸+包装数量全检")

    st.divider()

    # 检验用时：开始时间 / 结束时间
    st.markdown("##### ⏱️ 检验用时")
    colT1, colT2, colT3 = st.columns(3)
    with colT1:
        start_time = st.time_input("开始时间 / Start Time", value=datetime.now().time())
    with colT2:
        end_time = st.time_input("结束时间 / End Time", value=datetime.now().time())
    with colT3:
        # 计算用时（分钟）
        today = date.today()
        dt_start = datetime.combine(today, start_time)
        dt_end = datetime.combine(today, end_time)
        diff_min = (dt_end - dt_start).total_seconds() / 60
        if diff_min < 0:
            diff_min += 24 * 60  # 跨天保护
        st.metric("检验用时 / Time (分钟)", f"{diff_min:.0f}")

    st.divider()

    # 提交
    if st.button("💾 保存记录到 Google Sheets", type="primary", use_container_width=True):
        if not gsheet_ok:
            st.error("Google Sheets 未连接，无法保存。")
        elif not part_number:
            st.error("请选择有效料号。")
        elif inspector_sel == "其他" and not inspector.strip():
            st.error("请填写检验员姓名。")
        else:
            row = {
                "到货日期": str(arrival_date),
                "供应商": supplier,
                "零件料号": part_number,
                "零件名称": part_name,
                "生产日期": str(production_date),
                "累计批次数": cumulative,
                "执行动作": action,
                "开始时间": start_time.strftime("%H:%M"),
                "结束时间": end_time.strftime("%H:%M"),
                "检验用时(分钟)": f"{diff_min:.0f}",
                "结果": result,
                "检验员": inspector,
                "记录时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            try:
                append_record(ws, row)
                st.success("✅ 记录已保存！页面将刷新以更新历史。")
                st.cache_data.clear()  # 清缓存以重新读历史
                st.rerun()
            except Exception as e:
                st.error(f"保存失败：{e}")


# ------------------------------------------------------------
# Tab 2: 历史记录
# ------------------------------------------------------------
with tab2:
    st.subheader("历史检验记录")
    if history_df.empty:
        st.info("暂无历史记录。")
    else:
        # 筛选
        colF1, colF2, colF3 = st.columns(3)
        with colF1:
            f_supplier = st.multiselect("按供应商筛选", SUPPLIERS)
        with colF2:
            all_parts = history_df["零件料号"].astype(str).unique().tolist()
            f_part = st.multiselect("按料号筛选", all_parts)
        with colF3:
            f_result = st.multiselect("按结果筛选", RESULTS)

        view = history_df.copy()
        if f_supplier:
            view = view[view["供应商"].isin(f_supplier)]
        if f_part:
            view = view[view["零件料号"].astype(str).isin(f_part)]
        if f_result:
            view = view[view["结果"].isin(f_result)]

        st.dataframe(view, use_container_width=True, height=400)

        # 简单统计
        st.divider()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("总记录数", len(view))
        if "执行动作" in view.columns and len(view):
            skip_cnt = view["执行动作"].astype(str).str.contains("跳批（不检验尺寸）").sum()
            c2.metric("跳过尺寸检验批数", int(skip_cnt))
        if "结果" in view.columns and len(view):
            ng_cnt = (view["结果"] == "NG").sum()
            c3.metric("NG批数", int(ng_cnt))
        if "检验用时(分钟)" in view.columns and len(view):
            try:
                total_min = pd.to_numeric(view["检验用时(分钟)"], errors="coerce").sum()
                c4.metric("累计检验用时(小时)", f"{total_min/60:.1f}")
            except Exception:
                pass

        # 导出CSV
        csv = view.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "📥 导出当前视图为 CSV",
            csv,
            "inspection_records.csv",
            "text/csv",
        )
