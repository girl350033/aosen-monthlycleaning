import streamlit as st
import pandas as pd
import datetime
import calendar
import holidays
import io
import html

from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    Alignment,
    Border,
    Side,
    PatternFill,
)
from openpyxl.utils import get_column_letter


# =========================================================
# 頁面基本設定
# =========================================================
st.set_page_config(
    page_title="澳森托嬰中心 月清潔輪值與表單管理系統",
    page_icon="🏫",
    layout="wide",
)


# =========================================================
# 全站 CSS
# 1. 減少頁面留白
# 2. 排班五欄間距縮小
# 3. Selectbox 高度加高、文字可顯示約兩行
# 4. Expander 更緊湊
# =========================================================
st.markdown(
    """
<style>

/* ---------------------------------------------------------
   整體頁面
--------------------------------------------------------- */

.block-container {
    padding-top: 1rem !important;
    padding-bottom: 1rem !important;
    padding-left: 1.2rem !important;
    padding-right: 1.2rem !important;
    max-width: 100% !important;
}

/* 標題間距 */
h1 {
    margin-top: 0 !important;
    margin-bottom: 0.7rem !important;
}

h2, h3, h4 {
    margin-top: 0.4rem !important;
    margin-bottom: 0.45rem !important;
}

/* divider */
hr {
    margin-top: 0.7rem !important;
    margin-bottom: 0.7rem !important;
}


/* ---------------------------------------------------------
   Streamlit 橫向欄位
--------------------------------------------------------- */

div[data-testid="stHorizontalBlock"] {
    gap: 0.7rem !important;
}

/* 一般垂直元件間距 */
div[data-testid="stVerticalBlock"] {
    gap: 0.55rem !important;
}


/* ---------------------------------------------------------
   Expander
--------------------------------------------------------- */

div[data-testid="stExpander"] {
    margin-bottom: 0.7rem !important;
}

div[data-testid="stExpander"] details {
    border-radius: 10px !important;
}

div[data-testid="stExpander"] summary {
    padding-top: 0.45rem !important;
    padding-bottom: 0.45rem !important;
    font-weight: 600 !important;
}


/* ---------------------------------------------------------
   Selectbox
--------------------------------------------------------- */

/* 整個 selectbox 高度 */
div[data-baseweb="select"] > div {
    min-height: 58px !important;
    height: auto !important;
    border-radius: 9px !important;
}

/* 已選取文字 */
div[data-baseweb="select"] div {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    line-height: 1.3 !important;
}

/* 下拉項目的文字也允許換行 */
li[role="option"] {
    white-space: normal !important;
    height: auto !important;
    min-height: 42px !important;
    line-height: 1.35 !important;
}


/* ---------------------------------------------------------
   輸入框
--------------------------------------------------------- */

div[data-testid="stTextInput"] {
    margin-bottom: 0.2rem !important;
}


/* ---------------------------------------------------------
   星期＋日期
--------------------------------------------------------- */

.day-header {
    display: flex;
    align-items: center;
    gap: 7px;
    margin-top: 2px;
    margin-bottom: 4px;
    min-height: 30px;
}

.day-name {
    font-size: 18px;
    font-weight: 700;
    white-space: nowrap;
}

.day-date {
    font-size: 14px;
    color: #16823b;
    background: #f5f7f7;
    padding: 2px 7px;
    border-radius: 6px;
    white-space: nowrap;
}


/* ---------------------------------------------------------
   預覽表格
--------------------------------------------------------- */

.cleaning-preview-wrapper {
    width: 100%;
    overflow: hidden;
}

.cleaning-preview-table {
    width: 100%;
    table-layout: fixed;
    border-collapse: collapse;
    font-size: 13px;
    background: white;
}

.cleaning-preview-table th {
    border: 1px solid #d9dce1;
    background: #f5f6f8;
    padding: 7px 5px;
    text-align: center;
    vertical-align: middle;
    font-weight: 700;
}

.cleaning-preview-table td {
    border: 1px solid #d9dce1;
    padding: 7px 6px;
    vertical-align: top;
    white-space: normal;
    word-break: break-word;
    overflow-wrap: anywhere;
    line-height: 1.4;
}

.cleaning-preview-table .week-col {
    width: 6%;
    text-align: center;
    vertical-align: middle;
    font-weight: 700;
}

.cleaning-preview-table .day-col {
    width: 18.8%;
}

.preview-date {
    font-weight: 700;
    color: #16823b;
    margin-bottom: 4px;
}

.preview-task {
    margin-bottom: 5px;
    color: #262730;
}

.preview-teacher {
    color: #666666;
    font-size: 12px;
}


/* ---------------------------------------------------------
   較小螢幕
--------------------------------------------------------- */

@media (max-width: 1300px) {

    .block-container {
        padding-left: 0.7rem !important;
        padding-right: 0.7rem !important;
    }

    div[data-testid="stHorizontalBlock"] {
        gap: 0.4rem !important;
    }

    .day-name {
        font-size: 16px;
    }

    .day-date {
        font-size: 12px;
        padding: 2px 4px;
    }

    .cleaning-preview-table {
        font-size: 11px;
    }

    .cleaning-preview-table td {
        padding: 5px 4px;
    }
}

</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# 預設清潔項目
# =========================================================

DEFAULT_TASKS_A = [
    "樓下酵素熱水倒水槽(幼兒洗手水槽、門口小水槽、洗屁屁區、備餐區，使用2/3桶熱水配半杓酵素攪拌均勻)",
    "樓梯間掃地拖地(包含牆壁檯面灰塵清潔)",
    "晨檢區灰塵及洗手台(檯面及角落灰塵清潔、洗手台用漂白水噴式擦拭、掃地及拖地)",
    "樓上廁所(洗手台漂白水噴拭除黴、馬桶清潔刷淨、地板掃地拖地)",
    "辦公室灰塵(檯面整理、灰塵擦拭乾淨、後方櫃子除塵、監視器及縫隙除塵)",
    "清點備品",
    "除濕機濾網跟接水槽清潔",
    "刪除上個月Line相簿",
    "廚房門口上面紗窗清潔(使用雞毛撢子撥灰塵)",
    "掃地機器人/寢具表填寫/規劃戶外活動計畫表與回報(髒水倒掉、洗淨、補充乾淨水、清潔液1:50)",
]

DEFAULT_TASKS_B = [
    "酵素熱水倒水槽(幼兒洗手水槽、門口小水槽、洗屁屁區、備餐區，使用2/3桶熱水配半杓酵素攪拌均勻)",
    "樓梯間掃地拖地(包含玻璃檯面灰塵清潔)",
    "晨檢區灰塵及洗手台(檯面及角落灰塵清潔、洗手台用漂白水噴式擦拭、掃地及拖地)",
    "樓上廁所(洗手台漂白水噴拭除黴、馬桶清潔刷淨、地板掃地拖地)",
    "辦公室灰塵(檯面整理、灰塵擦拭乾淨、後方櫃子除塵、監視器及縫隙除塵)",
    "清點備品",
    "除濕機濾網跟接水槽清潔",
    "刪除上個月Line相簿",
    "廚房門口上面紗窗清潔(使用雞毛撢子撥灰塵)",
    "戶外掃落葉/寢具表填寫/規劃戶外活動計畫表與回報",
]


# =========================================================
# Session State 初始化
# =========================================================

DEFAULT_TEACHERS = [
    "主任",
    "均宜",
    "小安",
    "綺綺",
    "嘉鳳",
    "樺樺",
    "Candy",
    "Panda",
]

for prefix, tasks_init in [
    ("A", DEFAULT_TASKS_A),
    ("B", DEFAULT_TASKS_B),
]:

    if f"teachers_{prefix}" not in st.session_state:
        st.session_state[f"teachers_{prefix}"] = DEFAULT_TEACHERS.copy()

    if f"tasks_{prefix}" not in st.session_state:
        st.session_state[f"tasks_{prefix}"] = tasks_init.copy()

    if f"notes_{prefix}" not in st.session_state:
        st.session_state[f"notes_{prefix}"] = (
            "備註：請各位老師確實執行清潔項目，"
            "若遇請假請務必提前找職務代理人協助。"
        )

    if f"history_{prefix}" not in st.session_state:
        st.session_state[f"history_{prefix}"] = []


# =========================================================
# 假日調整
# =========================================================

def get_adjusted_workday(target_date, tw_holidays):
    """
    如果遇週末或國定假日：
    優先往前找最近一個工作日。
    """
    if (
        target_date.weekday() < 5
        and target_date not in tw_holidays
    ):
        return target_date

    curr = target_date - datetime.timedelta(days=1)

    while (
        curr.weekday() >= 5
        or curr in tw_holidays
    ):
        curr -= datetime.timedelta(days=1)

    return curr


# =========================================================
# Excel 產生
# =========================================================

def generate_cleaning_excel(
    branch_name,
    year_roc,
    month,
    schedule_df,
    notes_text,
):

    output = io.BytesIO()

    wb = Workbook()
    ws = wb.active
    ws.title = f"{month}月清潔輪值"

    # -----------------------------------------------------
    # 頁面設定
    # -----------------------------------------------------

    ws.sheet_view.showGridLines = False

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.freeze_panes = "C4"

    # -----------------------------------------------------
    # 樣式
    # -----------------------------------------------------

    thin_gray = Side(
        style="thin",
        color="B7B7B7",
    )

    border = Border(
        left=thin_gray,
        right=thin_gray,
        top=thin_gray,
        bottom=thin_gray,
    )

    title_fill = PatternFill(
        "solid",
        fgColor="D9EAF7",
    )

    header_fill = PatternFill(
        "solid",
        fgColor="EDEDED",
    )

    date_fill = PatternFill(
        "solid",
        fgColor="F7FBF8",
    )

    teacher_fill = PatternFill(
        "solid",
        fgColor="FFF7E6",
    )

    sign_fill = PatternFill(
        "solid",
        fgColor="FFFFFF",
    )

    # -----------------------------------------------------
    # 標題
    # -----------------------------------------------------

    ws.merge_cells("A1:G1")

    title_cell = ws["A1"]

    title_cell.value = (
        f"{branch_name} "
        f"{year_roc} 年 {month} 月清潔輪值紀錄表"
    )

    title_cell.font = Font(
        name="微軟正黑體",
        size=16,
        bold=True,
        color="1F497D",
    )

    title_cell.fill = title_fill

    title_cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
    )

    ws.row_dimensions[1].height = 30

    # -----------------------------------------------------
    # 表頭
    # -----------------------------------------------------

    headers = [
        "週次",
        "項目",
        "星期一",
        "星期二",
        "星期三",
        "星期四",
        "星期五",
    ]

    header_row = 3

    for col_idx, header in enumerate(
        headers,
        start=1,
    ):

        cell = ws.cell(
            row=header_row,
            column=col_idx,
            value=header,
        )

        cell.font = Font(
            name="微軟正黑體",
            size=11,
            bold=True,
        )

        cell.fill = header_fill

        cell.border = border

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    ws.row_dimensions[header_row].height = 25

    # -----------------------------------------------------
    # 排班資料
    # -----------------------------------------------------

    row_num = 4

    days_k = ["一", "二", "三", "四", "五"]

    for _, record in schedule_df.iterrows():

        start_row = row_num

        week_name = record["週次"]

        # 合併週次
        ws.merge_cells(
            start_row=start_row,
            start_column=1,
            end_row=start_row + 3,
            end_column=1,
        )

        week_cell = ws.cell(
            row=start_row,
            column=1,
            value=week_name,
        )

        week_cell.font = Font(
            name="微軟正黑體",
            size=11,
            bold=True,
        )

        week_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        week_cell.border = border

        # 四種資料列
        row_types = [
            ("日期", "_0"),
            ("清潔內容", "_1"),
            ("執行老師", "_2"),
            ("簽名", None),
        ]

        for offset, (
            row_label,
            suffix,
        ) in enumerate(row_types):

            current_row = start_row + offset

            label_cell = ws.cell(
                row=current_row,
                column=2,
                value=row_label,
            )

            label_cell.font = Font(
                name="微軟正黑體",
                size=10,
                bold=True,
            )

            label_cell.border = border

            label_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

            if row_label == "日期":
                label_cell.fill = date_fill

            elif row_label == "執行老師":
                label_cell.fill = teacher_fill

            elif row_label == "簽名":
                label_cell.fill = sign_fill

            # 星期一～星期五
            for i, dk in enumerate(
                days_k,
                start=3,
            ):

                if suffix is None:
                    value = ""

                else:
                    value = record[
                        f"{dk}{suffix}"
                    ]

                cell = ws.cell(
                    row=current_row,
                    column=i,
                    value=value,
                )

                cell.font = Font(
                    name="微軟正黑體",
                    size=10,
                )

                cell.border = border

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )

                if row_label == "日期":
                    cell.fill = date_fill
                    cell.font = Font(
                        name="微軟正黑體",
                        size=10,
                        bold=True,
                        color="00843D",
                    )

                elif row_label == "執行老師":
                    cell.fill = teacher_fill

                elif row_label == "簽名":
                    cell.fill = sign_fill

        # 列高
        ws.row_dimensions[start_row].height = 25
        ws.row_dimensions[start_row + 1].height = 65
        ws.row_dimensions[start_row + 2].height = 28
        ws.row_dimensions[start_row + 3].height = 38

        row_num += 4

    # -----------------------------------------------------
    # 修正合併格框線
    # -----------------------------------------------------

    for row in ws.iter_rows(
        min_row=4,
        max_row=row_num - 1,
        min_col=1,
        max_col=7,
    ):

        for cell in row:
            cell.border = border

    # -----------------------------------------------------
    # 備註
    # -----------------------------------------------------

    notes_row = row_num + 1

    ws.merge_cells(
        start_row=notes_row,
        start_column=1,
        end_row=notes_row,
        end_column=7,
    )

    notes_cell = ws.cell(
        row=notes_row,
        column=1,
        value=notes_text,
    )

    notes_cell.font = Font(
        name="微軟正黑體",
        size=10,
        color="333333",
    )

    notes_cell.alignment = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=True,
    )

    notes_cell.border = border

    ws.row_dimensions[notes_row].height = 38

    # -----------------------------------------------------
    # 欄寬
    # -----------------------------------------------------

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 11

    for col in [
        "C",
        "D",
        "E",
        "F",
        "G",
    ]:
        ws.column_dimensions[col].width = 27

    # -----------------------------------------------------
    # 列印範圍
    # -----------------------------------------------------

    ws.print_area = f"A1:G{notes_row}"

    ws.oddHeader.center.text = (
        f"&B{branch_name} 清潔輪值表"
    )

    ws.oddFooter.center.text = (
        "第 &P 頁／共 &N 頁"
    )

    # -----------------------------------------------------
    # 儲存
    # -----------------------------------------------------

    wb.save(output)

    output.seek(0)

    return output


# =========================================================
# APP 標題
# =========================================================

st.title("🏫 澳森托嬰中心 月清潔輪值與表單管理系統")

st.caption(
    "支援澳森與澳森文德雙分園｜"
    "排班預覽｜假日調整｜Excel 匯出｜歷史紀錄"
)


# =========================================================
# 雙分園 Tab
# =========================================================

tab_a, tab_b = st.tabs(
    [
        "🌳 澳森分園",
        "🌸 澳森文德分園",
    ]
)


# =========================================================
# 分園畫面
# =========================================================

def render_branch_tab(
    branch_name,
    prefix,
):

    # -----------------------------------------------------
    # 基本參數
    # -----------------------------------------------------

    st.subheader(
        f"📌 {branch_name}｜基本參數設定"
    )

    col1, col2, col3 = st.columns(
        [1, 1, 1.4]
    )

    with col1:

        year_roc = st.number_input(
            "民國年份",
            min_value=114,
            max_value=125,
            value=115,
            step=1,
            key=f"y_{prefix}",
        )

    with col2:

        month = st.selectbox(
            "月份",
            options=list(range(1, 13)),
            index=8,
            key=f"m_{prefix}",
        )

    with col3:

        st.markdown("<br>", unsafe_allow_html=True)

        save_clicked = st.button(
            "💾 儲存目前排班至歷史紀錄",
            key=f"save_hist_{prefix}",
            use_container_width=True,
        )

    # -----------------------------------------------------
    # 老師名單
    # -----------------------------------------------------

    st.markdown("#### 👥 執行清潔老師名單")

    teacher_text = st.text_input(
        "姓名以逗號分隔",
        value=", ".join(
            st.session_state[
                f"teachers_{prefix}"
            ]
        ),
        key=f"teacher_input_{prefix}",
    )

    teachers = [
        name.strip()
        for name in teacher_text.split(",")
        if name.strip()
    ]

    if not teachers:
        teachers = ["主任"]

    st.session_state[
        f"teachers_{prefix}"
    ] = teachers

    # -----------------------------------------------------
    # 清潔工作管理
    # -----------------------------------------------------

    with st.expander(
        "🧹 清潔工作項目管理",
        expanded=False,
    ):

        task_list = st.session_state[
            f"tasks_{prefix}"
        ]

        with st.form(
            key=f"add_task_form_{prefix}",
            clear_on_submit=True,
        ):

            new_task_name = st.text_input(
                "新增清潔工作項目名稱與說明"
            )

            submitted = st.form_submit_button(
                "➕ 新增工作項目"
            )

            if (
                submitted
                and new_task_name.strip()
            ):

                st.session_state[
                    f"tasks_{prefix}"
                ].append(
                    new_task_name.strip()
                )

                st.rerun()

        updated_tasks = []

        for idx, task_item in enumerate(
            task_list
        ):

            c1, c2 = st.columns(
                [8.5, 1.5]
            )

            with c1:

                edited_task = st.text_input(
                    f"項目 {idx + 1}",
                    value=task_item,
                    key=(
                        f"task_card_"
                        f"{prefix}_{idx}"
                    ),
                    label_visibility="collapsed",
                )

                if edited_task.strip():
                    updated_tasks.append(
                        edited_task.strip()
                    )

            with c2:

                if st.button(
                    "🗑️",
                    key=(
                        f"delete_task_"
                        f"{prefix}_{idx}"
                    ),
                    help="刪除此工作項目",
                    use_container_width=True,
                ):

                    task_list.pop(idx)

                    st.session_state[
                        f"tasks_{prefix}"
                    ] = task_list

                    st.rerun()

        if updated_tasks:

            st.session_state[
                f"tasks_{prefix}"
            ] = updated_tasks

    st.divider()

    # -----------------------------------------------------
    # 月曆產生
    # -----------------------------------------------------

    year_ad = int(year_roc) + 1911

    tw_holidays = holidays.Taiwan(
        years=year_ad
    )

    cal = calendar.monthcalendar(
        year_ad,
        month,
    )

    work_weeks = []

    for week in cal:

        monday_to_friday = week[0:5]

        if any(
            day > 0
            for day in monday_to_friday
        ):

            work_weeks.append(
                monday_to_friday
            )

    st.subheader(
        f"📅 {branch_name} "
        f"{year_roc} 年 {month} 月清潔輪值"
    )

    current_tasks = st.session_state[
        f"tasks_{prefix}"
    ]

    if not current_tasks:
        current_tasks = ["一般清潔"]

    teacher_pool = st.session_state[
        f"teachers_{prefix}"
    ]

    if not teacher_pool:
        teacher_pool = ["主任"]

    table_records = []

    task_index = 0

    week_labels = [
        "第一週",
        "第二週",
        "第三週",
        "第四週",
        "第五週",
        "第六週",
    ]

    # -----------------------------------------------------
    # 每週排班
    # -----------------------------------------------------

    for week_index, week in enumerate(
        work_weeks
    ):

        week_name = week_labels[
            week_index
        ]

        with st.expander(
            f"📌 {week_name}",
            expanded=True,
        ):

            row_dates = []
            row_tasks = []
            row_teachers = []

            cols = st.columns(5)

            days_label = [
                "週一",
                "週二",
                "週三",
                "週四",
                "週五",
            ]

            for day_index, day_value in enumerate(
                week
            ):

                with cols[day_index]:

                    # =====================================
                    # 日期
                    # =====================================

                    if day_value > 0:

                        target_dt = datetime.date(
                            year_ad,
                            month,
                            day_value,
                        )

                        adjusted_dt = (
                            get_adjusted_workday(
                                target_dt,
                                tw_holidays,
                            )
                        )

                        date_str = (
                            f"{month}/{day_value}"
                        )

                        if (
                            adjusted_dt
                            != target_dt
                        ):

                            date_str += (
                                " "
                                f"(調日至"
                                f"{adjusted_dt.month}/"
                                f"{adjusted_dt.day})"
                            )

                    else:

                        date_str = "－"

                    safe_date = html.escape(
                        date_str
                    )

                    safe_day_name = html.escape(
                        days_label[day_index]
                    )

                    st.markdown(
                        f"""
                        <div class="day-header">
                            <span class="day-name">
                                {safe_day_name}
                            </span>
                            <span class="day-date">
                                {safe_date}
                            </span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    # =====================================
                    # 空白日期
                    # =====================================

                    if day_value == 0:

                        chosen_task = (
                            st.selectbox(
                                (
                                    f"工作_"
                                    f"{week_index}_"
                                    f"{day_index}"
                                ),
                                options=["無"],
                                index=0,
                                key=(
                                    f"task_"
                                    f"{prefix}_"
                                    f"{year_roc}_"
                                    f"{month}_"
                                    f"{week_index}_"
                                    f"{day_index}"
                                ),
                                label_visibility=(
                                    "collapsed"
                                ),
                                disabled=True,
                            )
                        )

                        chosen_teacher = (
                            st.selectbox(
                                (
                                    f"老師_"
                                    f"{week_index}_"
                                    f"{day_index}"
                                ),
                                options=["無"],
                                index=0,
                                key=(
                                    f"teacher_"
                                    f"{prefix}_"
                                    f"{year_roc}_"
                                    f"{month}_"
                                    f"{week_index}_"
                                    f"{day_index}"
                                ),
                                label_visibility=(
                                    "collapsed"
                                ),
                                disabled=True,
                            )
                        )

                    # =====================================
                    # 有效日期
                    # =====================================

                    else:

                        # 特殊規則：
                        # 第2、4週星期二
                        # 預設清點備品
                        if (
                            day_index == 1
                            and (
                                week_index + 1
                                in [2, 4]
                            )
                        ):

                            matched = [
                                task
                                for task
                                in current_tasks
                                if "清點備品"
                                in task
                            ]

                            default_task = (
                                matched[0]
                                if matched
                                else current_tasks[0]
                            )

                        # 星期五
                        elif day_index == 4:

                            matched = [
                                task
                                for task
                                in current_tasks
                                if (
                                    "掃地機器人"
                                    in task
                                    or "戶外掃落葉"
                                    in task
                                    or "規劃戶外活動"
                                    in task
                                )
                            ]

                            default_task = (
                                matched[0]
                                if matched
                                else current_tasks[-1]
                            )

                        else:

                            default_task = (
                                current_tasks[
                                    task_index
                                    % len(
                                        current_tasks
                                    )
                                ]
                            )

                            task_index += 1

                        # -----------------------------
                        # 清潔項目
                        # -----------------------------

                        chosen_task = (
                            st.selectbox(
                                (
                                    f"工作_"
                                    f"{week_index}_"
                                    f"{day_index}"
                                ),
                                options=current_tasks,
                                index=(
                                    current_tasks.index(
                                        default_task
                                    )
                                    if default_task
                                    in current_tasks
                                    else 0
                                ),
                                key=(
                                    f"task_"
                                    f"{prefix}_"
                                    f"{year_roc}_"
                                    f"{month}_"
                                    f"{week_index}_"
                                    f"{day_index}"
                                ),
                                label_visibility=(
                                    "collapsed"
                                ),
                            )
                        )

                        # -----------------------------
                        # 老師
                        # -----------------------------

                        default_teacher = (
                            teacher_pool[
                                (
                                    week_index * 5
                                    + day_index
                                )
                                % len(
                                    teacher_pool
                                )
                            ]
                        )

                        chosen_teacher = (
                            st.selectbox(
                                (
                                    f"老師_"
                                    f"{week_index}_"
                                    f"{day_index}"
                                ),
                                options=teacher_pool,
                                index=(
                                    teacher_pool.index(
                                        default_teacher
                                    )
                                ),
                                key=(
                                    f"teacher_"
                                    f"{prefix}_"
                                    f"{year_roc}_"
                                    f"{month}_"
                                    f"{week_index}_"
                                    f"{day_index}"
                                ),
                                label_visibility=(
                                    "collapsed"
                                ),
                            )
                        )

                    row_dates.append(
                        date_str
                    )

                    row_tasks.append(
                        chosen_task
                    )

                    row_teachers.append(
                        chosen_teacher
                    )

            table_records.append(
                {
                    "週次": week_name,
                    "日期": row_dates,
                    "內容": row_tasks,
                    "老師": row_teachers,
                }
            )

    # -----------------------------------------------------
    # 備註
    # -----------------------------------------------------

    st.markdown(
        "#### 📝 備註與提醒事項"
    )

    st.session_state[
        f"notes_{prefix}"
    ] = st.text_area(
        "備註",
        value=st.session_state[
            f"notes_{prefix}"
        ],
        height=75,
        key=f"notes_area_{prefix}",
        label_visibility="collapsed",
    )

    # -----------------------------------------------------
    # DataFrame
    # -----------------------------------------------------

    export_rows = []

    for record in table_records:

        row_dict = {
            "週次": record["週次"]
        }

        days_k = [
            "一",
            "二",
            "三",
            "四",
            "五",
        ]

        for i, dk in enumerate(days_k):

            row_dict[
                f"{dk}_0"
            ] = record["日期"][i]

            row_dict[
                f"{dk}_1"
            ] = record["內容"][i]

            row_dict[
                f"{dk}_2"
            ] = record["老師"][i]

        export_rows.append(
            row_dict
        )

    final_df = pd.DataFrame(
        export_rows
    )

    # -----------------------------------------------------
    # 儲存歷史
    # -----------------------------------------------------

    if save_clicked:

        history_item = {
            "title": (
                f"{year_roc}年"
                f"{month}月 清潔輪值表"
            ),
            "timestamp": (
                datetime.datetime.now()
                .strftime(
                    "%Y-%m-%d %H:%M"
                )
            ),
            "branch_name": branch_name,
            "year_roc": int(year_roc),
            "month": int(month),
            "df": final_df.copy(),
            "notes": st.session_state[
                f"notes_{prefix}"
            ],
        }

        st.session_state[
            f"history_{prefix}"
        ].append(
            history_item
        )

        st.success(
            f"✅ 已儲存 "
            f"{branch_name} "
            f"{year_roc}年"
            f"{month}月排班紀錄"
        )

    # -----------------------------------------------------
    # 即時預覽
    # -----------------------------------------------------

    st.divider()

    st.subheader(
        "👁️ 月曆式即時預覽"
    )

    preview_html = """
    <div class="cleaning-preview-wrapper">
    <table class="cleaning-preview-table">

        <thead>
            <tr>
                <th class="week-col">週次</th>
                <th class="day-col">星期一</th>
                <th class="day-col">星期二</th>
                <th class="day-col">星期三</th>
                <th class="day-col">星期四</th>
                <th class="day-col">星期五</th>
            </tr>
        </thead>

        <tbody>
    """

    for record in table_records:

        preview_html += (
            "<tr>"
            f'<td class="week-col">'
            f'{html.escape(record["週次"])}'
            "</td>"
        )

        for i in range(5):

            safe_date = html.escape(
                str(record["日期"][i])
            )

            safe_task = html.escape(
                str(record["內容"][i])
            )

            safe_teacher = html.escape(
                str(record["老師"][i])
            )

            preview_html += f"""
                <td class="day-col">

                    <div class="preview-date">
                        {safe_date}
                    </div>

                    <div class="preview-task">
                        {safe_task}
                    </div>

                    <div class="preview-teacher">
                        負責：{safe_teacher}
                    </div>

                </td>
            """

        preview_html += "</tr>"

    preview_html += """
        </tbody>
    </table>
    </div>
    """

    st.markdown(
        preview_html,
        unsafe_allow_html=True,
    )

    # -----------------------------------------------------
    # Excel
    # -----------------------------------------------------

    st.divider()

    st.subheader(
        f"📥 匯出 {branch_name} Excel"
    )

    try:

        excel_bytes = (
            generate_cleaning_excel(
                branch_name=branch_name,
                year_roc=int(
                    year_roc
                ),
                month=int(
                    month
                ),
                schedule_df=final_df,
                notes_text=(
                    st.session_state[
                        f"notes_{prefix}"
                    ]
                ),
            )
        )

        st.download_button(
            label=(
                f"📥 下載【"
                f"{branch_name} "
                f"{year_roc}年"
                f"{month}月清潔輪值表】"
            ),
            data=excel_bytes,
            file_name=(
                f"{branch_name}_"
                f"{year_roc}年"
                f"{month}月"
                f"清潔輪值紀錄表.xlsx"
            ),
            mime=(
                "application/"
                "vnd.openxmlformats-"
                "officedocument."
                "spreadsheetml.sheet"
            ),
            key=(
                f"download_excel_"
                f"{prefix}_"
                f"{year_roc}_"
                f"{month}"
            ),
            use_container_width=True,
        )

    except Exception as error:

        st.error(
            "❌ Excel 產生失敗"
        )

        st.exception(error)

    # -----------------------------------------------------
    # 歷史紀錄
    # -----------------------------------------------------

    st.divider()

    st.subheader(
        f"📂 歷史輪值紀錄｜"
        f"{branch_name}"
    )

    history_list = st.session_state[
        f"history_{prefix}"
    ]

    if not history_list:

        st.info(
            "目前尚無歷史紀錄。"
            "完成排班後，可點擊上方"
            "「儲存目前排班至歷史紀錄」。"
        )

    else:

        reversed_history = list(
            reversed(
                history_list
            )
        )

        for history_index, history in enumerate(
            reversed_history
        ):

            with st.expander(
                (
                    f"📜 "
                    f"[{history['timestamp']}] "
                    f"{history['title']}"
                ),
                expanded=False,
            ):

                st.markdown(
                    "**備註：** "
                    + html.escape(
                        str(
                            history[
                                "notes"
                            ]
                        )
                    )
                )

                history_preview = """
                <div class="cleaning-preview-wrapper">
                <table class="cleaning-preview-table">

                    <thead>
                        <tr>
                            <th class="week-col">週次</th>
                            <th class="day-col">星期一</th>
                            <th class="day-col">星期二</th>
                            <th class="day-col">星期三</th>
                            <th class="day-col">星期四</th>
                            <th class="day-col">星期五</th>
                        </tr>
                    </thead>

                    <tbody>
                """

                for _, row in history[
                    "df"
                ].iterrows():

                    history_preview += (
                        "<tr>"
                        '<td class="week-col">'
                        f'{html.escape(str(row["週次"]))}'
                        "</td>"
                    )

                    days_k = [
                        "一",
                        "二",
                        "三",
                        "四",
                        "五",
                    ]

                    for dk in days_k:

                        h_date = html.escape(
                            str(
                                row[
                                    f"{dk}_0"
                                ]
                            )
                        )

                        h_task = html.escape(
                            str(
                                row[
                                    f"{dk}_1"
                                ]
                            )
                        )

                        h_teacher = html.escape(
                            str(
                                row[
                                    f"{dk}_2"
                                ]
                            )
                        )

                        history_preview += f"""
                        <td class="day-col">

                            <div class="preview-date">
                                {h_date}
                            </div>

                            <div class="preview-task">
                                {h_task}
                            </div>

                            <div class="preview-teacher">
                                負責：{h_teacher}
                            </div>

                        </td>
                        """

                    history_preview += "</tr>"

                history_preview += """
                    </tbody>
                </table>
                </div>
                """

                st.markdown(
                    history_preview,
                    unsafe_allow_html=True,
                )

                try:

                    history_excel = (
                        generate_cleaning_excel(
                            branch_name=history.get(
                                "branch_name",
                                branch_name,
                            ),
                            year_roc=history.get(
                                "year_roc",
                                int(year_roc),
                            ),
                            month=history.get(
                                "month",
                                int(month),
                            ),
                            schedule_df=history[
                                "df"
                            ],
                            notes_text=history[
                                "notes"
                            ],
                        )
                    )

                    st.download_button(
                        label=(
                            "📥 重新下載此歷史版本"
                        ),
                        data=history_excel,
                        file_name=(
                            f"{branch_name}_"
                            f"{history['title']}_"
                            f"歷史存檔.xlsx"
                        ),
                        mime=(
                            "application/"
                            "vnd.openxmlformats-"
                            "officedocument."
                            "spreadsheetml.sheet"
                        ),
                        key=(
                            f"history_download_"
                            f"{prefix}_"
                            f"{history_index}_"
                            f"{history['timestamp']}"
                        ),
                        use_container_width=True,
                    )

                except Exception as error:

                    st.error(
                        "此歷史版本 Excel "
                        "產生失敗"
                    )

                    st.exception(error)


# =========================================================
# 顯示兩個分園
# =========================================================

with tab_a:

    render_branch_tab(
        "澳森",
        "A",
    )


with tab_b:

    render_branch_tab(
        "澳森文德",
        "B",
    )
