import streamlit as st
import pandas as pd
import datetime
import calendar
import holidays
import io
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

st.set_page_config(page_title="澳森托嬰中心 月清潔輪值與表單管理系統", layout="wide")

# --- Word 排版樣式輔助工具 ---
def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=70, bottom=70, left=70, right=70):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="1F497D", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

DEFAULT_TASKS = [
    "樓下酵素熱水倒水槽\n(幼兒洗手水槽、門口小水槽、洗屁屁區、備餐區，使用2/3桶熱水配半杓酵素攪拌均勻)",
    "樓梯間掃地拖地\n(包含牆壁檯面灰塵清潔)",
    "晨檢區灰塵及洗手台\n(檯面及角落灰塵清潔、洗手台用漂白水噴式擦拭、掃地及拖地)",
    "樓上廁所\n(洗手台漂白水噴拭除黴、馬桶清潔刷淨、地板掃地拖地)",
    "辦公室灰塵\n(檯面整理、灰塵擦拭乾淨、後方櫃子除塵、監視器及縫隙除塵)",
    "清點備品",
    "除濕機濾網跟接水槽清潔",
    "刪除上個月Line相簿",
    "廚房門口上面紗窗清潔\n(使用雞毛撢子撥灰塵)",
    "掃地機器人/寢具表填寫/規劃戶外活動計畫表與回報\n(髒水倒掉、洗淨、補充乾淨水、清潔液1:50)"
]

# --- 初始化 Session State ---
for prefix in ["A", "B"]:
    if f"teachers_{prefix}" not in st.session_state:
        st.session_state[f"teachers_{prefix}"] = ["主任", "均宜", "小安", "綺綺", "嘉鳳", "樺樺", "Candy", "Panda"]
    if f"tasks_{prefix}" not in st.session_state:
        st.session_state[f"tasks_{prefix}"] = DEFAULT_TASKS.copy()
    if f"notes_{prefix}" not in st.session_state:
        st.session_state[f"notes_{prefix}"] = "備註：請各位老師確實執行清潔項目，若遇請假請務必提前找職務代理人協助。"
    if f"history_{prefix}" not in st.session_state:
        st.session_state[f"history_{prefix}"] = []

def get_adjusted_workday(target_date, tw_holidays):
    curr = target_date
    while curr in tw_holidays or curr.weekday() >= 5:
        prev_day = curr - datetime.timedelta(days=1)
        if prev_day.weekday() < 5 and prev_day not in tw_holidays:
            curr = prev_day
            break
        else:
            curr = curr + datetime.timedelta(days=1)
    return curr

def generate_cleaning_docx(branch_name, year_roc, month, schedule_df, notes_text):
    doc = Document()
    for section in doc.sections:
        section.page_width = Inches(11.69)
        section.page_height = Inches(8.27)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    r_title = title_p.add_run(f"{branch_name} {year_roc} 年 {month} 月清潔輪值紀錄表")
    r_title.font.name = "微軟正黑體"
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    table = doc.add_table(rows=len(schedule_df) * 3 + 1, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="1F497D", sz="4")

    col_widths = [Inches(1.2), Inches(1.9), Inches(1.9), Inches(1.9), Inches(1.9), Inches(1.9)]
    headers = ["週次", "星期一", "星期二", "星期三", "星期四", "星期五"]

    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.width = col_widths[i]
        set_cell_background(cell, "1F497D")
        set_cell_margins(cell, top=60, bottom=60, left=50, right=50)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.name = "微軟正黑體"
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r.font.size = Pt(12)

    for idx, row in schedule_df.iterrows():
        r_base = idx * 3 + 1
        labels = ["日期", "清潔內容", "執行老師"]
        for sub_i, label in enumerate(labels):
            cell_lbl = table.cell(r_base + sub_i, 0)
            cell_lbl.width = col_widths[0]
            set_cell_background(cell_lbl, "DCE6F1")
            set_cell_margins(cell_lbl, top=40, bottom=40, left=40, right=40)
            p_l = cell_lbl.paragraphs[0]
            p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_l = p_l.add_run(f"{row['週次']}\n{label}")
            r_l.font.name = "微軟正黑體"
            r_l.font.size = Pt(11)
            r_l.font.bold = True
            r_l.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

            days_keys = ['一', '二', '三', '四', '五']
            for d_idx, d_key in enumerate(days_keys):
                cell_d = table.cell(r_base + sub_i, d_idx + 1)
                cell_d.width = col_widths[d_idx + 1]
                set_cell_margins(cell_d, top=40, bottom=40, left=40, right=40)
                p_d = cell_d.paragraphs[0]
                p_d.alignment = WD_ALIGN_PARAGRAPH.LEFT

                val = row[f"{d_key}_{sub_i}"]
                r_d = p_d.add_run(str(val) if val else "")
                r_d.font.name = "微軟正黑體"
                r_d.font.size = Pt(12)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    note_p = doc.add_paragraph()
    r_note = note_p.add_run(notes_text)
    r_note.font.name = "微軟正黑體"
    r_note.font.size = Pt(12)
    r_note.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

st.title("🏫 澳森托嬰中心 月清潔輪值與表單管理系統")
st.markdown("支援 **澳森** 與 **澳森文德** 雙分園，具備 **即時預覽功能** 與 **歷史輪值紀錄儲存/查看**！")

tab_a, tab_b = st.tabs(["🌳 澳森分園", "🌸 澳森文德分園"])

def render_branch_tab(branch_name, prefix):
    st.subheader(f"📌 {branch_name} - 基本參數設定")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        year_roc = st.number_input("民國年份：", min_value=114, max_value=125, value=115, key=f"y_{branch_name}")
    with col2:
        month = st.selectbox("月份：", list(range(1, 13)), index=8, key=f"m_{branch_name}")
    with col3:
        st.markdown("<br>", unsafe_allow_html=True)
        save_clicked = st.button(f"💾 儲存目前排班至歷史紀錄", key=f"save_hist_{branch_name}")

    st.markdown("#### 👥 執行清潔老師名單（可編輯並自動記憶）")
    t_input = st.text_input("請輸入老師姓名（以逗號分隔）：", value=", ".join(st.session_state[f"teachers_{prefix}"]), key=f"t_in_{branch_name}")
    st.session_state[f"teachers_{prefix}"] = [t.strip() for t in t_input.split(",") if t.strip()]

    st.markdown("#### 🧹 清潔工作項目管理（可編輯與自訂）")
    task_text = st.text_area("清潔項目清單（每行一項）：", value="\n".join(st.session_state[f"tasks_{prefix}"]), height=130, key=f"task_in_{branch_name}")
    st.session_state[f"tasks_{prefix}"] = [tk.strip() for tk in task_text.split("\n") if tk.strip()]

    st.divider()

    year_ad = year_roc + 1911
    tw_holidays = holidays.Taiwan(years=year_ad)
    cal = calendar.monthcalendar(year_ad, month)
    work_weeks = []
    for w in cal:
        mon_to_fri = w[0:5]
        if any(d > 0 for d in mon_to_fri):
            work_weeks.append(mon_to_fri)

    st.subheader(f"📅 {branch_name} {year_roc} 年 {month} 月清潔輪值排班表（自動順延假日）")

    table_records = []
    teacher_pool = st.session_state[f"teachers_{prefix}"] if st.session_state[f"teachers_{prefix}"] else ["主任"]
    task_pool = st.session_state[f"tasks_{prefix}"] if st.session_state[f"tasks_{prefix}"] else DEFAULT_TASKS

    t_idx = 0
    for idx, w in enumerate(work_weeks[:4], start=1):
        week_name = f"第{['一','二','三','四'][idx-1]}週"
        with st.expander(f"📌 {week_name} 排班與任務細節", expanded=True):
            cols = st.columns(5)
            days_label = ['週一', '週二', '週三', '週四', '週五']
            
            row_dates = []
            row_tasks = []
            row_teachers = []

            for d_i, day_val in enumerate(w):
                with cols[d_i]:
                    if day_val > 0:
                        target_dt = datetime.date(year_ad, month, day_val)
                        adjusted_dt = get_adjusted_workday(target_dt, tw_holidays)
                        date_str = f"{month}/{day_val}"
                        if adjusted_dt != target_dt:
                            date_str += f"\n(調日至{adjusted_dt.month}/{adjusted_dt.day})"
                    else:
                        date_str = "-"

                    st.markdown(f"**{days_label[d_i]}**\n`{date_str}`")

                    default_task = ""
                    if day_val > 0:
                        if d_i == 1 and (idx in [2, 4]):
                            default_task = "清點備品"
                        else:
                            default_task = task_pool[t_idx % len(task_pool)]
                            t_idx += 1

                    chosen_task = st.text_area(f"工作_{idx}_{d_i}", value=default_task, height=75, key=f"task_{branch_name}_{idx}_{d_i}")
                    
                    default_teacher = teacher_pool[(idx * 5 + d_i) % len(teacher_pool)]
                    chosen_teacher = st.selectbox(f"老師_{idx}_{d_i}", teacher_options := (["主任"] + teacher_pool), index=teacher_options.index(default_teacher) if default_teacher in teacher_options else 0, key=f"tea_{branch_name}_{idx}_{d_i}")

                    row_dates.append(date_str)
                    row_tasks.append(chosen_task)
                    row_teachers.append(chosen_teacher)

            table_records.append({
                "週次": week_name,
                "日期": row_dates,
                "內容": row_tasks,
                "老師": row_teachers
            })

    st.markdown("#### 📝 備註與提醒事項")
    st.session_state[f"notes_{prefix}"] = st.text_area("備註欄位（供下次使用提醒）：", value=st.session_state[f"notes_{prefix}"], height=80, key=f"notes_area_{branch_name}")

    export_df_rows = []
    for rec in table_records:
        row_dict = {"週次": rec["週次"]}
        days_k = ['一', '二', '三', '四', '五']
        for i, dk in enumerate(days_k):
            row_dict[f"{dk}_0"] = rec["日期"][i]
            row_dict[f"{dk}_1"] = rec["內容"][i]
            row_dict[f"{dk}_2"] = rec["老師"][i]
        export_df_rows.append(row_dict)
    
    final_df = pd.DataFrame(export_df_rows)

    if save_clicked:
        history_item = {
            "title": f"{year_roc}年{month}月 清潔輪值表",
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "df": final_df,
            "notes": st.session_state[f"notes_{prefix}"]
        }
        st.session_state[f"history_{prefix}"].append(history_item)
        st.success(f"✅ 已成功儲存 {branch_name} {year_roc}年{month}月排班紀錄！")

    st.divider()
    st.subheader("👁️ 即時預覽清潔輪值表")
    preview_display_data = []
    for rec in table_records:
        preview_display_data.append({
            "週次": rec["週次"],
            "星期一": f"日期: {rec['日期'][0]}\n內容: {rec['內容'][0]}\n老師: {rec['老師'][0]}",
            "星期二": f"日期: {rec['日期'][1]}\n內容: {rec['內容'][1]}\n老師: {rec['老師'][1]}",
            "星期三": f"日期: {rec['日期'][2]}\n內容: {rec['內容'][2]}\n老師: {rec['老師'][2]}",
            "星期四": f"日期: {rec['日期'][3]}\n內容: {rec['內容'][3]}\n老師: {rec['老師'][3]}",
            "星期五": f"日期: {rec['日期'][4]}\n內容: {rec['內容'][4]}\n老師: {rec['老師'][4]}"
        })
    st.dataframe(pd.DataFrame(preview_display_data), use_container_width=True)

    st.divider()
    st.subheader(f"📥 匯出 {branch_name} 清潔輪值 Word 檔")
    doc_bytes = generate_cleaning_docx(branch_name, year_roc, month, final_df, st.session_state[f"notes_{prefix}"])

    st.download_button(
        label=f"📥 下載【{branch_name} {year_roc}年{month}月清潔輪值表】(Word 檔 / 12pt)",
        data=doc_bytes,
        file_name=f"{branch_name}_{year_roc}年{month}月清潔輪值紀錄表.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key=f"dl_btn_{branch_name}"
    )

    st.divider()
    st.subheader(f"📂 歷史儲存的輪值紀錄（{branch_name}）")
    hist_list = st.session_state[f"history_{prefix}"]
    if not hist_list:
        st.info("目前尚無儲存的歷史紀錄。點擊上方「💾 儲存目前排班至歷史紀錄」即可將排班封存檢視。")
    else:
        for h_idx, hist in enumerate(reversed(hist_list)):
            with st.expander(f"📜 [{hist['timestamp']}] {hist['title']}", expanded=False):
                st.markdown(f"**備註內容：** {hist['notes']}")
                h_display = []
                for idx_r, r_val in hist['df'].iterrows():
                    h_display.append({
                        "週次": r_val["週次"],
                        "星期一": f"日期: {r_val['一_0']}\n內容: {r_val['一_1']}\n老師: {r_val['一_2']}",
                        "星期二": f"日期: {r_val['二_0']}\n內容: {r_val['二_1']}\n老師: {r_val['二_2']}",
                        "星期三": f"日期: {r_val['三_0']}\n內容: {r_val['三_1']}\n老師: {r_val['三_2']}",
                        "星期四": f"日期: {r_val['四_0']}\n內容: {r_val['四_1']}\n老師: {r_val['四_2']}",
                        "星期五": f"日期: {r_val['五_0']}\n內容: {r_val['五_1']}\n老師: {r_val['五_2']}"
                    })
                st.dataframe(pd.DataFrame(h_display), use_container_width=True)
                
                h_doc_bytes = generate_cleaning_docx(branch_name, year_roc, month, hist['df'], hist['notes'])
                st.download_button(
                    label=f"📥 重新下載此歷史版本 ({hist['title']})",
                    data=h_doc_bytes,
                    file_name=f"{branch_name}_{hist['title']}_歷史存檔.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_hist_{prefix}_{h_idx}"
                )

with tab_a:
    render_branch_tab("澳森", "A")

with tab_b:
    render_branch_tab("澳森文德", "B")
