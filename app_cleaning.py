import streamlit as st
import pandas as pd
import datetime
import calendar
import holidays
import io

st.set_page_config(page_title="澳森托嬰中心 月清潔輪值與表單管理系統", layout="wide")

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
    "掃地機器人/寢具表填寫/規劃戶外活動計畫表與回報(髒水倒掉、洗淨、補充乾淨水、清潔液1:50)"
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
    "戶外掃落葉/寢具表填寫/規劃戶外活動計畫表與回報(髒水倒掉、洗淨、補充乾淨水、清潔液1:50)"
]

for prefix, tasks_init in [("A", DEFAULT_TASKS_A), ("B", DEFAULT_TASKS_B)]:
    if f"teachers_{prefix}" not in st.session_state:
        st.session_state[f"teachers_{prefix}"] = ["主任", "均宜", "小安", "綺綺", "嘉鳳", "樺樺", "Candy", "Panda"]
    if f"tasks_{prefix}" not in st.session_state:
        st.session_state[f"tasks_{prefix}"] = tasks_init
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

# --- Excel 檔案生成引擎 (自動套用格式、格線、自動換行與簽名列) ---
def generate_cleaning_excel(branch_name, year_roc, month, schedule_df, notes_text):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 將排班資料轉為乾淨的月曆表格結構
        excel_rows = []
        for idx, row in schedule_df.iterrows():
            w_name = row['週次']
            # 日期列
            excel_rows.append({
                "週次": w_name, "項目": "日期",
                "星期一": row['一_0'], "星期二": row['二_0'], "星期三": row['三_0'], "星期四": row['四_0'], "星期五": row['五_0']
            })
            # 清潔內容列
            excel_rows.append({
                "週次": w_name, "項目": "清潔內容",
                "星期一": row['一_1'], "星期二": row['二_1'], "星期三": row['三_1'], "星期四": row['四_1'], "星期五": row['五_1']
            })
            # 執行老師列
            excel_rows.append({
                "週次": w_name, "項目": "執行老師",
                "星期一": row['一_2'], "星期二": row['二_2'], "星期三": row['三_2'], "星期四": row['四_2'], "星期五": row['五_2']
            })
            # 簽名列
            excel_rows.append({
                "週次": w_name, "項目": "簽名",
                "星期一": "", "星期二": "", "星期三": "", "星期四": "", "星期五": ""
            })

        df_export = pd.DataFrame(excel_rows)
        df_export.to_excel(writer, sheet_name=f"{month}月清潔輪值", index=False, startrow=2)

        workbook = writer.book
        worksheet = writer.sheets[f"{month}月清潔輪值"]

        # 標題
        worksheet.merge_cells("A1:G1")
        title_cell = worksheet["A1"]
        title_cell.value = f"{branch_name} {year_roc} 年 {month} 月清潔輪值紀錄表"
        title_cell.font = openpyxl.styles.Font(name="微軟正黑體", size=16, bold=True, color="1F497D")
        title_cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")

        # 備註
        max_row = len(df_export) + 4
        worksheet.cell(row=max_row, column=1, value=notes_text).font = openpyxl.styles.Font(name="微軟正黑體", size=10, color="333333")

        # 設定欄寬與自動換行
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 15)

        for row in worksheet.iter_rows(min_row=3, max_row=len(df_export)+2, min_col=1, max_col=7):
            for cell in row:
                cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.font = openpyxl.styles.Font(name="微軟正黑體", size=10)

    output.seek(0)
    return output

st.title("🏫 澳森托嬰中心 月清潔輪值與表單管理系統")
st.markdown("支援 **澳森** 與 **澳森文德** 雙分園，自動避開國定假日，支援視窗內預覽與一鍵匯出 **Excel 表單**！")

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

    st.markdown("#### 🧹 清潔工作項目管理（小 Icon 卡片區）")
    task_list = st.session_state[f"tasks_{prefix}"]
    
    with st.form(key=f"add_task_form_{prefix}", clear_on_submit=True):
        new_task_name = st.text_input("➕ 新增清潔工作項目名稱與說明：")
        submitted = st.form_submit_button("新增至項目清單")
        if submitted and new_task_name.strip():
            st.session_state[f"tasks_{prefix}"].append(new_task_name.strip())
            st.rerun()

    updated_tasks = []
    for idx, t_item in enumerate(task_list):
        c_col1, c_col2, c_col3 = st.columns([0.05, 0.8, 0.15])
        with c_col1:
            st.markdown("🧹")
        with c_col2:
            edited_t = st.text_input(f"卡片_{idx}", value=t_item, key=f"task_card_{prefix}_{idx}", label_visibility="collapsed")
            updated_tasks.append(edited_t)
        with c_col3:
            if st.button("🗑️ 刪除", key=f"del_task_{prefix}_{idx}"):
                task_list.pop(idx)
                st.session_state[f"tasks_{prefix}"] = task_list
                st.rerun()
    st.session_state[f"tasks_{prefix}"] = updated_tasks

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
    current_tasks = st.session_state[f"tasks_{prefix}"] if st.session_state[f"tasks_{prefix}"] else ["一般清潔"]

    t_idx = 0
    for idx, w in enumerate(work_weeks[:4], start=1):
        week_name = f"第{['一','二','三','四'][idx-1]}週"
        with st.expander(f"📌 {week_name} 橫向月曆排班檢視", expanded=True):
            
            row_dates = []
            row_tasks = []
            row_teachers = []

            cols = st.columns(5)
            days_label = ['週一', '週二', '週三', '週四', '週五']

            for d_i, day_val in enumerate(w):
                with cols[d_i]:
                    if day_val > 0:
                        target_dt = datetime.date(year_ad, month, day_val)
                        adjusted_dt = get_adjusted_workday(target_dt, tw_holidays)
                        date_str = f"{month}/{day_val}"
                        if adjusted_dt != target_dt:
                            date_str += f" (調日至{adjusted_dt.month}/{adjusted_dt.day})"
                    else:
                        date_str = "-"

                    st.markdown(f"**{days_label[d_i]}**\n`{date_str}`")

                    default_task = ""
                    if day_val > 0:
                        if d_i == 1 and (idx in [2, 4]):
                            match_item = [t for t in current_tasks if "清點備品" in t]
                            default_task = match_item[0] if match_item else "清點備品"
                        elif d_i == 4:
                            match_item = [t for t in current_tasks if "掃地機器人" in t or "戶外掃落葉" in t or "規劃戶外活動" in t]
                            default_task = match_item[0] if match_item else current_tasks[-1]
                        else:
                            default_task = current_tasks[t_idx % len(current_tasks)]
                            t_idx += 1

                    chosen_task = st.selectbox(f"工作_{idx}_{d_i}", options=current_tasks, index=current_tasks.index(default_task) if default_task in current_tasks else 0, key=f"task_{branch_name}_{idx}_{d_i}", label_visibility="collapsed")
                    
                    default_teacher = teacher_pool[(idx * 5 + d_i) % len(teacher_pool)]
                    chosen_teacher = st.selectbox(f"老師_{idx}_{d_i}", teacher_options := (["主任"] + teacher_pool), index=teacher_options.index(default_teacher) if default_teacher in teacher_options else 0, key=f"tea_{branch_name}_{idx}_{d_i}", label_visibility="collapsed")
                
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
    st.subheader("👁️ 月曆式即時預覽（完整 Fit 於視窗內）")
    
    calendar_preview_rows = []
    for rec in table_records:
        calendar_preview_rows.append({
            "週次": rec["週次"],
            "星期一": f"{rec['日期'][0]}\n{rec['內容'][0]}\n(負責: {rec['老師'][0]})",
            "星期二": f"{rec['日期'][1]}\n{rec['內容'][1]}\n(負責: {rec['老師'][1]})",
            "星期三": f"{rec['日期'][2]}\n{rec['內容'][2]}\n(負責: {rec['老師'][2]})",
            "星期四": f"{rec['日期'][3]}\n{rec['內容'][3]}\n(負責: {rec['老師'][3]})",
            "星期五": f"{rec['日期'][4]}\n{rec['內容'][4]}\n(負責: {rec['老師'][4]})"
        })
    st.dataframe(pd.DataFrame(calendar_preview_rows), use_container_width=True, height=220)

    st.divider()
    import openpyxl
    st.subheader(f"📥 匯出 {branch_name} 清潔輪值 Excel 檔")
    excel_bytes = generate_cleaning_excel(branch_name, year_roc, month, final_df, st.session_state[f"notes_{prefix}"])

    st.download_button(
        label=f"📥 下載【{branch_name} {year_roc}年{month}月清潔輪值表】(Excel 檔)",
        data=excel_bytes,
        file_name=f"{branch_name}_{year_roc}年{month}月清潔輪值紀錄表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_btn_excel_{branch_name}"
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
                h_cal_rows = []
                for idx_r, r_val in hist['df'].iterrows():
                    days_k = ['一', '二', '三', '四', '五']
                    h_cal_rows.append({
                        "週次": r_val["週次"],
                        "星期一": f"{r_val['一_0']}\n{r_val['一_1']}\n(負責: {r_val['一_2']})",
                        "星期二": f"{r_val['二_0']}\n{r_val['二_1']}\n(負責: {r_val['二_2']})",
                        "星期三": f"{r_val['三_0']}\n{r_val['三_1']}\n(負責: {r_val['三_2']})",
                        "星期四": f"{r_val['四_0']}\n{r_val['四_1']}\n(負責: {r_val['四_2']})",
                        "星期五": f"{r_val['五_0']}\n{r_val['五_1']}\n(負責: {r_val['五_2']})"
                    })
                st.dataframe(pd.DataFrame(h_cal_rows), use_container_width=True, height=200)
                
                h_excel_bytes = generate_cleaning_excel(branch_name, year_roc, month, hist['df'], hist['notes'])
                st.download_button(
                    label=f"📥 重新下載此歷史版本 ({hist['title']})",
                    data=h_excel_bytes,
                    file_name=f"{branch_name}_{hist['title']}_歷史存檔.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_hist_excel_{prefix}_{h_idx}"
                )

with tab_a:
    render_branch_tab("澳森", "A")

with tab_b:
    render_branch_tab("澳森文德", "B")
