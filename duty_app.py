import streamlit as st
import pandas as pd
import io
import json
import os
import re
import tempfile
import hashlib
from collections import Counter
import xlsxwriter
from pathlib import Path
from openpyxl import load_workbook
from semesters import default_term, term_label, TermMismatchError
from logic import (
    parse_filename_for_name,
    parse_pdf_schedule,
    check_availability,
    SHIFTS, DAYS, DAYS_CN, SCHEDULE_PARSER_VERSION, MAX_WEEK, normalize_day,
    CAMPUS_NORTH, CAMPUS_SOUTH,
    ROLE_OFFICER, ROLE_CADRE, ROLE_MINISTER, ROLE_VICE_MINISTER, ROLE_DIRECTOR
)

# --- Configuration & Constants ---
st.set_page_config(page_title="智能值班表排班系统", layout="wide", page_icon="📅")

DUTY_TABLE_COLUMNS = ['Week', 'Day', 'Shift', 'Name', 'Class', 'Department', 'Role', 'Campus']
SHIFT_LABELS = dict(zip(SHIFTS, ['上午班 · 10:15–11:30', '下午一班 · 14:15–15:40', '下午二班 · 16:00–17:30']))
DUTY_IMPORT_COLUMN_MAP = {
    "周次": "Week",
    "星期": "Day",
    "班次": "Shift",
    "姓名": "Name",
    "班级": "Class",
    "专业班级": "Class",
    "部门": "Department",
    "职位": "Role",
    "职务": "Role",
    "校区": "Campus",
}


def normalize_week_values(value):
    if value is None:
        return set()

    if isinstance(value, str):
        return {int(item) for item in re.findall(r"\d+", value)}

    try:
        return {int(item) for item in value}
    except TypeError:
        try:
            return {int(value)}
        except (TypeError, ValueError):
            return set()


def normalize_week_number(value):
    if value is None or pd.isna(value):
        return None

    match = re.fullmatch(r"(?:第)?(\d{1,2})(?:\.0+)?(?:周)?", str(value).strip())
    if not match:
        return None
    week = int(match.group(1))
    return week if 1 <= week <= MAX_WEEK else None


def normalize_duty_table_dataframe(df):
    df = df.copy()
    df = df.rename(columns={col: DUTY_IMPORT_COLUMN_MAP.get(str(col).strip(), col) for col in df.columns})
    if df.empty:
        return pd.DataFrame(columns=DUTY_TABLE_COLUMNS)
    if df.columns.duplicated().any():
        raise ValueError('值班表存在重复表头，请保留每种字段的一列')
    required = {'Week', 'Day', 'Shift', 'Name'}
    missing = required - set(df.columns)
    if missing:
        raise ValueError('值班表缺少必填列：' + '、'.join(sorted(missing)))

    for col in DUTY_TABLE_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[DUTY_TABLE_COLUMNS].copy()
    for col in ["Day", "Shift", "Name", "Class", "Department", "Role", "Campus"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df = df[(df["Name"] != "") & (df["Name"] != "未安排")]
    weeks = df['Week'].apply(normalize_week_number)
    if weeks.isna().any():
        raise ValueError(f'值班表周次必须是 1-{MAX_WEEK} 内的单个整数，不能是负数、小数或周次范围')
    df['Week'] = weeks.astype(int)
    days = df['Day'].apply(normalize_day)
    if (days == '').any():
        raise ValueError('值班表有无法识别的星期：' + '、'.join(df.loc[days == '', 'Day'].unique()))
    df['Day'] = days
    shift_aliases = dict(zip(['上午班', '下午一班', '下午二班'], SHIFTS))
    df['Shift'] = df['Shift'].replace(shift_aliases)
    if (~df['Shift'].isin(SHIFTS)).any():
        raise ValueError('值班表有无法识别的班次：' + '、'.join(df.loc[~df['Shift'].isin(SHIFTS), 'Shift'].unique()))
    positions = ['Week', 'Day', 'Shift']
    if (df.groupby(positions)['Name'].nunique() > 1).any():
        raise ValueError('值班表同一周、星期、班次有多个不同人员；当前系统每个位置只支持一人，请先处理冲突')
    return df.drop_duplicates(subset=positions, keep='last')


def get_duty_weeks(duty_table):
    if duty_table is None or duty_table.empty or "Week" not in duty_table.columns:
        return set()

    weeks = duty_table["Week"].apply(normalize_week_number).dropna()
    return {int(week) for week in weeks}


def fill_duty_details_from_personnel(duty_table):
    if "personnel" not in st.session_state:
        return duty_table

    personnel = st.session_state["personnel"]
    if personnel.empty or "Name" not in personnel.columns:
        return duty_table

    detail_cols = ["Class", "Department", "Role", "Campus"]
    personnel_details = personnel.drop_duplicates("Name").set_index("Name")
    duty_table = duty_table.copy()

    for col in detail_cols:
        if col not in personnel_details.columns:
            continue
        missing = duty_table[col].fillna("").astype(str).str.strip() == ""
        duty_table.loc[missing, col] = duty_table.loc[missing, "Name"].map(personnel_details[col]).fillna("")

    return duty_table


def merge_imported_duty_table(import_df, replace_weeks=None):
    current = st.session_state.get("duty_table", pd.DataFrame(columns=DUTY_TABLE_COLUMNS))
    current = normalize_duty_table_dataframe(current)
    import_df = fill_duty_details_from_personnel(normalize_duty_table_dataframe(import_df))
    replace_weeks = set(replace_weeks or [])

    if replace_weeks:
        current = current[~current["Week"].isin(replace_weeks)]

    merged = pd.concat([current, import_df], ignore_index=True)
    merged = normalize_duty_table_dataframe(merged)
    st.session_state["duty_table"] = merged[DUTY_TABLE_COLUMNS]
    save_app_state()
    reset_planner_selections(get_duty_weeks(import_df) | replace_weeks)


def reset_planner_selections(weeks=None):
    for key in list(st.session_state.keys()):
        if key.startswith('sel_') and (weeks is None or int(key.split('_')[1]) in weeks):
            del st.session_state[key]


def reset_duty_import_widget(message):
    st.session_state["duty_import_message"] = message
    st.session_state["duty_import_uploader_version"] = st.session_state.get("duty_import_uploader_version", 0) + 1
    st.rerun()


def inspect_duty_table(duty_table):
    """Historical imports remain readable, with conflicts/unknown data clearly marked."""
    rows = []
    personnel = st.session_state.get('personnel', pd.DataFrame(columns=['Name']))
    schedules = st.session_state.get('schedules', {})
    for _, duty in duty_table.iterrows():
        matches = personnel[personnel['Name'] == duty['Name']]
        if matches.empty:
            free, code, reason = False, 'NO_PERSON', '人员不在当前名单中，无法校验'
        else:
            day = normalize_day(duty['Day'])
            free, code, reason = check_availability(
                matches.iloc[0], normalize_week_number(duty['Week']),
                DAYS_CN.index(day) if day else None, duty['Shift'], schedules,
            )
        if not free:
            rows.append({'周次': duty['Week'], '星期': duty['Day'], '班次': duty['Shift'],
                         '姓名': duty['Name'], '原因': reason, 'code': code})
    return pd.DataFrame(rows)


def build_duty_counts(personnel, duty_table):
    scheduled = duty_table['Name'].value_counts() if not duty_table.empty else pd.Series(dtype='int64')
    roster_names = personnel['Name'].tolist() if 'Name' in personnel else []
    names = list(dict.fromkeys(roster_names + scheduled.index.tolist()))
    return pd.DataFrame({'姓名': names, '已排班次数': [int(scheduled.get(name, 0)) for name in names]})


def render_duty_import_export_section():
    st.divider()
    st.subheader("导出与交接")

    if "duty_import_uploader_version" not in st.session_state:
        st.session_state["duty_import_uploader_version"] = 0

    if st.session_state.get("duty_import_message"):
        st.success(st.session_state.pop("duty_import_message"))

    if not st.session_state['duty_table'].empty:
        issues = inspect_duty_table(st.session_state['duty_table'])
        if not issues.empty:
            st.warning(f'现有排班有 {len(issues)} 条冲突或无法校验的记录，请检查后再使用或导出。')
            with st.expander('查看现有排班的冲突和未校验记录'):
                st.dataframe(issues.drop(columns=['code']), hide_index=True)

    col_ex1, col_ex2 = st.columns(2)

    with col_ex1:
        st.markdown("**导出Excel文件**")
        if st.session_state['duty_table'].empty:
            st.info("暂无已保存排班，保存后可导出 Excel。")
        else:
            export_df = build_duty_export_df(st.session_state['duty_table'])
            excel_bytes = build_duty_excel_bytes(export_df)
            default_export_name = default_duty_export_filename(st.session_state['duty_table'])

            st.download_button("下载当前学期值班表", excel_bytes, default_export_name,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               type='primary', on_click='ignore')
            with st.expander('本机保存到指定位置'):
                st.caption('保存窗口在运行程序的电脑上打开；通过其他电脑访问时请使用浏览器下载。')
                if st.button("💾 选择位置并保存"):
                    try:
                        export_path = choose_excel_save_path(default_export_name)
                        if export_path:
                            export_path.parent.mkdir(parents=True, exist_ok=True)
                            export_path.write_bytes(excel_bytes)
                            st.success(f"已保存到：{export_path}")
                        else:
                            st.info("已取消保存。")
                    except Exception as e:
                        st.error(f"保存失败：{e}")

    with col_ex2:
        st.markdown("**导入已有排班进行修改**")
        up_sched = st.file_uploader(
            "上传之前导出的值班表Excel",
            type=['xlsx'],
            key=f"duty_import_{st.session_state['active_term']}_{st.session_state['duty_import_uploader_version']}",
        )
        if up_sched:
            try:
                imp_df, has_term = read_duty_excel(up_sched)
                if not has_term:
                    st.warning('此旧版 Excel 没有学期标记，无法自动核实所属学期。')
                    if not st.checkbox(f"确认该文件属于 {term_label(st.session_state['active_term'])}",
                                       key=f"confirm_duty_term_{hashlib.sha256(up_sched.getvalue()).hexdigest()}"):
                        return
                import_weeks = get_duty_weeks(imp_df)
                existing_weeks = get_duty_weeks(st.session_state['duty_table'])
                conflict_weeks = sorted(import_weeks & existing_weeks)
                new_weeks = sorted(import_weeks - existing_weeks)

                if imp_df.empty or not import_weeks:
                    st.warning("导入文件中没有识别到有效排班记录，请检查表头和周次。")
                else:
                    st.write(f"导入文件包含周次：{', '.join(map(str, sorted(import_weeks)))}")

                    if conflict_weeks:
                        st.warning(
                            "系统中已存在这些周次的值班表："
                            + "、".join(map(str, conflict_weeks))
                            + "。请选择如何处理。"
                        )
                        replace_col, skip_col = st.columns(2)
                        with replace_col:
                            if st.button("替换已有周次并导入"):
                                merge_imported_duty_table(imp_df, replace_weeks=conflict_weeks)
                                reset_duty_import_widget(
                                    f"已导入 {len(imp_df)} 条记录，并替换第 "
                                    + "、".join(map(str, conflict_weeks))
                                    + " 周的原有值班表。"
                                )
                        with skip_col:
                            if new_weeks:
                                if st.button("只导入新周次"):
                                    new_df = imp_df[imp_df["Week"].isin(new_weeks)]
                                    merge_imported_duty_table(new_df)
                                    reset_duty_import_widget(
                                        f"已导入第 {'、'.join(map(str, new_weeks))} 周，已跳过冲突周次。"
                                    )
                            else:
                                st.info("导入文件全部周次都已存在，如需更新请使用替换导入。")
                    else:
                        if st.button('确认导入这些周次', type='primary'):
                            merge_imported_duty_table(imp_df)
                            reset_duty_import_widget(
                                f"导入成功，共导入 {len(imp_df)} 条记录，周次："
                                + "、".join(map(str, sorted(import_weeks)))
                                + "。"
                            )
            except Exception as e:
                st.error(f"导入失败: {e}")


def get_data_dir():
    base_dir = Path.cwd() / "data"
    try:
        base_dir.mkdir(parents=True, exist_ok=True)
    except OSError:
        base_dir = Path.home() / ".DutyScheduler" / "data"
        base_dir.mkdir(parents=True, exist_ok=True)
    return base_dir


def get_data_file():
    directory = get_data_dir() / 'terms'
    directory.mkdir(parents=True, exist_ok=True)
    return directory / f"{st.session_state['active_term']}.json"


def serialize_schedules(schedules):
    serializable = {}
    for name, slots in schedules.items():
        serializable[name] = []
        for slot in slots:
            item = dict(slot)
            item["weeks"] = sorted(normalize_week_values(item.get("weeks", [])))
            serializable[name].append(item)
    return serializable


def deserialize_schedules(schedules):
    restored = {}
    for name, slots in schedules.items():
        restored[name] = []
        for slot in slots:
            item = dict(slot)
            item["weeks"] = normalize_week_values(item.get("weeks", []))
            restored[name].append(item)
    return restored


def save_app_state():
    data_file = get_data_file()

    payload = {'academic_term': st.session_state['active_term']}
    if "personnel" in st.session_state:
        payload["personnel"] = st.session_state["personnel"].to_dict(orient="records")
    if "schedules" in st.session_state:
        payload["schedules"] = serialize_schedules(st.session_state["schedules"])
        payload["schedule_parser_version"] = SCHEDULE_PARSER_VERSION
    if "duty_table" in st.session_state:
        payload["duty_table"] = st.session_state["duty_table"].to_dict(orient="records")

    write_state_payload(data_file, payload)


def write_state_payload(data_file, payload):
    # Replace the whole file only after the complete new state has reached disk.
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=data_file.parent,
                                         prefix='app_state-', suffix='.tmp', delete=False) as temporary:
            temporary_path = Path(temporary.name)
            json.dump(payload, temporary, ensure_ascii=False, indent=2, default=str)
            temporary.flush()
            os.fsync(temporary.fileno())
        os.replace(temporary_path, data_file)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()


def load_app_state():
    data_file = get_data_file()
    if not data_file.exists():
        return

    try:
        payload = json.loads(data_file.read_text(encoding="utf-8"))
        if payload.get('academic_term') != st.session_state['active_term']:
            raise ValueError('状态文件的学期标记与当前学期不一致，未加载数据')
        if "personnel" in payload:
            st.session_state["personnel"] = pd.DataFrame(payload["personnel"])
        if "schedules" in payload:
            if payload.get("schedule_parser_version") == SCHEDULE_PARSER_VERSION:
                st.session_state["schedules"] = deserialize_schedules(payload["schedules"])
            else:
                st.session_state["schedules"] = {}
                if payload["schedules"]:
                    st.warning("旧版课表解析结果可能遗漏普通课程，请重新导入原始 PDF 课表后再排班。人员名单和值班记录已保留。")
        if "duty_table" in payload:
            duty_table = pd.DataFrame(payload["duty_table"])
            for col in DUTY_TABLE_COLUMNS:
                if col not in duty_table.columns:
                    duty_table[col] = pd.Series(dtype="object")
            st.session_state["duty_table"] = duty_table[DUTY_TABLE_COLUMNS]
    except Exception as e:
        st.warning(f"读取本机已保存数据失败：{e}")


def clear_saved_data_file():
    data_file = get_data_file()
    if data_file.exists():
        data_file.unlink()


def migrate_legacy_state():
    """Assign untagged data explicitly; preserve its exact bytes as a backup."""
    source = get_data_dir() / 'app_state.json'
    target = get_data_file()
    if target.exists():
        raise ValueError('当前学期已有保存数据，为避免覆盖，请选择一个空学期归入旧数据')
    payload = json.loads(source.read_text(encoding='utf-8'))
    payload['academic_term'] = st.session_state['active_term']
    backup = get_data_dir() / 'legacy_app_state.json'
    if backup.exists():
        raise ValueError('旧数据备份已存在，请先核对备份，未覆盖任何文件')
    write_state_payload(target, payload)
    source.rename(backup)
    return backup


def get_default_export_dir():
    export_dir = Path.cwd() / "值班表"
    export_dir.mkdir(parents=True, exist_ok=True)
    return export_dir


def sanitize_filename(filename):
    cleaned = str(filename or "").strip()
    for ch in '<>:"/\\|?*':
        cleaned = cleaned.replace(ch, "_")
    cleaned = cleaned.strip(" .")
    if not cleaned:
        cleaned = "值班表.xlsx"
    if not cleaned.lower().endswith(".xlsx"):
        cleaned += ".xlsx"
    return cleaned


def default_duty_export_filename(duty_table):
    prefix = st.session_state['active_term'] + '_'
    weeks = sorted(
        int(week)
        for week in pd.Series(duty_table.get("Week", [])).dropna().unique()
        if str(week).strip() != ""
    )
    if not weeks:
        return prefix + "值班表.xlsx"
    if len(weeks) == 1:
        return prefix + f"第{weeks[0]}周值班表.xlsx"
    return prefix + f"第{weeks[0]}-{weeks[-1]}周值班表.xlsx"


def build_duty_export_df(duty_table):
    return duty_table.rename(columns={
        "Week": "周次",
        "Day": "星期",
        "Shift": "班次",
        "Name": "姓名",
        "Class": "班级",
        "Department": "部门",
        "Role": "职位",
        "Campus": "校区",
    })


def build_duty_excel_bytes(export_df):
    out_buffer = io.BytesIO()
    with pd.ExcelWriter(out_buffer, engine="xlsxwriter") as writer:
        export_df.to_excel(writer, index=False, sheet_name="值班表")
        pd.DataFrame([{'学期标识': st.session_state['active_term'],
                       '学期': term_label(st.session_state['active_term'])}]).to_excel(
            writer, index=False, sheet_name='学期信息')
        worksheet = writer.sheets["值班表"]
        worksheet.freeze_panes(1, 0)
        widths = [8, 12, 24, 12, 24, 18, 12, 12]
        for idx, width in enumerate(widths[:len(export_df.columns)]):
            worksheet.set_column(idx, idx, width)
    return out_buffer.getvalue()


def read_duty_excel(file):
    with pd.ExcelFile(file) as workbook:
        has_term = '学期信息' in workbook.sheet_names
        if has_term:
            metadata = pd.read_excel(workbook, sheet_name='学期信息', dtype=str)
            if ('学期标识' not in metadata or len(metadata) != 1
                    or pd.isna(metadata.iloc[0]['学期标识'])):
                raise ValueError('Excel 学期信息不完整，未导入')
            actual = metadata.iloc[0]['学期标识']
            if actual != st.session_state['active_term']:
                raise TermMismatchError(f'Excel 属于 {actual}，当前选择 {st.session_state["active_term"]}，请切换学期后导入')
        sheet = '值班表' if '值班表' in workbook.sheet_names else workbook.sheet_names[0]
        return normalize_duty_table_dataframe(pd.read_excel(workbook, sheet_name=sheet)), has_term


def choose_excel_save_path(default_filename):
    import tkinter as tk
    from tkinter import filedialog

    default_dir = get_default_export_dir()
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.asksaveasfilename(
            parent=root,
            title="保存值班表",
            initialdir=str(default_dir),
            initialfile=sanitize_filename(default_filename),
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
    finally:
        root.destroy()

    if not selected:
        return None
    return Path(selected)


PERSONNEL_COLUMN_MAP = {
    "name": "Name",
    "姓名": "Name",
    "学生姓名": "Name",
    "class": "Class",
    "班级": "Class",
    "专业班级": "Class",
    "专业班": "Class",
    "major": "Major",
    "专业": "Major",
    "grade": "Grade",
    "年级": "Grade",
    "department": "Department",
    "部门": "Department",
    "role": "Role",
    "职位": "Role",
    "职务": "Role",
    "岗位": "Role",
    "角色": "Role",
    "campus": "Campus",
    "校区": "Campus",
    "序号": "Index",
}


def normalize_column_name(value):
    if value is None:
        return ""
    return str(value).replace("\n", "").replace(" ", "").strip()


def canonical_column_name(value):
    key = normalize_column_name(value)
    return PERSONNEL_COLUMN_MAP.get(key, PERSONNEL_COLUMN_MAP.get(key.lower(), key))


def read_excel_with_merged_cells(file_obj):
    wb = load_workbook(file_obj, data_only=True)
    ws = wb.active
    merged_values = {}

    for merged_range in ws.merged_cells.ranges:
        top_value = ws.cell(merged_range.min_row, merged_range.min_col).value
        for row in range(merged_range.min_row, merged_range.max_row + 1):
            for col in range(merged_range.min_col, merged_range.max_col + 1):
                merged_values[(row, col)] = top_value

    values = []
    for row in range(1, ws.max_row + 1):
        row_values = []
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row, col).value
            if value is None:
                value = merged_values.get((row, col))
            row_values.append(value)
        values.append(row_values)

    header_idx = None
    for idx, row in enumerate(values):
        mapped = [canonical_column_name(cell) for cell in row]
        hits = sum(col in {"Name", "Class", "Grade", "Department", "Role", "Campus"} for col in mapped)
        if "Name" in mapped and hits >= 2:
            header_idx = idx
            break

    if header_idx is None:
        return pd.DataFrame(values)

    columns = []
    seen = {}
    for idx, cell in enumerate(values[header_idx]):
        col = canonical_column_name(cell) or f"Unnamed_{idx + 1}"
        seen[col] = seen.get(col, 0) + 1
        if seen[col] > 1:
            if col in {'Name', 'Class', 'Grade', 'Department', 'Role', 'Campus'}:
                raise ValueError(f'人员名单存在重复表头：{col}，请保留一列')
            col = f"{col}_{seen[col]}"
        columns.append(col)

    return pd.DataFrame(values[header_idx + 1:], columns=columns)


def read_personnel_file(file_obj):
    file_name = getattr(file_obj, "name", "").lower()
    if file_name.endswith(".csv"):
        return pd.read_csv(file_obj)
    return read_excel_with_merged_cells(file_obj)


def normalize_campus(value):
    value = str(value).strip()
    if "南" in value and "北" not in value:
        return CAMPUS_SOUTH
    if "北" in value and "南" not in value:
        return CAMPUS_NORTH
    return value


def normalize_role(row):
    role = str(row.get("Role", "")).strip()
    department = str(row.get("Department", "")).strip()
    role_text = f"{role} {department}"
    if "主任团" in role_text:
        return ROLE_DIRECTOR
    if "副部" in role:
        return ROLE_VICE_MINISTER
    if "部长" in role:
        return ROLE_MINISTER
    if "干部" in role:
        return ROLE_CADRE
    if "干事" in role:
        return ROLE_OFFICER
    return role


def normalize_personnel_dataframe(df_p):
    df_p = df_p.copy()
    df_p = df_p.rename(columns={col: canonical_column_name(col) for col in df_p.columns})
    if df_p.columns.duplicated().any():
        raise ValueError('人员名单存在重复表头，请保留每种字段的一列')

    if "Name" in df_p.columns:
        df_p["Name"] = df_p["Name"].fillna("").astype(str).str.strip()
        df_p = df_p[df_p["Name"] != ""]
        if df_p.empty:
            raise ValueError('人员名单中没有有效姓名，不会覆盖已有名单')
        duplicate_names = df_p.loc[df_p['Name'].duplicated(keep=False), 'Name'].unique()
        if len(duplicate_names):
            raise ValueError('人员名单有重复或同名成员，当前系统以姓名关联课表，请先区分：' + '、'.join(duplicate_names))

    if "Class" not in df_p.columns:
        if "Major" in df_p.columns and "Grade" in df_p.columns:
            df_p["Class"] = (
                df_p["Major"].fillna("").astype(str).str.strip()
                + " "
                + df_p["Grade"].fillna("").astype(str).str.strip()
            ).str.strip()
        elif "Major" in df_p.columns:
            df_p["Class"] = df_p["Major"]
        elif "Grade" in df_p.columns:
            df_p["Class"] = df_p["Grade"]
        else:
            df_p["Class"] = ""

    if "Department" not in df_p.columns:
        df_p["Department"] = ""

    for col in ["Class", "Grade", "Department", "Role", "Campus"]:
        if col in df_p.columns:
            df_p[col] = df_p[col].fillna("").astype(str).str.strip()

    if "Department" in df_p.columns:
        df_p["Department"] = df_p["Department"].replace("", pd.NA).ffill().fillna("")

    if "Campus" in df_p.columns:
        df_p["Campus"] = df_p["Campus"].apply(normalize_campus)
    if "Role" in df_p.columns:
        df_p["Role"] = df_p.apply(normalize_role, axis=1)

    return df_p


def build_personnel_template():
    sample_data = pd.DataFrame({
        "序号": [1, 2, 3, 4],
        "姓名": ["张三", "李四", "王五", "赵六"],
        "部门": ["办公室", "活动部", "宣传部", "主任团"],
        "职务": ["干事", "干事", "干部", "干部"],
        "专业班级": ["会计学2501", "法学2502", "计算机科学与技术2401", "公共事业管理2301"],
        "年级": ["大一", "大一", "大二", "大三"],
        "校区": ["旗山北校区", "旗山南校区", "旗山北校区", "旗山北校区"],
    })
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        sample_data.to_excel(writer, index=False, sheet_name="人员名单")
        worksheet = writer.sheets["人员名单"]
        worksheet.freeze_panes(1, 0)
        widths = [8, 12, 18, 12, 26, 10, 16]
        for idx, width in enumerate(widths):
            worksheet.set_column(idx, idx, width)
    return buffer.getvalue()


def unique_filter_options(series):
    return sorted(
        value
        for value in series.fillna("").astype(str).str.strip().unique()
        if value
    )


def person_select_label(name, personnel_df):
    if name == "未安排":
        return name
    matches = personnel_df[personnel_df["Name"] == name]
    if matches.empty:
        return name

    person = matches.iloc[0]
    details = []
    for col in ["Role", "Grade", "Department"]:
        value = str(person.get(col, "")).strip()
        if value:
            details.append(value)

    if not details:
        return name
    return f"{name}（{'，'.join(details)}）"


def resolve_schedule_owner_name(file_name):
    fallback_name = parse_filename_for_name(file_name)

    personnel = st.session_state.get("personnel")
    if personnel is None or "Name" not in personnel.columns:
        raise ValueError('请先导入人员名单，再导入课表')

    names = (
        personnel["Name"]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: s != ""]
        .drop_duplicates()
        .tolist()
    )
    matched_names = [name for name in names if name in file_name]
    if not matched_names:
        raise ValueError(f'文件名未匹配到人员名单中的姓名：{fallback_name}')
    # A full name can legitimately contain a shorter person's name as a substring.
    full_matches = [name for name in matched_names
                    if not any(name != other and name in other for other in matched_names)]
    if len(full_matches) != 1:
        raise ValueError('文件名匹配到多位成员，无法确定课表归属：' + '、'.join(full_matches))
    return full_matches[0]


st.sidebar.header('🗓️ 学期设置')
initial_start, _, initial_semester = map(int, default_term().split('-'))
academic_start = st.sidebar.number_input('学年开始年份', min_value=2000, max_value=2100,
                                        value=initial_start, step=1, key='academic_start')
academic_semester = st.sidebar.selectbox('学期', [1, 2], index=initial_semester - 1,
                                        format_func=lambda value: f'第 {value} 学期', key='academic_semester')
selected_term = f'{academic_start}-{academic_start + 1}-{academic_semester}'
if st.session_state.get('active_term') != selected_term:
    # A new term must never inherit the previous term's data, selections or uploads.
    for state_key in list(st.session_state.keys()):
        if state_key not in {'academic_start', 'academic_semester'}:
            del st.session_state[state_key]
    st.session_state['active_term'] = selected_term
    load_app_state()

legacy_file = get_data_dir() / 'app_state.json'
if legacy_file.exists():
    with st.sidebar.expander('旧数据归属', expanded=not get_data_file().exists()):
        st.warning('发现未标记学期的旧数据，尚未用于当前排班。请先选择它实际所属的学年和学期，再归入；原文件会保留为备份。')
        if get_data_file().exists():
            st.caption('当前学期已有数据，为避免覆盖，归入按钮已停用。请先核对旧数据实际所属学期。')
        if st.button(f'将旧数据归入 {term_label(selected_term)}', disabled=get_data_file().exists()):
            try:
                migrate_legacy_state()
                load_app_state()
                st.rerun()
            except Exception as error:
                st.error(f'旧数据归入失败：{error}')


def parse_schedule_files(schedule_files):
    if not schedule_files:
        st.warning('没有选择课表文件。')
        return
    parsed = {}
    updated = dict(st.session_state.get('schedules', {}))
    logs = []
    total = len(schedule_files)
    bar = st.progress(0)
    status_text = st.empty()

    owners = []
    for f in schedule_files:
        file_name = getattr(f, "name", str(f))
        display_name = Path(file_name).name
        try:
            owners.append((f, display_name, resolve_schedule_owner_name(display_name), None))
        except ValueError as error:
            owners.append((f, display_name, None, str(error)))
    owner_counts = Counter(name for _, _, name, _ in owners if name is not None)

    for i, (f, display_name, name, owner_error) in enumerate(owners):
        try:
            status_text.text(f"正在处理: {display_name}...")
            if owner_error:
                raise ValueError(owner_error)
            if owner_counts[name] > 1:
                raise ValueError(f'本批次有多份课表匹配到 {name}，请仅保留当前学期的一份')
            slots = parse_pdf_schedule(f, expected_term=st.session_state['active_term'])
            parsed[name] = slots
            updated[name] = slots
            logs.append(f"✅ {name}: 成功提取 {len(slots)} 个课程时间段")
        except TermMismatchError as e:
            # An accidental import of another term must not delete a valid current PDF.
            logs.append(f"❌ {display_name}: 未导入 - {e}")
        except Exception as e:
            if name is not None:
                updated.pop(name, None)
            logs.append(f"❌ {display_name}: 解析失败 - {e}")
        bar.progress((i + 1) / total)

    status_text.text("处理完成！")
    st.session_state['schedules'] = updated
    save_app_state()
    if parsed:
        st.success(f"处理完成，共成功更新 {len(parsed)} 份课表，现有有效课表 {len(updated)} 份。")
    if len(parsed) < total:
        st.warning(f"有 {total - len(parsed)} 份课表解析失败，未采用这些文件，请查看详细处理日志。")

    if 'personnel' in st.session_state:
        personnel_names = st.session_state['personnel']['Name'].dropna().astype(str)
        all_names = set(personnel_names)
        parsed_names = set(updated.keys())
        missing = all_names - parsed_names
        missing = [str(m) for m in missing if m and str(m).strip()]

        if missing:
            st.warning(f"以下 {len(missing)} 位成员缺少有效课表（未导入或解析失败），不可值班：\n" + "、".join(sorted(missing)))
        else:
            st.success("✅ 完美！所有人都有课表。")

    st.session_state['schedule_import_logs'] = logs


def find_pdf_files(folder_path, recursive=False):
    if not folder_path or not folder_path.strip():
        raise ValueError("请先填写课表文件夹路径")

    folder = Path(folder_path.strip().strip('"').strip("'")).expanduser()
    if not folder.exists():
        raise FileNotFoundError(f"文件夹不存在：{folder}")
    if not folder.is_dir():
        raise NotADirectoryError(f"这不是文件夹：{folder}")

    entries = folder.rglob('*') if recursive else folder.iterdir()
    return sorted((p for p in entries if p.is_file() and p.suffix.lower() == '.pdf'), key=lambda p: str(p).lower())


# --- Main App ---

def render_state_summary(placeholder):
    with placeholder.container():
        personnel = st.session_state.get('personnel', pd.DataFrame())
        schedules = st.session_state.get('schedules', {})
        ready_count = sum(bool(schedules.get(person['Name'])) and person.get('Campus') in {CAMPUS_NORTH, CAMPUS_SOUTH}
                          for _, person in personnel.iterrows())
        summary_cols = st.columns(3)
        summary_cols[0].metric('本学期成员', f'{len(personnel)} 人')
        summary_cols[1].metric('数据就绪', f'{ready_count} 人', help='已有有效课表且校区明确；具体时段仍需排除课程冲突。')
        summary_cols[2].metric('已保存班次', f'{len(st.session_state.get("duty_table", []))} 班')

st.markdown('''<style>
[data-testid="stMainBlockContainer"] {padding-top: 2rem; padding-bottom: 2rem;}
[data-testid="stMain"] h1 {font-size: 1.85rem; padding-bottom: .35rem;}
[data-testid="stMain"] h2 {font-size: 1.3rem;}
[data-testid="stForm"] [data-testid="stVerticalBlock"] {gap: .45rem;}
[data-testid="stForm"] {background: #fafcfc;}
</style>''', unsafe_allow_html=True)
st.title("值班编排")
st.caption(f"{term_label(st.session_state['active_term'])} · 准备数据 → 查询空闲 → 编排值班 → 统计导出")
summary_placeholder = st.empty()
render_state_summary(summary_placeholder)

# Sidebar
st.sidebar.subheader("编排范围")
target_weeks = st.sidebar.multiselect("选择需要排班的周次", list(range(1, MAX_WEEK + 1)), default=[1])
st.sidebar.caption("选择周次后，在排班页切换编辑。切换学期或离开页面前请先保存排班。")

# Tabs
tab1, tab2, tab3, tab4 = st.tabs(["数据准备", "空闲查询", "排班制作", "统计与导出"])

# --- Tab 1: Data Upload ---
with tab1:
    st.subheader("准备本学期数据")
    st.caption("先导入人员名单，再读取课表。补传课表只更新对应成员。")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**1 · 人员名单**")
        st.caption("必填：姓名、职务、校区；支持 Excel／CSV 和中文／英文表头。")

        # Template
        st.download_button("下载人员名单模板", build_personnel_template(), "人员名单模板.xlsx", on_click='ignore')

        p_file = st.file_uploader("点击上传人员名单 (Excel/CSV)", type=['xlsx', 'csv'],
                                  key=f"personnel_upload_{selected_term}")
        if p_file and st.button('导入人员名单', key='import_personnel'):
            try:
                df_p = normalize_personnel_dataframe(read_personnel_file(p_file))

                # Keep only relevant columns if possible to avoid clutter, or just ensure existence
                # We need Name, Class, Department, Role, Campus

                # Validation: Check for required columns
                missing_cols = []
                if 'Name' not in df_p.columns: missing_cols.append('Name')
                if 'Role' not in df_p.columns: missing_cols.append('Role')
                if 'Campus' not in df_p.columns: missing_cols.append('Campus')

                if missing_cols:
                    st.error(f"❌ 上传的文件缺少关键列: {', '.join(missing_cols)}。请检查表头是否正确（支持中文：姓名、职务/职位、校区）。")
                else:
                    keep_cols = ['Name', 'Class', 'Grade', 'Department', 'Role', 'Campus']
                    keep_cols = [col for col in keep_cols if col in df_p.columns]
                    df_p = df_p[keep_cols]
                    st.session_state['personnel'] = df_p
                    save_app_state()
                    st.success(f"✅ 成功导入 {len(df_p)} 名人员信息！")
                    departments = unique_filter_options(df_p['Department'])
                    if departments:
                        st.info("已识别部门：" + "、".join(departments))
                    with st.expander("查看导入的人员列表"):
                        st.dataframe(df_p)
            except Exception as e:
                st.error(f"读取文件失败: {e}")

    with col2:
        st.markdown("**2 · 课表 PDF**")
        st.caption("文件名包含名单中的准确姓名；正文须带学期、节次及周次，且能选择文字。")
        can_parse = 'personnel' in st.session_state and not st.session_state['personnel'].empty
        if not can_parse:
            st.info('先完成左侧人员名单导入，再解析课表。')

        import_mode = st.radio(
            "课表导入方式",
            ["选择多个PDF文件", "选择课表文件夹", "填写本机文件夹路径"],
            horizontal=True,
        )

        if import_mode == "选择多个PDF文件":
            pdf_files = st.file_uploader(
                "点击选择PDF课表文件 (支持按住Ctrl/Command键多选)",
                type=['pdf'],
                accept_multiple_files=True,
                key=f"pdf_upload_{selected_term}",
            )

            if pdf_files:
                st.write(f"已选择 {len(pdf_files)} 个文件")
                if st.button("🚀 开始解析课表", key="parse_pdf_files", disabled=not can_parse):
                    parse_schedule_files(pdf_files)

        elif import_mode == "选择课表文件夹":
            folder_files = st.file_uploader(
                "点击选择课表文件夹",
                type=['pdf'],
                accept_multiple_files="directory",
                key=f"folder_upload_{selected_term}",
            )

            if folder_files:
                st.write(f"已从文件夹中选择 {len(folder_files)} 个PDF文件")
                if st.button("🚀 开始解析文件夹课表", key="parse_pdf_directory_upload", disabled=not can_parse):
                    parse_schedule_files(folder_files)

        else:
            folder_path = st.text_input("课表文件夹路径", placeholder=r"例如：D:\课表\2025-2026-1")
            include_subfolders = st.checkbox("同时读取子文件夹中的PDF", value=False)

            if st.button("📁 扫描并解析文件夹", key="parse_pdf_directory_path", disabled=not can_parse):
                try:
                    pdf_paths = find_pdf_files(folder_path, recursive=include_subfolders)
                    if not pdf_paths:
                        st.warning("这个文件夹里没有找到 PDF 文件。")
                    else:
                        st.write(f"已找到 {len(pdf_paths)} 个PDF文件")
                        parse_schedule_files(pdf_paths)
                except Exception as e:
                    st.error(f"读取文件夹失败：{e}")
        if st.session_state.get('schedule_import_logs'):
            with st.expander('最近一次课表处理日志'):
                for log in st.session_state['schedule_import_logs']:
                    st.write(log)

    st.divider()
    st.subheader("待处理的数据")

    personnel_count = len(st.session_state["personnel"]) if "personnel" in st.session_state else 0
    schedule_count = len(st.session_state["schedules"]) if "schedules" in st.session_state else 0
    duty_count = len(st.session_state["duty_table"]) if "duty_table" in st.session_state else 0

    if personnel_count:
        personnel = st.session_state['personnel']
        unknown_campus = []
        if 'Campus' in personnel.columns:
            unknown_campus = personnel.loc[~personnel['Campus'].isin([CAMPUS_NORTH, CAMPUS_SOUTH]), 'Name'].tolist()
            if unknown_campus:
                st.error('以下成员校区不明确，暂不参与排班：' + '、'.join(unknown_campus)
                         + '。请补全南校区或北校区后重新导入人员名单；已解析课表会保留。')
        missing_names = set(personnel['Name']) - set(st.session_state.get('schedules', {}))
        if missing_names:
            st.warning('以下成员缺少有效课表，暂不参与排班：' + '、'.join(sorted(missing_names)))
        if not unknown_campus and not missing_names:
            st.success('当前名单所有成员已具备有效课表和明确校区，可以开始查询与编排。')
        with st.expander('查看当前人员名单'):
            st.dataframe(personnel.rename(columns={'Name':'姓名','Class':'班级','Grade':'年级',
                         'Department':'部门','Role':'职务','Campus':'校区'}), hide_index=True)
    else:
        st.caption('尚未导入人员名单。')

    with st.expander('数据管理：删除课表与清空数据'):
        st.caption(f"仅操作当前学期。保存位置：{get_data_file()}")
        if schedule_count:
            delete_name = st.selectbox("删除某一份课表", sorted(st.session_state["schedules"].keys()))
            if st.button("🗑️ 删除选中课表"):
                st.session_state["schedules"].pop(delete_name, None)
                save_app_state()
                reset_planner_selections()
                st.rerun()
        c_clear1, c_clear2, c_clear3, c_clear4 = st.columns(4)
        clear_key = None
        if c_clear1.button("清空人员名单"):
            clear_key = 'personnel'
        if c_clear2.button("清空全部课表"):
            clear_key = 'schedules'
        if c_clear3.button("清空排班记录"):
            clear_key = 'duty_table'
        if clear_key:
            st.session_state.pop(clear_key, None)
            save_app_state()
            reset_planner_selections()
            st.rerun()
        if c_clear4.button("清空所有数据"):
            for key in ["personnel", "schedules", "duty_table", 'schedule_import_logs']:
                st.session_state.pop(key, None)
            clear_saved_data_file()
            reset_planner_selections()
            st.rerun()

# --- Tab 2: Search & Filter ---
with tab2:
    st.subheader("查询谁有空")
    st.caption("选择时间后查询。可按职务、校区、部门缩小范围，结果会说明不可值班的原因。")

    if 'personnel' in st.session_state and 'schedules' in st.session_state:
        c1, c2, c3 = st.columns(3)
        with c1:
            s_week = st.number_input("选择周次", 1, MAX_WEEK, 1, key='query_week')
        with c2:
            s_day_cn = st.selectbox("选择星期", DAYS_CN)
            s_day_idx = DAYS_CN.index(s_day_cn)
        with c3:
            # Create a display map for shifts
            s_shift = st.selectbox("选择班次", list(SHIFTS.keys()), format_func=SHIFT_LABELS.get)

        # Filters
        st.markdown("**筛选条件（可选）**")
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            f_role = st.multiselect("按职位筛选", unique_filter_options(st.session_state['personnel']['Role']), key='query_role')
        with f_col2:
            f_campus = st.multiselect("按校区筛选", unique_filter_options(st.session_state['personnel']['Campus']), key='query_campus')
        with f_col3:
            f_dept = st.multiselect("按部门筛选", unique_filter_options(st.session_state['personnel']['Department']), key='query_dept')

        if st.button("🔎 开始查询"):
            results = []
            excluded = []
            df = st.session_state['personnel']

            # Apply Pre-filters
            if f_role: df = df[df['Role'].isin(f_role)]
            if f_campus: df = df[df['Campus'].isin(f_campus)]
            if f_dept: df = df[df['Department'].isin(f_dept)]

            for _, person in df.iterrows():
                is_free, code, reason = check_availability(
                    person, s_week, s_day_idx, s_shift, st.session_state['schedules']
                )
                if is_free:
                    results.append({
                        "姓名": person['Name'],
                        "班级": person['Class'],
                        "部门": person['Department'],
                        "职位": person['Role'],
                        "校区": person['Campus'],
                        "状态": "✅ 可值班"
                    })
                else:
                    excluded.append({'姓名': person['Name'], '原因': reason})

            if results:
                st.success(f"共找到 {len(results)} 位符合条件且有空的同学！")
                st.dataframe(pd.DataFrame(results))
            else:
                st.warning("⚠️ 该时间段没有符合条件的空闲人员。")
            if excluded:
                with st.expander('查看不可值班人员及原因'):
                    st.dataframe(pd.DataFrame(excluded), hide_index=True)
    else:
        st.info("请先在「数据导入」页面上传人员名单和课表。")

# --- Tab 3: Scheduler ---
with tab3:
    st.subheader("编排一周值班")

    if 'personnel' in st.session_state and 'schedules' in st.session_state and target_weeks:
        # State for Duty Table
        if 'duty_table' not in st.session_state:
            st.session_state['duty_table'] = pd.DataFrame(columns=DUTY_TABLE_COLUMNS)

        # Select Week
        sch_week = st.selectbox("正在制作第几周的班表？", target_weeks, key="sch_week")

        st.caption(f"第 {sch_week} 周 · 自动排除有课、实践周及通勤冲突；每班一人，修改后点击下方保存。")

        with st.form("scheduler_form"):
            changes = []

            headings = st.columns([.6, 2, 2, 2])
            headings[0].markdown('**星期**')
            for heading, shift_label in zip(headings[1:], SHIFT_LABELS.values()):
                heading.markdown(f'**{shift_label}**')
            for day_i, day_cn in enumerate(DAYS_CN):
                row_cols = st.columns([.6, 2, 2, 2], vertical_alignment='center')
                row_cols[0].markdown(day_cn)
                cols = row_cols[1:]

                for shift_i, (shift_name, conf) in enumerate(SHIFTS.items()):
                    # Find current assignment
                    current = st.session_state['duty_table'][
                        (st.session_state['duty_table']['Week'] == sch_week) &
                        (st.session_state['duty_table']['Day'] == day_cn) &
                        (st.session_state['duty_table']['Shift'] == shift_name)
                    ]
                    current_name = current.iloc[0]['Name'] if not current.empty else "未安排"

                    # Find candidates
                    candidates = ["未安排"]

                    # Optimization: Filter personnel once? No, depends on slot.
                    for _, p in st.session_state['personnel'].iterrows():
                        is_free, _, _ = check_availability(p, sch_week, day_i, shift_name, st.session_state['schedules'])
                        if is_free:
                            candidates.append(p['Name'])

                    # Ensure current is in list
                    current_unavailable = current_name not in candidates
                    if current_unavailable:
                        candidates.append(current_name)

                    with cols[shift_i]:
                        # Improve label for readability
                        shift_label = SHIFT_LABELS[shift_name]

                        if current_unavailable:
                            st.warning(f"原排班中的 {current_name} 当前不可值班，请重新选择。")

                        sel = st.selectbox(
                            f"{day_cn} {shift_label}",
                            candidates,
                            index=candidates.index(current_name),
                            format_func=lambda name, personnel_df=st.session_state['personnel']: person_select_label(name, personnel_df),
                            key=f"sel_{sch_week}_{day_i}_{shift_i}",
                            label_visibility='collapsed',
                        )
                        changes.append({
                            "Week": sch_week,
                            "Day": day_cn,
                            "Shift": shift_name,
                            "Name": sel
                        })

            if st.form_submit_button("💾 保存当前周排班", type='primary'):
                conflicts = []
                for ch in changes:
                    if ch['Name'] == "未安排":
                        continue
                    people = st.session_state['personnel'][
                        st.session_state['personnel']['Name'] == ch['Name']
                    ]
                    if people.empty or not check_availability(
                        people.iloc[0], sch_week, DAYS_CN.index(ch['Day']),
                        ch['Shift'], st.session_state['schedules']
                    )[0]:
                        conflicts.append(f"{ch['Day']} {ch['Shift']}：{ch['Name']}")
                if conflicts:
                    st.error("以下安排存在有课、通勤或缺少有效课表的冲突，请调整后再保存：" + "；".join(conflicts))
                    st.stop()
                # Update State
                # Remove old entries for this week
                base_df = st.session_state['duty_table']
                base_df = base_df[base_df['Week'] != sch_week]

                new_rows = []
                for ch in changes:
                    if ch['Name'] != "未安排":
                        # Lookup details
                        p_details = st.session_state['personnel'][
                            st.session_state['personnel']['Name'] == ch['Name']
                        ].iloc[0]
                        new_rows.append({
                            "Week": ch['Week'],
                            "Day": ch['Day'],
                            "Shift": ch['Shift'],
                            "Name": ch['Name'],
                            "Class": p_details['Class'],
                            "Department": p_details['Department'],
                            "Role": p_details['Role'],
                            "Campus": p_details['Campus']
                        })

                st.session_state['duty_table'] = pd.concat([base_df, pd.DataFrame(new_rows)], ignore_index=True)
                save_app_state()
                st.success("✅ 保存成功！")
    else:
        if 'personnel' in st.session_state and 'schedules' in st.session_state:
            st.info('请先在左侧选择至少一个需要排班的周次。')
        else:
            st.info("请先在「数据导入」页面上传人员名单和课表后再手动制作排班。也可以先在下方导入已有值班表查看统计。")

    if 'duty_table' not in st.session_state:
        st.session_state['duty_table'] = pd.DataFrame(columns=DUTY_TABLE_COLUMNS)
    st.caption('保存后，到「统计与导出」下载值班表或导入交接文件。')

# --- Tab 4: Records ---
with tab4:
    st.subheader("当前学期记录")
    render_duty_import_export_section()
    st.caption('统计当前学期的已保存排班：一个班次记 1 次，同一天两个班次记 2 次；这是安排次数，尚未记录实际出勤或请假。')
    df = st.session_state.get('duty_table', pd.DataFrame(columns=DUTY_TABLE_COLUMNS))
    counts = build_duty_counts(st.session_state.get('personnel', pd.DataFrame()), df)
    if not counts.empty:
        st.subheader('当前学期已排班次数（含 0 次成员）')
        st.dataframe(counts, hide_index=True, height=280)
        with st.expander('查看次数分布图'):
            st.bar_chart(counts.set_index('姓名'))

    if 'duty_table' in st.session_state and not st.session_state['duty_table'].empty:
        # Detailed Log
        st.subheader("详细值班记录")
        search = st.text_input("🔍 搜索姓名")

        display_df = df.rename(columns={
            "Week": "周次", "Day": "星期", "Shift": "班次",
            "Name": "姓名", "Class": "班级", "Department": "部门", "Role": "职位", "Campus": "校区"
        })

        if search:
            st.dataframe(display_df[display_df['姓名'].str.contains(search, regex=False, na=False)])
        else:
            st.dataframe(display_df)
    else:
        st.info("暂无值班记录。请先在「排班制作」页面生成排班。")

render_state_summary(summary_placeholder)
