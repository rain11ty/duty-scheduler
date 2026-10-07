import pdfplumber
import os
import re

# Constants needed for logic
CAMPUS_NORTH = "北校区"
CAMPUS_SOUTH = "南校区"

# Roles
ROLE_OFFICER = "干事"
ROLE_CADRE = "干部"
ROLE_MINISTER = "部长"
ROLE_VICE_MINISTER = "副部长"
ROLE_DIRECTOR = "主任团"

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
DAYS_CN = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

SHIFTS = {
    "Morning (10:15-11:30)": {
        "periods": [3, 4],
        "next_periods": [5, 6],
        "type": "morning"
    },
    "Afternoon 1 (14:15-15:40)": {
        "periods": [5, 6],
        "next_periods": [7, 8],
        "type": "afternoon1"
    },
    "Afternoon 2 (16:00-17:30)": {
        "periods": [7, 8],
        "next_periods": [],
        "type": "afternoon2"
    }
}

ALL_PERIODS = list(range(1, 12))
MAX_WEEK = 20
DAY_TO_INDEX = {day: index for index, day in enumerate(DAYS_CN)}
DAY_TO_INDEX.update({f'周{suffix}': index for index, suffix in enumerate('一二三四五六日')})
DAY_TO_INDEX.update({'星期天': 6, '周天': 6})
SCHEDULE_PARSER_VERSION = 4
PERIOD_PATTERN = re.compile(r'\((\d{1,2})(?:[-－—](\d{1,2}))?节\)')


def normalize_schedule_text(text):
    return re.sub(r'\s+', '', str(text or '')
                  .replace('（', '(').replace('）', ')'))


def normalize_day(value):
    index = DAY_TO_INDEX.get(normalize_schedule_text(value))
    return DAYS_CN[index] if index is not None else ''


def parse_grid_schedule_tables(tables):
    """Reassemble weekday columns before parsing courses split across PDF pages."""
    day_columns = {}
    column_count = None
    day_texts = {day: [] for day in range(7)}
    for table in tables:
        for row in table:
            cells = [normalize_schedule_text(cell) for cell in row]
            header = {DAY_TO_INDEX[cell]: col for col, cell in enumerate(cells)
                      if cell in DAY_TO_INDEX}
            if len(header) >= 5:
                day_columns = header
                column_count = len(cells)
                continue
            if not day_columns or len(cells) != column_count:
                continue
            # Ignore titles/footers; an empty period cell can be a page continuation.
            if cells[1] and not (cells[1].isdigit() and int(cells[1]) in ALL_PERIODS):
                continue
            for day, col in day_columns.items():
                day_texts[day].append(cells[col])

    busy_slots = []
    for day, chunks in day_texts.items():
        content = ''.join(chunks)
        markers = list(PERIOD_PATTERN.finditer(content))
        if content and not markers:
            raise ValueError('网格课表只有课程名称等信息，缺少节次或周次详情，请重新导出完整课表')
        for i, marker in enumerate(markers):
            end = markers[i + 1].start() if i + 1 < len(markers) else len(content)
            course_text = content[marker.start():end]
            # Weeks belong to this course only, before its campus/location fields.
            weeks = parse_weeks_from_text(content[marker.end():end].split('/')[0])
            periods = parse_period_range(f'{marker.group(1)}-{marker.group(2) or marker.group(1)}')
            if not weeks or any(period not in ALL_PERIODS for period in periods):
                raise ValueError('网格课表中存在无法识别的课程周次或节次，请检查 PDF')
            campus = CAMPUS_SOUTH if '南校区' in course_text else CAMPUS_NORTH
            for period in periods:
                busy_slots.append({
                    'day': day, 'period': period, 'weeks': weeks,
                    'campus': campus, 'raw': course_text, 'type': 'class',
                })
    return busy_slots


def parse_practice_weeks_from_text(text):
    if not text:
        return set()

    normalized = (
        str(text)
        .replace("（", "(")
        .replace("）", ")")
        .replace("，", ",")
        .replace("、", ",")
        .replace("\n", "")
    )

    weeks = set()
    # Practice rows usually look like: 课程名(共2周)/18-19周/无;
    # Only parse the slash-delimited week field, not "(共2周)".
    for segment in re.findall(r'/([^/;]*周[^/;]*)/', normalized):
        weeks.update(parse_weeks_from_text(segment))
    return weeks


def add_practice_week_slots(busy_slots, text):
    if not text or "实践课程" not in text:
        return

    normalized = normalize_schedule_text(text).replace('，', ',').replace('；', ';')

    for match in re.finditer(r'实践课程[:：](.*?)(?:其他课程[:：]|打印时间|[:：]理论|$)', normalized):
        practice_text = match.group(1).strip()
        weeks = set()
        for entry in practice_text.split(';'):
            if not entry or entry.strip('。 ') in {'无', '暂无', '无实践课程'}:
                continue
            entry_weeks = parse_practice_weeks_from_text(entry)
            if not entry_weeks:
                raise ValueError(f'实践课程缺少可识别的周次：{entry[:60]}，请检查 PDF')
            weeks.update(entry_weeks)
        if not weeks:
            continue

        for day_idx in range(len(DAYS_CN)):
            for period in ALL_PERIODS:
                busy_slots.append({
                    'day': day_idx,
                    'period': period,
                    'weeks': weeks,
                    'campus': CAMPUS_NORTH,
                    'raw': f"实践课程：{practice_text}",
                    'type': 'practice'
                })


def parse_period_range(value):
    match = re.fullmatch(r'\s*(\d{1,2})(?:\s*[-－—]\s*(\d{1,2}))?\s*', str(value or ""))
    if not match:
        return None
    start, end = int(match.group(1)), int(match.group(2) or match.group(1))
    if start > end:
        start, end = end, start
    return list(range(start, end + 1))


def parse_list_schedule_table(table):
    busy_slots = []
    if not table:
        return busy_slots

    current_day = None
    current_periods = None

    for row in table:
        cells = [(str(cell).strip() if cell is not None else "") for cell in row]
        # The supported teaching-system list has weekday, periods, course info.
        # Grid rows must not inherit list context when formats are mixed.
        if len(cells) != 3:
            current_day = current_periods = None
            continue

        day_text, period_text, content = normalize_schedule_text(cells[0]), cells[1], cells[2]
        if day_text == '星期' and normalize_schedule_text(period_text) == '节次':
            continue
        if day_text in DAY_TO_INDEX:
            new_day = DAY_TO_INDEX[day_text]
            if new_day != current_day:
                current_periods = None
            current_day = new_day
        elif day_text:
            if content and (current_day is not None or '周' in content or day_text.startswith('星期')):
                raise ValueError(f'列表课表中的星期无法识别：{day_text}')
            current_day = current_periods = None
            continue

        parsed_periods = parse_period_range(period_text)
        if parsed_periods:
            current_periods = parsed_periods
        elif period_text and current_day is not None and content:
            raise ValueError(f'列表课表中的节次无法识别：{period_text}')

        if current_day is None or not current_periods:
            if content and (current_day is not None or parsed_periods or '周' in content):
                raise ValueError('列表课表中的课程缺少可识别的星期或节次')
            continue
        if not content:
            continue

        weeks = parse_weeks_from_text(content)
        if not weeks:
            raise ValueError('列表课表中存在无法识别的课程周次，请检查 PDF')
        if any(period not in ALL_PERIODS for period in current_periods):
            raise ValueError('列表课表中存在无法识别的课程节次，请检查 PDF')

        campus = CAMPUS_NORTH
        normalized_content = normalize_schedule_text(content)
        if "南校区" in normalized_content:
            campus = CAMPUS_SOUTH
        elif "北校区" in normalized_content:
            campus = CAMPUS_NORTH

        for period in current_periods:
            busy_slots.append({
                'day': current_day,
                'period': period,
                'weeks': weeks,
                'campus': campus,
                'raw': content,
                'type': 'class'
            })

    return busy_slots

def parse_filename_for_name(filename):
    """Extracts student name from filename like '张三(2023-2024-2)课表.pdf'"""
    base = os.path.basename(filename)
    name = re.split(r'[\(\[\.]', base)[0]
    return name.strip()

def parse_weeks_from_text(text):
    """
    Parses week ranges from text like "1-16周" or "1-8,10-16(双)周"
    Returns a set of integers.
    """
    weeks = set()
    if not text:
        return weeks

    normalized = (
        str(text)
        .replace("（", "(")
        .replace("）", ")")
        .replace("，", ",")
        .replace("、", ",")
        .replace("\n", "")
    )
    normalized = re.sub(r'\s+', '', normalized)
    normalized = normalized.replace('－', '-').replace('—', '-')

    # Avoid treating class periods such as "(1-2节)" as week ranges.
    clean_text = re.sub(r'\(\d{1,2}-\d{1,2}节\)', '', normalized)
    clean_text = re.sub(r'\d{1,2}-\d{1,2}节', '', clean_text)

    def add_range(start, end, parity=None):
        start, end = int(start), int(end or start)
        if not (1 <= start <= MAX_WEEK and 1 <= end <= MAX_WEEK):
            raise ValueError(f'课程周次超出支持范围 1-{MAX_WEEK}：{start}-{end}')
        if start > end:
            start, end = end, start
        for week in range(start, end + 1):
            if parity == "单" and week % 2 == 0:
                continue
            if parity == "双" and week % 2 != 0:
                continue
            weeks.add(week)

    # Handles both "4-6周(双)" and "1-8,10-16(双)周".
    week_expr = re.compile(
        r'(?<![\d-])(?P<body>\d{1,2}(?:-\d{1,2})?(?:,\d{1,2}(?:-\d{1,2})?)*)'
        r'(?P<pre_parity>\([单双]\))?\s*周'
        r'(?P<post_parity>\([单双]\))?'
    )
    for match in week_expr.finditer(clean_text):
        parity_text = match.group("pre_parity") or match.group("post_parity") or ""
        parity = "单" if "单" in parity_text else "双" if "双" in parity_text else None
        for token in re.finditer(r'(\d{1,2})(?:-(\d{1,2}))?', match.group("body")):
            add_range(token.group(1), token.group(2), parity)

    return weeks

def parse_pdf_schedule(pdf_file, expected_term=None):
    """
    Parses a PDF schedule.
    Returns busy slots for grid/list PDFs; raises ValueError for unreadable schedules.
    """
    busy_slots = []
    all_tables = []
    page_texts = []

    with pdfplumber.open(pdf_file) as pdf:
        for page_number, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            page_texts.append(text)
            tables = page.extract_tables()
            if not text.strip() and (page.images or tables or page.curves or page.lines or page.rects):
                if page.images and not page.chars:
                    raise ValueError(
                        f'第 {page_number} 页包含图片但没有文字层（图片版 PDF），当前不支持图片文字识别（OCR）；'
                        '请从教务系统重新导出能选择、复制文字的 PDF 课表')
                raise ValueError(f'第 {page_number} 页没有可提取文字，不能确认课程是否完整，请重新导出文字版课表')
            all_tables.extend(tables)
    full_text = '\n'.join(page_texts)
    if expected_term is not None:
        from semesters import validate_pdf_term
        validate_pdf_term(full_text, expected_term)
    # Repeated list headers and blank merged cells can continue on another page.
    busy_slots.extend(parse_list_schedule_table([row for table in all_tables for row in table]))
    add_practice_week_slots(busy_slots, full_text)
    busy_slots.extend(parse_grid_schedule_tables(all_tables))
    if not any(slot.get('type') == 'class' for slot in busy_slots) and PERIOD_PATTERN.search(
            normalize_schedule_text(full_text)):
        raise ValueError('PDF 中有普通课程，但未能识别课程表格，请重新导出课表')
    if not busy_slots:
        raise ValueError('未识别到课程时间；PDF 可能是图片、空白或不支持的格式，请导出文字版课表')
    return busy_slots

def check_availability(person, week, day_idx, shift_name, all_schedules):
    """
    Returns: (is_available, reason_code, debug_info)
    """
    if week not in range(1, MAX_WEEK + 1) or day_idx not in range(7) or shift_name not in SHIFTS:
        return False, 'INVALID_SELECTION', f'请选择 1-{MAX_WEEK} 周、有效星期和班次'
    name = person.get('Name', '')
    person_campus = person.get('Campus', '')

    # Check if schedule exists for this person
    if name not in all_schedules:
        return False, "NO_SCHEDULE", "未导入有效课表，或课表解析失败"

    schedule = all_schedules.get(name, [])
    if not schedule:
        return False, "INVALID_SCHEDULE", "课表没有有效课程时间，请重新导入"
    if person_campus not in {CAMPUS_NORTH, CAMPUS_SOUTH}:
        return False, "UNKNOWN_CAMPUS", "校区为空或无法识别，请先确认南北校区"

    shift_conf = SHIFTS[shift_name]
    shift_periods = shift_conf['periods']

    # 1. Check Direct Conflicts
    for slot in schedule:
        if slot['day'] == day_idx and slot['period'] in shift_periods:
            if week in slot['weeks']:
                if slot.get('type') == 'practice':
                    return False, "PRACTICE", f"第 {week} 周有实践课程，整周不可值班"
                return False, "CLASS", f"当日第 {slot['period']} 节有课"

    # 2. Check Commute
    if person_campus == CAMPUS_SOUTH and shift_conf['type'] == "afternoon1":
        next_periods = shift_conf['next_periods']
        for slot in schedule:
            if slot['day'] == day_idx and slot['period'] in next_periods:
                if week in slot['weeks']:
                    return False, "COMMUTE", "南校区成员当日第 7–8 节有课，下午一班存在通勤冲突"

    return True, "OK", "可值班"
