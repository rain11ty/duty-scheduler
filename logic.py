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
DAY_TO_INDEX = {day: index for index, day in enumerate(DAYS_CN)}


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

    normalized = (
        str(text)
        .replace("（", "(")
        .replace("）", ")")
        .replace("，", ",")
        .replace("\n", "")
    )

    for match in re.finditer(r'实践课程[:：](.*?)(?:其他课程[:：]|打印时间|$)', normalized):
        practice_text = match.group(1).strip()
        weeks = parse_practice_weeks_from_text(practice_text)
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
    match = re.fullmatch(r'\s*(\d{1,2})\s*[-－—]\s*(\d{1,2})\s*', str(value or ""))
    if not match:
        return None
    start, end = int(match.group(1)), int(match.group(2))
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
        if len(cells) < 3:
            continue

        day_text, period_text, content = cells[0], cells[1], cells[2]
        if day_text in DAY_TO_INDEX:
            current_day = DAY_TO_INDEX[day_text]

        parsed_periods = parse_period_range(period_text)
        if parsed_periods:
            current_periods = parsed_periods

        if current_day is None or not current_periods:
            continue
        if not content or "周数" not in content:
            continue

        weeks = parse_weeks_from_text(content)
        if not weeks:
            continue

        campus = CAMPUS_NORTH
        if "南校区" in content:
            campus = CAMPUS_SOUTH
        elif "北校区" in content:
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

    # Avoid treating class periods such as "(1-2节)" as week ranges.
    clean_text = re.sub(r'\(\d{1,2}-\d{1,2}节\)', '', normalized)
    clean_text = re.sub(r'\d{1,2}-\d{1,2}节', '', clean_text)

    def add_range(start, end, parity=None):
        start, end = int(start), int(end or start)
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
        r'(?P<body>(?:\d{1,2}(?:-\d{1,2})?\s*,?\s*)+)'
        r'(?P<pre_parity>\([单双]\))?\s*周'
        r'(?P<post_parity>\([单双]\))?'
    )
    for match in week_expr.finditer(clean_text):
        parity_text = match.group("pre_parity") or match.group("post_parity") or ""
        parity = "单" if "单" in parity_text else "双" if "双" in parity_text else None
        for token in re.finditer(r'(\d{1,2})(?:-(\d{1,2}))?', match.group("body")):
            add_range(token.group(1), token.group(2), parity)

    return weeks

def parse_pdf_schedule(pdf_file):
    """
    Parses a PDF schedule. 
    Returns a list of busy slots: {'day': 0-4, 'period': int, 'weeks': set, 'campus': str}
    """
    busy_slots = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            add_practice_week_slots(busy_slots, page.extract_text() or "")
            tables = page.extract_tables()
            for table in tables:
                busy_slots.extend(parse_list_schedule_table(table))
    return busy_slots

def check_availability(person, week, day_idx, shift_name, all_schedules):
    """
    Returns: (is_available, reason_code, debug_info)
    """
    name = person['Name']
    person_campus = person['Campus']
    
    # Check if schedule exists for this person
    if name not in all_schedules:
        return False, "NO_SCHEDULE", "No schedule uploaded"
        
    schedule = all_schedules.get(name, [])
    
    shift_conf = SHIFTS[shift_name]
    shift_periods = shift_conf['periods']
    
    # 1. Check Direct Conflicts
    for slot in schedule:
        if slot['day'] == day_idx and slot['period'] in shift_periods:
            if week in slot['weeks']:
                if slot.get('type') == 'practice':
                    return False, "PRACTICE", "Practice course week"
                return False, "CLASS", f"Class at Period {slot['period']}"

    # 2. Check Commute
    if person_campus == CAMPUS_SOUTH and shift_conf['type'] == "afternoon1":
        next_periods = shift_conf['next_periods']
        for slot in schedule:
            if slot['day'] == day_idx and slot['period'] in next_periods:
                if week in slot['weeks']:
                    return False, "COMMUTE", "Has class at 16:00 (Commute Risk)"

    return True, "OK", "Available"
