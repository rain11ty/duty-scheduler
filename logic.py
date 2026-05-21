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
            tables = page.extract_tables()
            for table in tables:
                if not table or len(table) < 3: continue
                
                header_row = table[1]
                day_map = {}
                for idx, col in enumerate(header_row):
                    if col:
                        c = col.replace('\n', '').strip()
                        if "星期一" in c: day_map[0] = idx
                        elif "星期二" in c: day_map[1] = idx
                        elif "星期三" in c: day_map[2] = idx
                        elif "星期四" in c: day_map[3] = idx
                        elif "星期五" in c: day_map[4] = idx
                        elif "星期六" in c: day_map[5] = idx
                        elif "星期日" in c: day_map[6] = idx
                
                if not day_map: continue
                
                for r_idx in range(2, len(table)):
                    row = table[r_idx]
                    period_val = row[1]
                    if period_val is None: continue
                    
                    try:
                        p_str = str(period_val).strip()
                        if not p_str.isdigit(): continue
                        period_num = int(p_str)
                    except:
                        continue
                        
                    for day_idx, col_idx in day_map.items():
                        if col_idx < len(row):
                            content = row[col_idx]
                            if content and len(content.strip()) > 2:
                                content = content.strip()
                                weeks = parse_weeks_from_text(content)
                                
                                campus = CAMPUS_NORTH 
                                if "南校区" in content:
                                    campus = CAMPUS_SOUTH
                                elif "北校区" in content:
                                    campus = CAMPUS_NORTH
                                
                                covered_periods = [period_num]
                                p_match = re.search(r'\((\d+)-(\d+)节\)', content)
                                if p_match:
                                    s, e = int(p_match.group(1)), int(p_match.group(2))
                                    covered_periods = list(range(s, e+1))
                                
                                for p in covered_periods:
                                    busy_slots.append({
                                        'day': day_idx,
                                        'period': p,
                                        'weeks': weeks,
                                        'campus': campus,
                                        'raw': content
                                    })
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
                return False, "CLASS", f"Class at Period {slot['period']}"

    # 2. Check Commute
    if person_campus == CAMPUS_SOUTH and shift_conf['type'] == "afternoon1":
        next_periods = shift_conf['next_periods']
        for slot in schedule:
            if slot['day'] == day_idx and slot['period'] in next_periods:
                if week in slot['weeks']:
                    return False, "COMMUTE", "Has class at 16:00 (Commute Risk)"

    return True, "OK", "Available"
