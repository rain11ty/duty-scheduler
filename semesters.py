"""Academic terms use the same year and semester labels as teaching-system PDFs."""
from datetime import date
import re


class TermMismatchError(ValueError):
    """A readable file belongs to a different term, not a failed refresh."""


def default_term(today=None):
    today = today or date.today()
    start = today.year if today.month >= 8 else today.year - 1
    semester = 1 if today.month >= 8 or today.month == 1 else 2
    return f'{start}-{start + 1}-{semester}'


def term_label(term):
    start, end, semester = term.split('-')
    return f'{start}–{end} 学年第 {semester} 学期'


def validate_pdf_term(text, expected_term):
    normalized = re.sub(r'\s+', '', text).replace('－', '-').replace('—', '-').replace('–', '-')
    matches = re.findall(r'(\d{4})-(\d{4})学年第?([12一二])学期', normalized)
    terms = set()
    for start, end, semester in matches:
        if int(end) != int(start) + 1:
            raise ValueError('课表学年起止年份不合法')
        semester = {'一': '1', '二': '2'}.get(semester, semester)
        terms.add(f'{start}-{end}-{semester}')
    if not terms:
        raise ValueError('PDF 正文未识别到学年和学期，请重新导出带学期标题的课表；不依据文件名猜测学期')
    if len(terms) != 1:
        raise ValueError('PDF 正文包含多个不同学期，不能混用')
    actual = terms.pop()
    if actual != expected_term:
        raise TermMismatchError(f'课表属于 {term_label(actual)}，当前选择 {term_label(expected_term)}，请切换学期后导入')
