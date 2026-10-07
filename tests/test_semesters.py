from datetime import date
import unittest

from semesters import default_term, validate_pdf_term, TermMismatchError


class SemesterTests(unittest.TestCase):
    def test_default_term_uses_academic_year_across_calendar_boundary(self):
        self.assertEqual(default_term(date(2026, 10, 7)), '2026-2027-1')
        self.assertEqual(default_term(date(2027, 1, 7)), '2026-2027-1')
        self.assertEqual(default_term(date(2027, 3, 7)), '2026-2027-2')

    def test_pdf_header_accepts_spacing_dashes_and_chinese_numerals(self):
        for header in ['2026-2027学年第1学期', '2026 — 2027 学年 第 一 学期', '2026－2027学年1学期']:
            validate_pdf_term(header, '2026-2027-1')

    def test_wrong_missing_mixed_and_invalid_year_headers_are_rejected(self):
        with self.assertRaises(TermMismatchError):
            validate_pdf_term('2026-2027学年第2学期', '2026-2027-1')
        for header in ['普通课表', '2026-2028学年第1学期',
                       '2026-2027学年第1学期\n2026-2027学年第2学期']:
            with self.assertRaises(ValueError):
                validate_pdf_term(header, '2026-2027-1')
