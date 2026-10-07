import unittest
from unittest.mock import Mock, patch

from logic import (
    CAMPUS_NORTH, CAMPUS_SOUTH, DAYS_CN, SHIFTS,
    check_availability, parse_pdf_schedule,
)


class ScheduleTests(unittest.TestCase):
    def parse_pages(self, pages, expected_term=None):
        pdf = Mock()
        pdf.pages = []
        for fixture in pages:
            text, tables = fixture[:2]
            page = Mock()
            page.extract_text.return_value = text
            page.extract_tables.return_value = tables
            page.images = fixture[2] if len(fixture) > 2 else []
            page.curves = fixture[3] if len(fixture) > 3 else []
            page.lines = page.rects = []
            page.chars = [{'text': char} for char in text]
            pdf.pages.append(page)
        with patch('logic.pdfplumber.open') as open_pdf:
            open_pdf.return_value.__enter__.return_value = pdf
            return parse_pdf_schedule('fixture.pdf', expected_term=expected_term)

    def available(self, slots, week=2, day=0, shift=0, campus=CAMPUS_NORTH):
        return check_availability(
            {'Name': 'sample', 'Campus': campus}, week, day,
            list(SHIFTS)[shift], {'sample': slots},
        )

    def grid(self, content, day=0, period='3'):
        row = ['上午', period] + [''] * 7
        row[day + 2] = content
        return [['时间段', '节次', *DAYS_CN], row,
                [None, '4', *([None] * 7)]]

    def test_grid_class_blocks_shift(self):
        slots = self.parse_pages([('课表', [self.grid(
            '课程(3-4节)2-3周,6周/校区:南校区')])])
        self.assertEqual(self.available(slots)[:2], (False, 'CLASS'))
        self.assertTrue(self.available(slots, week=4)[0])
        self.assertEqual({s['period'] for s in slots}, {3, 4})

    def test_pdf_import_validates_body_term_and_keeps_class_conflicts(self):
        pages = [('2026-2027学年第1学期', [self.grid('课程(3-4节)2周/北校区')])]
        slots = self.parse_pages(pages, expected_term='2026-2027-1')
        self.assertEqual(self.available(slots)[:2], (False, 'CLASS'))
        from semesters import TermMismatchError
        with self.assertRaises(TermMismatchError):
            self.parse_pages(pages, expected_term='2026-2027-2')

    def test_term_mismatch_is_detected_before_invalid_course_table(self):
        from semesters import TermMismatchError
        pages = [('2026-2027学年第2学期', [self.grid('课程(3-4节)未知周次/北校区')])]
        with self.assertRaises(TermMismatchError):
            self.parse_pages(pages, expected_term='2026-2027-1')

    def test_grid_multiple_courses_keep_periods_weeks_and_campus_separate(self):
        slots = self.parse_pages([('课表', [self.grid(
            '课程甲(5-6节)1-3周/校区:北校区\n'
            '课程乙（5-8节）12-15周/校区:南校\n区\n'
            '课程丙(7-8节)3-5周(单),6周/校区:南校区', period='5')])])
        self.assertTrue(self.available(slots, week=2, shift=2)[0])
        self.assertEqual(self.available(slots, week=3, shift=2)[:2], (False, 'CLASS'))
        self.assertTrue(self.available(slots, week=4, shift=2)[0])
        self.assertEqual(self.available(slots, week=6, shift=2)[:2], (False, 'CLASS'))
        self.assertEqual(self.available(slots, week=12, shift=2)[:2], (False, 'CLASS'))
        self.assertTrue(all(s['campus'] == CAMPUS_SOUTH for s in slots if 12 in s['weeks']))

    def test_grid_continues_across_pages_without_header_or_period(self):
        first = self.grid('课程(3-4节)2-3周,', day=4)
        continuation = [['', '', '', '', '', '', '6周/校区:南校区', '', ''],
                        ['下午', '5', '另一课程(5-6节)8周/北校区', '', '', '', '', '', '']]
        slots = self.parse_pages([('课表', [first]), ('续页', [continuation])])
        self.assertEqual(self.available(slots, week=6, day=4)[:2], (False, 'CLASS'))
        self.assertEqual(self.available(slots, week=8, shift=1)[:2], (False, 'CLASS'))

    def test_list_schedule_still_supports_merged_cells_and_parity(self):
        table = [['星期', '节次', '课程信息'],
                 ['星期一', '3-4', '课程甲 周数:1-8周(单)/南校区'],
                 [None, None, '课程乙 周数:10-16(双)周/北校区']]
        slots = self.parse_pages([('列表课表', [table])])
        self.assertEqual(self.available(slots, week=3)[:2], (False, 'CLASS'))
        self.assertTrue(self.available(slots, week=4)[0])
        self.assertEqual(self.available(slots, week=12)[:2], (False, 'CLASS'))

    def test_practice_only_blocks_entire_week(self):
        slots = self.parse_pages([('实践课程：实训(共2周)/18-19周/无;\n其他课程：无', [])])
        for day in range(7):
            for shift in range(3):
                self.assertEqual(self.available(slots, week=18, day=day, shift=shift)[:2],
                                 (False, 'PRACTICE'))
        self.assertTrue(self.available(slots, week=17)[0])

    def test_south_campus_commute_and_north_campus_exemption(self):
        slots = self.parse_pages([('课表', [self.grid(
            '课程(7-8节)2周/南校区', period='7')])])
        self.assertEqual(self.available(slots, shift=1, campus=CAMPUS_SOUTH)[:2],
                         (False, 'COMMUTE'))
        self.assertTrue(self.available(slots, shift=1, campus=CAMPUS_NORTH)[0])

    def test_unreadable_or_unrecognized_pdf_is_rejected(self):
        for text, tables in [('', []), ('无法识别的课表', []),
                             ('普通课程(3-4节)2周', []),
                             ('课程(3-4节)2周\n实践课程：实训(共1周)/18周/无;', []),
                             ('课表', [self.grid('课程(3-4节)未知周次/北校区')]),
                             ('实践课程：实训(共1周)/18周/无;',
                              [self.grid('数学分析 学分:4.5')])]:
            with self.subTest(text=text), self.assertRaises(ValueError):
                self.parse_pages([(text, tables)])

    def test_empty_cached_schedule_is_unavailable(self):
        self.assertEqual(self.available([])[:2], (False, 'INVALID_SCHEDULE'))

    def test_single_period_in_list_does_not_inherit_previous_range(self):
        table = [['星期一', '3-4', '课程甲 周数:2周/北校区'],
                 ['星期二', '5', '课程乙 周数:2周/北校区']]
        slots = self.parse_pages([('列表课表', [table])])
        self.assertEqual(self.available(slots, day=1, shift=1)[:2], (False, 'CLASS'))
        self.assertTrue(self.available(slots, day=1, shift=0)[0])

    def test_new_day_missing_period_is_rejected(self):
        table = [['星期一', '3-4', '课程甲 周数:2周/北校区'],
                 ['星期二', '', '课程乙 周数:2周/北校区']]
        with self.assertRaises(ValueError):
            self.parse_pages([('列表课表', [table])])

    def test_sunday_alias_is_not_silently_ignored(self):
        table = self.grid('课程甲(3-4节)2周/北校区')
        table[0][-1] = '星期天'
        table[1][-1] = '课程乙(3-4节)2周/北校区'
        slots = self.parse_pages([('课表', [table])])
        self.assertEqual(self.available(slots, day=6)[:2], (False, 'CLASS'))

    def test_unrecognized_practice_weeks_cannot_be_ignored(self):
        with self.assertRaises(ValueError):
            self.parse_pages([('实践课程：实训(共1周)/周次待定/无;',
                               [self.grid('课程(3-4节)2周/北校区')])])

    def test_one_valid_practice_course_cannot_hide_another_invalid_course(self):
        with self.assertRaises(ValueError):
            self.parse_pages([('实践课程：实训甲(共1周)/18周/无;实训乙(共1周)/周次待定/无;',
                               [self.grid('课程(3-4节)2周/北校区')])])

    def test_unknown_campus_is_not_assumed_to_be_north(self):
        slots = self.parse_pages([('课表', [self.grid('课程(7-8节)2周/南校区', period='7')])])
        self.assertEqual(self.available(slots, shift=1, campus='')[:2],
                         (False, 'UNKNOWN_CAMPUS'))

    def test_scanned_continuation_cannot_yield_partial_available_schedule(self):
        with self.assertRaisesRegex(ValueError, '2'):
            self.parse_pages([('课表', [self.grid('课程(3-4节)2周/北校区')]),
                              ('', [], [{'image': 'scanned schedule'}])])

    def test_image_only_pdf_reports_image_format_and_how_to_reexport(self):
        with self.assertRaisesRegex(ValueError, '第 1 页.*图片版 PDF.*OCR.*选择.*复制'):
            self.parse_pages([('', [], [{'image': 'full page timetable'}])])

    def test_vector_continuation_cannot_yield_partial_schedule(self):
        with self.assertRaisesRegex(ValueError, '第 2 页.*没有可提取文字'):
            self.parse_pages([('课表', [self.grid('课程(3-4节)2周/北校区')]),
                              ('', [], [], [{'path': 'outlined glyphs'}])])

    def test_true_blank_page_can_be_ignored(self):
        slots = self.parse_pages([('课表', [self.grid('课程(3-4节)2周/北校区')]), ('', [])])
        self.assertEqual(self.available(slots)[:2], (False, 'CLASS'))

    def test_list_continues_across_pages_with_repeated_header_and_merged_cells(self):
        first = [['星期', '节次', '课程信息'], ['星期一', '3-4', '课程甲 周数:2周/北校区']]
        second = [['星期', '节次', '课程信息'], [None, None, '课程乙 周次:3周/北校区']]
        slots = self.parse_pages([('列表课表', [first]), ('续页', [second])])
        self.assertEqual(self.available(slots, week=3)[:2], (False, 'CLASS'))

    def test_one_valid_list_course_cannot_hide_missing_course_weeks(self):
        table = [['星期一', '3-4', '课程甲 周数:2周/北校区'],
                 ['星期二', '5-6', '课程乙 学分:4.5']]
        with self.assertRaisesRegex(ValueError, '周次'):
            self.parse_pages([('列表课表', [table])])

    def test_unknown_list_weekday_cannot_inherit_previous_day(self):
        table = [['星期一', '3-4', '课程甲 周数:2周/北校区'],
                 ['星期八', '5-6', '课程乙 周数:3周/北校区']]
        with self.assertRaisesRegex(ValueError, '星期无法识别'):
            self.parse_pages([('列表课表', [table])])


if __name__ == '__main__':
    unittest.main()
