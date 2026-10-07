import json
import os
from pathlib import Path
import tempfile
import unittest

from streamlit.testing.v1 import AppTest

from logic import CAMPUS_NORTH, DAYS_CN, SHIFTS, SCHEDULE_PARSER_VERSION
from semesters import default_term


APP_FILE = Path(__file__).resolve().parents[1] / 'duty_app.py'
PERSON = {'Name': 'sample', 'Campus': CAMPUS_NORTH, 'Class': 'sample class',
          'Department': 'sample department', 'Role': '干事', 'Grade': '大一'}
DUTY = {**PERSON, 'Week': 2, 'Day': DAYS_CN[0], 'Shift': next(iter(SHIFTS))}
CLASS_SLOT = {'day': 0, 'period': 3, 'weeks': [2],
              'campus': CAMPUS_NORTH, 'type': 'class'}


class AppTests(unittest.TestCase):
    def setUp(self):
        self.original_cwd = Path.cwd()
        self.temp = tempfile.TemporaryDirectory(dir=APP_FILE.parent / 'tests')
        os.chdir(self.temp.name)
        Path('data').mkdir()
        Path('data/terms').mkdir()

    def tearDown(self):
        os.chdir(self.original_cwd)
        self.temp.cleanup()

    def start(self, version=None, slots=None, person=None):
        payload = {'academic_term': default_term(), 'personnel': [person or PERSON],
                   'schedules': {'sample': slots if slots is not None else [CLASS_SLOT]},
                   'duty_table': [DUTY]}
        if version is not None:
            payload['schedule_parser_version'] = version
        Path(f'data/terms/{default_term()}.json').write_text(json.dumps(payload), encoding='utf-8')
        app = AppTest.from_file(str(APP_FILE)).run(timeout=20)
        self.assertEqual(len(app.exception), 0)
        return app

    def test_old_parser_cache_requires_reimport_and_preserves_other_data(self):
        # The old parser often cached only practice slots, so checking emptiness is insufficient.
        app = self.start(slots=[{**CLASS_SLOT, 'weeks': [18], 'type': 'practice'}])
        self.assertEqual(app.session_state['schedules'], {})
        self.assertEqual(app.session_state['personnel'].iloc[0]['Name'], 'sample')
        self.assertEqual(len(app.session_state['duty_table']), 1)
        self.assertTrue(any('重新导入' in warning.value for warning in app.warning))

    def test_current_cache_restores_integer_week_sets(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        self.assertEqual(app.session_state['schedules']['sample'][0]['weeks'], {2})

    def test_conflicting_previous_assignment_cannot_be_saved(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        self.assertTrue(any('现有排班有' in warning.value for warning in app.warning))
        app.sidebar.multiselect[0].set_value([2]).run()
        self.assertTrue(any('当前不可值班' in warning.value for warning in app.warning))
        save = next(button for button in app.button if button.label == '💾 保存当前周排班')
        save.click().run()
        self.assertEqual(len(app.exception), 0)
        self.assertTrue(any('冲突' in error.value for error in app.error))
        self.assertEqual(len(app.session_state['duty_table']), 1)
        self.assertEqual(len(app.metric), 3)

    def test_valid_assignment_can_be_saved_and_cache_version_is_written(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION, slots=[{**CLASS_SLOT, 'day': 1}])
        app.sidebar.multiselect[0].set_value([2]).run()
        save = next(button for button in app.button if button.label == '💾 保存当前周排班')
        save.click().run()
        self.assertEqual(len(app.exception), 0)
        self.assertTrue(any('保存成功' in success.value for success in app.success))
        payload = json.loads(Path(f'data/terms/{default_term()}.json').read_text(encoding='utf-8'))
        self.assertEqual(payload['schedule_parser_version'], SCHEDULE_PARSER_VERSION)
        self.assertEqual(payload['duty_table'][0]['Name'], 'sample')

    def helpers(self, suffix):
        script = Path('helpers_app.py').resolve()
        script.write_text(APP_FILE.read_text(encoding='utf-8') + '\n' + suffix, encoding='utf-8')
        app = AppTest.from_file(str(script)).run(timeout=20)
        self.assertEqual(len(app.exception), 0)
        return app

    def test_incremental_pdf_import_preserves_unrelated_schedules(self):
        app = self.helpers('''
from types import SimpleNamespace
st.session_state['personnel'] = pd.DataFrame([
    {'Name':'sample', 'Campus':'北校区'}, {'Name':'other', 'Campus':'北校区'}])
st.session_state['schedules'] = {'sample': [{'day':0,'period':3,'weeks':{2},'campus':'北校区'}]}
parse_pdf_schedule = lambda file, **kwargs: [{'day':1,'period':5,'weeks':{2},'campus':'北校区'}]
parse_schedule_files([SimpleNamespace(name='other.pdf')])
''')
        self.assertEqual(set(app.session_state['schedules']), {'sample', 'other'})

    def test_duplicate_filename_owner_is_rejected_instead_of_last_file_winning(self):
        app = self.helpers('''
from types import SimpleNamespace
st.session_state['personnel'] = pd.DataFrame([{'Name':'sample', 'Campus':'北校区'}])
parse_pdf_schedule = lambda file, **kwargs: [{'day':1,'period':5,'weeks':{2},'campus':'北校区'}]
parse_schedule_files([SimpleNamespace(name='sample-a.pdf'), SimpleNamespace(name='sample-b.pdf')])
''')
        self.assertNotIn('sample', app.session_state['schedules'])
        self.assertTrue(any('失败' in warning.value for warning in app.warning))

    def test_invalid_duty_cells_are_rejected(self):
        app = self.helpers('''
bad_rows = [
    {'Week':-1,'Day':'星期一','Shift':next(iter(SHIFTS)),'Name':'sample'},
    {'Week':'1-2周','Day':'星期一','Shift':next(iter(SHIFTS)),'Name':'sample'},
    {'Week':21,'Day':'星期一','Shift':next(iter(SHIFTS)),'Name':'sample'},
    {'Week':2,'Day':'星期八','Shift':next(iter(SHIFTS)),'Name':'sample'},
    {'Week':2,'Day':'星期一','Shift':'未知班次','Name':'sample'},
]
rejected = []
for row in bad_rows:
    try:
        normalize_duty_table_dataframe(pd.DataFrame([row]))
        rejected.append(False)
    except ValueError:
        rejected.append(True)
st.session_state['test_rejected'] = rejected
''')
        self.assertEqual(app.session_state['test_rejected'], [True] * 5)

    def test_duplicate_duty_position_is_rejected_instead_of_silently_dropping_person(self):
        app = self.helpers('''
rows = [
    {'Week':2,'Day':'星期一','Shift':next(iter(SHIFTS)),'Name':'sample'},
    {'Week':2,'Day':'星期一','Shift':next(iter(SHIFTS)),'Name':'other'},
]
try:
    merge_imported_duty_table(normalize_duty_table_dataframe(pd.DataFrame(rows)))
    st.session_state['test_rejected'] = False
except ValueError:
    st.session_state['test_rejected'] = True
''')
        self.assertTrue(app.session_state['test_rejected'])

    def test_empty_week_selection_does_not_offer_saving_none_week(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        app.sidebar.multiselect[0].set_value([]).run()
        self.assertEqual(len(app.exception), 0)
        self.assertFalse(any(button.label == '💾 保存当前周排班' for button in app.button))

    def test_ambiguous_names_are_rejected(self):
        app = self.helpers('''
st.session_state['personnel'] = pd.DataFrame([{'Name':'sample'}, {'Name':'other'}])
try:
    resolve_schedule_owner_name('sample-other.pdf')
    st.session_state['test_rejected'] = False
except ValueError:
    st.session_state['test_rejected'] = True
''')
        self.assertTrue(app.session_state['test_rejected'])

    def test_duplicate_roster_names_are_rejected(self):
        app = self.helpers('''
try:
    normalize_personnel_dataframe(pd.DataFrame([
        {'Name':'sample','Role':'干事','Campus':'北校区'},
        {'Name':'sample','Role':'干事','Campus':'南校区'}]))
    st.session_state['test_rejected'] = False
except ValueError:
    st.session_state['test_rejected'] = True
''')
        self.assertTrue(app.session_state['test_rejected'])

    def test_failed_refresh_invalidates_only_its_owner(self):
        app = self.helpers('''
from types import SimpleNamespace
st.session_state['personnel'] = pd.DataFrame([{'Name':'sample'}, {'Name':'other'}])
slot = {'day':0,'period':3,'weeks':{2},'campus':'北校区'}
st.session_state['schedules'] = {'sample': [slot], 'other': [slot]}
def fail_parse(file, **kwargs):
    raise ValueError('invalid PDF')
parse_pdf_schedule = fail_parse
parse_schedule_files([SimpleNamespace(name='sample.pdf')])
''')
        self.assertEqual(set(app.session_state['schedules']), {'other'})

    def test_failed_atomic_save_keeps_previous_file_intact(self):
        app = self.helpers('''
from unittest.mock import patch
original = b'{"previous": true}'
get_data_file().write_bytes(original)
try:
    with patch('os.replace', side_effect=OSError('simulated disk error')):
        save_app_state()
except OSError:
    pass
st.session_state['test_intact'] = get_data_file().read_bytes() == original
st.session_state['test_temporary_files'] = list(get_data_file().parent.glob('app_state-*.tmp'))
''')
        self.assertTrue(app.session_state['test_intact'])
        self.assertEqual(app.session_state['test_temporary_files'], [])

    def test_excel_export_round_trip(self):
        app = self.helpers('''
row = {'Week':2,'Day':'周一','Shift':'上午班','Name':'sample','Class':'一班'}
normalized = normalize_duty_table_dataframe(pd.DataFrame([row]))
excel = build_duty_excel_bytes(build_duty_export_df(normalized))
restored = normalize_duty_table_dataframe(pd.read_excel(io.BytesIO(excel)))
st.session_state['test_round_trip'] = restored.to_dict(orient='records') == normalized.to_dict(orient='records')
''')
        self.assertTrue(app.session_state['test_round_trip'])

    def test_unknown_campus_error_persists_while_valid_schedule_is_kept(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION, person={**PERSON, 'Campus': '旗山校区'})
        self.assertTrue(any('校区不明确' in error.value for error in app.error))
        self.assertIn('sample', app.session_state['schedules'])
        self.assertTrue(all(not any(option.startswith('sample') for option in select.options) for select in app.selectbox
                            if select.key and select.key.startswith('sel_')))

    def test_term_switch_isolates_same_week_and_restores_saved_data(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        original_term = default_term()
        other_semester = 2 if original_term.endswith('-1') else 1
        app.sidebar.selectbox[0].set_value(other_semester).run()
        self.assertEqual(len(app.exception), 0)
        self.assertNotIn('personnel', app.session_state)
        self.assertNotIn('schedules', app.session_state)
        self.assertTrue(app.session_state['duty_table'].empty)
        app.sidebar.selectbox[0].set_value(int(original_term[-1])).run()
        self.assertEqual(app.session_state['personnel'].iloc[0]['Name'], 'sample')
        self.assertEqual(app.session_state['schedules']['sample'][0]['weeks'], {2})
        self.assertEqual(len(app.session_state['duty_table']), 1)

    def test_term_switch_clears_previous_manual_selection(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        original_term = default_term()
        key = 'sel_1_0_0'
        app.selectbox(key=key).set_value('sample').run()
        app.sidebar.number_input[0].set_value(int(original_term[:4]) - 1).run()
        self.assertEqual(len(app.exception), 0)
        self.assertNotIn(key, app.session_state)
        self.assertTrue(app.session_state['duty_table'].empty)

    def test_same_week_duties_in_two_terms_remain_separate(self):
        original_term = default_term()
        other_semester = 2 if original_term.endswith('-1') else 1
        other_term = original_term[:-1] + str(other_semester)
        other_person = {**PERSON, 'Name': 'other'}
        payload = {'academic_term': other_term, 'personnel': [other_person],
                   'schedules': {'other': [CLASS_SLOT]}, 'schedule_parser_version': SCHEDULE_PARSER_VERSION,
                   'duty_table': [{**DUTY, 'Name': 'other'}]}
        other_file = Path(f'data/terms/{other_term}.json')
        other_file.write_text(json.dumps(payload), encoding='utf-8')
        original_other_bytes = other_file.read_bytes()
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        app.sidebar.selectbox[0].set_value(other_semester).run()
        self.assertEqual(app.session_state['duty_table'].iloc[0]['Name'], 'other')
        app.sidebar.selectbox[0].set_value(int(original_term[-1])).run()
        self.assertEqual(app.session_state['duty_table'].iloc[0]['Name'], 'sample')
        self.assertEqual(other_file.read_bytes(), original_other_bytes)

    def test_legacy_data_requires_assignment_and_keeps_exact_backup(self):
        original = json.dumps({'personnel': [PERSON], 'duty_table': [DUTY]}, ensure_ascii=False).encode('utf-8')
        Path('data/app_state.json').write_bytes(original)
        app = AppTest.from_file(str(APP_FILE)).run(timeout=20)
        self.assertNotIn('personnel', app.session_state)
        self.assertTrue(any('未标记学期' in item.value for item in app.warning))
        next(button for button in app.button if button.label.startswith('将旧数据归入')).click().run()
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(len(app.session_state['duty_table']), 1)
        self.assertEqual(Path('data/legacy_app_state.json').read_bytes(), original)
        self.assertFalse(Path('data/app_state.json').exists())
        self.assertEqual(json.loads(Path(f'data/terms/{default_term()}.json').read_text(encoding='utf-8'))['academic_term'], default_term())

    def test_migration_never_overwrites_existing_term(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        target = Path(f'data/terms/{default_term()}.json')
        original = target.read_bytes()
        Path('data/app_state.json').write_text('{}', encoding='utf-8')
        app.run()
        button = next(button for button in app.button if button.label.startswith('将旧数据归入'))
        self.assertTrue(button.disabled)
        self.assertEqual(target.read_bytes(), original)
        self.assertTrue(Path('data/app_state.json').exists())

    def test_wrong_term_excel_is_rejected_and_legacy_is_detected(self):
        app = self.helpers('''
row = {'Week':2,'Day':'周一','Shift':'上午班','Name':'sample'}
normalized = normalize_duty_table_dataframe(pd.DataFrame([row]))
excel = build_duty_excel_bytes(build_duty_export_df(normalized))
restored, tagged = read_duty_excel(io.BytesIO(excel))
st.session_state['test_tagged'] = tagged and restored.to_dict('records') == normalized.to_dict('records')
original_term = st.session_state['active_term']
st.session_state['active_term'] = '2000-2001-1'
try:
    read_duty_excel(io.BytesIO(excel))
    st.session_state['test_rejected'] = False
except TermMismatchError:
    st.session_state['test_rejected'] = True
st.session_state['active_term'] = original_term
buffer = io.BytesIO()
normalized.to_excel(buffer, index=False)
_, tagged = read_duty_excel(io.BytesIO(buffer.getvalue()))
st.session_state['test_legacy_detected'] = not tagged
''')
        self.assertTrue(app.session_state['test_tagged'])
        self.assertTrue(app.session_state['test_rejected'])
        self.assertTrue(app.session_state['test_legacy_detected'])

    def test_wrong_term_pdf_does_not_invalidate_current_term_schedule(self):
        app = self.helpers('''
from types import SimpleNamespace
st.session_state['personnel'] = pd.DataFrame([{'Name':'sample'}])
slot = {'day':0,'period':3,'weeks':{2},'campus':'北校区'}
st.session_state['schedules'] = {'sample':[slot]}
def reject_other_term(file, **kwargs):
    raise TermMismatchError('wrong term')
parse_pdf_schedule = reject_other_term
parse_schedule_files([SimpleNamespace(name='sample.pdf')])
''')
        self.assertIn('sample', app.session_state['schedules'])

    def test_counts_include_zero_members_and_each_saved_shift_counts_once(self):
        app = self.helpers('''
people = pd.DataFrame([{'Name':'sample'}, {'Name':'other'}])
duties = pd.DataFrame([{'Name':'sample'}, {'Name':'sample'}, {'Name':'former'}])
st.session_state['test_counts'] = build_duty_counts(people, duties).set_index('姓名')['已排班次数'].to_dict()
''')
        self.assertEqual(app.session_state['test_counts'], {'sample': 2, 'other': 0, 'former': 1})

    def test_clear_current_term_preserves_other_term_file(self):
        app = self.helpers('''
other = get_data_file().parent / '2000-2001-1.json'
other.write_text('{"preserved": true}', encoding='utf-8')
save_app_state()
clear_saved_data_file()
st.session_state['test_other_preserved'] = other.exists()
st.session_state['test_current_removed'] = not get_data_file().exists()
''')
        self.assertTrue(app.session_state['test_other_preserved'])
        self.assertTrue(app.session_state['test_current_removed'])

    def test_import_replacement_drops_stale_planner_widget_values(self):
        app = self.helpers('''
st.session_state['personnel'] = pd.DataFrame([{**{'Name':'other'},'Campus':'北校区'}])
st.session_state['duty_table'] = normalize_duty_table_dataframe(pd.DataFrame([
    {'Week':2,'Day':'星期一','Shift':next(iter(SHIFTS)),'Name':'sample'}]))
st.session_state['sel_2_0_0'] = 'sample'
replacement = pd.DataFrame([{'Week':2,'Day':'星期一','Shift':next(iter(SHIFTS)),'Name':'other'}])
merge_imported_duty_table(replacement, replace_weeks=[2])
st.session_state['test_stale_removed'] = 'sel_2_0_0' not in st.session_state
''')
        self.assertTrue(app.session_state['test_stale_removed'])

    def test_clearing_duty_records_resets_visible_choices_and_cannot_recreate_old_duties(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION, slots=[{**CLASS_SLOT, 'day': 1}])
        app.sidebar.multiselect[0].set_value([2]).run()
        self.assertEqual(app.selectbox(key='sel_2_0_0').value, 'sample')
        next(button for button in app.button if button.label == '清空排班记录').click().run()
        self.assertEqual(len(app.exception), 0)
        self.assertTrue(app.session_state['duty_table'].empty)
        self.assertEqual(app.selectbox(key='sel_2_0_0').value, '未安排')
        next(button for button in app.button if button.label == '💾 保存当前周排班').click().run()
        self.assertTrue(app.session_state['duty_table'].empty)

    def test_template_and_csv_rosters_round_trip(self):
        app = self.helpers('''
template = normalize_personnel_dataframe(read_excel_with_merged_cells(io.BytesIO(build_personnel_template())))
st.session_state['test_template'] = template[['Name','Campus','Role']].to_dict('records')
csv = io.BytesIO('姓名,职务,校区,部门\n测试甲,副部,旗山北校区,摄影部\n'.encode('utf-8-sig'))
csv.name = 'roster.csv'
st.session_state['test_csv'] = normalize_personnel_dataframe(read_personnel_file(csv)).iloc[0].to_dict()
'''.replace("'姓名,职务,校区,部门\n测试甲,副部,旗山北校区,摄影部\n'", "'姓名,职务,校区,部门\\n测试甲,副部,旗山北校区,摄影部\\n'"))
        self.assertEqual(len(app.session_state['test_template']), 4)
        self.assertEqual(app.session_state['test_csv']['Name'], '测试甲')
        self.assertEqual(app.session_state['test_csv']['Role'], '副部长')
        self.assertEqual(app.session_state['test_csv']['Campus'], '北校区')

    def test_folder_discovery_recursive_uppercase_and_errors(self):
        app = self.helpers('''
folder = Path('schedule_folder')
(folder/'nested').mkdir(parents=True)
(folder/'a.PDF').write_bytes(b'fixture')
(folder/'ignore.txt').write_text('ignore')
(folder/'nested'/'b.pdf').write_bytes(b'fixture')
st.session_state['test_nonrecursive'] = [p.name for p in find_pdf_files(str(folder))]
st.session_state['test_recursive'] = [p.name for p in find_pdf_files(str(folder), recursive=True)]
errors = 0
for path in ['', 'missing-folder', str(folder/'ignore.txt')]:
    try:
        find_pdf_files(path)
    except (ValueError, FileNotFoundError, NotADirectoryError):
        errors += 1
st.session_state['test_folder_errors'] = errors
''')
        self.assertEqual(app.session_state['test_nonrecursive'], ['a.PDF'])
        self.assertEqual(app.session_state['test_recursive'], ['a.PDF', 'b.pdf'])
        self.assertEqual(app.session_state['test_folder_errors'], 3)

    def test_query_filters_and_conflicting_time_display_expected_results(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        app.number_input(key='query_week').set_value(2).run()
        next(button for button in app.button if button.label == '🔎 开始查询').click().run()
        self.assertTrue(any('没有符合条件' in item.value for item in app.warning))
        self.assertTrue(any('当日第 3 节有课' in str(frame.value) for frame in app.dataframe))
        app.number_input(key='query_week').set_value(1).run()
        app.multiselect(key='query_role').set_value(['干事']).run()
        next(button for button in app.button if button.label == '🔎 开始查询').click().run()
        self.assertTrue(any('1 位' in item.value for item in app.success))

    def test_deleting_schedule_disables_member_and_preserves_saved_history(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        next(button for button in app.button if button.label == '🗑️ 删除选中课表').click().run()
        self.assertEqual(app.session_state['schedules'], {})
        self.assertEqual(len(app.session_state['duty_table']), 1)
        self.assertTrue(any('现有排班有' in item.value for item in app.warning))
        self.assertTrue(all('sample' not in select.options for select in app.selectbox
                            if select.key and select.key.startswith('sel_')))

    def test_native_save_write_and_cancel_without_opening_a_real_window(self):
        from unittest.mock import patch
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        output = Path('chosen/location.xlsx').resolve()
        with patch('tkinter.Tk'), patch('tkinter.filedialog.asksaveasfilename', return_value=str(output)):
            next(button for button in app.button if button.label == '💾 选择位置并保存').click().run()
        self.assertEqual(len(app.exception), 0)
        self.assertTrue(output.exists())
        import pandas as pd
        self.assertEqual(pd.read_excel(output).iloc[0]['姓名'], 'sample')
        self.assertEqual(pd.read_excel(output, sheet_name='学期信息').iloc[0]['学期标识'], default_term())
        with patch('tkinter.Tk'), patch('tkinter.filedialog.asksaveasfilename', return_value=''):
            next(button for button in app.button if button.label == '💾 选择位置并保存').click().run()
        self.assertTrue(any('已取消保存' in item.value for item in app.info))

    def test_native_save_failure_is_shown_to_user(self):
        from unittest.mock import patch
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        with patch('tkinter.Tk'), patch('tkinter.filedialog.asksaveasfilename', side_effect=OSError('dialog failure')):
            next(button for button in app.button if button.label == '💾 选择位置并保存').click().run()
        self.assertEqual(len(app.exception), 0)
        self.assertTrue(any('保存失败' in item.value for item in app.error))

    def test_clear_all_is_current_term_only_and_leaves_no_loaded_data(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        other = Path('data/terms/2000-2001-1.json')
        other.write_text('{"preserved":true}')
        next(button for button in app.button if button.label == '清空所有数据').click().run()
        self.assertEqual(len(app.exception), 0)
        self.assertNotIn('personnel', app.session_state)
        self.assertNotIn('schedules', app.session_state)
        self.assertTrue(app.session_state['duty_table'].empty)
        self.assertTrue(other.exists())
        self.assertFalse(Path(f'data/terms/{default_term()}.json').exists())

    def test_clear_roster_and_clear_schedules_preserve_duty_history(self):
        for label, removed_key, retained_key in [('清空人员名单', 'personnel', 'schedules'),
                                                 ('清空全部课表', 'schedules', 'personnel')]:
            with self.subTest(action=label):
                app = self.start(version=SCHEDULE_PARSER_VERSION)
                next(button for button in app.button if button.label == label).click().run()
                self.assertEqual(len(app.exception), 0)
                self.assertNotIn(removed_key, app.session_state)
                self.assertIn(retained_key, app.session_state)
                self.assertEqual(len(app.session_state['duty_table']), 1)

    def test_import_log_remains_visible_after_unrelated_page_interaction(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION)
        app.session_state['schedule_import_logs'] = ['diagnostic: invalid sample.pdf']
        app.run()
        self.assertTrue(any('diagnostic: invalid sample.pdf' in item.value for item in app.markdown))
        app.number_input(key='query_week').set_value(3).run()
        self.assertTrue(any('diagnostic: invalid sample.pdf' in item.value for item in app.markdown))

    def test_saved_new_choices_restore_after_restart_without_double_counting(self):
        app = self.start(version=SCHEDULE_PARSER_VERSION, slots=[{**CLASS_SLOT, 'day': 1}])
        app.sidebar.multiselect[0].set_value([2]).run()
        app.selectbox(key='sel_2_0_1').set_value('sample').run()
        next(button for button in app.button if button.label == '💾 保存当前周排班').click().run()
        next(button for button in app.button if button.label == '💾 保存当前周排班').click().run()
        self.assertEqual(len(app.session_state['duty_table']), 2)
        restarted = AppTest.from_file(str(APP_FILE)).run(timeout=20)
        self.assertEqual(len(restarted.exception), 0)
        self.assertEqual(len(restarted.session_state['duty_table']), 2)
        restarted.sidebar.multiselect[0].set_value([2]).run()
        self.assertEqual(restarted.selectbox(key='sel_2_0_1').value, 'sample')


if __name__ == '__main__':
    unittest.main()
