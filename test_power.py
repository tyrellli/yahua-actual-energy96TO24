from datetime import datetime
from decimal import Decimal
from io import BytesIO
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from power import PowerError, extract_power, export_power


def sample(named=True, clocks=False, change=None, reverse=False):
    book = Workbook(); sheet = book.active
    points = list(range(1, 97))
    if reverse:
        points.reverse()
    headers = [f'{p//4:02d}:{p%4*15:02d}' if clocks else p for p in points]
    sheet.append(['日期']+(['出力类型'] if named else [None])+headers)
    sheet.append([datetime(2026, 10, 15), '风电总出力' if named else 1010010]+[p-50 for p in points])
    if named:
        sheet.append([datetime(2026, 10, 15), '光伏总出力']+[p*2 for p in points])
    if change:
        change(sheet)
    out = BytesIO(); book.save(out); book.close()
    return out.getvalue()


class PowerTests(unittest.TestCase):
    def test_formats_order_and_both_methods(self):
        for clocks in (False, True):
            group, = extract_power(sample(clocks=clocks, reverse=True))
            self.assertEqual(group.hours()['风电总出力'][0], Decimal('-47.5'))
            self.assertEqual(group.hours('sum')['风电总出力'][0], Decimal('-190'))
            self.assertEqual(group.hours()['风电总出力'][-1], Decimal('44.5'))
            self.assertEqual(group.hours()['光伏总出力'][0], Decimal('5'))

    def test_nameless_intertie(self):
        group, = extract_power(sample(named=False), '省间联络线.xlsx')
        self.assertEqual(group.identifier, '1010010')
        self.assertIn('省间联络线', group.curves)

    def test_reject_bad_points_and_headers(self):
        for value in (None, True, 'NaN', 'Infinity', '文字'):
            with self.subTest(value=value), self.assertRaises(PowerError):
                extract_power(sample(change=lambda s: setattr(s['C2'], 'value', value)))
        with self.assertRaisesRegex(PowerError, '缺少'):
            extract_power(sample(change=lambda s: s.delete_cols(98)))
        with self.assertRaisesRegex(PowerError, '重复时段'):
            extract_power(sample(change=lambda s: setattr(s['D1'], 'value', 1)))
        with self.assertRaisesRegex(PowerError, '日期'):
            extract_power(sample(change=lambda s: setattr(s['A2'], 'value', '错误日期')))
        with self.assertRaisesRegex(PowerError, '重复曲线'):
            extract_power(sample(change=lambda s: setattr(s['B3'], 'value', '风电总出力')))

    def test_dates_are_separate_and_excel_is_plain(self):
        groups = extract_power(sample(change=lambda s: setattr(s['A3'], 'value', datetime(2026, 10, 16))))
        self.assertEqual(len(groups), 2)
        group, = extract_power(sample())
        for method, expected in [('mean', -47.5), ('sum', -190)]:
            book = load_workbook(BytesIO(export_power(group, '测试.xlsx', method)), data_only=True)
            s = book['24时段']
            self.assertEqual(s['A25'].value, '24:00')
            self.assertEqual(s['C2'].value, expected)
            for sheet in book:
                for row in sheet:
                    for cell in row:
                        self.assertIn(cell.fill.patternType, (None, 'none'))
            book.close()

    def test_streamlit_power_page_and_switch(self):
        from streamlit.testing.v1 import AppTest
        upload = BytesIO(sample()); upload.name = '虚构出力.xlsx'
        app = AppTest.from_file('app.py').run()
        with patch('streamlit.file_uploader', return_value=[upload]), patch('streamlit.selectbox', return_value=0):
            app.sidebar.radio[0].set_value('出力／联络线').run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            self.assertEqual(app.dataframe[0].value.iloc[0,2], -47.5)
            app.radio[0].set_value('求和（15分钟电量）').run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.dataframe[0].value.iloc[0,2], -190)


if __name__ == '__main__':
    unittest.main()
