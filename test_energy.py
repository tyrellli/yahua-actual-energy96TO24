from datetime import date
from decimal import Decimal
from io import BytesIO
import unittest
from openpyxl import Workbook, load_workbook
from energy import BillError, extract, export


def sample(change=None, offset=0, reverse=False, extra=False):
    wb = Workbook()
    s = wb.active
    s.title = '移动后的工作表'
    s.cell(1+offset,1,'结算日期'); s.cell(1+offset,2,'时段')
    s.cell(1+offset,3,'省内现货日内'); s.cell(2+offset,3,'实际上网电量')
    s.cell(1+offset,4,'省内现货日前'); s.cell(2+offset,4,'实际上网电量')
    s.cell(1+offset,5,'结算单元名称')
    points = list(range(1,97))
    if reverse:
        points.reverse()
    for r,p in enumerate(points, 3+offset):
        for c,v in enumerate([date(2025,11,7),p,Decimal(p)/1000,999,'测试机组'],1):
            s.cell(r,c,float(v) if isinstance(v,Decimal) else v)
    if extra:
        for p in range(1,97):
            s.append([date(2025,11,8),p,2,999,'测试机组'])
    if change:
        change(s)
    out = BytesIO(); wb.save(out)
    return out.getvalue()


class EnergyTests(unittest.TestCase):
    def test_standard_and_wrong_metric(self):
        day, = extract(sample())
        self.assertEqual(day.total,Decimal('4.656'))
        self.assertEqual(day.hours[0],Decimal('.010'))
        self.assertEqual(day.hours[-1],Decimal('.378'))

    def test_header_moved_and_sorted(self):
        day, = extract(sample(offset=6, reverse=True))
        self.assertEqual(day.points[0],Decimal('.001'))
        self.assertEqual(day.points[-1],Decimal('.096'))

    def test_column_reordered(self):
        def swap(s):
            for row in s:
                row[2].value,row[3].value=row[3].value,row[2].value
        self.assertEqual(extract(sample(swap))[0].total,Decimal('4.656'))

    def test_multidate_separate(self):
        days=extract(sample(extra=True))
        self.assertEqual([d.date for d in days],['2025-11-07','2025-11-08'])
        self.assertEqual(days[1].total,Decimal('192'))

    def test_zero_preserved(self):
        day, = extract(sample(lambda s:setattr(s['C3'],'value',0)))
        self.assertEqual(day.points[0],0)

    def test_duplicate(self):
        with self.assertRaisesRegex(BillError,'重复'):
            extract(sample(lambda s:setattr(s['B4'],'value',1)))

    def test_missing(self):
        with self.assertRaisesRegex(BillError,'缺少时段96'):
            extract(sample(lambda s:s.delete_rows(98)))

    def test_blank(self):
        with self.assertRaisesRegex(BillError,'为空'):
            extract(sample(lambda s:setattr(s['C3'],'value',None)))

    def test_text_quantity(self):
        with self.assertRaisesRegex(BillError,'不是有效数字'):
            extract(sample(lambda s:setattr(s['C3'],'value','错误')))

    def test_bad_date(self):
        with self.assertRaisesRegex(BillError,'日期为空或无效'):
            extract(sample(lambda s:setattr(s['A3'],'value','不是日期')))

    def test_bad_period(self):
        with self.assertRaisesRegex(BillError,'时段必须'):
            extract(sample(lambda s:setattr(s['B3'],'value',1.5)))

    def test_wrong_category(self):
        with self.assertRaisesRegex(BillError,'未找到'):
            extract(sample(lambda s:setattr(s['C1'],'value','省间现货日内')))

    def test_duplicate_column(self):
        def duplicate(s):
            s['F1']='省内现货日内'; s['F2']='实际上网电量'
        with self.assertRaisesRegex(BillError,'多个目标字段'):
            extract(sample(duplicate))

    def test_corrupt(self):
        with self.assertRaisesRegex(BillError,'无法读取'):
            extract(b'not an excel file')

    def test_excel_export_cached_values_and_formulas(self):
        day, = extract(sample())
        result=export(day,'sample.xlsx')
        values=load_workbook(BytesIO(result),data_only=True)
        formulas=load_workbook(BytesIO(result),data_only=False)
        s=values['24时段汇总']
        self.assertEqual(s['D6'].value,.01)
        self.assertEqual(s['D30'].value,4.656)
        self.assertEqual(s['B34'].value,96)
        self.assertEqual(s['B37'].value,0)
        self.assertEqual(s['B38'].value,'一致')
        self.assertEqual(formulas['24时段汇总']['D29'].value,"=SUM('原始96点'!D100:D103)")
        self.assertEqual(values['原始96点']['D103'].value,.096)


if __name__=='__main__':
    unittest.main()
