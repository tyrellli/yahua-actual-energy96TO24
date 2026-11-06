from datetime import date, time
from decimal import Decimal
from io import BytesIO
import unittest

from openpyxl import Workbook, load_workbook
from price import PriceError, extract_prices, export_prices


def sample(change=None, reverse=False, offset=0):
    book=Workbook();sheet=book.active;sheet.title='实时市场出清结果'
    sheet.append(['说明'] if offset else ['时间','电量','价格','日期'])
    if offset:
        sheet.cell(offset+1,1,'时间');sheet.cell(offset+1,2,'电量')
        sheet.cell(offset+1,3,'价格');sheet.cell(offset+1,4,'日期')
    indices=list(range(1,97))
    if reverse:
        indices.reverse()
    for r,p in enumerate(indices,offset+2):
        minutes=p*15
        values=[f'{minutes//60:02d}:{minutes%60:02d}',0,p,date(2026,11,4)]
        for c,v in enumerate(values,1):
            sheet.cell(r,c,v)
    if change:
        change(sheet)
    out=BytesIO();book.save(out)
    return out.getvalue()


class PriceTests(unittest.TestCase):
    def test_mean_instead_of_sum(self):
        day,=extract_prices(sample())
        self.assertEqual(day.hours[0],Decimal('2.5'))
        self.assertEqual(day.hours[-1],Decimal('94.5'))
        self.assertEqual(day.mean,Decimal('48.5'))
        self.assertEqual(day.price_unit,'原表未标单位')

    def test_moved_header_and_reverse(self):
        day,=extract_prices(sample(reverse=True,offset=7))
        self.assertEqual(day.points[0],1)
        self.assertEqual(day.points[-1],96)

    def test_reordered_columns(self):
        def change(s):
            for row in s:
                row[0].value,row[2].value=row[2].value,row[0].value
        self.assertEqual(extract_prices(sample(change))[0].hours[0],Decimal('2.5'))

    def test_zero_negative(self):
        def change(s):
            s['C2']=0;s['C3']=-2
        self.assertEqual(extract_prices(sample(change))[0].hours[0],Decimal('1.25'))

    def test_explicit_unit(self):
        day,=extract_prices(sample(lambda s:setattr(s['C1'],'value','价格（元/MWh）')))
        self.assertEqual(day.price_unit,'元/MWh')

    def test_day_ahead_with_status(self):
        def change(s):
            s.title='日前市场机组出清结果';s['E1']='开停'
            for r in range(2,98):s.cell(r,5,'停')
        self.assertEqual(extract_prices(sample(change))[0].market,'日前')

    def test_period_numbers(self):
        def change(s):
            s['A1']='时段'
            for p in range(1,97):s.cell(p+1,1,p)
        self.assertEqual(extract_prices(sample(change))[0].hours[0],Decimal('2.5'))

    def test_native_time(self):
        self.assertEqual(extract_prices(sample(lambda s:setattr(s['A2'],'value',time(0,15))))[0].points[0],1)

    def test_excel_time_fraction(self):
        def change(s):
            for p in range(1,97):s.cell(p+1,1,p/96)
        self.assertEqual(extract_prices(sample(change))[0].hours[-1],Decimal('94.5'))

    def test_midnight_not_guessed(self):
        with self.assertRaisesRegex(PriceError,'00:00'):
            extract_prices(sample(lambda s:setattr(s['A97'],'value','00:00')))

    def test_duplicate(self):
        with self.assertRaisesRegex(PriceError,'重复'):
            extract_prices(sample(lambda s:setattr(s['A3'],'value','00:15')))

    def test_missing(self):
        with self.assertRaisesRegex(PriceError,'缺少时段96'):
            extract_prices(sample(lambda s:s.delete_rows(97)))

    def test_blank_and_bad_price(self):
        for value in (None,'bad',True,'NaN','Infinity'):
            with self.subTest(value=value),self.assertRaises(PriceError):
                extract_prices(sample(lambda s:setattr(s['C2'],'value',value)))

    def test_invalid_date(self):
        with self.assertRaisesRegex(PriceError,'日期为空或无效'):
            extract_prices(sample(lambda s:setattr(s['D2'],'value','bad')))

    def test_multiple_dates_and_units_separate(self):
        def change(s):
            s['E1']='机组名称'
            for r in range(2,98):s.cell(r,5,'1号机')
            original=list(s.values)[1:]
            for row in original:s.append([row[0],0,100,date(2026,11,5),'1号机'])
            for row in original:s.append([row[0],0,200,date(2026,11,4),'2号机'])
        days=extract_prices(sample(change))
        self.assertEqual(len(days),3)
        self.assertEqual([(d.date,d.unit,d.mean) for d in days],
                         [('2026-11-04','1号机',Decimal('48.5')),('2026-11-04','2号机',Decimal(200)),('2026-11-05','1号机',Decimal(100))])

    def test_ambiguous_price_columns(self):
        def change(s):s['E1']='电价'
        with self.assertRaisesRegex(PriceError,'多个'):
            extract_prices(sample(change))

    def test_duplicate_sheet_identity(self):
        contents=sample();book=load_workbook(BytesIO(contents));book.copy_worksheet(book.active)
        out=BytesIO();book.save(out)
        with self.assertRaisesRegex(PriceError,'相同日期'):
            extract_prices(out.getvalue())

    def test_missing_header(self):
        with self.assertRaisesRegex(PriceError,'未找到'):
            extract_prices(sample(lambda s:setattr(s['C1'],'value','电量')))

    def test_corrupt(self):
        with self.assertRaisesRegex(PriceError,'无法读取'):
            extract_prices(b'not xlsx')

    def test_export_mean_formula_and_caches(self):
        day,=extract_prices(sample())
        result=export_prices(day,'example.xlsx')
        values=load_workbook(BytesIO(result),data_only=True)
        formulas=load_workbook(BytesIO(result),data_only=False)
        s=values['24时段电价']
        self.assertEqual(s['D6'].value,2.5)
        self.assertEqual(s['D29'].value,94.5)
        self.assertEqual(s['D30'].value,48.5)
        self.assertEqual(s['B34'].value,96)
        self.assertEqual(s['B37'].value,0)
        self.assertEqual(s['B38'].value,'一致')
        self.assertIn("AVERAGE('原始96点电价'!D100:D103)",formulas['24时段电价']['D29'].value)
        self.assertEqual(values['原始96点电价']['C103'].value,'24:00')


if __name__=='__main__':unittest.main()
