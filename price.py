"""日前／实时出清电价：按等时长15分钟点求小时算术平均。"""
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation
from io import BytesIO
import re

from openpyxl import load_workbook
import xlsxwriter

from energy import clock


class PriceError(ValueError):
    pass


@dataclass(frozen=True)
class PriceDay:
    date: str
    unit: str
    market: str
    sheet: str
    header: str
    price_unit: str
    points: tuple
    cells: tuple

    @property
    def hours(self):
        return tuple(sum(self.points[i:i+4], Decimal(0)) / 4 for i in range(0, 96, 4))

    @property
    def mean(self):
        return sum(self.points, Decimal(0)) / 96

    @property
    def label(self):
        return f'{self.date} · {self.market} · {self.unit or "表内未标机组"} · {self.sheet}'


def clean(value):
    return re.sub(r'\s+', '', str(value if value is not None else '')).replace('（', '(').replace('）', ')')


def price_number(value, location):
    if value is None or isinstance(value, bool) or str(value).strip() == '':
        raise PriceError(f'{location}：电价为空或不是数字，空白不按0处理。')
    try:
        result = Decimal(str(value).strip())
    except InvalidOperation:
        raise PriceError(f'{location}：电价“{value}”不是有效数字。') from None
    if not result.is_finite():
        raise PriceError(f'{location}：电价不是有限数字。')
    return result


def price_date(value, location):
    if isinstance(value, (date, datetime)):
        return value.strftime('%Y-%m-%d')
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y年%m月%d日'):
        try:
            return datetime.strptime(str(value or '').strip(), fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    raise PriceError(f'{location}：日期为空或无效，日期须来自表内。')


def price_period(value, numeric_period, location):
    """时间列为结束时间00:15—24:00；时段列为整数1—96。"""
    if isinstance(value, bool) or value is None:
        raise PriceError(f'{location}：时间／时段为空或无效。')
    if numeric_period:
        try:
            p = Decimal(str(value))
            if p.is_finite() and p == p.to_integral_value() and 1 <= p <= 96:
                return int(p)
        except InvalidOperation:
            pass
        raise PriceError(f'{location}：时段必须为1—96的整数。')
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, time):
        minutes = value.hour * 60 + value.minute if not value.second and not value.microsecond else -1
    elif isinstance(value, (int, float, Decimal)):
        # Excel时间序列值，1表示当天结束的24:00，不将00:00猜成24:00。
        minute_value = Decimal(str(value)) * 1440
        nearest = minute_value.to_integral_value()
        minutes = int(nearest) if minute_value.is_finite() and abs(minute_value-nearest) < Decimal('0.000001') else -1
    else:
        match = re.fullmatch(r'(\d{1,2}):(\d{2})(?::00)?', str(value).strip())
        if not match:
            raise PriceError(f'{location}：时间须为00:15至24:00的15分钟结束时刻。')
        h, m = map(int, match.groups())
        minutes = h * 60 + m if 0 <= m < 60 and h <= 24 and (h < 24 or m == 0) else -1
    if 15 <= minutes <= 1440 and minutes % 15 == 0:
        return minutes // 15
    raise PriceError(f'{location}：时间须为00:15至24:00的15分钟结束时刻，不接受00:00。')


def extract_prices(contents):
    try:
        wb = load_workbook(BytesIO(contents), data_only=True)
    except Exception as exc:
        raise PriceError('无法读取Excel，请上传有效的.xlsx出清结果表。') from exc
    days = []
    try:
        for sheet in wb:
            candidates = []
            for r in range(1, min(sheet.max_row, 100)+1):
                fields = {}
                price_units = {}
                for c in range(1, min(sheet.max_column, 512)+1):
                    text = clean(sheet.cell(r, c).value)
                    match = re.fullmatch(r'(价格|电价|出清价格|节点电价)(?:\(([^()]+)\))?', text)
                    if match:
                        key = 'price'
                        price_units[c] = match.group(2) or '原表未标单位'
                    elif text in ('时间', '时段结束时间', '时段'):
                        key = 'period'
                    elif text in ('日期', '运行日期', '出清日期'):
                        key = 'date'
                    elif text in ('机组', '机组名称', '结算单元名称'):
                        key = 'unit'
                    else:
                        continue
                    fields.setdefault(key, []).append(c)
                if all(k in fields for k in ('price', 'period', 'date')):
                    if any(len(v) != 1 for v in fields.values()):
                        raise PriceError(f'“{sheet.title}”第{r}行存在多个价格、时间、日期或机组字段，请明确唯一字段。')
                    columns = {k:v[0] for k,v in fields.items()}
                    candidates.append((r, columns, price_units[columns['price']]))
            if not candidates:
                continue
            if len(candidates) != 1:
                raise PriceError(f'“{sheet.title}”有多处电价表头，请拆分后上传。')
            row, columns, price_unit = candidates[0]
            market = '日前' if '日前' in sheet.title else '实时' if '实时' in sheet.title else '市场未标注'
            groups = defaultdict(dict)
            for r in range(row+1, sheet.max_row+1):
                values = {k:sheet.cell(r,c).value for k,c in columns.items()}
                if all(v is None or (isinstance(v,str) and not v.strip()) for v in values.values()):
                    continue
                if clean(values['period']) in ('合计', '总计', '汇总', '平均', '均价'):
                    continue
                period = price_period(values['period'], clean(sheet.cell(row, columns['period']).value)=='时段',
                                      f'{sheet.title}!{sheet.cell(r,columns["period"]).coordinate}')
                d = price_date(values['date'], f'{sheet.title}第{r}行')
                unit = str(values.get('unit') or '').strip()
                cell = sheet.cell(r, columns['price']).coordinate
                group = groups[(d, unit)]
                if period in group:
                    raise PriceError(f'{d} {unit}：第{period}个15分钟时段重复，不覆盖或混加。')
                group[period] = (price_number(values['price'], f'{sheet.title}!{cell}'), cell)
            if not groups:
                raise PriceError(f'“{sheet.title}”电价表头下没有数据。')
            for (d, unit), group in groups.items():
                missing = sorted(set(range(1,97))-group.keys())
                if missing:
                    raise PriceError(f'{d} {unit}：缺少时段{", ".join(map(str,missing))}，需完整96点。')
                days.append(PriceDay(d, unit, market, sheet.title,
                                     sheet.cell(row,columns['price']).coordinate, price_unit,
                                     tuple(group[p][0] for p in range(1,97)),
                                     tuple(group[p][1] for p in range(1,97))))
    finally:
        wb.close()
    if not days:
        raise PriceError('未找到含“时间（或时段）”“价格（或电价）”“日期”的表头，请上传日前／实时出清结果表。')
    keys = [(d.date,d.unit,d.market) for d in days]
    if len(keys) != len(set(keys)):
        raise PriceError('多张工作表中出现相同日期、机组及市场的数据，请确认唯一数据来源。')
    return sorted(days, key=lambda d:(d.date,d.market,d.unit))


def export_prices(day, filename):
    """导出AVERAGE公式和缓存，不对价格求和、不依赖电量。"""
    out = BytesIO()
    book = xlsxwriter.Workbook(out, {'in_memory':True,'strings_to_formulas':False,'strings_to_urls':False})
    title = book.add_format({'bold':True,'font_size':15,'font_color':'#243E58'})
    header = book.add_format({'bold':True,'bg_color':'#243E58','font_color':'white','align':'center'})
    num = book.add_format({'num_format':'0.0000'})
    total = book.add_format({'num_format':'0.0000','bold':True,'bg_color':'#E7EDF3'})
    summary = book.add_worksheet('24时段电价')
    raw = book.add_worksheet('原始96点电价')
    for s in (summary,raw):
        s.hide_gridlines(2)
        s.set_default_row(22)
        s.set_column('A:A',20)
        s.set_column('B:B',25)
        s.set_column('C:C',24)
        s.set_column('D:D',30,num)
        s.set_column('E:E',30)
    summary.write('A2',f'{day.date} {day.market}电价96转24时段',title)
    summary.write('A3',day.unit or '表内未标机组')
    summary.write_row('A5',['小时序号','小时时段','原始点序号',f'平均电价（{day.price_unit}）'],header)
    for i,price in enumerate(day.hours):
        summary.write_row(i+5,0,[i+1,f'{clock(i*60)}—{clock((i+1)*60)}',f'{i*4+1}—{i*4+4}'])
        summary.write_formula(i+5,3,f"=IF(COUNT('原始96点电价'!D{i*4+8}:D{i*4+11})=4,AVERAGE('原始96点电价'!D{i*4+8}:D{i*4+11}),\"价格数据不完整\")",num,float(price))
    summary.write('C30','全天时间均价',total)
    summary.write_formula('D30','=IF(COUNT(D6:D29)=24,AVERAGE(D6:D29),"价格数据不完整")',total,float(day.mean))
    summary.write_row('A33',['核对项目','结果'],header)
    checks = [('原始价格点数',"=COUNT('原始96点电价'!D8:D103)",96),
              ('原始96点时间均价',"=AVERAGE('原始96点电价'!D8:D103)",float(day.mean)),
              ('24小时均价的平均值','=D30',float(day.mean)),
              ('均价差额','=ROUND(B36-B35,12)',0),
              ('核对结论','=IF(AND(B34=96,COUNT(D6:D29)=24,B37=0),"一致","请检查")','一致')]
    for r,(label,formula,cached) in enumerate(checks,33):
        summary.write(r,0,label)
        summary.write_formula(r,1,formula,num if 34<=r<=36 else None,cached)
    summary.conditional_format('B37',{'type':'cell','criteria':'!=','value':0,
                                      'format':book.add_format({'bg_color':'#FEE2E2','font_color':'#B91C1C'})})
    summary.write('A40','规则：连续4个等时长的15分钟电价取算术平均，价格不求和。')
    summary.write('A41','结束时间00:15—01:00归入00:00—01:00，依次至24:00。')
    summary.write('A42','单位沿用原表；未标单位时不猜测。显示4位小数，计算保留原始精度。')
    summary.write('A43','本表为时间均价，不作为电量加权结算均价。')
    raw.write('A2',f'{day.date} {day.market}原始96点电价',title)
    raw.write('A3',f'来源：{filename}；工作表：{day.sheet}；电价表头：{day.header}')
    raw.write('A4',f'机组：{day.unit or "表内未标机组"}；价格单位：{day.price_unit}')
    raw.write_row('A7',['日期','15分钟序号','结束时间',f'电价（{day.price_unit}）','原表单元格'],header)
    for i,(price,cell) in enumerate(zip(day.points,day.cells)):
        raw.write_row(i+7,0,[day.date,i+1,clock((i+1)*15),float(price),f'{day.sheet}!{cell}'])
    raw.write('C104','全天时间均价',total)
    raw.write_formula('D104','=AVERAGE(D8:D103)',total,float(day.mean))
    raw.freeze_panes(7,0)
    summary.freeze_panes(5,0)
    book.close()
    return out.getvalue()
