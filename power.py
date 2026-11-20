"""横向96点出力／联络线文件，按日期及标识隔离，输出无底色竖版。"""
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from io import BytesIO
import re

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import xlsxwriter
from price import price_date, price_number, price_period, PriceError
from energy import clock


class PowerError(ValueError):
    pass


@dataclass
class PowerGroup:
    date: str
    identifier: str
    sheet: str
    curves: dict
    cells: dict
    inferred: bool = False

    @property
    def label(self):
        return f'{self.date} · {self.identifier or "未标识单元"} · {self.sheet}'

    def hours(self, method='mean'):
        if method not in ('mean', 'sum'):
            raise PowerError('请选择取平均或求和。')
        divisor = 4 if method == 'mean' else 1
        return {name: tuple(sum(points[h:h+4], Decimal(0))/divisor
                            for h in range(0, 96, 4)) for name, points in self.curves.items()}


def _period(value):
    if value is None or isinstance(value, bool):
        return None
    text = str(value).strip()
    if re.fullmatch(r'\d{1,2}:\d{2}(?::00)?', text):
        return price_period(text, False, '时段表头')
    try:
        number = Decimal(text)
        if number.is_finite() and number == number.to_integral_value() and 1 <= number <= 96:
            return int(number)
    except InvalidOperation:
        pass
    return None


def _infer_fields(row):
    """按内容定位唯一连续96数值区间，不固定列号，允许额外元数据及备注。"""
    runs, start = [], None
    for i in range(len(row)+1):
        numeric = False
        if i < len(row):
            try:
                price_number(row[i], '无表头数据')
                numeric = True
            except PriceError:
                pass
        if numeric and start is None:
            start = i
        if not numeric and start is not None:
            runs.append((start, i))
            start = None
    candidates = [(a,b) for a,b in runs if b-a == 96]
    # 长于96点的数值段无法分清编号、合计和真实时段，不截取其中一段。
    if len(candidates) != 1 or any(b-a > 96 for a,b in runs):
        return None
    start, end = candidates[0]
    metadata = [i for i in range(len(row)) if not start <= i < end]
    dates = []
    for i in metadata:
        try:
            price_date(row[i], '无表头日期')
            dates.append(i)
        except PriceError:
            pass
    if len(dates) != 1:
        return None
    names = []
    for i in range(start):
        value = row[i]
        if i in dates or value is None or not str(value).strip():
            continue
        try:
            price_number(value, '元数据')
        except PriceError:
            names.append(i)
    # 数据区前最近的文字列作为候选名称，网页展示后由用户核对。
    name = names[-1] if names else None
    identifiers = [i for i in range(start) if i not in dates and i != name]
    return dates[0], name, identifiers, {p: start+p-1 for p in range(1, 97)}


def extract_power(contents, filename=''):
    try:
        book = load_workbook(BytesIO(contents), read_only=True, data_only=True)
    except Exception as exc:
        raise PowerError('无法读取Excel，请上传有效的.xlsx文件。') from exc
    groups = {}
    try:
        for sheet in book:
            sheet.reset_dimensions()
            iterator = sheet.iter_rows(values_only=True)
            fields = None
            inferred = False
            for row_index, row in enumerate(iterator, 1):
                if fields is None:
                    if row_index > 100:
                        break
                    labels = [str(v or '').strip() for v in row]
                    if '日期' not in labels:
                        fields = _infer_fields(row)
                        if fields is None:
                            numeric_count = 0
                            for value in row:
                                try:
                                    price_number(value, '待识别行')
                                    numeric_count += 1
                                except PriceError:
                                    pass
                            if numeric_count >= 90:
                                raise PowerError(f'{sheet.title}第{row_index}行：疑似无表头数据，但日期／名称／96点范围无法唯一识别，请核对。')
                            continue
                        inferred = True
                    else:
                        dates = [i for i, value in enumerate(labels) if value == '日期']
                        names = [i for i, value in enumerate(labels) if value in ('名称', '出力类型')]
                        periods = {}
                        for i, value in enumerate(row):
                            period = _period(value)
                            if period is not None:
                                if period in periods:
                                    raise PowerError(f'{sheet.title}：重复时段{period}。')
                                periods[period] = i
                        if not periods:
                            raise PowerError(f'{sheet.title}：有日期表头但无有效96点时段表头，请核对。')
                        missing = set(range(1, 97))-set(periods)
                        if missing:
                            raise PowerError(f'{sheet.title}：缺少时段' + '、'.join(map(str, sorted(missing))))
                        if len(dates) != 1 or len(names) > 1:
                            raise PowerError(f'{sheet.title}：日期／名称表头不唯一。')
                        if any(i >= min(periods.values()) for i in dates+names):
                            raise PowerError('日期和名称必须位于96点数据之前。')
                        identifiers = [i for i in range(min(periods.values())) if i not in dates+names]
                        fields = (dates[0], names[0] if names else None, identifiers, periods)
                        continue
                if not any(v is not None and str(v).strip() for v in row):
                    continue
                date_col, name_col, identifiers, periods = fields
                val = lambda i: row[i] if i < len(row) else None
                if inferred:
                    check = _infer_fields(row)
                    if check != fields:
                        raise PowerError(f'{sheet.title}第{row_index}行：无表头结构不一致或不是完整96点，请核对。')
                day = price_date(val(date_col), f'{sheet.title}第{row_index}行')
                identifier = ' / '.join(str(val(i)) for i in identifiers if val(i) is not None and str(val(i)).strip())
                if name_col is not None:
                    name = str(val(name_col) or '').strip()
                    if not name:
                        raise PowerError(f'{sheet.title}第{row_index}行：曲线名称为空。')
                else:
                    name = '省间联络线' if '联络线' in filename else '未命名曲线'
                key = (day, identifier, sheet.title)
                group = groups.setdefault(key, PowerGroup(day, identifier, sheet.title, {}, {}, inferred))
                if name in group.curves:
                    raise PowerError(f'{day}／{identifier}：重复曲线“{name}”，请核对后上传。')
                points = tuple(price_number(val(periods[p]), f'{sheet.title}!{get_column_letter(periods[p]+1)}{row_index}')
                               for p in range(1, 97))
                group.curves[name] = points
                group.cells[name] = tuple(f'{get_column_letter(periods[p]+1)}{row_index}' for p in range(1, 97))
        if not groups:
            raise PowerError('未识别到完整横向96点数据。无表头时需要明确日期和唯一连续96个有效数值；连续数值超过96个时不自动截取。')
        return list(groups.values())
    except PriceError as exc:
        raise PowerError(str(exc)) from exc
    finally:
        book.close()


def export_power(group, filename, method='mean'):
    hours = group.hours(method)
    out = BytesIO()
    book = xlsxwriter.Workbook(out, {'in_memory': True, 'strings_to_formulas': False, 'strings_to_urls': False})
    header = book.add_format({'bold': True, 'font_color': '#000000'})
    number = book.add_format({'num_format': '0.000', 'font_color': '#000000'})
    names = list(group.curves)
    for name, count, values in [('24时段', 24, hours), ('原始96点', 96, group.curves)]:
        sheet = book.add_worksheet(name)
        sheet.freeze_panes(1, 2)
        sheet.set_column(0, 0, 24)
        sheet.set_column(1, 1, 16)
        sheet.set_column(2, 1+len(names), 29)
        sheet.write_row(0, 0, ['时段（结束时刻）', '日期']+names, header)
        for i in range(count):
            sheet.write_string(i+1, 0, clock((i+1)*(60 if count == 24 else 15)))
            sheet.write_string(i+1, 1, group.date)
            for j, metric in enumerate(names, 2):
                sheet.write_number(i+1, j, float(values[metric][i]), number)
        row = count+3
        sheet.write_row(row, 0, ['计算口径', '每4点取算术平均' if method == 'mean' else '每4点求和'])
        sheet.write_row(row+1, 0, ['单位与方向', '沿用原表；未注明单位时不猜测；正负号原样保留。'])
        sheet.write_row(row+2, 0, ['来源文件', filename])
        sheet.write_row(row+3, 0, ['工作表／标识', f'{group.sheet} / {group.identifier}'])
        if group.inferred:
            sheet.write_row(row+4, 0, ['无表头说明', '按原列序作为00:15至24:00；转换前由用户确认时间顺序。'])
    book.close()
    return out.getvalue()
