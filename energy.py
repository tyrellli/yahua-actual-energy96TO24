"""日清分账单：严格识别省内现货日内实际上网电量，96点转24点。"""
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
import re

from openpyxl import load_workbook
import xlsxwriter


class BillError(ValueError):
    pass


@dataclass(frozen=True)
class Day:
    date: str
    member: str
    unit: str
    sheet: str
    headers: str
    points: tuple
    cells: tuple

    @property
    def hours(self):
        return tuple(sum(self.points[i:i + 4], Decimal(0)) for i in range(0, 96, 4))

    @property
    def total(self):
        return sum(self.points, Decimal(0))

    @property
    def label(self):
        return f'{self.date} · {self.unit or self.member or "未标注主体"} · {self.sheet}'


def clean(value):
    return re.sub(r'\s+', '', str(value or '')).replace('（', '(').replace('）', ')')


def number(value, location):
    if value is None or isinstance(value, bool) or str(value).strip() == '':
        raise BillError(f'{location}：电量为空或不是数字。')
    try:
        result = Decimal(str(value).strip())
    except InvalidOperation:
        raise BillError(f'{location}：电量“{value}”不是有效数字。') from None
    if not result.is_finite():
        raise BillError(f'{location}：电量不是有限数字。')
    return result


def parse_date(value, location):
    if isinstance(value, (date, datetime)):
        return value.strftime('%Y-%m-%d')
    text = str(value or '').strip()
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y年%m月%d日'):
        try:
            return datetime.strptime(text, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    raise BillError(f'{location}：结算日期为空或无效，不使用文件名猜测日期。')


def extract(contents):
    try:
        wb = load_workbook(BytesIO(contents), data_only=True)
    except Exception as exc:
        raise BillError('无法读取Excel，请上传有效的.xlsx账单。') from exc
    days = []
    try:
        for sheet in wb:
            # 按实际单元格读取，不使用依赖dimension的read_only模式。
            merges = list(sheet.merged_cells.ranges)

            def value(r, c):
                v = sheet.cell(r, c).value
                if v is None:
                    for area in merges:
                        if area.min_row <= r <= area.max_row and area.min_col <= c <= area.max_col:
                            return sheet.cell(area.min_row, area.min_col).value
                return v

            candidates = []
            for r in range(1, min(sheet.max_row, 100) + 1):
                for c in range(1, min(sheet.max_column, 512) + 1):
                    text = clean(value(r, c))
                    combined = text == '省内现货日内实际上网电量'
                    separate = text == '实际上网电量' and r > 1 and clean(value(r - 1, c)) == '省内现货日内'
                    if combined or separate:
                        # 合并的字段名只接受其左上角，避免重复识别。
                        if sheet.cell(r, c).value is not None:
                            candidates.append((r, c))
            if not candidates:
                continue
            if len(candidates) != 1:
                raise BillError(f'工作表“{sheet.title}”有多个目标字段，无法唯一定位。请拆分后上传。')
            row, col = candidates[0]
            columns = {}
            top = max(1, row - 1)
            for label in ('结算日期', '时段', '市场成员名称', '结算单元名称'):
                matches = {c for r in range(top, row + 1)
                           for c in range(1, min(sheet.max_column, 512) + 1)
                           if clean(value(r, c)) == label}
                if len(matches) > 1:
                    raise BillError(f'“{sheet.title}”字段“{label}”重复。')
                if matches:
                    columns[label] = matches.pop()
            for required in ('结算日期', '时段'):
                if required not in columns:
                    raise BillError(f'“{sheet.title}”缺少“{required}”字段。')
            groups = defaultdict(dict)
            for r in range(row + 1, sheet.max_row + 1):
                period = value(r, columns['时段'])
                quantity = value(r, col)
                # 完全空行和明确的总计行不作为15分钟记录。
                if all(sheet.cell(r, c).value is None for c in range(1, sheet.max_column + 1)):
                    continue
                if clean(period) in ('合计', '总计', '汇总'):
                    continue
                location = f'{sheet.title}!{sheet.cell(r, col).coordinate}'
                try:
                    p = Decimal(str(period))
                    if not p.is_finite() or p != p.to_integral_value() or not 1 <= p <= 96:
                        raise ValueError
                    p = int(p)
                except (InvalidOperation, ValueError):
                    raise BillError(f'{sheet.title}第{r}行：时段必须为1—96的整数。') from None
                d = parse_date(value(r, columns['结算日期']), f'{sheet.title}第{r}行')
                member = str(value(r, columns['市场成员名称']) or '') if '市场成员名称' in columns else ''
                unit = str(value(r, columns['结算单元名称']) or '') if '结算单元名称' in columns else ''
                group = groups[(d, member, unit)]
                if p in group:
                    raise BillError(f'{d} {unit or member}：时段{p}重复，不进行相加或覆盖。')
                group[p] = (number(quantity, location), sheet.cell(r, col).coordinate)
            if not groups:
                raise BillError(f'“{sheet.title}”目标字段下没有数据。')
            for (d, member, unit), group in groups.items():
                missing = sorted(set(range(1, 97)) - group.keys())
                if missing:
                    raise BillError(f'{d} {unit or member}：缺少时段{", ".join(map(str, missing))}。')
                day = Day(d, member, unit, sheet.title, f'{sheet.cell(top,col).coordinate}:{sheet.cell(row,col).coordinate}',
                          tuple(group[p][0] for p in range(1, 97)), tuple(group[p][1] for p in range(1, 97)))
                if day.total != sum(day.hours, Decimal(0)):
                    raise BillError(f'{d}：96点总和与24点总和不一致。')
                days.append(day)
    finally:
        wb.close()
    if not days:
        raise BillError('未找到“省内现货日内”下的“实际上网电量”。请检查账单，表头须在前100行、前512列内。')
    identities = [(d.date, d.member, d.unit) for d in days]
    if len(set(identities)) != len(identities):
        raise BillError('同一日期、主体的数据出现在多张工作表中，请先确认唯一数据来源。')
    return sorted(days, key=lambda d: (d.date, d.unit, d.member))


def clock(minutes):
    return f'{minutes // 60:02d}:{minutes % 60:02d}'


def export(day, filename):
    """每小时用SUM公式，并写入已核对的缓存值，支持无需重算的预览。"""
    output = BytesIO()
    book = xlsxwriter.Workbook(output, {'in_memory': True, 'strings_to_formulas': False, 'strings_to_urls': False})
    title = book.add_format({'bold': True, 'font_size': 15, 'font_color': '#243E58'})
    header = book.add_format({'bold': True, 'bg_color': '#243E58', 'font_color': 'white', 'align': 'center'})
    num = book.add_format({'num_format': '0.000'})
    total = book.add_format({'num_format': '0.000', 'bold': True, 'bg_color': '#E7EDF3'})
    for name in ('24时段汇总', '原始96点'):
        sheet = book.add_worksheet(name)
        sheet.hide_gridlines(2)
        sheet.set_column('A:A', 20)
        sheet.set_column('B:C', 24)
        sheet.set_column('D:D', 25, num)
        sheet.set_column('E:E', 32)
        sheet.set_default_row(22)
    summary, raw = book.worksheets()
    summary.write('A2', f'{day.date} 实际上网电量小时汇总', title)
    summary.write('A3', day.unit or day.member)
    summary.write_row('A5', ['小时序号', '小时时段', '原始点序号', '电量（原表单位）'], header)
    for i, quantity in enumerate(day.hours):
        summary.write_row(i + 5, 0, [i + 1, f'{clock(i*60)}—{clock((i+1)*60)}', f'{i*4+1}—{i*4+4}'])
        summary.write_formula(i + 5, 3, f"=SUM('原始96点'!D{i*4+8}:D{i*4+11})", num, float(quantity))
    summary.write('C30', '24时段合计', total)
    summary.write_formula('D30', '=SUM(D6:D29)', total, float(day.total))
    summary.write_row('A33', ['核对项目', '结果'], header)
    checks = [('原始点数', "=COUNT('原始96点'!D8:D103)", 96),
              ('96点总和', "=SUM('原始96点'!D8:D103)", float(day.total)),
              ('24点总和', '=D30', float(day.total)),
              ('总和差额', '=ROUND(B35-B36,12)', 0),
              ('核对结论', '=IF(AND(B34=96,B37=0),"一致","请检查")', '一致')]
    for r, (label, formula, cached) in enumerate(checks, 33):
        summary.write(r, 0, label)
        summary.write_formula(r, 1, formula, num if 34 <= r <= 36 else None, cached)
    summary.conditional_format('B37', {'type': 'cell', 'criteria': '!=', 'value': 0,
                                      'format': book.add_format({'bg_color': '#FEE2E2', 'font_color': '#B91C1C'})})
    summary.write('A40', '方法：时段1—96排序，每连续4点求和，不乘或除以0.25。')
    summary.write('A41', '单位沿用原表。显示三位小数，汇总采用原始精度。')
    raw.write('A2', f'{day.date} 原始96点电量', title)
    raw.write('A3', f'来源：{filename}，{day.sheet}，表头{day.headers}')
    raw.write('A4', '字段：省内现货日内 → 实际上网电量')
    raw.write('A5', day.unit or day.member)
    raw.write_row('A7', ['结算日期', '15分钟序号', '15分钟时段', '电量（原表单位）', '原表单元格'], header)
    for i, (quantity, cell) in enumerate(zip(day.points, day.cells)):
        raw.write_row(i + 7, 0, [day.date, i + 1, f'{clock(i*15)}—{clock((i+1)*15)}', float(quantity), f'{day.sheet}!{cell}'])
    raw.write('C104', '96点合计', total)
    raw.write_formula('D104', '=SUM(D8:D103)', total, float(day.total))
    raw.freeze_panes(7, 0)
    book.close()
    return output.getvalue()
