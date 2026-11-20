import logging
import hashlib

import pandas as pd
import streamlit as st
from energy import BillError, clock, export, extract
from price import PriceError, export_prices, extract_prices
from power import PowerError, extract_power, export_power


st.set_page_config(page_title='亚华 · 电量与电价96转24',page_icon='⚡',layout='wide')
st.title('亚华 · 96时段转24时段')
mode = st.sidebar.radio('选择转换功能',['实际上网电量','电价（日前／实时）','出力／联络线'])
st.sidebar.caption('电量：每4点求和\n\n电价：每4点取平均\n\n出力／联络线：每4点取平均')


def energy_page():
    st.header('实际上网电量')
    st.caption('日清分账单：每连续4个15分钟电量求和，生成24小时电量。')
    with st.expander('识别与计算口径'):
        st.markdown('''- 只提取**省内现货日内 → 实际上网电量**，按实际表头定位。
- 日期取自账单“结算日期”，不同日期及结算单元分别处理。
- 时段1—96排序，第1—4点对应00:00—01:00，末4点对应23:00—24:00。
- 电量直接相加，沿用原单位。显示三位小数，计算保留原始精度。
- 缺点、重复点、空白或非数字电量会停止计算。''')
    upload = st.file_uploader('上传电厂侧日清分账单',type=['xlsx'],key='energy_upload')
    if upload is None:
        st.info('请选择.xlsx账单。计算完成后可直接下载新的Excel。')
        return
    try:
        with st.spinner('正在识别账单并核对96点…'):
            days = extract(upload.getvalue())
        index = st.selectbox('选择结算日期及结算单元',range(len(days)),format_func=lambda i:days[i].label)
        day = days[index]
        excel = export(day,upload.name)
    except BillError as exc:
        st.error(str(exc))
        return
    except Exception:
        logging.exception('账单处理失败')
        st.error('账单处理失败，请确认文件未损坏，并保留原文件用于排查。')
        return
    st.success(f'核对通过：96点总和 = 24小时总和 = {day.total:,.3f}，差额为0。')
    c1,c2,c3=st.columns(3)
    c1.metric('结算日期',day.date)
    c2.metric('有效点数',len(day.points))
    c3.metric('全天电量（原表单位）',f'{day.total:,.3f}')
    data=pd.DataFrame({'小时时段':[f'{clock(i*60)}—{clock((i+1)*60)}' for i in range(24)],
                       '电量':[float(v) for v in day.hours]})
    st.subheader('24时段结果')
    left,right=st.columns([1,2])
    with left:
        st.dataframe(data.style.format({'电量':'{:.3f}'}),hide_index=True,use_container_width=True)
    with right:
        st.line_chart(data.set_index('小时时段'),y='电量')
    st.download_button('下载电量Excel（24时段＋原始96点＋总量校验）',excel,
                       file_name=f'{day.date}_实际上网电量_96点转24时段.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
    with st.expander('查看原始96点及数据来源'):
        st.write(f'工作表：{day.sheet}；表头：{day.headers}；数据：{day.cells[0]}—{day.cells[-1]}')
        st.dataframe(pd.DataFrame({'时段':range(1,97),'电量':[float(v) for v in day.points],
                                  '原表单元格':day.cells}),hide_index=True,use_container_width=True)


def price_page():
    st.header('电价（日前／实时）')
    st.caption('上传机组出清结果表：每连续4个15分钟电价取平均，生成24小时电价。')
    with st.expander('电价转换规则',expanded=True):
        st.markdown('''- **小时电价 =（第1点＋第2点＋第3点＋第4点）÷4**，各点均为15分钟。
- 按表内日期及时间识别。00:15、00:30、00:45、01:00归入00:00—01:00。
- 沿用原表价格单位，未标注时不猜测。0价、负价均保留；空白不当作0。
- 这是**时间平均电价**，不按电量加权。不同文件、日期、机组及日前／实时市场分别展示。
- 缺点、重复点、无效电价会报错。下载包含24小时电价、原始96点电价和均价核对。''')
    uploads=st.file_uploader('上传日前／实时市场机组出清结果（可多选）',type=['xlsx'],
                             accept_multiple_files=True,key='price_upload')
    if not uploads:
        st.info('可同时上传日前和实时表，分别选择查看、下载。')
        return
    records=[]
    with st.spinner('正在识别电价并核对96点…'):
        for upload in uploads:
            try:
                records.extend((upload.name,day) for day in extract_prices(upload.getvalue()))
            except PriceError as exc:
                st.error(f'{upload.name}：{exc}')
            except Exception:
                logging.exception('电价处理失败')
                st.error(f'{upload.name}：处理失败，请确认文件未损坏。')
    if not records:
        return
    index=st.selectbox('选择文件、日期及市场',range(len(records)),
                       format_func=lambda i:f'{records[i][0]} · {records[i][1].label}',key='price_record')
    filename,day=records[index]
    try:
        excel=export_prices(day,filename)
    except Exception:
        logging.exception('电价导出失败')
        st.error('电价导出失败，请保留原文件用于排查。')
        return
    st.success('核对通过：96点完整；24个小时均价的平均值与原始96点全天均价一致。')
    a,b,c=st.columns(3)
    a.metric('运行日期',day.date)
    b.metric('市场',day.market)
    c.metric('全天时间均价',f'{day.mean:,.4f}')
    st.caption(f'价格单位：{day.price_unit}。显示4位小数，计算使用原始精度。')
    data=pd.DataFrame({'小时时段':[f'{clock(i*60)}—{clock((i+1)*60)}' for i in range(24)],
                       '小时平均电价':[float(v) for v in day.hours]})
    st.subheader('24时段电价')
    left,right=st.columns([1,2])
    with left:
        st.dataframe(data.style.format({'小时平均电价':'{:.4f}'}),hide_index=True,use_container_width=True)
    with right:
        st.line_chart(data.set_index('小时时段'),y='小时平均电价')
    st.download_button('下载电价Excel（24时段＋原始96点＋均价核对）',excel,
                       file_name=f'{day.date}_{day.market}_电价96转24时段.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
    with st.expander('查看原始96点电价及数据来源'):
        st.write(f'工作表：{day.sheet}；电价表头：{day.header}')
        st.dataframe(pd.DataFrame({'结束时间':[clock((i+1)*15) for i in range(96)],
                                  '电价':[float(v) for v in day.points],'原表单元格':day.cells}),
                     hide_index=True,use_container_width=True)


def power_page():
    st.header('出力／联络线96转24')
    st.caption('支持新能源、非市场化机组出力、省间联络线横向96点表，也支持结构明确的无表头文件。结果竖排，Excel无底色、黑色文字。')
    method = 'mean'
    st.info('每小时值 = 对应4个15分钟值之和 ÷ 4。固定取算术平均，适用于小时平均功率／出力；单位沿用原表。')
    st.caption('不根据文件名猜测功率或电量。负数、零值保留；空白、重复或缺失时段会停止该文件的计算。')
    uploads = st.file_uploader('上传出力／联络线Excel（可多选）', type=['xlsx'], accept_multiple_files=True, key='power_upload')
    if not uploads:
        st.info('支持带时段表头的横向96点表；无表头文件会先显示识别范围，确认时间顺序后再转换。')
        return
    records = []
    for upload in uploads:
        try:
            records.extend((upload.name, group) for group in extract_power(upload.getvalue(), upload.name))
        except PowerError as exc:
            st.error(f'{upload.name}：{exc}')
        except Exception:
            logging.exception('出力／联络线读取失败')
            st.error(f'{upload.name}：读取失败，请核对原文件。')
    if not records:
        return
    selected = st.selectbox('选择文件、日期及单元', range(len(records)),
                            format_func=lambda i: f'{records[i][0]} · {records[i][1].label}', key='power_record')
    filename, group = records[selected]
    if group.inferred:
        st.warning('此文件没有时段表头。日期、名称与96点范围按内容识别，不固定列号；请核对名称、范围和原始时间顺序。')
        preview = [{'名称': name, '日期': group.date,
                    '数值范围': f'{group.cells[name][0]}—{group.cells[name][-1]}',
                    '前4点': ' / '.join(str(v) for v in points[:4]),
                    '后4点': ' / '.join(str(v) for v in points[-4:])}
                   for name, points in group.curves.items()]
        st.dataframe(pd.DataFrame(preview), hide_index=True, use_container_width=True)
        fingerprint = hashlib.sha256(repr((filename, group.date, group.identifier, group.sheet, group.curves, group.cells)).encode()).hexdigest()
        if not st.checkbox('我确认：以上96点从左到右依次对应00:15至24:00，每点15分钟', key='power_confirm_'+fingerprint):
            return
    values = group.hours(method)
    data = pd.DataFrame({'时段（结束时刻）': [clock((h+1)*60) for h in range(24)], '日期': [group.date]*24,
                         **{name: [float(v) for v in points] for name, points in values.items()}})
    st.success(f'96点完整，共{len(values)}条曲线，已生成24时段竖版结果。')
    st.caption('01:00代表00:00—01:00；24:00代表23:00—24:00。显示3位小数，计算保留原始精度。')
    st.dataframe(data.style.format({name: '{:.3f}' for name in values}), hide_index=True, use_container_width=True)
    try:
        excel = export_power(group, filename, method)
    except Exception:
        logging.exception('出力／联络线导出失败')
        st.error('Excel导出失败，请保留原文件用于排查。')
        return
    suffix = '平均' if method == 'mean' else '求和'
    stem = filename.rsplit('.', 1)[0]
    st.download_button('下载24时段竖版Excel（无底色）', excel,
                       file_name=f'{stem}_{group.date}_24时段_{suffix}_竖版_无底色.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')
    with st.expander('复制到Excel（选中下方文本后复制）'):
        st.text_area('制表符分隔文本', data.to_csv(index=False, sep='\t', float_format='%.3f'), height=260, key='power_copy')
    with st.expander('原始96点及来源'):
        st.write(f'工作表：{group.sheet}；单元标识：{group.identifier or "原表未标"}')
        st.dataframe(pd.DataFrame({'结束时刻': [clock((p+1)*15) for p in range(96)],
                                   **{name: [float(v) for v in points] for name, points in group.curves.items()}}),
                     hide_index=True, use_container_width=True)


if mode=='实际上网电量':
    energy_page()
elif mode=='电价（日前／实时）':
    price_page()
else:
    power_page()
