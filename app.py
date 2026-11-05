import pandas as pd
import streamlit as st
from energy import BillError, clock, export, extract

st.set_page_config(page_title='亚华 · 实际上网电量', page_icon='⚡', layout='wide')
st.title('亚华 · 实际上网电量汇总')
st.caption('日清分账单上传后自动计算：96个15分钟点，每连续4点求和为24小时。')

with st.expander('识别与计算口径'):
    st.markdown('''- 只提取**省内现货日内 → 实际上网电量**，按表头定位，支持工作表改名、表头移动及列重排。
- 日期取自账单“结算日期”。不同日期、市场成员及结算单元分别处理，请选择要导出的记录。
- 按时段1—96排序。第1—4点对应00:00—01:00，末4点对应23:00—24:00。
- 电量直接相加，单位沿用原表。显示三位小数，计算保留原始精度。
- 缺点、重复点、空白或非数字电量会停止计算并显示位置。''')

upload = st.file_uploader('上传电厂侧日清分账单', type=['xlsx'])
if upload is None:
    st.info('请选择.xlsx账单。计算完成后可直接下载新的Excel。')
    st.stop()

try:
    with st.spinner('正在识别账单并核对96点…'):
        days = extract(upload.getvalue())
    index = st.selectbox('选择结算日期及结算单元', range(len(days)), format_func=lambda i: days[i].label)
    day = days[index]
    excel = export(day, upload.name)
except BillError as exc:
    st.error(str(exc))
    st.stop()
except Exception:
    import logging
    logging.exception('账单处理失败')
    st.error('账单处理失败，请确认文件未损坏，并保留原文件用于排查。')
    st.stop()

st.success(f'核对通过：96点总和 = 24小时总和 = {day.total:,.3f}，差额为0。')
c1, c2, c3 = st.columns(3)
c1.metric('结算日期', day.date)
c2.metric('有效点数', len(day.points))
c3.metric('全天电量（原表单位）', f'{day.total:,.3f}')

df24 = pd.DataFrame({'小时时段': [f'{clock(i*60)}—{clock((i+1)*60)}' for i in range(24)],
                     '电量': [float(v) for v in day.hours]})
st.subheader('24时段结果')
left, right = st.columns([1, 2])
with left:
    st.dataframe(df24.style.format({'电量': '{:.3f}'}), hide_index=True, use_container_width=True)
with right:
    st.line_chart(df24.set_index('小时时段'), y='电量')

st.download_button('下载Excel（24时段汇总＋原始96点＋总量校验）', excel,
                   file_name=f'{day.date}_实际上网电量_96点转24时段.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')
with st.expander('查看原始96点及数据来源'):
    st.write(f'工作表：{day.sheet}；表头：{day.headers}；数据：{day.cells[0]}—{day.cells[-1]}')
    st.dataframe(pd.DataFrame({'时段': range(1,97), '电量': [float(v) for v in day.points],
                              '原表单元格': day.cells}).style.format({'电量': '{:.3f}'}),
                 hide_index=True, use_container_width=True)
