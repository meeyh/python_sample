import pandas as pd
import jieba
from pyecharts.charts import Bar
from pyecharts import options as opts
from pyecharts.charts import WordCloud
from collections import Counter

# 请先打开Excel表格加上这个表头，再执行这个代码
df = pd.read_excel('gestbook.xls', encoding='gbk', usecols=['qq号', '昵称', '日期', '时间', '内容'])

# 数据可视化，柱状图
qq_nums = []
nicknames = []
comment_times = []
for i in list(set(df['qq号'])):
    qq_nums.append(i)
    nicknames.append(df[df['qq号'].eq(i)]['昵称'].values[0])
    comment_times.append(len(df[df['qq号'].eq(i)]))
data = {'qq号': qq_nums, '昵称': nicknames, '次数': comment_times}
new_dataframe = pd.DataFrame(data).sort_values(by='次数', ascending=False).head(20)
# 自定义x轴文字替代完整昵称，保护隐私
leb = ['微*', '@一颗 り*']
bar = (
    Bar()
    .add_xaxis(list(new_dataframe['昵称']))
    .add_yaxis("留言次数", list(new_dataframe['次数']))
    .set_series_opts(label_opts=opts.LabelOpts(is_show=True))
    .set_global_opts(
        xaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=-45)),
        title_opts=opts.TitleOpts(title="好友留言次数", subtitle='前20位')
    )
)
bar.render('好友留言次数前20排名.html')
print('数据可视化，柱状图完成...')

# 数据可视化，词云图
dff = df['内容']
comment = ''
for i in dff.values:
    comment += str(i)
# text = ' '.join(jieba.cut(comment.replace('nan', ''), cut_all=False))
text = ' '.join(jieba.cut_for_search(comment.replace('nan', '')))
all_words = text.split()

c = Counter()
for x in all_words:
    if len(x) > 1 and x != '\r\n':  # \r\n 注意不同的系统，不一样
        c[x] += 1

lis = []
for (k, v) in c.most_common():
    tup = (k, v)
    lis.append(tup)
wdc = (
    WordCloud()
    .add(series_name="留言", data_pair=lis, word_size_range=[20, 66])
    .set_global_opts(
        title_opts=opts.TitleOpts(
            title="留言分析", title_textstyle_opts=opts.TextStyleOpts(font_size=18)
        ),
        tooltip_opts=opts.TooltipOpts(is_show=True),
    )
)
wdc.render('留言内容分析.html')
print('数据可视化，词云图完成...')
