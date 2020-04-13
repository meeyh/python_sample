import requests
import json
import re
import xlwt
import xlrd

# 函数，用于过滤非中文，有待优化
def dataTostr(data):
    if len(data) == 0:
        return ''
    else:
        datastr = ''
        for i in data:
            datastr += i
        return datastr

# 创建一个Excel表格，用来存数据
wordbook = xlwt.Workbook(encoding='utf-8')
wordsheet = wordbook.add_sheet('留言', cell_overwrite_ok=True)
header = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.162 Safari/537.36'
}
# cookie内容是浏览器里已登录QQ空间的账号身份信息，完整cookie自行替换自己的
cookie = {
    'Cookie': 'pgv_pvi=...'
}
# urls 空间留言板的数据接口链接
# 参数start是开始页码，改为start={0}循环中动态变化
# 完整链接，请复制自己登录空间查看留言板得到的链接替换
urls = 'https://user.qzone.qq.com/proxy/domain/m.qzone.qq.com/cgi-bin/new/get_msgb?start={0}'
current_row = 0    # Excel表格当前行号，0代表在第一行
for i in range(67):    # 67替换 是自己的留言板最大的页码
    url = urls.format(i*10)
    response = requests.get(url, headers=header, cookies=cookie)
    data = response.text.replace('_Callback(', '').replace(');', '')
    results = json.loads(data)['data']['commentList']
    for n, g in enumerate(results):
        try:
            uin = g['uin']  # QQ号
        except:
            continue
        pubtime = g['pubtime'].split(' ')[0]    # 日期
        Ttime = g['pubtime'].split(' ')[1]      # 时间
        nickname = g['nickname']
        htmlContent = g['htmlContent']
        patt = re.compile(r'[\u4e00-\u9fa5]+')
        Content = re.findall(patt, htmlContent)
        geststr = dataTostr(Content)
        wordsheet.write(current_row + n, 0, uin)        # QQ号
        wordsheet.write(current_row + n, 1, nickname)   # 昵称
        wordsheet.write(current_row + n, 2, pubtime)    # 日期
        wordsheet.write(current_row + n, 3, Ttime)      # 时间
        wordsheet.write(current_row + n, 4, geststr)    # 留言
    current_row += 10   # 留言板一页有10条记录，+10改变Excel表格当前行号
wordbook.save('gestbook.xls')
print('done...')
