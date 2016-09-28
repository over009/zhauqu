# coding=utf-8
'''
author:Chenwentao
date:2016-2-26
function:
抓取指定网页数据
'''
import urllib
import urllib2
import re
import xlwt
import sys
# 头设置
loginHeaders = {
    'Host':'www.dce.com.cn',
    'User-Agent' : 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:35.0) Gecko/20100101 Firefox/35.0',
    'Referer' : 'http://www.dce.com.cn/PublicWeb/MainServlet?action=Pu00011_search',
    'Content-Type': 'application/x-www-form-urlencoded',
    'Connection' : 'Keep-Alive'
}

# 参数设置
post = {
    'action':'Pu00011_result',
     #通过这个参数修改时间
    'Pu00011_Input.trade_date':'20160329',
    'Pu00011_Input.variety':'all',
    'Pu00011_Input.trade_type':'0'
}

reload(sys)
sys.setdefaultencoding('utf8')

url='http://www.dce.com.cn/PublicWeb/MainServlet'
postData = urllib.urlencode(post)
request = urllib2.Request(url, postData, loginHeaders)
opener = urllib2.build_opener()
response = opener.open(request)
content = response.read().decode('gbk')
res = r'<td>&nb|<td.*?nowrap.*?>(.*?)</td>'
m = re.findall(res,content,re.S)
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('dalian', cell_overwrite_ok=True)
for i in range(2506):
    sheet.write(i / 14, i % 14, m[i])
    book.save(r'e:\dalianqh.xls')












