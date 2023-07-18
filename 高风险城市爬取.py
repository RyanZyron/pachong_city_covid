
from scrapy.selector import Selector
import requests
import pycurl
import json
from io import BytesIO
import xlwings as xw
from datetime import datetime, timedelta
import pandas as pd
import datetime
app = xw.App()
workbook = app.books.open('/Users/ze/Desktop/python/zad/高风险地区表数据.xlsx')
worksheet = workbook.sheets['Sheet1']

cookies = {
    '__yjs_duid': '1_806f47a24780f7ed04c33ebf92148f3d1650262951915',
}

headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    # 'Cookie': '__yjs_duid=1_806f47a24780f7ed04c33ebf92148f3d1650262951915',
    'Referer': 'https://cn.bing.com/',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36 Edg/107.0.1418.68',
}

response = requests.get(
    'http://sz.bendibao.com/news/gelizhengce/fengxianmingdan.php',
    cookies=cookies,
    headers=headers,
    verify=False,
)
selector = Selector(response)
tr = selector.css(".info-list .detail-message .top-title")
all_value = []
date = datetime.date.today()
for x in tr:
    province = x.css('span span:nth-child(1)::text').extract_first().strip()
    city = x.css('span span:nth-child(2)::text').extract_first()
    if city is None:
        city = province
    else:
        city = city.strip()
    amount = x.css('.gao ::text').extract_first()
    if amount is None:
        continue
    else:
        amount = amount.strip()
    row_value = [date, province, city, amount]
    all_value.append(row_value)
print(all_value)
nrow = worksheet.used_range.shape
worksheet.range((nrow[0]+1, 1)).value = all_value

workbook.save('/Users/ze/Desktop/python/zad/高风险地区表数据.xlsx')
workbook.close()
app.quit()