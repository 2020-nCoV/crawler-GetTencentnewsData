"""
A script to parse data on the given url and save the serialized result into excel.

@Author : flamywhale
@Date   : 2020/1/31
"""
import json
import requests
import re
import time
from pandas import ExcelWriter, DataFrame

URL = 'https://view.inews.qq.com/g2/getOnsInfo?name=disease_h5&callback=jQuery34108760783335638791_1580372739739&_=1580372739740'

# use session to download data from the website
sess = requests.session()
crawledData = sess.get(URL)

# use regular expression to match useful data
group = re.search('\"areaTree\":\[\S+\]', crawledData.text.replace('\\', ''))
data = json.loads('{' + group[0] + '}')

# main process to serialize data
rows = []
for country in data['areaTree']:
    if country['name'] == '中国':
        for province in country['children']:
            for city in province['children']:
                row = {'region': province['name'] + city['name']}
                for key, total in city['total'].items():
                    row['total_' + key] = total
                for key, today in city['today'].items():
                    row['today_' + key] = today
                rows.append(row)
    else:
        row = {'region': country['name']}
        for key, total in country['total'].items():
            row['total_' + key] = total
        for key, today in country['today'].items():
            row['today_' + key] = today
        rows.append(row)

# create data frame from list and save it to excel
df = DataFrame(rows)
with ExcelWriter(f'result-{time.strftime("%m-%d-%H：%M")}.xlsx') as writer:
    df.to_excel(writer)
