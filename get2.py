import urllib.request
import json
from openpyxl import load_workbook
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

urlStart = "https://api.appmetrica.yandex.ru/stat/v1/data/?"
lang = "lang=ru&request_domain=ru"
ids="&id=1312957"
date1 = "&date1=2018-08-20"
date2 = "&date2=2018-08-26"
filters = "&filters=(ym%3Ace%3AoperatingSystemInfo%3D%3D'iOS')"
metrics = "&metrics=ym%3Ace%3Ausers%2Cym%3Ace%3Adevices%2Cym%3Ace%3AclientEvents"
dimensions ="&dimensions=ym%3Ace%3AeventLabel%2Cym%3Ace%3AparamsLevel1%2Cym%3Ace%3AparamsLevel2%2Cym%3Ace%3AparamsLevel3%2Cym%3Ace%3AparamsLevel4%2Cym%3Ace%3AparamsLevel5"
sort="&sort=-ym%3Ace%3AclientEvents"
limitAccur = "&limit=99999&accuracy=high&proposedAccuracy=true"
token ="&oauth_token=AQAAAAAfG7AEAASfc2O3zMIfhENQgzoavlpZq6A"

d1 = input("Введите начальную дату YYYY-MM-DD")
d2 = input("Введите конечную дату YYYY-MM-DD")

date1 = "&date1="+d1
date2 = "&date2="+d2

url = urlStart + lang + ids + date1 + date2 + filters + metrics + dimensions + sort + limitAccur + token

req = urllib.request.Request(url)
q = urllib.request.urlopen(req).read()

cont = json.loads(q.decode('utf-8'))

events = cont['data']

ws.cell (row=1, column = 1, value = "Название события")
for i in range (5):
    ws.cell (row = 1, column = i+2, value = "Параметр "+str(i+1))
ws.cell (row=1, column = 7, value = "Кол-во пользователей")
ws.cell (row=1, column = 8, value = "Кол-во устройств")
ws.cell (row=1, column = 9, value = "Кол-во событий")



for i in range(len(events)):
    for j in range(len(events[i]['dimensions'])):

        ws.cell (row = i + 1 + 1, column = j+1, value = events[i]['dimensions'][j]['name'])
    for j in range(len(events[i]['metrics'])):
        ws.cell (row = i + 1 + 1, column = j+1+6, value = events[i]['metrics'][j])               

wb.save("file2.xlsx")
