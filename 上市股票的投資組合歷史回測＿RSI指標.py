#!/usr/bin/env python
# coding: utf-8


#股票資訊爬蟲
import xlwings as xw
import requests
import pandas as pd
import numpy as np
import datetime
import json
import time

from notebook.services.config import ConfigManager
cm = ConfigManager().update('notebook', {'limit_output': 10000000})

#將輸入的起始日期與結束日期中間的所有年月製成一個list
def dateRange(beginDate, endDate):
    dates = []
    dt = datetime.datetime.strptime(beginDate, "%Y%m")
    date = beginDate[:]
    while date <= endDate:
        dates.append(date)
        dt = dt + datetime.timedelta(1)
        date = dt.strftime("%Y%m")
    return dates

def monthRange(beginDate, endDate):
    monthSet = set()
    for date in dateRange(beginDate, endDate):
        monthSet.add(date[0:7])
    monthList = []
    for month in monthSet:
        monthList.append(month)
    return sorted(monthList)
#將資料從url中擷取下來
def get_stock_history(d, sn):
    url="https://www.twse.com.tw/exchangeReport/STOCK_DAY?date=%s&stockNo=%s" %(d, sn)
    r=requests.get(url)
    datas=r.json()
    time.sleep(6)
    return transform(datas["data"])
#轉換資料格式
def transform_data(data):
    print(data)
    data[0]=datetime.datetime.strptime(transform_date(data[0]), "%Y/%m/%d")
    data[1]=int(data[1].replace(",", ""))
    data[2]=int(data[2].replace(",", ""))
    data[3]=float(data[3].replace(",", ""))
    data[4]=float(data[4].replace(",", ""))
    data[5]=float(data[5].replace(",", ""))
    data[6]=float(data[6].replace(",", ""))
    data[7]=float(0.0 if data[7].replace(",", "") =="X0.00" else data[7].replace(",", ""))
    data[8]=int(data[8].replace(",", ""))
    return data

def transform_date(date):
    y,m,d=date.split("/")
    return str(int(y)+1911)+"/"+m+"/"+d

def transform(data):
    return[transform_data(d) for d in data]

wb=xw.Book()
s=str(input("請輸入股票代碼 EX:2330,0050,2809："))
stock_symbol=s.split(",")
start_date=str(input("起始年月 EX:202001 "))
end_date=str(input("結束年月 EX:202101 "))
month_range=monthRange(start_date, end_date)
date=[f"{ym}01"  for ym in month_range]
info=[]
for i in stock_symbol:
    stock_no=[i]*len(month_range)
    d_sn=list(zip(date, stock_no))
    ws=wb.sheets.add(name=f"TW{i}")
    quotes=[]    
    for d, sn in d_sn:
        df_s=get_stock_history(d, sn)
        quotes.extend(df_s)
    ws.range("A1").value=["date","shares","amount","open","high","low","close","change","turnover"]
    ws.range("A2").value=quotes

wb.save("stock_information")


# In[15]:


nwb=xw.Book()
portfolio=nwb.sheets.add(name="portfolio")
portfolio.range("K2").value=["起始資金", 10000000]
portfolio.range("K3").value=["交易股數", 1000]

for i in stock_symbol:
    sheet=wb.sheets[f"TW{i}"]
    last_row = sheet.range("A1").end("down").row
    dates = sheet.range(f"A1:A{last_row}").options(ndim=2).value
    closes=sheet.range(f"G1:G{last_row}").options(ndim=2).value
    nws=nwb.sheets.add(name=f"{i}_strategy")
    nws.range("A1").value=dates
    nws.range("B1").value=closes
    nws.range("C1").value=["price change", "long rsi", "short rsi","買入股數","賣出股數","持有股數","持有資金","總資產"]
    #long rsi
    for n in range(3, last_row+1):
        price_change=(nws.range(f"B{n}").value-nws.range(f"B{n-1}").value)*100/nws.range(f"B{n-1}").value
        nws.range(f"C{n}").value=price_change
        if nws.range(f"C{n}").value>0:
            nws.range(f"C{n}").color=(255,0,0)
        elif nws.range(f"C{n}").value<0:
            nws.range(f"C{n}").color=(0,255,0)
        else:
            continue
    last_row = nws.range("A1").end("down").row
    lis=[]
    day_rsi_index=[]
    price_rise=[]
    price_fall=[]
    for i in range(3,last_row-12):
        day_pctchange=nws.range(f"C{i}:C{i+13}").value 
        lis.extend(day_pctchange)
        for i in lis:
            if i>0:
                price_rise.append(i)
            elif i<0:
                price_fall.append(i)
            else:
                continue
        day_rise_a=sum(price_rise)/14
        day_fall_a=sum(price_fall)/14*-1
        rsi_index=day_rise_a/(day_rise_a+day_fall_a)*100
        day_rsi_index.append(rsi_index)
    num=0
    for c in range(16, last_row+1):
        nws.range(f"D{c}").value=day_rsi_index[num]
        num+=1
    #short rsi    
    for n in range(3, last_row+1):
        price_change=(nws.range(f"B{n}").value-nws.range(f"B{n-1}").value)*100/nws.range(f"B{n-1}").value
        nws.range(f"C{n}").value=price_change
    last_row = nws.range("A1").end("down").row
    lis=[]
    day_rsi_index=[]
    price_rise=[]
    price_fall=[]
    for i in range(3,last_row-5):
        day_pctchange=nws.range(f"C{i}:C{i+6}").value 
        lis.extend(day_pctchange)
        for i in lis:
            if i>0:
                price_rise.append(i)
            elif i<0:
                price_fall.append(i)
            else:
                continue
        day_rise_a=sum(price_rise)/7
        day_fall_a=sum(price_fall)/7*-1
        rsi_index=day_rise_a/(day_rise_a+day_fall_a)*100
        day_rsi_index.append(rsi_index)
    num=0
    for c in range(9, last_row+1):
        nws.range(f"E{c}").value=day_rsi_index[num]
        num+=1
    #caluculate first day
    long_rsi = nws.range(f"D16").value
    short_rsi = nws.range(f"E16").value
    shares = portfolio.range("L3").value
    if short_rsi > long_rsi:
        nws.range(f"F16").value = shares
    else:
        nws.range(f"F16").value = 0
    if short_rsi < long_rsi:
        nws.range(f"G16").value = 0
    else:
        nws.range(f"G16").value = 0

    nws.range("H16").value = nws.range("F16").value - nws.range("G16").value
    nws.range("I16").value = portfolio.range("L2").value - nws.range("H16").value * nws.range("B16").value
    nws.range("J16").value = nws.range("I16").value + nws.range("H16").value * nws.range("B16").value
    #rsi strategy
    for i in range(17,last_row+1):
        short_term_rsi=nws.range(f"E{i}").value
        long_term_rsi=nws.range(f"D{i}").value
        price_today=nws.range(f"B{i}").value
        if (short_term_rsi>long_term_rsi) and (nws.range(f"I{i-1}").value>=price_today*1000):
            nws.range(f"F{i}").value=1000
        else:
            nws.range(f"F{i}").value=0
        if (short_term_rsi<long_term_rsi) and (nws.range(f"H{i-1}").value>=shares):
            nws.range(f"G{i}").value=1000
        else:
            nws.range(f"G{i}").value=0    
        nws.range(f"H{i}").value=nws.range(f"H{i-1}").value+nws.range(f"F{i}").value-nws.range(f"G{i}").value
        nws.range(f"I{i}").value=nws.range(f"I{i-1}").value-(nws.range(f"F{i}").value-nws.range(f"G{i}").value)*price_today
        nws.range(f"J{i}").value=nws.range(f"I{i}").value+nws.range(f"H{i}").value*price_today
    

#檢視投資組合收益
portfolio.range("A1").value="股票代碼"
portfolio.range("B1").value="投資收益"
num=2
for i in stock_symbol:
    portfolio.range(f"A{num}").value=i
    portfolio.range(f"B{num}").value=nwb.sheets[f"{i}_strategy"].range(f"J{last_row}").value-nwb.sheets[f"{i}_strategy"].range("J16").value
    if portfolio.range(f"B{num}").value>0:
        portfolio.range(f"B{num}").color=(255,0,0)
    elif portfolio.range(f"B{num}").value<0:
        portfolio.range(f"B{num}").color=(0,255,0)
    else:
        portfolio.range(f"B{num}").value=0
    num+=1
last_cell=portfolio.range("B1").end("down").row
portfolio.range(f"B{last_cell+1}").value=sum(portfolio.range(f"B2:B{last_cell}").value)
portfolio.range("K4").value=["總收益", portfolio.range(f"B{last_cell+1}").value]
nwb.save("rsistrategy")


#line notify提示總收益
content = "rsi strategy 投資收益："+ str(portfolio.range(f"B{last_cell+1}").value)
line_url = "https://notify-api.line.me/api/notify"
token = "E0u3UasxeGuNXuk7rX66W24ygb9sGds8mCCKidORnBm"
headers = {"Authorization": "Bearer " + token}
payload = {'message': content}
r = requests.post(line_url, headers = headers, params = payload)

