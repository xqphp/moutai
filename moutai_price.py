import requests,time,xlwt,os,base64
headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/537.36'}
today_time=time.strftime('%Y-%m-%d', time.localtime())

def moutai_data():
    url = 'aHR0cHM6Ly9wYXltZW50Lm5qeWh3bC50b3AvYXBpL3Byb3h5L3B1YmxpYy9hcGkvcHJpY2Uvc2VsZWN0UGFnZT9wYWdlPQ=='
    try:
        url=base64.b64decode(url).decode("utf-8")
        r=requests.get(url+'1',headers=headers,timeout=7)
        r.encoding=r.apparent_encoding
        r2 = requests.get(url + '2', headers=headers, timeout=7)
        r2.encoding = r2.apparent_encoding
        mt_data = r.json()['content'] +r2.json()['content']
        return  mt_data
    except Exception as e:
        print(f'接口失效!或者网络问题!\n{e}')

def moutai_Excel():
    if not os.path.exists(f'茅台今日行情-{today_time}.xls'):
        print('正在生成Excel请稍后*********')
        workBook = xlwt.Workbook(encoding='utf-8')
        sheet = workBook.add_sheet(f"今日茅台行情{today_time}")
        sheet.col(0).width = 400 * 20
        sheet.write(0, 0, f"今日茅台行情:{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())}")
        sheet.write(1, 0, '酒类')
        sheet.write(1, 1, '平均价(元)')
        sheet.write(1, 2, '昨日价(元)')
        mt_data=moutai_data()
        rowx=2
        for mt in mt_data:
            sheet.write(rowx, 0, mt['name'])
            sheet.write(rowx, 1, mt['averagePrice'])
            sheet.write(rowx, 2, mt['yesterdayPrice'])
            rowx+=1
        savePath=f'茅台今日行情-{today_time}.xls'
        workBook.save(savePath)
        print('生成Excel成功!')
    else:
        mt_data = moutai_data()
        print(f"Excel已经存在!\n今日茅台行情时间:{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())}")
        for i in mt_data:
            print(f"{i['name']} 平均:{i['averagePrice']}元 昨天:{i['yesterdayPrice']}元")

if __name__ == '__main__':
   moutai_Excel()



