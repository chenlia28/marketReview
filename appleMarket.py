# -*- encoding: utf-8 -*-
# ---
# @Software: PyCharm
# @File: appleMarket.py
# @Author: Liang.Chen
# @Site: 
# @Time: 10月 20, 2021
# ---

import urllib.request
import json
import xlsxwriter
# import sys
# import jenkinsapi
# from jenkinsapi.jenkins import Jenkins


# TODO
"""
#1 apple review
#2 google review
#3 评分按等级进行过滤
#4 集成到jenkins自动发报告
"""

# appid = sys.argv[1]
# country = sys.argv[2]
# lauguage = sys.argv[3]

page = 1
country = 'us'
language = 'en'
appid=input("Please enter your app's appid:");

# appid = 414478124


workbook = xlsxwriter.Workbook('App评论.xlsx')
worksheet = workbook.add_worksheet()

format = workbook.add_format()
format.set_border(1)
format.set_border(1)
format_title = workbook.add_format()
format_title.set_border(1)
format_title.set_bg_color('#cccccc')
format_title.set_align('left')
format_title.set_bold()

title = ['version', 'rating', 'time', 'name', 'title', 'content']
worksheet.write_row('A1', title, format_title)
row = 1
col = 0
count = 0

#每次50条数据，循环10次就是500条
while page < 11:

    myurl = "https://itunes.apple.com/rss/customerreviews/page=" + str(page) + \
        "/id=" + str(appid) + "/sortby=mostrecent/json?l=" + str(language) + "&&cc=" + str(country)

    response = urllib.request.urlopen(myurl)
    myjson = json.loads(response.read().decode())
    # print(myurl)
    # print(myjson)
    print("processing..., please wait a moment ======>"+str(page*10)+"%")
    if "entry" in myjson["feed"]:
        count += len(myjson["feed"]["entry"])

        # 循环写入第1列：version
        row = 1 + (page - 1) * 50
        for i in myjson["feed"]["entry"]:
            worksheet.write(row, col, i["im:version"]["label"], format)
            row += 1

        # 循环写入第2列：rating
        row = 1 + (page - 1) * 50
        for i in myjson["feed"]["entry"]:
            worksheet.write(row, col + 1, i["im:rating"]["label"], format)
            row += 1

        # 循环写入第3列：time
        row = 1 + (page - 1) * 50
        for i in myjson["feed"]["entry"]:
            worksheet.write(row, col + 2, i["updated"]["label"], format)
            row += 1

        #循环写入第4列：name
        row = 1 + (page - 1) * 50
        for i in myjson["feed"]["entry"]:
            worksheet.write(row,col + 3,i["author"]["name"]["label"],format)
            row += 1

        #循环写入第5列：title
        row=1+(page-1)*50
        for i in myjson["feed"]["entry"]:
            worksheet.write(row,col + 4,i["title"]["label"],format)
            row += 1

        #循环写入第6列：内容content
        row=1+(page-1)*50
        for i in myjson["feed"]["entry"]:
            worksheet.write(row,col + 5,i["content"]["label"],format)
            row += 1


        page = page + 1
        row = (page-1)*50+1
    else:
        print("processing..., please wait a moment ======>100%")
        break

if count == 0:
    print("\n Finished, but there was no data, please check it again!!！")

else:
    print("\n ======> Congratulations!!! \n please check the excel file, you have got "+str(count)+" reviews!!!")

workbook.close()