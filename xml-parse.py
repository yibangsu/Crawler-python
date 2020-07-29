#!/usr/bin/python
# -*- coding: UTF-8 -*- 

import requests
import bs4 
import xlwt

# 爬xml
url = "https://www.amazon.com/gp/bestsellers/pc/12879431/ref=pd_zg_hrsr_pc"
kv = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_3) AppleWebKit/601.4.4 (KHTML, like Gecko) Version/9.0.3 lang="cn"'}

try:
  response = requests.get(url, headers=kv)
  response.raise_for_status()
  response.encoding = response.apparent_encoding
  file=open("web-get.html","w+", encoding='utf-8')
  file.write(response.text)
  file.close()
  # print(response.text)
except:
  print("web get error: "+ str(response.status_code))
  exit


# 打开xml
soup = bs4.BeautifulSoup(response.text,'lxml')
ol = soup.find('ol')
li_list = ol.find_all('li',{'class':'zg-item-immersion'})
line = len(li_list)
print("line is:", line)

# 创建一个workbook 设置编码
workbook = xlwt.Workbook(encoding = 'utf-8')
# 创建一个worksheet
worksheet = workbook.add_sheet('Worksheet')
# 标题 - 产品
worksheet.write(0, 0, 'product')
# 标题 - 连接
worksheet.write(0, 1, 'link')
# 标题 - 标号
worksheet.write(0, 2, 'num')
# 标题 - 图片
worksheet.write(0, 3, 'image')
# 标题 - 星级
worksheet.write(0, 4, 'star')
# 标题 - 星级统计
worksheet.write(0, 5, 'star count')
# 标题 - 价格
worksheet.write(0, 6, 'price')
# 标题 - 版本数
worksheet.write(0, 7, 'version count')

print("worksheet start")

for i in  range(0, line):
  print("at", i)
  # 产品
  try:
    str_get = li_list[i].find('div',{'aria-hidden':'true'}).contents
    worksheet.write(i+1, 0, str_get)
    print("at",i,0,"product",str_get)
  except:
    print("at",i,0,"no product")
  # 连接
  try:
    str_get = li_list[i].find('a',{'class':'a-link-normal'})['href']
    worksheet.write(i+1, 1, "https://www.amazon.com/"+str_get)
    print("at",i,1,"link",str_get)
  except:
    print("at",i,1,"no link")
  # 标号
  try:
    str_get = li_list[i].find('span',{'class':'zg-badge-text'}).contents
    worksheet.write(i+1, 2, str_get)
    print("at",i,2,"num",str_get)
  except:
    print("at",i,2,"no num")
  # 图片
  try:
    str_get = li_list[i].find('img')['src']
    worksheet.write(i+1, 3, str_get)
    print("at",i,3,"image",str_get)
  except:
    print("at",i,3,"no image")
  # 星级
  try:
    str_get = li_list[i].find('span',{'class':'a-icon-alt'}).contents
    worksheet.write(i+1, 4, str_get)
    print("at",i,4,"star",str_get)
  except:
    print("at",i,4,"no star")
  # 星级统计
  try:
    str_get = li_list[i].find('a',{'class':'a-size-small a-link-normal'}).contents
    worksheet.write(i+1, 5, str_get)
    print("at",i,5,"star count",str_get)
  except:
    print("at",i,5,"no star count")
  # 价格
  try:
    str_get = li_list[i].find('span',{'class':'p13n-sc-price'}).contents
    worksheet.write(i+1, 6, str_get)
    print("at",i,5,"price",str_get)
  except:
    print("at",i,5,"no price")
  # 版本数
  try:
    str_get = li_list[i].find('span',{'class':'a-color-secondary'}).contents
    worksheet.write(i+1, 7, str_get)
    print("at",i,5,"version",str_get)
  except:
    print("at",i,"no versions")
  

# 保存
workbook.save('3.xls')
