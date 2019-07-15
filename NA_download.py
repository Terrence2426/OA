# 2019-07-16

# -*- coding: UTF-8 -*-

import os # 操作系统接口
import win32com.client as win32 # 格式转换
import openpyxl # 读写xlsx文件
from urllib.request import urlopen,urlretrieve # url处理
from bs4 import BeautifulSoup # 网页抓取


# 输入文件位置

path = os.getcwd() # 获取当前工作路径
fileinput = path + "\\" + "公司公告.xls" # Wind原始下载文件


# 将xls文件转为xlsx文件

excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = excel.Workbooks.Open(fileinput)
workbook.SaveAs(fileinput.replace("xls","xlsx"),FileFormat = 51) #FileFormat = 51 is for .xlsx extension

workbook.Close()
excel.Application.Quit()


# 提取xlsx文件信息

fileread = openpyxl.load_workbook(fileinput.replace("xls","xlsx")) # 读取xlsx文件
sheet = fileread[fileread.sheetnames[0]] # 读取工作表
data = [cell.value for cell in list(sheet.columns)[2]][1:-2] # 读取第三列数据（excel函数以文本显示）

dates = [str(cell.value).split(" ")[0] for cell in list(sheet.columns)[0]][1:-2] # 提取公告日期
symbols = [cell.value for cell in list(sheet.columns)[1]][1:-2] # 提取公司代码
hyperlinks = [string.split("\"")[1] for string in data] # 提取公告跳转链接
titles = [string.split("\"")[-2].replace(":","：") for string in data] # 提取公告标题

os.remove(fileinput.replace("xls","xlsx")) # 删除xlsx文件

# 获取下载链接

prefix = 'http://news.windin.com/ns/' # 下载链接前缀

def substract_download_link(link_original): # 从公告跳转链接得到下载链接后缀
    html = urlopen(link_original)
    object = BeautifulSoup(html, 'html.parser')
    link_downloadable = [obj.get('href') for obj in object.find_all('a')][0]
    return link_downloadable


# 下载文件

path_download = path + "\\" + "公告下载"
None if os.path.exists(path_download) else os.makedirs(path_download)
os.chdir(path_download) # 更改工作路径

print("\n开始下载\n")

for i,link in enumerate(hyperlinks):
    urlretrieve(prefix + substract_download_link(link), dates[i] + "-"+ symbols[i] + "-" + titles[i] + ".pdf")
    print("进度" + str(i + 1) + "/" + str(len(hyperlinks)) + " " + dates[i] + "-"+ symbols[i] + "-" + titles[i] + " " + "下载完成" +"\n")

print("全部下载完成\n")

# 下载文件测试
# if __name__ == '__main__':
#     link_original = "http://news.windin.com/ns/bulletin.php?code=6C85FB0DA6F5&id=106794452&type=1"
#     urlretrieve(prefix + substract_download_link(link_original),"test.pdf")
