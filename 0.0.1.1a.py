import urllib.request
from io import BytesIO
import gzip
from bs4 import BeautifulSoup
import re  # 正则表达式，进行文字匹配
import xlwt
import os
import tkinter as tk
from tkinter import simpledialog

# 创建正则表达式对象，表示规范（字符串的模式）
# 寻找串号
findthreadid = re.compile('<li><a href="/Admin/Content/sagePost/id/(.*?).html">SAGE</a></li>')
# 寻找饼干
findthreaduid = re.compile('<span class="h-threads-info-uid">ID:(.*?)</span>')
# 寻找跟串时间
findcreatedat = re.compile('<span class="h-threads-info-createdat">(.*?)</span>')
# 寻找内容 #re.S让换行包含在正则表达式中
findcontent = re.compile('<div class="h-threads-content">\n(.*?)</div>', re.S)


def main():
    # 创建一个弹出窗口让用户输入串号
    root = tk.Tk()
    root.withdraw()
    thread_id = simpledialog.askstring("输入", "请输入需要爬取的串号:")
    
    if not thread_id:
        print("未输入串号，程序终止。")
        return

    print('开始爬取......')
    baseurl = f'https://www.nmbxd1.com/t/{thread_id}?page='
    datalist = getData(baseurl)
    savepath = "data.xls"
    saveData(datalist, savepath)


def askURL(url):  # 定义获取网页数据的函数askURL
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.1",
        "Cookie": "PHPSESSID=6hsgf7hfecfr9gp6f0o6e6jsfs; memberUserspapapa=%E0an%0C%B5%03%BE%E67%B6%88%02iq%E0%14c%D7_%E57%96Y%D4%FD%04%5E%EB%E5%A6%E6%D5%91%9D%1F%5E%99w%BF%F9K%B4L%88%2F%5BV%AF%C64%DB%5D%89%24%11%B8%93%01%9A%22%C5%D8%93%19; userhash=%E7%DAe%DE%DB%DAV%AA%3D%7DKL%EF%28%E0%16G%CA%1F%B9%40%99%5B%E3"
    }  # 用户代理

    req = urllib.request.Request(url, headers=head)
    # 发送请求，由于urlopen无法传入参数，声明一个request对象
    html = ""
    with urllib.request.urlopen(req) as response:
        h = response.read()
        buff = BytesIO(h)
        f = gzip.GzipFile(fileobj=buff)
        html = f.read().decode('utf-8')
    return html


def parsePage(html):
    soup = BeautifulSoup(html, "html.parser")
    main_item = soup.find('div', class_="h-threads-item-main")  # 获取主串
    items = soup.find_all('div', class_="h-threads-item-reply-main")  # 获取回复串
    return main_item, items


def parseItem(item):
    data = []
    item = str(item)
    threadid = re.findall(findthreadid, item)
    threadid = threadid[0] if threadid else ""
    data.append(threadid)

    threaduid = re.findall(findthreaduid, item)
    data.append(threaduid[0] if threaduid else "")

    createdat = re.findall(findcreatedat, item)
    data.append(createdat[0] if createdat else "")

    content = re.findall(findcontent, item)
    if content:
        content = content[0]
        content = re.sub("<.*?>", "", content)  # 去掉所有HTML标签
        content = re.sub('&gt;',"＞",content) #替换&gt;字符
        content = re.sub('<font color="#789922">&gt;&gt;',">>",content) #替换>>字符
        content = re.sub('&lt;',"＜",content) #替换&lt;字符            
        content = re.sub("<br/>","",content)
        content = re.sub("<br>","",content)
        content = re.sub("</br>","",content)
        content = re.sub('<font color="#789922">',"",content)
        content = re.sub("</font>","",content)
        content = re.sub("<b>","",content)
        content = re.sub("</b>","",content)
        content = re.sub("</small>","",content)
        content = re.sub("<small>","",content)
        data.append(content.strip())
    else:
        data.append("")

    return data


def getData(baseurl):
    datalist = []  # 用来存储爬取的网页信息
    page = 1
    while True:
        url = baseurl + str(page)
        html = askURL(url)  # 保存获取到的网页源码
        print("第%d页" % page)
        main_item, items = parsePage(html)
        
        if page == 1 and main_item:  # 在第一页时提取主串信息
            data = parseItem(main_item)
            datalist.append(data)

        if page > 1 and len(items) <= 1:  # 如果除了第一页外，没有超过一个帖子，停止爬取
            break
        for item in items:
            data = parseItem(item)
            if data[0] != "9999999":  # 过滤掉广告信息
                datalist.append(data)  # 将处理好的一个回复数据放入datalist
        page += 1
    return datalist


def saveData(datalist, savepath):
    print("save....")
    workbook = xlwt.Workbook(encoding="utf-8", style_compression=0)  # 创建workbook对象
    worksheet = workbook.add_sheet('Sheet1', cell_overwrite_ok=True)  # 创建工作表
    col = ("串号", "饼干", "时间", "内容")  # 定义一个元组
    for a in range(0, 4):
        worksheet.write(0, a, col[a])  # 输入列名

    for i in range(len(datalist)):
        print("第%d个回复" % (i + 1))  # 显示写入进度
        data = datalist[i]
        for j in range(0, 4):
            worksheet.write(i + 1, j, data[j])  # 写入数据

    workbook.save(savepath)


if __name__ == "__main__":  # 当程序执行时
    main()
    print("爬取完毕！")
