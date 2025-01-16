import urllib.request
from io import BytesIO
import gzip
from bs4 import BeautifulSoup
import re  # 正则表达式，进行文字匹配
import xlwt
import os
import tkinter as tk
from tkinter import simpledialog
from datetime import datetime

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
    thread_id = simpledialog.askstring("输入", "请输入需要爬取的串号:").strip()

    if not thread_id:
        print("未输入串号，程序终止。")
        return

    print('开始爬取......')
    baseurl = f'https://www.nmbxd1.com/t/{thread_id}?page='
    datalist, author_id = getData(baseurl)
    save_date = datetime.now().strftime('%Y%m%d')
    savepath = f"{save_date}.xls"
    saveData(datalist, savepath, thread_id, author_id, save_date)


def askURL(url):  # 定义获取网页数据的函数askURL
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.1",
        "Cookie": "PHPSESSID=oh8pa19vddvtnn8p5c5mu9slih; memberUserspapapa=i%C3%7Cv%1B%F3t%BC%A63%F2%8A%E9%B1%5Ez%7C%EF~%18%B5%9B%00%D4w%C5%14%9A%B5%C9Yt%19%CE%FA%3D%85%8F%E8%D5%DC%DDmY%1F%EA%AA%1F%80%0B%E7%D3%3E%88g%F2%91%90%C5%98%DF%3A%A0%00; userhash=%F2%85%7B%93%A7%27%BE%85%228%24z%88%00%3D%87%18%91%CE%05%DF%F4%8Bn"
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
        content = re.sub('&gt;', "＞", content)  # 替换&gt;字符
        content = re.sub('<font color="#789922">&gt;&gt;', ">>", content)  # 替换>>字符
        content = re.sub('&lt;', "＜", content)  # 替换&lt;字符
        content = re.sub("<br/>", "", content)
        content = re.sub("<br>", "", content)
        content = re.sub("</br>", "", content)
        content = re.sub('<font color="#789922">', "", content)
        content = re.sub("</font>", "", content)
        content = re.sub("<b>", "", content)
        content = re.sub("</b>", "", content)
        content = re.sub("</small>", "", content)
        content = re.sub("<small>", "", content)
        data.append(content.strip())
    else:
        data.append("")

    return data


def getData(baseurl):
    datalist = []  # 用来存储爬取的网页信息
    page = 1
    author_id = None
    while True:
        url = baseurl + str(page)
        html = askURL(url)  # 保存获取到的网页源码
        print("第%d页" % page)
        main_item, items = parsePage(html)

        if page == 1 and main_item:  # 在第一页时提取主串信息
            data = parseItem(main_item)
            datalist.append(data)
            author_id = data[1]  # 主串作者ID

        if page > 1 and len(items) <= 1:  # 如果除了第一页外，没有超过一个帖子，停止爬取
            break
        for item in items:
            data = parseItem(item)
            if data[0] != "9999999":  # 过滤掉广告信息
                datalist.append(data)  # 将处理好的一个回复数据放入datalist
        page += 1
    return datalist, author_id


def contains_chinese(text):
    """判断文本中是否包含中文字符"""
    for ch in text:
        if '\u4e00' <= ch <= '\u9fff':
            return True
    return False


def saveData(datalist, savepath, thread_id, author_id, save_date):
    print("正在保存数据......")
    workbook = xlwt.Workbook(encoding="utf-8", style_compression=0)
    worksheet = workbook.add_sheet('Sheet1', cell_overwrite_ok=True)

    # 设置不同语言的字体样式
    style_english = xlwt.XFStyle()
    font_english = xlwt.Font()
    font_english.name = 'Times New Roman'
    font_english.bold = False
    style_english.font = font_english

    style_chinese = xlwt.XFStyle()
    font_chinese = xlwt.Font()
    font_chinese.name = '宋体'
    font_chinese.bold = False
    style_chinese.font = font_chinese

    # 设置列名
    col = ("串号", "饼干", "时间", "内容")
    for i in range(0, len(col)):
        worksheet.write(0, i, col[i], style_chinese)  # 列名使用中文样式
        

    # 写入数据
    for i in range(0, len(datalist)):
        print(f"正在写入第{i+1}条数据")
        data = datalist[i]
        for j in range(0, len(data)):
            # 根据内容判断使用哪种字体
            if contains_chinese(data[j]):
                worksheet.write(i + 1, j, data[j], style_chinese)
            else:
                worksheet.write(i + 1, j, data[j], style_english)

    workbook.save(savepath)
    print("数据保存成功！")
    
    # 保存作者所发信息到txt文件
    txt_savepath = f"{save_date}.txt"
    with open(txt_savepath, 'w', encoding='utf-8') as f:
        f.write(f"串号：{thread_id}\n")
        f.write(f"作者饼干：{author_id}\n")
        f.write(f"保存时间：{save_date}\n")
        f.write("\n======\n\n")
        
        for data in datalist:
            if data[1] == author_id:  # 判断是否为作者发的内容
                f.write(f"{data[3]}\n\n\n\n")
    print(f"作者信息保存成功！路径: {txt_savepath}")

if __name__ == "__main__":
    main()
    print("爬取完毕！")
