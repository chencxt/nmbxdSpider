import urllib.request
from io import BytesIO
import gzip
from bs4 import BeautifulSoup
import re
import xlwt
import os
import tkinter as tk
from tkinter import filedialog, simpledialog
from datetime import datetime

# 正则表达式对象
findthreadid = re.compile(r'<li><a href="/Admin/Content/sagePost/id/(.*?).html">SAGE</a></li>')
findthreaduid = re.compile(r'<span class="h-threads-info-uid">ID:(.*?)</span>')
findcreatedat = re.compile(r'<span class="h-threads-info-createdat">(.*?)</span>')
findcontent = re.compile(r'<div class="h-threads-content">\n(.*?)</div>', re.S)

def main():
    root = tk.Tk()
    root.withdraw()
    
    # 选择包含串号的TXT文件
    file_path = filedialog.askopenfilename(title="选择包含串号的TXT文件", filetypes=[("Text files", "*.txt")])
    if not file_path:
        print("未选择文件，程序终止。")
        return
    
    # 读取串号列表
    with open(file_path, 'r', encoding='utf-8') as f:
        thread_ids = [line.strip() for line in f.readlines() if line.strip()]
    
    # 遍历处理每个串号
    for thread_id in thread_ids:
        print(f"\n开始处理串号：{thread_id}")
        try:
            process_thread(thread_id)
        except Exception as e:
            print(f"处理串号 {thread_id} 时发生错误：{str(e)}")

def process_thread(thread_id):
    baseurl = f'https://www.nmbxd1.com/t/{thread_id}?page='
    datalist, author_id = getData(baseurl)
    save_date = datetime.now().strftime('%Y%m%d')
    
    # 创建图片保存目录
    image_dir = f"{thread_id}_{save_date}_images"
    os.makedirs(image_dir, exist_ok=True)
    
    # 下载图片并更新数据列表
    download_images(datalist, image_dir)
    
    # 保存Excel文件
    savepath = f"{thread_id}_{save_date}.xls"
    saveData(datalist, savepath, thread_id, author_id, save_date)

def askURL(url):
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.1",
        "Cookie": "PHPSESSID=n6estsf6nlru0lo0msn3m0b14j; memberUserspapapa=%12%AB%5ESnq%86%E1%88%92%A0%ED%A3%F7%87L%FA%C0%8F%EC%00%F3%BFf%1D%94%DB%0B%0EV%F5%C8%2A%D7%95%AAG1%B7%01%90%7B%A1%847%91%19L%FB%CB%08W%DB%D7u%161%8B%BE_0%16Y%E7; userhash=%FDEtl%17%8E%CC%06C%8Dn%DBs%C2%96%A6N%22%E1%9FY%7B%B3%91"
    }
    req = urllib.request.Request(url, headers=head)
    html = ""
    with urllib.request.urlopen(req) as response:
        h = response.read()
        buff = BytesIO(h)
        f = gzip.GzipFile(fileobj=buff)
        html = f.read().decode('utf-8')
    return html

def parsePage(html):
    soup = BeautifulSoup(html, "html.parser")
    main_item = soup.find('div', class_="h-threads-item-main")
    items = soup.find_all('div', class_="h-threads-item-reply-main")
    return main_item, items

def parseItem(item):
    data = []
    item = str(item)
    threadid = re.findall(findthreadid, item)
    data.append(threadid[0] if threadid else "")

    threaduid = re.findall(findthreaduid, item)
    data.append(threaduid[0] if threaduid else "")

    createdat = re.findall(findcreatedat, item)
    data.append(createdat[0] if createdat else "")

    content = re.findall(findcontent, item)
    if content:
        content = content[0]
        content = re.sub("<.*?>", "", content)
        content = re.sub('&gt;', "＞", content)
        content = re.sub('<font color="#789922">&gt;&gt;', ">>", content)
        content = re.sub('&lt;', "＜", content)
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

    # 提取图片URL
    soup = BeautifulSoup(item, 'html.parser')
    img_tags = soup.find_all('img', class_='h-threads-img')
    img_urls = []
    for img in img_tags:
        src = img.get('src')
        if src:
            if src.startswith('/'):
                src = 'https://www.nmbxd1.com' + src
            img_urls.append(src)
    data.append('; '.join(img_urls))  # 第五列为图片URL

    return data

def getData(baseurl):
    datalist = []
    page = 1
    author_id = None
    while True:
        url = baseurl + str(page)
        html = askURL(url)
        print(f"正在爬取第{page}页")
        main_item, items = parsePage(html)

        if page == 1 and main_item:
            data = parseItem(main_item)
            datalist.append(data)
            author_id = data[1]

        if page > 1 and len(items) <= 1:
            break
        for item in items:
            data = parseItem(item)
            if data[0] != "9999999":
                datalist.append(data)
        page += 1
    return datalist, author_id

def download_images(datalist, image_dir):
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.1",
        "Cookie": "PHPSESSID=jfard5ls79sbq3t4jqam34o1r8; memberUserspapapa=%2F%ED-%CFY%3E%94%06%A0R2%F5%C4c%95%E1%5E%DE%FC%09%3F%EC%15%98%AFM%E7%EF6%02u%3A%DF%D1%1Avb%A9%90%82x%10%02%ED%C3%E7%9F%9D7_%7B%ECP%CDm%A0%8Dh%A8%0F6%E9%C86; userhash=%E7%DAe%DE%DB%DAV%AA%3D%7DKL%EF%28%E0%16G%CA%1F%B9%40%99%5B%E3"
    }
    for data in datalist:
        if len(data) < 5:
            continue
        img_urls = data[4].split('; ') if data[4] else []
        img_paths = []
        for idx, url in enumerate(img_urls):
            if not url.strip():
                continue
            try:
                post_id = data[0]
                safe_post_id = re.sub(r'[\\/*?:"<>|]', '_', post_id)
                post_dir = os.path.join(image_dir, safe_post_id)
                os.makedirs(post_dir, exist_ok=True)
                
                req = urllib.request.Request(url, headers=head)
                with urllib.request.urlopen(req) as response:
                    content_type = response.headers.get('Content-Type', 'image/jpeg')
                    ext = 'jpg'
                    if content_type:
                        if 'image/jpeg' in content_type:
                            ext = 'jpg'
                        elif 'image/png' in content_type:
                            ext = 'png'
                        elif 'image/gif' in content_type:
                            ext = 'gif'
                        else:
                            ext = content_type.split('/')[-1]
                    filename = f"{idx}.{ext}"
                    filepath = os.path.join(post_dir, filename)
                    with open(filepath, 'wb') as f:
                        f.write(response.read())
                    img_paths.append(filepath)
            except Exception as e:
                print(f"下载失败：{url}，错误：{e}")
                img_paths.append(f"下载失败：{url}")
        data.append('; '.join(img_paths))

def contains_chinese(text):
    for ch in text:
        if '\u4e00' <= ch <= '\u9fff':
            return True
    return False

def saveData(datalist, savepath, thread_id, author_id, save_date):
    workbook = xlwt.Workbook(encoding="utf-8", style_compression=0)
    worksheet = workbook.add_sheet('Sheet1', cell_overwrite_ok=True)

    style_english = xlwt.XFStyle()
    font_english = xlwt.Font()
    font_english.name = 'Times New Roman'
    style_english.font = font_english

    style_chinese = xlwt.XFStyle()
    font_chinese = xlwt.Font()
    font_chinese.name = '宋体'
    style_chinese.font = font_chinese

    col = ("串号", "饼干", "时间", "内容", "图片URL", "图片路径")
    for i in range(len(col)):
        worksheet.write(0, i, col[i], style_chinese)

    for i in range(len(datalist)):
        data = datalist[i]
        for j in range(len(col)):
            if j >= len(data):
                content = ""
            else:
                content = data[j]
            if contains_chinese(str(content)):
                worksheet.write(i + 1, j, content, style_chinese)
            else:
                worksheet.write(i + 1, j, content, style_english)

    workbook.save(savepath)
    print(f"Excel文件已保存：{savepath}")

    # 保存作者信息到txt
    txt_savepath = f"{thread_id}_{save_date}.txt"
    with open(txt_savepath, 'w', encoding='utf-8') as f:
        f.write(f"串号：{thread_id}\n作者饼干：{author_id}\n保存时间：{save_date}\n\n======\n\n")
        for data in datalist:
            if data[1] == author_id:
                f.write(f"{data[3]}\n\n\n\n")

if __name__ == "__main__":
    main()
    print("所有串号处理完毕！")