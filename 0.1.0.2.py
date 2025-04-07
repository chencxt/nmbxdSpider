import urllib.request
from io import BytesIO
import gzip
from bs4 import BeautifulSoup
import re
import xlwt
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from datetime import datetime

# 正则表达式对象
findthreadid = re.compile(r'<li><a href="/Admin/Content/sagePost/id/(.*?).html">SAGE</a></li>')
findthreaduid = re.compile(r'<span class="h-threads-info-uid">ID:(.*?)</span>')
findcreatedat = re.compile(r'<span class="h-threads-info-createdat">(.*?)</span>')
findcontent = re.compile(r'<div class="h-threads-content">\n(.*?)</div>', re.S)

def main():
    root = tk.Tk()
    root.withdraw()
    
    # 获取用户Cookie
    cookie = get_user_cookie()
    if not cookie:
        print("未提供Cookie，程序终止。")
        return
    
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
            process_thread(thread_id, cookie)
        except Exception as e:
            print(f"处理串号 {thread_id} 时发生错误：{str(e)}")

def get_user_cookie():
    """获取用户输入的Cookie，优先使用缓存"""
    CACHE_DIR = os.path.join(os.getenv('APPDATA'), 'CookieCache')
    COOKIE_PATH = os.path.join(CACHE_DIR, 'cookie.txt')
    os.makedirs(CACHE_DIR, exist_ok=True)  # 确保目录存在

    # 尝试读取旧Cookie
    old_cookie = None
    if os.path.exists(COOKIE_PATH):
        with open(COOKIE_PATH, 'r', encoding='utf-8') as f:
            old_cookie = f.read().strip()

    use_new = True
    # 如果有旧Cookie，询问是否使用新Cookie
    if old_cookie:
        use_new = messagebox.askyesno("Cookie设置", "是否使用新Cookie？")

    if use_new:
        cookie = simpledialog.askstring("输入Cookie", "请输入Cookie值：")
        # 用户取消输入或关闭对话框
        if cookie is None:
            messagebox.showinfo("提示", "未输入Cookie，程序终止。")
            return None
        # 处理空输入
        while not cookie.strip():
            messagebox.showwarning("输入错误", "Cookie不能为空，请重新输入。")
            cookie = simpledialog.askstring("输入Cookie", "请输入Cookie值：")
            if cookie is None:
                messagebox.showinfo("提示", "未输入Cookie，程序终止。")
                return None
        # 保存新Cookie到缓存
        with open(COOKIE_PATH, 'w', encoding='utf-8') as f:
            f.write(cookie.strip())
    else:
        cookie = old_cookie

    return cookie

def process_thread(thread_id, cookie):
    baseurl = f'https://www.nmbxd1.com/t/{thread_id}?page='
    datalist, author_id = getData(baseurl, cookie)
    save_date = datetime.now().strftime('%Y%m%d')
    
    # 创建图片保存目录
    image_dir = f"{thread_id}_{save_date}_images"
    os.makedirs(image_dir, exist_ok=True)
    
    # 下载图片并更新数据列表
    download_images(datalist, image_dir, cookie)
    
    # 保存Excel文件
    savepath = f"{thread_id}_{save_date}.xls"
    saveData(datalist, savepath, thread_id, author_id, save_date)

def askURL(url, cookie):
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.1",
        "Cookie": cookie
    }
    req = urllib.request.Request(url, headers=head)
    html = ""
    try:
        with urllib.request.urlopen(req) as response:
            h = response.read()
            buff = BytesIO(h)
            f = gzip.GzipFile(fileobj=buff)
            html = f.read().decode('utf-8')
    except Exception as e:
        print(f"请求URL时发生错误：{str(e)}")
        html = ""
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

def getData(baseurl, cookie):
    datalist = []
    page = 1
    author_id = None
    while True:
        url = baseurl + str(page)
        html = askURL(url, cookie)
        if not html:
            break
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

def download_images(datalist, image_dir, cookie):
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:126.0) Gecko/20100101 Firefox/126.1",
        "Cookie": cookie
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