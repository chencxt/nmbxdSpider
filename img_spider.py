import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.drawing.image import Image
import re
import os
import time
from tqdm import tqdm
import datetime


# 修改点1：mian调用的cookies部分
# 修改点2：串号

def fetch_page(url, cookies=None):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    response = requests.get(url, headers=headers, cookies=cookies)
    response.raise_for_status()
    return response.text


def parse_page(html):
    soup = BeautifulSoup(html, 'html.parser')
    h2_tag = soup.find('h2', class_='h-title')
    if h2_tag is None:
        raise ValueError("No <h2 class='h-title'> found in the page")
    title = h2_tag.get_text(strip=True)

    # 使用正则表达式匹配以 "https://image.nmb.best/image/" 开头，且以 ".jpg" ".png" ".gif" 结尾的图片链接
    image_urls = re.findall(r'https://image\.nmb\.best/image/.*?\.(?:jpg|png|gif)', str(soup))

    # 检查是否存在下一页
    is_last_page = bool(soup.find('li', class_='uk-disabled', string=re.compile(r'下一页')))

    return title, image_urls, is_last_page


def download_image(url, folder):
    if not os.path.exists(folder):
        os.makedirs(folder)
    image_name = url.split('/')[-1]
    image_path = os.path.join(folder, image_name)
    response = requests.get(url)
    with open(image_path, 'wb') as file:
        file.write(response.content)
    return image_path


def save_to_excel(title, image_paths):
    save_date = datetime.datetime.now().strftime('%Y%m%d')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Images"
    ws.append(["Image Path"])
    for i, path in enumerate(image_paths, start=1):
        img = Image(path)
        img.anchor = f'A{i + 1}'
        ws.add_image(img)
        ws.append([path])
    filename = f"{title} - {save_date} - ImagesCollection.xlsx"
    wb.save(filename)
    print(f"Saved {filename} \n 你可以将扒好的.xlsx文件的后缀名改成.zip，解压后在xl/media/路径内，就是原始图片文件")


def imgmain(base_url, start_page=1, cookies=None):
    all_image_paths = []
    page = start_page

    while True:
        url = f"{base_url}?page={page}"
        print(f"Fetching {url} , every link will add 1000ms ping.")
        html = fetch_page(url, cookies)
        try:
            title, image_urls, is_last_page = parse_page(html)
            if image_urls:
                image_urls = image_urls[1:]
            for img_url in tqdm(image_urls, desc=f"Downloading images from page,please wait: {page}"):
                img_path = download_image(img_url, title)
                all_image_paths.append(img_path)
            if is_last_page:
                break
            page += 1
            time.sleep(1)  # 添加1秒延时
        except ValueError as e:
            print(f"Error parsing page {page}: {e}")
            break
        except Exception as e:
            print(f"An unexpected error occurred on page {page}: {e}")
            break

    # 保存所有爬取到的图片路径到一个Excel文件中
    save_to_excel(title, all_image_paths)


'''
if __name__ == "__main__":
    base_url = "https://www.nmbxd1.com/t/51414457"  # Replace with the actual base URL
    cookies = {
        "Cookie": "PHPSESSID=qdaqmuds779s9o8nmhutv0ghl1; memberUserspapapa=%E0an%0C%B5%03%BE%E67%B6%88%02iq%E0%14c%D7_%E57%96Y%D4%FD%04%5E%EB%E5%A6%E6%D5%91%9D%1F%5E%99w%BF%F9K%B4L%88%2F%5BV%AF%C64%DB%5D%89%24%11%B8%93%01%9A%22%C5%D8%93%19; userhash=%E7%DAe%DE%DB%DAV%AA%3D%7DKL%EF%28%E0%16G%CA%1F%B9%40%99%5B%E3"

    }  # Replace with the actual cookies required to access the site
    imgmain(base_url, start_page=1, cookies=cookies)  # Start from the first page
'''
