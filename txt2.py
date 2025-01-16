import xlrd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

def convert_xls_to_txt(xls_path):
    # 打开XLS文件
    workbook = xlrd.open_workbook(xls_path)
    sheet = workbook.sheet_by_index(0)

    # 获取当前日期作为文件名的一部分
    save_date = datetime.now().strftime('%Y%m%d')
    txt_savepath = f"{save_date}.txt"

    # 获取串号
    main_thread_id = sheet.cell_value(1, 0)  # 假设串号在第二行第一列
    # 获取po饼干
    author_id = sheet.cell_value(1, 1)  # 假设po饼干在第二行第二列
    # 获取时间
    created_at = sheet.cell_value(1, 2)  # 假设时间在第二行第三列

    with open(txt_savepath, 'w', encoding='utf-8') as f:
        f.write(f"串号：{main_thread_id}\n")
        f.write(f"po：{author_id}\n")
        f.write(f"时间：{created_at}\n")

        # 遍历XLS文件中的所有行并写入到TXT文件
        for row_idx in range(1, sheet.nrows):
            row = sheet.row_values(row_idx)
            if len(row) >= 4:
                content = row[3]  # 假设内容在第四列

                f.write(f"\n\n======\n\n{content}")  # 内容之间空三行

    print(f"数据已成功保存为 {txt_savepath}")

def main():
    root = tk.Tk()
    root.withdraw()
    xls_path = filedialog.askopenfilename(title="选择XLS文件", filetypes=[("XLS files", "*.xls")])

    if not xls_path:
        print("未选择文件，程序终止。")
        return

    convert_xls_to_txt(xls_path)

if __name__ == "__main__":
    main()