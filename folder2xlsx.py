import os
import pandas as pd


def extract_and_save_xls_filenames(folder_path, output_file):
    # 检查输出文件的目录是否存在，如果不存在则创建
    output_dir = os.path.dirname(output_file)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 获取指定文件夹中的所有文件名
    filenames = [f for f in os.listdir(folder_path) if f.endswith('.txt')]
    # txt|xls

    # 用于存储分割后的文件名数据
    data = []

    for filename in filenames:
        # 去掉文件扩展名
        name = os.path.splitext(filename)[0]
        # 使用“ - ”分隔文件名
        parts = name.split(' - ')
        # 确保第四列以数字形式存储
        if len(parts) > 3:
            try:
                parts[3] = int(parts[3])
            except ValueError:
                parts[3] = parts[3]  # 如果无法转换为数字，保持原样
        data.append(parts)

    # 将数据转换为DataFrame并添加标题
    df = pd.DataFrame(data, columns=["串号", "标题", "板块", "扒取时间"])

    # 保存为xlsx文件
    try:
        df.to_excel(output_file, index=False)
        print(f"文件已成功保存为 {output_file}")
    except PermissionError:
        print(f"无法写入文件 {output_file}，请检查文件是否已被打开或权限是否正确。")


# 使用示例
folder_path = 'E:/小说工作区/X岛——花园二号计划/part1.小说完结串51854427的备份（☆）'  # 修改为实际文件夹路径
output_file = 'E:/小说工作区/X岛——花园二号计划/part1.小说完结串51854427的备份（☆）/namelist.xlsx'  # 修改为实际输出文件路径
extract_and_save_xls_filenames(folder_path, output_file)
