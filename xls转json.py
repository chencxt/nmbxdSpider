import pandas as pd
import json
import xlrd
import time
import re


# 需要 把所有修改点全部按格式修改

# 读取xls文件
xls_file_path = 'E:/小说工作区/X岛——花园二号计划/workspace/第三档案馆今天也很和平_20240623.xls'  # 修改点1 替换为你的xls文件路径
df = pd.read_excel(xls_file_path)

# 将DataFrame保存为xlsx文件
xlsx_file_path = 'E:/小说工作区/X岛——花园二号计划/workspace/converted_file.xlsx'
# 修改点2 替换为你想要保存的xlsx文件路径，xlsx会作为中间文件。json也会和它在同一个目录

df.to_excel(xlsx_file_path, index=False)

print("xlsx转换完成，xlsx文件已保存到", xlsx_file_path)




# 读取xlsx文件
file_path = xlsx_file_path  # 替换为你的xlsx文件路径
df = pd.read_excel(file_path, engine='openpyxl')

# 获取标题和副标题
title = df.iloc[0, 0]
subtitle = df.iloc[0, 1]

# 生成文件名
timestamp = time.strftime("%Y%m%d%H%M%S")
file_name = f"{title}_{subtitle}_{timestamp}.json"

# 从第三行开始读取数据
data = df.iloc[2:, :4]  # 假设前四列是：串号、饼干、时间、内容
data.columns = ['串号', '饼干', '时间', '内容']

# 转换为字典列表
data_list = data.to_dict(orient='records')

# 保存为json文件
with open(file_name, 'w', encoding='utf-8') as json_file:
    json.dump(data_list, json_file, ensure_ascii=False, indent=4)

print(f"JSON文件已保存为: {file_name}")

