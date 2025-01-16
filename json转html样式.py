import json
from pathlib import Path

# 读取JSON文件
input_path = Path("E:/小说工作区/X岛——花园二号计划/workspace/第三档案馆今天也很和平_无名氏_20240623154220.json") #修改点1
with input_path.open("r", encoding="utf-8") as file:
    data = json.load(file)

# 构建HTML内容
html_content = '<html><head><meta charset="utf-8"><title>Output</title></head><body><div style="background-color: #FFFFEE;">'

for item in data:
    html_content += '<div style="background-color:#F0E0D6;">'
    html_content += f'<p style="color: #800000; font-size: 16px;">串号: {item["串号"]} <span style="color: #117743; font-size: 16px;"><b> {item["饼干"]}</b></span> 时间: {item["时间"]}</p>'
    content_with_br = item["内容"].replace('\n', '<br>')
    html_content += f'<p style="color: #800000; font-size: 16px;">{content_with_br}</p>'
    html_content += '<p> </p></div>'

html_content += '</div></body></html>'

# 写入HTML文件
output_path = Path("E:/小说工作区/X岛——花园二号计划/workspace/output.html") #修改点2
with output_path.open("w", encoding="utf-8") as file:
    file.write(html_content)

print("HTML文件已生成。")
