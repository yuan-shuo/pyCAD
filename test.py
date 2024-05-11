from pyautocad import Autocad, APoint
from pyautocad.contrib.tables import Table
import pandas as pd
import re

# 图纸提取
acad = Autocad()
acad.prompt("Hello, Autocad from Python\n")
print(acad.doc.Name)
# 表格数据记录
data = []

dp = APoint(10, 0)
for text in acad.iter_objects('Text'):
    # 使用正则表达式提取纯文本内容
    plain_text = re.sub(r'{\\.*?;', '', text.TextString)
    plain_text = plain_text.replace("}", "")
    # 如果纯文本内容不为空，则输出文本
    if plain_text.strip():
        # print('text: %s at: %s' % (plain_text, text.InsertionPoint))
        # text.InsertionPoint = APoint(text.InsertionPoint) + dp
        print(f'"text": {plain_text}')
        data.append([plain_text])

# 表格记录
table = Table()

# 创建 DataFrame
df = pd.DataFrame(data, columns=['TextString'])

# 保存为 Excel 文件
df.to_excel('data.xlsx', index=False)

