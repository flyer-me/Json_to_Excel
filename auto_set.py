# 导入模块
import openpyxl
import re
import pandas as pd
import json

# 打开xlsx文件
wb = openpyxl.load_workbook('测试.xlsx')
ws = wb.active

# 获取最大行数和列数
max_row = ws.max_row
max_col = ws.max_column

# 定义正则表达式匹配json字符串
pattern = re.compile(r'\{[\s\S]*?\}')
#new_dfs = []

# 遍历最后一列的单元格，提取json字符串，并生成新表格
for i in range(2, max_row + 1):
    # 跳过标题行
    cell = ws.cell(row=i, column=max_col)
    if cell.row == 1:
        continue
    
    # 获取单元格内容
    value = cell.value
    print(value)
    # 匹配json字符串
    match = pattern.search(value)
    print(match)
    # 如果匹配成功，提取json字符串，并转换为字典
    if match:
        json_str = match.group()
        json_dict = eval(json_str)
        
        # 根据字典生成新表格
        new_table = pd.DataFrame({key: [value] for key, value in json_dict.items()})

        cell_dict = json.loads(json_str)
        keys = list(cell_dict.keys())
        values = list(cell_dict.values())
        # 在第一行末尾添加键列表中的元素作为新的列名，并在相应的单元格下方添加值列表中的元素作为新数据 
        for j in range(len(keys)):
            # 计算新列名所在位置 
            new_col_index = max_col + j
            # 在第一行末尾添加新列名 
            ws.cell(row=1, column=new_col_index).value = keys[j]
            # 在相应位置添加新数据 
            ws.cell(row=i, column=new_col_index).value = values[j]

# 保存修改后的xlsx文件
wb.save('完成.xlsx')