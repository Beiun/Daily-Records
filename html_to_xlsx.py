from bs4 import BeautifulSoup
import pandas as pd
import os

# 指定包含HTML文件的目录
html_dir = 'test_dir'
xlsx_file = 'output.xlsx'
sheet_name = 'Sheet1'


html_files = [os.path.join(html_dir, file) for file in os.listdir(html_dir) if file.endswith('.html')]

# 首次添加文件覆盖，后续追加 
mode = 'w' 
if_sheet_exists = None
header_flag = True

# 创建一个新的Excel writer对象，使用openpyxl作为引擎
with pd.ExcelWriter(xlsx_file, mode=mode, engine='openpyxl', if_sheet_exists=if_sheet_exists) as writer:
    for html_file in html_files:

        # 加载HTML文件
        with open(html_file, "r", encoding="gb2312") as file:
            html = file.read()

        soup = BeautifulSoup(html, "html.parser")

        # 查找<meta>标签中name属性为"DC.Title"的内容
        dc_title = soup.find("meta", attrs={"name": "DC.Title"})

        # 找到HTML中的第一个表格
        table = soup.find("table")

        # 找到所有行
        rows = table.find_all("tr")

        # 解析表格头部
        headers = [header.text.strip() for header in rows[0].find_all("th")]
        headers.append("Type")

        # 初始化一个空的数据列表
        data = []

        # 解析表格数据
        for row in rows[1:]:
            cols = row.find_all("td")
            cols = [ele.text.strip() for ele in cols]
            cols.append(dc_title['content'])
            data.append(cols)  # 获取每行数据并添加到列表

        # 使用pandas创建DataFrame
        df = pd.DataFrame(data, columns=headers)

        # 根据是否第一次写入决定是否写入列名(表头)
        startrow = 0 if header_flag else writer.sheets[sheet_name].max_row

        #将数据写入excel中的aa表,从第一个空行开始写
        df.to_excel(writer, index=False, startrow=startrow, header=header_flag)
        header_flag = False
        mode = 'a' 
        if_sheet_exists = 'overlay'

print('end')

