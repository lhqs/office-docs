---
title: Excel与Python集成
slug: /excel/29-Excel与Python集成
sidebar_position: 29
---
# Excel与Python集成

## 本章概览

Python是强大的数据分析语言,与Excel结合可以实现更复杂的数据处理和分析。本章介绍如何使用Python操作Excel,以及如何在Excel中嵌入Python脚本。

**学习目标**
- 掌握openpyxl库操作Excel
- 学会pandas处理Excel数据
- 熟练使用xlwings双向交互
- 掌握数据分析和可视化
- 学会自动化报表生成

**前提条件**
- Python基础知识
- 已安装Python 3.7+

---

## 29.1 Python Excel库概览

### 29.1.1 常用库对比

**openpyxl**
```
功能: 读写xlsx文件
优点: 功能全面,支持样式/公式/图表
缺点: 速度较慢
适用: 需要精细控制格式

安装: pip install openpyxl
```

**pandas**
```
功能: 数据分析和处理
优点: 数据分析功能强大,速度快
缺点: 格式控制有限
适用: 数据分析,大批量处理

安装: pip install pandas openpyxl
```

**xlwings**
```
功能: Python与Excel双向交互
优点: 调用Excel功能,实时交互
缺点: 需要安装Excel
适用: 自动化,实时数据交换

安装: pip install xlwings
```

**xlrd/xlsxwriter**
```
xlrd: 只读xls/xlsx
xlsxwriter: 只写xlsx,性能好

适用: 单向操作,性能要求高
```

### 29.1.2 选择建议

**场景对应**
```
数据分析 → pandas
格式控制 → openpyxl
自动化报表 → xlsxwriter
Excel交互 → xlwings
批量处理 → pandas + openpyxl
```

---

## 29.2 openpyxl操作Excel

### 29.2.1 基本操作

**读取工作簿**
```python
from openpyxl import load_workbook

# 加载工作簿
wb = load_workbook('data.xlsx')

# 获取工作表
ws = wb['Sheet1']  # 按名称
ws = wb.active  # 活动工作表

# 所有工作表名
print(wb.sheetnames)
```

**读取单元格**
```python
# 方法1: 坐标
cell = ws['A1']
print(cell.value)

# 方法2: 行列号
cell = ws.cell(row=1, column=1)
print(cell.value)

# 读取范围
for row in ws['A1:C3']:
    for cell in row:
        print(cell.value)
```

**写入数据**
```python
# 写入单元格
ws['A1'] = '姓名'
ws['B1'] = '年龄'
ws.cell(row=2, column=1, value='张三')
ws.cell(row=2, column=2, value=28)

# 保存
wb.save('output.xlsx')
```

**创建新工作簿**
```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "销售数据"

ws['A1'] = '产品'
ws['B1'] = '销售额'

wb.save('new_file.xlsx')
```

### 29.2.2 批量操作

**遍历行**
```python
# iter_rows
for row in ws.iter_rows(min_row=2, max_row=10, values_only=True):
    name, age, score = row
    print(f"{name}: {age}岁, {score}分")
```

**遍历列**
```python
# iter_cols
for col in ws.iter_cols(min_col=1, max_col=3, values_only=True):
    print(col)
```

**批量写入**
```python
data = [
    ['张三', 28, 85],
    ['李四', 32, 92],
    ['王五', 25, 78]
]

for row in data:
    ws.append(row)

wb.save('data.xlsx')
```

### 29.2.3 格式设置

**字体样式**
```python
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 字体
ws['A1'].font = Font(
    name='黑体',
    size=14,
    bold=True,
    italic=False,
    color='FF0000'  # 红色(RGB)
)

# 填充
ws['A1'].fill = PatternFill(
    start_color='FFFF00',  # 黄色
    end_color='FFFF00',
    fill_type='solid'
)

# 对齐
ws['A1'].alignment = Alignment(
    horizontal='center',
    vertical='center'
)

# 边框
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
ws['A1'].border = thin_border
```

**数字格式**
```python
ws['B2'].number_format = '#,##0.00'  # 千分位
ws['C2'].number_format = '0.00%'  # 百分比
ws['D2'].number_format = 'yyyy-mm-dd'  # 日期
```

### 29.2.4 公式和图表

**插入公式**
```python
ws['D2'] = '=B2*C2'  # 公式
ws['D3'] = '=SUM(D2:D10)'

# 公式不会自动计算,需要Excel打开时计算
```

**创建图表**
```python
from openpyxl.chart import BarChart, Reference

# 创建图表
chart = BarChart()
chart.title = "销售额"
chart.x_axis.title = "产品"
chart.y_axis.title = "金额"

# 数据引用
data = Reference(ws, min_col=2, min_row=1, max_row=10)
cats = Reference(ws, min_col=1, min_row=2, max_row=10)

chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# 添加到工作表
ws.add_chart(chart, "E5")
wb.save('chart.xlsx')
```

---

## 29.3 pandas处理Excel

### 29.3.1 读取数据

**读取Excel**
```python
import pandas as pd

# 读取默认第一个工作表
df = pd.read_excel('data.xlsx')

# 指定工作表
df = pd.read_excel('data.xlsx', sheet_name='Sheet1')

# 读取多个工作表
dfs = pd.read_excel('data.xlsx', sheet_name=['Sheet1', 'Sheet2'])

# 读取所有工作表
dfs = pd.read_excel('data.xlsx', sheet_name=None)

# 指定列
df = pd.read_excel('data.xlsx', usecols='A:C')
df = pd.read_excel('data.xlsx', usecols=['姓名', '年龄'])

# 跳过行
df = pd.read_excel('data.xlsx', skiprows=2)  # 跳过前2行
```

**查看数据**
```python
print(df.head())  # 前5行
print(df.tail())  # 后5行
print(df.info())  # 数据信息
print(df.describe())  # 统计摘要
```

### 29.3.2 数据处理

**筛选**
```python
# 条件筛选
df_filtered = df[df['年龄'] > 25]
df_filtered = df[df['销售额'] > 10000]

# 多条件
df_filtered = df[(df['年龄'] > 25) & (df['部门'] == '销售')]
```

**排序**
```python
df_sorted = df.sort_values('年龄')  # 升序
df_sorted = df.sort_values('年龄', ascending=False)  # 降序

# 多列排序
df_sorted = df.sort_values(['部门', '年龄'], ascending=[True, False])
```

**分组聚合**
```python
# 按部门统计
grouped = df.groupby('部门')['销售额'].sum()

# 多种聚合
result = df.groupby('部门').agg({
    '销售额': 'sum',
    '年龄': 'mean',
    '姓名': 'count'
})
```

**数据透视表**
```python
pivot = pd.pivot_table(
    df,
    values='销售额',
    index='部门',
    columns='产品',
    aggfunc='sum',
    fill_value=0
)
```

### 29.3.3 写入Excel

**基本写入**
```python
df.to_excel('output.xlsx', index=False)
```

**写入多个工作表**
```python
with pd.ExcelWriter('output.xlsx') as writer:
    df1.to_excel(writer, sheet_name='销售数据', index=False)
    df2.to_excel(writer, sheet_name='统计汇总', index=False)
```

**追加数据**
```python
with pd.ExcelWriter('output.xlsx', mode='a', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='新数据', index=False)
```

**格式化输出**
```python
with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='数据', index=False)

    # 获取工作表
    workbook = writer.book
    worksheet = writer.sheets['数据']

    # 定义格式
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#4472C4',
        'font_color': 'white'
    })

    # 应用格式
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
```

---

## 29.4 xlwings双向交互

### 29.4.1 基本操作

**打开Excel**
```python
import xlwings as xw

# 新建工作簿
wb = xw.Book()

# 打开现有工作簿
wb = xw.Book('data.xlsx')

# 获取工作表
ws = wb.sheets['Sheet1']
ws = wb.sheets[0]  # 第一个工作表
```

**读写数据**
```python
# 读取单元格
value = ws.range('A1').value

# 写入单元格
ws.range('A1').value = '姓名'

# 读取区域
data = ws.range('A1:C10').value  # 返回列表

# 写入区域
ws.range('A1').value = [[1,2,3], [4,5,6]]

# 读取整个已用区域
data = ws.used_range.value
```

**Excel对象操作**
```python
# 执行Excel函数
result = ws.range('A1').formula = '=SUM(B1:B10)'

# 运行VBA宏
wb.macro('宏名称')()

# 保存
wb.save()
wb.save('new_name.xlsx')

# 关闭
wb.close()
```

### 29.4.2 与pandas集成

**DataFrame互转**
```python
import pandas as pd

# Excel → DataFrame
df = ws.range('A1').options(pd.DataFrame, index=False).value

# DataFrame → Excel
ws.range('A1').value = df
ws.range('A1').options(index=False).value = df  # 不包含索引
```

**实时更新**
```python
# 从Excel读取
df = ws.range('A1').options(pd.DataFrame).value

# 数据处理
df['总额'] = df['单价'] * df['数量']

# 写回Excel
ws.range('A1').value = df
```

### 29.4.3 自动化应用

**示例:数据刷新**
```python
def refresh_data():
    wb = xw.Book.caller()  # 获取调用工作簿
    ws = wb.sheets['数据']

    # 从数据库或API获取数据
    df = get_latest_data()  # 自定义函数

    # 更新Excel
    ws.range('A1').value = df

    # 刷新数据透视表
    wb.api.RefreshAll()
```

**嵌入Excel**
```python
# 在Excel中添加按钮,分配给Python函数
# 工具: xlwings → Python UDF
```

---

## 29.5 实战案例

### 案例1:批量处理Excel文件

**需求**
合并文件夹中所有Excel的第一个工作表

**代码**
```python
import pandas as pd
import os

def merge_excel_files(folder_path, output_file):
    all_data = []

    # 遍历文件
    for file in os.listdir(folder_path):
        if file.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file)
            df = pd.read_excel(file_path)
            df['来源文件'] = file  # 添加来源列
            all_data.append(df)

    # 合并
    result = pd.concat(all_data, ignore_index=True)

    # 保存
    result.to_excel(output_file, index=False)
    print(f"合并完成,共{len(result)}行")

# 使用
merge_excel_files('C:/数据/', 'merged.xlsx')
```

### 案例2:自动化报表生成

**需求**
从数据库读取数据,生成格式化Excel报表

**代码**
```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.chart import BarChart, Reference

def generate_report(data, output_file):
    # 数据处理
    df = pd.DataFrame(data)

    # 写入Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='数据', index=False)

    # 格式化
    wb = load_workbook(output_file)
    ws = wb['数据']

    # 标题行格式
    for cell in ws[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='4472C4', fill_type='solid')

    # 添加图表
    chart = BarChart()
    data_ref = Reference(ws, min_col=2, min_row=1, max_row=len(df)+1)
    cats = Reference(ws, min_col=1, min_row=2, max_row=len(df)+1)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "E2")

    wb.save(output_file)
    print("报表生成完成!")

# 使用
data = {
    '产品': ['A', 'B', 'C'],
    '销售额': [1000, 1500, 1200]
}
generate_report(data, 'report.xlsx')
```

### 案例3:数据清洗

**需求**
清洗脏数据并规范格式

**代码**
```python
import pandas as pd
import re

def clean_data(input_file, output_file):
    df = pd.read_excel(input_file)

    # 去除空格
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # 规范日期格式
    df['日期'] = pd.to_datetime(df['日期'], errors='coerce')

    # 提取数字(从"1000元"提取1000)
    df['金额'] = df['金额'].apply(lambda x:
        float(re.sub(r'[^0-9.]', '', str(x))) if pd.notna(x) else None
    )

    # 填充缺失值
    df['部门'].fillna('未知', inplace=True)

    # 删除重复行
    df.drop_duplicates(inplace=True)

    # 保存
    df.to_excel(output_file, index=False)
    print(f"清洗完成,剩余{len(df)}行")

clean_data('raw_data.xlsx', 'clean_data.xlsx')
```

### 案例4:动态仪表盘

**使用xlwings**
```python
import xlwings as xw
import pandas as pd

@xw.func
def get_sales_data(region):
    """Excel中调用此函数获取数据"""
    # 从数据库查询(示例)
    df = query_database(f"SELECT * FROM sales WHERE region='{region}'")
    return df.values.tolist()

@xw.sub
def refresh_dashboard():
    """刷新仪表盘"""
    wb = xw.Book.caller()
    ws = wb.sheets['Dashboard']

    # 更新数据
    data = get_latest_metrics()
    ws.range('A1').value = data

    # 刷新图表
    wb.api.RefreshAll()
```

---

## 29.6 性能优化

### 29.6.1 pandas优化

**读取优化**
```python
# 只读取需要的列
df = pd.read_excel('data.xlsx', usecols=['A', 'B', 'C'])

# 指定数据类型
df = pd.read_excel('data.xlsx', dtype={'年龄': int, '姓名': str})

# 分块读取大文件
chunks = pd.read_excel('large_file.xlsx', chunksize=1000)
for chunk in chunks:
    process(chunk)  # 处理每个块
```

**写入优化**
```python
# 使用xlsxwriter引擎(更快)
df.to_excel('output.xlsx', engine='xlsxwriter', index=False)

# 关闭自动列宽
with pd.ExcelWriter('output.xlsx', engine='xlsxwriter',
                    options={'strings_to_urls': False}) as writer:
    df.to_excel(writer, index=False)
```

### 29.6.2 批量处理优化

**并行处理**
```python
from multiprocessing import Pool
import os

def process_file(file_path):
    df = pd.read_excel(file_path)
    # 处理逻辑...
    return df

if __name__ == '__main__':
    files = ['file1.xlsx', 'file2.xlsx', 'file3.xlsx']

    with Pool(4) as pool:  # 4个进程
        results = pool.map(process_file, files)

    # 合并结果
    final_df = pd.concat(results)
```

---

## 29.7 部署和分发

### 29.7.1 打包Python脚本

**PyInstaller**
```bash
pip install pyinstaller
pyinstaller --onefile excel_tool.py

# 生成exe文件,无需Python环境即可运行
```

**配置文件**
```python
# config.ini
[database]
host = localhost
port = 3306
username = root
password = ****

# 读取配置
import configparser
config = configparser.ConfigParser()
config.read('config.ini')
host = config['database']['host']
```

### 29.7.2 xlwings插件

**创建Excel加载项**
```bash
xlwings addin install

# 在Excel中:
# 开发工具 → Excel加载项 → 勾选xlwings
```

**UDF(用户定义函数)**
```python
import xlwings as xw

@xw.func
def double_value(x):
    """将值翻倍"""
    return x * 2

# Excel中使用:
# =double_value(A1)
```

---

## 本章小结

**核心要点**
1. **openpyxl**: 读写xlsx,控制格式
2. **pandas**: 数据分析,快速处理
3. **xlwings**: Python与Excel双向交互
4. **实战**: 批量处理、自动化报表、数据清洗
5. **性能**: 优化读写速度,并行处理

**库选择建议**
| 需求 | 推荐库 |
|------|--------|
| 数据分析 | pandas |
| 格式控制 | openpyxl |
| 批量处理 | pandas + openpyxl |
| 实时交互 | xlwings |
| 性能优先 | xlsxwriter |

**学习建议**
1. 先掌握一个库(推荐pandas)
2. 结合实际需求学习
3. 多查阅官方文档
4. 参考GitHub优秀项目
5. 注意错误处理和日志

**下一步学习**
- 第30章: Excel自动化办公
- Python数据分析进阶
- 机器学习与Excel结合

---

## 思考练习

1. 用pandas读取Excel并进行数据透视分析
2. 用openpyxl批量格式化多个工作表
3. 用xlwings创建实时刷新的Excel仪表盘
4. 合并文件夹中所有Excel文件
5. 从数据库读取数据自动生成格式化报表

**练习提示**
- 先处理小数据集测试
- 添加异常处理
- 记录日志便于调试
- 考虑性能优化

