---
title: Python办公自动化入门
slug: /automation/60-Python办公自动化入门
sidebar_position: 60
---
# Python办公自动化入门

## 引言

Python是当今最流行的编程语言之一,在办公自动化领域有着强大的能力。相比VBA,Python更灵活、生态更丰富、跨平台支持更好。本章将带你零基础入门Python办公自动化,学会用Python处理Excel、Word、PPT等Office文件。

**为什么选择Python:**
- 语法简洁,易学易用
- 库丰富,功能强大
- 跨平台(Windows/Mac/Linux)
- 社区活跃,资源多

## 一、Python环境搭建

### 1.1 安装Python

**方式1:官网安装(推荐)**
```
1. 访问python.org
2. 下载最新版Python(3.11+)
3. 运行安装程序
4. ✓ 勾选"Add Python to PATH"
5. 点击"Install Now"
6. 等待安装完成
```

**方式2:Anaconda安装**
```
适合数据分析用户
包含Jupyter Notebook
集成常用科学计算库
下载地址:anaconda.com
```

**验证安装:**
```bash
# 打开命令提示符/终端
python --version
# 输出:Python 3.11.0

pip --version
# 输出:pip 22.3 from...
```

### 1.2 安装必要的库

**核心库安装:**
```bash
# Excel处理
pip install openpyxl
pip install pandas
pip install xlwings

# Word处理
pip install python-docx

# PPT处理
pip install python-pptx

# PDF处理
pip install PyPDF2
pip install pdfplumber

# 其他实用库
pip install pillow  # 图像处理
pip install requests  # 网络请求
```

**验证安装:**
```python
import openpyxl
import pandas as pd
from docx import Document
from pptx import Presentation

print("所有库安装成功!")
```

### 1.3 开发工具选择

**1. VS Code(推荐)**
```
优点:轻量、插件丰富、免费
安装:
1. 下载VS Code
2. 安装Python扩展
3. 配置Python解释器
```

**2. PyCharm**
```
优点:功能强大、智能提示
适合:专业开发
```

**3. Jupyter Notebook**
```
优点:交互式、适合数据分析
适合:学习、实验
```

## 二、Python基础速成

### 2.1 基本语法(15分钟掌握)

**变量与数据类型:**
```python
# 字符串
name = "张三"
company = '某某公司'

# 数字
age = 25
salary = 8500.5

# 布尔值
is_manager = True

# 列表
departments = ["销售部", "技术部", "财务部"]

# 字典
employee = {
    "姓名": "张三",
    "年龄": 25,
    "部门": "销售部"
}

# 打印
print(name)
print(f"{name}在{company}工作")  # 格式化字符串
```

**条件判断:**
```python
if salary > 10000:
    print("高薪")
elif salary > 5000:
    print("中等")
else:
    print("较低")
```

**循环:**
```python
# for循环
for dept in departments:
    print(dept)

# while循环
count = 0
while count < 5:
    print(count)
    count += 1

# range循环
for i in range(1, 6):  # 1到5
    print(i)
```

**函数:**
```python
def calculate_tax(income):
    """计算个人所得税"""
    if income <= 5000:
        return 0
    else:
        return (income - 5000) * 0.2

tax = calculate_tax(8000)
print(f"应纳税:{tax}元")
```

### 2.2 文件操作

**读取文件:**
```python
# 读取文本文件
with open('data.txt', 'r', encoding='utf-8') as f:
    content = f.read()
    print(content)

# 逐行读取
with open('data.txt', 'r', encoding='utf-8') as f:
    for line in f:
        print(line.strip())
```

**写入文件:**
```python
# 写入文件(覆盖)
with open('output.txt', 'w', encoding='utf-8') as f:
    f.write("Hello, World!\n")
    f.write("第二行内容")

# 追加内容
with open('output.txt', 'a', encoding='utf-8') as f:
    f.write("\n新增内容")
```

**路径操作:**
```python
import os

# 获取当前目录
current_dir = os.getcwd()
print(current_dir)

# 列出文件
files = os.listdir('D:\\数据')
for file in files:
    print(file)

# 判断文件是否存在
if os.path.exists('data.txt'):
    print("文件存在")

# 创建文件夹
os.makedirs('新文件夹/子文件夹', exist_ok=True)
```

## 三、Excel自动化(openpyxl)

### 3.1 基础操作

**创建和保存工作簿:**
```python
from openpyxl import Workbook

# 创建新工作簿
wb = Workbook()
ws = wb.active
ws.title = "销售数据"

# 写入数据
ws['A1'] = "产品名称"
ws['B1'] = "销售额"
ws['A2'] = "产品A"
ws['B2'] = 10000

# 保存
wb.save('销售数据.xlsx')
```

**打开和读取工作簿:**
```python
from openpyxl import load_workbook

# 打开工作簿
wb = load_workbook('销售数据.xlsx')
ws = wb['销售数据']

# 读取单元格
print(ws['A1'].value)  # 产品名称
print(ws.cell(row=2, column=2).value)  # 10000

# 遍历行
for row in ws.iter_rows(min_row=2, values_only=True):
    产品, 销售额 = row
    print(f"{产品}: {销售额}元")
```

### 3.2 批量数据处理

**示例1:批量写入数据**
```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

# 表头
headers = ["姓名", "部门", "工资"]
ws.append(headers)

# 数据
employees = [
    ["张三", "销售部", 8000],
    ["李四", "技术部", 12000],
    ["王五", "财务部", 9000]
]

for emp in employees:
    ws.append(emp)

wb.save('员工数据.xlsx')
```

**示例2:数据汇总**
```python
from openpyxl import load_workbook

wb = load_workbook('员工数据.xlsx')
ws = wb.active

total_salary = 0
for row in ws.iter_rows(min_row=2, max_col=3, values_only=True):
    name, dept, salary = row
    total_salary += salary

print(f"工资总额:{total_salary}元")

# 写入汇总
ws.append(["合计", "", total_salary])
wb.save('员工数据.xlsx')
```

### 3.3 格式设置

**单元格格式:**
```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = Workbook()
ws = wb.active

# 字体
ws['A1'].font = Font(name='微软雅黑', size=14, bold=True, color='FF0000')

# 背景色
ws['A1'].fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

# 对齐
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

# 边框
border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
ws['A1'].border = border

# 列宽
ws.column_dimensions['A'].width = 20

wb.save('格式化.xlsx')
```

### 3.4 公式与图表

**插入公式:**
```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

ws['A1'] = 100
ws['A2'] = 200
ws['A3'] = '=SUM(A1:A2)'  # 公式

wb.save('公式示例.xlsx')
```

**创建图表:**
```python
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

wb = Workbook()
ws = wb.active

# 数据
data = [
    ["产品", "销售额"],
    ["产品A", 100],
    ["产品B", 150],
    ["产品C", 120]
]
for row in data:
    ws.append(row)

# 创建图表
chart = BarChart()
chart.title = "产品销售额"
chart.x_axis.title = "产品"
chart.y_axis.title = "销售额"

# 设置数据
data_ref = Reference(ws, min_col=2, min_row=1, max_row=4)
cats_ref = Reference(ws, min_col=1, min_row=2, max_row=4)
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(cats_ref)

# 添加图表
ws.add_chart(chart, "D2")

wb.save('图表示例.xlsx')
```

## 四、Excel高级处理(pandas)

### 4.1 pandas基础

**读取Excel:**
```python
import pandas as pd

# 读取Excel
df = pd.read_excel('销售数据.xlsx')

# 查看前几行
print(df.head())

# 查看数据信息
print(df.info())

# 查看统计信息
print(df.describe())
```

**数据筛选:**
```python
# 筛选条件
高销售额 = df[df['销售额'] > 10000]
print(高销售额)

# 多条件筛选
结果 = df[(df['销售额'] > 5000) & (df['部门'] == '销售部')]

# 选择列
df_selected = df[['姓名', '销售额']]
```

### 4.2 数据处理

**数据清洗:**
```python
import pandas as pd

df = pd.read_excel('原始数据.xlsx')

# 删除空行
df = df.dropna()

# 删除重复行
df = df.drop_duplicates()

# 替换值
df['部门'] = df['部门'].replace('销售', '销售部')

# 填充缺失值
df['工资'].fillna(0, inplace=True)
```

**数据分组汇总:**
```python
# 按部门分组求和
dept_summary = df.groupby('部门')['工资'].sum()
print(dept_summary)

# 多重汇总
summary = df.groupby('部门').agg({
    '工资': ['sum', 'mean', 'count'],
    '销售额': 'sum'
})
print(summary)
```

### 4.3 实战案例:多文件合并

```python
import pandas as pd
import os

# 获取所有Excel文件
folder = 'D:\\月度数据\\'
files = [f for f in os.listdir(folder) if f.endswith('.xlsx')]

# 合并所有文件
dfs = []
for file in files:
    df = pd.read_excel(os.path.join(folder, file))
    dfs.append(df)

# 纵向合并
result = pd.concat(dfs, ignore_index=True)

# 保存结果
result.to_excel('汇总数据.xlsx', index=False)
print(f"合并完成,共{len(result)}行数据")
```

## 五、Word自动化(python-docx)

### 5.1 创建Word文档

**基础操作:**
```python
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# 创建文档
doc = Document()

# 添加标题
doc.add_heading('项目报告', level=1)

# 添加段落
p = doc.add_paragraph('这是第一段内容。')

# 设置字体
run = p.runs[0]
run.font.name = '微软雅黑'
run.font.size = Pt(12)
run.font.bold = True
run.font.color.rgb = RGBColor(255, 0, 0)

# 添加列表
doc.add_paragraph('第一项', style='List Bullet')
doc.add_paragraph('第二项', style='List Bullet')

# 保存
doc.save('报告.docx')
```

### 5.2 批量生成文档

**示例:批量生成合同**
```python
from docx import Document
import pandas as pd

# 读取客户信息
df = pd.read_excel('客户信息.xlsx')

# 合同模板内容
template = """
甲方:【公司名称】
乙方:我方公司

根据双方协商,达成以下协议:
1. 合作期限:【合同期限】
2. 合同金额:【合同金额】元
3. 联系人:【联系人】
"""

# 为每个客户生成合同
for index, row in df.iterrows():
    doc = Document()
    doc.add_heading('合作协议', level=1)

    # 替换模板内容
    content = template
    content = content.replace('【公司名称】', row['公司名称'])
    content = content.replace('【合同期限】', row['合同期限'])
    content = content.replace('【合同金额】', str(row['合同金额']))
    content = content.replace('【联系人】', row['联系人'])

    doc.add_paragraph(content)

    # 保存
    filename = f"合同_{row['公司名称']}.docx"
    doc.save(filename)
    print(f"生成:{filename}")
```

### 5.3 读取Word内容

```python
from docx import Document

doc = Document('报告.docx')

# 读取所有段落
for para in doc.paragraphs:
    print(para.text)

# 读取表格
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            print(cell.text, end='\t')
        print()
```

## 六、PPT自动化(python-pptx)

### 6.1 创建PPT

**基础操作:**
```python
from pptx import Presentation
from pptx.util import Inches, Pt

# 创建演示文稿
prs = Presentation()

# 添加标题幻灯片
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = "季度销售报告"
subtitle.text = "2025年Q1"

# 添加内容幻灯片
bullet_slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(bullet_slide_layout)
shapes = slide.shapes
title_shape = shapes.title
body_shape = shapes.placeholders[1]

title_shape.text = "销售数据"
tf = body_shape.text_frame
tf.text = "产品A销售额:100万"

p = tf.add_paragraph()
p.text = "产品B销售额:150万"
p.level = 1

# 保存
prs.save('季度报告.pptx')
```

### 6.2 批量生成PPT

**示例:根据数据生成报告PPT**
```python
from pptx import Presentation
from pptx.util import Inches
import pandas as pd

# 读取数据
df = pd.read_excel('月度数据.xlsx')

# 创建PPT
prs = Presentation()

# 封面
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "月度销售报告"

# 为每个产品创建一页
for _, row in df.iterrows():
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # 空白布局

    # 添加标题
    left = Inches(1)
    top = Inches(1)
    width = Inches(8)
    height = Inches(1)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.text = f"产品:{row['产品名称']}"

    # 添加数据
    top = Inches(2)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.text = f"销售额:{row['销售额']}万元\n增长率:{row['增长率']}%"

# 保存
prs.save('自动生成报告.pptx')
```

## 七、综合实战案例

### 7.1 自动化月报生成系统

**需求:**
```
1. 从多个Excel文件读取数据
2. 数据清洗和汇总
3. 生成Excel报表
4. 生成Word分析报告
5. 生成PPT汇报材料
6. 发送邮件
```

**完整代码:**
```python
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font
from docx import Document
from pptx import Presentation
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def merge_excel_files(folder):
    """合并Excel文件"""
    files = [f for f in os.listdir(folder) if f.endswith('.xlsx')]
    dfs = []
    for file in files:
        df = pd.read_excel(os.path.join(folder, file))
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

def generate_excel_report(df, filename):
    """生成Excel报表"""
    # 分组汇总
    summary = df.groupby('产品').agg({
        '销售额': 'sum',
        '数量': 'sum'
    }).reset_index()

    # 保存
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='明细', index=False)
        summary.to_excel(writer, sheet_name='汇总', index=False)

    print(f"Excel报表生成:{filename}")

def generate_word_report(df, filename):
    """生成Word报告"""
    doc = Document()
    doc.add_heading('月度销售分析报告', level=1)

    # 总体情况
    doc.add_heading('一、总体情况', level=2)
    total_sales = df['销售额'].sum()
    doc.add_paragraph(f"本月总销售额:{total_sales:,.2f}元")

    # 产品分析
    doc.add_heading('二、产品分析', level=2)
    summary = df.groupby('产品')['销售额'].sum().sort_values(ascending=False)
    for product, sales in summary.items():
        doc.add_paragraph(f"{product}:{sales:,.2f}元", style='List Bullet')

    doc.save(filename)
    print(f"Word报告生成:{filename}")

def generate_ppt_report(df, filename):
    """生成PPT报告"""
    prs = Presentation()

    # 封面
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "月度销售报告"

    # 数据页
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "销售数据"

    summary = df.groupby('产品')['销售额'].sum()
    body = slide.placeholders[1].text_frame
    for product, sales in summary.items():
        p = body.add_paragraph()
        p.text = f"{product}:{sales:,.0f}元"

    prs.save(filename)
    print(f"PPT报告生成:{filename}")

# 主流程
if __name__ == "__main__":
    # 1. 合并数据
    df = merge_excel_files('D:\\原始数据\\')

    # 2. 生成Excel报表
    generate_excel_report(df, '月度销售报表.xlsx')

    # 3. 生成Word报告
    generate_word_report(df, '月度销售报告.docx')

    # 4. 生成PPT
    generate_ppt_report(df, '月度销售汇报.pptx')

    print("所有报告生成完毕!")
```

## 实用技巧

### 技巧1:异常处理
```python
try:
    df = pd.read_excel('数据.xlsx')
except FileNotFoundError:
    print("文件不存在")
except Exception as e:
    print(f"发生错误:{e}")
```

### 技巧2:进度显示
```python
from tqdm import tqdm

for i in tqdm(range(100)):
    # 处理数据
    pass
```

### 技巧3:日志记录
```python
import logging

logging.basicConfig(
    filename='process.log',
    level=logging.INFO,
    format='%(asctime)s - %(message)s'
)

logging.info("开始处理")
logging.error("发生错误")
```

## 常见问题

**Q1:Python和VBA哪个更好?**
A:各有优势。VBA集成度高,Python更灵活强大。建议都学。

**Q2:学Python需要多久?**
A:基础语法1-2天,办公自动化1-2周实践即可上手。

**Q3:遇到问题怎么办?**
A:善用搜索引擎、Stack Overflow、ChatGPT。

## 检查清单

- [ ] 成功安装Python环境
- [ ] 掌握Python基础语法
- [ ] 能用Python读写Excel
- [ ] 能用Python生成Word文档
- [ ] 能用Python创建PPT
- [ ] 完成至少1个自动化项目
- [ ] 建立个人Python工具库

## 总结

Python办公自动化是一项值得投资的技能:

**核心价值:**
1. 更灵活强大的自动化能力
2. 跨平台支持
3. 丰富的第三方库
4. 可扩展到数据分析、AI等领域

**下一步行动:**
- 完成环境搭建
- 练习基础语法
- 实现一个小项目
- 持续学习提升

**下一章预告:**
《61-文件命名与组织系统》将分享高效的文件管理方法。

