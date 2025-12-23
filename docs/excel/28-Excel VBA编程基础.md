---
title: Excel VBA编程基础
slug: /excel/28-Excel-VBA编程基础
sidebar_position: 28
---
# Excel VBA编程基础

## 本章概览

VBA(Visual Basic for Applications)是Excel的编程语言,能够实现自动化操作、批量处理、自定义功能等。掌握VBA可以大幅提升工作效率。

**学习目标**
- 掌握VBA编辑器和基本语法
- 学会录制和编辑宏
- 熟练使用VBA对象模型
- 掌握常用VBA代码片段
- 学会调试和错误处理

---

## 27.1 VBA入门

### 28.1.1 启用开发工具选项卡

**显示开发工具**
```
文件 → 选项 → 自定义功能区
右侧勾选"开发工具"
确定
```

**开发工具功能**
```
- Visual Basic:打开VBA编辑器
- 宏:录制/查看/执行宏
- 录制宏:开始录制
- 使用相对引用:录制相对位置宏
- 宏安全性:设置宏安全级别
- 控件:插入窗体控件/ActiveX控件
```

### 28.1.2 VBA编辑器界面

**打开VBA编辑器**
```
开发工具 → Visual Basic
或快捷键: Alt + F11
```

**界面组成**
```
1. 项目资源管理器(左上):
   - VBAProject(工作簿名)
     ├ Microsoft Excel对象
     │ ├ ThisWorkbook(工作簿对象)
     │ └ Sheet1, Sheet2...(工作表对象)
     └ 模块
       └ 模块1, 模块2...

2. 属性窗口(左下):
   显示选中对象的属性

3. 代码窗口(右侧):
   编写VBA代码

4. 立即窗口(底部):
   调试和测试代码
   快捷键: Ctrl + G
```

### 28.1.3 第一个VBA程序

**插入模块**
```
插入 → 模块
在项目资源管理器中出现"模块1"
```

**编写代码**
```vba
Sub 第一个程序()
    MsgBox "Hello, VBA!"
End Sub
```

**运行程序**
```
方法1: F5
方法2: 运行 → 运行子过程/用户窗体
方法3: Excel中 开发工具 → 宏 → 选择宏名 → 运行
```

---

## 28.2 录制宏

### 28.2.1 录制宏基础

**开始录制**
```
开发工具 → 录制宏
设置:
- 宏名:MyMacro1(不能有空格)
- 快捷键:Ctrl + Shift + M(可选)
- 保存位置:当前工作簿
- 说明:宏的描述
```

**执行操作**
```
示例:格式化表格
1. 选中A1:D10
2. 设置字体:黑体,12号
3. 添加边框
4. 填充颜色:浅蓝色
```

**停止录制**
```
开发工具 → 停止录制
```

**查看代码**
```
开发工具 → 宏 → 选择录制的宏 → 编辑

生成的代码示例:
Sub MyMacro1()
    Range("A1:D10").Select
    With Selection
        .Font.Name = "黑体"
        .Font.Size = 12
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(173, 216, 230)
    End With
End Sub
```

### 28.2.2 相对引用录制

**使用相对引用**
```
开发工具 → 使用相对引用(切换状态)
录制宏

效果:
- 绝对引用:Range("A1")
- 相对引用:ActiveCell.Offset(1, 0)

应用:
宏可在不同位置执行,不固定在A1
```

**示例**
```vba
# 相对引用宏:向下移动并输入
Sub 相对移动()
    ActiveCell.Offset(1, 0).Select  # 下移一行
    ActiveCell.Value = "新数据"
End Sub

可以在任意位置执行,不固定单元格
```

---

## 28.3 VBA基本语法

### 28.3.1 变量和数据类型

**声明变量**
```vba
Dim 变量名 As 数据类型

示例:
Dim i As Integer        ' 整数
Dim name As String      ' 文本
Dim price As Double     ' 小数
Dim isActive As Boolean ' 布尔值
Dim today As Date       ' 日期
```

**数据类型**
```vba
Integer      ' 整数 (-32768 to 32767)
Long         ' 长整数
Single       ' 单精度浮点
Double       ' 双精度浮点
String       ' 字符串
Boolean      ' True/False
Date         ' 日期和时间
Variant      ' 任意类型(默认)
Object       ' 对象
```

**赋值**
```vba
i = 10
name = "张三"
price = 99.99
isActive = True
today = Date  ' 当前日期
```

**常量**
```vba
Const PI As Double = 3.14159
Const TAX_RATE As Double = 0.13
```

### 28.3.2 运算符

**算术运算符**
```vba
+   加
-   减
*   乘
/   除
\   整除
Mod 取余
^   乘方

示例:
result = 10 + 5  ' 15
result = 10 \ 3  ' 3(整除)
result = 10 Mod 3  ' 1(余数)
```

**比较运算符**
```vba
=   等于
<>  不等于
<   小于
>   大于
<=  小于等于
>=  大于等于

示例:
If age >= 18 Then
    MsgBox "成年"
End If
```

**逻辑运算符**
```vba
And  与
Or   或
Not  非

示例:
If age >= 18 And score >= 60 Then
    MsgBox "合格"
End If
```

### 28.3.3 流程控制

**If...Then...Else**
```vba
' 单条件
If score >= 60 Then
    MsgBox "及格"
End If

' 双分支
If score >= 60 Then
    MsgBox "及格"
Else
    MsgBox "不及格"
End If

' 多分支
If score >= 90 Then
    grade = "优秀"
ElseIf score >= 80 Then
    grade = "良好"
ElseIf score >= 70 Then
    grade = "中等"
ElseIf score >= 60 Then
    grade = "及格"
Else
    grade = "不及格"
End If
```

**Select Case**
```vba
Select Case score
    Case Is >= 90
        grade = "优秀"
    Case Is >= 80
        grade = "良好"
    Case Is >= 70
        grade = "中等"
    Case Is >= 60
        grade = "及格"
    Case Else
        grade = "不及格"
End Select
```

**For...Next循环**
```vba
' 基础循环
For i = 1 To 10
    Cells(i, 1).Value = i
Next i

' 步长
For i = 1 To 10 Step 2  ' 1, 3, 5, 7, 9
    Debug.Print i
Next i

' 倒序
For i = 10 To 1 Step -1
    Debug.Print i
Next i
```

**For Each循环**
```vba
' 遍历区域
Dim cell As Range
For Each cell In Range("A1:A10")
    cell.Value = cell.Value * 2
Next cell

' 遍历工作表
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Range("A1").Value = ws.Name
Next ws
```

**Do...Loop循环**
```vba
' Do While(先判断)
Dim i As Integer
i = 1
Do While i <= 10
    Cells(i, 1).Value = i
    i = i + 1
Loop

' Do Until
i = 1
Do Until i > 10
    Cells(i, 1).Value = i
    i = i + 1
Loop

' Do...Loop While(后判断)
i = 1
Do
    Cells(i, 1).Value = i
    i = i + 1
Loop While i <= 10
```

---

## 28.4 VBA对象模型

### 28.4.1 Excel对象层次

**层次结构**
```
Application (Excel应用程序)
└ Workbooks (工作簿集合)
  └ Workbook (单个工作簿)
    └ Worksheets (工作表集合)
      └ Worksheet (单个工作表)
        └ Range (单元格区域)
```

**引用方法**
```vba
' Application
Application.ScreenUpdating = False

' Workbooks
Workbooks("工作簿1.xlsx")
Workbooks(1)  ' 第一个工作簿
ThisWorkbook  ' 代码所在工作簿
ActiveWorkbook  ' 当前激活工作簿

' Worksheets
Worksheets("Sheet1")
Worksheets(1)  ' 第一个工作表
ActiveSheet  ' 当前激活工作表

' Range
Range("A1")
Range("A1:D10")
Cells(1, 1)  ' 第1行第1列
Cells(1, "A")  ' 第1行A列
```

### 28.4.2 Range对象

**引用单元格**
```vba
' 单个单元格
Range("A1").Value = 100
Cells(1, 1).Value = 100

' 区域
Range("A1:D10").Value = "批量填充"

' 选中区域
Selection.Value = "选中的区域"
ActiveCell.Value = "活动单元格"

' 相对引用
ActiveCell.Offset(1, 0).Value = "下一行"
ActiveCell.Offset(0, 1).Value = "右一列"

' Resize
Range("A1").Resize(5, 3).Value = "5行3列"
```

**Range属性**
```vba
' Value
Range("A1").Value = 100
Range("A1").Value2 = 100  ' 无格式值

' Text(只读)
txt = Range("A1").Text  ' 显示的文本(含格式)

' Formula
Range("A1").Formula = "=SUM(B1:B10)"
Range("A1").FormulaR1C1 = "=SUM(R[-1]C:R[-10]C)"

' 格式
Range("A1").Font.Name = "黑体"
Range("A1").Font.Size = 12
Range("A1").Font.Bold = True
Range("A1").Font.Color = RGB(255, 0, 0)
Range("A1").Interior.Color = RGB(255, 255, 0)

' 边框
Range("A1:D10").Borders.LineStyle = xlContinuous
Range("A1:D10").Borders.Weight = xlMedium
```

**Range方法**
```vba
' Clear
Range("A1:D10").Clear  ' 清除所有
Range("A1:D10").ClearContents  ' 清除内容
Range("A1:D10").ClearFormats  ' 清除格式

' Copy / Paste
Range("A1:D10").Copy
Range("F1").PasteSpecial xlPasteValues  ' 粘贴值

' Delete / Insert
Range("A1").EntireRow.Delete  ' 删除行
Range("A1").EntireColumn.Delete  ' 删除列
Range("A1").EntireRow.Insert  ' 插入行

' Select
Range("A1:D10").Select
```

### 28.4.3 Worksheet对象

**工作表操作**
```vba
' 添加工作表
Worksheets.Add
Worksheets.Add After:=Worksheets(Worksheets.Count)  ' 末尾添加

' 删除工作表
Application.DisplayAlerts = False  ' 禁用警告
Worksheets("Sheet2").Delete
Application.DisplayAlerts = True

' 重命名
Worksheets("Sheet1").Name = "销售数据"

' 复制
Worksheets("Sheet1").Copy After:=Worksheets("Sheet2")

' 移动
Worksheets("Sheet1").Move Before:=Worksheets("Sheet2")

' 隐藏/显示
Worksheets("Sheet1").Visible = xlSheetHidden  ' 隐藏
Worksheets("Sheet1").Visible = xlSheetVisible  ' 显示
Worksheets("Sheet1").Visible = xlSheetVeryHidden  ' 彻底隐藏

' 保护
Worksheets("Sheet1").Protect Password:="123456"
Worksheets("Sheet1").Unprotect Password:="123456"
```

### 28.4.4 Workbook对象

**工作簿操作**
```vba
' 新建
Workbooks.Add

' 打开
Workbooks.Open "C:\数据.xlsx"

' 保存
ThisWorkbook.Save
ThisWorkbook.SaveAs "C:\新文件.xlsx"

' 关闭
Workbooks("文件名.xlsx").Close SaveChanges:=True
ThisWorkbook.Close

' 激活
Workbooks("文件名.xlsx").Activate
```

---

## 28.5 常用代码片段

### 28.5.1 批量操作

**批量填充**
```vba
Sub 批量填充()
    Dim i As Integer
    For i = 1 To 100
        Cells(i, 1).Value = i
        Cells(i, 2).Value = "产品" & i
        Cells(i, 3).Value = Rnd() * 100  ' 随机数
    Next i
End Sub
```

**批量格式化**
```vba
Sub 格式化表格()
    With Range("A1:D10")
        .Font.Name = "微软雅黑"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(242, 242, 242)
    End With

    ' 标题行特殊格式
    With Range("A1:D1")
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(79, 129, 189)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub
```

### 28.5.2 数据处理

**查找替换**
```vba
Sub 查找替换()
    Cells.Replace What:="北京市", Replacement:="北京", LookAt:=xlPart
End Sub
```

**删除空行**
```vba
Sub 删除空行()
    Dim i As Long
    Dim lastRow As Long

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = lastRow To 1 Step -1
        If WorksheetFunction.CountA(Rows(i)) = 0 Then
            Rows(i).Delete
        End If
    Next i
End Sub
```

**去除重复**
```vba
Sub 去除重复()
    Range("A1").CurrentRegion.RemoveDuplicates _
        Columns:=1, Header:=xlYes
End Sub
```

### 28.5.3 文件操作

**遍历文件夹**
```vba
Sub 遍历文件()
    Dim folderPath As String
    Dim fileName As String

    folderPath = "C:\数据\"
    fileName = Dir(folderPath & "*.xlsx")

    Do While fileName <> ""
        Debug.Print fileName
        ' 处理文件...
        fileName = Dir  ' 下一个文件
    Loop
End Sub
```

**合并工作簿**
```vba
Sub 合并工作簿()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long

    folderPath = "C:\数据\"
    fileName = Dir(folderPath & "*.xlsx")

    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)
        Set ws = wb.Worksheets(1)

        ' 复制数据到主工作簿
        lastRow = ThisWorkbook.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
        ws.UsedRange.Copy ThisWorkbook.Worksheets(1).Cells(lastRow + 1, 1)

        wb.Close SaveChanges:=False
        fileName = Dir
    Loop
End Sub
```

---

## 28.6 用户窗体

### 28.6.1 创建用户窗体

**插入窗体**
```
VBA编辑器 → 插入 → 用户窗体
```

**添加控件**
```
工具箱:
- Label(标签)
- TextBox(文本框)
- CommandButton(按钮)
- ComboBox(下拉框)
- ListBox(列表框)
- CheckBox(复选框)
- OptionButton(单选按钮)
```

**示例:简单输入窗体**
```
控件布局:
Label1: "姓名:"
TextBox1: (输入框)
Label2: "年龄:"
TextBox2: (输入框)
CommandButton1: "确定"
CommandButton2: "取消"
```

**窗体代码**
```vba
Private Sub CommandButton1_Click()
    ' 确定按钮
    Dim name As String
    Dim age As Integer

    name = TextBox1.Value
    age = TextBox2.Value

    ' 写入工作表
    Range("A1").Value = name
    Range("B1").Value = age

    ' 关闭窗体
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    ' 取消按钮
    Unload Me
End Sub
```

**显示窗体**
```vba
Sub 显示窗体()
    UserForm1.Show
End Sub
```

---

## 28.7 调试和错误处理

### 28.7.1 调试工具

**Debug.Print**
```vba
Sub 调试示例()
    Dim i As Integer
    For i = 1 To 5
        Debug.Print "当前值: " & i
    Next i
End Sub

' 在立即窗口(Ctrl+G)查看输出
```

**断点**
```
在代码行左侧单击设置断点(红点)
F8: 逐句执行
F5: 继续执行
```

**监视窗口**
```
调试 → 添加监视
监视变量的值变化
```

### 28.7.2 错误处理

**On Error语句**
```vba
Sub 错误处理示例()
    On Error GoTo ErrorHandler

    ' 可能出错的代码
    Dim result As Double
    result = 10 / 0  ' 除零错误

    Exit Sub

ErrorHandler:
    MsgBox "发生错误: " & Err.Description
End Sub
```

**On Error Resume Next**
```vba
Sub 忽略错误()
    On Error Resume Next

    ' 即使出错也继续执行
    Range("不存在的名称").Value = 100

    ' 恢复正常错误处理
    On Error GoTo 0
End Sub
```

**Err对象**
```vba
Sub Err对象()
    On Error GoTo ErrorHandler

    ' 代码...

ErrorHandler:
    MsgBox "错误号: " & Err.Number & vbCrLf & _
           "错误描述: " & Err.Description
    Err.Clear  ' 清除错误
End Sub
```

---

## 28.8 性能优化

### 28.8.1 优化技巧

**关闭屏幕更新**
```vba
Sub 优化示例()
    Application.ScreenUpdating = False  ' 关闭屏幕更新
    Application.Calculation = xlCalculationManual  ' 手动计算
    Application.EnableEvents = False  ' 禁用事件

    ' 大量操作...
    Dim i As Long
    For i = 1 To 10000
        Cells(i, 1).Value = i
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

**使用数组**
```vba
Sub 数组优化()
    Dim arr() As Variant
    Dim i As Long

    ' 一次性读取到数组
    arr = Range("A1:A10000").Value

    ' 处理数组
    For i = 1 To UBound(arr)
        arr(i, 1) = arr(i, 1) * 2
    Next i

    ' 一次性写回
    Range("B1:B10000").Value = arr
End Sub
```

**避免Select和Activate**
```vba
' 慢(不推荐)
Range("A1").Select
Selection.Value = 100

' 快(推荐)
Range("A1").Value = 100
```

---

## 本章小结

**核心要点**
1. **VBA编辑器**:Alt+F11,模块中编写代码
2. **录制宏**:快速生成代码,需要编辑优化
3. **基本语法**:变量、循环、条件判断
4. **对象模型**:Application→Workbook→Worksheet→Range
5. **常用操作**:批量处理、文件操作、用户窗体
6. **调试**:Debug.Print、断点、错误处理
7. **优化**:关闭屏幕更新、使用数组、避免Select

**VBA学习路径**
```
1. 录制宏 → 查看代码 → 理解语法
2. 学习对象模型和属性方法
3. 编写简单自动化脚本
4. 学习用户窗体设计
5. 掌握调试和错误处理
6. 性能优化和高级技巧
```

**学习建议**
- 从录制宏开始,边学边改
- 多查帮助文档(F1)
- 善用Debug.Print调试
- 参考他人优秀代码
- 多做实际项目练习

**下一步学习**
- 第29章:Excel与Python集成
- 第30章:Excel自动化办公
- VBA高级主题:类模块、API调用

---

## 思考练习

1. 录制一个格式化表格的宏,并优化代码
2. 编写VBA批量生成100行测试数据
3. 创建用户窗体实现数据录入功能
4. 编写代码合并文件夹中的所有Excel文件
5. 实现自动生成月度报表的宏

**练习提示**
- 先用录制宏生成基础代码
- 逐步添加变量和循环
- 添加错误处理机制
- 测试不同场景下的表现
- 优化代码性能

