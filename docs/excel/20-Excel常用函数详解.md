---
title: Excel常用函数详解
slug: /excel/20-Excel常用函数详解
sidebar_position: 20
---
# Excel常用函数详解

## 引言

函数是Excel的灵魂,掌握常用函数可以解决80%的数据处理问题。本章将系统讲解职场最常用的函数及其实战应用。

## 一、函数基础

### 1.1 函数的结构

**基本语法:**
```
=函数名(参数1, 参数2, ...)

示例:
=SUM(A1:A10)
=IF(B2>100,"合格","不合格")
```

**参数类型:**
1. **数值**:10, 100, -50
2. **文本**:"张三", "合格"(需加引号)
3. **逻辑值**:TRUE, FALSE
4. **单元格引用**:A1, B2:C10
5. **数组**:&#123;1,2,3;4,5,6&#125;
6. **函数嵌套**:IF(SUM(A1:A10)&gt;100,"是","否")

### 1.2 函数输入方法

**方法1:直接输入**
```
1. 单元格输入=
2. 输入函数名
3. 输入左括号(
4. Excel显示参数提示
5. 输入参数
6. 输入右括号)
7. Enter确认
```

**方法2:插入函数(Shift+F3)**
```
1. 点击"插入函数"按钮
2. 搜索或选择函数
3. 填写参数对话框
4. 确定
```

**方法3:自动求和(Alt+=)**
```
快速插入SUM、AVERAGE等常用函数
```

### 1.3 常见错误

| 错误值 | 含义 | 常见原因 |
|--------|------|----------|
| #DIV/0! | 除数为零 | 公式中除以0或空单元格 |
| #N/A | 值不可用 | VLOOKUP找不到匹配值 |
| #NAME? | 名称错误 | 函数名拼写错误 |
| #NULL! | 空值错误 | 区域交集为空 |
| #NUM! | 数值错误 | 数值参数超出范围 |
| #REF! | 引用错误 | 引用的单元格被删除 |
| #VALUE! | 值错误 | 参数类型错误 |
| ##### | 列宽不够 | 加宽列即可 |

## 二、数学与统计函数

### 2.1 SUM系列

**SUM - 求和**
```
语法:=SUM(number1, [number2], ...)

示例:
=SUM(A1:A10)          单个区域求和
=SUM(A1:A10, C1:C10)  多个区域求和
=SUM(A1, A3, A5)      不连续单元格求和
```

**SUMIF - 单条件求和**
```
语法:=SUMIF(range, criteria, [sum_range])

参数:
- range:条件判断区域
- criteria:条件
- sum_range:求和区域(可选)

示例:
=SUMIF(A:A, "北京", B:B)     北京地区销售额求和
=SUMIF(B:B, ">1000", C:C)     销量>1000的金额求和
=SUMIF(A:A, A2, B:B)          与A2相同的所有值求和
```

**SUMIFS - 多条件求和(Excel 2007+)**
```
语法:=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)

示例:
=SUMIFS(D:D, A:A, "北京", B:B, "电脑")
解释:计算北京地区电脑的销售额

=SUMIFS(C:C, A:A, ">=2025-1-1", A:A, "<=2025-12-31")
解释:计算2025年的总销售额
```

### 2.2 AVERAGE系列

**AVERAGE - 平均值**
```
=AVERAGE(A1:A10)           算术平均值
=AVERAGEIF(A:A, ">60")     条件平均值
=AVERAGEIFS(C:C, A:A, "北京", B:B, ">100")  多条件平均值
```

### 2.3 COUNT系列

**COUNT - 计数**
```
=COUNT(A1:A10)         数值单元格个数
=COUNTA(A1:A10)        非空单元格个数
=COUNTBLANK(A1:A10)    空单元格个数
```

**COUNTIF - 条件计数**
```
=COUNTIF(A:A, "男")          男性人数
=COUNTIF(B:B, ">60")         分数>60的人数
=COUNTIF(C:C, "北*")         北字开头的城市数量
```

**COUNTIFS - 多条件计数**
```
=COUNTIFS(A:A, "北京", B:B, ">1000")
解释:北京地区销售额>1000的订单数
```

### 2.4 MAX/MIN - 最大最小值

```
=MAX(A1:A10)          最大值
=MIN(A1:A10)          最小值
=LARGE(A1:A10, 2)     第2大值
=SMALL(A1:A10, 3)     第3小值
```

### 2.5 ROUND系列 - 四舍五入

```
=ROUND(A1, 2)         四舍五入到2位小数
=ROUNDUP(A1, 0)       向上取整到整数
=ROUNDDOWN(A1, 0)     向下取整到整数
=INT(A1)              取整数部分
=MOD(A1, 2)           取余数
```

## 三、逻辑函数

### 3.1 IF - 条件判断

**基础用法:**
```
语法:=IF(logical_test, value_if_true, value_if_false)

示例:
=IF(A1>60, "及格", "不及格")
=IF(A1="", "空", "非空")
=IF(B1>0, B1, 0)     负数变0
```

**IF嵌套:**
```
多条件判断:
=IF(A1>=90, "优秀", IF(A1>=80, "良好", IF(A1>=60, "及格", "不及格")))

简化写法(Excel 2019+):
=IFS(A1>=90,"优秀", A1>=80,"良好", A1>=60,"及格", TRUE,"不及格")
```

### 3.2 AND/OR/NOT - 逻辑组合

**AND - 与(全部满足)**
```
=AND(A1>60, B1>60)        A1和B1都>60才返回TRUE
=IF(AND(A1>60, B1>60), "通过", "不通过")
```

**OR - 或(任一满足)**
```
=OR(A1>60, B1>60)         A1或B1任一>60就返回TRUE
=IF(OR(A1="", B1=""), "有空值", "无空值")
```

**NOT - 非(取反)**
```
=NOT(A1>60)               A1不大于60
=IF(NOT(ISBLANK(A1)), A1, 0)   不为空则显示值
```

### 3.3 IFERROR - 错误处理

```
语法:=IFERROR(value, value_if_error)

应用:
=IFERROR(VLOOKUP(A1,数据表,2,0), "未找到")
=IFERROR(A1/B1, 0)          除法错误时返回0
```

## 四、文本函数

### 4.1 LEFT/RIGHT/MID - 文本截取

**LEFT - 左边取值**
```
=LEFT(A1, 3)              取左边3个字符
=LEFT("北京市朝阳区", 3)  结果:"北京市"
```

**RIGHT - 右边取值**
```
=RIGHT(A1, 4)             取右边4个字符
=RIGHT("13812345678", 4)  结果:"5678"(后4位)
```

**MID - 中间取值**
```
=MID(A1, 4, 2)            从第4个字符开始取2个
=MID("北京市朝阳区", 4, 3)  结果:"朝阳区"
```

### 4.2 LEN - 文本长度

```
=LEN(A1)                  文本长度(字符数)
=IF(LEN(A1)>10, "太长", "正常")   长度检查
```

### 4.3 FIND/SEARCH - 查找文本

**FIND - 精确查找(区分大小写)**
```
=FIND("北京", A1)         返回"北京"在A1中的位置
=IF(ISNUMBER(FIND("北京",A1)), "包含", "不包含")
```

**SEARCH - 模糊查找(不区分大小写,支持通配符)**
```
=SEARCH("北*", A1)        查找以"北"开头的文本
```

### 4.4 CONCATENATE/&  - 文本合并

**&符号(推荐)**
```
=A1&B1                    简单合并
=A1&" "&B1                中间加空格
="姓名:"&A1&",年龄:"&B1  组合文本和值
```

**TEXTJOIN(Excel 2019+)**
```
=TEXTJOIN(",", TRUE, A1:A10)    用逗号连接,忽略空值
=TEXTJOIN(" ", FALSE, A1, B1, C1)   用空格连接
```

### 4.5 TRIM/CLEAN - 清理文本

```
=TRIM(A1)                 删除多余空格(保留单词间空格)
=CLEAN(A1)                删除不可打印字符
```

### 4.6 UPPER/LOWER/PROPER - 大小写转换

```
=UPPER(A1)                全部大写:HELLO
=LOWER(A1)                全部小写:hello
=PROPER(A1)               首字母大写:Hello World
```

### 4.7 SUBSTITUTE/REPLACE - 替换

**SUBSTITUTE - 内容替换**
```
=SUBSTITUTE(A1, "旧文本", "新文本")
=SUBSTITUTE("北京市朝阳区", "市", "")   删除"市"
=SUBSTITUTE(A1, " ", "")                删除所有空格
```

**REPLACE - 位置替换**
```
=REPLACE(A1, 起始位置, 字符数, 新文本)
=REPLACE("13812345678", 4, 4, "****")   手机号脱敏:138****5678
```

## 五、日期和时间函数

### 5.1 TODAY/NOW - 当前日期时间

```
=TODAY()                  当前日期:2025-12-21
=NOW()                    当前日期时间:2025-12-21 10:30:45
```

### 5.2 YEAR/MONTH/DAY - 提取日期部分

```
=YEAR(A1)                 提取年份:2025
=MONTH(A1)                提取月份:12
=DAY(A1)                  提取日:21
=WEEKDAY(A1)              星期几(1-7,周日为1)
=TEXT(A1, "aaaa")         星期几中文:星期六
```

### 5.3 DATE - 组合日期

```
=DATE(2025, 12, 21)       生成日期:2025-12-21
=DATE(YEAR(A1), MONTH(A1)+1, DAY(A1))  下个月同一天
```

### 5.4 DATEDIF - 日期差(隐藏函数)

```
语法:=DATEDIF(start_date, end_date, unit)

Unit:
- "Y":年数差
- "M":月数差
- "D":天数差
- "YM":忽略年的月数差
- "YD":忽略年的天数差
- "MD":忽略年月的天数差

示例:
=DATEDIF("2020-1-1", "2025-12-21", "Y")    5年
=DATEDIF(A1, TODAY(), "Y")                  年龄计算
```

### 5.5 EDATE/EOMONTH - 月份计算

```
=EDATE(A1, 3)             3个月后的日期
=EDATE(A1, -1)            1个月前的日期
=EOMONTH(A1, 0)           本月最后一天
=EOMONTH(A1, 1)           下月最后一天
```

### 5.6 NETWORKDAYS - 工作日计算

```
=NETWORKDAYS(开始日期, 结束日期, [节假日])
计算两个日期之间的工作日天数(排除周末和节假日)

示例:
=NETWORKDAYS("2025-1-1", "2025-12-31")    2025年工作日天数
```

## 六、查找与引用函数

### 6.1 VLOOKUP - 垂直查找(最常用)

**语法:**
```
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])

参数:
- lookup_value:要查找的值
- table_array:数据表区域
- col_index_num:返回第几列(从1开始)
- range_lookup:FALSE精确匹配,TRUE近似匹配
```

**示例:**
```
员工表(A:C):
工号 | 姓名 | 部门
1001 | 张三 | 销售
1002 | 李四 | 技术

查找公式:
=VLOOKUP(1001, A:C, 2, FALSE)    返回:"张三"
=VLOOKUP(1001, A:C, 3, 0)        返回:"销售"(0等同FALSE)
```

**常见问题:**
1. #N/A错误:找不到匹配值 → 检查查找值是否存在
2. 只能向右查找 → 使用INDEX+MATCH替代
3. 查找值必须在第一列 → 调整数据顺序或用其他函数

### 6.2 HLOOKUP - 水平查找

```
=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])

与VLOOKUP类似,但是横向查找
```

### 6.3 INDEX - 返回指定位置的值

**语法:**
```
=INDEX(array, row_num, [column_num])

示例:
=INDEX(A1:C10, 5, 2)     第5行第2列的值
=INDEX(B:B, 10)          B列第10个值
```

### 6.4 MATCH - 返回位置

**语法:**
```
=MATCH(lookup_value, lookup_array, [match_type])

match_type:
- 1:小于等于(默认,需升序)
- 0:精确匹配(推荐)
- -1:大于等于(需降序)

示例:
=MATCH("张三", A:A, 0)   返回"张三"在A列的行号
```

### 6.5 INDEX+MATCH组合(VLOOKUP替代)

**优势:**
- 可以左右查找
- 查找列可以改变
- 性能更好

**语法:**
```
=INDEX(返回列, MATCH(查找值, 查找列, 0))

示例:
=INDEX(B:B, MATCH(E1, A:A, 0))
解释:在A列查找E1的值,返回对应B列的内容
```

### 6.6 XLOOKUP(Excel 365/2021+)

**新一代查找函数,取代VLOOKUP**
```
语法:=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])

优势:
- 双向查找
- 返回整行/整列
- 自定义未找到提示
- 支持通配符
- 支持倒序查找

示例:
=XLOOKUP(E1, A:A, B:B, "未找到")
=XLOOKUP(E1, A:A, B:D)   返回多列
```

## 七、实战案例

### 7.1 员工考勤统计

**需求:**统计每位员工的出勤天数、迟到次数、请假天数

**公式:**
```
出勤天数:
=COUNTIF(C2:AF2, "√")

迟到次数:
=COUNTIF(C2:AF2, "迟到")

请假天数:
=COUNTIFS(C2:AF2, "事假")+COUNTIFS(C2:AF2, "病假")
```

### 7.2 销售提成计算

**需求:**
- 销售额≤10000:5%提成
- 10000&lt;销售额≤50000:8%提成
- 销售额&gt;50000:10%提成

**公式:**
```
方法1:IF嵌套
=IF(B2<=10000, B2*5%, IF(B2<=50000, B2*8%, B2*10%))

方法2:IFS(Excel 2019+)
=IFS(B2<=10000, B2*5%, B2<=50000, B2*8%, B2>50000, B2*10%)
```

### 7.3 身份证信息提取

**需求:**从18位身份证号提取出生日期、性别、年龄

**公式:**
```
出生日期:
=TEXT(MID(A2,7,8), "0000-00-00")

性别:
=IF(MOD(MID(A2,17,1), 2)=1, "男", "女")

年龄:
=DATEDIF(TEXT(MID(A2,7,8),"0000-00-00"), TODAY(), "Y")
```

### 7.4 多表联查

**需求:**从销售表和产品表获取完整信息

**公式:**
```
VLOOKUP单表查找:
=VLOOKUP(A2, 产品表!A:D, 2, 0)   产品名称
=VLOOKUP(A2, 产品表!A:D, 3, 0)   产品价格

INDEX+MATCH跨表查找:
=INDEX(产品表!B:B, MATCH(A2, 产品表!A:A, 0))
```

### 7.5 自动排名

**需求:**给成绩排名,并列情况跳过名次

**公式:**
```
=RANK(B2, $B$2:$B$100, 0)
参数:
- B2:要排名的值
- $B$2:$B$100:排名范围(绝对引用)
- 0:降序(分数高排名前)
```

## 八、函数学习建议

### 8.1 学习路线

**第1周:基础函数(10个)**
- SUM、AVERAGE、COUNT、COUNTA
- MAX、MIN
- IF
- LEFT、RIGHT、LEN

**第2周:进阶函数(10个)**
- SUMIF、SUMIFS
- COUNTIF、COUNTIFS
- VLOOKUP
- TEXT、DATE、YEAR、MONTH

**第3周:高级函数(10个)**
- INDEX、MATCH
- IFERROR
- SUBSTITUTE、FIND
- DATEDIF、NETWORKDAYS

**第4周:综合应用**
- 函数嵌套
- 实战案例
- 错误排查

### 8.2 记忆技巧

**分类记忆:**
- 数学统计:SUM系列、COUNT系列
- 逻辑:IF系列、AND/OR
- 文本:LEFT/MID/RIGHT、LEN、FIND
- 日期:DATE系列、YEAR/MONTH/DAY
- 查找:VLOOKUP、INDEX+MATCH

**联想记忆:**
- SUM:求和
- AVERAGE:平均
- COUNT:计数
- IF:如果
- LEFT:左边
- FIND:查找

### 8.3 避坑指南

**错误1:参数类型错误**
```
错误:=SUM("100", "200")  (文本无法求和)
正确:=SUM(100, 200)
```

**错误2:区域引用未锁定**
```
错误:=VLOOKUP(A2, B2:D10, 2, 0)  (下拉公式区域会移动)
正确:=VLOOKUP(A2, $B$2:$D$10, 2, 0)
```

**错误3:VLOOKUP第4参数误用**
```
错误:=VLOOKUP(A2, B:D, 2, TRUE)  (近似匹配可能出错)
正确:=VLOOKUP(A2, B:D, 2, FALSE)  (精确匹配)
```

**错误4:忘记错误处理**
```
错误:=VLOOKUP(A2, B:D, 2, 0)  (未找到显示#N/A)
正确:=IFERROR(VLOOKUP(A2,B:D,2,0), "未找到")
```

## 总结

Excel函数是提升效率的核心工具。掌握本章的30个常用函数,可以解决职场80%的数据处理需求。

**学习要点:**
1. 理解函数参数的含义
2. 掌握函数嵌套的技巧
3. 学会错误排查方法
4. 多练习实战案例
5. 建立个人函数库

**下一章预告:**
《21-查找引用函数深度应用》将深入讲解VLOOKUP、INDEX+MATCH等查找函数的高级技巧。

---

**函数掌握检查清单:**
- [ ] 掌握SUM系列(SUM、SUMIF、SUMIFS)
- [ ] 掌握COUNT系列
- [ ] 掌握IF逻辑判断
- [ ] 掌握文本函数(LEFT、MID、RIGHT、LEN)
- [ ] 掌握日期函数基础
- [ ] 掌握VLOOKUP查找
- [ ] 能进行函数嵌套
- [ ] 会使用IFERROR处理错误

