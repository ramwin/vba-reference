vba-reference
学习vba的笔记

[vba官方文档](https://docs.microsoft.com/zh-cn/office/vba/api/overview/excel)
[vba编程教程 w3cschool](https://www.w3cschool.cn/excelvba/?)
[网易课程](https://study.163.com/course/courseLearn.htm?courseId=1003088001)

# VBA语言基础
## 变量
* 基础
```
x = Cells(5, 3)  // x 等于 C5的值
Set x = Cells(5, 3)  // x 等于 C5这个单元格对象
```
* [单元格读取赋值](./01单元格读取赋值.xlsm)
```
Sub 做加法()
  Cells(7, 9) = Cells(7, 5) + Cells(7, 7)
End Sub
```
* [使用变量](./02使用变量.xlsm)
    * 没有定义的变量默认是0
```
Sub 做加法()
  x = Cells(3, 2)
  Cells(x, 6) = "+"
  Cells(x, 9) = Cells(x, 5) + Cells(x, 7)
End Sub
```
* 强制声明
所有的变量
```
Option Explicit  // 强制声明，这样所有的变量必须先定义
Sub Sub()
    Dim r1, s, v
    r1 = Cells(3, 2)
End sub
```
* 常量
```
const PI = 3.14159  // 后续就不能修改PI了
```

## 流程控制
### 循环
* For 循环
```
For i = 11 To 20 Step 1  // Step 1 可以省略
    Cells(i, 6) = Cells(i, 6) / rate
    // 中间可以修改i, 但是不建议
Next i  // i 可以省略。多个循环可以指定那个for的continue
```
* While 循环
```
Dim i
i = 2
While Cells(i, 1) <> ""
    If Cells(i, 2) < 60 Then
        Cells(i, 2).Font.Color = vbRed
    End If
    i = i + 1
WEND
# 或者用DO While
Do While Cells(i, 1) <> ""
    If Cells(i, 2) < 60 Then
        Cells(i, 2).Font.Color = vbRed
    End If
    i = i + 1
Loop
```

### 判断
```
if score >= 60 then
    Cells(8, 6) = "及格"
Elseif score <= 30 then
    Cells(8, 6) = "你退学吧"
Else  // else可以没有
    Cells(8, 6) = "不及格"
End If

if score < 10 Then Cells(8, 6) = "没得救了"
if score <> 100 Then  // 不等于用<>
End if
```

## 调试
* 设置断点
点击代码左侧. 然后单步执行。如果想知道这时候的某个变量是多少，只要把鼠标移动到变量上
* 单独执行
* 添加监视
可以一直看某个变量的值

## 字符串
* 用双括号括起来.
```
s1 = "你好, "
s2 = s1 & "Every One!"  // 拼接，前后要有空格
```

## 逻辑运算
And Or Not

## 单元格
* 修改字体颜色
```
Range("A3:B5").Font.Color = -16776961
```

## 录制宏
通过录制宏，然后查看宏的代码可以查看很多想知道的属性
```
Rows("6:6").Select
Selection.Delete Shift:=x1Up
Range("E7").Select
With Selection.Font
    .Color = -16776961
    .TintAndShare = 0
```

## 注释
使用Rem或者单引号
```
Rem i 代表行号
' j 代表列号
```

## 函数调用
### 过程
```
Sub sumAll()
End Sub

Sub allScore()
    Call sumAll  // Call可以省略
End Sub
```
### 函数
能返回结果的过程。 定义完函数后，在单元格里面可以直接使用`=客户分类(A2)`来得到结果
* 返回数据以及传入参数
```
Function 客户分类(s)
    客户分类 = level  // 最后函数名 = 值，就会把值返回
End Function
Sub 分析()
    score = 10
    level = 客户分类(score)
    Cells(i, 4) = level
End Sub
```

## 参考
### Application
* WorkBooks
### WorkBook
* WorkSheets
    * 获取某个工作表
    ```
    WorkSheets(1)
    WorkSheets("总分榜")
    ```
    * Count
    返回Worksheets的数量
    WorkSheets.count
    * Add
    新建一个新的工作表
    ```
    Set w1 = WorkSheets.Add
    w1.Cells(5, 3) = 100
    ```
    * Name
    返回工作表的名字
### WorkSheet
* Cells
```
Sub test()
    Dim w1 As Worksheet
    Set w1 = Worksheets(3)
    w1.Cells(5, 3) = 100
End Sub
```
### Range
