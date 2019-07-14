vba-reference
学习vba的笔记

[vba官方文档](https://docs.microsoft.com/zh-cn/office/vba/api/overview/excel)
[vba编程教程 w3cschool](https://www.w3cschool.cn/excelvba/?)
[网易课程](https://study.163.com/course/courseLearn.htm?courseId=1003088001)

# VBA语言基础
## 变量
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

## 循环
