#### Excel 常用 VBA 代码

- 合并相同内容单元格

```
Option Explicit

Sub 合并相同的单元格()
    '变量声明
    Dim strDep  As String
    Dim RowN    As Long
    Dim Rng     As Range
    '获取第一行的部门信息
    strDep = Cells(2, 1).Value
    Set Rng = Cells(2, 1)
    '先关闭警告提示
    Application.DisplayAlerts = False
    '循环遍历A列，从第2行至数据最后的下一行
    For RowN = 2 To 26
        If strDep = Cells(RowN, 1).Value Then
            '内容相同，获取合并的单元格区域对象
            Set Rng = Union(Rng, Cells(RowN, 1))
        Else
            '内容不同，先将相同内容区域进行合并单元格
            Rng.Merge
            '重新获取下一个内容信息
            strDep = Cells(RowN, 1).Value
            Set Rng = Cells(RowN, 1)
        End If
    Next RowN
    '再开启警告提示
    Application.DisplayAlerts = True
End Sub
```