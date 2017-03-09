#### Excel 常用 VBA 代码

- 合并相同内容单元格

```visual basic
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

- 条件查询并突出显示

  ```visual basic
  Option Explicit

  Sub VBA格式查找()
      '变量声明
      Dim RowN    As Long
      Dim dDate   As Date
      '变量初始化、设置当前日期为2013年5月1日
      dDate = DateSerial(2013, 5, 1)
      '循环
      For RowN = 2 To 36
          '多条件判断
          If Cells(RowN, "B").Value < dDate And Cells(RowN, "C") = "未付款" Then
              '超期未还款，设置填充色
              Columns("A:C").Rows(RowN).Interior.Color = RGB(230, 184, 183)
          Else
              '未超期或已付款，设置无填充
              Columns("A:C").Rows(RowN).Interior.Pattern = xlNone
          End If
      Next
  End Sub
  ```

- 查找最后一个数据（3种方法）

  ```visual basic
  Option Explicit
  'Find方法
  Sub 获取最后数据的行数1()
      '变量声明
      Dim RowN        As Long
      Dim Rng         As Range
      '利用Find方法查找
      Set Rng = Range("C:C").Find("*", Range("C1"), SearchDirection:=xlPrevious)
      RowN = Rng.Row
      '结果输出
      Debug.Print "最后一个单元格为:"; Rng.Address; "行号为:"; RowN
      '以下方法也可以找到
      Set Rng = Range("A:C").Find("*", Range("A1"), SearchOrder:=xlByColumns, _
                  SearchDirection:=xlPrevious)
      RowN = Rng.Row
      '结果输出
      Debug.Print "最后一个单元格为:"; Rng.Address; "行号为:"; RowN
  End Sub
      
  'For循环方法
  Sub 获取最后数据的行数2()
      '变量声明
      Dim rowN        As Long
      '利用For循环从最后一行开始遍历
      For rowN = Rows.Count To 1 Step -1
          If Cells(rowN, "C").Value <> "" Then
              Debug.Print "最后一个单元格为:C" & rowN; "行号为:"; rowN
              Exit For
          End If
      Next rowN
  End Sub
    
    'End方法 （加密时候不适用）
  Sub 获取最后数据的行数3()
      '变量声明
      Dim RowN        As Long
      Dim Rng         As Range
      '利用End属性获取最后一个非空单元格
      If Cells(Rows.Count, "C").Value = "" Then
          Set Rng = Cells(Rows.Count, "C").End(xlUp)
      Else
          Set Rng = Cells(Rows.Count, "C")
      End If
      RowN = Rng.Row
      '结果输出
      Debug.Print "最后一个单元格为:"; Rng.Address; "行号为:"; RowN
  End Sub
  ```

- 删除空白行

  ```vbscript
  Sub delBlankRow()
      '变量声明
      Dim RowN        As Long
      '从最后一行开始循环遍历
      For RowN = Cells(Rows.Count, "A").End(xlUp).Row To 2 Step -1
          '判断是否为空行
          If WorksheetFunction.CountA(Intersect(Rows(RowN), Columns("A:C"))) = 0 Then
              '若为空行则删除
              Rows(RowN).Delete shift:=xlShiftUp
          End If
      Next RowN
  End Sub
  ```

  ​