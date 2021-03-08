Attribute VB_Name = "Module1"
Sub 遞增排序()
Attribute 遞增排序.VB_Description = "口罩數量遞增"
Attribute 遞增排序.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' 遞增排序 巨集
' 口罩數量遞增
'
' 快速鍵: Ctrl+d
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("B1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 遞減排序()
Attribute 遞減排序.VB_Description = "口罩數量遞減排序"
Attribute 遞減排序.VB_ProcData.VB_Invoke_Func = "k\n14"
'
' 遞減排序 巨集
' 口罩數量遞減排序
'
' 快速鍵: Ctrl+k
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("B1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 口罩數量總和()
Attribute 口罩數量總和.VB_Description = "口罩數量總和"
Attribute 口罩數量總和.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' 口罩數量總和 巨集
' 口罩數量總和
'
' 快速鍵: Ctrl+p
'
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C[-3]:R1048576C[-3])"
    Range("G1").Select
End Sub
Sub 口罩數量平均()
Attribute 口罩數量平均.VB_Description = "口罩數量平均"
Attribute 口罩數量平均.VB_ProcData.VB_Invoke_Func = "l\n14"
'
' 口罩數量平均 巨集
' 口罩數量平均
'
' 快速鍵: Ctrl+l
'
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("H5").Select
End Sub
