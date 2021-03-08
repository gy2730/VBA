Attribute VB_Name = "Module1"
Sub 酒精遞增總和最大值最小值()
Attribute 酒精遞增總和最大值最小值.VB_Description = "酒精數量遞增排序總和最大值最小值"
Attribute 酒精遞增總和最大值最小值.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' 酒精遞增總和最大值最小值 巨集
' 酒精數量遞增排序總和最大值最小值
'
' 快速鍵: Ctrl+r
'
    Columns("B:B").Select
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 73
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 100
    ActiveWindow.ScrollRow = 103
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 109
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 117
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 123
    ActiveWindow.ScrollRow = 127
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 156
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 170
    ActiveWindow.ScrollRow = 174
    ActiveWindow.ScrollRow = 178
    ActiveWindow.ScrollRow = 183
    ActiveWindow.ScrollRow = 188
    ActiveWindow.ScrollRow = 191
    ActiveWindow.ScrollRow = 196
    ActiveWindow.ScrollRow = 201
    ActiveWindow.ScrollRow = 205
    ActiveWindow.ScrollRow = 210
    ActiveWindow.ScrollRow = 233
    ActiveWindow.ScrollRow = 236
    ActiveWindow.ScrollRow = 238
    ActiveWindow.ScrollRow = 240
    ActiveWindow.ScrollRow = 241
    ActiveWindow.ScrollRow = 243
    ActiveWindow.ScrollRow = 241
    ActiveWindow.ScrollRow = 239
    ActiveWindow.ScrollRow = 237
    ActiveWindow.ScrollRow = 235
    ActiveWindow.ScrollRow = 233
    ActiveWindow.ScrollRow = 229
    ActiveWindow.ScrollRow = 226
    ActiveWindow.ScrollRow = 220
    ActiveWindow.ScrollRow = 215
    ActiveWindow.ScrollRow = 208
    ActiveWindow.ScrollRow = 201
    ActiveWindow.ScrollRow = 194
    ActiveWindow.ScrollRow = 187
    ActiveWindow.ScrollRow = 180
    ActiveWindow.ScrollRow = 174
    ActiveWindow.ScrollRow = 167
    ActiveWindow.ScrollRow = 159
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 147
    ActiveWindow.ScrollRow = 140
    ActiveWindow.ScrollRow = 136
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 127
    ActiveWindow.ScrollRow = 125
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 119
    ActiveWindow.ScrollRow = 116
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 113
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 109
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 105
    ActiveWindow.ScrollRow = 104
    ActiveWindow.ScrollRow = 102
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 100
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B553"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("B1:B553")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=-6
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=MAX(C[-5])"
    Columns("I:I").Select
    ActiveCell.FormulaR1C1 = "=MIN(C[-7])"
    Range("F41").Select
    ActiveWindow.SmallScroll Down:=-30
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R2C[-5]:R1048576C[-5])"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C[-3]:R1048576C[-3])"
    Range("G4").Select
End Sub
Sub 酒精遞減總和最大值最小值()
Attribute 酒精遞減總和最大值最小值.VB_Description = "遞減+最大值+最小值"
Attribute 酒精遞減總和最大值最小值.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' 酒精遞減總和最大值最小值 巨集
' 遞減+最大值+最小值
'
' 快速鍵: Ctrl+t
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B553")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B553"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("B1:B553")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C[-3]:R1048576C[-3])"
    Range("I1").Select
End Sub
