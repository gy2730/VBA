Attribute VB_Name = "Module1"
Sub ���W�Ƨ�()
Attribute ���W�Ƨ�.VB_Description = "�f�n�ƶq���W"
Attribute ���W�Ƨ�.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' ���W�Ƨ� ����
' �f�n�ƶq���W
'
' �ֳt��: Ctrl+d
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("B1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub ����Ƨ�()
Attribute ����Ƨ�.VB_Description = "�f�n�ƶq����Ƨ�"
Attribute ����Ƨ�.VB_ProcData.VB_Invoke_Func = "k\n14"
'
' ����Ƨ� ����
' �f�n�ƶq����Ƨ�
'
' �ֳt��: Ctrl+k
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("B1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub �f�n�ƶq�`�M()
Attribute �f�n�ƶq�`�M.VB_Description = "�f�n�ƶq�`�M"
Attribute �f�n�ƶq�`�M.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' �f�n�ƶq�`�M ����
' �f�n�ƶq�`�M
'
' �ֳt��: Ctrl+p
'
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R2C[-3]:R1048576C[-3])"
    Range("G1").Select
End Sub
Sub �f�n�ƶq����()
Attribute �f�n�ƶq����.VB_Description = "�f�n�ƶq����"
Attribute �f�n�ƶq����.VB_ProcData.VB_Invoke_Func = "l\n14"
'
' �f�n�ƶq���� ����
' �f�n�ƶq����
'
' �ֳt��: Ctrl+l
'
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("H5").Select
End Sub
