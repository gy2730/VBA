Attribute VB_Name = "Module1"
Sub �۰ʨ���()
Dim deviceName As String
deviceName = Cells(2, 1).Value
MsgBox ("�п�J�˸m�W��:" & deviceName)

Dim modName As String
modName = Cells(2, 2).Value
MsgBox ("�п�J�ҫ��W��:" & modName)

Dim uPrice As Integer
uPrice = Cells(2, 3).Value
MsgBox ("���:" & uPrice)

Dim qty As Integer
qty = Cells(2, 4).Value
MsgBox ("�ƶq:" & qty)

Dim total As Integer
total = uPrice * qty
Cells(2, 5) = total
End Sub
