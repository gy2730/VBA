Attribute VB_Name = "Module1"
Sub 自動取值()
Dim deviceName As String
deviceName = Cells(2, 1).Value
MsgBox ("請輸入裝置名稱:" & deviceName)

Dim modName As String
modName = Cells(2, 2).Value
MsgBox ("請輸入模型名稱:" & modName)

Dim uPrice As Integer
uPrice = Cells(2, 3).Value
MsgBox ("單價:" & uPrice)

Dim qty As Integer
qty = Cells(2, 4).Value
MsgBox ("數量:" & qty)

Dim total As Integer
total = uPrice * qty
Cells(2, 5) = total
End Sub
