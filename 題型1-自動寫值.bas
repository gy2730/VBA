Attribute VB_Name = "Module1"
Sub 自動寫值()
Dim MachineName As String
MachineName = InputBox("請輸入設備名稱")
Cells(2, 1).Value = MachineName

Dim modeName As String
modeName = InputBox("請輸入模具名稱")
Cells(2, 2).Value = modeName

Dim uPrice As Integer
uPrice = InputBox("請輸入單價")
Cells(2, 3).Value = CInt(unitPrice)

Dim qty As Integer
qty = InputBox("請輸入數量")
Cells(2, 4).Value = CInt(qty)

Dim total As Integer
total = unitPrice * qty
Cells(2, 5) = total

End Sub

