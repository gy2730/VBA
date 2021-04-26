Attribute VB_Name = "Module2"
Option Explicit

Private Sub btninsert_Click()
Dim supplyName As String
supplyName = txtname1.Text
Cells(2, 1).Value = supplyName

Dim supplyPhone As String
supplyPhone = txtname2.Text
Cells(2, 2) = supplyPhone

Dim price As Integer
price = txtname3.Text
Cells(2, 3) = CInt(price)

Dim nprice As Integer
nprice = txtname4.Text
Cells(2, 4) = CInt(nprice)

Dim totalDiscount As Single
totalDiscount = (price - nprice) / price
Cells(2, 5) = totalDiscount

If (totalDiscount > 0.8) Then
Cells(2, 6).Value = "²§±`"
Else
Cells(2, 6).Value = "¥¿±`"

End If
End Sub

