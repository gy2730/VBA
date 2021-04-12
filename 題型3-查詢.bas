Attribute VB_Name = "Module1"
Sub ¬d¸ß()
Dim Man As String
Dim phone As String
Dim rowNum As Integer
Dim pay As Boolean

Man = Range("G1").Value
For rowNum = 2 To 7
If (Range("G1").Value = Cells(rowNum, "A").Value) Then
   Range("G2").Value = Cells(rowNum, "B").Value
   If (Cells(rowNum, "C").Value = "Y") Then
   pay = True
   phone = Man & ©µ¿ð¥æ³f & pay
   MsgBox (phone)
   Else
   pay = False
   End If
   Else
   End If
Next
End Sub
