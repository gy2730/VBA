Attribute VB_Name = "Module1"
Sub �E�E���k��()
Dim row, col As Integer
 For row = 1 To 9
  For col = 1 To 9
  Cells(row, col).Value = row & "*" & col & "=" & row * col
  Next
Next
End Sub
