Attribute VB_Name = "Module1"
Option Explicit

Sub 產能標計實習()

 Dim i, rowCnt As Integer
 Dim tagetValueUB As Integer
 Dim tagetValueB As Integer
 tagetValueUB = CInt(InputBox("請輸入標計上限值(0-1000)"))
 tagetValueB = CInt(InputBox("請輸入標計下限值(0-1000)"))
 
 Dim rangeStr As String
 rowCnt = Cells(Rows.Count, 1).End(xlUp).Row
 rangeStr = "b3:b" & rowCnt
 MsgBox "目前運算範圍" & rangeStr
 Range(rangeStr).Interior.Color = xlNone
  For i = 3 To rowCnt
     If Cells(i, "B") > tagetValueUB Then
        Cells(i, "B").Interior.Color = vbYellow
        End If
    If Cells(i, "B") < tagetValueB Then
        Cells(i, "B").Interior.Color = vbBlue
        End If
    Next
    Range("a1").CurrentRegion.Borders.LineStyle = xlContinuous


End Sub
