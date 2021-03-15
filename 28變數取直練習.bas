Attribute VB_Name = "Module1"
Sub 跑计絤策1()
Range("E1").Value = Range("A1").Value + Range("C1").Value
Range("E1").Value = Range("A2").Value - Range("C2").Value
Range("E1").Value = Range("A3").Value * Range("C3").Value
Range("E1").Value = Range("A3").Value / Range("C3").Value
End Sub
Sub 跑计絤策2()
Cells(1, "E").Value = Cells(1, 1).Value + Cells(1, 3).Value
Cells(2, "E").Value = Cells(1, 1).Value - Cells(1, 3).Value
Cells(3, "E").Value = Cells(1, 1).Value * Cells(1, 3).Value
Cells(4, "E").Value = Cells(1, 1).Value / Cells(1, 3).Value
End Sub

Sub 跑计絤策3()
Cells(1, 1).Value = Cells(1, "E").Value + Cells(1, "C").Value
Cells(2, 1).Value = Cells(1, "E").Value - Cells(1, "C").Value
Cells(3, 1).Value = Cells(1, "E").Value * Cells(1, "C").Value
Cells(4, 1).Value = Cells(1, "E").Value / Cells(1, "C").Value
End Sub
