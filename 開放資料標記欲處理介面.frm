VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 開放資料標記欲處理介面 
   Caption         =   "開放資料標記欲處理介面"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "開放資料標記欲處理介面.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "開放資料標記欲處理介面"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim arr() '宣告陣列
Dim arrIdx As Integer '宣告陣列索引
Dim tagetValueUB As Integer
Dim tagetValueLB As Integer
Dim tagetValue As Integer
tagetValue = CInt(TET03.Value)
tagetValueUB = CInt(TET01.Value)
tagetValueLB = CInt(TET02.Value)
Columns("h:h").Interior.Color = xlNone 'h欄位全部還原成無顏色
arr = Range(Cells(2, "H"), Cells(2, "H").End(xlDown))
'H2欄位最後一列=Cells(2, "H").End(xlDown)



For arrIdx = 1 To UBound(arr, 1) 'UBound = 陣列上限
  If arr(arrIdx, 1) > tagetValueUB Then '陣列第i元素值>目標值
   Cells(arrIdx + 1, "H").Interior.Color = vbCyan '天藍色
   End If
   If arr(arrIdx, 1) < tagetValueLB Then '小於
   Cells(arrIdx + 1, "H").Interior.Color = vbRed '紅色
   End If
   If arr(arrIdx, 1) = tagetValue Then '等於
   Cells(arrIdx + 1, "H").Interior.Color = vbYellow  '黃色
   End If
Next
Range("a1").CurrentRegion.Borders.LineStyle = xlContinuous
End Sub
