VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �}���ƼаO���B�z���� 
   Caption         =   "�}���ƼаO���B�z����"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "�}���ƼаO���B�z����.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "�}���ƼаO���B�z����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim arr() '�ŧi�}�C
Dim arrIdx As Integer '�ŧi�}�C����
Dim tagetValueUB As Integer
Dim tagetValueLB As Integer
Dim tagetValue As Integer
tagetValue = CInt(TET03.Value)
tagetValueUB = CInt(TET01.Value)
tagetValueLB = CInt(TET02.Value)
Columns("h:h").Interior.Color = xlNone 'h�������٭즨�L�C��
arr = Range(Cells(2, "H"), Cells(2, "H").End(xlDown))
'H2���̫�@�C=Cells(2, "H").End(xlDown)



For arrIdx = 1 To UBound(arr, 1) 'UBound = �}�C�W��
  If arr(arrIdx, 1) > tagetValueUB Then '�}�C��i������>�ؼЭ�
   Cells(arrIdx + 1, "H").Interior.Color = vbCyan '���Ŧ�
   End If
   If arr(arrIdx, 1) < tagetValueLB Then '�p��
   Cells(arrIdx + 1, "H").Interior.Color = vbRed '����
   End If
   If arr(arrIdx, 1) = tagetValue Then '����
   Cells(arrIdx + 1, "H").Interior.Color = vbYellow  '����
   End If
Next
Range("a1").CurrentRegion.Borders.LineStyle = xlContinuous
End Sub
