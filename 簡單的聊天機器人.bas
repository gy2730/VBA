Attribute VB_Name = "Module1"
Sub Workbook_Open()
Dim userString As String '�ŧi�ܼ�-�ŧi�@�Ӥ�r���A�ܼ�,�W�٥suserString'
userString = InputBox("�A���\�n�Y����?")
MsgBox "�ڭn�Y" & userString
userString01 = InputBox("���n�a��ı�o�A�ӭD�F���O��!")
MsgBox "���K�K�Y" & userString01
userString02 = InputBox("�����w�Y!���@�ӧa?")
MsgBox "�n�a�Y" & userString02
End Sub
