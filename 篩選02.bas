Attribute VB_Name = "Module1"
Sub �z��()
Dim targetCIdex As Integer
    Dim targetValue As String
   targetCIdex = CInt(InputBox("�п�J�n�z�諸���"))
   targetValue = InputBox("�ؼпz���")
     Dim username As String
   username = InputBox("�п�J�x�s�ɮצW��") '�ɮצW��
   '�ĤG���q��code
    If ActiveSheet.AutoFilterMode = True Then  '�p�G�O���A���w�z��
       ActiveSheet.AutoFilterMode = False '�����z��
    End If
    
   Range("A1").Select '����d��
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$I$2500").AutoFilter Field:=targetCIdex, Criteria1:=targetValue
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range("A1:I2497").Select
    Selection.Copy
    Workbooks.Add
  
    Application.CutCopyMode = False
    Sheets("�u�@��2").Select
    Sheets("�u�@��2").Name = username
    ChDir "C:\Users\user\Desktop"
    
    ActiveWorkbook.SaveAs Filename:="C:\Users\user\Desktop\" & username & ".xlsm", _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False '�x�s��C�Ѯୱ
    ActiveWorkbook.Close
End Sub
