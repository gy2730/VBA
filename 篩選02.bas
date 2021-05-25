Attribute VB_Name = "Module1"
Sub 篩選()
Dim targetCIdex As Integer
    Dim targetValue As String
   targetCIdex = CInt(InputBox("請輸入要篩選的欄位"))
   targetValue = InputBox("目標篩選值")
     Dim username As String
   username = InputBox("請輸入儲存檔案名稱") '檔案名稱
   '第二階段的code
    If ActiveSheet.AutoFilterMode = True Then  '如果是狀態為已篩選
       ActiveSheet.AutoFilterMode = False '取消篩選
    End If
    
   Range("A1").Select '選取範圍
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$I$2500").AutoFilter Field:=targetCIdex, Criteria1:=targetValue
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range("A1:I2497").Select
    Selection.Copy
    Workbooks.Add
  
    Application.CutCopyMode = False
    Sheets("工作表2").Select
    Sheets("工作表2").Name = username
    ChDir "C:\Users\user\Desktop"
    
    ActiveWorkbook.SaveAs Filename:="C:\Users\user\Desktop\" & username & ".xlsm", _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False '儲存至C槽桌面
    ActiveWorkbook.Close
End Sub
