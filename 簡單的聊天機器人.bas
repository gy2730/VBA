Attribute VB_Name = "Module1"
Sub Workbook_Open()
Dim userString As String '宣告變數-宣告一個文字型態變數,名稱叫userString'
userString = InputBox("你晚餐要吃什麼?")
MsgBox "我要吃" & userString
userString01 = InputBox("不好吧我覺得你太胖了換別的!")
MsgBox "那……吃" & userString01
userString02 = InputBox("不喜歡吃!換一個吧?")
MsgBox "好吧吃" & userString02
End Sub
