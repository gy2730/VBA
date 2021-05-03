VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "B10721128_朱玉絢_實習1.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button01_Click()

Dim week As String
Dim rowNum As Integer
Dim week01 As Integer
Dim week02 As Integer
Dim week03 As Integer
Dim rangeStr As String
  
   For rowNum = 2 To 10
  week = Tetinsert.Text
  
  If (Cells(rowNum, "A").Value = week) Then
  
     If (Cells(rowNum, "A").Value = "第一週") Then
     week01 = Cells(2, "B").Value + Cells(3, "B").Value + Cells(4, "B").Value
     Cells(2, "F").Value = week01
     
     End If
  If (Cells(rowNum, "A").Value = "第二週") Then
     week02 = Cells(5, "B").Value + Cells(6, "B").Value + Cells(7, "B").Value
    Cells(2, "F").Value = week02
     End If
     
  If (Cells(rowNum, "A").Value = "第三週") Then
     week03 = Cells(8, "B").Value + Cells(9, "B").Value + Cells(10, "B").Value
     Cells(2, "F").Value = week03
     
     End If

  End If
Next
    MsgBox "總數" & Cells(2, "F").Value
End Sub
