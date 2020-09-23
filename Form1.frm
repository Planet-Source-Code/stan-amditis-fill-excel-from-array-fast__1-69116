VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fill Excel From Array"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4080
   Begin VB.CommandButton Command2 
      Caption         =   "Fill Sheet using ""Resize"""
      Height          =   555
      Left            =   960
      TabIndex        =   1
      Top             =   1095
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill Sheet Cell-by-Cell"
      Height          =   630
      Left            =   945
      TabIndex        =   0
      Top             =   360
      Width           =   2190
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Create by Stan Amditis
'
' ExcelFillTest
'
' This program demonstrates the proper way to fill an excel sheet with data from an array
' References: None
' This code has been tested with Excel version 2000,XP, and 2003
' This code is untested for version prior to 2000 and versions after 2003
'
Option Explicit
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Dim aryTestData(0 To 99, 0 To 99) As Long
Dim oExcel As Object 'Excel.Application
Dim bNewExcelObjectCreated As Boolean
Dim oWorkBook1 As Object ' Excel.Workbook
Dim oWorkBook2 As Object ' Excel.Workbook
Private Function OpenExcel(ByRef bObjectCreated As Boolean) As Boolean
  On Error Resume Next
  Err.Clear
  ' If we haven't opened excel yet...
  If oExcel Is Nothing Then
    ' Assume we will use existing object
    bObjectCreated = False
    ' Attemtp to get existing excel application object
    Set oExcel = GetObject(, "Excel.Application")
    ' Existing excel not found, try to create a new one
    If Err.Number <> 0 Then
      Err.Clear
      Set oExcel = CreateObject("excel.application")
      oExcel.Visible = True
      If Err.Number <> 0 Then
        MsgBox "Cannot create excel object", vbCritical
        Exit Function
      End If
      ' We created a new instance of excel
      bObjectCreated = True
    End If
  End If
  ' We succesfully fetched an excel object
  OpenExcel = True
End Function
Private Sub CloseExcel()
  ' If we have no excel object, just return
  If oExcel Is Nothing Then Exit Sub
  ' If we have any workbooks open...
  If Not oWorkBook1 Is Nothing Then
    ' Close and do NOT save
    oWorkBook1.Close False
  End If
  If Not oWorkBook2 Is Nothing Then
    ' Close and do NOT save
    oWorkBook2.Close False
  End If
  ' If we had to create our own excel application object
  If bNewExcelObjectCreated Then
    ' Close it
    oExcel.Quit
  End If
  Set oExcel = Nothing
End Sub
Private Function AddWorkBook(oExcelApplication As Object) As Object
  ' Add a workbook and return object
  Set AddWorkBook = oExcelApplication.Workbooks.Add()
End Function
Private Function GetWorkSheet(oWorkBook As Object, n As Integer) As Object
  ' Get first worksheet and return object
  Set GetWorkSheet = oWorkBook.Sheets(n)
End Function

Private Sub FillSheetCellByCell(oSheet As Object, aryData)
  Dim n As Integer, m As Integer
  ' Loop through array and set values in sheet
  For n = LBound(aryData, 1) To UBound(aryData, 1)
    For m = LBound(aryData, 2) To UBound(aryData, 2)
      oSheet.Cells(n + 1, m + 1) = aryData(n, m)
    Next
  Next
  BringWindowToTop Me.hWnd
  Me.SetFocus
  MsgBox "Fill using Cell-by-cell...Done"
End Sub
Private Sub FillSheetUsingRange(oSheet As Object, aryData)
  ' Use "resize" function to fill cells from array
  ' Important:  Rows and Columns are 1-based in excel
  ' We start with range that includes only the first cell
  ' Using the resize function, we pump all the data into the sheet in one step
  oSheet.Range("A1", "A1").Resize(UBound(aryData, 1) - LBound(aryData, 1) + 1, UBound(aryData, 2) - LBound(aryData, 2) + 1).Value = aryData
  BringWindowToTop Me.hWnd
  Me.SetFocus
  MsgBox "Fill using Resize...Done"
End Sub
Private Sub InitDataArray()
  Dim n As Integer, m As Integer
  
  For n = 0 To 99
    For m = 0 To 99
      aryTestData(n, m) = n * 100 + m
    Next
  Next
  
End Sub

Private Sub Command1_Click()
  Screen.MousePointer = vbHourglass
  
  ' Load Excel Object
  If Not OpenExcel(bNewExcelObjectCreated) Then Exit Sub
  ' Add new workbook
  Set oWorkBook1 = AddWorkBook(oExcel)
  ' Fill first sheet with data
  FillSheetCellByCell GetWorkSheet(oWorkBook1, 1), aryTestData
  
  Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
  Screen.MousePointer = vbHourglass
  
  If Not OpenExcel(bNewExcelObjectCreated) Then Exit Sub
  ' Add new workbook
  Set oWorkBook2 = AddWorkBook(oExcel)
  ' Fill first sheet with data
  FillSheetUsingRange GetWorkSheet(oWorkBook2, 1), aryTestData
  
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  InitDataArray
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  CloseExcel
End Sub
