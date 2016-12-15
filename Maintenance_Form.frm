VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Maintenance_Form 
   Caption         =   "Maintenance Form"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3945
   OleObjectBlob   =   "Maintenance_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Maintenance_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Dim lastrow2 As Long, lastrow3 As Long, lastrow As Long
Dim ws As Worksheet, wsPages As Worksheet, wsMain As Worksheet

Set ws = ThisWorkbook.Sheets("Investor_Codes")
Set wsPages = ThisWorkbook.Sheets("Pages_Key")
Set wsMain = ThisWorkbook.Sheets("Sheet1")

With ws
    lastrow2 = .Cells(ws.Rows.Count, 3).End(xlUp).row
End With

With wsPages
    lastrow3 = .Cells(wsPages.Rows.Count, 1).End(xlUp).row
End With

With wsMain
    lastrow = .Cells(wsPages.Rows.Count, 19).End(xlUp).row
End With

With CreateBook_Form
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .ComboBox2.RowSource = "Investor_Codes!C2:C" & lastrow2
  .ListBox1.RowSource = "Pages_Key!E2:E" & lastrow3
  .ComboBox3.RowSource = "Sheet1!S1:S" & lastrow
  .ComboBox2.Value = "Select"
  .ComboBox3.Value = "Select"
  .Show
End With

End Sub


Private Sub CommandButton2_Click()
Dim lastrow2 As Long
Dim ws As Worksheet

Set ws = ThisWorkbook.Sheets("Standard_Books")

With ws
lastrow2 = .Cells(ws.Rows.Count, 1).End(xlUp).row
End With

With DeleteBook_Form
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .ListBox1.RowSource = "Standard_Books!A3:A" & lastrow2
  .Show
End With

End Sub

Private Sub CommandButton3_Click()
Dim lastrow2 As Long
Dim ws As Worksheet
Dim wb As Workbook

Set wb = ThisWorkbook
Set ws = wb.Sheets("Investor_Codes")


'wb.Connections(1).Refresh
ws.ListObjects("Table_sqlprd134").Refresh

With ws
lastrow2 = .Cells(ws.Rows.Count, 1).End(xlUp).row
End With


With CreateCodes_Form
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .ListBox1.RowSource = "Investor_Codes!A2:A" & lastrow2
  .Show
End With

End Sub

Private Sub CommandButton4_Click()
Dim lastrow2 As Long
Dim ws As Worksheet

Set ws = ThisWorkbook.Sheets("Investor_Codes")

With ws
lastrow2 = .Cells(ws.Rows.Count, 3).End(xlUp).row
End With

With DeleteCodes_Form
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .ListBox1.RowSource = "Investor_Codes!C3:C" & lastrow2
  .Show
End With

End Sub

Private Sub CommandButton5_Click()

Unload Me

End Sub
