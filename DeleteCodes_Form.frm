VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteCodes_Form 
   Caption         =   "Delete an Existing List"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8100
   OleObjectBlob   =   "DeleteCodes_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteCodes_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Dim ListName As String
Dim answer As Long
Dim lastrow As Long
Dim ws As Worksheet
Dim wb As Workbook
Dim wsBooks As Worksheet, wsInv As Worksheet
Dim i As Long, j As Long, k As Long
Dim report As String
Dim listStr As String, rgStr As String
Dim OldDeal As String

Set ws = ThisWorkbook.Sheets("Investor_Codes")
Set wsBooks = ThisWorkbook.Sheets("Standard_Books")
Set wsInv = ThisWorkbook.Sheets("Investor_Codes")
Set wb = ThisWorkbook

i = 0
j = 1
k = 2

If ListBox1.Value <> "" Then
    ListName = ListBox1.Value
Else
    MsgBox "Please select a list that you wish to view details."
    Exit Sub
End If

With wsBooks
    lastrow = .Cells(wsBooks.Rows.Count, 2).End(xlUp).row
    
    Do Until i > 0 Or j > lastrow
        If .Cells(j, 2).Value = ListName Then
            report = .Cells(j, 1).Value
            i = i + 1
        Else
            j = j + 1
        End If
    Loop
End With

Do Until wsInv.Range(ListName)(k) = ""
    listStr = listStr & wsInv.Range(ListName)(k) & ","
    k = k + 1
Loop

listStr = Left(listStr, (Len(listStr) - 1))
rgStr = wb.Names(ListName).RefersTo
rgStr = Right(rgStr, (Len(rgStr) - 16))
OldDeal = wsInv.Cells(1, wsInv.Range(rgStr).Column).Value

MsgBox "The list '" & ListName & "' is currently used for performance book '" & report & "'" & vbNewLine & vbNewLine & _
"It is associated with the Deal Level Name '" & OldDeal & "' and contains the following codes:" & vbNewLine & vbNewLine & listStr


End Sub

Private Sub CommandButton2_Click()
Dim ListName As String

ListName = ListBox1.Value

Call Delete_List(ListName)


End Sub

Private Sub CommandButton3_Click()

Unload Me

End Sub
