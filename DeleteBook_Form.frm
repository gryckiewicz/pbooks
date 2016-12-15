VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteBook_Form 
   Caption         =   "Delete Book"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8100
   OleObjectBlob   =   "DeleteBook_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteBook_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim bookName As String
Dim answer As Long
Dim ws As Worksheet, wsPages As Worksheet, wsMain As Worksheet, wsBooks As Worksheet
Dim lastrow As Long
Dim bookRow As Long
Dim pbArray(1 To 11) As String
Dim i As Long
Dim lien As String
Dim boardingStart As String
Dim boardingEnd As String

Set ws = ThisWorkbook.Sheets("Investor_Codes")
Set wsPages = ThisWorkbook.Sheets("Pages_Key")
Set wsMain = ThisWorkbook.Sheets("Sheet1")
Set wsBooks = ThisWorkbook.Sheets("Standard_Books")

If ListBox1.Value <> "" Then
    bookName = ListBox1.Value
Else
    MsgBox "Please select the book you wish to view details."
    Exit Sub
End If

With wsBooks
    bookRow = WorksheetFunction.Match(bookName, wsBooks.Range("A1:A1000"), 0)
    lastrow = .Cells(wsBooks.Rows.Count, 1).End(xlUp).row
End With

For i = LBound(pbArray) To UBound(pbArray)
    pbArray(i) = wsBooks.Range("A1:K" & lastrow).Cells(bookRow, i).Value
Next i

If pbArray(4) = 1 Then
    lien = "1st Liens"
ElseIf pbArray(4) = 2 Then
    lien = "2nd or Greater Liens"
ElseIf pbArray(4) = 3 Then
    lien = "All Liens"
Else
    MsgBox "You have a problem"
    Stop
End If

If pbArray(5) = "1" Then
    boardingStart = "All Dates"
Else
    boardingStart = pbArray(5)
End If

If pbArray(6) = "1" Then
    boardingEnd = "All Dates"
Else
    boardingEnd = pbArray(6)
End If


MsgBox "Here are the details of the book you've selected:" & vbNewLine & vbNewLine & _
"Book Name: " & bookName & vbNewLine & _
"Code List: " & pbArray(2) & vbNewLine & _
"Lien Position: " & lien & vbNewLine & _
"Boarding Dates from: " & boardingStart & " to: " & boardingEnd & vbNewLine & _
"Delinquency Type: " & pbArray(9) & vbNewLine & _
"Excluding Pages: " & pbArray(10) & vbNewLine & vbNewLine & _
"Client Folder: " & pbArray(11)


End Sub

Private Sub CommandButton2_Click()
Dim bookName As String

bookName = ListBox1.Value

Call Delete_Book(bookName)

End Sub

Private Sub CommandButton3_Click()

Unload Me

End Sub
