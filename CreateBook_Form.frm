VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateBook_Form 
   Caption         =   "Create New Book"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10200
   OleObjectBlob   =   "CreateBook_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateBook_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CheckBox9_Change()

If CheckBox9.Value = True Then
    CheckBox10.Enabled = False
    CheckBox11.Enabled = False
Else
    CheckBox10.Enabled = True
    CheckBox11.Enabled = True
End If

End Sub

Private Sub CheckBox10_Change()

If CheckBox10.Value = True Then
    CheckBox9.Enabled = False
    CheckBox11.Enabled = False
Else
    CheckBox9.Enabled = True
    CheckBox11.Enabled = True
End If

End Sub

Private Sub CheckBox11_Change()

If CheckBox11.Value = True Then
    CheckBox9.Enabled = False
    CheckBox10.Enabled = False
Else
    CheckBox9.Enabled = True
    CheckBox10.Enabled = True
End If

End Sub

Private Sub CheckBox2_Change()

If CheckBox2.Value = True Then
    With TextBox2
        .Enabled = False
    End With
    With TextBox3
        .Enabled = False
    End With
End If

If CheckBox2.Value = False Then
    With TextBox2
        .Enabled = True
    End With
    With TextBox3
        .Enabled = True
    End With
End If

End Sub

Private Sub ComboBox2_Change()

If ComboBox2.Value <> "Select" Then
    CheckBox8.Enabled = False
Else
    CheckBox8.Enabled = True
End If
    
End Sub

Private Sub CheckBox8_Change()
    
If CheckBox8.Value = True Then
    ComboBox2.Enabled = False
Else
    ComboBox2.Enabled = True
End If
    
End Sub

Private Sub ComboBox3_Change()

If ComboBox3.Value <> "Select" Then
    TextBox8.Value = ""
    TextBox8.Enabled = False
Else
    TextBox8.Enabled = True
End If

End Sub

Private Sub CheckBox5_Change()
    
If CheckBox5.Value = True Then
    CheckBox6.Enabled = False
Else
    CheckBox6.Enabled = True
End If
    
End Sub

Private Sub CheckBox6_Change()
    
If CheckBox6.Value = True Then
    CheckBox5.Enabled = False
Else
    CheckBox5.Enabled = True
End If
    
End Sub

Private Sub CommandButton1_Click()
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

Private Sub CommandButton2_Click()
Dim bookName As String
Dim codeList As String
Dim investorType As Long
Dim lien As String
Dim dateType As Long
Dim boardingStart As String
Dim boardingEnd As String
Dim delinq As String
Dim excludePages As String
Dim excludePageNums As String
Dim clientFolder As String
Dim MonthCheck1 As Long, MonthCheck2 As Long
Dim ws As Worksheet, wsPages As Worksheet, wsCodes As Worksheet
Dim j As Long, i As Long, k As Long
Dim answer As Long
Dim tmpBool As Boolean, boolCheck As Boolean
Dim lastrow As Long

Set ws = ThisWorkbook.Sheets("Boarding_Months")
Set wsPages = ThisWorkbook.Sheets("Pages_Key")
Set wsCodes = ThisWorkbook.Sheets("Investor_Codes")

With wsPages
    lastrow = .Cells(wsPages.Rows.Count, 1).End(xlUp).row
End With

If TextBox7.Value = "" Then
    MsgBox "A Performance Book name is required to proceed"
    Exit Sub
Else
    bookName = TextBox7.Value
End If

If Len(bookName) > 40 Then
    MsgBox "The Book Name is too long. Please shorten to ensure total length including spaces is 40 characters or less."
    Exit Sub
Else
End If


If (ComboBox2.Value = "Select" And CheckBox8.Value = False) Or (ComboBox2.Value <> "Select" And CheckBox8.Value = True) Then
    MsgBox "You must either select an existing Investor Code list OR check " _
        & "the box for " & "'All Codes.'" & vbNewLine & vbNewLine & _
        "If you wish to add a new Investor Code list, please select that option"
    Exit Sub
    
Else
    If ComboBox2.Value = "Select" And CheckBox8.Value = True Then
        codeList = "All Investor Codes"
        investorType = 3
    ElseIf ComboBox2.Value <> "Select" And CheckBox8.Value = False Then
        codeList = ComboBox2.Value
        k = 2
        Do Until wsCodes.Range(codeList)(k) = "" Or boolCheck = True
            If InStr(1, wsCodes.Range(codeList)(k), "_") Then
            boolCheck = True
            Else
            k = k + 1
            End If
        Loop
        
        If boolCheck Then
            investorType = 2
        Else
            investorType = 1
        End If
    Else
        MsgBox "Unknown Error"
        Stop
    End If
End If

If CheckBox9.Value = "" And CheckBox10.Value = "" And CheckBox11.Value = "" Then
    MsgBox "A lien position breakout must be selected to proceed."
    Exit Sub
Else
    If CheckBox9.Value = True Then
        lien = 1
    ElseIf CheckBox10.Value = True Then
        lien = 2
    ElseIf CheckBox11.Value = True Then
        lien = "All Liens"
    Else
        MsgBox "Unknown Error"
        Stop
    End If
End If

If CheckBox2.Value = False And (TextBox2.Value = "" Or TextBox3.Value = "") Then
    MsgBox "A range of boarding dates must be selected to proceed." & vbNewLine & vbNewLine & _
        "*note - If not selecting" & "'All Dates'" & " then be sure to include both a starting and ending date."
        Exit Sub
    ElseIf CheckBox2.Value = True And TextBox2.Value = "" And TextBox3.Value = "" Then
        dateType = 3
        boardingStart = "All Dates"
        boardingEnd = "All Dates"
    ElseIf CheckBox2.Value = False And TextBox2.Value <> "" And TextBox3.Value <> "" Then
        On Error GoTo MonthError
        MonthCheck1 = WorksheetFunction.Match(CLng(CDate(TextBox2.Value)), ws.Range("A1:A445"), 0)
        MonthCheck2 = WorksheetFunction.Match(CLng(CDate(TextBox3.Value)), ws.Range("A1:A445"), 0)
        On Error GoTo 0
        
        If CDate(TextBox3.Value) < CDate(TextBox2.Value) Then
            MsgBox "The date range is in the reverse order"
            Exit Sub
        Else
        End If
        
        dateType = 1
        boardingStart = TextBox2.Value
        boardingEnd = TextBox3.Value
        
    ElseIf CheckBox2.Value = True And (TextBox2.Value <> "" Or TextBox3.Value <> "") Then
        MsgBox "Please choose between either " & "'All Dates'" & " or a date range."
        Exit Sub
    
    Else
        MsgBox "You have a problem with the boarding dates"
        Stop
End If

i = 0
tmpBool = False
Do Until tmpBool = True Or i > ListBox1.ListCount - 1
    If ListBox1.Selected(i) = True Then
        tmpBool = True
    Else
        i = i + 1
    End If
Loop

If tmpBool Then
    excludePages = ListBox1.List(i)
    excludePageNums = WorksheetFunction.Index(wsPages.Range("A2:A" & lastrow), WorksheetFunction.Match(ListBox1.List(i), wsPages.Range("E2:E" & lastrow), 0))
    For i = i + 1 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            excludePages = excludePages & ", " & ListBox1.List(i)
            excludePageNums = excludePageNums & "," & WorksheetFunction.Index(wsPages.Range("A2:A" & lastrow), WorksheetFunction.Match(ListBox1.List(i), wsPages.Range("E2:E" & lastrow), 0))
        Else
        End If
    Next i
ElseIf i = ListBox1.ListCount Then
    answer = MsgBox("Be aware that the book will include ALL pages, with exception of pages that do not meet the minimum loan count requirement." & _
                vbNewLine & vbNewLine & "If you wish to proceed please select" & "'Yes'" & "below.", vbYesNo, "Exclude Pages?")
    If answer = 7 Then
        Exit Sub
    ElseIf answer = 6 Then
    excludePages = "None"
    excludePageNums = ""
    End If
Else
    MsgBox "You have a problem with the page selection"
    Stop
End If

If ComboBox3.Value = "Select" And TextBox8.Value = "" Then
    MsgBox "Please choose the name of a client folder for storing the final PDF for client relations"
    Exit Sub
ElseIf ComboBox3.Value <> "Select" And TextBox8.Value = "" Then
    clientFolder = ComboBox3.Value
ElseIf ComboBox3.Value = "Select" And TextBox8.Value <> "" Then
    clientFolder = TextBox8.Value
ElseIf ComboBox3.Value <> "Select" And TextBox8.Value <> "" Then
    MsgBox "Please choose between either an existing client folder OR enter a new folder name."
    Exit Sub
Else
    MsgBox "There is a problem with the client folder selection"
    Stop
End If

If CheckBox5.Value = False And CheckBox6.Value = False Then
    MsgBox "Please choose either a MBA or OTS version book in order to proceed."
    Exit Sub
ElseIf CheckBox5.Value = True And CheckBox6.Value = False Then
    delinq = "MBA"
ElseIf CheckBox6.Value = True And CheckBox5.Value = False Then
    delinq = "OTS"
Else
    MsgBox "There is a problem with the delinquency version selection"
    Stop
End If

   
Call NewBook_Create(bookName, codeList, investorType, lien, dateType, boardingStart, boardingEnd, _
                    excludePages, excludePageNums, clientFolder, delinq)
    
Unload Me
            
Exit Sub

MonthError:
    MsgBox "One of the two months entered is not a valid 'End of Month' value"
    Exit Sub

End Sub

Private Sub CommandButton3_Click()

Unload Me

End Sub

Private Sub CommandButton4_Click()
Dim ctrl As Control
Dim a As String

For Each ctrl In Controls
    a = ctrl.Name
    Select Case TypeName(ctrl)
            Case "TextBox"
                ctrl.Text = ""
            Case "CheckBox", "OptionButton", "ToggleButton"
                ctrl.Value = False
            Case "ComboBox"
                ctrl.ListIndex = -1
            Case "ListBox"
                ctrl.MultiSelect = fmMultiSelectSingle
                ctrl.ListIndex = -1
                ctrl.MultiSelect = fmMultiSelectMulti
    End Select
Next ctrl

ComboBox2.Value = "Select"
ComboBox3.Value = "Select"

End Sub



Private Sub TextBox2_Change()

If Len(TextBox2.Value) > 0 Then
    CheckBox2.Enabled = False
    TextBox3.Enabled = True
  
Else
    CheckBox2.Enabled = True
    TextBox3.Enabled = True
End If

End Sub

Private Sub TextBox3_Change()

If Len(TextBox3.Value) > 0 Then
    CheckBox2.Enabled = False
    TextBox2.Enabled = True
  
Else
    CheckBox2.Enabled = True
    TextBox2.Enabled = True
End If

End Sub

Private Sub TextBox8_Change()

If Len(TextBox8.Value) > 0 Then
    ComboBox3.Enabled = False
  
Else
    ComboBox3.Enabled = True
End If

End Sub

Private Sub UserForm_Click()

End Sub
