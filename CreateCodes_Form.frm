VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateCodes_Form 
   Caption         =   "Create a New Investor Code List"
   ClientHeight    =   11940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8100
   OleObjectBlob   =   "CreateCodes_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateCodes_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

Dim answer As Integer
Dim ListName As String, DealName As String, codeList As String
Dim j As Long, i As Long, k As Long, l As Long
Dim charallowed As String
Dim testQueryString As String
Dim testQueryResult() As String


charallowed = "1234567890acbdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_"

answer = MsgBox("Are you sure you want to create this list?", vbYesNo + vbQuestion, "Create List?")

If answer = vbYes Then
    
    If TextBox1.Text = "" Then
        MsgBox "The 'List Name' field is required", vbOKOnly, "Error"
        Exit Sub
    Else
        ListName = TextBox1.Text
    End If
    
    k = 0
    l = 1
    Do Until k > 0 Or l > Len(ListName)
    
        If Not InStr(1, charallowed, Mid(ListName, l, 1)) > 0 Then
            k = k + 1
        Else
            l = l + 1
        End If
    Loop
    
    If k > 0 Then
        MsgBox "Please enter only valid characters and NO spaces" & vbNewLine & vbNewLine & _
            "The character """ & Mid(ListName, l, 1) & """ is invalid" & vbNewLine & vbNewLine & _
            "Only alpha-numerics and the character ""_"" are allowed"
        Exit Sub
    Else
    End If
    
    j = 0
    i = 0
    Do Until j > 0 Or i > ListBox1.ListCount
        If ListBox1.Selected(i) = True Then
            j = i + 1
        Else
            i = i + 1
        End If
    Loop
    
    If j = 0 And TextBox2.Value <> "" Then
        DealName = "Custom"
        codeList = TextBox2.Value
    ElseIf j > 0 And TextBox2.Value = "" Then
        DealName = ListBox1.Value
        codeList = ""
    ElseIf j = 0 And TextBox2.Value = "" And TextBox3.Value <> "" Then
        testQueryString = TextBox3.Value
        Debug.Print testQueryString
        DealName = TextBox1.Value
        codeList = ""
        Unload Me
        Unload Maintenance_Form
        Unload UserForm1
        testQueryResult = Query_Validation(DealName, testQueryString)
        
        If testQueryResult(1) = "Y" Then
            MsgBox "The query provided does not return any data. Try writing the query first in SQL Server Management Studio and ensuring the result set" _
            & " is only one column of integer values"
            Exit Sub
        ElseIf testQueryResult(2) > 1 Then
            MsgBox "The query provided returns more than one column. Try writing the query first in SQL Server Management Studio and ensuring the result set" _
            & " is only one column of integer values"
            Exit Sub
        ElseIf testQueryResult(3) = "Y" Then
            MsgBox "The query provided returns a column of values that are NOT integers. Try writing the query first in SQL Server Management Studio and ensuring the result set" _
            & " is only one column of integer values"
            Exit Sub
        Else
        End If
    Else
        MsgBox "You must select only one method for creating a code list - either a Deal Level Name or input individual codes or a custom query", vbOKOnly, "Error"
        Exit Sub
    End If
    

    
    Call Create_Code_List(ListName, DealName, codeList)
    Unload Me
    
Else: Exit Sub
End If

End Sub

Private Sub CommandButton2_Click()

Unload Me

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Click()

End Sub
