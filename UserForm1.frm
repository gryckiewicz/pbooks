VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Performance Book Creator"
   ClientHeight    =   12780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16725
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CheckBox1_Change()
Dim ctrl As Control
Dim i As Integer
    
If CheckBox1.Value = True Then
    With ListBox1
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next i
        .Enabled = False
    End With
Else
    ListBox1.Enabled = True
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

Private Sub CheckBox3_Change()

If CheckBox3.Value = True Then
    ListBox2.Enabled = False
Else
    ListBox2.Enabled = True
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
Private Sub CheckBox8_Change()
    
If CheckBox8.Value = True Then
    TextBox1.Enabled = False
    TextBox5.Enabled = False
    TextBox6.Enabled = False
    ComboBox2.Enabled = False
Else
    TextBox1.Enabled = True
    TextBox5.Enabled = True
    TextBox6.Enabled = True
    ComboBox2.Enabled = True
End If
    
End Sub


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

Private Sub ComboBox1_Change()
Dim ctrl As Control
Dim i As Long


If ComboBox1.Value = "Standard Book(s)" Then
    For Each ctrl In UserForm1.Controls
        If (ctrl.Name <> "CheckBox1" Or ctrl.Name <> "CheckBox12" Or ctrl.Name <> "CheckBox13") And TypeName(ctrl) = "CheckBox" Then
            ctrl.Value = False
            ctrl.Enabled = False
        ElseIf ctrl.Name <> "ListBox1" And TypeName(ctrl) = "ListBox" Then
            For i = 0 To ctrl.ListCount - 1
            ctrl.Selected(i) = False
            Next i
            ctrl.Enabled = False
        ElseIf TypeName(ctrl) = "TextBox" Then
            ctrl.Value = ""
            ctrl.Enabled = False
        ElseIf TypeName(ctrl) = "Frame" And ctrl.Name <> "Frame3" Then
            ctrl.Enabled = False
        ElseIf TypeName(ctrl) = "Label" And ctrl.Name <> "Label1" And _
            ctrl.Name <> "Label20" And ctrl.Name <> "Label2" Then
            ctrl.Enabled = False
        End If
    Next ctrl
    CheckBox1.Enabled = True
    CheckBox12.Enabled = True
    CheckBox13.Enabled = True
    ListBox1.Enabled = True
    CommandButton1.Enabled = True
    Frame5.Enabled = False
    Frame3.Enabled = True
    Label1.Enabled = True
    Label2.Enabled = True
ElseIf ComboBox1.Value = "Custom Book" Then
    For Each ctrl In UserForm1.Controls
        If (ctrl.Name <> "CheckBox1" Or ctrl.Name <> "CheckBox12" Or ctrl.Name <> "CheckBox13") And TypeName(ctrl) = "CheckBox" Then
            ctrl.Enabled = True
        ElseIf ctrl.Name <> "ListBox1" And TypeName(ctrl) = "ListBox" Then
            ctrl.Enabled = True
        ElseIf TypeName(ctrl) = "TextBox" Then
            ctrl.Enabled = True
        ElseIf TypeName(ctrl) = "ComboBox" Then
            ctrl.Enabled = True
        ElseIf TypeName(ctrl) = "Frame" And ctrl.Name <> "Frame3" Then
            ctrl.Enabled = True
        ElseIf TypeName(ctrl) = "Label" And ctrl.Name <> "Label1" And _
            ctrl.Name <> "Label2" Then
            ctrl.Enabled = True
        End If
    Next ctrl
    With CheckBox1
        .Value = False
        .Enabled = False
    End With
    With CheckBox12
        .Value = False
        .Enabled = False
    End With
    With CheckBox13
        .Value = False
        .Enabled = False
    End With
    With ListBox1
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next i
        .Enabled = False
    End With
    CommandButton1.Enabled = True
    Frame5.Enabled = True
    Frame3.Enabled = False
Else
End If

End Sub

Private Sub ComboBox2_Change()

If ComboBox2.Value <> "Select" Then
    TextBox1.Enabled = False
    TextBox5.Enabled = False
    TextBox6.Enabled = False
    CheckBox8.Enabled = False
Else
    TextBox1.Enabled = True
    TextBox5.Enabled = True
    TextBox6.Enabled = True
    CheckBox8.Enabled = True
End If
    
End Sub

Private Sub CommandButton1_Click()

Dim answer As Integer


answer = MsgBox("Are you sure you want to produce these books?", vbYesNo + vbQuestion, "Produce Books?")

If answer = vbYes Then

    Call PBook_Main
    
Else: Exit Sub
End If

'Unload Me

End Sub


Private Sub CommandButton2_Click()

Unload Me

End Sub


Private Sub CommandButton3_Click()

With Process_Manual
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .MultiPage1.Value = 2
  .MultiPage2.Value = 0
  .Show
End With

End Sub


Private Sub CommandButton5_Click()

With Maintenance_Form
  .StartUpPosition = 0
  .Left = Application.Left + (0.75 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .Show
End With


End Sub

Private Sub ListBox1_Change()

Dim i, j As Long
 
For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) = True Then j = j + 1
Next i

If j > 0 Then
    CheckBox1.Enabled = False
Else
    CheckBox1.Enabled = True
End If

End Sub

Private Sub ListBox2_Change()

Dim i As Long
Dim j As Long
    
j = 0
For i = 0 To ListBox2.ListCount - 1
    If ListBox2.Selected(i) = True Then j = j + 1
Next i

If j > 0 Then
    With CheckBox3
        .Value = False
        .Enabled = False
    End With
Else
    With CheckBox3
        .Value = False
        .Enabled = True
    End With
End If

End Sub
Private Sub TextBox1_Change()

If TextBox1.Value = True Then
    TextBox5.Enabled = False
    TextBox6.Enabled = False
    ComboBox2.Enabled = False
    CheckBox8.Enabled = False
Else
    TextBox5.Enabled = True
    TextBox6.Enabled = True
    ComboBox2.Enabled = True
    CheckBox8.Enabled = True
End If
    
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
Private Sub TextBox5_Change()

If TextBox5.Value = True Then
    TextBox1.Enabled = False
    ComboBox2.Enabled = False
    CheckBox8.Enabled = False
Else
    TextBox1.Enabled = True
    ComboBox2.Enabled = True
    CheckBox8.Enabled = True
End If
    
End Sub
Private Sub TextBox6_Change()

If TextBox6.Value = True Then
    TextBox1.Enabled = False
    ComboBox2.Enabled = False
    CheckBox8.Enabled = False
Else
    TextBox1.Enabled = True
    ComboBox2.Enabled = True
    CheckBox8.Enabled = True
End If
    
End Sub
