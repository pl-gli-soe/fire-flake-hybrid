VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RozwinGodz 
   Caption         =   "set interval"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3645
   OleObjectBlob   =   "RozwinGodz.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RozwinGodz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub EnableMGOCheckBox_Click()
'    If RozwinGodz.EnableMGOCheckBox.Value = True Then
'        RozwinGodz.EnableMGOCheckBox.Value = False
'    ElseIf RozwinGodz.EnableMGOCheckBox.Value = False Then
'        RozwinGodz.EnableMGOCheckBox.Value = True
'    End If
'End Sub

Private Sub SubmitSet_Click()

    Dim catch_error As CatchError
    Set catch_error = New CatchError

    RozwinGodz.hide
    catch_error.exit_from_sub = True
    For Each r In Sheets("register").Range("IntervalComboBoxRange")
        If CStr(r) = CStr(RozwinGodz.IntervalComboBox.Value) Then
            catch_error.exit_from_sub = False
        End If
        
    Next r
    
    If catch_error.exit_from_sub = True Then
        MsgBox "You can't set this value!"
        Set catch_error = Nothing
        Exit Sub
    Else
        Set catch_error = Nothing
    End If
    
    
    If RozwinGodz.EnableMGOCheckBox.Value = True Then
        Sheets("register").Range("real_data") = "1"
    ElseIf RozwinGodz.EnableMGOCheckBox.Value = False Then
        Sheets("register").Range("real_data") = "0"
    End If
    
    
    
    Sheets("register").Range("itemInterval") = RozwinGodz.IntervalComboBox.Value
    
    Application.ScreenUpdating = False
    Sheets("register").Range("w_macro") = 1
    
    Dim l As ILayout
    Set l = New DailyLayout
    l.RozwinGodzinowke ActiveCell
    
    Sheets("register").Range("w_macro") = 0
    Application.ScreenUpdating = True
End Sub
