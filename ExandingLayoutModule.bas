Attribute VB_Name = "ExandingLayoutModule"
Public Sub rozwin_godz(ictrl As IRibbonControl)



    

    RozwinGodz.EnableMGOCheckBox.Value = True
    
    RozwinGodz.IntervalComboBox.Clear
    For Each r In ThisWorkbook.Sheets("register").Range("intervalsLib")
        RozwinGodz.IntervalComboBox.AddItem r.Value
    Next r


    RozwinGodz.IntervalComboBox.Value = "3"
    RozwinGodz.show
End Sub

Public Sub zwin_godz(ictrl As IRibbonControl)
    Dim l As ILayout
    Set l = New DailyLayout
    ThisWorkbook.Sheets("register").Range("w_macro") = 1
    l.ZwinGodzinowke ActiveCell
    ThisWorkbook.Sheets("register").Range("w_macro") = 0
End Sub

Public Sub tylkoEbal(ictrl As IRibbonControl)
    Dim l As IEbalLayout
    
    If Cells(1, 1) Like "*daily*" Then
        Set l = New DailyLayout
    ElseIf Cells(1, 1) Like "*weekly*" Then
        Set l = New WeeklyLayout
    End If
    
    
    l.EbalLayoutON
    ' l.EbalLayoutOFF
End Sub

Public Sub allWithEbal(ictrl As IRibbonControl)
    Dim l As IEbalLayout
    If Cells(1, 1) Like "*daily*" Then
        Set l = New DailyLayout
    ElseIf Cells(1, 1) Like "*weekly*" Then
        Set l = New WeeklyLayout
    End If
    ' l.EbalLayoutON
    l.EbalLayoutOFF
End Sub

Public Sub tylkoRqms(ictrl As IRibbonControl)
    Dim l As IRqmLayout
    If Cells(1, 1) Like "*daily*" Then
        Set l = New DailyLayout
    ElseIf Cells(1, 1) Like "*weekly*" Then
        Set l = New WeeklyLayout
    End If
    l.RqmLayoutON
End Sub

Public Sub allWithRqms(ictrl As IRibbonControl)
    Dim l As IRqmLayout
    If Cells(1, 1) Like "*daily*" Then
        Set l = New DailyLayout
    ElseIf Cells(1, 1) Like "*weekly*" Then
        Set l = New WeeklyLayout
    End If
    l.RqmLayoutOFF
End Sub
