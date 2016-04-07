VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Main"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8160
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' WEEEKLY RUN
Private Sub CommandButton1_Click()
    MainForm.hide
    
    
    
    Dim ff As FireFlakeHybrid
    Set ff = New FireFlakeHybrid
    
    Sheets("register").Range("redpink") = Me.ComboBoxColors.Value
    
    If MainForm.DTPicker1.Enabled = True Then
        ff.p_limit = CDate(MainForm.DTPicker1.Value)
        Sheets("register").Range("limitDate") = CStr(CDate(ff.p_limit))
    Else
        Sheets("register").Range("limitDate") = CStr(Format(CDate(Now + 100), "yyyy-mm-dd"))
        ff.p_limit = CDate(Now + 50 * 7)
    End If
    
    ' ffh 3.99
    If MainForm.DTPicker2.Enabled = True Then
        ff.p_limit_delivery = CDate(MainForm.DTPicker2.Value)
        Sheets("register").Range("limitDateDelivery") = CStr(CDate(ff.p_limit_delivery))
    Else
        Sheets("register").Range("limitDateDelivery") = CStr(Format(CDate(Now + 100), "yyyy-mm-dd"))
        ff.p_limit_delivery = CDate(Now + 100)
    End If
    ' tutaj w sposob prosty
    ';
    '
    ' jednak dziwi mnie nieco fakt w ktorym pomimo tego ze domyslnie
    ' ustawione sa argumenty byref to i tak nie robi mi
    ' problemu kreowania w dalszym ciagu
    ' algorytmu ...
    '
    '
    '
    ' ff.create_tear_down New ItemHourly
    
    
    
    ff.create_tear_down New ItemWeekly
    
    
    ' ff.createTeardown New ItemHourly
    
    ' MsgBox TypeName(ff.p_item)
End Sub

Private Sub EnableLimitDateCheckBox_Click()
    
    
    If MainForm.EnableLimitDateCheckBox.Value = True Then
        MainForm.DTPicker1.Enabled = True
    ElseIf MainForm.EnableLimitDateCheckBox.Value = False Then
        MainForm.DTPicker1.Enabled = False
    End If
End Sub

Private Sub EnableLimitDateCheckBox2_Click()


    If MainForm.EnableLimitDateCheckBox2.Value = True Then
        MainForm.DTPicker2.Enabled = True
    ElseIf MainForm.EnableLimitDateCheckBox2.Value = False Then
        MainForm.DTPicker2.Enabled = False
    End If

End Sub

Private Sub RunDaily_Click()
    MainForm.hide
    
    
    
    Dim ff As FireFlakeHybrid
    Set ff = New FireFlakeHybrid
    
    Sheets("register").Range("redpink") = Me.ComboBoxColors.Value
    
    
    ' na potrzeby tylko i wylacznie ffh'a 3.96.1
    If Me.CheckboxMiscFromDRqm.Value Then
        Sheets("register").Range("miscFromDailyRqm") = 1
    Else
        Sheets("register").Range("miscFromDailyRqm") = 0
    End If
    
    If MainForm.DTPicker1.Enabled = True Then
        ff.p_limit = CDate(MainForm.DTPicker1.Value)
        Sheets("register").Range("limitDate") = CStr(CDate(ff.p_limit))
    Else
        Sheets("register").Range("limitDate") = CStr(Format(CDate(Now + 100), "yyyy-mm-dd"))
        ff.p_limit = CDate(Now + 100)
    End If
    
    
    ' ffh 3.99
    If MainForm.DTPicker2.Enabled = True Then
        ff.p_limit_delivery = CDate(MainForm.DTPicker2.Value)
        Sheets("register").Range("limitDateDelivery") = CStr(CDate(ff.p_limit_delivery))
    Else
        Sheets("register").Range("limitDateDelivery") = CStr(Format(CDate(Now + 100), "yyyy-mm-dd"))
        ff.p_limit_delivery = CDate(Now + 100)
    End If
    ' tutaj w sposob prosty
    ';
    '
    ' jednak dziwi mnie nieco fakt w ktorym pomimo tego ze domyslnie
    ' ustawione sa argumenty byref to i tak nie robi mi
    ' problemu kreowania w dalszym ciagu
    ' algorytmu ...
    '
    '
    '
    ' ff.create_tear_down New ItemHourly
    
    
    
    ff.create_tear_down New ItemDaily
    
    
    ' ff.createTeardown New ItemHourly
    
    ' MsgBox TypeName(ff.p_item)
End Sub

Private Sub RunHourly_Click()
    MainForm.hide
    
    
    
    Dim ff As FireFlakeHybrid
    Set ff = New FireFlakeHybrid
    
    Sheets("register").Range("redpink") = Me.ComboBoxColors.Value
    
    If MainForm.DTPicker1.Enabled = True Then
        ff.p_limit = CDate(MainForm.DTPicker1.Value)
        Sheets("register").Range("limitDate") = CStr(CDate(ff.p_limit))
    Else
        Sheets("register").Range("limitDate") = CDate(Now + 100)
        ff.p_limit = CDate(Now + 100)
    End If
    
    ' ffh 3.99
    If MainForm.DTPicker2.Enabled = True Then
        ff.p_limit_delivery = CDate(MainForm.DTPicker2.Value)
        Sheets("register").Range("limitDateDelivery") = CStr(CDate(ff.p_limit_delivery))
    Else
        Sheets("register").Range("limitDateDelivery") = CStr(Format(CDate(Now + 100), "yyyy-mm-dd"))
        ff.p_limit_delivery = CDate(Now + 100)
    End If
    ' tutaj w sposob prosty
    ';
    '
    ' jednak dziwi mnie nieco fakt w ktorym pomimo tego ze domyslnie
    ' ustawione sa argumenty byref to i tak nie robi mi
    ' problemu kreowania w dalszym ciagu
    ' algorytmu ...
    '
    '
    '
    ' ff.create_tear_down New ItemHourly
    
    
    
    ff.create_tear_down New ItemHourly
    
    
    ' ff.createTeardown New ItemHourly
    
    ' MsgBox TypeName(ff.p_item)
End Sub

Private Sub TemplateBtn_Click()
    MainForm.hide
    TemplateConfig.StartDTPicker.Value = Now
    TemplateConfig.EndDTPicker.Value = Now
    
    Sheets("register").Range("redpink") = Me.ComboBoxColors.Value
    
    TemplateConfig.show
End Sub

