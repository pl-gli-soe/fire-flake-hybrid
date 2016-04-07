Attribute VB_Name = "MainModule"
Public Sub run_ff(ictrl As IRibbonControl)
    MainForm.DTPicker1.Value = Now
    MainForm.ComboBoxColors.Clear
    MainForm.ComboBoxColors.AddItem CStr(Sheets("register").Range("KOLORY"))
    MainForm.ComboBoxColors.AddItem CStr(Sheets("register").Range("KOLORY").Offset(1, 0))
    MainForm.ComboBoxColors.Value = CStr(Sheets("register").Range("KOLORY"))
    MainForm.show
End Sub


' ponizej znajduje sie osobna mozliwosc uruchomienia z listy makr w developerze co nie jest zbyt pro
' jednak z powodu instalacji zwiazanej z nowym SAP cos sie stalo z iRibbonem i do tego czasu musze to w ten sposob
' rozwiazywac


Public Sub run_daily()
    
    Dim ffh As FireFlakeHybrid
    Set ffh = New FireFlakeHybrid
    
    ffh.p_limit = CDate(Now) + 140
    ffh.create_tear_down New ItemDaily
    
    Set ffh = noting
End Sub


Public Sub run_weekly()
    Dim ffh As FireFlakeHybrid
    Set ffh = New FireFlakeHybrid
    
    ffh.p_limit = CDate(Now) + 140
    ffh.create_tear_down New ItemWeekly
    
    Set ffh = noting
End Sub

Public Sub run_hourly()
    Dim ffh As FireFlakeHybrid
    Set ffh = New FireFlakeHybrid
    
    ffh.p_limit = CDate(Now) + 140
    ffh.create_tear_down New ItemHourly
    
    Set ffh = noting
End Sub
