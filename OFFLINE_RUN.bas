Attribute VB_Name = "OFFLINE_RUN"
Public Sub offline_run_daily_standard()



    Dim ff As FireFlakeHybrid
    Set ff = New FireFlakeHybrid
    
    Sheets("register").Range("redpink") = CStr(Sheets("register").Range("KOLORY"))
    
    

    Sheets("register").Range("miscFromDailyRqm") = 0
    
    Sheets("register").Range("limitDate") = CStr(Format(CDate(Now + 100), "yyyy-mm-dd"))
    ff.p_limit = CDate(Now + 100)

    

    Sheets("register").Range("limitDateDelivery") = CStr(Format(CDate(Now + 100), "yyyy-mm-dd"))
    ff.p_limit_delivery = CDate(Now + 100)
    
    ff.create_tear_down New ItemDaily
End Sub


Public Sub offline_run_weekly_standard()


    Dim ff As FireFlakeHybrid
    Set ff = New FireFlakeHybrid
    
    Sheets("register").Range("redpink") = CStr(Sheets("register").Range("KOLORY"))
    
    Sheets("register").Range("limitDate") = CStr(Format(CDate(Now + 100), "yyyy-mm-dd"))
    ff.p_limit = CDate(Now + 50 * 7)

    

    Sheets("register").Range("limitDateDelivery") = CStr(Format(CDate(Now + 100), "yyyy-mm-dd"))
    ff.p_limit_delivery = CDate(Now + 100)

    
    
    
    ff.create_tear_down New ItemWeekly
End Sub
