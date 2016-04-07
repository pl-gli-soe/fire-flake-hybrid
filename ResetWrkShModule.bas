Attribute VB_Name = "ResetWrkShModule"
Public Sub reset_report(ictrl As IRibbonControl)


    Sheets("register").Range("scopeObliczen") = "all"
    przelicz_parametry_arkusza ActiveSheet, ActiveCell
    przelicz_arkusz ActiveSheet, ActiveCell, True
    Sheets("register").Range("scopeObliczen") = Sheets("register").Range("scopeObliczenDefault")
    DoEvents
    Sheets("register").Range("w_macro") = 0
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
