Attribute VB_Name = "ChangePercentModule"
Public Sub change_percent(ictrl As IRibbonControl)
    ChangePercentForm.ChangePercentTextBox = Sheets("register").Range("pinkOnHourly")
    ChangePercentForm.show
End Sub


