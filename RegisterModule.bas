Attribute VB_Name = "RegisterModule"
Public Sub show_register_sheet(ictrl As IRibbonControl)
    ret = MsgBox("Are you sure to open register?", vbQuestion + vbYesNo)
    If ret = vbYes Then
        Sheets("register").Visible = True
        Sheets("register").Activate
    End If
End Sub

Private Sub hide_register()
    Sheets("register").Visible = False
End Sub
