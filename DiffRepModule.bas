Attribute VB_Name = "DiffRepModule"

Public Sub cdiffrep(ictrl As IRibbonControl)
    ' MsgBox "diff report implementation in progress"
    Application.EnableEvents = False
    diffrep_inner True
    Application.EnableEvents = True
End Sub

' slabo ale musi byc global ponieaz korzystaja z innego scopu metody tez
Public Sub diffrep_inner(to_hide As Boolean)
    Dim difference_report As DiffReport
    
    If (CStr(Cells(1, 1)) Like "daily*") Or (CStr(Cells(1, 1)) Like "hourly*") Or (CStr(Cells(1, 1)) Like "weekly*") Then
        
        Set difference_report = New DiffReport
        
        difference_report.find_differences
        difference_report.create_diff_report
        
        Set difference_report = Nothing
        
    End If
End Sub

Public Sub cmaoi(ictrl As IRibbonControl)
    ' clear_manual_adjustment_on_intransit
    MsgBox "not yet implemented"
End Sub


' procedura ta prawdopodobnie zostanie juz na zawsze obsoletem z racji tego ze zbyt duzo czasu zajmuje jej wykonanie
' w stosunku do jej waznosci
' dzialac dziala ale to jednak nie o to chyba jednak chodzi
Private Sub clear_manual_adjustment_on_intransit()


    
End Sub
