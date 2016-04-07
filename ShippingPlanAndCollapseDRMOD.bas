Attribute VB_Name = "ShippingPlanAndCollapseDRMOD"
Public Sub sh_plan(ictrl As IRibbonControl)
    ' MsgBox "sp test proc"
    Application.EnableEvents = False
    
    If Cells(1, 1) Like "difference report*" Then
        ' przygotuj shipping plan na podstawie diff rep
        shplan ActiveSheet
    ElseIf (Cells(1, 1) Like "daily*") Or (Cells(1, 1) Like "hourly*") Then
        ' przygotuj napierw diff report zeby miec bazie dla shipping planu
        diffrep_inner True
        shplan ActiveSheet
    End If
    
    Application.EnableEvents = True
End Sub

Public Sub shplan(Optional ByRef wsh As Worksheet, Optional cont_runout As Boolean)

    Dim pivot_table_cache As PivotCache
    Dim pivot_table As PivotTable
    
    ' Debug.Print wsh.Cells(1, 1)
    
    

    ' Sheets.Add
    Dim szablon As ILayout
    Set szablon = New DailyLayout
    szablon.InitLayout
    Cells(1, 1) = "shipping plan " & CStr(Now) & " " & CStr(wsh.Cells(1, 1))
    
    adres = przygotowanie_adresu(wsh)
    Set pivot_table_cache = Nothing
    Set pivot_table_cache = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceData:=adres)
    Set pivot_table = ActiveSheet.PivotTables.Add(PivotCache:=pivot_table_cache, TableDestination:=Range("C6"))
    
    If Not cont_runout Then
        With pivot_table
        
            
        
            .PivotFields("plant").Orientation = xlRowField
            .PivotFields("plant").Position = 1
                
            .PivotFields("part number").Orientation = xlRowField
            .PivotFields("part number").Position = 2
            
            .PivotFields("name").Orientation = xlRowField
            .PivotFields("name").Position = 3
            
            '.PivotFields("regular transport").Orientation = xlRowField
            '.PivotFields("regular transport").Position = 4
            
            .PivotFields("pickup date").Orientation = xlColumnField
            .PivotFields("pickup date").Position = 1
            
            .PivotFields("qty for this transport").Orientation = xlDataField
            
            .PivotFields("regular transport").Orientation = xlPageField
            .PivotFields("regular transport").Position = 1
            
            .ColumnGrand = False
            .RowGrand = False
            
            .TableStyle2 = "PivotStyleMedium6"
            
            
            .PivotFields("plant").LayoutBlankLine = _
                True
            .PivotFields("part number"). _
                LayoutBlankLine = True
            .PivotFields("name").LayoutBlankLine = _
                True
            .PivotFields("delivery date"). _
                LayoutBlankLine = True
            .PivotFields("pickup date"). _
                LayoutBlankLine = True
            .PivotFields("qty for this transport"). _
                LayoutBlankLine = True
            .PivotFields("value in cell"). _
                LayoutBlankLine = True
            .PivotFields("difference"). _
                LayoutBlankLine = True
            .PivotFields("valid change"). _
                LayoutBlankLine = True
            .PivotFields("regular transport"). _
                LayoutBlankLine = True
            .PivotFields("qty for this transport"). _
                LayoutBlankLine = True
        End With
    ElseIf cont_runout Then
        
        
        With pivot_table
        
            .PivotFields("TRLR").Orientation = xlRowField
            .PivotFields("TRLR").Position = 1
            
        
            .PivotFields("part number").Orientation = xlRowField
            .PivotFields("part number").Position = 2
            
            .PivotFields("FST RUNOUT").Orientation = xlColumnField
            .PivotFields("FST RUNOUT").Position = 1
            
            .PivotFields("qty for this transport").Orientation = xlDataField
            .PivotFields("ST").Orientation = xlPageField
            .PivotFields("ST").Position = 1
            
            .PivotFields("TRLR").ShowDetail = False
            
            .ColumnGrand = False
            .RowGrand = False
            
            .TableStyle2 = "PivotStyleMedium15"
        End With
    End If
End Sub

Private Function przygotowanie_adresu(ByRef wsh As Worksheet)
    adr = "C6"
    lr = last_row("c6", CStr(wsh.Name))
    adr = adr & ":Q" & CStr(lr)
    przygotowanie_adresu = wsh.Name & "!" & adr
End Function




' maly diff rep (collapse)
Public Sub diffrep(ictrl As IRibbonControl)
    
    Application.EnableEvents = False

    If Cells(1, 1) Like "difference report*" Then
        ' przygotuj shipping plan na podstawie diff rep
        collapse_difference_report ActiveSheet
    ElseIf (Cells(1, 1) Like "daily*") Or (Cells(1, 1) Like "hourly*") Or (Cells(1, 1) Like "weekly*") Then
        ' przygotuj napierw diff report zeby miec bazie dla shipping planu
        diffrep_inner True
        collapse_difference_report ActiveSheet
    End If
    
    Application.EnableEvents = True
End Sub

Private Sub collapse_difference_report(ByRef wsh As Worksheet)

    ' Sheets.Add
    Dim szablon As ILayout
    Set szablon = New DailyLayout
    szablon.InitLayout
    Cells(1, 1) = "collapse difference report " & CStr(Now) & " " & CStr(wsh.Cells(1, 1))
    
    adres = przygotowanie_adresu(wsh)
    
    Dim rng As Range
    Set rng = Range(adres)
    ' rng.Address adres
    ' Debug.Print rng.item(1, 1)
    
    Dim r As Range
    Dim t As Range
    Set t = Range("c6")
    For Each r In rng
    
    
        ' Debug.Print r.Parent.Cells(r.row, 12).Value
        If (r.Parent.Cells(r.row, 12) Like "*manual*") Or (r.Parent.Cells(r.row, 12) = "regular transport") Then

            ' plt
            If r.Column = 3 Then
                t = r
                Set t = t.Offset(0, 1)
            
            ' pn
            ElseIf r.Column = 4 Then
                t = r
                Set t = t.Offset(0, 1)
            
            ' name
            ElseIf r.Column = 5 Then
                t = r
                Set t = t.Offset(0, 1)
            
            ' pu Date
            ElseIf r.Column = 6 Then
                t = CStr(r.Offset(0, 1))
                Set t = t.Offset(0, 1)
            
            ' eda
            ElseIf r.Column = 7 Then
                t = CStr(r.Offset(0, -1))
                Set t = t.Offset(0, 1)
            
            ' value
            ElseIf r.Column = 9 Then
                t = r
                Set t = t.Offset(0, 1)
            
            ' diff
            ElseIf r.Column = 10 Then
                t = r
                
                If (Not checksum_concat_on_row(t, -6)) And (t <> 0) Then
                    t.Offset(0, 1).FormulaR1C1 = "=RC[-2]-RC[-1]"
                    Set t = t.Offset(1, -6)
                Else
                
                    Range(t.Offset(0, -6), t).Clear
                    Set t = t.Offset(0, -6)
                End If
                
            End If
        End If
    Next r
    
    szablon.FillSolidGridLines Range("C7").CurrentRegion, RGB(0, 0, 0)
    szablon.FillSolidFrame Range("C7").CurrentRegion, RGB(0, 0, 0)
    Range("C6:J6").Interior.Color = RGB(200, 200, 200)
    
    Range("H6") = "value in cell"
    Range("J6") = "prev value in cell"
    Columns("C:J").AutoFit
    
End Sub


Private Function checksum_concat_on_row(ByRef t As Range, n As Integer) As Boolean

    curr = ""
    prev = ""
    
    Dim i As Range
    For Each i In Range(t.Offset(0, n), t)
        If (i.Column <> 5) And (i.Column <> 7) Then
            curr = curr + CStr(i)
        End If
        
    Next i
    

    For Each i In Range(t.Offset(-1, n), t.Offset(-1, 0))
        If (i.Column <> 5) And (i.Column <> 7) Then
            prev = prev + CStr(i)
        End If
    Next i
    
    If prev = curr Then
        checksum_concat_on_row = True
    Else
        checksum_concat_on_row = False
    End If
End Function
