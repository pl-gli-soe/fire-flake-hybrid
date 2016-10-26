Attribute VB_Name = "ChartForPartModule"
Private Sub show_chart()
    ChartOnPart.show vbModeless
End Sub

Public Sub hide_chart(ictrl As IRibbonControl)
    ChartOnPart.hide
End Sub

Private Function prepare_rows_on_hourly(ByRef part_num_row As Integer, ByRef ebal_row As Integer) As Boolean

    tmp = ActiveCell.row
    'in_iteration = tmp Mod 7
    'which_iteration = (tmp + 7 - in_iteration) / 7
    
    part_num_row = 7 * ((tmp + 7 - (tmp Mod 7)) / 7) - 5
    ebal_row = 7 * ((tmp + 7 - (tmp Mod 7)) / 7) - 1
    
    prepare_rows_on_hourly = True
End Function

Public Sub chart_for_part(ictrl As IRibbonControl)

    show_chart
    

    Dim tmp_chart As Chart
    Dim titles_c As Range
    Dim source_range As Range, source_data As Range
    Dim wszystkie_adresy As String
    Dim wszystkie_dni As String
    Dim wszystkie_wartosci As String
    
    Dim sh As Worksheet
    
    Set sh = Sheets("chart register")
    
    For x = 1 To sh.Shapes.COUNT
        sh.Shapes(x).Delete
    Next x
    
    Set tmp_chart = sh.Shapes.AddChart.Chart
    ' MsgBox tmp_chart.NAME
    
    wszsytkie_dni = ""
    wszystkie_wartosci = ""
    
    If ActiveSheet.Cells(1, 1) Like "daily*" Then
        If ActiveCell.row > 5 Then
            ChartOnPart.Caption = "Part number: " & CStr(Cells(ActiveCell.row, 2))
        
            For x = 17 To Int(Sheets("register").Range("lastColumn")) Step 3
                If x = 17 Then
                    wszystkie_dni = Replace(CStr(Cells(4, x).Address), "$", "")
                ElseIf x <= Int(Sheets("register").Range("lastColumn")) Then
                    wszystkie_dni = wszystkie_dni + "," + Replace(CStr(Cells(4, x).Address), "$", "")
                End If
            Next x
            
            For x = 19 To Int(Sheets("register").Range("lastColumn")) Step 3
                If x = 19 Then
                    wszystkie_wartosci = Replace(CStr(Cells(ActiveCell.row, x).Address), "$", "")
                ElseIf x <> Int(Sheets("register").Range("lastColumn")) Then
                    wszystkie_wartosci = wszystkie_wartosci + "," + Replace(CStr(Cells(ActiveCell.row, x).Address), "$", "")
                End If
            Next x
        End If
        
        
        
    ' HOURLY
    ' ===================================================================================================
    ElseIf ActiveSheet.Cells(1, 1) Like "hourly*" Then
        ' Exit Sub
        
        Dim row_on_part_number As Integer
        Dim row_ebal As Integer
        row_on_part_number = -1
        row_ebal = -1
        
        flag = prepare_rows_on_hourly(row_on_part_number, row_ebal)
        ChartOnPart.Caption = "Part number: " & CStr(Cells(row_on_part_number, 3))
        ' lastColumnHourly
        ost = Int(Sheets("register").Range("lastColumnHourly"))
        cap = Int(Sheets("register").Range("maxDataOnHourlyChart"))
        
        If ost > cap Then ost = cap
        
        For x = 9 To ost
            If x = 9 Then
                wszystkie_dni = Replace(CStr(Cells(row_on_part_number + 1, x).Address), "$", "")
                wszystkie_wartosci = Replace(CStr(Cells(row_ebal, x).Address), "$", "")
            ElseIf x <= Int(Sheets("register").Range("lastColumn")) Then
                wszystkie_dni = wszystkie_dni + "," + Replace(CStr(Cells(row_on_part_number + 1, x).Address), "$", "")
                wszystkie_wartosci = wszystkie_wartosci + "," + Replace(CStr(Cells(row_ebal, x).Address), "$", "")
            End If
        Next x
        
        
        If flag = False Then
            Exit Sub
        End If
    End If
    
    'MsgBox wszystkie_dni
    'MsgBox wszystkie_wartosci
    Set titles_c = Range(CStr(wszystkie_dni))
    Set source_range = Range(CStr(wszystkie_wartosci))
    ' titles_c.Select ' OK
    ' source_range.Select ' OK
    
    Set source_data = Union(titles_c, source_range)
    
    With tmp_chart
    
        .ChartType = xlLine
        .SetSourceData source_data, xlColumns
        .HasLegend = False
        .ChartArea.Format.Line.Visible = msoFalse
        '.ChartTitle.Text = CStr(Cells(ActiveCell.row, 2).Value)
        .SeriesCollection(1).values = source_range
        .SeriesCollection(1).XValues = titles_c
        
        With .SeriesCollection(1).Format
        
            With .Fill
                .Visible = msoTrue
                ' .ForeColor.RGB = RGB(255, 0, 0)
                .Transparency = 0
                .Solid
            End With
            
            With .Line
                .Visible = msoTrue
                .Weight = 5
            End With
        End With
        
        .AutoScaling = True
        
        With .Axes(xlValue)
            ChartOnPart.MaxValueTextBox = CStr(.MaximumScale)
            ChartOnPart.MinValueTextBox = CStr(.MinimumScale)
        End With
        
    End With


    
    tmp_chart.refresh
    fname = ThisWorkbook.Path & Application.PathSeparator & "temp.gif"
    tmp_chart.Export Filename:=fname, filtername:="GIF"
    ChartOnPart.ChartImage.Picture = LoadPicture(fname)
    
    
End Sub
