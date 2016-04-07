Attribute VB_Name = "Cont_RUNOUT_Module"
Sub cont_runout(ictrl As IRibbonControl)
Attribute cont_runout.VB_Description = "Creating Pivot table:\nRow -> Container\nColumn ->First Runout"
Attribute cont_runout.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Cont_RUNOUT Macro
' Creating Pivot table: Row -> Container Column ->First Runout
'

'
    shplan ActiveSheet, True


    'If ThisWorkbook.ActiveSheet.Name Like "difference report*" Then
    '    Dim rng As Range
    '    Set rng = Range("c6")
    'End If
    'Sheets.Add
    'ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    '    "Sheet1!R6C3:R15C17", Version:=xlPivotTableVersion15).CreatePivotTable _
    '    TableDestination:="Sheet2!R3C1", TableName:="PivotTable1", DefaultVersion _
    '    :=xlPivotTableVersion15
    'Sheets("Sheet2").Select
    'Cells(3, 1).Select
    'With ActiveSheet.PivotTables("PivotTable1").PivotFields("TRLR")
    '    .Orientation = xlRowField
    '    .Position = 1
    'End With
    'With ActiveSheet.PivotTables("PivotTable1").PivotFields("FST RUNOUT")
    '    .Orientation = xlColumnField
    '    .Position = 1
    'End With
    'ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
    '    "PivotTable1").PivotFields("qty for this transport"), _
    '    "Sum of qty for this transport", xlSum
    'With ActiveSheet.PivotTables("PivotTable1").PivotFields("part number")
    '    .Orientation = xlRowField
    '    .Position = 2
    'End With
    'Range("A6").Select
    'ActiveSheet.PivotTables("PivotTable1").PivotFields("TRLR").ShowDetail = False
    'With ActiveSheet.PivotTables("PivotTable1")
    '    .ColumnGrand = False
    '    .RowGrand = False
    'End With
    'ActiveSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleMedium15"
End Sub
