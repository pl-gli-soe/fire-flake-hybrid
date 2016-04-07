Attribute VB_Name = "FormulasForQueryInCommentModule"
Public Function componentPlantManualFill(Optional how_many_days_from_today As String, Optional rqm_table As Range, Optional transit_table As Range) As String
Attribute componentPlantManualFill.VB_Description = "This is function is allow to generate for you a Fire Flake query for custom data on requirements and transit"
Attribute componentPlantManualFill.VB_ProcData.VB_Invoke_Func = " \n14"
    
    
    ' componentPlantManualFill = "MAKE liczba MANUAL x RQM AND y TRANSIT"

    
    Dim rqm_adr As String
    Dim transit_adr As String
    
    If how_many_days_from_today = "" Then
        how_many_days_from_today = "20"
    End If
    
    If rqm_table Is Nothing Then
        rqm_adr = "EMPTY RQM"
    Else
        rqm_adr = rqm_table.Address & " " & rqm_table.Parent.Name & " " & rqm_table.Parent.Parent.Name & " RQM"
    End If
    
    If transit_table Is Nothing Then
        transit_adr = "EMPTY TRANSIT"
    Else
        transit_adr = transit_table.Address & " " & rqm_table.Parent.Name & " " & rqm_table.Parent.Parent.Name & " TRANSIT"
    End If
    
    componentPlantManualFill = "MAKE " & how_many_days_from_today & " MANUAL " & rqm_adr & " AND " & transit_adr
End Function

Public Function componentPlantRqmPopFill()


    componentPlantRqmPopFill = "MAKE X POP RQM"
End Function

