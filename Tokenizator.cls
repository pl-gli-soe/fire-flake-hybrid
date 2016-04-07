VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tokenizator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' zmienna ta pobiera caly ciag znakow odpowiedzialny za sterowanie poczynaniami Fire Flake'a
Private m_token As String

Private make_ As Boolean
Private iteration_value As Integer
Private screen_handler As String
Private manual_fill As Boolean
Private rqm_adr_table As Range
Private transit_adr_table As Range

Private tmp_local_adr As String
Private i_on_local_adres As Integer
Private tmp_sheet_name As String
Private tmp_workbook_name As String

Private flaga_czy_rqm_adr_table_jest_gotowy As Boolean

Property Get getMake_()
    getMake_ = make_
End Property

Property Get getIteration()
    getIteration = iteration_value
End Property

Property Let setIteration(arg As Integer)
    iteration_value = Int(arg)
End Property

Property Get mgoScreen()
    mgoScreen = screen_handler
End Property

Property Get getIsManualFill()
    getIsManualFill = manual_fill
End Property

Property Get getCustomRqmTable() As Range
    Set getCustomRqmTable = rqm_adr_table
End Property

Property Get getCustomTransitTable() As Range
    Set getCustomTransitTable = transit_adr_table
End Property




Property Let token(arg As String)

    If Trim(arg) <> "" Then
        m_token = CStr(arg)
        solve_query
    ElseIf Trim(arg) = "" Then
        m_token = "DEFAULT"
    End If
End Property

Property Get token() As String
    token = CStr(m_token)
End Property

Private Sub Class_Initialize()
    m_token = "DEFAULT"
    tmp_local_adr = ""
    tmp_sheet_name = ""
    tmp_workbook_name = ""
    i_on_local_adres = -1
    flaga_czy_rqm_adr_table_jest_gotowy = False
End Sub

Private Sub solve_query()

    Set rqm_adr_table = Nothing
    flaga_czy_rqm_adr_table_jest_gotowy = False

    tmp_local_adr = ""
    tmp_sheet_name = ""
    tmp_workbook_name = ""
    i_on_local_adres = -1

    arr = Split(m_token, " ")
    

    
    i = 1
    For Each s In arr
    
        ' sprawdz czy w ogóle jest make
        If s = "MAKE" Then
            make_ = True
        End If
        
        If (i = 2) And (IsNumeric(s)) Then
            iteration_value = Int(s)
        ElseIf (i = 2) And Not (IsNumeric(s)) Then
            iteration_value = -1
        End If
        
        If (i = 3) And (s = "MANUAL") Then
            manual_fill = True
        ElseIf (i = 3) And (s = "POP") Then
            manual_fill = False
            screen_handler = "MS9POP00"
            Exit Sub
        End If
        
        If flaga_czy_rqm_adr_table_jest_gotowy = False Then
        
            custom_data_ranges i, rqm_adr_table, CStr(s)
        Else
            ' liczba 5 trzeciego argumentu oznacza
            ' o ile trzeba przesunac slow zapytania zeby zalapac tabele dla rqmsow
            custom_data_ranges i, transit_adr_table, CStr(s)
        End If
        
        
        i = i + 1
    Next s
End Sub

Private Sub custom_data_ranges(ByVal i As Integer, ByRef rng As Range, s As String)


    If Left(s, 1) = "$" Then
        
        tmp_local_adr = s
        i_on_local_adres = i
    
    ' ElseIf (i = 4) And (s = "EMPTY") Then
        ' puste dane wypelnione ino zerami
    End If
    
    If (i = i_on_local_adres + 1) And (tmp_local_adr <> "") Then
        tmp_sheet_name = s
    End If
    
    If (i = i_on_local_adres + 2) And (tmp_local_adr <> "") Then
        tmp_workbook_name = s
        
        Set rng = Application.Workbooks(CStr(tmp_workbook_name)).Sheets(CStr(tmp_sheet_name)).Range(CStr(tmp_local_adr))
        
        tmp_local_adr = ""
        tmp_sheet_name = ""
        tmp_workbook_name = ""
        i_on_local_adres = -1
        flaga_czy_rqm_adr_table_jest_gotowy = True
    End If
End Sub

'    componentPlantManualFill = "MAKE " & how_many_days_from_today & " MANUAL " & rqm_adr & " AND " & transit_adr
'End Function'
'
'Public Function componentPlantRqmPopFill()
'
'
'    componentPlantPopFill = "MAKE X POP RQM"
'End Function
