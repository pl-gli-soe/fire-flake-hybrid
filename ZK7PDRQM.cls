VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PDRQM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sess0 As Object
' zmienna przechowujaca reszte danych z zk7pdrqm
Private daily_matrix As ZK7PDRQMDailyMatrix

Public Property Let sesja(ByRef arg As Object)
    Set Sess0 = arg
End Property

Public Property Get sesja() As Object
    Set sesja = Sess0
End Property

Public Property Get infocode() As String
    infocode = Trim(Sess0.screen.getString(22, 2, 5))
End Property

Public Property Get plt() As String
    plt = Trim(Sess0.screen.getString(3, 8, 2))
End Property

Public Property Let plt(arg As String)
    Sess0.screen.putString CStr(arg), 3, 8
End Property

Public Property Get pn() As String
    pn = Trim(Sess0.screen.getString(4, 8, 8))
End Property

Public Property Let pn(arg As String)
    Sess0.screen.putString CStr(arg), 4, 8
End Property

Public Property Get KANBAN() As String
    KANBAN = Trim(Sess0.screen.getString(5, 10, 4))
End Property

Public Property Let KANBAN(arg As String)
    Sess0.screen.putString CStr(arg), 5, 10
End Property

Public Property Get PART_DESC() As String
    PART_DESC = Trim(Sess0.screen.getString(4, 41, 10))
End Property


Public Property Get p_daily_matrix() As ZK7PDRQMDailyMatrix
    Set p_daily_matrix = daily_matrix
End Property

Public Property Let p_daily_matrix(arg As ZK7PDRQMDailyMatrix)
    Set p_daily_matrix = arg
End Property



Public Function dates(SCREEN_COLUMN As Integer) As String
    dates = Trim(Sess0.screen.getString(8, 24 + 8 * SCREEN_COLUMN, 5))
End Function

Public Function PLT_TOTAL(SCREEN_COLUMN As Integer, ByRef pf8_pressed As Boolean) As String
    
    ' zmienna pomocnicza dla ustawienia kolejnych elementów PLT TOTAL
    ' troche Ÿle poniewa¿ za ka¿dym razem gdy uruchamiam t¹ procedure
    ' wszystkie przeliczenia ciagle sie powtarzaja...
    Dim plttotalrow As Integer
    plttotalrow = fooplttotalrow(pf8_pressed)
    
    
    Dim misc_value As Long
    misc_value = 0
    
    
    If Int(ThisWorkbook.Sheets("register").Range("miscFromDailyRqm")) = 1 Then
    
        For x = 1 To 12
        
        
            If Trim(Sess0.screen.getString(plttotalrow - 1, 1, 10)) Like "*MISC*OTHR*" Then
                misc_value = CLng(Trim(Sess0.screen.getString(plttotalrow - 1, 22 + 8 * SCREEN_COLUMN, 12)))
                Exit For
            End If
            
        Next x
    
    End If
    
    

    PLT_TOTAL = _
        CLng(Trim(Sess0.screen.getString(plttotalrow, 22 + 8 * SCREEN_COLUMN, 7))) - CLng(misc_value)
            

    
    
End Function



Private Function fooplttotalrow(ByRef pf8_pressed As Boolean) As Integer

    ' While ((Trim(Sess0.screen.getString(22, 2, 5)) <> "R6101") And (Trim(Sess0.screen.getString(22, 2, 5)) <> "R6123"))
    While (Trim(Sess0.screen.getString(22, 2, 5)) = "R6693") Or ((Trim(Sess0.screen.getString(22, 2, 5)) = "R6102"))
        Sess0.screen.sendKeys ("<pf8>")
        pf8_pressed = True
        Do
            DoEvents
        Loop Until Sess0.screen.OIA.XStatus = 0
        
    Wend

    For x = 9 To 21
        If Trim(Sess0.screen.getString(x, 2, 9)) Like "*PLT TOTAL*" Then
        
            fooplttotalrow = x
            Exit For
        End If
    Next x
End Function
