VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PWRQM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sess0 As Object

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
    pn = Trim(Sess0.screen.getString(4, 11, 8))
End Property

Public Property Let pn(arg As String)
    Sess0.screen.putString CStr(arg), 4, 11
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



Public Function weeks(SCREEN_COLUMN As Integer) As String
    weeks = Trim(Sess0.screen.getString(8, 22 + 8 * SCREEN_COLUMN, 7))
End Function

Public Function day_from_cw(cw As Integer, which_day As Integer) As Date
    
    Dim wn As Range, wd As Range, tf As Range
    Set wn = ThisWorkbook.Sheets("register").Range("weekNum")
    Set wd = ThisWorkbook.Sheets("register").Range("weekDay")
    Set tf = ThisWorkbook.Sheets("register").Range("todayFormula")
    ' swn - string on week num
    ' swd - string on week day
    Dim swn As String, swd As String, stf As String
    stf = Replace(tf.formula, "=", "")
    swn = Replace(wn.formula, "-1", stf)
    swd = Replace(wd.formula, "-1", stf)
    
    current_cw = Evaluate(swn)
    current_weekday = Evaluate(swd)
    
    diff = cw - current_cw
    ' pamietaj ze jesli chodzi o formuly weekday to 1 - sunday, 2 - monday i tak dalej
    ' 8 dlatego, ze liczymy poniedzialek od 2 stad musimy kawalek dalej pojsc
    day_from_cw = Now + diff * 7 - 7 + (8 + which_day - current_weekday)
    
End Function

Public Function monday_from_cw(cw As Integer) As Date
    
    monday_from_cw = day_from_cw(cw, 1)
End Function

Public Function sunday_from_cw(cw As Integer) As Date
    
    sunday_from_cw = day_from_cw(cw, 7)
End Function

Public Function PLT_TOTAL(SCREEN_COLUMN As Integer, ByRef pf8_pressed As Boolean) As String
    
    ' zmienna pomocnicza dla ustawienia kolejnych elementów PLT TOTAL
    ' troche Ÿle poniewa¿ za ka¿dym razem gdy uruchamiam t¹ procedure
    ' wszystkie przeliczenia ciagle sie powtarzaja...
    Dim plttotalrow As Integer
    plttotalrow = fooplttotalrow(pf8_pressed)
    
    PLT_TOTAL = Trim(Sess0.screen.getString(plttotalrow, 22 + 8 * SCREEN_COLUMN, 7))
    
    
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

    For x = 10 To 21
        If Trim(Sess0.screen.getString(x, 8, 3)) Like "*TOT*" Then
            fooplttotalrow = x
            Exit For
        End If
    Next x
End Function

