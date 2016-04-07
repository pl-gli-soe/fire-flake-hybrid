VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CatchError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' code written by Mateusz Milewski
'
' mateusz.milewski@ gm.com
' FireFlakeHybrid
'
'
' ClassModule CatchError
' ======================
'
'
' this is quite simple class with a few proc
' i'm still wondering about creating tree of implementations from another interfaces
' it will be a good programming manner
' now it's only loose data without any hierarchy
' but this is not so simple - it requires a lot of changes in other modules as well

Public exit_from_sub As Boolean

Private Sub Class_Initialize()
    exit_from_sub = False
End Sub

Public Sub check_selection_on_date_to_extend(arg As Range)
    
    
    If arg.COUNT <> 1 Then
        MsgBox "Zaznaczono wiêcej ni¿ jedn¹ komórkê! Zaznacz datê w 4 wierszu, któr¹ chcesz rozszerzyæ do FireFlakeHourly!"
        exit_from_sub = True
        Exit Sub
    ElseIf arg.row <> 4 Then
        MsgBox "Zaznaczono nieprawid³ow¹ komórkê- z spoza 4 wierwsza! Zaznacz datê w 4 wierszu, któr¹ chcesz rozszerzyæ do FireFlakeHourly!"
        exit_from_sub = True
        Exit Sub
    ElseIf arg.Column < 17 Then
        MsgBox "Daty zaczynaj¹ siê od kolumny Q! Zaznacz datê w 4 wierszu, któr¹ chcesz rozszerzyæ do FireFlakeHourly!"
        exit_from_sub = True
        Exit Sub
    ElseIf Not arg Like "*????-??-?? (*)*" Then
        MsgBox "nieprawny format zaznaczonego elementu, upewnij sie ze zaznaczony rekord jest opisem dnia dla zrzutu!"
        exit_from_sub = True
        Exit Sub
    End If
    
End Sub

Public Sub sprawdz_czy_mozna_zwinac(arg As Range)

    ' MsgBox CStr(Cells(arg.Row, arg.Column - 3))
    If arg.Column <> 17 Then
        If Cells(arg.row, arg.Column - 3) Like "*????-??-?? (*)*" Then
            MsgBox "Obszar jest ju¿ zwiniêty!"
            exit_from_sub = True
            Exit Sub
        End If
    ElseIf arg.Column = 17 Then
        MsgBox "Aby zwin¹æ dane nale¿y zaznaczyæ date w kolumnie z danymi  daily!"
        exit_from_sub = True
        Exit Sub
    End If

End Sub

Public Sub sprawdz_czy_mozna_rozwinac(arg As Range)

    ' MsgBox CStr(Cells(arg.Row, arg.Column - 3))
    If arg.Column <> 17 Then
        If (Not Cells(arg.row, arg.Column - 3) Like "*????-??-?? (*)*") Then
            MsgBox "Obszar jest ju¿ rozwiniêty!"
            exit_from_sub = True
            Exit Sub
        End If
    ElseIf arg.Column = 17 Then
        If Not arg Like "*????-??-?? (*)*" Then
            MsgBox "Obszar jest ju¿ rozwiniêty!"
            exit_from_sub = True
            Exit Sub
        End If
    End If

End Sub

Public Sub sprawdz_activesheet(Name As String)
    If Not ActiveSheet.Cells(1, 1) Like CStr(Name) & "*" Then
        ' MsgBox "Aktywuj odpowiedni arkusz dla wybranej procedury!"
        exit_from_sub = True
        Exit Sub
    Else
        exit_from_sub = False
        Exit Sub
    End If
End Sub

' init test MGO
Public Sub catch_err_on_mgo(str As String)
    If str <> "PURCHASE PARTS MONITORING" Then
        MsgBox "Your MGO isn't active!"
        exit_from_sub = True
        Exit Sub
    End If
End Sub

Public Sub catch_error_on_ms9ph100_skip_if_not_found(err_code As String, part As String)
    If err_code Like "*I4017*" Then
        ' MsgBox "part " & CStr(part) & " history not found!"
        exit_from_sub = True
        Exit Sub
    End If
End Sub

Public Sub catch_error_if_pn_is_not_available_in_mgo(screen_name As String, infocode As String)
    
    If screen_name = "ms9pop00" Then
        ' I4266: DATA NOT FOUND FOR SCAN REQUEST - for example
        If infocode = "I4266" Then
            exit_from_sub = True
        End If

    ElseIf screen_name = "zk7pdrqm" Or screen_name = "zk7pwrqm" Then
    
        ' R6116: PART NUMBER NOT FOUND
        ' R6105: MUST INQUIRE BEFORE PF11 CAN BE USED
        If infocode = "R6116" Or infocode = "R6105" Then
            exit_from_sub = True
        End If
    End If
End Sub

