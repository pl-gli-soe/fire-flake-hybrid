VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EbalFormula"
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
' ClassModule DiffReport
' ======================
'
' nice class with built-in formulas

Private formula As String
Private przesuniecie As Integer
Private interval As Integer
Private temp As Range

Public Property Get fl() As String
    fl = CStr(formula)
End Property

Public Property Let fl(arg As String)
    fromula = CStr(arg)
End Property

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub

Public Function fill_first_day() As String
    '            CBAL    MISC   NYR    RQM   TRANS
    formula = "=RC[-15]-RC[-7]+RC[-6]-RC[-2]+RC[-1]"
    fill_first_day = formula
End Function

Public Function fill_other_days() As String
    '          EBAL-1   RQM    TRANS
    formula = "=RC[-3]-RC[-2]+RC[-1]"
    fill_other_days = formula
End Function

Public Function fill_other_days_przed_rozwinieciem() As String
    '            EBAL
    formula = "=RC[-3]"
    fill_other_days_przed_rozwinieciem = formula
End Function

Public Function hourly_fill_first_day() As String
    '            CBAL         MISC        NYR     RQM   TRANS
    formula = "=R[-2]C[-6]-R[-4]C[-2]+R[-3]C[-2]-R[-2]C+R[-1]C"
    hourly_fill_first_day = formula
End Function

Public Function hourly_fill_other_days() As String
    '          EBAL-1   TRANS  RQM
    formula = "=RC[-1]+R[-1]C-R[-2]C"
    hourly_fill_other_days = formula
End Function

Public Function fill_first_day_przed_zwinieciem(q As Range) As String
    interval = CDate((CDate(Cells(4, q.Column - 5)) - CDate(Cells(4, q.Column - 8))) * 24)
    If interval < 0 Then interval = interval + 24
    przesuniecie = Int(48 / CDbl(interval)) * 3
    '            CBAL         MISC        NYR     RQM   TRANS - - bardzo wazne ze trans oraz rqm nie trzeba przesuwac wzgledem rozszerzenia godzinowego
    formula = "=RC[-" & CStr(przesuniecie + 15) & "]-RC[-" & CStr(przesuniecie + 7) _
    & "]+RC[-" & CStr(przesuniecie + 6) & "]-RC[-" & CStr(0 + 2) & "]+RC[-" & CStr(0 + 1) & "]"
    fill_first_day_przed_zwinieciem = formula
End Function

Public Function fill_other_days_przed_zwinieciem(q As Range) As String
    ' interval = Int(CDbl(CDate(Cells(4, q.Column - 1) - CDate(Cells(4, q.Column - 2)))) * 24)
    ' MsgBox q.Address
    ' Debug.Print Cells(4, q.Column - 5)
    ' Debug.Print Cells(4, q.Column - 8)
    
    interval = CDate((CDate(Cells(4, q.Column - 5)) - CDate(Cells(4, q.Column - 8))) * 24)
    If interval < 0 Then interval = interval + 24
    
    
    
    przesuniecie = Int(48 / CDbl(interval)) * 3
    '          EBAL-1   TRANS  RQM - bardzo wazne ze trans oraz rqm nie trzeba przesuwac wzgledem rozszerzenia godzinowego
    formula = "=RC[-" & CStr(przesuniecie + 3) & "]-RC[-" & CStr(0 + 2) & "]+RC[-" & CStr(0 + 1) & "]"
    fill_other_days_przed_zwinieciem = formula
End Function

Public Function sum_requirements(q As Range) As String

    Set temp = q
    Dim i As Integer
    i = 1
    ' interval = Int(CDbl(CDate(Cells(4, q.Column - 1) - CDate(Cells(4, q.Column - 2)))) * 24)
    
    formula = "="
    
    Set temp = temp.Offset(0, -3)
    Do
    
        formula = formula & "RC[" & CStr(-3 * i) & "]+"
    
        i = i + 1
        Set temp = temp.Offset(0, -3)
        
        If (Cells(5, temp.Column) Like "*tandard*ack*") Then
            Exit Do
        End If
        
        If (Cells(4, temp.Column) Like "*????-??-?? *(*)*") Then
            Exit Do
        End If
        
        If Not (CStr(CDate(Cells(4, temp.Column))) Like "*??:??*") Then
            Exit Do
        End If
       
    Loop While True
    
    formula = Left(formula, Len(formula) - 1)
    
    sum_requirements = formula
End Function

Public Function sum_transit(q As Range) As String

    Set temp = q
    Dim i As Integer
    i = 1
    ' interval = Int(CDbl(CDate(Cells(4, q.Column - 1) - CDate(Cells(4, q.Column - 2)))) * 24)
    
    formula = "="
    
    Set temp = temp.Offset(0, -3)
    Do
    
        formula = formula & "RC[" & CStr(-3 * i) & "]+"
    
        i = i + 1
        Set temp = temp.Offset(0, -3)
        
        If (Cells(5, temp.Column - 1) Like "*tandard*ack*") Then
            Exit Do
        End If
        
        If (Cells(4, temp.Column - 1) Like "*????-??-?? *(*)*") Then
            Exit Do
        End If
        If Not (CStr(CDate(Cells(4, temp.Column - 1))) Like "*??:??*") Then
            Exit Do
        End If
       
    Loop While True
    
    formula = Left(formula, Len(formula) - 1)
    
    sum_transit = formula
End Function
