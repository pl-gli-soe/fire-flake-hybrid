VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WeeklyLayout"
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
' ClassModule WeeklyLayout
' ==========================
'
'



Implements ILayout
Implements IEbalLayout
Implements IRqmLayout


Dim catch_error As CatchError
Private ef As EbalFormula
Private p_filtracja As Filtracja
Private hybrid_iteration As IIteracja


' psuedo konstruktor dla daily od razu
' bedzie tworzyl caly layout wstepny dla dniowego layoutu
' z racji tego ¿e bedziemy mieli ff typowo dniowy bez zadnego
' urozmaicenia bedzie mozna sztywno przyjac ksztalt koncowy
' w jaki sposob kolory beda wystepowac na fireflake'u
Private Sub Class_Initialize()
End Sub

' implementacja dla rozwijania i zwijania rqm i intransit
Private Sub IEbalLayout_EbalLayoutOFF()

    Set catch_error = New CatchError
    catch_error.sprawdz_activesheet "weekly"
    
    If catch_error.exit_from_sub Then
        Exit Sub
    End If
    Set catch_error = Nothing

    ' MsgBox chrx((Int(catch_last_column())
    Dim adr As String
    Dim i As Range
    adr = chrx(Int(16)) & ":" & chrx(Int(catch_last_column()))
    Columns(CStr(adr)).EntireColumn.Hidden = False
    
    Rows("3:4").Select
    Selection.RowHeight = 15
    
    ILayout_ChangeTxtOrientation Selection, 0
    Range("b5").Select

End Sub

Private Sub IEbalLayout_EbalLayoutON()
    ' MsgBox chrx((Int(Sheets("register").Range("itemDays")) * 3) + 17)
    IEbalLayout_EbalLayoutOFF
    
    Set catch_error = New CatchError
    catch_error.sprawdz_activesheet "weekly"
    
    If catch_error.exit_from_sub Then
        Exit Sub
    End If
    Set catch_error = Nothing
    
    Dim adr As String
    Dim i As Range
    For x = 17 To Int(catch_last_column())
    
        If Cells(5, x) Like "*equirements*" Then
            adr = chrx(Int(x)) & ":" & chrx(Int(x + 1))
            Columns(CStr(adr)).EntireColumn.Hidden = True
            
        ' ElseIf x Mod 17 = 1 Then
            ' NOP
            ' poniewaz sprawe zalatwia juz pierwszy warunek
        ' ElseIf x Mod 17 = 2 Then
            ' NOP bo ma byc widoczne caly czas
        End If
    Next x
    
    Rows("3:4").Select
    Selection.RowHeight = 60
    
    ILayout_ChangeTxtOrientation Selection, 90
    Range("b5").Select
    
End Sub

Private Property Let ILayout_AllPartNumbers(RHS As Long)
    Sheets("register").Range("allParts") = RHS
End Property

Private Property Get ILayout_AllPartNumbers() As Long
    ILayout_AllPartNumbers = Sheets("register").Range("allParts")
End Property

Private Sub ILayout_BackColor(rng As Range, kolor As Long)
    With rng.Interior
        .Pattern = xlSolid
        .Color = kolor
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub ILayout_BoldOnEndingBalance(rng As Range)
    rng.Font.Bold = True
End Sub

Private Sub ILayout_ChangeTxtOrientation(rng As Range, angle As Integer)

    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = Int(angle)
        
    End With
End Sub

Private Sub ILayout_ClearLayoutWithData()
    Range("a1:" & chrx(16 + 3 * Int(Sheets("register").Range("itemWeeks"))) & Sheets("register").Range("lastRow")).Clear
    Range("a1:" & chrx(16 + 3 * Int(Sheets("register").Range("itemWeeks"))) & Sheets("register").Range("lastRow")).ClearComments
    Range("a1:" & chrx(16 + 3 * Int(Sheets("register").Range("itemWeeks"))) & Sheets("register").Range("lastRow")).ClearContents
    Range("a1:" & chrx(16 + 3 * Int(Sheets("register").Range("itemWeeks"))) & Sheets("register").Range("lastRow")).ClearFormats
    Range("a1:" & chrx(16 + 3 * Int(Sheets("register").Range("itemWeeks"))) & Sheets("register").Range("lastRow")).ClearOutline
End Sub

Private Sub ILayout_FillSolidFrame(rng As Range, kolor As Long)


    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
End Sub

Private Sub ILayout_FillSolidGridLines(rng As Range, kolor As Long)
    
    
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
End Sub

Private Sub ILayout_FillThinFrame(rng As Range, kolor As Long)


    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = kolor
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Private Sub ILayout_FillThinGridLines(rng As Range, kolor As Long)

    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub


Private Sub ILayout_GenerujReszteRusztownia()
    Dim last_column As String
    last_column = chrx(16 + 3 * Int(Sheets("register").Range("itemWeeks")))
    ' MsgBox last_column
    Dim i As Range
    Dim rng As Range
    Set rng = Range("q5:" & CStr(last_column) & CStr(Sheets("register").Range("lastRow")))
    ILayout_FillThinGridLines rng, 0
    ' pamietaj ze last kolumn zapisane jako string!
    Set rng = Range("q5:" & CStr(last_column) & CStr(5))
    ILayout_FillSolidGridLines rng, 0
    
    
    ' wypisywanie po kolei rqm, trans, ebal zgodnie z iloscia dni zachowana w rejestrze
    ' zmiana orientacji txt
    ' i backcolor dla kazdej nazwy kolumny
    For Each i In rng
    
    
            If ((i.Column - Sheets("register").Range("firstColumn") + 3) Mod 3) = 0 Then
                i = Sheets("register").Range("rqm")
            ElseIf ((i.Column - Sheets("register").Range("firstColumn") + 3) Mod 3) = 1 Then
                i = Sheets("register").Range("trans")
            ElseIf ((i.Column - Sheets("register").Range("firstColumn") + 3) Mod 3) = 2 Then
                i = Sheets("register").Range("ebal")
                ILayout_BoldOnEndingBalance Range(chrx(i.Column) & "6:" & chrx(i.Column) & CStr(Sheets("register").Range("lastRow")))
            End If
            ILayout_ChangeTxtOrientation i, 90
            ' ILayout_BackColor i, RGB(200, 200, 200)
    Next i
    
    
    ' merge oraz kolejne daty, weekdays oraz solid frame'y na datach oraz ending balance
    Dim item_d As Date
    Dim wday As String
    item_d = Format(Now, "yyyy-mm-dd")
    wday = Weekday(item_d, vbMonday)
    
    Dim zk7pw As ZK7PWRQM
    Set zk7pw = New ZK7PWRQM
    
    
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
    
    For x = Int(Sheets("register").Range("firstColumn")) To Int(16 + 3 * Int(Sheets("register").Range("itemWeeks"))) Step 3
        ILayout_MergeThreeCellsForDate 4, Int(x)
        Cells(4, x) = "CW " & CStr(Evaluate(swn))
        Cells(4, x).AddComment Format(zk7pw.monday_from_cw(Int(Evaluate(swn))), "yyyy-mm-dd") & Chr(10) & Format(zk7pw.sunday_from_cw(Int(Evaluate(swn))), "yyyy-mm-dd")
        Dim komentarz As Comment
        Set komentarz = Cells(4, x).Comment
        komentarz.Shape.Height = 30
        komentarz.Shape.Width = 70
        'If (wday = 6) Or (wday = 7) Then
        '    Cells(4, x).Interior.Color = RGB(200, 200, 0)
        'End If
        
        ' solid frame dla dat
        ILayout_FillSolidFrame Range(chrx(Int(x)) & CStr(4) & ":" & chrx(Int(x + 2)) & CStr(4)), RGB(0, 0, 0)
        Range(chrx(Int(x)) & CStr(4) & ":" & chrx(Int(x + 2)) & CStr(4)).HorizontalAlignment = xlCenter
        ' dla ending balance :)
        ILayout_FillSolidFrame Range(chrx(Int(x + 2)) & CStr(6) & ":" & chrx(Int(x + 2)) & CStr(Sheets("register").Range("lastRow"))), RGB(0, 0, 0)
        
        ' fill backcolor na co drugi item
        If (x Mod 2) = 1 Then
            ILayout_BackColor Range(chrx(Int(x)) & CStr(4) & ":" & chrx(Int(x + 2)) & CStr(Sheets("register").Range("lastRow"))), RGB(200, 200, 200)
        End If
        
        If (wday = 6) Or (wday = 7) Then
            ILayout_BackColor Cells(4, x), RGB(200, 100, 100)
        End If
        
        ' iteracja dnia oraz weekday
        item_d = item_d + 7
        wday = Weekday(item_d, vbMonday) - 1
        
        stf = Replace(tf.formula, "=", "")
        stf = stf & " + " & CStr(7 * (1 + (x - Int(Sheets("register").Range("firstColumn"))) / 3)) & "-" & wday
        swn = Replace(wn.formula, "-1", stf)
        swd = Replace(wd.formula, "-1", stf)
    
    Next x
    
    Set zk7pw = Nothing
    
    
End Sub

Private Sub ILayout_GenerujTabeleDlaListyDanychOgolnych()

    Columns("A:A").ColumnWidth = 2

    Dim rng As Range
    Set rng = Range("b5:p" & CStr(Int(ILayout_OstatniWiersz)))
    ILayout_FillThinGridLines rng, 0
    Set rng = Range("b5:p5")
    ILayout_FillSolidGridLines rng, 0
    
    Dim i As Range
    For Each i In rng
        ' ostatni 200 + i.column = 216
        ' ILayout_BackColor i, RGB(200, 200, 200)
        i = Sheets("register").Range("a" & i.Column)
        ILayout_ChangeTxtOrientation i, 90
    Next i
    
    Columns("B:P").AutoFit
    
    Set rng = Nothing
    
End Sub

Private Sub ILayout_InitLayout(Optional for_template As Boolean)
    ' Sheets("daily").Activate
    ' Sheets("register").Range("lastRow") = last_row() + 4
    ' Sheets("register").Range("allParts") = last_row() - 5
    Sheets.Add
    ' If IsMissing(for_template) Then
        ' ActiveSheet.NAME = "daily " & CStr(Sheets.COUNT)
    Sheets("register").Range("sheetName") = ActiveSheet.Name
    'ElseIf for_template = True Then
    '    ActiveSheet.NAME = "template " & CStr(Sheets.COUNT)
    '    Sheets("register").Range("sheetName") = ActiveSheet.NAME
    'End If
    
    
    ActiveWindow.Zoom = 80
    ActiveWindow.DisplayGridlines = False
    ' ILayout_ClearLayoutWithData
    ' ILayout_GenerujTabeleDlaListyDanychOgolnych
    ' ILayout_GenerujReszteRusztownia
End Sub

Private Sub ILayout_MergeThreeCellsForDate(mrow As Integer, mfirstcolumn As Integer)
    Range(CStr(chrx(mfirstcolumn)) & CStr(mrow) & ":" & CStr(chrx(mfirstcolumn + 2)) & CStr(mrow)).Merge
End Sub


Private Property Let ILayout_OstatniWiersz(RHS As Long)
    Sheets("register").Range("lastRow") = RHS
End Property

Private Property Get ILayout_OstatniWiersz() As Long
    ILayout_OstatniWiersz = Sheets("register").Range("lastRow")
End Property

Private Sub ILayout_RozwinGodzinowke(ktorydzien As Range)
        
        
    Dim sh As StatusHandler
    Application.Calculation = xlCalculationManual

    Set catch_error = New CatchError
    catch_error.sprawdz_activesheet "daily"
    catch_error.check_selection_on_date_to_extend ktorydzien
    catch_error.sprawdz_czy_mozna_rozwinac ktorydzien
    
    
    If catch_error.exit_from_sub Then
        Set catch_error = Nothing
        Exit Sub
    End If
    Set catch_error = Nothing
    
    Set sh = New StatusHandler
    sh.init_statusbar Int(Sheets("register").Range("lastRow")) - 5 + 3 * Int(Sheets("register").Range("lastRow")) + 9 ' 9 tutaj w module increasow :)
    sh.show
    
    'Dim s As Worksheet
    'Dim ff As Range
    'Set s = ActiveSheet
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
        
    End If
    
    
    sh.progress_increase
    
    'Dim ref_color As Long
    ' ref_color = RGB(200, 200, 200)

    sh.hide
    
End Sub

Private Sub ILayout_UnmergeThreeCellsForDate(mrow As Integer, mfirstcolumn As Integer)
    Range(CStr(chrx(mfirstcolumn)) & CStr(mrow)).UnMerge
End Sub

Private Sub ILayout_ZwinGodzinowke(ktorydzien As Range)
     
    Set catch_error = New CatchError
    catch_error.sprawdz_activesheet "weekly"
    catch_error.check_selection_on_date_to_extend ktorydzien
    catch_error.sprawdz_czy_mozna_zwinac ktorydzien
    
    If catch_error.exit_from_sub Then
        Set catch_error = Nothing
        Exit Sub
    End If
    Set catch_error = Nothing
    
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False

    
End Sub

Private Sub IRqmLayout_RqmLayoutOFF()
    IEbalLayout_EbalLayoutOFF
End Sub


' nie wiem czy jeszcze dzia³a!!
Private Sub IRqmLayout_RqmLayoutON()

    IEbalLayout_EbalLayoutOFF
    

    Set catch_error = New CatchError
    catch_error.sprawdz_activesheet "weekly"
    
    If catch_error.exit_from_sub Then
        Exit Sub
    End If
    Set catch_error = Nothing
    
    Dim adr As String
    Dim i As Range
    For x = 17 To Int(catch_last_column())
    
        If Cells(5, x) Like "*alance*" Then
            adr = chrx(Int(x)) & ":" & chrx(Int(x - 1))
            Columns(CStr(adr)).EntireColumn.Hidden = True
            
        ' ElseIf x Mod 17 = 1 Then
            ' NOP
            ' poniewaz sprawe zalatwia juz pierwszy warunek
        ' ElseIf x Mod 17 = 2 Then
            ' NOP bo ma byc widoczne caly czas
        End If
    Next x
    
    Rows("3:4").Select
    Selection.RowHeight = 60
    
    ILayout_ChangeTxtOrientation Selection, 90
    Range("b5").Select
    
End Sub

