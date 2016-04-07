VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HourlyLayout"
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
' ClassModule HourlyLayout
' ========================
'
' with hourly layout is less work
' the layout is static no need to implement more options


Implements ILayout


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

End Sub

Private Sub ILayout_ChangeTxtOrientation(rng As Range, angle As Integer)
    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = Int(angle)
        
    End With
End Sub

Private Sub ILayout_ClearLayoutWithData()
    ' wstepnie caly sheet usuwam potem postaram sie sprecyzowac procedure
    Range("a1:xfd1048576").Clear
    Range("a1:xfd1048576").ClearComments
    Range("a1:xfd1048576").ClearContents
    Range("a1:xfd1048576").ClearFormats
    Range("a1:xfd1048576").ClearOutline
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


' przygotowanie layoutu tak jak wyglada aktualnie wkly coverage od o'sea
' generowanie bedzie skaldalo sie z serii iteracji i teraz jeszcze tylko kwestia
' tego czy chce zrobic tak jak to ma miejsce w dniowce to znaczy dwie procedury
' 1) zajmuj¹ca sie ogolnymi danymi
' 2) seria ostawionych poditeracji (rqm,trans,ebal)
' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------
' pelna tabelka ustawiona mniej wiecej tak jak wkly coverage
Private Sub ILayout_GenerujReszteRusztownia()
    Columns("B:" & chrx(Sheets("register").Range("ActualQtyIntervals") + 8)).AutoFit
End Sub

' procedura zajmujaca sie danymi ogolnymi
Private Sub ILayout_GenerujTabeleDlaListyDanychOgolnych()

    ' ! DRY
    Dim parts As Integer
    parts = Sheets("register").Range("allParts")
    For x = 0 To parts - 1
    
    ILayout_FillThinGridLines Range("b" & CStr(x * 7 + 2) & ":" & chrx(Sheets("register").Range("ActualQtyIntervals") + 8) & CStr(x * 7 + 6)), 0
    
    
    ' od nowa ten cmnt, back kolor pracuje na calym secie
    ' co powoduje ze dane moga byc malo czytelne "na szybko"
    ' zrobic bym chcial to tak jak na daily to znaczy naprzemienne kolory :)
    ' -----------------------------------------------------------------------------------------
    ' -----------------------------------------------------------------------------------------
    
        ' ILayout_BackColor Range("b" & CStr(x * 7 + 2) & ":" & chrx(Sheets("register").Range("ActualQtyIntervals") + 8) & CStr(x * 7 + 6)), RGB(230, 230, 250)
        
        For y = 2 To 8
            ILayout_BackColor Range(chrx(Int(y)) & CStr(x * 7 + 2) & ":" & chrx(Int(y)) & CStr(x * 7 + 6)), RGB(240, 240, 240)
        Next y
    
        For y = 9 To Int(Sheets("register").Range("ActualQtyIntervals") + 8)
        
            If ((y - 7) Mod 2) = 1 Then
                ILayout_BackColor Range(chrx(Int(y)) & CStr(x * 7 + 2) & ":" & chrx(Int(y)) & CStr(x * 7 + 6)), RGB(240, 240, 240)
            ElseIf ((y - 7) Mod 2) = 0 Then
                ILayout_BackColor Range(chrx(Int(y)) & CStr(x * 7 + 2) & ":" & chrx(Int(y)) & CStr(x * 7 + 6)), RGB(200, 200, 200)
            End If
        Next y
    
    ' -----------------------------------------------------------------------------------------
    ' -----------------------------------------------------------------------------------------
    ' liczby pomocnicze okreslone w zwiazku z miejscem w jakim znajduja sie wartosci
    ' pierwotne w sheet "register"
        For y = 2 To 19
            If y >= 2 And y <= 6 Then
                
                Range("b" & CStr(x * 7 + y)) = Sheets("register").Range("a" & CStr(y))
                ' dziwna liczba 13 wynika z ustawien w sheet "register"
                Range("h" & CStr(x * 7 + y)) = Sheets("register").Range("c" & CStr(y + 13))
                
            ElseIf y >= 7 And y <= 11 Then
                Range("d" & CStr(x * 7 + y - 5)) = Sheets("register").Range("a" & CStr(y))
            ElseIf y >= 12 And y <= 16 Then
                Range("f" & CStr(x * 7 + y - 10)) = Sheets("register").Range("a" & CStr(y))
            End If
        Next y
    Next x
End Sub
' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------

Private Sub ILayout_InitLayout(Optional for_template As Boolean)
    ' Sheets("hourly").Activate
    
    'Sheets("register").Range("lastRow") = last_row()
    'Sheets("register").Range("allParts") = last_row() - 1
    DoEvents
    Sheets.Add
    DoEvents
    ' ActiveSheet.NAME = "hourly " & CStr(Sheets.COUNT)
    Sheets("register").Range("sheetName") = ActiveSheet.Name
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
End Sub

Private Sub ILayout_UnmergeThreeCellsForDate(mrow As Integer, mfirstcolumn As Integer)
    Range(CStr(chrx(mfirstcolumn)) & CStr(mrow)).UnMerge
End Sub

Private Sub ILayout_ZwinGodzinowke(ktorydzien As Range)

End Sub
