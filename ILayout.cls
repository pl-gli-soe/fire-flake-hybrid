VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILayout"
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
' ClassModule IIteracja
' ========================
'
' interface 4 daily & hourly


Public Sub FillThinGridLines(rng As Range, kolor As Long)
End Sub

Public Sub FillSolidGridLines(rng As Range, kolor As Long)
End Sub

Public Sub ZwinGodzinowke(ktorydzien As Range)
End Sub

Public Sub RozwinGodzinowke(ktorydzien As Range)
End Sub

Public Sub FillSolidFrame(rng As Range, kolor As Long)
End Sub

Public Sub FillThinFrame(rng As Range, kolor As Long)
End Sub

Public Sub BackColor(rng As Range, kolor As Long)
End Sub

Public Sub MergeThreeCellsForDate(mrow As Integer, mfirstcolumn As Integer)
End Sub

Public Sub UnmergeThreeCellsForDate(mrow As Integer, mfirstcolumn As Integer)
End Sub

Public Property Get OstatniWiersz() As Long
End Property

Public Property Let OstatniWiersz(arg As Long)
End Property

Public Property Get AllPartNumbers() As Long
End Property

Public Property Let AllPartNumbers(arg As Long)
End Property

Public Sub GenerujTabeleDlaListyDanychOgolnych()
End Sub

Public Sub GenerujReszteRusztownia()
End Sub

Public Sub ChangeTxtOrientation(rng As Range, angle As Integer)
End Sub

Public Sub ClearLayoutWithData()
End Sub

Public Sub InitLayout(Optional for_template As Boolean)

End Sub

Public Sub BoldOnEndingBalance(rng As Range)

End Sub



