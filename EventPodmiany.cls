VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventPodmiany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents zdarzenie As EventHandler
Attribute zdarzenie.VB_VarHelpID = -1



Private Sub Class_Initialize()
    Set zdarzenie = New EventHandler
End Sub

Private Sub Class_Terminate()
    Set zdarzenie = Nothing
End Sub

Private Sub zdarzenie_zmienDaneRejestru()
    Sheets("register").Range("ostatniaSelekcja") = Selection.Address
    Sheets("register").Range("ostatniArkusz") = ActiveSheet.Name
End Sub
