VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event zmienDaneRejestru()

Public Sub selection_zostal_zmieniony()
    
    If (ActiveSheet.Name <> "input") And (ActiveSheet.Name <> "register") Then
    
        If (Cells(1, 1) Like "daily*") Or (Cells(1, 1) Like "hourly*") Or (Cells(1, 1) Like "weekly*") Then
            RaiseEvent zmienDaneRejestru
        End If
    End If
End Sub



