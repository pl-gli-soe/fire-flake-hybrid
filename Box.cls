VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Box"
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
' ClassModule Box
' ===============
'
' the handler for global public function chrx()
' one of the independent pieces of code
' no relations with the main algorithm
' from the begining no changes simple
' public instance variables

Public counter As Integer
Public scope As Integer

Private Sub Class_Initialize()
    counter = 0
    scope = 0
End Sub

