VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FFTemplate"
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
' ClassModule FFTemplate
' ======================
'
' this is an independent class without any iterfaces
'
' simple boxing for FireFlakeHybrid object
' made for good name convention
' in higer abstraction level
' + create_template method -> also boxing for method with the same name in FireFlakeHybrid object

Private ffh As FireFlakeHybrid

Public Property Get p_ffh() As FireFlakeHybrid
    Set p_ffh = ffh
End Property


Public Sub create_template(item As IItem, startdate As Date, enddate As Date)
    ffh.create_template item, startdate, enddate
End Sub


Private Sub Class_Initialize()
    Set ffh = New FireFlakeHybrid
End Sub

Private Sub Class_Terminate()
    Set ffh = Nothing
End Sub
