VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Filtracja"
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
' ClassModule Filtracja
' ======================
'
' class that contain auto-handling algorithm for std Excel Filter
' now it's only nice prototype without components (2012-06-11)
' and only for daily data- because hourly requires complicated behavior on program
'
'
' datastore will be "register" sheet as well

Private p_filter As AutoFilter
Private criterias As Collection

Public Property Get pCriterias() As Collection
    Set pCriterias = criterias
End Property

Public Property Let pCriterias(arg As Collection)
    Set criterias = arg
End Property

Public Property Get pFilter() As AutoFilter
    pFilter = p_filter
End Property

Public Property Let pFilter(ByVal arg As AutoFilter)
    Set p_filter = arg
End Property


Public Function check_is_filter_on(sh As String) As Boolean

End Function

Public Sub get_settings(sh As String)

End Sub

Public Sub return_settings(sh As String)

End Sub

Public Sub clear_filter(sh As String)

End Sub

Private Sub Class_Initialize()
    Set criterias = New Collection
End Sub

Private Sub Class_Terminate()
    Set criterias = Nothing
End Sub
