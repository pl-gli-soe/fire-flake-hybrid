VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS7P5200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sess0 As Object

Public Property Let sesja(ByRef arg As Object)
    Set Sess0 = arg
End Property

Public Property Get sesja() As Object
    Set sesja = Sess0
End Property



Public Property Get plt() As String
    plt = Trim(Sess0.screen.getString(1, 16, 2))
End Property

Public Property Let plt(arg As String)
    Sess0.screen.putString CStr(arg), 1, 16
End Property

Public Property Get pn() As String
    pn = Trim(Sess0.screen.getString(5, 14, 8))
End Property

Public Property Let pn(arg As String)
    Sess0.screen.putString CStr(arg), 5, 14
End Property

Public Property Get RECV_TYPE() As String
    RECV_TYPE = Trim(Sess0.screen.getString(13, 31, 3))
End Property

Public Property Let RECV_TYPE(arg As String)
    Sess0.screen.putString CStr(arg), 13, 31
End Property

Public Property Get STD_PACK() As String
    STD_PACK = Trim(Sess0.screen.getString(8, 67, 9))
End Property

Public Property Get Name() As String
    Name = Trim(Sess0.screen.getString(6, 45, 35))
End Property

Public Property Get DESC() As String
    DUNS = Trim(Sess0.screen.getString(6, 7, 20))
End Property



Public Property Get KANBAN() As String
    KANBAN = Trim(Sess0.screen.getString(5, 34, 4))
End Property

Public Property Let KANBAN(arg As String)
    Sess0.screen.putString CStr(arg), 5, 34
End Property

Public Property Get DUNS() As String
    DUNS = Trim(Sess0.screen.getString(5, 50, 9))
End Property

Public Property Let DUNS(arg As String)
    Sess0.screen.putString CStr(arg), 5, 50
End Property


