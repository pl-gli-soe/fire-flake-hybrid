VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PPRHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sess0 As Object
' zmienna przechowujaca reszte danych z zk7pdrqm

Public Property Let sesja(ByRef arg As Object)
    Set Sess0 = arg
End Property

Public Property Get sesja() As Object
    Set sesja = Sess0
End Property

Public Property Get plt() As String
    plt = Trim(Sess0.screen.getString(3, 8, 2))
End Property

Public Property Let plt(arg As String)
    Sess0.screen.putString CStr(arg), 3, 8
End Property

Public Property Get pn() As String
    pn = Trim(Sess0.screen.getString(4, 8, 8))
End Property

Public Property Let pn(arg As String)
    Sess0.screen.putString CStr(arg), 4, 8
End Property

Public Property Get KANBAN() As String
    KANBAN = Trim(Sess0.screen.getString(5, 10, 4))
End Property

Public Property Let KANBAN(arg As String)
    Sess0.screen.putString CStr(arg), 5, 10
End Property

Public Property Get ULOC() As String
    ULOC = Trim(Sess0.screen.getString(4, 23, 8))
End Property

Public Property Let ULOC(arg As String)
    Sess0.screen.putString CStr(arg), 4, 23
End Property

Public Function HRQMDATE(r As Integer, c As Integer) As String
    HRQMDATE = Trim(Sess0.screen.getString(6 + 4 * r, 9 + 7 * c, 5))
End Function

Public Function HRQMTIME(r As Integer, c As Integer) As String
    HRQMTIME = Trim(Sess0.screen.getString(7 + 4 * r, 9 + 7 * c, 5))
End Function

Public Function HQTY(r As Integer, c As Integer) As String
    HQTY = Trim(Sess0.screen.getString(8 + 4 * r, 9 + 7 * c, 5))
End Function





