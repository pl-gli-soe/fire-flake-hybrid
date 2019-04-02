VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PROSL"
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
  plt = Trim(Sess0.screen.getString(4, 7, 2))
End Property

Public Property Let plt(arg As String)
  Sess0.screen.putString CStr(arg), 4, 7
End Property

Public Property Get pn() As String
  pn = Trim(Sess0.screen.getString(4, 18, 8))
End Property

Public Property Let pn(arg As String)
  Sess0.screen.putString CStr(arg), 4, 18
End Property

Public Property Get KANBAN_NUMBER() As String
  KANBAN_NUMBER = Trim(Sess0.screen.getString(4, 38, 4))
End Property

Public Property Let KANBAN_NUMBER(arg As String)
  Sess0.screen.putString CStr(arg), 4, 38
End Property



Public Property Get SUPPLIER_DUNS() As String
  SUPPLIER_DUNS = Trim(Sess0.screen.getString(4, 65, 9))
End Property

Public Property Let SUPPLIER_DUNS(arg As String)
  Sess0.screen.putString CStr(arg), 4, 65
End Property






Public Function FUNC(i As Integer) As String
    FUNC = Trim(Sess0.screen.getString(10 + 1 * i, 20, 6))
End Function

Public Function QTY(i As Integer) As String
    QTY = Trim(Sess0.screen.getString(10 + 1 * i, 10, 6))
End Function

Public Function SHIP_DATE(i As Integer) As String
    SHIP_DATE = Trim(Sess0.screen.getString(10 + 1 * i, 18, 8))
End Function

Public Function SHIP_TIME(i As Integer) As String
    SHIP_TIME = Trim(Sess0.screen.getString(10 + 1 * i, 29, 5))
End Function

Public Function DELIVERY_DATE(i As Integer) As String
    DELIVERY_DATE = Trim(Sess0.screen.getString(10 + 1 * i, 39, 8))
End Function

Public Function DELIVERY_TIME(i As Integer) As String
    DELIVERY_TIME = Trim(Sess0.screen.getString(10 + 1 * i, 53, 5))
End Function



Public Function PUS_NUMBER(i As Integer) As String
    PUS_NUMBER = Trim(Sess0.screen.getString(10 + 1 * i, 63, 14))
End Function




