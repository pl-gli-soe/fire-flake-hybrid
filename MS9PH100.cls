VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS9PH100"
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
  plt = Trim(Sess0.screen.getString(4, 6, 2))
End Property

Public Property Let plt(arg As String)
  Sess0.screen.putString CStr(arg), 4, 6
End Property

Public Property Get pn() As String
  pn = Trim(Sess0.screen.getString(4, 17, 8))
End Property

Public Property Let pn(arg As String)
  ' Sess0.screen.putString "00000000", 4, 17
  Sess0.screen.putString CStr(arg), 4, 17
End Property

Public Property Get DATE_FROM() As String
  DATE_FROM = Trim(Sess0.screen.getString(6, 8, 8))
End Property

Public Property Let DATE_FROM(arg As String)
  Sess0.screen.putString CStr(arg), 6, 8
End Property

Public Function fDTYP() As String
  DTYP = Trim(Sess0.screen.getString(7, 8, 11))
End Function

Public Sub setDTYP(arg1 As String, arg2 As String)
  Sess0.screen.putString CStr(arg1), 7, 8
  Sess0.screen.putString CStr(arg2), 7, 15
End Sub



Public Function SID(i As Integer) As String
    SID = Trim(Sess0.screen.getString(11 + 1 * i, 55, 9))
End Function

Public Function DUNS(i As Integer) As String
    DUNS = Trim(Sess0.screen.getString(11 + 1 * i, 65, 9))
End Function

Public Function Qty(i As Integer) As String
    Qty = Trim(Sess0.screen.getString(11 + 1 * i, 35, 10))
End Function
