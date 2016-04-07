VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PPUM0"
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

Public Property Get REC() As String
  REC = Trim(Sess0.screen.getString(3, 65, 5))
End Property

Public Property Get OF() As String
  OF = Trim(Sess0.screen.getString(3, 75, 5))
End Property



Public Function plt(i As Integer) As String
    plt = Trim(Sess0.screen.getString(7 + 1 * i, 5, 2))
End Function

Public Function SUPPLIER(i As Integer) As String
    SUPPLIER = Trim(Sess0.screen.getString(7 + 1 * i, 12, 9))
End Function
Public Function PU_DATE(i As Integer) As String
    PU_DATE = Trim(Sess0.screen.getString(7 + 1 * i, 23, 8))
End Function

Public Function PUS_NBR(i As Integer) As String
    PUS_NBR = Trim(Sess0.screen.getString(7 + 1 * i, 33, 9))
End Function

Public Function route(i As Integer) As String
    route = Trim(Sess0.screen.getString(7 + 1 * i, 44, 6))
End Function

Public Function DEL_DATE(i As Integer) As String
    DEL_DATE = Trim(Sess0.screen.getString(7 + 1 * i, 52, 8))
End Function

Public Function DEL_WINDW(i As Integer) As String
    DEL_WINDW = Trim(Sess0.screen.getString(7 + 1 * i, 62, 9))
End Function

Public Function BEGIN_DEL(i As Integer) As String
    BEGIN_DEL = Trim(Sess0.screen.getString(7 + 1 * i, 62, 4))
End Function
Public Function END_DEL(i As Integer) As String
    END_DEL = Trim(Sess0.screen.getString(7 + 1 * i, 67, 4))
End Function

Public Function PART_QTY(i As Integer) As String
    PART_QTY = Trim(Sess0.screen.getString(7 + 1 * i, 72, 9))
End Function
