VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS9POP00"
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

Public Property Let pn(arg As String)
    Sess0.screen.putString CStr(arg), 3, 22
End Property

Public Property Get pn() As String
    pn = Trim(Sess0.screen.getString(4, 22, 8))
End Property

Public Property Let plt(arg As String)
    Sess0.screen.putString CStr(arg), 3, 13
End Property

Public Property Get plt() As String
    plt = Trim(Sess0.screen.getString(4, 13, 2))
End Property

Public Property Let DS(arg As String)
    Sess0.screen.putString CStr(arg), 2, 18
End Property

Public Property Get DS() As String
    DS = Trim(Sess0.screen.getString(2, 18, 1))
End Property

Public Property Let DUNS(arg As String)
    Sess0.screen.putString CStr(arg), 3, 62
End Property

Public Property Get DUNS() As String
    DUNS = Trim(Sess0.screen.getString(7, 7, 9))
End Property

Public Property Get SUPPLIER() As String
    SUPPLIER = Trim(Sess0.screen.getString(7, 17, 20))
End Property

Public Property Let F_U(arg As String)
    Sess0.screen.putString CStr(arg), 3, 35
End Property

Public Property Get F_U() As String
    F_U = Trim(Sess0.screen.getString(4, 35, 2))
End Property

Public Property Let a(arg As String)
    Sess0.screen.putString CStr(arg), 3, 53
End Property

Public Property Get a() As String
    a = Trim(Sess0.screen.getString(4, 53, 2))
End Property


' pierwszy doh
Public Property Let firstDOH(arg As String)
    Sess0.screen.putString CStr(arg), 3, 43
End Property

Public Property Get firstDOH() As String
    firstDOH = Trim(Sess0.screen.getString(4, 43, 3))
End Property

' drugi doh
Public Property Let secDOH(arg As String)
    Sess0.screen.putString CStr(arg), 3, 47
End Property

Public Property Get secDOH() As String
    secDOH = Trim(Sess0.screen.getString(3, 47, 3))
End Property

Public Property Get MISC() As String
    MISC = Trim(Sess0.screen.getString(5, 73, 7))
End Property

Public Property Get OS() As String
    OS = Trim(Sess0.screen.getString(6, 67, 7))
End Property

Public Property Get Bank() As String
    Bank = Trim(Sess0.screen.getString(6, 54, 7))
End Property


Public Property Get BBAL() As String
    If Trim(Sess0.screen.getString(5, 15, 1)) = "-" Then
        BBAL = "-" & Trim(Sess0.screen.getString(5, 7, 8))
    Else
        BBAL = Trim(Sess0.screen.getString(5, 7, 8))
    End If
End Property


Public Property Get CBAL() As String
    If Trim(Sess0.screen.getString(5, 30, 1)) = "-" Then
        CBAL = "-" & Trim(Sess0.screen.getString(5, 22, 8))
    Else
        CBAL = Trim(Sess0.screen.getString(5, 22, 8))
    End If
End Property


Public Property Get PCS_TO_GO() As String
    PCS_TO_GO = Trim(Sess0.screen.getString(6, 23, 7))
End Property

Public Property Get DK() As String
    DK = Trim(Sess0.screen.getString(8, 6, 8))
End Property

Public Property Get MODE() As String
    MODE = Trim(Sess0.screen.getString(8, 21, 2))
End Property

Public Property Get TT() As String
    TT = Trim(Sess0.screen.getString(8, 31, 8))
End Property

Public Property Get MNPC() As String
    MNPC = Trim(Sess0.screen.getString(9, 7, 7))
End Property

Public Property Get NCX() As String
    NCX = Trim(Sess0.screen.getString(10, 6, 1))
End Property

Public Property Get OBS() As String
    OBS = Trim(Sess0.screen.getString(10, 14, 1))
End Property

Public Property Get STD_PACK() As String
    STD_PACK = Trim(Sess0.screen.getString(10, 30, 9))
End Property

Public Property Get oneJOB() As String
    oneJOB = Trim(Sess0.screen.getString(12, 7, 11))
End Property

Public Property Get ip() As String
    ip = Trim(Sess0.screen.getString(12, 27, 5))
End Property

Public Property Get COUNT() As String
    COUNT = Trim(Sess0.screen.getString(14, 8, 45))
End Property

Public Property Get PART_NAME() As String
    PART_NAME = Trim(Sess0.screen.getString(6, 2, 10))
End Property

Public Property Get QHD() As String
    QHD = Trim(Sess0.screen.getString(5, 48, 7))
End Property

Public Property Get O() As String
    O = Trim(Sess0.screen.getString(15, 4, 25))
End Property


Public Property Get infocode() As String
    infocode = Trim(Sess0.screen.getString(23, 2, 5))
End Property



Public Function transQTY(i As Integer) As String
    transQTY = Trim(Sess0.screen.getString(17 + (Int(i) * 2), 2, 9))
End Function

Public Function transSDATE(i As Integer) As String
    transSDATE = Trim(Sess0.screen.getString(17 + (Int(i) * 2), 18, 8))
End Function

Public Function transPCS(i As Integer) As String
    transPCS = Trim(Sess0.screen.getString(17 + (Int(i) * 2), 27, 3))
End Function

Public Function transWGT(i As Integer) As String
    transWGT = Trim(Sess0.screen.getString(17 + (Int(i) * 2), 31, 6))
End Function

Public Function transMD(i As Integer) As String
    transMD = Trim(Sess0.screen.getString(17 + (Int(i) * 2), 38, 2))
End Function

Public Function transCD(i As Integer) As String
    transCD = Trim(Sess0.screen.getString(17 + (Int(i) * 2), 41, 1))
End Function

Public Function transEDA(i As Integer) As String
    transEDA = Trim(Sess0.screen.getString(17 + (Int(i) * 2), 45, 8))
End Function

Public Function transETA(i As Integer) As String
    transETA = Trim(Sess0.screen.getString(17 + (Int(i) * 2), 54, 2)) & ":" & Trim(Sess0.screen.getString(17 + (Int(i) * 2), 56, 2))
End Function

Public Function transST(i As Integer) As String
    transST = Trim(Sess0.screen.getString(17 + (Int(i) * 2), 60, 1))
End Function

Public Function transBL(i As Integer) As String
    transBL = Trim(Sess0.screen.getString(17 + (Int(i) * 2), 63, 9))
End Function

Public Function transROUTE(i As Integer) As String
    transROUTE = Trim(Sess0.screen.getString(17 + (Int(i) * 2), 73, 7))
End Function

Public Function transCMNT(i As Integer) As String
    transCMNT = Trim(Sess0.screen.getString(18 + (Int(i) * 2), 39, 30))
End Function

Public Function transDUNS(i As Integer) As String
    transDUNS = Trim(Sess0.screen.getString(18 + (Int(i) * 2), 71, 9))
End Function

Public Function transCONT(i As Integer) As String
    transCONT = Trim(Sess0.screen.getString(18 + (Int(i) * 2), 2, 35))
End Function

' -------------------------- rqms on pop ----------------------------------------------------------
Public Function balance_on_pop_data(wiersz As Integer, col As Integer) As String
    'row 8 - 12
    'col 40, 46, 55, 65, 73, size 8
    Dim kolumna As Integer
    Select Case col
        Case 1
            kolumna = 40
        Case 2
            kolumna = 46
        Case 3
            kolumna = 55
        Case 4
            kolumna = 65
        Case 5
            kolumna = 73
    End Select
    
    balance_on_pop_data = Trim(Sess0.screen.getString(7 + wiersz, kolumna, 8))
End Function

