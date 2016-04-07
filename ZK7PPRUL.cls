VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZK7PPRUL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sess0 As Object
' zmienna przechowujaca reszte danych z zk7pdrqm
Private ULOC As Collection

Public Property Let sesja(ByRef arg As Object)
    Set Sess0 = arg
End Property

Public Property Get sesja() As Object
    Set sesja = Sess0
End Property


Public Property Let p_uloc(ByRef arg As Object)
    Set ULOC = arg
End Property

Public Property Get p_uloc() As Object
    Set p_uloc = ULOC
End Property

Private Sub Class_Initialize()
    Set ULOC = New Collection
End Sub

Private Sub Class_Terminate()
    Set ULOC = Nothing
End Sub




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
    KANBAN = Trim(Sess0.screen.getString(4, 26, 4))
End Property

Public Property Let KANBAN(arg As String)
    Sess0.screen.putString CStr(arg), 4, 26
End Property

Public Sub fill_uloc_collection(i_mgo As MGO)

    Do
        For x = 8 To 21
            If (Trim(Sess0.screen.getString(x, 6, 4)) <> "MISC") And (Sess0.screen.getString(x, 6, 4) <> "    ") Then
                ULOC.Add Trim(Sess0.screen.getString(x, 6, 8))
            End If
        Next x
        
        If Sess0.screen.getString(22, 2, 5) = "R6086" Then
            Exit Do
        End If
        
        If Sess0.screen.getString(22, 2, 5) = "R6017" Then
            Exit Do
        End If
        
        If Sess0.screen.getString(22, 2, 5) = "R6102" Then
            i_mgo.sendKeys "<pf8>"
            i_mgo.waitForMgo
        End If
        
    Loop While True
    
    ' MsgBox ULOC.COUNT
End Sub





