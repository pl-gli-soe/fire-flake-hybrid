VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIteracja"
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
' ClassModule IIteracja
' ========================
'
' interface 4 daily & hourly


Public Property Get Partnum() As String
End Property

Public Property Let Partnum(arg As String)
End Property

Public Property Get PlantName() As String
End Property

Public Property Let PlantName(arg As String)
End Property

Public Property Get CBAL() As String
End Property

Public Property Let CBAL(arg As String)
End Property

Public Property Get PartName() As String
End Property

Public Property Let PartName(arg As String)
End Property

Public Property Get SUPPLIER() As String
End Property

Public Property Let SUPPLIER(arg As String)
End Property

Public Property Get DUNS() As String
End Property

Public Property Let DUNS(arg As String)
End Property

Public Property Get Deck() As String
End Property

Public Property Let Deck(arg As String)
End Property

Public Property Get DOH() As String
End Property

Public Property Let DOH(arg As String)
End Property

Public Property Get FirstRunout() As String
End Property

Public Property Let FirstRunout(arg As String)
End Property

Public Property Get Qhold() As String
End Property

Public Property Let Qhold(arg As String)
End Property

Public Property Get MISC() As String
End Property

Public Property Let MISC(arg As String)
End Property

Public Property Get NotyetReceived() As Range
End Property

Public Property Let NotyetReceived(arg As Range)
End Property

Public Property Get StdPack() As String
End Property

Public Property Let StdPack(arg As String)
End Property

Public Property Get Cmnt1() As String
End Property

Public Property Let Cmnt1(arg As String)
End Property

Public Property Get Cmnt2() As String
End Property

Public Property Let Cmnt2(arg As String)
End Property

Public Property Get IterationRow() As Integer
End Property

Public Property Let IterationRow(arg As Integer)
End Property

Public Property Get ItemCollection() As Collection
End Property

Public Property Let ItemCollection(arg As Collection)
End Property

Public Property Get ProdDayBegin() As Date
End Property

Public Property Let ProdDayBegin(arg As Date)
End Property

Public Sub AddItem(arg As IItem)
End Sub

Public Sub PrepareItemCollection(i_mgo As MGO, item As IItem, limit As Date, limit_del As Date, Optional hourly_extension_day As Date, Optional ByRef sh As StatusHandler, Optional ByRef instance_tokenizator As Tokenizator)
End Sub

Public Sub DownloadGeneralData(i_mgo As MGO, ByRef sh As StatusHandler, limit As Date, limit_del As Date)
End Sub

Public Property Get NotYetReceivedCollection() As Collection

End Property

Public Property Get Bank() As Long

End Property

Public Property Let Bank(arg As Long)

End Property

Public Property Get SheetDef() As String

End Property

Public Property Let SheetDef(arg As String)

End Property



