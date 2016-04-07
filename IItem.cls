VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IItem"
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
' ClassModule IItem
' ========================
'
' interface for PUS and ASN

Public Property Let mDate(arg As Date)
End Property

Public Property Get mDate() As Date
End Property

'Public Property Let hDate(arg As Date)
'End Property
'
'Public Property Get hDate() As Date
'End Property
'
'Public Property Let hTime(arg As Date)
'End Property
'
'Public Property Get hTime() As Date
'End Property

Public Property Let mEbal(arg As String)
End Property

Public Property Get mEbal() As String
End Property

Public Property Let mReq(arg As Long)
End Property

Public Property Get mReq() As Long
End Property

Public Property Let mTrans(req As ITransit)
End Property

Public Property Get mTrans() As ITransit
End Property

Public Sub addTransToCollection(arg As ITransit)
End Sub

Public Property Get mTransCollection() As Collection
End Property

