VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' mateusz.milewski@gm.com
' FireFlakeHybrid



Implements IItem

Private daily_req As Long
Private daily_trans As Collection
Private item_trans As ITransit
Private daily_ebal As String
Private current_date As Date





Private Sub Class_Initialize()
    Set daily_trans = New Collection
End Sub

Private Sub Class_Terminate()
    Set daily_trans = Nothing
End Sub

Private Sub IItem_addTransToCollection(arg As ITransit)
    daily_trans.Add arg
End Sub

'Private Property Let IItem_hDate(RHS As Date)
'
'End Property
'
'Private Property Get IItem_hDate() As Date
'
'End Property
'
'Private Property Let IItem_hTime(RHS As Date)
'
'End Property
'
'Private Property Get IItem_hTime() As Date
'
'End Property

Private Property Let IItem_mDate(RHS As Date)
    current_date = RHS
End Property

Private Property Get IItem_mDate() As Date
    IItem_mDate = current_date
End Property

Private Property Get IItem_mEbal() As String
    IItem_mEbal = daily_ebal
End Property

Private Property Let IItem_mEbal(RHS As String)
    daily_ebal = RHS
End Property

Private Property Let IItem_mReq(RHS As Long)
    daily_req = RHS
End Property

Private Property Get IItem_mReq() As Long
    IItem_mReq = daily_req
End Property

Private Property Let IItem_mTrans(RHS As ITransit)
    Set item_trans = RHS
End Property

Private Property Get IItem_mTrans() As ITransit
    Set IItem_mTrans = item_trans
End Property


Private Property Get IItem_mTransCollection() As Collection
    Set IItem_mTransCollection = daily_trans
End Property
