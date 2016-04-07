VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemHourly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IItem

Private h_req As Long
Private h_ebal As String
Private h_trans As Collection
Private item_h_trans As ITransit
Private h_time As Date
'Private rhr_date As Date
'Private rhr_time As Date




Private Sub Class_Initialize()
    Set h_trans = New Collection
End Sub

Private Sub Class_Terminate()
    Set h_trans = Nothing
End Sub


Private Sub IItem_addTransToCollection(arg As ITransit)
    h_trans.Add arg
End Sub

'Private Property Let IItem_hDate(RHS As Date)
'   rhr_date = CDate(RHS)
'End Property
'
'Private Property Get IItem_hDate() As Date
'    IItem_hDate = CDate(rhr_date)
'End Property
'
'Private Property Let IItem_hTime(RHS As Date)
'   rhr_time = CDate(RHS)
'End Property
'
'Private Property Get IItem_hTime() As Date
'    IItem_hTime = CDate(rhr_time)
'End Property

Private Property Let IItem_mDate(RHS As Date)
    h_time = RHS
End Property

Private Property Get IItem_mDate() As Date
    IItem_mDate = h_time
End Property

Private Property Get IItem_mEbal() As String
    IItem_mEbal = h_ebal
End Property

Private Property Let IItem_mEbal(RHS As String)
    h_ebal = RHS
End Property

Private Property Let IItem_mReq(RHS As Long)
    h_req = RHS
End Property

Private Property Get IItem_mReq() As Long
    IItem_mReq = h_req
End Property

Private Property Let IItem_mTrans(RHS As ITransit)
    Set item_h_trans = RHS
End Property

Private Property Get IItem_mTrans() As ITransit
    Set IItem_mTrans = item_h_trans
End Property

Private Property Get IItem_mTransCollection() As Collection
    Set IItem_mTransCollection = h_trans
End Property
