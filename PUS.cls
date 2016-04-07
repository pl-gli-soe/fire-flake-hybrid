VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ITransit

Private route As String
Private iNAME As String
Private Qty As Long
Private delivery_date As Date
Private delivery_time As Date
Private pickup_date As Date
Private pickup_time As Date
Private kontener As String
Private st As String


Private not_yet_received As Boolean
Private out_of_scope As Boolean
Private on_pop As Boolean
Private tt_less_than_24h As Boolean



Private Sub Class_Initialize()
    not_yet_received = True
    out_of_scope = False
End Sub

Private Sub Class_Terminate()
    
End Sub

Private Property Let ITransit_mDeliveryDate(RHS As Date)
    delivery_date = RHS
End Property

Private Property Get ITransit_mDeliveryDate() As Date
    ITransit_mDeliveryDate = delivery_date
End Property

Private Property Let ITransit_mDeliveryTime(RHS As Date)
    delivery_time = RHS
End Property

Private Property Get ITransit_mDeliveryTime() As Date
    ITransit_mDeliveryTime = delivery_time
End Property

Private Property Let ITransit_mName(RHS As String)
    iNAME = CStr(RHS)
End Property

Private Property Get ITransit_mName() As String
    ITransit_mName = iNAME
End Property

Private Property Let ITransit_mNotYetReceived(RHS As Boolean)
    not_yet_received = RHS
End Property

Private Property Get ITransit_mNotYetReceived() As Boolean
    ITransit_mNotYetReceived = not_yet_received
End Property

Private Property Let ITransit_mOnPop(RHS As Boolean)
    on_pop = RHS
End Property

Private Property Get ITransit_mOnPop() As Boolean
    ITransit_mOnPop = on_pop
End Property

Private Property Let ITransit_mOutOfScope(RHS As Boolean)
    out_of_scope = RHS
End Property

Private Property Get ITransit_mOutOfScope() As Boolean
    ITransit_mOutOfScope = out_of_scope
End Property

Private Property Let ITransit_mPickupDate(RHS As Date)
    pickup_date = RHS
End Property

Private Property Get ITransit_mPickupDate() As Date
    ITransit_mPickupDate = pickup_date
End Property

Private Property Let ITransit_mPickupTime(RHS As Date)
    pickup_time = RHS
End Property

Private Property Get ITransit_mPickupTime() As Date
    ITransit_mPickupTime = pickup_time
End Property

Private Property Let ITransit_mQty(RHS As Long)
    Qty = RHS
End Property

Private Property Get ITransit_mQty() As Long
    ITransit_mQty = Qty
End Property

Private Property Let ITransit_mRoute(RHS As String)
    route = RHS
End Property

Private Property Get ITransit_mRoute() As String
    ITransit_mRoute = route
End Property

Private Property Let ITransit_mST(RHS As String)
    st = CStr(RHS)
End Property

Private Property Get ITransit_mST() As String
    ITransit_mST = CStr(st)
End Property

Private Property Let ITransit_mTRLR(RHS As String)
    kontener = CStr(RHS)
End Property

Private Property Get ITransit_mTRLR() As String
    ITransit_mTRLR = CStr(kontener)
End Property

Private Property Let ITransit_mTTLessThan24h(RHS As Boolean)
    tt_less_than_24h = RHS
End Property

Private Property Get ITransit_mTTLessThan24h() As Boolean
    ITransit_mTTLessThan24h = tt_less_than_24h
End Property


