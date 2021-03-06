VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASN"
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
' ClassModule ASN
' ===============
'
' class for transit electronic document
' most of are O'Sea vessels
' have interaface from ITransit only as well as PUS type

Implements ITransit

Private route As String
Private iNAME As String
Private QTY As Long
Private DELIVERY_DATE As Date
Private DELIVERY_TIME As Date
Private SDATE As Date
Private kontener As String
Private st As String


Private not_yet_received As Boolean
Private out_of_scope As Boolean
Private on_pop As Boolean
Private on_prosl As Boolean



Private Sub Class_Initialize()
    not_yet_received = True
    out_of_scope = False
End Sub

Private Property Let ITransit_mDeliveryDate(RHS As Date)
    DELIVERY_DATE = RHS
End Property

Private Property Get ITransit_mDeliveryDate() As Date
    ITransit_mDeliveryDate = DELIVERY_DATE
End Property

Private Property Let ITransit_mDeliveryTime(RHS As Date)
    DELIVERY_TIME = RHS
End Property

Private Property Get ITransit_mDeliveryTime() As Date
    ITransit_mDeliveryTime = DELIVERY_TIME
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

Private Property Let ITransit_mOnProsl(RHS As Boolean)
    on_prosl = RHS
End Property

Private Property Get ITransit_mOnProsl() As Boolean
    ITransit_mOnProsl = on_prosl
End Property

Private Property Let ITransit_mOutOfScope(RHS As Boolean)
    out_of_scope = RHS
End Property

Private Property Get ITransit_mOutOfScope() As Boolean
    ITransit_mOutOfScope = out_of_scope
End Property

Private Property Let ITransit_mPickupDate(RHS As Date)
    SDATE = RHS
End Property

Private Property Get ITransit_mPickupDate() As Date
    ITransit_mPickupDate = SDATE
End Property

Private Property Let ITransit_mPickupTime(RHS As Date)
    
End Property

Private Property Get ITransit_mPickupTime() As Date
    ITransit_mPickupTime = CDate("00:00")
End Property

Private Property Let ITransit_mQty(RHS As Long)
    QTY = RHS
End Property

Private Property Get ITransit_mQty() As Long
    ITransit_mQty = QTY
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
    ITransit_mTTLessThan24h = RHS
End Property

Private Property Get ITransit_mTTLessThan24h() As Boolean
    ITransit_mTTLessThan24h = False
End Property
