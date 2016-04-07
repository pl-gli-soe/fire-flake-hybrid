VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemWeekly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' mateusz.milewski@gm.com
' FireFlakeHybrid



Implements IItem

Private weekly_req As Long
Private weekly_trans As Collection
Private item_trans As ITransit
Private weekly_ebal As String
' wstepnie zalozylem poniedzialek jednak z powodu problemow
' z implementacja wczesniejsza dopasowana przede wszystkim do
' godzinowki musze wybrac jako date current_date niedziele z racji tego
' ffh zawsze dopiera pickupy posymistycznie w stosunku do przedzialow dat
' a tak sie nie fajnie sklada ze jesli poniedzialek jest poczatkiem i zawsze bedzie przegdywac
' przyklad
' Cw 12 --------------------------------- Cw 13
' ------------------ASN/PUS--------------------
' delivery date wspasowane jeszcze w cw12 jednak z racji implementacji dobierze z przedzialu gorna granice
Private current_date As Date





Private Sub Class_Initialize()
    Set weekly_trans = New Collection
End Sub

Private Sub Class_Terminate()
    Set weekly_trans = Nothing
End Sub

Private Sub IItem_addTransToCollection(arg As ITransit)
    weekly_trans.Add arg
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
    IItem_mEbal = weekly_ebal
End Property

Private Property Let IItem_mEbal(RHS As String)
    weekly_ebal = RHS
End Property

Private Property Let IItem_mReq(RHS As Long)
    weekly_req = RHS
End Property

Private Property Get IItem_mReq() As Long
    IItem_mReq = weekly_req
End Property

Private Property Let IItem_mTrans(RHS As ITransit)
    Set item_trans = RHS
End Property

Private Property Get IItem_mTrans() As ITransit
    Set IItem_mTrans = item_trans
End Property


Private Property Get IItem_mTransCollection() As Collection
    Set IItem_mTransCollection = weekly_trans
End Property

