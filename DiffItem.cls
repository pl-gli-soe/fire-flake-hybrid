VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiffItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public pn As String
Public plt As String
Public Name As String
Public pickup_date As Date
Public delivery_date As Date
Public Qty As Long
Public sum_of_regular_transit As Long
Public sum_of_qty_in_parent_cell As Long
Public istransit As Boolean
Public regular As Boolean
Private exact_value As Long

Public TRLR As String
Public st As String
Public route As String
Public fst_runout As String

Public Property Get differenceQty() As Long
    differenceQty = sum_of_qty_in_parent_cell - Qty
End Property

Public Property Get manualAdjustment() As Long
    manualAdjustment = sum_of_qty_in_parent_cell - sum_of_regular_transit
End Property



Public Sub set_exact_value()

    If istransit Then
        If manualAdjustment = 0 Then
            exact_value = Qty
        Else
    
                exact_value = manualAdjustment
        End If
    End If
End Sub

Public Property Get getExactValue() As Long
    getExactValue = CLng(exact_value)
End Property
