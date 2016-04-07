Attribute VB_Name = "AddManualPUSModule"
Public Sub add_pickup(ictrl As IRibbonControl)
    
    AddPickupForm.PickupDateDTPicker = CStr(Format(Now, "yyyy-mm-dd"))
    AddPickupForm.show
End Sub
