Attribute VB_Name = "CloudModule"
Public Sub init_cloud(ictrl As IRibbonControl)
    Set cloud_item = New Cloud
    cloud_item.create_cloud
    Dim Qty As Long
    Qty = CLng(InputBox("Capacity limitation: ", "Capacity limit"))
    cloud_item.set_capacity_limit Qty
    cloud_item.config_limit
End Sub

Public Sub catch_cloud(ictrl As IRibbonControl)
    Set cloud_item = New Cloud
    cloud_item.catch_cloud
    Dim Qty As Long
    Qty = CLng(InputBox("Capacity limitation: ", "Capacity limit"))
    cloud_item.set_capacity_limit Qty
    cloud_item.config_limit
End Sub

Public Sub destroy_cloud(ictrl As IRibbonControl)
    cloud_item.delete_shape
    Set cloud_item = Nothing
End Sub
