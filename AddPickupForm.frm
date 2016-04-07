VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddPickupForm 
   Caption         =   "Add Pickup"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4095
   OleObjectBlob   =   "AddPickupForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddPickupForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SubmitAddPickup_Click()
    AddPickupForm.hide
    
    Dim r As Range
    Dim temp() As String
    Dim how_many_lines As Integer
    Dim act_txt As String
    Set r = ActiveCell
    
        If (CStr(Cells(5, r.Column)) = CStr(Sheets("register").Range("trans"))) Or (CStr(Cells(r.row, 8)) = CStr(Sheets("register").Range("C18"))) Then
            If r.Comment Is Nothing Then
            
                
                r.AddComment "DeliveryDate: " & Cells(3, r.Column - 1) & " " & Left(Cells(4, r.Column - 1), 10) & Chr(10) & _
                            "DeliveryTime: " & "00:00" & Chr(10) & _
                            AddPickupForm.Label1 & AddPickupForm.NameTextBox & Chr(10) & _
                            AddPickupForm.Label4 & AddPickupForm.PickupDateDTPicker.Value & Chr(10) & _
                            AddPickupForm.Label3 & AddPickupForm.QtyTextBox & Chr(10) & _
                            AddPickupForm.Label2 & AddPickupForm.RouteTextBox & Chr(10) & _
                            "-----------------------------------------------------"
                            
                            
                r = r + AddPickupForm.QtyTextBox
            Else
                act_txt = r.Comment.Text
                r.Comment.Delete
                
                r.AddComment "DeliveryDate: " & Cells(3, r.Column - 1) & " " & Left(Cells(4, r.Column - 1), 10) & Chr(10) & _
                            "DeliveryTime: " & "00:00" & Chr(10) & _
                            AddPickupForm.Label1 & AddPickupForm.NameTextBox & Chr(10) & _
                            AddPickupForm.Label4 & AddPickupForm.PickupDateDTPicker.Value & Chr(10) & _
                            AddPickupForm.Label3 & AddPickupForm.QtyTextBox & Chr(10) & _
                            AddPickupForm.Label2 & AddPickupForm.RouteTextBox & Chr(10) & _
                            "-----------------------------------------------------" & Chr(10) & _
                            act_txt
                            
                r = r + AddPickupForm.QtyTextBox
                
            End If
        Else
            MsgBox "You can't put PUS/ASN here!"
            Exit Sub
        End If
        
        r.Comment.Shape.Width = 200
        temp = Split(r.Comment.Text, Chr(10))
        how_many_lines = 0
        For x = LBound(temp) To UBound(temp)
            how_many_lines = how_many_lines + 1
        Next x
        
        r.Comment.Shape.Height = 12 * how_many_lines
        
End Sub
