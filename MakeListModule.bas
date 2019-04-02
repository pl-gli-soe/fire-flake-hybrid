Attribute VB_Name = "MakeListModule"
Public Sub listmaker(ic As IRibbonControl)
    ' MsgBox "make list test procedure"
    MakeListForm.show
End Sub

Public Sub clearlist(c As IRibbonControl)
    inner_clearlist
End Sub

Public Sub inner_clearlist()


    If Sheets("input").FilterMode = True Then
        Sheets("input").ShowAllData
    End If
    Sheets("input").Range("a2:l1048576").Clear
    Sheets("input").Range("a2:l1048576").ClearComments
End Sub

Public Function makelistaftershow(m As MGO, pop As MS9POP00, Optional ByRef start As Range, Optional fu As String, Optional a As String, Optional plt As String) As Range

    Dim i As Integer
    i = 0
    

    
    m.sendKeys "<Clear>"
    m.sendKeys "ms9pop00 <Enter>"
    
    
    pop.DS = MakeListForm.TextBoxDS
    If fu <> "" Then
        pop.F_U = fu
    End If
    
    If MakeListForm.TextBoxDUNS <> "" Then
        pop.DUNS = MakeListForm.TextBoxDUNS
    End If
    
    If a <> "" Then
        pop.a = a
    End If
    pop.firstDOH = MakeListForm.TextBoxDOH1
    pop.secDOH = MakeListForm.TextBoxDOH2
    If plt = "" Then
        m.putString ThisWorkbook.Sheets("register").Range("makelistregion").Value, 3, 5
    Else
        m.putString CStr(plt), 3, 13
    End If
    m.sendKeys "<Enter>"
    
    Do
        
        
        If Trim(pop.plt) <> "" Then
            start = pop.plt
            start.Offset(0, 1) = pop.pn
            If pop.transQTY(0) <> "" Then
                cmnt_string = "First PUS/ASN on MS9POP00: " & Chr(10) & _
                "QTY: " & CStr(pop.transQTY(0)) & Chr(10) & _
                "CONTAINER: " & CStr(pop.transCONT(0)) & Chr(10) & _
                "SDATE: " & CStr(pop.transSDATE(0)) & Chr(10) & _
                "EDA: " & CStr(pop.transEDA(0)) & Chr(10) & _
                "ETA: " & CStr(pop.transETA(0)) & Chr(10) & _
                "CMNT: " & CStr(pop.transCMNT(0)) & Chr(10) & _
                "ETA: " & CStr(pop.transETA(0)) & Chr(10) & _
                "DUNS: " & CStr(pop.transDUNS(0)) & Chr(10) & _
                "ROUTE: " & CStr(pop.transROUTE(0)) & Chr(10)
                
                start.Offset(0, 1).AddComment CStr(cmnt_string)
                
                start.Offset(0, 1).Comment.Shape.Width = 200
                start.Offset(0, 1).Comment.Shape.Height = 150
                
                start.Offset(0, 1).Interior.Color = RGB(200, 200, 200)
            Else
                start.Offset(0, 1).AddComment "no active PUS/ASN on MS9POP00"
            End If
            start.Offset(0, 2) = pop.firstDOH
            start.Offset(0, 3) = pop.SUPPLIER
            start.Offset(0, 4) = pop.DUNS
            start.Offset(0, 5) = pop.F_U
            start.Offset(0, 6) = pop.a
            start.Offset(0, 7) = pop.COUNT
            start.Offset(0, 8) = pop.O
            start.Offset(0, 11) = pop.PCS_TO_GO
            
            i = 0
        Else
            start = "null"
            start.Offset(0, 1) = "null"
        End If
        
        MakeListStatusForm.LabelStatus = "PN: " & CStr(start.Value) & ", PLT: " & CStr(start.Offset(0, 1).Value)
        
        m.sendKeys "<pf8>"
        
        If m.getString(23, 2, 5) = "I4028" Then
            i = i + 1
        End If
        
        If m.getString(23, 2, 5) = "I4265" Then
            Set start = start.Offset(1, 0)
            Exit Do
        End If
        
        If i > MAKE_LIST_TIMES_F8 Then
            Set start = start.Offset(1, 0)
            Exit Do
        End If
        
        Set start = start.Offset(1, 0)
        
    Loop While True
    
    
    Set makelistaftershow = start
    
    
End Function

Private Function check_is_the_end(ByRef i As Integer)

    If i < 5 Then
        
    Else
        check_is_the_end = False
    End If
End Function
