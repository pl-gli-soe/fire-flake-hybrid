VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cloud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private actual_sum As Range
Private capacity_limit As Range
Private x_pos As Range
Private y_pos As Range
Private dx As Range
Private dy As Range
Private shape_range As ShapeRange
Private tmp As ShapeRange
Private v As Variant

Private Sub Class_Initialize()

    Set v = Nothing
    Set shape_range = Nothing
    
    Set actual_sum = Sheets("register").Range("actSumCloud")
    actual_sum = 0
    Set capacity_limit = Sheets("register").Range("capcityLimit")
    capacity_limit = -1
    
    Set x_pos = Sheets("register").Range("xpos")
    x_pos = 10
    Set y_pos = Sheets("register").Range("ypos")
    y_pos = 10
    
    Set dx = Sheets("register").Range("dx")
    dx = 0
    Set dy = Sheets("register").Range("dy")
    dy = 0
    
    
End Sub

Private Sub Class_Terminate()
    Set v = Nothing
    Set shape_range = Nothing
End Sub

Public Sub config_limit()
shape_range.TextFrame2.TextRange.Font.size = 8
    shape_range.TextFrame2.TextRange.Characters.Text = CStr(actual_sum) & Chr(10) & "Capacity: " & CStr(capacity_limit)
    shape_range.Width = 70
    shape_range.Height = 50
End Sub

Public Sub create_cloud()
    Set v = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, pos_x, pos_y, 70, 30)
    Set shape_range = ActiveSheet.Shapes.Range(Array(v.Name))
End Sub

Public Sub catch_cloud()
    Set v = Selection
    Set shape_range = ActiveSheet.Shapes.Range(Array(v.Name))
    Dim tmp As Variant
    tmp = Split(shape_range.TextFrame2.TextRange.Characters.Text, Chr(10))
    actual_sum = tmp(0)
    x_pos = shape_range.Left
    y_pos = shape_range.Top
End Sub

Public Sub delete_shape()
    shape_range.Delete
End Sub

Public Sub set_capacity_limit(arg As Long)
    capacity_limit = arg
End Sub

Public Sub move_on_selection_change(t As Range)


    If IsNumeric(t) Then
        If Not shape_range Is Nothing Then
            dx = x_pos
            dy = y_pos
            x_pos = 0
            y_pos = 0
            
            ' actual_sum = actual_sum + CLng(t)
            Dim r As Range
            For ix = 1 To t.Column
                    Set r = Cells(1, ix)
                    x_pos = x_pos + r.Width
            Next ix
            
            For iy = 1 To t.row
                Set r = Cells(iy, 1)
                y_pos = y_pos + r.Height
            Next iy
            
            y_pos = y_pos + 100
            ' ActiveSheet.Shapes.Range(Array(v.name)).Select
            shape_range.Select
    
            ' Selection.ShapeRange.IncrementLeft x_pos - dx
            shape_range.IncrementLeft x_pos - dx
            
            ' Selection.ShapeRange.IncrementTop y_pos - dy
            shape_range.IncrementTop y_pos - dy
            
            t.Select
        End If
    End If
End Sub

' chyba zrobie to tak jak mam do dyspozycji dane to znaczy jesli
' zmieniona zostala wartosc to podkrels na niebiekso
' tam implementacja dziala na zasadzie porownania poprzedniej wartosci
' z aktualnym rangem wiec takie dwa argumenty dam wlasnie
Public Sub add_the_difference_to_textbox(prev_value As Long, r As Range)
    If IsNumeric(r) Then
        actual_sum = actual_sum + CLng(r.Value - prev_value)
        shape_range.TextFrame2.TextRange.Characters.Text = CStr(actual_sum) & Chr(10) & "Capacity: " & CStr(capacity_limit)
        
        If capacity_limit > 0 Then
            If actual_sum > capacity_limit Then
                shape_range.TextFrame2.TextRange.Font.Bold = msoCTrue
                shape_range.TextFrame2.TextRange.Characters.Text = _
                shape_range.TextFrame2.TextRange.Characters.Text & Chr(10) & _
                "You are over the capacity!"
                
                shape_range.Width = 70
                shape_range.Height = 50
            Else
                shape_range.TextFrame2.TextRange.Font.Bold = msoFalse
                shape_range.Width = 70
                shape_range.Height = 50
            End If
        End If
    End If
End Sub


