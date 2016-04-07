VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MakeListForm 
   Caption         =   "Make list"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3825
   OleObjectBlob   =   "MakeListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MakeListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SubmitMakeListButton_Click()
    hide
    MakeListStatusForm.show vbModeless
    
    
    Sheets("register").Range("makelistregion").Value = Left(CStr(MakeListForm.ComboBox1.Value), 3)
    
    Application.EnableEvents = False
    
    If Sheets("input").FilterMode = True Then
        Sheets("input").ShowAllData
    End If
    
    Sheets("input").Range("a2:k1048576").Clear
    
    Dim arr() As String
    Dim plt_arr() As String
    Dim start As Range
    
    Dim m As MGO
    Set m = New MGO
    
    Dim pop As MS9POP00
    Set pop = m.pMS9POP00
    
    
    Set start = Sheets("input").Range("a2")
    arr = Split(MakeListForm.TextBoxFU, " ")
    plt_arr = Split(MakeListForm.TextBoxPLT, " ")
    
    
    ' arr reprezentuje tablice F/U
    If UBound(arr) <> -1 Then
        For i = LBound(arr) To UBound(arr)
            If MakeListForm.TextBoxPLT.Text = "" Then
                start = makelistaftershow(m, pop, start, CStr(arr(i)), CStr(TextBoxA.Text))
            Else
                If UBound(plt_arr) <> -1 Then
                    For x = LBound(plt_arr) To UBound(plt_arr)
                        start = makelistaftershow(m, pop, start, CStr(arr(i)), CStr(TextBoxA.Text), CStr(plt_arr(x)))
                    Next x
                Else
                    start = makelistaftershow(m, pop, start, CStr(arr(i)), CStr(TextBoxA.Text), "")
                End If
            End If
        Next i
    Else
        If MakeListForm.TextBoxPLT.Text = "" Then
            start = makelistaftershow(m, pop, start, "", CStr(TextBoxA.Text))
        Else
            If UBound(plt_arr) <> -1 Then
                For x = LBound(plt_arr) To UBound(plt_arr)
                    start = makelistaftershow(m, pop, start, "", CStr(TextBoxA.Text), CStr(plt_arr(x)))
                Next x
            Else
                start = makelistaftershow(m, pop, start, "", CStr(TextBoxA.Text), "")
            End If
        End If
    End If
    
    Dim rng As Range
    Dim tmp As Range
    Set rng = Range("a2")
    
    Do
        If Trim(rng) = "null" Then
            Set tmp = rng.Offset(1, 0)
            Rows(CStr(rng.row) & ":" & CStr(rng.row)).Delete Shift:=xlUp
            Set rng = tmp
        Else
            Set rng = rng.Offset(1, 0)
        End If
    Loop While rng <> ""
    
    MakeListStatusForm.LabelStatus = ""
    MakeListStatusForm.hide
    
    
    Application.EnableEvents = True
End Sub

Private Sub UserForm_Initialize()
    ComboBox1.AddItem "GME - for Europe"
    ComboBox1.AddItem "MGO - for NA"
    ComboBox1.Value = "GME - for Europe"
End Sub
