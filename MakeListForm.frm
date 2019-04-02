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
    
    
    ThisWorkbook.Sheets("register").Range("makelistregion").Value = Left(CStr(MakeListForm.ComboBox1.Value), 3)
    
    Application.EnableEvents = False
    
    If ThisWorkbook.Sheets("input").FilterMode = True Then
        ThisWorkbook.Sheets("input").ShowAllData
    End If
    
    inner_clearlist
    
    Dim arr() As String
    Dim plt_arr() As String
    Dim start As Range
    
    Dim m As MGO
    Set m = New MGO
    
    Dim pop As MS9POP00
    Set pop = m.pMS9POP00
    
    
    Set start = ThisWorkbook.Sheets("input").Range("a2")
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
    Set rng = ThisWorkbook.ActiveSheet.Range("a2")
    
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
    
    'update - Paulina 28-07-2017
    MsgBox "Check Results - Make List Completed!"
End Sub

Private Sub TextBoxDOH1_Change()
    If Len(Me.TextBoxDOH1.Text) > 3 Then
        MsgBox "Only 3 digits max can be provided in this box!"
        Me.TextBoxDOH1.Text = Left(Me.TextBoxDOH1.Text, 3)
    End If
End Sub

Private Sub TextBoxDOH2_Change()
    If Len(Me.TextBoxDOH2.Text) > 3 Then
        MsgBox "Only 3 digits max can be provided in this box!"
        Me.TextBoxDOH2.Text = Left(Me.TextBoxDOH2.Text, 3)
    End If
End Sub

Private Sub TextBoxDUNS_Change()
    If Len(Me.TextBoxDUNS.Text) > 9 Then
        MsgBox "DUNS requires 9 digits!"
        Me.TextBoxDUNS.Text = Left(Me.TextBoxDUNS.Text, 9)
    End If
    
    If Me.TextBoxDS.Text <> "8" Then
        MsgBox "I see that you want to make list by DUNS, so I put DS = 8 for you..."
        Me.TextBoxDS.Text = 8
    End If
    
    If Me.TextBoxDUNS.Text = "" Then
        If Me.TextBoxDS.Text = "8" Then
            MsgBox "I am removing 8 from DS"
            Me.TextBoxDS.Text = ""
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    ComboBox1.AddItem "GME - for Europe"
    ComboBox1.AddItem "MGO - for NA"
    ComboBox1.Value = "GME - for Europe"
End Sub
