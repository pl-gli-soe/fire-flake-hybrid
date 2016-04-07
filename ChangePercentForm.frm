VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangePercentForm 
   Caption         =   "Change Percent"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3750
   OleObjectBlob   =   "ChangePercentForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChangePercentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChangePercentButton_Click()
    hide
    Sheets("register").Range("pinkOnHourly") = ChangePercentForm.ChangePercentTextBox
    przelicz_arkusz ActiveSheet, ActiveCell, True
End Sub
