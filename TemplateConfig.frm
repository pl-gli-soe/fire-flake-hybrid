VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateConfig 
   Caption         =   "Template Config"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3765
   OleObjectBlob   =   "TemplateConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TemplateConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CreateTempBtn_Click()
    TemplateConfig.hide
    
    
    
    Dim t As FFTemplate
    Set t = New FFTemplate
    t.create_template New ItemDaily, CDate(TemplateConfig.StartDTPicker), CDate(TemplateConfig.EndDTPicker)
End Sub
