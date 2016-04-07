VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChartOnPart 
   Caption         =   "Chart for active part"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9915
   OleObjectBlob   =   "ChartOnPart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChartOnPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ChangeScaleButton_Click()
    Dim tmp_chart As ChartObject
    Set sh = Sheets("chart register")
    Set tmp_chart = sh.ChartObjects(1)
    
    
    With tmp_chart.Chart
    
        With .Axes(xlValue)
            .MaximumScale = Int(ChartOnPart.MaxValueTextBox.Value)
            .MinimumScale = Int(ChartOnPart.MinValueTextBox.Value)
        End With
        
    End With
    
    With tmp_chart
        .Width = 500
        .Height = 250
    End With
    
    tmp_chart.Chart.Refresh
    fname = ThisWorkbook.Path & Application.PathSeparator & "temp.gif"
    tmp_chart.Chart.Export Filename:=fname, filtername:="GIF"
    ChartOnPart.ChartImage.Picture = LoadPicture(fname)
    
End Sub

Private Sub ExportChartButton_Click()
    MsgBox "Implementation in progress"
End Sub
