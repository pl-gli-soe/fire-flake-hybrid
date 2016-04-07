VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StatusHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public szerokosc_paska As Double
' bardzo wazne zeby nie przekroczyc zakresu mozliwosci tych dwoch zmiennych
Public act As Double
Public max As Double
Public dx As Double
Public jeden_procent As Double
' 360/255 <-  pod gradientowa zmiane koloru
Public grad_color As Double
Private increment_decrement_flag As Boolean

Public Sub init_statusbar(m_max, Optional m_min)
    max = m_max
    If Not IsMissing(m_min) Then
        act = m_min
    End If
    dx = CDbl(szerokosc_paska) / CDbl(max)
    
End Sub

Public Sub progress_increase()
    If act < max Then
        act = act + 1
        
        
        StatusForm.PasekPostepu.Width = StatusForm.PasekPostepu.Width + dx
        If StatusForm.PasekPostepu.Width >= 360 Then
            StatusForm.PasekPostepu.Width = 360
        End If
        DoEvents
        StatusForm.FrameProgress.Caption = CStr(Int((CDbl(act) / CDbl(max)) * 100)) & " %"
        StatusForm.Repaint
    End If
End Sub

Public Sub show()
    StatusForm.show vbModeless
End Sub

Public Sub hide()
    StatusForm.hide
End Sub



Private Sub Class_Initialize()
    szerokosc_paska = 360
    grad_color = 155 ' od 155 do 255 kolor 100 razy zmieniany
    jeden_procent = CDbl(szerokosc_paska) / CDbl(100)
    act = 0
    StatusForm.PasekPostepu.Width = 0
    StatusForm.PasekPostepu.BackColor = RGB(255, 0, 0)
    
End Sub

Private Sub Class_Terminate()

End Sub
