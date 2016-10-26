VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FireFlakeHybrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' code written by Mateusz Milewski
'
' mateusz.milewski@ gm.com
' FireFlakeHybrid
'
'
' ClassModule DiffReport
' ======================
' the main class with all variable instances
' the complex algorithm- make it as component
' and You'll able do to everything with this code :)


Private instance_item As IItem
Private instances_item As Collection

Private instance_iteration As IIteracja
Private instances_iteration As Collection

Private instance_diff_report As DiffReport
Private instance_make_list As MakeList
Private instance_shipping_plan As ShippingPlan
Private instance_mgo As MGO
Private instance_matrix_handler As MatrixHandler
Private instance_layout As ILayout
Private instance_status_handler As StatusHandler
Private instance_filter As Filtracja

' instancja jest potrzebna na calym scopie
Private instance_tokenizator As Tokenizator


Private catch_err As CatchError
Private date_limit As Date
Private date_limit_delivery As Date
Private ip_zero As Boolean
Private first_item As IItem

' set, let dla pol - limit pickup & delivery
Public Property Get p_limit() As Date
    p_limit = date_limit
End Property

Public Property Let p_limit(arg As Date)

    date_limit = arg
End Property


Public Property Get ip() As Date
    ip = ip_zero
End Property


Public Property Let ip(arg As Date)

    ip_zero = arg
End Property




Public Property Let p_limit_delivery(arg As Date)
    date_limit_delivery = arg
End Property

Public Property Get p_limit_delivery() As Date
    p_limit = date_limit_delivery
End Property

Public Property Get p_instance_iteration() As IIteracja
    Set p_instance_iteration = instance_iteration
End Property

Public Property Let p_instance_iteration(arg As IIteracja)
    Set instance_iteration = arg
End Property

Public Property Get p_item() As IItem
    Set p_item = instance_item
End Property

Public Property Let p_item(i As IItem)
    Set instance_item = i
End Property

Public Property Get p_diff_report() As DiffReport
    Set p_item = instance_diff_report
End Property

Public Property Let p_diff_report(i As DiffReport)
    Set instance_diff_report = i
End Property

Public Property Get p_make_list() As MakeList
    Set p_make_list = instance_make_list
End Property

Public Property Get p_shipping_plan() As ShippingPlan
    Set p_shipping_plan = instance_shipping_plan
End Property

Public Property Let p_shipping_plan(i As ShippingPlan)
    Set instance_shipping_plan = i
End Property

Public Property Get p_mgo() As MGO
    Set p_mgo = instance_mgo
End Property

Public Property Let p_mgo(i As MGO)
    Set instance_mgo = i
End Property

Public Property Get p_matrix() As MatrixHandler
    Set p_matrix = instance_matrix_handler
End Property

Public Property Let p_matrix(i As MatrixHandler)
    Set instance_matrix_handler = i
End Property

Public Property Get p_layout() As ILayout
    Set p_layout = instance_layout
End Property

Public Property Let p_layout(i As ILayout)
    Set instance_layout = i
End Property

Public Property Get p_status_handler() As StatusHandler
    p_status_handler = instance_status_handler
End Property

Public Property Let p_status_handler(i As StatusHandler)
    Set instance_status_handler = i
End Property



Private Sub Class_Initialize()

    ' Set instance_shipping_plan = New ShippingPlan
    ' Set instance_make_list = New MakeList
    ' Set instance_mgo = New MGO
    Set instances_item = New Collection
    Set instances_iteration = New Collection
    Set instance_status_handler = New StatusHandler
    Set instance_filter = New Filtracja
    Set instance_tokenizator = New Tokenizator
End Sub

Private Sub Class_Terminate()

    Set instance_iteration = Nothing
    ' Set instance_shipping_plan = Nothing
    ' Set instance_make_list = Nothing
    Set instance_diff_report = Nothing
    Set instance_item = Nothing
    ' Set instance_mgo = Nothing
    Set instance_layout = Nothing
    Set instance_status_handler = Nothing
    Set INPUT_DATA = Nothing
    Set instances_item = Nothing
    Set instances_iteration = Nothing
    Set instance_status_handler = Nothing
    Set instance_filter = Nothing
    Set instance_tokenizator = Nothing
End Sub

Public Sub create_template(item As IItem, startdate As Date, enddate As Date)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Sheets("register").Range("w_macro") = 1
    
    If Int(enddate - startdate) > 0 Then
        Sheets("register").Range("difference") = Int(enddate - startdate)
    Else
        Sheets("register").Range("difference") = 20
    End If
    ' Sheets("register").Range("itemDays") = Sheets("register").Range("difference")
    
    prepare_input_data
    Set instance_item = item
    
    If CStr(TypeName(item)) Like "*Daily*" Then
        Set instance_layout = New DailyLayout
        Set instance_item = item
        
        
    ElseIf CStr(TypeName(item)) Like "*Hourly*" Then
        Set instance_layout = New HourlyLayout
        Set instance_item = item
    End If
    Sheets("register").Range("itemDays") = Sheets("register").Range("difference")
    
    instance_layout.InitLayout True
    instance_layout.GenerujTabeleDlaListyDanychOgolnych
    instance_layout.GenerujReszteRusztownia

    
    
    Dim kol As Long
    Dim cl As Range
    Dim ef As EbalFormula
    
    If instance_tokenizator.token = "DEFAULT" Then
        For r = 6 To Int(Sheets("register").Range("lastRow"))
        
            Cells(r, 2) = INPUT_DATA.item(r - 5).part_num
            Cells(r, 3) = INPUT_DATA.item(r - 5).plt_name
            Cells(r, 15) = INPUT_DATA.item(r - 5).comment1
            Cells(r, 16) = INPUT_DATA.item(r - 5).comment2
        
            kol = Sheets("register").Range("firstColumn") + 2
        
            For d = 1 To Int(Sheets("register").Range("itemDays"))
        
                Set ef = New EbalFormula
                If kol = 19 Then
                    Cells(r, kol) = ef.fill_first_day()
                Else
                    Cells(r, kol) = ef.fill_other_days()
                End If
                Set ef = Nothing
                kol = kol + 3
            Next d
            
        Next r
    End If
    
    Sheets("register").Range("scopeObliczen") = "all"
    przelicz_parametry_arkusza ActiveSheet, ActiveCell
    przelicz_arkusz ActiveSheet, ActiveCell, True
    Sheets("register").Range("scopeObliczen") = Sheets("register").Range("scopeObliczenDefault")
    DoEvents
    Sheets("register").Range("w_macro") = 0
    Application.Calculation = xlCalculationAutomatic
    
    Columns("B:P").AutoFit
    
    If TypeName(item) Like "*Daily*" Then
        If Not ActiveSheet.AutoFilterMode Then
            Range("b5").AutoFilter
        End If
    End If
    
End Sub

Public Sub create_tear_down(item As IItem)
    ' MsgBox "assume that your report is ready :P"
    ' MsgBox CStr(TypeName(item))
    
    ' INPUT_DATA jest to zmienna globalna ale pseudo stala przyciagajaca dane z sheet- input
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Sheets("register").Range("w_macro") = 1
    
    prepare_input_data
    Set instance_item = item
    Set instance_mgo = New MGO
    
    If CStr(TypeName(item)) Like "*Daily*" Then
        Set instance_layout = New DailyLayout
        Set instance_item = item
        
        
    ElseIf CStr(TypeName(item)) Like "*Hourly*" Then
        Set instance_layout = New HourlyLayout
        Set instance_item = item
        
    ElseIf CStr(TypeName(item)) Like "*Weekly*" Then
        Set instance_layout = New WeeklyLayout
        Set instance_item = item
    End If
    
    instance_layout.InitLayout
    Sheets("register").Range("ActualQtyIntervals") = ""
    
    Dim id As InputData
    
    Dim part As Integer
    Dim parts As Integer
    Dim actual_row As Integer
    parts = Sheets("register").Range("allParts")
    
    Dim ile_progress_incr_w_petli As Integer
    ile_progress_incr_w_petli = 6
    
    
    ' tutaj sa heurystycznie dopierane wartosci na postawie ilosci wsadzonych sh w poszczegolne miejsca kodu
    ' poprzednich nie pamietam, ale ostatni mnoznik jest wynikiem dla ilosci progress_increase w implementancji tworzenia tranzytów
    ' 13 dla tworzenie tranzytow :P
    instance_status_handler.init_statusbar ((parts) * ile_progress_incr_w_petli) + 2 * parts + 3 * parts + 10 * parts
    instance_status_handler.show
    
    For part = 0 To parts - 1
    
        If CStr(TypeName(instance_item)) Like "*Daily*" Then
            Set instance_layout = New DailyLayout
            Set item = New ItemDaily
        ElseIf CStr(TypeName(instance_item)) Like "*Hourly*" Then
            Set instance_layout = New HourlyLayout
            Set item = New ItemHourly
        ElseIf CStr(TypeName(instance_item)) Like "*Weekly*" Then
            Set instance_layout = New WeeklyLayout
            Set item = New ItemWeekly
        End If
        instance_status_handler.progress_increase
    
    
        If TypeName(item) Like "*Daily*" Then
            
            ' ---------------------------------------------------
            Set instance_iteration = Nothing
            Set instance_iteration = New DailyIteration
            
            instance_iteration.IterationRow = part + 6
            
            ' ---------------------------------------------------
        ElseIf TypeName(item) Like "*Weekly*" Then
            
            ' ---------------------------------------------------
            Set instance_iteration = Nothing
            Set instance_iteration = New WeeklyIteration
            
            instance_iteration.IterationRow = part + 6
            
            ' ---------------------------------------------------
        ElseIf TypeName(item) Like "*Hourly*" Then
            
            ' ---------------------------------------------------
            Set instance_iteration = Nothing
            Set instance_iteration = New HourlyIteration
            
            instance_iteration.IterationRow = part * 7 + 2
            
            ' ---------------------------------------------------
        End If
        
        instance_status_handler.progress_increase
        
        Set id = INPUT_DATA.item(part + 1)
        instance_iteration.Partnum = id.part_num
        instance_iteration.PlantName = id.plt_name
        instance_iteration.Cmnt1 = id.comment1
        instance_iteration.Cmnt2 = id.comment2
        instance_tokenizator.token = CStr(id.comment1)
        
        instance_status_handler.progress_increase
        
        ' instance_iteration.take_first_17th_columns instance_mgo
        ' download std data from ms9pop00
        instance_iteration.DownloadGeneralData instance_mgo, instance_status_handler, date_limit, date_limit_delivery
        instance_status_handler.progress_increase
        
        instance_iteration.PrepareItemCollection instance_mgo, item, date_limit, date_limit_delivery, , instance_status_handler, instance_tokenizator
        instance_status_handler.progress_increase
            
        ' problematyczne to wait_for_mgo ;/
        ' w takim razie chyba juz z tym problemu nie bedzie
        ' zalozenie dosyc proste do kazdej procedury
        ' dokladam wbudowane oczekiwanie- wait for mgo
        ' i w ten oto prosty sposob pozbylem sie tloku w tym kodzie
        ' oczywiscie zostawie zaremowania upamietniajace, ale kazde nastepne bedzie juz nago :)
        'instance_mgo.sendKeys "<Clear>"
        ' instance_mgo.waitForMgo
        'instance_mgo.sendKeys "ms9pop00 <Enter>"
        ' instance_mgo.waitForMgo
        'instance_mgo.pMS9POP00.DS = "6"
        ' instance_mgo.waitForMgo
        'instance_mgo.pMS9POP00.PN = id.part_num
        ' instance_mgo.waitForMgo
        'instance_mgo.pMS9POP00.PLT = id.plt_name
        ' instance_mgo.waitForMgo
        'instance_mgo.sendKeys "<Enter>"
        ' instance_mgo.waitForMgo
        ' dane z MGO bezposrednio pamietaj nie ma zadnych posrednikow w postaci zmiennej
        ' MsgBox instance_mgo.pMS9POP00.PN
            
        'instance_iteration.CBAL = instance_mgo.pMS9POP00.CBAL
        'instance_iteration.PartName = instance_mgo.pMS9POP00.PART_NAME
        'instance_iteration.SUPPLIER = instance_mgo.pMS9POP00.SUPPLIER
        'instance_iteration.DUNS = instance_mgo.pMS9POP00.DUNS
        'instance_iteration.Deck = instance_mgo.pMS9POP00.F_U
        'instance_iteration.DOH = instance_mgo.pMS9POP00.firstDOH
        'instance_iteration.FirstRunout = ' to dopiero trzeba przeliczyc na podstawie danych z innych scr
        'instance_iteration.Qhold = instance_mgo.pMS9POP00.QHD
        'instance_iteration.MISC = instance_mgo.pMS9POP00.MISC
        'instance_iteration.StdPack = instance_mgo.pMS9POP00.STD_PACK
        
        ' powyzej obiekty jakie zostaly uzyte:
        ' ms9pop00- tyle, ¿e zawiera sie w instance_mgo, wiec mozna zaniechac, instance_mgo, instance_iteration
        ' teoretycznie mozna zrobic osobna procedure w ktorej odbywaja sie taka cala sielanka :)
        
        
        ' ok wszystko zrobione poza pustymi danymi z niedzieli ktore zostaly pominiete do teraz, aby je wkleic do MISC
        
            
            
        Set id = Nothing
        Set instance_iteration = Nothing
        instance_status_handler.progress_increase
    Next part
    
    instance_layout.GenerujTabeleDlaListyDanychOgolnych
    instance_layout.GenerujReszteRusztownia
    Sheets("register").Range("scopeObliczen") = "all"
    przelicz_parametry_arkusza ActiveSheet, ActiveCell
    przelicz_arkusz ActiveSheet, ActiveCell, True
    Sheets("register").Range("scopeObliczen") = Sheets("register").Range("scopeObliczenDefault")
    DoEvents
    
    Sheets("register").Range("w_macro") = 0
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    instance_status_handler.hide
    
    Columns("B:P").AutoFit
    
    Application.ScreenUpdating = True
    
    If (ActiveSheet.Cells(1, 1) Like "daily*") Or (ActiveSheet.Cells(1, 1) Like "weekly*") Then
        If Not ActiveSheet.AutoFilterMode Then
            Range("b5").AutoFilter
        End If
    End If
    
    
    Set instance_item = Nothing
    Set instance_mgo = Nothing
    
    MsgBox "ready! " & CStr(Now)
    
End Sub

Private Sub prepare_input_data()
    Set INPUT_DATA = New Collection

    
    Dim lr As Long
    Dim item As InputData
    lr = last_row()
    For x = 2 To lr
    
        If Cells(x, 1).EntireRow.Hidden = False Then
        
            Set item = New InputData
            item.plt_name = Cells(x, 1)
            item.part_num = Cells(x, 2)
            item.comment1 = Cells(x, 10)
            item.comment2 = Cells(x, 11)
            INPUT_DATA.Add item
            Set item = Nothing
        End If
    Next x
    
    Sheets("register").Range("itemDays") = 0
    Sheets("register").Range("allParts") = INPUT_DATA.COUNT
    ' MsgBox Sheets("register").Range("allParts")
    Sheets("register").Range("lastRow") = INPUT_DATA.COUNT + 5
    date_limit = CDate(Sheets("register").Range("limitDate"))
End Sub
