VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WeeklyIteration"
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
' ClassModule WeeklyIteration
' ==========================
'

Implements IIteracja

Private ir As Integer
Private production_day_begining As Date
' wlasciwie jest to kolekcja przetrzymujaca doslownie wszystkie dane
' item -> zawiera po kolei rqm, trans as collection, ebal as formula i date
Private item_collection As Collection
Private not_yet_received_collection As Collection
Private tworzenie_tranzytow As TworzenieTranzytow


Private ef As EbalFormula

Private pop As MS9POP00
Private catch_error As CatchError


Private Sub Class_Initialize()
    Set item_collection = New Collection
    Set not_yet_received_collection = New Collection
    Set tworzenie_tranzytow = New TworzenieTranzytow
    Set catch_error = New CatchError
End Sub

Private Sub Class_Terminate()
    Set item_collection = Nothing
    Set not_yet_received_collection = Nothing
    Set tworzenie_tranzytow = Nothing
    Set catch_error = Nothing
End Sub

Private Sub IIteracja_AddItem(arg As IItem)
    item_collection.Add arg
End Sub

Private Property Let IIteracja_Bank(RHS As Long)
    ' dla weekly not important
End Property

Private Property Get IIteracja_Bank() As Long
    IIteracja_Bank = CLng(-1)
End Property

Private Property Let IIteracja_CBAL(RHS As String)
    Cells(ir, 4) = CStr(RHS)
End Property

Private Property Get IIteracja_CBAL() As String
    IIteracja_CBAL = CStr(Cells(ir, 4))
End Property

Private Property Let IIteracja_Cmnt1(RHS As String)
    Cells(ir, 15) = CStr(RHS)
End Property

Private Property Get IIteracja_Cmnt1() As String
    IIteracja_Cmnt1 = CStr(Cells(ir, 15))
End Property

Private Property Let IIteracja_Cmnt2(RHS As String)
    Cells(ir, 16) = CStr(RHS)
End Property

Private Property Get IIteracja_Cmnt2() As String
    IIteracja_Cmnt2 = CStr(Cells(ir, 16))
End Property

Private Property Let IIteracja_Deck(RHS As String)
    Cells(ir, 8) = CStr(RHS)
End Property

Private Property Get IIteracja_Deck() As String
    IIteracja_Deck = CStr(Cells(ir, 8))
End Property

Private Property Let IIteracja_DOH(RHS As String)
    Cells(ir, 9) = CStr(RHS)
End Property

Private Property Get IIteracja_DOH() As String
    IIteracja_DOH = CStr(Cells(ir, 9))
End Property

Private Sub IIteracja_DownloadGeneralData(ByRef i_mgo As MGO, ByRef sh As StatusHandler, limit As Date, limit_del As Date)

        i_mgo.sendKeys "<Clear>"
        i_mgo.sendKeys "ms9pop00 <Enter>"
        i_mgo.pMS9POP00.DS = "6"
        i_mgo.pMS9POP00.pn = Trim(CStr(IIteracja_Partnum))
        i_mgo.pMS9POP00.plt = Trim(CStr(IIteracja_PlantName))
        i_mgo.sendKeys "<Enter>"
        
        Set catch_error = New CatchError
        catch_error.catch_error_if_pn_is_not_available_in_mgo "ms9pop00", i_mgo.pMS9POP00.infocode
        If catch_error.exit_from_sub Then
            ' MsgBox "no data on screen ms9pop00 on part: " & CStr(IIteracja_Partnum) & " infocode: " & CStr(i_mgo.pMS9POP00.infocode)
            Exit Sub
        End If
        
        
        sh.progress_increase
            
        IIteracja_CBAL = i_mgo.pMS9POP00.CBAL
        IIteracja_PartName = i_mgo.pMS9POP00.PART_NAME
        IIteracja_Supplier = i_mgo.pMS9POP00.SUPPLIER
        IIteracja_Duns = i_mgo.pMS9POP00.DUNS
        IIteracja_Deck = i_mgo.pMS9POP00.F_U
        IIteracja_DOH = i_mgo.pMS9POP00.firstDOH
        ' IIteracja_FirstRunout = ' to dopiero trzeba przeliczyc na podstawie danych z innych scr
        IIteracja_Qhold = i_mgo.pMS9POP00.QHD
        IIteracja_MISC = i_mgo.pMS9POP00.MISC
        IIteracja_StdPack = i_mgo.pMS9POP00.STD_PACK
        IIteracja_SheetDef = "weekly " & CStr(Sheets("register").Range("redpink")) & " " & CStr(Now) & " limit on scheduled PUSes: " & limit & " and del limit: " & limit_del & " " & ActiveSheet.Name
        
        
        ' komentarz z kompleksowymi danymi z POPa
        Set pop = i_mgo.pMS9POP00
        txt = _
        "COUNT: " & pop.COUNT & Chr(10) & _
        "O: " & pop.O & Chr(10) & _
        "BBAL: " & pop.BBAL & Chr(10) & _
        "BANK: " & pop.Bank & Chr(10) & _
        "PCS TO GO: " & pop.PCS_TO_GO & Chr(10) & _
        "TT: " & pop.TT & Chr(10) & _
        "O/S: " & pop.OS & Chr(10) & _
        "FMA FU: " & pop.a & Chr(10) & _
        "DK: " & pop.DK & Chr(10) & _
        "IP: " & pop.ip & Chr(10) & _
        "MODE: " & pop.MODE & Chr(10)

        
        Cells(IIteracja_IterationRow, 2).AddComment CStr(txt)
        
        Cells(IIteracja_IterationRow, 2).Comment.Shape.Height = 160
        Cells(IIteracja_IterationRow, 2).Comment.Shape.Width = 300
        
        
        sh.progress_increase
End Sub

Private Property Let IIteracja_Duns(RHS As String)
    Cells(ir, 7) = CStr(RHS)
End Property

Private Property Get IIteracja_Duns() As String
    IIteracja_Duns = CStr(Cells(ir, 7))
End Property

Private Property Let IIteracja_FirstRunout(RHS As String)
    Cells(ir, 10) = CStr(RHS)
End Property

Private Property Get IIteracja_FirstRunout() As String
    IIteracja_FirstRunout = CStr(Cells(interacja_row, 10))
End Property

' tutaj publiczne odniesienie sie do kolecji danych...
' niestety struktura obiektowa krzyzuje moje plany o niezale�nym sciaganiu
' danych z roznych ekranow wszystko z gory musi byc ulozone
' ma to swoje zalety (przejrzystosc kodu)
' ale i wady (latwo napisac brzydki kod, kt�ry nie ma w sobie nawet krzty optymalizacji)
' w koncu jednak jestem (a przynajmniej uwazam sie) za nie byle jakiego programiste
' wiec wybralem ta trudniejsza sciezke (wbrew pozorom), gdzie byc i moze jest pelno
' sidel czychajacych na mnie jednak cel jest znakomisty
' otrzymac spojny kod pracujacy w pieknym paradygmacie obiektowym
' ============================================================
Private Property Let IIteracja_ItemCollection(RHS As Collection)
    Set item_collection = RHS
End Property


Private Property Get IIteracja_ItemCollection() As Collection
    Set IIteracja_ItemCollection = item_collection
End Property
' ============================================================



Private Property Let IIteracja_IterationRow(RHS As Integer)
    ir = Int(RHS)
End Property

Private Property Get IIteracja_IterationRow() As Integer
    IIteracja_IterationRow = Int(ir)
End Property

Private Property Let IIteracja_MISC(RHS As String)
    Cells(ir, 12) = CStr(RHS)
End Property

Private Property Get IIteracja_MISC() As String
    IIteracja_MISC = CStr(Cells(interacja_row, 12))
End Property

Private Property Let IIteracja_NotyetReceived(RHS As Range)
    Cells(ir, 13) = RHS
End Property

Private Property Get IIteracja_NotyetReceived() As Range
    Set IIteracja_NotyetReceived = Cells(ir, 13)
End Property

Private Property Get IIteracja_NotYetReceivedCollection() As Collection
    Set IIteracja_NotYetReceivedCollection = not_yet_received_collection
End Property

Private Property Let IIteracja_PartName(RHS As String)
    Cells(ir, 5) = CStr(RHS)
End Property

Private Property Get IIteracja_PartName() As String
    IIteracja_PartName = CStr(Cells(ir, 5))
End Property

Private Property Let IIteracja_Partnum(RHS As String)
    Cells(ir, 2) = CStr(RHS)
End Property

Private Property Get IIteracja_Partnum() As String
    IIteracja_Partnum = CStr(Cells(ir, 2))
End Property

Private Property Let IIteracja_PlantName(RHS As String)
    Cells(ir, 3) = CStr(RHS)
End Property

Private Property Get IIteracja_PlantName() As String
    IIteracja_PlantName = CStr(Cells(ir, 3))
End Property


' wypelnianie danych w item_collection na podstawie danych z mgo
' =======================================================================================
Private Sub IIteracja_PrepareItemCollection(i_mgo As MGO, item As IItem, limit As Date, limit_del As Date, Optional hourly_extension_day As Date, Optional ByRef sh As StatusHandler, Optional ByRef instance_tokenizator As Tokenizator)

    ' rusztowanie powinno byc zbudowane na bazie rqms
    ' i dopiero na polozyc bedzie mozna PUSy i ASNy
    ' ebal jest tylko i wylacznie formula ktora musi pracowac na zywych danych
    
    Dim is_f8_pressed As Boolean
    is_f8_pressed = False
    Dim i As ITransit
    ' zmienna dla zk7ppus0 w przypadku cofania sie z sprawdzaniem PUSow
    Dim check_past As Integer
    ' tekst na czas delivery troche poza enkapsulacja ale niechaj i zdarzy mi sie taki babuch
    ' taki mark ze to jest moj kod :P
    Dim txt_time As String
    Dim tempdate As Date
    tempdate = CDate(Format(Now, "yyyy-mm-dd"))
    
    
    ' ZK7PWRQM
    ' ===================================================================================
    If instance_tokenizator.token = "DEFAULT" Then
    
        tworzenie_rqmsow i_mgo, item, limit, hourly_extension_day, sh, instance_tokenizator, tempdate
        
        
    ' wypelnienie niemanual
    ElseIf Not instance_tokenizator.getIsManualFill Then
    
        ' nop poniewaz ta metoda ponizej pracuje dobrze tylko dla daily
        ' tworzenie_c_rqmsow i_mgo, item, limit, hourly_extension_day, sh, instance_tokenizator, tempdate
    
    ' wypelnienie manual
    ElseIf instance_tokenizator.getIsManualFill Then
    
        ' nop poniewaz ta metoda ponizej pracuje dobrze tylko dla daily
        ' tworzenie_manualych_rqmsow item, limit, hourly_extension_day, sh, instance_tokenizator, tempdate
        
        ' mimo to sprobujemy dodac na pale implementacje
        tworzenie_manualych_rqmsow_dla_weekly item, limit, hourly_extension_day, sh, instance_tokenizator, tempdate
    
    End If
    
    
    
    Sheets("register").Range("itemDays") = Sheets("register").Range("itemWeeks")
    ' MsgBox Int(Sheets("register").Range("itemDays")) & "    " & Int(tempdate - CDate(Format(Now, "yyyy-mm-dd")))
    
    'If Int(Sheets("register").Range("itemDays")) < Int(tempdate - CDate(Format(Now, "yyyy-mm-dd"))) Then
    '    Sheets("register").Range("itemDays") = Int(tempdate - CDate(Format(Now, "yyyy-mm-dd")))
    '    ' MsgBox Sheets("register").Range("itemDays")
    'End If
    
    ' ale teraz to nie wszystko poniewaz kolekcja itemow
    ' jest zapelniona miejscami tylko i wylacznie tam gdzie faktycznie sa wartosci rqm
    ' teraz jesli PUS albo ASN znajduje sie poza scope'em
    ' bedzie trzeba stworzyc nowy obiekt dodatkowo - a wlasciwie wszsytkie ponieaz i tak
    ' potem trzeba bedzie na koniec wypelnic dane za pomoca ending balance i tak musi sie
    ' odniesc do poprzedniej wartosci, ale jest to opcja gdy chcemy na bank wypelnic wszystko
    ' juz w tej procedurze- musze sie na tym zastanowic czy ending balance moze byc jednak
    ' wypelniany dopiero po layout'cie...
    
    ' ===================================================================================
    
    ' ITRANS( RQM, date_limit )
    ' ===================================================================================
    sh.progress_increase
    
    ' Debug.Assert IIteracja_Partnum <> "13364342"
    
    ' tworzenie_kolekcji_tranzytow i_mgo As MGO, item As IItem, limit As Date, txt_time, item_collection, not_yet_received_collection, is_f8_pressed
    tworzenie_tranzytow.towrzenie_tranzytow_2 sh, ir, IIteracja_PlantName, IIteracja_Partnum, i_mgo, item, limit, limit_del, txt_time, item_collection, not_yet_received_collection, is_f8_pressed
    
    
    
    ' TEST item_collection ' OK
    ' ======================================================================================================
    '    Debug.Print item_collection.COUNT
    '    For x = 1 To item_collection.COUNT
    '        Debug.Print item_collection.item(x).mDate
    '    Next x
    ' ======================================================================================================
    
    
    sh.progress_increase
    ' ===================================================================================
    
    ' EBAL( RQM, ITRANS )
    ' ===================================================================================
    Dim kol As Long
    Dim cl As Range
    ' z lokalnej zmiennej na zmienna instancyjna
    ' Dim ef As EbalFormula
    Set ef = New EbalFormula
    kol = Sheets("register").Range("firstColumn") + 2
    
    For Each item In item_collection
    
        Set ef = New EbalFormula
        If kol = 19 Then
            item.mEbal = ef.fill_first_day()
        Else
            item.mEbal = ef.fill_other_days()
        End If
        Set ef = Nothing
        kol = kol + 3
        
    Next item
    
    
    ' ===================================================================================
    
    ' WYRZUCENIE DANYCH DO ARKUSZA
    ' ===================================================================================
    
    
    
    
        ' z powodu tego ze w PUSach moga byc pickupy juz received nie moge sie posluzyc
        ' tylko i wylacznie countem dla ich kolekcji potrzebuje dodatkowego iteratora
        ' sprawdzajacego faktyczny stan PUSow na not yet received
        Dim var_helper As Integer
        Dim nyrtrans As ITransit
        var_helper = 0
        Dim cmnt As String
        kol = Sheets("register").Range("firstColumn")
        
        cmnt = ""
        For Each nyrtrans In not_yet_received_collection
            If nyrtrans.mNotYetReceived = True Then
                If nyrtrans.mOutOfScope = True Then
                    If nyrtrans.mOnPop = True Then
                        If (nyrtrans.mQty <> 0) Then
                            IIteracja_NotyetReceived.Value = IIteracja_NotyetReceived.Value + nyrtrans.mQty
                            cmnt = cmnt & _
                            "DeliveryDate: " & CStr(nyrtrans.mDeliveryDate) & Chr(10) & _
                            "DeliveryTime: " & CStr(nyrtrans.mDeliveryTime) & Chr(10) & _
                            "Name: " & CStr(nyrtrans.mName) & Chr(10) & _
                            "PickupDate: " & CStr(nyrtrans.mPickupDate) & Chr(10) & _
                            "Qty: " & CStr(nyrtrans.mQty) & Chr(10) & _
                            "Route: " & CStr(nyrtrans.mRoute) & Chr(10) & _
                            "TRLR: " & CStr(nyrtrans.mTRLR) & Chr(10) & _
                            "ST: " & CStr(nyrtrans.mST) & Chr(10) & _
                            "----------------------------------------" & Chr(10) & Chr(10)
                            
                            var_helper = var_helper + 1
                        End If
                    End If
                End If
            End If
        Next nyrtrans
        If var_helper > 0 Then
        
            IIteracja_NotyetReceived.AddComment CStr(cmnt)
            ' 10 wynika z ilosci danych na kazdy transport
            ' 1 - delivery
            ' 9 - time
            ' 2 - name
            ' 3 - pickupdate
            ' 4 - qty
            ' 5 - route
            ' 6,7,8
            ' 9, 10 kontener i status line and chr(10)*2
            ' 11 buffer heurystycznie
            IIteracja_NotyetReceived.Comment.Shape.Height = (11 * var_helper) * 10
            ' heurystycznie :P
            IIteracja_NotyetReceived.Comment.Shape.Width = 200
        End If
        
        
        
        For Each item In item_collection
        
            var_helper = 0
        
            'If Cells(4, kol) = item.mDate Then
            Cells(ir, kol) = item.mReq
            
            ' MsgBox item.mTransCollection.COUNT & " " & item.mDate ' PROBLEM!
            If item.mTransCollection.COUNT > 0 Then
            Set cl = Cells(ir, kol + 1)
            cmnt = ""
            For Each i In item.mTransCollection
            
                If (i.mNotYetReceived = True) And (i.mQty <> 0) Then
                    Cells(ir, kol + 1) = Cells(ir, kol + 1) + i.mQty
                    cmnt = cmnt & _
                    "DeliveryDate: " & CStr(i.mDeliveryDate) & Chr(10) & _
                    "DeliveryTime: " & CStr(i.mDeliveryTime) & Chr(10) & _
                    "Name: " & CStr(i.mName) & Chr(10) & _
                    "PickupDate: " & CStr(i.mPickupDate) & Chr(10) & _
                    "Qty: " & CStr(i.mQty) & Chr(10) & _
                    "Route: " & CStr(i.mRoute) & Chr(10) & _
                    "TRLR: " & CStr(i.mTRLR) & Chr(10) & _
                    "ST: " & CStr(i.mST) & Chr(10) & _
                    "----------------------------------------" & Chr(10) & Chr(10)
                    var_helper = var_helper + 1
                End If
                
            Next i
            
              If var_helper > 0 Then
                cl.AddComment CStr(cmnt)
                ' 10 wynika z ilosci danych na kazdy transport
                ' 1 - delivery
                ' 9 - time
                ' 2 - name
                ' 3 - pickupdate
                ' 4 - qty
                ' 5 - route
                ' 6,7,8
                ' 9, 10 kontener i status line and chr(10)*2
                ' 11 buffer heurystycznie
                cl.Comment.Shape.Height = (11 * var_helper) * 10
                ' heurystycznie :P
                cl.Comment.Shape.Width = 200
              End If
            End If
            
            Set cl = Cells(ir, kol + 2)
            cl.FormulaR1C1 = item.mEbal
            
            'Else
            '    MsgBox "problem!"
            'End If
            kol = kol + 3
        Next item
    
    ' ===================================================================================
    
    

End Sub



Private Property Let IIteracja_ProdDayBegin(RHS As Date)

End Property

Private Property Get IIteracja_ProdDayBegin() As Date

End Property

Private Property Let IIteracja_Qhold(RHS As String)
    Cells(ir, 11) = CStr(RHS)
End Property

Private Property Get IIteracja_Qhold() As String
    IIteracja_Qhold = CStr(Cells(interacja_row, 11))
End Property

Private Property Let IIteracja_SheetDef(RHS As String)
    Cells(1, 1) = CStr(RHS)
End Property

Private Property Get IIteracja_SheetDef() As String
    IIteracja_SheetDef = CStr(Cells(1, 1))
End Property

Private Property Let IIteracja_StdPack(RHS As String)
    Cells(ir, 14) = CStr(RHS)
End Property

Private Property Get IIteracja_StdPack() As String
    IIteracja_StdPack = CStr(Cells(ir, 14))
End Property

Private Property Let IIteracja_Supplier(RHS As String)
    Cells(ir, 6) = CStr(RHS)
End Property

Private Property Get IIteracja_Supplier() As String
    IIteracja_Supplier = CStr(Cells(ir, 6))
End Property



Private Sub tworzenie_rqmsow(i_mgo As MGO, item As IItem, limit As Date, Optional hourly_extension_day As Date, Optional ByRef sh As StatusHandler, Optional ByRef instance_tokenizator As Tokenizator, Optional tempdate As Date)
    
    Dim empty_item As IItem
    Dim data_do_sprawdzenia As Date
    Dim wn As Range, wd As Range, tf As Range
    Set wn = ThisWorkbook.Sheets("register").Range("weekNum")
    Set wd = ThisWorkbook.Sheets("register").Range("weekDay")
    Set tf = ThisWorkbook.Sheets("register").Range("todayFormula")
    ' swn - string on week num
    ' swd - string on week day
    Dim swn As String, swd As String, stf As String
    stf = Replace(tf.formula, "=", "")
    swn = Replace(wn.formula, "-1", stf)
    swd = Replace(wd.formula, "-1", stf)
    
    ' tutaj zk7pdrqm dniowe tylko kilka
    i_mgo.sendKeys "<Clear>"
    i_mgo.sendKeys "zk7pdrqm <Enter>"
    i_mgo.pZK7PDRQM.pn = IIteracja_Partnum
    i_mgo.pZK7PDRQM.plt = IIteracja_PlantName
    i_mgo.pZK7PDRQM.KANBAN = ""
    i_mgo.sendKeys "<Enter>"
    
    Set catch_error = New CatchError
    catch_error.catch_error_if_pn_is_not_available_in_mgo "zk7pwrqm", i_mgo.pZK7PWRQM.infocode
    If catch_error.exit_from_sub Then
        ' MsgBox "no data on screen zk7pwrqm on part: " & CStr(IIteracja_Partnum) & ", infocode: " & CStr(i_mgo.i_mgo.pZK7PWRQM.infocode)
        Exit Sub
    End If
    
    sh.progress_increase
    ' od teraz implementacja zajmujaca sie sciaganiem danych daily
    ' ===========================================================
    
    ' iteracja po wszystkich dniach
    ' 2012-11-09 poprawka dla GMNA
    ' tutaj jest masakra poniewaz MGO zachowuje sie naprawde bardzo srednio
    ' w przypadku posiadania duzej ilosci MY na ekranie zk7pdrqm
    ' jest mus robienia na okolo kolejnych durnych przeskakiwan miedzy screenami
    ' na szczescie architektura ffh jest na tyle znosna ze poszlo bez bolu
    ' jednak musze miec to na wzgledzie ze kod zostal przyrzadzony na predce i srednio
    ' mu do naprawde pieknego kodu
    ' duzo proceduralnego shitu
    Dim temp_rqm_from_daily As Long
    temp_rqm_from_daily = 0
    
    
    Dim pf8_pressed As Boolean
    Dim ktora_iteracja_na_glownej_petli_do
    ktora_iteracja_na_glownej_petli_do = 0
    pf8_pressed = False
    Do
    
        ' wraz z nowa iteracja tabularaza jesli chodzi o klikanie w F8!
        ktora_iteracja_na_glownej_petli_do = ktora_iteracja_na_glownej_petli_do + 1
    
        ' czesci bez danych
        If i_mgo.getString(22, 2, 5) Like "*R6017*" Then
            Exit Do
        End If
        
        ' part not found
        If i_mgo.getString(22, 2, 5) Like "*R6117*" Then
            Exit Do
        End If
    
        For col = 0 To 4
        
            check_this_date_if_sunday = CDate(i_mgo.convertToDateFromZK7PDRQMDate(i_mgo.pZK7PDRQM.dates(Int(col))))
            If (ktora_iteracja_na_glownej_petli_do = 1) And (col = 0) And _
                (Weekday(check_this_date_if_sunday) = 1) Then
                
                Cells(ir, 12) = Cells(ir, 12) + i_mgo.pZK7PDRQM.PLT_TOTAL(Int(col), pf8_pressed)
                Cells(ir, 12).AddComment "Sunday production: " & CStr(i_mgo.pZK7PDRQM.PLT_TOTAL(Int(col), pf8_pressed)) & Chr(10)
                Cells(ir, 12).Comment.Shape.Height = Cells(ir, 12).Comment.Shape.Height * 0.4
            Else
            
                If TypeName(item) Like "*Daily*" Then
                    Set item = Nothing
                    Set item = New ItemDaily
                ElseIf TypeName(item) Like "*Hourly*" Then
                    Set item = Nothing
                    Set item = New ItemHourly
                End If
                
                If Trim(i_mgo.pZK7PDRQM.dates(Int(col))) <> "" Then
                    
                    data_do_sprawdzenia = CDate(i_mgo.convertToDateFromZK7PDRQMDate(i_mgo.pZK7PDRQM.dates(Int(col))))
                    ' tutaj zle dobiera CW!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    ' cw_daty = Evaluate(Replace(wn.formula, "-1", """ & CStr(data_do_sprawdzenia) & """))
                    ' MsgBox data_do_sprawdzenia
                    cw_daty = Evaluate(Replace(wn.formula, "-1", """" & CStr(data_do_sprawdzenia) & """"))
                    
                    If Int(cw_daty) = Int(Evaluate(swn)) Then
                        temp_rqm_from_daily = temp_rqm_from_daily + CLng(i_mgo.pZK7PDRQM.PLT_TOTAL(Int(col), pf8_pressed))
                    ElseIf Int(cw_daty) > Int(Evaluate(swn)) Then
                        Exit Do
                    End If
                End If
            End If
        Next col
        
        
        If pf8_pressed = False Then
            i_mgo.sendKeys "<pf11>"
        ElseIf pf8_pressed = True Then
            i_mgo.sendKeys "<Enter>"
            For x = 1 To ktora_iteracja_na_glownej_petli_do
                i_mgo.sendKeys "<pf11>"
            Next x
            
            ' i_mgo.sendKeys "<pf8>"
            
        End If
        
        If i_mgo.getString(22, 2, 5) Like "*R6120*" Then
            Exit Do
        End If
        
        If i_mgo.getString(22, 2, 5) Like "*R6123*" Then
            Exit Do
        End If
        
        If i_mgo.getString(22, 2, 5) Like "*R6017*" Then
            Exit Do
        End If
    Loop While True
    
    ' ===========================================================
    
    
    ' czesc z tygodniowym zapotrzebowaniem od 2 CW aktualnego
    i_mgo.sendKeys "<Clear>"
    i_mgo.sendKeys "zk7pwrqm <Enter>"
    i_mgo.pZK7PWRQM.pn = IIteracja_Partnum
    i_mgo.pZK7PWRQM.plt = IIteracja_PlantName
    i_mgo.pZK7PWRQM.KANBAN = ""
    i_mgo.sendKeys "<Enter>"
    
    sh.progress_increase
    
    
    ' iteracja po wszystkich dniach
    ' 2012-11-09 poprawka dla GMNA
    ' tutaj jest masakra poniewaz MGO zachowuje sie naprawde bardzo srednio
    ' w przypadku posiadania duzej ilosci MY na ekranie zk7pdrqm
    ' jest mus robienia na okolo kolejnych durnych przeskakiwan miedzy screenami
    ' na szczescie architektura ffh jest na tyle znosna ze poszlo bez bolu
    ' jednak musze miec to na wzgledzie ze kod zostal przyrzadzony na predce i srednio
    ' mu do naprawde pieknego kodu
    ' duzo proceduralnego shitu
    'Dim pf8_pressed As Boolean
    'Dim ktora_iteracja_na_glownej_petli_do
    ktora_iteracja_na_glownej_petli_do = 0
    pf8_pressed = False
    Do
    
        ' wraz z nowa iteracja tabularaza jesli chodzi o klikanie w F8!
        ktora_iteracja_na_glownej_petli_do = ktora_iteracja_na_glownej_petli_do + 1
    
        ' czesci bez danych
        If i_mgo.getString(22, 2, 5) Like "*R6017*" Then
            Exit Do
        End If
        
        ' part not found
        If i_mgo.getString(22, 2, 5) Like "*R6117*" Then
            Exit Do
        End If
    
        For col = 0 To 4
            If TypeName(item) Like "*Daily*" Then
                Set item = Nothing
                Set item = New ItemDaily
            ElseIf TypeName(item) Like "*Hourly*" Then
                Set item = Nothing
                Set item = New ItemHourly
            ElseIf TypeName(item) Like "*Weekly*" Then
                Set item = Nothing
                Set item = New ItemWeekly
            End If
            
            If (Trim(i_mgo.pZK7PWRQM.weeks(Int(col))) <> "") Then
            
            
                ' wyodrebniony numer z stringu ktory zawsze zawiera np
                ' WEEK  1, WEEK  2 ... WEEK 12 i tak dalej
                uchwyt_stra_WEEK_x = i_mgo.pZK7PWRQM.weeks(Int(col))
                ' najpierw wyrzucenie spacji, potem wyrzucenie literalu WEEK
                'uchwyt_stra_WEEK_x = Replace(uchwyt_stra_WEEK_x, " ", "")
                ' uchwyt_stra_WEEK_x = Replace(uchwyt_stra_WEEK_x, "WEEK", "")
                uchwyt_stra_WEEK_x = Trim(Right(uchwyt_stra_WEEK_x, 2))
        
                If Int(uchwyt_stra_WEEK_x) <> 1 Then
                    item.mReq = i_mgo.pZK7PWRQM.PLT_TOTAL(Int(col), pf8_pressed)
                ElseIf Int(uchwyt_stra_WEEK_x) = 1 Then
                    item.mReq = temp_rqm_from_daily
                End If
                
                ' ta implementacja jest kiepsko napisana
                ' curr_cw = Int(uchwyt_stra_WEEK_x) + Evaluate(swn) - 1
                
                curr_cw = foo_stworz_str_na_podstawie(Int(Evaluate(swn)), Int(uchwyt_stra_WEEK_x), i_mgo)
                
                If curr_cw = -1 Then
                    MsgBox "cos poszlo nie tak z proba przypisania aktualnego cw!"
                    End
                End If
                
                item.mDate = i_mgo.pZK7PWRQM.sunday_from_cw(Int(curr_cw))
                
                ' procedura zajmujaca sie wypelnianiem kolekcji pustymi itemami,
                ' ktore nie maja daty uwzglednionej w zk7pdrqm
                ' check_empty_event item_collection, item, item.mDate - tempdate
                ' jednak pierwotna implementacja zakladala podliczanie pustych dni a nie weekow
                ' bedzie trzeba dopasowac
                ' okazuje sie ze w przypadku weekly nie trzeba sie o to martwic gdyz wszystkie weeki
                ' po kolei sie pokazuja bez bolu tylka
    
                item_collection.Add item
            End If
        Next col
        
        
        If pf8_pressed = False Then
            i_mgo.sendKeys "<pf11>"
        ElseIf pf8_pressed = True Then
            i_mgo.sendKeys "<Enter>"
            For x = 1 To ktora_iteracja_na_glownej_petli_do
                i_mgo.sendKeys "<pf11>"
            Next x
            
            ' i_mgo.sendKeys "<pf8>"
            
        End If
        
        If i_mgo.getString(22, 2, 5) Like "*R6120*" Then
            Exit Do
        End If
        
        If i_mgo.getString(22, 2, 5) Like "*R6123*" Then
            Exit Do
        End If
        
        If i_mgo.getString(22, 2, 5) Like "*R6017*" Then
            Exit Do
        End If
    Loop While True
End Sub

Private Function foo_stworz_str_na_podstawie(current_cw As Integer, ofst As Integer, i_mgo As MGO) As Integer
    
    foo_stworz_str_na_podstawie = -1
    
    ' ale prawdziwy ofst trzeba zmniejszyc o jeden
    ' bo liczymy od WEEK 1, anie od od WEEK 0
    ofst = ofst - 1
    
    If ofst > 0 Then
        Dim poniedzialek As Date
        poniedzialek = CDate(i_mgo.pZK7PWRQM.monday_from_cw(current_cw))
        
        foo_stworz_str_na_podstawie = Application.WorksheetFunction.IsoWeekNum(CDbl(CDate(poniedzialek + ofst * 7)))
    Else
        foo_stworz_str_na_podstawie = current_cw
    End If
    
    
    
    
End Function

Private Sub tworzenie_manualych_rqmsow_dla_weekly(item As IItem, limit As Date, Optional hourly_extension_day As Date, Optional ByRef sh As StatusHandler, Optional ByRef instance_tokenizator As Tokenizator, Optional tempdate As Date)
    If instance_tokenizator.getIteration < -1 Then
        If Not instance_tokenizator.getCustomRqmTable Is Nothing Then
            instance_tokenizator.setIteration = instance_tokenizator.getCustomRqmTable.COUNT
        End If
    End If
    
    If instance_tokenizator.getIteration > 0 Then
        
        ' empty
        If instance_tokenizator.token Like "*EMPTY*RQM*" Then
            
            fill_collection_with_manual_data item, limit, hourly_extension_day, sh, instance_tokenizator, tempdate
        ElseIf Not instance_tokenizator.getCustomRqmTable Is Nothing Then
            
            fill_collection_with_manual_data item, limit, hourly_extension_day, sh, instance_tokenizator, tempdate, instance_tokenizator.getCustomRqmTable
        End If
        
        ' tabele
    
    End If
End Sub

Private Sub fill_collection_with_manual_data(item As IItem, limit As Date, Optional hourly_extension_day As Date, Optional ByRef sh As StatusHandler, Optional ByRef instance_tokenizator As Tokenizator, Optional tempdate As Date, Optional dataFromRange As Range)

    Dim empty_item As IItem
    
    tempdate = getNextSundayForWeeklyManualRqm(tempdate)
    
    For x = 1 To instance_tokenizator.getIteration
        
        If TypeName(item) Like "*Daily*" Then
            Set item = Nothing
            Set item = New ItemDaily
        ElseIf TypeName(item) Like "*Hourly*" Then
            Set item = Nothing
            Set item = New ItemHourly
        ElseIf TypeName(item) Like "*Weekly*" Then
            Set item = Nothing
            Set item = New ItemWeekly
        End If
        
        item.mDate = tempdate
        
        If (instance_tokenizator.getCustomRqmTable Is Nothing) Then
        
            item.mReq = 0
        ElseIf Not instance_tokenizator.getCustomRqmTable Is Nothing Then
            
            If x > instance_tokenizator.getCustomRqmTable.COUNT Then
                item.mReq = 0
            Else
                item.mReq = instance_tokenizator.getCustomRqmTable.item(x)
            End If
        End If
        
        'If (instance_tokenizator.getCustomTransitTable Is Nothing) Or (x > instance_tokenizator.getCustomTransitTable.COUNT) Then
            '
        'Else
            '
        'End If
        
        item_collection.Add item
        tempdate = tempdate + 7
    Next x
End Sub


Private Function getNextSundayForWeeklyManualRqm(d As Date) As Date

    jakiDzien = Weekday(d)
    
    getNextSundayForWeeklyManualRqm = CDate(d - jakiDzien + 1 + 7)
End Function

Private Function getMondayForWeeklyManualRqm(d As Date) As Date

    jakiDzien = Weekday(d)
    
    getMondayForWeeklyManualRqm = CDate(d - jakiDzien + 2)
End Function


Private Sub tworzenie_c_rqmsow(i_mgo As MGO, item As IItem, limit As Date, Optional hourly_extension_day As Date, Optional ByRef sh As StatusHandler, Optional ByRef instance_tokenizator As Tokenizator, Optional tempdate As Date)

    ' nothing for now
    
End Sub


