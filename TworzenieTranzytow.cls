VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TworzenieTranzytow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private flaga_miejsca_w_kolekcji As Integer
Private part_posiada_pusy As Boolean
' przygotowanie warunków brzegowych dla obliczeñ
' pierwsza data w kolekcji itemow
Private first_date As Date
' ostatnia data w kolekcji itemow
Private last_date As Date


Private Sub Class_Initialize()
    flaga_miejsca_w_kolekcji = 1
End Sub

Public Sub towrzenie_tranzytow_2(ByRef sth As StatusHandler, ByRef ir As Integer, ByVal IIteracja_PlantName As String, _
    ByVal IIteracja_Partnum As String, ByRef i_mgo As MGO, ByRef item As IItem, ByRef limit As Date, ByRef limit_del As Date, _
    ByRef txt_time As String, ByRef item_collection As Collection, _
    ByRef not_yet_received_collection As Collection, ByRef is_f8_pressed As Boolean, Optional poczatek_dnia_prod As Date)
    
    
    
    ' Debug.Assert IIteracja_Partnum <> "13364342"
    
    ' kolekcja nyrc od samego poczatku byla upierdliwa jak tysiac pincet
    Set not_yet_received_collection = Nothing
    Set not_yet_received_collection = New Collection
    
    check_past = Int(Sheets("register").Range("History"))
    
    
    ' pierwsza rzecz otworzyc zk7ppus0 i zobaczyc czy cos mamy
    i_mgo.sendKeys "<Clear>"
    i_mgo.sendKeys "zk7ppus0 <Enter>"
    i_mgo.pZK7PPUS0.plt = IIteracja_PlantName
    i_mgo.pZK7PPUS0.pn = IIteracja_Partnum
    i_mgo.pZK7PPUS0.DELIVERY_DATE_FROM = i_mgo.convertToMgoLongDate(Now - Int(check_past))
    i_mgo.sendKeys "<Enter>"
    i_mgo.waitForMgo
    i_mgo.sendKeys "<pf12>"
    i_mgo.waitForMgo
    
    
    ' ten debug assert dalem zeby sprawdzic co jest nie tak
    ' okazuje sie, ze ktos wystawil pusa z del time jako 24:00
    ' Debug.Assert IIteracja_Partnum <> "13447214"
    
    ' procedura z zaldaowaniem wszystkich danych do nyrc
    zaladuj_do_ sth, ir, IIteracja_PlantName, IIteracja_Partnum, i_mgo, item, limit, limit_del, txt_time, item_collection, not_yet_received_collection, is_f8_pressed, poczatek_dnia_prod
    
    ' sprawdz ktore faktycznie tranzyty nie sa jeszcze zabookowane
    ' oczywiscie przyda sie tutaj przeszukiwanie binarne bedzie szybciuchniej
    compare_for_not_yet_received_and_delete_duplicates sth, ir, IIteracja_PlantName, IIteracja_Partnum, i_mgo, item, limit, txt_time, item_collection, not_yet_received_collection, is_f8_pressed, poczatek_dnia_prod
    
    ' wpasuj wszystko w dane w item_collection
    dopasuj_dane_ sth, ir, IIteracja_PlantName, IIteracja_Partnum, i_mgo, item, limit, txt_time, item_collection, not_yet_received_collection, is_f8_pressed, poczatek_dnia_prod
    
End Sub

Private Sub zaladuj_do_(ByRef sth As StatusHandler, ByRef ir As Integer, ByVal IIteracja_PlantName As String, _
    ByVal IIteracja_Partnum As String, ByRef i_mgo As MGO, ByRef item As IItem, ByRef limit As Date, ByRef limit_del As Date, _
    ByRef txt_time As String, ByRef item_collection As Collection, _
    ByRef not_yet_received_collection As Collection, ByRef is_f8_pressed As Boolean, Optional poczatek_dnia_prod As Date)
    
    If CStr(i_mgo.getString(1, 2, 8)) = "ZK7PPUM0" Then
        part_posiada_pusy = True
        p_tworzenie_tranzytow_on_zk7ppum0 sth, ir, IIteracja_PlantName, IIteracja_Partnum, i_mgo, item, limit, limit_del, txt_time, item_collection, not_yet_received_collection, is_f8_pressed, poczatek_dnia_prod
    Else
        part_posiada_pusy = False
    End If
    p_tworzenie_tranzytow_on_ms9po400 sth, ir, IIteracja_PlantName, IIteracja_Partnum, i_mgo, item, limit, limit_del, txt_time, item_collection, not_yet_received_collection, is_f8_pressed, poczatek_dnia_prod
    

End Sub

Private Sub compare_for_not_yet_received_and_delete_duplicates(ByRef sth As StatusHandler, ByRef ir As Integer, ByVal IIteracja_PlantName As String, _
    ByVal IIteracja_Partnum As String, ByRef i_mgo As MGO, ByRef item As IItem, ByRef limit As Date, _
    ByRef txt_time As String, ByRef item_collection As Collection, _
    ByRef not_yet_received_collection As Collection, ByRef is_f8_pressed As Boolean, Optional poczatek_dnia_prod As Date)
    
    first_date_for_pickup_date = CDate(Format(Now, "yyyy-mm-dd"))
    ' first_date = Now
    
    
    ' transit outside loop
    Dim t_o As ITransit
    ' transit inside loop
    Dim t_i As ITransit
    
    'For i = 1 To not_yet_received_collection.COUNT
    '      Set t_o = not_yet_received_collection.item(i)
    '
    '      MsgBox t_o.mName & " " & t_o.mPickupDate & " " & t_o.mPickupTime & " " & t_o.mQty & " " & t_o.mDeliveryTime
    'Next i
    
    If not_yet_received_collection.COUNT > 1 Then
        
        For i = 1 To not_yet_received_collection.COUNT
        
            Set t_o = not_yet_received_collection.item(i)
            
            ' Debug.Assert Not t_o.mName Like "*EIS101*"
            
            For j = i + 1 To not_yet_received_collection.COUNT
            
            
                ' oczywiscie non sens porownywac itemy w tych samych miejscach bo to ten sam po krotce
                ' wiec:
                If j > i Then
                
                    Set t_i = not_yet_received_collection.item(j)
                    
                    ' Debug.Assert CStr(t_o.mName) <> "PO1146232"
                    ' tutaj nalezy mi sie krotkie wyjasnienie algorytmu poniewaz pisalem to juz
                    ' jakis czas temu
                    ' i myslalem ze jest to klarowne a jednak okazalo sie ze nie tak bardzo
                    '
                    ' po pierwsze sprawdzamy czy zawartosci pickupa nie jest pusta a to wynika z elementu ktory znajduje
                    ' sie pozniej, czyli:
                    '
                    ' If t_o.mOnPop Then
                    '   t_o.mQty = 0
                    ' ElseIf t_i.mOnPop Then
                    '   t_i.mQty = 0
                    ' End If
                    ' tutaj troche ryzykjowny
                    ' kodzik poniewaz zmienia dane ktore jeszcze iteracja nie przejrzala
                    ' co za tym idzie jeden falszywy krok i jestem w powaznych tarapatach :P 2014-03-17
                    '
                    If (t_o.mName = t_i.mName) And (CStr(t_o.mDeliveryDate) = CStr(t_i.mDeliveryDate)) Then
                        
                        
                        
                        ' tutaj znajduje sie ten kluczowy element kodu
                        ' ktory zeruje wszystko na wariata
                        ' na szczescie iteracja jest tak napsiana ze nigdy dwa razy to samo porownanie nie nastapi
                        ' zatem moge sobie pozwolic na taka ekstrawagancje
                        If t_o.mOnPop Then
                            t_o.mQty = 0
                        ElseIf t_i.mOnPop Then
                            t_i.mQty = 0
                        End If
                        
                        
                        ' ten zapis ponizej nadaje pewnosci ze niezaleznie od elementu kolekcji
                        ' co jest calkiem super wg mnie
                        t_o.mOnPop = True
                        t_i.mOnPop = True
                        Exit For
                    End If
                    
                End If
            
            
            Next j
            
            
            ' wyzerowane zostaly wszystkie transporty ktore zostaly zdublowane na po400
            ' trzeba teraz sprawdzic wszystkie not yet received oraz inne glupie warunki
            ' Debug.Print t_o.mName
            
            
            ' OBSOLETE
            'If t_o.mDeliveryDate < CDate(Format(Now, "yyyy-mm-dd")) Then
            '    If Not t_o.mOnPop Then
            '            t_o.mQty = 0
            '    End If
            'Else
            '    If Not t_o.mOnPop Then
            '        If t_o.mDeliveryDate = t_o.mPickupDate Then
            '                If t_o.mPickupDate < CDate(Format(Now, "yyyy-mm-dd")) Then
            '                    t_o.mQty = 0
            '                End If
            '        Else
            '            If t_o.mPickupTime < CDate(Format(Now, "hh:mm")) Then
            '                t_o.mQty = 0
            '            End If
            '        End If
            '    End If
            'End If
            
            
            ' new 2014-03-17
            ' sprawdzam scope dla podwojonych z MS9PO400 - property on POP
            ' dodatkowo od razu ignorujemy "puste"
            If t_o.mQty > 0 Then
                If t_o.mOnPop Then
                
                    ' troche restrykcyjne ale jesli stanie sie to skasuje takie male ryzyko zawodowe
                    ' Debug.Print t_o.mName
                    If CDate(CStr(t_o.mPickupDate) & " " & CStr(t_o.mPickupTime)) > Now Then
                        t_o.mQty = 0
                    End If
                
                Else ' not on pop
                    If CDate(CStr(t_o.mPickupDate) & " " & CStr(t_o.mPickupTime)) <= Now Then
                        t_o.mQty = 0
                    End If
                End If
            End If
            
        Next i
    End If
End Sub

Private Sub dopasuj_dane_(ByRef sth As StatusHandler, ByRef ir As Integer, ByVal IIteracja_PlantName As String, _
    ByVal IIteracja_Partnum As String, ByRef i_mgo As MGO, ByRef item As IItem, ByRef limit As Date, _
    ByRef txt_time As String, ByRef item_collection As Collection, _
    ByRef not_yet_received_collection As Collection, ByRef is_f8_pressed As Boolean, Optional poczatek_dnia_prod As Date)
    
    '   ==============================================================================================================================
    ' i teraz moze uda sie uzyc przeszukiwania binarnego zeby je wkleic do kolekcji regularnych itemow :P
    ' item_collection
    Dim lewy As Integer
    Dim prawy As Integer
    Dim srodek As Integer
    Dim znalazlem As Boolean
    Dim ii As IItem
    Dim temp_trans As ITransit
    
    sth.progress_increase
    
    If item_collection.COUNT > 0 Then
        ' MsgBox TypeName(item_collection.item(1))
        If (TypeName(item_collection.item(1)) Like "*eekly*") Then
            sth.progress_increase
            
            
            For y = 1 To not_yet_received_collection.COUNT
            ' wlasciwie implementacja przeszukiwania binarnego :D szybkosc algorytmu O(log_2 N)
                Set temp_trans = not_yet_received_collection.item(y)
                
                If temp_trans.mOutOfScope = False Then
                
                    lewy = 1
                    prawy = item_collection.COUNT
                    znalazlem = False
                    Do While (lewy <= prawy) And (znalazlem = False)
                        srodek = Int((lewy + prawy) / 2)
                        ' if item_collection.Item(srodek)
                        Set ii = Nothing
                        Set ii = item_collection.item(srodek)
                        
                        If CDate(Format(ii.mDate, "yyyy-mm-dd")) >= CDate(Format(temp_trans.mDeliveryDate, "yyyy-mm-dd")) And _
                        CDate((CDate(Format(ii.mDate, "yyyy-mm-dd")) - 6) <= CDate(Format(temp_trans.mDeliveryDate, "yyyy-mm-dd"))) Then
                            
                            ' narazie na samej dacie teraz jeszcze godzina!
                            ' poniewaz korzystam z algorytmu binarnego nie moge tak po prostu
                            znalazlem = True
                            ii.addTransToCollection temp_trans
                            Exit Do
                            
                            
                        Else
                            If CDate(Format(ii.mDate, "yyyy-mm-dd")) < CDate(Format(temp_trans.mDeliveryDate, "yyyy-mm-dd")) Then
                                lewy = srodek + 1
                            ElseIf CDate(Format(ii.mDate, "yyyy-mm-dd")) > CDate(Format(temp_trans.mDeliveryDate, "yyyy-mm-dd")) Then
                                prawy = srodek - 1
                            End If
                            
                        End If
                    Loop
                    
                    If znalazlem = False Then
                        ii.addTransToCollection temp_trans
                    End If
                End If
            Next y
            
            
        ElseIf (TypeName(item_collection.item(1)) Like "*ourly*") Then
            
            ' algorytm na dniowego ff- prostszy

            sth.progress_increase
            
            ' Debug.Print not_yet_received_collection.COUNT
            
            For y = 1 To not_yet_received_collection.COUNT
            ' wlasciwie implementacja przeszukiwania binarnego :D szybkosc algorytmu O(log_2 N)
                Set temp_trans = not_yet_received_collection.item(y)
                
                If temp_trans.mOutOfScope = False Then
                
                    lewy = 1
                    prawy = item_collection.COUNT
                    znalazlem = False
                    Do While (lewy <= prawy) And (znalazlem = False)
                        srodek = Int((lewy + prawy) / 2)
                        ' if item_collection.Item(srodek)
                        Set ii = Nothing
                        Set ii = item_collection.item(srodek)
                        If CDate(Format(ii.mDate, "yyyy-mm-dd")) = CDate(Format(temp_trans.mDeliveryDate, "yyyy-mm-dd")) Then
                            
                            ' narazie na samej dacie teraz jeszcze godzina!
                            ' poniewaz korzystam z algorytmu binarnego nie moge tak po prostu
                            If ii.mDate > CDate(CStr(temp_trans.mDeliveryDate) & " " & CStr(temp_trans.mDeliveryTime)) Then
                                
                                Do While ii.mDate > CDate(CStr(temp_trans.mDeliveryDate) & " " & CStr(temp_trans.mDeliveryTime))
                                    srodek = srodek - 1
                                    If srodek > 0 Then
                                        Set ii = item_collection.item(srodek)
                                    Else
                                        Set ii = item_collection.item(1)
                                        Exit Do
                                    End If
                                Loop
                                
                                znalazlem = True
                                ii.addTransToCollection temp_trans
                                Exit Do
                            ElseIf ii.mDate < CDate(CStr(temp_trans.mDeliveryDate) & " " & CStr(temp_trans.mDeliveryTime)) Then
                            
                                Do While ii.mDate < CDate(CStr(temp_trans.mDeliveryDate) & " " & CStr(temp_trans.mDeliveryTime))
                                    srodek = srodek + 1
                                    Set ii = item_collection.item(srodek)
                                Loop
                                
                                znalazlem = True
                                ii.addTransToCollection temp_trans
                                Exit Do
                            ElseIf ii.mDate = CDate(CStr(temp_trans.mDeliveryDate) & " " & CStr(temp_trans.mDeliveryTime)) Then
                                znalazlem = True
                                ii.addTransToCollection temp_trans
                                Exit Do
                            End If
                            
                        Else
                            If CDate(Format(ii.mDate, "yyyy-mm-dd")) < CDate(Format(temp_trans.mDeliveryDate, "yyyy-mm-dd")) Then
                                lewy = srodek + 1
                            ElseIf CDate(Format(ii.mDate, "yyyy-mm-dd")) > CDate(Format(temp_trans.mDeliveryDate, "yyyy-mm-dd")) Then
                                prawy = srodek - 1
                            End If
                            
                        End If
                    Loop
                    
                    If znalazlem = False Then
                        ii.addTransToCollection temp_trans
                    End If
                End If
            Next y
            
            
        ElseIf TypeName(item_collection.item(1)) Like "*aily*" Then
            
            
            sth.progress_increase
            For y = 1 To not_yet_received_collection.COUNT
            ' wlasciwie implementacja przeszukiwania binarnego :D szybkosc algorytmu O(log_2 N)
                Set temp_trans = not_yet_received_collection.item(y)
                
                If temp_trans.mOutOfScope = False Then
                
                    lewy = 1
                    prawy = item_collection.COUNT
                    znalazlem = False
                    Do While (lewy <= prawy) And (znalazlem = False)
                        srodek = Int((lewy + prawy) / 2)
                        ' if item_collection.Item(srodek)
                        Set ii = item_collection.item(srodek)
                        If ii.mDate = temp_trans.mDeliveryDate Then
                            znalazlem = True
                            ii.addTransToCollection temp_trans
                            Exit Do
                        Else
                            If ii.mDate < temp_trans.mDeliveryDate Then
                                lewy = srodek + 1
                            ElseIf ii.mDate > temp_trans.mDeliveryDate Then
                                prawy = srodek - 1
                            End If
                            
                        End If
                    Loop
                End If
            Next y
            
        End If
        
        sth.progress_increase
    End If
    
End Sub


Public Sub tworzenie_tranzytow(ByRef sth As StatusHandler, ByRef ir As Integer, ByVal IIteracja_PlantName As String, _
    ByVal IIteracja_Partnum As String, ByRef i_mgo As MGO, ByRef item As IItem, ByRef limit As Date, ByRef limit_del As Date, _
    ByRef txt_time As String, ByRef item_collection As Collection, _
    ByRef not_yet_received_collection As Collection, ByRef is_f8_pressed As Boolean, Optional poczatek_dnia_prod As Date)
    
    ' 13 progress increasow
    
    ' procedura ta ma byc jednolinijkowa podmiana pod kawa³ kodu pracujacy pod ms9po400, zk7ppus0, zk7ppum0, ms9ph100
    ' =====================================================================================================================
     
    ' 1 otworzenie zk7ppus0 i sprawdzenie czy ten part number pracuje w takim paradygmacie
    ' 2 warunek tworzenia procedury pod PUSy lub ASNy
    '   ASN - wlasciwie moze zostac stary algorytm
    '   PUS - napisanie wszystkiego od nowa gdy¿ kuleje poprzedni algorytm szczególnie jeœli chodzi o nyrt
    '   a poniewa¿ wiekosc starego kodu byla mocno ze soba powiazana nie ma sensu go ratowaæ a lepiej napisaæ
    '   od nowa lepiej zorganizowany nowy kodzik, ktory lepiej dziala z MGO oraz jest na wyzszym poziomie enkapsulacji
    '   i podzielenia na niezale¿ne procedury które o wiele latwiej jest koordynowac oraz kleic i update'owac :)
    '   ==============================================================================================================================
    
    Set not_yet_received_collection = Nothing
    Set not_yet_received_collection = New Collection
    
    check_past = Int(Sheets("register").Range("History"))
    
    
    ' pierwsza rzecz otworzyc zk7ppus0 i zobaczyc czy cos mamy
    i_mgo.sendKeys "<Clear>"
    i_mgo.sendKeys "zk7ppus0 <Enter>"
    i_mgo.pZK7PPUS0.plt = IIteracja_PlantName
    i_mgo.pZK7PPUS0.pn = IIteracja_Partnum
    i_mgo.pZK7PPUS0.DELIVERY_DATE_FROM = i_mgo.convertToMgoLongDate(Now - Int(check_past))
    i_mgo.sendKeys "<Enter>"
    i_mgo.waitForMgo
    i_mgo.sendKeys "<pf12>"
    
    sth.progress_increase
    
    If CStr(i_mgo.getString(22, 2, 5)) = "R6285" Then
        ' ASNs
        p_tworzenie_tranzytow_on_ms9po400 sth, ir, IIteracja_PlantName, IIteracja_Partnum, i_mgo, item, limit, limit_del, txt_time, item_collection, not_yet_received_collection, is_f8_pressed, poczatek_dnia_prod
        sth.progress_increase
        ' jesli przeszlismy do strony z danymi znaczy ze mozna kontynuowac algorytm
    ElseIf CStr(i_mgo.getString(1, 2, 8)) = "ZK7PPUM0" Then
        ' PUSes
        
        p_tworzenie_tranzytow_on_zk7ppum0 sth, ir, IIteracja_PlantName, IIteracja_Partnum, i_mgo, item, limit, limit_del, txt_time, item_collection, not_yet_received_collection, is_f8_pressed, poczatek_dnia_prod
        p_towrzenie_tranzytow_on_ms9ph100 sth, ir, IIteracja_PlantName, IIteracja_Partnum, i_mgo, item, limit, limit_del, txt_time, item_collection, not_yet_received_collection, is_f8_pressed, poczatek_dnia_prod
        sth.progress_increase
    End If
    
    dopasuj_dane_ sth, ir, IIteracja_PlantName, IIteracja_Partnum, i_mgo, item, limit, txt_time, item_collection, not_yet_received_collection, is_f8_pressed, poczatek_dnia_prod
    
End Sub

Private Sub p_tworzenie_tranzytow_on_ms9po400(ByRef sth As StatusHandler, ByRef ir As Integer, ByVal IIteracja_PlantName As String, _
    ByVal IIteracja_Partnum As String, ByRef i_mgo As MGO, ByRef item As IItem, ByRef limit As Date, ByRef limit_del As Date, _
    ByRef txt_time As String, ByRef item_collection As Collection, _
    ByRef not_yet_received_collection As Collection, ByRef is_f8_pressed As Boolean, Optional poczatek_dnia_prod As Date)


        ' temp item
        Dim temp_item As IItem
        ' handler pod tranzyt
        Dim temp_trans As ITransit
        
        sth.progress_increase
        
        ' w zaleznosci od kolekcji itemow ustawiamy warunki brzegowe
        If item_collection.COUNT > 0 Then
            Set temp_item = Nothing
            Set temp_item = item_collection.item(1)
            ' tutaj pogrzebany pies poniewaz caly czas patrzylem na niedziele stad
            ' warunek brzegowy byl zawsze przesuniety o tydzien do przodu
            ' oczywiscie inny warunek na daily inny na weekly stad dodatkowy podwojny if ifelse
            If Cells(1, 1) Like "daily*" Then
                first_date = CDate(temp_item.mDate)
            ElseIf Cells(1, 1) Like "weekly*" Then
                first_date = CDate(Format(Now, "yyyy-mm-dd"))
            End If
            
            Set temp_item = Nothing
            Set temp_item = item_collection.item(item_collection.COUNT)
            last_date = CDate(temp_item.mDate)
        Else
            first_date = CDate(Format(Now, "yyyy-mm-dd"))
            last_date = CDate(Format(Now, "yyyy-mm-dd"))
        End If
        
        
        ' ten kod jest wysoce upierdliwy i do tego nie dziala za bardzo tak jak trzeba wiec narazie sie tego pozbede
        'If limit > last_date Then
        '    limit = last_date
        'End If
        
        sth.progress_increase
        ' open po400
        i_mgo.sendKeys "<Clear>"
        i_mgo.sendKeys "ms9po400 <Enter>"
        i_mgo.pMS9PO400.plt = IIteracja_PlantName
        i_mgo.pMS9PO400.pn = IIteracja_Partnum
        i_mgo.pMS9PO400.KANBAN = ""
        i_mgo.sendKeys "<Enter>"
        
        
        sth.progress_increase
        ' tutaj petla zapelniajaca nyrt
        ' ===================================================================
        
        Do
            For x = 0 To 7
                
                m_str = CStr(i_mgo.pMS9PO400.EDA(Int(x)))
                If m_str <> "______" Then
                
                    If CDate(i_mgo.convertToDateFromMS9PO400Date(i_mgo.pMS9PO400.EDA(Int(x)))) <= CDate(Format(limit_del, "yyyy-mm-dd")) Then
                    If CDate(i_mgo.convertToDateFromMS9PO400Date(i_mgo.pMS9PO400.SDATE(Int(x)))) <= CDate(Format(limit, "yyyy-mm-dd")) Then
                
                        txt_time = Left(i_mgo.pMS9PO400.ETA(Int(x)), 2) & ":" & Right(i_mgo.pMS9PO400.ETA(Int(x)), 2)
                        If (txt_time = "__:__") Or (txt_time = "24:00") Then
                            txt_time = "00:00"
                        End If
                        
                        Set temp_trans = Nothing
                        Set temp_trans = New ASN
                        On Error Resume Next
                        temp_trans.mDeliveryDate = i_mgo.convertToDateFromMS9PO400Date(i_mgo.pMS9PO400.EDA(Int(x)))
                        On Error Resume Next
                        temp_trans.mDeliveryTime = CDate(Format(txt_time, "hh:mm"))
                        temp_trans.mNotYetReceived = True
                        temp_trans.mName = i_mgo.pMS9PO400.SID(Int(x))
                        temp_trans.mTRLR = i_mgo.pMS9PO400.TRLR(Int(x))
                        temp_trans.mST = i_mgo.pMS9PO400.st(Int(x))
                        temp_trans.mPickupDate = i_mgo.convertToDateFromMS9PO400Date(CStr(i_mgo.pMS9PO400.SDATE(Int(x))))
                        temp_trans.mQty = CLng(i_mgo.pMS9PO400.Qty(Int(x)))
                        temp_trans.mRoute = i_mgo.pMS9PO400.route(Int(x))
                        
                        temp_trans.mOnPop = True
                        
                        Dim tmp_lng_date As Date
                        tmp_lng_date = CDate(CStr(temp_trans.mDeliveryDate) & " " & CStr(temp_trans.mDeliveryTime))
                        
                        If (first_date > tmp_lng_date) Or (last_date < tmp_lng_date) Then
                            temp_trans.mOutOfScope = True
                        End If
                        
                        If last_date >= tmp_lng_date Then
                            not_yet_received_collection.Add temp_trans
                        End If
                        
                    End If
                    End If
                End If
            Next x
            
            If (CStr(i_mgo.getString(22, 2, 5)) = "I5487") Or (CStr(i_mgo.getString(22, 2, 5)) = "I6155") Or (CStr(i_mgo.getString(22, 2, 5)) = "I6293") Then
                Exit Do
            End If
            
            i_mgo.sendKeys "<pf8>"
            is_f8_pressed = True
        Loop While True
        
        sth.progress_increase
        
        is_f8_pressed = False
        
        ' ===================================================================
End Sub

Private Sub p_tworzenie_tranzytow_on_zk7ppum0(ByRef sth As StatusHandler, ByRef ir As Integer, ByVal IIteracja_PlantName As String, _
    ByVal IIteracja_Partnum As String, ByRef i_mgo As MGO, ByRef item As IItem, ByRef limit As Date, ByRef limit_del As Date, _
    ByRef txt_time As String, ByRef item_collection As Collection, _
    ByRef not_yet_received_collection As Collection, ByRef is_f8_pressed As Boolean, Optional poczatek_dnia_prod As Date)
    
        check_past = Int(Sheets("register").Range("History"))
        
        
        ' przygotowanie warunków brzegowych dla obliczeñ
        ' pierwsza data w kolekcji itemow
        'Dim first_date As Date
        ' ostatnia data w kolekcji itemow
        'Dim last_date As Date
        ' temp item
        Dim temp_item As IItem
        ' handler pod tranzyt
        Dim temp_trans As ITransit
        
        
        ' w zaleznosci od kolekcji itemow ustawiamy warunki brzegowe
        If item_collection.COUNT > 0 Then
            Set temp_item = Nothing
            Set temp_item = item_collection.item(1)
            ' tutaj pogrzebany pies poniewaz caly czas patrzylem na niedziele stad
            ' warunek brzegowy byl zawsze przesuniety o tydzien do przodu
            ' oczywiscie inny warunek na daily inny na weekly stad dodatkowy podwojny if ifelse
            If Cells(1, 1) Like "daily*" Then
                first_date = CDate(temp_item.mDate)
            ElseIf Cells(1, 1) Like "weekly*" Then
                first_date = CDate(Format(Now, "yyyy-mm-dd"))
            End If
            
            Set temp_item = Nothing
            Set temp_item = item_collection.item(item_collection.COUNT)
            last_date = CDate(temp_item.mDate)
        Else
            first_date = CDate(Format(Now, "yyyy-mm-dd"))
            last_date = CDate(Format(Now, "yyyy-mm-dd"))
        End If
        
        'If limit > last_date Then
        '    limit = last_date
        'End If
        
        sth.progress_increase
        
        
        ' jednak ten licznik sluzy przede wszystkim sprawdzeniu czy bylo wejscie na ppud0 a jesli bylo to zeby sie nie pomyslic
        'i zbyt duzo nie klikac w to f8
        licznik_f8 = 0
        
        Do
            For x = 0 To 11
            
            
                Debug.Assert CStr(i_mgo.pZK7PPUM0.PUS_NBR(Int(x))) <> "ZAS192067"
                ' Debug.Print CStr(i_mgo.pZK7PPUM0.PUS_NBR(Int(x)))
            
                If i_mgo.pZK7PPUM0.plt(Int(x)) <> "" Then
                    If CDate(i_mgo.convertToDateFromZK7PPUM0Date(i_mgo.pZK7PPUM0.DEL_DATE(Int(x)))) <= CDate(Format(limit_del, "yyyy-mm-dd")) Then
                        If CDate(i_mgo.convertToDateFromZK7PPUM0Date(i_mgo.pZK7PPUM0.PU_DATE(Int(x)))) <= CDate(Format(limit, "yyyy-mm-dd")) Then
                    
                            txt_time = take_time_on_zk7ppum0(Int(x), i_mgo)
                            
                            
                            
                            Set temp_trans = Nothing
                            Set temp_trans = New pus
                            On Error Resume Next
                            temp_trans.mDeliveryDate = i_mgo.convertToDateFromZK7PPUM0Date(i_mgo.pZK7PPUM0.DEL_DATE(Int(x)))
                            On Error Resume Next
                            temp_trans.mDeliveryTime = CDate(Format(CStr(txt_time), "hh:mm"))
                            temp_trans.mName = CStr(i_mgo.pZK7PPUM0.PUS_NBR(Int(x)))
                            temp_trans.mPickupDate = i_mgo.convertToDateFromZK7PPUM0Date(i_mgo.pZK7PPUM0.PU_DATE(Int(x)))
                            
                            ' Debug.Assert Not temp_trans.mName = "EIS101076"
                            
                            If CStr(Format(temp_trans.mPickupDate, "yyyy-mm-dd")) = CStr(Format(Now, "yyyy-mm-dd")) Then
                                
                                ' temp_trans.mTTLessThan24h = True
                                ' sciaganie pus time z ppud0
                                ' funkcja pickup_time_z_ppud0
                                pu_time = pickup_time_z_ppud0(i_mgo, CStr(i_mgo.pZK7PPUM0.PUS_NBR(Int(x))))
                                
                                
                                ' go back to current view from ppus0
                                x = go_back_to_current_view_in_ppus0(i_mgo, check_past, IIteracja_PlantName, IIteracja_Partnum, temp_trans.mName, Int(licznik_f8))
                            Else
                                pu_time = "00:00"
                            End If
                            
                            'temp_trans.mPickupTime = CDate(Format(CStr(txt_time), "00:00"))
                            temp_trans.mPickupTime = CDate(Format(CStr(pu_time), "hh:mm"))
                            temp_trans.mQty = i_mgo.pZK7PPUM0.PART_QTY(Int(x))
                            temp_trans.mRoute = i_mgo.pZK7PPUM0.route(Int(x))
                            ' domyslnie true na not yet received
                            ' domyslnie false na out of the scope
                            
                            temp_trans.mOnPop = False
                            
                            Dim tmp_lng_date As Date
                            tmp_lng_date = CDate(CStr(temp_trans.mDeliveryDate) & " " & CStr(temp_trans.mDeliveryTime))
                            
                            If (first_date > tmp_lng_date) Or (last_date < tmp_lng_date) Then
                                temp_trans.mOutOfScope = True
                            End If
                            
                            If last_date >= tmp_lng_date Then
                                not_yet_received_collection.Add temp_trans
                            End If
                        End If
                    End If
                End If
            Next x
            
            If CStr(i_mgo.getString(22, 2, 5)) = "R6086" Then
                Exit Do
            End If
            i_mgo.sendKeys "<pf8>"
            licznik_f8 = licznik_f8 + 1
            is_f8_pressed = True
        Loop While True
        
        is_f8_pressed = False
        
        Dim left_trans As ITransit
        Dim right_trans As ITransit
        Set right_trans = Nothing
        Set left_trans = Nothing
        Set temp_trans = Nothing
        
        sth.progress_increase
        
        ' szybkosc algorytmu O(log_2 N) :D:D:D
        ' ale zwiesza sie niestety
        ' trza wrocic do bubble'a ;/
        ' quicksort_on not_yet_received_collection, 1, not_yet_received_collection.COUNT, right_trans, left_trans, temp_trans
        bubble_sort not_yet_received_collection, right_trans, left_trans, temp_trans
        
        
        sth.progress_increase
        ' mam ustawione teraz chronologicznie ze wzlgedu na delivery date i time pickupy na not yet received collection
        ' ==============================================================================================================
        ' teraz moge sprawdzic na historii ktore sa not yet received badz out of scope
        
        ' zaczynamy od nowa wiec tez zerujmy poprzednie zatruowanie na zk7ppum0
        is_f8_pressed = False
End Sub


Private Function pickup_time_z_ppud0(ByRef i_mgo As MGO, nazwa_pusu As String) As String

    ' Debug.Assert CStr(nazwa_pusu) <> "EIS101076"

    i_mgo.sendKeys "<Clear>"
    i_mgo.sendKeys "zk7ppud0 <Enter>"
    i_mgo.pZK7PPUD0.pus = nazwa_pusu
    i_mgo.sendKeys "<Enter>"
    
    pickup_time_z_ppud0 = Left(CStr(i_mgo.pZK7PPUD0.fst_pickup_time_window), 2) & ":" & Right(CStr(i_mgo.pZK7PPUD0.fst_pickup_time_window), 2)
End Function

Private Function go_back_to_current_view_in_ppus0(i_mgo As MGO, check_past, IIteracja_PlantName As String, IIteracja_Partnum As String, mName As String, pf8 As Integer)
    
    i_mgo.sendKeys "<Clear>"
    i_mgo.sendKeys "zk7ppus0 <Enter>"
    i_mgo.pZK7PPUS0.plt = IIteracja_PlantName
    i_mgo.pZK7PPUS0.pn = IIteracja_Partnum
    i_mgo.pZK7PPUS0.PICKUP_DATE_FROM = i_mgo.convertToMgoLongDate(Now)
    i_mgo.sendKeys "<Enter>"
    i_mgo.waitForMgo
    i_mgo.sendKeys "<pf12>"
    i_mgo.waitForMgo
    
    
    Do
        For x = 0 To 11
            If mName = CStr(i_mgo.pZK7PPUM0.PUS_NBR(Int(x))) Then
                go_back_to_current_view_in_ppus0 = x
                Exit Function
            End If
        Next x
    
                
        If CStr(i_mgo.getString(22, 2, 5)) = "R6086" Then
            go_back_to_current_view_in_ppus0 = 11
            Exit Do
        End If
        i_mgo.sendKeys "<pf8>"
        
    Loop While True
    
    
End Function

Private Sub p_towrzenie_tranzytow_on_ms9ph100(ByRef sth As StatusHandler, ByRef ir As Integer, ByVal IIteracja_PlantName As String, _
    ByVal IIteracja_Partnum As String, ByRef i_mgo As MGO, ByRef item As IItem, ByRef limit As Date, ByRef limit_del As Date, _
    ByRef txt_time As String, ByRef item_collection As Collection, _
    ByRef not_yet_received_collection As Collection, ByRef is_f8_pressed As Boolean, Optional poczatek_dnia_prod As Date)
        
        i_mgo.sendKeys "<Clear>"
        i_mgo.sendKeys "ms9ph100 <Enter>"
        i_mgo.pMS9PH100.plt = IIteracja_PlantName
        i_mgo.pMS9PH100.pn = "        "
        i_mgo.pMS9PH100.pn = IIteracja_Partnum
        i_mgo.pMS9PH100.DATE_FROM = i_mgo.convertToMgoShortDate(Now - Int(check_past))
        i_mgo.pMS9PH100.setDTYP "RECV", "ASN"
        i_mgo.sendKeys "<Enter>"
        
        sth.progress_increase
        
        Dim catch_err As CatchError
        Set catch_err = New CatchError
        catch_err.catch_error_on_ms9ph100_skip_if_not_found CStr(i_mgo.getString(22, 2, 5)), CStr(i_mgo.getString(4, 17, 8))
        If catch_err.exit_from_sub Then
            ' nop on history
        Else
            ' history algorithm
            
            Do
                For x = 0 To 9
                
                    If CStr(i_mgo.pMS9PH100.Qty(Int(x))) <> "" Then
                        If Int(i_mgo.pMS9PH100.Qty(Int(x))) <> 0 Then
                            For y = 1 To not_yet_received_collection.COUNT
                            
                                Set temp_trans = Nothing
                                Set temp_trans = not_yet_received_collection.item(y)
                            
                                If temp_trans.mName = CStr(i_mgo.pMS9PH100.SID(Int(x))) Then
                                    ' temp_trans.mQty = 0
                                    temp_trans.mNotYetReceived = False
                                End If
                            Next y
                        End If
                    End If
                Next x
                
                If CStr(i_mgo.getString(22, 2, 5)) = "R6086" Then
                    Exit Do
                End If
                
                i_mgo.sendKeys "<pf8>"
                is_f8_pressed = True
            Loop While True
        End If
        


End Sub

Private Function take_time_on_zk7ppum0(x As Integer, i_mgo As MGO)
    Dim tmp_str As String
    tmp_str = Left(i_mgo.pZK7PPUM0.END_DEL(Int(x)), 2) & ":" & Right(i_mgo.pZK7PPUM0.END_DEL(Int(x)), 2)
    
    If tmp_str = "  :  " Then
        tmp_str = "00:00"
    ElseIf tmp_str = "24:00" Then
        tmp_str = "23:59"
    End If
    
    take_time_on_zk7ppum0 = tmp_str
End Function

Private Sub bubble_sort(not_yet_received_collection As Collection, _
    Optional ByRef i_trans As ITransit, Optional ByRef left_trans As ITransit, Optional ByRef temp_trans As ITransit)

    
    For x = 1 To not_yet_received_collection.COUNT - 1
        Set left_trans = not_yet_received_collection.item(x)
        Set i_trans = not_yet_received_collection(x + 1)
        
        If CDate(CStr(left_trans.mDeliveryDate) & " " & CStr(left_trans.mDeliveryTime)) > CDate(CStr(i_trans.mDeliveryDate) & " " & CStr(i_trans.mDeliveryTime)) Then
            
            Set temp_trans = Nothing
            Set temp_trans = New pus
            swap_transits temp_trans, left_trans
            swap_transits left_trans, i_trans
            swap_transits i_trans, temp_trans
            x = 1
        End If
    Next x
    
    'For x = 1 To not_yet_received_collection.COUNT
    '    Set left_trans = not_yet_received_collection.item(x)
    '    Debug.Print left_trans.mDeliveryDate & " " & left_trans.mDeliveryTime & " " & left_trans.mName
    'Next x
End Sub

Private Sub swap_transits(ByRef tr1 As ITransit, ByRef tr2 As ITransit)
    tr1.mDeliveryDate = tr2.mDeliveryDate
    tr1.mDeliveryTime = tr2.mDeliveryTime
    tr1.mName = tr2.mName
    tr1.mNotYetReceived = tr2.mNotYetReceived
    tr1.mOutOfScope = tr2.mOutOfScope
    tr1.mPickupDate = tr2.mPickupDate
    tr1.mPickupTime = tr2.mPickupTime
    tr1.mQty = tr2.mQty
    tr1.mRoute = tr2.mRoute
End Sub


Private Sub quicksort_on(not_yet_received_collection As Collection, left_wsk As Integer, right_wsk As Integer, _
    Optional ByRef i_trans As ITransit, Optional ByRef left_trans As ITransit, Optional ByRef temp_trans As ITransit)
    
    
    If left_wsk < right_wsk Then
    
        Dim m As Integer
        m = Int(left_wsk)
    
        Set left_trans = Nothing
        Set left_trans = not_yet_received_collection.item(left_wsk)
        
        For i = left_wsk + 1 To right_wsk
            Set i_trans = Nothing
            Set i_trans = not_yet_received_collection.item(i)
            
            If CDate(CStr(i_trans.mDeliveryDate) & " " & CStr(i_trans.mDeliveryTime)) < CDate(CStr(left_trans.mDeliveryDate) & " " & CStr(left_trans.mDeliveryTime)) Then
                m = m + 1
                swap_objects not_yet_received_collection.item(m), not_yet_received_collection.item(i), temp_trans
            End If
            swap_objects not_yet_received_collection.item(left_wsk), not_yet_received_collection.item(m), temp_trans
            
            quicksort_on not_yet_received_collection, left_wsk, m - 1, i_trans, left_trans, temp_trans
            quicksort_on not_yet_received_collection, m + 1, right_wsk, i_trans, left_trans, temp_trans
        Next i
    End If
End Sub

Private Sub swap_objects(ByRef obj1 As ITransit, ByRef obj2 As ITransit, ByRef tmp As ITransit)
    Set tmp = obj1
    Set obj1 = obj2
    Set obj2 = tmp
End Sub

Private Function check_possibility_2_put_transit_in_this_item(prev_item As String, act_item As String, mgo_date As String, mgo_time As String) As Boolean
    ' MsgBox item_date & " " & mgo_date & " " & mgo_time
    
    If (CDate(prev_item) < CDate(mgo_date & " " & mgo_time)) And (CDate(act_item) >= CDate(mgo_date & " " & mgo_time)) Then
        check_possibility_2_put_transit_in_this_item = True
    End If
End Function
