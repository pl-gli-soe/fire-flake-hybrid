VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiffReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private wrk As Worksheet
Private wrr As Range
Private adr As String
Private kolekcja_manualnych_dopasowan As Collection
Private diff_item As DiffItem
Private szablon As ILayout
Private rep_diff As String
Private sh As StatusHandler
Private diff_on_daily As Boolean


Private var As DiffItem

Private Sub Class_Initialize()

    Set sh = New StatusHandler
    adr = ""

    Set wrk = ActiveSheet
    zrob_adres
    Set wrr = work_range_on()
    sh.init_statusbar wrr.COUNT
    sh.show
    ' od tutaj mamy gotowy obaszar przeszukiwania
    
    Set kolekcja_manualnych_dopasowan = New Collection
    
End Sub

Private Sub Class_Terminate()
    sh.hide
    Set wrk = Nothing
    Set wrr = Nothing
    Set kolekcja_manualnych_dopasowan = Nothing
    Set sh = Nothing
End Sub

Public Sub find_differences()
    
    
    Dim r As Range
    If Not wrr Is Nothing Then
        
        For Each r In wrr
        
        
            ' ok nastapic musi nieco zmiana jesli chodzi o nadawanie kolejnych elementow do kolekcji szczegolnie
            ' jesli wywalilem kolor niebieski i nie zawsze bedzie wiadomo ze jest to na przyklad manualna zmiana
            ' bo nie zawsze jest komment
            ' bedzie trzeba sie odniesc do 5 wiersza raportu w ktorym to widnieje
            If check_if_this_cell_has_sth_in_transit(r, CStr(wrk.Cells(1, 1))) Then
                rep_diff = ""
                add_item_to_collection r, CStr(wrk.Cells(1, 1))
            End If
            
            sh.progress_increase
        Next r
    End If
End Sub


Private Function check_if_this_cell_has_sth_in_transit(ByRef r As Range, jaki_to_raport As String) As Boolean
    check_if_this_cell_has_sth_in_transit = False
    
    
    ' ta niby masakra to bardzo prosty warunek
    ' ktory:
    ' albo masz komentarz albo u gory widnieje napis intransit i sam cell nie jest pusty :) - dosyc proste zdanie po polsku
    ' ale ile lewakowania poprzez warunek if :D :P
    If (jaki_to_raport Like "*daily*") Or (jaki_to_raport Like "*weekly*") Then
        If (Cells(5, r.Column) Like "*transit*") And (r <> "") Then
            check_if_this_cell_has_sth_in_transit = True
        End If
    ElseIf jaki_to_raport Like "*hourly*" Then
        If (Cells(r.row, 8) Like "*PUS*ASN") And (r <> "") Then
            check_if_this_cell_has_sth_in_transit = True
            
        End If
    End If
    
    
End Function


Private Sub add_item_to_collection(ByRef r As Range, n As String)
    
    ' 0 - delivery date 10 z prawej
    ' 1 - delivery time 8 z prawej
    ' 2 - name - usun 6 z lewej
    ' 3 - pickup date 10 z prawej
    ' 4 - qty 4 usun z lewej
    ' 5 - route ususn 7 z lewej
    ' 6 - dash breaks
    ' 7 - puste
    ' 8 - puste ?
    ' troche sie zmienilo od teraz
    
    '1 DeliveryDate: 2014-05-21
    '2 DeliveryTime: 12:30:00
    '3 Name: 864670
    '4 PickupDate: 2014-04-10
    '5 Qty: 7800
    '6 Route: VLU -312
    '7 TRLR: HLXU5092737
    '8 ST: 2
    '9 ----------------------------------------
    
    
    ' ok trzeba rozwazyc opcji kilka teraz
    ' daily - manulne z komentarzem i bez
    ' weekly - manualne z komentarzem i bez - tutaj problematyczne daty delivery
    ' hourly - manualne z komentarzem i bez - tu jeszcze pol biedy
    
    
    ' to jest dosyc proste sprawdzamy sobie komorke i tworzymy na jej podstawie DiffItemy
    ' ktore potem pakujemy do kolekcji kolekcja_manualnych_dopasowan
    
    
    If n Like "daily*" Then
        preprare_items r, n, DAILY
    ElseIf n Like "weekly*" Then
        preprare_items r, n, WEEKLY
    ElseIf n Like "hourly*" Then
        preprare_items r, n, HOURLY
    End If
        
    
    
    
End Sub


Private Sub preprare_items(ByRef r As Range, n As String, typ As E_REPORT_TYPE)

    Dim komorka As Range

    Dim ile_transportow As Integer
    Dim suma_real_tranzytow As Long
    suma_real_tranzytow = 0
    ile_transportow = podlicz_wystepowanie_oraz_zsumuj_real_tranzyty(r, suma_real_tranzytow) ' przelicza suma_real_tranzytow
    
    ' funkcja powyzej zrobila dwie rzeczy
    ' jedna std zwrocila ile transportow jest
    ' druga jako side effect oddala suma real transitow :D
    '0 1 DeliveryDate: 2014-05-21
    '1 2 DeliveryTime: 12:30:00
    '2 3 Name: 864670
    '3 4 PickupDate: 2014-04-10
    '4 5 Qty: 7800
    '5 6 Route: VLU -312
    '6 7 TRLR: HLXU5092737
    '7 8 ST: 2
    '8 9 ---------------------------------------- dash breaks
    
    
    ' najpierw zrobimy diffitemy tylko i wylacznie pod same komentarze
    If ile_transportow > 0 Then
        Dim str_arr() As String
        str_arr = Split(r.Comment.Text, Chr(10))
        Dim inc_delivery_appears As Integer
        For t = 1 To ile_transportow
        
            Set diff_item = Nothing
            Set diff_item = New DiffItem
            inc_delivery_appears = 0
            
            ' po kolei wszystkie komentowane real PUS/ASNy
            
            For x = LBound(str_arr) To UBound(str_arr)
                If str_arr(x) Like "*Delivery*Date*" Then
                    inc_delivery_appears = inc_delivery_appears + 1
                    
                    If inc_delivery_appears = t Then
                        diff_item.delivery_date = CDate(Right(str_arr(x), 10))
                        diff_item.pickup_date = CDate(Right(str_arr(x + 3), 10))
                        ' diff_item.differenceQty = r.Value
                        ' diff_item.getExactValue = r.Value
                        diff_item.istransit = True
                        ' diff_item.manualAdjustment = r.Value
                        diff_item.Name = Right(str_arr(x + 2), Len(str_arr(x + 2)) - 6)
                        
                        If typ = DAILY Or typ = WEEKLY Then
                            diff_item.plt = Cells(r.row, 3).Value ' daily & weekly
                            diff_item.pn = Cells(r.row, 2).Value ' daily & weekly
                            
                            If typ = DAILY Then
                                diff_item.fst_runout = Left(Cells(r.row, 10), 10) ' daily & weekly fst runout
                            Else
                                diff_item.fst_runout = Cells(r.row, 10) ' daily & weekly fst runout
                            End If
                        ElseIf typ = HOURLY Then
                            diff_item.plt = Cells(r.row - 2, 3).Value
                            diff_item.pn = Cells(r.row - 3, 3).Value
                            diff_item.fst_runout = Cells(r.row, 5)
                        End If
                        
                        diff_item.Qty = Right(str_arr(x + 4), Len(str_arr(x + 4)) - 4)
                        diff_item.route = Right(str_arr(x + 5), Len(str_arr(x + 5)) - 7)
                        ' trlr
                        diff_item.TRLR = Right(str_arr(x + 6), Len(str_arr(x + 6)) - 6)
                        ' st
                        diff_item.st = Right(str_arr(x + 7), Len(str_arr(x + 7)) - 4)
                        diff_item.regular = True
                        diff_item.sum_of_qty_in_parent_cell = r.Value
                        diff_item.sum_of_regular_transit = suma_real_tranzytow
                        diff_item.set_exact_value
                        
                        kolekcja_manualnych_dopasowan.Add diff_item
                        
                        Set diff_item = Nothing
                        
                    End If
                End If
            Next x
        
        Next t
    End If
    
    
    ' dodatkowy item poza komentarzem
    ' pamietaj ze taki item moze byc ujemny z racji perspektywy zmniejszania manula w stosunku do reala
    If suma_real_tranzytow <> r.Value Then
        Set diff_item = New DiffItem

        ' diff_item.delivery_date = Now ' CDate(Right(str_arr(x), 10))
        ' tutaj musi byc bardziej skomplikowana opcja poniewaz nie ma komentarza i latwy sposob nie idzie wyciagnac info o delivery date
        ' oinaczej wylada na kazdym raporcie
        
        If typ = DAILY Then
            diff_item.delivery_date = CDate(Left(Cells(4, r.Column - 1), 10))
        ElseIf typ = HOURLY Then
            ' diff_item.delivery_date = CDate(Now)
            diff_item.delivery_date = CDate(r.Offset(-3, 0))
        ElseIf typ = WEEKLY Then
            ' komentarze ponizej pokazuja jak dolny kod mozna zwinac / rozwinac
            ' Set komorka = Cells(4, r.Column - 1)
            ' i nie ma co dwa razy sprawdzac poniewaz komentarz jest domyslnie w tym miejscu i jesli go nie bedzie to cos wczesniej zmascilem
            ' cmnt_txt = CStr(komorka.Comment.Text)
            ' pierwsza_data = Left(cmnt_txt, 10)
            ' diff_item.delivery_date = CDate(Left(cmnt_txt, 10))
            diff_item.delivery_date = CDate(Left(CStr(Cells(4, r.Column - 1).Comment.Text), 10))
        End If
        
        
        ' diff_item.pickup_date =  ' CDate(Right(str_arr(x + 3), 10))
        ' diff_item.differenceQty = r.Value
        ' diff_item.getExactValue = r.Value
        diff_item.istransit = True
        ' diff_item.manualAdjustment = r.Value
        diff_item.Name = "manual"
        
        If typ = DAILY Or typ = WEEKLY Then
            diff_item.plt = Cells(r.row, 3).Value ' daily & weekly
            diff_item.pn = Cells(r.row, 2).Value ' daily & weekly
            
            If typ = DAILY Then
                diff_item.fst_runout = Left(Cells(r.row, 10), 10) ' daily & weekly fst runout
            Else
                diff_item.fst_runout = Cells(r.row, 10) ' daily & weekly fst runout
            End If
        ElseIf typ = HOURLY Then
            diff_item.plt = Cells(r.row - 2, 3).Value
            diff_item.pn = Cells(r.row - 3, 3).Value
            diff_item.fst_runout = Cells(r.row, 5) ' daily & weekly fst runout
        End If
        
        ' diff_item.Qty = r.Value
        diff_item.route = "manual"
        ' trlr
        diff_item.TRLR = "manual"
        ' st
        diff_item.st = ""
        diff_item.regular = False
        diff_item.sum_of_qty_in_parent_cell = r.Value
        diff_item.sum_of_regular_transit = suma_real_tranzytow
        diff_item.set_exact_value
        diff_item.Qty = diff_item.manualAdjustment
        
        kolekcja_manualnych_dopasowan.Add diff_item
        
        Set diff_item = Nothing
    End If

End Sub

Private Function podlicz_wystepowanie_oraz_zsumuj_real_tranzyty(ByRef r As Range, ByRef suma As Long) As Integer

    podlicz_wystepowanie = 0
    
    ' tablica do przechowywania kolejnych linii komentarzy
    Dim str_arr() As String
    If Not r.Comment Is Nothing Then
        str_arr = Split(r.Comment.Text, Chr(10))
        For x = LBound(str_arr) To UBound(str_arr)
            If str_arr(x) Like "*Qty*" Then
                ' podliczam sobie kolejne wystepwanie ciagu znakow *Qty*
                podlicz_wystepowanie = podlicz_wystepowanie + 1
                ' suma = suma + liczba
                suma = suma + CLng(Right(str_arr(x), Len(str_arr(x)) - 4))
            End If
        Next x
    Else
        suma = 0
    End If
    podlicz_wystepowanie_oraz_zsumuj_real_tranzyty = podlicz_wystepowanie
    
End Function

Public Sub create_diff_report()

    If kolekcja_manualnych_dopasowan.COUNT > 0 Then
        Set szablon = New DailyLayout
        
        szablon.InitLayout
        
        Dim tmp As String
        tmp = "difference report " & CStr(Now) & " on " & CStr(ActiveSheet.Name) & " " & CStr(wrk.Range("a1"))
        tmp = Replace(tmp, "daily", "D")
        tmp = Replace(tmp, "hourly", "H")
        Range("a1") = tmp
        
        
        
        
        ' A B C D E F G H I J K L M N O P Q R S T U V W X Y Z
        Range("C6") = "plant"
        Range("D6") = "part number"
        Range("E6") = "name"
        Range("F6") = "delivery date"
        Range("G6") = "pickup date"
        Range("H6") = "qty for this transport"
        Range("I6") = "value in cell"
        Range("J6") = "difference"
        Range("K6") = "valid change"
        Range("L6") = "regular transport"
        Range("M6") = "exact value"
        Range("N6") = "route"
        Range("O6") = "TRLR"
        Range("P6") = "ST"
        Range("Q6") = "FST RUNOUT"
        
        Range("C6:Q6").Interior.Color = RGB(200, 200, 200)
        
        
        Dim r As Range
        
        Dim color_toggle As Boolean
        color_toggle = False
        
        ile_odjac = 0
        Set r = Range("C7")
        For Each diff_item In kolekcja_manualnych_dopasowan

            If diff_item.getExactValue <> 0 Then
            
                ' tutaj mala zagrywka logiczna poniewaz latwo jest sprawdzic calkowite porownanie ale
                ' w druga strone juz gorzej zatem z musu za kazdym razem od razu robie toggle
                ' i tylko w przypadku kiedy jednak okazuje sie ze to ta sama komorka robie z powrotem
                ' zostawiajac ten sam kolor
                toggle color_toggle
                If CStr(r.Offset(-1, 0)) = diff_item.plt Then
                    If CStr(r.Offset(-1, 1)) = diff_item.pn Then
                        If CStr(r.Offset(-1, 3) = CStr(Format(diff_item.delivery_date, "yyyy-mm-dd"))) Then
                            toggle color_toggle
                        End If
                    End If
                End If
            
                r = diff_item.plt
                r.Offset(0, 1) = diff_item.pn
                r.Offset(0, 2) = diff_item.Name
                r.Offset(0, 3) = diff_item.delivery_date
                If diff_item.regular Then
                    r.Offset(0, 4) = diff_item.pickup_date
                ElseIf Not diff_item.regular Then
                    r.Offset(0, 4).Font.Color = RGB(200, 0, 0)
                    r.Offset(0, 4) = "yyyy-mm-dd"
                    r.Offset(0, 4).Font.Bold = True
                End If
                r.Offset(0, 5) = diff_item.Qty
                r.Offset(0, 6) = diff_item.sum_of_qty_in_parent_cell
                r.Offset(0, 7) = diff_item.manualAdjustment
                r.Offset(0, 8) = diff_item.istransit
                If diff_item.regular Then
                    r.Offset(0, 9) = "regular"
                ElseIf Not diff_item.regular Then
                    r.Offset(0, 9) = "manual"
                End If
                r.Offset(0, 10) = diff_item.getExactValue
                
                If color_toggle Then
                    Range(r, r.Offset(0, 14)).Interior.Color = RGB(200, 200, 200)
                    Range(r, r.Offset(0, 14)).Font.Bold = True
                ElseIf Not color_toggle Then
                    Range(r, r.Offset(0, 14)).Interior.Color = RGB(255, 255, 255)
                    Range(r, r.Offset(0, 14)).Font.Bold = True
                End If
                
                r.Offset(0, 11) = diff_item.route
                r.Offset(0, 12) = diff_item.TRLR
                r.Offset(0, 13) = diff_item.st
                r.Offset(0, 14) = diff_item.fst_runout
                
                
                Set r = r.Offset(1, 0)
            ElseIf diff_item.getExactValue = 0 Then
                ile_odjac = ile_odjac + 1
            End If
            
            
        Next diff_item
        
        wierszy = kolekcja_manualnych_dopasowan.COUNT
        adres_dla_tabeli = "C6:Q" & CStr(6 + (wierszy - ile_odjac)) ' plus jeden dla labelek
        szablon.FillThinGridLines Range(CStr(adres_dla_tabeli)), RGB(0, 0, 0)

        Range("C6:Q6").AutoFilter
        Columns("C:Z").AutoFit
    End If

End Sub

Private Sub toggle(ByRef arg As Boolean)
    If arg = True Then
        arg = False
    ElseIf arg = False Then
        arg = True
    End If
End Sub

Private Function work_range_on() As Range
    ' adr = ""
    Set work_range_on = wrk.Range(adr)
End Function

Private Sub zrob_adres()
    If (wrk.Range("a1") Like "daily*") Or (wrk.Range("a1") Like "weekly*") Then
        adr = "Q6"
        adr = adr & ":" & chrx(Sheets("register").Range("lastColumn")) & CStr(Sheets("register").Range("lastRow"))
        diff_on_daily = True
    ElseIf wrk.Range("a1") Like "hourly*" Then
    
        adr = "I2"
        adr = adr & ":" & chrx(Sheets("register").Range("lastColumnHourly")) & CStr(Sheets("register").Range("lastRowHourly"))
        diff_on_daily = False
    End If
End Sub
