VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HourlyIteration"
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
' ClassModule DailyIteration
' ==========================
'
' this class generating data for each set of data (row)
' RQM, TRANS(RQM), EBAL(EBAL-1, TRANS, RQM) <- from the main proc PrepareItemCollection
' the main data from ms9pop00 <- proc DownloadGeneralData
'
' there's no mgo instance varibale- every time it's working as arg
' + type TworzenieTranzytow for algorithm on PUSes and ASNs
'
'
' raw hourly no extended version from daily
' hourly from daily is on Rozwijanie*

Implements IIteracja

' wiersz dla pierwszego elementu wiec dla hourly
' part num bedzie Cells(ir, 3)
' plt name bedzie Cells(ir+1,3)
Private production_day_begining As Date
Private ir As Integer
Private item As IItem
Private item_collection As Collection
Private not_yet_received_collection As Collection
Private tworzenie_tranzytow As TworzenieTranzytow

' ciezko stwierdzic czy to dobry pomysl, ale potrzebuje kontenera danych na bank
Private instance_in_iteration_bank As Long

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
    instance_in_iteration_bank = CLng(RHS)
End Property

Private Property Get IIteracja_Bank() As Long
    IIteracja_Bank = CLng(instance_in_iteration_bank)
End Property

Private Property Let IIteracja_CBAL(RHS As String)
    Cells(ir + 2, 3) = CStr(RHS)
End Property

Private Property Get IIteracja_CBAL() As String
    IIteracja_CBAL = CStr(Cells(ir + 2, 3))
End Property

Private Property Let IIteracja_Cmnt1(RHS As String)
    Cells(ir + 3, 7) = CStr(RHS)
End Property

Private Property Get IIteracja_Cmnt1() As String
    IIteracja_Cmnt1 = CStr(Cells(ir + 3, 7))
End Property

Private Property Let IIteracja_Cmnt2(RHS As String)
    Cells(ir + 4, 7) = CStr(RHS)
End Property

Private Property Get IIteracja_Cmnt2() As String
    IIteracja_Cmnt1 = CStr(Cells(ir + 4, 7))
End Property

Private Property Let IIteracja_Deck(RHS As String)
    Cells(ir + 1, 5) = CStr(RHS)
End Property

Private Property Get IIteracja_Deck() As String
    IIteracja_Deck = CStr(Cells(ir + 1, 5))
End Property

Private Property Get IIteracja_DOH() As String
    Cells(ir + 2, 5) = CStr(RHS)
End Property

Private Property Let IIteracja_DOH(RHS As String)
    Cells(ir + 2, 5) = CStr(RHS)
End Property

Private Sub IIteracja_DownloadGeneralData(i_mgo As MGO, ByRef sh As StatusHandler, limit As Date, limit_del As Date)
        i_mgo.sendKeys "<Clear>"
        i_mgo.sendKeys "ms9pop00 <Enter>"
        i_mgo.pMS9POP00.DS = "6"
        i_mgo.pMS9POP00.pn = IIteracja_Partnum
        i_mgo.pMS9POP00.plt = IIteracja_PlantName
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
        IIteracja_SheetDef = "hourly " & CStr(ThisWorkbook.Sheets("register").Range("redpink")) & " " & CStr(Now) & " limit on scheduled PUSes: " & limit & " and del limit: " & limit_del & " " & ActiveSheet.Name
        
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
        
        instance_in_iteration_bank = CLng(pop.Bank)
        Cells(IIteracja_IterationRow, 3).AddComment CStr(txt)
        
        Cells(IIteracja_IterationRow, 3).Comment.Shape.Height = 160
        Cells(IIteracja_IterationRow, 3).Comment.Shape.Width = 300
        
        sh.progress_increase
End Sub

Private Property Let IIteracja_Duns(RHS As String)
    Cells(ir, 5) = CStr(RHS)
End Property

Private Property Get IIteracja_Duns() As String
    IIteracja_Duns = CStr(Cells(ir, 5))
End Property

Private Property Let IIteracja_FirstRunout(RHS As String)
    Cells(ir + 3, 5) = CStr(RHS)
End Property

Private Property Get IIteracja_FirstRunout() As String
    iteracja_firstrunout = CStr(Cells(ir + 3, 5))
End Property

Private Property Let IIteracja_ItemCollection(RHS As Collection)
    Set item_collection = RHS
End Property

Private Property Get IIteracja_ItemCollection() As Collection
    Set IIteracja_ItemCollection = item_collection
End Property

Private Property Let IIteracja_IterationRow(RHS As Integer)
    ir = Int(RHS)
End Property

Private Property Get IIteracja_IterationRow() As Integer
    IIteracja_IterationRow = Int(ir)
End Property

Private Property Let IIteracja_MISC(RHS As String)
    Cells(ir, 7) = CStr(RHS)
End Property

Private Property Get IIteracja_MISC() As String
    IIteracja_MISC = CStr(Cells(ir, 7))
End Property

Private Property Let IIteracja_NotyetReceived(RHS As Range)
    Cells(ir + 1, 7) = CStr(RHS)
End Property

Private Property Get IIteracja_NotyetReceived() As Range
    Set IIteracja_NotyetReceived = Cells(ir + 1, 7)
End Property

Private Property Get IIteracja_NotYetReceivedCollection() As Collection
    Set IIteracja_NotYetReceivedCollection = not_yet_received_collection
End Property

Private Property Let IIteracja_PartName(RHS As String)
    Cells(ir + 3, 3) = CStr(RHS)
End Property

Private Property Get IIteracja_PartName() As String
    IIteracja_PartName = CStr(Cells(ir + 1, 3))
End Property

Private Property Let IIteracja_Partnum(RHS As String)
    Cells(ir, 3) = CStr(RHS)
End Property

Private Property Get IIteracja_Partnum() As String
    IIteracja_Partnum = CStr(Cells(ir, 3))
End Property

Private Property Let IIteracja_PlantName(RHS As String)
    Cells(ir + 1, 3) = CStr(RHS)
End Property

Private Property Get IIteracja_PlantName() As String
    IIteracja_PlantName = CStr(Cells(ir + 1, 3))
End Property

Private Sub IIteracja_PrepareItemCollection(i_mgo As MGO, item As IItem, limit As Date, limit_del As Date, Optional hourly_extension_day As Date, Optional ByRef sh As StatusHandler, Optional ByRef instance_tokenizator As Tokenizator)

    ' rusztowanie powinno byc zbudowane na bazie rqms
    ' i dopiero na polozyc bedzie mozna PUSy i ASNy
    ' ebal jest tylko i wylacznie formula ktora musi pracowac na zywych danych
    
    ' Dim interwal As Double
    ' interwal = CDbl((1 / 24) * Sheets("register").Range("itemInterval"))
    Dim first_uloc_ready As Boolean
    Dim index_for_coll As Long
    index_for_coll = 1
    
    Dim is_f8_pressed As Boolean
    is_f8_pressed = False
    Dim i As ITransit
    ' zmienna dla zk7ppus0 w przypadku cofania sie z sprawdzaniem PUSow
    Dim check_past As Integer
    ' tekst na czas delivery troche poza enkapsulacja ale niechaj i zdarzy mi sie taki babuch
    ' taki mark ze to jest moj kod :P
    Dim txt_time As String
    Dim tempdate As Date
    Dim e_item As IItem
    tempdate = CDate(Format(Now, "yyyy-mm-dd"))
    tempdate = CDate(CStr(tempdate) & " 00:00")
    
    ' RQM
    ' ===================================================================================
    
    i_mgo.sendKeys "<Clear>"
    i_mgo.sendKeys "zk7pprul <Enter>"
    i_mgo.pZK7PPRUL.pn = IIteracja_Partnum
    i_mgo.pZK7PPRUL.plt = IIteracja_PlantName
    i_mgo.pZK7PPRUL.KANBAN = ""
    i_mgo.sendKeys "<Enter>"
    
    
    ' pobranie uloc'ów
    ' -------------------------
    i_mgo.pZK7PPRUL.p_uloc = Nothing
    i_mgo.pZK7PPRUL.p_uloc = New Collection
    i_mgo.pZK7PPRUL.fill_uloc_collection i_mgo
    ' -------------------------
    
    
    sh.progress_increase
    
    first_uloc_ready = False
    For Each u In i_mgo.pZK7PPRUL.p_uloc
    
        ' fill
        i_mgo.sendKeys "<Clear>"
        i_mgo.sendKeys "zk7pprhr <Enter>"
        i_mgo.pZK7PPRHR.pn = IIteracja_Partnum
        i_mgo.pZK7PPRHR.plt = IIteracja_PlantName
        i_mgo.pZK7PPRHR.KANBAN = ""
        i_mgo.pZK7PPRHR.ULOC = CStr(u)
        i_mgo.sendKeys "<Enter>"
        
        
        ' take items
        
        
        
        ' RQMs =========================================================================================
        index_for_coll = 1
        Do
            
            For r = 0 To 3
                For c = 0 To 9
                
                    If i_mgo.pZK7PPRHR.HRQMDATE(Int(r), Int(c)) <> "" Then
                
                
                        Set item = Nothing
                        Set item = New ItemHourly
                    
                    
                    
                        If Not first_uloc_ready Then
                    
                            ' dopiero przygotowanie itemow :)
                            item.mReq = i_mgo.pZK7PPRHR.HQTY(Int(r), Int(c))
                            item.mDate = CDate(i_mgo.convertToDateFromZK7PDRQMDate(i_mgo.pZK7PPRHR.HRQMDATE(Int(r), Int(c))) & _
                                " " & CStr(i_mgo.pZK7PPRHR.HRQMTIME(Int(r), Int(c))))
                                    
                            item_collection.Add item
                        Else
                                Set e_item = item_collection.item(Int(index_for_coll))
                                If e_item.mDate = CDate(i_mgo.convertToDateFromZK7PDRQMDate(i_mgo.pZK7PPRHR.HRQMDATE(Int(r), Int(c))) & _
                                    " " & CStr(i_mgo.pZK7PPRHR.HRQMTIME(Int(r), Int(c)))) Then
                                
                                        e_item.mReq = e_item.mReq + i_mgo.pZK7PPRHR.HQTY(Int(r), Int(c))
                                        index_for_coll = index_for_coll + 1
                                End If
                        End If
                    Else
                        Exit Do
                    End If
                    
                Next c
            Next r
            
            i_mgo.sendKeys "<pf8>"
            
            If CStr(i_mgo.getString(22, 2, 5)) = "R6123" Then
                Exit Do
            End If
        
        Loop While True
        first_uloc_ready = True
    Next u
    ' ==============================================================================================
    
    
    sh.progress_increase
    
    ' TRANS
    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    tworzenie_tranzytow.towrzenie_tranzytow_2 sh, ir, IIteracja_PlantName, IIteracja_Partnum, i_mgo, item, limit, limit_del, txt_time, item_collection, not_yet_received_collection, is_f8_pressed
    ' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    
    sh.progress_increase
    
    ' EBAL
    ' --------------------------------------------------------------------
    Dim kol As Long
    Dim cl As Range
    ' zmienna z lokalnej na instancyjna
    ' Dim ef As EbalFormulaz    Set ef = New EbalFormula
    kol = ThisWorkbook.Sheets("register").Range("firstColumnHourly")
    
    For Each item In item_collection
    
        Set ef = New EbalFormula
        If kol = 9 Then
            item.mEbal = ef.hourly_fill_first_day()
        Else
            item.mEbal = ef.hourly_fill_other_days()
        End If
        Set ef = Nothing
        kol = kol + 1
        
    Next item
    ' --------------------------------------------------------------------
    
    
    
    ' wyrzuc dane
    ' =================================================================================
    
    ' z powodu tego ze w PUSach moga byc pickupy juz received nie moge sie posluzyc
    ' tylko i wylacznie countem dla ich kolekcji potrzebuje dodatkowego iteratora
    ' sprawdzajacego faktyczny stan PUSow na not yet received
    Dim var_helper As Integer
    Dim first_interval As Date
    Dim ih As IItem
    
    If item_collection.COUNT > 0 Then
        Set ih = item_collection.item(1)
        first_interval = CDate(CStr(ih.mDate))
    Else
        first_interval = Now
    End If
    var_helper = 0
    Dim cmnt As String
    kol = ThisWorkbook.Sheets("register").Range("firstColumn")
    
    cmnt = ""
    For Each nyrtrans In not_yet_received_collection
        If nyrtrans.mNotYetReceived = True Then
            If nyrtrans.mOutOfScope = True Then
                If CDate(CStr(nyrtrans.mDeliveryDate) & " " & CStr(nyrtrans.mDeliveryTime)) < first_interval Then
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
        
    If item_collection.COUNT > 0 Then
    
    
        iteracja = 0
        For Each e_item In item_collection
        
            Cells(ir, 9 + iteracja) = Format(e_item.mDate, "yyyy-mm-dd")
            Cells(ir + 1, 9 + iteracja) = Format(e_item.mDate, "hh:mm")
            Cells(ir + 2, 9 + iteracja) = e_item.mReq
            ' Cells(ir + 3, 9 + iteracja)
            
            var_helper = 0
            
            If e_item.mTransCollection.COUNT > 0 Then
                cmnt = ""
                For Each i In e_item.mTransCollection
                
                    If (i.mNotYetReceived = True) And (i.mQty <> 0) Then
                        Cells(ir + 3, 9 + iteracja) = Cells(ir + 3, 9 + iteracja) + i.mQty
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
                    Cells(ir + 3, 9 + iteracja).AddComment CStr(cmnt)
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
                    Cells(ir + 3, 9 + iteracja).Comment.Shape.Height = (11 * var_helper) * 10
                    ' heurystycznie :P
                    Cells(ir + 3, 9 + iteracja).Comment.Shape.Width = 200
                  End If
            End If
            
            
            Cells(ir + 4, 9 + iteracja) = e_item.mEbal
            iteracja = iteracja + 1
        Next e_item
        ' =================================================================================
        
        If ThisWorkbook.Sheets("register").Range("ActualQtyIntervals") < iteracja Then
            ThisWorkbook.Sheets("register").Range("ActualQtyIntervals") = iteracja
        End If
    
    End If
    

    
End Sub

Private Property Let IIteracja_ProdDayBegin(RHS As Date)
End Property

Private Property Get IIteracja_ProdDayBegin() As Date
End Property

Private Property Let IIteracja_Qhold(RHS As String)
    Cells(ir + 4, 5) = CStr(RHS)
End Property

Private Property Get IIteracja_Qhold() As String
    IIteracja_Qhold = CStr(Cells(ir + 4, 5))
End Property

Private Property Let IIteracja_SheetDef(RHS As String)
    Cells(1, 1) = CStr(RHS)
End Property

Private Property Get IIteracja_SheetDef() As String
    IIteracja_SheetDef = CStr(Cells(1, 1))
End Property

Private Property Let IIteracja_StdPack(RHS As String)
    Cells(ir + 2, 7) = CStr(RHS)
End Property

Private Property Get IIteracja_StdPack() As String
    IIteracja_StdPack = CStr(Cells(ir + 2, 7))
End Property

Private Property Let IIteracja_Supplier(RHS As String)
    Cells(ir + 4, 3) = CStr(RHS)
End Property

Private Property Get IIteracja_Supplier() As String
    IIteracja_Supplier = CStr(Cells(ir + 4, 3))
End Property
