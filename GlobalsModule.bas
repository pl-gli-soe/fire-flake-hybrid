Attribute VB_Name = "GlobalsModule"
' delay time
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' pobierz dynamiczna library
Private Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)

Global Const LETTERS = 26
Global Const MAX_COLUMNS = 16384 ' ostatnia kolumna
Global Const C_HOUR = (0.041 + 0.001 * (2 / 3))
Global Const MAKE_LIST_TIMES_F8 = 11
Global Const WaitDynamicForm_postepZmian_max = 100

Global INPUT_DATA As Collection
Global poprzedni_select_range As Range
Global poprzedni_range_value As Long

' handler pod podswietlanie zmian
Global rng As Range
Global value_before As Range

Global cloud_item As Cloud
Global event_podmiany As EventPodmiany


Enum E_REPORT_TYPE
    WEEKLY
    DAILY
    HOURLY
End Enum

' Global Const KOLORY = Sheets("register").Range("g29")



Public Function chrx(col As Integer, Optional ByRef s As Box) As String

    If col <= MAX_COLUMNS And col > 0 Then
    If s Is Nothing Then
        Set s = New Box
    End If
    
    If col > LETTERS Then
        s.counter = s.counter + 1
        If s.counter = 26 Then
        ' wersja prostsza
            s.counter = 0
            s.scope = s.scope + 1
        End If
        chrx = chrx(col - LETTERS, s)
    Else
        If s.counter = 0 And s.scope = 0 Then
            chrx = chrx + Chr(64 + col)
        ElseIf s.counter <> 0 And s.scope = 0 Then
            chrx = chrx + Chr(64 + s.counter) + Chr(64 + col)
        ElseIf s.counter <> 0 And s.scope <> 0 Then
            chrx = chrx + Chr(64 + s.scope) + Chr(64 + s.counter) + Chr(64 + col)
        End If
    End If
    Else
        MsgBox "out of scope mf! MAX_COLUMNS = 16384"
    End If
    
   
End Function

' wersja prototypowa bardzo uniwersalnego last_row jeszcze nie testowana :P 2012-04-17
Public Function last_row(Optional adr As String, Optional sh As String)
        Dim rng As Range
        
        If adr = "" Then
            Sheets("input").Activate
            Set rng = Range("a2")
            rng.Select
        ElseIf adr <> "" And Len(adr) = 2 Then
            If sh = "" Then
                Set rng = Range(CStr(adr))
            ElseIf sh <> "" Then
                Set rng = Sheets(CStr(sh)).Range(CStr(adr))
            End If
        End If
        
        While (rng.EntireRow.Hidden = True) Or (rng.Value <> "")
            Set rng = rng.Offset(1, 0)
        Wend

    
        last_row = rng.row - 1
End Function

Public Function catch_last_column()
    ' Range("B5").Select
    Range("B5").End(xlToRight).Select
    catch_last_column = Selection.Column
    Range("B5").Select
End Function

Public Function first_daily_ebal(r As Range)
    
    Set r = Cells(r.row, r.Column + 12)
    Do
        
        Set r = Cells(r.row, r.Column + 3)
        If Cells(4, r.Column - 2).Value = "" Then
            Exit Do
        End If
    Loop While Not Cells(4, r.Column - 2).Value Like "????-??-?? *"
    
    first_daily_ebal = r
End Function

Public Function next_daily_ebal(r As Range)

    ' Set r = Cells(r.Row, r.Column + 3)
    If r.Column >= 19 Then
    Do
        Set r = Cells(r.row, r.Column + 3)
        If Cells(4, r.Column - 2).Value = "" Then
            Exit Do
        End If
    Loop While (Not Cells(4, r.Column - 2).Value Like "????-??-?? *")
    
    next_daily_ebal = r
    End If
End Function

Public Sub przelicz_arkusz(sh As Object, rr As Range, Optional first_time As Boolean)

    ' pora nieco poprawic implementacje czesciowego odswiezania danych poniewaz
    ' czekanie za kazdym razem az wszystko sie odswiezy jest masakratorem
    ' szczegolnie kiedy raport przekracza wiecej niz kilkanascie juz czesci [sic!]
    
    '
    '
    ' 1 bede potrzebowal konkretne informacje na temat zmiany w konkretnym wierszu
    ' 2 a moze nawet ostatnich wszystkich zmian jakie sie dokonaly na arkuszu
    ' teraz czy excel posiada mozliwosci sprawdzania jakie komorki ostatnio zostaly zmienione
    ' MsgBox Sheets("register").Range("ostatniaSelekcja")

    ' Debug.Print rr.Address
    ' ten na ktory zostal przeklikany
    
    

        
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    WaitDynamicForm.show vbModeless

    Dim nastepny_wiersz As Boolean
    
    ' Dim status_symulacji As StatusHandler
    
    nastepny_wiersz = True
            
            
    ' Set sh = ActiveSheet
    ' MsgBox sh.name
    If (sh.Cells(1, 1) Like "daily*") Or (sh.Cells(1, 1) Like "weekly*") Then
        Dim all_data As Range
        Set all_data = Range("b6:" & chrx(Sheets("register").Range("lastColumn")) & Sheets("register").Range("lastRow"))
        
        If Sheets("register").Range("redpink") = Sheets("register").Range("KOLORY").Offset(1, 0) Then
            all_data.NumberFormat = "0_ ;[Red]-0 "
        Else
            all_data.NumberFormat = "0_ ;[Black]-0 "
            
        
                ' WaitDynamicForm.postepZmian.max = all_data.COUNT + 1
                ' MsgBox Range("b6:" & chrx(Sheets("register").Range("itemDays")) & Sheets("register").Range("lastRow")).Address
                ' Dim r As Range
                
                
                'Set status_symulacji = New StatusHandler
                'status_symulacji.init_statusbar all_data.COUNT
                'status_symulacji.show
                'Dim r As Range
                For Each r In all_data
                    ' MsgBox r
                    
                    ' status_symulacji.progress_increase
                    
                    If Not r.EntireRow.Hidden = True Then
                    
                        If r.Column = 4 Then
                        
                            ' MsgBox Cells(r.row, 4) & " " & Cells(r.row, 12) & " " & Cells(r.row, 17)
                        
                            If r < 0 Then
                                r.Interior.Color = RGB(240, 0, 0)
                            ElseIf ((CLng(Cells(r.row, 4)) - CLng(Cells(r.row, 12)) - CLng(Cells(r.row, 17))) < 0) And (r.row > 5) Then
                                r.Interior.Color = RGB(250, 180, 200)
                            Else
                                r.Interior.Color = RGB(255, 255, 255)
                            End If
                        ElseIf (r.Column >= 17) And ((r.Column - 16) Mod 3) = 0 Then
                        
                            If (Cells(4, r.Column - 2).Value Like "????-??-?? *") Or (Cells(4, r.Column - 2).Value Like "CW *") Then
                            
                                ' ------------------------------------------------------------------------
                                If r < 0 Then
                                    If nastepny_wiersz = True Then
                                        Cells(r.row, 10) = Cells(4, r.Column - 2)
                                        nastepny_wiersz = False
                                    End If
                                    r.Interior.Color = RGB(240, 0, 0)
                                    
                                ElseIf (r >= 0) And (r.Column >= 19) And (Cells(r.row, r.Column + 1) > r) Then
                                    r.Interior.Color = RGB(250, 180, 200)
                                ElseIf r >= 0 And next_daily_ebal(Cells(r.row, r.Column)) < 0 Then
                                    r.Interior.Color = RGB(250, 180, 200)
                                Else
                                    ' r.Interior.Color = r.Offset(0, -1).Interior.Color
                                     ' bialy, czy fioletowy
                                     
                                     If (((r.Column - 16) / 3) Mod 2) = 1 Then
                                        r.Interior.Color = RGB(200, 200, 200)
                                     ElseIf (((r.Column - 16) / 3) Mod 2) = 0 Then
                                        r.Interior.Color = RGB(255, 255, 255)
                                     End If
                                End If
                                ' ------------------------------------------------------------------------
                            End If
                        End If
                        
                        
                        If Int(r.Column) = Int(Sheets("register").Range("lastColumn")) Then
                            If nastepny_wiersz = True Then
                                If r.row > 5 Then
                                    Cells(r.row, 10) = ""
                                End If
                            Else
                                nastepny_wiersz = True
                            End If
                        End If
                    End If
                    
                Next r
        End If
        
        'status_symulacji.hide
        'Set status_symulacji = Nothing
        
    ElseIf sh.Cells(1, 1) Like "hourly*" Then
        
        Dim beg_of_loop As Integer
        Dim end_of_loop As Integer
        
        Dim pierwszy_runout As Boolean
        
        If IsMissing(first_time) Then
            init_for_loop_ beg_of_loop, end_of_loop
        Else
            beg_of_loop = 6
            end_of_loop = Int(Sheets("register").Range("lastRowHourly"))
        End If
        
    
        ' Dim r As Range
        Dim tb As Range
        Dim arr_c() As String
        Dim inner_arr() As String
        Dim tmp_bank As Integer
        ' For x = 6 To Int(Sheets("register").Range("lastRowHourly")) Step 7
        
        
        'Set status_symulacji = New StatusHandler
        'status_symulacji.init_statusbar (end_of_loop - beg_of_loop) / 7
        'status_symulacji.show
        ' WaitDynamicForm.postepZmian.max = end_of_loop - beg_of_loop + 7
        
        For x = beg_of_loop To end_of_loop Step 7
        
        
            'status_symulacji.progress_increase
        
            ' CBAL layoucik
            ' ==================================================================
            
            Set tb = Cells(x - 2, 3)
            
            If Sheets("register").Range("redpink") = Sheets("register").Range("KOLORY").Offset(1, 0) Then ' simplified RED
                tb.NumberFormat = "0_ ;[Red]-0 "
            Else
                tb.NumberFormat = "0_ ;[Black]-0 "
                
            
                    pink_flag = Int(tb.Offset(-2, 4).Value) + Int(tb.Offset(0, 6).Value)
                    
                    If tb < 0 Then
                        tb.Interior.Color = RGB(240, 0, 0)
                        tb.Font.Bold = True
                    ElseIf tb < pink_flag Then
                        tb.Interior.Color = RGB(250, 180, 200)
                        tb.Font.Bold = True
                    Else
                        tb.Interior.Color = tb.Offset(-1, 0).Interior.Color
                        tb.Font.Bold = False
                    End If
            End If
            ' ==================================================================
        
            ' For r = Int(Sheets("register").Range("firstColumnHourly")) To Int(Sheets("register").Range("lastRowHourly"))
            ' MsgBox Range(chrx(Sheets("register").Range("firstColumnHourly")) & CStr(x) & ":" & chrx(Sheets("register").Range("lastColumnHourly")) & CStr(x)).Address
            
            Dim big_hourly_range As Range
            Set big_hourly_range = Range(chrx(Sheets("register").Range("firstColumnHourly")) & CStr(x) & ":" & chrx(Sheets("register").Range("lastColumnHourly")) & CStr(x))
            
            If Sheets("register").Range("redpink") = Sheets("register").Range("KOLORY").Offset(1, 0) Then ' simplified RED
                big_hourly_range.NumberFormat = "0_ ;[Red]-0 "
            Else
                big_hourly_range.NumberFormat = "0_ ;[Black]-0 "
                
                
                    pierwszy_runout = False
                    For Each r In big_hourly_range
        
                        If r.Column = Int(Sheets("register").Range("firstColumnHourly")) Then
                            Cells(r.row - 1, 5) = ""
                        End If
                    
                        r.Font.Bold = True
                        r.Font.size = 13
                        
                        Set tb = Cells(r.row - 4, 3)
                        arr_c = Split(tb.Comment.Text, Chr(10))
                        
                        For q = LBound(arr_c) To UBound(arr_c)
                            If arr_c(q) Like "BANK:*" Then
                                inner_arr = Split(arr_c(q), " ")
                                tmp_bank = Int(inner_arr(UBound(inner_arr)))
                                Exit For
                            End If
                        Next q
                        
                        pink_flag = CLng(CLng(Sheets("register").Range("pinkOnHourly")) * tmp_bank * 0.01)
                        If CStr(r) <> "" Then
                            If r < 0 Then
                            
                                If pierwszy_runout = False Then
                                    ' wypisz tylko raz w przypadku pojawienia sie w jednym wierszu
                                    Cells(r.row - 1, 5) = CStr(Cells(r.row - 4, r.Column)) & " " & CStr(Format(Cells(r.row - 3, r.Column), "hh:mm"))
                                    Cells(r.row - 1, 5).Font.Bold = True
                                End If
                            
                                pierwszy_runout = True
                                
                                ' r.Font.Color = RGB(240, 0, 0)
                                r.Interior.Color = RGB(240, 0, 0)
                                r.Font.Bold = True
                            ElseIf r < pink_flag Then
                                r.Interior.Color = RGB(250, 180, 200)
                                r.Font.Bold = True
                            ElseIf r >= 0 Then
                                r.Interior.Color = r.Offset(-1, 0).Interior.Color
                                ' r.Font.Bold = False
                            End If
                        End If
                    Next r
                    
            End If
            
        Next x

    End If
    
    
    WaitDynamicForm.hide
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ' Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub init_for_loop_(ByRef b As Integer, ByRef e As Integer)
    If ActiveCell.row < 6 + 6 * 7 Then
        b = 6
        e = 6 + 6 * 7
    Else
        tmp = ActiveCell.row
        tmp = tmp - 6
        b = 6 + tmp - 3 * 7
        e = 6 + tmp + 3 * 7
    End If
End Sub


Public Sub przelicz_parametry_arkusza(sh As Object, ByRef rr As Range)



        Application.EnableEvents = False
        If sh.Name <> "register" Then
            If sh.Name <> "input" Then
                Sheets("register").Range("sheetName") = sh.Name
                
                If sh.Cells(1, 1) Like "daily*" Then
                    Sheets("register").Range("lastColumn") = catch_last_column
                    Sheets("register").Range("lastRow") = last_row("b5", sh.Name)
                    Sheets("register").Range("allParts") = Sheets("register").Range("lastRow") - 5
                    Range("b5").Select
                ElseIf sh.Cells(1, 1) Like "weekly*" Then
                    Sheets("register").Range("lastColumn") = catch_last_column
                    Sheets("register").Range("lastRow") = last_row("b5", sh.Name)
                    Sheets("register").Range("allParts") = Sheets("register").Range("lastRow") - 5
                    Range("b5").Select
                ElseIf sh.Cells(1, 1) Like "hourly*" Then
                    Sheets("register").Range("lastColumnHourly") = catch_last_coloured_column(2, 2)
                    Sheets("register").Range("lastRowHourly") = catch_last_hourly_row()
                End If
                
                If sh.Cells(1, 1) Like "*RED and PINK*" Then
                    Sheets("register").Range("redpink") = Sheets("register").Range("KOLORY")
                ElseIf sh.Cells(1, 1) Like "*simplified RED*" Then
                    Sheets("register").Range("redpink") = Sheets("register").Range("KOLORY").Offset(1, 0) ' simplified RED
                End If
                
            End If
        End If
        Application.EnableEvents = True
End Sub

Public Function catch_last_coloured_column(row As Integer, col As Integer) As Long
    Dim cell As Range
    Set cell = Cells(row, col)
    Do
        If (cell.Interior.Color = RGB(200, 200, 200)) Or (cell.Interior.Color = RGB(240, 240, 240)) Then 'Or (cell.Interior.Color = RGB(0, 50, 220)) Then
            Set cell = cell.Offset(0, 1)
        Else
            catch_last_coloured_column = cell.Column - 1
            ' MsgBox chrx(cell.Column) ' o jeden za duzo bo ostatnie przesuniecie bylo na ostatnim kolorze :)
            Exit Do
        End If
    Loop While True
End Function

Public Function catch_last_hourly_row() As Long
    Dim cell As Range
    Set cell = Range("b6")
    Do
        If cell = "Supplier" Then
            Set cell = cell.Offset(7, 0)
        Else
            catch_last_hourly_row = cell.row - 7
            Exit Do
        End If
    Loop While True

End Function

