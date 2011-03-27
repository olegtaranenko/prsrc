Attribute VB_Name = "AppUtils"
Option Explicit

Dim quantity As Long
Public gain2 As Single, gain3 As Single, gain4 As Single

'константы для whoReserved
Public Const rtNomZak = 1
Public Const rtReserv = 2
Public Const rtCeh = 3
Public Const rtData = 4
Public Const rtMen = 5
Public Const rtStatus = 6
Public Const rtFirma = 7
Public Const rtProduct = 8
Public Const rtZakazano = 9
Public Const rtOplacheno = 10

' Константы для Справочника фирм по продажам
Public Const bfNazwFirm = 1
Public Const bfM = 2
Public Const bfRegion = 3
Public Const bfBayInfo = 4
Public Const bfTools = 5
Public Const bfBayStatus = 6
Public Const bfDirector = 7
Public Const bfFIO = 8
Public Const bf2001 = 9
Public Const bf2002 = 10
Public Const bf2003 = 11
Public Const bf2004 = 12

Public Const bfId = 0
Public Const bfType = 13
Public Const bfOborud = 14
Public Const bfSale = 15

Public sc ' ScriptControl
Private CenaFreight As String, CenaFact As String
Const DLM = vbTab

'Public Const bfOtklik = 7
'Public Const bfTlf = 13
'Public Const bfFax = 14
''Public Const bfEmail = 15
'Public Const bfLogin = 17
'Public Const bfPass = 18



' Этот файл разделяется между prior, stime и rowmat.
' не использовать в cfg

Private Sub changeCaseOfTheVariables()
'Dim IsEmpty, Numorder, StatusId, Rollback, Outdatetime, p_numOrder, tbWorktime, Left, RemoveItem, J, Value, X, Y, Table, IL, Name, L, Equip, Worktime, ManagId, ColWidth, Index, W, K, Visible, Field, Fields, WerkId, FirmId, Edizm2, V, Key, RemoveAll, Remove, Frm, xGroup, Delim, Item, ListBox, Nomname, Formula, Nomnom, perList, Ves, prSeriaId
Dim IsEmpty, Numorder, StatusId, Rollback, Outdatetime, p_numOrder, tbWorktime, Left, RemoveItem, J, Value, X, Y, Table, IL, Name, L, Equip, Worktime, ManagId, ColWidth, Index, W, K, Visible, Field, Fields, WerkId, FirmId, Edizm2, V, Key, RemoveAll, Remove, Frm, xGroup, Delim, Item, ListBox, Nomname, Formula, Nomnom, perList, Ves, prSeriaId

End Sub


Sub GridToExcel(Grid As MSFlexGrid, Optional title As String = "")
Dim ColWidth As String, Note As String


Dim objExel As Excel.Application, c As Long, r As Long
Dim I As Integer, strA() As String, begRow As Integer, str As String

begRow = 3
If title = "" Then begRow = 1

Set objExel = New Excel.Application
objExel.Visible = True
objExel.SheetsInNewWorkbook = 1
objExel.Workbooks.Add
With objExel.ActiveSheet
.Cells(1, 2).Value = title
ReDim Preserve strA(Grid.Cols + 1)
For r = 0 To Grid.Rows - 1
    Dim curColumn As Integer
    curColumn = 1
    For c = 1 To Grid.Cols - 1
        If Grid.ColWidth(c) > 0 Then
            str = Grid.TextMatrix(r, c) '=' - наверно зарезервирован для ввода формул
            Dim firstLetter As String
            firstLetter = Left$(str, 1)
            Dim doEscape As Boolean
            
            If firstLetter = "=" Or firstLetter = "+" Then
                doEscape = True
            End If
            
            If str = "--" Then
                doEscape = True
            End If
            
            If doEscape Then
                str = "'" & str
            End If
'иногда символы Cr и Lf (поле MEMO в базе) дают Err в Excel, поэтому из поля
            I = InStr(str, vbCr) 'MEMO берем только первую строчку
            If I > 0 Then str = Left$(str, I - 1)
            I = InStr(str, vbLf) 'MEMO берем только первую строчку
            If I > 0 Then str = Left$(str, I - 1)
            If IsNumeric(str) And r > 0 Then
                strA(curColumn - 1) = CStr(CDbl(str))
            Else
                If Len(str) > 255 Then
                    str = Left(str, 252) & "..."
                End If
                strA(curColumn - 1) = str
            End If
            curColumn = curColumn + 1
        End If
    Next c
'    On Error Resume Next
   .Range(.Cells(begRow + r, 1), .Cells(begRow + r, Grid.Cols)).FormulaArray = strA
Next r

'objExel.ActiveSheet.Range("A" & begRow & ":U" & Grid.Rows + begRow).FormulaArray = strA
'.Range(.Cells(begRow, 1), .Cells(Grid.Rows + begRow, Grid.Rows)).FormulaArray = strA
End With
Set objExel = Nothing
End Sub



Function existsInTreeview(ByRef tTree As TreeView, Key As String) As Boolean
Dim I As Integer
    For I = 1 To tTree.Nodes.Count
        If tTree.Nodes(I).Key = Key Then
            existsInTreeview = True
            Exit Function
        End If
    Next I
    existsInTreeview = False
End Function



Function newNumorder(Value As String) As Numorder
    Dim ret As Numorder
    Set ret = New Numorder
    ret.val = Value
    Set newNumorder = ret
End Function


Function getNextDocNum() As Long
Dim valueorder As Numorder

    Set valueorder = New Numorder
    valueorder.val = getSystemField("lastDocNum")
    If valueorder.IsEmpty Then
        valueorder.docs = True
    End If
    If Not valueorder.isCurrentDay Then
        Set valueorder = New Numorder
        valueorder.docs = True
    End If
    valueorder.nextNum
    sql = "UPDATE SYSTEM SET lastDocNum = " & valueorder.val
    'Debug.Print sql
    myBase.Execute (sql)
    numDoc = valueorder.val
    
    getNextDocNum = valueorder.val
    
End Function

Sub showRezerv(ByVal dostupOstatki As Double, ByVal factOstatki As Double, ByVal Edizm2 As String, ByRef f As Form)
    If Round(factOstatki, 2) > Round(dostupOstatki, 2) Then
        If MsgBox("Если Вы хотите просмотреть список всех заказов, под " & _
        "которые была зарезервирована эта номенклатура, нажмите <Да>.", _
        vbYesNo Or vbDefaultButton2, "Посмотреть, кто резервировал? '" & _
        gNomNom & "' ?") = vbYes Then
            Report.Edizm2 = Edizm2
            Report.Regim = "whoRezerved"
            Set Report.Caller = f
            Report.Sortable = True
            Report.Show vbModal
        End If
    Else
        MsgBox "Эта номенклатура никем не резервировалась.", , ""
    End If

End Sub



Function whoRezerved(ByRef Grid As MSFlexGrid, Optional p_term_index As Integer = 0) As Integer
Dim groupklassid As Integer, rowStr As String
Dim p_days_start As Integer, p_days_end As Integer
    

    Grid.Visible = False
    
    Grid.FormatString = "|<№ заказа|>кол-во|^Цех |^Дата |^ М|Статус" & _
    "|<Название Фирмы|<Изделия|>Заказано|>Согласовано"
    Grid.ColWidth(0) = 0
    'Grid.ColWidth(rtNomZak) =
    Grid.ColWidth(rtReserv) = 765
    Grid.ColWidth(rtCeh) = 765
    Grid.ColWidth(rtData) = 1600
    'Grid.ColWidth(rtMen) =
    Grid.ColWidth(rtStatus) = 930
    Grid.ColWidth(rtFirma) = 3270
    Grid.ColWidth(rtProduct) = 1950
    'Grid.ColWidth(rtZakazano) =
    Grid.ColWidth(rtOplacheno) = 810

    If p_term_index = 0 Then
        p_days_start = 10000
        p_days_end = 0
    ElseIf p_term_index = 1 Then
        p_days_start = 30
        p_days_end = 0
    ElseIf p_term_index = 2 Then
        p_days_start = 60
        p_days_end = 30
    ElseIf p_term_index = 3 Then
        p_days_start = 120
        p_days_end = 60
    ElseIf p_term_index = 4 Then
        p_days_start = 180
        p_days_end = 120
    ElseIf p_term_index = 5 Then
        p_days_start = 10000
        p_days_end = 120
    End If
    
    
    sql = "call wf_order_reserved ('" & gNomNom & "', " & p_days_start & ", " & p_days_end & ")"
    
    Set tbOrders = myOpenRecordSet("##350", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Function
    quantity = 0
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF

            quantity = quantity + 1
            Grid.TextMatrix(quantity, rtNomZak) = tbOrders!Numorder
            Grid.TextMatrix(quantity, rtReserv) = Format(tbOrders!quant, "# ##0.00")
            If Not IsNull(tbOrders!Werk) Then _
                Grid.TextMatrix(quantity, rtCeh) = tbOrders!Werk
            
            Grid.TextMatrix(quantity, rtData) = tbOrders!date1
            If Not IsNull(tbOrders!Manager) Then _
                Grid.TextMatrix(quantity, rtMen) = tbOrders!Manager
            
            If Not IsNull(tbOrders!Status) Then _
                Grid.TextMatrix(quantity, rtStatus) = tbOrders!Status
            
            If Not IsNull(tbOrders!client) Then _
                Grid.TextMatrix(quantity, rtFirma) = tbOrders!client
            
            If Not IsNull(tbOrders!Note) Then _
                Grid.TextMatrix(quantity, rtProduct) = tbOrders!Note
            
            If Not IsNull(tbOrders!sm_zakazano) Then _
                Grid.TextMatrix(quantity, rtZakazano) = Format(tbOrders!sm_zakazano, "# ##0.00")
                
            If Not IsNull(tbOrders!sm_paid) Then _
                Grid.TextMatrix(quantity, rtOplacheno) = Format(tbOrders!sm_paid, "# ##0.00")
                
            Grid.AddItem ""
            tbOrders.MoveNext
        Wend
    End If
  tbOrders.Close

'laRecSum.Caption = Round(sum, 2)
If quantity > 0 Then
    Grid.RemoveItem quantity + 1
End If
trigger = False
Grid.Visible = True
whoRezerved = quantity


End Function


Function RuDate2Date(ByVal dt As String) As Variant
Dim dotPos As Integer
Dim v_dd As Integer
Dim v_mm As Integer
Dim v_yyyy As Integer
    
    On Error GoTo catch
    dotPos = InStr(dt, ".")
    If IsNull(dotPos) Or dotPos = 0 Then
        v_dd = CInt(dt)
        dt = ""
    Else
        v_dd = CInt(Left(dt, dotPos))
        dt = Mid(dt, dotPos + 1)
    End If
    
    If Len(dt) > 0 Then
        dotPos = InStr(dt, ".")
        If IsNull(dotPos) Or dotPos = 0 Then
            v_mm = CInt(dt)
            dt = ""
        Else
            v_mm = CInt(Left(dt, dotPos))
            dt = Mid(dt, dotPos + 1)
        End If
    Else
        v_mm = Format(Now(), "mm")
    End If
    
    If Len(dt) <= 2 And Len(dt) > 0 Then
        v_yyyy = CInt("20" & Format(CInt(dt), "0#"))
    Else
        v_yyyy = Format(Now(), "yyyy")
    End If
    
    RuDate2Date = CDate(CStr(v_yyyy) & "-" & Format(v_mm, "0#") & "-" & Format(v_dd, "0#"))
    Exit Function
catch:
    RuDate2Date = False
End Function

Function myIsDate(ByVal str As String) As Variant
    Dim dt As Variant
    dt = RuDate2Date(str)
    If IsDate(dt) Then
        myIsDate = Format(dt, "dd.mm.yy")
    Else
        myIsDate = dt
    End If
End Function

Public Sub initListbox(ByVal InitSql As String, ByVal lb As ListBox, keyField As String, valueField As String, Optional ListBoxMode As Integer = 0)

    If ListBoxMode = 0 Then
        ' Сначала удаляем старые значения
        While lb.ListCount
            lb.RemoveItem (0)
        Wend
        
        lb.AddItem "", 0
    End If

    Set Table = myOpenRecordSet("##72.0", InitSql, dbOpenForwardOnly)
    
    If Table Is Nothing Then myBase.Close: End
    
    
    While Not Table.EOF
        lb.AddItem "" & Table(valueField) & ""
        lb.ItemData(lb.ListCount - 1) = Table(keyField)
        Table.MoveNext
    Wend
    Table.Close
    lb.Height = 225 * lb.ListCount

End Sub

Public Sub initCombobox(ByVal InitSql As String, ByVal lb As ComboBox, keyField As String, valueField As String, Optional ListBoxMode As Integer = 0)
    If ListBoxMode = 0 Then
        ' Сначала удаляем старые значения
        While lb.ListCount
            lb.RemoveItem (0)
        Wend
        
        lb.AddItem "", 0
    End If

    Set Table = myOpenRecordSet("##72.0", InitSql, dbOpenForwardOnly)
    
    If Table Is Nothing Then myBase.Close: End
    
    
    While Not Table.EOF
        lb.AddItem "" & Table(valueField) & ""
        lb.ItemData(lb.ListCount - 1) = Table(keyField)
        Table.MoveNext
    Wend
    Table.Close

End Sub


Function OstatToCSV(ByRef Frm As VB.Form, Regim As String, csvFile As String, curRate As Double) As String
Dim prices(4) As Double, I As Integer
Dim csvRow As String, RPF_Rate As Single
Dim ret As String


Frm.MousePointer = flexHourglass


  sql = "call wf_report_mat_ost(1)"
    
    'Debug.Print sql

    Set tbNomenk = myOpenRecordSet("##415", sql, dbOpenDynaset)
    If Not tbNomenk Is Nothing Then
        
        'On Error GoTo ERR1
                
        If Not tbNomenk.BOF Then
            Open csvFile For Output As #1
            Print #1, "cod" & DLM & "description" & DLM & "size" _
                     & DLM & "quantity" & DLM & "edizm" _
                     & DLM & "price1" & DLM & "price2" & DLM & "price3" & DLM & "price4"
            
            While Not tbNomenk.EOF
            
            
                csvRow = tbNomenk!cod & DLM & tbNomenk!Nomname _
                    & DLM & tbNomenk!Size & DLM & Int(tbNomenk!qty_dost) & DLM & tbNomenk!ed_Izmer2
    
                CalcKolonPrices prices, curRate
    
                For I = 0 To UBound(prices) - 1
                    csvRow = csvRow & DLM & Round(prices(I), 1)
                Next I
    
                Print #1, csvRow
                
                tbNomenk.MoveNext
            Wend
            Close #1
        End If
        tbNomenk.Close
    End If
GoTo EN2
        
ERR1:
If Err = 76 Then
    MsgBox "Невозможно создать файл " & csvFile, , "Error: Не обнаружен ПК или Путь к файлу"
ElseIf Err = 53 Then
    Resume Next ' файла м.не быть
ElseIf Err = 47 Then
    MsgBox "Невозможно создать файл " & csvFile, , "Error: Нет доступа на запись."
Else
    MsgBox Error, , "Ошибка 47-" & Err '##47
    
End If
GoTo done


EN2:
On Error Resume Next 'нужен, если фокус после нажатия передали другому приложению
MsgBox "Файл " & csvFile & " успешно сформирован.", , "Файлы для WEB"

done:
    Frm.MousePointer = flexDefault

End Function


Private Function convertToCsv(Cells() As Variant) As String

Dim I As Integer
Dim ret As String
    ret = CStr(Cells(1))
    For I = 2 To UBound(Cells)
        ret = ret & DLM & Cells(I)
    Next I
    convertToCsv = ret
    
End Function

Public Sub BrightAwardsRestToCsv(csvFile As String, csvHeaders As String, Optional Regim As String = "", Optional RubRate As Double = 1, Optional commonRabbat As Single = 0)

Dim currentSeriaId As Integer
Dim currentProductId As Integer
Dim currentVariative As String
Dim csvRow As String, RPF_Rate As Single
Dim priceType As Integer
Dim izdeliePrices(4) As Single, I As Integer
Dim saveHeaders As Boolean, headMap() As MapEntry

' да, если строку уже вывели в файл. Встречается у вариативных изделий
Dim submitted As Boolean


'будем хранить значения таблицы, аналог ячейки экселя
' для изделий
Dim Cells(10) As Variant
' для вариантной комплектации
Dim vCells(10) As Variant

    
    ReDim headMap(0)
    saveHeaders = csvHeaders <> ""
    
    priceType = 0
    
    ' Если изделие состоит из номенклатуры типа "web",
    ' то его цена для дилеров определяется как сумма по справочнику номенклатуры
    Dim WebIzdelie() As Double
    
    Dim priceRegim As String
    
    
    Screen.MousePointer = flexHourglass
    On Error GoTo ERR1
    
    
    If priceType = 0 Then     ' RPF
        priceRegim = "agency"
    ElseIf priceType = 1 Then ' dealer
        priceRegim = "dealer"
    Else
        priceRegim = "default"
    End If

    'On Error GoTo ERR2
        
    If saveHeaders Then
        Open csvHeaders For Output As #2
        Print #2, "seriaId" & DLM & "seria" _
                & DLM & "head1" & DLM & "head2" _
                & DLM & "head3" & DLM & "head4"
    End If
        
    ' если в Изделии нет вариативных изделий, то не печатать комплектацию,
    ' а только кол-во доступных, определяемых по наименьшему доступному из всех комплектующих
    Dim izdeliaQty As Long, lastIzdeliaType As String
    Dim dostOstatok As String, nomDostOst As Integer
    izdeliaQty = -1
    
    
    sql = "call wf_report_bright_ostat"
    Set tbProduct = myOpenRecordSet("##331", sql, dbOpenDynaset)
    
    If Not tbProduct.BOF Then
        Open csvFile For Output As #1
        
        Print #1, "cod" & DLM & "description" & DLM & "size" & DLM & "page" & DLM & "quantity" _
                & DLM & "price1" & DLM & "price2" & DLM & "price3" & DLM & "price4" & DLM & "seriaid"
                
        While Not tbProduct.EOF

            If saveHeaders Then
                Dim seriaId As String
                seriaId = tbProduct!prSeriaId
                If getMapEntryIndex(headMap, seriaId) = -1 Then
                    appendKeyValue headMap, seriaId, ""
                    Print #2, seriaId & DLM & tbProduct!seriaName _
                            & DLM & tbProduct!head1 & DLM & tbProduct!head2 _
                            & DLM & tbProduct!head3 & DLM & tbProduct!head4
                                            
                Else
                    
                End If
                
            End If
            
            
            If Not IsNull(tbProduct!qty_dost) Then
                nomDostOst = tbProduct!qty_dost
            Else
                nomDostOst = 0
            End If
            
            If currentProductId <> tbProduct!prId Then
            
                ' первый проход пропускаем
                If currentProductId > 0 And Not submitted Then
                    csvRow = convertToCsv(Cells)
                    Print #1, csvRow
                End If
                
                ' при смене изделия - обнуляем для нового цикла
                izdeliaQty = -1
                submitted = False
                
                Cells(1) = CStr(tbProduct!prName)
                Cells(2) = CStr(tbProduct!prDescript)
                Cells(3) = CStr(tbProduct!prSize)
                Cells(4) = CInt(tbProduct!Page)
                Cells(5) = izdeliaQty
                Cells(10) = tbProduct!prSeriaId

                'gain2 = tbProduct!gain2
                'gain3 = tbProduct!gain3
                'gain4 = tbProduct!gain4
                'ExcelProductPrices RPF_Rate, priceRegim, RubRate,  6, commonRabbat
                csvRow = CsvProductPrices(izdeliePrices, RPF_Rate, priceRegim, RubRate, commonRabbat)
                If Not csvRow = "" Then
                    MsgBox "Ошибка при вычислении цены для изделия " _
                    & vbCr & "Текст ошибки: " & csvRow, , "Изделие " & tbProduct!prName & "  " & tbProduct!prDescript
                    MsgBox "Генерация файла прервана. Сначала нужно исправить ошибку", , csvFile
                    Close #1
                    GoTo done
                End If
                
                For I = 0 To 3
                    Cells(I + 6) = Round(izdeliePrices(I), 1)
                    izdeliePrices(I) = 0
                Next I
                
            End If
                
            If tbProduct!quantEd <> 1 Then
                nomDostOst = nomDostOst * tbProduct!quantEd
            End If
            
            currentVariative = tbProduct!variative
            If currentVariative = "V" Then
                
                If Not submitted Then
                    submitted = True
                    csvRow = convertToCsv(Cells)
                    Print #1, csvRow
                End If
                
                vCells(1) = tbProduct!Ncod
                vCells(2) = tbProduct!Nomname
                vCells(3) = tbProduct!Nsize
                'vCells(4) = tbProduct!ed_Izmer2
                If tbProduct!quantEd <> 1 Then
                    dostOstatok = nomDostOst / tbProduct!quantEd
                Else
                    dostOstatok = nomDostOst
                End If
                If dostOstatok > 0 Then
                    vCells(5) = dostOstatok
                Else
                    vCells(5) = 0
                End If

                csvRow = convertToCsv(vCells)
                Print #1, csvRow


'                    If Regim = "awards" Then
'                        ExcelKolonPrices  6, RubRate, RPF_Rate
'                    End If
            
            Else
            
                If nomDostOst < izdeliaQty Or izdeliaQty = -1 Then
                    If nomDostOst >= 0 Then
                        Cells(5) = nomDostOst
                   Else
                        Cells(5) = 0
                    End If
                End If
                izdeliaQty = nomDostOst
            End If
            

            currentSeriaId = tbProduct!prSeriaId
            currentProductId = tbProduct!prId
            lastIzdeliaType = tbProduct!variative

            tbProduct.MoveNext
        Wend
    
        ' если последняя строчка не была V, тогда последнее изделие выводим только сейчас
        If currentVariative <> "V" Then
            csvRow = convertToCsv(Cells)
            Print #1, csvRow
        End If
        
        Close #1
    
    End If
    
    tbProduct.Close

    If saveHeaders Then
        Close #2
    End If

GoTo EN2
ERR1:
    If Err = 76 Then
        MsgBox "Невозможно создать файл " & csvFile, , "Error: Не обнаружен ПК или Путь к файлу"
    ElseIf Err = 53 Then
        Resume Next ' файла м.не быть
    ElseIf Err = 47 Then
        MsgBox "Невозможно создать файл " & csvFile, , "Error: Нет доступа на запись."
    Else
        MsgBox Error, , "Ошибка 47-" & Err '##47
        
    End If
    GoTo done

EN2:
    MsgBox "Файл " & csvFile & " успешно сформирован.", , "Файлы для WEB"

done:
    Screen.MousePointer = flexDefault
        
        
End Sub



Function PriceToCSV(ByRef Frm As VB.Form, Regim As String, csvFile As String, curRate As Double _
        , Optional prodCategoryId As Integer = 0, Optional commonRabbat As Single = 1 _
        , Optional csvHeaders As String = "") As String
        
Dim izdeliePrices(4) As Single, I As Integer
Dim csvRow As String, RPF_Rate As Single
Dim ret As String
Dim strCrit() As String, critIndex As Integer
Dim saveHeaders As Boolean, headMap() As MapEntry



    Frm.MousePointer = flexHourglass
    On Error GoTo ERR1

    ReDim headMap(0)
    saveHeaders = csvHeaders <> ""
    
    sql = "SELECT p.prId, p.prName, p.prDescript, p.prSize" _
        & vbCr & " , p.Cena4, p.page, p.rabbat as productRabbat" _
        & vbCr & " , f.formula" _
        & vbCr & " , w.prId as hasWeb " _
        & vbCr & " , s.gain2, s.gain3, s.gain4" _
        & vbCr & " , p.prseriaId as seriaId, s.seriaName" _
        & vbCr & " , s.head1, s.head2, s.head3, s.head4" _
        & vbCr & " From sGuideProducts p " _
        & vbCr & " join sGuideSeries s on p.prSeriaId = s.seriaId " _
        & vbCr & " left JOIN sGuideFormuls f on f.nomer = p.formulaNom" _
        & vbCr & " left JOIN wf_izdeliaWithWeb w on w.prId  = p.prId"
    
    critIndex = 0
    If prodCategoryId = -1 Then
        ReDim strCrit(critIndex)
        strCrit(critIndex) = "isnull(p.prodCategoryId, 0) <> 0"
        critIndex = critIndex + 1
    ElseIf prodCategoryId > 0 Then
        ReDim strCrit(critIndex)
        strCrit(critIndex) = "p.prodCategoryId = " & prodCategoryId
        critIndex = critIndex + 1
    End If
    
    If Not Regim = "pricePM" Then
        ReDim Preserve strCrit(critIndex)
        strCrit(critIndex) = "isnumeric(p.page) = 1"
        critIndex = critIndex + 1
    End If
    
    If Regim = "agency" Then
        ReDim Preserve strCrit(critIndex)
        strCrit(critIndex) = "w.prId is null"
        critIndex = critIndex + 1
    End If
    
    If critIndex > 0 Then
        sql = sql & vbCr & " WHERE "
        For I = 0 To critIndex - 1
            If I > 0 Then
                sql = sql & " AND "
            End If
            sql = sql & strCrit(I)
        Next I
    End If

    sql = sql & " order by p.prName"
    'Debug.Print sql


    If saveHeaders Then
        Open csvHeaders For Output As #2
        Print #2, "seriaId" & DLM & "seria" _
                & DLM & "head1" & DLM & "head2" _
                & DLM & "head3" & DLM & "head4"
    End If

    Set tbProduct = myOpenRecordSet("##415", sql, dbOpenDynaset)
    If Not tbProduct Is Nothing Then
        
                
        If Not tbProduct.BOF Then
            Open csvFile For Output As #1
            Print #1, "cod" & DLM & "size" & DLM & "description" _
                    & DLM & "price1" & DLM & "price2" & DLM & "price3" & DLM & "price4" & DLM & "page" & DLM & "seriaid"
            
            While Not tbProduct.EOF
            
                If saveHeaders Then
                    Dim seriaId As String
                    seriaId = tbProduct!seriaId
                    If getMapEntryIndex(headMap, seriaId) = -1 Then
                        appendKeyValue headMap, seriaId, ""
                        Print #2, seriaId & DLM & tbProduct!seriaName _
                                & DLM & tbProduct!head1 & DLM & tbProduct!head2 _
                                & DLM & tbProduct!head3 & DLM & tbProduct!head4
                                                
                    Else
                        
                    End If
                    
                End If
                
                ret = CsvProductPrices(izdeliePrices, RPF_Rate, Regim, curRate, commonRabbat)
                
                If Not ret = "" Then
                    MsgBox "Ошибка при вычислении цены для изделия " _
                    & vbCr & "Текст ошибки: " & ret, , "Изделие " & tbProduct!prName & "  " & tbProduct!prDescript
                    MsgBox "Генерация файла прервана. Сначала нужно исправить ошибку", , csvFile
                    Close #1
                    GoTo done
                End If
                csvRow = tbProduct!prName & DLM & tbProduct!prSize _
                    & DLM & tbProduct!prDescript
    
                For I = 0 To 3
                    csvRow = csvRow & DLM & Round(izdeliePrices(I), 1)
                    izdeliePrices(I) = 0
                Next I
    
                csvRow = csvRow & DLM & tbProduct!Page & DLM & tbProduct!seriaId
    
                Print #1, csvRow
                
                tbProduct.MoveNext
            Wend
            Close #1
        End If
        tbProduct.Close
    End If
        
    If saveHeaders Then
        Close #2
    End If

    GoTo EN2
    
ERR1:
    If Err = 76 Then
        MsgBox "Невозможно создать файл " & csvFile, , "Error: Не обнаружен ПК или Путь к файлу"
    ElseIf Err = 53 Then
        Resume Next ' файла м.не быть
    ElseIf Err = 47 Then
        MsgBox "Невозможно создать файл " & csvFile, , "Error: Нет доступа на запись."
    Else
        MsgBox Error, , "Ошибка 47-" & Err '##47
        
    End If
    GoTo done


EN2:
    On Error Resume Next 'нужен, если фокус после нажатия передали другому приложению
    MsgBox "Файл " & csvFile & " успешно сформирован.", , "Файлы для WEB"
    
done:
    Frm.MousePointer = flexDefault

End Function


Function getCurrentRate() As Double
Dim S As String

    sql = "SELECT Kurs FROM System;"
    If byErrSqlGetValues("##321", sql, S) Then
        getCurrentRate = Abs(S)
    End If

End Function

' если при расчете возникла ошибка, то функция возвращает сообщение об ошибке
' в этом случае цены в массив izdeliePrices не выставляются

Public Function CsvProductPrices(ByRef izdeliePrices() As Single, ByRef RPF_Rate As Single, Regim As String, curRate As Double, commonRabbat As Single) As String
    Dim ret As String
    Dim baseCena As String
    
    gProductId = tbProduct!prId

    ret = calcBaseCenaAndRpfRate(Regim, baseCena, RPF_Rate, tbProduct!Cena4, tbProduct!productRabbat, commonRabbat, tbProduct!hasWeb)
    If ret = "" Then
        izdeliePrices(0) = CSng(baseCena) * curRate
        izdeliePrices(1) = izdeliePrices(0) * tbProduct!gain2
        izdeliePrices(2) = izdeliePrices(0) * tbProduct!gain3
        izdeliePrices(3) = izdeliePrices(0) * tbProduct!gain4
    Else
        CsvProductPrices = ret
    End If
End Function


Public Function calcBaseCenaAndRpfRate(Regim As String, ByRef baseCena As String, ByRef RPF_Rate As Single, _
        tCena4 As Variant, tProductRabbat As Variant, commonRabbat As Single, hasWebField As Variant) As String
    
    Dim rbt As Single, SumCenaFreight As String, hasWeb As Boolean, CenaFreightOK As Boolean
    If IsNull(hasWebField) Then
        hasWeb = False
    Else
        hasWeb = True
    End If
    
    RPF_Rate = 1
    If Regim = "dealer" Or Regim = "agency" Or hasWeb Then
        productFormula SumCenaFreight, baseCena, "noOpen"
        If Not IsNumeric(SumCenaFreight) Then
            calcBaseCenaAndRpfRate = SumCenaFreight
            Exit Function
        End If
    End If
    
    If Regim = "agency" And Not hasWeb Then
        RPF_Rate = baseCena 'temp storage
        If tProductRabbat = 0 Then
            rbt = commonRabbat
        Else
            rbt = tProductRabbat
        End If
        baseCena = tCena4 * (1 - rbt / 100)
        If RPF_Rate > 0 Then
            RPF_Rate = baseCena / RPF_Rate
        Else
            RPF_Rate = RPF_Rate
            
        End If


    ElseIf Regim = "default" Or Regim = "pricePM" Then
        baseCena = tCena4
    End If
    
End Function

Function calcKolonValue(ByVal freight As Double, ByVal marginProc As Double, ByVal rabbat As Double, ByVal kolonok As Double, ByVal curentKolon As Integer)
    Dim marginRate As Double, MarginValue As Double, maxUstupka As Double, stepUstupka As Double
    
    'marginRate = marginProc / 100
    'maxUstupka =
    MarginValue = freight * rabbat / 100
    If kolonok > 1 Then
        stepUstupka = MarginValue / (kolonok - 1)
    Else
        stepUstupka = 0
    End If
    
    calcKolonValue = freight - stepUstupka * (curentKolon - 1)
    
End Function



' Используется при вычислении цен и остатков по складу материалов
Private Sub CalcKolonPrices(ByRef prices() As Double, RubRate As Double, Optional RPF_Rate As Single = 1)
Dim I As Integer

    For I = 1 To UBound(prices) - 1
        prices(I) = 0
    Next I
    
    prices(0) = RPF_Rate * tbNomenk!CENA_W * RubRate
    
    Dim kolonok As Integer, optBasePrice As Double, margin As Double, iKolon As Integer, manualOpt As Boolean
    kolonok = tbNomenk!kolonok
    margin = tbNomenk!margin
    optBasePrice = tbNomenk!CENA_W
    
    If kolonok > 0 Then
        manualOpt = False
    Else
        manualOpt = True
    End If
    
    For iKolon = 1 To Abs(kolonok) - 1
        If manualOpt Then
            If iKolon = 1 Then
                prices(iKolon) = RPF_Rate * tbNomenk("Cena_W") * RubRate
            Else
                prices(iKolon) = RPF_Rate * tbNomenk("CenaOpt" & CStr(iKolon + 1)) * RubRate
            End If
        Else
            prices(iKolon) = prices(iKolon) + RPF_Rate * RubRate * calcKolonValue(optBasePrice, margin, tbNomenk!rabbat, Abs(kolonok), iKolon + 1)
        End If
    Next iKolon

End Sub




Function productFormula(ByRef SumCenaFreight As String, ByRef SumCenaSale As String, Optional noOpen As String = "")
Dim str As String

If noOpen = "" Then
    sql = "SELECT sGuideProducts.*, sGuideFormuls.Formula FROM sGuideFormuls " & _
    "INNER JOIN sGuideProducts ON sGuideFormuls.nomer = sGuideProducts.formulaNom " & _
    "WHERE (((sGuideProducts.prId)=" & gProductId & "));"
    'MsgBox sql
    Set tbProduct = myOpenRecordSet("##316", sql, dbOpenDynaset)
    If tbProduct Is Nothing Then Exit Function
    If tbProduct.BOF Then tbProduct.Close: Exit Function
End If

SumCenaFreight = getSumCena(tbProduct!prId)
If InStr(tbProduct!Formula, "SumCenaFreight") > 0 Then
  If IsNumeric(SumCenaFreight) Then
    sc.ExecuteStatement "SumCenaFreight=" & SumCenaFreight
    SumCenaFreight = Round(CSng(SumCenaFreight), 2)
  Else
    productFormula = "error СумЦ.доставка" 'текст ошибки
    'tbProduct.Close
    GoTo EN1
  End If
End If

SumCenaSale = getSumCena(tbProduct!prId, "Sale")
If InStr(tbProduct!Formula, "SumCenaSale") > 0 Then
  If IsNumeric(SumCenaSale) Then
    sc.ExecuteStatement "SumCenaSale=" & SumCenaSale
    SumCenaSale = Round(CSng(SumCenaSale), 2)
  Else
    'tbProduct.Close
    productFormula = "error СумЦоПродажа" 'текст ошибки
    GoTo EN1
  End If
End If

On Error GoTo ERR2
sc.ExecuteStatement "VremObr = " & tbProduct!VremObr
productFormula = Round(sc.Eval(tbProduct!Formula), 2)
GoTo EN1
ERR2:
    productFormula = "error: " & Error
'    MsgBox Error & " - при выполнении формулы '" & tbProduct!formula & _
'    "' для изделия '" & tbProduct!prName & "' (" & tmpStr & ")", , _
'    "Ошибка 316 - " & Err & ":  " '##316
EN1:
tmpStr = tbProduct!Formula
If noOpen = "" Then tbProduct.Close
End Function



'reg = "" => SumCenaFreight
'reg <> "" => SumCenaSale
Function getSumCena(productId As Integer, Optional reg As String = "") As String
Dim sum As Single, V, S As Single, prevGroup As String, max As Single

sum = 0

sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup, sGuideNomenk.perList, " & _
"sGuideNomenk.CENA1, sGuideNomenk.VES, sGuideNomenk.STAVKA, sGuideFormuls.Formula, " & _
"sGuideNomenk.CENA_W " & _
"FROM (sGuideFormuls INNER JOIN sGuideNomenk ON sGuideFormuls.nomer = " & _
"sGuideNomenk.formulaNom) INNER JOIN sProducts ON sGuideNomenk.nomNom " & _
"= sProducts.nomNom WHERE (((sProducts.ProductId)=" & productId & "))" & _
"ORDER BY sProducts.xgroup;"
Set tbNomenk = myOpenRecordSet("##313", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then
    'tbProduct.Close
    If reg = "" Then
        getSumCena = "Error ##31 в SumCenaFreight"
    Else
        getSumCena = "Error ##313 в SumCenaSale"
    End If
    Exit Function
End If
If tbNomenk.BOF Then
    getSumCena = "Error: Не обнаружены комплектующие"
    GoTo er
End If
'If tbProduct!prName = "S202M" Then
'    max = max
'End If
max = -1
While Not tbNomenk.EOF
'    If tbNomenk!formula = "" Then
'        getSumCena = "Error: Не определена формула для номенклатуры '" & tbNomenk!nomNom & "'"
'        GoTo ER
'    End If
    If reg = "" Then
        If tbNomenk!Formula = "" Then
            getSumCena = "Error: Не определена формула для номенклатуры '" & tbNomenk!Nomnom & "'"
            GoTo er
        End If
        V = nomenkFormula("noOpen") 'Цена2
    Else
        V = tbNomenk!CENA_W
    End If
        
    If IsNumeric(V) Then
        S = V * tbNomenk!quantity / tbNomenk!perList
        If tbNomenk!xGroup = "" Then
            sum = sum + S
            prevGroup = tbNomenk!xGroup
        ElseIf prevGroup = tbNomenk!xGroup Then
            If max < S Then max = S
        Else
            If prevGroup <> "" Then sum = sum + max
            max = S
            prevGroup = tbNomenk!xGroup
        End If
    Else
        getSumCena = V & " в формуле для  '" & tbNomenk!Nomnom & "'"
        GoTo er
    End If

    tbNomenk.MoveNext
Wend
If max = -1 Then '  - не было групп
    getSumCena = sum
Else
    getSumCena = sum + max
End If
er:
tbNomenk.Close
End Function



Function nomenkFormula(Optional noOpen As String = "", Optional Web As String = "", Optional Cena1 As Double = -1) As String

Dim str As String
Dim vCena1 As Double

If noOpen = "" Then
    sql = "SELECT sGuideNomenk.formulaNom" & Web & " , sGuideNomenk.CENA1, " & _
    "sGuideNomenk.VES, sGuideNomenk.STAVKA, sGuideFormuls.Formula as formula" & Web & _
    " FROM sGuideFormuls INNER JOIN sGuideNomenk ON sGuideFormuls.nomer = " & _
    "sGuideNomenk.formulaNom" & Web & _
    " WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
'MsgBox sql
    Set tbNomenk = myOpenRecordSet("##317", sql, dbOpenDynaset)
    If tbNomenk Is Nothing Then Exit Function
    If tbNomenk.BOF Then tbNomenk.Close: Exit Function
End If
tmpStr = tbNomenk!Formula
tmpStr = tbNomenk.Fields("formula" & Web)
'If tbNomenk!formula = "" Then
If tmpStr = "" Then
    nomenkFormula = "error: Формула не задана"
    Exit Function
End If

If Cena1 < 0 Then
    vCena1 = tbNomenk!Cena1
Else
    vCena1 = Cena1
End If


If Web = "" Then
    str = "CENA1=" & vCena1 & ": VES=" & _
    tbNomenk!Ves & ": STAVKA=" & tbNomenk!STAVKA
    sc.ExecuteStatement (str)
Else
    str = "CenaFreight=" & CenaFreight & ": CenaFact=" & CenaFact
    On Error GoTo ERR2
    sc.ExecuteStatement (str)
End If
On Error GoTo ERR1
nomenkFormula = Round(sc.Eval(tmpStr), 2)

GoTo en
ERR2:
  nomenkFormula = "error в  CenaFreight или cenaFact"
  GoTo en
ERR1:
  nomenkFormula = "error: " & Error
'  If noMsg = "" Then
'    MsgBox Error & " - при выполнении формулы '" & tbNomenk!formula & _
'    "' для номенклатуры '" & tbNomenk!nomNom & "' (" & tmpStr & ")", , _
'    "Ошибка 314 - " & Err & ":  " '##314
 ' End If
en:
If noOpen = "" Then tbNomenk.Close
End Function



Public Sub initProdCategoryBox(ByRef lbPrWeb, Optional extended As Boolean = False)
Dim Table As Recordset
Dim Name As String

    ' Сначала удаляем старые значения
    While lbPrWeb.ListCount
        lbPrWeb.RemoveItem (0)
    Wend
    
    sql = "select * from GuideProdCategory"
    
    Set Table = myOpenRecordSet("##72", sql, dbOpenForwardOnly)
    If Table Is Nothing Then myBase.Close: End
    
    lbPrWeb.AddItem "", 0
    While Not Table.EOF
        If extended Then
            Name = Table!nameRu & " (" & Table!sysname & ")"
        Else
            Name = Table!sysname
        End If
        lbPrWeb.AddItem Table!sysname
        lbPrWeb.ItemData(lbPrWeb.ListCount - 1) = Table!prodCategoryId
        Table.MoveNext
    Wend
    Table.Close
    
    If TypeOf lbPrWeb Is ListBox Then
        lbPrWeb.Height = 225 * lbPrWeb.ListCount
    End If

End Sub


Public Function initFomulConstats() As Boolean
    initFomulConstats = True
    On Error GoTo er
    ' init the Global constants to use its in formulas
    sql = "select * from GuideConstants"
    Set tbGuide = myOpenRecordSet("##0.3", sql, dbOpenForwardOnly)
    If tbGuide Is Nothing Then GoTo er
    While Not tbGuide.EOF
        Dim initStr As String
        initStr = tbGuide!Constants & "=" & CDbl(tbGuide!Value)
        sc.ExecuteStatement (initStr)
        tbGuide.MoveNext
    Wend
    tbGuide.Close
    Exit Function
er:
    initFomulConstats = False
End Function

Function makeCsvFilePath(csvFileName As String) As String

    Dim csvPath As String
    
    csvPath = getEffectiveSetting("ProductsPath", ".")
    If csvPath <> "" And Right(csvPath, 1) <> "\" Then
        csvPath = csvPath & "\"
    End If
    csvPath = csvPath & csvFileName
    
    If Dir$(csvPath) <> "" Then
        If MsgBox("По кнопке 'Дa'(Yes) будет перезаписан файл " & csvPath, vbDefaultButton2 Or vbYesNo, "Подтвердите запись") = vbNo Then
            Exit Function
        End If
    End If
    
    makeCsvFilePath = csvPath
End Function



' динамически располагает форму мемо так, чтобы была видна и ячейка грида
Sub positionMemoFrame(ByRef Grid As MSFlexGrid, ByRef frmRemark As Frame)

Dim meForm As VB.Form
Set meForm = frmRemark.Parent

    If Grid.CellTop + frmRemark.Height < Grid.Height Then
        frmRemark.Top = Grid.CellTop + Grid.Top + Grid.CellHeight
    Else
        frmRemark.Top = Grid.CellTop + Grid.Top - frmRemark.Height '+ Grid.CellHeight
    End If
    Dim shiftRight As Long
    If frmRemark.Top < 0 Then
        frmRemark.Top = 0
        shiftRight = Grid.CellWidth
    Else
    
    End If
    
    frmRemark.Left = Grid.CellLeft + Grid.Left + shiftRight
    If frmRemark.Left + frmRemark.Width > meForm.Width Then
        If frmRemark.Top = 0 Then
            frmRemark.Left = Grid.CellLeft + Grid.Left - frmRemark.Width
        Else
            frmRemark.Left = meForm.Width - frmRemark.Width
        End If
    End If
    If frmRemark.Left < 0 Then
        frmRemark.Left = 0
    End If

End Sub

Public Sub FlexGridColumnColor(FlexGrid As MSFlexGrid, ByVal lngColumn As Long, ByVal lngColor As Long)
    Dim lngPrevCol As Long
    Dim lngPrevColSel As Long
    Dim lngPrevRow As Long
    Dim lngPrevRowSel As Long
    Dim lngPrevFillStyle As Long
    If lngColumn > FlexGrid.Cols - 1 Then
        Exit Sub
    End If
    lngPrevCol = FlexGrid.col
    lngPrevRow = FlexGrid.row
    lngPrevColSel = FlexGrid.ColSel
    lngPrevRowSel = FlexGrid.RowSel
    lngPrevFillStyle = FlexGrid.FillStyle
    FlexGrid.col = lngColumn
    FlexGrid.row = FlexGrid.FixedRows
    FlexGrid.ColSel = lngColumn
    FlexGrid.RowSel = FlexGrid.Rows - 1
    FlexGrid.FillStyle = flexFillRepeat
    FlexGrid.CellBackColor = lngColor
    FlexGrid.col = lngPrevCol
    FlexGrid.row = lngPrevRow
    FlexGrid.ColSel = lngPrevColSel
    FlexGrid.RowSel = lngPrevRowSel
    FlexGrid.FillStyle = lngPrevFillStyle
End Sub

Public Sub FlexGridRowColor(FlexGrid As MSFlexGrid, ByVal lngRow As Long, ByVal lngColor As Long)
    Dim lngPrevCol As Long
    Dim lngPrevColSel As Long
    Dim lngPrevRow As Long
    Dim lngPrevRowSel As Long
    Dim lngPrevFillStyle As Long
    If lngRow > FlexGrid.Rows - 1 Then
        Exit Sub
    End If
    lngPrevCol = FlexGrid.col
    lngPrevRow = FlexGrid.row
    lngPrevColSel = FlexGrid.ColSel
    lngPrevRowSel = FlexGrid.RowSel
    lngPrevFillStyle = FlexGrid.FillStyle
    FlexGrid.col = FlexGrid.FixedCols
    FlexGrid.row = lngRow
    FlexGrid.ColSel = FlexGrid.Cols - 1
    FlexGrid.RowSel = lngRow
    FlexGrid.FillStyle = flexFillRepeat
    FlexGrid.CellBackColor = lngColor
    FlexGrid.col = lngPrevCol
    FlexGrid.row = lngPrevRow
    FlexGrid.ColSel = lngPrevColSel
    FlexGrid.RowSel = lngPrevRowSel
    FlexGrid.FillStyle = lngPrevFillStyle
End Sub

