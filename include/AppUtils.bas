Attribute VB_Name = "AppUtils"
Option Explicit

Dim quantity As Long

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

Public Sub initListbox(ByVal InitSql As String, ByVal lb As listBox, keyField As String, valueField As String, Optional ListBoxMode As Integer = 0)

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

Function PriceToCSV(ByRef Frm As VB.Form, Regim As String, csvFile As String, curRate As Double, Optional prodCategoryId As Integer = 0, Optional commonRabbat As Single = 1) As String
Dim izdeliePrices(4) As Single, I As Integer
Dim csvRow As String, RPF_Rate As Single
Dim ret As String
Dim strCrit() As String, critIndex As Integer


Frm.MousePointer = flexHourglass


    sql = "SELECT p.prId, p.prName, p.prDescript, p.prSize, p.Cena4, p.page, p.rabbat as productRabbat" _
        & vbCr & " , f.formula" _
        & vbCr & " , w.prId as hasWeb " _
        & vbCr & " , s.gain2, s.gain3, s.gain4 " _
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
    
    If Regim = "agency" Or Regim = "default" Then
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

    Set tbProduct = myOpenRecordSet("##415", sql, dbOpenDynaset)
    If Not tbProduct Is Nothing Then
        
        On Error GoTo ERR1
                
        If Not tbProduct.BOF Then
            Open csvFile For Output As #1
            Print #1, "cod" & DLM & "size" & DLM & "description" _
                    & DLM & "price1" & DLM & "price2" & DLM & "price3" & DLM & "price4" & DLM & "page"
            
            While Not tbProduct.EOF
            
            
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
                    csvRow = csvRow & DLM & Round(izdeliePrices(I), 2)
                    izdeliePrices(I) = 0
                Next I
    
                csvRow = csvRow & DLM & tbProduct!Page
    
                Print #1, csvRow
                
                tbProduct.MoveNext
            Wend
            Close #1
        End If
        tbProduct.Close
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


Public Function CsvProductPrices(ByRef izdeliePrices() As Single, ByRef RPF_Rate As Single, Regim As String, curRate As Double, commonRabbat As Single) As String
    Dim ret As String
    Dim baseCena As String
    
    gProductId = tbProduct!prId
    If gProductId = 1192 Then
        gProductId = gProductId
    End If
    
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
        RPF_Rate = baseCena / RPF_Rate


    ElseIf Regim = "default" Or Regim = "pricePM" Then
        baseCena = tCena4
    End If
    
End Function




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
        S = V * tbNomenk!quantity / tbNomenk!perlist
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
tmpStr = tbNomenk.fields("formula" & Web)
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
    tbNomenk!ves & ": STAVKA=" & tbNomenk!STAVKA
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
    
    If TypeOf lbPrWeb Is listBox Then
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

