Attribute VB_Name = "Common"
Option Explicit
'Проeкт\Свойства\Создание\Аргументы компиляции:
' - onErrorOtlad = 1 ' ловля редких err

' Размер страницы (к-во строк) при печати накладной.
' Задается в настройках, используется в модуле Nakladna.frm

Public gCfgOrderPageSize As Integer

Public isOrders As Boolean
Public isWerkOrders As Boolean
Public isZagruz As Boolean
Public isFindFirm As Boolean
Public mainTitle As String
Public flReportArhivOrders As Boolean

Public tbOrders As Recordset
Public tqOrders As Recordset
Public tbSystem As Recordset
Public tbFirms As Recordset
Public tbNomenk  As Recordset
Public tbProduct As Recordset
Public tbDMC As Recordset
Public tbDocs As Recordset
Public tbGuide As Recordset
Public tbSeries As Recordset
Public Node As Node
Public isAdmin As Boolean


Public isBlock As Boolean
Public Equip() As String
Public Werk() As String
Public werkSourceId() As Integer

'Public gEquipId As Integer
Public Const lenStatus = 20
Public statId(lenStatus) As Integer
Public status(lenStatus) As String
Public lenProblem As Integer
Public Problems(20) As String
Public manId() As Integer '$$7
Public Manag() As String  '
Public Managers() As MapEntry

Public insideId() As String
Public Const begWerkProblemId = 10 ' начало цеховых проблем в справочнике
Public neVipolnen As Double, neVipolnen_O As Double
Public maxDay As Integer ' число дней в реестре
Public befDays As Integer ' число дней до даты реестра (когда сменилась дата), вычисляется как разница между текущим днем, и днем из последнего сохраненного System.lastNumorder
Public webSvodkaPath As String
Public webLoginsPath As String
Public webNomenks As String '- есть и в sTime
Public webProducts As String

' Перенесено из cfg.frm
Public loginsPath As String
Public SvodkaPath As String
Public NomenksPath As String
Public ProductsPath As String


'Public baseNamePath As String
Public begDate As Date ' Дата вступительных остатков
Public logFile As String
Public dostup As String
Public sessionCurrency As Integer
Public otlad As Variant
Public tbSize As Integer
Public cErr As String 'позволяет выявить место возникновения Err, если по
                      'всем местам сообщение об Err выдает один MsgBox
Public zakazNum As Long  ' кол-во заказов в  Mен.реестре
Public gNzak As String  ' тек номер заказа
'Public gFirmId As Integer
Public gFirmId As String
Public gProductId As String
Public gProduct As String
Public gDocDate As Date
Public gSeriaId As String
Public gKlassId As String
Public gNomNom As String
Public numDoc As Long, numExt As Integer

Public oldValue As String 'старое значение поля, измененного последним
Public curDate As Date
Public lastYear As Integer

Public begDay As Integer ' день первого куска заказа
Public endDay As Integer ' день последнего куска заказа
Public begDayMO As Integer ' день первого куска МО заказа
Public endDayMO As Integer ' день последнего куска МО заказа
Public flEdit As String ' редактируется ресурс
Public Nstan As Double
Public KPD As Double
Public newRes As Double ' смена по умолчанию
Public nr As Double ', dr As Double 'убываощие ном. и доп. ресурсы
Public isLive As Boolean ' флаг - заказ живой
Public zagAll As Double, zagLive As Double

Public table As Recordset '
Public myQuery As QueryDef
Public sql As String      ' коллективного пользования
Public strWhere As String '
Public sortGrid As MSFlexGrid
Public trigger As Boolean '
Public tmpDate As Date    '
Public tmpStr As String
Public tmpVar As Variant
Public tmpSng As Double
Public day As Integer     '
Public tiki As Integer    '
Public flClickDouble As Boolean
Public ClickItem As ListItem
Public noClick As Boolean 'делаем True, чтобы подавить onClick lb
Public bilo As Boolean
Public cep As Boolean  'активизирована какая либо цепь послед-ных действий
Public oldCellColor As Long
Public prExt As Integer, pType As String
Public gridIsLoad As Boolean

Public orColNumber As Integer ' число колонок в Orders
Public orSqlWhere() As String
Public orSqlFields() As String  '
Public orNomZak As Integer, orWerk As Integer, orEquip As Integer, orData As Integer, orTema As Integer
Public orMen As Integer, orStatus As Integer, orProblem As Integer
Public orDataRS As Integer, orFirma As Integer, orDataVid As Integer
Public orVrVid As Integer, orVrVip As Integer, orM As Integer, orO As Integer
Public orType As Integer, orInvoice As Integer
Public orMOData As Integer, orMOVrVid As Integer, orOVrVip As Integer
Public orLogo As Integer, orIzdelia As Integer, orZakazano As Integer
Public orZalog As Integer, orNal As Integer, orRate As Integer
Public orOplacheno As Integer, orOtgrugeno As Integer, orLastMen As Integer
Public orVenture As Integer
Public orlastModified As Integer
Public orBillId As Integer
Public orVocnameId As Integer
Public orServername As Integer

Public NN() As String, QQ() As Double ' откатываемая номенклатура и кол-во
Public QQ2() As Double, QQ3() As Double
Public skladId As Integer

Private Const dhcMissing = -2 'нужна для quickSort
Public Const cDELLwidth = 19200 ' это порог а мах = 19290

Public Const dcSourId = 0 ' скрыт
Public Const dcDate = 1
Public Const dcNumDoc = 2
Public Const dcSour = 3
Public Const dcDest = 4
Public Const dcVenture = 6
Public Const dcNote = 5

'Grid в FirmComtex
Public Const fcId = 0
Public Const fcFirmName = 1
Public Const fcInn = 2
Public Const fcOkonx = 3
Public Const fcOkpo = 4
Public Const fcKpp = 5
Public Const fcAddress = 6
Public Const fcPhone = 7

Public Const fcFormatString = _
  "|< Название  фирмы" _
& "|>ИНН" _
& "|>ОКОНХ" _
& "|>ОКПО" _
& "|>КПП" _
& "|<Адрес" _
& "|<Телефон" _


Public Const gfNazwFirm = 1
Public Const gfM = 2
Public Const gfKategor = 3
Public Const gfSale = 4
Public Const gf2001 = 5
Public Const gf2002 = 6
Public Const gf2003 = 7
Public Const gf2004 = 8
Public Const gfFIO = 9
Public Const gfTlf = 10
Public Const gfFax = 11
Public Const gfEmail = 12
Public Const gfType = 13
Public Const gfKatalog = 14
Public Const gfLevel = 15
Public Const gfAdres = 16 'примеч
Public Const gfAtr1 = 17
Public Const gfAtr2 = 18
Public Const gfAtr3 = 19
Public Const gfLogin = 20
Public Const gfPass = 21
Public Const gfId = 22

Public Const chNomZak = 1
Public Const chM = 2
Public Const chEquip = 3
Public Const chStatus = 4
Public Const chVrVip = 5
Public Const chProcVip = 6
Public Const chProblem = 7
Public Const chDataVid = 8
Public Const chVrVid = 9
Public Const chDataRes = 10
Public Const chFirma = 11
Public Const chLogo = 12
Public Const chIzdelia = 13
Public Const chKey = 14

Public Const zgPrinato = 1
Public Const zgNomRes = 2
'Public Const zgDopRes = 3
'Public Const zgRaspred = 4
Public Const zgResurs = 3
Public Const zgZagruz = 4
Public Const zgOstatki = 5
Public Const zgLive = 6
'прием заказа
Public Const zkPrinato = 1
Public Const zkFirmKolvo = 2
Public Const zkResurs = 3
Public Const zkMzagr = 4
Public Const zkMbef = 5
Public Const zkHide = 6 'невидим
Public Const zkMost = 7
Public Const zkCzagr = 8
Public Const zkCost = 9
Public Const zkCliv = 10
'номенклатура в документе из sProducts Grid2
Public Const fnNomNom = 1
Public Const fnNomName = 2
Public Const fnEdIzm = 3
Public Const fnQuant = 4

'Grid5 в sProducts и Otgruz
Public Const prId = 0
Public Const prType = 1
Public Const prName = 2
Public Const prDescript = 3
Public Const prEdizm = 4
Public Const prCenaEd = 5
Public Const prQuant = 6
Public Const prSumm = 7
Public Const prEtap = 8
Public Const prEQuant = 9
Public Const prOutQuant = 8
Public Const prOutSum = 9
Public Const prNowQuant = 10
Public Const prNowSum = 11

'массив начал зел.коридора начиная от дня прошлого запуска проги до Сегодня
Public stDays() As Integer        ' включая дни пропуска (Сб,Вс,праздники)
Public stDay As Integer 'равен последнему stDays(Сегодня)
                            
Public nomRes() As Double
Public delta() As Double
Public tmp() As Double
Public tmpL() As Long
Public ost() As Double, befOst() As Double

' список выбранных позиций в таблице предметов к заказу
' (по CtrlLeftClick) в sProducts.Grid5
Public selectedItems() As Long
Public Const otladColor = &H80C0FF

Public Const CC_RUBLE As Integer = 1
Public Const CC_UE As Integer = 2

' на сколько нужно увеличивать ширину колонок, если выбраны рубли
Public Const ColWidthForRuble As Single = 1.3



Function tuneCurencyAndGranularity(tunedValue, currentRate, valueCurrency As Integer, Optional quantity As Double = 1, Optional perlist As Long = 1) As Double
    Dim Left As String, StatusId As String, Outdatetime As String, Rollback As String, IsEmpty As String, ExeName As String
    '
    Dim totalInRubles As Double
    Dim singleInRubles As Double
    Dim totalInUE As Double
    Dim singleInUE As Double
    '
    If valueCurrency = CC_RUBLE Then
        singleInRubles = Round(CDbl(tunedValue) / CDbl(quantity), 2)
    Else
        singleInRubles = Round(CDbl(tunedValue) / CDbl(quantity) * CDbl(currentRate), 2)
    End If
    totalInRubles = singleInRubles * quantity
    totalInUE = totalInRubles / currentRate
    singleInUE = totalInUE / quantity
    tuneCurencyAndGranularity = totalInUE

End Function



Function rated(geld, rate) As Variant
    If IsNull(geld) Then
        rated = Null
        Exit Function
    End If
    If sessionCurrency = CC_RUBLE Then
        rated = CDbl(geld) * CDbl(rate)
    Else
        rated = geld
    End If
End Function

Function serverIsAccessible(ventureName As String) As Boolean
Dim I As Integer

    serverIsAccessible = False
    For I = 0 To Orders.lbVenture.ListCount
        If Orders.lbVenture.List(I) = ventureName Then
            serverIsAccessible = True
            Exit For
        End If
    Next I
    
End Function


'Если первый пораметр ="W.." - не выдавать Err по невып-ю Where, а все
'параметры обнулить, если для всех них нуль это возможное значение, то в sql
'м. задать константу "1" и принять ее в I. Тогда если I=0 то была Err Where
'$odbc15$
Function byErrSqlGetValues(ParamArray val() As Variant) As Boolean
Dim tabl As Recordset, I As Integer, maxi As Integer, str As String, c As String

byErrSqlGetValues = False
maxi = UBound(val())
If maxi < 1 Then
    wrkDefault.Rollback
    MsgBox "мало параметров для п\п byErrSqlGetValues()"
    Exit Function
End If
str = CStr(val(0)): c = Left$(str, 1)
If c = "W" Then str = Mid$(str, 2)
Set tabl = myOpenRecordSet(str, CStr(val(1)), dbOpenForwardOnly) 'dbOpenDynaset)$#$
'If tabl Is Nothing Then Exit Function
If tabl.BOF Then
    If c = "W" Then
        For I = 2 To maxi: val(I) = 0: Next I
        GoTo EN1
    ElseIf c = "w" Then
        GoTo EN1
    Else
'        msgOfEnd CStr(val(0)), "Нет записей удовлетворяющих Where."
        wrkDefault.Rollback
        MsgBox "Нет записей удовлетворяющих Where!", , "Error-" & str
        GoTo EN2
    End If
End If
'tabl.MoveFirst $#$
For I = 2 To maxi
    str = TypeName(val(I))
    If (str = "Single" Or str = "Integer" Or str = "Long" Or str = "Double") _
    And IsNull(tabl.fields(I - 2)) Then
        val(I) = 0
    ElseIf str = "String" And IsNull(tabl.fields(I - 2)) Then
        val(I) = ""
    ElseIf str = "Date" And IsNull(tabl.fields(I - 2)) Then
        'do nothing the date remain quasi-null
'        val(I) = tabl.Fil
    Else
        val(I) = tabl.fields(I - 2)
    End If
Next I
EN1:
byErrSqlGetValues = True
EN2:
tabl.Close
End Function

Sub clearGrid(Grid As MSFlexGrid, Optional fixed As Integer = 1)
If fixed = 1 Then
    Grid.Rows = 2
    clearGridRow Grid, 1
Else
    Grid.Rows = 3
    clearGridRow Grid, 2
End If
End Sub

Sub clearGridRow(Grid As MSFlexGrid, row As Long)
Dim il As Long
    noClick = True
    Grid.row = row
    For il = 0 To Grid.Cols - 1
        Grid.col = il
        If il > 0 Then Grid.CellBackColor = Grid.BackColor
        Grid.CellForeColor = Grid.ForeColor
        Grid.CellFontStrikeThrough = False
        Grid.TextMatrix(row, il) = ""
    Next il
    Grid.col = 1
    noClick = False
End Sub

Sub colorGridRow(Grid As MSFlexGrid, row As Long, color As Long)
Dim il As Long
    Grid.row = row
    For il = 0 To Grid.Cols - 1
        Grid.col = il
        If il > 0 Then Grid.CellBackColor = color
    Next il
    Grid.col = 1
End Sub

Sub dayMassLenght(Optional newLen As Integer = 0)
Dim maxLen As Integer

On Error GoTo ERR1
    maxLen = UBound(nomRes)
On Error GoTo 0
    If newLen > maxLen Then
        maxLen = newLen
    Else
        If newLen > 0 Then Exit Sub
        maxLen = maxLen + 10
    End If
    If 511 > maxLen And maxLen > 499 Then _
        MsgBox "Ресурсы программы превысили 500 дней? Сообщите Администратору!", , "Err в dayMassLenght()"
    ReDim Preserve stDays(maxLen)
    ReDim Preserve nomRes(maxLen)
    ReDim Preserve delta(maxLen)
    ReDim Preserve tmp(maxLen)
Exit Sub

ERR1:
If Err = 9 Then
    maxLen = 0
    Resume Next
Else
    MsgBox Error, , "Ошибка 17-" & Err & ":  " '##17
    End
End If

End Sub

Sub myRedim(Mass As Variant, newLen As Integer)
Dim maxLen As Integer

maxLen = 0
On Error Resume Next
maxLen = UBound(Mass)
On Error GoTo 0
If newLen < maxLen Then Exit Sub
ReDim Preserve Mass(newLen + 20)
End Sub

Sub delay(tau As Double)
Dim S As Double
    S = Timer
    While Timer - S < tau ' 1 сек
        DoEvents
    Wend

End Sub

Sub delZakazFromReplaceRS()
sql = "DELETE From ReplaceRS " & "WHERE numOrder = " & gNzak
myExecute "##79", sql, 0 ' удаляем, если есть
End Sub


Sub exitAll()
If isOrders Then Unload Orders
If isWerkOrders Then Unload WerkOrders
If isZagruz Then Unload Zagruz
If isFindFirm Then Unload FindFirm

If sDocs.isLoad Then Unload sDocs

If cfg.isLoad Then Unload cfg '$$2
'If isZagruzM Then Unload ZagruzM
'myBase.Close

End Sub

Function findValInCol(Grid As MSFlexGrid, value, col As Integer) As Boolean
Dim il As Long
findValInCol = False
For il = 1 To Grid.Rows - 1
    If value = Grid.TextMatrix(il, orNomZak) Then
        Grid.TopRow = il
        Grid.row = il
        findValInCol = True
        Exit For
    End If
Next il

End Function
        
Function findExValInCol(Grid As MSFlexGrid, value As String, _
            col As Integer, Optional pos As Long = -1) As Long
Dim il As Long, str  As String, beg As Long

If pos < 1 Then
    beg = 1
Else
    beg = pos
End If
value = UCase(value)
For il = beg To Grid.Rows - 1
    str = UCase(Grid.TextMatrix(il, col))
    If InStr(str, value) > 0 Then
        Grid.TopRow = il
        Grid.row = il
        findExValInCol = il
        Exit Function
    End If
Next il
findExValInCol = -1

End Function

'$odbc08$
Function existValueInTableFielf(ByVal value As Variant, tabl As String, field) As Boolean
Dim table As Recordset

existValueInTableFielf = False

If Not IsNumeric(value) Then value = "'" & value & "'"

sql = "SELECT " & field & " From " & tabl & " WHERE (((" & field & ") = " & _
value & "));"
'MsgBox sql
Set table = myOpenRecordSet("##390", sql, dbOpenForwardOnly)
'If table Is Nothing Then myBase.Close: End

If Not table.BOF Then existValueInTableFielf = True

table.Close

End Function

'год ставит на первое место, а день на последнее - для правильной сортировки
Function yymmdd(dateStr As String) As String
yymmdd = Right$(dateStr, 2) & "." & Mid$(dateStr, 4, 2) & "." & Left$(dateStr, 2)
End Function


Function getValueFromTable(tabl As String, field As String, Where As String) As Variant
Dim table As Recordset

getValueFromTable = Null
sql = "SELECT " & field & " as fff  From " & tabl & _
      " WHERE " & Where
Set table = myOpenRecordSet("##59.1", sql, dbOpenForwardOnly)
If table Is Nothing Then Exit Function
If Not table.BOF Then getValueFromTable = table!fff
table.Close
End Function


Function getNextDay(tmpDay As Integer) As Integer

tmpDate = DateAdd("d", tmpDay - 1, curDate)
'tmpDate = CurDate + tmpDay - 1
day = Weekday(tmpDate)
If day = vbFriday Then
    getNextDay = tmpDay + 3
ElseIf day = vbSunday Then
    getNextDay = tmpDay + 2
Else
    getNextDay = tmpDay + 1
End If

End Function

Function getStrDocExtNum(nmDoc As Long, nmExt As Integer) As String

If nmExt > 0 And nmExt < 254 Then
    getStrDocExtNum = nmDoc & "/" & nmExt
Else
    getStrDocExtNum = nmDoc
End If

End Function

Function getStrPrEx(name As String, ext As Integer) As String
If ext = 0 Then
    getStrPrEx = name
Else
    getStrPrEx = ext & "/ " & name
End If
End Function


Function getNumFromStr(str As String) As String
Dim I As Integer, ch As String

For I = 1 To Len(str)
    getNumFromStr = Mid$(str, I, 1)
    If Not IsNumeric(getNumFromStr) Then Exit For
Next I
gNzak = Left$(str, I - 1)

End Function
'$odbc10$
Function getResurs(equipId As Integer) As Integer
Dim I As Integer, J As Integer, rMaxDay As Integer, S As Double
' rMaxDay - Resource max day - максимальное значение из таблицы ResursCEH (CO2, etc)

Set tbSystem = myOpenRecordSet("##93", "select * from GuideResurs where equipId = " & equipId, dbOpenForwardOnly)
If tbSystem Is Nothing Then myBase.Close: End
KPD = tbSystem!KPD
Nstan = tbSystem!Nstan
newRes = tbSystem!newRes
tbSystem.Close

sql = "SELECT nomRes from Resurs where equipId = " & equipId & " ORDER BY xDate"
Set table = myOpenRecordSet("##10", sql, dbOpenForwardOnly)
'If table Is Nothing Then Exit Function

'j = -1
'If flEdit <> "" Then _
'    j = Mid$(Zagruz.lv.SelectedItem.key, 2)
rMaxDay = 0
On Error GoTo ERR1
'If Not table.BOF Then
' table.MoveFirst
' Do
While Not table.EOF
    rMaxDay = rMaxDay + 1
    If rMaxDay = J Then
'        table.Edit
'            table!nomRes = Zagruz.tbMobile.Text
'        table.Update
    End If
    nomRes(rMaxDay) = table!nomRes
    table.MoveNext
' Loop While Not table.EOF
Wend
table.Close
'End If

addDays max(stDay, rMaxDay) 'добавляем дни, если Дата Выдачи всех заказов
                            'ближе чем stDay или rMaxDay(например заказов нет)
For I = rMaxDay + 1 To maxDay
'    table.AddNew
    tmpDate = DateAdd("d", I - 1, curDate)
    day = Weekday(tmpDate)
'    table!Date = Format(tmpDate, "dd.mm.yy")
    If day = vbSunday Or day = vbSaturday Then
'        table!nomRes = 0
        nomRes(I) = 0
    Else
'        table!nomRes = newRes
        nomRes(I) = newRes
    End If
'    table.Update
Next I
'table.Close

'*********************** убывающий ресурс **************************
S = Timer / 3600:
I = Int(S)
If I < 9 Then
ElseIf I < 13 Then
    nr = Round(nomRes(1) - S + 9, 1)
ElseIf I < 14 Then
    nr = Round(nomRes(1) - 4, 1)
Else
    nr = Round(nomRes(1) - S + 10, 1)
End If
If nr < 0 Then
    nr = 0
End If

getResurs = maxDay '1:
Exit Function

ERR1:
If Err = 9 Then
    dayMassLenght 'корректируем размерности, если надо
    Resume
Else
    MsgBox Error, , "Ошибка 18-" & Err & ":  " '##18
    myBase.Close: End
End If

End Function
'$NOodbc$
'выдает "error"- если нарушен формат дат (не следует запускать SQL) .
'reg="" -  выдает аргумент для WHERE для промежутка между датами
'          либо "" - ограничния во WHERE по дате не требуется(с учетом begDate и CurDate)
'          либо "error" если даты не пересекаются
'reg<>"" - выдает аргумент для WHERE для промежутка До startDate
'          либо "" если startDate раньше begDate(не следует запускать SQL)
Function getWhereByDateBoxes(Frm As Form, dateField As String, _
begDate As Date, Optional reg As String = "") As String

Dim str As String, ckStart As Boolean, ckEnd  As Boolean

getWhereByDateBoxes = "": str = "":

ckStart = False: ckEnd = False
On Error Resume Next ' на случай, если в этой форме у дат нет флажков
If Frm.ckEndDate.value > 0 Then ckEnd = True  'то они как бы установлены
If Frm.ckStartDate.value > 0 And Frm.ckStartDate.Visible Then ckStart = True
On Error GoTo 0

If ckStart Then
    If Not isDateTbox(Frm.tbStartDate) Then GoTo ERRd  'tmpDate
End If
If reg = "" Then ' если период Между
    If DateDiff("d", begDate, tmpDate) > 0 And ckStart Then _
        str = "(" & dateField & ") >=" & Format(tmpDate, "'yyyy-mm-dd'")
    If ckEnd Then
      If Not isDateTbox(Frm.tbEndDate) Then GoTo ERRd
      If ckStart Then
        If DateDiff("d", Frm.tbStartDate.Text, tmpDate) < 0 Then
          MsgBox "Начальная дата периода загрузки не должна превышать конечную ", , "Предупреждение"
ERRd:     getWhereByDateBoxes = "error"
          Exit Function
        End If
      End If
      If DateDiff("d", tmpDate, curDate) > 0 Then getWhereByDateBoxes = _
          "(" & dateField & ")<='" & Format(tmpDate, "yyyy-mm-dd") & " 11:59:59 PM'"
    End If
ElseIf ckStart Then ' если период До
    If DateDiff("d", begDate, tmpDate) <= 0 Then Exit Function
    tmpDate = DateAdd("d", -1, tmpDate) ' "-1" день т.к. далее "+ 23ч59м59с
    If DateDiff("d", tmpDate, curDate) > 0 Then getWhereByDateBoxes = _
        "(" & dateField & ")<='" & Format(tmpDate, "'yyyy-mm-dd'") & " 11:59:59 PM'"
End If
If str <> "" And getWhereByDateBoxes <> "" Then
    getWhereByDateBoxes = str & " AND " & getWhereByDateBoxes
Else
    getWhereByDateBoxes = str & getWhereByDateBoxes
End If
End Function

Sub textBoxInGridCell(tb As TextBox, Grid As MSFlexGrid)
    tb.Width = Grid.CellWidth
'    tb.Text = Grid.TextMatrix(mousRow, mousCol)
    tb.Text = Grid.TextMatrix(Grid.row, Grid.col)
    tb.Left = Grid.CellLeft + Grid.Left
    tb.Top = Grid.CellTop + Grid.Top
    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)
    tb.Visible = True
    tb.SetFocus
    tb.ZOrder
    Grid.Enabled = False 'иначе курсор по ней бегает
End Sub

Sub listBoxSelectByText(lb As listBox, obrazec As String)
Dim I As Integer
    
    For I = 0 To lb.ListCount - 1 '
'       noClick = True
        If obrazec = lb.List(I) Then
            lb.Selected(I) = True 'вызывает ложное onClick'
        Else
            lb.Selected(I) = False
        End If
'       noClick = False
    Next I

End Sub

Sub lbDeSelectAll(listBox As listBox)
Dim I As Integer

For I = 0 To listBox.ListCount - 1
    listBox.Selected(I) = False
Next I
End Sub
'$NOodbc$
Function lbToOrSqlWhere(listBox As listBox, col As Integer, Optional _
notAll As String = "") As String
Dim str As String, I As Integer, strWhere As String, beAll As Boolean
Dim beNothing As Boolean
strWhere = ""
beAll = True
beNothing = True
For I = 0 To listBox.ListCount - 1
    If listBox.Selected(I) Then
        If notAll = "byId" Then
            str = I
        Else
            str = listBox.List(I)
        End If
        str = Orders.strWhereByValCol(str, col)
        If strWhere = "" Then
            strWhere = str
        Else
            strWhere = strWhere & " OR " & str
        End If
        beNothing = False
    Else
        beAll = False
    End If
Next I
orSqlWhere(col) = strWhere
If notAll = "byId" Then notAll = ""
'If beAll And notAll = "" Then orSqlWhere(col) = ""
If beAll And notAll = "" Then orSqlWhere(col) = ""

If (beAll Or beNothing) And Not Orders.tbEnable.Visible And col = orStatus Then
    orSqlWhere(col) = "(GuideStatus.Status)<>'закрыт'"
Else
End If
End Function
'$NOodbc$
Sub listBoxInGridCell(lb As listBox, Grid As MSFlexGrid, Optional sel As String = "")
Dim I As Integer
    If Grid.CellTop + lb.Height < Grid.Height Then
        lb.Top = Grid.CellTop + Grid.Top
    Else
        lb.Top = Grid.CellTop + Grid.Top - lb.Height + Grid.CellHeight
    End If
    lb.Left = Grid.CellLeft + Grid.Left
    lb.ListIndex = 0
    If sel <> "" Then
        For I = 0 To lb.ListCount - 1 '
            If Grid.Text = lb.List(I) Then
'                noClick = True
                lb.ListIndex = I 'вызывает ложное onClick
'                noClick = False
                Exit For
            End If
        Next I
    End If
    lb.Visible = True
    lb.ZOrder
    lb.SetFocus
    Grid.Enabled = False 'иначе курсор по ней бегает
'    lbIsActiv = True
End Sub

Function LoadNumeric(Grid As MSFlexGrid, row As Long, col As Integer, _
        val As Variant, Optional myErr As String = "", Optional fmt As String) As Double
 If IsNull(val) Then
    Grid.TextMatrix(row, col) = ""
    LoadNumeric = 0 ' для log файла
    If myErr <> "" Then msgOfZakaz (myErr)
 Else
    LoadNumeric = Round(val, 2)
    If Round(val, 2) <> Round(val, 0) Then
        If IsMissing(fmt) Then
            Grid.TextMatrix(row, col) = Format(LoadNumeric, "####.00")
        Else
            Grid.TextMatrix(row, col) = Format(LoadNumeric, fmt)
        End If
    Else
        Grid.TextMatrix(row, col) = LoadNumeric
    End If
 End If
End Function

Function LoadDate(Grid As MSFlexGrid, row As Long, col As Integer, _
val As Variant, formatStr As String, Optional myErr As String = "") As String
Dim str As String

 If IsNull(val) Then
    Grid.TextMatrix(row, col) = ""
    LoadDate = "" ' для log файла
    If myErr <> "" Then
        msgOfZakaz (myErr)
        Grid.TextMatrix(row, col) = "??"
    End If
 Else
    LoadDate = Format(val, formatStr)
    If LoadDate = "00" Then LoadDate = "" '    замена для 00 часов
    Grid.TextMatrix(row, col) = LoadDate
 End If
End Function



'$NOodbc$
Sub initOrCol(colNum As Integer, Optional field As String = "")
orColNumber = orColNumber + 1
colNum = orColNumber
ReDim Preserve orSqlFields(orColNumber + 1)
orSqlFields(orColNumber) = field
End Sub



'$odbc10$
Sub Main()
Dim I As Integer, S As Double, str As String, str1 As String, str2 As String
Dim isXP As Boolean
If App.PrevInstance = True Then
    MsgBox "Программа уже запущена", , "Error"
    End
End If

ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): ReDim QQ3(0) 'чтобы Ubound никогда не давала Err

flReportArhivOrders = False
ReDim tmpL(0)

cfg.isLoad = False  '$$2
loadEffectiveSettingsApp
dostup = getEffectiveSetting("dostup")
sessionCurrency = getEffectiveSetting("currency", CC_RUBLE)
loginsPath = getEffectiveSetting("loginsPath")
SvodkaPath = getEffectiveSetting("SvodkaPath")
NomenksPath = getEffectiveSetting("NomenksPath")
ProductsPath = getEffectiveSetting("ProductsPath")

initLogFileName
isAdmin = getEffectiveSetting("dostup", "") = "a"

checkReloadCfg

isXP = (Dir$("C:\WINDOWS\net.exe") = "") 'в XP нет файла
On Error GoTo ERRs ' не дает Err если в сети не б.найден server, хотя из под DOS дает сист.Err=53
otlad = getEffectiveSetting("otlad")

On Error GoTo 0
If dostup = "с" Then dostup = "c"
If dostup = "0" Then
    I = 5 / I  'проверка наличия библиотеки сообщений
    End
End If
If dostup <> "a" And dostup <> "m" And dostup <> "" And dostup <> "b" _
And dostup <> "c" And dostup <> "y" _
And dostup <> "s" Then '$$$ceh
    MsgBox "'" & dostup & "' - неверный параметр запуска!", , ""
    End
End If

baseOpen
    
    ' Используется при определении того, произошло ли на сервере возникновение Бизнес Проблемы или нет.
    sql = "create variable @issueMarker varchar(32)"
    If myExecute("##issueMarker", sql, 0) = 0 Then End

    sql = "create variable @managerId varchar(20)"
    If myExecute("##@managerId", sql, 0) = 0 Then End
    
mainTitle = getMainTitle

If Not IsEmpty(otlad) Then '
  webSvodkaPath = "C:\WINDOWS\TEMP\svodkaW."
  webLoginsPath = "C:\WINDOWS\TEMP\logins."

Else
    webSvodkaPath = SvodkaPath          '$$2
    webLoginsPath = loginsPath          '
End If

On Error GoTo 0

'проверка настройки Win98
str = "05.08.2004"
tmpDate = str
If str <> Format(tmpDate, "dd.mm.yyyy") Then ' для "Хорватский" в Win98 не совпадет
    str = "Региональные стандарты " & Chr(151) & " установите  'Русский'."
    GoTo AA
End If
'begDate = "01.01.2003"
sql = "SELECT begDate, lastYear From System" ' WHERE (((System.begDate) Like '*##.##.20##*'));"
Set tbSystem = myOpenRecordSet("##181", sql, dbOpenForwardOnly)
'If tbSystem Is Nothing Then myBase.Close: End
If tbSystem.BOF Then
    tbSystem.Close: myBase.Close
    str = "Дата\Краткий формат " & Chr(151) & " установите  'дд.ММ.гггг'."
    '"Хорватский" в XP попадпет сюда
    If isXP Then str = str & " Также проверте, что установлен 'Русский' регион."
AA: MsgBox "Пуск\Настройка\Панель Управления\Язык и стандарты\" & str, , _
    "Для правильной работы программы необходимо выполнить настройку Win98:"
    End
Else
    lastYear = tbSystem!lastYear
    begDate = tbSystem!begDate
End If
tbSystem.Close

'не менять порядок след. 3х строчек
nextDayDetect ' здесь определ-ся CurDate
stDay = startDays() ' в т.ч. устанавливаем начальные размерности dayMassLenght
If befDays <> 0 Then nextDay ' перевод базы на новую дату

checkNextYear '$$3 если сменился год - пересчет статистики посещений

' обрезание лог файла*********************************************
'If Not (dostup = "c" Or dostup = "y") Then
If dostup = "a" Or dostup = "m" Or dostup = "" Or dostup = "b" Then
 'logFile = "C:\Windows\Orders" ' без расширения
 logFile = App.path & "\" & App.ExeName
 str2 = logFile & "$.log" ' временный файл
 logFile = logFile & ".log"
 
 On Error GoTo ENop
 Open logFile For Input As #2
 Open str2 For Output As #3
 While Not EOF(2)
    Input #2, str
    I = InStr(str, vbTab)
    If I < 9 Then GoTo ENlog
    str1 = Left$(str, I - 1)
    If Not IsDate(str1) Then GoTo ENlog
    If DateDiff("d", str1, curDate) <= 7 Then Print #3, str ' удаляем > 7ми дней давности
 Wend
ENlog:
 Close #2
 Close #3
 Kill logFile
 Name str2 As logFile
End If '***************************************
ENop:
isBlock = False
noClick = False

Set table = myOpenRecordSet("##05", "GuideStatus", dbOpenForwardOnly)
While Not table.EOF
    If table!StatusId > lenStatus + 1 Then
        MsgBox "Err в Orders\FormLoad"
        End
    End If
    status(table!StatusId) = table!status
    table.MoveNext
Wend
table.Close

Set table = myOpenRecordSet("##04", "GuideProblem", dbOpenForwardOnly)
'If table Is Nothing Then myBase.Close: End

For I = 0 To 20
    Problems(I) = "no"
Next I
lenProblem = -1
CC:
    If lenProblem < table!ProblemId Then lenProblem = table!ProblemId
    If table!ProblemId > 20 Then
        MsgBox "число проблем в базе превысило 20"
        End
    End If
    Problems(table!ProblemId) = table!problem
    table.MoveNext
    If Not table.EOF Then GoTo CC
table.Close

'Проверка режима совместной работы с Комтех
CheckIntegration

If dostup = "y" Then
    WerkOrders.idWerk = 1: WerkOrders.Show
ElseIf dostup = "c" Then    '$$$ceh
    WerkOrders.idWerk = 2: WerkOrders.Show
Else
    Orders.Show
End If
Exit Sub
ERRf:
MsgBox "Пуск\Настройка\Панель Управления\Язык и стандарты\Числа\" & _
      "Разделитель целой и дробной частей числа " & Chr(151) & _
      " Установите точку вместо запятой!", , "Для правильной работы " & _
      "программы необходимо выполнить настройку Win98: "
End

ERRs:
MsgBox "Система не смогла синхронизировать часы", , "Сообщите администратору!"
Resume Next

End Sub


Sub CheckIntegration()
Dim servers As Recordset
Dim fromComtexRS As Recordset
Dim msgOk As VbMsgBoxResult
Dim fromComtex As Integer



'sql = "call wf_load_session()"

'If myExecute("##0.1", sql) = 1 Then
    
'End If

' для каждого сервера проверяем условия соответствия режима интеграции
' установленного в настройках и действительной доступности серверов
' на момент запуска программы
' Если обнаружено рассогласование, то выдаем предупреждение
' Предупреждение можно подавлять настройкой в System
sql = "select * from guideVenture"
Set servers = myOpenRecordSet("##0.2", sql, 0)
If servers Is Nothing Then Exit Sub
While Not servers.EOF
    On Error GoTo no_access
    If byErrSqlGetValues("##0.3" _
        , "select get_standalone_remote ('" & servers!sysname & "')" _
        , fromComtex) _
    Then
        
    End If
        
    If fromComtex = 1 And servers!standalone = 0 Then
        msgOk = MsgBox("Сервер """ & servers!ventureName & """ (" & servers!sysname & ") " _
        & " доступен, но он со своей стороны НЕ настроен на режим совместного использования с программой " _
        & vbCr & "Чтобы выйти и исправить ситуацию нажмите кнопку Отмена(Cancel)" _
        & vbCr & "Если же вы все-таки хотите продолжить работу, нажмите кнопку Ок" _
        , vbOKCancel, "Предупреждение")
        
        If msgOk <> vbOK Then myBase.Close: End
         
    ElseIf fromComtex = 0 And servers!standalone = 1 Then
        msgOk = MsgBox("Сервер """ & servers!ventureName & """ (" & servers!sysname & ") " _
        & " доступен и настроен на режим совместной работы с программой." _
        & vbCr & "В тоже время сама программа настроена так, что она не будет работать с этим сервером." _
        & " Поэтому некоторая информации не будет попадать туда." _
        & vbCr & "Чтобы выйти и исправить ситуацию нажмите кнопку Отмена(Cancel)" _
        & vbCr & "Если же вы все-таки хотите продолжить работу, нажмите кнопку Ок" _
        , vbOKCancel, "Предупреждение")
        
        If msgOk <> vbOK Then myBase.Close: End
    
    ElseIf fromComtex = -1 And servers!standalone <> 1 Then
        msgOk = MsgBox("Сервер """ & servers!ventureName & """ (" & servers!sysname & ") " _
        & " НЕ ДОСТУПЕН, хотя в настройках указано, что программа будет работать совместно. " _
        & vbCr & vbCr & " Такой режим может вызывать ошибки в работе программы!" _
        & vbCr & "Чтобы выйти и исправить ситуацию нажмите кнопку Отмена(Cancel)" _
        , vbOKCancel, "Предупреждение")
        
        If msgOk <> vbOK Then myBase.Close: End
    End If
no_access:
cont:
    servers.MoveNext
Wend
servers.Close


End Sub
'$odbc10$
Function statisticReplace(tabl As String) As Boolean

statisticReplace = False

On Error GoTo EN1
sql = "UPDATE " & tabl & " SET year01 =[year01]+[year02], year02 =[year03], " & _
"year03 = [year04], year04 = 0;"
If myExecute("##390", sql) <> 0 Then Exit Function
'sql = "SELECT year01, year02, year03, year04 FROM " & tabl & ";"
'Set tbFirms = myOpenRecordSet("##390", sql, dbOpenDynaset)
'If tbFirms Is Nothing Then Exit Function
'While Not tbFirms.EOF
'    tbFirms.Edit
'    tbFirms!year01 = tbFirms!year01 + tbFirms!year02
'    tbFirms!year02 = tbFirms!year03
'    tbFirms!year03 = tbFirms!year04
''    tbFirms!year04 = 0
'    tbFirms.Update
'    tbFirms.MoveNext
'Wend
'tbFirms.Close
statisticReplace = True
EN1:
End Function
'$NOodbc$
Sub checkNextYear()
Dim I As Integer, S As Double

I = Format(Now, "yyyy")
If I <= lastYear Then Exit Sub

If MsgBox("Обнаружена необходимость перевода Статистики посешений на новый год. " & _
"Перевести?", vbDefaultButton2 Or vbYesNo, "Подтвердите!") = vbNo Then Exit Sub

wrkDefault.BeginTrans

If Not statisticReplace("GuideFirms") Then GoTo ER1
If Not statisticReplace("BayGuideFirms") Then GoTo ER1

If valueToSystemField("##389", I, "lastYear") Then
    wrkDefault.CommitTrans
    lastYear = I
    MsgBox "База переведена в новый(" & I & ") год!"
Else
ER1: wrkDefault.Rollback
    MsgBox "Программа не смогла перевести базу в новый(" & I & ") год! " & _
    "Перезапустите программу еще раз или свяжитесь с Администратором.", , "Error"
End If
End Sub

Function min(val1, val2)
If val2 < val1 Then
    min = val2
Else
    min = val1
End If
End Function

Function max(val1, val2)
If val2 > val1 Then
    max = val2
Else
    max = val1
End If
End Function

Sub msgOfZakaz(myErrCod As String, Optional msg As String = "", Optional mng = "")
    myErrCod = Mid$(myErrCod, 3)
    
    If msg = "" Then
        msg = "Нарушена целостность данных."
    End If
    msg = msg & "Обнаружен сбой базы (Err=" & myErrCod & ") в заказе № " & gNzak
    
    If Not IsNull(mng) Then
        msg = msg & "  " & CStr(mng)
    End If
    
    MsgBox msg & " Работа с этим заказом пока " & _
    "невозможна. Сообщите администратору!", , msg

End Sub

Sub msgOfEnd(myErrCod As String, Optional msg As String = "")
    myErrCod = Mid$(myErrCod, 3)
    MsgBox msg & " Сообщите администратору!", , "Ошибка " & myErrCod
    End
End Sub

' Возвращает issueId нового BusinessIssue если добавление действительно произошло
' иначе - возвращаем 0 - если нормальное завершение
' или -1 если произошла другая ошибка и было выдан диалог с описанием ошибки
' Если произошла ошибка на этапе добавления нового ишью, то выдаем наружу значение myExecute
Function myExecuteWithIssue(ByVal pSql As String, ByVal passErr As Integer, ByRef issueId As Integer) As Integer
On Error GoTo viewAtTheErrorNumber
    wrkDefault.BeginTrans
    myBase.Execute pSql ', dbFailOnError  ' выдавать Err если все или часть записей заблокировано
    wrkDefault.CommitTrans
    Exit Function

viewAtTheErrorNumber:
    On Error GoTo 0
    Dim strMsg As String, strSource As String, issued As Boolean
    Dim errLoop
       
    issued = False
    For Each errLoop In Errors
       With errLoop
          If .Number = passErr Then
             wrkDefault.CommitTrans
             issued = True
             strMsg = .Description
             strSource = .Source
             Exit For
          End If
       End With
    Next
    
    If issued Then
        sql = "select wi_file_new_issue (" & passErr & ", '" & strMsg & "')"
        byErrSqlGetValues "##", sql, issueId
        myExecuteWithIssue = 1
        wrkDefault.CommitTrans
    Else
        If errorCodAndMsg(cErr, passErr) Then
            myExecuteWithIssue = -2
        Else
            myExecuteWithIssue = 1
        End If
        wrkDefault.Rollback
    End If

End Function


' если passErr=-11111 или не указано то выдаются все сообщения
' если passErr=0  - подавляем сообщение "...WHERE..."
' если passErr<0  - подавляем все сообщения, кроме 3262 Or 3261
' если passErr>0  - подавляем сообщение только для ошибок с кодом= passErr
' в случае успешного вып-я возвращает myExecute=0 иначе выдает код ошибки
' системы myExecute >0; myExecute=-1 означает что записи не обнаружены
'$odbc15!$
Function myExecute(myErrCod, sql, Optional passErr As Integer = -11111) As Integer
myExecute = -1
On Error GoTo ERR1
RETR:
'wrkDefault.BeginTrans ' так рекомендуется обрамлять Execute но нельзя без wrkDefault.Rollback
myBase.Execute sql ', dbFailOnError  ' выдавать Err если все или часть записей заблокировано
'Debug.Print sql
If myBase.RecordsAffected < 1 Then
  If passErr > 0 Or passErr = -11111 Then _
    MsgBox "Нет записей, удовлетворяющих условию WHERE. Сообщите " & _
    "Администратору!", , "Error " & myErrCod & " в myExecute:"
  Exit Function
End If
myExecute = 0
Exit Function

ERR1:
wrkDefault.Rollback
cErr = Mid$(myErrCod, 3) ' - использовался наруже только в Prior
    
'MsgBox Error, , "Error " & cErr & "-" & Err & ":  "

If errorCodAndMsg(cErr, passErr) Then
    myExecute = -2
Else
    myExecute = 1
End If

End Function

'$odbc15!$
Function errorCodAndMsg(line As String, Optional passErr As Integer = 22222) As Boolean
Dim strError As String
Dim errLoop
   
   strError = "": errorCodAndMsg = True
   For Each errLoop In Errors
      With errLoop
         If .Number = passErr Then Exit Function
         
         strError = strError & _
            "******** Error: '" & .Number & "' *********" & vbCr
         strError = strError & _
            .Description & vbCr
         strError = strError & _
            "(Source:   " & .Source & ")" & vbCr & vbCr
      End With
'      MsgBox strError
   Next
errorCodAndMsg = False
MsgBox strError, , "sourceErr = " & line
End Function

'$odbc15$
Function myOpenRecordSet(myErrCod As String, sours As String, _
                passErr As Integer) As Recordset

On Error GoTo ErrorHandler

Set myOpenRecordSet = myBase.Connection.OpenRecordset(sours, dbOpenDynaset, dbExecDirect, dbPessimistic)

Exit Function

ErrorHandler:
    
If Not errorCodAndMsg(Mid$(myErrCod, 3), passErr) Then
    myBase.Close: End
End If

End Function


'$odbc08!$
Sub nextDayDetect() 'см также Orders.cmAdd_Click
Dim str As String ', intNum As Integer
Dim strNow As String, dNow As Date
Dim serverDate As String

strNow = Format(Now, "dd.mm.yyyy")
curDate = strNow 'без часов и минут

sql = "select convert(varchar(10), now(), 104)"
byErrSqlGetValues "##chksrvdate", sql, serverDate

If serverDate <> curDate Then
    fatalError "Время на компьютере очень сильно отличается от времени сервера." _
    & vbCr & "Дата на сервере: " & serverDate _
    & vbCr & "Работа программы будет завершена.", _
    "Если не получается или вы не знаете как, обратитесь к администратору"
End If

dNow = strNow
strNow = Right$(Format(curDate, "yymmdd"), 6)
 
befDays = 0

wrkDefault.BeginTrans 'lock01
sql = "update system set resursLock = resursLock" 'lock02
myBase.Execute (sql) 'lock03

Set tbSystem = myOpenRecordSet("##91", "System", dbOpenTable) ', dbOpenForwardOnly)
If tbSystem Is Nothing Then Exit Sub

'Строчки lock01-04 блокируют от конкурентного изменения в Sybase
'tbSystem.Edit '$odbs?$ необходимо, чтобы Другие не могли ни читать ни писать
'вплоть до Update

Dim doUpdateNum As Boolean
doUpdateNum = False

If tbSystem!resursLock = "nextDay" Then
   wrkDefault.Rollback
   MsgBox "Срочно сообщите Администратору! А пока можно работать с программой, " & _
    "но c ограниченной функциональностью.", , "Error при переводе Базы на новую дату!"

Else
    str = tbSystem!lastPrivatNum
    Dim valueorder As Numorder
    Set valueorder = newNumorder(str)
    If Not valueorder.IsEmpty Then
        tmpDate = valueorder.dat
        If tmpDate < dNow Then
            befDays = DateDiff("d", tmpDate, Now)
            doUpdateNum = True
            Set valueorder = New Numorder
        End If
     Else ' т.е. если lastPrivatNum не была еще инициализирована
        doUpdateNum = True
     End If
End If

If doUpdateNum Then
    sql = "UPDATE SYSTEM SET lastPrivatNum = " & valueorder.val
    'Debug.Print sql
    myBase.Execute (sql)
End If
     
If befDays <> 0 Then
    myBase.Execute ("UPDATE SYSTEM SET resursLock = 'nextDay'")
End If
wrkDefault.CommitTrans 'lock04
tbSystem.Close

End Sub

'$NOodbc$
Public Sub quickSort(varArray As Variant, _
 Optional lngLeft As Long = dhcMissing, Optional lngRight As Long = dhcMissing)
Dim I As Long, J As Long, varTestVal As Variant, lngMid As Long

    If lngLeft = dhcMissing Then lngLeft = LBound(varArray)
    If lngRight = dhcMissing Then lngRight = UBound(varArray)
   
    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
        varTestVal = varArray(lngMid)
        I = lngLeft
        J = lngRight
        Do
            Do While varArray(I) < varTestVal
                I = I + 1
            Loop
            Do While varArray(J) > varTestVal
                J = J - 1
            Loop
            If I <= J Then
                Call SwapElements(varArray, I, J)
                I = I + 1
                J = J - 1
            End If
        Loop Until I > J
        ' To optimize the sort, always sort the
        ' smallest segment first.
        If J <= lngMid Then
            Call quickSort(varArray, lngLeft, J)
            Call quickSort(varArray, I, lngRight)
        Else
            Call quickSort(varArray, I, lngRight)
            Call quickSort(varArray, lngLeft, J)
        End If
    End If
End Sub
'нужна для quickSort
Private Sub SwapElements(varItems As Variant, _
 lngItem1 As Long, lngItem2 As Long)
    Dim varTemp As Variant

    varTemp = varItems(lngItem2)
    varItems(lngItem2) = varItems(lngItem1)
    varItems(lngItem1) = varTemp
End Sub

Sub replaceDate(aTable As String, aField As String, checkDate As Date, pK As Integer)
Dim strDate As String
If Not IsNull(checkDate) Then
    If DateDiff("d", curDate, checkDate) < 0 Then
        strDate = Format(curDate, "yyyy-mm-dd 10:00:00")
        sql = "update " & aTable & " set " & aField & " = '" & strDate & "' where numOrder = " & pK
'        Debug.Print sql
        myBase.Execute (sql)
    End If
End If
End Sub

'$odbc10$

Sub SortCol(Grid As MSFlexGrid, col As Long, _
                        Optional ColisNum As String = "")
Static ascSort As Integer, dscSort As Integer
Grid.MousePointer = flexHourglass
    If ColisNum = "" Then
        ascSort = 5
        dscSort = 6
    ElseIf ColisNum = "date" Then
        Set sortGrid = Grid
        ascSort = 9
        dscSort = 9
    Else
        ascSort = 3
        dscSort = 4
    End If
    Grid.col = col
    Grid.ColSel = col
    trigger = Not trigger
    
    If trigger Then
        Grid.Sort = dscSort
    Else
        Grid.Sort = ascSort
    End If
Grid.MousePointer = flexDefault
End Sub


Function StatParamsLoad(row As Long, Optional redraw As Boolean = False)
Dim S As Double, log As String, str As String
Dim Grid As MSFlexGrid

Set Grid = Orders.Grid

If redraw Then
    Grid.col = orStatus
    Grid.row = row
    If tqOrders!equipStatusSync <> 0 Then
        Grid.CellForeColor = vbRed
    Else
        Grid.CellForeColor = vbBlack
    End If
End If

 log = Format(Now(), "dd.mm.yy hh:nn") & vbTab & Orders.cbM.Text & " " & gNzak ' именно vbTab
 str = status(tqOrders!StatusId): log = log & " " & str
 Orders.Grid.TextMatrix(row, orStatus) = str
 
 str = LoadDate(Orders.Grid, row, orDataVid, tqOrders!Outdatetime, "dd.mm.yy")
 If str <> "" Then log = log & " Out=" & str
 str = LoadNumeric(Orders.Grid, row, orVrVid, tqOrders!outTime)
 If str <> "" Then log = log & "_" & str
 
 str = LoadNumeric(Orders.Grid, row, orVrVip, tqOrders!Worktime, , "#0.0")
 log = log & " Вр.вып=" & str
 
 Orders.Grid.TextMatrix(row, orProblem) = tqOrders!problem
 
 str = LoadDate(Orders.Grid, row, orDataRS, tqOrders!DateRS, "dd.mm.yy")
 If str <> "" Then log = log & " РС=" & str
 
 gNzak = tqOrders!Numorder
 If IsNull(tqOrders!DateTimeMO) Then
    Orders.Grid.TextMatrix(row, orMOData) = ""
    Orders.Grid.TextMatrix(row, orMOVrVid) = ""
    str = ""
 Else
    str = LoadDate(Orders.Grid, row, orMOVrVid, tqOrders!DateTimeMO, "hh")
    If str <> "" Then
        str = LoadDate(Orders.Grid, row, orMOData, tqOrders!DateTimeMO, "dd.mm.yy") & "_" & str
    Else
        str = LoadDate(Orders.Grid, row, orMOData, tqOrders!DateTimeMO, "dd.mm.yy")
    End If
 End If
 
 If IsNull(tqOrders!StatM) Then
    Orders.Grid.TextMatrix(row, orM) = ""
 Else
    Orders.Grid.TextMatrix(row, orM) = tqOrders!StatM
    log = log & " Mк(" & tqOrders!StatM & "):" & str ' Дата выд
 End If
 If IsNull(tqOrders!StatO) Then
    Orders.Grid.TextMatrix(row, orO) = ""
    Orders.Grid.TextMatrix(row, orOVrVip) = ""
 Else
    Orders.Grid.TextMatrix(row, orO) = tqOrders!StatO
    If tqOrders!StatO = "в работе" Or tqOrders!StatO = "готов" Then
        If IsNull(tqOrders!DateTimeMO) Then
            msgOfZakaz "##313", "Отсутствует 'Дата MO'."
            str = " !Нет Даты MO! "
        End If
        log = log & " Oб(" & tqOrders!StatO & "):" & str ' Дата выд
        If IsNull(tqOrders!workTimeMO) Then
            msgOfZakaz "##314", "Отсутствует 'Время выполнения MO'."
        Else
            Orders.Grid.TextMatrix(row, orOVrVip) = tqOrders!workTimeMO
            str = LoadNumeric(Orders.Grid, row, orOVrVip, tqOrders!workTimeMO)
            log = log & "=" & str
        End If
    End If
 End If
StatParamsLoad = log
End Function

Sub rowViem(numRow As Long, Grid As MSFlexGrid)
Dim I As Long

I = Grid.Height \ Grid.RowHeight(1) - 1 ' столько умещается строк
I = numRow - I \ 2 ' в центр
If I < 1 Then I = 1
Grid.TopRow = I
End Sub

'эта ф-я д.заменить и startDay() и getNextDay() и getPrevDay()
' возвращает смещение до треб. дня
Function getWorkDay(offsDay As Integer, Optional baseDate As String = "") As Integer
Dim I As Integer, J As Integer, step  As Integer
getWorkDay = -1
If baseDate = "" Then
    tmpDate = curDate
Else
    If Not IsDate(baseDate) Then Exit Function
    tmpDate = baseDate
End If

step = 1
If offsDay < 0 Then step = -1

J = 0: I = 0
While step * J < step * offsDay '
    I = I + step
    day = Weekday(DateAdd("d", I, tmpDate))
    If Not (day = vbSunday Or day = vbSaturday) Then J = J + step
Wend
getWorkDay = I
tmpDate = DateAdd("d", I, tmpDate)

End Function

Function startDays() As Integer
Dim I As Integer, J  As Integer, K As Integer
ReDim Preserve stDays(befDays + 1)

For K = 0 To befDays '    *********************************************

J = 0
I = 1
While J < 3 '         задание смещения зеленого коридора (3-й день)

    day = Weekday(DateAdd("d", K + I - befDays, curDate))
'    day = Weekday(CurDate - befDays + K + I)
    If Not (day = vbSunday Or day = vbSaturday) Then J = J + 1
    I = I + 1
Wend
stDays(K) = I + K ' "+k" т.к. пока нумерация начинается befDays дней назад

Next K          '       ***********************************************
dayMassLenght (stDays(befDays) + 1)
startDays = stDays(befDays) - befDays ' для Сегодня, которое под №1
End Function

Sub statistic(Optional year As String = "")
Dim nRow As Long, nCol As Long, str As String, I As Integer, J As Integer
Dim iMonth As Integer, iYear As Integer, iCount As Integer, strWhere As String
Dim nMonth As Integer, nYear As Integer, mCount As Integer, lastCol As Integer
Dim wtSum As Double, paidSum As Double, orderSum As Double, visits As Integer, visitSum As Integer
Dim year01 As Integer, year02 As Integer, year03 As Integer, year04 As Integer
Dim errCurYear As Integer, errBefYear As Integer, whereByTemaAndType As String


errCurYear = 0:   errBefYear = 0

whereByTemaAndType = ""
If year = "" Then
 str = Reports.tbStartDate.Text
 Report.laHeader.Caption = "Статистика посещений фирм за период с " & str & _
                " по " & Reports.tbEndDate.Text
 nMonth = Left$(str, 2)
 nYear = Right$(str, 4)
 mCount = DateDiff("m", str, Reports.tbEndDate.Text) + 1

 str = "|<Название фирмы|^М |Kатегория|Скидки"
 iCount = mCount
 lastCol = 5 ' в 2х местах
 iMonth = nMonth
 Do
    If iMonth = 13 Then iMonth = 1
    str = str & "|" & Format(iMonth, "00")
    iMonth = iMonth + 1
    lastCol = lastCol + 1
    iCount = iCount - 1
 Loop While iCount > 0
 str = str & "|Итого|Вр.вып|Заказано|Оплачено"
 Report.Grid.FormatString = str
 Report.Grid.ColWidth(0) = 0
 Report.Grid.ColWidth(1) = 1875
 Report.Grid.ColWidth(3) = 375
 
 Report.nCols = lastCol + 2
 If Report.Regim = "KK" Then
    strWhere = "WHERE (((GuideFirms.Kategor)='К'));"
    Report.Grid.ColWidth(4) = 0
 ElseIf Report.Regim = "RA" Then
    strWhere = "WHERE (((GuideFirms.Kategor)='П' Or (GuideFirms.Kategor)='Д'));"
    Report.Grid.ColWidth(4) = 375
 Else
    Exit Sub
 End If
 
 If Reports.lbType.Text <> "все" Then
    lbToOrSqlWhere Reports.lbType, orType
    whereByTemaAndType = "(" & orSqlWhere(orType) & ") AND "
    Report.laHeader.Caption = Report.laHeader.Caption & _
    "  (заказы категории '" & Reports.lbType.Text & "'"
    If Reports.lbTema.Enabled Then
        lbToOrSqlWhere Reports.lbTema, orTema, "byId"
        str = ""
        For I = 0 To Reports.lbTema.ListCount - 1
            If Reports.lbTema.Selected(I) Then
                str = str & " " & Reports.lbTema.List(I)
            End If
        Next I
        If str <> "" Then
            whereByTemaAndType = whereByTemaAndType & "(" & _
                                        orSqlWhere(orTema) & ") AND "
            Report.laHeader.Caption = Report.laHeader.Caption & "  по теме:" & str
        End If
    End If
        Report.laHeader.Caption = Report.laHeader.Caption & ")"
 End If
 
 nRow = 1
 'sql = "SELECT GuideFirms.FirmId, GuideFirms.Name, GuideFirms.Kategor, " & _
 "GuideFirms.year01, GuideFirms.year02, GuideFirms.year03, GuideFirms.year04, " & _
 "GuideFirms.Sale, GuideManag.Manag FROM GuideFirms LEFT JOIN GuideManag " & _
 "ON GuideFirms.ManagId = GuideManag.ManagId " & strWhere
Else 'пересчет статистики для справ-ка фирм
 nMonth = 1
 nYear = lastYear - 3 '$$3
 mCount = DateDiff("m", "01.01." & nYear, curDate) + 1
 strWhere = ""
 'sql = "SELECT FirmId, Name, Kategor, " & _
 "year01, year02, year03, year04, " & _
 "Sale, ManagId FROM GuideFirms"
End If



 sql = "SELECT FirmId, Name, Kategor, " & _
 "year01, year02, year03, year04, " & _
 "Sale, ManagID FROM GuideFirms " & strWhere
'MsgBox sql
Set tbFirms = myOpenRecordSet("##68", sql, dbOpenDynaset) 'ForwardOnly)
If tbFirms Is Nothing Then Exit Sub
If tbFirms.BOF Then GoTo EN1:
tbFirms.MoveFirst
While Not tbFirms.EOF '                         *******************

 iMonth = nMonth
 iYear = nYear
 iCount = mCount
 visitSum = 0
 wtSum = 0
 paidSum = 0
 orderSum = 0
 bilo = False
 nCol = 5 ' в 2х местах
 year01 = 0: year02 = 0: year03 = 0: year04 = 0
 Do
'    str = Format(iMonth, "00") & "." & iYear
    str = iYear & "-" & Format(iMonth, "00")
    
    
    
    sql = "SELECT Orders.numOrder, Orders.workTime, Orders.paid, Orders.ordered From Orders " & _
    "WHERE (" & whereByTemaAndType & " ((Orders.inDate) Like '" & str & "-%') AND " & _
    "(Not ((Orders.StatusId)=0 Or (Orders.StatusId)=7)) AND " & _
    "((Orders.FirmId)=" & tbFirms!firmId & ") AND ((Orders.workTime) Is Not Null));"
'Debug.Print "1:" & sql
'If tbFirms!firmId > 0 Then MsgBox sql
    Set tbOrders = myOpenRecordSet("##69", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    visits = 0:
    If Not tbOrders.BOF Then
'      tbOrders.MoveFirst
      While Not tbOrders.EOF '$$3
            str = tbOrders!Numorder
          If year <> "" Then
            If iYear = lastYear - 3 Then
                year01 = year01 + 1 'не исп-ся
            ElseIf iYear = lastYear - 2 Then
                year02 = year02 + 1
            ElseIf iYear = lastYear - 1 Then
                year03 = year03 + 1
            ElseIf iYear = lastYear Then
                year04 = year04 + 1
            End If
          End If
          visits = visits + 1
          wtSum = wtSum + tbOrders!Worktime
          If Not IsNull(tbOrders!paid) Then _
                paidSum = paidSum + tbOrders!paid
          If Not IsNull(tbOrders!ordered) Then _
                orderSum = orderSum + tbOrders!ordered
          tbOrders.MoveNext
      Wend
    End If
    If visits > 0 And year = "" Then
        If Not bilo Then
            Report.Grid.TextMatrix(nRow, 1) = tbFirms!name
            If Not IsNull(tbFirms!ManagId) Then _
                    Report.Grid.TextMatrix(nRow, 2) = Manag(tbFirms!ManagId)
            Report.Grid.TextMatrix(nRow, 3) = tbFirms!Kategor
            If Not IsNull(tbFirms!Sale) Then _
                    Report.Grid.TextMatrix(nRow, 4) = tbFirms!Sale
            bilo = True
        End If
        Report.Grid.TextMatrix(nRow, nCol) = visits
        visitSum = visitSum + visits
    End If
    
    If iMonth = 12 Then
        iMonth = 1
        iYear = iYear + 1
    Else
        iMonth = iMonth + 1
    End If
    
    nCol = nCol + 1
    iCount = iCount - 1
 Loop While iCount > 0

 If bilo And year = "" Then
    Report.Grid.TextMatrix(nRow, lastCol) = visitSum
    Report.Grid.TextMatrix(nRow, lastCol + 1) = Round(wtSum, 1)
    Report.Grid.TextMatrix(nRow, lastCol + 2) = Round(orderSum, 1)
    Report.Grid.TextMatrix(nRow, lastCol + 3) = Round(paidSum, 1)
    Report.Grid.AddItem ""
    nRow = nRow + 1
 End If
NXT:
 If year <> "" Then 'пересчет статистики
    tbFirms.Edit
    I = getLockYear 'не пересчитываем года, кот учавствовали в отсечении базы
'первую тоже не пересчитываем, т.к. незивестно насколько там ранние года
'        If tbFirms!year01 <> year01 Then errBefYear = errBefYear + 1
'        tbFirms!year01 = year01
    If lastYear - 2 > I Then
        If tbFirms!year02 <> year02 Then errBefYear = errBefYear + 1
        tbFirms!year02 = year02
    End If
    If lastYear - 1 > I Then
        If tbFirms!year03 <> year03 Then errBefYear = errBefYear + 1
        tbFirms!year03 = year03
    End If
    If lastYear > I Then
        If tbFirms!year04 <> year04 Then errCurYear = errCurYear + 1
        tbFirms!year04 = year04
    End If
    tbFirms.update
 End If
 tbFirms.MoveNext
Wend '*******************
EN1:
tbFirms.Close
If year = "" Then
  If nRow > 1 Then Report.Grid.removeItem (nRow)
  Report.laCount.Caption = nRow - 1
Else
'  If errBefYear > 0 Then !!!не стирать
'     MsgBox "В прошлых годах обнаружено " & errBefYear & " фирм с неверно " & _
'     "подсчитанным количеством посещений.  Все ошибки устранены.", , "Обнаружены ошибки"
'  End If
'  If errCurYear > 0 Then
'     MsgBox "В текущем году обнаружено " & errCurYear & " фирм с неверно " & _
'     "подсчитанным количеством посещений.  Все ошибки устранены.", , "Обнаружены ошибки"
'  End If
End If
End Sub

Function isNumericTbox(tBox As TextBox, Optional minVal, Optional maxVal) As Boolean

If checkNumeric(tBox.Text, minVal, maxVal) Then
    isNumericTbox = True
Else
    isNumericTbox = False
    tBox.SetFocus
    tBox.SelStart = 0
    tBox.SelLength = Len(tBox.Text)
End If
End Function

Function checkNumeric(val As String, Optional minVal, Optional maxVal) As Boolean

checkNumeric = True
If IsNumeric(val) And InStr(val, " ") = 0 Then
    tmpSng = val
    If Not IsMissing(maxVal) Then
        If (minVal > tmpSng Or tmpSng > maxVal) Then
            MsgBox "значение должно быть в диапазоне от " & minVal & _
            "  до " & maxVal, , "Error"
            checkNumeric = False
        End If
    ElseIf Not IsMissing(minVal) Then
        If minVal > tmpSng Then
            MsgBox "значение должно быть больше " & minVal
            checkNumeric = False
        End If
    End If
Else
    MsgBox "Недопустимое значение", , "Error"
    checkNumeric = False
End If
End Function

'в случеу true также возвращает дату в tmpDate
Function isDateTbox(tBox As TextBox, Optional fryDays As String = "", Optional doEmptyCheck As Boolean = True) As Boolean
Dim str As String

isDateTbox = False
str = tBox.Text
If str <> "" Then
    str = "20" & Right$(str, 2) & "-" & Mid$(str, 4, 2) & "-" & Left$(str, 2)
    If IsDate(str) Then
        isDateTbox = True
        tmpDate = str
        If fryDays <> "" Then
            day = Weekday(tmpDate)
            If day = vbSunday Or day = vbSaturday Then
                If MsgBox(str & " - выходной день. Продолжить?", vbYesNo, "Предупреждение!") <> vbYes Then
                    isDateTbox = False
                End If
            End If
        End If
    Else
        MsgBox "Неверный формат даты или дня с такой датой не существует ", , "Ошибка"
    End If
Else
    If doEmptyCheck Then
        MsgBox "Заполните поле Даты!", , "Ошибка"
    End If
End If
If Not isDateTbox Then
    tBox.SelStart = 0
    tBox.SelLength = Len(tBox.Text)
    On Error Resume Next
    tBox.SetFocus
End If
End Function


Function valueToSystemField(myErr As String, val As Variant, field As String) As Boolean

valueToSystemField = False
'sql = "select * from System"
'Set tbSystem = myOpenRecordSet(myErr, sql, dbOpenForwardOnly)
'If tbSystem Is Nothing Then myBase.Close: End
'Debug.Print val
If val = "" Then val = "''"
myBase.Execute ("UPDATE SYSTEM SET " & field & " = " & val)

'tbSystem.Edit
'tbSystem.Fields(field) = val
'tbSystem.Update
'tbSystem.Close
valueToSystemField = True
End Function

'не записыват неуникальное значение, для полей, где такие
'значения запрещены. А  генерит при этом error?
Function ValueToTableField(myErrCod As String, ByVal value As String, ByVal table As String, _
ByVal field As String, Optional by As String = "", Optional Numorder As Variant) As Integer
Dim sql As String, byStr As String  ', numOrd As String


ValueToTableField = False
'If value = "" Then value = Chr(34) & Chr(34)

If value = "" Then value = "''"
If by = "" Then
    Dim nzak As String
    If IsMissing(Numorder) Then
        nzak = gNzak
    Else
        nzak = Numorder
    End If
        
    byStr = ".numOrder = " & nzak
ElseIf by = "byFirmId" Then
    byStr = ".FirmId = " & gFirmId
ElseIf by = "byKlassId" Then
    byStr = ".klassId = " & gKlassId
ElseIf by = "byNomNom" Then
    byStr = ".nomNom = " & "'" & gNomNom & "'"
ElseIf by = "bySeriaId" Then
    byStr = ".seriaId = " & gSeriaId
ElseIf by = "byProductId" Then
    byStr = ".prId = " & gProductId
ElseIf by = "byWerkId" Then
    byStr = ".numOrder = " & gNzak
ElseIf by = "byNumDoc" Then
    sql = "UPDATE " & table & " SET " & table & "." & field & "=" & value _
        & " WHERE " & table & ".numDoc =" & numDoc & " AND " & table & _
        ".numExt =" & numExt
    GoTo AA
Else
    Exit Function
End If
sql = "UPDATE " & table & " SET " & table & "." & field & _
" = " & value & " WHERE " & table & byStr
AA:
'MsgBox "sql = " & sql

If Left$(myErrCod, 1) = "W" Then
    myErrCod = Mid$(myErrCod, 2)
    ValueToTableField = myExecute(myErrCod, sql, 0) 'не сообщать если не WHERE
ElseIf Left$(myErrCod, 1) = "L" Then
    ' Ожидаем, что выполненную операцию нужно записать в серверный журнал (lBusinessIssues).
    ' пример формата строки ошибки - "L-17002"
    myErrCod = Mid$(myErrCod, 2)
    Dim issueId As Integer
    ValueToTableField = myExecuteWithIssue(sql, CInt(myErrCod), issueId)
    ValueToTableField = issueId
Else
    ValueToTableField = myExecute(myErrCod, sql)
End If
End Function

Sub unLockBase()
valueToSystemField "##148", "", "resursLock"
End Sub

Sub getIdFromGrid5Row(Frm As Form, Optional p_row As Long = -1)
Dim str As String, I As Integer
Dim v_row As Long

If IsMissing(p_row) Or p_row = -1 Then
    v_row = Frm.mousRow5
Else
    v_row = p_row
End If

If Frm.Grid5.TextMatrix(v_row, prType) = "изделие" Then
    str = Frm.Grid5.TextMatrix(v_row, prName) '
    I = InStr(str, "/")
    prExt = 0: If I > 1 Then prExt = Left$(str, I - 1)   'номер поставки
    gProductId = Frm.Grid5.TextMatrix(v_row, prId)
Else
    gNomNom = Frm.Grid5.TextMatrix(v_row, prId)
End If
End Sub

Function getNevip(day As Integer, equipId As Integer)
sql = "SELECT Sum(oe.workTime * oc.Nevip) AS wSum " & _
"FROM OrdersInCeh oc " & _
"JOIN OrdersEquip oe ON oe.numOrder = oc.numOrder " & _
"WHERE DateDiff(day,'" & Format(curDate, "yyyy-mm-dd") & "',oe.outDateTime) =" & day - 1 _
& " AND oe.equipId =" & equipId
'MsgBox sql
getNevip = 0
byErrSqlGetValues "W##382", sql, getNevip
End Function

Sub addDays(outDay As Integer)
Dim J As Integer
        If maxDay < outDay Then
            dayMassLenght outDay + 1 'если дольше , корректируем размерности
            For J = maxDay + 1 To outDay 'новые дни
                delta(J) = 0
            Next J
            maxDay = outDay
        End If
End Sub

Function getLockYear() As Integer
getLockYear = Format(begDate, "yyyy")
If Format(begDate, "dd.mm") = "01.01" Then _
    getLockYear = getLockYear - 1 'считаем, что этот год не учавствовал в отсечении базы
End Function

Function getYearField(checkDate As Date) As String '$$3
Dim I As Integer, lockYear As Integer

lockYear = getLockYear
I = Format(checkDate, "yyyy")
'If I <= lockYear Then
'    getYearField = "lock" 'этот год учавствовал в отсечении базы
'    Exit Function
'End If
I = I - lastYear + 4 'номер колонки
If I < 1 Then     'если это не последние 3 года, то в кучу
    getYearField = "year01"
Else
    getYearField = "year" & Format(I, "00")
End If
End Function


Sub visits(oper As String, Optional firm As String = "") '$$3
Dim str As String, I As Integer, statId As Integer

sql = "SELECT Orders.inDate, Orders.StatusId , Orders.FirmId From Orders " & _
"WHERE Orders.numOrder = " & gNzak
'MsgBox sql
If Not byErrSqlGetValues("##88", sql, tmpDate, statId, I) Then GoTo ER1

If I = 0 Then Exit Sub
If firm <> "" And (statId = 0 Or statId = 7) Then Exit Sub ' если меняем фирму

str = getYearField(tmpDate)

'If str = "lock" Then Exit Sub ' если год участвовал в отсечении базы , то его не пересчитываем

sql = "UPDATE GuideFirms SET GuideFirms." & str & " = [GuideFirms].[" & _
str & "] " & oper & " 1  WHERE GuideFirms.FirmId =" & I
'Debug.Print sql
I = myExecute("##87", sql, -143)

'If I <> 3061 And I <> 0 Then '3061 - колонки этого года уже(или еще) нет в базе
If I = -2 Then '3061 - колонки этого года уже(или еще) нет в базе
ER1:    MsgBox "Ошибка коррекции посещения фирм. Сообщите администратору!", , "Error-87"
End If
End Sub


Sub zagruzFromCeh(equipId As Integer, Optional passZakazNom As String = "")
Dim outDay As Integer, J As Integer, passSql As String, str As String
Dim tbCeh As Recordset


If IsNumeric(passZakazNom) Then
    passSql = " AND oe.numOrder <> " & passZakazNom
Else
    passSql = ""
End If

'    "OrdersInCeh.numOrder, OrdersInCeh.VrVipParts, OrdersInCeh.rowLock "
sql = "SELECT oe.outDateTime, o.StatusId, o.numOrder" _
    & " FROM Orders o " _
    & " JOIN OrdersEquip oe ON oe.numOrder = o.numOrder" _
    & " JOIN OrdersInCeh oc ON oc.numOrder = o.numOrder" _
    & " WHERE oe.EquipId = " & equipId & " AND (isnull(worktime, 0) > 0 OR isnull(worktimeMO, 0) > 0 ) " & passSql

'Debug.Print sql

Set tbCeh = myOpenRecordSet("##14", sql, dbOpenForwardOnly)
If tbCeh Is Nothing Then Exit Sub

'1:MaxDay = 0
If Not tbCeh.BOF Then
    While Not tbCeh.EOF
        isLive = False ' неживой заказ
        If tbCeh!StatusId = 1 Then
            isLive = True
        End If
        outDay = DateDiff("d", curDate, tbCeh!Outdatetime) + 1
        If outDay < 1 Then outDay = 1
                
        addDays outDay '1:добавляем дни,  т.к. Дата  Выд тек.заказа может
                       '  оказаться  дальше  чем stDay и rMaxDay
        tbCeh.MoveNext
    Wend
End If
tbCeh.Close
End Sub


Function beNaklads(Optional reg As String = "") As Boolean
beNaklads = True
Dim S As Double
'отпущено
sql = "SELECT Sum(sDMC.quant) AS Sum_quant From sDMC " & _
"WHERE (((sDMC.numExt)< 254) AND ((sDMC.numDoc)=" & numDoc & "));"
If Not byErrSqlGetValues("##140", sql, S) Then Exit Function
If S > 0.005 Then ' что-то отпущено
    If reg = "" Then
        MsgBox "По этому заказу выписывались накладные, поэтому изменять " & _
        "предметы нельзя. Если изменения все-же требуются, то прежде надо " & _
        "удалить все накладные к заказу.", , "Редактирование запрещено!"
    End If
Else
    sql = "SELECT Sum(curQuant) AS Sum_curQuant " & _
    "From sDMCrez WHERE (((numDoc)=" & numDoc & "));"
    If Not byErrSqlGetValues("##367", sql, S) Then Exit Function
    If S > 0.005 Then ' что-то отпущено
        If reg = "" Then
            MsgBox "По этому заказу в Цеху уже выписана накладная (заполнена " & _
            "колонка 'Кол-во') и возможно она уже распечатана. Прежде чем " & _
            "менять предметы заказа, Цех должен уничтожить распечатку, а затем " & _
            "обнулить колонку 'Кол-во'.", , "Редактирование запрещено!"
        End If
    Else
        beNaklads = False
    End If
End If

End Function

'$odbc08$ Function docLock не исп-ся
'--------------------------------------------------------------------------

Sub getNakladnieList(Optional from As String = "") '
Dim I As Integer, str As String, l As Long

'pusto=-1 если хотя бы одного предмета нет ни в одной накладной (иначе pusto=0)
'нужно, т.к. при этом delta=Null а не quantity
If from = "Buh" Then str = "3" Else str = "2" 'для буха только те prior_заказы, где проставлены кол-ва в Цеху

sql = "SELECT numDoc, Max(quantity - IsNull( Sum_quant, 0)) AS delta, " _
& " Min(IsNull(Sum_quant,0)) AS pusto  " _
& " From wCloseNomenk" & str _
& " GROUP BY numDoc ORDER BY numDoc;"

'Debug.Print sql
Set tbDMC = myOpenRecordSet("##142", sql, dbOpenDynaset)
If tbDMC Is Nothing Then Exit Sub

I = 0
ReDim tmpL(0)
While Not tbDMC.EOF
 
  If tbDMC!pusto = -1 Then GoTo AA
  If tbDMC!delta < 0.005 Then ' все предметы закрыты
        gNzak = -tbDMC!numDoc
  Else
AA:     gNzak = tbDMC!numDoc
    
    If from = "werk" Then ' для цеха проверяем еще и этапность
      sql = "SELECT numOrder From xEtapByNomenk Where numOrder = " _
      & gNzak _
      & " UNION ALL SELECT numOrder From xEtapByIzdelia " & _
      "WHERE numOrder = " & gNzak
      If Not byErrSqlGetValues("W##352", sql, l) Then GoTo NXT
      
      If l > 0 Then 'для убыстрения делаем только для этапных
          If predmetiIsClose("etap") Then GoTo NXT 'закрытые по этапу пропускаем
      End If
    End If
    
  End If
    
    I = I + 1
    ReDim Preserve tmpL(I)
    tmpL(I) = gNzak
NXT:
    tbDMC.MoveNext
Wend
tbDMC.Close
End Sub

Function getNextNumExt() As Integer
Dim v As Variant

getNextNumExt = 0
sql = "SELECT Max(sDocs.numExt) AS Max_numExt From sDocs " & _
"WHERE (((sDocs.numDoc)=" & numDoc & " AND (sDocs.numExt) < 254));"

If Not byErrSqlGetValues("##128", sql, v) Then Exit Function
If IsNumeric(v) Then
    getNextNumExt = v + 1
Else
    getNextNumExt = 1
End If

End Function


'reg=""  - проверка полного списания предметов
'reg = "prev" - проверка, что списано ровно по пред.этапу, не больше
'иначе - проверка, что списано по этапу, не менее
Function predmetiIsClose(Optional reg As String = "") As Boolean
Dim I As Integer, S As Double

#If onErrorOtlad Then
    On Error GoTo errMsg
    GoTo START
errMsg:
    MsgBox Error, , "Ошибка  " & Err & " в п\п predmetiIsClose" '
    End
START:
#End If

predmetiIsClose = False

If Not sProducts.zakazNomenkToNNQQ() Then Exit Function
For I = 1 To UBound(NN)
    sql = "SELECT Sum(quant) AS Sum_quant From sDMC " & _
    "WHERE (((sDMC.numDoc)=" & gNzak & ") AND ((nomNom)='" & NN(I) & "'));"
    If Not byErrSqlGetValues("##164", sql, S) Then Exit Function
    If reg = "prev" Then
        If Abs(QQ3(I) - S) > 0.005 Then Exit Function
    ElseIf reg = "" Or QQ2(0) = 0 Then 'вызов не из цеха или для неэтапного заказа
        If QQ(I) - S > 0.005 Then Exit Function
    Else
        If QQ2(I) - S > 0.005 Then Exit Function
    End If
Next I
predmetiIsClose = True
End Function


Function PrihodRashod(reg As String, skladId As Integer) As Double
Dim qWhere As String, S As Double

PrihodRashod = 0

If reg = "+" Then
    If skladId = 0 Then
        qWhere = ") AND ((sDocs.destId) < -1000)"
    ElseIf skladId = 2 Then
        qWhere = ") AND ((sDocs.destId) = -1001 Or (sDocs.destId) = -1002)"
    Else
        qWhere = ") AND ((sDocs.destId) =" & skladId & ")"
    End If
ElseIf reg = "-" Then
    If skladId = 0 Then
        qWhere = ") AND ((sDocs.sourId) < -1000)"
    ElseIf skladId = 2 Then
        qWhere = ") AND ((sDocs.sourId) = -1001 Or (sDocs.sourId) = -1002)"
    Else
        qWhere = ") AND ((sDocs.sourId) =" & skladId & ")"
    End If
End If
sql = "SELECT Sum(sDMC.quant) AS Sum_quantity FROM sDocs INNER JOIN " & _
"sDMC ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc) " & _
"WHERE (((sDMC.nomNom) = '" & gNomNom & "' " & qWhere & ");"
'Debug.Print sql
byErrSqlGetValues "##157", sql, PrihodRashod

End Function
'$odbc15$
Function ostatCorr(delta As Double) As Boolean
Dim sId As Integer, dId As Integer

ostatCorr = False

sql = "SELECT sDocs.sourId, sDocs.destId, sDocs.numDoc, sDocs.numExt " & _
"From sDocs " & _
"WHERE (((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & "));"
If Not byErrSqlGetValues("##180", sql, sId, dId) Then Exit Function

If sId < -1000 And dId < -1000 Then ' для межскладских не корректируем
        ostatCorr = True
Else
    'корректируем остатки
'    ostatCorr = False
'    Set tbNomenk = myOpenRecordSet("##163", "select * from sGuideNomenk", dbOpenForwardOnly)
'    If tbNomenk Is Nothing Then Exit Function
'    tbNomenk.index = "PrimaryKey"
'    tbNomenk.Seek "=", gNomNom
'    If Not tbNomenk.NoMatch Then
'        tbNomenk.Edit
'        tbNomenk!nowOstatki = Round(tbNomenk!nowOstatki - delta, 2)
'        tbNomenk.Update
'        ostatCorr = True
'    End If
'    tbNomenk.Close
    
    sql = "UPDATE sGuideNomenk SET nowOstatki = [nowOstatki]-" & delta & _
    " WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
    If myExecute("##163", sql) <> 0 Then Exit Function
    ostatCorr = True
End If
End Function

'исп-ся при формировании предметов, а также отвечает за часть Надлежит
'отпустить в Otgruz.frm
Sub loadPredmeti(Frm As Form, orderRate As Double, Optional reg As String = "", Optional needToRefresh As Boolean = False)
Dim I As Integer

Screen.MousePointer = flexHourglass
Frm.Grid5.Visible = False
Frm.quantity5 = 0
I = 0: If reg = "fromOtgruz" Then I = 1

clearGrid Frm.Grid5, 1 + I

'******** изделия ************************************************
sql = "SELECT sGuideProducts.prName, sGuideProducts.prDescript, " & _
"xPredmetyByIzdelia.*, xEtapByIzdelia.eQuant, xEtapByIzdelia.prevQuant " & _
"FROM (sGuideProducts INNER JOIN xPredmetyByIzdelia ON sGuideProducts.prId = " & _
"xPredmetyByIzdelia.prId) LEFT JOIN xEtapByIzdelia ON (xPredmetyByIzdelia." & _
"prExt = xEtapByIzdelia.prExt ) AND (xPredmetyByIzdelia.prId = " & _
"xEtapByIzdelia.prId) AND (xPredmetyByIzdelia.numOrder = xEtapByIzdelia.numOrder)" & _
"WHERE (((xPredmetyByIzdelia.numOrder)= " & gNzak & "));"
'Debug.Print sql

Set tbNomenk = myOpenRecordSet("##183", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then Exit Sub
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    Frm.quantity5 = Frm.quantity5 + 1
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prId) = tbNomenk!prId
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prType) = "изделие"
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prName) = getStrPrEx(tbNomenk!prName, tbNomenk!prExt)
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prDescript) = tbNomenk!prDescript
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prEdizm) = "шт."
    If Not IsNull(tbNomenk!cenaEd) Then
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prCenaEd) = Round(rated(tbNomenk!cenaEd, orderRate), 2)
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prSumm) = _
                                Round(rated(tbNomenk!cenaEd * tbNomenk!quant, orderRate), 2)
    End If
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prQuant) = Round(tbNomenk!quant, 2)
' все изменения проделать и для ном-ры (см. ниже)
    If reg = "fromOtgruz" Then
        Otgruz.getOtgrugeno Frm.quantity5 + I
    ElseIf Not IsNull(tbNomenk!eQuant) Then
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prEtap) = tbNomenk!eQuant
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prEQuant) = _
                            Round(tbNomenk!eQuant - tbNomenk!prevQuant, 2)
    End If
    
    Frm.Grid5.AddItem ""
    tbNomenk.MoveNext
  Wend
End If
tbNomenk.Close

'****** номенклатура ********************************************
sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, " & _
"sGuideNomenk.Size, sGuideNomenk.nomNom, sGuideNomenk.cod, " & _
"sGuideNomenk.ed_Izmer, xPredmetyByNomenk.quant, xPredmetyByNomenk.cenaEd, " & _
"xEtapByNomenk.eQuant, xEtapByNomenk.prevQuant " & _
"FROM (sGuideNomenk INNER JOIN xPredmetyByNomenk ON sGuideNomenk.nomNom = " & _
"xPredmetyByNomenk.nomNom) LEFT JOIN xEtapByNomenk ON (xPredmetyByNomenk." & _
"nomNom = xEtapByNomenk.nomNom) AND (xPredmetyByNomenk.numOrder = xEtapByNomenk.numOrder) " & _
"WHERE (((xPredmetyByNomenk.numOrder)=" & gNzak & "));"

'Debug.Print sql
Set tbNomenk = myOpenRecordSet("##184", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then Exit Sub
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    Frm.quantity5 = Frm.quantity5 + 1
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prId) = tbNomenk!nomNom
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prType) = "номенклатура"
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prName) = tbNomenk!cod
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prDescript) = _
        tbNomenk!nomName & " " & tbNomenk!Size
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prEdizm) = tbNomenk!ed_Izmer
    If Not IsNull(tbNomenk!cenaEd) Then
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prCenaEd) = Round(rated(tbNomenk!cenaEd, orderRate), 2)
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prSumm) = _
                                Round(rated(tbNomenk!cenaEd * tbNomenk!quant, orderRate), 2)
    End If

    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prQuant) = Round(tbNomenk!quant, 2)

    If reg = "fromOtgruz" Then
        Otgruz.getOtgrugeno Frm.quantity5 + I, "byNomenk"
    ElseIf Not IsNull(tbNomenk!eQuant) Then
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prEtap) = Round(tbNomenk!eQuant, 2)
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prEQuant) = _
                            Round(tbNomenk!eQuant - tbNomenk!prevQuant, 2)
    End If
    
    
    Frm.Grid5.AddItem ""
    tbNomenk.MoveNext
  Wend
End If
tbNomenk.Close

If Frm.quantity5 > 0 Then
    Frm.Grid5.row = Frm.quantity5 + 1 + I
    Frm.Grid5.col = prQuant
    Frm.Grid5.Text = "Итого:"
    Frm.Grid5.col = prSumm
    Frm.Grid5.Text = Round(rated(sProducts.saveOrdered(orderRate, needToRefresh), orderRate), 2)
    Frm.Grid5.CellFontBold = True
    If reg = "fromOtgruz" Then
        Frm.Grid5.col = prOutSum
        Frm.Grid5.Text = Round(rated(Otgruz.saveShipped(False), orderRate), 2)
        Frm.Grid5.CellFontBold = True
        Frm.Grid5.col = prNowSum
        Frm.Grid5.Text = "0"
        Frm.Grid5.CellFontBold = True
    End If
End If
Frm.Grid5.Visible = True

Screen.MousePointer = flexDefault
End Sub

Function lockSklad(Optional back As String = "") As Boolean
Dim str As String
lockSklad = True: Exit Function '!!! отключаем чтобы проверить, нужно все это реально
lockSklad = False
RETR:
sql = "select * from System"
Set tbSystem = myOpenRecordSet("##94", sql, dbOpenForwardOnly)
If tbSystem Is Nothing Then myBase.Close: End
''''''LOCK''''''

'tbSystem.Edit
str = tbSystem!skladLock
If Not back = "" Then
    If str = Orders.cbM.Text Then
                'tbSystem!skladLock = ""
        myBase.Execute ("update System set skladLock = ''")
    End If
Else
    If str <> "" And str <> Orders.cbM.Text Then
        'tbSystem.Update:
        tbSystem.Close
        
        If MsgBox("Доступ к остаткам на складе временно занят менеджером '" & _
        str & "'. Повторите попытку или обратитесь к Администратору.", _
        vbRetryCancel, "Нет доступа !!!") = vbRetry Then
            GoTo RETR
        Else
            Exit Function
        End If
    End If
    'tbSystem!skladLock = Orders.cbM.Text
        myBase.Execute ("update System set skladLock = " & Orders.cbM.Text)
End If
'tbSystem.Update
tbSystem.Close
lockSklad = True
End Function
    
Function orderUpdateWithIssue(ByVal issueMarker As String, value As String, table As String, _
field As String, Optional by As String = "", Optional Numorder As Variant) As Integer
Dim nzak As String
Dim issueId As Variant
    If IsMissing(Numorder) Then
        nzak = gNzak
    Else
        nzak = Numorder
    End If
    orderUpdateWithIssue = ValueToTableField("##orderUpdateWithIssue", value, table, field, by, nzak)
    
    sql = "select wi_check_business_issue(' " & issueMarker & "')"
    byErrSqlGetValues "##check_issue", sql, issueId
    If Not IsNull(issueId) And issueId <> 0 Then
        orderUpdateWithIssue = CInt(issueId)
    End If
    
    If table = "Orders" Then
        refreshTimestamp (nzak)
    End If
End Function
    
    
Function orderUpdate(ByVal myErrCod As String, value As String, table As String, _
field As String, Optional by As String = "", Optional Numorder As Variant) As Integer
Dim nzak As String
    If IsMissing(Numorder) Then
        nzak = gNzak
    Else
        nzak = Numorder
    End If
    orderUpdate = ValueToTableField(myErrCod, value, table, field, by, nzak)
    If table = "Orders" Then
        refreshTimestamp (nzak)
    End If
End Function

Function refreshTimestamp(nzak As String)
    Dim orderTimestamp As Date
    Dim zakRow As Long
    
    sql = "SELECT O.lastModified From Orders o " _
        & " WHERE O.numOrder = " & nzak
    If Not byErrSqlGetValues("##174.2", sql, orderTimestamp) Then Exit Function
    
    zakRow = searchZakRow(Orders.Grid, nzak)

    Orders.Grid.TextMatrix(zakRow, orlastModified) = orderTimestamp
End Function

Function searchZakRow(ByRef Grid As MSFlexGrid, nzak As String) As Long
Dim irow As Long

    searchZakRow = -1
    For irow = 1 To Grid.Rows - 1
        If Grid.TextMatrix(irow, orNomZak) = nzak Then
            searchZakRow = irow
            Exit Function
        End If
    Next irow

End Function

Sub loadSeria(ByRef p_tv As TreeView)
Dim Key As String, pKey As String, K() As String, pK()  As String
Dim I As Integer, iErr As Integer
bilo = False
sql = "SELECT sGuideSeries.*  From sGuideSeries ORDER BY sGuideSeries.seriaId;"
Set tbSeries = myOpenRecordSet("##110", sql, dbOpenForwardOnly)
If tbSeries Is Nothing Then myBase.Close: End
If Not tbSeries.BOF Then
 p_tv.Nodes.Clear
 Set Node = p_tv.Nodes.Add(, , "k0", "Справочник по сериям")
 Node.Sorted = True
 
 ReDim K(0): ReDim pK(0): ReDim NN(0): iErr = 0
 While Not tbSeries.EOF
    If tbSeries!seriaId = 0 Then GoTo NXT1
    Key = "k" & tbSeries!seriaId
    pKey = "k" & tbSeries!parentSeriaId
    On Error GoTo ERR1 ' назначить второй проход
    Set Node = p_tv.Nodes.Add(pKey, tvwChild, Key, tbSeries!seriaName)
    On Error GoTo 0
    Node.Sorted = True
NXT1:
    tbSeries.MoveNext
 Wend
End If
tbSeries.Close

While bilo ' необходимы еще проходы
  bilo = False
  For I = 1 To UBound(K())
    If K(I) <> "" Then
        On Error GoTo ERR2 ' назначить еще проход
        Set Node = p_tv.Nodes.Add(pK(I), tvwChild, K(I), NN(I))
        On Error GoTo 0
        K(I) = ""
        Node.Sorted = True
    End If
NXT:
  Next I
Wend
p_tv.Nodes.item("k0").Expanded = True
Exit Sub
ERR1:
 iErr = iErr + 1: bilo = True
 ReDim Preserve K(iErr): ReDim Preserve pK(iErr): ReDim Preserve NN(iErr)
 K(iErr) = Key: pK(iErr) = pKey: NN(iErr) = tbSeries!seriaName
 Resume Next

ERR2: bilo = True: Resume NXT

End Sub

Public Function getManagById(ManagId As Variant) As String
Dim I As Integer
    getManagById = ""
    If Not IsNull(ManagId) Then
        Dim imanagId As String
        imanagId = CStr(ManagId)
        If imanagId <> "" Then
            For I = 0 To UBound(Managers)
                If Managers(I).Key = imanagId Then
                    getManagById = CStr(Managers(I).value)
                    Exit Function
                End If
            Next I
        End If
    End If
End Function


