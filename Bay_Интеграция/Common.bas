Attribute VB_Name = "Common"
Option Explicit
'Проэкт\Свойства\Создание\Аргументы компиляции:
' - onErrorOtlad = 1 ' ловля редких err

'Public otlStr As String
Public isOrders As Boolean
Public isFindFirm As Boolean
Public mainTitle As String
Public flReportArhivOrders As Boolean
Public base() As String         '$$2
Public basePath() As String     '

Public myBase As Database
Public wrkDefault As Workspace
Public tbOrders As Recordset
Public tqOrders As Recordset
Public tbSystem As Recordset
Public tbFirms As Recordset
Public tbNomenk  As Recordset
Public tbProduct  As Recordset
Public tbDMC As Recordset
Public tbGuide As Recordset
Public isBlock As Boolean
'Public statId(lenStatus) As Integer
Public status() As String
Public Problems() As String
'Public manags() As String

Public manId() As Integer
Public Manag() As String ' индекс = id из GuideManag

Public insideId() As String
Public Const begCehProblemId = 10 ' начало цеховых проблем в справочнике
'Public temaId() As Integer
Public neVipolnen As Single, neVipolnen_O As Single
Public MaxDay As Integer ' число дней в реестре
'Public tmpMaxDay As Integer 'число дней в окне Zakaz
Public befDays As Integer ' число дней до даты реестра (когда сменилась дата)

'Public baseNamePath As String
'Public otherBase As String
Public begDate As Date ' Дата вступительных остатков
Public logFile As String
Public dostup As String
Public otlad As String
Public tbSIze As Integer
Public cErr As String 'позволяет выявить место возникновения Err, если по
                      'всем местам сообщение об Err выдает один MsgBox
'Public iDate As Date
Public zakazNum As Long  ' кол-во заказов в  Mен.реестре
Public gNzak As String  ' тек номер заказа
Public gFirmId As String
Public gProductId As String
Public gProduct As String
Public gDocDate As Date
Public gSeriaId As String
Public gKlassId As String
Public gNomNom As String
Public numDoc As Long, numExt As Integer


Public oldValue As String 'старое значение поля, измененного последним
Public CurDate As Date
Public lastYear As Integer

Public begDay As Integer ' день первого куска заказа
Public endDay As Integer ' день последнего куска заказа
Public begDayMO As Integer ' день первого куска МО заказа
Public endDayMO As Integer ' день последнего куска МО заказа
Public flEdit As String ' редактируется ресурс
Public Nstan As Single
Public KPD As Single
Public newRes As Single ' смена по умолчанию
Public nr As Single, dr As Single 'убываощие ном. и доп. ресурсы
'Public isDoMO As Boolean ' МО готов или пред. МО был готов - заказ достоверно начал делаться
Public isLive As Boolean ' флаг - заказ живой
Public zagAll As Single, zagLive As Single
Public drobleDopRes As Boolean

Public table As Recordset '
Public myQuery As QueryDef
Public sql As String      ' коллективного пользования
Public strWhere As String '
'Public mousRow As Long
'Public mousCol As Long    '
Public sortGrid As MSFlexGrid
Public trigger As Boolean '
Public tmpDate As Date    '
Public tmpStr As String
Public tmpVar As Variant
Public tmpSng As Single
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
Public orNomZak As Integer, orData As Integer ', orTema As Integer
Public orMen As Integer, orStatus As Integer, orProblem As Integer
Public orFirma As Integer, orDataVid As Integer
Public orVrVid As Integer, orInvoice As Integer, orZakazano As Integer
Public orVes As Integer, orSize As Integer, orPlaces As Integer
Public orOplacheno As Integer, orOtgrugeno As Integer, orLastMen As Integer
Public orVenture As Integer

Public NN() As String, QQ() As Single ' откатываемая номенклатура и кол-во
Public QQ2() As Single, QQ3() As Single
'Public tmpNum As Single ' временая в т.ч. для isNunericTbox()
'Public cehNom As Integer
Public skladId As Integer

Public Const cDELLwidth = 19200 ' это порог а мах = 19290
Private Const dhcMissing = -2 'нужна для quickSort

Public Const gfNazwFirm = 1
Public Const gfM = 2
Public Const gfOborud = 3
Public Const gfRegion = 4 'Регион
Public Const gfSale = 5
Public Const gfKontakt = 6
Public Const gfOtklik = 7
Public Const gf2001 = 8
Public Const gf2002 = 9
Public Const gf2003 = 10
Public Const gf2004 = 11
Public Const gfFIO = 12
Public Const gfTlf = 13
Public Const gfFax = 14
Public Const gfEmail = 15
Public Const gfType = 16
Public Const gfLogin = 17
Public Const gfPass = 18
Public Const gfId = 19

'Если первый пораметр ="W.." - не выдавать Err по невып-ю Where, а все
'параметры обнулить, если для всех них нуль это возможное значение, то в sql
'м. задать константу "1" и принять ее в i. Тогда если i=0 то была Err Where
Function byErrSqlGetValues(ParamArray val() As Variant) As Boolean
Dim tabl As Recordset, i As Integer, maxi As Integer, str As String
Dim c As String

byErrSqlGetValues = False
maxi = UBound(val())
If maxi < 1 Then
    wrkDefault.Rollback
    MsgBox "мало параметров для п\п byErrSqlGetValues()"
    Exit Function
End If
str = CStr(val(0)): c = Left$(str, 1)
If c = "W" Then str = Mid$(str, 2)
'str = Mid$(str, 3)
Set tabl = myOpenRecordSet(str, CStr(val(1)), dbOpenForwardOnly) 'dbOpenDynaset)$#$
'If tabl Is Nothing Then Exit Function
If tabl.BOF Then
    If c <> "W" Then
        wrkDefault.Rollback
        MsgBox "Нет записей удовлетворяющих Where!", , "Error-" & Mid$(str, 3)
        GoTo EN2
    End If
Else
    c = ""
End If

For i = 2 To maxi
    If IsNull(tabl.Fields(i - 2)) Or c = "W" Then
        str = TypeName(val(i))
'        If str = "Single" Or str = "Integer" Or str = "Long" Or str = "Double" Then
        If str = "String" Then
            val(i) = ""
        Else
            val(i) = 0
        End If
    Else
        val(i) = tabl.Fields(i - 2)
    End If
Next i
'EN1:
byErrSqlGetValues = True
EN2:
tabl.Close
End Function

Sub clearGrid(Grid As MSFlexGrid, Optional fixed As Integer = 1)
'Dim il As Long
' On Error GoTo AA
' For il = Grid.Rows To 3 Step -1
'    Grid.RemoveItem (il)
' Next il
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


Sub myRedim(Mass As Variant, newLen As Integer)
Dim maxLen As Integer

maxLen = 0
On Error Resume Next
maxLen = UBound(Mass)
On Error GoTo 0
If newLen < maxLen Then Exit Sub
ReDim Preserve Mass(newLen + 20)
End Sub

Sub delay(tau As Single)
Dim s As Single
    s = Timer
    While Timer - s < tau ' 1 сек
        DoEvents
    Wend

End Sub

Sub exitAll()
If isOrders Then Unload Orders
If isFindFirm Then Unload FindFirm
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


Function getSystemField(field As String) As Variant
getSystemField = Null
Set tbSystem = myOpenRecordSet("##147", "System", dbOpenForwardOnly)
If tbSystem Is Nothing Then myBase.Close: End
getSystemField = tbSystem.Fields(field)
tbSystem.Close
End Function

        
Sub fitFormToGrid(frm As Form, Grid As MSFlexGrid)
Dim i As Long, delta As Long

i = 350 + (Grid.CellHeight + 17) * Grid.Rows
delta = i - Grid.Height
If frm.Height + delta > (Screen.Height - 400) Then _
    delta = (Screen.Height - 400) - frm.Height
frm.Height = frm.Height + delta

'delta = 0
'For i = 0 To Grid.Cols - 1
'    delta = delta + Grid.ColWidth(i)
'Next i
'frm.Width = delta + 700

End Sub

Function getOrdered(numZak As String) As Single
Dim s As Single

getOrdered = -1

sql = "SELECT Sum([sDMCrez].[quantity]*[sDMCrez].[intQuant]/[sGuideNomenk].[perList]) AS cSum " & _
"FROM sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom " & _
"WHERE (((sDMCrez.numDoc)=" & numZak & "));"
If Not byErrSqlGetValues("W##209", sql, s) Then Exit Function
getOrdered = Round(s, 2)
End Function
'Orders.Grid.TextMatrix(Orders.Grid.row, orOtgrugeno)=getShipped()
Function getShipped(numZak As String) As Single
Dim s As Single, s1 As Single, str As String

getShipped = 0
'sql = "SELECT Sum([sDMC].[quant]*[sDMCrez].[intQuant]/[sGuideNomenk].[perList]) AS Выражение1 " & _
"FROM (sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom) INNER JOIN sDMC ON (sDMCrez.nomNom = sDMC.nomNom) AND (sDMCrez.numDoc = sDMC.numDoc) " & _
"WHERE (((sDMCrez.numDoc)=" & numZak & "));"

sql = "SELECT Sum([bayNomenkOut].[quant]*[sDMCrez].[intQuant]) AS bSum " & _
"FROM bayNomenkOut INNER JOIN sDMCrez ON (bayNomenkOut.nomNom = sDMCrez.nomNom) AND (bayNomenkOut.numOrder = sDMCrez.numDoc) " & _
"WHERE (((sDMCrez.numDoc)=" & numZak & "));"
'Debug.Print sql

If Not byErrSqlGetValues("W##209", sql, s) Then Exit Function

getShipped = Round(s, 2)
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


Sub listBoxInGridCell(lb As ListBox, Grid As MSFlexGrid, Optional sel As String = "")
Dim i As Integer
    If Grid.CellTop + lb.Height < Grid.Height Then
        lb.Top = Grid.CellTop + Grid.Top
    Else
        lb.Top = Grid.CellTop + Grid.Top - lb.Height + Grid.CellHeight
    End If
    lb.Left = Grid.CellLeft + Grid.Left
    lb.ListIndex = 0
    If sel <> "" Then
        For i = 0 To lb.ListCount - 1 '
            If Grid.Text = lb.List(i) Then
'                noClick = True
                lb.ListIndex = i 'вызывает ложное onClick
'                noClick = False
                Exit For
            End If
        Next i
    End If
    lb.Visible = True
    lb.ZOrder
    lb.SetFocus
    Grid.Enabled = False 'иначе курсор по ней бегает
'    lbIsActiv = True
End Sub

Function LoadNumeric(Grid As MSFlexGrid, row As Long, col As Integer, _
        val As Variant, Optional myErr As String = "") As Single
 If IsNull(val) Then
    Grid.TextMatrix(row, col) = ""
    LoadNumeric = 0 ' для log файла
    If myErr <> "" Then msgOfZakaz (myErr)
 Else
    LoadNumeric = Round(val, 2)
    Grid.TextMatrix(row, col) = LoadNumeric
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

Sub loadLbMassFromGuide(lbMass() As String, tableName As String)
Dim i As Integer

Set table = myOpenRecordSet("##04", tableName, dbOpenForwardOnly)
If table Is Nothing Then myBase.Close: End
ReDim lbMass(0)
While Not table.EOF
    i = table.Fields(0)
    ReDim Preserve lbMass(i)
    If tableName = "GuideStatus" Then
        If table.Fields(1) = "в работе" Then
            lbMass(i) = "собран" '
        ElseIf table.Fields(1) = "готов" Then
            lbMass(i) = "выдан" '
        Else
            lbMass(i) = table.Fields(1)
        End If
    Else
        lbMass(i) = table.Fields(1)
    End If
    table.MoveNext
Wend
table.Close
'lb.Height = lb.Height + 195 * (lb.ListCount - 1)
End Sub



Sub GridToExcel(Grid As MSFlexGrid, Optional title As String = "")

Dim objExel As Excel.Application, c As Long, r As Long
Dim i As Integer, strA() As String, begRow As Integer, str As String

begRow = 3
If title = "" Then begRow = 1

Set objExel = New Excel.Application
objExel.Visible = True
objExel.SheetsInNewWorkbook = 1
objExel.Workbooks.Add
With objExel.ActiveSheet
.Cells(1, 2).value = title
ReDim Preserve strA(Grid.Cols + 1)
For r = 0 To Grid.Rows - 1
    For c = 1 To Grid.Cols - 1
        str = Grid.TextMatrix(r, c) '=' - наверно зарезервирован для ввода формул
        If Left$(str, 1) = "=" Then str = "." & str
'иногда символы Cr и Lf (поле MEMO в базе) дают Err в Excel, поэтому из поля
        i = InStr(str, vbCr) 'MEMO берем только первую строчку
        If i > 0 Then str = Left$(str, i - 1)
        i = InStr(str, vbLf) 'MEMO берем только первую строчку
        If i > 0 Then str = Left$(str, i - 1)
        strA(c - 1) = str
    Next c
'    On Error Resume Next
   .Range(.Cells(begRow + r, 1), .Cells(begRow + r, Grid.Cols)).FormulaArray = strA
Next r

'objExel.ActiveSheet.Range("A" & begRow & ":U" & Grid.Rows + begRow).FormulaArray = strA
'.Range(.Cells(begRow, 1), .Cells(Grid.Rows + begRow, Grid.Rows)).FormulaArray = strA
End With
Set objExel = Nothing
End Sub

Sub Main()
Dim i As Integer, s As Single, str As String, str1 As String, str2 As String
Dim isXP As Boolean

If App.PrevInstance = True Then
    MsgBox "Программа уже запущена", , "Error"
    End
End If

ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): ReDim QQ3(0) 'чтобы Ubound никогда не давала Err

flReportArhivOrders = False
ReDim tmpL(0)

'If InStr(Command(), ":\") > 0 Then '$$2
'    dostup = "a"
'    otlad = Command()
'Else
If Len(Command()) > 4 Then
    dostup = Mid$(Command(), 6)
    otlad = Left$(Command(), 5)
Else
    dostup = Command()
    otlad = ""
End If
cfg.isLoad = False  '$$2
If Not cfg.loadCfg Then End '$$2


On Error GoTo ERRf 'проверка настройки Win98
s = "1.6"

On Error GoTo ERRs ' не дает Err если в сети не б.найден server, хотя из под DOS дает сист.Err=53
If otlad <> "otlaD" And InStr(otlad, ":\") = 0 Then '
      
'  If dostup = "a" Then 'временно это б. означать  winXP
  If Dir$("C:\WINDOWS\net.exe") = "" Then 'не файла
    Shell "C:\WINDOWS\system32\net time \\server /SET /YES", vbHide ' winXP
  Else
    Shell "C:\WINDOWS\net time \\server /WORKGROUP:JOBSHOP /SET /YES", vbHide
  End If
End If
On Error GoTo 0

'If InStr(otlad, ":\") > 0 Then '$$2
'  str = "\"
'  If Right$(otlad, 1) = "\" Then str = ""
'  baseNamePath = otlad & str & "dlsricN.mdb"
'  mainTitle = "    " & baseNamePath
'Else
If otlad = "otlaD" Then '
'  baseNamePath = "C:\VB_DIMA\dlsricN.mdb"
'  mainTitle = "    " & baseNamePath $$2
  cfg.baseOpen '"C:\VB_DIMA\dlsricN.mdb" $$2
'ElseIf otlad = "otlad" Then
'    baseNamePath = "\\Server\D\!INSTAL!\EPILOG\RADIUS.V20\pitchN.mdb"
'    mainTitle = "      Учебная"
Else
'    mainTitle = ""
'    baseNamePath = "\\Server\D\!INSTAL!\EPILOG\RADIUS.V20\dlsricN.mdb "
    mainTitle = "              New"
    cfg.baseOpen cfg.curBaseInd  '$$2
End If

If dostup = "0" Then i = 5 / i  'проверка наличия библиотеки сообщений

'On Error GoTo ERRb '$$2
'                                                                                                                                                                            Set myBase = OpenDatabase(baseNamePath, False, False, ";PWD=play")
'Set myBase = OpenDatabase(baseNamePath) '$$2
'If myBase Is Nothing Then End

On Error GoTo 0

'Set wrkDefault = DBEngine.Workspaces(0) ' для орг-ии транзакций


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

str = Format(Now, "dd.mm.yyyy")
CurDate = str 'без часов и минут

' обрезание лог файла*********************************************
 
 'logFile = "C:\Windows\OrdersBay" ' без расширения
 logFile = App.Path & "\" & App.EXEName
 str2 = logFile & "$$.log" ' временный файл
 logFile = logFile & ".log"
 
 On Error GoTo ENop
 Open logFile For Input As #2
 Open str2 For Output As #3
 While Not EOF(2)
    Input #2, str
    i = InStr(str, vbTab)
    If i < 9 Then GoTo ENlog
    str1 = Left$(str, i - 1)
    If Not IsDate(str1) Then GoTo ENlog
    'tmpDate = str
    If DateDiff("d", str1, CurDate) <= 7 Then Print #3, str ' удаляем > 7ми дней давности
 Wend
ENlog:
 Close #2
 Close #3
 Kill logFile
 Name str2 As logFile
ENop:
isBlock = False
noClick = False

loadLbMassFromGuide Problems(), "BayGuideProblem"
loadLbMassFromGuide status(), "GuideStatus"
'loadLbMassFromGuide manags, "GuideManag" $$7
Orders.Show

Exit Sub
ERRf:
MsgBox "Пуск\Настройка\Панель Управления\Язык и стандарты\Числа\" & _
      "Разделитель целой и дробной частей числа " & Chr(151) & _
      " Установите точку вместо запятой!", , "Для правильной работы " & _
      "программы необходимо выполнить настройку Win98: "
End

'ERRb:
'MsgBox "Не удалось подключиться к базе " & mainTitle
'End

ERRs:
MsgBox "Система не смогла синхронизировать часы", , "Сообщите администратору!"
Resume Next

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

Sub msgOfZakaz(myErrCod As String, Optional msg As String = "")
    wrkDefault.Rollback

    myErrCod = Mid$(myErrCod, 3)
    If msg = "" Then msg = "Нарушена целостность данных."
    MsgBox msg & " Работа с этим заказом пока " & _
    "невозможна. Сообщите администратору!", , _
    "Обнаружен сбой базы (Err=" & myErrCod & ") в заказе № " & gNzak
End Sub

'Sub msgOfEnd(myErrCod As String, Optional msg As String = "")
'    wrkDefault.Rollback
'
'    myErrCod = Mid$(myErrCod, 3)
'    MsgBox msg & " Сообщите администратору!", , "Ошибка " & myErrCod
'    End
'End Sub

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
                typ As Integer) As Recordset

On Error GoTo ErrorHandler

Set myOpenRecordSet = myBase.Connection.OpenRecordset(sours, dbOpenDynaset, dbExecDirect, dbPessimistic)

Exit Function

ErrorHandler:
    
errorCodAndMsg (Mid$(myErrCod, 3))

myBase.Close: End

End Function


'NULL - предметов нет вообщее
'skladId=0 - cуммарно по всем складам
'skladId=2 - cуммарно по 1 и 2му складам





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


Sub rowViem(numRow As Long, Grid As MSFlexGrid)
Dim i As Integer

i = Grid.Height \ Grid.RowHeight(1) - 1 ' столько умещается строк
i = numRow - i \ 2 ' в центр
If i < 1 Then i = 1
Grid.TopRow = i
End Sub

'эта ф-я д.заменить и startDay() и getNextDay() и getPrevDay()
' возвращает смещение до треб. дня
Function getWorkDay(offsDay As Integer, Optional baseDate As String = "") As Integer
Dim i As Integer, j As Integer, step  As Integer
getWorkDay = -1
If baseDate = "" Then
    tmpDate = CurDate
Else
    If Not IsDate(baseDate) Then Exit Function
    tmpDate = baseDate
End If

step = 1
If offsDay < 0 Then step = -1

j = 0: i = 0
While step * j < step * offsDay '
    i = i + step
'    day = Weekday(tmpDate + i)
    day = Weekday(DateAdd("d", i, tmpDate))
    If Not (day = vbSunday Or day = vbSaturday) Then j = j + step
Wend
getWorkDay = i
'tmpDate = tmpDate + i
tmpDate = DateAdd("d", i, tmpDate)

End Function


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
Function isDateTbox(tBox As TextBox, Optional fryDays As String = "") As Boolean
Dim str As String

isDateTbox = True
str = tBox.Text
If str = "" Then
        MsgBox "Заполните поле Даты!", , "Ошибка"
Else
'    If Not IsDate(str) Then
'    If Len(str) <> 8 Or Not IsDate(str) Then
'        MsgBox "неверный формат даты", , "Ошибка"
'    Else
        'str = Left$(str, 6) & "20" & Mid$(str, 7, 2)
        str = "20" & Right$(str, 2) & "-" & Mid$(str, 4, 2) & "-" & Left$(str, 2)
        If IsDate(str) Then
            tmpDate = str
            If fryDays = "" Then
                Exit Function
            Else
                day = Weekday(tmpDate)
                If day = vbSunday Or day = vbSaturday Then
                    If MsgBox(str & " - выходной день. Продолжить?", vbYesNo, _
                    "Предупреждение!") = vbYes Then Exit Function
                Else
                    Exit Function
                End If
            End If
        Else
            MsgBox "Неверный формат даты или дня с такой датой не существует ", , "Ошибка"
        End If
'    End If
End If
 '   tBox.Text = oldValue
tBox.SetFocus
tBox.SelStart = 0
tBox.SelLength = Len(tBox.Text)
isDateTbox = False
End Function

Sub valueToSystemField(myErr As String, val As Variant, field As String)
Set tbSystem = myOpenRecordSet(myErr, "System", dbOpenTable)
If tbSystem Is Nothing Then myBase.Close: End
tbSystem.Edit
tbSystem.Fields(field) = val
tbSystem.Update
tbSystem.Close
End Sub

'не записыват неуникальное значение, для полей, где такие
'значения запрещены. А  генерит при этом error?
Function ValueToTableField(myErrCod As String, value As String, table As String, _
field As String, Optional by As String = "") As Integer
Dim sql As String, byStr As String ', numOrd As String

ValueToTableField = False
If value = "" Then value = Chr(34) & Chr(34)
If by = "" Then
    byStr = ".numOrder)= " & gNzak
ElseIf by = "byFirmId" Then
    byStr = ".FirmId)= " & gFirmId
ElseIf by = "byKlassId" Then
    byStr = ".klassId)= " & gKlassId
ElseIf by = "byNomNom" Then
    byStr = ".nomNom)= " & "'" & gNomNom & "'"
ElseIf by = "bySeriaId" Then
    byStr = ".seriaId)= " & gSeriaId
ElseIf by = "byProductId" Then
    byStr = ".prId)= " & gProductId
'ElseIf by = "bySourceId" Then
'    byStr = ".sourceId)= " & gSourceId
ElseIf by = "byNumDoc" Then
    sql = "UPDATE " & table & " SET " & table & "." & field & "=" & value _
        & " WHERE (((" & table & ".numDoc)=" & numDoc & " AND (" & table & _
        ".numExt)=" & numExt & " ));"
    GoTo AA
Else
    Exit Function
End If
sql = "UPDATE " & table & " SET " & table & "." & field & _
" = " & value & " WHERE (((" & table & byStr & " ));"
AA:
'MsgBox "sql = " & sql

If Left$(myErrCod, 1) = "W" Then
    myErrCod = Mid$(myErrCod, 2)
    ValueToTableField = myExecute(myErrCod, sql, 0) 'не сообщать если не WHERE
Else
    ValueToTableField = myExecute(myErrCod, sql)
End If
End Function



Function beNaklads(Optional reg As String = "") As Boolean
beNaklads = True
Dim s As Single
'отпущено
sql = "SELECT Sum(sDMC.quant) AS Sum_quant From sDMC " & _
"WHERE (((sDMC.numExt)< 254) AND ((sDMC.numDoc)=" & numDoc & "));"
If Not byErrSqlGetValues("##140", sql, s) Then Exit Function
If s > 0.005 Then ' что-то отпущено
    If reg = "" Then
        MsgBox "По этому заказу выписывались накладные, поэтому изменять " & _
        "предметы нельзя. Если изменения все-же требуются, то прежде надо " & _
        "удалить все накладные к заказу.", , "Редактирование запрещено!"
    End If
Else
    beNaklads = False
End If

End Function
    
Function PrihodRashod(reg As String, skladId As Integer) As Single
Dim qWhere As String, s As Single

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
'MsgBox sql
byErrSqlGetValues "##157", sql, PrihodRashod

'If skladId >= -1001 And reg = "+" Then
'    sql = "SELECT sGuideNomenk.begOstatki From sGuideNomenk " & _
'    "WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
'    If Not byErrSqlGetValues("##161", sql, s) Then Exit Function
'    PrihodRashod = PrihodRashod + s
'End If
End Function
    
Sub unLockBase()
valueToSystemField "##148", "", "resursLock"
End Sub
Function getLockYear() As Integer '$$3
getLockYear = Format(begDate, "yyyy")
If Format(begDate, "dd.mm") = "01.01" Then _
    getLockYear = getLockYear - 1 'считаем, что этот год не учавствовал в отсечении базы
End Function

Function ostatCorr(delta As Single) As Boolean
Dim sId As Integer, dId As Integer

ostatCorr = False

sql = "SELECT sDocs.sourId, sDocs.destId, sDocs.numDoc, sDocs.numExt " & _
"From sDocs " & _
"WHERE (((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & "));"
If Not byErrSqlGetValues("##180", sql, sId, dId) Then Exit Function

If sId < -1000 And dId < -1000 Then ' для межскладских не корректируем
        ostatCorr = True
Else
    sql = "UPDATE sGuideNomenk SET nowOstatki = [nowOstatki]-" & delta & _
    " WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
    If myExecute("##163", sql) <> 0 Then Exit Function
    ostatCorr = True
End If
End Function


Function predmetiIsClose() As Variant
Dim i As Integer, s As Single

predmetiIsClose = Null
'If gNzak = 4092402 Then
'i = i
'End If

'sql = "SELECT sDMCrez.quantity, sDMC.quant " & _
"FROM sDMCrez LEFT JOIN sDMC ON (sDMCrez.nomNom = sDMC.nomNom) AND (sDMCrez.numDoc = sDMC.numDoc) " & _
"Where (((sDMCrez.numDoc) = " & gNzak & "));"
sql = "SELECT sDMCrez.nomNom, sDMCrez.quantity, Sum(sDMC.quant) AS Sum_quant " & _
"FROM sDMCrez LEFT JOIN sDMC ON (sDMCrez.nomNom = sDMC.nomNom) AND (sDMCrez.numDoc = sDMC.numDoc) " & _
"Where (((sDMCrez.numDoc) = " & gNzak & ")) " & _
"GROUP BY sDMCrez.nomNom, sDMCrez.quantity;"
Set tbDMC = myOpenRecordSet("##350", sql, dbOpenForwardOnly)
If tbDMC Is Nothing Then Exit Function
If Not tbDMC.BOF Then
  While Not tbDMC.EOF
    If IsNull(tbDMC!Sum_quant) Then
        GoTo AA
    ElseIf tbDMC!quantity > tbDMC!Sum_quant + 0.005 Then
AA:     predmetiIsClose = False
        GoTo EN1
    End If
    tbDMC.MoveNext
  Wend
  predmetiIsClose = True
End If
EN1:
tbDMC.Close
End Function

Function getYearField(checkDate As Date) As String '$$3
Dim i As Integer, lockYear As Integer

lockYear = getLockYear
i = Format(checkDate, "yyyy")
'If i <= lockYear Then
'    getYearField = "lock" 'этот год учавствовал в отсечении базы
'    Exit Function
'End If
i = i - lastYear + 4 'номер колонки
If i < 1 Then     'если это не последние 3 года, то в кучу
    getYearField = "year01"
Else
    getYearField = "year" & Format(i, "00")
End If
End Function


Sub visits(oper As String, Optional firm As String = "")
Dim str As String, i As Integer, statId As Integer

sql = "SELECT inDate, StatusId , FirmId From BayOrders " & _
"WHERE (((numOrder)=" & gNzak & "));"
'MsgBox sql
If Not byErrSqlGetValues("##88", sql, tmpDate, statId, i) Then GoTo ER1

If i = 0 Then Exit Sub
If firm <> "" And (statId = 0 Or statId = 7) Then Exit Sub ' если меняем фирму

'str = "year" & Format(tmpDate, "yy")
str = getYearField(tmpDate) '$$3

sql = "UPDATE BayGuideFirms SET BayGuideFirms." & str & " = [BayGuideFirms].[" & _
str & "] " & oper & " 1  WHERE (((BayGuideFirms.FirmId)=" & i & "));"
'MsgBox sql
i = myExecute("##87", sql, -143)

'If i <> 3061 And i <> 0 Then '3061 - колонки этого года уже(или еще) нет в базе
If i = -2 Then '3061 - колонки этого года уже(или еще) нет в базе
ER1:    MsgBox "Ошибка коррекции посещения фирм. Сообщите администратору!", , "Error-87"
End If
End Sub

Function lockSklad(Optional back As String = "") As Boolean
Dim str As String

lockSklad = True: Exit Function '!!! временно отключаем
lockSklad = False
RETR:
Set tbSystem = myOpenRecordSet("##94", "System", dbOpenTable) ', dbOpenForwardOnly)
If tbSystem Is Nothing Then myBase.Close: End
tbSystem.Edit
str = tbSystem!skladLock
If Not back = "" Then
    If str = Orders.cbM.Text Then tbSystem!skladLock = ""
Else
    If str <> "" And str <> Orders.cbM.Text Then
        tbSystem.Update: tbSystem.Close
        
        If MsgBox("Доступ к остаткам на складе временно занят менеджером '" & _
        str & "'. Повторите попытку или обратитесь к Администратору.", _
        vbRetryCancel, "Нет доступа !!!") = vbRetry Then
            GoTo RETR
        Else
            Exit Function
        End If
    End If
    tbSystem!skladLock = Orders.cbM.Text
End If
tbSystem.Update
tbSystem.Close
lockSklad = True
End Function

