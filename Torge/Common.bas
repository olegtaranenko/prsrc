Attribute VB_Name = "Common"
Option Explicit

Private Const dhcMissing = -2 'нужна для quickSort
Public objExel As Excel.Application, exRow As Long
Public gain2 As Single, gain3 As Single, gain4 As Single
Public head1 As String, head2 As String, head3 As String, head4 As String

Public sql As String, strWhere As String
Public webProducts As String
Public webNomenks As String
Public mainTitle As String

Public loginsPath As String
Public SvodkaPath As String
Public NomenksPath As String
Public ProductsPath As String


Public tbSystem As Recordset
Public table As Recordset
Public tbNomenk As Recordset
Public tbProduct As Recordset
Public tbDocs As Recordset
Public tbDMC As Recordset
Public tbGuide As Recordset
Public tbOrders As Recordset
Public gNzak As String  ' тек номер заказа
Public gFirmId As String
Public sortGrid As MSFlexGrid
Public trigger As Boolean
Public dostup As String
Public otlad As String
Public CurDate As Date
Public tmpDate As Date
Public tmpSingle As Single
Public tmpStr As String
Public tmpLong As Long
Public tmpSng As Single
Public cErr As String '

Public day As Integer     '
Public bilo As Boolean
Public otherBase As String
Public manId() As Integer
Public Manag() As String
Public Status() As String
Public insideId() As String
Public gKlassId As String
Public gKlassType As String
Public gNomNom As String
Public gSeriaId As String
Public gProduct As String
Public gProductId As String
Public prExt As Integer
Public cenaFreight As String, cenaFact As String

Public gSourceId As String
Public gDocDate As Date
Public mousRight As Integer
Public nodeKey As String
Public prevRow As Long
Public gridIsLoad As Boolean
Public DMCnomNom() As String ' номер(а), кот в загруж.карточке
Public DMCklass As String ' номер группы, кот в загруж.карточке
Public tmpNum As Single ' временая в т.ч. для isNunericTbox()
'Public CENA1 As Single, VES As Single, STAVKA As Single
Public sc ' ScriptControl
Public noClick As Boolean
Public beChange As Boolean 'была правка в textBox
Public debit As String, subDebit As String, kredit As String, subKredit As String
Public detailId As Integer, purposeId As Integer, KredDebitor As Integer
Public numDoc As Long, numExt As Integer
Public begDate As Date ' Дата вступительных остатков
Public NN() As String, QQ() As Single ' откатываемая номенклатура и кол-во
Public QQ3() As Single, QQ2() As Single ' вспомагательое откатываемое кол-во
Public bulkChangEnabled As Boolean
Public Const otladColor = &H80C0FF
Public sqlRowDetail() As String
Public aRowText() As String
Public rowFormatting() As String
Public aRowSortable() As Boolean
Public arowSubtitle() As Boolean
Public startDate As String, endDate As String
Public rate As Variant


Function RateAsString(ByVal curRate As Double) As String
    
    Const rubleRoot As String = "рубл"

    Dim strRate As String, strRate00 As String
    Dim rubleSuffix As String
    
    strRate00 = CDbl(Format(getCurrentRate, "##0.00"))
    strRate = CDbl(Format(getCurrentRate, "##0"))
    If CDbl(strRate) <> CDbl(strRate00) Then
        strRate = strRate00
        rubleSuffix = "я"
    Else
        Dim strLastDigit As String, strLastTwoDigit As String
        Dim digit As Integer
        If Len(strRate) >= 2 Then
            strLastTwoDigit = right(strRate, 2)
            digit = CInt(strLastTwoDigit)
            If digit >= 5 And digit <= 20 Then
                rubleSuffix = "ей"
            End If
        End If
        If rubleSuffix = "" Then
            strLastDigit = right(strRate, 1)
            digit = CInt(strLastDigit)
            If digit = 1 Then
                rubleSuffix = "ь"
            ElseIf digit > 1 And digit < 5 Then
                rubleSuffix = "я"
            Else
                rubleSuffix = "ей"
            End If
        End If

    End If

  
    RateAsString = "1 у.е. = " & strRate & " " & rubleRoot & rubleSuffix
End Function


Function getCurrentRate() As Double
Dim s As String

    sql = "SELECT Kurs FROM System;"
    If byErrSqlGetValues("##321", sql, s) Then
        getCurrentRate = Abs(s)
    End If

End Function

Function dateBasic2Sybase(aDay As String)
Dim dt_str As String

dt_str = "20" & right$(aDay, 2) & "-" & Mid$(aDay, 4, 2) & "-" & left$(aDay, 2)
dateBasic2Sybase = CDate(dt_str)

End Function


Function dateSybase2Basic(aDay As String)
Dim dt_str As String

dt_str = left(aDay, 4) & "-" & Mid(aDay, 5, 2) & "-" & right(aDay, 2)
dateSybase2Basic = dt_str

End Function

Sub setStartEndDates(tbStartDate As TextBox, Optional tbEndDate As TextBox)
'    setStartEndDates tbStartDate, tbEndDate
    startDate = "null"
    If isDateTbox(tbStartDate) Then
        startDate = "'" & Format(tmpDate, "yyyymmdd") & "'"
    End If
    
    endDate = "null"
    If isDateTbox(tbEndDate) Then
        If isDateTbox(tbEndDate) Then
            endDate = "'" & Format(tmpDate, "yyyy-mm-dd") & " 11:59:59 PM'"
        End If
    End If

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
    
    If servers!sysname = "stime" Then
        If servers!standalone = 1 Then
            bulkChangEnabled = False
        Else
            bulkChangEnabled = True
        End If
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
    GoTo cont
no_access:
cont:
    servers.MoveNext
Wend
servers.Close


End Sub


Sub clearGridRow(Grid As MSFlexGrid, row As Long)
Dim il As Long
    noClick = True
    Grid.row = row
    For il = 0 To Grid.Cols - 1
      If il > 0 Then
        Grid.col = il
        Grid.CellBackColor = Grid.BackColor
        Grid.CellForeColor = Grid.ForeColor
      End If
      Grid.TextMatrix(row, il) = ""
    Next il
    Grid.col = 1
    noClick = False
End Sub

Sub colorGridRow(Grid As MSFlexGrid, row As Long, color As Long)
Dim il As Long
    noClick = True
    Grid.row = row
    For il = 0 To Grid.Cols - 1
        Grid.col = il
        If il > 0 Then Grid.CellBackColor = color
    Next il
    Grid.col = 1
    noClick = False
End Sub

Sub foreColorGridRow(Grid As MSFlexGrid, row As Long, color As Long, ccol As Long)
Dim il As Long
    noClick = True
    Grid.row = row
    For il = 0 To Grid.Cols - 1
        Grid.col = il
        If il > 0 Then Grid.CellForeColor = color
    Next il
    Grid.col = ccol
    noClick = False
End Sub

Sub colorGridCell(Grid As MSFlexGrid, ByVal row As Long, ByVal col As Long, ByVal color As Long)
Dim il As Long
    noClick = True
    Grid.row = row
    Grid.col = col
    Grid.CellForeColor = color
    noClick = False
End Sub


Function myIsDate(ByVal dt As String) As Variant
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
        v_dd = CInt(left(dt, dotPos))
        dt = Mid(dt, dotPos + 1)
    End If
    
    If Len(dt) > 0 Then
        dotPos = InStr(dt, ".")
        If IsNull(dotPos) Or dotPos = 0 Then
            v_mm = CInt(dt)
            dt = ""
        Else
            v_mm = CInt(left(dt, dotPos))
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
    
    myIsDate = Format(CDate(CStr(v_yyyy) & "-" & Format(v_mm, "0#") & "-" & Format(v_dd, "0#")), "dd.mm.yy")
    Exit Function
catch:
    myIsDate = False
End Function


Function isDateEmpty(tBox As TextBox, Optional warn As Boolean = True) As Boolean
Dim str As String
Dim dt As String
Dim ret As Variant
    
    dt = tBox.Text
    If IsEmpty(dt) Or Len(CStr(dt)) = 0 Then
        isDateEmpty = False
        Exit Function
    End If
    
    dt = tBox.Text
    
    ret = myIsDate(dt)
    If IsDate(ret) Then
        isDateEmpty = True
        tBox.Text = ret
    Else
        If warn Then
            MsgBox "Неверная дата: " & CStr(dt)
        End If
        isDateEmpty = False
        tBox.SetFocus
        tBox.SelStart = 0
        tBox.SelLength = Len(tBox.Text)
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
        str = "20" & right$(str, 2) & "-" & Mid$(str, 4, 2) & "-" & left$(str, 2)
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


Function LoadNumeric(Grid As MSFlexGrid, row As Long, col As Integer, _
        val As Variant, Optional myErr As String = "") As Single
 If IsNull(val) Then
    Grid.TextMatrix(row, col) = ""
    LoadNumeric = 0 ' для log файла
    If myErr <> "" Then msgOfZakaz (myErr)
 Else
    LoadNumeric = val
    Grid.TextMatrix(row, col) = LoadNumeric
 End If
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

Sub listBoxInGridCell(lb As ListBox, Grid As MSFlexGrid, Optional sel As String = "")
Dim I As Integer, l As Long
    If lb.ListCount < 200 Then
        l = CLng(195) * CLng(lb.ListCount) + 100 ' Для больших списков
    Else
        l = Grid.Height / 2 + 500
    End If
        
    If l > Grid.Height / 2 + 500 Then l = Grid.Height / 2 + 500
    lb.Height = l
    
    If Grid.CellTop + lb.Height < Grid.Height Then
        lb.Top = Grid.CellTop + Grid.Top
    Else
        lb.Top = Grid.CellTop + Grid.Top - lb.Height + Grid.CellHeight
    End If
    lb.left = Grid.CellLeft + Grid.left
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


Sub Main()
Dim str As String, I As Integer

If App.PrevInstance = True Then
    MsgBox "Программа уже запущена", , "Error"
    End
End If

mainTitle = getMainTitle


Set sc = CreateObject("ScriptControl")
sc.Language = "VBScript"

GuideSource.isLoad = False
KartaDMC.isLoad = False
Nomenklatura.isRegimLoad = False
Products.isLoad = False
ReDim DMCnomNom(0)



CurDate = Now

loadEffectiveSettingsApp

webProducts = getEffectiveSetting("ProductsPath")
webNomenks = getEffectiveSetting("NomenksPath")
loginsPath = getEffectiveSetting("loginsPath")
SvodkaPath = getEffectiveSetting("SvodkaPath")

checkReloadCfg

baseOpen

CheckIntegration

    sql = "create variable @manager varchar(20)"
    If myExecute("##0.2", sql, 0) = 0 Then End
    
    If Not initFomulConstats Then
        MsgBox "Ошибка при инициализации формул" _
            & vbCr & "Работа тем не менее будет продолжена", vbOKOnly Or vbExclamation, "Обратитесь к администратору"
    End If


    AUTO.Show

End Sub

Private Function initFomulConstats() As Boolean
    initFomulConstats = True
    On Error GoTo er
    ' init the Global constants to use its in formulas
    sql = "select * from GuideConstants"
    Set tbGuide = myOpenRecordSet("##0.3", sql, dbOpenForwardOnly)
    If tbGuide Is Nothing Then GoTo er
    While Not tbGuide.EOF
        Dim initStr As String
        initStr = tbGuide!Constants & "=" & CDbl(tbGuide!value)
        sc.ExecuteStatement (initStr)
        tbGuide.MoveNext
    Wend
    tbGuide.Close
    Exit Function
er:
    initFomulConstats = False
End Function

Function max(val1, val2)
If val2 > val1 Then
    max = val2
Else
    max = val1
End If
End Function

Sub msgOfEnd(myErrCod As String, Optional msg As String = "")
    wrkDefault.Rollback

    myErrCod = Mid$(myErrCod, 3)
    MsgBox msg & " Сообщите администратору!", , "Ошибка " & myErrCod
    myBase.Close
    End
End Sub

Sub msgOfZakaz(myErrCod As String)
    wrkDefault.Rollback
    myErrCod = Mid$(myErrCod, 3)
    MsgBox "Нарушена целостность данных. Работа с этим заказом пока " & _
    "невозможна. Сообщите администратору!", , _
    "Ошибка " & myErrCod & " в заказе № " & gNzak
End Sub

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
'  If passErr > 0 Or passErr = -11111 Then
  If passErr <> 0 Then _
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

Function ValueToGuideSourceField(myErrCod As String, value As String, _
field As String, Optional passErr As Integer = -11111) As Integer
Dim I As Integer

ValueToGuideSourceField = False
sql = "UPDATE sGuideSource SET [" & field & _
"] = '" & value & "' WHERE (((sourceId)=" & gSourceId & "));"
'MsgBox "sql = " & sql

ValueToGuideSourceField = myExecute(myErrCod, sql, passErr)
End Function

Function errorCodAndMsg(line As String, Optional passErr As Integer = 22222) As Boolean
Dim strError As String
Dim errLoop
   
   strError = "": errorCodAndMsg = True
   For Each errLoop In Errors
      With errLoop
         If .number = passErr Then Exit Function
         
         strError = strError & _
            "******** Error: '" & .number & "' *********" & vbCr
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


Function myOpenRecordSet(myErrCod As String, sours As String, _
                typ As Integer) As Recordset

On Error GoTo ErrorHandler

Set myOpenRecordSet = myBase.Connection.OpenRecordset(sours, dbOpenDynaset, dbExecDirect, dbPessimistic)

Exit Function

ErrorHandler:
    
errorCodAndMsg (Mid$(myErrCod, 3))

myBase.Close: End

End Function



Sub myRedim(Mass As Variant, newLen As Integer)
Dim maxLen As Integer

maxLen = 0
On Error Resume Next
maxLen = UBound(Mass)
On Error GoTo 0
If newLen < maxLen Then Exit Sub
ReDim Preserve Mass(newLen + 20)
End Sub

Function nomenkFormula(Optional noOpen As String = "", Optional web As String = "")
Dim str As String

If noOpen = "" Then
    sql = "SELECT sGuideNomenk.formulaNom" & web & " , sGuideNomenk.CENA1, " & _
    "sGuideNomenk.VES, sGuideNomenk.STAVKA, sGuideFormuls.Formula as formula" & web & _
    " FROM sGuideFormuls INNER JOIN sGuideNomenk ON sGuideFormuls.nomer = " & _
    "sGuideNomenk.formulaNom" & web & _
    " WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
'MsgBox sql
    Set tbNomenk = myOpenRecordSet("##317", sql, dbOpenDynaset)
    If tbNomenk Is Nothing Then Exit Function
    If tbNomenk.BOF Then tbNomenk.Close: Exit Function
End If
tmpStr = tbNomenk!formula
tmpStr = tbNomenk.fields("formula" & web)
'If tbNomenk!formula = "" Then
If tmpStr = "" Then
    nomenkFormula = "error: Формула не задана"
    Exit Function
End If
If web = "" Then
    str = "CENA1=" & tbNomenk!Cena1 & ": VES=" & _
    tbNomenk!ves & ": STAVKA=" & tbNomenk!STAVKA
    sc.ExecuteStatement (str)
Else
    str = "CenaFreight=" & cenaFreight & ": CenaFact=" & cenaFact
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

Sub rowViem(numRow As Long, Grid As MSFlexGrid)
Dim I As Integer

I = Grid.Height \ Grid.RowHeight(1) - 1 ' столько умещается строк
I = numRow - I \ 2 ' в центр
If I < 1 Then I = 1
Grid.TopRow = I

End Sub

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


Sub backNomenk()
Dim q As Single, I As Integer, str As String, n As Integer, rr As Integer

wrkDefault.BeginTrans

For rr = 1 To UBound(QQ)
  Set tbDMC = myOpenRecordSet("##157", "sDMC", dbOpenTable)
  If tbDMC Is Nothing Then GoTo EN1
  tbDMC.index = "nomDoc"
  tbDMC.Seek "=", numDoc, numExt, NN(rr)
  If tbDMC.NoMatch Then
      tbDMC.Close
      GoTo EN1
  End If
  q = Round(tbDMC!quantity - QQ(rr), 2)
  If q = 0 Then
     tbDMC.Delete
  Else
    tbDMC.Edit
    tbDMC!quantity = q
    tbDMC.Update
  End If
  tbDMC.Close
  
    Set tbNomenk = myOpenRecordSet("##156", "sGuideNomenk", dbOpenTable)
    If tbNomenk Is Nothing Then GoTo EN1
    tbNomenk.index = "PrimaryKey"
    tbNomenk.Seek "=", NN(rr)
    If tbNomenk.NoMatch Then
        tbNomenk.Close
EN1:    wrkDefault.Rollback
        MsgBox "Не удалось отменить операцию по этому документу, поэтому " & _
        "проделайте необходимые изменения вручную.", , "Внимание!"
        GoTo EN2
    End If
    tbNomenk.Edit
    tbNomenk!nowOstatki = Round(tbNomenk!nowOstatki - QQ(rr), 2)
    tbNomenk.Update
    tbNomenk.Close

Next rr
wrkDefault.CommitTrans
EN2:
ReDim QQ(0)
End Sub

Sub getDocExtNomFromStr(nom As String)
Dim I As Integer
I = InStr(nom, "/")
If I = 0 Then
    numDoc = nom
    numExt = 254
'ElseIf i = Len(nom) Then
'    numDoc = Left$(nom, i - 1)
'    numExt = 0
Else
    numDoc = left$(nom, I - 1)
    numExt = Mid$(nom, I + 1)
End If
End Sub

Function getPurposeIdByDescript(purpose As String) As Integer
Dim id As Integer, exists As Integer

getPurposeIdByDescript = 255

sql = "SELECT 1, pId From yGuidePurpose " & _
"WHERE (((Debit)='" & debit & "') AND ((subDebit)='" & _
subDebit & "') AND ((Kredit)='" & kredit & "') AND ((" & _
"subKredit)='" & subKredit & "') AND ((pDescript)='" & purpose & "'));"

Debug.Print sql
If byErrSqlGetValues("W#356", sql, exists, id) Then
    If exists = 0 Then Exit Function
End If

getPurposeIdByDescript = id
End Function


Function getSystemField(field As String) As Variant
getSystemField = Null
Set tbSystem = myOpenRecordSet("##147", "System", dbOpenForwardOnly)
If tbSystem Is Nothing Then Exit Function
getSystemField = tbSystem.fields(field)
tbSystem.Close
End Function

Function valueToSystemField(val As Variant, field As String) As Boolean
valueToSystemField = False
Set tbSystem = myOpenRecordSet("##148", "System", dbOpenTable)
If tbSystem Is Nothing Then Exit Function
tbSystem.Edit
tbSystem.fields(field) = val
tbSystem.Update
tbSystem.Close
valueToSystemField = True
End Function

'выдает "error"- если нарушен формат дат (не следует запускать SQL) .
'reg="" -  выдает аргумент для WHERE для промежутка между датами
'          либо "" - ограничния во WHERE по дате не требуется(с учетом begDate и CurDate)
'          либо "error" если даты не пересекаются
'reg<>"" - выдает аргумент для WHERE для промежутка До startDate
'          либо "" если startDate раньше begDate(не следует запускать SQL)
Function getWhereByDateBoxes(frm As Form, dateField As String, _
begDate As Date, Optional reg As String = "") As String

Dim str As String, ckStart As Boolean, ckEnd  As Boolean

getWhereByDateBoxes = "": str = "":

ckStart = False: ckEnd = False
On Error Resume Next ' на случай, если в этой форме у дат нет флажков
If frm.ckEndDate.value > 0 Then ckEnd = True  'то они как бы установлены
If frm.ckStartDate.value > 0 And frm.ckStartDate.Visible Then ckStart = True
On Error GoTo 0

If ckStart Then
    If Not isDateTbox(frm.tbStartDate) Then GoTo ERRd  'tmpDate
End If
If reg = "" Then ' если период Между
    If DateDiff("d", begDate, tmpDate) > 0 And ckStart Then _
        str = "(" & dateField & ") >='" & Format(tmpDate, "yyyy-mm-dd") & "'"
    If ckEnd Then
      If Not isDateTbox(frm.tbEndDate) Then GoTo ERRd
      If ckStart Then
        If DateDiff("d", frm.tbStartDate.Text, tmpDate) < 0 Then
          MsgBox "Начальная дата периода загрузки не должна превышать конечную ", , "Предупреждение"
ERRd:     getWhereByDateBoxes = "error"
          Exit Function
        End If
      End If
      If DateDiff("d", tmpDate, CurDate) > 0 Then getWhereByDateBoxes = _
          "(" & dateField & ")<='" & Format(tmpDate, "yyyy-mm-dd") & " 11:59:59 PM'"
    End If
ElseIf ckStart Then ' если период До
    If DateDiff("d", begDate, tmpDate) <= 0 Then Exit Function
    tmpDate = DateAdd("d", -1, tmpDate) ' "-1" день т.к. далее "+ 23ч59м59с
    If DateDiff("d", tmpDate, CurDate) > 0 Then getWhereByDateBoxes = _
        "(" & dateField & ")<='" & Format(tmpDate, "yyyy-mm-dd") & " 11:59:59 PM'"
End If
If str <> "" And getWhereByDateBoxes <> "" Then
    getWhereByDateBoxes = str & " AND " & getWhereByDateBoxes
Else
    getWhereByDateBoxes = str & getWhereByDateBoxes
End If
End Function

'Function getStrNumDoc(N_Doc As Long, N_Ext As Integer) As String
'    getStrNumDoc = N_Doc
'    If N_Ext = 0 Then
'        getStrNumDoc = getStrNumDoc & "/"
'    ElseIf N_Ext < 255 Then
'        getStrNumDoc = getStrNumDoc & "/" & N_Ext
'    End If
'End Function
    
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
    tmpNum = val
    If Not IsMissing(maxVal) Then
        If (minVal > tmpNum Or tmpNum > maxVal) Then
            MsgBox "значение должно быть в диапазоне от " & minVal & _
            "  до " & maxVal, , "Error"
            checkNumeric = False
        End If
    ElseIf Not IsMissing(minVal) Then
        If minVal > tmpNum Then
            MsgBox "значение должно быть больше " & minVal
            checkNumeric = False
        End If
    End If
Else
    MsgBox "Недопустимое значение", , "Error"
    checkNumeric = False
End If
End Function
Sub clearGrid1(Grid As MSFlexGrid)
    Grid.Rows = 1
End Sub


Sub clearGrid(Grid As MSFlexGrid)
Dim il As Long
Grid.Rows = 2
' For il = Grid.Rows To 3 Step -1
'    Grid.RemoveItem (il)
' Next il
 clearGridRow Grid, 1
End Sub

'Если первый пораметр ="W.." - не выдавать Err по невып-ю Where, а все
'параметры обнулить, если для всех них нуль это возможное значение, то в sql
'м. задать константу "1" и принять ее в i. Тогда если i=0 то была Err Where
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
str = CStr(val(0)): c = left$(str, 1)
If c = "W" Then str = Mid$(str, 2)
Set tabl = myOpenRecordSet(str, CStr(val(1)), dbOpenForwardOnly) 'dbOpenDynaset)$#$
'If tabl Is Nothing Then Exit Function
If tabl.BOF Then
    If c = "W" Then
        For I = 2 To maxi: val(I) = 0: Next I
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
    Else
        val(I) = tabl.fields(I - 2)
    End If
Next I
EN1:
byErrSqlGetValues = True
EN2:
tabl.Close
End Function

' фактически это и блокировка и соовт. записей в sDMC(rez)
' nnExt=0 исп-ся для разблок-ки оставшегося резерва при частичном списании
Function docLock(Optional unLok As String = "", Optional nnExt) As Boolean
Dim str As String
If IsMissing(nnExt) Then nnExt = numExt

Set tbDocs = myOpenRecordSet("##158", "sDocs", dbOpenTable) 'dbOpenForwardOnly)
If tbDocs Is Nothing Then Exit Function

docLock = False
tbDocs.index = "PrimaryKey"
tbDocs.Seek "=", numDoc, nnExt
If tbDocs.NoMatch Then
    MsgBox "Похоже документ уже удалили", , "Error - 166"
Else
    tbDocs.Edit ' блокируем
    str = tbDocs!rowLock
    If str <> "" And str <> AUTO.cbM.Text Then
       tbDocs.Update ' снимаем блокировку
       If unLok = "" Then _
       MsgBox "Документ '" & tbDocs!numDoc & "/" & tbDocs!numExt & _
       "' временно занят другим менеджером (" & str & ")"
       GoTo EN1
    End If
    If unLok = "" Then
        tbDocs!rowLock = AUTO.cbM.Text
    Else
        tbDocs!rowLock = ""
    End If
    tbDocs.Update
    docLock = True
End If
EN1:
tbDocs.Close
End Function


Function sumInGridCol(Grid As MSFlexGrid, col As Long) As Single
Dim v, I As Integer
    
    sumInGridCol = 0
    For I = Grid.row To Grid.RowSel
        v = Grid.TextMatrix(I, col)
        If Not IsNumeric(v) Then
            v = 0
        Else
            If v < 10000000 Then
                sumInGridCol = sumInGridCol + v
            End If
        End If
        
    Next I
End Function


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


Sub textBoxInGridCell(tb As TextBox, Grid As MSFlexGrid, Optional value As String = "", Optional pRow As Long = -1)
    Dim vRow As Long
    If pRow = -1 Then
        vRow = Grid.row
    Else
        vRow = pRow
    End If
    tb.Width = Grid.CellWidth + 50
'    tb.Text = Grid.TextMatrix(mousRow, mousCol)
    If value = "" Then
        tb.Text = Grid.TextMatrix(vRow, Grid.col)
    Else
        tb.Text = value
    End If
    tb.left = Grid.CellLeft + Grid.left
    If pRow = 0 Then
        tb.Top = Grid.Top
    Else
        tb.Top = Grid.CellTop + Grid.Top
    End If
    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)
    tb.Visible = True
    tb.SetFocus
    tb.ZOrder
    Grid.Enabled = False 'иначе курсор по ней бегает
End Sub

'не записыват неуникальное значение, для полей, где такие
'значения запрещены. А  генерит при этом error?
Function ValueToTableField(myErrCod As String, value As String, table As String, _
field As String, Optional by As String = "") As Boolean
Dim sql As String, byStr As String  ', numOrd As String

ValueToTableField = False
If value = "" Then value = "''"
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
    byStr = "." & by
    'Exit Function
End If
sql = "UPDATE " & table & " SET " & table & "." & field & _
" = " & value & " WHERE (((" & table & byStr & " ));"
AA:
'MsgBox "sql = " & sql
'Debug.Print sql
If myExecute(myErrCod, sql) = 0 Then ValueToTableField = True
End Function

Public Function vo_deleteNomnom(nomnom As String, numDoc As String) As Boolean
    vo_deleteNomnom = False
    sql = " delete from sDmcVenture where " _
        & " nomnom = '" & nomnom & "'" _
        & " and sdv_id = " & numDoc
    If myExecute("##122.2", sql) = 0 Then
        vo_deleteNomnom = True
    End If
End Function

Function getValueFromTable(tabl As String, field As String, where As String) As Variant
Dim table As Recordset

getValueFromTable = Null
sql = "SELECT " & field & " as fff  From " & tabl & _
      " WHERE " & where & ";"
Set table = myOpenRecordSet("##59.1", sql, dbOpenForwardOnly)
If table Is Nothing Then Exit Function
If Not table.BOF Then getValueFromTable = table!fff
table.Close
End Function

Function calcKolonValue(ByVal freight As Double, ByVal marginProc As Double, ByVal kodel As Double, ByVal kolonok As Double, ByVal curentKolon As Integer)
    Dim marginRate As Double, MarginValue As Double, maxUstupka As Double, stepUstupka As Double
    
    marginRate = marginProc / 100
    MarginValue = freight * marginRate / (1 - marginRate)
    maxUstupka = MarginValue * (1 - kodel)
    If kolonok > 1 Then
        stepUstupka = maxUstupka / (kolonok - 1)
    Else
        stepUstupka = 0
    End If
    
    calcKolonValue = freight + MarginValue - stepUstupka * (curentKolon - 1)
    
End Function

Function excelCombiSchapka(ByRef objExel, ByVal RubRate As Double, ByVal mainTitle As String, ByVal leftBound As String) As Integer
    Dim I As Integer
    Const ret As Integer = 16
    excelCombiSchapka = ret
    With objExel.ActiveSheet.Cells(1, 1)
        .value = Format(Now(), "dd.mm.yyyy")
        .HorizontalAlignment = xlHAlignCenter
        .Font.Bold = True
    End With
    
    With objExel.ActiveSheet.Range("D1:" & leftBound & "1")
        .Merge (True)
        .value = "ПЕТРОВСКИЕ МАСТЕРСКИЕ"
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    With objExel.ActiveSheet.Range("A2:" & leftBound & "2")
        .Merge (True)
        .value = "www.petmas.ru, e-mail: petmas@dol.ru"
        .HorizontalAlignment = xlHAlignCenter
    End With
    With objExel.ActiveSheet.Range("A3:" & leftBound & "3")
        .Merge (True)
        .value = "тел.: (495) 333-02-78, (499) 743-00-70, (499) 788-73-64; Факс: (495) 720-54-56"
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    With objExel.ActiveSheet.Range("A5:" & leftBound & "5")
        .Merge (True)
        .value = "ОБЗОРНЫЙ ПРАЙС-ЛИСТ ДЛЯ РАБОТЫ  ПО КАТАЛОГУ ""ПЕТРОВСКИХ МАСТЕРСКИХ"""
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    
    With objExel.ActiveSheet.Range("A7:" & leftBound & "7")
        .Merge (True)
        .value = mainTitle
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
    End With

    For I = 0 To 2
        With objExel.ActiveSheet.Cells(9 + I, 1)
            .value = "Колонка"
        End With
        With objExel.ActiveSheet.Cells(9 + I, 2)
            .value = ChrB(Asc("A") + I)
            .Font.Bold = True
        End With
        With objExel.ActiveSheet.Cells(9 + I, 3)
            .value = Choose(I + 1, _
                "Цена готового изделия с нанесением, в Москве - для конечного клиента", _
                "Цена готового изделия с нанесением - для РА", _
                "Суммарная цена комплектующих")
        End With
    Next I
    
    With objExel.ActiveSheet.Range("A13:" & leftBound & "13")
        .Merge (True)
        .value = "Для запроса по наличию комплектующих достаточно указать код изделия и количество"
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    With objExel.ActiveSheet.Range("A" & CStr(ret - 1) & ":" & leftBound & CStr(ret - 1))
        .Merge (True)
        If RubRate = 1 Then
            .value = "Цены указаны в у.е., исчисляются в USD по курсу ЦБ и включают НДС"
        Else
            .value = "Цены указаны рублях и включают НДС"
        End If
        .HorizontalAlignment = xlHAlignRight
    End With
    
End Function

Function excelStdSchapka(ByRef objExel, ByVal RubRate As Double, ByVal mainTitle As String, ByVal leftBound As String) As Integer

    excelStdSchapka = 6
    With objExel.ActiveSheet.Cells(1, 1)
        .value = Format(Now(), "dd.mm.yyyy")
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    With objExel.ActiveSheet.Range("E1:" & leftBound & "1")
        .Merge (True)
        .value = "ПЕТРОВСКИЕ МАСТЕРСКИЕ"
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    With objExel.ActiveSheet.Range("A2:" & leftBound & "2")
        .Merge (True)
        .value = mainTitle
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
    End With
    With objExel.ActiveSheet.Range("A3:" & leftBound & "3")
        .Merge (True)
        .value = "www.petmas.ru, e-mail: petmas@dol.ru"
        .HorizontalAlignment = xlHAlignCenter
    End With
    With objExel.ActiveSheet.Range("A4:" & leftBound & "4")
        .Merge (True)
        .value = "тел.: (495) 333-02-78, (499) 743-00-70, (499) 788-73-64; Факс: (495) 720-54-56"
        .HorizontalAlignment = xlHAlignCenter
    End With
    With objExel.ActiveSheet.Range("A5:" & leftBound & "5")
        .Merge (True)
        If RubRate = 1 Then
            .value = "Цены указаны в у.е., исчисляются в USD по курсу ЦБ и включают НДС"
        Else
            .value = "Цены указаны рублях и включают НДС"
        End If
        .HorizontalAlignment = xlHAlignRight
        .Font.Bold = True
    End With
End Function

Sub PriceToExcel(Regim As String, curRate As Double, mainReportTitle As String, kegl As Integer, Optional prodCategoryId As Integer = 1)
Dim I As Integer, findId As Integer, str As String

' столбец - последний. В зависимости от режима - разный
Dim lastCol As String, lastColInt As Integer

'Из Спарвочника Готовых изделий получаем Список Id всех групп(серий),
'в которых есть хотя бы одно изделие.
sql = "SELECT DISTINCT prSeriaId from sGuideProducts"
If prodCategoryId <> 0 Then
    sql = sql & " where prodCategoryId = " & prodCategoryId
End If

Set tbProduct = myOpenRecordSet("##412", sql, dbOpenDynaset)
If tbProduct Is Nothing Then Exit Sub

ReDim NN(0): I = 0
While Not tbProduct.EOF
    I = I + 1
    ReDim Preserve NN(I): NN(I) = Format(tbProduct!prSeriaId, "0000")
    findId = tbProduct!prSeriaId

AA: ' tbGuide.Seek "=", findId
'    If tbGuide.NoMatch Then msgOfEnd ("##414")
    sql = "SELECT seriaName, parentSeriaId from sGuideSeries " & _
    "WHERE seriaId = " & findId
    If Not byErrSqlGetValues("##414", sql, str, findId) Then tbProduct.Close: Exit Sub
    
'    NN(i) = tbGuide!seriaName & " / " & NN(i) ' к имени добавляем Id
    NN(I) = str & " / " & NN(I) ' к имени добавляем Id
'    findId = tbGuide!parentSeriaId
    If findId > 0 Then GoTo AA 'к имени текущей группы спереди приклеиваются
                               'имена всех групп дерева, в которые она входит
    tbProduct.MoveNext
Wend

tbProduct.Close


'Этот блок не требует изменения -------------------------------------------

quickSort NN, 1

On Error GoTo ERR2

    Set objExel = New Excel.Application
    objExel.Visible = True
    objExel.SheetsInNewWorkbook = 1
    objExel.Workbooks.Add
    objExel.ActiveSheet.Cells.Font.Size = kegl
    
    If Regim = "default" Then
        lastCol = "H"
        ' печать стандартной шапки
        exRow = excelStdSchapka(objExel, curRate, mainReportTitle, lastCol)
    ElseIf Regim = "combi" Then
        lastCol = "G"
        exRow = excelCombiSchapka(objExel, curRate, mainReportTitle, lastCol)
    End If
    
    lastColInt = Asc(lastCol) - Asc("A") + 1
    
    objExel.ActiveSheet.Columns(1).columnWidth = 10
    objExel.ActiveSheet.Columns(2).columnWidth = 10
    objExel.ActiveSheet.Columns(3).columnWidth = 50
    objExel.ActiveSheet.Columns(4).HorizontalAlignment = xlHAlignRight
    objExel.ActiveSheet.Columns(5).HorizontalAlignment = xlHAlignRight
    objExel.ActiveSheet.Columns(6).HorizontalAlignment = xlHAlignRight
    objExel.ActiveSheet.Columns(7).HorizontalAlignment = xlHAlignRight

'------------------------------------------------------------------------

For I = 1 To UBound(NN) ' перебор всех групп
  str = NN(I)
  findId = right$(str, 4) ' извлекаем из имен группы id группы

'$comtec$  Далее ссылки на табл.sGuideProducts и на ее поля надо заменить на
'эквиваленты из базы Comtec исходя из след.соответствия с колонками
'Справ-ка Готовых изделий из программы stime:
'"Номер"    "Код"   "web"   "Описание"    Размер   "1-5"   "Стр."
'SortNom   prName    web    prDescript    prSize   Cena4    page

sql = "SELECT prId, prName, prDescript, prSize, Cena4, page, rabbat " _
    & " From sGuideProducts " _
    & " Where prSeriaId = " & findId & " AND prodCategoryId = " & prodCategoryId _
    & " ORDER BY SortNom"

  Set tbProduct = myOpenRecordSet("##415", sql, dbOpenDynaset)
  If Not tbProduct Is Nothing Then
    If Not tbProduct.BOF Then
      bilo = False
      While Not tbProduct.EOF

'Этот блок не требует изменения (здесь выдаются заголовки групп)------------
        If Not bilo Then
            bilo = True
            
            With objExel.ActiveSheet.Range("A" & exRow & ":" & lastCol & exRow)
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).Weight = xlThin
            End With
            
            str = left$(str, Len(str) - 6)
            With objExel.ActiveSheet.Cells(exRow, 2)
                .value = str
                .Font.Bold = True
            End With
            objExel.ActiveSheet.Cells(exRow, lastColInt).Borders(xlEdgeRight).Weight = xlMedium
            
            exRow = exRow + 1
            objExel.ActiveSheet.Range("A" & exRow & ":" & lastCol & exRow). _
            Borders(xlEdgeBottom).Weight = xlThin
            
            objExel.ActiveSheet.Cells(exRow, 1).value = "Код"
            objExel.ActiveSheet.Cells(exRow, 2).value = "Размер[см]"
            objExel.ActiveSheet.Cells(exRow, 3).value = "Описание"
            
            gain2 = 0
            gSeriaId = findId
            
            If Regim = "default" Then
                If getGainAndHead Then
                    objExel.ActiveSheet.Cells(exRow, 4).value = " " & head1
                    objExel.ActiveSheet.Cells(exRow, 5).value = " " & head2
                    objExel.ActiveSheet.Cells(exRow, 6).value = " " & head3
                    objExel.ActiveSheet.Cells(exRow, 7).value = " " & head4
                End If
            ElseIf Regim = "combi" Then
                Dim J As Integer
                For J = 0 To 2
                    With objExel.ActiveSheet.Cells(exRow, 4 + J)
                        .value = ChrB(Asc("A") + J)
                    End With
                Next J
            End If
            
            objExel.ActiveSheet.Cells(exRow, lastColInt).value = "    стр."
            cErr = setVertBorders(objExel, xlThin, lastColInt)
            exRow = exRow + 1
        End If

'---------------------------------------------------------------------------
'Далее выдаются параметры по каждому изделию группы
        
        objExel.ActiveSheet.Cells(exRow, 1).value = tbProduct!prName
        objExel.ActiveSheet.Cells(exRow, 2).value = tbProduct!prSize
        objExel.ActiveSheet.Cells(exRow, 3).value = tbProduct!prDescript
        objExel.ActiveSheet.Cells(exRow, 4).value = Chr(160) & Format(tbProduct!Cena4 * curRate, "0.00")
        
        If Regim = "default" Then
            If gain4 > 0 Then
                objExel.ActiveSheet.Cells(exRow, 5).value = Chr(160) & Format(Round(tbProduct!Cena4 * curRate * gain2, 1), "0.00")
                objExel.ActiveSheet.Cells(exRow, 6).value = Chr(160) & Format(Round(tbProduct!Cena4 * curRate * gain3, 1), "0.00")
                objExel.ActiveSheet.Cells(exRow, 7).value = Chr(160) & Format(Round(tbProduct!Cena4 * curRate * gain4, 1), "0.00")
            End If
        ElseIf Regim = "combi" Then
            gain2 = getRabbat(tbProduct!Cena4, tbProduct!rabbat)
            objExel.ActiveSheet.Cells(exRow, 5).value = Chr(160) & Format(Round(curRate * gain2, 1), "0.00")
            gain3 = getCenaSale(tbProduct!prId)
            objExel.ActiveSheet.Cells(exRow, 6).value = Chr(160) & Format(Round(curRate * gain3, 1), "0.00")
        End If
        
        
        objExel.ActiveSheet.Cells(exRow, lastColInt).value = " " & tbProduct!Page
        cErr = setVertBorders(objExel, xlThin, lastColInt)
        If cErr <> 0 Then GoTo ERR2
        exRow = exRow + 1:
    
        tbProduct.MoveNext
      Wend
    End If
    tbProduct.Close
  End If
Next I
With objExel.ActiveSheet.Range("A" & exRow & ":" & lastCol & exRow)
    .Borders(xlEdgeTop).Weight = xlMedium
End With

Set objExel = Nothing
Exit Sub

ERR2:
If cErr <> 424 And Err <> 424 Then  ' 424 - не дождались конца вывода закрыли док-т
    MsgBox Error, , "Ошибка 421-" & cErr '##421
End If
Set objExel = Nothing

End Sub

Function getRabbat(cenaProd As Double, rabbat As Integer) As Double
    getRabbat = (1 - rabbat / 100) * cenaProd
End Function

Function getCenaSale(productId As Integer) As Double
    Dim ret As String
    ret = getSumCena(productId, "sale")
    If IsNumeric(ret) Then
        getCenaSale = CDbl(ret)
    Else
        getCenaSale = 0
    End If
End Function

Function getGainAndHead() As Boolean
getGainAndHead = False
sql = "SELECT head1, head2, head3, head4, gain2, gain3, gain4 " & _
"from sGuideSeries WHERE (((sGuideSeries.seriaId)=" & gSeriaId & "));"
If Not byErrSqlGetValues("##416", sql, head1, head2, head3, head4, gain2, _
gain3, gain4) Then Exit Function
getGainAndHead = True
End Function


Function setVertBorders(ByRef objExel, lineWeight As Long, Optional lastCol = 8) As Integer
Dim I As Integer
    For I = 1 To lastCol
        If I < lastCol Then
            objExel.ActiveSheet.Cells(exRow, I).Borders(xlEdgeRight).Weight = lineWeight
        Else
            objExel.ActiveSheet.Cells(exRow, I).Borders(xlEdgeRight).Weight = xlMedium
        End If
    Next I
End Function

'reg = "" => SumCenaFreight
'reg <> "" => SumCenaSale
Function getSumCena(productId As Integer, Optional reg As String = "") As String
Dim sum As Single, v, s As Single, prevGroup As String, max As Single

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
        If tbNomenk!formula = "" Then
            getSumCena = "Error: Не определена формула для номенклатуры '" & tbNomenk!nomnom & "'"
            GoTo er
        End If
        v = nomenkFormula("noOpen") 'Цена2
    Else
        v = tbNomenk!CENA_W
    End If
        
    If IsNumeric(v) Then
        s = v * tbNomenk!quantity / tbNomenk!perList
        If tbNomenk!xgroup = "" Then
            sum = sum + s
            prevGroup = tbNomenk!xgroup
        ElseIf prevGroup = tbNomenk!xgroup Then
            If max < s Then max = s
        Else
            If prevGroup <> "" Then sum = sum + max
            max = s
            prevGroup = tbNomenk!xgroup
        End If
    Else
        getSumCena = v & " в формуле для  '" & tbNomenk!nomnom & "'"
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


Public Sub initProdCategoryBox(ByRef lbPrWeb, Optional extended As Boolean = False)
Dim table As Recordset
Dim name As String

    ' Сначала удаляем старые значения
    While lbPrWeb.ListCount
        lbPrWeb.RemoveItem (0)
    Wend
    
    sql = "select * from GuideProdCategory"
    
    Set table = myOpenRecordSet("##72", sql, dbOpenForwardOnly)
    If table Is Nothing Then myBase.Close: End
    
    lbPrWeb.AddItem "", 0
    While Not table.EOF
        If extended Then
            name = table!nameRu & " (" & table!sysname & ")"
        Else
            name = table!sysname
        End If
        lbPrWeb.AddItem table!sysname
        lbPrWeb.ItemData(lbPrWeb.ListCount - 1) = table!prodCategoryId
        table.MoveNext
    Wend
    table.Close
    
    If TypeOf lbPrWeb Is ListBox Then
        lbPrWeb.Height = 225 * lbPrWeb.ListCount
    End If

End Sub

