Attribute VB_Name = "Common"
Option Explicit

Private Const dhcMissing = -2 'нужна для quickSort
Public objExel As Excel.Application, exRow As Long
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
Public Table As Recordset
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

Public gSrokPostav As Single ' переменная испольуется для определения,нужно ли пересчитывать запасы и к заявке, если срок доставки изменился
Public gSrokBetweenPostav As Single ' переменная испольуется для определения,нужно ли пересчитывать запасы и к заявке, если срок между поставками изменился

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
            strLastTwoDigit = Right(strRate, 2)
            digit = CInt(strLastTwoDigit)
            If digit >= 5 And digit <= 20 Then
                rubleSuffix = "ей"
            End If
        End If
        If rubleSuffix = "" Then
            strLastDigit = Right(strRate, 1)
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


Function dateBasic2Sybase(aDay As String)
Dim dt_str As String

dt_str = "20" & Right$(aDay, 2) & "-" & Mid$(aDay, 4, 2) & "-" & Left$(aDay, 2)
dateBasic2Sybase = CDate(dt_str)

End Function


Function dateSybase2Basic(aDay As String)
Dim dt_str As String

dt_str = Left(aDay, 4) & "-" & Mid(aDay, 5, 2) & "-" & Right(aDay, 2)
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
Dim IL As Long
    noClick = True
    Grid.row = row
    For IL = 0 To Grid.Cols - 1
      If IL > 0 Then
        Grid.col = IL
        Grid.CellBackColor = Grid.BackColor
        Grid.CellForeColor = Grid.ForeColor
      End If
      Grid.TextMatrix(row, IL) = ""
    Next IL
    Grid.col = 1
    noClick = False
End Sub

Sub colorGridRow(Grid As MSFlexGrid, row As Long, color As Long)
Dim IL As Long
    noClick = True
    Grid.row = row
    For IL = 0 To Grid.Cols - 1
        Grid.col = IL
        If IL > 0 Then Grid.CellBackColor = color
    Next IL
    Grid.col = 1
    noClick = False
End Sub

Sub foreColorGridRow(Grid As MSFlexGrid, row As Long, color As Long, ccol As Long)
Dim IL As Long
    noClick = True
    Grid.row = row
    For IL = 0 To Grid.Cols - 1
        Grid.col = IL
        If IL > 0 Then Grid.CellForeColor = color
    Next IL
    Grid.col = ccol
    noClick = False
End Sub

Sub colorGridCell(Grid As MSFlexGrid, ByVal row As Long, ByVal col As Long, ByVal color As Long)
Dim IL As Long
    noClick = True
    Grid.row = row
    Grid.col = col
    Grid.CellForeColor = color
    noClick = False
End Sub



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


Function findExValInCol(Grid As MSFlexGrid, Value As String, _
            col As Integer, Optional pos As Long = -1) As Long
Dim IL As Long, str  As String, beg As Long

If pos < 1 Then
    beg = 1
Else
    beg = pos
End If
Value = UCase(Value)
For IL = beg To Grid.Rows - 1
    str = UCase(Grid.TextMatrix(IL, col))
    If InStr(str, Value) > 0 Then
        Grid.TopRow = IL
        Grid.row = IL
        findExValInCol = IL
        Exit Function
    End If
Next IL
findExValInCol = -1

End Function

Sub listBoxInGridCell(lb As ListBox, Grid As MSFlexGrid, Optional sel As String = "")
Dim I As Integer, L As Long
    If lb.ListCount < 200 Then
        L = CLng(195) * CLng(lb.ListCount) + 100 ' Для больших списков
    Else
        L = Grid.Height / 2 + 500
    End If
        
    If L > Grid.Height / 2 + 500 Then L = Grid.Height / 2 + 500
    lb.Height = L
    
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

gSrokPostav = getEffectiveSetting("gSrokPostav")
gSrokBetweenPostav = getEffectiveSetting("gSrokBetweenPostav")


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

Function ValueToGuideSourceField(myErrCod As String, Value As String, _
Field As String, Optional passErr As Integer = -11111) As Integer
Dim I As Integer

ValueToGuideSourceField = False
sql = "UPDATE sGuideSource SET [" & Field & _
"] = '" & Value & "' WHERE (((sourceId)=" & gSourceId & "));"
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
  tbDMC.Index = "nomDoc"
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
    tbNomenk.Index = "PrimaryKey"
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
    numDoc = Left$(nom, I - 1)
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


Function valueToSystemField(val As Variant, Field As String) As Boolean
valueToSystemField = False
Set tbSystem = myOpenRecordSet("##148", "System", dbOpenTable)
If tbSystem Is Nothing Then Exit Function
tbSystem.Edit
tbSystem.Fields(Field) = val
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
Function getWhereByDateBoxes(Frm As Form, dateField As String, _
begDate As Date, Optional reg As String = "") As String

Dim str As String, ckStart As Boolean, ckEnd  As Boolean

getWhereByDateBoxes = "": str = "":

ckStart = False: ckEnd = False
On Error Resume Next ' на случай, если в этой форме у дат нет флажков
If Frm.ckEndDate.Value > 0 Then ckEnd = True  'то они как бы установлены
If Frm.ckStartDate.Value > 0 And Frm.ckStartDate.Visible Then ckStart = True
On Error GoTo 0

If ckStart Then
    If Not isDateTbox(Frm.tbStartDate) Then GoTo ERRd  'tmpDate
End If
If reg = "" Then ' если период Между
    If DateDiff("d", begDate, tmpDate) > 0 And ckStart Then _
        str = "(" & dateField & ") >='" & Format(tmpDate, "yyyy-mm-dd") & "'"
    If ckEnd Then
      If Not isDateTbox(Frm.tbEndDate) Then GoTo ERRd
      If ckStart Then
        If DateDiff("d", Frm.tbStartDate.Text, tmpDate) < 0 Then
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
Dim IL As Long
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
str = CStr(val(0)): c = Left$(str, 1)
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
    And IsNull(tabl.Fields(I - 2)) Then
        val(I) = 0
    ElseIf str = "String" And IsNull(tabl.Fields(I - 2)) Then
        val(I) = ""
    Else
        val(I) = tabl.Fields(I - 2)
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
tbDocs.Index = "PrimaryKey"
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
Dim V, I As Integer
    
    sumInGridCol = 0
    For I = Grid.row To Grid.RowSel
        V = Grid.TextMatrix(I, col)
        If Not IsNumeric(V) Then
            V = 0
        Else
            If V < 10000000 Then
                sumInGridCol = sumInGridCol + V
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


Sub textBoxInGridCell(tb As TextBox, Grid As MSFlexGrid, Optional Value As String = "", Optional pRow As Long = -1)
    Dim vRow As Long
    If pRow = -1 Then
        vRow = Grid.row
    Else
        vRow = pRow
    End If
    tb.Width = Grid.CellWidth + 50
'    tb.Text = Grid.TextMatrix(mousRow, mousCol)
    If Value = "" Then
        tb.Text = Grid.TextMatrix(vRow, Grid.col)
    Else
        tb.Text = Value
    End If
    tb.Left = Grid.CellLeft + Grid.Left
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
Function ValueToTableField(myErrCod As String, Value As String, Table As String, _
Field As String, Optional by As String = "") As Boolean
Dim sql As String, byStr As String  ', numOrd As String

ValueToTableField = False
If Value = "" Then Value = "''"
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
    sql = "UPDATE " & Table & " SET " & Table & "." & Field & "=" & Value _
        & " WHERE (((" & Table & ".numDoc)=" & numDoc & " AND (" & Table & _
        ".numExt)=" & numExt & " ));"
    GoTo AA
Else
    byStr = "." & by
    'Exit Function
End If
sql = "UPDATE " & Table & " SET " & Table & "." & Field & _
" = " & Value & " WHERE (((" & Table & byStr & " ));"
AA:
'MsgBox "sql = " & sql
'Debug.Print sql
If myExecute(myErrCod, sql) = 0 Then ValueToTableField = True
End Function

Public Function vo_deleteNomnom(Nomnom As String, numDoc As String) As Boolean
    vo_deleteNomnom = False
    sql = " delete from sDmcVenture where " _
        & " nomnom = '" & Nomnom & "'" _
        & " and sdv_id = " & numDoc
    If myExecute("##122.2", sql) = 0 Then
        vo_deleteNomnom = True
    End If
End Function

Function getValueFromTable(tabl As String, Field As String, where As String) As Variant
Dim Table As Recordset

getValueFromTable = Null
sql = "SELECT " & Field & " as fff  From " & tabl & _
      " WHERE " & where & ";"
Set Table = myOpenRecordSet("##59.1", sql, dbOpenForwardOnly)
If Table Is Nothing Then Exit Function
If Not Table.BOF Then getValueFromTable = Table!fff
Table.Close
End Function

Function excelDealerSchapka(ByRef objExel, ByVal RubRate As Double, ByVal mainTitle As String, ByVal leftBound As String) As Integer
    Dim I As Integer
    Const ret As Integer = 16
    excelDealerSchapka = ret
    With objExel.ActiveSheet.Cells(1, 1)
        .Value = Format(Now(), "dd.mm.yyyy")
        .HorizontalAlignment = xlHAlignCenter
        .Font.Bold = True
    End With
    
    With objExel.ActiveSheet.Range("D1:" & leftBound & "1")
        .Merge (True)
        .Value = "ПЕТРОВСКИЕ МАСТЕРСКИЕ"
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    With objExel.ActiveSheet.Range("A2:" & leftBound & "2")
        .Merge (True)
        .Value = "www.petmas.ru, e-mail: petmas@dol.ru"
        .HorizontalAlignment = xlHAlignCenter
    End With
    With objExel.ActiveSheet.Range("A3:" & leftBound & "3")
        .Merge (True)
        .Value = "тел.: (495) 333-02-78, (499) 743-00-70, (499) 788-73-64; Факс: (495) 720-54-56"
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    With objExel.ActiveSheet.Range("A5:" & leftBound & "5")
        .Merge (True)
        .Value = "ОБЗОРНЫЙ ПРАЙС-ЛИСТ ДЛЯ РАБОТЫ  ПО КАТАЛОГУ ""ПЕТРОВСКИХ МАСТЕРСКИХ"""
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    
    With objExel.ActiveSheet.Range("A7:" & leftBound & "7")
        .Merge (True)
        .Value = mainTitle
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
    End With

    For I = 0 To 2
        With objExel.ActiveSheet.Cells(9 + I, 1)
            .Value = "Колонка"
        End With
        With objExel.ActiveSheet.Cells(9 + I, 2)
            .Value = ChrB(Asc("A") + I)
            .Font.Bold = True
        End With
        With objExel.ActiveSheet.Cells(9 + I, 3)
            .Value = Choose(I + 1, _
                "Цена готового изделия с нанесением, в Москве - для конечного клиента", _
                "Цена готового изделия с нанесением - для РА", _
                "Суммарная цена комплектующих")
        End With
    Next I
    
    With objExel.ActiveSheet.Range("A13:" & leftBound & "13")
        .Merge (True)
        .Value = "Для запроса по наличию комплектующих достаточно указать код изделия и количество"
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    With objExel.ActiveSheet.Range("A" & CStr(ret - 1) & ":" & leftBound & CStr(ret - 1))
        .Merge (True)
        If RubRate = 1 Then
            .Value = "Цены указаны в у.е., исчисляются в USD по курсу ЦБ и включают НДС"
        Else
            .Value = "Цены указаны рублях и включают НДС"
        End If
        .HorizontalAlignment = xlHAlignRight
    End With
    
End Function


Function excelStdSchapka(ByRef objExel, ByVal RubRate As Double, ByVal mainTitle As String _
    , ByVal leftBound As String, Optional ventureName As String = "ПЕТРОВСКИЕ МАСТЕРСКИЕ", Optional paymentType = False _
    , Optional contact1 As String = "", Optional contact2 As String = "") As Integer

    Const titleUE = "цены указаны в у.е. и исчисляются в USD по курсу ЦБ (НДС не облагается)"
    Const titleRub = "цены указаны в рублях (НДС не облагается)"
    
    excelStdSchapka = 7
    With objExel.ActiveSheet.Cells(1, 1)
        .Value = Format(Now(), "dd.mm.yyyy")
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    
    With objExel.ActiveSheet.Range("B1:" & leftBound & "1")
        .Merge (True)
        .Value = ventureName
        .HorizontalAlignment = xlHAlignRight
    End With
    
    With objExel.ActiveSheet.Range("A2:" & leftBound & "2")
        .Merge (True)
        .Value = mainTitle
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
    End With
    With objExel.ActiveSheet.Range("A3:" & leftBound & "3")
        .Merge (True)
        .Value = contact1
        .HorizontalAlignment = xlHAlignCenter
    End With
    With objExel.ActiveSheet.Range("A4:" & leftBound & "4")
        .Merge (True)
        .Value = contact2
        .HorizontalAlignment = xlHAlignCenter
    End With
    If paymentType Then
        With objExel.ActiveSheet.Range("A5:" & leftBound & "5")
            .Merge (True)
            If RubRate = 1 Then
                .Value = titleUE
            Else
                .Value = titleRub
            End If
            .HorizontalAlignment = xlHAlignRight
            .Font.Bold = True
        End With
    End If
End Function

Function excelSchapkaBAClient(ByRef objExel, ByVal RubRate As Double, ByVal mainTitle As String _
    , ByVal leftBound As String, Optional paymentType = False) As Integer

    Const titleUE = "Цены на готовые изделия с нанесением указаны в у.е., исчисляются в USD по курсу ЦБ"
    Const titleRub = "Цены на готовые изделия с нанесением указаны в рублях"
    
    excelSchapkaBAClient = 7
    With objExel.ActiveSheet.Cells(1, 1)
        .Value = Format(Now(), "dd.mm.yyyy")
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    
    With objExel.ActiveSheet.Range("A2:" & leftBound & "2")
        .Merge (True)
        .Value = mainTitle
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    With objExel.ActiveSheet.Range("A3:" & leftBound & "3")
        .Merge (True)
        .Value = "КОРПОРАТИВНЫЕ ПРИЗЫ И НАГРАДЫ"
        .HorizontalAlignment = xlHAlignCenter
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    If paymentType Then
        With objExel.ActiveSheet.Range("A5:" & leftBound & "5")
            .Merge (True)
            If RubRate = 1 Then
                .Value = titleUE
            Else
                .Value = titleRub
            End If
            .HorizontalAlignment = xlHAlignCenter
        End With
    End If
End Function








Sub PriceToExcel(Regim As String, curRate As Double, mainReportTitle As String, kegl As Integer, Optional prodCategoryId As Integer = 1, Optional commonRabbat As Single = 1 _
        , Optional contact1 As String = "", Optional contact2 As String = "")
        
Dim I As Integer, findId As Integer, str As String
Dim venture As String

Const venturePM = "Петровские Мастерские"
Const ventureMM = "Маркмастер"
' столбец - последний. В зависимости от режима - разный
Dim lastCol As String, lastColInt As Integer
Dim RPF_Rate As Single
Dim paymentType As Boolean

'Из Спарвочника Готовых изделий получаем Список Id всех групп(серий),
'в которых есть хотя бы одно изделие.
sql = "SELECT DISTINCT prSeriaId from sGuideProducts"
If prodCategoryId <> 0 Then
    sql = sql & " where prodCategoryId = " & prodCategoryId
End If

If Regim = "pricePM" Then
    venture = venturePM
    paymentType = False
Else
    venture = ventureMM
    paymentType = True
End If

Set tbProduct = myOpenRecordSet("##412", sql, dbOpenDynaset)
If tbProduct Is Nothing Then Exit Sub

ReDim NN(0): I = 0
While Not tbProduct.EOF
    I = I + 1
    ReDim Preserve NN(I): NN(I) = Format(tbProduct!prSeriaId, "0000")
    findId = tbProduct!prSeriaId

AA:
    sql = "SELECT seriaName, parentSeriaId from sGuideSeries " & _
    "WHERE seriaId = " & findId
    If Not byErrSqlGetValues("##414", sql, str, findId) Then tbProduct.Close: Exit Sub
    
    NN(I) = str & " / " & NN(I) ' к имени добавляем Id
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
With objExel.ActiveSheet
        .Cells.Font.Size = kegl
    
    lastCol = "H"
    
    exRow = excelStdSchapka(objExel, curRate, mainReportTitle, lastCol, venture, paymentType, contact1, contact2)
    
    lastColInt = Asc(lastCol) - Asc("A") + 1
    
        .Columns(1).columnWidth = 10
        .Columns(2).columnWidth = 10
        .Columns(3).columnWidth = 50
        .Columns(4).HorizontalAlignment = xlHAlignRight
        .Columns(5).HorizontalAlignment = xlHAlignRight
        .Columns(6).HorizontalAlignment = xlHAlignRight
        .Columns(7).HorizontalAlignment = xlHAlignRight

'------------------------------------------------------------------------

    For I = 1 To UBound(NN) ' перебор всех групп
      str = NN(I)
      findId = Right$(str, 4) ' извлекаем из имен группы id группы
    
    '$comtec$  Далее ссылки на табл.sGuideProducts и на ее поля надо заменить на
    'эквиваленты из базы Comtec исходя из след.соответствия с колонками
    'Справ-ка Готовых изделий из программы stime:
    '"Номер"    "Код"   "web"   "Описание"    Размер   "1-5"   "Стр."
    'SortNom   prName    web    prDescript    prSize   Cena4    page
    
    sql = "SELECT p.prId, p.prName, p.prDescript, p.prSize, p.Cena4, p.page, p.rabbat as productRabbat, f.formula, w.prId as hasWeb " _
        & vbCr & " From sGuideProducts p " _
        & vbCr & " left JOIN sGuideFormuls f on f.nomer = p.formulaNom" _
        & vbCr & " left JOIN wf_izdeliaWithWeb w on w.prId  = p.prId" _
        & vbCr & " left join wf_izdeliaVmtOnly vmt on vmt.productid = p.prId" _
        & vbCr & " Where p.prSeriaId = " & findId _
        & " AND p.prodCategoryId = " & prodCategoryId
    
    'Debug.Print sql
    If Regim = "dealer" Then
        sql = sql & " AND vmt.productId is null AND w.prId is null"
    End If
    
    If Not Regim = "pricePM" Then
        sql = sql & " AND isnumeric(p.page) = 1"
    End If
    
    sql = sql & " ORDER BY p.SortNom"
    
    'Debug.Print sql
    
      Set tbProduct = myOpenRecordSet("##415", sql, dbOpenDynaset)
      If Not tbProduct Is Nothing Then
        If Not tbProduct.BOF Then
          bilo = False
          While Not tbProduct.EOF
    
    'Этот блок не требует изменения (здесь выдаются заголовки групп)------------
            If Not bilo Then
                bilo = True
                
                With .Range("A" & exRow & ":" & lastCol & exRow)
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeBottom).Weight = xlThin
                End With
                
                str = Left$(str, Len(str) - 6)
                With .Cells(exRow, 2)
                    .Value = str
                    .Font.Bold = True
                End With
                .Cells(exRow, lastColInt).Borders(xlEdgeRight).Weight = xlMedium
                
                exRow = exRow + 1
                .Range("A" & exRow & ":" & lastCol & exRow). _
                Borders(xlEdgeBottom).Weight = xlThin
                
                .Cells(exRow, 1).Value = "Код"
                .Cells(exRow, 2).Value = "Размер[см]"
                .Cells(exRow, 3).Value = "Описание"
                
                gain2 = 0
                gSeriaId = findId
                
                If getGainAndHead Then
                    .Cells(exRow, 4).Value = " " & head1
                    .Cells(exRow, 5).Value = " " & head2
                    .Cells(exRow, 6).Value = " " & head3
                    .Cells(exRow, 7).Value = " " & head4
                End If
                
                .Cells(exRow, lastColInt).Value = "    стр."
                cErr = setVertBorders(objExel, xlThin, lastColInt)
                exRow = exRow + 1
            End If
    
    '---------------------------------------------------------------------------
    'Далее выдаются параметры по каждому изделию группы
            
            .Cells(exRow, 1).Value = tbProduct!prName
            .Cells(exRow, 2).Value = tbProduct!prSize
            .Cells(exRow, 3).Value = tbProduct!prDescript
            Dim res As String
            
            res = ExcelProductPrices(RPF_Rate, Regim, curRate, exRow, 4, commonRabbat)
            If res <> "" Then
                MsgBox "Ошибка при вычислении цены" & vbCr & "Текст ошибки: " & res, , "Изделие " & tbProduct!prName
                tbProduct.Close
                GoTo fail
            End If
             
            
            .Cells(exRow, lastColInt).Value = " " & tbProduct!Page
            cErr = setVertBorders(objExel, xlThin, lastColInt)
            If cErr <> 0 Then GoTo ERR2
            exRow = exRow + 1:
        
            tbProduct.MoveNext
          Wend
        End If
        tbProduct.Close
      End If
    Next I
    With .Range("A" & exRow & ":" & lastCol & exRow)
        .Borders(xlEdgeTop).Weight = xlMedium
    End With

End With

fail:
Set objExel = Nothing
Exit Sub

ERR2:
If cErr <> "424" And Err <> 424 Then  ' 424 - не дождались конца вывода закрыли док-т
    MsgBox Error, , "Ошибка 421-" & cErr '##421
End If
Set objExel = Nothing

End Sub

Public Function ExcelProductPrices(ByRef RPF_Rate As Single, Regim As String, curRate As Double, exRow As Long, exCol As Long, Optional commonRabbat As Single) As String

    
    Dim baseCena As String
    Dim rbt As Single
    Dim hasWeb As Boolean
    Dim res As String
        
    gProductId = tbProduct!prId
    
    If Not IsNull(tbProduct!hasWeb) Then
        res = res
    End If
    res = calcBaseCenaAndRpfRate(Regim, baseCena, RPF_Rate, tbProduct!Cena4, tbProduct!productRabbat, commonRabbat, tbProduct!hasWeb)
    
    If res <> "" Then
        ExcelProductPrices = res
        Exit Function
        
    End If
    
    With objExel.ActiveSheet
        .Cells(exRow, exCol).Value = Chr(160) & Format(baseCena * curRate, "0.00")
        If gain4 > 0 Then
            .Cells(exRow, exCol + 1).Value = Chr(160) & Format(Round(baseCena * curRate * gain2, 1), "0.00")
            .Cells(exRow, exCol + 2).Value = Chr(160) & Format(Round(baseCena * curRate * gain3, 1), "0.00")
            .Cells(exRow, exCol + 3).Value = Chr(160) & Format(Round(baseCena * curRate * gain4, 1), "0.00")
        End If
    End With
        

End Function



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
"from sGuideSeries WHERE sGuideSeries.seriaId = " & gSeriaId
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




