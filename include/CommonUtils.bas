Attribute VB_Name = "CommonUtils"
Option Explicit

Type VersionInfo
    path As String
    maj As Long
    min As Long
    rev As Long
    bld As Long
End Type


Type OrderNumInfo
    Numorder As Integer
    oDate As Date
    num As Integer
    ext As Integer
End Type

Sub fatalError(msg As String)
    MsgBox msg & vbCr & "Обратитесь к администратору", vbCritical, "Критическая ошибка"
    End
End Sub

Sub getAppInfo(ByRef version As VersionInfo)
    version.path = App.path & "\" & App.exeName & ".exe"
    version.maj = App.Major
    version.min = App.Minor
    version.bld = App.Revision
    version.rev = 0
End Sub

Function compareVersion(left As VersionInfo, Right As VersionInfo) As Integer
Dim lessThen As Boolean
Dim greateThen As Boolean

    If left.maj < Right.maj Then
        lessThen = True
    ElseIf left.maj > Right.maj Then
        greateThen = True
    End If
    
    If Not lessThen And Not greateThen Then
        If left.min < Right.min Then
            lessThen = True
        ElseIf left.min > Right.min Then
            greateThen = True
        End If
    End If

    If Not lessThen And Not greateThen Then
        If left.rev < Right.rev Then
            lessThen = True
        ElseIf left.rev > Right.rev Then
            greateThen = True
        End If
    End If

    If Not lessThen And Not greateThen Then
        If left.bld < Right.bld Then
            lessThen = True
        ElseIf left.bld > Right.bld Then
            greateThen = True
        End If
    End If
    
    If lessThen Then
        compareVersion = -1
        Exit Function
    End If
    If greateThen Then
        compareVersion = 1
        Exit Function
    End If
    compareVersion = 0

End Function

Function infoToString(info As VersionInfo) As String
    infoToString = info.maj & "." & info.min & "." & info.bld
End Function


Function getMainTitle() As String
Dim version As VersionInfo

    getAppInfo version
    getMainTitle = " [версия " & infoToString(version) & "]"

End Function

Function existsInTreeview(ByRef tTree As TreeView, Key As String) As Boolean
Dim I As Integer
    For I = 1 To tTree.Nodes.count
        If tTree.Nodes(I).Key = Key Then
            existsInTreeview = True
            Exit Function
        End If
    Next I
    existsInTreeview = False
End Function



Function newNumorder(value As String) As Numorder
    Dim ret As Numorder
    Set ret = New Numorder
    ret.val = value
    Set newNumorder = ret
End Function


Function getNextDocNum() As Long
Dim valueorder As Numorder

    Set valueorder = New Numorder
    valueorder.val = getSystemField("lastDocNum")
    If valueorder.isEmpty Then
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


Sub nextDay()  'возможен прыжок на неск дней
Dim I As Integer, str As String, str1 As String, j As Integer, s As Double
Dim ch As String
'MsgBox "переход на новую дату"

'wrkDefault.BeginTrans

sql = "DELETE from OrdersInCeh WHERE (((Stat)='готов'));"
If myExecute("##63", sql, 0) > 0 Then GoTo ER1

'Set tbCeh = myOpenRecordSet("##63", "OrdersInCeh", dbOpenTable) 'dbOpenForwardOnly)
'If Not tbCeh Is Nothing Then
'  If Not tbCeh.BOF Then
'    tbCeh.MoveFirst
'    While Not tbCeh.EOF
'        If tbCeh!stat = "готов" Then
'            tbCeh.Delete
'        End If
'        tbCeh.MoveNext
'    Wend
'  End If
'  tbCeh.Close
'End If


sql = "UPDATE Orders INNER JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " & _
"SET Orders.DateRS = '" & Format(curDate, "yyyy-mm-dd 10:00:00") & _
"' WHERE (((Orders.DateRS) < '" & Format(curDate, "yyyy-mm-dd 00:00:00") & _
"' And Not (Orders.DateRS) Is Null));"
'MsgBox sql
If myExecute("##11", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE Orders INNER JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " & _
"SET Orders.outDateTime = '" & Format(curDate, "yyyy-mm-dd 10:00:00") & _
"' WHERE (((Orders.outDateTime)<'" & Format(curDate, "yyyy-mm-dd 0:0:0") & "'));"
'MsgBox sql
If myExecute("##404", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE OrdersMO SET DateTimeMO = '" & Format(curDate, "yyyy-mm-dd 10:00:00") & _
"' WHERE (((DateTimeMO)<'" & Format(curDate, "yyyy-mm-dd 00:00:00") & "'));"
If myExecute("##405", sql, 0) > 0 Then GoTo ER1

''      "OrdersMO.workTimeMO, OrdersInCeh.VrVipParts, OrdersInCeh.Stat  "
'sql = "SELECT Orders.outDateTime, Orders.StatusId, OrdersMO.DateTimeMO, " & _
'      "Orders.DateRs, " & _
'      "OrdersMO.workTimeMO, OrdersInCeh.Stat  " & _
'      "FROM (Orders RIGHT JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder) " & _
'      "LEFT JOIN OrdersMO ON OrdersInCeh.numOrder = OrdersMO.numOrder;"

'Set table = myOpenRecordSet("##11", sql, dbOpenDynaset) 'dbOpenForwardOnly)
'If table Is Nothing Then myBase.Close: End
  
'If Not table.BOF Then
'  table.MoveFirst
'  While Not table.EOF
'    table.Edit
'    replaceDate table!outDateTime
'    replaceDate table!dateRS
'    replaceDate table!DateTimeMO
'    table.Update
'NXT:
'    table.MoveNext
'  Wend
'End If
'table.Close

If Not replaceResurs(1) Then GoTo ER1
If Not replaceResurs(2) Then GoTo ER1
If Not replaceResurs(3) Then GoTo ER1  '$$ceh

sql = "UPDATE System SET resursLock = '', Kurs = -Abs([Kurs]);"
If myExecute("##90", sql, 0) > 0 Then GoTo ER1

'  Set tbSystem = myOpenRecordSet("##90", "System", dbOpenTable) ', dbOpenForwardOnly)
'  If tbSystem Is Nothing Then myBase.Close: End
'  tbSystem.Edit
'  tbSystem!resursLock = ""
'  tbSystem!Kurs = -Abs(tbSystem!Kurs)
'  tbSystem.Update
 ' tbSystem.Close

wrkDefault.CommitTrans
MsgBox "База переведена на новую дату!"
Exit Sub

ER1:
wrkDefault.rollback
End Sub


Function replaceResurs(id As Integer) As Boolean
Dim oldRes As Double, s As Double, n As Double, I As Integer, j As Integer
Dim newRes As Double

replaceResurs = False
        

oldRes = 0
newRes = getSystemField("newRes" & Ceh(id))


For I = 1 To befDays
    tmpDate = DateAdd("d", -I, curDate)
    
    sql = "SELECT 1, nomRes FROM Resurs" & Ceh(id) & _
    " WHERE xDate = '" & Format(tmpDate, "yy.mm.dd") & "'"
    
    If Not byErrSqlGetValues("W##12", sql, j, s) Then Exit Function
    If j = 0 Then ' нет этого дня
        day = Weekday(tmpDate)
        If Not (day = vbSunday Or day = vbSaturday) Then
            oldRes = oldRes + newRes
        End If
    Else
        oldRes = oldRes + s
    End If
Next I

'sql = "SELECT Sum(nomRes) AS rSum from Resurs" & Ceh(id) & _
" WHERE (((xDate)< '" & Format(curDate, "yy.mm.dd") & "'));"
'Debug.Print sql
'If Not byErrSqlGetValues("W##12", sql, oldRes) Then Exit Function

sql = "DELETE from Resurs" & Ceh(id) & _
" WHERE (((xDate)<'" & Format(curDate, "yy.mm.dd") & "'));"
If myExecute("##406", sql, 0) > 0 Then Exit Function


'****** отстреливаем итоги ***********
tmpSng = 0 'сумма невыполнено живых
sql = "SELECT Sum(Orders.workTime*OrdersInCeh.Nevip) AS nevip " & _
"FROM Orders INNER JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " & _
"WHERE (((Orders.StatusId)=1) AND ((Orders.CehId)=" & id & "));"
byErrSqlGetValues "##372", sql, tmpSng
s = 0 ' плюс неготовые образцы
sql = "SELECT Sum(OrdersMO.workTimeMO) AS Sum_workTimeMO " & _
"FROM Orders INNER JOIN OrdersMO ON Orders.numOrder = OrdersMO.numOrder " & _
"WHERE (((OrdersMO.StatO)='в работе') AND ((Orders.CehId)=" & id & "));"
byErrSqlGetValues "##378", sql, s
tmpSng = tmpSng + s

sql = "SELECT Nstan" & Ceh(id) & ", KPD_" & Ceh(id) & " FROM System;"
byErrSqlGetValues "##379", sql, n, s

On Error GoTo EN1
'записываем ресурс и КПД в пред.день
'!!! Если Мастер хочет изменить число станков и КПД на завтра, то этом.делать
'только завтра, поскольку новые значения применятся ко всему текущему дню
'(у дат впереди год - чтобы корректно работала сортировка)

sql = "SELECT Max(xDate) AS dLast FROM Itogi_" & Ceh(id) & ";"
byErrSqlGetValues "##407", sql, tmpStr
If tmpStr = Format(curDate, "yy.mm.dd") Then GoTo EN1 ' запись сегодня уже была

'numOrder = 0 ' признак ресурса
sql = "INSERT INTO Itogi_" & Ceh(id) & " ( [xDate], numOrder, Virabotka ) " & _
"SELECT '" & tmpStr & "', 0, " & Round(oldRes * n, 2) & ";"
'MsgBox sql
myExecute "##408", sql
sql = "INSERT INTO Itogi_" & Ceh(id) & " ( [xDate], numOrder, Virabotka ) " & _
"SELECT '" & tmpStr & "', 1, " & s & ";"
myExecute "##409", sql
'записываем сумму невыполнено живых(относятся к сегодня)
'numOrder = 2 ' признак суммы невыполнено живых
sql = "INSERT INTO Itogi_" & Ceh(id) & " ( [xDate], numOrder, Virabotka ) " & _
"SELECT '" & Format(curDate, "yy.mm.dd") & "', 2, " & tmpSng & ";"
myExecute "##410", sql

'оставляем только историю последнего месяца
sql = "DELETE from Itogi_" & Ceh(id) & _
" WHERE (((xDate)<'" & Format(DateAdd("m", -1, curDate), "yy.mm.dd") & "'));"
myExecute "##411", sql, 0
EN1:
replaceResurs = True
On Error Resume Next
End Function


