Attribute VB_Name = "OrderUtils"
Option Explicit


Sub nextDay()  'возможен прыжок на неск дней
Dim Ceh As String

Dim I As Integer, str As String, str1 As String, j As Integer, s As Double
Dim ch As String
'MsgBox "переход на новую дату"

'wrkDefault.BeginTrans

sql = "DELETE from OrdersInCeh WHERE (((Stat)='готов'));"
If myExecute("##63", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE Orders INNER JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " & _
"SET Orders.DateRS = '" & Format(curDate, "yyyy-mm-dd 10:00:00") & _
"' WHERE (((Orders.DateRS) < '" & Format(curDate, "yyyy-mm-dd 00:00:00") & _
"' And Not (Orders.DateRS) Is Null));"
'MsgBox sql
If myExecute("##11", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE OrdersEquip " _
& " INNER JOIN OrdersInCeh ON OrdersEquip.numOrder = OrdersInCeh.numOrder " _
& "SET OrdersEquip.outDateTime = '" & Format(curDate, "yyyy-mm-dd 10:00:00") & _
"' WHERE Orders.outDateTime <'" & Format(curDate, "yyyy-mm-dd 0:0:0") & "'"

If myExecute("##404", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE OrdersEquip SET DateTimeMO = '" & Format(curDate, "yyyy-mm-dd 10:00:00") & _
"' WHERE DateTimeMO < '" & Format(curDate, "yyyy-mm-dd 00:00:00") & "'"
If myExecute("##405", sql, 0) > 0 Then GoTo ER1


If Not replaceResurs(1) Then GoTo ER1
If Not replaceResurs(2) Then GoTo ER1
If Not replaceResurs(3) Then GoTo ER1  '$$ceh

sql = "UPDATE System SET resursLock = '', Kurs = -Abs([Kurs]);"
If myExecute("##90", sql, 0) > 0 Then GoTo ER1

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
" WHERE xDate < '" & Format(curDate, "yy.mm.dd") & "'"
'Debug.Print sql
'If Not byErrSqlGetValues("W##12", sql, oldRes) Then Exit Function

sql = "DELETE from Resurs" & Ceh(id) & _
" WHERE (((xDate)<'" & Format(curDate, "yy.mm.dd") & "'));"
If myExecute("##406", sql, 0) > 0 Then Exit Function


'****** отстреливаем итоги ***********
tmpSng = 0 'сумма невыполнено живых
'' equipment
sql = "SELECT Sum(OrdersEquip.workTime*OrdersInCeh.Nevip) AS nevip " & _
"FROM OrdersInCeh " _
& " JOIN Orders ON Orders.numOrder = OrdersInCeh.numOrder " _
& " JOIN OrdersEquip ON OrdersEquip.numOrder = OrdersInCeh.numOrder " _
& " WHERE Orders.StatusId =1 AND OrdersEquip.CehId =" & id
byErrSqlGetValues "##372", sql, tmpSng

s = 0 ' плюс неготовые образцы
sql = "SELECT Sum(OrdersMO.workTimeMO) AS Sum_workTimeMO " _
& " FROM OrdersMO " _
& " JOIN Orders ON Orders.numOrder = OrdersMO.numOrder " _
& " JOIN OrdersEquip ON OrdersEquip.numOrder = OrdersMO.numOrder " _
& " WHERE OrdersMO.StatO = 'в работе' AND OrdersEquip.CehId = " & id
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
" WHERE xDate < '" & Format(DateAdd("m", -1, curDate), "yy.mm.dd") & "'"
myExecute "##411", sql, 0
EN1:
replaceResurs = True
On Error Resume Next
End Function




