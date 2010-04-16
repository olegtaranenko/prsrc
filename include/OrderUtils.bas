Attribute VB_Name = "OrderUtils"
Option Explicit


Sub nextDay()  'возможен прыжок на неск дней
Dim Ceh As String

Dim I As Integer, str As String, str1 As String, J As Integer, S As Double
Dim ch As String, tenOclock As String, Midnight As String
'MsgBox "переход на новую дату"

'wrkDefault.BeginTrans

sql = "DELETE from OrdersInCeh WHERE Stat = 'готов'"
If myExecute("##63", sql, 0) > 0 Then GoTo ER1

tenOclock = "'" & Format(curDate, "yyyy-mm-dd 10:00:00") & "'"
Midnight = "'" & Format(curDate, "yyyy-mm-dd 00:00:00") & "'"

sql = "UPDATE Orders INNER JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " _
& " SET Orders.DateRS = " & tenOclock & ", OrdersInCeh.DateTimeMO = " & tenOclock _
& " WHERE Orders.DateRS  < " & Midnight & " And Orders.DateRS Is Not Null"

'Debug.Print sql
If myExecute("##11", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE OrdersEquip " _
& " SET OrdersEquip.outDateTime = " & tenOclock _
& " WHERE OrdersEquip.outDateTime < " & Midnight

If myExecute("##404", sql, 0) > 0 Then GoTo ER1

'sql = "UPDATE OrdersInCeh SET DateTimeMO = '" & Format(curDate, "yyyy-mm-dd 10:00:00") & "'" _
& " WHERE DateTimeMO < '" & Format(curDate, "yyyy-mm-dd 00:00:00") & "'"
'If myExecute("##405", sql, 0) > 0 Then GoTo ER1


If Not replaceResurs(1) Then GoTo ER1
If Not replaceResurs(2) Then GoTo ER1
If Not replaceResurs(3) Then GoTo ER1  '$$ceh

sql = "UPDATE System SET resursLock = '', Kurs = -Abs(Kurs)"
If myExecute("##90", sql, 0) > 0 Then GoTo ER1

wrkDefault.CommitTrans
MsgBox "База переведена на новую дату!"
Exit Sub

ER1:
wrkDefault.Rollback
End Sub


Function replaceResurs(id As Integer) As Boolean
Dim oldRes As Double, S As Double, n As Double, I As Integer, J As Integer
Dim newRes As Double

replaceResurs = False
        

oldRes = 0
newRes = getSystemField("newRes" & Ceh(id))


For I = 1 To befDays
    tmpDate = DateAdd("d", -I, curDate)
    
    sql = "SELECT 1, nomRes FROM Resurs" & Ceh(id) & _
    " WHERE xDate = '" & Format(tmpDate, "yy.mm.dd") & "'"
    
    If Not byErrSqlGetValues("W##12", sql, J, S) Then Exit Function
    If J = 0 Then ' нет этого дня
        day = Weekday(tmpDate)
        If Not (day = vbSunday Or day = vbSaturday) Then
            oldRes = oldRes + newRes
        End If
    Else
        oldRes = oldRes + S
    End If
Next I

'sql = "SELECT Sum(nomRes) AS rSum from Resurs" & Ceh(id) & _
" WHERE xDate < '" & Format(curDate, "yy.mm.dd") & "'"
'Debug.Print sql
'If Not byErrSqlGetValues("W##12", sql, oldRes) Then Exit Function

sql = "DELETE from Resurs" & Ceh(id) & _
" WHERE xDate < '" & Format(curDate, "yy.mm.dd") & "'"
If myExecute("##406", sql, 0) > 0 Then Exit Function


'****** отстреливаем итоги ***********
tmpSng = 0 'сумма невыполнено живых
'' equipment
sql = "SELECT Sum(oe.workTime * oc.Nevip) AS nevip " & _
"FROM Orders      o " _
& " JOIN OrdersInCeh oc ON o.numOrder = oc.numOrder " _
& " JOIN OrdersEquip oe  ON oe.numOrder = oc.numOrder AND oe.cehId = oc.cehId " _
& " WHERE o.StatusId = 1 AND oc.CehId = " & id
byErrSqlGetValues "##372", sql, tmpSng

S = 0 ' плюс неготовые образцы
sql = "SELECT Sum(oc.workTimeMO) AS Sum_workTimeMO " _
& " FROM OrdersInCeh oc" _
& " JOIN Orders o ON o.numOrder = oc.numOrder " _
& " WHERE oc.StatO = 'в работе' AND oc.CehId = " & id
byErrSqlGetValues "##378", sql, S
tmpSng = tmpSng + S

sql = "SELECT Nstan" & Ceh(id) & ", KPD_" & Ceh(id) & " FROM System"
byErrSqlGetValues "##379", sql, n, S

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
"SELECT '" & tmpStr & "', 1, " & S & ";"
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


