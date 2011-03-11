Attribute VB_Name = "OrderUtils"
Option Explicit

'
' ��� ������������ ��� ���������� ���������� �������� �����,
' � ����������� �� ����, ����� ������������� ������ ������� (��� ������� Orders)
'
Type WerkColumn
    Field As String ' ��������� �������� �������. ������������� ����� ���� � �������
    WerkId As Integer ' ���� 0 - ���������� ��� ���� �������������
    columnWidth As Integer ' ������ �������. ����� ����������� �������� � ����������� �� ������ ������� � �����
    GridColIndex As Long ' ���������� ����� � �����
    
    Visible As Boolean  ' false ���� ������� � �������� �� ������������. �������� �������, ����������� � �.�.
                        ' �������� ��� ������������� ���� ��������� ����� ����� 0
End Type



Sub nextDay()  '�������� ������ �� ���� ����
Dim Werk As String

Dim I As Integer, str As String, str1 As String, J As Integer, S As Double
Dim ch As String, tenOclock As String, Midnight As String
Dim checkNextDay As Boolean


checkNextDay = getEffectiveSetting("checkNextDay", True)
If Not checkNextDay Then
    Exit Sub
End If


'MsgBox "������� �� ����� ����"

'wrkDefault.BeginTrans

sql = "delete from OrdersInCeh" _
& vbCr & " WHERE EXISTS (" _
& vbCr & "     select 1 from OrdersEquip oe where oe.Stat = '�����' and oe.numorder = OrdersInCeh.numorder" _
& vbCr & ") and not exists (" _
& vbCr & "     select 1 from vw_reestr r where r.numorder = ordersInCeh.numorder group by r.numorder having min(isnull(r.stat, '')) != max(isnull(r.stat, ''))" _
& vbCr & ")"

'Debug.Print sql
If myExecute("##63", sql, 0) > 0 Then GoTo ER1

tenOclock = "'" & Format(curDate, "yyyy-mm-dd 10:00:00") & "'"
Midnight = "'" & Format(curDate, "yyyy-mm-dd 00:00:00") & "'"

sql = "UPDATE Orders JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " _
& " SET Orders.DateRS = " & tenOclock & ", OrdersInCeh.DateTimeMO = " & tenOclock _
& " WHERE Orders.DateRS  < " & Midnight & " And Orders.DateRS Is Not Null"

If myExecute("##11", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE OrdersEquip " _
& " SET outDateTime = " & tenOclock _
& " WHERE outDateTime < " & Midnight _
& " AND exists (select 1 from OrdersInCeh oc where oc.numorder = OrdersEquip.Numorder) "
If myExecute("##404", sql, 0) > 0 Then GoTo ER1


sql = "UPDATE OrdersInCeh SET DateTimeMO = " & tenOclock _
& " WHERE DateTimeMO < " & Midnight

If myExecute("##405", sql, 0) > 0 Then GoTo ER1


If Not replaceResurs Then GoTo ER1

sql = "UPDATE System SET resursLock = '', Kurs = -Abs(Kurs)"
If myExecute("##90", sql, 0) > 0 Then GoTo ER1

wrkDefault.CommitTrans
MsgBox "���� ���������� �� ����� ����!"
Exit Sub

ER1:
wrkDefault.Rollback
End Sub


Function replaceResurs() As Boolean
Dim oldRes As Double, S As Double, N As Double, I As Integer, J As Integer
Dim newRes As Double, EquipId As Integer, KPD As Double

replaceResurs = False
        

oldRes = 0
sql = "select equipId, newRes, Nstan, KPD  from GuideResurs"
'byErrSqlGetValues "##newRes", sql, newRes

Set tbOrders = myOpenRecordSet("##newRes", sql, dbOpenForwardOnly)

If tbOrders Is Nothing Then Exit Function
While Not tbOrders.EOF
    EquipId = tbOrders!EquipId
    newRes = tbOrders!newRes
    N = tbOrders!Nstan
    KPD = tbOrders!KPD
    oldRes = 0
    
    For I = 1 To befDays
        tmpDate = DateAdd("d", -I, curDate)
        
        sql = "SELECT 1, nomRes FROM Resurs " & _
        " WHERE xDate = '" & Format(tmpDate, "yy.mm.dd") & "' and equipId = " & EquipId
        
        If Not byErrSqlGetValues("W##12", sql, J, S) Then Exit Function
        If J = 0 Then ' ��� ����� ���
            day = Weekday(tmpDate)
            If Not (day = vbSunday Or day = vbSaturday) Then
                oldRes = oldRes + newRes
            End If
        Else
            oldRes = oldRes + S
        End If
    Next I
    
    
    sql = "DELETE from Resurs" _
    & " WHERE xDate < '" & Format(curDate, "yy.mm.dd") & "' and equipId = " & EquipId
    If myExecute("##406", sql, 0) > 0 Then Exit Function
    
  
    
    '****** ������������ ����� ***********
    tmpSng = 0 '����� ����������� �����
    sql = "SELECT Sum(oe.workTime * oe.Nevip) AS nevip" _
    & " FROM Orders o " _
    & " JOIN OrdersEquip oe ON oe.numOrder = o.numOrder" _
    & " WHERE o.StatusId = 1 AND oe.equipId = " & EquipId _
    & " AND EXISTS (select 1 from OrdersInCeh oc where oc.numorder = oe.numorder) "
    byErrSqlGetValues "##372", sql, tmpSng
    
    S = 0 ' ���� ��������� �������
    sql = "SELECT sum(oe.worktimeMO) as Sum_worktimeMO " _
    & " FROM OrdersEquip oe " _
    & " WHERE oe.StatO ='� ������' AND oe.equipId = " & EquipId _
    & " AND EXISTS (select 1 from OrdersInCeh oc where oc.numorder = oe.numorder) "
    byErrSqlGetValues "##372", sql, S
    
    tmpSng = tmpSng + S
    
    On Error GoTo EN1
    '���������� ������ � ��� � ����.����
    '!!! ���� ������ ����� �������� ����� ������� � ��� �� ������, �� ����.������
    '������ ������, ��������� ����� �������� ���������� �� ����� �������� ���
    '(� ��� ������� ��� - ����� ��������� �������� ����������)
    
    sql = "SELECT Max(xDate) AS dLast FROM Itogi WHERE equipId = " & EquipId
    byErrSqlGetValues "##407", sql, tmpStr
    If tmpStr = Format(curDate, "yy.mm.dd") Then GoTo EN1 ' ������ ������� ��� ����
    
    'numOrder = 0 ' ������� �������
    sql = "INSERT INTO Itogi ( equipId, [xDate], numOrder, Virabotka ) " & _
    "SELECT " & EquipId & ", '" & tmpStr & "', 0, " & Round(oldRes * N, 2)
    'MsgBox sql
    myExecute "##408", sql
    
    sql = "INSERT INTO Itogi ( equipId, [xDate], numOrder, Virabotka ) " & _
    "SELECT " & EquipId & ", '" & tmpStr & "', 1, " & KPD
    myExecute "##409", sql
    
    '���������� ����� ����������� �����(��������� � �������)
    'numOrder = 2 ' ������� ����� ����������� �����
    sql = "INSERT INTO Itogi (equipId, [xDate], numOrder, Virabotka ) " & _
    "SELECT " & EquipId & ", '" & Format(curDate, "yy.mm.dd") & "', 2, " & tmpSng
    myExecute "##410", sql
    
    '��������� ������ ������� ���������� ������
    sql = "DELETE from Itogi" _
    & " WHERE xDate < '" & Format(DateAdd("m", -1, curDate), "yy.mm.dd") & "' AND equipId = " & EquipId
    myExecute "##411", sql, 0
    tbOrders.MoveNext
Wend
EN1:
tbOrders.Close

replaceResurs = True
On Error Resume Next
End Function


    

