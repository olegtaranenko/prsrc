Attribute VB_Name = "OrderUtils"
Option Explicit


Sub nextDay()  '�������� ������ �� ���� ����
Dim Werk As String

Dim I As Integer, str As String, str1 As String, J As Integer, S As Double
Dim ch As String, tenOclock As String, Midnight As String
'MsgBox "������� �� ����� ����"

'wrkDefault.BeginTrans

sql = "DELETE from OrdersInCeh WHERE Stat = '�����'"
If myExecute("##63", sql, 0) > 0 Then GoTo ER1

tenOclock = "'" & Format(curDate, "yyyy-mm-dd 10:00:00") & "'"
Midnight = "'" & Format(curDate, "yyyy-mm-dd 00:00:00") & "'"

sql = "UPDATE Orders INNER JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " _
& " SET Orders.DateRS = " & tenOclock & ", OrdersInCeh.DateTimeMO = " & tenOclock _
& " WHERE Orders.DateRS  < " & Midnight & " And Orders.DateRS Is Not Null"

'Debug.Print sql
If myExecute("##11", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE OrdersEquip " _
& " SET outDateTime = " & tenOclock _
& " WHERE outDateTime < " & Midnight
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
Dim newRes As Double, equipId As Integer, KPD As Double

replaceResurs = False
        

oldRes = 0
sql = "select equipId, newRes, Nstan, KPD  from GuideResurs"
'byErrSqlGetValues "##newRes", sql, newRes

Set tbOrders = myOpenRecordSet("##newRes", sql, dbOpenForwardOnly)

If tbOrders Is Nothing Then Exit Function
While Not tbOrders.EOF
    equipId = tbOrders!equipId
    newRes = tbOrders!newRes
    N = tbOrders!Nstan
    KPD = tbOrders!KPD
    oldRes = 0
    
    For I = 1 To befDays
        tmpDate = DateAdd("d", -I, curDate)
        
        sql = "SELECT 1, nomRes FROM Resurs " & _
        " WHERE xDate = '" & Format(tmpDate, "yy.mm.dd") & "' and equipId = " & equipId
        
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
    & " WHERE xDate < '" & Format(curDate, "yy.mm.dd") & "' and equipId = " & equipId
    If myExecute("##406", sql, 0) > 0 Then Exit Function
    
  
    
    '****** ������������ ����� ***********
    tmpSng = 0 '����� ����������� �����
    sql = "SELECT Sum(oe.workTime * oc.Nevip) AS nevip" _
    & " FROM Orders o " _
    & " JOIN OrdersInCeh oc ON o.numOrder = oc.numOrder " _
    & " JOIN OrdersEquip oe ON oe.numOrder = oc.numOrder" _
    & " WHERE o.StatusId = 1 AND oe.equipId = " & equipId
    byErrSqlGetValues "##372", sql, tmpSng
    
    S = 0 ' ���� ��������� �������
    sql = "SELECT sum(oe.worktimeMO) as Sum_worktimeMO " _
    & " FROM Orders o " _
    & " JOIN OrdersInCeh oc ON o.numOrder = oc.numOrder " _
    & " JOIN OrdersEquip oe ON oe.numOrder = oc.numOrder" _
    & " WHERE oc.StatO ='� ������' AND oe.equipId = " & equipId
    byErrSqlGetValues "##372", sql, S
    
    tmpSng = tmpSng + S
    
    On Error GoTo EN1
    '���������� ������ � ��� � ����.����
    '!!! ���� ������ ����� �������� ����� ������� � ��� �� ������, �� ����.������
    '������ ������, ��������� ����� �������� ���������� �� ����� �������� ���
    '(� ��� ������� ��� - ����� ��������� �������� ����������)
    
    sql = "SELECT Max(xDate) AS dLast FROM Itogi WHERE equipId = " & equipId
    byErrSqlGetValues "##407", sql, tmpStr
    If tmpStr = Format(curDate, "yy.mm.dd") Then GoTo EN1 ' ������ ������� ��� ����
    
    'numOrder = 0 ' ������� �������
    sql = "INSERT INTO Itogi ( equipId, [xDate], numOrder, Virabotka ) " & _
    "SELECT " & equipId & ", '" & tmpStr & "', 0, " & Round(oldRes * N, 2)
    'MsgBox sql
    myExecute "##408", sql
    
    sql = "INSERT INTO Itogi ( equipId, [xDate], numOrder, Virabotka ) " & _
    "SELECT " & equipId & ", '" & tmpStr & "', 1, " & KPD
    myExecute "##409", sql
    
    '���������� ����� ����������� �����(��������� � �������)
    'numOrder = 2 ' ������� ����� ����������� �����
    sql = "INSERT INTO Itogi (equipId, [xDate], numOrder, Virabotka ) " & _
    "SELECT " & equipId & ", '" & Format(curDate, "yy.mm.dd") & "', 2, " & tmpSng
    myExecute "##410", sql
    
    '��������� ������ ������� ���������� ������
    sql = "DELETE from Itogi" _
    & " WHERE xDate < '" & Format(DateAdd("m", -1, curDate), "yy.mm.dd") & "' AND equipId = " & equipId
    myExecute "##411", sql, 0
    tbOrders.MoveNext
Wend
EN1:
tbOrders.Close

replaceResurs = True
On Error Resume Next
End Function


Private Sub addToCbStatus(ByRef statusComboBox As ComboBox, id, Optional begin As String = "")

    Static I As Integer
    If begin <> "" Then I = 0
    If id > lenStatus Then
        MsgBox "Err � Orders\addToCbStatus"
    End If

    statusComboBox.AddItem status(id)
    statId(I) = id
    I = I + 1

End Sub
    
    
Public Sub cbBuildStatuses(ByRef statusComboBox As ComboBox, ByRef statusIdOld As Integer)
    
    statusComboBox.Clear
    
    If statusIdOld = 4 Then
        addToCbStatus statusComboBox, 6 '"������"
    End If
    
    addToCbStatus statusComboBox, 7, "b" '"�������."
    If statusIdOld = 5 Then
        addToCbStatus statusComboBox, 5    '"�������"
    ElseIf statusIdOld = 8 Then
        statusIdOld = 1
        addToCbStatus statusComboBox, 1 '"� ������"
    ElseIf statusIdOld = 4 Then '"�����"
        addToCbStatus statusComboBox, 0
        addToCbStatus statusComboBox, 4
    Else
        addToCbStatus statusComboBox, 0 '"������"  '�� ��������� � �.�. ���
        addToCbStatus statusComboBox, 1 '"� ������"
        addToCbStatus statusComboBox, 2 '"������"  '������-� � ������� ��������
        addToCbStatus statusComboBox, 3 '"��������."
    End If

End Sub

Public Function cbMOsetByText(cb As ComboBox, stat As Variant, Optional baseIndex As Integer = 1) As Boolean
    cbMOsetByText = False
Dim I As Integer, txt As String
    txt = ""
    If Not IsNull(stat) Then txt = CStr(stat)
    If txt = "�����" Then
        If cb.List(baseIndex + 2) <> "�����" Then cb.AddItem "�����", baseIndex + 2
        If cb.List(baseIndex + 3) <> "���������" Then cb.AddItem "���������", baseIndex + 3
        cb.ListIndex = baseIndex + 2
        cbMOsetByText = True
    ElseIf txt = "���������" Then
        If cb.List(baseIndex + 2) = "�����" Then
            I = baseIndex + 3
        Else
            I = baseIndex + 2
        End If
        If cb.List(I) <> "���������" Then cb.AddItem "���������", I
        cb.ListIndex = I
    ElseIf txt = "� ������" Then
        cb.ListIndex = baseIndex + 1
        cbMOsetByText = True
    ElseIf txt = "�����" Or txt = "�������" Then
        cb.ListIndex = 1
    Else
        cb.ListIndex = 0
    End If

End Function



