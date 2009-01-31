Attribute VB_Name = "Common"
Option Explicit
'���e��\��������\��������\��������� ����������:
' - onErrorOtlad = 1 ' ����� ������ err

' ������ �������� (�-�� �����) ��� ������ ���������.
' �������� � ����������, ������������ � ������ Nakladna.frm

Public gCfgOrderPageSize As Integer

Public isOrders As Boolean
Public isCehOrders As Boolean
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
Public tbCeh As Recordset
Public tbDMC As Recordset
Public tbDocs As Recordset
Public tbGuide As Recordset
Public tbSeries As Recordset
Public Node As Node


Public isBlock As Boolean
Public Ceh(10) As String
Public cehId As Integer
Public Const lenStatus = 20
Public statId(lenStatus) As Integer
Public status(lenStatus) As String
Public lenProblem As Integer
Public Problems(20) As String
Public manId() As Integer '$$7
Public Manag() As String  '
Public insideId() As String
Public Const begCehProblemId = 10 ' ������ ������� ������� � �����������
Public neVipolnen As Double, neVipolnen_O As Double
Public maxDay As Integer ' ����� ���� � �������
Public befDays As Integer ' ����� ���� �� ���� ������� (����� ��������� ����)
Public webSvodkaPath As String
Public webLoginsPath As String
Public webNomenks As String '- ���� � � sTime
Public webProducts As String

' ���������� �� cfg.frm
Public loginsPath As String
Public SvodkaPath As String
Public NomenksPath As String
Public ProductsPath As String


'Public baseNamePath As String
Public begDate As Date ' ���� ������������� ��������
Public logFile As String
Public dostup As String
Public otlad As Variant
Public tbSize As Integer
Public cErr As String '��������� ������� ����� ������������� Err, ���� ��
                      '���� ������ ��������� �� Err ������ ���� MsgBox
Public zakazNum As Long  ' ���-�� ������� �  M��.�������
Public gNzak As String  ' ��� ����� ������
'Public gFirmId As Integer
Public gFirmId As String
Public gProductId As String
Public gProduct As String
Public gDocDate As Date
Public gSeriaId As String
Public gKlassId As String
Public gNomNom As String
Public numDoc As Long, numExt As Integer

Public oldValue As String '������ �������� ����, ����������� ���������
Public curDate As Date
Public lastYear As Integer

Public begDay As Integer ' ���� ������� ����� ������
Public endDay As Integer ' ���� ���������� ����� ������
Public begDayMO As Integer ' ���� ������� ����� �� ������
Public endDayMO As Integer ' ���� ���������� ����� �� ������
Public flEdit As String ' ������������� ������
Public Nstan As Double
Public kpd As Double
Public newRes As Double ' ����� �� ���������
Public nr As Double ', dr As Double '��������� ���. � ���. �������
Public isLive As Boolean ' ���� - ����� �����
Public zagAll As Double, zagLive As Double

Public table As Recordset '
Public myQuery As QueryDef
Public sql As String      ' ������������� �����������
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
Public noClick As Boolean '������ True, ����� �������� onClick lb
Public bilo As Boolean
Public cep As Boolean  '�������������� ����� ���� ���� ������-��� ��������
Public oldCellColor As Long
Public prExt As Integer, pType As String
Public gridIsLoad As Boolean

Public orColNumber As Integer ' ����� ������� � Orders
Public orSqlWhere() As String
Public orSqlFields() As String  '
Public orNomZak As Integer, orCeh As Integer, orData As Integer, orTema As Integer
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

Public NN() As String, QQ() As Double ' ������������ ������������ � ���-��
Public QQ2() As Double, QQ3() As Double
Public skladId As Integer

Private Const dhcMissing = -2 '����� ��� quickSort
Public Const cDELLwidth = 19200 ' ��� ����� � ��� = 19290

Public Const dcSourId = 0 ' �����
Public Const dcDate = 1
Public Const dcNumDoc = 2
Public Const dcSour = 3
Public Const dcDest = 4
Public Const dcVenture = 6
Public Const dcNote = 5

'Grid � FirmComtex
Public Const fcId = 0
Public Const fcFirmName = 1
Public Const fcInn = 2
Public Const fcOkonx = 3
Public Const fcOkpo = 4
Public Const fcKpp = 5
Public Const fcAddress = 6
Public Const fcPhone = 7

Public Const fcFormatString = _
  "|< ��������  �����" _
& "|>���" _
& "|>�����" _
& "|>����" _
& "|>���" _
& "|<�����" _
& "|<�������" _


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
Public Const gfAdres = 16 '������
Public Const gfAtr1 = 17
Public Const gfAtr2 = 18
Public Const gfAtr3 = 19
Public Const gfLogin = 20
Public Const gfPass = 21
Public Const gfId = 22

Public Const chNomZak = 1
Public Const chM = 2
Public Const chStatus = 3
Public Const chVrVip = 4
Public Const chProcVip = 5
Public Const chProblem = 6
Public Const chDataVid = 7
Public Const chVrVid = 8
Public Const chDataRes = 9
Public Const chFirma = 10
Public Const chLogo = 11
Public Const chIzdelia = 12
Public Const chKey = 13

Public Const zgPrinato = 1
Public Const zgNomRes = 2
'Public Const zgDopRes = 3
'Public Const zgRaspred = 4
Public Const zgResurs = 3
Public Const zgZagruz = 4
Public Const zgOstatki = 5
Public Const zgLive = 6
'����� ������
Public Const zkPrinato = 1
Public Const zkFirmKolvo = 2
Public Const zkResurs = 3
Public Const zkMzagr = 4
Public Const zkMbef = 5
Public Const zkHide = 6 '�������
Public Const zkMost = 7
Public Const zkCzagr = 8
Public Const zkCost = 9
Public Const zkCliv = 10
'������������ � ��������� �� sProducts Grid2
Public Const fnNomNom = 1
Public Const fnNomName = 2
Public Const fnEdIzm = 3
Public Const fnQuant = 4

'Grid5 � sProducts � Otgruz
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

'������ ����� ���.�������� ������� �� ��� �������� ������� ����� �� �������
Public stDays() As Integer        ' ������� ��� �������� (��,��,���������)
Public stDay As Integer '����� ���������� stDays(�������)
                            
Public nomRes() As Double
Public delta() As Double
Public tmp() As Double
Public tmpL() As Long
Public ost() As Double, befOst() As Double

' ������ ��������� ������� � ������� ��������� � ������
' (�� CtrlLeftClick) � sProducts.Grid5
Public selectedItems() As Long
Public Const otladColor = &H80C0FF




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


'���� ������ �������� ="W.." - �� �������� Err �� �����-� Where, � ���
'��������� ��������, ���� ��� ���� ��� ���� ��� ��������� ��������, �� � sql
'�. ������ ��������� "1" � ������� �� � I. ����� ���� I=0 �� ���� Err Where
'$odbc15$
Function byErrSqlGetValues(ParamArray val() As Variant) As Boolean
Dim tabl As Recordset, I As Integer, maxi As Integer, str As String, c As String

byErrSqlGetValues = False
maxi = UBound(val())
If maxi < 1 Then
    wrkDefault.rollback
    MsgBox "���� ���������� ��� �\� byErrSqlGetValues()"
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
'        msgOfEnd CStr(val(0)), "��� ������� ��������������� Where."
        wrkDefault.rollback
        MsgBox "��� ������� ��������������� Where!", , "Error-" & str
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
    Grid.rows = 2
    clearGridRow Grid, 1
Else
    Grid.rows = 3
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
        MsgBox "������� ��������� ��������� 500 ����? �������� ��������������!", , "Err � dayMassLenght()"
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
    MsgBox Error, , "������ 17-" & Err & ":  " '##17
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
Dim s As Double
    s = Timer
    While Timer - s < tau ' 1 ���
        DoEvents
    Wend

End Sub

Sub delZakazFromReplaceRS()

'sql = "DELETE ReplaceRS.* From ReplaceRS " & _
"WHERE (((ReplaceRS.numOrder)=" & gNzak & "));"
sql = "DELETE From ReplaceRS " & _
"WHERE (((numOrder)=" & gNzak & "));"
'MsgBox sql
myExecute "##79", sql, 0 ' �������, ���� ����
End Sub


Sub exitAll()
If isOrders Then Unload Orders
If isCehOrders Then Unload CehOrders
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
For il = 1 To Grid.rows - 1
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
For il = beg To Grid.rows - 1
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

'��� ������ �� ������ �����, � ���� �� ��������� - ��� ���������� ����������
Function yymmdd(dateStr As String) As String
yymmdd = right$(dateStr, 2) & "." & Mid$(dateStr, 4, 2) & "." & left$(dateStr, 2)
End Function

Function getSystemField(field As String) As Variant
getSystemField = Null
sql = "SELECT " & field & " From System"
Set tbSystem = myOpenRecordSet("##147", sql, dbOpenForwardOnly)
'If tbSystem Is Nothing Then myBase.Close: End
getSystemField = tbSystem.fields(field)
tbSystem.Close
End Function

'$odbc08$----------------------------------------------------------
Function getTableField(tabl As String, field As String) As Variant
Dim table As Recordset

getTableField = Null
sql = "SELECT " & tabl & "." & field & " AS fff From " & tabl & _
      " WHERE (((OrdersMO.numOrder)=" & gNzak & "));"
Set table = myOpenRecordSet("##59", sql, dbOpenForwardOnly)
If table Is Nothing Then Exit Function
If Not table.BOF Then getTableField = table!fff
table.Close
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
gNzak = left$(str, I - 1)

End Function
'$odbc10$
Function getResurs() As Integer
Dim I As Integer, j As Integer, rMaxDay As Integer, s As Double

Set tbSystem = myOpenRecordSet("##93", "System", dbOpenForwardOnly)
If tbSystem Is Nothing Then myBase.Close: End
If cehId = 1 Then
    kpd = tbSystem!KPD_YAG
    Nstan = tbSystem!NstanYAG
    newRes = tbSystem!newResYAG
ElseIf cehId = 2 Then
    kpd = tbSystem!KPD_CO2
    Nstan = tbSystem!NstanCO2
    newRes = tbSystem!newResCO2
ElseIf cehId = 3 Then           '$$ceh
    kpd = tbSystem!KPD_SUB      '
    Nstan = tbSystem!NstanSUB   '
    newRes = tbSystem!newResSUB '
Else
    getResurs = 1
    Exit Function
End If
tbSystem.Close

sql = "SELECT nomRes from Resurs" & Ceh(cehId) & " ORDER BY xDate;"
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
    If rMaxDay = j Then
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

addDays max(stDay, rMaxDay) '��������� ���, ���� ���� ���. ���� �������
                            '����� ��� stDay ��� rMaxDay(����. ������� ���)
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

'*********************** ��������� ������ **************************
s = Timer / 3600:
I = Int(s)
If I < 9 Then
ElseIf I < 13 Then
    nr = Round(nomRes(1) - s + 9, 1)
ElseIf I < 14 Then
    nr = Round(nomRes(1) - 4, 1)
Else
    nr = Round(nomRes(1) - s + 10, 1)
End If
If nr < 0 Then
    nr = 0
End If

getResurs = maxDay '1:
Exit Function

ERR1:
If Err = 9 Then
    dayMassLenght '������������ �����������, ���� ����
    Resume
Else
    MsgBox Error, , "������ 18-" & Err & ":  " '##18
    myBase.Close: End
End If

End Function
'$NOodbc$
'������ "error"- ���� ������� ������ ��� (�� ������� ��������� SQL) .
'reg="" -  ������ �������� ��� WHERE ��� ���������� ����� ������
'          ���� "" - ���������� �� WHERE �� ���� �� ���������(� ������ begDate � CurDate)
'          ���� "error" ���� ���� �� ������������
'reg<>"" - ������ �������� ��� WHERE ��� ���������� �� startDate
'          ���� "" ���� startDate ������ begDate(�� ������� ��������� SQL)
Function getWhereByDateBoxes(Frm As Form, dateField As String, _
begDate As Date, Optional reg As String = "") As String

Dim str As String, ckStart As Boolean, ckEnd  As Boolean

getWhereByDateBoxes = "": str = "":

ckStart = False: ckEnd = False
On Error Resume Next ' �� ������, ���� � ���� ����� � ��� ��� �������
If Frm.ckEndDate.value > 0 Then ckEnd = True  '�� ��� ��� �� �����������
If Frm.ckStartDate.value > 0 And Frm.ckStartDate.Visible Then ckStart = True
On Error GoTo 0

If ckStart Then
    If Not isDateTbox(Frm.tbStartDate) Then GoTo ERRd  'tmpDate
End If
If reg = "" Then ' ���� ������ �����
    If DateDiff("d", begDate, tmpDate) > 0 And ckStart Then _
        str = "(" & dateField & ") >=" & Format(tmpDate, "'yyyy-mm-dd'")
    If ckEnd Then
      If Not isDateTbox(Frm.tbEndDate) Then GoTo ERRd
      If ckStart Then
        If DateDiff("d", Frm.tbStartDate.Text, tmpDate) < 0 Then
          MsgBox "��������� ���� ������� �������� �� ������ ��������� �������� ", , "��������������"
ERRd:     getWhereByDateBoxes = "error"
          Exit Function
        End If
      End If
      If DateDiff("d", tmpDate, curDate) > 0 Then getWhereByDateBoxes = _
          "(" & dateField & ")<='" & Format(tmpDate, "yyyy-mm-dd") & " 11:59:59 PM'"
    End If
ElseIf ckStart Then ' ���� ������ ��
    If DateDiff("d", begDate, tmpDate) <= 0 Then Exit Function
    tmpDate = DateAdd("d", -1, tmpDate) ' "-1" ���� �.�. ����� "+ 23�59�59�
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
    tb.left = Grid.CellLeft + Grid.left
    tb.Top = Grid.CellTop + Grid.Top
    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)
    tb.Visible = True
    tb.SetFocus
    tb.ZOrder
    Grid.Enabled = False '����� ������ �� ��� ������
End Sub

Sub listBoxSelectByText(lb As listBox, obrazec As String)
Dim I As Integer
    
    For I = 0 To lb.ListCount - 1 '
'       noClick = True
        If obrazec = lb.List(I) Then
            lb.Selected(I) = True '�������� ������ onClick'
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
    orSqlWhere(col) = "(GuideStatus.Status)<>'������'"
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
    lb.left = Grid.CellLeft + Grid.left
    lb.ListIndex = 0
    If sel <> "" Then
        For I = 0 To lb.ListCount - 1 '
            If Grid.Text = lb.List(I) Then
'                noClick = True
                lb.ListIndex = I '�������� ������ onClick
'                noClick = False
                Exit For
            End If
        Next I
    End If
    lb.Visible = True
    lb.ZOrder
    lb.SetFocus
    Grid.Enabled = False '����� ������ �� ��� ������
'    lbIsActiv = True
End Sub

Function LoadNumeric(Grid As MSFlexGrid, row As Long, col As Integer, _
        val As Variant, Optional myErr As String = "", Optional fmt As String) As Double
 If IsNull(val) Then
    Grid.TextMatrix(row, col) = ""
    LoadNumeric = 0 ' ��� log �����
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
    LoadDate = "" ' ��� log �����
    If myErr <> "" Then
        msgOfZakaz (myErr)
        Grid.TextMatrix(row, col) = "??"
    End If
 Else
    LoadDate = Format(val, formatStr)
    If LoadDate = "00" Then LoadDate = "" '    ������ ��� 00 �����
    Grid.TextMatrix(row, col) = LoadDate
 End If
End Function
'$NOodbc$
Sub GridToExcel(Grid As MSFlexGrid, Optional title As String = "")

Dim objExel As Excel.Application, c As Long, r As Long
Dim I As Integer, strA() As String, begRow As Integer, str As String

begRow = 3
If title = "" Then begRow = 1

Set objExel = New Excel.Application
objExel.Visible = True
objExel.SheetsInNewWorkbook = 1
objExel.Workbooks.Add
With objExel.ActiveSheet
.Cells(1, 2).value = title
ReDim Preserve strA(Grid.Cols + 1)
For r = 0 To Grid.rows - 1
    For c = 1 To Grid.Cols - 1
        str = Grid.TextMatrix(r, c) '=' - ������� �������������� ��� ����� ������
        If left$(str, 1) = "=" Then str = "." & str
        strA(c - 1) = str
    Next c
   .Range(.Cells(begRow + r, 1), .Cells(begRow + r, Grid.Cols)).FormulaArray = strA
Next r

End With
Set objExel = Nothing
End Sub


Sub initOrCol(colNum As Integer, Optional field As String = "")
orColNumber = orColNumber + 1
colNum = orColNumber
ReDim Preserve orSqlFields(orColNumber + 1)
orSqlFields(orColNumber) = field
End Sub



'$odbc10$
Sub Main()
Dim I As Integer, s As Double, str As String, str1 As String, str2 As String
Dim isXP As Boolean
If App.PrevInstance = True Then
    MsgBox "��������� ��� ��������", , "Error"
    End
End If

ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): ReDim QQ3(0) '����� Ubound ������� �� ������ Err

flReportArhivOrders = False
ReDim tmpL(0)

cfg.isLoad = False  '$$2
loadEffectiveSettingsApp
dostup = getEffectiveSetting("dostup")
loginsPath = getEffectiveSetting("loginsPath")
SvodkaPath = getEffectiveSetting("SvodkaPath")
NomenksPath = getEffectiveSetting("NomenksPath")
ProductsPath = getEffectiveSetting("ProductsPath")

initLogFileName

checkReloadCfg

isXP = (Dir$("C:\WINDOWS\net.exe") = "") '� XP ��� �����
On Error GoTo ERRs ' �� ���� Err ���� � ���� �� �.������ server, ���� �� ��� DOS ���� ����.Err=53
otlad = getEffectiveSetting("otlad")

On Error GoTo 0
If dostup = "�" Then dostup = "c"
If dostup = "0" Then
    I = 5 / I  '�������� ������� ���������� ���������
    End
End If
If dostup <> "a" And dostup <> "m" And dostup <> "" And dostup <> "b" _
And dostup <> "c" And dostup <> "y" _
And dostup <> "s" Then '$$$ceh
    MsgBox "'" & dostup & "' - �������� �������� �������!", , ""
    End
End If

baseOpen

mainTitle = getMainTitle

If Not IsEmpty(otlad) Then '
  webSvodkaPath = "C:\WINDOWS\TEMP\svodkaW."
  webLoginsPath = "C:\WINDOWS\TEMP\logins."

Else
    webSvodkaPath = SvodkaPath          '$$2
    webLoginsPath = loginsPath          '
End If

On Error GoTo 0

'�������� ��������� Win98
str = "05.08.2004"
tmpDate = str
If str <> Format(tmpDate, "dd.mm.yyyy") Then ' ��� "����������" � Win98 �� ��������
    str = "������������ ��������� " & Chr(151) & " ����������  '�������'."
    GoTo AA
End If
'begDate = "01.01.2003"
sql = "SELECT begDate, lastYear From System" ' WHERE (((System.begDate) Like '*##.##.20##*'));"
Set tbSystem = myOpenRecordSet("##181", sql, dbOpenForwardOnly)
'If tbSystem Is Nothing Then myBase.Close: End
If tbSystem.BOF Then
    tbSystem.Close: myBase.Close
    str = "����\������� ������ " & Chr(151) & " ����������  '��.��.����'."
    '"����������" � XP �������� ����
    If isXP Then str = str & " ����� ��������, ��� ���������� '�������' ������."
AA: MsgBox "����\���������\������ ����������\���� � ���������\" & str, , _
    "��� ���������� ������ ��������� ���������� ��������� ��������� Win98:"
    End
Else
    lastYear = tbSystem!lastYear
    begDate = tbSystem!begDate
End If
tbSystem.Close

Ceh(1) = "YAG"
Ceh(2) = "CO2"
Ceh(3) = "SUB" '$$ceh

'�� ������ ������� ����. 3� �������
nextDayDetect ' ����� �������-�� CurDate
stDay = startDays() ' � �.�. ������������� ��������� ����������� dayMassLenght
If befDays <> 0 Then nextDay ' ������� ���� �� ����� ����

checkNextYear '$$3 ���� �������� ��� - �������� ���������� ���������

' ��������� ��� �����*********************************************
'If Not (dostup = "c" Or dostup = "y") Then
If dostup = "a" Or dostup = "m" Or dostup = "" Or dostup = "b" Then
 'logFile = "C:\Windows\Orders" ' ��� ����������
 logFile = App.path & "\" & App.exeName
 str2 = logFile & "$.log" ' ��������� ����
 logFile = logFile & ".log"
 
 On Error GoTo ENop
 Open logFile For Input As #2
 Open str2 For Output As #3
 While Not EOF(2)
    Input #2, str
    I = InStr(str, vbTab)
    If I < 9 Then GoTo ENlog
    str1 = left$(str, I - 1)
    If Not IsDate(str1) Then GoTo ENlog
    If DateDiff("d", str1, curDate) <= 7 Then Print #3, str ' ������� > 7�� ���� ��������
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
        MsgBox "Err � Orders\FormLoad"
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
        MsgBox "����� ������� � ���� ��������� 20"
        End
    End If
    Problems(table!ProblemId) = table!problem
    table.MoveNext
    If Not table.EOF Then GoTo CC
table.Close

'�������� ������ ���������� ������ � ������
CheckIntegration

If dostup = "y" Then
    cehId = 1: CehOrders.Show
ElseIf dostup = "c" Then
    cehId = 2: CehOrders.Show
ElseIf dostup = "s" Then        '$$$ceh
    cehId = 3: CehOrders.Show   '
Else
    Orders.Show
End If
Exit Sub
ERRf:
MsgBox "����\���������\������ ����������\���� � ���������\�����\" & _
      "����������� ����� � ������� ������ ����� " & Chr(151) & _
      " ���������� ����� ������ �������!", , "��� ���������� ������ " & _
      "��������� ���������� ��������� ��������� Win98: "
End

ERRs:
MsgBox "������� �� ������ ���������������� ����", , "�������� ��������������!"
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

' ��� ������� ������� ��������� ������� ������������ ������ ����������
' �������������� � ���������� � �������������� ����������� ��������
' �� ������ ������� ���������
' ���� ���������� ���������������, �� ������ ��������������
' �������������� ����� ��������� ���������� � System
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
        msgOk = MsgBox("������ """ & servers!ventureName & """ (" & servers!sysname & ") " _
        & " ��������, �� �� �� ����� ������� �� �������� �� ����� ����������� ������������� � ���������� " _
        & vbCr & "����� ����� � ��������� �������� ������� ������ ������(Cancel)" _
        & vbCr & "���� �� �� ���-���� ������ ���������� ������, ������� ������ ��" _
        , vbOKCancel, "��������������")
        
        If msgOk <> vbOK Then myBase.Close: End
         
    ElseIf fromComtex = 0 And servers!standalone = 1 Then
        msgOk = MsgBox("������ """ & servers!ventureName & """ (" & servers!sysname & ") " _
        & " �������� � �������� �� ����� ���������� ������ � ����������." _
        & vbCr & "� ���� ����� ���� ��������� ��������� ���, ��� ��� �� ����� �������� � ���� ��������." _
        & " ������� ��������� ���������� �� ����� �������� ����." _
        & vbCr & "����� ����� � ��������� �������� ������� ������ ������(Cancel)" _
        & vbCr & "���� �� �� ���-���� ������ ���������� ������, ������� ������ ��" _
        , vbOKCancel, "��������������")
        
        If msgOk <> vbOK Then myBase.Close: End
    
    ElseIf fromComtex = -1 And servers!standalone <> 1 Then
        msgOk = MsgBox("������ """ & servers!ventureName & """ (" & servers!sysname & ") " _
        & " �� ��������, ���� � ���������� �������, ��� ��������� ����� �������� ���������. " _
        & vbCr & vbCr & " ����� ����� ����� �������� ������ � ������ ���������!" _
        & vbCr & "����� ����� � ��������� �������� ������� ������ ������(Cancel)" _
        , vbOKCancel, "��������������")
        
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
Dim I As Integer, s As Double

I = Format(Now, "yyyy")
If I <= lastYear Then Exit Sub

If MsgBox("���������� ������������� �������� ���������� ��������� �� ����� ���. " & _
"���������?", vbDefaultButton2 Or vbYesNo, "�����������!") = vbNo Then Exit Sub

wrkDefault.BeginTrans

If Not statisticReplace("GuideFirms") Then GoTo ER1
If Not statisticReplace("BayGuideFirms") Then GoTo ER1

If valueToSystemField("##389", I, "lastYear") Then
    wrkDefault.CommitTrans
    lastYear = I
    MsgBox "���� ���������� � �����(" & I & ") ���!"
Else
ER1: wrkDefault.rollback
    MsgBox "��������� �� ������ ��������� ���� � �����(" & I & ") ���! " & _
    "������������� ��������� ��� ��� ��� ��������� � ���������������.", , "Error"
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

Sub msgOfZakaz(myErrCod As String, Optional msg As String = "")
    myErrCod = Mid$(myErrCod, 3)
    If msg = "" Then msg = "�������� ����������� ������."
    MsgBox msg & " ������ � ���� ������� ���� " & _
    "����������. �������� ��������������!", , _
    "��������� ���� ���� (Err=" & myErrCod & ") � ������ � " & gNzak & _
    "  " & tbCeh!Manag
End Sub

Sub msgOfEnd(myErrCod As String, Optional msg As String = "")
    myErrCod = Mid$(myErrCod, 3)
    MsgBox msg & " �������� ��������������!", , "������ " & myErrCod
    End
End Sub

' ���� passErr=-11111 ��� �� ������� �� �������� ��� ���������
' ���� passErr=0  - ��������� ��������� "...WHERE..."
' ���� passErr<0  - ��������� ��� ���������, ����� 3262 Or 3261
' ���� passErr>0  - ��������� ��������� ������ ��� ������ � �����= passErr
' � ������ ��������� ���-� ���������� myExecute=0 ����� ������ ��� ������
' ������� myExecute >0; myExecute=-1 �������� ��� ������ �� ����������
'$odbc15!$
Function myExecute(myErrCod, sql, Optional passErr As Integer = -11111) As Integer
myExecute = -1
On Error GoTo ERR1
RETR:
'wrkDefault.BeginTrans ' ��� ������������� ��������� Execute �� ������ ��� wrkDefault.Rollback
myBase.Execute sql ', dbFailOnError  ' �������� Err ���� ��� ��� ����� ������� �������������
'Debug.Print sql
If myBase.RecordsAffected < 1 Then
  If passErr > 0 Or passErr = -11111 Then _
    MsgBox "��� �������, ��������������� ������� WHERE. �������� " & _
    "��������������!", , "Error " & myErrCod & " � myExecute:"
  Exit Function
End If
myExecute = 0
Exit Function

ERR1:
wrkDefault.rollback
cErr = Mid$(myErrCod, 3) ' - ������������� ������ ������ � Prior
    
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

'$odbc10$
Sub nextDay()  '�������� ������ �� ���� ����
Dim I As Integer, str As String, str1 As String, j As Integer, s As Double
Dim ch As String
'MsgBox "������� �� ����� ����"

'wrkDefault.BeginTrans

sql = "DELETE from OrdersInCeh WHERE (((Stat)='�����'));"
If myExecute("##63", sql, 0) > 0 Then GoTo ER1

'Set tbCeh = myOpenRecordSet("##63", "OrdersInCeh", dbOpenTable) 'dbOpenForwardOnly)
'If Not tbCeh Is Nothing Then
'  If Not tbCeh.BOF Then
'    tbCeh.MoveFirst
'    While Not tbCeh.EOF
'        If tbCeh!stat = "�����" Then
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
MsgBox "���� ���������� �� ����� ����!"
Exit Sub

ER1:
wrkDefault.rollback
End Sub

'$odbc08!$
Sub nextDayDetect() '�� ����� Orders.cmAdd_Click
Dim str As String ', intNum As Integer
Dim strNow As String, DateFromNum As String, dNow As Date

strNow = Format(Now, "dd.mm.yyyy")
curDate = strNow '��� ����� � �����
dNow = strNow
strNow = right$(Format(curDate, "yymmdd"), 5)
 
befDays = 0

wrkDefault.BeginTrans 'lock01
sql = "update system set resursLock = resursLock" 'lock02
myBase.Execute (sql) 'lock03

Set tbSystem = myOpenRecordSet("##91", "System", dbOpenTable) ', dbOpenForwardOnly)
If tbSystem Is Nothing Then Exit Sub

'������� lock01-04 ��������� �� ������������� ��������� � Sybase
'tbSystem.Edit '$odbs?$ ����������, ����� ������ �� ����� �� ������ �� ������
                      '������ �� Update

If tbSystem!resursLock = "nextDay" Then
   wrkDefault.rollback
   MsgBox "������ �������� ��������������! � ���� ����� �������� � ����������, " & _
    "�� c ������������ �����������������.", , "Error ��� �������� ���� �� ����� ����!"

Else
 str = tbSystem!lastPrivatNum
 If Len(str) > 6 Then
    DateFromNum = Mid$(str, 4, 2) & "." & Mid$(str, 2, 2) & ".200" & left$(str, 1)
    tmpDate = DateFromNum
    
    If tmpDate < dNow Then
        befDays = DateDiff("d", tmpDate, Now)
        GoTo AA
    End If
 Else ' �.�. ���� lastPrivatNum �� ���� ��� ����������������
AA:  'tbSystem!lastPrivatNum = strNow & "00"
    sql = "UPDATE SYSTEM SET lastPrivatNum = " & strNow & "00"
'    Debug.Print sql
    myBase.Execute (sql)
 End If
End If
 
If befDays <> 0 Then
'tbSystem!resursLock = "nextDay"
    myBase.Execute ("UPDATE SYSTEM SET resursLock = 'nextDay'")
        
End If
'tbSystem.Update
wrkDefault.CommitTrans 'lock04
tbSystem.Close

End Sub
'$NOodbc$
Public Sub quickSort(varArray As Variant, _
 Optional lngLeft As Long = dhcMissing, Optional lngRight As Long = dhcMissing)
Dim I As Long, j As Long, varTestVal As Variant, lngMid As Long

    If lngLeft = dhcMissing Then lngLeft = LBound(varArray)
    If lngRight = dhcMissing Then lngRight = UBound(varArray)
   
    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2
        varTestVal = varArray(lngMid)
        I = lngLeft
        j = lngRight
        Do
            Do While varArray(I) < varTestVal
                I = I + 1
            Loop
            Do While varArray(j) > varTestVal
                j = j - 1
            Loop
            If I <= j Then
                Call SwapElements(varArray, I, j)
                I = I + 1
                j = j - 1
            End If
        Loop Until I > j
        ' To optimize the sort, always sort the
        ' smallest segment first.
        If j <= lngMid Then
            Call quickSort(varArray, lngLeft, j)
            Call quickSort(varArray, I, lngRight)
        Else
            Call quickSort(varArray, I, lngRight)
            Call quickSort(varArray, lngLeft, j)
        End If
    End If
End Sub
'����� ��� quickSort
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
    If j = 0 Then ' ��� ����� ���
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


'****** ������������ ����� ***********
tmpSng = 0 '����� ����������� �����
sql = "SELECT Sum(Orders.workTime*OrdersInCeh.Nevip) AS nevip " & _
"FROM Orders INNER JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " & _
"WHERE (((Orders.StatusId)=1) AND ((Orders.CehId)=" & id & "));"
byErrSqlGetValues "##372", sql, tmpSng
s = 0 ' ���� ��������� �������
sql = "SELECT Sum(OrdersMO.workTimeMO) AS Sum_workTimeMO " & _
"FROM Orders INNER JOIN OrdersMO ON Orders.numOrder = OrdersMO.numOrder " & _
"WHERE (((OrdersMO.StatO)='� ������') AND ((Orders.CehId)=" & id & "));"
byErrSqlGetValues "##378", sql, s
tmpSng = tmpSng + s

sql = "SELECT Nstan" & Ceh(id) & ", KPD_" & Ceh(id) & " FROM System;"
byErrSqlGetValues "##379", sql, n, s

On Error GoTo EN1
'���������� ������ � ��� � ����.����
'!!! ���� ������ ����� �������� ����� ������� � ��� �� ������, �� ����.������
'������ ������, ��������� ����� �������� ���������� �� ����� �������� ���
'(� ��� ������� ��� - ����� ��������� �������� ����������)

sql = "SELECT Max(xDate) AS dLast FROM Itogi_" & Ceh(id) & ";"
byErrSqlGetValues "##407", sql, tmpStr
If tmpStr = Format(curDate, "yy.mm.dd") Then GoTo EN1 ' ������ ������� ��� ����

'numOrder = 0 ' ������� �������
sql = "INSERT INTO Itogi_" & Ceh(id) & " ( [xDate], numOrder, Virabotka ) " & _
"SELECT '" & tmpStr & "', 0, " & Round(oldRes * n, 2) & ";"
'MsgBox sql
myExecute "##408", sql
sql = "INSERT INTO Itogi_" & Ceh(id) & " ( [xDate], numOrder, Virabotka ) " & _
"SELECT '" & tmpStr & "', 1, " & s & ";"
myExecute "##409", sql
'���������� ����� ����������� �����(��������� � �������)
'numOrder = 2 ' ������� ����� ����������� �����
sql = "INSERT INTO Itogi_" & Ceh(id) & " ( [xDate], numOrder, Virabotka ) " & _
"SELECT '" & Format(curDate, "yy.mm.dd") & "', 2, " & tmpSng & ";"
myExecute "##410", sql

'��������� ������ ������� ���������� ������
sql = "DELETE from Itogi_" & Ceh(id) & _
" WHERE (((xDate)<'" & Format(DateAdd("m", -1, curDate), "yy.mm.dd") & "'));"
myExecute "##411", sql, 0
EN1:
replaceResurs = True
On Error Resume Next
End Function

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

Function StatParamsLoad(row As Long)
Dim s As Double, log As String, str As String


 log = Format(Now(), "dd.mm.yy hh:nn") & vbTab & Orders.cbM.Text & " " & gNzak ' ������ vbTab
 str = status(tqOrders!StatusId): log = log & " " & str
 Orders.Grid.TextMatrix(row, orStatus) = str
 
 str = LoadDate(Orders.Grid, row, orDataVid, tqOrders!outDateTime, "dd.mm.yy")
 If str <> "" Then log = log & " Out=" & str
 str = LoadDate(Orders.Grid, row, orVrVid, tqOrders!outDateTime, "hh")
 If str <> "" Then log = log & "_" & str
 
 str = LoadNumeric(Orders.Grid, row, orVrVip, tqOrders!workTime, , "#0.0")
 log = log & " ��.���=" & str
 
 Orders.Grid.TextMatrix(row, orProblem) = tqOrders!problem
 
 str = LoadDate(Orders.Grid, row, orDataRS, tqOrders!dateRS, "dd.mm.yy")
 If str <> "" Then log = log & " ��=" & str
 
 gNzak = tqOrders!numorder
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
 
 If IsNull(tqOrders!statM) Then
    Orders.Grid.TextMatrix(row, orM) = ""
 Else
    Orders.Grid.TextMatrix(row, orM) = tqOrders!statM
    log = log & " M�(" & tqOrders!statM & "):" & str ' ���� ���
 End If
 If IsNull(tqOrders!statO) Then
    Orders.Grid.TextMatrix(row, orO) = ""
    Orders.Grid.TextMatrix(row, orOVrVip) = ""
 Else
    Orders.Grid.TextMatrix(row, orO) = tqOrders!statO
    If tqOrders!statO = "� ������" Or tqOrders!statO = "�����" Then
        If IsNull(tqOrders!DateTimeMO) Then
            msgOfZakaz "##313", "����������� '���� MO'."
            str = " !��� ���� MO! "
        End If
        log = log & " O�(" & tqOrders!statO & "):" & str ' ���� ���
        If IsNull(tqOrders!workTimeMO) Then
            msgOfZakaz "##314", "����������� '����� ���������� MO'."
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
Dim I As Integer

I = Grid.Height \ Grid.RowHeight(1) - 1 ' ������� ��������� �����
I = numRow - I \ 2 ' � �����
If I < 1 Then I = 1
Grid.TopRow = I
End Sub

'��� �-� �.�������� � startDay() � getNextDay() � getPrevDay()
' ���������� �������� �� ����. ���
Function getWorkDay(offsDay As Integer, Optional baseDate As String = "") As Integer
Dim I As Integer, j As Integer, step  As Integer
getWorkDay = -1
If baseDate = "" Then
    tmpDate = curDate
Else
    If Not IsDate(baseDate) Then Exit Function
    tmpDate = baseDate
End If

step = 1
If offsDay < 0 Then step = -1

j = 0: I = 0
While step * j < step * offsDay '
    I = I + step
    day = Weekday(DateAdd("d", I, tmpDate))
    If Not (day = vbSunday Or day = vbSaturday) Then j = j + step
Wend
getWorkDay = I
tmpDate = DateAdd("d", I, tmpDate)

End Function

Function startDays() As Integer
Dim I As Integer, j  As Integer, k   As Integer
ReDim Preserve stDays(befDays + 1)

For k = 0 To befDays '    *********************************************

j = 0
I = 1
While j < 3 '         ������� �������� �������� �������� (3-� ����)

    day = Weekday(DateAdd("d", k + I - befDays, curDate))
'    day = Weekday(CurDate - befDays + K + I)
    If Not (day = vbSunday Or day = vbSaturday) Then j = j + 1
    I = I + 1
Wend
stDays(k) = I + k ' "+k" �.�. ���� ��������� ���������� befDays ���� �����

Next k          '       ***********************************************
dayMassLenght (stDays(befDays) + 1)
startDays = stDays(befDays) - befDays ' ��� �������, ������� ��� �1
End Function

Sub statistic(Optional year As String = "")
Dim nRow As Long, nCol As Long, str As String, I As Integer, j As Integer
Dim iMonth As Integer, iYear As Integer, iCount As Integer, strWhere As String
Dim nMonth As Integer, nYear As Integer, mCount As Integer, lastCol As Integer
Dim wtSum As Double, paidSum As Double, orderSum As Double, visits As Integer, visitSum As Integer
Dim year01 As Integer, year02 As Integer, year03 As Integer, year04 As Integer
Dim errCurYear As Integer, errBefYear As Integer, whereByTemaAndType As String


errCurYear = 0:   errBefYear = 0

whereByTemaAndType = ""
If year = "" Then
 str = Reports.tbStartDate.Text
 Report.laHeader.Caption = "���������� ��������� ���� �� ������ � " & str & _
                " �� " & Reports.tbEndDate.Text
 nMonth = left$(str, 2)
 nYear = right$(str, 4)
 mCount = DateDiff("m", str, Reports.tbEndDate.Text) + 1

 str = "|<�������� �����|^� |K��������|������"
 iCount = mCount
 lastCol = 5 ' � 2� ������
 iMonth = nMonth
 Do
    If iMonth = 13 Then iMonth = 1
    str = str & "|" & Format(iMonth, "00")
    iMonth = iMonth + 1
    lastCol = lastCol + 1
    iCount = iCount - 1
 Loop While iCount > 0
 str = str & "|�����|��.���|��������|��������"
 Report.Grid.FormatString = str
 Report.Grid.ColWidth(0) = 0
 Report.Grid.ColWidth(1) = 1875
 Report.Grid.ColWidth(3) = 375
 
 Report.nCols = lastCol + 2
 If Report.Regim = "KK" Then
    strWhere = "WHERE (((GuideFirms.Kategor)='�'));"
    Report.Grid.ColWidth(4) = 0
 ElseIf Report.Regim = "RA" Then
    strWhere = "WHERE (((GuideFirms.Kategor)='�' Or (GuideFirms.Kategor)='�'));"
    Report.Grid.ColWidth(4) = 375
 Else
    Exit Sub
 End If
 
 If Reports.lbType.Text <> "���" Then
    lbToOrSqlWhere Reports.lbType, orType
    whereByTemaAndType = "(" & orSqlWhere(orType) & ") AND "
    Report.laHeader.Caption = Report.laHeader.Caption & _
    "  (������ ��������� '" & Reports.lbType.Text & "'"
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
            Report.laHeader.Caption = Report.laHeader.Caption & "  �� ����:" & str
        End If
    End If
        Report.laHeader.Caption = Report.laHeader.Caption & ")"
 End If
 
 nRow = 1
 'sql = "SELECT GuideFirms.FirmId, GuideFirms.Name, GuideFirms.Kategor, " & _
 "GuideFirms.year01, GuideFirms.year02, GuideFirms.year03, GuideFirms.year04, " & _
 "GuideFirms.Sale, GuideManag.Manag FROM GuideFirms LEFT JOIN GuideManag " & _
 "ON GuideFirms.ManagId = GuideManag.ManagId " & strWhere
Else '�������� ���������� ��� �����-�� ����
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
 nCol = 5 ' � 2� ������
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
            str = tbOrders!numorder
          If year <> "" Then
            If iYear = lastYear - 3 Then
                year01 = year01 + 1 '�� ���-��
            ElseIf iYear = lastYear - 2 Then
                year02 = year02 + 1
            ElseIf iYear = lastYear - 1 Then
                year03 = year03 + 1
            ElseIf iYear = lastYear Then
                year04 = year04 + 1
            End If
          End If
          visits = visits + 1
          wtSum = wtSum + tbOrders!workTime
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
 If year <> "" Then '�������� ����������
    tbFirms.Edit
    I = getLockYear '�� ������������� ����, ��� ������������ � ��������� ����
'������ ���� �� �������������, �.�. ���������� ��������� ��� ������ ����
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
'  If errBefYear > 0 Then !!!�� �������
'     MsgBox "� ������� ����� ���������� " & errBefYear & " ���� � ������� " & _
'     "������������ ����������� ���������.  ��� ������ ���������.", , "���������� ������"
'  End If
'  If errCurYear > 0 Then
'     MsgBox "� ������� ���� ���������� " & errCurYear & " ���� � ������� " & _
'     "������������ ����������� ���������.  ��� ������ ���������.", , "���������� ������"
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
            MsgBox "�������� ������ ���� � ��������� �� " & minVal & _
            "  �� " & maxVal, , "Error"
            checkNumeric = False
        End If
    ElseIf Not IsMissing(minVal) Then
        If minVal > tmpSng Then
            MsgBox "�������� ������ ���� ������ " & minVal
            checkNumeric = False
        End If
    End If
Else
    MsgBox "������������ ��������", , "Error"
    checkNumeric = False
End If
End Function

'� ������ true ����� ���������� ���� � tmpDate
Function isDateTbox(tBox As TextBox, Optional fryDays As String = "") As Boolean
Dim str As String

isDateTbox = True
str = tBox.Text
If str = "" Then
        MsgBox "��������� ���� ����!", , "������"
Else
'    If Not IsDate(str) Then
'    If Len(str) <> 8 Or Not IsDate(str) Then
'        MsgBox "�������� ������ ����", , "������"
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
                    If MsgBox(str & " - �������� ����. ����������?", vbYesNo, _
                    "��������������!") = vbYes Then Exit Function
                Else
                    Exit Function
                End If
            End If
        Else
            MsgBox "�������� ������ ���� ��� ��� � ����� ����� �� ���������� ", , "������"
        End If
'    End If
End If
 '   tBox.Text = oldValue
tBox.SetFocus
tBox.SelStart = 0
tBox.SelLength = Len(tBox.Text)
isDateTbox = False
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

'�� ��������� ������������ ��������, ��� �����, ��� �����
'�������� ���������. �  ������� ��� ���� error?
Function ValueToTableField(myErrCod As String, value As String, table As String, _
field As String, Optional by As String = "", Optional numorder As Variant) As Integer
Dim sql As String, byStr As String  ', numOrd As String


ValueToTableField = False
'If value = "" Then value = Chr(34) & Chr(34)
If value = "" Then value = "''"
If by = "" Then
    Dim nzak As String
    If IsMissing(numorder) Then
        nzak = gNzak
    Else
        nzak = numorder
    End If
        
    byStr = ".numOrder)= " & nzak
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

If left$(myErrCod, 1) = "W" Then
    myErrCod = Mid$(myErrCod, 2)
    ValueToTableField = myExecute(myErrCod, sql, 0) '�� �������� ���� �� WHERE
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

If Frm.Grid5.TextMatrix(v_row, prType) = "�������" Then
    str = Frm.Grid5.TextMatrix(v_row, prName) '
    I = InStr(str, "/")
    prExt = 0: If I > 1 Then prExt = left$(str, I - 1)   '����� ��������
    gProductId = Frm.Grid5.TextMatrix(v_row, prId)
Else
    gNomNom = Frm.Grid5.TextMatrix(v_row, prId)
End If
End Sub

Function getNevip(day As Integer)
sql = "SELECT Sum(Orders.workTime*OrdersInCeh.Nevip) AS wSum " & _
"FROM Orders INNER JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " & _
"WHERE (((DateDiff(day,'" & Format(curDate, "yyyy-mm-dd") & _
"',Orders.outDateTime))=" & day - 1 & " AND (Orders.CehId)=" & cehId & "));"
'MsgBox sql
getNevip = 0
byErrSqlGetValues "W##382", sql, getNevip
End Function

Sub addDays(outDay As Integer)
Dim j As Integer
        If maxDay < outDay Then
            dayMassLenght outDay + 1 '���� ������ , ������������ �����������
            For j = maxDay + 1 To outDay '����� ���
                delta(j) = 0
            Next j
            maxDay = outDay
        End If
End Sub

Function getLockYear() As Integer
getLockYear = Format(begDate, "yyyy")
If Format(begDate, "dd.mm") = "01.01" Then _
    getLockYear = getLockYear - 1 '�������, ��� ���� ��� �� ����������� � ��������� ����
End Function

Function getYearField(checkDate As Date) As String '$$3
Dim I As Integer, lockYear As Integer

lockYear = getLockYear
I = Format(checkDate, "yyyy")
'If I <= lockYear Then
'    getYearField = "lock" '���� ��� ����������� � ��������� ����
'    Exit Function
'End If
I = I - lastYear + 4 '����� �������
If I < 1 Then     '���� ��� �� ��������� 3 ����, �� � ����
    getYearField = "year01"
Else
    getYearField = "year" & Format(I, "00")
End If
End Function


Sub visits(oper As String, Optional firm As String = "") '$$3
Dim str As String, I As Integer, statId As Integer

sql = "SELECT Orders.inDate, Orders.StatusId , Orders.FirmId From Orders " & _
"WHERE (((Orders.numOrder)=" & gNzak & "));"
'MsgBox sql
If Not byErrSqlGetValues("##88", sql, tmpDate, statId, I) Then GoTo ER1

If I = 0 Then Exit Sub
If firm <> "" And (statId = 0 Or statId = 7) Then Exit Sub ' ���� ������ �����

str = getYearField(tmpDate)

'If str = "lock" Then Exit Sub ' ���� ��� ����������� � ��������� ���� , �� ��� �� �������������

sql = "UPDATE GuideFirms SET GuideFirms." & str & " = [GuideFirms].[" & _
str & "] " & oper & " 1  WHERE (((GuideFirms.FirmId)=" & I & "));"
'Debug.Print sql
I = myExecute("##87", sql, -143)

'If I <> 3061 And I <> 0 Then '3061 - ������� ����� ���� ���(��� ���) ��� � ����
If I = -2 Then '3061 - ������� ����� ���� ���(��� ���) ��� � ����
ER1:    MsgBox "������ ��������� ��������� ����. �������� ��������������!", , "Error-87"
End If
End Sub

Sub zagruzFromCeh(Optional passZakazNom As String = "")
Dim outDay As Integer, j As Integer, passSql As String, str As String

If IsNumeric(passZakazNom) Then
    passSql = "((OrdersInCeh.numOrder)<>" & passZakazNom & ") AND "
Else
    passSql = ""
End If

'    "OrdersInCeh.numOrder, OrdersInCeh.VrVipParts, OrdersInCeh.rowLock "
sql = "SELECT Orders.outDateTime, Orders.StatusId, " & _
    "OrdersInCeh.numOrder " & _
    "FROM OrdersInCeh LEFT JOIN Orders " & _
    "ON OrdersInCeh.numOrder = Orders.numOrder " & _
    "WHERE (" & passSql & "((Orders.CehId)=" & cehId & "));"

Set tbCeh = myOpenRecordSet("##14", sql, dbOpenForwardOnly)
If tbCeh Is Nothing Then Exit Sub

'1:MaxDay = 0
If Not tbCeh.BOF Then
    While Not tbCeh.EOF
        isLive = False ' ������� �����
        If tbCeh!StatusId = 1 Then isLive = True
        outDay = DateDiff("d", curDate, tbCeh!outDateTime) + 1
        If outDay < 1 Then outDay = 1
                
        addDays outDay '1:��������� ���,  �.�. ����  ��� ���.������ �����
                       '  ���������  ������  ��� stDay � rMaxDay
        tbCeh.MoveNext
    Wend
End If
tbCeh.Close
End Sub


Function beNaklads(Optional reg As String = "") As Boolean
beNaklads = True
Dim s As Double
'��������
sql = "SELECT Sum(sDMC.quant) AS Sum_quant From sDMC " & _
"WHERE (((sDMC.numExt)< 254) AND ((sDMC.numDoc)=" & numDoc & "));"
If Not byErrSqlGetValues("##140", sql, s) Then Exit Function
If s > 0.005 Then ' ���-�� ��������
    If reg = "" Then
        MsgBox "�� ����� ������ ������������ ���������, ������� �������� " & _
        "�������� ������. ���� ��������� ���-�� ���������, �� ������ ���� " & _
        "������� ��� ��������� � ������.", , "�������������� ���������!"
    End If
Else
    sql = "SELECT Sum(curQuant) AS Sum_curQuant " & _
    "From sDMCrez WHERE (((numDoc)=" & numDoc & "));"
    If Not byErrSqlGetValues("##367", sql, s) Then Exit Function
    If s > 0.005 Then ' ���-�� ��������
        If reg = "" Then
            MsgBox "�� ����� ������ � ���� ��� �������� ��������� (��������� " & _
            "������� '���-��') � �������� ��� ��� �����������. ������ ��� " & _
            "������ �������� ������, ��� ������ ���������� ����������, � ����� " & _
            "�������� ������� '���-��'.", , "�������������� ���������!"
        End If
    Else
        beNaklads = False
    End If
End If

End Function

'$odbc08$ Function docLock �� ���-��
'--------------------------------------------------------------------------

Sub getNakladnieList(Optional from As String = "") '
Dim I As Integer, str As String, l As Long

'pusto=-1 ���� ���� �� ������ �������� ��� �� � ����� ��������� (����� pusto=0)
'�����, �.�. ��� ���� delta=Null � �� quantity
If from = "Buh" Then str = "3" Else str = "2" '��� ���� ������ �� prior_������, ��� ����������� ���-�� � ����
sql = "SELECT numDoc, Max([quantity]-" & _
"IsNull([Sum_quant],0)) AS delta, Min(IsNull(" & _
"[Sum_quant],0)) AS pusto  From wCloseNomenk" & str & " GROUP BY numDoc ORDER BY numDoc;"
'Debug.Print sql
Set tbDMC = myOpenRecordSet("##142", sql, dbOpenDynaset)
If tbDMC Is Nothing Then Exit Sub

I = 0
ReDim tmpL(0)
While Not tbDMC.EOF
 
  If tbDMC!pusto = -1 Then GoTo AA
  If tbDMC!delta < 0.005 Then ' ��� �������� �������
        gNzak = -tbDMC!numDoc
  Else
AA:     gNzak = tbDMC!numDoc
    
    If from = "ceh" Then ' ��� ���� ��������� ��� � ���������
      sql = "SELECT numOrder From xEtapByNomenk Where (((numOrder) = " & gNzak & _
      ")) UNION ALL SELECT numOrder From xEtapByIzdelia " & _
      "WHERE (((numOrder)= " & gNzak & "));"
      If Not byErrSqlGetValues("W##352", sql, l) Then GoTo NXT
      If l > 0 Then '��� ���������� ������ ������ ��� �������
          If predmetiIsClose("etap") Then GoTo NXT '�������� �� ����� ����������
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


'reg=""  - �������� ������� �������� ���������
'reg = "prev" - ��������, ��� ������� ����� �� ����.�����, �� ������
'����� - ��������, ��� ������� �� �����, �� �����
Function predmetiIsClose(Optional reg As String = "") As Boolean
Dim I As Integer, s As Double

#If onErrorOtlad Then
    On Error GoTo errMsg
    GoTo START
errMsg:
    MsgBox Error, , "������  " & Err & " � �\� predmetiIsClose" '
    End
START:
#End If

predmetiIsClose = False

If Not sProducts.zakazNomenkToNNQQ() Then Exit Function
For I = 1 To UBound(NN)
    sql = "SELECT Sum(quant) AS Sum_quant From sDMC " & _
    "WHERE (((sDMC.numDoc)=" & gNzak & ") AND ((nomNom)='" & NN(I) & "'));"
    If Not byErrSqlGetValues("##164", sql, s) Then Exit Function
    If reg = "prev" Then
        If Abs(QQ3(I) - s) > 0.005 Then Exit Function
    ElseIf reg = "" Or QQ2(0) = 0 Then '����� �� �� ���� ��� ��� ���������� ������
        If QQ(I) - s > 0.005 Then Exit Function
    Else
        If QQ2(I) - s > 0.005 Then Exit Function
    End If
Next I
predmetiIsClose = True
End Function


Function PrihodRashod(reg As String, skladId As Integer) As Double
Dim qWhere As String, s As Double

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
Debug.Print sql
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

If sId < -1000 And dId < -1000 Then ' ��� ������������ �� ������������
        ostatCorr = True
Else
    '������������ �������
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

'���-�� ��� ������������ ���������, � ����� �������� �� ����� ��������
'��������� � Otgruz.frm
Sub loadPredmeti(Frm As Form, Optional reg As String = "", Optional needToRefresh As Boolean = False)
Dim I As Integer

Screen.MousePointer = flexHourglass
Frm.Grid5.Visible = False
Frm.quantity5 = 0
I = 0: If reg = "fromOtgruz" Then I = 1

clearGrid Frm.Grid5, 1 + I

'******** ������� ************************************************
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
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prType) = "�������"
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prName) = getStrPrEx(tbNomenk!prName, tbNomenk!prExt)
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prDescript) = tbNomenk!prDescript
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prEdizm) = "��."
    If Not IsNull(tbNomenk!cenaEd) Then
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prCenaEd) = Round(tbNomenk!cenaEd, 2)
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prSumm) = _
                                Round(tbNomenk!cenaEd * tbNomenk!quant, 2)
    End If
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prQuant) = tbNomenk!quant
' ��� ��������� ��������� � ��� ���-�� (��. ����)
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

'****** ������������ ********************************************
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
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prType) = "������������"
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prName) = tbNomenk!cod
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prDescript) = _
        tbNomenk!nomName & " " & tbNomenk!Size
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prEdizm) = tbNomenk!ed_Izmer
    If Not IsNull(tbNomenk!cenaEd) Then
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prCenaEd) = Round(tbNomenk!cenaEd, 2)
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prSumm) = _
                                Round(tbNomenk!cenaEd * tbNomenk!quant, 2)
    End If

    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prQuant) = tbNomenk!quant

    If reg = "fromOtgruz" Then
        Otgruz.getOtgrugeno Frm.quantity5 + I, "byNomenk"
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

If Frm.quantity5 > 0 Then
    Frm.Grid5.row = Frm.quantity5 + 1 + I
    Frm.Grid5.col = prQuant
    Frm.Grid5.Text = "�����:"
    Frm.Grid5.col = prSumm
    Frm.Grid5.Text = sProducts.saveOrdered(needToRefresh)
    Frm.Grid5.CellFontBold = True
    If reg = "fromOtgruz" Then
        Frm.Grid5.col = prOutSum
        Frm.Grid5.Text = Otgruz.saveShipped
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
lockSklad = True: Exit Function '!!! ��������� ����� ���������, ����� ��� ��� �������
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
        
        If MsgBox("������ � �������� �� ������ �������� ����� ���������� '" & _
        str & "'. ��������� ������� ��� ���������� � ��������������.", _
        vbRetryCancel, "��� ������� !!!") = vbRetry Then
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
    
Function orderUpdate(myErrCod As String, value As String, table As String, _
field As String, Optional by As String = "", Optional numorder As Variant) As Integer
Dim nzak As String
    If IsMissing(numorder) Then
        nzak = gNzak
    Else
        nzak = numorder
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
    For irow = 1 To Grid.rows - 1
        If Grid.TextMatrix(irow, orNomZak) = nzak Then
            searchZakRow = irow
            Exit Function
        End If
    Next irow

End Function

Sub loadSeria(ByRef p_tv As TreeView)
Dim Key As String, pKey As String, k() As String, pK()  As String
Dim I As Integer, iErr As Integer
bilo = False
sql = "SELECT sGuideSeries.*  From sGuideSeries ORDER BY sGuideSeries.seriaId;"
Set tbSeries = myOpenRecordSet("##110", sql, dbOpenForwardOnly)
If tbSeries Is Nothing Then myBase.Close: End
If Not tbSeries.BOF Then
 p_tv.Nodes.Clear
 Set Node = p_tv.Nodes.Add(, , "k0", "���������� �� ������")
 Node.Sorted = True
 
 ReDim k(0): ReDim pK(0): ReDim NN(0): iErr = 0
 While Not tbSeries.EOF
    If tbSeries!seriaId = 0 Then GoTo NXT1
    Key = "k" & tbSeries!seriaId
    pKey = "k" & tbSeries!parentSeriaId
    On Error GoTo ERR1 ' ��������� ������ ������
    Set Node = p_tv.Nodes.Add(pKey, tvwChild, Key, tbSeries!seriaName)
    On Error GoTo 0
    Node.Sorted = True
NXT1:
    tbSeries.MoveNext
 Wend
End If
tbSeries.Close

While bilo ' ���������� ��� �������
  bilo = False
  For I = 1 To UBound(k())
    If k(I) <> "" Then
        On Error GoTo ERR2 ' ��������� ��� ������
        Set Node = p_tv.Nodes.Add(pK(I), tvwChild, k(I), NN(I))
        On Error GoTo 0
        k(I) = ""
        Node.Sorted = True
    End If
NXT:
  Next I
Wend
p_tv.Nodes.item("k0").Expanded = True
Exit Sub
ERR1:
 iErr = iErr + 1: bilo = True
 ReDim Preserve k(iErr): ReDim Preserve pK(iErr): ReDim Preserve NN(iErr)
 k(iErr) = Key: pK(iErr) = pKey: NN(iErr) = tbSeries!seriaName
 Resume Next

ERR2: bilo = True: Resume NXT

End Sub

