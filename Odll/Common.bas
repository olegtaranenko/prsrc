Attribute VB_Name = "Common"
Option Explicit
'���e��\��������\��������\��������� ����������:
' - onErrorOtlad = 1 ' ����� ������ err

'���������� ��� ������������
Public nomnomCache As NomnomFactory
'Public nomnomCache As New NomnomFactory
Dim cacheResetTimer As Timer

' ������ �������� (�-�� �����) ��� ������ ���������.
' �������� � ����������, ������������ � ������ Nakladna.frm

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
Public EquipFullName() As String
Public Werk() As String
Public werkSourceId() As Integer

Public gWerkId As Integer
Public gEquipId As Integer
Public Const lenStatus = 20

Public Status(lenStatus) As String
Public lenProblem As Integer
Public Problems(20) As String
Public manId() As Integer '$$7
Public Manag() As String  '
Public Managers() As MapEntry

Public insideId() As String
Public Const begWerkProblemId = 10 ' ������ ������� ������� � �����������
Public maxDay As Integer ' ����� ���� � �������
Public befDays As Integer ' ����� ���� �� ���� ������� (����� ��������� ����), ����������� ��� ������� ����� ������� ����, � ���� �� ���������� ������������ System.lastNumorder
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
Public sessionCurrency As Integer
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
Public gAsWhole As Integer ' ��� ������: ����������, � ����� ������� ���������� �������� ������������


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
Public KPD As Double
Public newRes As Double ' ����� �� ���������
Public nr As Double ', dr As Double '��������� ���. � ���. �������
Public isLive As Boolean ' ���� - ����� �����
Public zagAll As Double, zagLive As Double

Public Table As Recordset '
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
Public orRemark As Integer, orSize As Integer, orPlaces As Integer
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
Public Const chEquipId = 15
Public Const chWerkId = 16
Public Const chRemark = chIzdelia

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
Public Const prVes = prEtap
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

Public Const CC_RUBLE As Integer = 1
Public Const CC_UE As Integer = 2

' �� ������� ����� ����������� ������ �������, ���� ������� �����
Public Const ColWidthForRuble As Single = 1.3

Function dummy()
    Dim Left, StatusId, Outdatetime, Rollback, IsEmpty, ExeName
    Dim WorktimeMO

End Function

Function tuneCurencyAndGranularity(tunedValue, currentRate, valueCurrency As Integer, Optional quantity As Double = 1, Optional perlist As Long = 1) As Double
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


'���� ������ �������� ="W.." - �� �������� Err �� �����-� Where, � ���
'��������� ��������, ���� ��� ���� ��� ���� ��� ��������� ��������, �� � sql
'�. ������ ��������� "1" � ������� �� � I. ����� ���� I=0 �� ���� Err Where
'$odbc15$
Function byErrSqlGetValues(ParamArray val() As Variant) As Boolean
Dim tabl As Recordset, I As Integer, maxi As Integer, str As String, c As String

byErrSqlGetValues = False
maxi = UBound(val())
If maxi < 1 Then
    wrkDefault.Rollback
    MsgBox "���� ���������� ��� �\� byErrSqlGetValues()"
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
'        msgOfEnd CStr(val(0)), "��� ������� ��������������� Where."
        wrkDefault.Rollback
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
Dim I As Long
    noClick = True
    Grid.row = row
    For I = 0 To Grid.Cols - 1
        Grid.col = I
        If I > 0 Then Grid.CellBackColor = Grid.BackColor
        Grid.CellForeColor = Grid.ForeColor
        Grid.CellFontStrikeThrough = False
        Grid.TextMatrix(row, I) = ""
    Next I
    Grid.col = 1
    noClick = False
End Sub

Sub colorGridRow(Grid As MSFlexGrid, row As Long, color As Long)
Dim IL As Long
    Grid.row = row
    For IL = 0 To Grid.Cols - 1
        Grid.col = IL
        If IL > 0 Then Grid.CellBackColor = color
    Next IL
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
Dim S As Double
    S = Timer
    While Timer - S < tau ' 1 ���
        DoEvents
    Wend

End Sub

Sub delZakazFromReplaceRS()
    sql = "DELETE From ReplaceRS " & "WHERE numOrder = " & gNzak
    myExecute "##79", sql, 0 ' �������, ���� ����
End Sub


Sub exitAll()
    If isOrders Then Unload Orders
    If isWerkOrders Then Unload WerkOrders
    If isZagruz Then Unload Zagruz
    If isFindFirm Then Unload FindFirm
    
    If sDocs.isLoad Then Unload sDocs
    
    If cfg.isLoad Then Unload cfg '$$2

End Sub

Function findValInCol(Grid As MSFlexGrid, Value As Variant, col As Integer) As Boolean
Dim IL As Long
findValInCol = False
For IL = 1 To Grid.Rows - 1
    If Value = Grid.TextMatrix(IL, orNomZak) Then
        Grid.TopRow = IL
        Grid.row = IL
        findValInCol = True
        Exit For
    End If
Next IL

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

'$odbc08$
Function existValueInTableFielf(ByVal Value As Variant, tabl As String, Field) As Boolean
Dim Table As Recordset

existValueInTableFielf = False

If Not IsNumeric(Value) Then Value = "'" & Value & "'"

sql = "SELECT " & Field & " From " & tabl & " WHERE (((" & Field & ") = " & _
Value & "));"
'MsgBox sql
Set Table = myOpenRecordSet("##390", sql, dbOpenForwardOnly)
'If table Is Nothing Then myBase.Close: End

If Not Table.BOF Then existValueInTableFielf = True

Table.Close

End Function

'��� ������ �� ������ �����, � ���� �� ��������� - ��� ���������� ����������
Function yymmdd(dateStr As String) As String
yymmdd = Right$(dateStr, 2) & "." & Mid$(dateStr, 4, 2) & "." & Left$(dateStr, 2)
End Function


Function getValueFromTable(tabl As String, Field As String, Where As String) As Variant
Dim Table As Recordset

getValueFromTable = Null
sql = "SELECT " & Field & " as fff  From " & tabl & _
      " WHERE " & Where
Set Table = myOpenRecordSet("##59.1", sql, dbOpenForwardOnly)
If Table Is Nothing Then Exit Function
If Not Table.BOF Then getValueFromTable = Table!fff
Table.Close
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

Function getStrPrEx(Name As String, ext As Integer) As String
If ext = 0 Then
    getStrPrEx = Name
Else
    getStrPrEx = ext & "/ " & Name
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
Function getResurs(EquipId As Integer) As Integer
Dim I As Integer, J As Integer, rMaxDay As Integer, S As Double
' rMaxDay - Resource max day - ������������ �������� �� ������� ResursCEH (CO2, etc)

Set tbSystem = myOpenRecordSet("##93", "select * from GuideResurs where equipId = " & EquipId, dbOpenForwardOnly)
If tbSystem Is Nothing Then myBase.Close: End
KPD = tbSystem!KPD
Nstan = tbSystem!Nstan
newRes = tbSystem!newRes
tbSystem.Close

sql = "SELECT nomRes from Resurs where equipId = " & EquipId & " ORDER BY xDate"
Set Table = myOpenRecordSet("##10", sql, dbOpenForwardOnly)
'If table Is Nothing Then Exit Function

'j = -1
'If flEdit <> "" Then _
'    j = Mid$(Zagruz.lv.SelectedItem.key, 2)
rMaxDay = 0
On Error GoTo ERR1
'If Not table.BOF Then
' table.MoveFirst
' Do
While Not Table.EOF
    rMaxDay = rMaxDay + 1
    If rMaxDay = J Then
'        table.Edit
'            table!nomRes = Zagruz.tbMobile.Text
'        table.Update
    End If
    nomRes(rMaxDay) = Table!nomRes
    Table.MoveNext
' Loop While Not table.EOF
Wend
Table.Close
'End If

addDays max(stDay, rMaxDay) '��������� ���, ���� ���� ������ ���� �������
                            '����� ��� stDay ��� rMaxDay(�������� ������� ���)
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
If Frm.ckEndDate.Value > 0 Then ckEnd = True  '�� ��� ��� �� �����������
If Frm.ckStartDate.Value > 0 And Frm.ckStartDate.Visible Then ckStart = True
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
    tb.Left = Grid.CellLeft + Grid.Left
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
    orSqlWhere(col) = "(s.Status)<>'������'"
Else
End If
End Function


Sub listBoxInGridCell(lb As listBox, Grid As MSFlexGrid, Optional sel As String = "", Optional ColWidth As Long = -1)
Dim I As Integer
    If Grid.CellTop + lb.Height < Grid.Height Then
        lb.Top = Grid.CellTop + Grid.Top
    Else
        lb.Top = Grid.CellTop + Grid.Top - lb.Height + Grid.CellHeight
    End If
    lb.Left = Grid.CellLeft + Grid.Left
    If ColWidth > 0 Then
        lb.Width = ColWidth
    End If
    lb.ListIndex = 0
    
    Dim saveNoClick As Boolean
    saveNoClick = noClick
    If sel <> "" Then
        For I = 0 To lb.ListCount - 1 '
            If Grid.Text = lb.List(I) Then
                noClick = True
                lb.ListIndex = I '�������� ������ onClick
                noClick = False
                Exit For
            End If
        Next I
    End If
    noClick = saveNoClick
    
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





'$odbc10$
Sub Main()
Dim I As Integer, S As Double, str As String, str1 As String, str2 As String
Dim isXP As Boolean
If App.PrevInstance = True Then
    MsgBox "��������� ��� ��������", , "Error"
    End
End If

' ��� ��� ������������
Set nomnomCache = New NomnomFactory

'nomnomCache.CompareMode = TextCompare
'nomnomCache.Add "", New Nomnom

ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): ReDim QQ3(0) '����� Ubound ������� �� ������ Err

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
    
    ' ������������ ��� ����������� ����, ��������� �� �� ������� ������������� ������ �������� ��� ���.
    sql = "create variable @issueMarker varchar(32)"
    If myExecute("##issueMarker", sql, 0) = 0 Then End

    sql = "create variable @managerId varchar(20)"
    If myExecute("##@managerId", sql, 0) = 0 Then End
    
mainTitle = getMainTitle
gAsWhole = -1

Set sc = CreateObject("ScriptControl")
sc.Language = "VBScript"


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

'�� ������ ������� ����. 3� �������
nextDayDetect ' ����� �������-�� CurDate
stDay = startDays() ' � �.�. ������������� ��������� ����������� dayMassLenght
If befDays <> 0 Then nextDay ' ������� ���� �� ����� ����

checkNextYear '$$3 ���� �������� ��� - �������� ���������� ���������

' ��������� ��� �����*********************************************
'If Not (dostup = "c" Or dostup = "y") Then
If dostup = "a" Or dostup = "m" Or dostup = "" Or dostup = "b" Then
 'logFile = "C:\Windows\Orders" ' ��� ����������
 logFile = App.path & "\" & App.ExeName
 str2 = logFile & "$.log" ' ��������� ����
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

Set Table = myOpenRecordSet("##05", "StatusGuide", dbOpenForwardOnly)
While Not Table.EOF
    If Table!StatusId > lenStatus + 1 Then
        MsgBox "Err � Orders\FormLoad"
        End
    End If
    Status(Table!StatusId) = Table!Status
    Table.MoveNext
Wend
Table.Close

Set Table = myOpenRecordSet("##04", "GuideProblem", dbOpenForwardOnly)
'If table Is Nothing Then myBase.Close: End

For I = 0 To 20
    Problems(I) = "no"
Next I
lenProblem = -1
CC:
    If lenProblem < Table!ProblemId Then lenProblem = Table!ProblemId
    If Table!ProblemId > 20 Then
        MsgBox "����� ������� � ���� ��������� 20"
        End
    End If
    Problems(Table!ProblemId) = Table!problem
    Table.MoveNext
    If Not Table.EOF Then GoTo CC
Table.Close

'�������� ������ ���������� ������ � ������
CheckIntegration

    gWerkId = getEffectiveSetting("werkId")
    gEquipId = getEffectiveSetting("equipId")


' ������ ���� � Orders.Form.Load
Dim mysql As String

mysql = "select werkId, werkCode, werkSourceId from GuideWerk order by 1"

Set Table = myOpenRecordSet("##72.2", mysql, dbOpenForwardOnly)
If Table Is Nothing Then myBase.Close: End
' ReDim werkSourceId(Orders.lbWerk.ListCount - 1)
I = 1
While Not Table.EOF
    ReDim Preserve Werk(I)
    Werk(I) = Table!WerkCode
    
    ReDim Preserve werkSourceId(I)
    werkSourceId(I) = Table!werkSourceId
    Table.MoveNext
    I = I + I
Wend
Table.Close

Werk(0) = ""

' ��������, �������� ����������� �������� ������ ���� equipId ���������� � 1 � �� ����� ���������
' ����� ������ � ������� ��������� � ��
mysql = "select equipId, equipName, equipFullName from GuideEquip where equipName <> '' order by 1"

Set Table = myOpenRecordSet("##72.3", mysql, dbOpenForwardOnly)
If Table Is Nothing Then myBase.Close: End

I = 1
While Not Table.EOF
    ReDim Preserve Equip(I)
    ReDim Preserve EquipFullName(I)
    Equip(I) = Table!equipName
    EquipFullName(I) = Table!EquipFullName
    Table.MoveNext
    I = I + 1
Wend
Table.Close


If Not initFomulConstats Then
    MsgBox "������ ��� ������������� ������" _
        & vbCr & "������ ��� �� ����� ����� ����������", vbOKOnly Or vbExclamation, "���������� � ��������������"
End If



If dostup = "y" Then
    WerkOrders.idWerk = 1: WerkOrders.Show
ElseIf dostup = "c" Then    '$$$ceh
    WerkOrders.idWerk = 2: WerkOrders.Show
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
statisticReplace = True
EN1:
End Function
'$NOodbc$
Sub checkNextYear()
Dim I As Integer, S As Double

I = Format(Now, "yyyy")
If I <= lastYear Then Exit Sub

If MsgBox("���������� ������������� �������� ���������� ��������� �� ����� ���. " & _
"���������?", vbDefaultButton2 Or vbYesNo, "�����������!") = vbNo Then Exit Sub

wrkDefault.BeginTrans

If Not statisticReplace("FirmGuide") Then GoTo ER1

If valueToSystemField("##389", I, "lastYear") Then
    wrkDefault.CommitTrans
    lastYear = I
    MsgBox "���� ���������� � �����(" & I & ") ���!"
Else
ER1: wrkDefault.Rollback
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

Sub msgOfZakaz(myErrCod As String, Optional msg As String = "", Optional mng = "")
    myErrCod = Mid$(myErrCod, 3)
    
    If msg = "" Then
        msg = "�������� ����������� ������."
    End If
    msg = msg & "��������� ���� ���� (Err=" & myErrCod & ") � ������ � " & gNzak
    
    If Not IsNull(mng) Then
        msg = msg & "  " & CStr(mng)
    End If
    
    MsgBox msg & " ������ � ���� ������� ���� " & _
    "����������. �������� ��������������!", , msg

End Sub

Sub msgOfEnd(myErrCod As String, Optional msg As String = "")
    myErrCod = Mid$(myErrCod, 3)
    MsgBox msg & " �������� ��������������!", , "������ " & myErrCod
    End
End Sub

' ���������� issueId ������ BusinessIssue ���� ���������� ������������� ���������
' ����� - ���������� 0 - ���� ���������� ����������
' ��� -1 ���� ��������� ������ ������ � ���� ����� ������ � ��������� ������
' ���� ��������� ������ �� ����� ���������� ������ ����, �� ������ ������ �������� myExecute
Function myExecuteWithIssue(ByVal pSql As String, ByVal passErr As Integer, ByRef issueId As Integer) As Integer
On Error GoTo viewAtTheErrorNumber
    wrkDefault.BeginTrans
    myBase.Execute pSql ', dbFailOnError  ' �������� Err ���� ��� ��� ����� ������� �������������
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
wrkDefault.Rollback
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


'$odbc08!$
Sub nextDayDetect() '�� ����� Orders.cmAdd_Click
Dim str As String ', intNum As Integer
Dim strNow As String, dNow As Date
Dim serverDate As String

strNow = Format(Now, "dd.mm.yyyy")
curDate = strNow '��� ����� � �����

sql = "select convert(varchar(10), now(), 104)"
byErrSqlGetValues "##chksrvdate", sql, serverDate

If serverDate <> curDate Then
    fatalError "����� �� ���������� ����� ������ ���������� �� ������� �������. " _
    & vbCr & "���� �� �������: " & serverDate _
    & vbCr & "������ ��������� ����� ���������.", _
    "��������� ���� �� ����� ����������. ���� �� ���������� ��� �� �� ������ ��� ��� �������, ���������� � ��������������."
End If

dNow = strNow
strNow = Right$(Format(curDate, "yymmdd"), 6)
 
befDays = 0

wrkDefault.BeginTrans 'lock01
sql = "update system set resursLock = resursLock" 'lock02
myBase.Execute (sql) 'lock03

Set tbSystem = myOpenRecordSet("##91", "System", dbOpenTable) ', dbOpenForwardOnly)
If tbSystem Is Nothing Then Exit Sub

'������� lock01-04 ��������� �� ������������� ��������� � Sybase
'tbSystem.Edit '$odbs?$ ����������, ����� ������ �� ����� �� ������ �� ������
'������ �� Update

Dim doUpdateNum As Boolean
doUpdateNum = False

If tbSystem!resursLock = "nextDay" Then
   wrkDefault.Rollback
   MsgBox "������ �������� ��������������! � ���� ����� �������� � ����������, " & _
    "�� c ������������ �����������������.", , "Error ��� �������� ���� �� ����� ����!"

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
     Else ' �.�. ���� lastPrivatNum �� ���� ��� ����������������
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


Function StatParamsLoad(row As Long, ByRef orderBean As ZakazVO, Optional redraw As Boolean = False)
Dim S As Double, log As String, str As String
Dim Grid As MSFlexGrid

Set Grid = Orders.Grid

If redraw Then
    Grid.col = orStatus
    Grid.row = row
    If orderBean.equipStatusSync <> 0 Then
        Grid.CellForeColor = vbRed
    Else
        Grid.CellForeColor = vbBlack
    End If
End If

 log = Format(Now(), "dd.mm.yy hh:nn") & vbTab & Orders.cbM.Text & " " & gNzak ' ������ vbTab
 If tqOrders!WerkId = 1 And IsNull(tqOrders!Equip) Then
    str = tqOrders!Status
 Else
    str = Status(tqOrders!StatusId)
 End If
 log = log & " " & str
 Orders.Grid.TextMatrix(row, orStatus) = str
 
 str = LoadDate(Orders.Grid, row, orDataVid, tqOrders!Outdatetime, "dd.mm.yy")
 If str <> "" Then log = log & " Out=" & str
 str = LoadNumeric(Orders.Grid, row, orVrVid, tqOrders!Outtime)
 If str <> "" Then log = log & "_" & str
 
 
 
 str = LoadNumeric(Orders.Grid, row, orVrVip, orderBean.Worktime, , "#0.0")
 log = log & " ��.���=" & str
 
 Orders.Grid.TextMatrix(row, orProblem) = tqOrders!problem
 
 str = LoadDate(Orders.Grid, row, orDataRS, tqOrders!DateRS, "dd.mm.yy")
 If str <> "" Then log = log & " ��=" & str
 
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
    log = log & " M�(" & tqOrders!StatM & "):" & str ' ���� ���
 End If
 If IsNull(orderBean.StatO) Then
    Orders.Grid.TextMatrix(row, orO) = ""
    Orders.Grid.TextMatrix(row, orOVrVip) = ""
 Else
    Orders.Grid.TextMatrix(row, orO) = orderBean.StatO
    If orderBean.StatO = "� ������" Or orderBean.StatO = "�����" Then
        'If IsNull(orderBean.DateTimeMO) Then
        '    msgOfZakaz "##313", "����������� '���� MO'."
        '    str = " !��� ���� MO! "
        'End If
        log = log & " O�(" & orderBean.StatO & "):" & str ' ���� ���
        If IsNull(orderBean.WorktimeMO) Then
            msgOfZakaz "##314", "����������� '����� ���������� MO'."
        Else
            Orders.Grid.TextMatrix(row, orOVrVip) = orderBean.WorktimeMO
            str = LoadNumeric(Orders.Grid, row, orOVrVip, orderBean.WorktimeMO)
            log = log & "=" & str
        End If
    End If
 End If
StatParamsLoad = log
End Function

Sub rowViem(numRow As Long, Grid As MSFlexGrid)
Dim I As Long

I = Grid.Height \ Grid.RowHeight(1) - 1 ' ������� ��������� �����
I = numRow - I + 2
If I < 1 Then I = 1
Grid.TopRow = I
End Sub

'��� �-� �.�������� � startDay() � getNextDay() � getPrevDay()
' ���������� �������� �� ����. ���
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
While J < 3 '         ������� �������� �������� �������� (3-� ����)

    day = Weekday(DateAdd("d", K + I - befDays, curDate))
'    day = Weekday(CurDate - befDays + K + I)
    If Not (day = vbSunday Or day = vbSaturday) Then J = J + 1
    I = I + 1
Wend
stDays(K) = I + K ' "+k" �.�. ���� ��������� ���������� befDays ���� �����

Next K          '       ***********************************************
dayMassLenght (stDays(befDays) + 1)
startDays = stDays(befDays) - befDays ' ��� �������, ������� ��� �1
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
 Report.laHeader.Caption = "���������� ��������� ���� �� ������ � " & str & _
                " �� " & Reports.tbEndDate.Text
 nMonth = Left$(str, 2)
 nYear = Right$(str, 4)
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
    strWhere = "WHERE (((FirmGuide.Kategor)='�'));"
    Report.Grid.ColWidth(4) = 0
 ElseIf Report.Regim = "RA" Then
    strWhere = "WHERE (((FirmGuide.Kategor)='�' Or (FirmGuide.Kategor)='�'));"
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
 "Sale, ManagID FROM FirmGuide " & strWhere
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
    "((Orders.FirmId)=" & tbFirms!FirmId & ") AND ((Orders.workTime) Is Not Null));"
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
            Report.Grid.TextMatrix(nRow, 1) = tbFirms!Name
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
  If nRow > 1 Then Report.Grid.RemoveItem (nRow)
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

Function checkInputDateFormat(inp As String) As String
    checkInputDateFormat = ""
    If Len(inp) = 0 Then
        '������ ���� �������� ��������
        Exit Function
    End If
    If Len(inp) <> 8 Or Mid(inp, 3, 1) <> "." Or Mid(inp, 6, 1) <> "." Then
        checkInputDateFormat = "��������� �������� ������ ��������������� ������� dd.mm.yy"
    End If
End Function

'� ������ true ����� ���������� ���� � tmpDate
Function isDateTbox(tBox As TextBox, Optional fryDays As String = "", Optional doEmptyCheck As Boolean = True) As Boolean
Dim str As String, checkDt As String

isDateTbox = False
str = tBox.Text

checkDt = checkInputDateFormat(str)
If checkDt <> "" Then
    MsgBox checkDt, , "������ ��� ����� ����"
    Exit Function
End If

If str <> "" Then
    str = "20" & Right$(str, 2) & "-" & Mid$(str, 4, 2) & "-" & Left$(str, 2)
    If IsDate(str) Then
        isDateTbox = True
        tmpDate = str
        If fryDays <> "" Then
            day = Weekday(tmpDate)
            If day = vbSunday Or day = vbSaturday Then
                If MsgBox(str & " - �������� ����. ����������?", vbYesNo, "��������������!") <> vbYes Then
                    isDateTbox = False
                End If
            End If
        End If
    Else
        MsgBox "�������� ������ ���� ��� ��� � ����� ����� �� ���������� ", , "������"
    End If
Else
    If doEmptyCheck Then
        MsgBox "��������� ���� ����!", , "������"
    End If
End If
If Not isDateTbox Then
    tBox.SelStart = 0
    tBox.SelLength = Len(tBox.Text)
    On Error Resume Next
    tBox.SetFocus
End If
End Function


Function valueToSystemField(myErr As String, val As Variant, Field As String) As Boolean

valueToSystemField = False
'sql = "select * from System"
'Set tbSystem = myOpenRecordSet(myErr, sql, dbOpenForwardOnly)
'If tbSystem Is Nothing Then myBase.Close: End
'Debug.Print val
If val = "" Then val = "''"
myBase.Execute ("UPDATE SYSTEM SET " & Field & " = " & val)

'tbSystem.Edit
'tbSystem.Fields(field) = val
'tbSystem.Update
'tbSystem.Close
valueToSystemField = True
End Function

'�� ��������� ������������ ��������, ��� �����, ��� �����
'�������� ���������. �  ������� ��� ���� error?
Function ValueToTableField(myErrCod As String, ByVal Value As String, ByVal Table As String, _
ByVal Field As String, Optional by As String = "", Optional Numorder As Variant) As Integer
Dim sql As String, byStr As String  ', numOrd As String


ValueToTableField = False
'If value = "" Then value = Chr(34) & Chr(34)

If Value = "" Then Value = "''"
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
ElseIf by = "byEquipId" Then
    byStr = ".numOrder = " & gNzak & " AND " & Table & ".equipId = " & gEquipId
ElseIf by = "byNumDoc" Then
    sql = "UPDATE " & Table & " SET " & Table & "." & Field & "=" & Value _
        & " WHERE " & Table & ".numDoc =" & numDoc & " AND " & Table & _
        ".numExt =" & numExt
    GoTo AA
Else
    Exit Function
End If
sql = "UPDATE " & Table & " SET " & Table & "." & Field & _
" = " & Value & " WHERE " & Table & byStr
AA:
'MsgBox "sql = " & sql

If Left$(myErrCod, 1) = "W" Then
    myErrCod = Mid$(myErrCod, 2)
    ValueToTableField = myExecute(myErrCod, sql, 0) '�� �������� ���� �� WHERE
ElseIf Left$(myErrCod, 1) = "L" Then
    ' �������, ��� ����������� �������� ����� �������� � ��������� ������ (lBusinessIssues).
    ' ������ ������� ������ ������ - "L-17002"
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

If Frm.Grid5.TextMatrix(v_row, prType) = "�������" Then
    str = Frm.Grid5.TextMatrix(v_row, prName) '
    I = InStr(str, "/")
    prExt = 0: If I > 1 Then prExt = Left$(str, I - 1)   '����� ��������
    gProductId = Frm.Grid5.TextMatrix(v_row, prId)
Else
    gNomNom = Frm.Grid5.TextMatrix(v_row, prId)
End If
End Sub

Function getNevip(day As Integer, EquipId As Integer)
sql = "SELECT Sum(oe.workTime * oe.Nevip) AS wSum " & _
"FROM OrdersEquip oe " & _
"WHERE DateDiff(day,'" & Format(curDate, "yyyy-mm-dd") & "',oe.outDateTime) =" & day - 1 _
& " AND oe.equipId =" & EquipId
'MsgBox sql
getNevip = 0
byErrSqlGetValues "W##382", sql, getNevip
End Function

Sub addDays(outDay As Integer)
Dim J As Integer
        If maxDay < outDay Then
            dayMassLenght outDay + 1 '���� ������ , ������������ �����������
            For J = maxDay + 1 To outDay '����� ���
                delta(J) = 0
            Next J
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
Dim str As String, I As Integer, myStatId As Integer

sql = "SELECT Orders.inDate, Orders.StatusId , Orders.FirmId From Orders " & _
"WHERE Orders.numOrder = " & gNzak
'MsgBox sql
If Not byErrSqlGetValues("##88", sql, tmpDate, myStatId, I) Then GoTo ER1

If I = 0 Then Exit Sub
If firm <> "" And (myStatId = 0 Or myStatId = 7) Then Exit Sub ' ���� ������ �����

str = getYearField(tmpDate)

'If str = "lock" Then Exit Sub ' ���� ��� ���������� � ��������� ���� , �� ��� �� �������������

sql = "UPDATE FirmGuide SET " & str & " = " & _
str & " " & oper & " 1  WHERE FirmId = " & I
'Debug.Print sql
I = myExecute("##87", sql, -143)

'If I <> 3061 And I <> 0 Then '3061 - ������� ����� ���� ���(��� ���) ��� � ����
If I = -2 Then '3061 - ������� ����� ���� ���(��� ���) ��� � ����
ER1:    MsgBox "������ ��������� ��������� ����. �������� ��������������!", , "Error-87"
End If
End Sub


Sub zagruzFromCeh(EquipId As Integer, Optional passZakazNom As String = "")
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
    & " WHERE oe.EquipId = " & EquipId & " AND (isnull(worktime, 0) > 0 OR isnull(worktimeMO, 0) > 0 ) " & passSql

'Debug.Print sql

Set tbCeh = myOpenRecordSet("##14", sql, dbOpenForwardOnly)
If tbCeh Is Nothing Then Exit Sub

'1:MaxDay = 0
If Not tbCeh.BOF Then
    While Not tbCeh.EOF
        isLive = False ' ������� �����
        If tbCeh!StatusId = 1 Then
            isLive = True
        End If
        If Not IsNull(tbCeh!Outdatetime) Then
            outDay = DateDiff("d", curDate, tbCeh!Outdatetime) + 1
        Else
            outDay = 0
        End If
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
Dim S As Double
'��������
sql = "SELECT Sum(sDMC.quant) AS Sum_quant From sDMC " & _
"WHERE (((sDMC.numExt)< 254) AND ((sDMC.numDoc)=" & numDoc & "));"
If Not byErrSqlGetValues("##140", sql, S) Then Exit Function
If S > 0.005 Then ' ���-�� ��������
    If reg = "" Then
        MsgBox "�� ����� ������ ������������ ���������, ������� �������� " & _
        "�������� ������. ���� ��������� ���-�� ���������, �� ������ ���� " & _
        "������� ��� ��������� � ������.", , "�������������� ���������!"
    End If
Else
    sql = "SELECT Sum(curQuant) AS Sum_curQuant " & _
    "From sDMCrez WHERE (((numDoc)=" & numDoc & "));"
    If Not byErrSqlGetValues("##367", sql, S) Then Exit Function
    If S > 0.005 Then ' ���-�� ��������
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


Sub getNakladnieList(Optional from As String = "") '
Dim I As Integer, str As String, L As Long

'pusto=-1 ���� ���� �� ������ �������� ��� �� � ����� ��������� (����� pusto=0)
'�����, �.�. ��� ���� delta=Null � �� quantity
If from = "Buh" Then str = "3" Else str = "2" '��� ���� ������ �� prior_������, ��� ����������� ���-�� � ����

sql = "SELECT numDoc, Max(quantity - IsNull( Sum_quant, 0)) AS delta, " _
& " Min(IsNull(Sum_quant,0)) AS pusto  " _
& " From wCloseNomenk" & str _
& " GROUP BY numDoc ORDER BY numDoc"

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
    
    If from = "werk" Then ' ��� ���� ��������� ��� � ���������
      sql = "SELECT numOrder From xEtapByNomenk Where numOrder = " _
      & gNzak _
      & " UNION ALL SELECT numOrder From xEtapByIzdelia " & _
      "WHERE numOrder = " & gNzak
      If Not byErrSqlGetValues("W##352", sql, L) Then GoTo NXT
      
      If L > 0 Then '��� ���������� ������ ������ ��� �������
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
Dim V As Variant

getNextNumExt = 0
sql = "SELECT Max(sDocs.numExt) AS Max_numExt From sDocs " & _
"WHERE (((sDocs.numDoc)=" & numDoc & " AND (sDocs.numExt) < 254));"

If Not byErrSqlGetValues("##128", sql, V) Then Exit Function
If IsNumeric(V) Then
    getNextNumExt = V + 1
Else
    getNextNumExt = 1
End If

End Function


'reg=""  - �������� ������� �������� ���������
'reg = "prev" - ��������, ��� ������� ����� �� ����.�����, �� ������
'����� - ��������, ��� ������� �� �����, �� �����
Function predmetiIsClose(Optional reg As String = "") As Boolean
Dim I As Integer, S As Double

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
    "WHERE sDMC.numDoc =" & gNzak & " AND nomNom = '" & NN(I) & "'"
    If Not byErrSqlGetValues("##164", sql, S) Then Exit Function
    If reg = "prev" Then
        If Abs(QQ3(I) - S) > 0.005 Then Exit Function
    ElseIf reg = "" Or QQ2(0) = 0 Then '����� �� �� ���� ��� ��� ���������� ������
        If QQ(I) - S > 0.005 Then Exit Function
    Else
        If QQ2(I) - S > 0.005 Then Exit Function
    End If
Next I
predmetiIsClose = True
End Function


Function PrihodRashod(reg As String, skladId As Integer) As Double
Dim qWhere As String, S As Double, targetSklad As String

PrihodRashod = 0

If reg = "+" Then
    targetSklad = "destId"
ElseIf reg = "-" Then
    targetSklad = "sourId"
End If
    
If skladId = 0 Then
    qWhere = "sDocs." & targetSklad & " < -1000"
ElseIf skladId = 2 Then
    qWhere = "(sDocs." & targetSklad & " = -1001 OR sDocs." & targetSklad & " = -1002)"
Else
    qWhere = "sDocs." & targetSklad & " = " & skladId
End If
sql = "SELECT Sum(sDMC.quant) AS Sum_quantity " _
& " FROM sDocs" _
& " INNER JOIN sDMC ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc) " & _
" WHERE sDMC.nomNom  = '" & gNomNom & "' AND " & qWhere
'Debug.Print sql
byErrSqlGetValues "##157", sql, PrihodRashod

End Function
'$odbc15$
Function ostatCorr(delta As Double) As Boolean
Dim sId As Integer, dId As Integer

ostatCorr = False

sql = "SELECT sDocs.sourId, sDocs.destId, sDocs.numDoc, sDocs.numExt " & _
"From sDocs " & _
"WHERE sDocs.numDoc = " & numDoc & " AND sDocs.numExt = " & numExt
If Not byErrSqlGetValues("##180", sql, sId, dId) Then Exit Function

If sId < -1000 And dId < -1000 Then ' ��� ������������ �� ������������
        ostatCorr = True
Else
    '������������ �������
    sql = "UPDATE sGuideNomenk SET nowOstatki = [nowOstatki]-" & delta & _
    " WHERE sGuideNomenk.nomNom = '" & gNomNom & "'"
    If myExecute("##163", sql) <> 0 Then Exit Function
    ostatCorr = True
End If
End Function

'���-�� ��� ������������ ���������, � ����� �������� �� ����� ��������
'��������� � Otgruz.frm
Sub loadPredmeti(Frm As Form, orderRate As Double, idWerk As Integer, asWhole As Integer, Optional reg As String = "", Optional needToRefresh As Boolean = False)
Dim I As Integer
Dim Nomnom1 As Nomnom, myAsWhole As Integer
Dim sumVes As Double


Screen.MousePointer = flexHourglass
Frm.Grid5.Visible = False
Frm.quantity5 = 0
I = 0: If reg = "fromOtgruz" Then I = 1

clearGrid Frm.Grid5, 1 + I

'******** ������� ************************************************
sql = "call wf_sell_ordered_ves (" & gNzak & ")"
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
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prCenaEd) = Round(rated(tbNomenk!cenaEd, orderRate), 2)
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prSumm) = _
                                Round(rated(tbNomenk!cenaEd * tbNomenk!quant, orderRate), 2)
    End If
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prQuant) = Round(tbNomenk!quant, 2)
    
    If idWerk = 1 Then
        Dim Ves As Variant 'null
        Ves = tbNomenk("totVes").Value
        If Not IsNull(Ves) Then
            Frm.Grid5.TextMatrix(Frm.quantity5 + I, prVes) = Round(Ves, 2)
            sumVes = sumVes + Ves
        End If
    End If
    
' ��� ��������� ��������� � ��� ���-�� (��. ����)
    If reg = "fromOtgruz" Then
        Otgruz.getOtgrugeno Frm.quantity5 + I, tbNomenk!cenaEd
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
sql = "SELECT n.nomNom, n.nomName" & _
", n.Size, n.nomNom, n.cod" & _
", n.ed_Izmer, n.ed_Izmer2, k.quant, k.cenaEd, n.perList" & _
", en.eQuant, en.prevQuant, n.ves " & _
" FROM (sGuideNomenk n " & _
" INNER JOIN xPredmetyByNomenk k ON n.nomNom = k.nomNom) " & _
" LEFT JOIN xEtapByNomenk en ON (k.nomNom = en.nomNom) AND (k.Numorder = en.Numorder) " & _
" WHERE k.Numorder =" & gNzak


'Debug.Print sql
Set tbNomenk = myOpenRecordSet("##184", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then Exit Sub
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    Frm.quantity5 = Frm.quantity5 + 1
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prId) = tbNomenk!Nomnom
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prType) = "������������"
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prName) = tbNomenk!cod
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prDescript) = tbNomenk!nomName & " " & tbNomenk!Size
    
    Dim quant As Double, cenaEd As Double
    quant = tbNomenk!quant
    cenaEd = tbNomenk!cenaEd
    
    
    Set Nomnom1 = nomnomCache.getNomnom(tbNomenk!Nomnom, True)
    
    If idWerk <> 1 Then
        myAsWhole = 0
    Else
        myAsWhole = asWhole
    End If
    
    
    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prEdizm) = Nomnom1.getEdizm(myAsWhole)
    
    Dim myQuant As Double
    If idWerk = 1 Then
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prVes) = Nomnom1.getVesEd(myAsWhole) * Nomnom1.getQuantity(quant, myAsWhole)
        sumVes = sumVes + Nomnom1.getVesEd(myAsWhole) * Nomnom1.getQuantity(quant, myAsWhole)
    End If

    If Not IsNull(cenaEd) Then
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prCenaEd) = Nomnom1.getCenaEd(rated(cenaEd, orderRate), myAsWhole)
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prSumm) = calcRounded(rated(quant * cenaEd, orderRate))
    End If

    Frm.Grid5.TextMatrix(Frm.quantity5 + I, prQuant) = Nomnom1.getQuantity(quant, myAsWhole)

    If reg = "fromOtgruz" Then
        Otgruz.getOtgrugeno Frm.quantity5 + I, tbNomenk!cenaEd, "byNomenk"
    ElseIf Not IsNull(tbNomenk!eQuant) Then
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prEtap) = Round(tbNomenk!eQuant, 2)
        Frm.Grid5.TextMatrix(Frm.quantity5 + I, prEQuant) = Round(tbNomenk!eQuant - tbNomenk!prevQuant, 2)
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
    Frm.Grid5.CellFontBold = True
    Frm.Grid5.MergeCol(prQuant) = True
    'Frm.Grid5.col = prQuant - 1
    'Frm.Grid5.Text = "�����:"
    'Frm.Grid5.CellFontBold = True
    'Frm.Grid5.MergeCol(prQuant - 1) = True
    
    Frm.Grid5.col = prSumm
    Frm.Grid5.Text = Round(rated(sProducts.saveOrdered(orderRate, needToRefresh), orderRate), 2)
    Frm.Grid5.CellFontBold = True
    
    If idWerk = 1 Then
        Frm.Grid5.col = prVes
        Frm.Grid5.CellFontBold = True
        Frm.Grid5.Text = Round(sumVes, 2)
    End If
    
    If reg = "fromOtgruz" Then
        Frm.Grid5.col = prOutSum
        Frm.Grid5.Text = Round(rated(Otgruz.saveShipped(False), orderRate), 2)
        Frm.Grid5.CellFontBold = True
        
        Frm.Grid5.col = prNowSum
        Frm.Grid5.Text = "0"
        Frm.Grid5.CellFontBold = True
    End If
    If needToRefresh Then
        Orders.openOrdersRowToGrid "##212.1"
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
    
Function orderUpdateWithIssue(ByVal issueMarker As String, Value As String, Table As String, _
Field As String, Optional by As String = "", Optional Numorder As Variant) As Integer
Dim nzak As String
Dim issueId As Variant
    If IsMissing(Numorder) Then
        nzak = gNzak
    Else
        nzak = Numorder
    End If
    orderUpdateWithIssue = ValueToTableField("##orderUpdateWithIssue", Value, Table, Field, by, nzak)
    
    sql = "select wi_check_business_issue(' " & issueMarker & "')"
    byErrSqlGetValues "##check_issue", sql, issueId
    If Not IsNull(issueId) And issueId <> 0 Then
        orderUpdateWithIssue = CInt(issueId)
    End If
    
    If Table = "Orders" Then
        refreshTimestamp (nzak)
    End If
End Function
    
    
Function orderUpdate(ByVal myErrCod As String, Value As String, Table As String, _
Field As String, Optional by As String = "", Optional Numorder As Variant) As Integer
Dim nzak As String
    If IsMissing(Numorder) Then
        nzak = gNzak
    Else
        nzak = Numorder
    End If
    orderUpdate = ValueToTableField(myErrCod, Value, Table, Field, by, nzak)
    If Table = "Orders" Then
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
 Set Node = p_tv.Nodes.Add(, , "k0", "���������� �� ������")
 Node.Sorted = True
 
 ReDim K(0): ReDim pK(0): ReDim NN(0): iErr = 0
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
  For I = 1 To UBound(K())
    If K(I) <> "" Then
        On Error GoTo ERR2 ' ��������� ��� ������
        Set Node = p_tv.Nodes.Add(pK(I), tvwChild, K(I), NN(I))
        On Error GoTo 0
        K(I) = ""
        Node.Sorted = True
    End If
NXT:
  Next I
Wend
p_tv.Nodes.Item("k0").Expanded = True
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
                    getManagById = CStr(Managers(I).Value)
                    Exit Function
                End If
            Next I
        End If
    End If
End Function

Private Sub addToCbStatus(ByRef statusComboBox As ComboBox, id)

    'Static I As Integer
    If id > lenStatus Then
        MsgBox "Err � Orders\addToCbStatus"
    End If

    statusComboBox.AddItem Status(id)
    statusComboBox.ItemData(statusComboBox.ListCount - 1) = id

End Sub
    


Public Sub cbBuildStatuses(ByRef statusComboBox As ComboBox, ByRef statusIdOld As Integer)
    
    statusComboBox.Clear
    
    If statusIdOld = 4 Then
        addToCbStatus statusComboBox, 6 '"������"
    End If
    
    addToCbStatus statusComboBox, 7 ', "b" '"�������."
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

Public Function cbMOsetByText(cb As ComboBox, Status As Variant, Optional baseIndex As Integer = 1) As Boolean
Dim I As Integer, txt As String
    If noClick Then Exit Function
    cbMOsetByText = False
    txt = ""
    If Not IsNull(Status) Then txt = CStr(Status)
    If txt = "�����" Then
        If cb.List(baseIndex + 2) <> "�����" Then cb.AddItem "�����", baseIndex + 2
        If cb.List(baseIndex + 3) <> "���������" Then cb.AddItem "���������", baseIndex + 3
        cb.ListIndex = baseIndex + 2
        cb.Enabled = True
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

Function addCurrencyToCaption(Caption As String) As String
    If sessionCurrency = CC_RUBLE Then
        addCurrencyToCaption = Caption & " - �����"
    ElseIf sessionCurrency = CC_UE Then
        addCurrencyToCaption = Caption & " - �������"
    End If
End Function

Function calcRounded(ByVal inp As Double) As Double

    Dim rounded As Double
    rounded = Round(inp, 2)
    If Abs(rounded) < 0.005 Then
        rounded = Round(inp, 3)
    End If
    calcRounded = rounded
End Function


Public Sub initEquipCombo(ByRef cbEquips As ComboBox, idWerk As Integer, Optional idEquip As Integer = 0)

    Dim werkSql As String, I As Integer
    If idWerk > 0 Then
        werkSql = " AND we.werkId = " & idWerk
    End If
    
    sql = "select e.equipId, e.equipFullName, we.equipId as IsPresent " _
        & vbCr & " from GuideEquip e " _
        & vbCr & " LEFT JOIN WerkEquip we ON we.equipId = e.equipId" & werkSql _
        & vbCr & " WHERE e.equipId > 0" _
        & vbCr & " group by e.equipId, e.equipFullName, ispresent, we.equipOrder" _
        & vbCr & " having ispresent is not null" _
        & vbCr & " order by isnull(we.equipOrder, 100000), e.equipId"
    'Debug.Print sql
    
    Dim nEquips As Integer
    nEquips = cbEquips.ListCount - 1
    For I = nEquips To 0 Step -1
        cbEquips.RemoveItem I
    Next I
    
    If idWerk >= 0 Then
        cbEquips.AddItem "��� ���� ������������", 0
    Else
        ' ������� ������ �� ���������
        If idEquip <= 0 Then
            idEquip = 1
        End If
    End If
    initCombobox sql, cbEquips, "equipId", "equipFullName", 1
    Dim saveNoClick As Boolean
    saveNoClick = noClick
    noClick = True
    For I = 0 To cbEquips.ListCount - 1
        If cbEquips.ItemData(I) = idEquip Then
            cbEquips.ListIndex = I
            Exit For
        End If
    Next I
    noClick = saveNoClick


End Sub

Public Sub initWerkCombo(ByRef cbWerks As ComboBox, Optional idWerk As Integer = 0)

    Dim I As Integer
    
    Dim nWerks As Integer
    nWerks = cbWerks.ListCount - 1
    For I = nWerks To 0 Step -1
        cbWerks.RemoveItem I
    Next I
    
    If idWerk >= 0 Then
        cbWerks.AddItem "���", 0
    End If
    
    For I = 1 To UBound(Werk)
        cbWerks.AddItem Werk(I), I
        cbWerks.ItemData(I) = I
    Next I
    
    
    Dim saveNoClick As Boolean
    saveNoClick = noClick
    noClick = True
    For I = 0 To cbWerks.ListCount - 1
        If cbWerks.ItemData(I) = idWerk Then
            cbWerks.ListIndex = I
            Exit For
        End If
    Next I
    noClick = saveNoClick


End Sub


