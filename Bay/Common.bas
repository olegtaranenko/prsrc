Attribute VB_Name = "Common"
Option Explicit

'Public otlStr As String
Public isOrders As Boolean
Public isFindFirm As Boolean
Public mainTitle As String
Public flReportArhivOrders As Boolean

Public tbOrders As Recordset
Public tqOrders As Recordset
Public tbSystem As Recordset
Public tbFirms As Recordset
Public tbNomenk  As Recordset
Public tbProduct  As Recordset
Public tbDMC As Recordset
Public tbGuide As Recordset
Public isBlock As Boolean
Public status() As String
Public Problems() As String
Public Ceh(10) As String

Public manId() As Integer
Public Manag() As String ' ������ = id �� GuideManag

Public insideId() As String
Public Const begCehProblemId = 10 ' ������ ������� ������� � �����������
Public neVipolnen As Double, neVipolnen_O As Double
Public MaxDay As Integer ' ����� ���� � �������
'Public tmpMaxDay As Integer '����� ���� � ���� Zakaz
Public befDays As Integer ' ����� ���� �� ���� ������� (����� ��������� ����)

Public begDate As Date ' ���� ������������� ��������
Public logFile As String
Public dostup As String
Public otlad As String
Public tbSIze As Integer
Public cErr As String '��������� ������� ����� ������������� Err, ���� ��
                      '���� ������ ��������� �� Err ������ ���� MsgBox
'Public iDate As Date
Public zakazNum As Long  ' ���-�� ������� �  M��.�������
Public gNzak As String  ' ��� ����� ������
Public gFirmId As String
Public gProductId As String
Public gProduct As String
Public gDocDate As Date
Public gSeriaId As String
Public gKlassId As String
Public gNomNom As String
Public numDoc As Long, numExt As Integer


Public oldValue As String '������ �������� ����, ����������� ���������
Public CurDate As Date
Public lastYear As Integer

Public begDay As Integer ' ���� ������� ����� ������
Public endDay As Integer ' ���� ���������� ����� ������
Public begDayMO As Integer ' ���� ������� ����� �� ������
Public endDayMO As Integer ' ���� ���������� ����� �� ������
Public flEdit As String ' ������������� ������
Public Nstan As Double
Public KPD As Double
Public newRes As Double ' ����� �� ���������
Public nr As Double, dr As Double '��������� ���. � ���. �������
'Public isDoMO As Boolean ' �� ����� ��� ����. �� ��� ����� - ����� ���������� ����� ��������
Public isLive As Boolean ' ���� - ����� �����
Public zagAll As Double, zagLive As Double
Public drobleDopRes As Boolean

Public Table As Recordset '
Public myQuery As QueryDef
Public sql As String      ' ������������� �����������
Public strWhere As String '
'Public mousRow As Long
'Public mousCol As Long    '
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
Public orNomZak As Integer, orData As Integer ', orTema As Integer
Public orMen As Integer, orStatus As Integer, orProblem As Integer
Public orFirma As Integer, orDataVid As Integer
Public orVrVid As Integer, orInvoice As Integer, orZakazano As Integer, orRate As Integer
Public orVes As Integer, orSize As Integer, orPlaces As Integer
Public orOplacheno As Integer, orOtgrugeno As Integer, orLastMen As Integer
Public orVenture As Integer
Public orBillId As Integer
Public orVocnameId As Integer
Public orServername As Integer

Public sessionCurrency As Integer


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


Public NN() As String, QQ() As Double ' ������������ ������������ � ���-��
Public QQ2() As Double, QQ3() As Double
'Public tmpNum As Double ' �������� � �.�. ��� isNunericTbox()
'Public cehNom As Integer
Public skladId As Integer

Public Const cDELLwidth = 19200 ' ��� ����� � ��� = 19290
Private Const dhcMissing = -2 '����� ��� quickSort

Public Const otladColor = &H80C0FF

Public Const CC_RUBLE As Integer = 1
Public Const CC_UE As Integer = 2

' �� ������� ����� ����������� ������ �������, ���� ������� �����
Public Const ColWidthForRuble As Single = 1.3



Function tuneCurencyAndGranularity(tunedValue, currentRate, valueCurrency As Integer, Optional quantity As Double = 1) As Double
    '
    Dim totalInRubles As Double
    Dim singleInRubles As Double
    Dim totalInUE As Double
    Dim singleInUE As Double
    '
    If valueCurrency = CC_RUBLE Then
        singleInRubles = Round(tunedValue / quantity, 2)
    Else
        singleInRubles = Round(tunedValue / quantity * currentRate, 2)
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
    rated = geld
    If rate <= 0 Then
        Exit Function
    End If
    If sessionCurrency = CC_RUBLE Then
        rated = CDbl(geld) * CDbl(rate)
    End If
End Function



Function getValueFromTable(tabl As String, field As String, where As String) As Variant
Dim Table As Recordset

getValueFromTable = Null
sql = "SELECT " & field & " as fff  From " & tabl & _
      " WHERE " & where & ";"
Set Table = myOpenRecordSet("##59.1", sql, dbOpenForwardOnly)
If Table Is Nothing Then Exit Function
If Not Table.BOF Then getValueFromTable = Table!fff
Table.Close
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



'���� ������ �a������ ="W.." - �� �������� Err �� �����-� Where, � ���
'��������� ��������, ���� ��� ���� ��� ���� ��� ��������� ��������, �� � sql
'�. ������ ��������� "1" � ������� �� � i. ����� ���� i=0 �� ���� Err Where
Function byErrSqlGetValues(ParamArray val() As Variant) As Boolean
Dim tabl As Recordset, I As Integer, maxi As Integer, str As String
Dim c As String

byErrSqlGetValues = False
maxi = UBound(val())
If maxi < 1 Then
    wrkDefault.Rollback
    MsgBox "���� ���������� ��� �\� byErrSqlGetValues()"
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
        MsgBox "��� ������� ��������������� Where!", , "Error-" & Mid$(str, 3)
        GoTo EN2
    End If
Else
    c = ""
End If

For I = 2 To maxi
    If IsNull(tabl.Fields(I - 2)) Or c = "W" Then
        str = TypeName(val(I))
'        If str = "Double" Or str = "Integer" Or str = "Long" Or str = "Double" Then
        If str = "String" Then
            val(I) = ""
        Else
            val(I) = 0
        End If
    Else
        val(I) = tabl.Fields(I - 2)
    End If
Next I
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

Sub delay(tau As Double)
Dim S As Double
    S = Timer
    While Timer - S < tau ' 1 ���
        DoEvents
    Wend

End Sub

Sub exitAll()
If isOrders Then Unload Orders
If isFindFirm Then Unload FindFirm
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
            ByVal col As Integer, Optional pos As Long = -1) As Long
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
        If Not Grid.RowIsVisible(il) Then
            Grid.TopRow = il
        End If
        Grid.row = il
        findExValInCol = il
        Exit Function
    End If
Next il
findExValInCol = -1

End Function

Function existValueInTableFielf(ByVal value As Variant, tabl As String, field) As Boolean
Dim Table As Recordset

existValueInTableFielf = False

If Not IsNumeric(value) Then value = "'" & value & "'"

sql = "SELECT " & field & " From " & tabl & " WHERE (((" & field & ") = " & _
value & "));"
'MsgBox sql
Set Table = myOpenRecordSet("##390", sql, dbOpenForwardOnly)
'If table Is Nothing Then myBase.Close: End

If Not Table.BOF Then existValueInTableFielf = True

Table.Close

End Function



        
Sub fitFormToGrid(frm As Form, Grid As MSFlexGrid)
Dim I As Long, delta As Long

I = 350 + (Grid.CellHeight + 17) * Grid.Rows
delta = I - Grid.Height
If frm.Height + delta > (Screen.Height - 400) Then _
    delta = (Screen.Height - 400) - frm.Height
frm.Height = frm.Height + delta

'delta = 0
'For i = 0 To Grid.Cols - 1
'    delta = delta + Grid.ColWidth(i)
'Next i
'frm.Width = delta + 700

End Sub

Function getOrdered(numZak As String) As Double
Dim S As Double

getOrdered = -1

sql = "SELECT Sum([sDMCrez].[quantity]*[sDMCrez].[intQuant]/[sGuideNomenk].[perList]) AS cSum " & _
"FROM sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom " & _
"WHERE (((sDMCrez.numDoc)=" & numZak & "));"
If Not byErrSqlGetValues("W##209", sql, S) Then Exit Function
getOrdered = S
End Function
'Orders.Grid.TextMatrix(Orders.Grid.row, orOtgrugeno)=getShipped()
Function getShipped(numZak As String) As Double
Dim S As Double, s1 As Double, str As String

getShipped = 0
'sql = "SELECT Sum([sDMC].[quant]*[sDMCrez].[intQuant]/[sGuideNomenk].[perList]) AS ���������1 " & _
"FROM (sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom) INNER JOIN sDMC ON (sDMCrez.nomNom = sDMC.nomNom) AND (sDMCrez.numDoc = sDMC.numDoc) " & _
"WHERE (((sDMCrez.numDoc)=" & numZak & "));"

sql = "SELECT Sum([bayNomenkOut].[quant]*[sDMCrez].[intQuant]) AS bSum " & _
"FROM bayNomenkOut INNER JOIN sDMCrez ON (bayNomenkOut.nomNom = sDMCrez.nomNom) AND (bayNomenkOut.numOrder = sDMCrez.numDoc) " & _
"WHERE (((sDMCrez.numDoc)=" & numZak & "));"
'Debug.Print sql

If Not byErrSqlGetValues("W##209", sql, S) Then Exit Function

getShipped = S
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


Sub listBoxInGridCell(lb As ListBox, Grid As MSFlexGrid, Optional sel As String = "")
Dim I As Integer
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
        val As Variant, Optional myErr As String = "") As Double
 If IsNull(val) Then
    Grid.TextMatrix(row, col) = ""
    LoadNumeric = 0 ' ��� log �����
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

Sub loadLbMassFromGuide(lbMass() As String, tableName As String)
Dim I As Integer

Set Table = myOpenRecordSet("##04", tableName, dbOpenForwardOnly)
If Table Is Nothing Then myBase.Close: End
ReDim lbMass(0)
While Not Table.EOF
    I = Table.Fields(0)
    ReDim Preserve lbMass(I)
    If tableName = "GuideStatus" Then
        If Table.Fields(1) = "� ������" Then
            lbMass(I) = "������" '
        ElseIf Table.Fields(1) = "�����" Then
            lbMass(I) = "�����" '
        Else
            lbMass(I) = Table.Fields(1)
        End If
    Else
        lbMass(I) = Table.Fields(1)
    End If
    Table.MoveNext
Wend
Table.Close
'lb.Height = lb.Height + 195 * (lb.ListCount - 1)
End Sub



Sub Main()
Dim I As Integer, S As Double, str As String, str1 As String, str2 As String
Dim isXP As Boolean

If App.PrevInstance = True Then
    MsgBox "��������� ��� ��������", , "Error"
    End
End If

mainTitle = getMainTitle

ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): ReDim QQ3(0) '����� Ubound ������� �� ������ Err

flReportArhivOrders = False
ReDim tmpL(0)

loadEffectiveSettingsApp
initLogFileName

checkReloadCfg

baseOpen
dostup = getEffectiveSetting("dostup")
otlad = getEffectiveSetting("otlad")
sessionCurrency = getEffectiveSetting("currency", CC_RUBLE)


On Error GoTo ERRf '�������� ��������� Win98
S = "1.6"

On Error GoTo ERRs ' �� ���� Err ���� � ���� �� �.������ server, ���� �� ��� DOS ���� ����.Err=53
If otlad <> "otlaD" And InStr(otlad, ":\") = 0 Then '
      
'  If dostup = "a" Then '�������� ��� �. ��������  winXP
  If Dir$("C:\WINDOWS\net.exe") = "" Then '�� �����
    Shell "C:\WINDOWS\system32\net time \\server /SET /YES", vbHide ' winXP
  Else
    Shell "C:\WINDOWS\net time \\server /WORKGROUP:JOBSHOP /SET /YES", vbHide
  End If
End If
On Error GoTo 0


If dostup = "0" Then I = 5 / I  '�������� ������� ���������� ���������

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
If tbSystem Is Nothing Then myBase.Close: End
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


str = Format(Now, "dd.mm.yyyy")
CurDate = str '��� ����� � �����

' ��������� ��� �����*********************************************
 
 'logFile = "C:\Windows\OrdersBay" ' ��� ����������
 logFile = App.path & "\" & App.EXEName
 str2 = logFile & "$$.log" ' ��������� ����
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
    'tmpDate = str
    If DateDiff("d", str1, CurDate) <= 7 Then Print #3, str ' ������� > 7�� ���� ��������
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
MsgBox "����\���������\������ ����������\���� � ���������\�����\" & _
      "����������� ����� � ������� ������ ����� " & Chr(151) & _
      " ���������� ����� ������ �������!", , "��� ���������� ������ " & _
      "��������� ���������� ��������� ��������� Win98: "
End

'ERRb:
'MsgBox "�� ������� ������������ � ���� " & mainTitle
'End

ERRs:
MsgBox "������� �� ������ ���������������� ����", , "�������� ��������������!"
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
    If msg = "" Then msg = "�������� ����������� ������."
    MsgBox msg & " ������ � ���� ������� ���� " & _
    "����������. �������� ��������������!", , _
    "��������� ���� ���� (Err=" & myErrCod & ") � ������ � " & gNzak
End Sub

'Sub msgOfEnd(myErrCod As String, Optional msg As String = "")
'    wrkDefault.Rollback
'
'    myErrCod = Mid$(myErrCod, 3)
'    MsgBox msg & " �������� ��������������!", , "������ " & myErrCod
'    End
'End Sub

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

'myBase.Close: End

End Function


'NULL - ��������� ��� �������
'skladId=0 - c������� �� ���� �������
'skladId=2 - c������� �� 1 � 2�� �������





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
Dim I As Integer

I = Grid.Height \ Grid.RowHeight(1) - 1 ' ������� ��������� �����
I = numRow - I \ 2 ' � �����
If I < 1 Then I = 1
Grid.TopRow = I
End Sub

'��� �-� �.�������� � startDay() � getNextDay() � getPrevDay()
' ���������� �������� �� ����. ���
Function getWorkDay(offsDay As Integer, Optional baseDate As String = "") As Integer
Dim I As Integer, J As Integer, step  As Integer
getWorkDay = -1
If baseDate = "" Then
    tmpDate = CurDate
Else
    If Not IsDate(baseDate) Then Exit Function
    tmpDate = baseDate
End If

step = 1
If offsDay < 0 Then step = -1

J = 0: I = 0
While step * J < step * offsDay '
    I = I + step
'    day = Weekday(tmpDate + i)
    day = Weekday(DateAdd("d", I, tmpDate))
    If Not (day = vbSunday Or day = vbSaturday) Then J = J + step
Wend
getWorkDay = I
'tmpDate = tmpDate + i
tmpDate = DateAdd("d", I, tmpDate)

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
        str = "20" & Right$(str, 2) & "-" & Mid$(str, 4, 2) & "-" & Left$(str, 2)
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

Sub valueToSystemField(myErr As String, val As Variant, field As String)
Set tbSystem = myOpenRecordSet(myErr, "System", dbOpenTable)
If tbSystem Is Nothing Then myBase.Close: End
tbSystem.Edit
tbSystem.Fields(field) = val
tbSystem.Update
tbSystem.Close
End Sub

'�� ��������� ������������ ��������, ��� �����, ��� �����
'�������� ���������. �  ������� ��� ���� error?
Function ValueToTableField(myErrCod As String, value As String, Table As String, _
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
    sql = "UPDATE " & Table & " SET " & Table & "." & field & "=" & value _
        & " WHERE (((" & Table & ".numDoc)=" & numDoc & " AND (" & Table & _
        ".numExt)=" & numExt & " ));"
    GoTo AA
Else
    Exit Function
End If
sql = "UPDATE " & Table & " SET " & Table & "." & field & _
" = " & value & " WHERE (((" & Table & byStr & " ));"
AA:
'MsgBox "sql = " & sql

If Left$(myErrCod, 1) = "W" Then
    myErrCod = Mid$(myErrCod, 2)
    ValueToTableField = myExecute(myErrCod, sql, 0) '�� �������� ���� �� WHERE
Else
    ValueToTableField = myExecute(myErrCod, sql)
End If
End Function



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
    beNaklads = False
End If

End Function
    
Function PrihodRashod(reg As String, skladId As Integer) As Double
Dim qWhere As String, S As Double

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
    getLockYear = getLockYear - 1 '�������, ��� ���� ��� �� ����������� � ��������� ����
End Function

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
    sql = "UPDATE sGuideNomenk SET nowOstatki = [nowOstatki]-" & delta & _
    " WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
    If myExecute("##163", sql) <> 0 Then Exit Function
    ostatCorr = True
End If
End Function


Function predmetiIsClose() As Variant
Dim I As Integer, S As Double

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
Dim I As Integer, lockYear As Integer

lockYear = getLockYear
I = Format(checkDate, "yyyy")
'If i <= lockYear Then
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


Sub visits(oper As String, Optional firm As String = "")
Dim str As String, I As Integer, statId As Integer

sql = "SELECT inDate, StatusId , FirmId From BayOrders " & _
"WHERE (((numOrder)=" & gNzak & "));"
'MsgBox sql
If Not byErrSqlGetValues("##88", sql, tmpDate, statId, I) Then GoTo ER1

If I = 0 Then Exit Sub
If firm <> "" And (statId = 0 Or statId = 7) Then Exit Sub ' ���� ������ �����

'str = "year" & Format(tmpDate, "yy")
str = getYearField(tmpDate) '$$3

sql = "UPDATE BayGuideFirms SET BayGuideFirms." & str & " = [BayGuideFirms].[" & _
str & "] " & oper & " 1  WHERE (((BayGuideFirms.FirmId)=" & I & "));"
'MsgBox sql
I = myExecute("##87", sql, -143)

'If i <> 3061 And i <> 0 Then '3061 - ������� ����� ���� ���(��� ���) ��� � ����
If I = -2 Then '3061 - ������� ����� ���� ���(��� ���) ��� � ����
ER1:    MsgBox "������ ��������� ��������� ����. �������� ��������������!", , "Error-87"
End If
End Sub

Function lockSklad(Optional back As String = "") As Boolean
Dim str As String

lockSklad = True: Exit Function '!!! �������� ���������
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
        
        If MsgBox("������ � �������� �� ������ �������� ����� ���������� '" & _
        str & "'. ��������� ������� ��� ���������� � ��������������.", _
        vbRetryCancel, "��� ������� !!!") = vbRetry Then
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

