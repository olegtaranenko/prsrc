VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Report 
   BackColor       =   &H8000000A&
   Caption         =   "�����"
   ClientHeight    =   8184
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8184
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmPrev 
      Caption         =   "<"
      Height          =   255
      Left            =   11280
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmNext 
      Caption         =   ">"
      Height          =   255
      Left            =   11520
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "������"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "�����"
      Height          =   315
      Left            =   10980
      TabIndex        =   4
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "������ � Exel"
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   7800
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7455
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   13145
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label laHeader 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   11775
   End
   Begin VB.Label laRecCount 
      Caption         =   "����� �������:"
      Height          =   195
      Left            =   2460
      TabIndex        =   2
      Top             =   7860
      Width           =   1335
   End
   Begin VB.Label laCount 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   7800
      Width           =   975
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Regim As String
Public idEquip As Integer
Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����
Dim zakazano As Double, Oplacheno As Double, Otgrugeno As Double
Public nCols As Integer ' ����� ���-�� �������
Public mousRow As Long
Public mousCol As Long
Public firmID As Integer
Public firmNazv As String

Public Edizm2 As String
Public Caller As Form

Public Sortable As Boolean
    '��������� ���������� - ����� ��� ��� ����� �������������.


Dim workSum As Double, paidSum As Double, quantity As Long
'��������� ��� firmOrders
Const rpNomZak = 1
Const rpM = 2
Const rpStatus = 3
Const rpProblem = 4
Const rpDataVid = 5
Const rpVrVid = 6
Const rpLogo = 7
Const rpIzdelia = 8
Const rpZakazano = 9
Const rpOplacheno = 10
Const rpOtgrugeno = 11
'��������� ��� managStat
Const rpM2 = 1
Const rpFirmRA = 2
Const rpFirmKK = 3
Const rpFirmAll = 4
Const rpQuantNoClose = 5
Const rpQuantAll = 6
Const rpWorkNoClose = 7
Const rpWorkAll = 8
Const rpPaidNoClose = 9
Const rpPaidAll = 10
'��������� ��� whoReserved
Const rtNomZak = 1
Const rtReserv = 2
Const rtCeh = 3
Const rtData = 4
Const rtMen = 5
Const rtStatus = 6
Const rtFirma = 7
Const rtProduct = 8
Const rtZakazano = 9
Const rtOplacheno = 10

Private Sub cmExel_Click()
Dim Equip As String, Left As String, X As String
    GridToExcel Grid, laHeader.Caption
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
End Sub

Private Sub cmNext_Click()
doVirabotka "next"
End Sub

Private Sub cmPrev_Click()
doVirabotka "prev"
End Sub

Private Sub cmPrint_Click()
Me.PrintForm

End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width
If Regim = "KK" Or Regim = "RA" Then
    statistic
ElseIf Regim = "Manag" Then
    managStat
ElseIf Regim = "whoRezerved" Then
    laHeader.Caption = "������ �������, ���. ������������� ���-�� '" & gNomNom & "' [" & Me.Edizm2 & "]."
    Me.MousePointer = flexHourglass
    laCount.Caption = whoRezerved(Me.Grid)
    Me.MousePointer = flexDefault
ElseIf Regim = "fromCehNaklad" Then
    productSostav
ElseIf Regim = "Virabotka" Then
    cmPrev.Visible = True
    cmNext.Visible = True
    laRecCount.Visible = False
    laCount.Visible = False
    cmExel.Visible = False
    doVirabotka
Else
    firmOrders
End If
End Sub
'str As String,
Sub doVirabotka(Optional direct As String = "")
Static prevDay As String, nextDay As String, str As String
Dim curDay As String, Resurs As Double, live As Double, sum As Double
Dim kpd_ As Double, res As Double, I As Integer
Const crNomZak = 1
Const crM = 2
Const crStatus = 3
Const crVrVip = 4
Const crProcVip = 5
Const crVirab = 6
Const crProblem = 7
Const crDataVid = 8
Const crVrVid = 9
Const crFirma = 10
Const crLogo = 11
Const crIzdelia = 12
    

curDay = Format(curDate, "yy.mm.dd")
If direct = "next" Then
    If nextDay = curDay Then
        direct = ""
        GoTo AA
    End If
    curDay = nextDay
ElseIf direct = "prev" Then
    If curDay = prevDay Then
        direct = ""
        GoTo AA
    End If
    curDay = prevDay
Else
End If
AA:
'������ ������� ��� ����
tmpStr = Right$(curDay, 2)
tmpStr = tmpStr & Mid$(curDay, 3, 4)
tmpStr = tmpStr & Left$(curDay, 2)
laHeader.Caption = "��������� �� ������������ " & Equip(idEquip) & " �� " & tmpStr

Grid.Rows = 2
Grid.Cols = 13
Grid.Clear
    Grid.ColWidth(0) = 0
    Grid.ColWidth(crNomZak) = 1000
    Grid.ColWidth(crM) = 270
    Grid.ColWidth(crVrVip) = 540
    Grid.ColWidth(crStatus) = 870
    Grid.ColWidth(crProcVip) = 420
    Grid.ColWidth(crVirab) = 930
    Grid.ColWidth(crProblem) = 900
    Grid.ColWidth(crVrVid) = 330
    Grid.ColWidth(crDataVid) = 735
    Grid.ColWidth(crFirma) = 2000
    Grid.ColWidth(crLogo) = 870
    Grid.ColWidth(crIzdelia) = 2450

sql = "SELECT numOrder, obrazec, Virabotka From Itogi" _
& " WHERE xDate ='" & curDay & "' AND equipId = " & idEquip & " ORDER BY numOrder, obrazec DESC;"
'MsgBox sql
Set tbOrders = myOpenRecordSet("##377", sql, dbOpenForwardOnly)
If tbOrders Is Nothing Then Exit Sub
If tbOrders.BOF Then GoTo EN1
Resurs = -1: live = -1
If tbOrders!Numorder = 0 Then
    Resurs = Round(tbOrders!virabotka, 2)
    tbOrders.MoveNext
End If
If tbOrders.EOF Then GoTo EN2

KPD = -1
If tbOrders!Numorder = 1 Then
    kpd_ = Round(tbOrders!virabotka, 2)
    tbOrders.MoveNext
End If
If tbOrders.EOF Then GoTo EN2

If tbOrders!Numorder = 2 Then
    live = Round(tbOrders!virabotka, 2)
    tbOrders.MoveNext
End If
sum = 0: quantity = 0
ReDim NN(0): ReDim QQ(0): ReDim QQ2(0)
While Not tbOrders.EOF
    quantity = quantity + 1
    ReDim Preserve NN(quantity): ReDim Preserve QQ(quantity): ReDim Preserve QQ2(quantity)
    NN(quantity) = tbOrders!Numorder
If tbOrders!Numorder = 4080201 Then
    I = I
End If
    QQ2(quantity) = (tbOrders!obrazec = "o") ' = -1 ��� �������
    QQ(quantity) = Round(tbOrders!virabotka, 2)
    sum = sum + QQ(quantity)
    tbOrders.MoveNext
Wend
EN1:
tbOrders.Close
EN2:

If direct = "" Then
    res = Zagruz.laUsed.Caption ' � ������ ���
    Resurs = Round(res / Zagruz.tbKPD.Text, 2)
Else
    res = Round(Resurs * kpd_, 2) ' � ������ ���(�� ����������)
End If
sum = Round(sum, 2)

Grid.MergeCells = flexMergeRestrictRows 'flexMergeRestrictAll 'flexMergeRestrictColumns
Grid.TextMatrix(0, 1) = "��������"
Grid.TextMatrix(1, 1) = "���������"
Grid.AddItem vbTab & "������ � ������ �������������"
Grid.AddItem vbTab & "������ ��� ����� �������������"
Grid.AddItem vbTab & "�������� �������������."
Grid.AddItem vbTab & "����� �����"

Grid.MergeRow(0) = True
Grid.MergeRow(1) = True
Grid.MergeRow(2) = True
Grid.MergeRow(3) = True
Grid.MergeRow(4) = True
Grid.MergeRow(5) = True
For I = 2 To crVirab - 1
    Grid.TextMatrix(0, I) = Grid.TextMatrix(0, 1)
    Grid.TextMatrix(1, I) = Grid.TextMatrix(1, 1)
    Grid.TextMatrix(2, I) = Grid.TextMatrix(2, 1)
    Grid.TextMatrix(3, I) = Grid.TextMatrix(3, 1)
    Grid.TextMatrix(4, I) = Grid.TextMatrix(4, 1)
    Grid.TextMatrix(5, I) = Grid.TextMatrix(5, 1)
Next I
Grid.AddItem ""
Grid.MergeRow(6) = True
I = Grid.Rows - 1
Grid.row = I: Grid.col = 1: Grid.CellFontBold = True
quantity = I + 1
For I = 1 To Grid.Cols - 1
    Grid.TextMatrix(6, I) = "                                                                   " & _
    "����������� ��������� �� �������:"
Next I
Grid.ColAlignment(crNomZak) = flexAlignLeftCenter 'flexAlignCenterCenter 'crStatus

Grid.TextMatrix(0, crVirab) = "��������"
Grid.TextMatrix(1, crVirab) = Round(sum, 2)
If Resurs > -1 Then
    Grid.TextMatrix(2, crVirab) = res
    Grid.TextMatrix(3, crVirab) = Resurs
End If
If Resurs > 0.01 Then Grid.TextMatrix(4, crVirab) = Round(sum / Resurs, 2)
If live > -1 Then Grid.TextMatrix(5, crVirab) = Round(live, 2)

If sum > 0 Then
    
    Grid.AddItem vbTab & "� ������" & vbTab & "�" & vbTab & "������" & vbTab & _
    "��.����������" & vbTab & "%����������" & vbTab & "���������" & vbTab & _
    "��������" & vbTab & "���� ������" & vbTab & "��.���" & _
    vbTab & "��������" & vbTab & "����" & vbTab & "�������"
    Grid.row = quantity
    For I = 1 To Grid.Cols - 1
        Grid.col = I
        Grid.CellBackColor = vbButtonFace
    Next I
    
  For I = 1 To UBound(QQ)
    Grid.AddItem ""
    If QQ2(I) = 0 Then
        Grid.TextMatrix(quantity + I, crNomZak) = NN(I)
        sql = "SELECT o.ManagId, o.Logo, oe.Stat, " _
        & "o.Product, o.ProblemId, oe.outDateTime, " _
        & "f.Name, oe.workTime, o.StatusId, oe.Nevip " _
        & " FROM Orders o" _
        & " JOIN FirmGuide f ON f.FirmId = o.FirmId " _
        & " JOIN OrdersEquip oe ON o.numOrder = oe.numOrder AND oe.equipId = " & idEquip _
        & "WHERE o.numOrder = " & NN(I)
    Else '�������
        Grid.TextMatrix(quantity + I, crNomZak) = NN(I) & "o"
        sql = "SELECT o.ManagId, o.Logo, oe.StatO As Stat, " & _
        "o.Product, o.ProblemId, oc.DateTimeMO As outDateTime, " _
        & " f.Name, oe.workTimeMO As workTime" _
        & " FROM Orders o" _
        & " JOIN FirmGuide f ON f.FirmId = o.FirmId" _
        & " JOIN OrdersEquip oe ON o.numOrder = oe.numOrder AND oe.equipId = " & idEquip _
        & " LEFT JOIN OrdersInCeh oc ON o.numOrder = oc.numOrder" _
        & " WHERE o.numOrder = " & NN(I)
    End If
    Grid.TextMatrix(quantity + I, crVirab) = Round(QQ(I), 2)
    'Debug.Print sql
    Set tbOrders = myOpenRecordSet("##380", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then GoTo NXT1
    If Not tbOrders.BOF Then
        Grid.TextMatrix(quantity + I, crM) = Manag(tbOrders!ManagId)
        '������� ��� �.�� ����, ����� ���� IsNull
        If Not IsNull(tbOrders!Worktime) Then _
            Grid.TextMatrix(quantity + I, crVrVip) = Round(tbOrders!Worktime, 1)
        If IsNull(tbOrders!Stat) Then
            Grid.TextMatrix(quantity + I, crStatus) = "���"
        Else
            Grid.TextMatrix(quantity + I, crStatus) = tbOrders!Stat
        End If
        If QQ2(I) = 0 Then ' �� �������
            If tbOrders!StatusId = 5 Then
                Grid.TextMatrix(quantity + I, crStatus) = "�������"
                Grid.TextMatrix(quantity + I, crProblem) = Problems(tbOrders!ProblemId)
            End If
            If Not IsNull(tbOrders!nevip) Then _
                Grid.TextMatrix(quantity + I, crProcVip) = Round(100 * (1 - tbOrders!nevip), 1)
        End If
'        Grid.TextMatrix(quantity + i, crProblem) = Problems(tbOrders!ProblemId)
        LoadDate Grid, quantity + I, crDataVid, tbOrders!Outdatetime, "dd.mm.yy"
        LoadDate Grid, quantity + I, crVrVid, tbOrders!Outdatetime, "hh"
        Grid.TextMatrix(quantity + I, crFirma) = tbOrders!Name
        Grid.TextMatrix(quantity + I, crLogo) = tbOrders!Logo
        Grid.TextMatrix(quantity + I, crIzdelia) = tbOrders!Product
    End If
  Next I
End If

NXT1:
'���� �� �������� ���
sql = "SELECT Max(xDate) AS Prev From Itogi" _
& " WHERE xDate <'" & curDay & "' AND equipId = " & idEquip

If Not byErrSqlGetValues("##376", sql, prevDay) Then Exit Sub
cmPrev.Enabled = (prevDay <> "")

sql = "SELECT Min(xDate) AS Next From Itogi" & _
" WHERE xDate > '" & curDay & "' AND equipId = " & idEquip
If Not byErrSqlGetValues("##376", sql, nextDay) Then Exit Sub
cmNext.Enabled = (nextDay <> "")

fitFormToGrid
End Sub


Sub managStat()
Dim L As Long, I As Integer, J As Integer, line As Integer, id  As Integer
Dim str As String, strFrom As String, strWhere As String

laRecCount.Visible = False
laCount.Visible = False
Grid.Rows = 3
Grid.FixedRows = 2
Grid.MergeRow(0) = True
str = "|���-�� ���� �� �����������"
strFrom = str & str & str
str = "|���������� �������"
strFrom = strFrom & str & str
str = "|��������� ��.����������"
strFrom = strFrom & str & str
If dostup = "" Then
    Grid.FormatString = "| " & strFrom
Else
    str = "|�������� ��������"
    Grid.FormatString = "| " & strFrom & str & str
End If

Grid.TextMatrix(1, rpM2) = "M"
Grid.TextMatrix(1, rpFirmRA) = "����������"
Grid.TextMatrix(1, rpFirmKK) = "���������"
Grid.TextMatrix(1, rpFirmAll) = "   �����"
Grid.TextMatrix(1, rpQuantNoClose) = "����������"
Grid.TextMatrix(1, rpQuantAll) = "    ���"
Grid.TextMatrix(1, rpWorkNoClose) = "����������"
Grid.TextMatrix(1, rpWorkAll) = "    ���"
Grid.ColWidth(0) = 0
Grid.ColWidth(rpM2) = 660
Grid.ColWidth(rpFirmRA) = 675
Grid.ColWidth(rpFirmKK) = 810
Grid.ColWidth(rpFirmAll) = 825
Grid.ColWidth(rpQuantAll) = 825
Grid.ColWidth(rpQuantNoClose) = 825
Grid.ColWidth(rpWorkAll) = 825
Grid.ColWidth(rpWorkNoClose) = 825
If dostup <> "" Then
    Grid.TextMatrix(1, rpPaidNoClose) = "����������"
    Grid.TextMatrix(1, rpPaidAll) = "       ���"
    Grid.ColWidth(rpPaidNoClose) = 1100
    Grid.ColWidth(rpPaidAll) = 1035
End If

sql = "SELECT GuideManag.ManagId, GuideManag.Manag From GuideManag " & _
      "ORDER BY GuideManag.ForSort;"
Set Table = myOpenRecordSet("##75", sql, dbOpenForwardOnly)
If Table Is Nothing Then Exit Sub
'Table.MoveFirst
If Table.BOF Then
    Table.Close
    Exit Sub
End If
line = 2
Dim sumKK As Integer, sumRA As Integer, sumAll As Integer
sumKK = 0: sumRA = 0: sumAll = 0
While Not Table.EOF '    ********************
    Grid.TextMatrix(line, rpM2) = Table!Manag
    id = Table!ManagId
    I = getCount(id, "KK"): sumKK = sumKK + I
    Grid.TextMatrix(line, rpFirmKK) = I
    I = getCount(id, "RA"): sumRA = sumRA + I
    Grid.TextMatrix(line, rpFirmRA) = I
    I = getCount(id, "SUM"): sumAll = sumAll + I
    Grid.TextMatrix(line, rpFirmAll) = I
    I = getCountAndSumm(id, "noClose")
    Grid.TextMatrix(line, rpQuantNoClose) = I
    Grid.TextMatrix(line, rpWorkNoClose) = Format(workSum, "0.0")
    If dostup <> "" Then _
        Grid.TextMatrix(line, rpPaidNoClose) = Format(paidSum, "0.00")
    I = getCountAndSumm(id, "All")
    Grid.TextMatrix(line, rpQuantAll) = I
    Grid.TextMatrix(line, rpWorkAll) = Format(workSum, "0.0")
    If dostup <> "" Then _
        Grid.TextMatrix(line, rpPaidAll) = Format(paidSum, "0.00")
    line = line + 1
    Grid.AddItem ""
    Table.MoveNext
Wend '    ********************
Grid.RowHeight(line) = 50
Grid.AddItem ""
Grid.row = line + 1
Grid.col = rpM2: Grid.CellFontBold = True: Grid.Text = "�����:"
Grid.col = rpFirmKK: Grid.CellFontBold = True: Grid.Text = sumKK
Grid.col = rpFirmRA: Grid.CellFontBold = True: Grid.Text = sumRA
Grid.col = rpFirmAll: Grid.CellFontBold = True: Grid.Text = sumAll

Table.Close

End Sub
'$odbc15$
Function getCountAndSumm(id As Integer, Stat As String) As Integer
Dim strWhere As String, myStatId As String, str As String, I As Integer, J As Integer
getCountAndSumm = 0
workSum = 0
paidSum = 0
If Stat = "All" Then
    myStatId = 7
Else
    myStatId = 6
End If

str = Reports.tbStartDate2.Text
'strWhere = Left$(str, 2) & "/1/" & Right$(str, 4)
strWhere = "'" & Right$(str, 4) & "-" & Left$(str, 2) & "-01'"
str = Reports.tbEndDate2.Text
' ��������� ����� ������ ���� ������
I = Left$(str, 2) ' �����
J = Right$(str, 4) '���
I = I + 1:
If I > 12 Then I = 1: J = J + 1
'strWhere = strWhere & "# And (Orders.inDate)<#" & i & "/1/" & j
strWhere = strWhere & " And (Orders.inDate)< '" & Format(J, "0000") & _
"-" & Format(I, "00") & "-01'"

sql = "SELECT Count(Orders.numOrder) AS Kolvo, Sum(Orders.workTime) " & _
"AS Sum_workTime, Sum(Orders.paid) AS Sum_paid   From Orders " & _
"WHERE (((Orders.ManagId)=" & id & ") AND ((Orders.StatusId)<" & myStatId & _
") AND ((Orders.inDate)>=" & strWhere & "));"
'MsgBox sql
Set tbOrders = myOpenRecordSet("##74", sql, dbOpenForwardOnly)
If tbOrders Is Nothing Then Exit Function
If tbOrders.BOF Then GoTo EN1
getCountAndSumm = tbOrders!Kolvo
If Not IsNull(tbOrders!Sum_workTime) Then workSum = tbOrders!Sum_workTime
If Not IsNull(tbOrders!Sum_paid) Then paidSum = tbOrders!Sum_paid
EN1:
tbOrders.Close

End Function

Function getCount(id As Integer, typ As String) As Integer
Dim strWhere As String
strWhere = "(FirmGuide.Kategor)"
If typ = "KK" Then
    strWhere = "(" & strWhere & "='�') AND"
ElseIf typ = "RA" Then
    strWhere = "(" & strWhere & "='�' Or " & strWhere & "='�') AND"
Else
    strWhere = ""
End If
getCount = 0
sql = "SELECT Count(FirmGuide.FirmId) AS Kolvo From FirmGuide " & _
"WHERE (" & strWhere & " ((FirmGuide.ManagId)=" & id & "));"
'MsgBox sql
Set tbFirms = myOpenRecordSet("##458", sql, dbOpenForwardOnly)
If tbFirms Is Nothing Then Exit Function
If tbFirms.BOF Then GoTo EN1
getCount = tbFirms!Kolvo
EN1:
tbFirms.Close
End Function

'Regim = "Orders"       FindFirm    <����� "���������� ������">
'Regim = "allOrders"    FindFirm    <���."��� ������ �����">
'Regim = "FromFirms"    FirmGuide  <����� "���������� ������">
'Regim = "allFromFirms" FirmGuide  <���."��� ������ �����>
'Regim = "fromCehNaklad Nakladna    <������ ���.>
'       ������� ����.���� ��� ����� � ���� ����� � Orders
'Regim = "allOrdersByFirmName" '����� "��� ������ �����"'
'Regim = "OrdersByFirmName"    '����� "���������� ������"'
Sub firmOrders()
Dim L As Long, str As String, I As Integer, J As Integer
Dim strFirm As String, strFrom As String, strWhere As String
Grid.FormatString = "|<� ������|^M |<������|<��������|" & _
"<���� ������|<����� ������|<����|<�������|��������|��������|���������"

Grid.ColWidth(0) = 0
Grid.ColWidth(rpStatus) = 645
Grid.ColWidth(rpDataVid) = 735
Grid.ColWidth(rpVrVid) = 285
Grid.ColWidth(rpLogo) = 1890
Grid.ColWidth(rpIzdelia) = 3240 ' 3570

If Regim = "Orders" Or Regim = "allOrders" Then '�� FindFirm
    strFirm = FindFirm.lb.Text
    strWhere = "((Orders.FirmId)=" & FindFirm.FirmId & ")"
    strFrom = "FROM GuideManag INNER JOIN Orders ON GuideManag.ManagId = Orders.ManagId"
ElseIf Regim = "FromFirms" Or Regim = "allFromFirms" Then
    strFirm = firmNazv
    strWhere = "((Orders.FirmId)=" & FirmId & ")"
    strFrom = "FROM GuideManag INNER JOIN Orders ON GuideManag.ManagId = Orders.ManagId"
Else                                            '�� ����. ����
    strFirm = Orders.Grid.TextMatrix(Orders.mousRow, orFirma)
    strWhere = "((FirmGuide.Name)='" & strFirm & "')"
    strFrom = "FROM FirmGuide INNER JOIN (GuideManag INNER JOIN Orders ON GuideManag.ManagId = Orders.ManagId) ON FirmGuide.FirmId = Orders.FirmId"
End If
If Regim = "allOrdersByFirmName" Or Regim = "allOrders" Or Regim = "allFromFirms" Then
    flReportArhivOrders = True
    laHeader.Caption = "��� ������ ����� " & strFirm
Else
    laHeader.Caption = "���������� ������ ����� " & strFirm
    strWhere = "((Orders.StatusId)<>6) AND " & strWhere
End If

sql = "SELECT Orders.numOrder, Orders.StatusId, Orders.ProblemId, " & _
"Orders.DateRS, Orders.FirmId, Orders.outDateTime, Orders.Logo, " & _
"Orders.Product, Orders.ordered, Orders.paid, Orders.shipped, " & _
"GuideManag.Manag " & _
strFrom _
& " WHERE " & strWhere & " ORDER BY Orders.outDateTime"

Set tqOrders = myOpenRecordSet("##65", sql, dbOpenDynaset)
L = 1
zakazano = 0
Oplacheno = 0
Otgrugeno = 0
If tqOrders Is Nothing Then GoTo ENs
If Not tqOrders.BOF Then
  While Not tqOrders.EOF
    Grid.TextMatrix(L, rpNomZak) = tqOrders!Numorder
    J = tqOrders!StatusId
    If J = 2 Or J = 3 Or J = 9 Then
        Grid.MergeRow(L) = True
        str = Status(J) & " �� " & tqOrders!DateRS
        Grid.TextMatrix(L, rpStatus) = str
        Grid.row = L
        Grid.col = rpStatus
        Grid.CellFontBold = True
        If J = 2 Then
           Grid.CellForeColor = vbBlue
        Else
           Grid.CellForeColor = &HAA00& ' �.���.
        End If
        Grid.TextMatrix(L, rpProblem) = str
    Else
        Grid.TextMatrix(L, rpStatus) = Status(J)
        Grid.TextMatrix(L, rpProblem) = Problems(tqOrders!ProblemId)
    End If
    LoadDate Grid, L, rpDataVid, tqOrders!Outdatetime, "dd.mm.yy"
    LoadDate Grid, L, rpVrVid, tqOrders!Outdatetime, "hh"
    Grid.TextMatrix(L, rpM) = tqOrders!Manag
    Grid.TextMatrix(L, rpLogo) = tqOrders!Logo
    Grid.TextMatrix(L, rpIzdelia) = tqOrders!Product
    zakazano = zakazano + numericToReport(L, rpZakazano, tqOrders!ordered)
    Oplacheno = Oplacheno + numericToReport(L, rpOplacheno, tqOrders!paid)
    Otgrugeno = Otgrugeno + numericToReport(L, rpOtgrugeno, tqOrders!shipped)
    L = L + 1
    Grid.AddItem ""
    tqOrders.MoveNext
  Wend
End If
tqOrders.Close
ENs:
Grid.MergeRow(L) = True
str = "�����:"
Grid.TextMatrix(L, rpNomZak) = str
Grid.TextMatrix(L, rpStatus) = str
Grid.TextMatrix(L, rpProblem) = str
Grid.TextMatrix(L, rpStatus) = str
Grid.TextMatrix(L, rpProblem) = str
Grid.TextMatrix(L, rpDataVid) = str
Grid.TextMatrix(L, rpVrVid) = str
Grid.TextMatrix(L, rpLogo) = str
Grid.TextMatrix(L, rpIzdelia) = str
Grid.TextMatrix(L, rpZakazano) = Round(zakazano, 2)
Grid.TextMatrix(L, rpOplacheno) = Round(Oplacheno, 2) & " "
Grid.TextMatrix(L, rpOtgrugeno) = Round(Otgrugeno, 2)

Grid.row = L
Grid.col = 1
Grid.CellFontBold = True
Grid.col = rpZakazano
Grid.CellFontBold = True
Grid.col = rpOplacheno
Grid.CellFontBold = True
Grid.col = rpOtgrugeno
Grid.CellFontBold = True
laCount.Caption = L - 1
Grid.col = 0
End Sub

Sub fitFormToGrid()
Dim I As Long, delta As Long

I = 350 + (Grid.CellHeight + 17) * Grid.Rows
delta = I - Grid.Height
If Me.Height + delta > (Screen.Height - 400) Then _
    delta = (Screen.Height - 400) - Me.Height
Me.Height = Me.Height + delta
delta = 0
For I = 0 To Grid.Cols - 1
    delta = delta + Grid.ColWidth(I)
Next I
Me.Width = delta + 700

End Sub

Function numericToReport(row As Long, col As Integer, Value As Variant) _
As Double
    If Not IsNumeric(Value) Then
        numericToReport = 0
    Else
        numericToReport = Value
    End If
    If Round(numericToReport, 0) = numericToReport Then
        Grid.TextMatrix(row, col) = numericToReport
    Else
        Grid.TextMatrix(row, col) = Format(numericToReport, "###0.00")
    End If

End Function

Private Sub Form_Resize()
Dim H As Integer, W As Integer

If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next

H = Me.Height - oldHeight
oldHeight = Me.Height
W = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + H
Grid.Width = Grid.Width + W
laRecCount.Top = laRecCount.Top + H
laCount.Top = laCount.Top + H
laHeader.Width = laHeader.Width + W
cmExel.Top = cmExel.Top + H
cmPrint.Top = cmPrint.Top + H
cmExit.Top = cmExit.Top + H
cmExit.Left = cmExit.Left + W
cmPrev.Left = cmPrev.Left + W
cmNext.Left = cmNext.Left + W
End Sub

Private Sub Form_Unload(Cancel As Integer)
flReportArhivOrders = False
End Sub

Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If mousRow = 0 And (Regim = "KK" Or Regim = "RA") Then
    Grid.CellBackColor = Grid.BackColor
    If mousCol = 0 Then Exit Sub
    If mousCol > 3 Then
        SortCol Grid, mousCol, "numeric"
    Else
        SortCol Grid, mousCol
    End If
    Grid.row = 1    ' ������ ����� ����� ���������
End If

End Sub
Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
End Sub


Function getOrdered(numZak As String) As Double
Dim S As Double

getOrdered = -1

sql = "SELECT Sum([sDMCrez].[quantity]*[sDMCrez].[intQuant]/[sGuideNomenk].[perList]) AS cSum " & _
"FROM sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom " & _
"WHERE (((sDMCrez.numDoc)=" & numZak & "));"
If Not byErrSqlGetValues("W##209", sql, S) Then Exit Function
getOrdered = Round(S, 2)
End Function


Sub productSostav()
Dim str As String, I As Integer, delta As Integer
laHeader.Caption = "������ ������� �������, �������� � ����� " & gNzak
Grid.FormatString = "|<�����|<��������|���-��"
Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 1500
Grid.ColWidth(2) = 5000

While Not tbProduct.EOF
  Grid.AddItem Chr(9) & tbProduct!prName & Chr(9) & tbProduct!prDescript & _
  Chr(9) & "<--�������"
  Grid.row = Grid.Rows - 1: Grid.col = 1: Grid.CellFontBold = True
  Grid.col = 2: Grid.CellFontBold = True
  ReDim NN(0): ReDim QQ(0)
  gProductId = tbProduct!prId
  prExt = tbProduct!prExt
  If Not sProducts.productNomenkToNNQQ(1, 0, 0) Then GoTo NXT
  For I = 1 To UBound(NN)
    sql = "SELECT nomName From sGuideNomenk WHERE (((nomNom)='" & NN(I) & "'));"
    byErrSqlGetValues "##333", sql, str
    Grid.AddItem Chr(9) & NN(I) & Chr(9) & str & Chr(9) & Round(QQ(I), 2)
  Next I
  Grid.AddItem ""
NXT:
  tbProduct.MoveNext
Wend
Grid.RemoveItem Grid.Rows
Grid.RemoveItem 1

I = 350 + (Grid.CellHeight + 17) * Grid.Rows
delta = I - Grid.Height
If Me.Height + delta > (Screen.Height - 400) Then _
    delta = (Screen.Height - 400) - Me.Height
Me.Height = Me.Height + delta
delta = 0
For I = 0 To Grid.Cols - 1
    delta = delta + Grid.ColWidth(I)
Next I
Me.Width = delta + 700

End Sub



