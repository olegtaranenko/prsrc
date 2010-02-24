VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Report 
   BackColor       =   &H8000000A&
   Caption         =   "Отчет"
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
      Caption         =   "Печать"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   10980
      TabIndex        =   4
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Exel"
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
      Caption         =   "Число записей:"
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
'Public Regim As String
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim zakazano As Double, Oplacheno As Double, Otgrugeno As Double
Public nCols As Integer ' общее кол-во колонок
Public mousRow As Long
Public mousCol As Long
Dim workSum As Double, paidSum As Double, quantity As Long
'константы для firmOrders
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
'константы для managStat
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
'константы для whoReserved
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
Dim Ceh As String, Left As String, X As String
    GridToExcel Grid, laHeader.Caption
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
End Sub

Private Sub cmNext_Click()
virabotka "next"
End Sub

Private Sub cmPrev_Click()
virabotka "prev"
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
#If Not COMTEC = 1 Then '----------------------------------------------------
ElseIf Regim = "whoRezerved" Then
    whoRezerved
ElseIf Regim = "fromCehNaklad" Then
    productSostav
#End If '--------------------------------------------------------------
ElseIf Regim = "Virabotka" Then
    cmPrev.Visible = True
    cmNext.Visible = True
    laRecCount.Visible = False
    laCount.Visible = False
    cmExel.Visible = False
    virabotka
Else
    firmOrders
End If
End Sub
'str As String,
Sub virabotka(Optional direct As String = "")
Static prevDay As String, nextDay As String, str As String
Dim curDay As String, resurs As Double, live As Double, sum As Double
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
'делаем обычный вид даты
tmpStr = Right$(curDay, 2)
tmpStr = tmpStr & Mid$(curDay, 3, 4)
tmpStr = tmpStr & Left$(curDay, 2)
laHeader.Caption = "Выработка по цеху " & Ceh(cehId) & " на " & tmpStr

Grid.rows = 2
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

sql = "SELECT numOrder, obrazec, Virabotka From Itogi_" & Ceh(cehId) & _
" WHERE (((xDate)='" & curDay & "')) ORDER BY numOrder, obrazec DESC;"
'MsgBox sql
Set tbOrders = myOpenRecordSet("##377", sql, dbOpenForwardOnly)
If tbOrders Is Nothing Then Exit Sub
If tbOrders.BOF Then GoTo EN1
resurs = -1: live = -1
If tbOrders!Numorder = 0 Then
    resurs = Round(tbOrders!virabotka, 2)
    tbOrders.MoveNext
End If
If tbOrders.EOF Then GoTo EN2

kpd = -1
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
    QQ2(quantity) = (tbOrders!obrazec = "o") ' = -1 для образца
    QQ(quantity) = Round(tbOrders!virabotka, 2)
    sum = sum + QQ(quantity)
    tbOrders.MoveNext
Wend
EN1:
tbOrders.Close
EN2:

If direct = "" Then
    res = Zagruz.laUsed.Caption ' с учетом КПД
    resurs = Round(res / Zagruz.tbKPD.Text, 2)
Else
    res = Round(resurs * kpd_, 2) ' с учетом КПД(из хронологии)
End If
sum = Round(sum, 2)

Grid.MergeCells = flexMergeRestrictRows 'flexMergeRestrictAll 'flexMergeRestrictColumns
Grid.TextMatrix(0, 1) = "Параметр"
Grid.TextMatrix(1, 1) = "Выработка"
Grid.AddItem vbTab & "Ресурс с учетом эффективности"
Grid.AddItem vbTab & "Ресурс без учета эффективности"
Grid.AddItem vbTab & "Реальная Эффективность."
Grid.AddItem vbTab & "Сумма живых"

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
I = Grid.rows - 1
Grid.row = I: Grid.col = 1: Grid.CellFontBold = True
quantity = I + 1
For I = 1 To Grid.Cols - 1
    Grid.TextMatrix(6, I) = "                                                                   " & _
    "Детализация выработки по заказам:"
Next I
Grid.ColAlignment(crNomZak) = flexAlignLeftCenter 'flexAlignCenterCenter 'crStatus

Grid.TextMatrix(0, crVirab) = "значение"
Grid.TextMatrix(1, crVirab) = sum
If resurs > -1 Then
    Grid.TextMatrix(2, crVirab) = res
    Grid.TextMatrix(3, crVirab) = resurs
End If
If resurs > 0.01 Then Grid.TextMatrix(4, crVirab) = Round(sum / resurs, 2)
If live > -1 Then Grid.TextMatrix(5, crVirab) = live

If sum > 0 Then
    
    Grid.AddItem vbTab & "№ заказа" & vbTab & "М" & vbTab & "Статус" & vbTab & _
    "Вр.выполнения" & vbTab & "%выполнения" & vbTab & "Выработка" & vbTab & _
    "Проблемы" & vbTab & "Дата выдачи" & vbTab & "Вр.выд" & _
    vbTab & "Заказчик" & vbTab & "Лого" & vbTab & "Изделия"
    Grid.row = quantity
    For I = 1 To Grid.Cols - 1
        Grid.col = I
        Grid.CellBackColor = vbButtonFace
    Next I
    
  For I = 1 To UBound(QQ)
    Grid.AddItem ""
    If QQ2(I) = 0 Then
        Grid.TextMatrix(quantity + I, crNomZak) = NN(I)
        sql = "SELECT Orders.ManagId, Orders.Logo, OrdersInCeh.Stat, " & _
        "Orders.Product, Orders.ProblemId, Orders.outDateTime, " & _
        "GuideFirms.Name, Orders.workTime, Orders.StatusId, OrdersInCeh.Nevip " & _
        "FROM (GuideFirms INNER JOIN Orders ON (GuideFirms.FirmId = Orders.FirmId) AND (GuideFirms.FirmId = Orders.FirmId)) LEFT JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " & _
        "WHERE (((Orders.numOrder)=" & NN(I) & "));"
    Else 'образец
        Grid.TextMatrix(quantity + I, crNomZak) = NN(I) & "o"
        sql = "SELECT Orders.ManagId, Orders.Logo, OrdersMO.StatO As Stat, " & _
        "Orders.Product, Orders.ProblemId, oe.DateTimeMO As outDateTime, " _
        & "GuideFirms.Name, oe.workTimeMO As workTime " _
        & "FROM Orders " _
        & "JOIN GuideFirms ON GuideFirms.FirmId = Orders.FirmId " _
        & "LEFT JOIN vw_OrdersEquipSummary oe ON Orders.numOrder = oe.numOrder " _
        & "WHERE Orders.numOrder = " & NN(I)
    End If
    Grid.TextMatrix(quantity + I, crVirab) = QQ(I)
    
    Set tbOrders = myOpenRecordSet("##380", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then GoTo NXT1
    If Not tbOrders.BOF Then
        Grid.TextMatrix(quantity + I, crM) = Manag(tbOrders!ManagId)
        'образца уже м.не быть, тогда поле IsNull
        If Not IsNull(tbOrders!Worktime) Then _
            Grid.TextMatrix(quantity + I, crVrVip) = tbOrders!Worktime
        If IsNull(tbOrders!stat) Then
            Grid.TextMatrix(quantity + I, crStatus) = "нет"
        Else
            Grid.TextMatrix(quantity + I, crStatus) = tbOrders!stat
        End If
        If QQ2(I) = 0 Then ' не образец
            If tbOrders!StatusId = 5 Then
                Grid.TextMatrix(quantity + I, crStatus) = "отложен"
                Grid.TextMatrix(quantity + I, crProblem) = Problems(tbOrders!ProblemId)
            End If
            If Not IsNull(tbOrders!nevip) Then _
                Grid.TextMatrix(quantity + I, crProcVip) = Round(100 * (1 - tbOrders!nevip), 1)
        End If
'        Grid.TextMatrix(quantity + i, crProblem) = Problems(tbOrders!ProblemId)
        LoadDate Grid, quantity + I, crDataVid, tbOrders!outDateTime, "dd.mm.yy"
        LoadDate Grid, quantity + I, crVrVid, tbOrders!outDateTime, "hh"
        Grid.TextMatrix(quantity + I, crFirma) = tbOrders!name
        Grid.TextMatrix(quantity + I, crLogo) = tbOrders!Logo
        Grid.TextMatrix(quantity + I, crIzdelia) = tbOrders!Product
    End If
  Next I
End If

NXT1:
'есть ли соседние дни
sql = "SELECT Max(xDate) AS Prev From Itogi_" & Ceh(cehId) & _
" WHERE (((xDate)<'" & curDay & "'));"
If Not byErrSqlGetValues("##376", sql, prevDay) Then Exit Sub
cmPrev.Enabled = (prevDay <> "")

sql = "SELECT Min(xDate) AS Next From Itogi_" & Ceh(cehId) & _
" WHERE (((xDate)>'" & curDay & "'));"
If Not byErrSqlGetValues("##376", sql, nextDay) Then Exit Sub
cmNext.Enabled = (nextDay <> "")

fitFormToGrid
End Sub


Sub managStat()
Dim l As Long, I As Integer, j As Integer, line As Integer, id  As Integer
Dim str As String, strFrom As String, strWhere As String

laRecCount.Visible = False
laCount.Visible = False
Grid.rows = 3
Grid.FixedRows = 2
Grid.MergeRow(0) = True
str = "|Кол-во фирм по Справочнику"
strFrom = str & str & str
str = "|Количество заказов"
strFrom = strFrom & str & str
str = "|Суммарное вр.выполнения"
strFrom = strFrom & str & str
If dostup = "" Then
    Grid.FormatString = "| " & strFrom
Else
    str = "|Суммарно оплачено"
    Grid.FormatString = "| " & strFrom & str & str
End If

Grid.TextMatrix(1, rpM2) = "M"
Grid.TextMatrix(1, rpFirmRA) = "Рекламщики"
Grid.TextMatrix(1, rpFirmKK) = "Конечники"
Grid.TextMatrix(1, rpFirmAll) = "   Всего"
Grid.TextMatrix(1, rpQuantNoClose) = "незакрытые"
Grid.TextMatrix(1, rpQuantAll) = "    все"
Grid.TextMatrix(1, rpWorkNoClose) = "незакрытые"
Grid.TextMatrix(1, rpWorkAll) = "    все"
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
    Grid.TextMatrix(1, rpPaidNoClose) = "незакрытые"
    Grid.TextMatrix(1, rpPaidAll) = "       все"
    Grid.ColWidth(rpPaidNoClose) = 1100
    Grid.ColWidth(rpPaidAll) = 1035
End If

sql = "SELECT GuideManag.ManagId, GuideManag.Manag From GuideManag " & _
      "ORDER BY GuideManag.ForSort;"
Set table = myOpenRecordSet("##75", sql, dbOpenForwardOnly)
If table Is Nothing Then Exit Sub
'Table.MoveFirst
If table.BOF Then Exit Sub
line = 2
Dim sumKK As Integer, sumRA As Integer, sumAll As Integer
sumKK = 0: sumRA = 0: sumAll = 0
While Not table.EOF '    ********************
    Grid.TextMatrix(line, rpM2) = table!Manag
    id = table!ManagId
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
    table.MoveNext
Wend '    ********************
Grid.RowHeight(line) = 50
Grid.AddItem ""
Grid.row = line + 1
Grid.col = rpM2: Grid.CellFontBold = True: Grid.Text = "Итого:"
Grid.col = rpFirmKK: Grid.CellFontBold = True: Grid.Text = sumKK
Grid.col = rpFirmRA: Grid.CellFontBold = True: Grid.Text = sumRA
Grid.col = rpFirmAll: Grid.CellFontBold = True: Grid.Text = sumAll

table.Close

End Sub
'$odbc15$
Function getCountAndSumm(id As Integer, stat As String) As Integer
Dim strWhere As String, statId As String, str As String, I As Integer, j As Integer
getCountAndSumm = 0
workSum = 0
paidSum = 0
If stat = "All" Then
    statId = 7
Else
    statId = 6
End If

str = Reports.tbStartDate2.Text
'strWhere = Left$(str, 2) & "/1/" & Right$(str, 4)
strWhere = "'" & Right$(str, 4) & "-" & Left$(str, 2) & "-01'"
str = Reports.tbEndDate2.Text
' формируем самое начало след месяца
I = Left$(str, 2) ' месяц
j = Right$(str, 4) 'год
I = I + 1:
If I > 12 Then I = 1: j = j + 1
'strWhere = strWhere & "# And (Orders.inDate)<#" & i & "/1/" & j
strWhere = strWhere & " And (Orders.inDate)< '" & Format(j, "0000") & _
"-" & Format(I, "00") & "-01'"

sql = "SELECT Count(Orders.numOrder) AS Kolvo, Sum(Orders.workTime) " & _
"AS Sum_workTime, Sum(Orders.paid) AS Sum_paid   From Orders " & _
"WHERE (((Orders.ManagId)=" & id & ") AND ((Orders.StatusId)<" & statId & _
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
strWhere = "(GuideFirms.Kategor)"
If typ = "KK" Then
    strWhere = "(" & strWhere & "='К') AND"
ElseIf typ = "RA" Then
    strWhere = "(" & strWhere & "='П' Or " & strWhere & "='Д') AND"
Else
    strWhere = ""
End If
getCount = 0
sql = "SELECT Count(GuideFirms.FirmId) AS Kolvo From GuideFirms " & _
"WHERE (" & strWhere & " ((GuideFirms.ManagId)=" & id & "));"
'MsgBox sql
Set tbFirms = myOpenRecordSet("##458", sql, dbOpenForwardOnly)
If tbFirms Is Nothing Then Exit Function
If tbFirms.BOF Then GoTo EN1
getCount = tbFirms!Kolvo
EN1:
tbFirms.Close
End Function

'Regim = "Orders"       FindFirm    <Отчет "Незакрытые заказы">
'Regim = "allOrders"    FindFirm    <Отч."Все заказы фирмы">
'Regim = "FromFirms"    GuideFirms  <Отчет "Незакрытые заказы">
'Regim = "allFromFirms" GuideFirms  <Отч."Все заказы фирмы>
'Regim = "fromCehNaklad Nakladna    <Состав изд.>
'       Команды конт.меню при клике в поле Фирма в Orders
'Regim = "allOrdersByFirmName" 'Отчет "Все заказы Фирмы"'
'Regim = "OrdersByFirmName"    'Отчет "Незакрытые заказы"'
Sub firmOrders()
Dim l As Long, str As String, I As Integer, j As Integer
Dim strFirm As String, strFrom As String, strWhere As String
Grid.FormatString = "|<№ заказа|^M |<Статус|<Проблемы|" & _
"<Дата выдачи|<Время выдачи|<Лого|<Изделия|Заказано|Оплачено|Отгружено"

Grid.ColWidth(0) = 0
Grid.ColWidth(rpStatus) = 645
Grid.ColWidth(rpDataVid) = 735
Grid.ColWidth(rpVrVid) = 285
Grid.ColWidth(rpLogo) = 1890
Grid.ColWidth(rpIzdelia) = 3240 ' 3570

If Regim = "Orders" Or Regim = "allOrders" Then 'из FindFirm
    strFirm = FindFirm.lb.Text
    strWhere = "((Orders.FirmId)=" & FindFirm.firmId & ")"
    strFrom = "FROM GuideManag INNER JOIN Orders ON GuideManag.ManagId = Orders.ManagId"
ElseIf Regim = "FromFirms" Or Regim = "allFromFirms" Then
    strFirm = GuideFirms.Grid.TextMatrix(GuideFirms.mousRow, gfNazwFirm)
    strWhere = "((Orders.FirmId)=" & GuideFirms.Grid.TextMatrix(GuideFirms.mousRow, gfId) & ")"
    strFrom = "FROM GuideManag INNER JOIN Orders ON GuideManag.ManagId = Orders.ManagId"
Else                                            'из конт. меню
    strFirm = Orders.Grid.TextMatrix(Orders.mousRow, orFirma)
    strWhere = "((GuideFirms.Name)='" & strFirm & "')"
    strFrom = "FROM GuideFirms INNER JOIN (GuideManag INNER JOIN Orders ON GuideManag.ManagId = Orders.ManagId) ON GuideFirms.FirmId = Orders.FirmId"
End If
If Regim = "allOrdersByFirmName" Or Regim = "allOrders" Or Regim = "allFromFirms" Then
    flReportArhivOrders = True
    laHeader.Caption = "Все заказы фирмы " & strFirm
Else
    laHeader.Caption = "Незакрытые заказы фирмы " & strFirm
    strWhere = "((Orders.StatusId)<>6) AND " & strWhere
End If

sql = "SELECT Orders.numOrder, Orders.StatusId, Orders.ProblemId, " & _
"Orders.DateRS, Orders.FirmId, Orders.outDateTime, Orders.Logo, " & _
"Orders.Product, Orders.ordered, Orders.paid, Orders.shipped, " & _
"GuideManag.Manag " & _
strFrom & " WHERE (" & strWhere & ") ORDER BY Orders.outDateTime;"

Set tqOrders = myOpenRecordSet("##65", sql, dbOpenDynaset)
l = 1
zakazano = 0
Oplacheno = 0
Otgrugeno = 0
If tqOrders Is Nothing Then GoTo ENs
If Not tqOrders.BOF Then
  While Not tqOrders.EOF
    Grid.TextMatrix(l, rpNomZak) = tqOrders!Numorder
    j = tqOrders!StatusId
    If j = 2 Or j = 3 Or j = 9 Then
        Grid.MergeRow(l) = True
        str = status(j) & " на " & tqOrders!dateRS
        Grid.TextMatrix(l, rpStatus) = str
        Grid.row = l
        Grid.col = rpStatus
        Grid.CellFontBold = True
        If j = 2 Then
           Grid.CellForeColor = vbBlue
        Else
           Grid.CellForeColor = &HAA00& ' т.зел.
        End If
        Grid.TextMatrix(l, rpProblem) = str
    Else
        Grid.TextMatrix(l, rpStatus) = status(j)
        Grid.TextMatrix(l, rpProblem) = Problems(tqOrders!ProblemId)
    End If
    LoadDate Grid, l, rpDataVid, tqOrders!outDateTime, "dd.mm.yy"
    LoadDate Grid, l, rpVrVid, tqOrders!outDateTime, "hh"
    Grid.TextMatrix(l, rpM) = tqOrders!Manag
    Grid.TextMatrix(l, rpLogo) = tqOrders!Logo
    Grid.TextMatrix(l, rpIzdelia) = tqOrders!Product
    zakazano = zakazano + numericToReport(l, rpZakazano, tqOrders!ordered)
    Oplacheno = Oplacheno + numericToReport(l, rpOplacheno, tqOrders!paid)
    Otgrugeno = Otgrugeno + numericToReport(l, rpOtgrugeno, tqOrders!shipped)
    l = l + 1
    Grid.AddItem ""
    tqOrders.MoveNext
  Wend
End If
tqOrders.Close
ENs:
Grid.MergeRow(l) = True
str = "Итого:"
Grid.TextMatrix(l, rpNomZak) = str
Grid.TextMatrix(l, rpStatus) = str
Grid.TextMatrix(l, rpProblem) = str
Grid.TextMatrix(l, rpStatus) = str
Grid.TextMatrix(l, rpProblem) = str
Grid.TextMatrix(l, rpDataVid) = str
Grid.TextMatrix(l, rpVrVid) = str
Grid.TextMatrix(l, rpLogo) = str
Grid.TextMatrix(l, rpIzdelia) = str
Grid.TextMatrix(l, rpZakazano) = Round(zakazano, 2)
Grid.TextMatrix(l, rpOplacheno) = Round(Oplacheno, 2) & " "
Grid.TextMatrix(l, rpOtgrugeno) = Round(Otgrugeno, 2)

Grid.row = l
Grid.col = 1
Grid.CellFontBold = True
Grid.col = rpZakazano
Grid.CellFontBold = True
Grid.col = rpOplacheno
Grid.CellFontBold = True
Grid.col = rpOtgrugeno
Grid.CellFontBold = True
laCount.Caption = l - 1
Grid.col = 0
End Sub

Sub fitFormToGrid()
Dim I As Long, delta As Long

I = 350 + (Grid.CellHeight + 17) * Grid.rows
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

Function numericToReport(row As Long, col As Integer, value As Variant) _
As Double
    If Not IsNumeric(value) Then
        numericToReport = 0
    Else
        numericToReport = value
    End If
    If Round(numericToReport, 0) = numericToReport Then
        Grid.TextMatrix(row, col) = numericToReport
    Else
        Grid.TextMatrix(row, col) = Format(numericToReport, "###0.00")
    End If

End Function

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next

h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w
laRecCount.Top = laRecCount.Top + h
laCount.Top = laCount.Top + h
laHeader.Width = laHeader.Width + w
cmExel.Top = cmExel.Top + h
cmPrint.Top = cmPrint.Top + h
cmExit.Top = cmExit.Top + h
cmExit.Left = cmExit.Left + w
cmPrev.Left = cmPrev.Left + w
cmNext.Left = cmNext.Left + w
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
    Grid.row = 1    ' только чтобы снять выделение
End If

End Sub
Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
End Sub

Sub whoRezerved()
Dim v, s As Double, ed2 As String, per As Double, sum As Double
sql = "SELECT  ed_Izmer2, perList From sGuideNomenk WHERE (((nomNom)='" & gNomNom & "'));"
'MsgBox sql
If Not byErrSqlGetValues("##349", sql, ed2, per) Then Exit Sub

laHeader.Caption = "Список заказов, кот. резервировали ном-ру '" & gNomNom & _
"' [" & ed2 & "]."
Grid.FormatString = "|>№ заказа|кол-во|^Цех |^Дата |^ М|<Статус" & _
"|<Название Фирмы|Изделия|заказано|согласовано"
laRecCount.Caption = "Сумма резервов:"
Grid.ColWidth(0) = 0
Grid.ColWidth(rtReserv) = 765
Grid.ColWidth(rtCeh) = 765
Grid.ColWidth(rtData) = 870
Grid.ColWidth(rtStatus) = 930
Grid.ColWidth(rtFirma) = 3270
Grid.ColWidth(rtProduct) = 1950
Grid.ColWidth(rtOplacheno) = 810

sql = "SELECT Orders.numOrder, GuideCeh.Ceh, Orders.inDate, " & _
"GuideManag.Manag, Orders.Product, " & _
"GuideStatus.Status, GuideFirms.Name, Orders.ordered, Orders.paid, " & _
"sDMCrez.quantity, Sum(sDMC.quant) AS Sum_quant " & _
"FROM GuideStatus INNER JOIN (GuideManag INNER JOIN (GuideFirms INNER " & _
"JOIN (GuideCeh INNER JOIN (sDMC RIGHT JOIN (sDMCrez INNER JOIN Orders " & _
"ON sDMCrez.numDoc = Orders.numOrder) ON (sDMC.nomNom = sDMCrez.nomNom) " & _
"AND (sDMC.numDoc = sDMCrez.numDoc)) ON GuideCeh.CehId = Orders.CehId) " & _
"ON GuideFirms.FirmId = Orders.FirmId) ON GuideManag.ManagId = " & _
"Orders.ManagId) ON GuideStatus.StatusId = Orders.StatusId " & _
"Where(((sDMCrez.nomNom) = '" & gNomNom & "')) " & _
"GROUP BY Orders.numOrder, GuideCeh.Ceh, Orders.inDate, GuideManag.Manag, " & _
"GuideStatus.Status, GuideFirms.Name, Orders.Product, Orders.ordered, Orders.paid, sDMCrez.quantity;"
'MsgBox sql
Set tbOrders = myOpenRecordSet("##139", sql, dbOpenForwardOnly) ', dbOpenDynaset)
If tbOrders Is Nothing Then Exit Sub
quantity = 0: sum = 0
If Not tbOrders.BOF Then
 While Not tbOrders.EOF
    v = tbOrders!Sum_quant: If IsNull(v) Then v = 0
    s = Round((tbOrders!quantity - v) / per, 2)
    If s > 0 Then
        quantity = quantity + 1
        Grid.TextMatrix(quantity, rtNomZak) = tbOrders!Numorder
        Grid.TextMatrix(quantity, rtCeh) = tbOrders!Ceh
        LoadDate Grid, quantity, rtData, tbOrders!inDate, "dd.mm.yy"
        Grid.TextMatrix(quantity, rtMen) = tbOrders!Manag
        Grid.TextMatrix(quantity, rtStatus) = tbOrders!status
        Grid.TextMatrix(quantity, rtFirma) = tbOrders!name
        
        If Not IsNull(tbOrders!Product) Then _
            Grid.TextMatrix(quantity, rtProduct) = tbOrders!Product
        If Not IsNull(tbOrders!ordered) Then _
            Grid.TextMatrix(quantity, rtZakazano) = Round(tbOrders!ordered, 2)
        If Not IsNull(tbOrders!paid) Then _
            Grid.TextMatrix(quantity, rtOplacheno) = Round(tbOrders!paid, 2)
        Grid.TextMatrix(quantity, rtReserv) = s
    
        Grid.AddItem ""
        sum = sum + s
    End If
    tbOrders.MoveNext
 Wend
End If
tbOrders.Close

'выписанные в цеху накладные со склада целых
sql = "SELECT sDMCrez.numDoc, sDMCrez.quantity, sDocs.Note, sDocs.xDate " & _
"FROM sDocs INNER JOIN sDMCrez ON sDocs.numDoc = sDMCrez.numDoc " & _
"Where (((sDMCrez.nomNom) = '" & gNomNom & "') And ((sDocs.numExt) = 0));"

Set tbOrders = myOpenRecordSet("##342", sql, dbOpenForwardOnly) ', dbOpenDynaset)
If Not tbOrders Is Nothing Then
  If Not tbOrders.BOF Then
    While Not tbOrders.EOF
        quantity = quantity + 1
        Grid.TextMatrix(quantity, rtNomZak) = tbOrders!numDoc
        Grid.TextMatrix(quantity, rtCeh) = tbOrders!Note
        LoadDate Grid, quantity, rtData, tbOrders!xDate, "dd.mm.yy"
'        Grid.TextMatrix(quantity, rtStatus) = tbOrders!status
        Grid.TextMatrix(quantity, rtFirma) = "Выписанная в Цеху накладная"
        Grid.TextMatrix(quantity, rtReserv) = Round(tbOrders!quantity / per, 2)
        Grid.AddItem ""
        tbOrders.MoveNext
    Wend
  End If
End If
tbOrders.Close

'заказы на продажу
sql = "SELECT BayOrders.numOrder, BayOrders.inDate, BayOrders.ManagId, " & _
"BayOrders.StatusId, BayGuideProblem.Problem, BayGuideFirms.Name, " & _
" BayOrders.ordered, BayOrders.paid, sDMCrez.quantity, sDMC.quant " & _
"FROM BayGuideFirms INNER JOIN (BayGuideProblem INNER JOIN ((sDMCrez " & _
"LEFT JOIN sDMC ON (sDMCrez.nomNom = sDMC.nomNom) AND (sDMCrez.numDoc = sDMC.numDoc)) INNER JOIN BayOrders ON sDMCrez.numDoc = BayOrders.numOrder) ON BayGuideProblem.ProblemId = " & _
"BayOrders.ProblemId) ON BayGuideFirms.FirmId = BayOrders.FirmId " & _
"WHERE (((sDMCrez.nomNom)='" & gNomNom & "'));"

Set tbOrders = myOpenRecordSet("##350", sql, dbOpenForwardOnly) ', dbOpenDynaset)
If Not tbOrders Is Nothing Then
  If Not tbOrders.BOF Then
    While Not tbOrders.EOF
      v = tbOrders!quant: If IsNull(v) Then v = 0
      s = Round((tbOrders!quantity - v) / per, 2)
      If s > 0 Then
        quantity = quantity + 1
        Grid.TextMatrix(quantity, rtNomZak) = tbOrders!Numorder
        Grid.TextMatrix(quantity, rtCeh) = "Продажа"
        LoadDate Grid, quantity, rtData, tbOrders!inDate, "dd.mm.yy"
        Grid.TextMatrix(quantity, rtMen) = Manag(tbOrders!ManagId)
        Grid.TextMatrix(quantity, rtStatus) = status(tbOrders!StatusId)
        Grid.TextMatrix(quantity, rtFirma) = tbOrders!name
        Grid.TextMatrix(quantity, rtReserv) = s
'        If Not IsNull(tbOrders!ordered) Then _
            Grid.TextMatrix(quantity, rtZakazano) = tbOrders!ordered
         Grid.TextMatrix(quantity, rtZakazano) = Round(getOrdered(tbOrders!Numorder), 2)
        
        If Not IsNull(tbOrders!paid) Then _
            Grid.TextMatrix(quantity, rtOplacheno) = Round(tbOrders!paid, 2)
        Grid.AddItem ""
      End If
      tbOrders.MoveNext
    Wend
  End If
  tbOrders.Close
End If

laCount.Caption = Round(sum, 2)
If quantity > 0 Then
    Grid.removeItem quantity + 1
End If
trigger = False
SortCol Grid, rtReserv, "numeric"
Grid.Visible = True
Me.MousePointer = flexDefault

End Sub

Function getOrdered(numZak As String) As Double
Dim s As Double

getOrdered = -1

sql = "SELECT Sum([sDMCrez].[quantity]*[sDMCrez].[intQuant]/[sGuideNomenk].[perList]) AS cSum " & _
"FROM sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom " & _
"WHERE (((sDMCrez.numDoc)=" & numZak & "));"
If Not byErrSqlGetValues("W##209", sql, s) Then Exit Function
getOrdered = Round(s, 2)
End Function


Sub productSostav()
Dim str As String, I As Integer, delta As Integer
laHeader.Caption = "Состав готовых изделий, входящих в заказ " & gNzak
Grid.FormatString = "|<Номер|<Описание|кол-во"
Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 1500
Grid.ColWidth(2) = 5000

While Not tbProduct.EOF
  Grid.AddItem Chr(9) & tbProduct!prName & Chr(9) & tbProduct!prDescript & _
  Chr(9) & "<--Изделие"
  Grid.row = Grid.rows - 1: Grid.col = 1: Grid.CellFontBold = True
  Grid.col = 2: Grid.CellFontBold = True
  ReDim NN(0): ReDim QQ(0)
  gProductId = tbProduct!prId
  prExt = tbProduct!prExt
  If Not sProducts.productNomenkToNNQQ(1, 0, 0) Then GoTo NXT
  For I = 1 To UBound(NN)
    sql = "SELECT nomName From sGuideNomenk WHERE (((nomNom)='" & NN(I) & "'));"
    byErrSqlGetValues "##333", sql, str
    Grid.AddItem Chr(9) & NN(I) & Chr(9) & str & Chr(9) & QQ(I)
  Next I
  Grid.AddItem ""
NXT:
  tbProduct.MoveNext
Wend
Grid.removeItem Grid.rows
Grid.removeItem 1

I = 350 + (Grid.CellHeight + 17) * Grid.rows
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



