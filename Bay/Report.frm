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
   Begin VB.CommandButton cmPrint 
      Caption         =   "Печать"
      Height          =   315
      Left            =   2760
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
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   3780
      TabIndex        =   3
      Top             =   7800
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7212
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11652
      _ExtentX        =   20553
      _ExtentY        =   12721
      _Version        =   393216
      MergeCells      =   2
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
      Height          =   432
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   11772
   End
   Begin VB.Label laRecCount 
      Caption         =   "Число записей:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   7830
      Width           =   1335
   End
   Begin VB.Label laCount 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   7800
      Width           =   615
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
Dim zakazano As Single, Oplacheno As Single, Otgrugeno As Single
Public nCols As Integer ' общее кол-во колонок
Public mousRow As Long
Public mousCol As Long
Dim workSum As Single, paidSum As Single, quantity As Long
'константы для firmOrders
Const rpNomZak = 1
Const rpM = 2
Const rpStatus = 3
Const rpProblem = 4
Const rpDataVid = 5
Const rpVrVid = 6
Const rpZakazano = 7
Const rpOplacheno = 8
Const rpOtgrugeno = 9
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
'Const rtEdIzm = 2
Const rtCeh = 3
Const rtData = 4
Const rtMen = 5
Const rtStatus = 6
Const rtFirma = 7
Const rtProduct = 8
Const rtZakazano = 9
Const rtOplacheno = 10

Private Sub cmExel_Click()
'If InStr(Regim, "Orders") > 0 Then
    GridToExcel Grid, laHeader.Caption
'ElseIf Regim = "KK" Or Regim = "RA" Then
'    GridToExcel Grid, laHeader.Caption
'ElseIf Regim = "Manag" Then
'    GridToExcel Grid, laHeader.Caption
'End If
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmPrint_Click()
Me.PrintForm

End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width
If Regim = "stat" Then
    statistic
ElseIf Regim = "whoRezerved" Then
    whoRezerved
Else
    firmOrders
End If
End Sub

Sub statistic(Optional year As String = "")
Dim nRow As Long, nCol As Long, str As String, i As Integer, j As Integer
Dim iMonth As Integer, iYear As Integer, iCount As Integer, strWhere As String
Dim nMonth As Integer, nYear As Integer, mCount As Integer, lastCol As Integer
Dim wtSum As Single, paidSum As Single, orderSum As Single, visits As Integer, visitSum As Integer
Dim year01 As Integer, year02 As Integer, year03 As Integer, year04 As Integer
Dim errCurYear As Integer, errBefYear As Integer ', whereByTemaAndType As String
errCurYear = 0:   errBefYear = 0
Dim mDate As Date

'whereByTemaAndType = ""
If year = "" Then
 str = Reports.tbStartDate.Text
 Report.laHeader.Caption = "Статистика посещений фирм за период с " & str & _
                " по " & Reports.tbEndDate.Text
 nMonth = left$(str, 2)
 nYear = right$(str, 4)
 mCount = DateDiff("m", str, Reports.tbEndDate.Text) + 1

 str = "|<Название фирмы|^М |Регион|Скидки"
 iCount = mCount
 lastCol = 5 ' в 2х местах
 iMonth = nMonth

 Do
    If iMonth = 13 Then iMonth = 1
    str = str & "|" & Format(iMonth, "00")
    iMonth = iMonth + 1
    lastCol = lastCol + 1
    iCount = iCount - 1
 Loop While iCount > 0
 str = str & "|Итого|Вр.вып|Заказано|Оплачено"
 Report.Grid.FormatString = str
 Report.Grid.ColWidth(0) = 0
 Report.Grid.ColWidth(1) = 1875
 Report.Grid.ColWidth(3) = 1605
'Grid.ColWidth(lastCol + 2) = 795
 Report.nCols = lastCol + 2
  
 nRow = 1
Else
 nMonth = 1
 nYear = lastYear - 3
 mCount = DateDiff("m", "01.01." & nYear, CurDate) + 1
 strWhere = ""
End If

sql = "SELECT f.FirmId, f.Name, isnull(r.region, '') as Kategor, f.year01, f.year02, f.year03, f.year04, f.Sale, f.ManagId " _
& " FROM BayGuideFirms f" _
& " left join bayRegion r on f.regionid = r.regionid" _
& strWhere _
& " order by f.name"

'MsgBox sql
Set tbFirms = myOpenRecordSet("##68", sql, dbOpenDynaset) 'ForwardOnly)
If tbFirms Is Nothing Then Exit Sub
If tbFirms.BOF Then GoTo EN1:
tbFirms.MoveFirst
While Not tbFirms.EOF '                         *******************
 If year <> "all" Then
     mDate = "01." & Reports.tbStartDate.Text
 Else
     mDate = "01.01." & lastYear - 3
 End If
 iMonth = nMonth
 iYear = nYear
 iCount = mCount
 visitSum = 0
 wtSum = 0
 paidSum = 0
 orderSum = 0
 bilo = False
 nCol = 5 ' в 2х местах
 year01 = 0: year02 = 0: year03 = 0: year04 = 0
 Do '$$6
    str = "(inDate)"
    strWhere = str & " >= '" & Format(mDate, "yyyy-mm-dd 00:00:00") & "' AND "
    tmpDate = DateAdd("m", 1, mDate)
    strWhere = strWhere & str & " < '" & Format(tmpDate, "yyyy-mm-dd 00:00:00") & "'"

    str = Format(iMonth, "00") & "." & iYear
    sql = "SELECT paid,numOrder From BayOrders " & _
    "WHERE ((" & strWhere & ") AND " & _
    "(Not ((StatusId)=0 Or (StatusId)=7)) AND " & _
    "((FirmId)=" & tbFirms!firmId & "));"
'Debug.Print sql
    Set tbOrders = myOpenRecordSet("##69", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    visits = 0:
    If Not tbOrders.BOF Then
'      tbOrders.MoveFirst
      While Not tbOrders.EOF
          If year = "" Then
            visits = visits + 1
            If Not IsNull(tbOrders!paid) Then _
                    paidSum = paidSum + tbOrders!paid
'             If Not IsNull(tbOrders!ordered) Then _
                    orderSum = orderSum + tbOrders!ordered'$$6
               orderSum = orderSum + getOrdered(tbOrders!numOrder) '$$6 tbOrders!ordered
          Else
            If iYear = lastYear - 3 Then
                year01 = year01 + 1 'не исп-ся
            ElseIf iYear = lastYear - 2 Then
                year02 = year02 + 1
            ElseIf iYear = lastYear - 1 Then
                year03 = year03 + 1
            ElseIf iYear = lastYear Then
                year04 = year04 + 1
            End If
          End If
          tbOrders.MoveNext
      Wend
      tbOrders.Close
      
    End If
    If visits > 0 And year = "" Then
        If Not bilo Then
            Report.Grid.TextMatrix(nRow, 1) = tbFirms!Name
            If Not IsNull(tbFirms!managId) Then _
                    Report.Grid.TextMatrix(nRow, 2) = Manag(tbFirms!managId)
            Report.Grid.TextMatrix(nRow, 3) = tbFirms!Kategor
            If Not IsNull(tbFirms!Sale) Then _
                    Report.Grid.TextMatrix(nRow, 4) = tbFirms!Sale
            bilo = True
        End If
        Report.Grid.TextMatrix(nRow, nCol) = visits
        visitSum = visitSum + visits
'        nCol = nCol + 1
    End If
    
    If iMonth = 12 Then
        iMonth = 1
        iYear = iYear + 1
    Else
        iMonth = iMonth + 1
    End If
    mDate = DateAdd("m", 1, mDate)
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
 If year <> "" Then
    tbFirms.Edit
    If tbFirms!year01 <> year01 Then errBefYear = errBefYear + 1
    tbFirms!year01 = year01
    If tbFirms!year02 <> year02 Then errBefYear = errBefYear + 1
    tbFirms!year02 = year02
    If tbFirms!year03 <> year03 Then errBefYear = errBefYear + 1
    tbFirms!year03 = year03
    If tbFirms!year04 <> year04 Then errCurYear = errCurYear + 1
    tbFirms!year04 = year04
    tbFirms.Update
 End If
 tbFirms.MoveNext
Wend '*******************
EN1:
tbFirms.Close
'tbOrders.Close
If year = "" Then
  If nRow > 1 Then Report.Grid.RemoveItem (nRow)
  Report.laCount.Caption = nRow - 1
Else
'  If errBefYear > 0 Then
'     MsgBox "В прошлых годах обнаружено " & errBefYear & " фирм с неверно " & _
'     "подсчитанным количеством посещений.  Все ошибки устранены.", , "Обнаружены ошибки"
'  End If
'  If errCurYear > 0 Then
'     MsgBox "В текущем году обнаружено " & errCurYear & " фирм с неверно " & _
'     "подсчитанным количеством посещений.  Все ошибки устранены.", , "Обнаружены ошибки"
'  End If
End If
End Sub

'3052715
Sub whoRezerved()
Dim v, s As Single, ed2 As String, per As Single, sum As Single  ', ed1 As String, obr As String
'obrez, ed_Izmer, obr, ed1,
sql = "SELECT  ed_Izmer2, perList From sGuideNomenk " & _
"WHERE (((nomNom)='" & gNomNom & "'));"
'MsgBox sql
If Not byErrSqlGetValues("##349", sql, ed2, per) Then Exit Sub
'If obr <> "" Then ed1 = ed2
laHeader.Caption = "Список заказов, кот. резервировали ном-ру '" & gNomNom & _
"' [" & ed2 & "]."
Grid.FormatString = "|>№ заказа|кол-во|^Цех |^Дата |^ М" & _
"|<Статус|<Название Фирмы|<Изделия|заказано|согласовано"
Grid.ColWidth(0) = 0
'Grid.ColWidth(rtNomZak) =
'Grid.ColWidth(rtReserv) = 765
'Grid.ColWidth(rtEdIzm)=
Grid.ColWidth(rtCeh) = 765
Grid.ColWidth(rtData) = 870
'Grid.ColWidth(rtMen) =
Grid.ColWidth(rtStatus) = 930
Grid.ColWidth(rtFirma) = 3270
Grid.ColWidth(rtProduct) = 1950
'Grid.ColWidth(rtZakazano) =
Grid.ColWidth(rtOplacheno) = 810

'******************************* Prior заказы
quantity = 0: sum = 0
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
If Not tbOrders.BOF Then
 While Not tbOrders.EOF
    v = tbOrders!Sum_quant: If IsNull(v) Then v = 0
    s = Round((tbOrders!quantity - v) / per, 2)
    If s > 0 Then
        quantity = quantity + 1
        Grid.TextMatrix(quantity, rtNomZak) = tbOrders!numOrder
        Grid.TextMatrix(quantity, rtCeh) = tbOrders!Ceh
    '    Grid.TextMatrix(quantity, rtData) = tbOrders!inDate
        LoadDate Grid, quantity, rtData, tbOrders!inDate, "dd.mm.yy"
        Grid.TextMatrix(quantity, rtMen) = tbOrders!Manag
        Grid.TextMatrix(quantity, rtStatus) = tbOrders!status
        Grid.TextMatrix(quantity, rtFirma) = tbOrders!Name
        
        If Not IsNull(tbOrders!Product) Then _
            Grid.TextMatrix(quantity, rtProduct) = tbOrders!Product
        If Not IsNull(tbOrders!ordered) Then _
            Grid.TextMatrix(quantity, rtZakazano) = tbOrders!ordered 'Prior
        If Not IsNull(tbOrders!paid) Then _
            Grid.TextMatrix(quantity, rtOplacheno) = tbOrders!paid
        Grid.TextMatrix(quantity, rtReserv) = s
    
        Grid.AddItem ""
        sum = sum + s
    End If
    tbOrders.MoveNext
 Wend
End If
tbOrders.Close

'выписанные в цеху накладные но не межскладские
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
'"BayOrders.ordered, BayOrders.paid, sDMCrez.quantity, sDMC.quant " & $$6
sql = "SELECT BayOrders.numOrder, BayOrders.inDate, BayOrders.ManagId, " & _
"BayOrders.StatusId, BayGuideProblem.Problem, BayGuideFirms.Name, " & _
"BayOrders.paid, sDMCrez.quantity, sDMC.quant " & _
"FROM BayGuideFirms INNER JOIN (BayGuideProblem INNER JOIN ((sDMCrez LEFT JOIN sDMC ON (sDMCrez.nomNom = sDMC.nomNom) AND (sDMCrez.numDoc = " & _
"sDMC.numDoc)) INNER JOIN BayOrders ON sDMCrez.numDoc = BayOrders.numOrder) " & _
"ON BayGuideProblem.ProblemId = " & _
"BayOrders.ProblemId) ON BayGuideFirms.FirmId = BayOrders.FirmId " & _
"WHERE (((sDMCrez.nomNom)='" & gNomNom & "'));"

Set tbOrders = myOpenRecordSet("##351", sql, dbOpenForwardOnly) ', dbOpenDynaset)
If Not tbOrders Is Nothing Then
  If Not tbOrders.BOF Then
    While Not tbOrders.EOF
      v = tbOrders!quant: If IsNull(v) Then v = 0
      s = Round((tbOrders!quantity - v) / per, 2)
      If s > 0 Then
        quantity = quantity + 1
        Grid.TextMatrix(quantity, rtNomZak) = tbOrders!numOrder
        Grid.TextMatrix(quantity, rtCeh) = "Продажа"
        LoadDate Grid, quantity, rtData, tbOrders!inDate, "dd.mm.yy"
        Grid.TextMatrix(quantity, rtMen) = Manag(tbOrders!managId)
        Grid.TextMatrix(quantity, rtStatus) = status(tbOrders!StatusId)
        Grid.TextMatrix(quantity, rtFirma) = tbOrders!Name
        Grid.TextMatrix(quantity, rtReserv) = s
'        If Not IsNull(tbOrders!ordered) Then _
            Grid.TextMatrix(quantity, rtZakazano) = tbOrders!ordered  $$6
        Grid.TextMatrix(quantity, rtZakazano) = getOrdered(tbOrders!numOrder) '$$6
            
        If Not IsNull(tbOrders!paid) Then _
            Grid.TextMatrix(quantity, rtOplacheno) = tbOrders!paid
        Grid.AddItem ""
      End If
      tbOrders.MoveNext
    Wend
  End If
  tbOrders.Close
End If


laCount.Caption = Round(sum, 2)
laRecCount.Caption = "Сумма резервов:"

If quantity > 0 Then
    Grid.RemoveItem quantity + 1
End If
trigger = False
SortCol Grid, rtReserv, "numeric"
Grid.Visible = True
Me.MousePointer = flexDefault

End Sub

'Regim = "Orders"       FindFirm    <Отчет "Незакрытые заказы">
'Regim = "allOrders"    FindFirm    <Отч."Все заказы фирмы">
'Regim = "FromFirms"    GuideFirms  <Отчет "Незакрытые заказы">
'Regim = "allFromFirms" GuideFirms  <Отч."Все заказы фирмы>
'Regim = "fromCehNaklad Nakladna    <Состав изд.>
'       Команды конт.меню при клике в поле Фирма в Orders
'Regim = "allOrdersByFirmName" 'Отчет "Все заказы Фирмы"'
'Regim = "OrdersByFirmName"    'Отчет "Незакрытые заказы"'
Sub firmOrders()
Dim l As Long, str As String, i As Integer, j As Integer
Dim strFirm As String, strFrom As String, strWhere As String
Grid.FormatString = "|<№ заказа|^M |<Статус|<Проблемы|" & _
"<Дата выдачи|<Время выдачи|Заказано|Оплачено|Отгружено"

Grid.ColWidth(0) = 0
'Grid.ColWidth(rpNomZak) = 840
Grid.ColWidth(rpStatus) = 720
Grid.ColWidth(rpProblem) = 975
Grid.ColWidth(rpDataVid) = 1095
Grid.ColWidth(rpVrVid) = 615

If Regim = "Orders" Or Regim = "allOrders" Then 'из FindFirm
    strFirm = FindFirm.lb.Text
    strWhere = "((BayOrders.FirmId)=" & FindFirm.firmId & ")"
'    strFrom = "From Orders"
    strFrom = "FROM GuideManag INNER JOIN BayOrders ON GuideManag.ManagId = BayOrders.ManagId"
ElseIf Regim = "FromFirms" Or Regim = "allFromFirms" Then
    strFirm = GuideFirms.Grid.TextMatrix(GuideFirms.mousRow, gfNazwFirm)
    strWhere = "((BayOrders.FirmId)=" & GuideFirms.Grid.TextMatrix(GuideFirms.mousRow, gfId) & ")"
'    strFrom = "From Orders"
    strFrom = "FROM GuideManag INNER JOIN BayOrders ON GuideManag.ManagId = BayOrders.ManagId"
Else                                            'из конт. меню
    strFirm = Orders.Grid.TextMatrix(Orders.mousRow, orFirma)
    strWhere = "((BayGuideFirms.Name)='" & strFirm & "')"
    strFrom = "FROM BayGuideFirms INNER JOIN (GuideManag INNER JOIN BayOrders ON GuideManag.ManagId = BayOrders.ManagId) ON BayGuideFirms.FirmId = BayOrders.FirmId"
'    strFrom = "From BayGuideFirms INNER JOIN BayOrders ON BayGuideFirms.FirmId = BayOrders.FirmId"
End If
If Regim = "allOrdersByFirmName" Or Regim = "allOrders" Or Regim = "allFromFirms" Then
    flReportArhivOrders = True
    laHeader.Caption = "Все заказы фирмы " & strFirm
Else
    laHeader.Caption = "Незакрытые заказы фирмы " & strFirm
    strWhere = "((BayOrders.StatusId)<>6) AND " & strWhere
End If
'"BayOrders.ordered, BayOrders.paid, BayOrders.shipped, " & $$6
sql = "SELECT BayOrders.numOrder, BayOrders.StatusId, BayOrders.ProblemId, " & _
"BayOrders.FirmId, BayOrders.outDateTime, " & _
"BayOrders.paid, " & _
"GuideManag.Manag " & _
strFrom & " WHERE (" & strWhere & ") ORDER BY BayOrders.outDateTime;"
'MsgBox sql
Set tqOrders = myOpenRecordSet("##65", sql, dbOpenDynaset)
l = 1
zakazano = 0
Oplacheno = 0
Otgrugeno = 0
If tqOrders Is Nothing Then GoTo ENs
If Not tqOrders.BOF Then
  While Not tqOrders.EOF
    
'  Grid.MergeRow(2) = True

    Grid.TextMatrix(l, rpNomZak) = tqOrders!numOrder
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
        Grid.TextMatrix(l, rpProblem) = Problems(tqOrders!problemId)
    End If
    LoadDate Grid, l, rpDataVid, tqOrders!outDateTime, "dd.mm.yy"
    LoadDate Grid, l, rpVrVid, tqOrders!outDateTime, "hh"
    Grid.TextMatrix(l, rpM) = tqOrders!Manag
    'zakazano = zakazano + numericToReport(l, rpZakazano, tqOrders!ordered) '$$6
    zakazano = zakazano + numericToReport(l, rpZakazano, getOrdered(tqOrders!numOrder)) '$$6
    Oplacheno = Oplacheno + numericToReport(l, rpOplacheno, tqOrders!paid)
    'Otgrugeno = Otgrugeno + numericToReport(l, rpOtgrugeno, tqOrders!shipped) '$$6
    Otgrugeno = Otgrugeno + numericToReport(l, rpOtgrugeno, getShipped(tqOrders!numOrder)) '$$6
    l = l + 1
    Grid.AddItem ""
    tqOrders.MoveNext
  Wend
'  If l > 1 Then Grid.RemoveItem (l)
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

Function numericToReport(row As Long, col As Integer, value As Variant) _
As Single
    If Not IsNumeric(value) Then
        numericToReport = 0
    Else
        numericToReport = value
    End If
    Grid.TextMatrix(row, col) = numericToReport

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
cmExit.left = cmExit.left + w
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

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor

End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
End Sub

