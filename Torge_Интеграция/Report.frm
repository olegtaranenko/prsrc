VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Report 
   BackColor       =   &H8000000A&
   Caption         =   "Отчет"
   ClientHeight    =   8184
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11808
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8184
   ScaleWidth      =   11808
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmPrint 
      Caption         =   "Печать"
      Height          =   315
      Left            =   3060
      TabIndex        =   4
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   10980
      TabIndex        =   2
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Exel"
      Height          =   315
      Left            =   3960
      TabIndex        =   1
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
   Begin VB.Label laRecCount 
      Caption         =   "Число записей:"
      Height          =   195
      Left            =   2220
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label laCount 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label laRecSum 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1980
      TabIndex        =   6
      Top             =   7800
      Width           =   915
   End
   Begin VB.Label laSum 
      Alignment       =   1  'Right Justify
      Caption         =   "Сумма:"
      Height          =   195
      Left            =   660
      TabIndex        =   5
      Top             =   7860
      Width           =   1035
   End
   Begin VB.Label laHeader 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Regim As String, param1 As String, param2 As String, param3 As String
'Public Regim As String
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Public nCols As Integer ' общее кол-во колонок
Public mousRow As Long
Public mousCol As Long
Dim quantity As Long
Dim Cena()  As Single

Const rrNumOrder = 1
Const rrDate = 2
Const rrFirm = 3
Const rrProduct = 4
Const rrMater = 5
Const rrReliz = 6
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

Const rzZatratName = 1
Const rzMainCosts = 2
Const rzAddCosts = 3

Const zdDate = 1
Const zdSumm = 2
Const zdProvodka = 3
Const zdAgent = 4
Const zdNazn = 5
Const zdUtochn = 6

'otlaDwkdh - отладочная база, дебаг режим


'если col <> "" - проверяется, какая колонка
Sub laSumControl(Optional col As String = "")
If col <> "" And Grid.col <> rrFirm Then GoTo AA
If InStr(Regim, "tatistic") Then
   laSum.Caption = "Кол-во фирм:"
   If col = "" Then laRecSum.Caption = Grid.Rows - 1
Else
AA:
   laSum.Caption = "Сумма:"
End If
End Sub

Sub fitFormToGrid()
Dim i As Long, delta As Long

i = 350 + (Grid.CellHeight + 17) * Grid.Rows
delta = i - Grid.Height
If Me.Height + delta > (Screen.Height - 900) Then _
    delta = (Screen.Height - 900) - Me.Height
Me.Height = Me.Height + delta
'Grid.Height = i
delta = 0
For i = 0 To Grid.Cols - 1
    delta = delta + Grid.ColWidth(i)
Next i
Me.Width = delta + 700

End Sub


Private Sub cmExel_Click()
GridToExcel Grid, laHeader.Caption
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmPrint_Click()
Me.PrintForm

End Sub

Private Sub Form_Load()
Dim prevDate As Date, prevNom As Long

oldHeight = Me.Height
oldWidth = Me.Width
Me.MousePointer = flexHourglass
If Regim = "subDetail" Then
    laHeader.Caption = "Детализация сумм " & param3 & "  по отгрузке от " & _
    Left$(param2, 8) & " по заказу №" & gNzak
    subDetail
ElseIf Regim = "subDetailMat" Then
    laHeader.Caption = "Детализация суммы" & param3 & " по накладной №" & gNzak
    subDetail
ElseIf Regim = "aReport" Then
    laHeader.Caption = "Отчет 'А' на " & Format(Now(), "dd.mm.yy")
    aReport
ElseIf Regim = "whoRezerved" Then
    whoRezerved
ElseIf Regim = "" Then 'отчет Реализация - заказы производства
    laHeader.Caption = "Детализация сумм " & param2 & "(Материалы) и " & _
    param1 & "(Реализация) по датам отгрузок заказов производства."
    relizDetail
ElseIf Regim = "relizStatistic" Then 'отчет Реализация - заказы производства
    laHeader.Caption = "Детализация сумм " & param2 & "(Материалы) и " & _
    param1 & "(Реализация) по фирмам."
    relizDetail "statistic"
ElseIf Regim = "uslug" Then 'отчет Реализация - заказы производства
    laHeader.Caption = "Детализация суммы " & param1 & "(Услуги)" & _
    " по датам отгрузок заказов производства."
    uslugDetail
ElseIf Regim = "uslugStatistic" Then 'отчет Реализация - заказы производства
    laHeader.Caption = "Детализация суммы " & param1 & "(Услуги)" & _
    " по датам отгрузок заказов производства."
    uslugDetail "statistic"
ElseIf Regim = "bay" Then 'отчет Реализация - заказы продаж
    laHeader.Caption = "Детализация сумм " & param2 & "(Материалы) и " & _
    param1 & "(Реализация) по датам списания под заказы продаж."
    relizDetailBay
ElseIf Regim = "bayStatistic" Then 'отчет Реализация - заказы продаж
    laHeader.Caption = "Детализация сумм " & param2 & "(Материалы) и " & _
    param1 & "(Реализация) по фирмам."
    relizDetailBay "statistic"
ElseIf Regim = "mat" Then 'отчет Реализация - материалы не под заказы
    laHeader.Caption = "Детализация суммы " & _
    param1 & " по датам списания материалов не под заказы."
    relizDetailMat
ElseIf Regim = "venture" Then 'детализация и статистика по предприятиям
    laHeader.Caption = "Детализация сумм " & param2 & "(Материалы) и " & _
    param1 & "(Реализация) по """ & Pribil.ventureId & """"
    ventureReport Pribil.ventureId
ElseIf Regim = "ventureZatrat" Then 'статистика по предприятиям
    laHeader.Caption = "Детализация сумм по основным и вспомогательным затратам  по """ & Pribil.ventureId & """"
    ventureZatrat Pribil.ventureId
ElseIf Regim = "ventureZatratDetail" Then 'детализация по предприятиям
    laHeader.Caption = "Детализация сумм по статье затрат """ & Grid.TextMatrix(Grid.row, rzZatratName) & """ для """ & Pribil.ventureId & """"
    ventureZatratDetail Pribil.ventureId, Grid.TextMatrix(Grid.row, 0)
End If
laSumControl
If InStr(Regim, "tatistic") Then
    trigger = False
    SortCol Grid, rrReliz, "numeric"
End If
fitFormToGrid
Me.MousePointer = flexDefault
End Sub

Sub ventureZatratDetail(ventureId As Integer, id_shiz As String)
Dim sum As Single

    Grid.FormatString = "|Дата|>Сумма|Проводка|Через|Назначение|Уточнение"
    Grid.ColWidth(0) = 0
    Grid.ColWidth(zdDate) = 850
    Grid.ColWidth(zdSumm) = 1000
    Grid.ColWidth(zdProvodka) = 1200
    Grid.ColWidth(zdAgent) = 2300
    Grid.ColWidth(zdNazn) = 3000
    Grid.ColWidth(zdUtochn) = 3000
    'Grid.ColWidth() =
    
    sql = "select xdate, uesumm, b.debit + '-' + b.subdebit + ' => ' + b.kredit + '-' + b.subkredit as provodka" _
    & " , k.name, p.pdescript as nazn, b.descript as utochn" _
    & " from ybook b" _
    & " join ydebkreditor k on k.id = b.kreddebitor" _
    & " join yguidepurpose p on p.debit = b.debit and p.subdebit = b.subdebit and p.kredit = b.kredit and p.subkredit = b.subkredit and p.pid = b.purposeid" _
    & " where id_shiz is not null " _
    & " and ventureid = " & ventureId & " and id_shiz = " & param1
 
    If Pribil.costsDateWhere <> "" Then
        sql = sql & " and " & Pribil.costsDateWhere
    End If
    sql = sql & " order by xdate, provodka, uesumm desc"

    'Debug.Print sql
    
    Set tbOrders = myOpenRecordSet("##vnt_det", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    quantity = 0: sum = 0
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF
            quantity = quantity + 1
            Grid.TextMatrix(quantity, zdDate) = tbOrders!xDate
            Grid.TextMatrix(quantity, zdSumm) = Format(tbOrders!uesumm, "## ##0.00")
            Grid.TextMatrix(quantity, zdProvodka) = tbOrders!provodka
            If Not IsNull(tbOrders!Name) Then Grid.TextMatrix(quantity, zdAgent) = tbOrders!Name
            If Not IsNull(tbOrders!nazn) Then Grid.TextMatrix(quantity, zdNazn) = tbOrders!nazn
            If Not IsNull(tbOrders!utochn) Then Grid.TextMatrix(quantity, zdUtochn) = tbOrders!utochn
            
            Grid.AddItem ""
            tbOrders.MoveNext
        Wend
    End If
    Grid.row = quantity + 1
    Grid.col = rzMainCosts: Grid.CellFontBold = True
    Grid.col = rzAddCosts: Grid.CellFontBold = True
    Grid.TextMatrix(quantity + 1, rzMainCosts) = Format(param2, "## ##0.00")
    Grid.TextMatrix(quantity + 1, rzAddCosts) = Format(param1, "## ##0.00")

End Sub


Sub ventureZatrat(ventureId As Integer)
Dim sum As Single

    Grid.FormatString = "|Наименование|>Осн.Затраты|>Вспом.Затраты"
    Grid.ColWidth(0) = 0
    Grid.ColWidth(rzZatratName) = 3600
    Grid.ColWidth(rzMainCosts) = 1500
    Grid.ColWidth(rzAddCosts) = 1500
    
    sql = "select sum(uesumm) as sm, ventureid, is_main_costs, s.nm as nm, b.id_shiz" _
    & " from ybook b" _
    & " join shiz s on s.id = b.id_shiz" _
    & " where " _
    & "     ventureid = " & ventureId
 
    If Pribil.costsDateWhere <> "" Then
        sql = sql & " and " & Pribil.costsDateWhere
    End If
    sql = sql & " group by ventureid, is_main_costs, nm, b.id_shiz" _
    & " order by nm"

    'Debug.Print sql
    
    Set tbOrders = myOpenRecordSet("##vnt_det", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    quantity = 0: sum = 0
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF
            quantity = quantity + 1
            Grid.TextMatrix(quantity, rzZatratName) = tbOrders!nm
            Grid.TextMatrix(quantity, 0) = tbOrders!id_shiz
            
            If tbOrders!is_main_costs = 1 Then
                Grid.TextMatrix(quantity, rzMainCosts) = Format(tbOrders!sm, "## ##0.00")
            Else
                Grid.TextMatrix(quantity, rzAddCosts) = Format(tbOrders!sm, "## ##0.00")
            End If
            Grid.AddItem ""
            tbOrders.MoveNext
        Wend
    End If
    Grid.row = quantity + 1
    Grid.col = rzMainCosts: Grid.CellFontBold = True
    Grid.col = rzAddCosts: Grid.CellFontBold = True
    Grid.TextMatrix(quantity + 1, rzMainCosts) = Format(param2, "## ##0.00")
    Grid.TextMatrix(quantity + 1, rzAddCosts) = Format(param1, "## ##0.00")
End Sub

Sub ventureReport(ventureId As Integer)
Dim sum As Single

    Grid.FormatString = "|Заказ|<Дата|<Фирма|<Комментарий|>Затраты|>Реализация"
    Grid.ColWidth(0) = 250
    Grid.ColWidth(rrNumOrder) = 885
    Grid.ColWidth(rrDate) = 765
    Grid.ColWidth(rrFirm) = 3855
    Grid.ColWidth(rrProduct) = 0
    Grid.ColWidth(rrMater) = 1005
    Grid.ColWidth(rrReliz) = 1005
    
    sql = "SELECT * from vPredmetyOutSummary where ventureid = " & ventureId
    If Pribil.nDateWhere <> "" Then
        sql = sql & " and " & Pribil.nDateWhere
    End If
    sql = sql & " order by outdate"
    Set tbOrders = myOpenRecordSet("##vnt_det", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    quantity = 0: sum = 0
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF
            quantity = quantity + 1
            Grid.TextMatrix(quantity, rrNumOrder) = tbOrders!numOrder
            Grid.TextMatrix(quantity, rrDate) = Format(tbOrders!outDate, "dd/mm/yy hh:nn:ss")
            Grid.TextMatrix(quantity, rrFirm) = tbOrders!firmName
            'Grid.TextMatrix(quantity, rrProduct) = tbOrders!Text
            If tbOrders!Type = 1 Then
                Grid.TextMatrix(quantity, 0) = "p"
            ElseIf tbOrders!Type = 2 Then
                Grid.TextMatrix(quantity, 0) = "n"
            ElseIf tbOrders!Type = 3 Then
                Grid.TextMatrix(quantity, 0) = "w"
            ElseIf tbOrders!Type = 4 Then
                Grid.TextMatrix(quantity, 0) = "u"
            ElseIf tbOrders!Type = 8 Then
                Grid.TextMatrix(quantity, 0) = "b"
            End If
            
            Grid.TextMatrix(quantity, rrMater) = Format(tbOrders!prime_cost, "## ##0.00")
            Grid.TextMatrix(quantity, rrReliz) = Format(tbOrders!shipped, "## ##0.00")
            Grid.AddItem ""
            tbOrders.MoveNext
        Wend
    End If

End Sub


Sub whoRezerved()
Dim v, s As Single, ed2 As String, per As Single, sum As Single
', obr As String
sql = "SELECT  ed_Izmer2, perList From sGuideNomenk WHERE (((nomNom)='" & gNomNom & "'));"
'MsgBox sql
If Not byErrSqlGetValues("##349", sql, ed2, per) Then Exit Sub

laHeader.Caption = "Список заказов, кот. резервировали ном-ру '" & gNomNom & _
"' [" & ed2 & "]."
Grid.FormatString = "|<№ заказа|кол-во|^Цех |^Дата |^ М|Статус" & _
"|<Название Фирмы|<Изделия|заказано|согласовано"
Grid.ColWidth(0) = 0
'Grid.ColWidth(rtNomZak) =
Grid.ColWidth(rtReserv) = 765
Grid.ColWidth(rtCeh) = 765
Grid.ColWidth(rtData) = 870
'Grid.ColWidth(rtMen) =
Grid.ColWidth(rtStatus) = 930
Grid.ColWidth(rtFirma) = 3270
Grid.ColWidth(rtProduct) = 1950
'Grid.ColWidth(rtZakazano) =
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
    v = tbOrders!sum_quant: If IsNull(v) Then v = 0
    s = Round((tbOrders!quantity - v) / per, 2)
    If s > 0 Then
        quantity = quantity + 1
        Grid.TextMatrix(quantity, rtNomZak) = tbOrders!numOrder
        Grid.TextMatrix(quantity, rtCeh) = tbOrders!Ceh
    '    Grid.TextMatrix(quantity, rtData) = tbOrders!inDate
        LoadDate Grid, quantity, rtData, tbOrders!inDate, "dd.mm.yy"
        Grid.TextMatrix(quantity, rtMen) = tbOrders!Manag
        Grid.TextMatrix(quantity, rtStatus) = tbOrders!Status
        Grid.TextMatrix(quantity, rtFirma) = tbOrders!Name
        
        If Not IsNull(tbOrders!Product) Then _
            Grid.TextMatrix(quantity, rtProduct) = tbOrders!Product
        If Not IsNull(tbOrders!ordered) Then _
            Grid.TextMatrix(quantity, rtZakazano) = tbOrders!ordered
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
        Grid.TextMatrix(quantity, rtCeh) = tbOrders!note
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

'If obr <> "" Then obr = "2"
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
      If s > 0 And tbOrders!StatusId <> 6 Then
        quantity = quantity + 1
        Grid.TextMatrix(quantity, rtNomZak) = tbOrders!numOrder
        Grid.TextMatrix(quantity, rtCeh) = "Продажа"
        LoadDate Grid, quantity, rtData, tbOrders!inDate, "dd.mm.yy"
        Grid.TextMatrix(quantity, rtMen) = Manag(tbOrders!ManagId)
        Grid.TextMatrix(quantity, rtStatus) = Status(tbOrders!StatusId)
        Grid.TextMatrix(quantity, rtFirma) = tbOrders!Name
        Grid.TextMatrix(quantity, rtReserv) = s
        If Not IsNull(tbOrders!ordered) Then _
            Grid.TextMatrix(quantity, rtZakazano) = tbOrders!ordered
        If Not IsNull(tbOrders!paid) Then _
            Grid.TextMatrix(quantity, rtOplacheno) = tbOrders!paid
        Grid.AddItem ""
      End If
      tbOrders.MoveNext
    Wend
  End If
  tbOrders.Close
End If

laCount.Caption = quantity
laRecSum.Caption = Round(sum, 2)
If quantity > 0 Then
    Grid.RemoveItem quantity + 1
End If
trigger = False
SortCol Grid, rtReserv, "numeric"
Grid.Visible = True
Me.MousePointer = flexDefault

End Sub




Sub aReport()
Dim s As Single, k As Single, d As Single, sumD As Single, sumK As Single
Dim s2 As Single

Grid.FormatString = "||ПЛЮС|МИНУС|"
Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 4000
Grid.ColWidth(2) = 1060
Grid.ColWidth(3) = 1060
Grid.ColWidth(4) = 1200
'сделать исправления в переводом цен на Целые кол-ва ном-р !!!
sumD = 0: sumK = 0
Grid.TextMatrix(1, 1) = "Склад - максимальный запас"
sql = "SELECT " _
    & " Sum((if mark='Used' then Zakup else nowOstatki endif) " _
    & " * cost/perList) AS sum " _
    & "FROM sGuideNomenk "
'Debug.Print sql

byErrSqlGetValues "##387", sql, s
Grid.TextMatrix(1, 3) = Round(-s, 2)
sumK = sumK - s

sql = "SELECT Sum([cost]*[nowOstatki]/[perList]) AS sum FROM sGuideNomenk;"
byErrSqlGetValues "##388", sql, s
Grid.AddItem Chr(9) & "Склад -фактический запас" & Chr(9) & Round(s, 2)
sumD = sumD + s


' сумма списанной и еще неотгруженной ном-ры по незакрытим заказам !!!
' здесь не м.б. заказов Продаж т.к. у них отгрузка - это списание
sql = "SELECT Sum([sDMC].[quant]*[sGuideNomenk].[cost]/[sGuideNomenk].[perList]) AS sum " & _
"FROM Orders INNER JOIN (sGuideNomenk INNER JOIN sDMC ON " & _
"sGuideNomenk.nomNom = sDMC.nomNom) ON Orders.numOrder = sDMC.numDoc " & _
"WHERE (((StatusId)<6));"
byErrSqlGetValues "##386", sql, s
s = s - otgruzNomenk()
Grid.AddItem Chr(9) & "Незавершенное производство" & Chr(9) & Round(s, 2)
sumD = sumD + s

s = Round(schetOstat("60"), 2)
If s < 0 Then k = -s: s = 0 Else k = 0
Grid.AddItem Chr(9) & "Товары в пути" & Chr(9) & s
Grid.AddItem Chr(9) & "Долги по товарам" & Chr(9) & Chr(9) & k
sumD = sumD + s
sumK = sumK + k

s = Round(schetOstat("51"), 2)
If s < 0 Then k = -s: s = 0 Else k = 0
Grid.AddItem Chr(9) & "Р/счет" & Chr(9) & s & Chr(9) & k
sumD = sumD + s
sumK = sumK + k

s = Round(schetOstat("50"), 2)
If s < 0 Then k = -s: s = 0 Else k = 0
Grid.AddItem Chr(9) & "Касса" & Chr(9) & s & Chr(9) & k
sumD = sumD + s
sumK = sumK + k

s = Round(schetOstat("57"), 2)
If s < 0 Then k = -s: s = 0 Else k = 0
Grid.AddItem Chr(9) & "Долги" & Chr(9) & s & Chr(9) & k
sumD = sumD + s
sumK = sumK + k


d = 0: k = 0
sql = "SELECT Sum(if paid > shipped then shipped - paid endif ) AS k" _
        & "    , Sum(if paid < shipped then paid - shipped endif) AS d " _
        & "    from Orders WHERE StatusId < 6 "
'не вылавливает строки где paid или shipped Null
Debug.Print sql

byErrSqlGetValues "##392", sql, k, d
s = 0: s2 = 0
sql = "SELECT Sum(if paid > shipped then shipped - paid endif ) AS k" _
        & "    , Sum(if paid < shipped then paid - shipped endif) AS d " _
        & "    from BayOrders WHERE StatusId < 6 "
'не вылавливает строки где paid или shipped Null
byErrSqlGetValues "##392", sql, s, s2
k = k + s
d = d + s2

s = 0
sql = "SELECT Sum(shipped) AS Sum_shipped from Orders " & _
"WHERE (((paid) Is Null) AND ((StatusId)<6));"
byErrSqlGetValues "##393", sql, s
d = d + s
s = 0
sql = "SELECT Sum(shipped) AS Sum_shipped from bayOrders " & _
"WHERE (((paid) Is Null) AND ((StatusId)<6));"
byErrSqlGetValues "##393", sql, s
d = d + s
Grid.AddItem Chr(9) & "Дебиторы" & Chr(9) & d

s = 0
sql = "SELECT Sum(paid) AS Sum_paid from Orders " & _
"WHERE (((shipped) Is Null) AND ((StatusId)<6));"
byErrSqlGetValues "##394", sql, s
k = k + s
s = 0
sql = "SELECT Sum(paid) AS Sum_paid from bayOrders " & _
"WHERE (((shipped) Is Null) AND ((StatusId)<6));"
byErrSqlGetValues "##394", sql, s
k = k + s
Grid.AddItem Chr(9) & "Кредиторы" & Chr(9) & Chr(9) & k
sumD = Round(sumD + d, 2)
sumK = Round(sumK + k, 2)

Grid.AddItem Chr(9) & "                                       ИТОГО:" & _
Chr(9) & sumD & Chr(9) & sumK & Chr(9) & Round(sumD - sumK, 2)
Grid.row = Grid.Rows - 1
Grid.col = 1: Grid.CellFontBold = True
Grid.col = 2: Grid.CellFontBold = True
Grid.col = 3: Grid.CellFontBold = True
Grid.col = 4: Grid.CellFontBold = True

End Sub

Function schetOstat(schet As String)
Dim d As Single, k As Single

schetOstat = 0
sql = "SELECT Sum(begDebit) AS Sum_begDebit, Sum(begKredit) AS Sum_begKredit " & _
"From yGuideSchets GROUP BY number HAVING (((number)=" & schet & "));"

If Not byErrSqlGetValues("W##389", sql, d, k) Then GoTo EN1 '$$4 в самом начале счета м.и не быть
schetOstat = d - k

d = 0: k = 0
sql = "SELECT Sum(UEsumm) AS Sum_UEsumm from yBook " & _
"WHERE (((Debit)=" & schet & "));"
If Not byErrSqlGetValues("##390", sql, d) Then GoTo EN1

sql = "SELECT Sum(UEsumm) AS Sum_UEsumm from yBook " & _
"WHERE (((Kredit)=" & schet & "));"
If Not byErrSqlGetValues("##391", sql, k) Then GoTo EN1
schetOstat = schetOstat + d - k

EN1:
End Function

Sub subDetail()
Dim str As String, i As Integer, delta As Integer, ed_izm As String
Dim str2 As String, str3 As String
'Const rdNomer = 1
'Const rdName = 2
'Const rdQuant = 3
'str = Left$(param2, 8)
'laHeader.Caption = "Детализация отгрузки от " & str & " по заказу " & gNzak
Grid.FormatString = "|<Номер|<Описание|кол-во в одном |кол-во общее|" & _
"<Ед.измерения|Цена|Сумма|Реализация"
Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 1500
Grid.ColWidth(2) = 3840
Grid.ColWidth(3) = 765
Grid.ColWidth(4) = 720
Grid.ColWidth(5) = 420
Grid.ColWidth(6) = 1080

'strWhere = Left$(param2, 6) & "20" & Mid$(param2, 7)
strWhere = "20" & Mid$(param2, 7, 2) & "-" & Mid$(param2, 4, 2) & "-" & _
Left$(param2, 2) & Mid$(param2, 9)

If param1 = "p" Or param1 = "w" Then 'есть  гот.изделия
'  sql = "SELECT r.prId, r.prExt, " & _
  "r.quant, sGuideProducts.prName, sGuideProducts.prDescript " & _
  "FROM xPredmetyByIzdeliaOut INNER JOIN sGuideProducts ON r.prId = sGuideProducts.prId " & _

  sql = "SELECT r.prId, r.prExt, " & _
  "r.quant, sGuideProducts.prName, " & _
  "sGuideProducts.prDescript, p.cenaEd " & _
  "FROM sGuideProducts INNER JOIN (xPredmetyByIzdelia p INNER JOIN xPredmetyByIzdeliaOut r ON (p.prExt = r.prExt) AND (p.prId = r.prId) AND (p.numOrder = r.numOrder)) ON sGuideProducts.prId = p.prId " & _
  "WHERE (((r.numOrder)=" & gNzak & ") AND " & _
  "((r.outDate) ='" & strWhere & "'));"
'  "((r.outDate) Like  '" & strWhere & "*'));"
  'MsgBox sql
  Set tbProduct = myOpenRecordSet("##382", sql, dbOpenForwardOnly)
  If tbProduct Is Nothing Then Exit Sub

    
  While Not tbProduct.EOF
    Grid.AddItem Chr(9) & tbProduct!prName & Chr(9) & tbProduct!prDescript & _
    Chr(9) & "<--Изделие" & Chr(9) & tbProduct!quant & Chr(9) & "шт." & Chr(9) & _
    "(" & Round(tbProduct!cenaEd, 2) & ")" & Chr(9) & Chr(9) & _
    Round(tbProduct!quant * tbProduct!cenaEd, 2)
    Grid.row = Grid.Rows - 1: Grid.col = 1: Grid.CellFontBold = True
    Grid.col = 2: Grid.CellFontBold = True
    ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): ReDim QQ3(0): ReDim Cena(0)
    gProductId = tbProduct!prId
    prExt = tbProduct!prExt
    If Not productNomenkToNNQQ(1, tbProduct!quant) Then GoTo NXT '1
    For i = 1 To UBound(NN)
      sql = "SELECT nomName, ed_izmer, Size, cod From sGuideNomenk WHERE (((nomNom)='" & NN(i) & "'));"
      byErrSqlGetValues "##333", sql, str, ed_izm, str2, str3
      Grid.AddItem Chr(9) & NN(i) & Chr(9) & str3 & " " & str & " " & str2 & Chr(9) & QQ(i) & _
      Chr(9) & QQ2(i) & Chr(9) & ed_izm & Chr(9) & Cena(i) & Chr(9) & QQ3(i)
    Next i
    Grid.AddItem ""
NXT:
    tbProduct.MoveNext
  Wend
  tbProduct.Close
End If
'  Grid.RemoveItem Grid.Rows
'  Grid.RemoveItem 1
If param1 = "n" Or param1 = "w" Then
  sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.cost, " & _
  "sGuideNomenk.ed_izmer, sGuideNomenk.Size, sGuideNomenk.cod, " & _
  "sGuideNomenk.perList, xPredmetyByNomenk.cenaEd, xPredmetyByNomenkOut.quant " & _
  "FROM sGuideNomenk INNER JOIN (xPredmetyByNomenk INNER JOIN xPredmetyByNomenkOut ON (xPredmetyByNomenk.nomNom = xPredmetyByNomenkOut.nomNom) AND (xPredmetyByNomenk.numOrder = xPredmetyByNomenkOut.numOrder)) ON sGuideNomenk.nomNom = xPredmetyByNomenk.nomNom " & _
  "WHERE (((xPredmetyByNomenkOut.numOrder)=" & gNzak & ") AND " & _
  "((xPredmetyByNomenkOut.outDate) =  '" & strWhere & "'));"
'  "((xPredmetyByNomenkOut.outDate) Like  '" & strWhere & "*'));"
  
  Set tbNomenk = myOpenRecordSet("##383", sql, dbOpenDynaset)
  If tbNomenk Is Nothing Then Exit Sub
  While Not tbNomenk.EOF '!!!
    Grid.AddItem Chr(9) & tbNomenk!nomnom & Chr(9) & tbNomenk!cod & " " & _
    tbNomenk!nomName & " " & tbNomenk!Size & _
    Chr(9) & "<--Номенклатура" & Chr(9) & tbNomenk!quant & Chr(9) & _
    tbNomenk!ed_izmer & Chr(9) & _
    Round(tbNomenk!cost / tbNomenk!perList, 2) & " (" & Round(tbNomenk!cenaEd, 2) & ")" & Chr(9) & _
    Round(tbNomenk!quant * tbNomenk!cost / tbNomenk!perList, 2) & Chr(9) & _
    Round(tbNomenk!quant * tbNomenk!cenaEd, 2)
    Grid.row = Grid.Rows - 1: Grid.col = 1: Grid.CellFontBold = True
    Grid.col = 2: Grid.CellFontBold = True
    Grid.AddItem ""

'NXT2:
    tbNomenk.MoveNext
  Wend
  tbNomenk.Close
End If

If param1 = "b" Then
  Grid.ColWidth(3) = 0
'  sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.cost, " & _
  "sGuideNomenk.ed_izmer2, sGuideNomenk.Size, sGuideNomenk.cod, " & _
  "sGuideNomenk.perList, sDMCrez.quantity, BayOrders.shipped, sDMCrez.numDoc " & _
  "FROM BayOrders INNER JOIN (sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom) ON BayOrders.numOrder = sDMCrez.numDoc " & _
  "WHERE (((sDMCrez.numDoc)=" & gNzak & "));"
  sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.cost, " & _
  "sGuideNomenk.ed_izmer2, sGuideNomenk.Size, sGuideNomenk.cod, " & _
  "sGuideNomenk.perList, sDMC.quant, sDMCrez.intQuant,  sDMCrez.numDoc " & _
  "FROM sGuideNomenk INNER JOIN ((BayOrders INNER JOIN sDocs ON BayOrders.numOrder = sDocs.numDoc) INNER JOIN (sDMC INNER JOIN sDMCrez ON sDMC.nomNom = sDMCrez.nomNom) ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc) AND (BayOrders.numOrder = sDMCrez.numDoc)) ON sGuideNomenk.nomNom = sDMC.nomNom " & _
  "WHERE (((sDMCrez.numDoc)=" & gNzak & ") AND " & _
  "((dateformat(sDocs.xDate, 'yyyy-mm-dd hh:nn:ss')) = '" & strWhere & "'));"
  
  Set tbNomenk = myOpenRecordSet("##432", sql, dbOpenDynaset)
  If tbNomenk Is Nothing Then Exit Sub
'  Grid.AddItem Chr(9) & Chr(9) & "Отгружено по заказу:" & Chr(9) & Chr(9) & _
  Chr(9) & Chr(9) & Chr(9) & Chr(9) & _
  Round(tbNomenk!quant * tbNomenk!intQuant / tbNomenk!perList, 2)
'  Grid.row = 2: Grid.col = 2: Grid.CellFontBold = True
  While Not tbNomenk.EOF '!!!
    Grid.AddItem Chr(9) & tbNomenk!nomnom & Chr(9) & tbNomenk!cod & " " & _
    tbNomenk!nomName & " " & tbNomenk!Size & _
    Chr(9) & "<--Номенклатура" & Chr(9) & _
    Round(tbNomenk!quant / tbNomenk!perList, 2) & Chr(9) & _
    tbNomenk!ed_Izmer2 & Chr(9) & tbNomenk!cost & Chr(9) & _
    Round(tbNomenk!quant * tbNomenk!cost / tbNomenk!perList, 2) & Chr(9) & _
    Round(tbNomenk!quant * tbNomenk!intQuant / tbNomenk!perList, 2)
    tbNomenk.MoveNext
  Wend
  tbNomenk.Close
  Debug.Print sql
End If

If param1 = "m" Then
  Grid.ColWidth(3) = 0
  Grid.ColWidth(8) = 0
  sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.cod, sGuideNomenk.nomName, " & _
  "sGuideNomenk.Size, sDMC.quant, sGuideNomenk.cost, sGuideNomenk.perList, " & _
  "sGuideNomenk.ed_Izmer2 " & _
  "FROM sGuideNomenk INNER JOIN sDMC ON sGuideNomenk.nomNom = sDMC.nomNom " & _
  "GROUP BY sGuideNomenk.nomNom, sGuideNomenk.cod, sGuideNomenk.nomName, sGuideNomenk.Size, sDMC.quant, sGuideNomenk.cost, sGuideNomenk.perList, sGuideNomenk.ed_Izmer2, sDMC.numDoc " & _
  "HAVING (((sDMC.numDoc)=" & gNzak & "));"

  Set tbNomenk = myOpenRecordSet("##435", sql, dbOpenDynaset)
  If tbNomenk Is Nothing Then Exit Sub
  While Not tbNomenk.EOF '!!!
    Grid.AddItem Chr(9) & tbNomenk!nomnom & Chr(9) & tbNomenk!cod & " " & _
    tbNomenk!nomName & " " & tbNomenk!Size & _
    Chr(9) & Chr(9) & Round(tbNomenk!quant / tbNomenk!perList, 2) & Chr(9) & _
    tbNomenk!ed_Izmer2 & Chr(9) & tbNomenk!cost & Chr(9) & _
    Round(tbNomenk!quant * tbNomenk!cost / tbNomenk!perList, 2)
'    Grid.col = 2: Grid.CellFontBold = True
'    Grid.AddItem ""

'NXT2:
    tbNomenk.MoveNext
  Wend
  tbNomenk.Close
End If

End Sub


Sub nomenkToNNQQ(pQuant As Single, eQuant As Single, prQuant As Single)
Dim j As Integer, leng As Integer

leng = UBound(NN)

    For j = 1 To leng
        If NN(j) = tbNomenk!nomnom Then
            QQ(j) = QQ(j) + pQuant * tbNomenk!quantity
            If eQuant > 0 Then _
                QQ2(j) = QQ2(j) + eQuant * tbNomenk!quantity
            If prQuant > 0 Then _
                QQ3(j) = QQ3(j) + prQuant * tbNomenk!quantity
            Exit Sub
        End If
    Next j
    leng = leng + 1
    ReDim Preserve NN(leng): NN(leng) = tbNomenk!nomnom
    ReDim Preserve Cena(leng): Cena(leng) = tbNomenk!cost
    ReDim Preserve QQ(leng): QQ(leng) = pQuant * tbNomenk!quantity
    ReDim Preserve QQ2(leng): QQ2(leng) = eQuant * tbNomenk!quantity
'    QQ2(leng) = 0: If eQuant > 0 Then QQ2(leng) = eQuant * tbNomenk!quantity
    ReDim Preserve QQ3(leng): QQ3(leng) = prQuant * tbNomenk!quantity
    

End Sub
'сумма( по себест-ти) уже отгруженной номенклатуры(незакрытые заказы)
Function otgruzNomenk() As Single
Dim i As Integer
otgruzNomenk = 0

ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): QQ2(0) = 0: ReDim QQ3(0)

'ном-ра входящих в заказы изделий
sql = "SELECT r.* " & _
"FROM xPredmetyByIzdeliaOut r INNER JOIN Orders ON r.numOrder = Orders.numOrder " & _
"WHERE (((Orders.StatusId)<6));"

Set tbProduct = myOpenRecordSet("##384", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Function

While Not tbProduct.EOF
    gNzak = tbProduct!numOrder
    gProductId = tbProduct!prId
    prExt = tbProduct!prExt
    productNomenkToNNQQ 0, tbProduct!quant '2
    tbProduct.MoveNext
Wend
tbProduct.Close

'отдельные ном-ры заказов
sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.cost, sGuideNomenk.perList, " & _
"xPredmetyByNomenkOut.quant as quantity FROM (xPredmetyByNomenkOut INNER JOIN sGuideNomenk ON xPredmetyByNomenkOut.nomNom = sGuideNomenk.nomNom) INNER JOIN Orders ON xPredmetyByNomenkOut.numOrder = Orders.numOrder " & _
"WHERE (((Orders.StatusId)<6));"
Set tbNomenk = myOpenRecordSet("##385", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
While Not tbNomenk.EOF
  Dim str As String: str = tbNomenk!nomnom
  nomenkToNNQQ 0, 0, tbNomenk!cost / tbNomenk!perList '!!!
  tbNomenk.MoveNext
Wend
tbNomenk.Close

otgruzNomenk = 0
For i = 1 To UBound(NN)
    otgruzNomenk = otgruzNomenk + QQ3(i)
Next i

End Function

'в QQ3 накапливается суммарная себест-ть ном-ры в пересчете на цел.ед.изм!!!
'перед исп-ем надо ReDim NN(0): ReDim QQ(0): ReDim QQ2(0) : ReDim QQ3(0):QQ2(0)=0 - не б.этапа
Function productNomenkToNNQQ(pQuant As Single, eQuant As Single) As Boolean
Dim i As Integer, gr() As String

productNomenkToNNQQ = False
'ReDim NN(0): ReDim QQ(0)

'вариантная ном-ра изделия
sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup, " & _
"sGuideNomenk.cost, sGuideNomenk.perList " & _
"FROM sGuideNomenk INNER JOIN (sProducts INNER JOIN xVariantNomenc ON (sProducts.nomNom = xVariantNomenc.nomNom) AND (sProducts.ProductId = xVariantNomenc.prId)) ON sGuideNomenk.nomNom = sProducts.nomNom " & _
"WHERE (((xVariantNomenc.numOrder)=" & gNzak & ") AND (" & _
"(xVariantNomenc.prId)=" & gProductId & ") AND ((xVariantNomenc.prExt)=" & prExt & "));"
'MsgBox sql
Set tbNomenk = myOpenRecordSet("##192", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
ReDim gr(0): i = 0
While Not tbNomenk.EOF
    nomenkToNNQQ pQuant, eQuant, eQuant * tbNomenk!cost / tbNomenk!perList '!!!
    i = i + 1
    ReDim Preserve gr(i): gr(i) = tbNomenk!xgroup
    tbNomenk.MoveNext
Wend
tbNomenk.Close
    
'НЕвариантная ном-ра изделия
'sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup " & _
"From sProducts WHERE (((sProducts.ProductId)=" & gProductId & "));"
sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup, " & _
"sGuideNomenk.cost, sGuideNomenk.perList " & _
"FROM sGuideNomenk INNER JOIN sProducts ON sGuideNomenk.nomNom = sProducts.nomNom " & _
"WHERE (((sProducts.ProductId)=" & gProductId & "));"
'MsgBox sql
Set tbNomenk = myOpenRecordSet("##177", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
While Not tbNomenk.EOF
    For i = 1 To UBound(gr) ' если группа состоит из одной ном-ры, то она
        If gr(i) = tbNomenk!xgroup Then GoTo NXT ' НЕвариантна, т.к. не
    Next i                                      ' не попала в xVariantNomenc
    nomenkToNNQQ pQuant, eQuant, eQuant * tbNomenk!cost / tbNomenk!perList '!!!
NXT: tbNomenk.MoveNext
Wend
tbNomenk.Close

productNomenkToNNQQ = True
End Function
Sub relizDetailMat()
Dim r As Single ', typ As String, prevTyp As String

sql = "SELECT sDocs.numDoc, sDocs.xDate, sGuideSource.sourceName, " & _
"sGuideSource_1.sourceName AS destName, sDocs.Note, " & _
"Sum([sDMC].[quant]*[sGuideNomenk].[cost]/[sGuideNomenk].[perList]) AS cSum " & _
"FROM sGuideNomenk INNER JOIN (sGuideSource AS sGuideSource_1 INNER JOIN ((sGuideSource INNER JOIN sDocs ON sGuideSource.sourceId = sDocs.sourId) INNER JOIN sDMC ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc)) ON sGuideSource_1.sourceId = sDocs.destId) ON sGuideNomenk.nomNom = sDMC.nomNom " & _
"WHERE (" & Pribil.mDateWhere & ") " & _
"GROUP BY sDocs.numDoc, sDocs.xDate, sGuideSource.sourceName, " & _
"sGuideSource_1.sourceName, sDocs.Note ORDER BY sDocs.numDoc;"
'MsgBox sql
Set tbProduct = myOpenRecordSet("##434", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Sub
Grid.FormatString = "|<Накладная №|<Дата|<Откуда|<Куда|<Примечание|<Материалы"
Grid.ColWidth(0) = 0
Grid.ColWidth(rrNumOrder) = 930
Grid.ColWidth(rrDate) = 765
Grid.ColWidth(rrFirm) = 1300
Grid.ColWidth(rrProduct) = 1300
Grid.ColWidth(rrMater) = 1035
Grid.ColWidth(rrReliz) = 1035
quantity = 0
While Not tbProduct.EOF
    quantity = quantity + 1
    Grid.TextMatrix(quantity, 0) = "m"
    Grid.TextMatrix(quantity, rrNumOrder) = tbProduct!numDoc
    Grid.TextMatrix(quantity, rrDate) = Format(tbProduct!xDate, "dd/mm/yy hh:nn:ss")
    Grid.TextMatrix(quantity, rrFirm) = tbProduct!SourceName
    Grid.TextMatrix(quantity, rrProduct) = tbProduct!destName
    Grid.TextMatrix(quantity, rrMater) = tbProduct!note
    Grid.TextMatrix(quantity, rrReliz) = Format(tbProduct!cSum, "0.00") ' сумма цен вход.номенклатур в пересчете на целые
    Grid.AddItem ""
    tbProduct.MoveNext
Wend
tbProduct.Close

End Sub


Sub relizDetailBay(Optional statistic As String = "")
Dim bSum As Single, cSum As Single, prevName As String, prevNom As Long

strWhere = Pribil.bDateWhere
If strWhere <> "" Then strWhere = "where " & strWhere & " "

sql = "SELECT o.numOrder, d.xDate, f.Name " _
& ", Sum(n.cost * i.quant / n.perList) AS cSum " _
& ", Sum(r.intQuant * i.quant / n.perList) AS bSum  " _
& "FROM BayOrders o  " _
& "JOIN sDocs d ON d.numDoc = o.numOrder  " _
& "join sDMC i on d.numExt = i.numExt AND d.numDoc = i.numDoc " _
& "join sDMCrez r on i.nomNom = r.nomNom and o.numOrder = r.numDoc " _
& "JOIN sGuideNomenk n ON n.nomNom = i.nomNom and r.nomNom = i.nomNom  " _
& "JOIN BayGuideFirms f on f.FirmId = o.FirmId " _
& strWhere _
& "GROUP BY o.numOrder, d.xDate, f.Name "

If statistic = "" Then
    sql = sql & " ORDER BY o.numOrder, d.xDate;"
Else
    sql = sql & " ORDER BY f.Name, o.numOrder;"
End If

'Debug.Print sql
Set tbProduct = myOpenRecordSet("##433", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Sub
If statistic = "" Then
    Grid.FormatString = "|<Заказ|<Дата|<Фирма||>Материалы|>Реализация"
    Grid.ColWidth(rrDate) = 765
Else
    Grid.FormatString = "|<Закаpов||<Фирма||>Материалы|>Реализация"
    Grid.ColWidth(rrDate) = 0
End If
Grid.ColWidth(0) = 0
Grid.ColWidth(rrNumOrder) = 885
Grid.ColWidth(rrFirm) = 3855
Grid.ColWidth(rrProduct) = 0
Grid.ColWidth(rrReliz) = 1005
Grid.ColWidth(rrMater) = 1005

quantity = 0: prevName = "$$$$#####@@@@"
While Not tbProduct.EOF
  gNzak = tbProduct!numOrder
  If statistic = "" Or tbProduct!Name <> prevName Then
    quantity = quantity + 1
    bSum = tbProduct!bSum
    cSum = tbProduct!cSum
    Grid.TextMatrix(quantity, rrReliz) = Format(bSum, "0.00")
    Grid.TextMatrix(quantity, rrMater) = Format(cSum, "0.00") ' сумма цен вход.номенклатур в пересчете на целые
    If statistic = "" Then
        Grid.TextMatrix(quantity, rrNumOrder) = gNzak
    Else
        Grid.TextMatrix(quantity, rrNumOrder) = 1
    End If
    Grid.TextMatrix(quantity, rrDate) = Format(tbProduct!xDate, "dd/mm/yy hh:nn:ss")
    Grid.TextMatrix(quantity, rrFirm) = tbProduct!Name
    Grid.TextMatrix(quantity, 0) = "b"
'    Grid.TextMatrix(quantity, rrReliz) = Round(r, 2)
'    Grid.TextMatrix(quantity, rrMater) = Round(tbProduct!cSum, 2) ' сумма цен вход.номенклатур в пересчете на целые
    Grid.AddItem ""
  Else ' только для статистики
    If prevNom <> gNzak Then _
      Grid.TextMatrix(quantity, rrNumOrder) = 1 + Grid.TextMatrix(quantity, rrNumOrder)
    bSum = bSum + tbProduct!bSum
    Grid.TextMatrix(quantity, rrReliz) = Format(bSum, "0.00")
    cSum = cSum + tbProduct!cSum
    Grid.TextMatrix(quantity, rrMater) = Format(cSum, "0.00") ' сумма цен вход.номенклатур в пересчете на целые
  End If
  prevName = tbProduct!Name
  prevNom = gNzak
  tbProduct.MoveNext
Wend
tbProduct.Close

End Sub

Sub uslugDetail(Optional statistic As String = "")
'Dim prevDate As Date, prevNom As Long, prevReliz As Single, prevMater As Single
Dim prevName As String, cSum As Single, prevNom As Long

'strWhere = Pribil.bDateWhere
'If strWhere <> "" Then strWhere = "HAVING ((" & strWhere & ")) "
If statistic = "" Then
    strWhere = " ORDER BY u.numOrder, u.outDate;"
Else
    strWhere = " ORDER BY GuideFirms.Name, u.numOrder;"
End If

sql = "SELECT u.numOrder, u.outDate, " & _
"u.quant, 1 AS cenaEd, GuideFirms.Name, Orders.Product " & _
"FROM GuideFirms INNER JOIN (Orders INNER JOIN xUslugOut u ON Orders.numOrder = u.numOrder) ON GuideFirms.FirmId = Orders.FirmId " & _
Pribil.uDateWhere & strWhere
'" ORDER BY xUslugOut.numOrder, xUslugOut.outDate;"

'MsgBox sql
Set tbProduct = myOpenRecordSet("##383", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Sub
If statistic = "" Then
    Grid.FormatString = "|Заказ|<Дата|<Фирма|<Изделия||>Реализация"
    Grid.ColWidth(rrDate) = 765
    Grid.ColWidth(rrProduct) = 2400
Else
    Grid.FormatString = "|Заказов||<Фирма|||>Реализация"
    Grid.ColWidth(rrDate) = 0
    Grid.ColWidth(rrProduct) = 0
End If
Grid.ColWidth(0) = 0
Grid.ColWidth(rrNumOrder) = 885
Grid.ColWidth(rrFirm) = 3855
Grid.ColWidth(rrReliz) = 1005
Grid.ColWidth(rrMater) = 0 '1005

'prevDate = 0: prevNom = 0: quantity = 0: prevReliz = 0: prevMater = 0
quantity = 0: prevName = "$$$$#####@@@@"
While Not tbProduct.EOF
  gNzak = tbProduct!numOrder
  If statistic = "" Or tbProduct!Name <> prevName Then
  'If 1 = 1 Then
    quantity = quantity + 1
    If statistic = "" Then
        Grid.TextMatrix(quantity, rrNumOrder) = gNzak
    Else
        Grid.TextMatrix(quantity, rrNumOrder) = 1
    End If
    Grid.TextMatrix(quantity, rrDate) = Format(tbProduct!outDate, "dd/mm/yy hh:nn:ss")
    Grid.TextMatrix(quantity, rrFirm) = tbProduct!Name
    Grid.TextMatrix(quantity, rrProduct) = tbProduct!Product
    cSum = tbProduct!cenaEd * tbProduct!quant
    Grid.TextMatrix(quantity, rrReliz) = Format(cSum, "0.00")
    Grid.AddItem ""
  Else ' только для статистики
    If prevNom <> gNzak Then _
      Grid.TextMatrix(quantity, rrNumOrder) = 1 + Grid.TextMatrix(quantity, rrNumOrder)
    cSum = cSum + tbProduct!cenaEd * tbProduct!quant
    Grid.TextMatrix(quantity, rrReliz) = Format(cSum, "0.00") ' сумма цен вход.номенклатур в пересчете на целые
  End If
  prevName = tbProduct!Name
  prevNom = gNzak
  tbProduct.MoveNext
Wend
tbProduct.Close

End Sub

Sub relizDetail(Optional statistic As String = "")
Dim prevDate As Date, prevNom As Long, prevReliz As Single, prevMater As Single
Dim m As Single, r As Single, typ As String, prevTyp As String, prevName As String


If statistic = "" Then
'    strWhere = " ORDER BY r.numOrder, r.outDate;"
    strWhere = " ORDER BY 1, 2;"
Else
'    strWhere = " ORDER BY GuideFirms.Name, r.numOrder;"
    strWhere = " ORDER BY 8, 1;"
End If
sql = "SELECT r.numOrder, r.outDate, " & _
"r.prId, r.prExt, -1 AS costI, " & _
"r.quant, p.cenaEd, f.Name, o.Product " & _
"FROM (GuideFirms f INNER JOIN Orders o ON f.FirmId = o.FirmId) INNER JOIN (xPredmetyByIzdelia p INNER JOIN xPredmetyByIzdeliaOut r ON (p.prExt = r.prExt) AND (p.prId = r.prId) AND (p.numOrder = r.numOrder)) ON o.numOrder = p.numOrder " & _
Pribil.pDateWhere & _
" UNION ALL SELECT pno.numOrder, pno.outDate, " & _
"-1 AS prId, -1 AS prExt, n.cost/n.perList as costI, " & _
"pno.quant, pn.cenaEd, f.Name, o.Product " & _
"FROM sGuideNomenk n INNER JOIN ((GuideFirms f INNER JOIN Orders o ON f.FirmId = o.FirmId) " & _
"INNER JOIN (xPredmetyByNomenk pn INNER JOIN xPredmetyByNomenkOut pno ON " & _
"(pn.nomNom = pno.nomNom) AND (pn.numOrder = pno.numOrder)) ON o.numOrder = pn.numOrder) ON n.nomNom = pn.nomNom " & _
" where " & Pribil.nDateWhere & strWhere

'Debug.Print sql
Set tbProduct = myOpenRecordSet("##381", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Sub
Grid.FormatString = "|Заказ|<Дата|<Фирма|<Изделия|>Материалы|>Реализация"
If statistic = "" Then
    Grid.FormatString = "|Заказ|<Дата|<Фирма|<Изделия|>Материалы|>Реализация"
    Grid.ColWidth(0) = 300
    Grid.ColWidth(rrDate) = 765
    Grid.ColWidth(rrProduct) = 2400
Else
    Grid.FormatString = "|Заказов||<Фирма||>Материалы|>Реализация"
    Grid.ColWidth(0) = 0
    Grid.ColWidth(rrDate) = 0
    Grid.ColWidth(rrProduct) = 0
End If
Grid.ColWidth(rrNumOrder) = 885
Grid.ColWidth(rrFirm) = 3855
Grid.ColWidth(rrReliz) = 1005
Grid.ColWidth(rrMater) = 1005

prevDate = 0: prevNom = 0: quantity = 0: prevReliz = 0: prevMater = 0
While Not tbProduct.EOF
    
  gNzak = tbProduct!numOrder
  If tbProduct!costI = -1 Then ' готовое изделие
        gProductId = tbProduct!prId
        prExt = tbProduct!prExt
        m = Pribil.getProductNomenkSum * tbProduct!quant
        typ = "p"
        GoTo AA
'  ElseIf tbProduct!costI = -2 Then ' услуга
'        m = 0: typ = "u"
'        GoTo AA
  Else ' отд.ном-ра
        typ = "n"
        m = tbProduct!costI * tbProduct!quant
AA:     r = tbProduct!cenaEd * tbProduct!quant
  End If
'If gNzak = "3102201" Then
'    gNzak = gNzak
'End If
  If statistic = "" Then
      bilo = (prevNom <> gNzak Or prevDate <> tbProduct!outDate)
  Else
      bilo = (prevName <> tbProduct!Name)
  End If
'  bilo = True
  If bilo Then
'  If prevNom <> gNzak Or prevDate <> tbProduct!outDate Then
    quantity = quantity + 1
    If statistic = "" Then
        Grid.TextMatrix(quantity, rrNumOrder) = gNzak
    Else
        Grid.TextMatrix(quantity, rrNumOrder) = 1 'кол-во заказов
    End If
    Grid.TextMatrix(quantity, rrDate) = Format(tbProduct!outDate, "dd/mm/yy hh:nn:ss")
    Grid.TextMatrix(quantity, rrFirm) = tbProduct!Name
    Grid.TextMatrix(quantity, rrProduct) = tbProduct!Product
    Grid.AddItem ""
    prevReliz = r
    prevMater = m
    prevTyp = typ
  Else ' это строка с тем же заказом и с той же датой - если отгрузка и изделий и отдель.номенклатур
    If statistic <> "" And prevNom <> gNzak Then _
        Grid.TextMatrix(quantity, rrNumOrder) = 1 + Grid.TextMatrix(quantity, rrNumOrder)
    prevReliz = r + prevReliz
    prevMater = m + prevMater
    If typ <> prevTyp Then prevTyp = "w" 'здесь не м.б."u"
  End If
  Grid.TextMatrix(quantity, 0) = prevTyp
  Grid.TextMatrix(quantity, rrReliz) = Format(prevReliz, "0.00")
  Grid.TextMatrix(quantity, rrMater) = Format(prevMater, "0.00")
  prevNom = gNzak: prevDate = tbProduct!outDate
  prevName = tbProduct!Name
  tbProduct.MoveNext
Wend
tbProduct.Close

End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If WindowState = vbMinimized Then Exit Sub
On Error Resume Next

h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w

laSum.Top = laSum.Top + h
laRecSum.Top = laRecSum.Top + h
laHeader.Width = laHeader.Width + w
cmExel.Top = cmExel.Top + h
cmPrint.Top = cmPrint.Top + h
cmExit.Top = cmExit.Top + h
cmExit.Left = cmExit.Left + w
laRecCount.Top = laRecCount.Top + h

End Sub

Private Sub Grid_Click()
If Regim <> "" Then Exit Sub
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
'If mousRow = 0 And (Regim = "KK" Or Regim = "RA") Then
'Grid.CellBackColor = Grid.BackColor
    If mousCol = 0 Then Exit Sub
'    If mousCol > 3 Then
'        SortCol Grid, mousCol, "numeric"
'    Else
'        SortCol Grid, mousCol
'    End If
'    Grid.row = 1    ' только чтобы снять выделение
'End If

End Sub

Private Sub Grid_DblClick()
Dim str As String

If Grid.CellBackColor <> &H88FF88 Then Exit Sub
'If Grid.CellBackColor <> &H88FF88 Or Regim <> "" Then Exit Sub

gNzak = Grid.TextMatrix(mousRow, rrNumOrder)
If Grid.TextMatrix(mousRow, 0) = "u" Then
    MsgBox "Заказ №" & gNzak & " не содержит предметов, поэтому далее он не " & _
    "детализируется!", , "Предупреждение"
    Exit Sub
End If
    
Dim Report2 As New Report

If Regim = "ventureZatrat" Then
    Report2.Regim = "ventureZatratDetail"
ElseIf Regim = "mat" Then

    Report2.Regim = "subDetailMat"
    str = Grid.TextMatrix(mousRow, rrReliz)
    If MsgBox("Вы хотите посмотреть записи, которые образуют сумму " & str _
    , vbDefaultButton2 Or vbYesNo, "Продолжить?") = vbNo Then Exit Sub
Else
    Report2.Regim = "subDetail"
    str = Grid.TextMatrix(mousRow, rrMater) & " и " & Grid.TextMatrix(mousRow, rrReliz)
    If MsgBox("Вы хотите посмотреть записи, которые образуют суммы " & str _
    , vbDefaultButton2 Or vbYesNo, "Продолжить?") = vbNo Then Exit Sub
End If

Report2.param1 = Grid.TextMatrix(mousRow, 0) '
Report2.param2 = Grid.TextMatrix(mousRow, rrDate)
Report2.param3 = str
Report2.Show vbModal
End Sub

Private Sub Grid_EnterCell()
If quantity = 0 Or Not (Regim = "" Or Regim = "bay" Or Regim = "mat" Or Regim = "venture" Or Regim = "ventureZatrat") Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col
If (mousCol = rrReliz Or (mousCol = rrMater And Regim <> "mat") Or (Regim = "ventureZatrat" And Grid.col >= rzMainCosts)) _
Then
   Grid.CellBackColor = &H88FF88
Else
   Grid.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor

End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
Else
'ElseIf Grid.col = rrReliz Or Grid.col = rrMater Then
    laSumControl "col"
    laRecSum.Caption = Round(sumInGridCol(Grid, Grid.col), 2)
End If
End Sub

