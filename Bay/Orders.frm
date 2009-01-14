VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Orders 
   Appearance      =   0  'Flat
   Caption         =   "Продажа"
   ClientHeight    =   6132
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   11880
   Icon            =   "Orders.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6132
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lbVenture 
      Appearance      =   0  'Flat
      Height          =   600
      Left            =   5500
      TabIndex        =   26
      Top             =   1000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   0
      TabIndex        =   16
      Top             =   -80
      Width           =   11835
      Begin VB.TextBox tbEndDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   21
         Top             =   180
         Width           =   795
      End
      Begin VB.ComboBox cbM 
         Height          =   315
         ItemData        =   "Orders.frx":030A
         Left            =   11160
         List            =   "Orders.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   615
      End
      Begin VB.CheckBox cbClose 
         Caption         =   "  "
         Height          =   195
         Left            =   5040
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox cbEndDate 
         Caption         =   " "
         Height          =   315
         Left            =   2460
         TabIndex        =   19
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox tbStartDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1260
         TabIndex        =   18
         Text            =   "01.09.02"
         Top             =   180
         Width           =   795
      End
      Begin VB.CheckBox cbStartDate 
         Caption         =   " "
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Top             =   180
         Width           =   315
      End
      Begin VB.Label laFiltr 
         Caption         =   "Включен фильтр !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7260
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label laClos 
         Caption         =   ",  в т. ч. закрытые"
         Height          =   195
         Left            =   3600
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label laPo 
         Caption         =   "пос"
         Height          =   195
         Left            =   2160
         TabIndex        =   23
         Top             =   240
         Width           =   195
      End
      Begin VB.Label laPeriod 
         Caption         =   "Период с  "
         Height          =   195
         Left            =   60
         TabIndex        =   22
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.ListBox lbProblem 
      Height          =   240
      Left            =   2580
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lbAnnul 
      Height          =   816
      ItemData        =   "Orders.frx":030E
      Left            =   240
      List            =   "Orders.frx":031E
      TabIndex        =   13
      Top             =   1980
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   10560
      Top             =   5340
   End
   Begin VB.TextBox tbEnable 
      BackColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   11460
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5460
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox tbInform 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   9015
   End
   Begin VB.ListBox lbClose 
      Height          =   1008
      ItemData        =   "Orders.frx":0348
      Left            =   240
      List            =   "Orders.frx":035B
      TabIndex        =   11
      Top             =   3180
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lbDel 
      Height          =   432
      ItemData        =   "Orders.frx":038B
      Left            =   240
      List            =   "Orders.frx":0395
      TabIndex        =   10
      Top             =   4380
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmExvel 
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   8940
      TabIndex        =   5
      Top             =   5580
      Width           =   1515
   End
   Begin VB.ListBox lbM 
      Height          =   240
      Left            =   1560
      TabIndex        =   9
      Top             =   1020
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   240
      MaxLength       =   10
      TabIndex        =   8
      Text            =   "tbMobile"
      Top             =   1620
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4455
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   11835
      _ExtentX        =   20870
      _ExtentY        =   7853
      _Version        =   393216
      BackColor       =   16777215
      ForeColorFixed  =   0
      BackColorSel    =   65535
      ForeColorSel    =   -2147483630
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   $"Orders.frx":03AF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   5580
      Width           =   1275
   End
   Begin VB.CommandButton cmRefr 
      Caption         =   "Загрузить"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   5580
      Width           =   975
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   396
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   699
      ButtonWidth     =   635
      ButtonHeight    =   572
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Менеджер:"
      Height          =   195
      Left            =   10320
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label laInform 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1260
      TabIndex        =   6
      Top             =   5580
      Width           =   1575
   End
   Begin VB.Menu mnMenu 
      Caption         =   "Меню"
      Begin VB.Menu mnGuideFirms 
         Caption         =   "Справочник сторонних организаций F11"
      End
      Begin VB.Menu mnFirmFind 
         Caption         =   "Поиск фирмы по названию               F12"
      End
      Begin VB.Menu mnSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Выход из программы                Alt F4"
      End
   End
   Begin VB.Menu mnMeassure 
      Caption         =   "Настройка"
   End
   Begin VB.Menu mnSklad 
      Caption         =   "Склад"
      Begin VB.Menu mnNomenk 
         Caption         =   "Остатки по ном-ре    F4"
      End
      Begin VB.Menu mnRecalc 
         Caption         =   "Пересчет статистики"
      End
      Begin VB.Menu mnSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnReports 
         Caption         =   "Посещения фирм по месяцам"
      End
      Begin VB.Menu mnAnalityc 
         Caption         =   "Аналитика по продажам"
      End
   End
   Begin VB.Menu mnContext 
      Caption         =   "aa"
      Visible         =   0   'False
      Begin VB.Menu mnFirmsGuide 
         Caption         =   "Вход в справочник организаций"
      End
      Begin VB.Menu mnNoArhivFiltr 
         Caption         =   "Фильтр ""Заказы в обработке"""
      End
      Begin VB.Menu mnSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnNoCloseFiltr 
         Caption         =   "Фильтр ""Незакрытые заказы"""
         Visible         =   0   'False
      End
      Begin VB.Menu mnNoClose 
         Caption         =   "Отчет ""Незакрытые заказы"""
         Visible         =   0   'False
      End
      Begin VB.Menu mnAllOrders 
         Caption         =   "Отчет ""Все заказы Фирмы"""
         Visible         =   0   'False
      End
      Begin VB.Menu mnBillFirma 
         Caption         =   "Плательщик: "
         Visible         =   0   'False
      End
      Begin VB.Menu mnQuickBill 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mousRow As Long
Public mousCol As Long
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim beClick As Boolean
Dim flDelRowInMobile As Boolean
Dim minut As Integer
Public g_id_bill As String


Const AddCaption = "Добавить"
Const t17_00 = 61200 ' в секундах
'"BayGuideFirms.Name, BayOrders.outDateTime, BayOrders.ordered, BayOrders.paid, " & _
"BayOrders.shipped, BayOrders.lastManagId, BayOrders.Invoice, " & $$6
Const rowFromOrdersSQL = "SELECT BayOrders.numOrder, BayOrders.inDate, " & _
"BayOrders.ManagId, BayOrders.StatusId, BayOrders.ProblemId, " & _
"BayGuideFirms.Name, BayOrders.outDateTime, BayOrders.paid, " & _
"BayOrders.lastManagId, BayOrders.Invoice, " & _
"BayOrders.Ves, BayOrders.Size, BayOrders.Places " & _
",guideventure.venturename as venture " & _
",id_bill, rate " & _
",BayGuideFirms.id_voc_names as id_voc_names" & _
",guideventure.sysname as servername" & _
" FROM BayOrders INNER JOIN BayGuideFirms ON BayGuideFirms.FirmId = BayOrders.FirmId " & _
" left join guideventure on guideventure.ventureId = bayOrders.ventureid "

Private Sub cbClose_Click()
cmRefr.Caption = "Загрузить"
End Sub

Private Sub cbEndDate_Click()
cmRefr.Caption = "Загрузить"
tbEndDate.Enabled = Not tbEndDate.Enabled
End Sub

Private Sub cbM_Click()
If zakazNum = 0 Then
    On Error Resume Next ' т.к. устанавливаем cbM из Load
    cmRefr.SetFocus
Else
'If cbM.ListIndex > -1 Then cmAdd.Enabled = True
    lbHide
End If
cbM.TabStop = False
End Sub

Private Sub cbM_LostFocus()
If cbM.ListIndex < 0 Then
    MsgBox "Заполните поле 'Менеджер'", , "Предупреждение"
    cbM.SetFocus
End If

End Sub

Private Sub cbStartDate_Click()
cmRefr.Caption = "Загрузить"
tbStartDate.Enabled = Not tbStartDate.Enabled
End Sub
    
Sub begFiltrDisable()
    laPeriod.Enabled = False
    laPo.Enabled = False
    laClos.Enabled = False
'    laPerson.Enabled = False
    cbStartDate.Enabled = False
    tbStartDate.Enabled = False
    cbEndDate.Enabled = False
    tbEndDate.Enabled = False
    cbClose.Enabled = False

End Sub

Sub begFiltrEnable()
    laPeriod.Enabled = True
    laPo.Enabled = True
    laClos.Enabled = True
    cbStartDate.Enabled = True
    If cbStartDate.value = 1 Then tbStartDate.Enabled = True
    cbEndDate.Enabled = True
    If cbEndDate.value = 1 Then tbEndDate.Enabled = True
    cbClose.Enabled = True

End Sub


Private Sub cmAdd_Click()
Dim str As String, intNum As Integer, l As Long
Dim DateFromNum As String
  
wrkDefault.BeginTrans 'lock01
sql = "update system set resursLock = resursLock" 'lock02
myBase.Execute (sql) 'lock03

str = getSystemField("lastPrivatNum")
DateFromNum = left$(str, 5)
intNum = right$(str, Len(str) - 5)

intNum = intNum + 1
If intNum < 100 Then
    str = Format(intNum, "00")
Else
    str = Format(intNum, "000")
End If
l = DateFromNum & str
'tbSystem!lastPrivatNum = DateFromNum & str
myBase.Execute ("update system set lastPrivatNum = " & DateFromNum & str)
'tbSystem.Update
BB:
wrkDefault.CommitTrans
'tbSystem.Close
Dim isBaseOrder As Boolean
Dim baseFirmId As Integer, baseFirm As String
Dim baseProblemId As Integer, baseProblem As String, begPubNum As Long

gNzak = Grid.TextMatrix(Orders.mousRow, orNomZak)
If InStr(Orders.cmAdd.Caption, "+") > 0 Then
  sql = "SELECT BayOrders.CehId, BayOrders.ProblemId, BayOrders.FirmId, " & _
        "GuideCeh.Ceh, GuideProblem.Problem, BayGuideFirms.Name " & _
        "FROM GuideProblem INNER JOIN (BayGuideFirms INNER JOIN " & _
        "(GuideCeh INNER JOIN BayOrders ON GuideCeh.CehId = BayOrders.CehId) " & _
        "ON BayGuideFirms.FirmId = BayOrders.FirmId) ON GuideProblem.ProblemId " & _
        "= BayOrders.ProblemId WHERE (((BayOrders.numOrder)=" & gNzak & "));"
'  On Error GoTo NXT1
  Set tbOrders = myBase.OpenRecordset(sql, dbOpenForwardOnly)
  
  baseFirmId = tbOrders!firmId
  baseProblemId = tbOrders!problemId
  
  baseFirm = tbOrders!Name
  baseProblem = tbOrders!problem
  isBaseOrder = True
  tbOrders.Close
Else
  isBaseOrder = False
End If
NXT1:
cmAdd.Caption = AddCaption

sql = "select * from BayOrders where numOrder = " & l
'MsgBox sql
Set tbOrders = myOpenRecordSet("##07", sql, dbOpenForwardOnly)
'If tbOrders Is Nothing Then Exit Sub

'If Not uniqOrderNum(tbOrders, l) Then
If Not tbOrders.BOF Then
    MsgBox "номер " & l & " не уникален (см. заказ от " _
    & tbOrders!inDate & ").  Повторите попытку или обратитесь к Администратору!", , ""
    tbOrders.Close
    Exit Sub
End If

'On Error GoTo ERR1
tbOrders.AddNew
tbOrders!StatusId = 0
tbOrders!numorder = l
tbOrders!inDate = Now
tbOrders!managId = manId(cbM.ListIndex)
'tbOrders!firmId = 0
If isBaseOrder Then
  tbOrders!firmId = baseFirmId
  tbOrders!problemId = baseProblemId
End If
tbOrders.Update

If zakazNum > 0 Then Grid.AddItem ""
zakazNum = zakazNum + 1
Grid.TextMatrix(zakazNum, 0) = zakazNum
Grid.TextMatrix(zakazNum, orInvoice) = "счет ?"
Grid.TextMatrix(zakazNum, orNomZak) = l
Grid.TextMatrix(zakazNum, orData) = Format(Now, "dd.mm.yy")
Grid.TextMatrix(zakazNum, orMen) = Orders.cbM.Text
Grid.TextMatrix(zakazNum, orStatus) = status(0)
If isBaseOrder Then
  Grid.TextMatrix(zakazNum, orProblem) = baseProblem
  Grid.TextMatrix(zakazNum, orFirma) = baseFirm
End If
rowViem Grid.Rows - 1, Grid
tbOrders.Close
Grid.row = zakazNum
Grid.col = orFirma
Grid.LeftCol = orNomZak
Grid.SetFocus


End Sub

Private Sub cmExvel_Click()
GridToExcel Grid
End Sub

Private Sub cmRefr_Click()
Dim minDate As Date, maxDate As Date

begFiltrEnable
If cbStartDate.value = 1 And cbEndDate.value = 1 Then
    minDate = tbStartDate.Text
    maxDate = tbEndDate.Text
    If minDate > maxDate Then
        MsgBox "начало периода должно быть раньше конца", , "ERROR"
        Exit Sub
    End If
End If
beClick = False
Me.MousePointer = flexHourglass
begFiltr
LoadBase
Me.MousePointer = flexDefault

cmRefr.Caption = "Обновить"
laFiltr.Visible = False

End Sub


Sub lbHide(Optional noFocus As String = "")
tbMobile.Visible = False
lbM.Visible = False
lbDel.Visible = False
lbClose.Visible = False
lbAnnul.Visible = False
lbProblem.Visible = False
lbVenture.Visible = False

Grid.Enabled = True
If noFocus = "" Then
    Grid.SetFocus
    Grid_EnterCell
End If
End Sub


Sub loadWithFiltr(Optional nomZak As String = "")
'bilo = True
If IsNumeric(nomZak) Then ' поиск номера из Цеха
    orSqlWhere(0) = "" 'исп-ся только и сложного фильтра
    orSqlWhere(orNomZak) = strWhereByValCol(nomZak, orNomZak)
ElseIf nomZak = "" Then
    orSqlWhere(0) = ""
    orSqlWhere(mousCol) = strWhereByValCol(Grid.Text, CInt(mousCol))
    If orSqlWhere(mousCol) = "" Then Exit Sub ' в этом поле не предусмотрен фильтр
End If
Me.MousePointer = flexHourglass
laFiltr.Visible = True
LoadBase
cmRefr.Caption = "Загрузить"
Me.MousePointer = flexDefault
orSqlWhere(0) = "" 'исп-ся однократно (для сложного фильтра)
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, value As String, i As Integer, il As Long

If cbM.ListIndex < 0 Then
'    cbM_LostFocus
    Exit Sub
End If

If LCase(tbEnable.Text) <> "arh" And LCase(tbEnable.Text) <> "фкр" _
And tbEnable.Visible Then Exit Sub
If KeyCode = vbKeyEscape Then
    cmAdd.Caption = AddCaption
    lbHide
ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    tbEnable.Text = ""
    tbEnable.Visible = True
    tbEnable.SetFocus
ElseIf KeyCode = vbKeyF4 Then
    mnNomenk_Click 'не прописываем hotkey в меню, т.к. cbM_LostFocus
ElseIf KeyCode = vbKeyF5 Then
    cmAdd_Click
ElseIf KeyCode = vbKeyF7 Then
    If mousCol = orNomZak Then
        value = ""
AA:     value = InputBox("Введите номер заказа", "Поиск", value)
        If value = "" Then Exit Sub
        If Not IsNumeric(value) Then
            MsgBox "Номер должен быть числом"
            GoTo AA
        End If
        If findValInCol(Grid, value, orNomZak) Then Exit Sub
        If MsgBox("Выполнить поиск заказа по всей базе?", vbYesNo, _
        "Среди загруженных заказ не найден!") = vbNo Then Exit Sub
        For i = 1 To orColNumber
            orSqlWhere(i) = ""
        Next i
        loadWithFiltr value
        Grid_EnterCell 'поскольку одна строчка
    ElseIf mousCol = orFirma Then
'        BayGuideFirms.Regim = "F7" ' не показать кнопку Выбрать и Добавить
'        GuideFirmOnOff
        value = Grid.TextMatrix(mousRow, orFirma)
        value = InputBox("Укажите полное название или фрагмент.", "Поиск в колонке 'Название Фирмы'", value)
        If value = "" Then Exit Sub
        If findExValInCol(Grid, value, orFirma) > 0 Then Exit Sub
        If MsgBox("Выполнить расширенный поиск фирмы '" & value & "' ?", vbYesNo, _
        "Среди загруженных заказ этой фирмы не найден!") = vbNo Then Exit Sub
        If tbEnable.Visible Then
            FindFirm.cmAllOrders.Visible = True
            FindFirm.cmNoClose.Visible = True
            FindFirm.cmNoCloseFiltr.Visible = True
        End If
        FindFirm.tb.Text = value
        FindFirm.Show vbModal
    Else
        value = Grid.TextMatrix(mousRow, mousCol)
        value = InputBox("Укажите образец поиска.", "Поиск", value)
        If findExValInCol(Grid, value, CInt(mousCol)) > 0 Then Exit Sub
        MsgBox "Фрагмент не найден"
'        MsgBox "По этому полю поиск не предусмотрен", , "Предупреждение"
    End If
ElseIf KeyCode = vbKeyF11 Then
    mnGuideFirms_Click 'не прописываем hotkey в меню, т.к. cbM_LostFocus
ElseIf KeyCode = vbKeyF12 Then
    mnFirmFind_Click
ElseIf KeyCode = vbKeyMenu Then
    If cmAdd.Enabled And beClick And cmAdd.Caption = AddCaption Then _
                    cmAdd.Caption = AddCaption & " +"
End If
End Sub
'curCol As Integer, colName As String, colWdth As Integer
'Sub initOrCol(colNum As Integer, Optional field As String = "")
Sub initOrCol(curCol As Integer, colName As String, colWdth As Integer, _
Optional field As String = "")

If orColNumber = 0 Then
    Grid.Cols = 2
    Grid.colWidth(0) = 0
Else
    Grid.Cols = Grid.Cols + 1
End If
orColNumber = orColNumber + 1

'If orColNumber > 1 Then Grid.Cols = Grid.Cols + 1
curCol = orColNumber

ReDim Preserve orSqlFields(orColNumber + 1)
orSqlFields(orColNumber) = field

If colWdth >= 0 Then Grid.colWidth(orColNumber) = colWdth
Grid.TextMatrix(0, orColNumber) = colName

End Sub


Private Sub Form_Load()
Dim i As Integer, str As String

oldHeight = Me.Height
oldWidth = Me.Width


If otlad = "otlaD" Then
    Frame1.BackColor = otladColor
    Me.BackColor = otladColor
    mnReports.Visible = True
'    tbEnable.Visible = True
    tbEnable.Text = "arh"
End If

'lb.AddItem table!Problem



beClick = False
flDelRowInMobile = False
Me.Caption = Me.Caption & mainTitle
mousCol = 1
orColNumber = 0
initOrCol orNomZak, "№ заказа", 1050, "nBayOrders.numOrder"
initOrCol orInvoice, "№ счета", 950, "sBayOrders.Invoice"
initOrCol orVenture, "Предпр.", 700, "sOrders.ventureId"
initOrCol orData, "Дата", 810, "dBayOrders.inDate"
initOrCol orMen, "М", 255, "sGuideManag.Manag"
initOrCol orStatus, "Статус", 810, "sGuideStatus.Status"
initOrCol orProblem, "Проблемы", -1, "sGuideProblem.Problem"
initOrCol orFirma, "Название Фирмы", 2625, "sGuideFirms.Name"
initOrCol orDataVid, "Дата выдачи", 990, "dBayOrders.outDateTime"
initOrCol orVrVid, "Вр.выдачи", 645
initOrCol orVes, "Вес", 630
initOrCol orSize, "Размер", 830
initOrCol orPlaces, "Мест", 600
initOrCol orZakazano, "заказано", 660 ', "nBayOrders.ordered" $$6
initOrCol orOplacheno, "согласовано", 645, "nBayOrders.paid"
initOrCol orOtgrugeno, "отгружено", 645 ', "nBayOrders.shipped"$$6
initOrCol orLastMen, "M", 255, "sGuideManag_1.Manag"
initOrCol orBillId, "", 0, "sOrders.id_bill"
initOrCol orVocnameId, "", 0, "nOrders.id_voc_names"
initOrCol orServername, "", 0, "sOrders.servername"

ReDim Preserve orSqlWhere(orColNumber)
zakazNum = 0
'tbStartDate.Text = Format(CurDate - 7, "dd/mm/yy")
tbStartDate.Text = Format(DateAdd("d", -7, CurDate), "dd/mm/yy")
tbEndDate.Text = Format(CurDate, "dd/mm/yy")

'*********************************************************************$$7
sql = "SELECT * From GuideManag ORDER BY forSort;"
Set table = myOpenRecordSet("##03", sql, dbOpenForwardOnly)
If table Is Nothing Then myBase.Close: End
i = 0: ReDim manId(0):
Dim imax As Integer: imax = 0: ReDim Manag(0)
While Not table.EOF
    str = table!Manag
    If str = "not" Then
        GoTo AA
    ElseIf LCase(table!forSort) <> "unused" Then
        If table!managId <> 0 Then cbM.AddItem str
        lbM.AddItem str
        manId(i) = table!managId
        i = i + 1
        ReDim Preserve manId(i):
AA:     If imax < table!managId Then
            imax = table!managId
            ReDim Preserve Manag(imax)
        End If
        Manag(table!managId) = str
    End If
    table.MoveNext
Wend
table.Close

lbM.Height = lbM.Height + 195 * (lbM.ListCount - 1)

If otlad = "otlaD" Then cbM.ListIndex = cbM.ListCount - 1

For i = 0 To UBound(Problems)
    lbProblem.AddItem Problems(i)
Next i
lbProblem.Height = lbProblem.Height + 195 * (lbProblem.ListCount - 1)

'*******
Set table = myOpenRecordSet("##72", "GuideVenture", dbOpenForwardOnly)
If table Is Nothing Then myBase.Close: End

lbVenture.AddItem "", 0
While Not table.EOF
    lbVenture.AddItem "" & table!ventureName & "", table!ventureId
    table.MoveNext
Wend
table.Close

'begFiltr '******* начальный фильтр
isOrders = True
trigger = True

End Sub
 

Sub begFiltr() '******* начальный фильтр
Dim stDate As String, enDate As String, i As Integer
Dim addNullDate As String, strWhere As String
 
 For i = 1 To orColNumber
    orSqlWhere(i) = ""
 Next i
 
 
 If cbStartDate.value = 1 Then
    stDate = "(BayOrders.inDate)>='" & _
             Format(Orders.tbStartDate.Text, "yyyy-mm-dd") & "'"
    addNullDate = ""
 Else
    stDate = ""
    addNullDate = " OR (BayOrders.inDate) Is Null"
 End If

 If cbEndDate.value = 1 Then
    enDate = "(BayOrders.inDate)<='" & _
            Format(Orders.tbEndDate.Text, "yyyy-mm-dd") & " 11:59:59 PM'"
 Else
    enDate = ""
 End If
 If stDate <> "" And enDate <> "" Then
    strWhere = stDate & ") AND( " & enDate
 ElseIf stDate <> "" Or enDate <> "" Then
    strWhere = stDate & enDate
 Else
    addNullDate = ""
    strWhere = ""
 End If
 orSqlWhere(orData) = strWhere & addNullDate
 
 If cbClose.value = 0 Or Not tbEnable.Visible Then
    orSqlWhere(orStatus) = "(BayOrders.StatusId)<>6" 'закрыт
 Else
    orSqlWhere(orStatus) = ""
 End If
 
 setWhereInvoice "check"
 
End Sub
Sub setWhereInvoice(Optional check As String = "")
 If Not tbEnable.Visible Or check = "" Then
'    orSqlWhere(orInvoice) = "isNumeric(BayOrders.Invoice) =true OR (BayOrders.shipped) Is Null" $$6
'    orSqlWhere(orInvoice) = "isNumeric(BayOrders.Invoice) = 1 " '$$6
orSqlWhere(orInvoice) = "(isNumeric(BayOrders.Invoice) =1) OR NOT EXISTS(" & _
"SELECT * FROM bayNomenkOut WHERE BayOrders.numOrder = bayNomenkOut.numOrder) "
    
 End If
End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer
lbHide "noFocus"


If Me.WindowState = vbMinimized Then Exit Sub

On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w
cmRefr.Top = cmRefr.Top + h
laInform.Top = laInform.Top + h
cmAdd.Top = cmAdd.Top + h
'cmExAll.Top = cmExAll.Top + h
cmExvel.Top = cmExvel.Top + h
tbEnable.Top = tbEnable.Top + h
tbEnable.left = tbEnable.left + w
End Sub

Private Sub Form_Unload(Cancel As Integer)

isOrders = False
exitAll
'exitLast

End Sub

Private Sub Grid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Dim date1 As Date, date2 As Date ' в 2 х местах
Dim date1S, date2S As String

date1S = sortGrid.TextMatrix(Row1, mousCol)
date2S = sortGrid.TextMatrix(Row2, mousCol)

'If Not IsDate(date1S) = "" And date2S = "" Then
'    Cmp = 0
'    Exit Sub
If Not IsDate(date1S) Then
    Cmp = -1
    GoTo CC:
ElseIf Not IsDate(date2S) Then
    Cmp = 1
    GoTo CC:
End If

date1 = date1S
date2 = date2S
If date1 > date2 Then
    Cmp = 1
ElseIf date1 < date2 Then
    Cmp = -1
Else
    Cmp = 0
End If
CC:
If trigger Then Cmp = -Cmp

End Sub

Private Sub Grid_Click()
'laInform.Caption = laInform.Caption & "   cRow=" & Grid.row & "  cCol=" & Grid.col
If zakazNum = 0 Then Exit Sub
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow

If mousRow = 0 Then
    Grid.CellBackColor = Grid.BackColor
    If mousCol = 0 Then Exit Sub
    If mousCol = orNomZak Or mousCol = orZakazano Or mousCol = orOplacheno _
    Or mousCol = orOtgrugeno Then
        SortCol Grid, mousCol, "numeric"
    ElseIf mousCol = orData Or mousCol = orDataVid Then
        SortCol Grid, mousCol, "date"
    Else
        SortCol Grid, mousCol
    End If
    Grid.row = 1    ' только чтобы снять выделение
'    Grid_EnterCell
End If
Grid_EnterCell
End Sub
    
Sub GuideFirmOnOff()
Dim tmpRow As Long, tmpCol As Long
'    tmpRow = mousRow
'    tmpCol = mousCol
    GuideFirms.Show vbModal
'    mousRow = tmpRow
'    mousCol = tmpCol
'    BayOrders.Enabled = True
    Orders.SetFocus

End Sub

Function haveUslugi() As Boolean
Dim s As Single

End Function

Function havePredmeti() As Boolean
Dim s As Single
havePredmeti = False
sql = "SELECT quantity From sDMCrez WHERE (((numDoc)=" & gNzak & "));"
If Not byErrSqlGetValues("W##199", sql, s) Then myBase.Close: End
If s > 0 Then havePredmeti = True
End Function

Private Sub Grid_DblClick()
Dim str As String, statId As Integer
Dim billCompany As String
Dim i As Integer


If zakazNum = 0 Then Exit Sub
If mousRow = 0 Then Exit Sub

gNzak = Grid.TextMatrix(mousRow, orNomZak)
sql = "SELECT StatusId From BayOrders WHERE (((numOrder)=" & gNzak & "));"
If Not byErrSqlGetValues("##174", sql, statId) Then Exit Sub

If mousCol = orNomZak Then
  
  If statId = 7 Then
    MsgBox "У заказа с данным статусом не может быть предметов!", , "Предупреждение"
    Exit Sub
  End If
  
'  If Grid.CellForeColor = 200 Or Grid.CellForeColor = vbBlue Then
  tmpStr = ""
  If havePredmeti Then
    str = "посмотреть"
  ElseIf statId >= 6 Then
    Exit Sub
  Else
    str = "сформировать"
  End If
  numDoc = gNzak
  numExt = 0 ' это флаг для некот. п\п, что нужно считать именно доступные остатки
  If MsgBox("Вы хотите " & str & " предметы к заказу? " & tmpStr, _
  vbYesNo Or vbDefaultButton2, "Заказ № " & numDoc) = vbYes Then
        sql = "DELETE From xUslugOut WHERE (((numOrder)=" & gNzak & "));"
        myExecute "##304", sql, 0 'удаляем если есть
        
        If statId < 6 Then
            sProducts.Regim = ""
        Else
            sProducts.Regim = "closeZakaz"
        End If
        sProducts.Show vbModal
  End If
  Exit Sub
End If

If Grid.CellBackColor = vbYellow Then Exit Sub

If mousCol = orVenture Then
     listBoxInGridCell lbVenture, Grid, Grid.TextMatrix(mousRow, orVenture)
ElseIf mousCol = orFirma Then
     
    If Grid.TextMatrix(mousRow, orVenture) <> "" Then
        
        billCompany = "Установить"
    
        If Grid.CellForeColor = vbRed Then
            sql = "select wf_retrieve_bill_company(" + Grid.TextMatrix(mousRow, orBillId) + ", '" + Grid.TextMatrix(mousRow, orVenture) + "')"
'            Debug.Print sql
            If byErrSqlGetValues("W##102.1", sql, billCompany) Then
                mnBillFirma.Tag = Grid.TextMatrix(mousRow, orBillId)
            End If
            If billCompany = "" Then
                billCompany = "Id = [" & Grid.TextMatrix(mousRow, orBillId) & "]"
            End If
        Else
            mnBillFirma.Tag = ""
        End If
        
        mnBillFirma.Visible = True
        mnBillFirma.Caption = "Плательщик: " + billCompany
        
        For i = mnQuickBill.UBound To 1 Step -1
            Unload mnQuickBill(i)
        Next i
        
        If serverIsAccessible(Grid.TextMatrix(mousRow, orVenture)) Then
        
            sql = _
                 " select o.id_bill, max(o.inDate) as lastDate " _
                & " from bayorders o" _
                & " join bayorders z on z.firmid = o.firmid and z.ventureid = o.ventureid and z.numorder = " & gNzak _
                & " where " _
                & "     o.id_bill is not null " _
                & " group by o.id_bill" _
                & " order by lastDate desc"
                  
            
            Set tbOrders = myOpenRecordSet("##102.2", sql, 0)
            If Not tbOrders.BOF Then
    '            Load mnQuickBill(0)
    '            mnQuickBill(0).Caption = "-"
                i = 0
                While Not tbOrders.EOF
                    If CStr(tbOrders!id_bill) <> Grid.TextMatrix(mousRow, orBillId) Then
                        mnQuickBill(0).Visible = True
                        Load mnQuickBill(1 + i)
                        mnQuickBill(i + 1).Tag = tbOrders!id_bill
                        sql = "select wf_retrieve_bill_company(" + CStr(tbOrders!id_bill) + ", '" + Grid.TextMatrix(mousRow, orVenture) + "')"
                        byErrSqlGetValues "W##102.1", sql, billCompany
                        mnQuickBill(i + 1).Caption = billCompany
                        i = i + 1
                    End If
                    tbOrders.MoveNext
                Wend
                tbOrders.Close
            End If
        End If
        If i = 0 Then
            mnQuickBill(0).Visible = False
        End If
        
'        success = byErrSqlGetValues("##102.2", sql, lastBillCompany)
        
    Else
        mnBillFirma.Visible = False
        mnQuickBill(0).Visible = False
        For i = mnQuickBill.UBound To 1 Step -1
            Unload mnQuickBill(i)
        Next i
    End If
    Me.PopupMenu mnContext
ElseIf mousCol = orZakazano Then
    sql = "SELECT nomNom From sDMCrez WHERE (((sDMCrez.numDoc)='" & gNzak & "'));"
    byErrSqlGetValues "W##362", sql, str
    If str = "" Then
        MsgBox "У  заказа № " & gNzak & " нет пердметов!", , ""
    Else
        Nakladna.Show vbModal
    End If

ElseIf mousCol = orStatus Then
    
    sql = "SELECT  StatusId FROM bayOrders WHERE numOrder = " & gNzak
    Set tbOrders = myOpenRecordSet("##29", sql, dbOpenForwardOnly)
'    MsgBox sql
    If tbOrders.BOF Then
       tbOrders.Close
       MsgBox "Возможно он уже удален. Обновите Реестр", , "Заказ не найден!!!"
       Exit Sub
    End If
    statId = tbOrders!StatusId
    tbOrders.Close
   
   If statId = 7 Then ' "аннулирован"
     If dostup = "a" Then
        listBoxInGridCell lbAnnul, Grid, "select"
     Else
        listBoxInGridCell lbDel, Grid, "select"
     End If
   Else
     listBoxInGridCell lbClose, Grid, "select"
   End If
   Exit Sub
ElseIf mousCol = orProblem Then
    listBoxInGridCell lbProblem, Grid, Grid.TextMatrix(mousRow, mousCol)
ElseIf mousCol = orMen Then
    listBoxInGridCell lbM, Grid, "select"
ElseIf mousCol = orOtgrugeno Then
    If IsNumeric(Grid.TextMatrix(mousRow, orInvoice)) Or tbEnable.Visible Then
        GoTo AA
'        textBoxInGridCell tbMobile, Grid
    ElseIf MsgBox("", vbYesNo, "Счет ?") = vbYes Then
        Grid.col = orInvoice
        Grid.LeftCol = orInvoice
        Grid.SetFocus
    Else
'            flDelRowInMobile = True 'спрятать заказ
AA:     Otgruz.closeZakaz = (Grid.TextMatrix(mousRow, orStatus) = "закрыт")
        Otgruz.laZakaz = Grid.TextMatrix(mousRow, orNomZak)
        Otgruz.laFirm = Grid.TextMatrix(mousRow, orFirma)
        Otgruz.Show vbModal
'        textBoxInGridCell tbMobile, Grid
    End If
ElseIf mousCol = orDataVid Then
    textBoxInGridCell tbMobile, Grid
    If tbMobile.Text = "" Then
        tbMobile.Text = Format(Now(), "dd.mm.yy")
        tbMobile.SelLength = 2
    End If
Else
    textBoxInGridCell tbMobile, Grid
End If

End Sub


Public Sub Grid_EnterCell()
mousRow = Grid.row
mousCol = Grid.col
flDelRowInMobile = False
If zakazNum = 0 Then Exit Sub
beClick = True
tbInform.Text = Grid.TextMatrix(mousRow, mousCol)

If (dostup <> "a" And Grid.TextMatrix(mousRow, orStatus) = "закрыт") _
Or mousCol = orNomZak Or mousCol = orData Or mousCol = orLastMen _
Or (mousCol = orVrVid And Grid.TextMatrix(mousRow, orDataVid) = "") _
Or (mousCol = orVenture And Not (Grid.TextMatrix(mousRow, orOtgrugeno) = "" Or Grid.TextMatrix(mousRow, orOtgrugeno) = "0")) _
Then
    Grid.CellBackColor = vbYellow
    tbInform.Locked = True
Else
    Grid.CellBackColor = &H88FF88
    If mousCol = orOplacheno Then
        tbInform.Locked = False
    Else
        tbInform.Locked = True
    End If
End If

End Sub

Private Sub Grid_GotFocus()
'tbInform.Enabled = True
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
'laInform.Caption = "rrrow=" & Grid.row & "  cccol=" & Grid.col
If KeyCode = vbKeyReturn Then

    If mousCol = orFirma Then
'        If cbM.ListIndex < 0 Then
'        MsgBox "Заполните поле 'Менеджер'"
'        Exit Sub
'        End If

        gNzak = Grid.TextMatrix(mousRow, orNomZak)
    
        If zakazNum = 0 Then Exit Sub
        FindFirm.Regim = "edit"
        FindFirm.cmSelect.Visible = True
        FindFirm.tb.Text = Grid.TextMatrix(mousRow, orFirma)
        FindFirm.Show vbModal
    Else
        Grid_DblClick
    End If
End If
End Sub

Private Sub Grid_LeaveCell()
'laInform.Caption = laInform.Caption & "  lRow=" & Grid.row & "  lCol=" & Grid.col
Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_LostFocus()
'Grid.CellBackColor = Grid.BackColor
'tbInform.Enabled = False
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.colWidth(Grid.MouseCol)
End Sub

Private Sub Grid_RowColChange()
'laInform.Caption = laInform.Caption & "  chRow=" & Grid.row & " chCol=" & Grid.col

End Sub


Private Sub lbAnnul_DblClick()
Dim str As String, id As String

If noClick Then Exit Sub
' здесь изм-ся статус "закрыт" и "аннулирован"
str = Grid.TextMatrix(mousRow, mousCol) ' старое значение
If lbAnnul.Text = str Then GoTo EN1 '  значение не  поменялось
If lbAnnul.Text = "аннулирован" Then
    do_Annul
ElseIf lbAnnul.Text = "закрыт" Then
    If orderClose Then
        visits "+"    ' коррекция посещения фирмой
        Grid.TextMatrix(mousRow, mousCol) = lbAnnul.Text
    End If
ElseIf lbAnnul.Text = "удалить" Then
    do_Del
ElseIf lbAnnul.Text = "принят" Then
    id = 0: GoTo AA
ElseIf lbAnnul.Text = "готов" Then
    id = 4
AA: If MsgBox("Такое изменение статуса можно применить только в нештатных " & _
    "ситуациях. Если Вы уверены , нажмите <Да>, затем внимательно " & _
    "просмотрите все поля заказа на соответствие новому статусу." & _
    vbCrLf & vbCrLf & "Если после этого надо будет еще удалить и отгрузку " & _
    "по заказу, то поле 'Цех' оставте пустым.", _
    vbDefaultButton2 Or vbYesNo, "Внимание!!") = vbNo Then GoTo EN1
    wrkDefault.BeginTrans
BB: str = manId(cbM.ListIndex)
    ValueToTableField "##50", str, "BayOrders", "lastManagId"
    If ValueToTableField("##50", id, "BayOrders", "StatusId") = 0 Then
        Grid.TextMatrix(mousRow, mousCol) = lbAnnul.Text
        wrkDefault.CommitTrans
    Else
        wrkDefault.Rollback
    End If
End If
EN1:
lbHide
End Sub

Private Sub lbAnnul_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbAnnul_DblClick
End Sub




Private Sub lbClose_DblClick()
Dim str As String

If noClick Then Exit Sub
If lbClose.Visible = False Then Exit Sub
If lbClose.Text = "закрыт" Then ' из принят
    If orderClose Then
        visits "+"    ' коррекция посещения фирмой
        Grid.TextMatrix(mousRow, mousCol) = lbClose.Text
        Grid.CellBackColor = vbYellow
    End If
ElseIf lbClose.Text = "аннулирован" Then
    do_Annul "no_visit"
ElseIf lbClose.Text = "принят" Then
    str = "0"
    GoTo AA
ElseIf lbClose.Text = "выдан" Then
    str = "4"
    GoTo AA
ElseIf lbClose.Text = "собран" Then
    str = "1" '
AA: ValueToTableField "##50", str, "BayOrders", "StatusId"
    str = manId(cbM.ListIndex)
    ValueToTableField "##50", str, "BayOrders", "lastManagId"
    Grid.TextMatrix(mousRow, mousCol) = lbClose.Text
    'visits "-"  ' коррекция посещения фирмой здесь не нужна
End If
lbHide
    
End Sub

Sub do_Annul(Optional txt As String = "")
Dim str As String
    numDoc = gNzak
    If beNaklads("noMsg") Then
        MsgBox "У этого заказа есть накладные. Сначала удалите их.", , "Аннулирование невозможно!"
        Exit Sub
    End If
    If havePredmeti Then
        MsgBox "У этого заказа есть предметы. Сначала удалите их.", , "Аннулирование невозможно!"
        Exit Sub
    End If

    wrkDefault.BeginTrans
    If txt = "" Then visits "-"  ' коррекция посещения фирмой
    str = manId(cbM.ListIndex)
    ValueToTableField "##326", str, "BayOrders", "lastManagId"
    If ValueToTableField("##326", 7, "BayOrders", "StatusId") = 0 Then
        Grid.TextMatrix(mousRow, mousCol) = "аннулирован"
        wrkDefault.CommitTrans
    Else
        wrkDefault.Rollback
    End If

End Sub


Private Sub lbClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbClose_DblClick

End Sub

Sub do_Del()
  If MsgBox("По кнопке <Да> вся информация по заказу будет безвозвратно " & _
  "удалена из базы!", vbDefaultButton2 Or vbYesNo, "Удалить заказ " & _
  gNzak & " ?") = vbYes Then
    wrkDefault.BeginTrans
    
    'услуги удал-ся автоматом (каскадно)
    
    'по идее его там уже нет, т.к. не позволяет аннулировать с пердметами
    sql = "DELETE From sDMCrez WHERE (((numDoc)=" & gNzak & "));"
    myExecute "##305", sql, 0
    
    sql = "DELETE FROM BayOrders WHERE (((numOrder)=" & gNzak & "));"
'    myBase.Execute sql
    If myExecute("##136", sql) = 0 Then
        delZakazFromGrid
        wrkDefault.CommitTrans
    Else
ERR1:   wrkDefault.Rollback
    End If
  End If

End Sub

Private Sub lbDel_DblClick()
If noClick Then Exit Sub
If lbDel.Visible = False Then Exit Sub
If lbDel.Text = "удалить" Then
  do_Del
End If
lbHide

End Sub

Sub delZakazFromGrid()
    zakazNum = zakazNum - 1 '
    If zakazNum = 0 Then
        clearGridRow Grid, mousRow
    Else
        Grid.RemoveItem mousRow
    End If
    Grid.col = orNomZak

End Sub

Private Sub lbDel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbDel_DblClick
End Sub

Private Sub lbM_DblClick()
Dim str As String, i As Integer

If noClick Then Exit Sub
If lbM.Visible = False Then Exit Sub
Grid.Text = lbM.Text
str = manId(lbM.ListIndex)
ValueToTableField "##22", str, "BayOrders", "ManagId"
str = manId(cbM.ListIndex)
ValueToTableField "##49", str, "BayOrders", "lastManagId"

lbHide


End Sub

Private Sub lbM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbM_DblClick
End Sub

Function isConflict() As Boolean
Dim problemId As String, ordered, paid, shipped, stat As String
Dim titl As String, msg As String

Const ukagite = " Укажите правильно проблему!"
isConflict = True


titl = "Заказ № " & gNzak & " с противоречиями!"
  
sql = "SELECT problemId,  paid,  StatusId " & _
"FROM bayOrders WHERE numOrder = " & gNzak
'MsgBox sql
If Not byErrSqlGetValues("##357", sql, problemId, paid, stat) _
Then Exit Function

ordered = getOrdered(gNzak)
shipped = getShipped(gNzak)

msg = "Заказ "
'msg = "Заказ"
'If IsNull(ordered) Then GoTo AA
'If Not IsNumeric(ordered) Then GoTo AA
If ordered < 0.01 Then
'AA: isConflict = True
    MsgBox msg & " не заказан.", , titl
    Exit Function
End If

'If IsNull(paid) Then GoTo BB
'If Not IsNumeric(paid) Then GoTo BB
If ordered - paid > 0.01 Then
'BB:
  If problemId <> 1 Then  'оплата
'   isConflict = True
    MsgBox msg & " недоплачен." & ukagite, , titl
    Exit Function
  End If
  GoTo EN1
End If
    
'If IsNull(shipped) Then GoTo СС
'If Not IsNumeric(shipped) Then GoTo СС
If ordered - shipped > 0.01 Then
'СС:
  If problemId <> 3 Then  'отгрузка
'    isConflict = True
    MsgBox msg & " не полностью отгружен." & ukagite, , titl
    Exit Function
  End If
  GoTo EN1
End If
    
If paid - ordered > 0.01 Then
  If problemId <> 4 Then  'переплата
'    isConflict = True
    MsgBox msg & " переплачен." & ukagite, , titl
    Exit Function
  End If
End If

EN1:
If problemId = 0 Then
    isConflict = False
Else
    MsgBox "Невозможно закрыть заказ поскольку у него установлена " & _
    "поблема", , "Заказ с проблемами!"
End If

End Function


Function orderClose() As Boolean
Dim sql2 As String, str As String, account_is_closed As Integer

orderClose = False
lbHide

If isConflict() Then Exit Function

'If Grid.TextMatrix(mousRow, orProblem) = "" Then
    
    If Not predmetiIsClose Then ' эта проверка нужна для заказов без работы
        MsgBox "У этого заказа есть несписанные предметы.", , "Закрытие невозможно!"
        Exit Function
    End If
    sql = "select wf_order_closed_comtex (" & gNzak & ", '" & Grid.TextMatrix(mousRow, orServername) & "')"
    byErrSqlGetValues "##45.1", sql, account_is_closed
    If account_is_closed <> 1 Then
        MsgBox "Нельзя закрыть заказ, до тех пор, пока он не закрыт в Бухгалтерии.", , "Закрытие невозможно!"
        Exit Function
    End If
        
    
    wrkDefault.BeginTrans   ' начало транзакции
        
    str = manId(cbM.ListIndex)
    ValueToTableField "##45", 6, "BayOrders", "StatusId"
    ValueToTableField "##48", str, "BayOrders", "lastManagId"
        
      
    wrkDefault.CommitTrans  ' подтверждение транзакции
    orderClose = True
'Else внес в isConflict
'    MsgBox "Невозможно закрыть заказ поскольку у него установлена " & _
'    "поблема", , "Заказ с проблемами!"
'End If
  
End Function


Private Sub lbProblem_DblClick()
Dim str As String, i As Integer, DNM As String

If noClick Then Exit Sub
If lbProblem.Visible = False Then Exit Sub

sql = "UPDATE BayOrders, BayGuideProblem SET BayOrders.lastManagId = " & _
manId(cbM.ListIndex) & ", BayOrders.ProblemId = [BayGuideProblem].[ProblemId] " & _
"WHERE (((BayGuideProblem.Problem)='" & lbProblem.Text & _
"') AND ((BayOrders.numOrder)=" & gNzak & "));"
'MsgBox sql
If myExecute("##49", sql) <> 0 Then GoTo EN1
Grid.Text = lbProblem.Text

DNM = Format(Now(), "dd.mm.yy hh:nn") & vbTab & Orders.cbM.Text & " " & gNzak ' именно vbTab
On Error Resume Next ' в некот.ситуациях один из Open logFile дает Err: файл уже открыт
Open logFile For Append As #2
Print #2, DNM & " проблема=" & lbProblem.Text & _
"   зак=" & Grid.TextMatrix(mousRow, orZakazano) & _
" опл=" & Grid.TextMatrix(mousRow, orOplacheno) & _
" отг=" & Grid.TextMatrix(mousRow, orOtgrugeno)
Close #2
EN1:
lbHide
End Sub

Private Sub lbProblem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbProblem_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub lbVenture_DblClick()
Dim str As Variant, i As Integer, newInv As String

If noClick Then Exit Sub
If lbVenture.Visible = False Then Exit Sub
i = ValueToTableField("##72", lbVenture.ListIndex, "BayOrders", "ventureId")
If i = 0 Then
    Grid.Text = lbVenture.Text
    If (lbVenture.ListIndex = 0) Then Grid.Text = ""
    sql = "select invoice from bayorders where numOrder = " & Grid.TextMatrix(mousRow, orNomZak)
    If Not byErrSqlGetValues("W##72.1", sql, newInv) Then myBase.Close: End
'    newInv = getValueFromTable("Orders", "invoice", "numOrder = " & Grid.TextMatrix(mousRow, orNomZak))
    Grid.TextMatrix(mousRow, orInvoice) = newInv
    str = getValueFromTable("GuideVenture", "sysname", "ventureId = " & lbVenture.ListIndex)
    If IsNull(str) Then str = ""
    Grid.TextMatrix(mousRow, orServername) = str
End If

lbHide

End Sub

Private Sub lbVenture_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbVenture_DblClick

End Sub

Private Sub mnAllOrders_Click()
Me.MousePointer = flexHourglass
Report.Regim = "allOrdersByFirmName"
Report.Show vbModal
Grid.SetFocus
Me.MousePointer = flexDefault
End Sub

Private Sub mnAnalityc_Click()
    Me.MousePointer = flexHourglass
    Analityc.managId = Orders.cbM.Text
    Analityc.applicationType = "bay"
    Analityc.Show vbModeless, Me
    
    Me.MousePointer = flexDefault
End Sub

Private Sub mnBillFirma_Click()
Dim ventureName As String

    ventureName = Grid.TextMatrix(mousRow, orVenture)
    If serverIsAccessible(ventureName) Then
        g_id_bill = mnBillFirma.Tag
        FirmComtex.Show vbModal
    Else
        MsgBox "Сервер " & ventureName & " не доступен ", , "Предупреждение"
    End If
    
End Sub

Private Sub mnExit_Click()
    exitAll
End Sub

Private Sub mnFirmFind_Click()
        If tbEnable.Visible Then
            FindFirm.cmAllOrders.Visible = True
            FindFirm.cmNoClose.Visible = True
            FindFirm.cmNoCloseFiltr.Visible = True
        End If
'        FindFirm.tb.Text = value
        FindFirm.Show vbModal

End Sub


Private Sub mnFirmsGuide_Click()
Me.MousePointer = flexHourglass
GuideFirms.Regim = "fromContext"
GuideFirms.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub mnGuideFirms_Click()
Me.MousePointer = flexHourglass
GuideFirms.Regim = "fromMenu"
GuideFirms.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub mnMeassure_Click()
cbM_LostFocus '$$2
End Sub

Private Sub mnMenu_Click()
cbM_LostFocus
End Sub


Private Sub mnNoArhivFiltr_Click()
loadFirmOrders "noArhiv"
End Sub

Private Sub mnNoClose_Click()
Me.MousePointer = flexHourglass
Report.Regim = "OrdersByFirmName"
Report.Show vbModal
Grid.SetFocus
Me.MousePointer = flexDefault

End Sub

Private Sub mnNoCloseFiltr_Click()
loadFirmOrders ""
End Sub

Private Sub mnNomenk_Click()
sProducts.Regim = "ostat"
sProducts.Show vbModal
End Sub

Private Sub mnQuickBill_Click(index As Integer)
    If index = 0 Then Exit Sub
    FirmComtex.makeBillChoice mnQuickBill(index).Tag, Grid.TextMatrix(mousRow, orServername)

End Sub

Private Sub mnRecalc_Click()
Me.MousePointer = flexHourglass
    'statistic
    'Report.statistic "all"
Me.MousePointer = flexDefault

End Sub

Private Sub mnReports_Click()
Reports.Show vbModal
End Sub

Private Sub mnSklad_Click()
cbM_LostFocus
End Sub

Private Sub tbEnable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If LCase(tbEnable.Text) <> "arh" And LCase(tbEnable.Text) <> "фкр" Then
        tbEnable.Text = ""
        tbEnable.Visible = False
    End If
    Grid.SetFocus
ElseIf KeyCode = vbKeyDelete Then
    minut = 0
    Timer1_Timer
End If
End Sub

Private Sub tbEnable_LostFocus()
If LCase(tbEnable.Text) = "arh" Or LCase(tbEnable.Text) = "фкр" Then ' см еще и onKeyDown
    laClos.Visible = True
    cbClose.Visible = True
    mnAllOrders.Visible = True
    mnSep2.Visible = True
    mnNoCloseFiltr.Visible = True
    mnNoClose.Visible = True
    mnReports.Visible = True
    If dostup = "a" Then
 '       mnReports.Visible = True
    Else
        minut = 5
        Timer1.Interval = 60000 ' 1 минута
        Timer1.Enabled = True
    End If
'    tbEnable.Text = ""
Else
    tbEnable.Visible = False
End If
On Error Resume Next ' например когда не нажав <Enter> нажимаем <F11>
Grid.SetFocus
End Sub

Private Sub tbEndDate_Change()
cmRefr.Caption = "Загрузить"

End Sub

Function isFloatFromMobile(field As String) As Boolean

        If checkNumeric(tbMobile.Text, 0) Then
            ValueToTableField "##23", tbMobile.Text, "BayOrders", field
            Grid.TextMatrix(mousRow, mousCol) = tbMobile.Text
            isFloatFromMobile = True
        Else
            tbMobile.SelStart = 0
            tbMobile.SelLength = Len(tbMobile.Text)
            isFloatFromMobile = False
        End If
End Function

Private Sub tbInform_GotFocus()
'If cbM.ListIndex < 0 Then
'    MsgBox "Заполните поле 'Менеджер'", , "Предупреждение"
'    Grid.SetFocus
'Else
    tbInform.SelStart = Len(tbInform.Text)
'End If

End Sub

Private Sub tbInform_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    gNzak = Grid.TextMatrix(mousRow, orNomZak)
    tbMobile.Text = tbInform.Text
    tbMobile_KeyDown vbKeyReturn, 0
ElseIf KeyCode = vbKeyEscape Then
    Grid.SetFocus
End If
End Sub

Private Sub tbMobile_Change()
tbInform.Text = tbMobile.Text
End Sub

Private Sub tbMobile_DblClick()
lbHide
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, DNM As String, s As Single

If KeyCode = vbKeyReturn Then
DNM = Format(Now(), "dd.mm.yy hh:nn") & vbTab & cbM.Text & " " & gNzak ' именно vbTab
   
    If mousCol = orVrVid Then
        If Not isNumericTbox(tbMobile, 9, 21) Then Exit Sub
        Grid.TextMatrix(mousRow, orVrVid) = tbMobile.Text
        GoTo BB
    ElseIf mousCol = orDataVid Then
        If Not isDateTbox(tbMobile, "fry") Then Exit Sub
        Grid.TextMatrix(mousRow, orDataVid) = tbMobile.Text
BB:     tmpDate = Grid.TextMatrix(mousRow, orDataVid)
        str = Grid.TextMatrix(mousRow, orVrVid)
        If str <> "" Then str = " " & str & ":00:00"
        str = "'" & Format(tmpDate, "yyyy-mm-dd") & str & "'"
        ValueToTableField "##24", str, "BayOrders", "outDateTime"
'    ElseIf mousCol = orZakazano Then
'        If Not isFloatFromMobile("ordered") Then Exit Sub
    ElseIf mousCol = orOplacheno Then
        If Not isFloatFromMobile("paid") Then Exit Sub
'    ElseIf mousCol = orOtgrugeno Then
'        If Not isFloatFromMobile("shipped") Then Exit Sub
'        s = Round(tbMobile.Text, 2)
'        If s = 0 Then
'            ValueToTableField "##78", "Null", "BayOrders", "shipped"
'            Grid.TextMatrix(mousRow, orOtgrugeno) = ""
'        ElseIf flDelRowInMobile Then
'            flDelRowInMobile = False
'            delZakazFromGrid
'        End If
    ElseIf mousCol = orInvoice Then
        If InStr(tbMobile.Text, "счет") > 0 Or tbMobile.Text = "0" Then
            str = Grid.TextMatrix(mousRow, orOtgrugeno)
            If IsNumeric(str) Then
                delZakazFromGrid
            Else 'если в "отгружено ничего нет"
                Grid.TextMatrix(mousRow, mousCol) = "счет ?"
            End If
            ValueToTableField "##77", "'" & "счет ?" & "'", "BayOrders", "Invoice"
        Else
            If Not isFloatFromMobile("Invoice") Then Exit Sub
        End If
    ElseIf mousCol = orVes Then
        str = "Ves": GoTo AA
    ElseIf mousCol = orSize Then
        str = "Size": GoTo AA
    ElseIf mousCol = orPlaces Then
        str = "Places"
AA:     If ValueToTableField("##77", "'" & tbMobile.Text & "'", "BayOrders", _
        str) = 0 Then
            Grid.TextMatrix(mousRow, mousCol) = tbMobile.Text
        End If
    End If
    str = manId(cbM.ListIndex)
    ValueToTableField "##48", str, "BayOrders", "lastManagId"
    
'    tbMobile.Visible = False
    GoTo CC
ElseIf KeyCode = vbKeyEscape Then
CC:
lbHide
End If

End Sub


Private Sub tbStartDate_Change()
cmRefr.Caption = "Загрузить"
End Sub

Private Sub tbStartDate_GotFocus()
oldValue = tbStartDate.Text
End Sub

Private Sub tbStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tbStartDate_LostFocus
End If
End Sub

Private Sub tbStartDate_LostFocus()
isDateTbox tbStartDate
End Sub

Sub LoadBase(Optional reg As String = "")
Dim i As Integer, v As Variant

laInform.Caption = ""
Grid.Visible = False
clearGrid Grid


zakazNum = 0
'LoadOrders********************************************************
sql = rowFromOrdersSQL & getSqlWhere & " ORDER BY BayOrders.inDate;"
'Debug.Print getSqlWhere
'Debug.Print sql
Set tqOrders = myOpenRecordSet("##08", sql, dbOpenForwardOnly)
'If tqOrders Is Nothing Then myBase.Close: End
If Not tqOrders.BOF Then
While Not tqOrders.EOF
 
 gNzak = tqOrders!numorder
  
 
 If zakazNum > 0 Then Grid.AddItem ""
 zakazNum = zakazNum + 1
 
 Grid.TextMatrix(zakazNum, orNomZak) = gNzak

    If Not IsNull(tqOrders!id_bill) Then 'срочный
         Grid.col = orFirma
         Grid.row = zakazNum
         Grid.CellForeColor = vbRed
    End If
 
' If tqOrders!StatusId < 6 Then '***************
    v = predmetiIsClose
    If Not v Then
        Grid.col = orNomZak
        Grid.row = zakazNum
        Grid.CellForeColor = 200
        Grid.col = orZakazano
        Grid.CellForeColor = 200
    
    ElseIf Not IsNull(v) And tqOrders!StatusId = 6 Then 'есть предметы и закрыт
        Grid.col = orNomZak
        Grid.row = zakazNum
        Grid.CellForeColor = &H8800& ' т.зел.
    ElseIf v Then 'все накладные закрыты
        Grid.col = orNomZak
        Grid.row = zakazNum
        Grid.CellForeColor = vbBlue
        
        Grid.col = orZakazano
        Grid.CellForeColor = vbBlue
    End If
' End If '*************************************
 'If reg = "list" And Not bilo Then GoTo NXT
 
 copyRowToGrid (zakazNum)

NXT:
 tqOrders.MoveNext
Wend

End If 'Not tqOrders.BOF
tqOrders.Close '*********************************************

laInform.Caption = " кол-во зап.: " & zakazNum

Grid.Visible = True
'rowViem zakazNum, Grid
If zakazNum > 0 Then
    Grid.col = 1
    Grid.row = zakazNum
    Grid.SetFocus
End If
End Sub

Function getSqlWhere() As String
Dim i As Integer

getSqlWhere = ""
For i = 0 To orColNumber
  If orSqlWhere(i) <> "" Then
    If getSqlWhere = "" Then
        getSqlWhere = "(" & orSqlWhere(i) & ")"
    Else
        getSqlWhere = getSqlWhere & " AND " & "(" & orSqlWhere(i) & ")"
    End If
'    MsgBox "orSqlWhere=" & orSqlWhere(i) & "  getSqlWhere=" & getSqlWhere
  End If
Next i
If getSqlWhere <> "" Then getSqlWhere = " WHERE (" & getSqlWhere & ")"
'MsgBox "Where = " & getSqlWhere
    
End Function

Function strWhereByValCol(value As String, col As Integer, Optional _
operator As String = "=") As String
Dim str As String, typ As String, oper As String

oper = " " & operator & " "
strWhereByValCol = ""
str = orSqlFields(col)
If str = "" Then
    MsgBox "По этому полю фильтр не предусмотрен"
    Exit Function
End If
typ = left$(str, 1)
str = Mid$(str, 2)
If typ = "d" Then
    If value = "" Then
        value = " Is Null"
    Else
        If operator = "=" Then
            value = left$(value, 6) & "20" & Mid$(value, 7, 2) 'это нужно если в Win98 установлен "гггг" - формат года
            value = " Like '" & value & "%'"
        ElseIf operator = "<" Then
            value = " <= '" & Format(value, "yyyy-mm-dd") & " 11:59:59 PM'"
        Else
            value = " >= '" & Format(value, "yyyy-mm-dd") & "'"
        End If
    End If
ElseIf typ = "s" Then
    value = " = '" & value & "'"
Else
    If value = "" Then
        value = " Is Null"
    Else
'        value = " = " & value
        value = oper & value
    End If
End If
strWhereByValCol = "(" & str & ")" & value

End Function

Sub loadFirmOrders(stat As String, Optional ordNom As String = "")
Dim i As Integer

For i = 1 To orColNumber
    orSqlWhere(i) = ""
Next i
If stat = "noArhiv" Then
    stat = ""
    setWhereInvoice ' только заказы со счетом или с еще отгрузка не начата
'    orSqlWhere(orInvoice) = "isNumeric(BayOrders.Invoice) = 1 OR " & _
    "(BayOrders.Invoice) Is Null " '$$6
End If
If stat <> "all" And stat <> "" Then
    orSqlWhere(orFirma) = "(BayGuideFirms.Name) = '" & stat & "'"
Else
    orSqlWhere(orFirma) = "(BayGuideFirms.Name) = '" & Grid.Text & "'"
End If
If stat <> "all" Then _
    orSqlWhere(orStatus) = "(BayOrders.StatusId)<>6"

MousePointer = flexHourglass
LoadBase
If ordNom <> "" Then findValInCol Grid, ordNom, orNomZak
MousePointer = flexDefault
begFiltrDisable
laFiltr.Visible = True
End Sub

Sub copyRowToGrid(row As Long)
Dim str  As String

 'Grid.TextMatrix(row, orNomZak) = numZak
 Grid.TextMatrix(row, orInvoice) = tqOrders!Invoice
 Grid.TextMatrix(row, orMen) = Manag(tqOrders!managId)
 Grid.TextMatrix(row, orFirma) = tqOrders!Name
 Grid.TextMatrix(row, orStatus) = status(tqOrders!StatusId)
 Orders.Grid.TextMatrix(row, orProblem) = Problems(tqOrders!problemId)
 
 LoadDate Grid, row, orData, tqOrders!inDate, "dd.mm.yy"
 LoadDate Orders.Grid, row, orDataVid, tqOrders!outDateTime, "dd.mm.yy"
 LoadDate Orders.Grid, row, orVrVid, tqOrders!outDateTime, "hh"
 
 
 gNzak = tqOrders!numorder
 
 
 'LoadNumeric Grid, row, orZakazano, tqOrders!ordered
 Grid.TextMatrix(row, orZakazano) = getOrdered(gNzak)
 LoadNumeric Grid, row, orOplacheno, tqOrders!paid
 'LoadNumeric Grid, row, orOtgrugeno, tqOrders!shipped
 Grid.TextMatrix(row, orOtgrugeno) = getShipped(gNzak)
 
 Grid.TextMatrix(row, orVes) = tqOrders!VES
 Grid.TextMatrix(row, orSize) = tqOrders!Size
 Grid.TextMatrix(row, orPlaces) = tqOrders!Places
 Grid.TextMatrix(row, orLastMen) = Manag(tqOrders!lastManagId)
 If Not IsNull(tqOrders!Venture) Then
    Grid.TextMatrix(row, orVenture) = tqOrders!Venture
 End If
 If Not IsNull(tqOrders!id_bill) Then
    Grid.TextMatrix(row, orBillId) = CStr(tqOrders!id_bill)
 End If
 If Not IsNull(tqOrders!id_voc_names) Then
    Grid.TextMatrix(row, orVocnameId) = CStr(tqOrders!id_voc_names)
 End If
 If Not IsNull(tqOrders!serverName) Then
    Grid.TextMatrix(row, orServername) = CStr(tqOrders!serverName)
 End If

End Sub

    
Function cbMOsetByText(cb As ComboBox, stat As Variant) As Boolean
    cbMOsetByText = False
Dim i As Integer, txt As String
    txt = ""
    If Not IsNull(stat) Then txt = CStr(stat)
    If txt = "готов" Then
        If cb.List(3) <> "готов" Then cb.AddItem "готов", 3
        If cb.List(4) <> "утвержден" Then cb.AddItem "утвержден", 4
        cb.ListIndex = 3
        cbMOsetByText = True
    ElseIf txt = "утвержден" Then
        If cb.List(3) = "готов" Then
            i = 4
        Else
            i = 3
        End If
        If cb.List(i) <> "утвержден" Then cb.AddItem "утвержден", i
        cb.ListIndex = i
    ElseIf txt = "собран" Then
        cb.ListIndex = 2
        cbMOsetByText = True
    ElseIf txt = "макет" Or txt = "образец" Then
        cb.ListIndex = 1
    Else
        cb.ListIndex = 0
    End If

End Function


Private Sub Timer1_Timer()
minut = minut - 1
If minut <= 0 Then
    cbClose.value = 0
    
    Timer1.Enabled = False
    tbEnable.Visible = False
    laClos.Visible = False
    cbClose.Visible = False
'    If flReportArhivOrders Then Unload Report не стирать
'    cmRefr_Click
End If
End Sub
