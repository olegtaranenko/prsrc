VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Zakaz 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Перемещение заказа в цеховую сводку"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ckCeh 
      Caption         =   "Цех"
      Height          =   255
      Left            =   6480
      TabIndex        =   29
      Top             =   5280
      Visible         =   0   'False
      Width           =   675
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3555
      Left            =   60
      TabIndex        =   27
      Top             =   4920
      Visible         =   0   'False
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   6271
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmNewUklad 
      Caption         =   "Новая укладка"
      Height          =   375
      Left            =   6480
      TabIndex        =   26
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cbO 
      Height          =   315
      ItemData        =   "Zakaz.frx":0000
      Left            =   8160
      List            =   "Zakaz.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2940
      Width           =   1035
   End
   Begin VB.ComboBox cbM 
      Height          =   315
      ItemData        =   "Zakaz.frx":0026
      Left            =   6660
      List            =   "Zakaz.frx":0033
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2940
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Left            =   7080
      Top             =   4140
   End
   Begin VB.TextBox tbDateMO 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   7
      Top             =   3360
      Width           =   915
   End
   Begin VB.TextBox tbDateRS 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   4
      Top             =   2220
      Width           =   915
   End
   Begin VB.TextBox tbReadyDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   3
      Top             =   1740
      Width           =   915
   End
   Begin VB.TextBox tbVrVipO 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   8
      Top             =   3780
      Width           =   915
   End
   Begin VB.ComboBox cbStatus 
      Height          =   315
      Left            =   8220
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton cmExit 
      Cancel          =   -1  'True
      Caption         =   "Выход"
      Height          =   315
      Left            =   8340
      TabIndex        =   12
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmRepit 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   8340
      TabIndex        =   11
      Top             =   4740
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmZapros 
      Caption         =   "Запросить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7560
      TabIndex        =   9
      Top             =   4260
      Width           =   975
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4515
      Left            =   60
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7964
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Дата"
         Object.Width           =   1535
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Принят"
         Object.Width           =   1296
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Фирм"
         Object.Width           =   1164
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ресурс"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Загруз"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Остатки"
         Object.Width           =   1429
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Тек.зак"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Нов.ост"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Ц.заг"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Ц.ост"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Ж.заг"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Ст.заг"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Ст.ост"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "заказ"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Ст.ост"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6780
      TabIndex        =   10
      Top             =   4740
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox tbWorkTime 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   2
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Left            =   8340
      TabIndex        =   28
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label laMO 
      Caption         =   "Макет                    Образец"
      Height          =   195
      Left            =   6840
      TabIndex        =   25
      Top             =   2700
      Width           =   2115
   End
   Begin VB.Label laZapas 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Запас:"
      Height          =   195
      Left            =   4680
      TabIndex        =   23
      Top             =   60
      Width           =   675
   End
   Begin VB.Label laError 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1080
      TabIndex        =   22
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label laNomZak 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8220
      TabIndex        =   0
      Top             =   330
      Width           =   915
   End
   Begin VB.Label laVrVipO 
      Caption         =   "Вр.вып. Образца"
      Enabled         =   0   'False
      Height          =   195
      Left            =   6540
      TabIndex        =   21
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label laDateMO 
      Caption         =   "Дата Мак.\Обр."
      Enabled         =   0   'False
      Height          =   195
      Left            =   6540
      TabIndex        =   20
      Top             =   3420
      Width           =   1275
   End
   Begin VB.Label laHnomZak 
      Caption         =   "№ Заказа"
      Height          =   195
      Left            =   6540
      TabIndex        =   19
      Top             =   360
      Width           =   855
   End
   Begin VB.Label laStatus 
      Caption         =   "Статус:"
      Height          =   195
      Left            =   6540
      TabIndex        =   18
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label laDateRS 
      Alignment       =   2  'Center
      Caption         =   "Дата Р\С (не позже)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   6540
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label laMess 
      Height          =   555
      Left            =   420
      TabIndex        =   16
      Top             =   5100
      Width           =   5835
   End
   Begin VB.Label laReadyDate 
      Caption         =   "Дата выдачи"
      Height          =   195
      Left            =   6540
      TabIndex        =   15
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label laWorkTime 
      Caption         =   "Время выполнения"
      Height          =   255
      Left            =   6540
      TabIndex        =   14
      Top             =   1320
      Width           =   1515
   End
End
Attribute VB_Name = "Zakaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public urgent As String '"y" - срочный заказ
Public Regim As String
Dim isTimeZakaz As Boolean ' тождественен "заказ передается в цех"
Dim oldHeight As Integer, oldWidth As Integer
Dim ZbDay As Integer, ZeDay As Integer, ObDay As Integer, OeDay As Integer

Dim FormIsActiv As Boolean
Dim perekr As Double  'перекрываемая часть MO
Dim parts As String
Dim be_cmRepit As Boolean
Dim tmpMaxDay As Integer
Dim perenos As Integer ' этапы переноса
Dim workChange As Boolean
Dim isMzagruz As Boolean
Dim quantity As Integer
Const zgNomZak = 1
Const zgStatus = 2
Const zgVrVip = 3
Const zgNevip = 4
Const zgInDate = 5
Const zgOutDate = 6
Const zgOtlad = 7

Sub lvAddDay(I As Integer)
Dim item As ListItem, str As String
    str = Format(DateAdd("d", I - 1, curDate), "dd/mm/yy")
    Set item = Zakaz.lv.ListItems.Add(, "k" & I, str)
    day = Weekday(DateAdd("d", I - 1, curDate))
    If day = vbSunday Or day = vbSaturday Then item.ForeColor = &HFF
End Sub

Sub lvAddDays(Optional newLen As Integer = -1)
Dim I As Integer, j As Integer

If newLen < 0 Then newLen = maxDay

j = lv.ListItems.count
If newLen > j Then ' j=0 когда startParams вызыв-ся первый раз
    For I = j + 1 To newLen
        lvAddDay I
    Next I
ElseIf newLen < j Then
    For I = newLen + 1 To j
        lv.ListItems.Remove "k" & I
    Next I
    End If
End Sub

Private Sub cbM_Click()
cmZapros.Enabled = True
If cbM.Text = "в работе" Or cbM.Text = "готов" Then
    If FormIsActiv Then Zakaz.cmZapros.Enabled = True
    laDateMO.Enabled = True
    tbDateMO.Enabled = True
ElseIf Not (cbO.Text = "в работе" Or cbO.Text = "готов") Then
    laDateMO.Enabled = False
    tbDateMO.Enabled = False
    tbDateMO.Text = ""
End If
End Sub

Private Sub cbO_Click()
cmZapros.Enabled = True
If cbO.Text = "в работе" Or cbO.Text = "готов" Then
    If FormIsActiv Then Zakaz.cmZapros.Enabled = True
    laDateMO.Enabled = True
    tbDateMO.Enabled = True
    laVrVipO.Enabled = True
    tbVrVipO.Enabled = True
Else
    If Not (cbM.Text = "в работе" Or cbM.Text = "готов") Then
        laDateMO.Enabled = False
        tbDateMO.Enabled = False
        tbDateMO.Text = ""
    End If
    laVrVipO.Enabled = False
    tbVrVipO.Enabled = False
    tbVrVipO.Text = ""
End If

End Sub


Private Sub cmNewUklad_Click()

newZagruz

End Sub

Sub getBegEndDays(Optional stat As String = "")
Dim priemData As String

If Regim = "setka" Then ' по F1 F2
    priemData = Format(curDate, "dd.mm.yy")
Else
    priemData = Orders.Grid.TextMatrix(Orders.Grid.row, orData)
End If
If stat = "образец" Then
    Grid.TextMatrix(quantity, zgInDate) = priemData
    Grid.TextMatrix(quantity, zgOutDate) = tbDateMO.Text
    Grid.TextMatrix(quantity, zgOtlad) = ObDay & " " & OeDay
    Exit Sub
ElseIf stat <> "" Then
    Grid.TextMatrix(quantity, zgOtlad) = ZbDay & " " & ZeDay
    If IsDate(tbDateRS.Text) Then
        Grid.TextMatrix(quantity, zgInDate) = tbDateRS.Text
    Else
        Grid.TextMatrix(quantity, zgInDate) = priemData
    End If
    Grid.TextMatrix(quantity, zgOutDate) = tbReadyDate.Text
    Exit Sub
End If

'ZbDay = DateDiff("d", curDate, priemData)
'ObDay = ZbDay

If IsDate(tbReadyDate.Text) Then
    ZbDay = DateDiff("d", curDate, priemData)
    ZeDay = DateDiff("d", curDate, tbReadyDate.Text)
End If
If IsDate(tbDateRS.Text) Then
    ZbDay = DateDiff("d", curDate, tbDateRS.Text)
End If
If tbVrVipO.Enabled Then
    ObDay = DateDiff("d", curDate, priemData)
    OeDay = DateDiff("d", curDate, tbDateMO.Text)
End If

End Sub
' заказ в реестр отладки
Function zakazToGrid(reg As String, stat As String, nevip As String) As Boolean
zakazToGrid = False

If reg = "" Then
    quantity = quantity + 1
    Grid.TextMatrix(quantity, zgNomZak) = laNomZak.Caption
    Grid.TextMatrix(quantity, zgStatus) = stat
    If stat = "образец" Then
        Grid.TextMatrix(quantity, zgVrVip) = tbVrVipO.Text
    Else
        Grid.TextMatrix(quantity, zgVrVip) = tbWorkTime.Text
    End If
    Grid.TextMatrix(quantity, zgNevip) = nevip
    
    getBegEndDays stat
    
    Grid.AddItem ""
Else
    If laNomZak.Caption = Grid.TextMatrix(Grid.row, zgNomZak) Then
        If Grid.TextMatrix(Grid.row, zgStatus) = "образец" Then
            If stat = "образец" Then Exit Function
        ElseIf stat <> "образец" Then
            Exit Function ' обрабатываем только до(включит-но) отмеч.заказа
        End If
    End If
End If
zakazToGrid = True
End Function

'для новой раскладки заказов
'reg="tail" для позаказного просмотра(только при вызове из этой формы)
'reg = "fromCehZagruz"
'reg = "setka" по F1,F2 - т.е. первый раз
Sub newZagruz(Optional reg As String = "")
Dim s As Double, nevip As Double, I As Integer
Dim bDay As Integer, eDay As Integer, stat As String

'isMzagruz = (frm.name = "Zakaz")
If reg = "fromCehZagruz" Then
    isMzagruz = False
Else
    isMzagruz = (ckCeh.value = 0)
End If

'ZeDay = maxDay + 1  'чтобы не сработали !!эти значение исп-ся далее 4 раза
'OeDay = ZeDay
ObDay = -32000: ZbDay = ObDay ' флаг, что соотв.даты не заполнены (в 4х местах)
If isMzagruz Then getBegEndDays 'ZbDay,ObDay,ZbDay,ObDay (если заполнены)

If reg = "" Then ' не из Enter_cell
    Grid.Clear
    Grid.Rows = 2
    Grid.FormatString = "|<№ Заказа|<Статус|Вр.вып|Нев-но|Нач.дата|Кон.дата|"
    Grid.ColWidth(0) = 0
    Grid.ColWidth(zgStatus) = 765
    Grid.ColWidth(zgOutDate) = 825
    Grid.ColWidth(zgOtlad) = 525
    quantity = 0
End If


ReDim ost(maxDay): ReDim befOst(maxDay)
Dim firstRes As Double
'firstRes = Round(nr * Nstan * kpd, 1)  '23.11.04
firstRes = nr * Nstan * kpd             '
ost(1) = firstRes
befOst(1) = firstRes
For I = 2 To maxDay
    ost(I) = nomRes(I) * kpd * Nstan                '23.11.04
    befOst(I) = nomRes(I) * kpd * Nstan             '
'    ost(i) = Round(nomRes(i) * kpd * Nstan, 1)     '
'    befOst(i) = Round(nomRes(i) * kpd * Nstan, 1)  '
Next I
'VrVipParts заменнили на Nevip
sql = "SELECT Orders.numOrder, Orders.workTime, " & _
"DateDiff(day,Now(),Orders.outDateTime) AS endDay, " & _
"DateDiff(day,Now(),Orders.inDate) AS begDay, Orders.outDateTime, " & _
"Orders.inDate, Orders.StatusId, OrdersInCeh.Nevip, OrdersInCeh.urgent " & _
"FROM Orders INNER JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " & _
"Where (((Orders.StatusId) = 1 Or (Orders.StatusId) = 5) AND ((Orders.CehId)= " & cehId & ")) " & _
"UNION ALL " & _
"SELECT Orders.numOrder, Orders.workTime, " & _
"DateDiff(day,Now(),Orders.outDateTime) AS endDay, " & _
"DateDiff(day,Now(),Orders.DateRS) AS begDay, Orders.outDateTime, " & _
"Orders.DateRS, Orders.StatusId, OrdersInCeh.Nevip, OrdersInCeh.urgent " & _
"FROM Orders INNER JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " & _
"Where (((Orders.StatusId) = 2 Or (Orders.StatusId) = 3) AND ((Orders.CehId)= " & cehId & ")) " & _
"UNION ALL " & _
"SELECT Orders.numOrder, OrdersMO.workTimeMO, " & _
"DateDiff(day,Now(),OrdersMO.DateTimeMO) AS endDay, " & _
"DateDiff(day,Now(),Orders.inDate) AS begDay, OrdersMO.DateTimeMO, " & _
"Orders.inDate, 1 AS StatusId, -1 AS Nevip, '' AS urgent " & _
"FROM Orders INNER JOIN OrdersMO ON Orders.numOrder = OrdersMO.numOrder " & _
"Where (((OrdersMO.statO) = 'в работе') AND ((Orders.CehId)= " & cehId & _
")) ORDER BY "
'")) ORDER BY endDay, begDay DESC;"

If isMzagruz Then
    sql = sql & "4 DESC;" ' в порядке уменьшения Даты Начала
Else
    sql = sql & "3;" ' в порядке увеличения  Даты Конца
End If
'Debug.Print sql
Set tbOrders = myOpenRecordSet("##370", sql, dbOpenForwardOnly) ', dbOpenDynaset)
If tbOrders Is Nothing Then Exit Sub
While Not tbOrders.EOF
If tbOrders!Numorder = 5021505 Then
    quantity = quantity
End If
    bDay = tbOrders!begDay '  отн. Now()
    eDay = tbOrders!endDay '  отн. Now()
'    If eDay > maxDay Then msgOfEnd ("##371")
    
    If isMzagruz Then 'у Менеджеров тек.заказ б.браться всегда из формы, поэ-
        If tbOrders!Numorder = laNomZak.Caption Then GoTo NXT ' тому в базе
        'его пропускаем, т.к. у него будет новое место среди других
        
'вставляем тек.заказ так, чтобы не нарушить порядок уменьшения Даты Начала
'        If eDay > OeDay Or (eDay = OeDay And bDay <= ObDay) Then ' не нарушаем сортировку
        If bDay <= ObDay Then  ' не нарушаем сортировку
            dayCorrect ObDay, OeDay
            ukladka ost, OeDay, ObDay, tbVrVipO.Text 'обратная укладка  (в bef не попадает)
            If Not zakazToGrid(reg, "образец", tbVrVipO.Text) Then GoTo EN1
'            OeDay = maxDay + 1 ' чтобы более не срабатывал
            ObDay = -32000 ' чтобы более не срабатывал
        End If
'        If eDay > ZeDay Or (eDay = ZeDay And bDay <= ZbDay) Then ' не нарушаем сортировку
        If bDay <= ZbDay Then ' не нарушаем сортировку
            dayCorrect ZbDay, ZeDay, urgent
            ukladka ost, ZeDay, ZbDay, tbWorkTime.Text 'обратная укладка (в bef не попадает)
            If Not zakazToGrid(reg, cbStatus.Text, tbWorkTime.Text) Then GoTo EN1
            'ZeDay = maxDay + 1 ' чтобы более не срабатывал
            ZbDay = -32000 ' чтобы более не срабатывал
        End If
'        If tbOrders!numOrder = laNomZak.Caption Then GoTo NXT 'а не из базы
    End If
    
    If eDay > maxDay Then msgOfEnd ("##371")
    
'    dayCorrect bDay, eDay, tbOrders!urgent спец.распределение срочн. заказов приводит к тому, что м.б. разные минусы у М и в цеху
    dayCorrect bDay, eDay, ""
    
    If tbOrders!nevip = -1 Then '"образец"
        nevip = tbOrders!workTime
    Else
        nevip = Round(tbOrders!workTime * tbOrders!nevip, 2)
    End If
    If isMzagruz Then
        ukladka ost, eDay, bDay, nevip 'обратная укладка
        ukladka befOst, eDay, bDay, nevip 'обратная укладка
    Else
        ukladka ost, bDay, eDay, nevip
        If tbOrders!StatusId = 1 Or tbOrders!StatusId = -1 Then _
            ukladka befOst, bDay, eDay, nevip ' жывые(в раб. и образец)
    End If
    
    If reg = "" Then
      quantity = quantity + 1
      Grid.TextMatrix(quantity, zgNomZak) = tbOrders!Numorder
'      If tbOrders!StatusId = -1 Then
      If tbOrders!nevip = -1 Then '"образец"
        Grid.TextMatrix(quantity, zgStatus) = "образец"
      Else
        Grid.TextMatrix(quantity, zgStatus) = status(tbOrders!StatusId)
      End If
      Grid.TextMatrix(quantity, zgVrVip) = tbOrders!workTime
      Grid.TextMatrix(quantity, zgNevip) = nevip
      Grid.TextMatrix(quantity, zgInDate) = Format(tbOrders!inDate, "dd.mm.yy")
      Grid.TextMatrix(quantity, zgOutDate) = Format(tbOrders!outDateTime, "dd.mm.yy")
      Grid.TextMatrix(quantity, zgOtlad) = bDay & " " & eDay
      Grid.AddItem ""
    End If
    
    If reg = "tail" Then ' из Enter_cell
      If tbOrders!Numorder = Grid.TextMatrix(Grid.row, zgNomZak) Then
        If Grid.TextMatrix(Grid.row, zgStatus) = "образец" Then
          If tbOrders!nevip = -1 Then GoTo EN1 '"образец"
        ElseIf tbOrders!nevip <> -1 Then
          GoTo EN1 ' обрабатываем только до(включит-но) отмеч.заказа
        End If
      End If
    End If
NXT:
    tbOrders.MoveNext
Wend

If isMzagruz Then
'если в базе нет заказов кот позже чем текущий то ZeDay и м.б. и OeDay не
'не сработали, т.е. надо проверить и сработать
'  If OzDay < maxDay + 1 Then  '
  If ObDay > -32000 Then  '
    dayCorrect ObDay, OeDay
    ukladka ost, OeDay, ObDay, tbVrVipO.Text 'обратная укладка  (в bef не попадает)
    zakazToGrid reg, "образец", tbVrVipO.Text
  End If
'  If ZeDay < maxDay + 1 Then
  If ZbDay > -32000 Then
    dayCorrect ZbDay, ZeDay
    ukladka ost, ZeDay, ZbDay, tbWorkTime.Text 'обратная укладка (в bef не попадает)
    zakazToGrid reg, cbStatus.Text, tbWorkTime.Text
  End If
End If

If reg = "" And quantity > 0 Then Grid.RemoveItem Grid.Rows - 1
EN1:
tbOrders.Close

If reg = "fromCehZagruz" Then Exit Sub

If ckCeh.value = 0 Then
  For I = 1 To maxDay
    lv.ListItems("k" & I).SubItems(zkMbef) = Round(befOst(I), 1) '23.11.04
    lv.ListItems("k" & I).SubItems(zkMzagr) = _
               Round(nomRes(I) * kpd * Nstan - befOst(I), 1)
    lv.ListItems("k" & I).ListSubItems(zkMbef).Bold = False
    lv.ListItems("k" & I).ListSubItems(zkMbef).ForeColor = 0
    If reg = "setka" Then
        If befOst(I) < 0 Then
            lv.ListItems("k" & I).ListSubItems(zkMbef).Bold = True
            lv.ListItems("k" & I).ListSubItems(zkMbef).ForeColor = 200
        End If
    Else
        lv.ListItems("k" & I).SubItems(zkMost) = Round(ost(I), 1) '23.11.04
   
        lv.ListItems("k" & I).ListSubItems(zkMost).Bold = False
        lv.ListItems("k" & I).ListSubItems(zkMost).ForeColor = 0
        If befOst(I) < 0 Then
            lv.ListItems("k" & I).ListSubItems(zkMbef).Bold = True
            lv.ListItems("k" & I).ListSubItems(zkMbef).ForeColor = 200
            If ost(I) < befOst(I) Then GoTo AA
        ElseIf ost(I) < 0 Then
AA:         lv.ListItems("k" & I).ListSubItems(zkMost).Bold = True
            lv.ListItems("k" & I).ListSubItems(zkMost).ForeColor = 200
        ElseIf ost(I) <> befOst(I) Then
            lv.ListItems("k" & I).ListSubItems(zkMost).Bold = True
        End If
    End If
  Next I
  lv.ListItems("k1").SubItems(zkMzagr) = Round(firstRes - befOst(1), 1) '23.11.04
Else
  For I = 1 To maxDay
   lv.ListItems("k" & I).SubItems(zkCost) = Round(ost(I), 1) '23.11.04
   lv.ListItems("k" & I).SubItems(zkCliv) = _
            Round(nomRes(I) * kpd * Nstan - befOst(I), 1)
   lv.ListItems("k" & I).SubItems(zkCzagr) = _
            Round(nomRes(I) * kpd * Nstan - ost(I), 1)
  Next I
  lv.ListItems("k1").SubItems(zkCzagr) = Round(firstRes - ost(1), 1) '23.11.04
  lv.ListItems("k1").SubItems(zkCliv) = Round(firstRes - befOst(1), 1) '23.11.04
End If

End Sub
    
Sub dayCorrect(bDay As Integer, eDay As Integer, Optional urgen As String = "")
    bDay = bDay + 1: eDay = eDay + 1 'корр-я отн-но DateDiff(,now())
    If bDay < 1 Then bDay = 1
    If urgen = "" Then ' не срочный
        eDay = getPrev2DayRes_(eDay) 'за 2 дня
    End If
    If bDay > eDay Then bDay = eDay
End Sub

Sub ukladka(ost() As Double, bDay As Integer, eDay As Integer, ByVal nevip As Double)
Dim I As Integer, stp As Integer

stp = 1
If bDay > eDay Then stp = -1
For I = bDay To eDay Step stp
    If ost(I) > 0 Then ' на отриц ресурс не распределяем
        ost(I) = Round(ost(I) - nevip, 2)
        If ost(I) >= 0 Then
            nevip = 0
            Exit Sub
        End If
        nevip = -ost(I)
        ost(I) = 0
    End If
Next I
If nevip > 0 Then
    I = max(bDay, eDay)
    ost(I) = ost(I) - nevip
End If
End Sub

Sub formMaximize()
Dim oldWidth As Integer
    Me.WindowState = vbMaximized
     cmNewUklad.Visible = True
    Grid.Visible = True
    ckCeh.Visible = True
    Label1.Visible = True
    lv.ColumnHeaders(zkCzagr + 1).Width = 680
    lv.ColumnHeaders(zkCost + 1).Width = 680
    lv.ColumnHeaders(zkCliv + 1).Width = 680
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If Shift = vbCtrlMask And KeyCode = vbKeyO Then
    formMaximize
End If
End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If Me.WindowState = vbMinimized Then Exit Sub
'lbHide

On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width

'lv.Height = lv.Height + h
lv.Width = lv.Width + w
laHnomZak.Left = laHnomZak.Left + w
laNomZak.Left = laNomZak.Left + w
laStatus.Left = laStatus.Left + w
cbStatus.Left = cbStatus.Left + w
laWorkTime.Left = laWorkTime.Left + w
tbWorkTime.Left = tbWorkTime.Left + w
laReadyDate.Left = laReadyDate.Left + w
tbReadyDate.Left = tbReadyDate.Left + w
laDateRS.Left = laDateRS.Left + w
tbDateRS.Left = tbDateRS.Left + w
laMO.Left = laMO.Left + w
cbM.Left = cbM.Left + w
cbO.Left = cbO.Left + w
laDateMO.Left = laDateMO.Left + w
tbDateMO.Left = tbDateMO.Left + w
laVrVipO.Left = laVrVipO.Left + w
tbVrVipO.Left = tbVrVipO.Left + w
cmZapros.Left = cmZapros.Left + w
cmAdd.Left = cmAdd.Left + w
cmRepit.Left = cmRepit.Left + w
cmExit.Left = cmExit.Left + w
cmExit.Top = cmExit.Top + h

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Если именно мы блокировали:
If getSystemField("resursLock") = Orders.cbM.Text Then unLockBase
Orders.Grid_EnterCell ' подсветка ячейки
End Sub

Private Sub Grid_EnterCell()
Static I As Integer

Grid.CellBackColor = vbButtonFace
If quantity > 0 Then
    I = I + 1
    Label1.Caption = I
    newZagruz "tail"
End If
End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_LostFocus()
Grid_LeaveCell
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)

End Sub

Private Sub tbDateMO_GotFocus()
If FormIsActiv Then Zakaz.cmZapros.Enabled = True
If tbDateMO.Text = "" Then
    tbDateMO.Text = Format(curDate, "dd/mm/yy")
End If
tbDateMO.SelStart = 0
tbDateMO.SelLength = 2

End Sub

Private Sub cbStatus_Click()
Dim noClick As Boolean

'If noClick Then
'    noClick = False
'    Exit Sub
'End If
If FormIsActiv Then Zakaz.cmZapros.Enabled = True
If cbStatus.Text = "в работе" Then
'    If ((cbO.Text <> "" And cbO.Text <> "утвержден") _
'    Or (cbM.Text <> "" And cbM.Text <> "утвержден")) And FormIsActiv Then
'        noClick = True
'        MsgBox "Для перевода заказа в работу необходимо, чтобы макет " & _
'        "и(или) образец были утверждены.", , "Недопустимый статус!"
'        cbStatus.Text = "согласов"
'        Exit Sub
'    End If
    laMO.Enabled = False
    cbM.Enabled = False
    cbO.Enabled = False
    tbDateMO.Enabled = False
    laVrVipO.Enabled = False
    tbVrVipO.Enabled = False
'    End If
'End If
'If cbStatus.Text = "согласов" Then
ElseIf cbStatus.Text = "согласов" Then
    cbM.Enabled = True
    cbO.Enabled = True
'    tbDateMO.Enabled = True
'    tbVrVipO.Enabled = True
    laMO.Enabled = True
Else
    laMO.Enabled = False
    cbM.Enabled = False
    cbO.Enabled = False
    tbDateMO.Enabled = False
    laVrVipO.Enabled = False
    tbVrVipO.Enabled = False
'    cbM.ListIndex = 0
'    cbO.ListIndex = 0
'    tbVrVipO.Text = ""
'    tbDateMO.Text = ""
End If

'If cbStatus.ListIndex = 2 Or cbStatus.ListIndex = 3 Then
If cbStatus.Text = "согласов" Or cbStatus.Text = "резерв" Then
    tbDateRS.Enabled = True             ' резерв согласование
    laDateRS.Enabled = True
    Zakaz.laWorkTime.Enabled = True
    Zakaz.laReadyDate.Enabled = True
    Zakaz.tbReadyDate.Enabled = True
    Zakaz.tbWorkTime.Enabled = True
'ElseIf cbStatus.ListIndex = 1 Then      'в работу
ElseIf cbStatus.Text = "в работе" Or cbStatus.Text = "отложен" Then
'    tbDateRS.Text = ""
    tbDateRS.Enabled = False
    laDateRS.Enabled = False
    Zakaz.laWorkTime.Enabled = True
    Zakaz.laReadyDate.Enabled = True
    Zakaz.tbReadyDate.Enabled = True
    Zakaz.tbWorkTime.Enabled = True
Else
    laWorkTime.Enabled = False
    laReadyDate.Enabled = False
    tbReadyDate.Text = ""
    tbReadyDate.Enabled = False
    tbWorkTime.Text = ""
    tbWorkTime.Enabled = False
    tbDateRS.Text = ""
    tbDateRS.Enabled = False
    laDateRS.Enabled = False
End If
'MsgBox "ListIndex=" & cbStatus.ListIndex & "  tbDateRS.Enabled=" & tbDateRS.Enabled
End Sub
'$odbc08$
Private Sub cmAdd_Click()
Dim I As Integer, str As String, item As ListItem, s As Double, t As Double
Dim id As Integer, VrVip As String, VrVipO As String, editWorkTime As Boolean
Dim workTime As String

'MaxDay = tmpMaxDay наверно это уже никому не нужно

Timer1.Enabled = False
'Set tbOrders = myOpenRecordSet("##01", "OrdersInCeh", dbOpenTable)
'If tbOrders Is Nothing Then GoTo AA

sql = "SELECT * from Orders WHERE (((numOrder)=" & gNzak & "));"
Set tbOrders = myOpenRecordSet("##30", sql, dbOpenForwardOnly) '$#$
'If tbOrders Is Nothing Then GoTo AA

'If Not findZakazInTable_("Orders") Then '$#$
If tbOrders.BOF Then
    tbOrders.Close
AA: If getSystemField("resursLock") = Orders.cbM.Text Then unLockBase 'Если именно мы блокирова
    MsgBox "Возможно он уже удален. Обновите Реестр", , "Заказ не найден!!!"
    Exit Sub
End If
id = statId(Zakaz.cbStatus.ListIndex)

Dim workTimeOld As Double, statIdOld As Integer
workTimeOld = 0
If Not IsNull(tbOrders!workTime) Then workTimeOld = tbOrders!workTime
statIdOld = tbOrders!StatusId
tbOrders.Close

wrkDefault.BeginTrans

'tbOrders.Edit
' коррекция посещения фирмой
If (statIdOld = 0 Or statIdOld = 7) And id <> 0 And id <> 7 Then
    visits "+"
ElseIf Not (statIdOld = 0 Or statIdOld = 7) And (id = 0 Or id = 7) Then
    visits "-"
End If

If id = 7 Then delZakazFromReplaceRS ' если аннулируемый заказ там есть

'tbOrders!rowLock = "
If id <> statIdOld Or (neVipolnen_O = 0 And neVipolnen = 0) Then
    editWorkTime = False    '
Else                        'если что-то недовып-но и статус не меняется
    editWorkTime = True     'то это необх. условие изменения Вр.Вып.
End If

str = tbReadyDate.Text
If str = "" Then
'    tbOrders!outDateTime = Null
    str = "Null"
Else
    str = "'" & "20" & Mid$(str, 7, 2) & "-" & Mid$(str, 4, 2) & "-" & Left$(str, 2)
    sql = Orders.Grid.TextMatrix(Orders.mousRow, orVrVid)
    If sql = "" Then
'        tbOrders!outDateTime = tbReadyDate.Text
        str = str & "'"
    Else
        str = str & " " & sql & ":00'"
'       tbOrders!outDateTime = str
    End If
End If

sql = "UPDATE Orders SET outDateTime = " & str & _
" WHERE (((numOrder)=" & gNzak & "));"
'MsgBox sql
'Debug.Print sql
If myExecute("##391", sql) <> 0 Then GoTo ER1

If tbDateRS.Enabled = True Then
'    tbOrders!dateRS = tbDateRS.Text
    str = tbDateRS.Text
    str = "'" & "20" & Mid$(str, 7, 2) & "-" & Mid$(str, 4, 2) & _
    "-" & Left$(str, 2) & "'"
    'str = "'" & Mid$(str, 4, 2) & "/" & Left$(str, 2) & "/20" & _
    Mid$(str, 7, 2) & "'"
Else
'    tbOrders!dateRS = Null
    str = "Null"
End If
sql = "UPDATE Orders SET dateRS = " & str & _
" WHERE (((Orders.numOrder)=" & gNzak & "));"
'MsgBox sql
If myExecute("##392", sql) <> 0 Then GoTo ER1

sql = "SELECT * from OrdersInCeh WHERE (((numOrder)=" & gNzak & "));"
Set tbOrders = myOpenRecordSet("##01", sql, dbOpenForwardOnly)
'If tbOrders Is Nothing Then GoTo AA

workTime = workTimeOld ' для случая, если не менялось
If Not tbOrders.BOF Then
    If isTimeZakaz Then
'        tbOrders.Edit
'        tbOrders!rowLock = ""
       If workChange Then
         If (id = 1 Or id = 5) And editWorkTime Then 'остается в работе или отложен
'            tbOrders!workTime = Round(tbOrders!workTime + tbWorkTime.Text _
                     - neVipolnen, 1) 'время с учетом коррекции
            workTime = Round(workTimeOld + tbWorkTime.Text _
                     - neVipolnen, 1) 'время с учетом коррекции
'            tbOrders!nevip = tbWorkTime.Text / tbOrders!workTime
            sql = "UPDATE OrdersInCeh SET Nevip = " & _
            tbWorkTime.Text / workTime & " WHERE (((numOrder)=" & gNzak & "));"
            If myExecute("##393", sql) <> 0 Then GoTo ER1
         Else
'            tbOrders!workTime = tbWorkTime.Text
            workTime = tbWorkTime.Text
         End If
       End If
       sql = "UPDATE OrdersInCeh SET urgent = '" & urgent & _
       "' WHERE (((OrdersInCeh.numOrder)=" & gNzak & "));"
       If myExecute("##403", sql) <> 0 Then GoTo ER1
       GoTo DD
    Else
'        tbOrders.Delete
        sql = "DELETE from OrdersInCeh WHERE (((numOrder)=" & gNzak & "));"
        If myExecute("##394", sql) <> 0 Then GoTo ER1
'        tbOrders!workTime = 0
        workTime = 0
    End If
Else
    If isTimeZakaz Then
'        tbOrders.AddNew
'        tbOrders!numOrder = gNzak
'        On Error GoTo ERRp
'        On Error GoTo 0
'        tbOrders!workTime = tbWorkTime.Text
        workTime = tbWorkTime.Text
        sql = "INSERT INTO OrdersInCeh ( numOrder, urgent )" & _
        "SELECT " & gNzak & ",'" & urgent & "';"
        If myExecute("##395", sql) <> 0 Then GoTo ER1
DD:     noClick = True
        Orders.Grid.col = orCeh
'        Grid.row = zakazNum
        If urgent = "y" Then
'            tbOrders!urgent = "y"
            Orders.Grid.CellForeColor = 200
        Else
'            tbOrders!urgent = ""
            Orders.Grid.CellForeColor = vbBlack
        End If
'        tbOrders.Update
        Orders.Grid.col = orStatus
        noClick = False
    End If
End If

sql = "UPDATE Orders SET workTime = " & workTime & _
", statusId = " & id & ", lastManagId = " & manId(Orders.cbM.ListIndex) & _
" WHERE (((Orders.numOrder)=" & gNzak & "));"
If myExecute("##396", sql) <> 0 Then GoTo ER1


' согласование или из согласования в работу
sql = "SELECT * from OrdersMO WHERE (((numOrder)=" & gNzak & "));"
Set table = myOpenRecordSet("##02", sql, dbOpenForwardOnly)
'If Not table Is Nothing Then '
 bilo = Not table.BOF
table.Close

' bilo = findZakazInTable(table) '1:
 
 If id = 3 Then ' согласов
'  If bilo Then      '
'    table.Edit      '
'  Else              '
'    table.AddNew    '
'    table!numOrder = gNzak
'  End If            '
'  table!statM = cbM.Text '
'  table!statO = cbO.Text '
  If cbM.Text = "в работе" Or cbM.Text = "готов" Or _
    cbO.Text = "в работе" Or cbO.Text = "готов" Then
    str = tbDateMO.Text
    str = "'" & "20" & Mid$(str, 7, 2) & "-" & Mid$(str, 4, 2) & "-" & Left$(str, 2)
    sql = Orders.Grid.TextMatrix(Orders.mousRow, orMOVrVid)
    If sql = "" Then
        str = str & "'"
'        table!DateTimeMO = tbDateMO.Text
    Else
        str = str & " " & sql & ":00'"
'        str = tbDateMO.Text & " " & str & ":00"
'        table!DateTimeMO = str
    End If
  Else
'    table!DateTimeMO = Null
    str = "Null"
  End If
  If cbO.Text = "в работе" Or cbO.Text = "готов" Then
'    table!workTimeMO = tbVrVipO
    workTime = tbVrVipO.Text
  Else
'    table!workTimeMO = Null
    workTime = "Null"
  End If
'  table.Update
  If bilo Then      '
    sql = "UPDATE OrdersMO SET DateTimeMO = " & str & ", workTimeMO = " & _
    workTime & ", StatM = '" & cbM.Text & "', StatO = '" & cbO.Text & _
    "' WHERE (((numOrder)=" & gNzak & "));"
  Else
    sql = "INSERT INTO OrdersMO ( numOrder, DateTimeMO, workTimeMO, StatM, " & _
    "StatO ) SELECT " & gNzak & ", " & str & ", " & workTime & ", '" & _
    cbM.Text & "', '" & cbO.Text & "';"
  End If
'  MsgBox sql
  If myExecute("##397", sql) <> 0 Then GoTo ER1
 Else
'  If bilo Then table.Delete
  If bilo Then
    sql = "DELETE from OrdersMO WHERE (((numOrder)=" & gNzak & "));"
    If myExecute("##398", sql) <> 0 Then GoTo ER1
  End If
 End If ' согласов
'End If ' If Not table Is Nothing
'table.Close
tbOrders.Close

' коррекция посещения фирмой
'If (tbOrders!statusId = 0 Or tbOrders!statusId = 7) And id <> 0 And id <> 7 Then
'    visits "+"
'ElseIf Not (tbOrders!statusId = 0 Or tbOrders!statusId = 7) And (id = 0 Or id = 7) Then
'    visits "-"
'End If
'tbOrders!statusId = id
'tbOrders!lastManagId = manId(Orders.cbM.ListIndex)
'tbOrders.Update
'tbOrders.Close
    
'******** перенос Даты RS ***********************************
'sql = "SELECT ReplaceRS.numOrder, ReplaceRS.newDateIn, ReplaceRS.newDateRS, " & _
'"ReplaceRS.newDateOut  From ReplaceRS " & _
'"Where (((ReplaceRS.numOrder) = " & gNzak & ")) ORDER BY ReplaceRS.newDateIn;"

'Set table = myOpenRecordSet("##22", sql, dbOpenForwardOnly)
'If Not table Is Nothing Then
If perenos = 1 Then ' был подтвержден перенос РС
    sql = "INSERT INTO ReplaceRS ( numOrder, newDateIn, newDateRS, newDateOut) " & _
    "SELECT " & gNzak & ", '" & _
    yymmdd(Orders.Grid.TextMatrix(Orders.mousRow, orData)) & "', '" & _
    yymmdd(Orders.Grid.TextMatrix(Orders.mousRow, orDataRS)) & "', '" & _
    yymmdd(Orders.Grid.TextMatrix(Orders.mousRow, orDataVid)) & "';"
'    MsgBox sql
    If myExecute("##399", sql) <> 0 Then GoTo ER1
'    table.AddNew
'    table!numOrder = gNzak 'yymmdd(
'    table!newDateIn = Orders.Grid.TextMatrix(Orders.mousRow, orData)
'    table!newDateRS = Orders.Grid.TextMatrix(Orders.mousRow, orDataRS)
'    table!newDateOut = Orders.Grid.TextMatrix(Orders.mousRow, orDataVid)
'    table.Update
    GoTo СС
  ElseIf perenos = 2 Then ' был подтвержден перенос РС
'    table.MoveLast
СС: ' table.AddNew
    sql = "INSERT INTO ReplaceRS ( numOrder, newDateIn, newDateRS, newDateOut) " & _
    "SELECT " & gNzak & ", '" & Format(Now(), "yyyy-mm-dd") & "', '" & _
    yymmdd(tbDateRS.Text) & "', '" & yymmdd(tbReadyDate.Text) & "';"
    If myExecute("##400", sql) <> 0 Then GoTo ER1
'    table!numOrder = gNzak
'    GoTo BB
  ElseIf perenos = 3 Then ' был подтвержден перенос РС
    sql = "SELECT Max(newDateIn) AS MaxDate from ReplaceRS " & _
    "WHERE (((numOrder)=" & gNzak & "));"
    If byErrSqlGetValues("##22", sql, str) Then
      If str <> "" Then
        sql = "UPDATE ReplaceRS SET newDateIn = '" & Format(Now(), "yyyy-mm-dd") & _
        "', newDateRS = '" & yymmdd(tbDateRS.Text) & _
        "', newDateOut = '" & yymmdd(tbReadyDate.Text) & _
        "' WHERE (((numOrder)=" & gNzak & ") AND (newDateIn)= '" & str & "');"
        If myExecute("##401", sql) <> 0 Then GoTo ER1
      End If
    End If
'    table.MoveLast
'    table.Edit
'BB: table!newDateIn = Format(Now(), "dd.mm.yy")
'    table!newDateRS = tbDateRS.Text
'    table!newDateOut = tbReadyDate.Text
'    table.Update
End If
'End If
'table.Close
'******************************************************************

If getSystemField("resursLock") = Orders.cbM.Text Then unLockBase 'Если именно мы блокирова

wrkDefault.CommitTrans

'обновить Окно Orders
sql = "SELECT Orders.StatusId, GuideProblem.Problem, Orders.DateRS, " & _
"Orders.outDateTime, Orders.workTime, Orders.numOrder, OrdersMO.DateTimeMO, " & _
"OrdersMO.StatM, OrdersMO.StatO, " & _
"OrdersMO.workTimeMO FROM (GuideStatus INNER JOIN (GuideProblem " & _
"INNER JOIN Orders ON GuideProblem.ProblemId = Orders.ProblemId) ON " & _
"GuideStatus.StatusId = Orders.StatusId) LEFT JOIN OrdersMO ON " & _
"Orders.numOrder = OrdersMO.numOrder WHERE (((Orders.numOrder)=" & gNzak & "));"

Set tqOrders = myOpenRecordSet("##16", sql, dbOpenForwardOnly) ', dbDenyWrite)
'If tqOrders Is Nothing Then Exit Sub
str = StatParamsLoad(Orders.mousRow)
tqOrders.Close

On Error Resume Next ' в некот.ситуациях один из Open logFile дает Err: файл уже открыт
Open logFile For Append As #2
Print #2, str
Close #2

Unload Me
Exit Sub

ER1:
wrkDefault.Rollback
On Error Resume Next
'table.Close
'tbOrders.Close
'tqOrders.Close
tbOrders.Close
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmRepit_Click()
workChange = False
    cmAdd.Enabled = False
    tbReadyDate.Enabled = True
    tbWorkTime.Enabled = True
    tbReadyDate.SetFocus
    Orders.startParams
Timer1.Enabled = False
If getSystemField("resursLock") = Orders.cbM.Text Then unLockBase 'Если именно мы блокирова
be_cmRepit = True
laMess.Caption = ""
End Sub

Private Sub cmStatus_Click()

End Sub

Function getNextDayRes(tmpDay As Integer) As Integer
Dim I As Integer

getNextDayRes = maxDay
If tmpDay = maxDay Then Exit Function
I = tmpDay + 1
While nomRes(I) = 0
    If I = maxDay Then Exit Function
    I = I + 1
Wend
If I = maxDay Then Exit Function
getNextDayRes = I
End Function

Function getPrevDayRes(ByVal iDay As Integer) As Integer
Dim I As Integer


If iDay < 2 Then GoTo EN1

While iDay > 1
    iDay = iDay - 1
    If nomRes(iDay) > 0 Then GoTo EN2
Wend
If iDay > 0 Then GoTo EN2
EN1:
For iDay = 1 To maxDay '
    If nomRes(iDay) > 0 Then Exit For
Next iDay
EN2:
getPrevDayRes = iDay
End Function
Function getPrev2DayRes_(ByVal iDay As Integer) As Integer
Dim I As Integer


If iDay < 3 Then GoTo EN1

While iDay > 1
    iDay = iDay - 1
    If nomRes(iDay) > 0 Then GoTo EN0
Wend
EN0:
While iDay > 1
    iDay = iDay - 1
    If nomRes(iDay) > 0 Then GoTo EN2
Wend
If iDay > 0 Then GoTo EN2
EN1:
For iDay = 1 To maxDay '
    If nomRes(iDay) > 0 Then Exit For
Next iDay
EN2:
getPrev2DayRes_ = iDay
End Function

Function getPrev2DayRes(tmpDay As Integer) As Integer
Dim I As Integer
getPrev2DayRes = 1
If tmpDay < 2 Then Exit Function

I = tmpDay - 1
While nomRes(I) = 0
    If I < 2 Then Exit Function
    I = I - 1
Wend
If I < 2 Then Exit Function
I = I - 1
While nomRes(I) = 0
    If I < 2 Then Exit Function
    I = I - 1
Wend
If I < 2 Then Exit Function
getPrev2DayRes = I
End Function

Function getPrev2Day(tmpDay As Integer) As Integer
getPrev2Day = tmpDay - 1
day = Weekday(DateAdd("d", getPrev2Day - 1, curDate))
While day = vbSaturday Or day = vbSunday
    getPrev2Day = getPrev2Day - 1
    day = Weekday(DateAdd("d", getPrev2Day - 1, curDate))
Wend

getPrev2Day = getPrev2Day - 1
day = Weekday(DateAdd("d", getPrev2Day - 1, curDate))
While day = vbSaturday Or day = vbSunday
    getPrev2Day = getPrev2Day - 1
    day = Weekday(DateAdd("d", getPrev2Day - 1, curDate))
Wend

If getPrev2Day < 1 Then getPrev2Day = 1
End Function

Private Sub cmZapros_Click() ' zagruzFromCeh использует глобальные end(beg)Day(MO)
Dim I As Integer, str As String, num As Integer, v As Variant
Dim begDay As Integer, endDay As Integer, begDayMO As Integer, endDayMO As Integer
Dim begDay_ As Integer, endDay_ As Integer ', begDayMO_ As Integer, endDayMO_ As Integer
Dim title As String, msg As String

cmZapros.Enabled = True
cmAdd.Enabled = False
laMess.Caption = ""
isTimeZakaz = True
perenos = 0
I = statId(cbStatus.ListIndex)

If I = 7 Then ' аннулир.
    If Not Orders.do_Annul("no_Do") Then Exit Sub
    GoTo BB
ElseIf I = 0 Then  ' принят  (готов и закрыт здесь не м.б.)
    ' не освежаем данные
BB: isTimeZakaz = False
    
    For I = 1 To lv.ListItems.count
        lv.ListItems("k" & I).SubItems(zkMost) = lv.ListItems("k" & I).SubItems(zkMbef)
       lv.ListItems("k" & I).ListSubItems(zkMost).Bold = False
        lv.ListItems("k" & I).ListSubItems(zkMost).ForeColor = 0
    Next I
    cmAdd.Enabled = True
    Exit Sub
End If

If Not isNumericTbox(tbWorkTime, 0, 2000) Then Exit Sub
tbWorkTime.Text = Round(tbWorkTime.Text, 1)
If Not isDateTbox(tbReadyDate, "fri") Then Exit Sub

tmpDate = CDate(tbReadyDate.Text)
endDay = DateDiff("d", curDate, tmpDate) + 1

maxDay = 0     'добавляем дни, т.к. Дата Выд тек.заказа может оказаться
addDays endDay '1: дальше чем всех других, либо чем stDay и rMaxDay

If endDay < 1 Then
ErrDate: MsgBox "Одна из дат уже в прошлом.", , "Недопустимое значение"
        Exit Sub
End If
If endDay > 100 Then _
    If MsgBox("Получается дата выдачи  через " & endDay & " дней. " & _
        "Подтверждаете?", vbYesNo, "Внимание!!!") = vbNo Then Exit Sub
        
If tbDateRS.Enabled = True Then
    If Not isDateTbox(tbDateRS, "fri") Then Exit Sub
    tmpDate = CDate(tbDateRS.Text)
    begDay = DateDiff("d", curDate, tmpDate) + 1
    If begDay < 1 Then GoTo ErrDate
                
    If begDay > endDay Then
        MsgBox "Дата Р\С не может быть позже Даты выдачи", , "Недопустимая дата"
        Exit Sub
    End If
                
    str = "Между Датой Р\С и Датой выдачи должен быть по крайней " & _
        "мере два рабочих дня!" & Chr(13) & "Иначе при дальнейшем переводе заказа " & _
        "в работу, он станет срочным." & Chr(13) & Chr(13) & "Если вы уверены, что " & _
            "это нарушение не затруднит выполнение Заказов, нажмите - <Да>"
    sql = "Нарушен Коридор:"
Else ' " в работу
    begDay = 1
    str = "Вы задали Срочный заказ. Подтверждаете?"
    sql = "Внимание!!!"
End If
begDay_ = begDay
endDay_ = getPrev2Day(endDay)
endDay = getPrev2DayRes(endDay)

urgent = ""
If endDay_ <= begDay_ Then
    If MsgBox(str, vbYesNo, sql) = vbNo Then Exit Sub
    urgent = "y"
End If
begDay = getNextDayRes(begDay)
If endDay < begDay Then begDay = begDay_ 'сначала возвращаем begDay
If endDay < begDay Then endDay = begDay  'если не помогло, то откат endDay

'******** перенос Даты RS ***********************************
If tbDateRS.Enabled = True Then
If IsDate(Orders.Grid.TextMatrix(Orders.mousRow, orDataRS)) Then ' если Дата РС
tmpDate = Orders.Grid.TextMatrix(Orders.mousRow, orDataRS)       ' поменялась
If DateDiff("d", tmpDate, tbDateRS.Text) <> 0 Then        '

tmpDate = Orders.Grid.TextMatrix(Orders.mousRow, orData)         ' и Сегодня не
If DateDiff("d", tmpDate, curDate) > 0 Then               ' день приема заказа
  title = "Перенос № 1  Подтвеждаете?"
  str = "Всего допустимо только 2 переноса Даты РС (и даты выдачи)." & _
  Chr(10) & "На 3-й раз необходимо аннулировать заказ!" & Chr(10)
  msg = str & Chr(10) & "Если перенос еще допустим нажмите <Да>"
  
  sql = "SELECT ReplaceRS.newDateIn, ReplaceRS.newDateRS, ReplaceRS.newDateOut " & _
  "From ReplaceRS  Where (((ReplaceRS.numOrder) = " & gNzak & ")) " & _
  "ORDER BY ReplaceRS.newDateIn;"
  
  Set table = myOpenRecordSet("##22", sql, dbOpenDynaset) 'dbOpenTable)
  If Not table Is Nothing Then
    If table.BOF Then ' заказа пока нет в ReplaceRS
      If MsgBox(msg, vbYesNo, title) = vbNo Then Exit Sub
         perenos = 1
    Else
      table.MoveFirst: I = 0
      While Not table.EOF
        I = I + 1
        table.MoveNext
      Wend
      table.MoveLast
      If DateDiff("d", table!newDateIn, curDate) > 0 Then ' Дата РС изменилась
         str = I                                      ' первый раз за
         Mid(title, 11) = str                             ' за сегодня
         If MsgBox(msg, vbYesNo, title) = vbNo Then Exit Sub
         perenos = 2
      Else
         title = "Перенос № " & I - 1
         MsgBox str, , title
         perenos = 3
      End If
    End If 'Table.BOF
    table.Close
  End If 'Not Table Is Nothing
End If ' и Сегодня не день приема заказа

End If ' если Дата РС
End If ' поменялась
End If 'tbDateRS.Enabled = True
'*********************************************************
If cbStatus.Text = "согласов" Then
    title = "Недопустимый статус МО"
    If (cbM.Text = "в работе" Or cbM.Text = "готов") And _
    (cbO.Text = "в работе" Or cbO.Text = "готов") Then
        MsgBox "Макет и образец не могут одновременно быть переданы в цех", , title
        Exit Sub
    ElseIf cbM.Text = "" And cbO.Text = "" Then
        MsgBox "Для заказа 'согласование' необходимо установить статус макета и(или) образца", , title
        Exit Sub
    End If
ElseIf cbStatus.Text = "отложен" Then
    GoTo EE
ElseIf cbStatus.Text = "в работе" Then
    If ((cbO.Text <> "" And cbO.Text <> "утвержден") _
    Or (cbM.Text <> "" And cbM.Text <> "утвержден")) And FormIsActiv Then
        MsgBox "Для перевода заказа в работу необходимо, чтобы макет " & _
        "и(или) образец были утверждены.", , "Недопустимый статус!"
        cbStatus.Text = "согласов"
        Exit Sub
    Else
EE:     tbDateRS.Text = ""
    GoTo DD
    End If
Else
DD: cbM.ListIndex = 0
    cbO.ListIndex = 0
    tbVrVipO.Text = ""
    tbDateMO.Text = ""
End If

endDayMO = 0 ' номер дня MO
begDayMO = 0
If cbM.Text = "в работе" Then GoTo AA  'Макет
If cbO.Text = "в работе" Then          'образец
    If Not isNumericTbox(tbVrVipO, 0.1, 2000) Then Exit Sub
    tbVrVipO.Text = Round(tbVrVipO, 1)
AA:
    If Not isDateTbox(tbDateMO, "fri") Then Exit Sub
    tmpDate = CDate(tbDateMO.Text)
    endDayMO = DateDiff("d", curDate, tmpDate) + 1
    If endDayMO < 1 Then GoTo ErrDate
    If endDayMO > begDay_ Then ' не подправленное
        MsgBox "Дата Mак.\Обр. не может быть позже Даты Р\С"
        Exit Sub
    End If
    endDayMO = getPrev2DayRes(endDayMO)
    begDayMO = 1
    I = getNextDayRes(begDayMO)
    If I <= endDayMO Then begDayMO = I
    If endDayMO < begDayMO Then endDayMO = begDayMO
End If

If endDayMO - begDayMO + endDay - begDay > 40 Then
    MsgBox "Заказ сильно растянут, что превышает возможности системы. " & _
    "Если такой интервал действительно необходим, сообщите администратору!" _
      , , "Система не может разместить этот Заказ!"
    Exit Sub
End If

wrkDefault.BeginTrans
myBase.Execute ("update system set resursLock = resursLock")

sql = "select * from System"
'Set tbSystem = myOpenRecordSet("##94", sql, dbOpenForwardOnly)
'If tbSystem Is Nothing Then myBase.Close: End
'tbSystem.Edit
I = 0
     be_cmRepit = False
      str = getSystemField("resursLock")
'     str = tbSystem!resursLock
     If str = "nextDay" Then
'        tbSystem.Update
        wrkDefault.Rollback
        MsgBox "Обнаружено, что был сбой при переводе базы на новый день. " & _
        "Сообщите Администратору или Мастеру Цеха, чтобы он произвел Сброс и " & _
        "переустановку ресурсов в Цехах.", , _
        "Доступ к ресурсам заблокирован!"
        GoTo CC
     End If
     While str <> "" And str <> Orders.cbM.Text
'        tbSystem.Update
        wrkDefault.Rollback
        cmZapros.Enabled = False
        laMess.ForeColor = 200
        laMess.Caption = I & " сек: Доступ к ресурсам временно занят " & _
        "менеджером " & Chr(34) & str & Chr(34) & Chr(13) _
        & Chr(10) & ". Ждите."
        delay (1)
        I = I + 1
        If be_cmRepit Then
            cmZapros.Enabled = True
CC:         'tbSystem.Close
            Exit Sub
        End If
        wrkDefault.BeginTrans
        myBase.Execute ("update system set resursLock = resursLock")
'        tbSystem.Edit
        str = getSystemField("resursLock")
        'str = tbSystem!resursLock
     Wend
     cmZapros.Enabled = True
     myBase.Execute ("update system set resursLock = '" & Orders.cbM.Text & "'")
'tbSystem!resursLock = Orders.cbM.Text
'tbSystem.Update
wrkDefault.CommitTrans
'tbSystem.Close
laMess.Caption = ""

zagruzFromCeh gNzak ' в delta(), Ostatki()  !!!кроме gNzak

tmpMaxDay = getResurs ' выч-е nomRes()
Zakaz.lvAddDays tmpMaxDay 'удаляем или добавляем последние строки(дни) в
'таблице загрузки т.к. Менеджер м. пробывать разные даты выдачи
    
For I = 1 To tmpMaxDay
    lv.ListItems("k" & I).SubItems(zkResurs) = Round(nomRes(I) * kpd * Nstan, 1)
Next I

newZagruz

v = lv.ListItems("k1").SubItems(zkMost)
If Not IsNumeric(v) Then v = 0
I = getNextDay(1)
laZapas.Caption = Round(nomRes(I) * kpd * Nstan + v, 1)

If cmRepit.Visible Then '  не по <F1> <F2>
    tiki = 11
    cmAdd.Enabled = True
    Timer1.Interval = 1 ' перпвый вход сразу
    Timer1.Enabled = True
End If

End Sub

Private Sub Form_Activate()
FormIsActiv = True
End Sub

Private Sub Form_Load()
Dim I As Integer, str As String
FormIsActiv = False
be_cmRepit = True
workChange = False
oldHeight = Me.Height
oldWidth = Me.Width

lv.ColumnHeaders(zkHide + 1).Width = 0

End Sub


Private Sub opM_Click()
End Sub

Private Sub opO_Click()
End Sub

Private Sub tbNomZak_Change()

End Sub

Private Sub tbDateRS_GotFocus()
If FormIsActiv Then Zakaz.cmZapros.Enabled = True
tbDateRS.SelStart = 0
tbDateRS.SelLength = 2

End Sub

Private Sub tbReadyDate_GotFocus()
If FormIsActiv Then Zakaz.cmZapros.Enabled = True
tbReadyDate.SelStart = 0
tbReadyDate.SelLength = 2

End Sub

Private Sub tbReadyDate_KeyDown(KeyCode As Integer, Shift As Integer)
Dim s As Double, I As Integer
If KeyCode = vbKeyReturn Then

If tbDateRS.Enabled Then
  If isDateTbox(tbReadyDate, "fri") Then
    s = Round(CDbl(tbWorkTime.Text), 1)
    I = -(Int((CDbl(s) - 0.05) / 3) + 1 + 2) ' + 2 - дата выд от посл. куска
    getWorkDay I, tbReadyDate.Text ' дает tmpDate
    If tmpDate < curDate Then tmpDate = curDate
    tbDateRS.Text = Format(tmpDate, "dd.mm.yy")
  End If
End If

End If

End Sub

Private Sub tbVrVipO_Change()
If FormIsActiv Then Zakaz.cmZapros.Enabled = True
End Sub

Private Sub tbWorkTime_Change()
If FormIsActiv Then
    Zakaz.cmZapros.Enabled = True
    workChange = True
End If
End Sub

Private Sub tbWorkTime_KeyDown(KeyCode As Integer, Shift As Integer)
Dim s As Double, I As Integer

If KeyCode = vbKeyReturn Then

  If isNumericTbox(tbWorkTime, 0, 2000) Then
     If cbStatus.Text = "в работе" Then
        s = Round(CDbl(tbWorkTime.Text), 1)
        tbWorkTime.Text = s
        I = Int((CDbl(s) - 0.05) / 3)
        getWorkDay 3 + I ' дает tmpDate
        tbReadyDate.Text = Format(tmpDate, "dd.mm.yy")
     Else
        tbReadyDate.Text = "00." & Format(tmpDate, "mm.yy")
     End If
  End If
End If

End Sub

Private Sub Timer1_Timer()
tiki = tiki - 1
If tiki > 0 Then
    laMess.ForeColor = 0
    laMess.Caption = "Для нажатия на кнопку <Ok>" & Chr(13) & Chr(10) & _
    "у Вас осталось несколько секунд: " & tiki
    Timer1.Interval = 1000 ' 1c
Else
    Timer1.Enabled = False
    laMess.Caption = ""
    cmAdd.Enabled = False
    unLockBase
End If
End Sub

