VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Zakaz 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Перемещение заказа в цеховую сводку"
   ClientHeight    =   5892
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9468
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Zakaz.frx":0000
   ScaleHeight     =   5892
   ScaleWidth      =   9468
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmCeh 
      Caption         =   "YAG"
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   29
      Top             =   5520
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CheckBox ckCeh 
      BackColor       =   &H8000000A&
      Caption         =   "Цех"
      Height          =   255
      Left            =   6720
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   675
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4572
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Visible         =   0   'False
      Width           =   6312
      _ExtentX        =   11134
      _ExtentY        =   8065
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.ComboBox cbO 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Zakaz.frx":0342
      Left            =   8160
      List            =   "Zakaz.frx":034F
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2940
      Width           =   1035
   End
   Begin VB.ComboBox cbMaket 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Zakaz.frx":0368
      Left            =   6660
      List            =   "Zakaz.frx":0375
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
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton cmExit 
      BackColor       =   &H8000000A&
      Cancel          =   -1  'True
      Caption         =   "Выход"
      Height          =   315
      Left            =   8340
      TabIndex        =   12
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmRepit 
      BackColor       =   &H8000000A&
      Caption         =   "Cancel"
      Height          =   315
      Left            =   8340
      TabIndex        =   11
      Top             =   4740
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmZapros 
      BackColor       =   &H8000000A&
      Caption         =   "Запросить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7560
      TabIndex        =   9
      Top             =   4260
      Width           =   975
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4512
      Left            =   60
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   360
      Width           =   6372
      _ExtentX        =   11240
      _ExtentY        =   7959
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
      BackColor       =   &H8000000A&
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6780
      TabIndex        =   10
      Top             =   4740
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox tbWorktime 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   2
      Top             =   1320
      Width           =   915
   End
   Begin VB.CheckBox ckCehDone 
      BackColor       =   &H8000000A&
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   5520
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label laStatusId 
      BackColor       =   &H8000000A&
      Caption         =   "Статус заказа"
      Height          =   432
      Left            =   3240
      TabIndex        =   33
      Top             =   5040
      Width           =   732
   End
   Begin VB.Label laStatusText 
      BackColor       =   &H8000000A&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   4080
      TabIndex        =   32
      Top             =   5040
      Width           =   1692
   End
   Begin VB.Label lbEquip 
      BackColor       =   &H8000000A&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   1200
      TabIndex        =   31
      Top             =   5040
      Width           =   1932
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Обору- дование"
      Height          =   432
      Left            =   240
      TabIndex        =   27
      Top             =   5040
      Width           =   852
   End
   Begin VB.Label laMO 
      BackColor       =   &H8000000A&
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
      BackColor       =   &H8000000A&
      Caption         =   "Запас:"
      Height          =   195
      Left            =   4680
      TabIndex        =   23
      Top             =   60
      Width           =   675
   End
   Begin VB.Label laError 
      BackColor       =   &H8000000A&
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
      Left            =   1080
      TabIndex        =   22
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label laNomZak 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   0
      Top             =   330
      Width           =   915
   End
   Begin VB.Label laVrVipO 
      BackColor       =   &H8000000A&
      Caption         =   "Вр.вып. Образца"
      Enabled         =   0   'False
      Height          =   195
      Left            =   6540
      TabIndex        =   21
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label laDateMO 
      BackColor       =   &H8000000A&
      Caption         =   "Дата Мак.\Обр."
      Enabled         =   0   'False
      Height          =   195
      Left            =   6540
      TabIndex        =   20
      Top             =   3420
      Width           =   1275
   End
   Begin VB.Label laHnomZak 
      BackColor       =   &H8000000A&
      Caption         =   "№ Заказа"
      Height          =   195
      Left            =   6540
      TabIndex        =   19
      Top             =   360
      Width           =   855
   End
   Begin VB.Label laStatus 
      BackColor       =   &H8000000A&
      Caption         =   "Статус:"
      Height          =   195
      Left            =   6540
      TabIndex        =   18
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label laDateRS 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Дата Р\С (не позже)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   6540
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label laReadyDate 
      BackColor       =   &H8000000A&
      Caption         =   "Дата выдачи"
      Height          =   195
      Left            =   6540
      TabIndex        =   15
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label laWorkTime 
      BackColor       =   &H8000000A&
      Caption         =   "Время выполнения"
      Height          =   255
      Left            =   6540
      TabIndex        =   14
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Label laMess 
      BackColor       =   &H8000000A&
      Height          =   732
      Left            =   5820
      TabIndex        =   16
      Top             =   5160
      Width           =   2472
   End
End
Attribute VB_Name = "Zakaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public urgent As Variant ' As String '"y" - срочный заказ
Public Regim As String
Public idEquip As Integer
Public isUpdated As Boolean  ' выставлен в true если что-то в статусе заказа действительно поменялось

Dim neVipolnen As Double, neVipolnen_O As Double ' в часах, сколько еще не выполнено

' M125 - таким будет статус заказа
Public festStatusId As Integer


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

' Это должны быть статусы по оборудованию, не по общему по заказу!!!
Dim statusIdNew As Integer, statusIdOld As Integer
Dim zakazBean As ZakazVO
Dim someWasChanged As Boolean ' true если хотя бы раз был нажат ОК



Const zgNomZak = 1
Const zgStatus = 2
Const zgVrVip = 3
Const zgNevip = 4
Const zgInDate = 5
Const zgOutDate = 6
Const zgOtlad = 7



Sub lvAddDay(I As Integer)

Dim Item As ListItem, str As String
    str = Format(DateAdd("d", I - 1, curDate), "dd/mm/yy")
    Set Item = lv.ListItems.Add(, "k" & I, str)
    day = Weekday(DateAdd("d", I - 1, curDate))
    If day = vbSunday Or day = vbSaturday Then Item.ForeColor = &HFF
End Sub

Sub lvAddDays(Optional newLen As Integer = -1)
Dim I As Integer, J As Integer

If newLen < 0 Then newLen = maxDay

J = lv.ListItems.Count
If newLen > J Then ' j=0 когда startParams вызыв-ся первый раз
    For I = J + 1 To newLen
        lvAddDay I
    Next I
ElseIf newLen < J Then
    For I = newLen + 1 To J
        lv.ListItems.Remove "k" & I
    Next I
    End If
End Sub

Private Sub cbMaket_Click()
'cmZapros.Enabled = True
If cbMaket.Text = "в работе" Or cbMaket.Text = "готов" Then
    If FormIsActiv Then
        cmZapros.Enabled = True
    End If
    laDateMO.Enabled = True
    tbDateMO.Enabled = True
ElseIf Not (cbO.Text = "в работе" Or cbO.Text = "готов") Then
    laDateMO.Enabled = False
    tbDateMO.Enabled = False
    tbDateMO.Text = ""
End If
End Sub

Private Sub cbO_Click()
'cmZapros.Enabled = True
If cbO.Text = "в работе" Or cbO.Text = "готов" Then
    If FormIsActiv Then
        cmZapros.Enabled = True
    End If
    laDateMO.Enabled = True
    tbDateMO.Enabled = True
    laVrVipO.Enabled = True
    tbVrVipO.Enabled = True
    If Not IsNull(zakazBean.WorktimeMO) Then
        tbVrVipO.Text = zakazBean.WorktimeMO
    Else
        tbVrVipO.Text = ""
    End If
Else
    If Not (cbMaket.Text = "в работе" Or cbMaket.Text = "готов") Then
        laDateMO.Enabled = False
        tbDateMO.Enabled = False
        tbDateMO.Text = ""
    End If
    laVrVipO.Enabled = False
    tbVrVipO.Enabled = False
    tbVrVipO.Text = ""
End If
End Sub

Private Sub cmCeh_Click(Index As Integer)
    idEquip = Index + 1
    gEquipId = idEquip
    'statusIdOld = statusIdNew
    startParams
    'newZagruz ' вызывается в startParams (!?)
End Sub

Sub getBegEndDays(Optional Status As String = "")
Dim priemData As String

If Regim = "setka" Then ' по F1 F2
    priemData = Format(curDate, "dd.mm.yy")
Else
    priemData = Orders.Grid.TextMatrix(Orders.Grid.row, orData)
End If
If Status = "образец" Then
    Grid.TextMatrix(quantity, zgInDate) = priemData
    Grid.TextMatrix(quantity, zgOutDate) = tbDateMO.Text
    Grid.TextMatrix(quantity, zgOtlad) = ObDay & " " & OeDay
    Exit Sub
ElseIf Status <> "" Then
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
If tbVrVipO.Enabled And tbDateMO.Text <> "" Then
    ObDay = DateDiff("d", curDate, priemData)
    OeDay = DateDiff("d", curDate, tbDateMO.Text)
End If

End Sub
' заказ в реестр отладки
Function zakazToGrid(reg As String, Status As String, nevip As String) As Boolean
zakazToGrid = False

If reg = "" Then
    quantity = quantity + 1
    Grid.TextMatrix(quantity, zgNomZak) = laNomZak.Caption
    Grid.TextMatrix(quantity, zgStatus) = Status
    If Status = "образец" Then
        Grid.TextMatrix(quantity, zgVrVip) = tbVrVipO.Text
    Else
        Grid.TextMatrix(quantity, zgVrVip) = tbWorktime.Text
    End If
    Grid.TextMatrix(quantity, zgNevip) = nevip
    
    getBegEndDays Status
    
    Grid.AddItem ""
Else
    If laNomZak.Caption = Grid.TextMatrix(Grid.row, zgNomZak) Then
        If Grid.TextMatrix(Grid.row, zgStatus) = "образец" Then
            If Status = "образец" Then Exit Function
        ElseIf Status <> "образец" Then
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
'reg = "" - double click at status cell

Sub newZagruz(Optional reg As String = "", Optional EquipId As Integer)
Dim S As Double, nevip As Double, I As Integer
Dim bDay As Integer, eDay As Integer, vEquipId As Integer

If EquipId <> 0 Then
    vEquipId = EquipId
Else
    vEquipId = Me.idEquip
End If

'isMzagruz - true: если вызвали загрузку НЕ из цеха, то есть Менеджер.
If reg = "fromCehZagruz" Then
    isMzagruz = False
Else
    isMzagruz = (ckCeh.Value = 0)
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
firstRes = nr * Nstan * KPD             '
ost(1) = firstRes
befOst(1) = firstRes
For I = 2 To maxDay
    ost(I) = nomRes(I) * KPD * Nstan
    befOst(I) = nomRes(I) * KPD * Nstan
Next I

'VrVipParts заменнили на Nevip
sql = "SELECT o.numOrder, oe.workTime, " & _
" DateDiff(day,Now(),oe.outDateTime) AS endDay, " & _
" DateDiff(day,Now(),o.inDate) AS begDay, dateadd(hour, isnull(o.outtime, 12), oe.outDateTime) as outdatetime, " & _
" o.inDate, o.StatusId, isnull(oe.nevip, 1) as nevip, oc.urgent " & _
vbCr & " FROM Orders o " & _
" JOIN OrdersEquip oe ON oe.numorder = o.numorder AND oe.equipId = " & vEquipId & _
" JOIN OrdersInCeh oc ON o.numOrder = oc.numOrder " & _
" Where (o.StatusId = 1 Or o.StatusId = 5) " & _
vbCr & " UNION ALL " _
& vbCr & " SELECT o.numOrder, oe.workTime, DateDiff(day,Now(),oe.outDateTime) AS endDay, " & _
" DateDiff(day,Now(),isnull(o.DateRS, oe.outdatetime)) AS begDay, dateadd(hour, isnull(o.outtime, 12), oe.outDateTime) as outdatetime, " & _
" o.DateRS, o.StatusId, isnull(oe.nevip, 1) as nevip, oc.urgent " & _
vbCr & " FROM Orders o " & _
" JOIN OrdersEquip oe ON oe.numorder = o.numorder AND oe.equipId = " & vEquipId & _
" JOIN OrdersInCeh oc ON o.numOrder = oc.numOrder " & _
" Where (o.StatusId = 2 Or o.StatusId = 3) " & _
vbCr & " UNION ALL " _
& vbCr & " SELECT o.numOrder, oe.workTimeMO, DateDiff(day,Now(),oc.DateTimeMO) AS endDay, " & _
" DateDiff(day,Now(),o.inDate) AS begDay, dateadd(hour, isnull(o.outtime, 12), oc.DateTimeMO) as outdatetime, " & _
" o.inDate, 1 AS StatusId, -1 AS Nevip, '' AS urgent " & _
vbCr & " FROM Orders o " & _
" JOIN OrdersInCeh oc ON o.numOrder = oc.numOrder " & _
" JOIN OrdersEquip oe ON oe.numorder = o.numorder AND oe.equipId = " & vEquipId & _
" Where oe.statO = 'в работе' " & " ORDER BY "

If isMzagruz Then
    sql = sql & "4 DESC" ' в порядке уменьшения Даты Начала
Else
    sql = sql & "3" ' в порядке увеличения  Даты Конца
End If
'Debug.Print sql
Set tbOrders = myOpenRecordSet("##370", sql, dbOpenForwardOnly) ', dbOpenDynaset)
If tbOrders Is Nothing Then Exit Sub
While Not tbOrders.EOF
    bDay = tbOrders!begDay '  отн. Now()
    If Not IsNull(tbOrders!endDay) Then
        eDay = tbOrders!endDay '  отн. Now()
    Else
        eDay = 0
    End If
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
            ObDay = -32000 ' чтобы более не срабатывал
        End If
'        If eDay > ZeDay Or (eDay = ZeDay And bDay <= ZbDay) Then ' не нарушаем сортировку
        If bDay <= ZbDay Then ' не нарушаем сортировку
            dayCorrect ZbDay, ZeDay, CStr(Me.urgent)
            ukladka ost, ZeDay, ZbDay, tbWorktime.Text 'обратная укладка (в bef не попадает)
            If Not zakazToGrid(reg, cbStatus.Text, tbWorktime.Text) Then GoTo EN1
            ZbDay = -32000 ' чтобы более не срабатывал
        End If
    End If
    
    If eDay > maxDay Then
        msgOfEnd "##371", "Заказ " & CStr(tbOrders!Numorder) & vbCr & "Неверная дата!"
    End If

'    dayCorrect bDay, eDay, tbOrders!urgent спец.распределение срочн. заказов приводит к тому, что м.б. разные минусы у М и в цеху
    dayCorrect bDay, eDay, ""
    
    If tbOrders!nevip = -1 Then '"образец"
        nevip = tbOrders!Worktime
    Else
        nevip = Round(tbOrders!Worktime * tbOrders!nevip, 2)
    End If
    If isMzagruz Then
        ukladka ost, eDay, bDay, nevip 'обратная укладка
        ukladka befOst, eDay, bDay, nevip 'обратная укладка
    Else
        ukladka ost, bDay, eDay, nevip
        If tbOrders!StatusId = 1 Or tbOrders!StatusId = -1 Then _
            ukladka befOst, bDay, eDay, nevip ' живые(в раб. и образец)
    End If
    
    If reg = "" Then
      quantity = quantity + 1
      Grid.TextMatrix(quantity, zgNomZak) = tbOrders!Numorder
'      If tbOrders!StatusId = -1 Then
      If tbOrders!nevip = -1 Then '"образец"
        Grid.TextMatrix(quantity, zgStatus) = "образец"
      Else
        Grid.TextMatrix(quantity, zgStatus) = Status(tbOrders!StatusId)
      End If
      Grid.TextMatrix(quantity, zgVrVip) = tbOrders!Worktime
      Grid.TextMatrix(quantity, zgNevip) = nevip
      Grid.TextMatrix(quantity, zgInDate) = Format(tbOrders!inDate, "dd.mm.yy")
      Grid.TextMatrix(quantity, zgOutDate) = Format(tbOrders!Outdatetime, "dd.mm.yy")
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
    ukladka ost, ZeDay, ZbDay, tbWorktime.Text 'обратная укладка (в bef не попадает)
    zakazToGrid reg, cbStatus.Text, tbWorktime.Text
  End If
End If

If reg = "" And quantity > 0 Then Grid.RemoveItem Grid.Rows - 1
EN1:
tbOrders.Close

If reg = "fromCehZagruz" Then Exit Sub

If ckCeh.Value = 0 Then
  For I = 1 To maxDay
    lv.ListItems("k" & I).SubItems(zkMbef) = Round(befOst(I), 1) '23.11.04
    lv.ListItems("k" & I).SubItems(zkMzagr) = _
               Round(nomRes(I) * KPD * Nstan - befOst(I), 1)
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
            Round(nomRes(I) * KPD * Nstan - befOst(I), 1)
   lv.ListItems("k" & I).SubItems(zkCzagr) = _
            Round(nomRes(I) * KPD * Nstan - ost(I), 1)
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
    ' cmNewUklad.Visible = True
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
Dim H As Integer, W As Integer

If Me.WindowState = vbMinimized Then Exit Sub
'lbHide

On Error Resume Next
H = Me.Height - oldHeight
oldHeight = Me.Height
W = Me.Width - oldWidth
oldWidth = Me.Width

'lv.Height = lv.Height + h
lv.Width = lv.Width + W
laHnomZak.Left = laHnomZak.Left + W
laNomZak.Left = laNomZak.Left + W
laStatus.Left = laStatus.Left + W
cbStatus.Left = cbStatus.Left + W
laWorkTime.Left = laWorkTime.Left + W
tbWorktime.Left = tbWorktime.Left + W
laReadyDate.Left = laReadyDate.Left + W
tbReadyDate.Left = tbReadyDate.Left + W
laDateRS.Left = laDateRS.Left + W
tbDateRS.Left = tbDateRS.Left + W
laMO.Left = laMO.Left + W
cbMaket.Left = cbMaket.Left + W
cbO.Left = cbO.Left + W
laDateMO.Left = laDateMO.Left + W
tbDateMO.Left = tbDateMO.Left + W
laVrVipO.Left = laVrVipO.Left + W
tbVrVipO.Left = tbVrVipO.Left + W
cmZapros.Left = cmZapros.Left + W
cmAdd.Left = cmAdd.Left + W
cmRepit.Left = cmRepit.Left + W
cmExit.Left = cmExit.Left + W
cmExit.Top = cmExit.Top + H

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Если именно мы блокировали:
    If getSystemField("resursLock") = Orders.cbM.Text Then unLockBase
    'Orders.Grid_EnterCell ' подсветка ячейки
    
    Unload Equipment

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

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)

End Sub


Private Sub laNomZak_Click()
    Dim Left As String, Worktime As String, tbWorktime As String, Rollback As String, Value
End Sub

Private Sub tbDateMO_GotFocus()
If FormIsActiv Then
    cmZapros.Enabled = True
End If
If tbDateMO.Text = "" Then
    tbDateMO.Text = Format(curDate, "dd/mm/yy")
End If
tbDateMO.SelStart = 0
tbDateMO.SelLength = 2

End Sub

Private Sub cbStatus_Click()

If noClick Then
    Exit Sub
End If
If FormIsActiv Then
    cmZapros.Enabled = True
End If
Dim I As Integer
'Exit Sub
statusIdNew = cbStatus.ItemData(cbStatus.ListIndex)

'cmZapros.Enabled = statusIdOld <> statusIdNew

'For I = 0 To UBound(Equip) - 1
'    If ckCehDone(I).Tag = CStr(statusIdNew) Then
'        ckCehDone(I).Value = 1
'    Else
'        ckCehDone(I).Value = 0
'    End If
'Next I

tbWorktime.Text = zakazBean.Worktime
If Not IsNull(zakazBean.DateRS) Then
    If statusIdNew > 1 Then
        'если не в работе и не принят
        tbDateRS.Text = Format(zakazBean.DateRS, "dd.mm.yy")
    Else
        ' переводим в работу
        tbDateRS.Text = ""
    End If
End If

If Not IsNull(zakazBean.Outdatetime) Then
    tbReadyDate.Text = Format(zakazBean.Outdatetime, "dd.mm.yy")
End If

setTheControls cbStatus.ItemData(cbStatus.ListIndex)

End Sub



Private Sub setTheControls(equipStatusId As Integer)

If equipStatusId = 1 Then 'cbStatus.Text = "в работе" Then
    laMO.Enabled = False
    cbMaket.Enabled = False
    cbO.Enabled = False
    tbDateMO.Enabled = False
    laVrVipO.Enabled = False
    tbVrVipO.Enabled = False
    tbWorktime.SelStart = 0
    tbWorktime.SelLength = Len(tbWorktime.Text)
    'tbWorktime.SetFocus
ElseIf equipStatusId = 3 Then 'cbStatus.Text = "согласов" Then
    cbMaket.Enabled = True
    cbO.Enabled = True
    laMO.Enabled = True
Else
    laMO.Enabled = False
    cbMaket.Enabled = False
    cbO.Enabled = False
    tbDateMO.Enabled = False
    laVrVipO.Enabled = False
    tbVrVipO.Enabled = False
End If

If equipStatusId = 2 Or equipStatusId = 3 Then 'cbStatus.Text = "согласов" Or cbStatus.Text = "резерв" Then
    tbDateRS.Enabled = True             ' резерв согласование
    laDateRS.Enabled = True
    laWorkTime.Enabled = True
    laReadyDate.Enabled = True
    tbReadyDate.Enabled = True
    tbWorktime.Enabled = True
ElseIf equipStatusId = 1 Or equipStatusId = 5 Then ' cbStatus.Text = "в работе" Or cbStatus.Text = "отложен" Then
    tbDateRS.Enabled = False
    laDateRS.Enabled = False
    laWorkTime.Enabled = True
    laReadyDate.Enabled = True
    tbReadyDate.Enabled = True
    tbWorktime.Enabled = True
    If tbWorktime.Enabled And tbWorktime.Visible Then tbWorktime.SetFocus
Else
    laWorkTime.Enabled = False
    laReadyDate.Enabled = False
    tbReadyDate.Text = ""
    tbReadyDate.Enabled = False
    tbWorktime.Text = ""
    tbWorktime.Enabled = False
    tbDateRS.Text = ""
    tbDateRS.Enabled = False
    laDateRS.Enabled = False
End If
End Sub


'$odbc08$
Private Sub cmAdd_Click()
Dim I As Integer, str As String, Item As ListItem, S As Double, T As Double
Dim VrVip As String, VrVipO As String, editWorkTime As Boolean
Dim Worktime As String, hasRecord As Integer
Dim workTimeOld As Double, Otrabotano As Double, nevip As Double

'MaxDay = tmpMaxDay наверно это уже никому не нужно

Timer1.Enabled = False
cmAdd.Enabled = False
laMess.Visible = False


sql = "SELECT 1, oe.worktime, oe.Nevip" _
& " from OrdersEquip oe " _
& " WHERE oe.numOrder = " & gNzak & " and oe.equipId = " & idEquip
byErrSqlGetValues "##30", sql, hasRecord, workTimeOld, nevip

If hasRecord = 0 Then
    If getSystemField("resursLock") = Orders.cbM.Text Then
        unLockBase 'Если именно мы блокировали
    End If
    MsgBox "Возможно он уже удален. Обновите Реестр", , "Заказ не найден"
    Exit Sub
Else
    neVipolnen = nevip * workTimeOld
    Otrabotano = workTimeOld - neVipolnen
End If



wrkDefault.BeginTrans

If statusIdNew <> statusIdOld Or (neVipolnen_O = 0 And neVipolnen = 0) Then
    editWorkTime = False    '
Else                        'если что-то недовып-но и статус не меняется
    editWorkTime = True     'то нужно поменять Вр.Вып.
End If

Dim v_outDateTime As String
v_outDateTime = tbReadyDate.Text

If v_outDateTime <> "" Then
    v_outDateTime = "'" & "20" & Mid$(v_outDateTime, 7, 2) & "-" & Mid$(v_outDateTime, 4, 2) & "-" & Left$(v_outDateTime, 2) & "'"
Else
    v_outDateTime = "Null"
End If




If Not tbDateRS.Enabled And tbDateRS.Text = "" Then
    str = "Null"
Else
    str = tbDateRS.Text
    str = "'" & "20" & Mid$(str, 7, 2) & "-" & Mid$(str, 4, 2) & "-" & Left$(str, 2) & "'"
End If

sql = "UPDATE Orders SET dateRS = " & str & " WHERE Orders.numOrder = " & gNzak
If myExecute("##392", sql) <> 0 Then GoTo ER1

sql = "SELECT * from OrdersInCeh WHERE numOrder = " & gNzak
Set tbOrders = myOpenRecordSet("##01", sql, dbOpenForwardOnly)


Worktime = workTimeOld ' для случая, если не менялось
If Not tbOrders.BOF Then
    If isTimeZakaz Then
        If workChange Then
            If (statusIdNew = 1 Or statusIdNew = 5) And editWorkTime Then 'остается в работе или отложен
               Worktime = Round(tbWorktime.Text, 1) 'время с учетом коррекции
               nevip = (Worktime - Otrabotano) / Worktime
               sql = "UPDATE OrdersEquip SET Nevip = " & nevip _
                & " WHERE numOrder =" & gNzak & " AND equipId = " & idEquip
               If myExecute("##393", sql) <> 0 Then GoTo ER1
            Else
               Worktime = tbWorktime.Text
            End If
        End If
       sql = "UPDATE OrdersInCeh SET urgent = '" & urgent & _
       "' WHERE OrdersInCeh.numOrder = " & gNzak
       If myExecute("##403", sql) <> 0 Then GoTo ER1
       GoTo DD
    Else
        'sql = "select count(*) from vw_Reestr"
    
        sql = "DELETE from OrdersInCeh WHERE numOrder = " & gNzak
        If myExecute("##394", sql) <> 0 Then GoTo ER1
        Worktime = 0
    End If
Else
    If isTimeZakaz Then
        Worktime = tbWorktime.Text
        sql = "INSERT INTO OrdersInCeh ( numOrder, urgent)" & _
        "SELECT " & gNzak & ",'" & urgent & "'"
        If myExecute("##395", sql) <> 0 Then GoTo ER1
DD:     noClick = True
        Orders.Grid.col = orWerk
        If urgent = "y" Then
            Orders.Grid.CellForeColor = 200
        Else
            Orders.Grid.CellForeColor = vbBlack
        End If
        Orders.Grid.col = orStatus
        noClick = False
    End If
End If


sql = "UPDATE OrdersEquip SET outDateTime = " & v_outDateTime _
    & ", workTime = " & Worktime _
    & ", statusEquipId = " & statusIdNew _
    & " WHERE numOrder = " & gNzak & " and equipId =" & idEquip
'Debug.Print sql
If myExecute("##391", sql) <> 0 Then GoTo ER1


If zakazBean.StatusId <> statusIdNew Then
    ' если основной статус заказа поменялся ...
    sql = "UPDATE Orders SET statusId = " & statusIdNew & " WHERE Orders.numOrder =" & gNzak
    If myExecute("##396", sql) <> 0 Then
        GoTo ER1
    Else
        ' ... коррекция посещения фирмой
        If (zakazBean.StatusId = 0 Or zakazBean.StatusId = 7) And statusIdNew <> 0 And statusIdNew <> 7 Then
            visits "+"
        ElseIf Not (zakazBean.StatusId = 0 Or zakazBean.StatusId = 7) And (statusIdNew = 0 Or statusIdNew = 7) Then
            visits "-"
        End If
        ' ... если аннулируемый заказ там есть
        If statusIdNew = 7 Then delZakazFromReplaceRS
    End If
End If


' в согласование или из согласования в работу
sql = "SELECT * from OrdersInCeh WHERE numOrder =" & gNzak
Set Table = myOpenRecordSet("##02", sql, dbOpenForwardOnly)
bilo = Not Table.BOF
Table.Close

 If statusIdNew = 3 Then ' согласов
  If cbMaket.Text = "в работе" Or cbMaket.Text = "готов" Or _
    cbO.Text = "в работе" Or cbO.Text = "готов" Then
    str = tbDateMO.Text
    str = "'" & "20" & Mid$(str, 7, 2) & Mid$(str, 4, 2) & Left$(str, 2)
    sql = Orders.Grid.TextMatrix(Orders.mousRow, orMOVrVid)
    If sql = "" Then
        str = str & "'"
    Else
        str = str & " " & sql & ":00'"
    End If
  Else
    str = "Null"
  End If
  If cbO.Text = "в работе" Or cbO.Text = "готов" Then
    Worktime = tbVrVipO.Text
  Else
    Worktime = "Null"
  End If
  If bilo Then      '
    sql = "UPDATE OrdersInCeh SET StatM = '" & cbMaket.Text & "'" _
    & ", DateTimeMO = " & str & _
    " WHERE numOrder = " & gNzak
  Else
    sql = "INSERT INTO OrdersInCeh ( numOrder, StatM, DatetimeMO ) " & _
    "SELECT " & gNzak & ", '" & cbMaket.Text & "', " & str
  End If
  'Debug.Print sql
  If myExecute("##397", sql) <> 0 Then GoTo ER1
    
  sql = "UPDATE OrdersEquip SET workTimeMO = " & Worktime _
    & ", statO = '" & cbO.Text & "'" _
    & " WHERE numOrder = " & gNzak & " and equipId = " & idEquip
    
  If myExecute("##397.2", sql) <> 0 Then GoTo ER1
 Else
  'Quasi Delete MO info from Equip table ...
  sql = "update OrdersEquip SET " _
  & " nevip = 1, stat = '', WorktimeMO = NULL, statO = NULL " _
  & " WHERE NumOrder = " & gNzak ' & " AND equipId = " & idEquip ' для всех оборудований ...
  If myExecute("##397.3", sql) <> 0 Then GoTo ER1
  ' ... and from InCeh table
  sql = "update OrdersInCeh SET DateTimeMO = NULL, statM = NULL WHERE NumOrder = " & gNzak
  myExecute "##397.4", sql, -1
 End If ' согласов
tbOrders.Close

    
'******** перенос Даты RS ***********************************
If perenos = 1 Then ' был подтвержден перенос РС
    sql = "INSERT INTO ReplaceRS ( numOrder, newDateIn, newDateRS, newDateOut) " & _
    "SELECT " & gNzak & ", '" & _
    yymmdd(Orders.Grid.TextMatrix(Orders.mousRow, orData)) & "', '" & _
    yymmdd(Orders.Grid.TextMatrix(Orders.mousRow, orDataRS)) & "', '" & _
    yymmdd(Orders.Grid.TextMatrix(Orders.mousRow, orDataVid)) & "';"
'    MsgBox sql
    If myExecute("##399", sql) <> 0 Then GoTo ER1
    GoTo СС
  ElseIf perenos = 2 Then ' был подтвержден перенос РС
СС: ' table.AddNew
    sql = "INSERT INTO ReplaceRS ( numOrder, newDateIn, newDateRS, newDateOut) " & _
    "SELECT " & gNzak & ", '" & Format(Now(), "yyyy-mm-dd") & "', '" & _
    yymmdd(tbDateRS.Text) & "', '" & yymmdd(tbReadyDate.Text) & "';"
    If myExecute("##400", sql) <> 0 Then GoTo ER1
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
End If
'******************************************************************

If getSystemField("resursLock") = Orders.cbM.Text Then unLockBase 'Если именно мы блокирова

wrkDefault.CommitTrans

str = Orders.openOrdersRowToGrid("##16")

On Error Resume Next ' в некот.ситуациях один из Open logFile дает Err: файл уже открыт
Open logFile For Append As #2
Print #2, str
Close #2

Dim nextEquipId As Integer
ckCehDone(idEquip - 1).Tag = statusIdNew
 
someWasChanged = True
If Not chooseTheEquipment(statusIdNew, nextEquipId) Then
    ' refresh the Orders.Grid row
    Orders.refreshCurrentRow = True
    Unload Me
Else
    idEquip = nextEquipId
    InitZagruz
    startParams
End If

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
Orders.refreshCurrentRow = someWasChanged
Unload Me
End Sub

Private Sub cmRepit_Click()
workChange = False
    cmAdd.Enabled = False
    tbReadyDate.Enabled = True
    tbWorktime.Enabled = True
    tbReadyDate.SetFocus
    cmZapros.Enabled = False
    startParams
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
Dim I As Integer, str As String, num As Integer, V As Variant
Dim begDay As Integer, endDay As Integer, begDayMO As Integer, endDayMO As Integer
Dim begDay_ As Integer, endDay_ As Integer ', begDayMO_ As Integer, endDayMO_ As Integer
Dim title As String, msg As String

'cmZapros.Enabled = True
cmAdd.Enabled = False
laMess.Caption = ""
laMess.Visible = True
isTimeZakaz = True
perenos = 0
I = cbStatus.ItemData(cbStatus.ListIndex)

If I = 7 Then ' аннулир.
    If Not Orders.do_Annul("no_Do") Then Exit Sub
    GoTo BB
ElseIf I = 0 Then  ' принят  (готов и закрыт здесь не м.б.)
    ' не освежаем данные
BB: isTimeZakaz = False
    
    For I = 1 To lv.ListItems.Count
        lv.ListItems("k" & I).SubItems(zkMost) = lv.ListItems("k" & I).SubItems(zkMbef)
       lv.ListItems("k" & I).ListSubItems(zkMost).Bold = False
        lv.ListItems("k" & I).ListSubItems(zkMost).ForeColor = 0
    Next I
    cmAdd.Enabled = True
    Exit Sub
End If

If Not isNumericTbox(tbWorktime, 0, 2000) Then Exit Sub
tbWorktime.Text = Round(tbWorktime.Text, 1)
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
  
  Set Table = myOpenRecordSet("##22", sql, dbOpenDynaset) 'dbOpenTable)
  If Not Table Is Nothing Then
    If Table.BOF Then ' заказа пока нет в ReplaceRS
      If MsgBox(msg, vbYesNo, title) = vbNo Then Exit Sub
         perenos = 1
    Else
      Table.MoveFirst: I = 0
      While Not Table.EOF
        I = I + 1
        Table.MoveNext
      Wend
      Table.MoveLast
      If DateDiff("d", Table!newDateIn, curDate) > 0 Then ' Дата РС изменилась
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
    Table.Close
  End If 'Not Table Is Nothing
End If ' и Сегодня не день приема заказа

End If ' если Дата РС
End If ' поменялась
End If 'tbDateRS.Enabled = True
'*********************************************************
If cbStatus.Text = "согласов" Then
    title = "Недопустимый статус МО"
    If (cbMaket.Text = "в работе" Or cbMaket.Text = "готов") And _
    (cbO.Text = "в работе" Or cbO.Text = "готов") Then
        MsgBox "Макет и образец не могут одновременно быть переданы в цех", , title
        Exit Sub
    ElseIf cbMaket.Text = "" And cbO.Text = "" Then
        MsgBox "Для заказа 'согласование' необходимо установить статус макета и(или) образца", , title
        Exit Sub
    End If
ElseIf cbStatus.Text = "отложен" Then
    GoTo EE
ElseIf cbStatus.Text = "в работе" Then
    If ((cbO.Text <> "" And cbO.Text <> "утвержден") _
    Or (cbMaket.Text <> "" And cbMaket.Text <> "утвержден")) And FormIsActiv Then
        MsgBox "Для перевода заказа в работу необходимо, чтобы макет " & _
        "и(или) образец были утверждены.", , "Недопустимый статус!"
        cbStatus.Text = "согласов"
        Exit Sub
    Else
EE:     tbDateRS.Text = ""
    GoTo DD
    End If
Else
DD: cbMaket.ListIndex = 0
    cbO.ListIndex = 0
    tbVrVipO.Text = ""
    tbDateMO.Text = ""
End If

endDayMO = 0 ' номер дня MO
begDayMO = 0
If cbMaket.Text = "в работе" Then GoTo AA  'Макет
If cbO.Text = "в работе" Then          'образец
    If Not isNumericTbox(tbVrVipO, 0.1, 2000) Then Exit Sub
    tbVrVipO.Text = Round(tbVrVipO.Text, 1)
AA:
    If Not isDateTbox(tbDateMO, "fri") Then Exit Sub
    'tmpDate = CDate(tbDateMO.Text)
    endDayMO = DateDiff("d", curDate, tmpDate) + 1
    If endDayMO < 1 Then GoTo ErrDate
    'If endDayMO > begDay_ Then ' не подправленное
    '    MsgBox "Дата Mак.\Обр. не может быть позже Даты Р\С"
    '    Exit Sub
    'End If
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

'sql = "select * from System"
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
            'cmZapros.Enabled = True
CC:         'tbSystem.Close
            Exit Sub
        End If
        wrkDefault.BeginTrans
        myBase.Execute ("update system set resursLock = resursLock")
'        tbSystem.Edit
        str = getSystemField("resursLock")
        'str = tbSystem!resursLock
     Wend
     'cmZapros.Enabled = True
     myBase.Execute ("update system set resursLock = '" & Orders.cbM.Text & "'")
'tbSystem!resursLock = Orders.cbM.Text
'tbSystem.Update
wrkDefault.CommitTrans
'tbSystem.Close
laMess.Caption = ""

zagruzFromCeh idEquip, gNzak ' в otstup(), Ostatki()  !!!кроме gNzak

tmpMaxDay = getResurs(idEquip)  ' выч-е nomRes()
lvAddDays tmpMaxDay 'удаляем или добавляем последние строки(дни) в
'таблице загрузки т.к. Менеджер м. пробывать разные даты выдачи
    
For I = 1 To tmpMaxDay
    lv.ListItems("k" & I).SubItems(zkResurs) = Round(nomRes(I) * KPD * Nstan, 1)
Next I

newZagruz

V = lv.ListItems("k1").SubItems(zkMost)
If Not IsNumeric(V) Then V = 0
I = getNextDay(1)
laZapas.Caption = Round(nomRes(I) * KPD * Nstan + V, 1)

If cmRepit.Visible Then '  не по <F1> <F2>
    tiki = 11
    cmAdd.Enabled = True
    Timer1.Interval = 1 ' первый вход сразу
    Timer1.Enabled = True
End If

End Sub

Private Sub Form_Activate()
FormIsActiv = True
End Sub


Private Sub cehSelectorAccess(cehIndex As Integer, action As Boolean, syncStatus As Boolean)
    ckCehDone(cehIndex).Visible = action
    ckCehDone(cehIndex).Value = IIf(syncStatus, 1, 0)
    cmCeh(cehIndex).Visible = action
    cmCeh(cehIndex).Enabled = action
End Sub


Private Sub cehSelectorsInit(action As Boolean)
Dim I As Integer
    For I = 0 To UBound(Equip) - 1
        If Regim = "" Then
            ckCehDone(I).Visible = action
        End If
        cmCeh(I).Visible = action
        cmCeh(I).Enabled = action
    Next I
    
End Sub


Private Function chooseTheEquipment(orderStatusId As Integer, ByRef suggestedCehId As Integer) As Boolean

    Dim I As Integer
    Dim firstVisibleId As Integer
    firstVisibleId = -1
    chooseTheEquipment = True
    For I = 0 To UBound(Equip) - 1
        If ckCehDone(I).Tag <> CStr(orderStatusId) And ckCehDone(I).Tag <> "" Then
            suggestedCehId = I + 1
            Exit Function
        End If
        If firstVisibleId <> 0 And ckCehDone(I).Tag <> "" Then
            firstVisibleId = I
        End If
    Next I
    chooseTheEquipment = False
    suggestedCehId = firstVisibleId + 1
    
End Function

' returns -1 если нет ни одного оборудования
' в противном случае - statusId заказа
Private Function InitZagruz() As Integer

    Dim myCehId As Integer, cehCtlIndex As Integer, statusIsSync As Boolean
    Dim statusIdZakaz As Integer
    
    
    sql = "select oe.*, o.statusId  " _
    & " from OrdersEquip oe " _
    & " join orders o on o.numorder = oe.numorder" _
    & " where oe.numorder = " & gNzak
    
    
    Dim atLeastOne As Boolean
    atLeastOne = False
    Set tbOrders = myOpenRecordSet("##273", sql, dbOpenForwardOnly) ', dbOpenDynaset)
    statusIdZakaz = 0
    
    If tbOrders Is Nothing Then Exit Function
    While Not tbOrders.EOF
        myCehId = tbOrders("equipId")
        cehCtlIndex = myCehId - 1
        statusIdZakaz = tbOrders!StatusId
        If Not IsNull(tbOrders!statusEquipID) Then
            statusIsSync = tbOrders!statusEquipID = tbOrders!StatusId
            ' в Таг контрола ckCehDone() кладем статус по оборудованию на момент открытия формы.
            ' Потом при смене статуса заказа будем смотреть совпадает ли он со статусом по оборорудованию
            ckCehDone(cehCtlIndex).Tag = tbOrders!statusEquipID
        Else
            statusIsSync = tbOrders!StatusId = 0
            ckCehDone(cehCtlIndex).Tag = 0
        End If
        cehSelectorAccess cehCtlIndex, True, statusIsSync
        atLeastOne = True
        tbOrders.MoveNext
    Wend
    tbOrders.Close
    
    If Not atLeastOne Then
        ' warning: no ceh assigned
    Else
        Dim newEquipId As Integer
        chooseTheEquipment statusIdZakaz, newEquipId
        idEquip = newEquipId
        
    End If
    
    InitZagruz = atLeastOne
    
End Function

Private Sub Form_Load()
Dim I As Integer, str As String
FormIsActiv = False
be_cmRepit = True
workChange = False
oldHeight = Me.Height
oldWidth = Me.Width


someWasChanged = False

lv.ColumnHeaders(zkHide + 1).Width = 0

    For I = 1 To UBound(Equip) - 1
        Const HShift = 960
        Load ckCehDone(I)
        Load cmCeh(I)
        ckCehDone(I).Left = ckCehDone(I - 1).Left + HShift
        cmCeh(I).Left = cmCeh(I - 1).Left + HShift
        cmCeh(I).Caption = Equip(I + 1)
    Next I


cehSelectorsInit False

statusIdNew = -1

If festStatusId > 0 Then
    statusIdNew = festStatusId
End If


If Regim = "" Then
    If InitZagruz Then
    End If
Else
    cehSelectorsInit True
End If

startParams

End Sub

Private Sub tbDateRS_GotFocus()
If FormIsActiv Then
    cmZapros.Enabled = True
End If
tbDateRS.SelStart = 0
tbDateRS.SelLength = 2

End Sub

Private Sub tbReadyDate_GotFocus()
tbReadyDate.SelStart = 0
tbReadyDate.SelLength = 2

Me.cmZapros.Enabled = IsNumeric(tbWorktime.Text) And IsDate(tbReadyDate)

End Sub

Private Sub tbReadyDate_KeyDown(KeyCode As Integer, Shift As Integer)
Dim S As Double, I As Integer
If KeyCode = vbKeyReturn Then

If cbStatus.Text = "в работе" Then
    cmZapros.SetFocus
End If
     
If tbDateRS.Enabled Then
  If isDateTbox(tbReadyDate, "fri") Then
    S = Round(CDbl(tbWorktime.Text), 1)
    I = -(Int((CDbl(S) - 0.05) / 3) + 1 + 2) ' + 2 - дата выд от посл. куска
    getWorkDay I, tbReadyDate.Text ' дает tmpDate
    If tmpDate < curDate Then tmpDate = curDate
    tbDateRS.Text = Format(tmpDate, "dd.mm.yy")
  End If
End If

End If
Me.cmZapros.Enabled = IsNumeric(tbWorktime.Text) And IsDate(tbReadyDate)

End Sub

Private Sub tbVrVipO_Change()
If FormIsActiv Then
    cmZapros.Enabled = True
End If
End Sub


Private Sub tbWorktime_KeyDown(KeyCode As Integer, Shift As Integer)
Dim S As Double, I As Integer

If KeyCode = vbKeyReturn Then

  If isNumericTbox(tbWorktime, 0, 2000) Then
     If cbStatus.Text = "в работе" Then
        S = Round(CDbl(tbWorktime.Text), 1)
        tbWorktime.Text = S
        I = Int((CDbl(S) - 0.05) / 3)
        getWorkDay 3 + I ' дает tmpDate
        tbReadyDate.Text = Format(tmpDate, "dd.mm.yy")
        tbReadyDate.SetFocus
     Else
        tbReadyDate.Text = "00." & Format(tmpDate, "mm.yy")
     End If
  End If
Else
    cmZapros.Enabled = IsNumeric(tbWorktime.Text)
    If cmZapros.Enabled Then
        If CDbl(tbWorktime.Text) = zakazBean.Worktime Then
            workChange = True
        Else
            workChange = False
        End If
    End If
End If

End Sub

Private Sub Timer1_Timer()
tiki = tiki - 1
If tiki > 0 Then
    laMess.ForeColor = 0
    laMess.Caption = "Для нажатия на кнопку <Ok>" & _
    " у Вас осталось несколько секунд: " & tiki
    Timer1.Interval = 1000 ' 1c
Else
    Timer1.Enabled = False
    laMess.Caption = ""
    cmAdd.Enabled = False
    unLockBase
End If
End Sub



Public Function startParams() As Boolean
Dim I As Integer, str As String, J As Integer ', sumSroch As Double
Dim Item As ListItem, V As Variant, S As Double
startParams = False

maxDay = 0

Set zakazBean = New ZakazVO
If gNzak = "" Then ' вызов в режиме Сетки заказов
    Me.cmAdd.Visible = False
    Me.cmRepit.Visible = False
    gNzak = ""
    statusIdOld = 0
    Me.urgent = ""
Else

    Me.laNomZak.Caption = gNzak
    Me.cmAdd.Visible = True
    Me.cmRepit.Visible = True

    sql = "SELECT o.numorder, o.StatusId, o.DateRS, o.outTime, o.werkId, o.FirmId" _
    & ", oe.outDateTime, oe.statusEquipId, oe.equipId, oe.worktime, oe.workTimeMO" _
    & ", oc.DateTimeMO, oc.StatM, oe.StatO" _
    & ", oe.stat as statusInCeh, oe.nevip, oc.urgent" _
    & ", o.lastModified, o.lastManagId, oe.lastManagId as lastManagEquipId, 0 as presentationFormat" _
    & " from Orders o" _
    & " JOIN OrdersEquip oe on oe.numorder = o.numorder" _
    & " LEFT JOIN OrdersInCeh oc on oc.numorder = o.numorder" _
    & " WHERE o.numOrder =" & gNzak & " AND oe.equipId = " & CStr(idEquip)
    Set tbOrders = myOpenRecordSet("##402", sql, dbOpenForwardOnly)
    
    zakazBean.initFromDb
    
    tbOrders.Close
    
    If Not zakazBean.inited Then
        Exit Function
    End If
    Me.urgent = zakazBean.urgent
    
    If IsDate(zakazBean.Outdatetime) Then
        I = DateDiff("d", curDate, zakazBean.Outdatetime) + 1
        addDays I 'добавляем дни, т.к. Дата Выд тек.заказа может оказаться
                  'дальше чем всех других, либо чем stDay и rMaxDay
    End If
    
    statusIdOld = zakazBean.statusEquipID
    
    
End If
    
zagruzFromCeh idEquip, gNzak '              1| в otstup(), Ostatki() !!! кроме текущего
getResurs idEquip

Me.lvAddDays  ' добавляем стороки и даты
For I = 1 To maxDay
    Me.lv.ListItems("k" & I).SubItems(zkPrinato) = Round(getNevip(I, idEquip), 1)
    Me.lv.ListItems("k" & I).SubItems(zkResurs) = Round(nomRes(I) * KPD * Nstan, 1)
Next I
Me.lv.ListItems("k1").SubItems(zkResurs) = Round(nr * Nstan * KPD, 1)

   
lbEquip.Caption = EquipFullName(idEquip)
laStatusText.Caption = Status(zakazBean.StatusId)

tbWorktime.Text = zakazBean.Worktime

If statusIdOld = 0 Or statusIdOld = 7 Then 'принят или аннулир
    neVipolnen = 0
    neVipolnen_O = 0
    Me.Caption = "Сетка по оборудованию " & Equip(idEquip)
    
    'tbWorktime.Text = ""
    'tbReadyDate.Text = ""
Else
    Me.Caption = "Заказ " & gNzak & " - " & EquipFullName(idEquip)
    If Not IsNull(zakazBean.DateRS) Then
        Me.tbDateRS.Text = Format(zakazBean.DateRS, "dd.mm.yy")
    End If
    Me.tbReadyDate.Text = Format(zakazBean.Outdatetime, "dd.mm.yy")
          
    V = zakazBean.StatM
    'noClick = True
    If cbMOsetByText(Me.cbMaket, V) Then
        Me.tbDateMO.Text = Format(zakazBean.DateTimeMO, "dd.mm.yy")
        cbMaket.Enabled = True
    End If
    'noClick = False
     
    V = zakazBean.StatO
    If cbMOsetByText(Me.cbO, V) Then
        If Not IsNull(zakazBean.DateTimeMO) Then
            Me.tbDateMO = Format(zakazBean.DateTimeMO, "dd.mm.yy")
        Else
            Me.tbDateMO = ""
        End If
        If Me.cbO.Text = "готов" Then
            'tbVrVipO.Text = Orders.Grid.TextMatrix(Orders.mousRow, orOVrVip)
            tbVrVipO.Text = zakazBean.Worktime
            tbVrVipO.Enabled = False
            tbDateMO.Enabled = False
        Else
            neVipolnen_O = zakazBean.WorktimeMO
            tbVrVipO.Text = neVipolnen_O
            'tbVrVipO.Text = zakazBean.workTimeMO
        End If
    End If
End If

I = getNextDay(1)
V = Me.lv.ListItems("k1").SubItems(zkMost)
If Not IsNumeric(V) Then V = 0
Me.laZapas.Caption = Round(nomRes(I) * KPD * Nstan + V, 1)

'количесво фирм по дням выдачи
For I = 1 To maxDay
    otstup(I) = 0
Next I
str = "DateDiff(day, now(), oe.outDateTime)"
sql = "SELECT " & str & " AS day, o.FirmId" _
& " From Orders o" _
& " join OrdersEquip oe on oe.numorder = o.numorder and oe.equipId = " & idEquip _
& " join OrdersInCeh oc on oc.numorder = o.numorder" _
& " Where o.StatusId < 4" _
& " GROUP BY " & str & ", o.FirmId" _
& " HAVING " & str & " >= 0"

'MsgBox str & Chr(13) & Chr(13) & sql
Debug.Print sql

Set tbOrders = myOpenRecordSet("##76", sql, dbOpenForwardOnly)
If Not tbOrders Is Nothing Then
 If Not tbOrders.BOF Then
 While Not tbOrders.EOF
    I = tbOrders!day + 1
    otstup(I) = otstup(I) + 1
    tbOrders.MoveNext
 Wend
 End If
 tbOrders.Close
End If
For I = 1 To maxDay
    Me.lv.ListItems("k" & I).SubItems(zkFirmKolvo) = Round(otstup(I), 1)
Next I

cbBuildStatuses Me.cbStatus, zakazBean.StatusId


' M227 -
For I = 0 To Me.cbStatus.ListCount - 1
    If cbStatus.ItemData(I) = zakazBean.statusEquipID Then
        noClick = True
        Me.cbStatus.ListIndex = I
        noClick = False
        GoTo NN
    End If
Next I

    noClick = True
    Me.cbStatus.ListIndex = 1
    noClick = False

NN:

Me.cmZapros.Enabled = (IsNumeric(tbWorktime.Text) And IsDate(tbReadyDate)) Or statusIdNew = 0

Me.lv.ListItems("k" & stDay).ForeColor = &HBB00&
Me.lv.ListItems("k" & stDay).Bold = True

Me.newZagruz Me.Regim  'влияет только один раз

startParams = True
End Function

