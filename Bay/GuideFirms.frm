VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GuideFirms 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Справочник сторонних организаций"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.ListBox lbOborud 
      Height          =   1620
      ItemData        =   "GuideFirms.frx":0000
      Left            =   1320
      List            =   "GuideFirms.frx":001C
      TabIndex        =   25
      Top             =   2640
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmLoad 
      Caption         =   "Обновить"
      Height          =   315
      Left            =   180
      TabIndex        =   24
      Top             =   7680
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3675
      Left            =   5760
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   4755
      Begin VB.TextBox tbType 
         Height          =   2835
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   240
         Width           =   4515
      End
      Begin VB.CommandButton cmCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   3600
         TabIndex        =   23
         Top             =   3240
         Width           =   795
      End
      Begin VB.CommandButton cmOk 
         Caption         =   "Ok"
         Height          =   315
         Left            =   660
         TabIndex        =   22
         Top             =   3240
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Вид  деятельности"
         Height          =   255
         Left            =   1620
         TabIndex        =   21
         Top             =   0
         Width           =   1515
      End
   End
   Begin VB.ListBox lbRegion 
      Height          =   5715
      ItemData        =   "GuideFirms.frx":005E
      Left            =   2640
      List            =   "GuideFirms.frx":0060
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Exel"
      Height          =   315
      Left            =   6600
      TabIndex        =   17
      Top             =   7680
      Width           =   1215
   End
   Begin VB.ComboBox cbM 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   540
      Width           =   1635
   End
   Begin VB.TextBox tbInform 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   555
      Width           =   8835
   End
   Begin VB.Timer Timer1 
      Left            =   3540
      Top             =   7740
   End
   Begin VB.CommandButton cmAllOrders 
      Caption         =   "Отч.""Все заказы фирмы"
      Height          =   315
      Left            =   7020
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmNoClose 
      Caption         =   "Отчет ""Незакрытые заказы""  "
      Height          =   315
      Left            =   4380
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmNoCloseFiltr 
      Caption         =   "Фильтр""Незакрытые заказы"""
      Height          =   315
      Left            =   9180
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox lbM 
      Height          =   255
      Left            =   420
      TabIndex        =   12
      Top             =   1860
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2460
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox tbMobile 
      Height          =   285
      Left            =   600
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "tbMobile"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   7680
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmFind 
      Caption         =   "Поиск"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox tbFind 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmSel 
      Caption         =   "Выбрать"
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   11100
      TabIndex        =   10
      Top             =   7680
      Width           =   675
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6435
      Left            =   180
      TabIndex        =   5
      Top             =   1020
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   11351
      _Version        =   393216
      MergeCells      =   2
      AllowUserResizing=   1
      FormatString    =   " "
   End
   Begin VB.Label Label2 
      Caption         =   "Фильтр:"
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   600
      Width           =   735
   End
   Begin VB.Label laQuant 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   315
      Left            =   5520
      TabIndex        =   14
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label laHeadQ 
      Caption         =   "Число записей:"
      Height          =   195
      Left            =   4260
      TabIndex        =   13
      Top             =   7740
      Width           =   1215
   End
End
Attribute VB_Name = "GuideFirms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Regim As String ' режим окна
Public mousRow As Long    '
Public mousCol As Long    '
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim quantity As Integer 'количество найденных фирм
Dim pos As Long 'поиция поиска
Const cEmpty = "пустой менеджер"

Private Sub chClose_Click()

End Sub

Private Sub cbM_Click()
loadGuide
On Error Resume Next ' требуется при вызове из Load
Grid.SetFocus
End Sub

Private Sub cmAdd_Click()
If Grid.TextMatrix(Grid.Rows - 1, gfId) <> "" Then Grid.AddItem ("")

'Grid.col = gfId ' чтобы наверняка было соб.EnterCell по Grid.col = gfNazwFirm
Grid.row = Grid.Rows - 1
Grid.col = gfNazwFirm 'название
Grid.SetFocus
textBoxInGridCell tbMobile, Grid
End Sub

Private Sub cmAllOrders_Click()
Me.MousePointer = flexHourglass
Report.Regim = "allFromFirms"
Report.Show vbModal
Grid.SetFocus
Me.MousePointer = flexDefault


End Sub

Private Sub cmCancel_Click()
    Frame1.Visible = False
tbInform.Text = Grid.TextMatrix(mousRow, gfType)
Grid.SetFocus
End Sub

Private Sub cmDel_Click()
Dim strId As String, I As Integer

If MsgBox("По кнопке <Да> вся информация по фирме будет безвозвратно " & _
"удалена из базы!", vbYesNo, "Удалить Фирму?") = vbNo Then Exit Sub

strId = Grid.TextMatrix(mousRow, gfId)
'sql = "SELECT BayGuideFirms.FirmId, BayGuideFirms.Name  From BayGuideFirms " & _
'"WHERE (((BayGuideFirms.FirmId)=" & strId & "));"
'Set tbFirms = myOpenRecordSet("##67", sql, dbOpenDynaset)
'If tbFirms Is Nothing Then Exit Sub
'On Error GoTo ERR1
'tbFirms.MoveFirst
'tbFirms.Delete
'tbFirms.Close

sql = "DELETE FROM BayGuideFirms WHERE FirmId = " & strId
I = myExecute("##67", sql, -198)
If I = -2 Then
    MsgBox "У этой фирмы есть заказы. Перед ее удалением необходимо " & _
    "в этих заказах выбать другую фирму, либо удалить эти заказы", , _
    "Удаление невозможно!"
    GoTo EN1
ElseIf I <> 0 Then
    GoTo EN1
End If

quantity = quantity - 1
If quantity = 0 Then
    clearGridRow Grid, mousRow
Else
    Grid.RemoveItem mousRow
End If
'Grid.SetFocus
'Exit Sub'

'ERR1:
'If Err = 3200 Then
'    MsgBox "У этой фирмы есть заказы. Перед ее удалением необходимо " & _
'    "в этих заказах выбать другую фирму, либо удалить эти заказы", , _
'    "Удаление невозможно!"
'Else
'    MsgBox Error, , "Ошибка 352-" & Err & ":  " '##352
'End If
EN1:
Grid.SetFocus
End Sub

Private Sub cmExel_Click()
GridToExcel Grid, "Справочник сторонних организаций (" & cbM.Text & ")"

End Sub

Private Sub cmExit_Click()
    Unload Me
End Sub

Public Sub cmFind_Click()
'Static pos As Long
pos = findExValInCol(Grid, tbFind.Text, gfNazwFirm, pos)
If pos > 0 Then
    cmSel.Enabled = True
    cmFind.Caption = "Далее"
    Grid.SetFocus
Else
    cmSel.Enabled = False
    tbFind.SetFocus
End If
pos = pos + 1

End Sub

Sub lbHide(Optional noGrid As String)
tbMobile.Visible = False
lbM.Visible = False
lbRegion.Visible = False
lbOborud.Visible = False
Frame1.Visible = False
Grid.Enabled = True
If noGrid <> "" Then Exit Sub
Grid.SetFocus
Grid_EnterCell
End Sub

Sub loadGuide()
Dim I As Long, strWhere As String, str As String

Me.MousePointer = flexHourglass
Grid.Visible = False
clearGrid Grid
strWhere = Trim(tbFind.Text)
If Not strWhere = "" Then
'    strWhere = "Where (((BayGuideFirms.Name) = '" & strWhere & "' )) "
    strWhere = "(BayGuideFirms.Name) = '" & strWhere & "'"
End If
str = ""
If cbM.ListIndex > 0 Then str = "(BayGuideFirms.ManagId) = " & _
    manId(cbM.ListIndex - 1)
If strWhere <> "" And str <> "" Then
    strWhere = strWhere & " AND " & str
Else
    strWhere = strWhere & str
End If
If strWhere <> "" Then strWhere = "Where ((" & strWhere & ")) "
'MsgBox "strWhere = " & strWhere
quantity = 0

sql = "SELECT f.*, isnull(r.region, '') as region, isnull(u.oborud, '') as oborud " _
& " FROM BayGuideFirms f " _
& " left join BayRegion r on r.regionid = f.regionid" _
& " left join GuideOborud u on u.oborudId = f.oborudId" _
& strWhere _
& " ORDER BY Name"
'MsgBox sql
Set tbFirms = myOpenRecordSet("##15", sql, dbOpenForwardOnly)
If tbFirms Is Nothing Then GoTo EN1

If Not tbFirms.BOF Then
  'tbFirms.MoveFirst
  While Not tbFirms.EOF
    If tbFirms!firmId = 0 Then GoTo AA
    quantity = quantity + 1
If tbFirms!firmId = 39 Then
I = I
End If
    Grid.TextMatrix(quantity, gfId) = tbFirms!firmId
    Grid.TextMatrix(quantity, gfNazwFirm) = tbFirms!Name
    Grid.TextMatrix(quantity, gfM) = Manag(tbFirms!ManagId)
    fieldToCol tbFirms!Oborud, gfOborud
    fieldToCol tbFirms!Sale, gfSale
    fieldToCol tbFirms!Kontakt, gfKontakt
    fieldToCol tbFirms!Otklik, gfOtklik
    fieldToCol tbFirms!year01, gf2001 '$$3
    fieldToCol tbFirms!year02, gf2002 '
    fieldToCol tbFirms!year03, gf2003 '
    fieldToCol tbFirms!year04, gf2004

    fieldToCol tbFirms!FIO, gfFIO
    fieldToCol tbFirms!Fax, gfFax
    fieldToCol tbFirms!Email, gfEmail
    fieldToCol tbFirms!Type, gfType
    fieldToCol tbFirms!Pass, gfPass
    fieldToCol tbFirms!region, gfRegion
    fieldToCol tbFirms!xLogin, gfLogin
    fieldToCol tbFirms!Phone, gfTlf
    Grid.AddItem ("")
AA: tbFirms.MoveNext
  Wend
  If quantity > 0 Then Grid.RemoveItem (quantity + 1)
End If
tbFirms.Close
EN1:
Grid.Visible = True
laQuant.Caption = quantity
Me.MousePointer = flexDefault

End Sub

Sub fieldToCol(field As Variant, col As Long)
If Not IsNull(field) Then Grid.TextMatrix(quantity, col) = field
End Sub

Private Sub cmLoad_Click()
loadGuide
Grid.SetFocus
End Sub

Private Sub cmNoClose_Click()
Me.MousePointer = flexHourglass
Report.Regim = "FromFirms"
Report.Show vbModal
Grid.SetFocus
Me.MousePointer = flexDefault

End Sub

Private Sub cmNoCloseFiltr_Click()
Dim str As String
str = Grid.TextMatrix(mousRow, gfNazwFirm)
Unload Me
Orders.loadFirmOrders str

End Sub

Private Sub cmOk_Click()
If ValueToTableField("##353", "'" & tbType.Text & "'", "BayGuideFirms", _
"Type", "byFirmId") = 0 Then
    Grid.TextMatrix(mousRow, gfType) = tbType.Text
End If
Frame1.Visible = False
Grid.SetFocus
End Sub

Private Sub cmSel_Click()
Dim sqlReq As String, firmId As String, DNM As String

    Orders.Grid.Text = Grid.Text

    gNzak = Orders.Grid.TextMatrix(Orders.Grid.row, orNomZak)
    visits "-", "firm" ' уменьщаем посещения у старой фирмы, если она была
    firmId = Grid.TextMatrix(Grid.row, gfId)
    ValueToTableField "##20", firmId, "BayOrders", "FirmId"
    visits "+", "firm" ' увеличиваем посещения у новой фирмы

    DNM = Format(Now(), "dd.mm.yy hh:nn") & vbTab & Orders.cbM.Text & " " & gNzak ' именно vbTab
    On Error Resume Next ' в некот.ситуациях один из Open logFile дает Err: файл уже открыт
    Open logFile For Append As #2
    Print #2, DNM & " фирма=" & Grid.Text
    Close #2
    
Unload Me

'Orders.SetFocus
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then lbHide
End Sub

Private Sub Form_Load()
Dim I As Integer
quantity = 0
pos = 0
oldHeight = Me.Height
oldWidth = Me.Width

Grid.FormatString = "|< Название  фирмы|^ M|<Оборудование|<Регион|Скидки в %|" & _
"<Контакт|<Отклик|200x|2001|2002|2004|<Конт.лицо|<Телефон|<Факс|<e-mail|<Вид деятельности" & _
"|<Логин|<Пароль|>Id"
'сделать чтобы в 2004 - высвечивалась колонка, в 2005 - 2,2006 - 3, далее -4

Grid.TextMatrix(0, gf2002) = Format(lastYear - 2, "0000") '$$3
Grid.TextMatrix(0, gf2003) = Format(lastYear - 1, "0000")
Grid.TextMatrix(0, gf2004) = Format(lastYear, "0000")

If lastYear < 2007 Then Grid.ColWidth(gf2001) = 0 '$$3
If lastYear < 2006 Then Grid.ColWidth(gf2002) = 0
If lastYear < 2005 Then Grid.ColWidth(gf2003) = 0

Grid.MergeRow(0) = True
Grid.ColWidth(0) = 0
Grid.ColWidth(gfM) = 330
Grid.ColWidth(gfNazwFirm) = 2730
Grid.ColWidth(gfOborud) = 735
Grid.ColWidth(gfRegion) = 1140
Grid.ColWidth(gfSale) = 655
Grid.ColWidth(gfKontakt) = 700
Grid.ColWidth(gfOtklik) = 645
Grid.ColWidth(gfFIO) = 1410
Grid.ColWidth(gfTlf) = 1140
'Grid.ColWidth(gfType) = 615 в Resize
Grid.ColWidth(gfLogin) = 780
Grid.ColWidth(gfId) = 480

cbM.AddItem "все менеджеры"
lbM.AddItem "not"
For I = 0 To Orders.lbM.ListCount - 1
    If I < Orders.lbM.ListCount - 1 Then lbM.AddItem Orders.lbM.List(I)
    If Orders.lbM.List(I) = "" Then
        cbM.AddItem cEmpty
    Else
        cbM.AddItem "менеджер " & Orders.lbM.List(I)
    End If
Next I
cbM.ListIndex = 0
lbM.Height = lbM.Height + 195 * (lbM.ListCount - 1)

Me.Caption = "Справочник сторонних организаций"

sql = "SELECT Region FROM BayRegion ORDER BY Region"
Set tbGuide = myOpenRecordSet("##349", sql, dbOpenForwardOnly)
'tbGuide.Index = "Region"
'If Not tbGuide Is Nothing Then
  While Not tbGuide.EOF
    lbRegion.AddItem tbGuide!region
    tbGuide.MoveNext
  Wend
  tbGuide.Close
'End If

If tbFind.Text <> "" Then cmFind.Enabled = True
If Regim = "fromContext" Then 'из Orders
    tbFind.Text = Orders.Grid.Text
    tbFind.SelStart = 0
    tbFind.SelLength = Len(GuideFirms.tbFind.Text)
    cmSel.Visible = True
    GoTo AA
ElseIf Regim = "fromFindFirm" Then
'    tbFind.Text = FindFirm.lb.Text
    tbFind.SelStart = 0
    tbFind.SelLength = Len(GuideFirms.tbFind.Text)
'    cmSel.Visible = FindFirm.cmSelect.Visible
ElseIf Regim = "fromMenu" Then 'по <F11> из Orders
    cmLoad.Visible = True
AA: If Orders.tbEnable.Visible Then
        cmNoClose.Visible = True
        cmAllOrders.Visible = True
        cmNoCloseFiltr.Visible = True
    End If
End If
cmAdd.Visible = True
cmDel.Visible = True

Set table = myOpenRecordSet("##72", "bayRegion", dbOpenForwardOnly)
If table Is Nothing Then myBase.Close: End

'loadGuide не надо, т.к. при загрузке cbM_Click
Timer1.Interval = 100
Timer1.Enabled = True

End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If Me.WindowState = vbMinimized Then Exit Sub
If Me.WindowState = vbMaximized And Me.Width > cDELLwidth Then 'экран DELL
    Grid.ColWidth(gfType) = 5430
Else
    Grid.ColWidth(gfType) = 615
End If
On Error Resume Next

lbHide "noGrid"
w = Me.Width - oldWidth
oldWidth = Me.Width
h = Me.Height - oldHeight
oldHeight = Me.Height

Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w
cmSel.Top = cmSel.Top + h
cmExit.Top = cmExit.Top + h
cmExit.left = cmExit.left + w
cmDel.Top = cmDel.Top + h
cmAdd.Top = cmAdd.Top + h
laHeadQ.Top = laHeadQ.Top + h
laQuant.Top = laQuant.Top + h
cmExel.Top = cmExel.Top + h
cmLoad.Top = cmLoad.Top + h
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Me.Visible = False
'Orders.Enabled = True
'Orders.SetFocus
End Sub

Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If quantity = 0 Then Exit Sub
If Grid.MouseRow = 0 Then
    Grid.CellBackColor = Grid.BackColor
    If mousCol = gf2004 Then
        SortCol Grid, mousCol, "numeric"
    Else
        SortCol Grid, mousCol
    End If
    Grid.row = 1    ' только чтобы снять выделение
    Grid_EnterCell
End If

End Sub

Private Sub Grid_DblClick()
Dim I As Integer

If Grid.CellBackColor = vbYellow Then Exit Sub

gFirmId = Grid.TextMatrix(mousRow, gfId)

If mousCol = gfM Then
    listBoxInGridCell lbM, Grid, "select"
ElseIf mousCol = gfOborud Then
    listBoxInGridCell lbOborud, Grid, "select"
ElseIf mousCol = gfRegion Then ' Регион
        For I = 0 To lbRegion.ListCount - 1 '
            If Grid.Text = lbRegion.List(I) Then
'                noClick = True
                lbRegion.ListIndex = I 'вызывает ложное onClick
'                noClick = False
                Exit For
            End If
        Next I
    lbRegion.Visible = True
    lbRegion.ZOrder
    lbRegion.SetFocus
    Grid.Enabled = False 'иначе курсор по ней бегает
'    listBoxInGridCell lbRegion, Grid, "select"
ElseIf mousCol = gfType Then
    tbType.Text = Grid.Text
'    tbType.SelLength = Len(tbType.Text)
    Frame1.Visible = True
    Frame1.ZOrder
    tbType.SetFocus
Else
    textBoxInGridCell tbMobile, Grid
'ElseIf cmSel.Enabled Then
'    If cmSel.Visible Then cmSel_Click
End If
End Sub

Private Sub Grid_EnterCell()
If noClick Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col
'If quantity = 0 Or Regim = "F7" Then Exit Sub
'If Regim = "F7" Then Exit Sub
If mousCol = gfNazwFirm Then
    cmSel.Enabled = True
    cmDel.Enabled = True
Else
    cmSel.Enabled = False
    cmDel.Enabled = False
End If

If mousCol = gfNazwFirm Or mousCol = gfFIO Then
    tbInform.MaxLength = 80
    tbMobile.MaxLength = 80
ElseIf mousCol = gfType Then
    tbInform.MaxLength = 255
'    tbMobile.MaxLength = 255
ElseIf mousCol = gfTlf Or mousCol = gfSale Or _
mousCol = gfFax Or mousCol = gfEmail Then
    tbInform.MaxLength = 50
    tbMobile.MaxLength = 50
Else
    tbInform.MaxLength = 10
    tbMobile.MaxLength = 10
End If
tbInform.Text = Grid.TextMatrix(mousRow, mousCol)

If mousCol = gfId Or mousCol = gfLogin Or mousCol = gfPass Or _
mousCol = gf2004 Or Grid.TextMatrix(mousRow, gfId) = "" Then
    Grid.CellBackColor = vbYellow
    tbInform.Locked = True
Else
    Grid.CellBackColor = &H88FF88 ' бл.зел
'   Grid.CellBackColor = &H8888FF    ' бл.кр
    If mousCol = gfM Or mousCol = gfRegion Or mousCol = gfType Then
        tbInform.Locked = True
    Else
        tbInform.Locked = False
    End If
End If

End Sub

Private Sub Grid_GotFocus()
Grid_EnterCell
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
        Grid_DblClick
ElseIf KeyCode = vbKeyF4 Then
    If mousCol = gfLogin Or mousCol = gfPass Then _
                textBoxInGridCell tbMobile, Grid
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If
End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
End Sub

Private Sub lbOborud_DblClick()

If lbOborud.Text = "" Then
    sql = "update bayguidefirms set Oborudid = null " _
    & " where bayguidefirms.firmId = " & gFirmId
Else
    sql = "update bayguidefirms set Oborudid = u.Oborudid " _
    & " from GuideOborud u" _
    & " where bayguidefirms.firmId = " & gFirmId _
    & " and u.Oborud = '" & lbOborud.Text & "'"
End If
myExecute "##354", sql

Grid.TextMatrix(mousRow, gfOborud) = lbOborud.Text

lbHide

End Sub

Private Sub lbOborud_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbOborud_DblClick
End Sub

Private Sub lbRegion_DblClick()
'Region

If lbRegion.Text = "" Then
    sql = "update bayguidefirms set regionid = null " _
    & " where bayguidefirms.firmId = " & gFirmId
Else
    sql = "update bayguidefirms set regionid = r.regionid " _
    & " from bayregion r" _
    & " where bayguidefirms.firmId = " & gFirmId _
    & " and r.region = '" & lbRegion.Text & "'"
End If

myExecute "##354", sql

Grid.TextMatrix(mousRow, gfRegion) = lbRegion.Text
lbHide
cmAdd.Enabled = True
cmExit.Enabled = True
cbM.Enabled = True
cmExel.Enabled = True
End Sub

Private Sub lbRegion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbRegion_DblClick

End Sub

Private Sub lbM_DblClick()
Dim str As String

If lbM.ListIndex = 0 Then
    str = "14" ' not
Else
    str = manId(lbM.ListIndex - 1)
End If
ValueToTableField "##355", str, "BayGuideFirms", "ManagId", "byFirmId"
Grid.TextMatrix(mousRow, gfM) = lbM.Text
lbHide

End Sub

Private Sub lbM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbM_DblClick
End Sub

Private Sub tbFind_Change()
If tbFind.Text <> "" Then cmFind.Enabled = True
End Sub

Private Sub tbFind_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmFind_Click
End Sub

Private Sub tbInform_GotFocus()
    tbInform.SelStart = Len(tbInform.Text)
End Sub

Private Sub tbInform_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
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
Dim str As String, I As Integer, strId As String
If KeyCode = vbKeyReturn Then
 str = Trim(tbMobile.Text)
 gFirmId = Grid.TextMatrix(mousRow, gfId)

 If mousCol = gfNazwFirm Then
   strId = Grid.TextMatrix(mousRow, gfId)
'   On Error GoTo ERR1
   If strId = "" Then
    wrkDefault.BeginTrans
    sql = "update bayGuideFirms set firmId = firmId where firmId = 0"
    
    sql = "select max(firmid) from bayGuideFirms;"
    If Not byErrSqlGetValues("##50", sql, gFirmId) Then GoTo ERR1
    gFirmId = gFirmId + 1
    
    sql = "insert into bayGuideFirms (firmId, name, ManagId) values (" & _
    gFirmId & ", '" & str & "', 14)"
    I = myExecute("##50", sql, -196)
    If I <> 0 Then GoTo ERR0:
    wrkDefault.CommitTrans
    
    
    Grid.TextMatrix(mousRow, gfId) = gFirmId
    quantity = quantity + 1
'    cep = True ' запускаем цепь посл.ввода еще 2х полей
'    cmAdd.Enabled = False
'    cmDel.Enabled = False
'    cmExit.Enabled = False
'    cbM.Enabled = False
'    cmExel.Enabled = False
    Grid.TextMatrix(mousRow, mousCol) = str
    Grid.TextMatrix(mousRow, gfM) = "not"
    lbHide
    Grid.col = gfM: mousCol = gfM
    listBoxInGridCell lbM, Grid
    Exit Sub
   Else
    sql = "UPDATE BayGuideFirms SET Name = '" & str & _
    "' WHERE (((FirmId)=" & strId & "));"
'    MsgBox sql
    I = myExecute("##356", sql, -196)
    If I <> 0 Then GoTo ERR0:
    
   End If
   On Error GoTo 0
 ElseIf mousCol = gfSale Then
    ValueToTableField "##66", "'" & str & "'", "BayGuideFirms", "Sale", "byFirmId"
 ElseIf mousCol = gfKontakt Then
    ValueToTableField "##66", "'" & str & "'", "BayGuideFirms", "Kontakt", "byFirmId"
 ElseIf mousCol = gfOtklik Then
    ValueToTableField "##66", "'" & str & "'", "BayGuideFirms", "Otklik", "byFirmId"
 ElseIf mousCol = gfFIO Then
    ValueToTableField "##66", "'" & str & "'", "BayGuideFirms", "FIO", "byFirmId"
 ElseIf mousCol = gfTlf Then
    ValueToTableField "##66", "'" & str & "'", "BayGuideFirms", "Phone", "byFirmId"
 ElseIf mousCol = gfFax Then
    ValueToTableField "##66", "'" & str & "'", "BayGuideFirms", "Fax", "byFirmId"
 ElseIf mousCol = gfEmail Then
    ValueToTableField "##66", "'" & str & "'", "BayGuideFirms", "Email", "byFirmId"
 ElseIf mousCol = gfLogin Then
    If str <> "" And str <> Grid.TextMatrix(mousRow, gfLogin) Then
        If existValueInTableFielf(str, "BayGuideFirms", "xLogin") Then
            MsgBox "Такой логин уже есть.", , "Недопустимое значение"
            tbMobile.SelStart = Len(tbMobile.Text)
            tbMobile.SetFocus
            Exit Sub
        End If
    End If
    ValueToTableField "##66", "'" & str & "'", "BayGuideFirms", "xLogin", "byFirmId"
 ElseIf mousCol = gfPass Then
    ValueToTableField "##66", "'" & str & "'", "BayGuideFirms", "Pass", "byFirmId"
End If
'GoTo AA
'ElseIf KeyCode = vbKeyEscape Then
AA:
 Grid.TextMatrix(mousRow, mousCol) = str
 lbHide
End If
Exit Sub

ERR0: If I = -2 Then
        MsgBox "Такая фирма уже есть", , "Ошибка!"
        tbMobile.SetFocus
      Else
ERR1:   lbHide
      End If
End Sub

Private Sub tbType_Change()
tbInform.Text = tbType.Text
End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False
 If Regim = "fromMenu" Then  'по F11 из Orders
    tbFind.SetFocus
 Else ' из контекстного меню
    cmFind.Caption = "Поиск"
   'cmFind_Click
    If findValInCol(Grid, tbFind.Text, gfNazwFirm) Then
        Grid.SetFocus
    Else
        tbFind.SetFocus
    End If
 End If

End Sub
