VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GuideFirms 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Справочник сторонних организаций"
   ClientHeight    =   8184
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   11892
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8184
   ScaleWidth      =   11892
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Exel"
      Height          =   315
      Left            =   9180
      TabIndex        =   18
      Top             =   7680
      Width           =   1215
   End
   Begin VB.ComboBox cbM 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   540
      Width           =   1635
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
      Height          =   240
      Left            =   3780
      TabIndex        =   13
      Top             =   2100
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2340
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ListBox lbKP 
      Height          =   816
      ItemData        =   "GuideFirms.frx":0000
      Left            =   3300
      List            =   "GuideFirms.frx":0010
      TabIndex        =   12
      Top             =   1980
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox tbMobile 
      Height          =   285
      Left            =   4560
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "tbMobile"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   1200
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
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   11715
      _ExtentX        =   20659
      _ExtentY        =   11345
      _Version        =   393216
      MergeCells      =   2
      AllowUserResizing=   1
      FormatString    =   " "
   End
   Begin VB.CommandButton cmLoad 
      Caption         =   "Обновить"
      Height          =   315
      Left            =   180
      TabIndex        =   19
      Top             =   7680
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Фильтр:"
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   600
      Width           =   735
   End
   Begin VB.Label laQuant 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   315
      Left            =   5520
      TabIndex        =   15
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Число записей:"
      Height          =   195
      Left            =   4260
      TabIndex        =   14
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
Const cEmpty = "пустой менеджер"

Private Sub chClose_Click()

End Sub

Private Sub cbM_Click()
loadGuide
On Error Resume Next ' требуется при вызове из Load
Grid.SetFocus
End Sub

Private Sub cmAdd_Click()
If Grid.TextMatrix(Grid.rows - 1, gfId) <> "" Then Grid.AddItem ("")

Grid.row = Grid.rows - 1
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

Private Sub cmDel_Click()
Dim strId As String, I As Integer

If MsgBox("По кнопке <Да> вся информация по фирме будет безвозвратно " & _
"удалена из базы!", vbYesNo, "Удалить Фирму?") = vbNo Then Exit Sub

strId = Grid.TextMatrix(mousRow, gfId)
'sql = "SELECT GuideFirms.FirmId, GuideFirms.Name  From GuideFirms " & _
'"WHERE (((GuideFirms.FirmId)=" & strId & "));"
'Set tbFirms = myOpenRecordSet("##67", sql, dbOpenDynaset)
'If tbFirms Is Nothing Then Exit Sub
'On Error GoTo ERR1
'tbFirms.MoveFirst
'tbFirms.Delete
'tbFirms.Close
sql = "DELETE FROM GuideFirms WHERE FirmId = " & strId
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
    Grid.removeItem mousRow
End If

'Grid.SetFocus
'Exit Sub

'ERR1:
'If Err = 3200 Then
'    MsgBox "У этой фирмы есть заказы. Перед ее удалением необходимо " & _
'    "в этих заказах выбать другую фирму, либо удалить эти заказы", , _
'    "Удаление невозможно!"
'Else
'    errorCodAndMsg ("66")
'    MsgBox Error, , "Ошибка 66-" & Err & ":  " '##66
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
Static pos As Long
pos = findExValInCol(Grid, tbFind.Text, gfNazwFirm, pos)
On Error Resume Next
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
lbKP.Visible = False
lbM.Visible = False
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
strWhere = trimAll(tbFind.Text)
If Not strWhere = "" Then
'    strWhere = "Where (((GuideFirms.Name) = '" & strWhere & "' )) "
    strWhere = "(GuideFirms.Name) = '" & strWhere & "'"
End If
str = ""
If cbM.ListIndex > 0 Then str = "(GuideFirms.ManagId) = " & _
    manId(cbM.ListIndex - 1)
If strWhere <> "" And str <> "" Then
    strWhere = strWhere & " AND " & str
Else
    strWhere = strWhere & str
End If
If strWhere <> "" Then strWhere = "Where ((" & strWhere & ")) "
'MsgBox "strWhere = " & strWhere
quantity = 0

sql = "SELECT GuideFirms.FirmId, GuideFirms.Name, GuideFirms.xLogin, " & _
"GuideFirms.Address, GuideFirms.Phone, GuideFirms.Kategor, GuideFirms.Sale, " & _
"GuideFirms.year01, GuideFirms.year02, GuideFirms.year03, GuideFirms.year04, " & _
"GuideManag.Manag, GuideFirms.FIO, GuideFirms.Fax, GuideFirms.Email, " & _
"GuideFirms.Atr1, GuideFirms.Atr2, GuideFirms.Atr3, GuideFirms.Pass, " & _
"GuideFirms.Level, GuideFirms.Type, GuideFirms.Katalog " & _
"FROM GuideManag RIGHT JOIN GuideFirms ON GuideManag.ManagId = GuideFirms.ManagId " & _
strWhere & "ORDER BY GuideFirms.Name;"
'MsgBox sql
Set tbFirms = myOpenRecordSet("##15", sql, dbOpenForwardOnly) 'dbOpenSnapshot)
If tbFirms Is Nothing Then GoTo EN1

If Not tbFirms.BOF Then
  'tbFirms.MoveFirst
  While Not tbFirms.EOF
    If tbFirms!firmId = 0 Then GoTo AA
    quantity = quantity + 1
'    Grid.TextMatrix(quantity, 0) = quantity
    Grid.TextMatrix(quantity, gfId) = tbFirms!firmId
    Grid.TextMatrix(quantity, gfNazwFirm) = tbFirms!name
'    Grid.TextMatrix(quantity, gfM) = Manag(tbFirms!ManagId)
'    If Not IsNull(tbFirms!Manag) Then
    fieldToCol tbFirms!Manag, gfM
    fieldToCol tbFirms!Sale, gfSale
    fieldToCol tbFirms!year01, gf2001
    fieldToCol tbFirms!year02, gf2002
    fieldToCol tbFirms!year03, gf2003
    fieldToCol tbFirms!year04, gf2004
    fieldToCol tbFirms!FIO, gfFIO
    fieldToCol tbFirms!Fax, gfFax
    fieldToCol tbFirms!Email, gfEmail
    fieldToCol tbFirms!Level, gfLevel
    fieldToCol tbFirms!Type, gfType
    fieldToCol tbFirms!Katalog, gfKatalog
    fieldToCol tbFirms!Atr1, gfAtr1
    fieldToCol tbFirms!Atr2, gfAtr2
    fieldToCol tbFirms!Atr3, gfAtr3
    fieldToCol tbFirms!PASS, gfPass
    fieldToCol tbFirms!Kategor, gfKategor
    fieldToCol tbFirms!xLogin, gfLogin
    fieldToCol tbFirms!Address, gfAdres
    fieldToCol tbFirms!Phone, gfTlf
    Grid.AddItem ("")
AA: tbFirms.MoveNext
  Wend
  If quantity > 0 Then Grid.removeItem (quantity + 1)
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

Private Sub cmSel_Click()
Dim sqlReq As String, firmId As String, DNM As String

    Orders.Grid.Text = Grid.Text

    gNzak = Orders.Grid.TextMatrix(Orders.Grid.row, orNomZak)
    visits "-", "firm" ' уменьщаем посещения у старой фирмы, если она была
    firmId = Grid.TextMatrix(Grid.row, gfId)
    ValueToTableField "##20", firmId, "Orders", "FirmId"
    visits "+", "firm" ' увеличиваем посещения у новой фирмы

    DNM = Format(Now(), "dd.mm.yy hh:nn") & vbTab & Orders.cbM.Text & " " & gNzak ' именно vbTab
    On Error Resume Next ' в некот.ситуациях один из Open logFile дает Err: файл уже открыт
    Open logFile For Append As #2
    Print #2, DNM & " фирма=" & Grid.Text
    Close #2
    Filtr.lbFirm.AddItem Grid.Text, 0
    Filtr.lbFirm.Selected(0) = True
    refreshTimestamp gNzak
    
    
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape And Not cep Then lbHide
End Sub

Private Sub Form_Load()
Dim I As Integer
quantity = 0

Grid.FormatString = "|< Название  фирмы|^ M|Kатегория|Скидки в %|200x|2002" & _
"|2003|2004|<Конт.лицо|<Телефон|<Факс|<e-mail|<Вид деятельности|<Каталог" & _
"|<Специализация|<Примечание|Атрибуты|Атрибуты|Атрибуты|<Логин|<Пароль|>Id"

Grid.TextMatrix(0, gf2002) = Format(lastYear - 2, "0000") '$$3
Grid.TextMatrix(0, gf2003) = Format(lastYear - 1, "0000")
Grid.TextMatrix(0, gf2004) = Format(lastYear, "0000")

Grid.MergeRow(0) = True
Grid.ColWidth(0) = 0
Grid.ColWidth(gfM) = 330
Grid.ColWidth(gfNazwFirm) = 2730
Grid.ColWidth(gfKategor) = 525
Grid.ColWidth(gfSale) = 655
Grid.ColWidth(gfTlf) = 1140
Grid.ColWidth(gfLevel) = 750
Grid.ColWidth(gfType) = 615
Grid.ColWidth(gfAtr1) = 300
Grid.ColWidth(gfAtr2) = 300
Grid.ColWidth(gfAtr3) = 700
Grid.ColWidth(gfLogin) = 780
Grid.ColWidth(gfAdres) = 1665
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

If tbFind.Text <> "" Then cmFind.Enabled = True
If Regim = "fromContext" Then 'из Orders
    tbFind.Text = Orders.Grid.Text
    tbFind.SelStart = 0
    tbFind.SelLength = Len(GuideFirms.tbFind.Text)
    cmSel.Visible = True
    cmLoad.Visible = False
    GoTo AA
ElseIf Regim = "fromFindFirm" Then
'    tbFind.Text = FindFirm.lb.Text
    tbFind.SelStart = 0
    tbFind.SelLength = Len(GuideFirms.tbFind.Text)
ElseIf Regim = "fromMenu" Then 'из Orders
    cmLoad.Visible = True
AA: If Orders.tbEnable.Visible Then
        cmNoClose.Visible = True
        cmAllOrders.Visible = True
        cmNoCloseFiltr.Visible = True
    End If
End If
cmAdd.Visible = True
cmDel.Visible = True

Timer1.Interval = 100
Timer1.Enabled = True

oldHeight = Me.Height
oldWidth = Me.Width
End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If Me.WindowState = vbMinimized Then Exit Sub
If Me.WindowState = vbMaximized And Me.Width > cDELLwidth Then 'экран DELL
    Grid.ColWidth(gfFIO) = 3090
Else
    Grid.ColWidth(gfFIO) = 1410
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
'cmExit.Left = cmExit.Left + w
cmDel.Top = cmDel.Top + h
'tbFind.Top = tbFind.Top + h
'cmFind.Top = cmFind.Top + h
cmAdd.Top = cmAdd.Top + h
cmExel.Top = cmExel.Top + h
cmExel.Left = cmExel.Left + w
End Sub

Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If quantity = 0 Then Exit Sub
If Grid.MouseRow = 0 Then
    Grid.CellBackColor = Grid.BackColor
    If mousCol = gf2004 Or mousCol = gf2003 Or mousCol = gf2002 Or mousCol = gf2001 Then
        SortCol Grid, mousCol, "numeric"
    Else
        SortCol Grid, mousCol
    End If
    Grid.row = 1    ' только чтобы снять выделение
    Grid_EnterCell
End If

End Sub

Private Sub Grid_DblClick()
If Grid.CellBackColor = vbYellow Then Exit Sub

gFirmId = Grid.TextMatrix(mousRow, gfId)

If mousCol = gfKategor Then
    listBoxInGridCell lbKP, Grid
ElseIf mousCol = gfM Then
    listBoxInGridCell lbM, Grid
ElseIf mousCol = gfLogin Or mousCol = gfPass Then
' редакти-ся по <F4>
Else
    textBoxInGridCell tbMobile, Grid
End If
End Sub

Private Sub Grid_EnterCell()
If noClick Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col
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
    tbInform.MaxLength = 100
    tbMobile.MaxLength = 100
ElseIf mousCol = gfTlf Or mousCol = gfSale Or mousCol = gfAdres Or _
mousCol = gfFax Or mousCol = gfEmail Or mousCol = gfLevel Or _
mousCol = gfKatalog Then
    tbInform.MaxLength = 50
    tbMobile.MaxLength = 50
Else
    tbInform.MaxLength = 10
    tbMobile.MaxLength = 10
End If
tbInform.Text = Grid.TextMatrix(mousRow, mousCol)

If mousCol = gfId Or mousCol = gfLogin Or mousCol = gfPass Or _
mousCol = gf2003 Or mousCol = gf2001 Or mousCol = gf2002 Or Grid.TextMatrix(mousRow, gfId) = "" Then
    Grid.CellBackColor = vbYellow
    tbInform.Locked = True
Else
    Grid.CellBackColor = &H88FF88  ' бл.зел
'   Grid.CellBackColor = &H8888FF    ' бл.кр
    If mousCol = gfM Or mousCol = gfKategor Then
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

Private Sub lbKP_DblClick()

ValueToTableField "##66", "'" & lbKP.Text & "'", "GuideFirms", "Kategor", "byFirmId"
Grid.TextMatrix(mousRow, gfKategor) = lbKP.Text
lbHide
cep = False
cmAdd.Enabled = True
cmExit.Enabled = True
cbM.Enabled = True
cmExel.Enabled = True

End Sub

Private Sub lbKP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbKP_DblClick
End Sub

Private Sub lbM_DblClick()
Dim str As String

If lbM.ListIndex = 0 Then
    str = "14"  'not'
Else
    str = manId(lbM.ListIndex - 1)
End If
ValueToTableField "##66", str, "GuideFirms", "ManagId", "byFirmId"
Grid.TextMatrix(mousRow, gfM) = lbM.Text
lbHide
If cep Then
   Grid.col = gfKategor: mousCol = gfKategor
   listBoxInGridCell lbKP, Grid
End If
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
'$odbc08!$
Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, I As Integer, strId As String
If KeyCode = vbKeyReturn Then
 str = trimAll(tbMobile.Text)
 gFirmId = Grid.TextMatrix(mousRow, gfId)

 If mousCol = gfNazwFirm Then
   strId = Grid.TextMatrix(mousRow, gfId)
   On Error GoTo ERR1
   If strId = "" Then
    wrkDefault.BeginTrans
    sql = "update guidefirms set firmId = firmId where firmId = 0"
    myBase.Execute sql
    
    sql = "select max(firmid) from guidefirms;"
    If Not byErrSqlGetValues("##50", sql, gFirmId) Then GoTo ERR1
    gFirmId = gFirmId + 1
    
    sql = "insert into guidefirms (firmId, name) values (" & gFirmId & ", '" & str & "')"
    I = myExecute("##50", sql, -196)
    If I <> 0 Then GoTo ERR0:
    wrkDefault.CommitTrans
    
    Grid.TextMatrix(mousRow, gfId) = gFirmId
    quantity = quantity + 1
    cep = True ' запускаем цепь посл.ввода tot 2х полей
    cmAdd.Enabled = False
    cmExit.Enabled = False
    cbM.Enabled = False
    cmExel.Enabled = False
    Grid.TextMatrix(mousRow, mousCol) = str
    lbHide
    Grid.col = gfM: mousCol = gfM
    listBoxInGridCell lbM, Grid
    Exit Sub
   Else
    sql = "UPDATE GuideFirms SET Name = '" & str & _
    "' WHERE (((FirmId)=" & strId & "));"
'    MsgBox sql
    I = myExecute("##356", sql, -196)
    If I <> 0 Then GoTo ERR0:
   End If
   On Error GoTo 0
 ElseIf mousCol = gfSale Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Sale", "byFirmId"
 ElseIf mousCol = gfFIO Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "FIO", "byFirmId"
 ElseIf mousCol = gfTlf Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Phone", "byFirmId"
 ElseIf mousCol = gfFax Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Fax", "byFirmId"
 ElseIf mousCol = gfEmail Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Email", "byFirmId"
 ElseIf mousCol = gfAdres Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Address", "byFirmId"
 ElseIf mousCol = gfAtr1 Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Atr1", "byFirmId"
 ElseIf mousCol = gfAtr2 Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Atr2", "byFirmId"
 ElseIf mousCol = gfAtr3 Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Atr3", "byFirmId"
 ElseIf mousCol = gfLogin Then
    If str <> "" And str <> Grid.TextMatrix(mousRow, gfLogin) Then
        If existValueInTableFielf(str, "GuideFirms", "xLogin") Then '$#$
            MsgBox "Такой логин уже есть.", , "Недопустимое значение"
            tbMobile.SelStart = Len(tbMobile.Text)
            tbMobile.SetFocus
            Exit Sub
        End If
    End If
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "xLogin", "byFirmId"
 ElseIf mousCol = gfPass Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Pass", "byFirmId"
 ElseIf mousCol = gfLevel Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Level", "byFirmId"
 ElseIf mousCol = gfType Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Type", "byFirmId"
 ElseIf mousCol = gfKatalog Then
    ValueToTableField "##66", "'" & str & "'", "GuideFirms", "Katalog", "byFirmId"
End If
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


Private Sub Timer1_Timer()

Timer1.Enabled = False
 If Regim = "fromMenu" Then  'по F11 из Orders
    tbFind.SetFocus
 Else ' из контекстного меню
    cmFind.Caption = "Поиск"
    If findValInCol(Grid, tbFind.Text, gfNazwFirm) Then
        Grid.SetFocus
    Else
        tbFind.SetFocus
    End If
    Grid.SetFocus
 End If

End Sub
