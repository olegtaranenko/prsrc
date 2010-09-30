VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FirmComtex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Выбор фирмы плательщика из бух.базы Комтех"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11895
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Exel"
      Height          =   315
      Left            =   9180
      TabIndex        =   2
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox tbInform 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   6495
   End
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   7560
   End
   Begin VB.TextBox tbMobile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "tbMobile"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   7680
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmFind 
      Caption         =   "Поиск"
      Height          =   360
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox tbFind 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmSel 
      Caption         =   "Выбрать"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   11100
      TabIndex        =   8
      Top             =   7680
      Width           =   675
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6915
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   12197
      _Version        =   393216
      MergeCells      =   2
      AllowUserResizing=   1
      FormatString    =   " "
   End
   Begin VB.Label laQuant 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   315
      Left            =   5520
      TabIndex        =   10
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Число записей:"
      Height          =   195
      Left            =   4260
      TabIndex        =   7
      Top             =   7740
      Width           =   1215
   End
End
Attribute VB_Name = "FirmComtex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mousRow As Long    '
Public mousCol As Long    '
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim quantity As Integer 'количество найденных фирм
Dim findBy As Long ' искать по полю
Dim lastSortField As Long  ' столбец сортировки
Dim lastSortDir As Boolean ' true - по убыванию, false - по возрастанию
Public serverName As String


Private Sub cmAdd_Click()
    If Grid.TextMatrix(Grid.Rows - 1, fcFirmName) <> "" Then Grid.AddItem ("")
    
    Grid.row = Grid.Rows - 1
    Grid.col = fcFirmName 'название
    Grid.SetFocus
    textBoxInGridCell tbMobile, Grid
End Sub


Private Sub cmExel_Click()
'GridToExcel Grid, "Справочник сторонних организаций (" & cbM.Text & ")"

End Sub

Private Sub cmExit_Click()
    Unload Me
End Sub
Private Function getFieldFind(fld As Long)

    If fld = fcInn Then
        getFieldFind = "p.inn"
    ElseIf fld = fcOkonx Then
        getFieldFind = "p.okonx"
    ElseIf fld = fcOkpo Then
        getFieldFind = "p.okpo"
    ElseIf fld = fcKpp Then
        getFieldFind = "p.kpp"
    ElseIf fld = fcAddress Then
        getFieldFind = "v.address"
    ElseIf fld = fcPhone Then
        getFieldFind = "v.phone"
    Else
        getFieldFind = "v.nm"
    End If

End Function


Public Sub cmFind_Click()
    loadGuide , tbFind.Text, getFieldFind(findBy)
End Sub

Sub lbHide(Optional noGrid As String)
    tbMobile.Visible = False
    Grid.Enabled = True
    If noGrid <> "" Then Exit Sub
    Grid.SetFocus
    Grid_EnterCell
End Sub

Sub loadGuide( _
      Optional p_id_show As Integer _
    , Optional match As String _
    , Optional matchField As String _
)

Dim I As Long, strWhere As String, strOrder As String, strField As String
Dim orderField As String


    Me.MousePointer = flexHourglass
    Grid.Visible = False
    clearGrid Grid
    'strWhere = trimAll(tbFind.Text)
    'MsgBox "strWhere = " & strWhere
    quantity = 0
    
    If Not (IsNull(p_id_show) Or IsMissing(p_id_show) Or p_id_show = 0) Then
        strWhere = _
            "   v.id = " & p_id_show
    ElseIf Not (IsNull(match) Or IsMissing(match) Or match = "") Then
        If Not (IsNull(matchField) Or IsMissing(matchField) Or matchField = "") Then
            strField = matchField
        Else
            strField = "v.nm"
        End If
            strWhere = _
        strField & " like ('%" & match & "%')"
    End If
    
    If Not (IsNull(orderField) Or IsMissing(orderField) Or orderField = "") Then
        strOrder = " order by " & orderField
    Else
        strOrder = " order by v.nm "
    End If
    
    sql = _
          "select " _
        & "  v.id      as id" _
        & " ,v.nm      as FirmName " _
        & " ,p.inn     as inn" _
        & " ,p.okonx   as okonx" _
        & " ,p.okpo    as okpo" _
        & " ,p.kpp     as kpp" _
        & " ,v.address as address" _
        & " ,v.phone   as phone" _
        & " from voc_names_" & serverName & " v" _
        & " join post_" & serverName & " p on p.id = v.id"
    
    sql = sql _
        & " where v.id > 0"
    
    If Len(strWhere) > 0 Then
        sql = sql _
            & " and " & strWhere
    End If
    
    orderField = getFieldFind(lastSortField)
    If (lastSortDir) Then
        orderField = orderField & " desc"
    End If
    
    sql = sql _
        & strOrder
    
    
    
    'Debug.Print sql
    'Set tbFirms = myOpenRecordSet("##15.1", sql, dbOpenForwardOnly)
    On Error GoTo sqlex
    Set tbFirms = myBase.Connection.OpenRecordset(sql, dbOpenForwardOnly, dbExecDirect, dbPessimistic)
    If tbFirms Is Nothing Then GoTo EN1
    
    If Not tbFirms.BOF Then
      'tbFirms.MoveFirst
      While Not tbFirms.EOF
        If tbFirms!id = 0 Then GoTo AA
        quantity = quantity + 1
        Grid.TextMatrix(quantity, fcId) = tbFirms!id
        Grid.TextMatrix(quantity, fcFirmName) = tbFirms!FirmName
        fieldToCol tbFirms!Inn, fcInn
        fieldToCol tbFirms!Okonx, fcOkonx
        fieldToCol tbFirms!Okpo, fcOkpo
        fieldToCol tbFirms!Kpp, fcKpp
        fieldToCol tbFirms!Address, fcAddress
        fieldToCol tbFirms!Phone, fcPhone
        Grid.AddItem ("")
AA:     tbFirms.MoveNext
      Wend
      If quantity > 0 Then Grid.RemoveItem (quantity + 1)
    End If
    tbFirms.Close
    wrkDefault.CommitTrans
    
EN1:
    Grid.Visible = True
    laQuant.Caption = quantity
    Me.MousePointer = flexDefault
    Exit Sub
    
sqlex:
    If Not errorCodAndMsg("r_list_customer") Then
        GoTo EN1
    End If
    
End Sub

Sub fieldToCol(field As Variant, col As Long)
If Not IsNull(field) Then Grid.TextMatrix(quantity, col) = field
End Sub
Private Sub updateRemote(info As String _
    , tableName As String _
    , fieldName As String _
    , Value As Variant _
    , condititon As String _
    , p_ServerName As String _
)
    On Error GoTo sqle
    sql = "call update_remote ('" & p_ServerName & "', " _
        & "'" & tableName & "'" _
        & ", '" & fieldName & "'" _
        & ","
    If IsNumeric(Value) Then
        sql = sql & _
            Value
    Else
        sql = sql & _
            "'''''" & Value & "'''''"
    End If
    
    ' условие
    sql = sql _
       & ", '" & condititon & "')"
    Debug.Print sql
     myBase.Execute sql
     wrkDefault.CommitTrans
     Exit Sub
sqle:
    errorCodAndMsg info
    Resume Next
End Sub
Private Sub cmSel_Click()
    makeBillChoice Grid.TextMatrix(mousRow, fcId), serverName
    Unload Me
End Sub

Public Sub makeBillChoice(p_id_bill As String, p_ServerName As String)
Dim id_jscet As String

    'Orders.Grid.Text = Grid.Text
    'wrkDefault.BeginTrans
    On Error GoTo fail
    id_jscet = getValueFromTable("Orders", "id_jscet", "numOrder = " & gNzak)
    wrkDefault.Rollback
    
    updateRemote "Выбор плательщика", "jscet", "id_d" _
            , p_id_bill _
            , "id = " & id_jscet _
            , p_ServerName
    updateRemote "Выбор грузополучателя", "jscet", "id_d_cargo" _
            , p_id_bill _
            , "id = " & id_jscet _
            , p_ServerName
    Orders.Grid.TextMatrix(Orders.Grid.row, orBillId) = p_id_bill
    Orders.Grid.CellForeColor = vbRed
    refreshTimestamp gNzak
    wrkDefault.CommitTrans
    GoTo finally
fail:
    MsgBox "Ошибка при выборе плательщика" & vbCr & "", , "Сообщите администратору"
    wrkDefault.Rollback
finally:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape And Not cep Then lbHide
End Sub



Private Sub Form_Load()

Dim I As Integer
Dim v_id_bill As String
Dim id_voc_names As String
Dim v_id_show As Integer
Dim match_def As String

serverName = Orders.Grid.TextMatrix(Orders.mousRow, orServername)

If IsEmpty(serverName) Or serverName = "" Then
    Exit Sub
End If

Me.Caption = "Выбор фирмы плательщика из базы """ & Orders.Grid.TextMatrix(Orders.mousRow, orVenture) & """"
quantity = 0

Grid.FormatString = fcFormatString

Grid.ColWidth(fcId) = 0
Grid.ColWidth(fcFirmName) = 3000
Grid.ColWidth(fcInn) = 1100
Grid.ColWidth(fcOkonx) = 800
Grid.ColWidth(fcOkpo) = 800
Grid.ColWidth(fcKpp) = 800
Grid.ColWidth(fcAddress) = 3200
Grid.ColWidth(fcPhone) = 1400



'    v_id_bill = Orders.Grid.TextMatrix(Orders.mousRow, orBillId)
    v_id_bill = Orders.g_id_bill
    Orders.g_id_bill = ""
    id_voc_names = Orders.Grid.TextMatrix(Orders.mousRow, orVocnameId)
    If v_id_bill = "" And id_voc_names = "" Then
        match_def = "1"
    End If
    If Not IsEmpty(v_id_bill) And v_id_bill <> "" Then
        loadGuide CInt(v_id_bill)
'        v_id_show = CInt(v_id_bill)
    ElseIf (Not IsEmpty(id_voc_names) And id_voc_names <> "") Or match_def <> "" Then
        match_def = Mid(Orders.Grid.TextMatrix(Orders.mousRow, orFirma), 1, 3)
        tbFind.Text = match_def
        tbFind.SelStart = 0
        tbFind.SelLength = Len(tbFind.Text)
        tbFind.Enabled = True
'        tbFind.SetFocus
        
        
        loadGuide , match_def
        'v_id_show = CInt(id_voc_names)
    End If
    
    If Not IsEmpty(v_id_show) Then
        'tbFind.Text = v_id_show
    End If
    
'    tbFind.SelStart = 0
'    tbFind.SelLength = Len(tbFind.Text)
'    cmSel.Visible = True
    cmAdd.Visible = True

'Timer1.Interval = 100
'Timer1.Enabled = True

oldHeight = Me.Height
oldWidth = Me.Width
End Sub

Private Sub Form_Resize()
Dim H As Integer, W As Integer

If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next
lbHide "noGrid"
W = Me.Width - oldWidth
oldWidth = Me.Width
H = Me.Height - oldHeight
oldHeight = Me.Height


Grid.Height = Grid.Height + H
Grid.Width = Grid.Width + W
cmSel.Top = cmSel.Top + H
cmExit.Top = cmExit.Top + H
'cmExit.Left = cmExit.Left + w
'tbFind.Top = tbFind.Top + h
'cmFind.Top = cmFind.Top + h
cmAdd.Top = cmAdd.Top + H
cmExel.Top = cmExel.Top + H
cmExel.Left = cmExel.Left + W
End Sub

Private Sub Grid_Click()
    mousCol = Grid.MouseCol
    mousRow = Grid.MouseRow
    If quantity = 0 Then Exit Sub
    If Grid.MouseRow = 0 Then
        Grid.CellBackColor = Grid.BackColor
        If mousCol = 0 Then Exit Sub
        SortCol Grid, mousCol
        Grid.row = 1    ' только чтобы снять выделение
        Grid_EnterCell
    End If

End Sub

Private Sub Grid_DblClick()
    If Grid.MouseRow = 0 Then Exit Sub
    textBoxInGridCell tbMobile, Grid
End Sub

Private Sub Grid_EnterCell()
If noClick Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col

If mousCol = fcFirmName Then
    tbMobile.MaxLength = 98
ElseIf mousCol = fcInn Then
    tbMobile.MaxLength = 14
ElseIf mousCol = fcOkonx Then
    tbMobile.MaxLength = 5
ElseIf mousCol = fcOkpo Then
    tbMobile.MaxLength = 10
ElseIf mousCol = fcKpp Then
    tbMobile.MaxLength = 10
ElseIf mousCol = fcAddress Then
    tbMobile.MaxLength = 98
ElseIf mousCol = fcPhone Then
    tbMobile.MaxLength = 37
End If
tbInform.MaxLength = tbMobile.MaxLength


tbInform.Text = Grid.TextMatrix(mousRow, mousCol)

    Grid.CellBackColor = &H88FF88  ' бл.зел

End Sub

Private Sub Grid_GotFocus()
Grid_EnterCell
End Sub

Private Function mapColToRussianName(fc As Long)
    If fc = fcFirmName Then
        mapColToRussianName = "фирме"
    ElseIf fc = fcInn Then
        mapColToRussianName = "ИНН"
    ElseIf fc = fcOkonx Then
        mapColToRussianName = "ОКОНХ"
    ElseIf fc = fcOkpo Then
        mapColToRussianName = "ОКПО"
    ElseIf fc = fcKpp Then
        mapColToRussianName = "КПП"
    ElseIf fc = fcAddress Then
        mapColToRussianName = "адресу"
    ElseIf fc = fcPhone Then
        mapColToRussianName = "тел-ну"
    End If
End Function


Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And Shift = 2 Then
    cmSel_Click
ElseIf KeyCode = vbKeyReturn Then
    Grid_DblClick
ElseIf KeyCode = vbKeyF7 Then
    cmFind.Caption = "Поиск по " & mapColToRussianName(mousCol)
    findBy = mousCol
    tbFind.SetFocus
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If
End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
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
Dim nvalue As String, I As Integer, strId As String
Dim id_voc_names As Integer
Dim id_jscet As Integer
Dim id_belong As Integer

If KeyCode = vbKeyReturn Then
    nvalue = trimAll(tbMobile.Text)
    If nvalue = "" Then
        Exit Sub
    End If
    
    If Grid.TextMatrix(mousRow, fcId) <> "" Then
        id_voc_names = CInt(Grid.TextMatrix(mousRow, fcId))
    End If
    
    On Error GoTo fail
    id_jscet = getValueFromTable("Orders", "id_jscet", "numOrder = " & gNzak)
    wrkDefault.Rollback
    
    If mousCol = fcFirmName Then
        strId = Grid.TextMatrix(mousRow, fcId)
        On Error GoTo ERR1
        If strId = "" Then
            wrkDefault.BeginTrans

            id_belong = getValueFromTable("FirmGuide", "id_voc_names", "firmid = 0")
            
            sql = "select insert_count_remote('" _
                & serverName _
                & "', 'voc_names', 'nm, belong_id', '''''" _
                & nvalue & "'''', " & id_belong & "')"

            byErrSqlGetValues "add firm", sql, id_voc_names
            
            sql = "select insert_remote('" _
                & serverName _
                & "', 'post', 'id', '" _
                & id_voc_names & "')"
                
            myBase.Execute sql
            
            wrkDefault.CommitTrans
            
            Grid.TextMatrix(mousRow, fcId) = id_voc_names
            quantity = quantity + 1
            GoTo AA
        Else
            updateRemote "Имя плательщика", "voc_names", "nm", nvalue, "id = " & id_voc_names, serverName

        End If
    ElseIf mousCol = fcInn Then
        updateRemote "ИНН", "post", "inn", nvalue, "id = " & id_voc_names, serverName
    ElseIf mousCol = fcOkonx Then
        updateRemote "ОКОНХ", "post", "okonx", nvalue, "id = " & id_voc_names, serverName
    ElseIf mousCol = fcOkpo Then
        updateRemote "ОКПО", "post", "okpo", nvalue, "id = " & id_voc_names, serverName
    ElseIf mousCol = fcKpp Then
        updateRemote "КПП", "post", "kpp", nvalue, "id = " & id_voc_names, serverName
    ElseIf mousCol = fcAddress Then
        updateRemote "Адрес", "voc_names", "address", nvalue, "id = " & id_voc_names, serverName
    ElseIf mousCol = fcPhone Then
        updateRemote "Телефон", "voc_names", "phone", nvalue, "id = " & id_voc_names, serverName
    End If
AA:
    Grid.TextMatrix(mousRow, mousCol) = nvalue
    lbHide
End If

Exit Sub
fail:
    wrkDefault.Rollback
    GoTo finally

ERR0:
    If I = -2 Then
        MsgBox "Такая фирма уже есть", , "Ошибка!"
        tbMobile.SetFocus
    Else
finally:
ERR1:
        lbHide
    End If

End Sub


