VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form VentureHistory 
   Caption         =   "История номенклатуры по предприятиям"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ckCumulative 
      Caption         =   "Сводные"
      Height          =   255
      Left            =   10560
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   2280
      TabIndex        =   14
      Text            =   "tbMobile"
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   10680
      TabIndex        =   13
      Top             =   5880
      Width           =   915
   End
   Begin VB.ListBox lbVenture 
      Height          =   255
      ItemData        =   "VentureHistory.frx":0000
      Left            =   4320
      List            =   "VentureHistory.frx":0002
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmExcel 
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CheckBox ckPerList 
      Caption         =   "В целых"
      Height          =   255
      Left            =   3780
      TabIndex        =   11
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "Печать"
      Height          =   315
      Left            =   9480
      TabIndex        =   10
      Top             =   5880
      Width           =   915
   End
   Begin VB.CommandButton cmLoad 
      Caption         =   "Загрузить"
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox tbStartDate 
      Height          =   285
      Left            =   1620
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "01.01.01"
      Top             =   60
      Width           =   795
   End
   Begin VB.TextBox tbEndDate 
      Height          =   285
      Left            =   2700
      MaxLength       =   8
      TabIndex        =   1
      Top             =   60
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
      _Version        =   393216
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      Enabled         =   -1  'True
      MergeCells      =   1
      AllowUserResizing=   1
   End
   Begin VB.Label laQuant 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2460
      TabIndex        =   9
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Число записей:"
      Height          =   195
      Left            =   1140
      TabIndex        =   8
      Top             =   5940
      Width           =   1215
   End
   Begin VB.Label laBegin 
      Caption         =   "Установите Период загрузки и нажмите <Загрузить>"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   2880
      Width           =   4155
   End
   Begin VB.Label laPeriod 
      Caption         =   "Период Загрузки  с  "
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label laPo 
      Caption         =   "по"
      Height          =   195
      Left            =   2460
      TabIndex        =   5
      Top             =   120
      Width           =   195
   End
   Begin VB.Menu mnPopup 
      Caption         =   "mnPopup"
      Visible         =   0   'False
      Begin VB.Menu mnZachet 
         Caption         =   "Компенсировать через взаимозачет"
      End
   End
   Begin VB.Menu mpVenture 
      Caption         =   "mpVenture"
      Visible         =   0   'False
      Begin VB.Menu mnDelete 
         Caption         =   "Удалить взаимозачет"
      End
      Begin VB.Menu mnDateChange 
         Caption         =   "Изменить дату зачета"
      End
      Begin VB.Menu mnDmcMove 
         Caption         =   "Перенести в другую накладную"
      End
   End
End
Attribute VB_Name = "VentureHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public isLoad As Boolean
Public nomenkName As String
Public Regim As String
Public quantity  As Long
Public DMCnomNomCur As String ' текущая номенклатура в групповой карте

Dim ventureIds() As Integer
Dim ventureRests() As Single
Dim ventureAbbrevs() As String

Dim defaultVentureIndex As Integer

Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim mousCol As Long, mousRow As Long

'Столбцы 0, 1, 2 скрыты
Const ktNomnom = 0
Const ktSourId = 1
Const ktDestId = 2
Const ktDate = 3
Const ktDocNum = 4
Const ktOperation = 5
Const ktQty = 6
Const ktVenture = 7
Const ktOstat = 8

Sub lbHide()
    lbVenture.Visible = False
    tbMobile.Visible = False
    Grid.Enabled = True
    On Error Resume Next
    Grid.SetFocus
    Grid_EnterCell
End Sub


Private Function orderIsVentureOrder(rowIndex As Integer) As Boolean
    orderIsVentureOrder = False
    If InStr(Grid.TextMatrix(rowIndex, ktOperation), "Взаимозачет с") = 1 Then
        orderIsVentureOrder = True
    End If
    
End Function

Private Sub ckCumulative_Click()
    If Not isLoad Then Exit Sub
    fillGrid
End Sub

Private Sub ckPerList_Click()
    If Not isLoad Then Exit Sub
    fillGrid
End Sub

Private Sub cmExcel_Click()
    GridToExcel Grid, "Карта движения по номенклатуре '" & gNomNom & "' по предприятиям"
End Sub

Private Sub cmExit_Click()
    Unload Me
End Sub
Public Sub fillGrid()
Dim i As Integer, tmpTopRow As Long

    tmpTopRow = Grid.TopRow
    
    Me.MousePointer = flexHourglass
    Grid.Visible = False
    quantity = 0
    For i = 1 To UBound(DMCnomNom())
        getVentureHistory DMCnomNom(i)
    Next i
    If tmpTopRow < Grid.Rows Then
        Grid.TopRow = tmpTopRow
    End If
    Grid.Visible = True
    Me.MousePointer = flexDefault

End Sub

Private Sub cmLoad_Click()
    fillGrid
End Sub

Private Function getOperation( _
                sourId, destId _
                , activeOper As Integer _
                , ventureId, destVentureId As Integer _
                , srcName As String, dstName As String _
                , invoice, firmName _
            )
    If destId = -8 Then
        ' Продажа
        getOperation = dstName & ": " & firmName & " №" & invoice
        Exit Function
    End If
    
    If destId = -7 Then
        ' Инвентаризация
        getOperation = "Инвентаризации: РАСХОД ": Exit Function
    End If
    
    If sourId = 34 Then
        ' Инвентаризация
        getOperation = "Инвентаризации: ПРИХОД ": Exit Function
    End If
    
    If ventureId > 0 And destVentureId > 0 Then
        getOperation = "Взаимозачет с " & getVentureNameById(destVentureId)
        Exit Function
    End If
    
    If sourId <= -1001 And destId <= -1001 Then
        getOperation = "Внутреннее перемещение (межсклад)"
    ElseIf sourId <= -1001 Then
        getOperation = "Расход в '" & dstName & "'"
        If Not IsNull(firmName) Then
            getOperation = getOperation & " (" & firmName
            If invoice <> "счет ?" Then
                getOperation = getOperation & " №" & invoice
            End If
            getOperation = getOperation & ")"
        End If
        
    ElseIf destId <= -1001 Then
        getOperation = "Приход от '" & srcName & "'"
    End If
    
End Function

Private Sub showRest(dstRow As Long)
Dim totalRest As Single
Dim i As Integer

    totalRest = 0
    For i = 1 To UBound(ventureIds)
        If (Round(ventureRests(i), 3) < 0) Then
            Grid.col = ktOstat + i
            Grid.row = dstRow
            Grid.CellForeColor = vbRed
        End If
        Grid.TextMatrix(dstRow, ktOstat + i) = Format(ventureRests(i), "#0.00")
        totalRest = totalRest + ventureRests(i)
    Next i
    Grid.TextMatrix(dstRow, ktOstat) = Format(totalRest, "#0.00")
End Sub

Public Sub getVentureHistory(nNom As String)
Dim ed_izm As String, ed_izm2 As String, v_perList As Single
Dim v_cod As String, v_size As String, v_nameDisplay As String
Dim dstRow As Long
Dim i As Integer
Dim totalRest As Single
Dim restrictDate As Variant
Dim cumulative_predikat As String
Dim historyStart As Variant
Dim firstEntry As Boolean
Dim doShowRow As Boolean


    
    firstEntry = False
    
    If isDateEmpty(tbEndDate) Then
        restrictDate = "convert(date, '" & Format(tbEndDate.Text, "yyyymmdd") & "')"
    Else
        restrictDate = Null
    End If
        
    If isDateEmpty(tbStartDate, False) Then
        historyStart = CDate(tbStartDate.Text)
    Else
        historyStart = Null
    End If
    
    If ckCumulative.value = 0 Then
        cumulative_predikat = "not"
    End If


    sql = "SELECT sGuideNomenk.nomName, " & _
    "sGuideNomenk.ed_Izmer , sGuideNomenk.ed_Izmer2, sGuideNomenk.perList " _
    & ",  cod, size" _
    & " " _
    & "From sGuideNomenk WHERE (((sGuideNomenk.nomNom)='" & nNom & "'));"
    
    byErrSqlGetValues "##132", sql, nomenkName, ed_izm, ed_izm2, v_perList, v_cod, v_size
    
    If ckPerList.value = 0 Then
        v_perList = 1
    End If
    
    If quantity = 0 Then clearGrid1 Grid
    Grid.AddItem vbTab & vbTab & vbTab & nNom
    dstRow = Grid.Rows - 1
    Grid.MergeRow(dstRow) = True
    'Высота заголовока
    Grid.RowHeight(dstRow) = 320
    Grid.row = dstRow
    'Grid.col = 3: Grid.CellAlignment = flexAlignRightCenter
    Grid.col = ktDocNum: Grid.CellAlignment = flexAlignRightCenter
    Grid.col = ktOperation: Grid.CellAlignment = flexAlignRightCenter
    'Grid.col = 6: Grid.CellAlignment = flexAlignRightCenter
    If Not IsNull(v_cod) Then
        v_nameDisplay = v_cod & " "
    End If
    v_nameDisplay = v_nameDisplay & nomenkName
    If Not IsNull(v_size) Then
        v_nameDisplay = v_nameDisplay & " " & v_size
    End If
     
    
    Grid.TextMatrix(dstRow, ktDocNum) = v_nameDisplay
    Grid.TextMatrix(dstRow, ktOperation) = v_nameDisplay
    
    ' Единица измерения
    'Grid.col = ktQty: Grid.CellFontBold = False: Grid.CellAlignment = flexAlignRightCenter
    If (ckPerList = 1) Then
        Grid.TextMatrix(dstRow, ktQty) = ed_izm2
    Else
        Grid.TextMatrix(dstRow, ktQty) = ed_izm
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = _
        "select " _
        & "     if destId <= -1001 then 0 else 2 endif " _
        & "         as sec_sort" _
        & "     , convert(date, xDate) as xDate" _
        & "     , convert(varchar(10), n.numdoc) + if m.numext < 254 then '/' + convert(varchar(3), m.numext) endif" _
        & "         as numdoc" _
        & "     , n.sourid, n.destId" _
        & "     , quant as qty" _
        & "     , if (n.sourid <= -1001 and n.destid <= -1001) then " _
        & "         0 " _
        & "         else " _
        & "             if n.destid <= -1001 then " _
        & "                 1" _
        & "             else" _
        & "                 -1" _
        & "             endif" _
        & "         endif " _
        & "     as activeOper" _
        & "     , isnull(o.invoice, bo.invoice) as invoice" _



    sql = sql _
        & " , if (n.sourid <= -1001 and n.destid <= -1001) then " _
        & "     null " _
        & " else " _
        & "     if n.destid <= -1001 then " _
        & "         isnull(n.ventureid, v.ventureid) " _
        & "     else " _
        & "      isnull(" _
        & "         isnull(" _
        & "             isnull(o.ventureid, bo.ventureid)" _
        & "             , if substring(isnull(o.invoice, bo.invoice), 1, 2) = '55' then 2 else 1 endif)" _
        & "         , v.ventureid) " _
        & "     endif" _
        & " endif ventureid" _
        & " , 0 as destVentureId" _
        & " , s.sourceName as srcName, d.sourceName as dstName" _
        & " , isnull(o.numorder, bo.numorder) as numorder, isnull(f.name, bf.name) as firmName"

    sql = sql _
        & " from sdocs n" _
        & "     join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext " _
        & "     join sguidesource s on s.sourceId = n.sourId" _
        & "     join sguidesource d on d.sourceId = n.destId" _
        & "     join system sys on 1 = 1" _
        & "     join guideventure v on v.id_analytic = sys.id_analytic_default" _
        & "     left join orders o on o.numorder = n.numdoc" _
        & "     left join bayOrders bo on bo.numorder = n.numdoc" _
        & "     left join GuideFirms f on f.firmid = o.firmid" _
        & "     left join BayGuideFirms bf on bf.firmid = bo.firmid" _
        & " where" _
        & "     m.nomnom = '" & nNom & "'"
        
    If Not IsNull(restrictDate) Then
        sql = sql & " and convert(date, n.xDate) <= " & restrictDate
    End If
    
    sql = sql _
        & "         union" _
        & "     select 1, " _
        & "         convert(date, n.nDate) as xDate" _
        & "         , convert(varchar(10), n.id)" _
        & "         , null, null" _
        & "         , m.quant  as qty" _
        & "         , 0 as income" _
        & "         , null as invoice" _
        & "         , srcVentureId, dstVentureId" _
        & "         , '', ''" _
        & "         , null, null" _
        & " from sdmcventure m" _
        & " join sdocsventure n on m.sdv_id = n.id and cumulative_id is " & cumulative_predikat & " null "
        
    If Not IsNull(restrictDate) Then
        sql = sql & " and convert(date, n.nDate) <= " & restrictDate
    End If
        
    sql = sql _
        & "     where m.nomnom = '" & nNom & "'" _
        & "     order by 2, 1"


    Set tbDMC = myOpenRecordSet("##130", sql, dbOpenForwardOnly)

    If tbDMC Is Nothing Then Exit Sub
    'Grid.Visible = False
    
    resetRests
    
    While Not tbDMC.EOF
        
        If Not IsNull(historyStart) Then
            If tbDMC!xDate < historyStart Then
                doShowRow = False
            Else
                doShowRow = True
                If Not firstEntry Then
                    firstEntry = True
                    showRest Grid.Rows - 1
                End If
            End If
        Else
            doShowRow = True
        End If
        
    
        If tbDMC!ventureId > 0 And tbDMC!destVentureId > 0 Then
            orderVentureRest tbDMC!qty, tbDMC!ventureId, tbDMC!destVentureId, v_perList
        ElseIf tbDMC!sourId > -1001 Or tbDMC!destId > -1001 Then
            orderRest tbDMC!qty, tbDMC!ventureId, tbDMC!activeOper, v_perList
        End If
        
        If doShowRow Then
            Grid.AddItem nNom & vbTab & tbDMC!sourId & vbTab & tbDMC!destId
            dstRow = Grid.Rows - 1
            Grid.TextMatrix(dstRow, ktDate) = Format(tbDMC!xDate, "dd.mm.yy")
            Grid.TextMatrix(dstRow, ktDocNum) = tbDMC!numDoc
            Grid.TextMatrix(dstRow, ktQty) = Format(Round(tbDMC!qty / v_perList, 2), "#0.00")
            
            Grid.TextMatrix(dstRow, ktOperation) = _
                getOperation( _
                      tbDMC!sourId, tbDMC!destId _
                    , tbDMC!activeOper _
                    , tbDMC!ventureId, tbDMC!destVentureId _
                    , tbDMC!srcName, tbDMC!dstName _
                    , tbDMC!invoice, tbDMC!firmName _
                )
            Grid.TextMatrix(dstRow, ktVenture) = getVentureNameById(tbDMC!ventureId)
    
    
            showRest dstRow
    
            quantity = quantity + 1
        End If
        tbDMC.MoveNext
    Wend

    tbDMC.Close
    
    laQuant.Caption = quantity
    isLoad = True
    
End Sub

Private Sub resetRests()
Dim i As Integer

    For i = 1 To UBound(ventureIds)
        ventureRests(i) = 0
    Next i
    
End Sub
Private Function vlIndex(ventureId)
Dim i As Integer

    For i = 1 To UBound(ventureIds)
        If ventureId = ventureIds(i) Then
            vlIndex = i
            Exit Function
        End If
    Next i
    vlIndex = defaultVentureIndex
    
End Function


Private Sub orderVentureRest(ByVal qty, srcventureId, dstVentureId, perList)
Dim i As Integer

    qty = qty / perList
    i = vlIndex(srcventureId)
    ventureRests(i) = ventureRests(i) - qty
    i = vlIndex(dstVentureId)
    ventureRests(i) = ventureRests(i) + qty
End Sub
            
            
Private Sub orderRest(ByVal qty, ventureId, activeOper, perList)
Dim i As Integer
    i = vlIndex(ventureId)
    ventureRests(i) = ventureRests(i) + qty / perList * activeOper
End Sub



Private Function getVentureNameById(destVentureId)
Dim v_ventureId As Integer
Dim i As Integer

    i = vlIndex(destVentureId)
    getVentureNameById = lbVenture.List(i - 1)
    Exit Function

    If IsNull(destVentureId) Then
        v_ventureId = -1
    Else
        v_ventureId = destVentureId
    End If
    
   
    For i = 0 To lbVenture.ListCount - 1
        If lbVenture.ItemData(i) = v_ventureId Then
            getVentureNameById = lbVenture.List(i)
            Exit Function
        End If
    Next i
    
End Function

Private Sub cmPrint_Click()
    Me.PrintForm
End Sub


Private Sub Form_Activate()

 '   If UBound(DMCnomNom) = 1 Then
        Me.Caption = "Карточка движения по предприятиям " & _
                       " "
  '  Else
'        Me.Caption = "Карточка движения по группе позиций"
   ' End If
End Sub

Private Sub Form_Load()
Dim i As Integer, sz As Integer


    oldHeight = Me.Height
    oldWidth = Me.Width

    sql = "SELECT ventureId, ventureName, rusAbbrev, s.id_analytic_default " _
        & " From GuideVenture v" _
        & " left join system s on v.id_analytic = s.id_analytic_default " _
        & "WHERE id_analytic is not null order by ventureName"
    Set Table = myOpenRecordSet("##144", sql, dbOpenForwardOnly)
    If Table Is Nothing Then End
    ReDim ventureIds(0)
    ReDim ventureRests(0)
    ReDim ventureAbbrevs(0)
    
    
    While Not Table.EOF
        lbVenture.AddItem Table!ventureName
        lbVenture.ItemData(lbVenture.ListCount - 1) = Table!ventureId
        sz = UBound(ventureIds)
        ReDim Preserve ventureIds(sz + 1)
        ventureIds(sz + 1) = Table!ventureId

        ReDim Preserve ventureRests(sz + 1)
        ventureRests(sz + 1) = 0

        ReDim Preserve ventureAbbrevs(sz + 1)
        ventureAbbrevs(sz + 1) = Table!rusAbbrev
        
        If Not IsNull(Table!id_analytic_default) Then
            defaultVentureIndex = sz + 1
        End If

        Table.MoveNext
    Wend
    Table.Close
    lbVenture.Height = lbVenture.ListCount * 205 + 50

    tbStartDate.Text = Format(begDate, "dd/mm/yy")
    'VentureHistory.Caption = "Карточка движения по позиции № " & DMCnomNom & _
                       " (" & nomenkName & ")"
    'laStart.Caption = Format(begDate, "dd/mm/yy")
    tbEndDate.Text = Format(CurDate, "dd/mm/yy")
    'laEnd.Caption = tbEndDate.Text
    
    Grid.FormatString = "|||^Дата|<Документ|<Операция|>Кол-во|<Через|>Остаток"
    Grid.ColWidth(0) = 0
    Grid.ColWidth(1) = 0
    Grid.ColWidth(2) = 0
    Grid.ColWidth(ktDate) = 850
    Grid.ColWidth(ktDocNum) = 930
    Grid.ColWidth(ktOperation) = 4400
    Grid.ColWidth(ktQty) = 800
    Grid.ColWidth(ktVenture) = 800
    Grid.ColWidth(ktOstat) = 900
    Grid.Cols = Grid.Cols + lbVenture.ListCount
    
    For i = 1 To lbVenture.ListCount
        Grid.TextMatrix(0, ktOstat + i) = lbVenture.List(i - 1)
        Grid.ColWidth(ktOstat + i) = 600
        Grid.ColAlignment(ktOstat + i) = flexAlignGeneral
    Next i
    
    isLoad = True
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
    cmLoad.Top = cmLoad.Top + h
    Label1.Top = Label1.Top + h
    laQuant.Top = laQuant.Top + h
    cmExit.Top = cmExit.Top + h
    cmExit.Left = Grid.Left + Grid.Width - cmExit.Width
    
    cmPrint.Top = cmPrint.Top + h
    cmPrint.Left = cmExit.Left - 50 - cmPrint.Width
    
    cmExcel.Top = cmExcel.Top + h
    
End Sub


Private Sub Grid_Click()
    mousCol = Grid.MouseCol
    mousRow = Grid.MouseRow
    'If quantity = 0 Then Exit Sub

End Sub


Private Sub Grid_DblClick()
    If Grid.CellBackColor = &H88FF88 Then
        If mousCol = ktQty Then
            textBoxInGridCell tbMobile, Grid
        End If
    End If

End Sub

Private Sub Grid_EnterCell()
Dim isVentureOrder As Boolean
    mousRow = Grid.row
    mousCol = Grid.col

    If mousCol = 0 Then Exit Sub

    isVentureOrder = orderIsVentureOrder(Grid.row)
    If _
        (isVentureOrder _
            And (mousCol = ktDate Or mousCol = ktDocNum Or mousCol = ktQty) _
        ) _
    Then
        Grid.CellBackColor = &H88FF88
    Else
        Grid.CellBackColor = vbYellow
    End If
    
End Sub

Private Sub Grid_LeaveCell()
    If Grid.col <> 0 Then Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim qty As Single

    If Grid.MouseRow = 0 And Shift = 2 Then
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
    ElseIf Button = 2 Then
        Grid.col = Grid.MouseCol
        Grid.row = Grid.MouseRow
        On Error Resume Next
        Grid.SetFocus
        'Grid.CellBackColor = vbButtonFace
        qty = CSng(Grid.Text)
        If Grid.col > ktQty And qty < 0 Then
            mnZachet.Visible = True
            Me.PopupMenu mnPopup
        Else
            mnZachet.Visible = False
        End If
        
        If orderIsVentureOrder(Grid.row) Then
            Me.PopupMenu mpVenture
        End If

    End If
End Sub

Private Sub mnDateChange_Click()
    Grid.col = ktDate
    textBoxInGridCell tbMobile, Grid
End Sub

Private Sub mnDelete_Click()
Dim tmpTopRow As Long

    If MsgBox("Нажмите <Да>, если вы действительно хотите удалить позицию ", vbOKCancel, "Подтверждение") <> vbOK Then Exit Sub
    If vo_deleteNomnom(Grid.TextMatrix(mousRow, ktNomnom), Grid.TextMatrix(mousRow, ktDocNum)) Then
        fillGrid
        ' откорректировать сумму по накладной
        ' сохраняем, какая сумма у накладной была
        ' получаем сумму по удаленной позиции
        ' и итоги за период по предприятию
    End If

End Sub

Private Sub mnDmcMove_Click()
    Grid.col = ktDocNum
    textBoxInGridCell tbMobile, Grid
End Sub

Private Sub mnZachet_Click()
Dim iRow As Long, iCol As Long, curNomnom As String, targetDate As Date, qty As Single
Dim endQty As Single, corrQty As Single, corrCol As Integer
Dim ivoId As Integer, saveCol As Long, saveRow As Long

    iRow = Grid.row
    curNomnom = Grid.TextMatrix(Grid.row, ktNomnom)
    endQty = CSng(Grid.Text)
    For iCol = ktOstat + 1 To ktOstat + UBound(ventureIds)
        If iCol <> Grid.col And CSng(Grid.TextMatrix(iRow, iCol)) > 0 Then
            corrCol = iCol
            Exit For
        End If
    Next iCol
    
    Do
        qty = CSng(Grid.TextMatrix(iRow, Grid.col))
        If qty >= 0 Then
            Exit Do
        End If
        targetDate = CDate(Grid.TextMatrix(iRow, ktDate))
        
        ' !todo corrCol может быть равным 0!

        corrQty = CSng(Grid.TextMatrix(iRow, corrCol))
        iRow = iRow - 1
    Loop While curNomnom = Grid.TextMatrix(iRow, ktNomnom) _
                And (Fix(qty + 0.5) + corrQty) > 0
                
    targetDate = CDate("01." & Month(targetDate) & "." & Year(targetDate))
    Do While Weekday(targetDate) = vbSaturday Or Weekday(targetDate) = vbSunday
        targetDate = DateAdd("d", 1, targetDate)
    Loop
    
    endQty = Abs(Int(endQty))
    If MsgBox( _
        "Для компенсации отрицательного остатка по позиции " & curNomnom _
        & " будет создана накладная. " & vbCr & "Количество (в основных единицах измерения " & CStr(endQty) & " " & vbCr _
        & " Дата проведения накладной " & Format(targetDate, "dd.mm.yyyy") _
        , vbOK Or vbDefaultButton2, "Подтвердите" _
    ) = vbOK Then
        sql = "select wf_put_ivo_nomnom (convert (date, " & Format(targetDate, "yyyymmdd") & ") " _
            & ", '" & Grid.TextMatrix(Grid.row, ktNomnom) & "'" _
            & ", " & CStr(endQty) _
            & ", " & VentureOrder.tbProcent.Text _
            & ", " & CStr(ventureIds(corrCol - ktOstat)) _
            & ", " & CStr(ventureIds(Grid.col - ktOstat)) _
            & ")"
        'Debug.Print sql
        On Error Resume Next
        byErrSqlGetValues "##132.20", sql, ivoId
        MsgBox "Зачет по номнклатуре создан. Номер накладной " & CStr(ivoId), , "Успех"
        saveCol = Grid.col
        saveRow = Grid.row
        fillGrid
        Grid.col = saveCol
        Grid.row = saveRow + 1
        Exit Sub
AA:
        MsgBox "При создании зачет по номнклатуре произошла ошибка ", , "Ошибка!"
        Exit Sub
    End If

    
End Sub


Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String
    If KeyCode = vbKeyReturn Then
        str = tbMobile.Text
        If mousCol = ktQty Then
            If Not isNumericTbox(tbMobile) Then Exit Sub
            
            sql = "update sdmcventure d set quant = n.perlist * " & str _
                & " from sguidenomenk n where n.nomnom = d.nomnom " _
                & " and d.nomnom = '" & Grid.TextMatrix(mousRow, ktNomnom) & "'" _
                & " and d.sdv_id = " & Grid.TextMatrix(mousRow, ktDocNum)
            If myExecute("##119", sql) = 0 Then
                Grid.TextMatrix(mousRow, mousCol) = str
                fillGrid
                lbHide
            End If
        End If
    ElseIf KeyCode = vbKeyEscape Then
        lbHide
    End If
    
End Sub

