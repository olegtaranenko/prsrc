VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Results 
   Caption         =   "Результаты анализа"
   ClientHeight    =   7812
   ClientLeft      =   48
   ClientTop       =   588
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   7812
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmFind 
      Caption         =   "Поиск"
      Height          =   315
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "F7"
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   2100
      TabIndex        =   4
      Top             =   7320
      Width           =   1332
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   9300
      TabIndex        =   3
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "Печать"
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   7320
      Width           =   735
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11892
      _ExtentX        =   20976
      _ExtentY        =   656
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Количество"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Сумма"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5292
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11892
      _ExtentX        =   20976
      _ExtentY        =   9335
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   192
      Left            =   480
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Label lbTotalQty 
      AutoSize        =   -1  'True
      Caption         =   "25 фирм"
      Height          =   192
      Left            =   5520
      TabIndex        =   6
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label lbTotal 
      AutoSize        =   -1  'True
      Caption         =   "Найдено"
      Height          =   192
      Left            =   3480
      TabIndex        =   5
      Top             =   7320
      Width           =   720
   End
End
Attribute VB_Name = "Results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public filterId As Integer
Public applyTriggered As Boolean
Public startDate As Date
Public endDate As Date
Public ManagId As String
Dim filterSettings() As MapEntry
Dim PreHeaderCount As Integer, PostHeaderCount As Integer, multiplyCols As Integer
Dim tableSettingNoRowDetail As Integer
Dim periodCount As Integer ' количество периодов (столбцов)
Dim activeTab As Integer
Dim mousCol As Integer
Dim searchValue As String, searchPos As Long, searchAgain As Boolean



' переменные используемые в сортировке таблицы
Dim colType As String
' определяет тип текущей сортировки.
    
Const CT_NUMBER = "numeric"
Const CT_DATE = "date"
Const CT_STRING = ""
Const CT_EMPTY = "empty"
Const CT_CUSTOM = "custom"
Const CT_SCHET = "schet"


' будут храниться итоги по столбам для показа их внизу таблицы
Dim columnTotals() As Double


Private Sub cmExel_Click()
    GridToExcel Grid
End Sub


Private Sub cmExit_Click()
    Unload Me
End Sub


Private Sub cmFind_Click()
Const orFirma = 1
Dim searchedFirm As Long
    If Not searchAgain Then
        searchValue = InputBox("Укажите полное название или фрагмент.", "Поиск по названию фирмы", searchValue)
        If searchValue = "" Then
            Exit Sub
        End If
    End If
    searchedFirm = findExValInCol(Grid, searchValue, orFirma, searchPos)
    If searchedFirm > 0 Then
        Grid.row = searchedFirm
        searchPos = searchedFirm + 1
    Else
        MsgBox "Достигнут конец таблицы", , "Поиск по " & searchValue
        searchPos = -1
    End If
End Sub


Private Sub cmFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub


Private Sub cmFind_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub


Private Sub cmPrint_Click()
    Me.PrintForm
End Sub


Private Sub Form_Activate()
    If applyTriggered Then
        applyTriggered = False
        LoadTable
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        cmFind_Click
    End If
    If ((KeyCode And vbKeyShift) = vbKeyShift) And ((Shift And vbShiftMask) > 0) And (searchValue <> "") Then
        searchAgain = True
        cmFind.Caption = "Дальше"
        cmFind.ToolTipText = "Shift+F7"
    End If
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode And vbKeyShift) = vbKeyShift Then
        searchAgain = False
        cmFind.Caption = "Поиск"
        cmFind.ToolTipText = "F7"
    End If
End Sub


Private Sub Form_Load()
    ReDim filterSettings(0)
End Sub


Private Sub Form_Resize()
    Grid.Left = 100
    Grid.Width = Me.Width - 300
    TabStrip1.Top = 100
    TabStrip1.Width = Grid.Width
    TabStrip1.Left = Grid.Left
    Grid.Top = TabStrip1.Top + TabStrip1.Height
    Grid.Height = Me.Height - Grid.Top - 1200
    cmExit.Left = Grid.Left + Grid.Width - cmExit.Width
    cmExit.Top = Grid.Top + Grid.Height + 50
    cmExel.Top = cmExit.Top
    cmPrint.Top = cmExit.Top
    cmExel.Left = 500
    cmPrint.Left = cmExel.Left + cmExel.Width + 300
    cmExit.Visible = True
    lbTotal.Left = cmPrint.Left + cmPrint.Width + 300
    lbTotal.Top = cmExit.Top + 50
    lbTotalQty.Top = lbTotal.Top
    lbTotalQty.Left = lbTotal.Left + lbTotal.Width + 50
    cmFind.Left = lbTotalQty.Left + lbTotalQty.Width + 300
    cmFind.Top = cmExit.Top
    Grid.Visible = True

End Sub


Private Sub Grid_Click()
    If Grid.MouseRow = 0 Then
        Grid_LeaveCell  ' только чтобы снять выделение
        mousCol = Grid.MouseCol
        colType = determineColType(mousCol)
        Grid.Sort = 9
        trigger = Not trigger
    End If
End Sub


Private Sub Grid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Dim cell_1, cell_2 As String
Dim date1, date2 As Date
Dim num1, num2 As Double

    ' Всегда проверяем у строки 0-й столбец. если он пустой - считаем, что это строка с итогами по столбцам.
    ' потому что у всех остальных там должен быть id (фирмы, регионах и т.д.)
    
    ' Эта строка всегда в конце таблицы при любой сортировке
    If Grid.TextMatrix(Row1, 0) = "" Then
        Cmp = 1
        Exit Sub
    End If
    If Grid.TextMatrix(Row2, 0) = "" Then
        Cmp = -1
        Exit Sub
    End If
    
    
    cell_1 = Grid.TextMatrix(Row1, mousCol)
    cell_2 = Grid.TextMatrix(Row2, mousCol)
    
    If cell_1 = "" And cell_2 = "" Then
        Cmp = 0: Exit Sub
    End If
    
    If cell_1 = "" Then
        Cmp = 1: Exit Sub
    End If
    If cell_2 = "" Then
        Cmp = -1: Exit Sub
    End If
    
    If colType = CT_NUMBER Then
        num1 = Round(CDbl(cell_1), 5)
        num2 = Round(CDbl(cell_2), 5)
        Cmp = Sgn(num1 - num2)
    ElseIf colType = CT_STRING Then
        If cell_1 > cell_2 Then
            Cmp = 1
        ElseIf cell_1 < cell_2 Then
            Cmp = -1
        Else
            Cmp = 0
        End If
    ElseIf colType = CT_DATE Then
        Cmp = Sgn(CDate(cell_1) - CDate(cell_2))
    End If
    If trigger Then Cmp = -Cmp
End Sub


Private Sub Grid_DblClick()
Dim columnNo As Long, periodNo As Long
Dim FirmId As Long, periodId As Integer
    'Dim PreHeaderCount As Integer, PostHeaderCount As Integer, multiplyCols As Integer
    If Grid.CellBackColor = vbYellow Then Exit Sub

    columnNo = Grid.col
    Portrait.filterId = filterId
    FirmId = CInt(Grid.TextMatrix(Grid.row, 0))
    If columnNo = 1 Then
        ' название фирмы (главного атрибута, по которому происходит группировка
        '
        Portrait.mode = "portrait"
        Portrait.byRowId = FirmId
        Portrait.byColumnId = 0
        Portrait.Show , Me
    ElseIf columnNo >= PreHeaderCount And columnNo < PreHeaderCount + multiplyCols * periodCount Then
        ' Нажали на ячейку с периодом
        '
        periodNo = getPeriodNoByColumn(columnNo)
        periodId = periods(periodNo).periodId
        
        Portrait.mode = "detail"
        Portrait.byRowId = FirmId
        Portrait.byColumnId = periodId
        Portrait.Show , Me
    ElseIf columnNo >= PreHeaderCount + multiplyCols * periodCount Then
        ' нажали на итог по строке
        '
        Portrait.mode = "detail"
        Portrait.byRowId = FirmId
        Portrait.byColumnId = 0
        
        Portrait.Show , Me
    End If
End Sub

Private Sub Grid_EnterCell()

    'Не менять цвет фиксированных элементов
    If Grid.row = 0 Or Grid.col = 0 Then
        Exit Sub
    End If
    
    'У отчета не может быть детализации по умолчанию.
    If tableSettingNoRowDetail = 1 Then
        Grid.CellBackColor = vbYellow
        Exit Sub
    End If
    
    
    If (Grid.col >= PreHeaderCount) And (Grid.row <> Grid.Rows - 1) Then
        Grid.CellBackColor = &H88FF88
    Else
        Grid.CellBackColor = vbYellow
    End If
    
End Sub


Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub


Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub


Private Sub Grid_LeaveCell()
    If Grid.col <> 0 Then
        Grid.CellBackColor = Grid.BackColor
    End If
End Sub


Private Sub Grid_LostFocus()
    saveGridColWidth
End Sub


Private Sub TabStrip1_Click()
Dim currentTab As Tabs
Dim curTab As Variant

    Set curTab = TabStrip1.SelectedItem
    activateTab curTab.Index
End Sub


Private Sub LoadTable()
' Номер строки в таблице
Dim rownum As Integer
Dim groupSelectorColumn As String, prevSelector As Variant
Dim checkResult As String
Dim I As Integer ' номер столбца
Dim columnBaseIndex As Integer, periodCount As Integer
Dim orderQty As Integer, orderOrdered As Double, materialQty As Double, materialSaled As Double
Dim rowTotals() As Double
Dim columnIndex As Integer
Dim skipFixedInit As Boolean
Dim periodColumnName As String
Dim totalQtyLabel As String
Dim totalBaseIndex As Integer
Dim curValue As Variant
Dim periodIndex As Integer


    cleanTable
    Grid.Visible = False
    Me.MousePointer = flexArrowHourGlass
    
    sql = "select n_check_filter( " & filterId & ", '" & ManagId & "')"
    byErrSqlGetValues "##loadTable.1", sql, checkResult
    
    If checkResult <> "ok" Then
        MsgBox "При проверке фильтра возника ошибка: " _
        & " '" & checkResult & "'." _
        & vbCr & "Исправьте и попробуйте снова." _
        , vbExclamation, "Ошибка"
        
        GoTo cleanup
    End If
    
    setFilterParams
    
    groupSelectorColumn = getMapEntry(filterSettings, "groupSelectorColumn")
    If Not setGridHeaders(filterId) Then
        'MsgBox "Отчет не содержит данных", vbExclamation
        Me.Caption = "Отчет не содержит данных"
        GoTo cleanup
    End If
    Dim defTimeout As Integer
    defTimeout = myBase.QueryTimeout
    
    myBase.QueryTimeout = myBase.QueryTimeout * 5
    
    sql = "call n_exec_filter( " & filterId & ")"
    Set Table = myOpenRecordSet("##Results.1", sql, dbOpenDynaset)
    
    myBase.QueryTimeout = defTimeout
    
    If Table Is Nothing Then
        Me.Caption = "Ошибка при загрузке данных из базы"
        GoTo cleanup
    End If
    
    If Table.BOF Then
        Table.Close
        'MsgBox "Отчет не содержит данных", vbExclamation
        Me.Caption = "Отчет не содержит данных"
        GoTo cleanup
    End If
    
    Table.MoveFirst
    
    periodCount = UBound(periods) + 1
    ReDim rowTotals(multiplyCols)
    ReDim columnTotals((periodCount + 1) * multiplyCols)
    
    periodColumnName = getMapEntry(filterSettings, "periodId4detail")

    rownum = 0
    prevSelector = Null
    skipFixedInit = False
    While Not Table.EOF

        If prevSelector <> Table(groupSelectorColumn) Or IsNull(prevSelector) Then
            'totalBaseIndex = getPeriodShift(table("periodId")) * periodCount
            I = PreHeaderCount + multiplyCols * periodCount
            If rownum > 0 Then
                For columnIndex = 0 To UBound(GridHeaderTailDef)
                    'columnTotals(totalBaseIndex + columnIndex) = columnTotals(totalBaseIndex + columnIndex) + rowTotals(columnIndex)
                    rowTotals(columnIndex) = 0
                Next columnIndex
                Grid.AddItem ""
            End If
            rownum = rownum + 1
        End If

        columnBaseIndex = 0
        Dim val As Variant
        For columnIndex = 0 To UBound(GridHeaderHeadDef)
            val = Table(GridHeaderHeadDef(columnIndex).columnName)
            If Not IsNull(val) Then
                If GridHeaderHeadDef(columnIndex).columnFormat <> "" Then
                    Grid.TextMatrix(rownum, columnBaseIndex + columnIndex) = Format(val, GridHeaderHeadDef(columnIndex).columnFormat)
                Else
                    Grid.TextMatrix(rownum, columnBaseIndex + columnIndex) = val
                End If
            End If
        Next columnIndex
        
        
        totalBaseIndex = getPeriodShift(Table("periodId")) * multiplyCols
        columnBaseIndex = totalBaseIndex + UBound(GridHeaderHeadDef) + 1
        For columnIndex = 0 To UBound(GridHeaderTailDef)
            curValue = Table(GridHeaderTailDef(columnIndex).columnName)
            If IsNull(curValue) Then curValue = CDbl(0)
            Grid.TextMatrix(rownum, columnBaseIndex + columnIndex) = Format(curValue, GridHeaderTailDef(columnIndex).columnFormat)
            rowTotals(columnIndex) = rowTotals(columnIndex) + curValue
            columnTotals(totalBaseIndex + columnIndex) = columnTotals(totalBaseIndex + columnIndex) + curValue
        Next
        
        columnBaseIndex = periodCount * multiplyCols + PreHeaderCount
        For columnIndex = 0 To UBound(GridHeaderTailDef)
            Grid.TextMatrix(rownum, columnBaseIndex + columnIndex) = Format(rowTotals(columnIndex), GridHeaderTailDef(columnIndex).columnFormat)
        Next columnIndex
        
        prevSelector = Table(groupSelectorColumn)
        
        Table.MoveNext
    Wend
    Table.Close
    
    Grid.AddItem ""
    Grid.col = 1: Grid.row = rownum + 1
    Grid.CellFontBold = True
    Grid.Text = "Итоги"
    
    For periodIndex = 0 To periodCount
        columnBaseIndex = periodIndex * multiplyCols + PreHeaderCount
        totalBaseIndex = periodCount * multiplyCols
        For columnIndex = 0 To UBound(GridHeaderTailDef)

            Grid.col = columnBaseIndex + columnIndex
            Grid.CellFontBold = True
            curValue = columnTotals(periodIndex * multiplyCols + columnIndex)
            If curValue > 0 Then
                Grid.Text = Format(curValue, GridHeaderTailDef(columnIndex).columnFormat)
            End If

            If periodIndex <> periodCount Then
                columnTotals(totalBaseIndex + columnIndex) = columnTotals(totalBaseIndex + columnIndex) + curValue
            End If
        Next columnIndex
        I = I
    Next periodIndex
    
    AjustColumnWidths Me.Grid, Label1
    totalQtyLabel = getMapEntry(filterSettings, "totalQtyLabel")
    lbTotalQty.Caption = CStr(rownum) & " " & totalQtyLabel
    cmFind.Left = lbTotalQty.Left + lbTotalQty.Width + 100
    
    activateTab 1
    
    Grid.Visible = True
finally:

    Me.MousePointer = flexDefault
    Exit Sub
    
cleanup:
    Grid.Visible = False
    Me.MousePointer = flexDefault
    Exit Sub
    
    
End Sub



Private Function getPeriodShift(periodId As Integer) As Integer
Dim I As Integer
Dim ln As Integer
    ln = UBound(periods)
    For I = 0 To ln
        If periods(I).periodId = periodId Then
            getPeriodShift = periods(I).Index
            Exit Function
        End If
    Next I
End Function

Private Sub appendToHeader(GridHeaderHead As String, ByRef headerColumn As columnDef, ByRef delimCount As Integer)
Dim Delim As String

    If Not headerColumn.saved Then
        Exit Sub
    End If
    If delimCount > 0 Then
        Delim = "|"
    Else
        Delim = ""
    End If
    GridHeaderHead = GridHeaderHead & Delim
    delimCount = delimCount + 1
    
    If headerColumn.hidden <> 1 Then
        GridHeaderHead = GridHeaderHead & headerColumn.align & headerColumn.nameRu
    End If
End Sub


Private Function setGridHeaders(filterId As Integer) As Boolean
Dim periodType As Variant
Dim Index As Integer
Dim colInfo As PeriodDef
Dim colIndex As Integer, I As Integer
Dim GridHeaderHead As String
Dim GridHeaderTail As String
Dim titleStartStr As String, titleEndStr As String, ResultTitle As String
Dim headerList() As columnDef
Dim headerColumn As columnDef
Dim Delim As String, delimHead As Integer, delimTail As Integer
Dim periodColumnName As String

    'Optimistic view
    setGridHeaders = True

    initColumns GridHeaderHeadDef, 1, ManagId, filterId
    initColumns GridHeaderTailDef, 2, ManagId, filterId
    
    For Index = 0 To UBound(GridHeaderHeadDef)
        headerColumn = GridHeaderHeadDef(Index)
        appendToHeader GridHeaderHead, headerColumn, delimHead
    Next Index
    
    For Index = 0 To UBound(GridHeaderTailDef)
        headerColumn = GridHeaderTailDef(Index)
        appendToHeader GridHeaderTail, headerColumn, delimTail
    Next Index
    
    PreHeaderCount = UBound(GridHeaderHeadDef) + 1
    PostHeaderCount = UBound(GridHeaderTailDef) + 1
    multiplyCols = parseTabStrip(GridHeaderTail, TabStrip1)

    ResultTitle = getMapEntry(filterSettings, "resultTitle")
    If ResultTitle = "" Then
        If startDate > "2000-01-01" Then
            titleStartStr = Format(startDate, "dd.mm.yyyy")
        End If
        
        If endDate > "2000-01-01" Then
            titleEndStr = Format(endDate, "dd.mm.yyyy")
        End If
        
        
        Me.Caption = "Продажи: "
        If titleStartStr = "" And titleEndStr = "" Then
            Me.Caption = "весь учет"
        Else
            If titleStartStr <> "" Then
                Me.Caption = Me.Caption & "с " & titleStartStr & " "
            Else
                Me.Caption = Me.Caption & "от начала учета "
            End If
            If titleEndStr <> "" Then
                Me.Caption = Me.Caption & "по " & titleEndStr & " "
            Else
                Me.Caption = Me.Caption & "до окончания учета"
            End If
            
            If titleStartStr <> "" And titleEndStr <> "" Then
                Me.Caption = Me.Caption & "включительно"
            End If
        End If
    Else
        Me.Caption = ELExpressionCheck(ResultTitle)
    End If
    
    
    sql = "call n_exec_header( " & filterId & ") "
    
    Set Table = myOpenRecordSet("##Results.2", sql, dbOpenDynaset)
    ReDim periods(0)
    Index = 0
    If Table.BOF Then
        setGridHeaders = False
        Exit Function
    End If

    'может ли получить деталировку?
    tableSettingNoRowDetail = getMapEntry(filterSettings, "noRowDetail")
    
    periodColumnName = getMapEntry(filterSettings, "periodId4detail")
    
    

    While Not Table.EOF
        colInfo.label = Table("label")
        If Not IsNull(Table!year) Then colInfo.year = Table!year
        If Not IsNull(Table!st) Then colInfo.stDate = Table!st
        If Not IsNull(Table!en) Then colInfo.enDate = Table!en
        colInfo.periodId = Table(periodColumnName)
        colInfo.Index = Index
        colInfo.ColWidth = getColumnWidth(Index, Table!label)


        ReDim Preserve periods(Index)
        periods(Index) = colInfo
        Table.MoveNext
        Index = Index + 1
    Wend
    Table.Close
    
    
    periodCount = UBound(periods) + 1
    Grid.row = 0
    
    
    For I = 0 To periodCount - 1
        For colIndex = 0 To multiplyCols - 1
            GridHeaderHead = GridHeaderHead & "|" & GridHeaderTailDef(colIndex).align & periods(I).label
        Next colIndex
    Next I
    If GridHeaderTail <> "" Then
        GridHeaderHead = GridHeaderHead & "|" & GridHeaderTail
    End If
    Grid.FormatString = GridHeaderHead
    
    For Index = 0 To UBound(GridHeaderHeadDef)
        headerColumn = GridHeaderHeadDef(Index)
        If headerColumn.hidden = 1 Then
            Grid.ColWidth(Index) = 0
        End If
    Next Index
    
End Function

Function ELExpressionCheck(msg As String) As String
    Dim curCheck As String
    curCheck = msg
    Do While InStr(curCheck, "${") > 0
        Dim dollarPos As Integer, closedPos As Integer
        Dim leftStr As String, rightStr As String, varName As String
        
        dollarPos = InStr(curCheck, "${")
        closedPos = InStr(curCheck, "}")
        leftStr = Mid(curCheck, 1, dollarPos - 1)
        varName = Mid(curCheck, dollarPos + 2, closedPos - dollarPos - 2)
        rightStr = Mid(curCheck, closedPos + 1, Len(curCheck))
        curCheck = leftStr & resolveEL(varName) & rightStr
        
    Loop
    ELExpressionCheck = curCheck
        
End Function

Function resolveEL(varName As String) As String
    resolveEL = varName
    If varName = "clientName" Then
        resolveEL = Analityc.clientName
    End If
End Function


Function getPeriodNoByColumn(columnNo As Long) As Integer
    getPeriodNoByColumn = (columnNo - PreHeaderCount) \ multiplyCols
End Function

Private Function getColumnWidth(I As Integer, label As String)
    getColumnWidth = 200
End Function


Private Sub setFilterParams()
Dim entry As MapEntry

'    Grid.FormatString = "|Фирма|Регион"
    sql = "call n_boot_filter(" & filterId & ", '" & ManagId & "')"
    
    Set Table = myOpenRecordSet("##Results.3", sql, dbOpenDynaset)
    While Not Table.EOF
        entry.Key = Table!paramName
        entry.Value = Table!paramValue
        append filterSettings, entry
        Table.MoveNext
    Wend
    Table.Close
End Sub

Private Sub cleanTable()
    cleanSettings filterSettings
    clearGrid Me.Grid
    Me.Grid.Cols = 2
    TabStrip1.Tabs.Clear
    searchPos = -1
    cmFind.Caption = "Поиск"

End Sub


Private Function parseHeaderMetrics(header As String) As Integer
Dim I As Integer, ln As Integer

    ln = Len(header)
    If ln > 0 Then
        parseHeaderMetrics = 1
    Else
        parseHeaderMetrics = 0
        Exit Function
    End If

    For I = 1 To ln
        If Mid(header, I, 1) = "|" Then
            parseHeaderMetrics = parseHeaderMetrics + 1
        End If
    Next I
End Function

Private Function parseTabStrip(formatStr As String, tabStrip As tabStrip) As Integer
Dim I As Integer
Dim loopDone As Boolean, delimitorPos As Long
Dim headerRest As String, headerRestLn As Long, tabName As String

    headerRest = formatStr
    loopDone = False
    I = 1
    While Not loopDone
        headerRestLn = Len(headerRest)
        If headerRestLn > 0 Then
            parseTabStrip = parseTabStrip + 1
            delimitorPos = InStr(1, headerRest, "|", vbBinaryCompare)
            If delimitorPos = 0 Then
                tabName = headerRest
                headerRest = ""
            Else
                tabName = Left(headerRest, delimitorPos - 1)
                headerRest = Mid(headerRest, delimitorPos + 1)
            End If
            If Len(tabName) > 0 Then
                Dim controlChar As String
                controlChar = Left(tabName, 1)
                If InStr(1, "^<>", controlChar, vbBinaryCompare) > 0 Then
                    tabName = Mid(tabName, 2)
                End If
                tabStrip.Tabs.Add , "tab" & CStr(I), tabName
            End If
            I = I + 1
        Else
            loopDone = True
        End If
    Wend
End Function



Private Sub activateTab(tabNumber As Integer)
Dim I As Integer, J As Integer, colIndex As Integer

    activeTab = tabNumber

    For I = 0 To periodCount - 1
        For J = 0 To multiplyCols - 1
            colIndex = PreHeaderCount + (I * multiplyCols) + J
            If J + 1 = tabNumber Then
                Grid.ColWidth(colIndex) = periods(I).ColWidth
            Else
                Grid.ColWidth(colIndex) = 0
            End If
        Next J
    Next I
End Sub


Sub saveGridColWidth()
Dim I As Integer, colIndex As Integer
Dim ln As Integer

    For I = 0 To periodCount - 1
        colIndex = PreHeaderCount + (I * multiplyCols) + activeTab - 1
        periods(I).ColWidth = Grid.ColWidth(colIndex)
    Next I
    
End Sub



Private Function determineColType(ByVal colIndex As Long) As String
Dim rowIndex As Long, cellText As String
Dim asNumber As Integer, asString As Integer, asEmpty As Integer, asDate As Integer, asUnknown As Integer, asSchet As Integer

    For rowIndex = 2 To Grid.Rows
        cellText = Grid.TextMatrix(rowIndex - 1, colIndex)
        If IsNumeric(cellText) Then
            asNumber = asNumber + 1
        ElseIf IsDate(cellText) Then
            asDate = asDate + 1
        ElseIf cellText = "" Then
            asEmpty = asEmpty + 1
        ElseIf IsDate(cellText) Then
            asDate = asDate + 1
        ElseIf InStr(cellText, "=>") > 1 Then
            asSchet = asSchet + 1
        ElseIf Len(cellText) > 1 Then
            asString = asString + 1
        End If
    Next rowIndex
    
    Dim totalRows As Integer
    totalRows = Grid.Rows - asEmpty - 1
    If totalRows = 0 Then
        determineColType = CT_EMPTY
    ElseIf asNumber / totalRows > 0.9 Then
        determineColType = CT_NUMBER
    ElseIf asDate / totalRows > 0.9 Then
        determineColType = CT_DATE
    ElseIf asSchet / totalRows > 0.9 Then
        determineColType = CT_SCHET
    Else
        determineColType = CT_STRING
    End If
    
End Function


Private Sub TabStrip1_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub TabStrip1_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub
