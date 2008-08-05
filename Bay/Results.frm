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
Public StartDate As Date
Public endDate As Date
Public managId As String
Dim filterSettings() As MapEntry
Dim PreHeaderCount As Integer, PostHeaderCount As Integer, multiplyCols As Integer
Dim periodCount As Integer ' количество периодов (столбцов)
Dim activeTab As Integer
Dim mousCol As Integer
Dim searchValue As String, searchPos As Long, searchAgain As Boolean

Dim GridHeaderHeadDef() As columnDef
Dim GridHeaderTailDef() As columnDef

Private Type ColumnInfo
    PeriodId As Integer
    label As String
    year As Integer
    index As Integer
    colWidth As Integer
    stDate As Date
    enDate As Date
End Type


Dim columns() As ColumnInfo


' переменные используемые в сортировке таблицы
Dim colType As String
' определяет тип текущей сортировки.
    
Const CT_NUMBER = "numeric"
Const CT_DATE = "date"
Const CT_STRING = ""
Const CT_EMPTY = "empty"
Const CT_CUSTOM = "custom"
Const CT_SCHET = "schet"


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

    Grid.left = 100
    Grid.Width = Me.Width - 300
    TabStrip1.Top = 100
    TabStrip1.Width = Grid.Width
    TabStrip1.left = Grid.left
    Grid.Top = TabStrip1.Top + TabStrip1.Height
    Grid.Height = Me.Height - Grid.Top - 1200
    cmExit.left = Grid.left + Grid.Width - cmExit.Width
    cmExit.Top = Grid.Top + Grid.Height + 50
    cmExel.Top = cmExit.Top
    cmPrint.Top = cmExit.Top
    cmExel.left = 500
    cmPrint.left = cmExel.left + cmExel.Width + 300
    cmExit.Visible = True
    lbTotal.left = cmPrint.left + cmPrint.Width + 300
    lbTotal.Top = cmExit.Top + 50
    lbTotalQty.Top = lbTotal.Top
    lbTotalQty.left = lbTotal.left + lbTotal.Width + 50
    cmFind.left = lbTotalQty.left + lbTotalQty.Width + 300
    cmFind.Top = cmExit.Top
    Grid.Visible = True

End Sub


Private Sub Grid_Click()
    If Grid.MouseRow = 0 Then
        Grid_LeaveCell  ' только чтобы снять выделение
        mousCol = Grid.MouseCol
        colType = determineColType(mousCol)
        Grid.Sort = 9
    End If
    trigger = Not trigger

End Sub


Private Sub Grid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Dim cell_1, cell_2 As String
Dim date1, date2 As Date
Dim num1, num2 As Single

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
        num1 = Round(CSng(cell_1), 5)
        num2 = Round(CSng(cell_2), 5)
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
Dim firmId As Long, PeriodId As Integer
    'Dim PreHeaderCount As Integer, PostHeaderCount As Integer, multiplyCols As Integer
    columnNo = Grid.col
    Portrait.filterId = filterId
    firmId = CInt(Grid.TextMatrix(Grid.row, 0))
    If columnNo = 1 Then
        ' название фирмы (главного атрибута, по которому происходит группировка
        '
        Portrait.mode = "portrait"
        Portrait.byRowId = firmId
        Portrait.byColumnId = 0
        Portrait.Show , Me
    ElseIf columnNo >= PreHeaderCount And columnNo < PreHeaderCount + multiplyCols * periodCount Then
        ' Нажали на ячейку с периодом
        '
        periodNo = getPeriodNoByColumn(columnNo)
        PeriodId = columns(periodNo).PeriodId
        
        Portrait.mode = "detail"
        Portrait.byRowId = firmId
        Portrait.byColumnId = PeriodId
        Portrait.Show , Me
    ElseIf columnNo >= PreHeaderCount + multiplyCols * periodCount Then
        ' нажали на итог по строке
        '
        Portrait.mode = "detail"
        Portrait.byRowId = firmId
        Portrait.byColumnId = 0
        
        Portrait.Show , Me
    End If
End Sub

Private Sub Grid_EnterCell()
    If Grid.row = 0 Or Grid.col = 0 Then
        Exit Sub
    End If
    If Grid.col = 1 Or Grid.col >= PreHeaderCount Then
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
'    Debug.Print "Grid_LeaveCell() => col = " & Grid.col & ", row = " & Grid.row
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
    activateTab curTab.index
End Sub


Private Sub LoadTable()
' Номер строки в таблице
Dim rownum As Integer
Dim groupSelectorColumn As String, prevSelector As Variant



    cleanTable
    
    setFilterParams
    
    groupSelectorColumn = getCurrentSetting("groupSelectorColumn", filterSettings)
    If Not setGridHeaders(filterId) Then
        MsgBox "Отчет не содержит данных", vbExclamation
        Exit Sub
    End If
    
    sql = "call n_exec_filter( " & filterId & ")"
    Set table = myOpenRecordSet("##Results.1", sql, dbOpenDynaset)
    If table Is Nothing Then
        table.Close
        MsgBox "Ошибка при загрузки данных из базы", vbCritical
        Exit Sub
    End If
    If table.BOF Then
        table.Close
        MsgBox "Отчет не содержит данных", vbExclamation
        Exit Sub
    End If
    
    table.MoveFirst
    Dim I As Integer ' номер столбца
    Dim colShift As Integer, ln As Integer
    Dim orderQty As Integer, orderOrdered As Single, materialQty As Single, materialSaled As Single
'    Dim prevFirm As String
    Dim totals() As Double
    Dim columnIndex As Integer
    Dim skipFixedInit As Boolean
    
    ln = UBound(columns)
    ReDim totals(multiplyCols)
    
    
    rownum = 0
    prevSelector = Null
    skipFixedInit = False
    While Not table.EOF

        If prevSelector <> table(groupSelectorColumn) Or IsNull(prevSelector) Then
            I = PreHeaderCount + multiplyCols * ln
            If rownum > 0 Then
                
                colShift = periodCount * multiplyCols + PreHeaderCount
                For columnIndex = 0 To UBound(GridHeaderTailDef)
                    Grid.TextMatrix(rownum, colShift + columnIndex) = Format(totals(columnIndex), GridHeaderTailDef(columnIndex).columnFormat)
                    totals(columnIndex) = 0
                Next columnIndex
                
                Grid.AddItem ""
            End If
            rownum = rownum + 1
        End If

        colShift = 0
        For columnIndex = 0 To UBound(GridHeaderHeadDef)
            Grid.TextMatrix(rownum, colShift + columnIndex) = table(GridHeaderHeadDef(columnIndex).columnName)
        Next columnIndex
        
        
        colShift = getPeriodShift(table!PeriodId) * multiplyCols + UBound(GridHeaderHeadDef) + 1
        For columnIndex = 0 To UBound(GridHeaderTailDef)
            Dim curValue As Double
            curValue = table(GridHeaderTailDef(columnIndex).columnName)
            Grid.TextMatrix(rownum, colShift + columnIndex) = Format(curValue, GridHeaderTailDef(columnIndex).columnFormat)
            totals(columnIndex) = totals(columnIndex) + curValue
        Next columnIndex
        
        prevSelector = table(groupSelectorColumn)
        
        table.MoveNext
    Wend
    table.Close
    
    If rownum > 1 Then
        Grid.RemoveItem rownum
    End If
    If rownum > 0 Then
        rownum = rownum - 1
    End If
    lbTotalQty.Caption = CStr(rownum) & " фирм"
    activateTab 1
End Sub



Private Function getPeriodShift(PeriodId As Integer) As Integer
Dim I As Integer
Dim ln As Integer
    ln = UBound(columns)
    For I = 0 To ln
        If columns(I).PeriodId = PeriodId Then
            getPeriodShift = columns(I).index
            Exit Function
        End If
    Next I
End Function

Private Sub appendToHeader(GridHeaderHead As String, ByRef headerColumn As columnDef, ByRef delimCount As Integer)
Dim delim As String

    If Not headerColumn.saved Then
        Exit Sub
    End If
    If delimCount > 0 Then
        delim = "|"
    Else
        delim = ""
    End If
    GridHeaderHead = GridHeaderHead & delim
    delimCount = delimCount + 1
    
    If headerColumn.hidden <> 1 Then
        GridHeaderHead = GridHeaderHead & headerColumn.align & headerColumn.nameRu
    End If
End Sub


Private Function setGridHeaders(filterId As Integer) As Boolean
Dim periodType As Variant
Dim index As Integer
Dim colInfo As ColumnInfo
Dim colIndex As Integer, I As Integer
Dim GridHeaderHead As String
Dim GridHeaderTail As String
Dim titleStartStr As String, titleEndStr As String
Dim headerList() As columnDef
Dim headerColumn As columnDef
Dim delim As String, delimHead As Integer, delimTail As Integer

    'Optimistic view
    setGridHeaders = True

    initColumns GridHeaderHeadDef, 1, managId, filterId
    initColumns GridHeaderTailDef, 2, managId, filterId
    
    For index = 0 To UBound(GridHeaderHeadDef)
        headerColumn = GridHeaderHeadDef(index)
        appendToHeader GridHeaderHead, headerColumn, delimHead
    Next index
    
    For index = 0 To UBound(GridHeaderTailDef)
        headerColumn = GridHeaderTailDef(index)
        appendToHeader GridHeaderTail, headerColumn, delimTail
    Next index
    
    'GridHeaderHead = "|<Название фирмы|<Регион|>Перв.Визит|>Посл.Визит"
    'GridHeaderTail = ">Заказов|>Сумма заказов|>К-во матер.|>Сумма матер."
    
    'periodType = getCurrentSetting("periodType", filterSettings)

    
    PreHeaderCount = parseHeaderMetrics(GridHeaderHead)
    PostHeaderCount = parseHeaderMetrics(GridHeaderTail)
    multiplyCols = parseTabStrip(GridHeaderTail, TabStrip1)

    If StartDate > "2000-01-01" Then
        titleStartStr = Format(StartDate, "dd.mm.yyyy")
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
    
    
    sql = "call n_exec_header( " & filterId & ") "
    
    Set table = myOpenRecordSet("##Results.2", sql, dbOpenDynaset)
    ReDim columns(0)
    index = 0
    If table.BOF Then
        setGridHeaders = False
    End If
    
    While Not table.EOF
        colInfo.label = table("label")
        colInfo.year = table!year
        colInfo.stDate = table!st
        colInfo.enDate = table!EN
        colInfo.PeriodId = table!PeriodId
        colInfo.index = index
        colInfo.colWidth = getColumnWidth(index, table!label)


        ReDim Preserve columns(index)
        columns(index) = colInfo
        table.MoveNext
        index = index + 1
    Wend
    table.Close
    
    
    periodCount = UBound(columns) + 1
    Grid.row = 0
    
    
    For I = 0 To periodCount - 1
        For colIndex = 0 To multiplyCols - 1
            GridHeaderHead = GridHeaderHead & "|>" & columns(I).label
        Next colIndex
    Next I
    If GridHeaderTail <> "" Then
        GridHeaderHead = GridHeaderHead & "|" & GridHeaderTail
    End If
    Grid.FormatString = GridHeaderHead
'    For I = 0 To periodCount - 1
'        For colIndex = 0 To multiplyCols - 1
'        For colIndex = PreHeaderCount + (I * multiplyCols) To PreHeaderCount + (I + 1) * multiplyCols - 1
'            Grid.colWidth(colIndex) = getColumnWidth(I, columns(I).label)
'        Next colIndex
'    Next I
    
End Function

Function getPeriodNoByColumn(columnNo As Long) As Integer
    getPeriodNoByColumn = (columnNo - PreHeaderCount) \ multiplyCols
End Function

Private Function getColumnWidth(I As Integer, label As String)
    getColumnWidth = 500
End Function


Private Sub setFilterParams()
Dim entry As MapEntry

    Grid.FormatString = "|Фирма|Регион"
    sql = "call n_boot_filter(" & filterId & ", '" & Orders.cbM.Text & "')"
    
    Set table = myOpenRecordSet("##Results.3", sql, dbOpenDynaset)
    While Not table.EOF
        entry.key = table!paramName
        entry.Value = table!paramValue
        append filterSettings, entry
        table.MoveNext
    Wend
    table.Close
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
                tabName = left(headerRest, delimitorPos - 1)
                headerRest = Mid(headerRest, delimitorPos + 1)
            End If
            If Len(tabName) > 0 Then
                Dim controlChar As String
                controlChar = left(tabName, 1)
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
Dim I As Integer, j As Integer, colIndex As Integer

    activeTab = tabNumber

    For I = 0 To periodCount - 1
        For j = 0 To multiplyCols - 1
            colIndex = PreHeaderCount + (I * multiplyCols) + j
            If j + 1 = tabNumber Then
                Grid.colWidth(colIndex) = columns(I).colWidth
            Else
                Grid.colWidth(colIndex) = 0
            End If
        Next j
    Next I
End Sub


Sub saveGridColWidth()
Dim I As Integer, colIndex As Integer
Dim ln As Integer

    For I = 0 To periodCount - 1
        colIndex = PreHeaderCount + (I * multiplyCols) + activeTab - 1
        columns(I).colWidth = Grid.colWidth(colIndex)
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
