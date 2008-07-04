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
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   2100
      TabIndex        =   4
      Top             =   7320
      Width           =   1215
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5292
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11892
      _ExtentX        =   20976
      _ExtentY        =   9335
      _Version        =   393216
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
Dim filterSettings() As MapEntry
Dim PreHeaderCount As Integer, PostHeaderCount As Integer, multiplyCols As Integer
Dim periodCount As Integer ' количество периодов (столбцов)
Dim activeTab As Integer
Dim mousCol As Integer

Private Type ColumnInfo
    periodId As Integer
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
    'определяет тип текущей сортировки.
    
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


Private Sub cmPrint_Click()
    Me.PrintForm

End Sub


Private Sub Form_Activate()
    If applyTriggered Then
        applyTriggered = False
        LoadTable
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
    
    Grid.Visible = True

End Sub


Private Sub Grid_Click()
    If Grid.MouseRow = 0 Then
        mousCol = Grid.MouseCol
        colType = determineColType(mousCol)
        Grid.Sort = 9
        Grid.row = 1    ' только чтобы снять выделение
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


Private Sub Grid_LeaveCell()
    If Grid.col <> 0 Then Grid.CellBackColor = Grid.BackColor
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

    cleanTable
    
    setFilterParams
    
    If Not setGridHeaders() Then
        MsgBox "Отчет не содержит данных", vbExclamation
        Exit Sub
    End If
    
    sql = "call n_exec_filter( " & filterId & ")"
    Set table = myOpenRecordSet("##Results.1", sql, dbOpenDynaset)
    If table Is Nothing Then
        MsgBox "Ошибка при загрузки данных из базы", vbCritical
        Exit Sub
    End If
    If table.BOF Then
        table.Close
        MsgBox "Отчет не содержит данных", vbExclamation
        Exit Sub
    End If
    
    table.MoveFirst
    Dim i As Integer ' номер столбца
    Dim colShift As Integer, ln As Integer
    Dim saled As Single, paid As Single, orders As Integer
    Dim prevFirmId As Integer
    Dim totals() As Double
    
    ln = UBound(columns)
    ReDim totals(multiplyCols)
    
    
    rownum = 0
    prevFirmId = 0
    While Not table.EOF

        If prevFirmId <> table!firmId Then
            i = PreHeaderCount + multiplyCols * ln
            If rownum > 0 Then
                colShift = periodCount * multiplyCols + PreHeaderCount
                Grid.TextMatrix(rownum, colShift) = Format(orders, "# ##0")
                orders = 0
                Grid.TextMatrix(rownum, colShift + 1) = Format(saled, "# ###.00")
                saled = 0
                Grid.TextMatrix(rownum, colShift + 2) = Format(paid, "# ###.00")
                paid = 0
                Grid.AddItem ""
            End If
            rownum = rownum + 1
        End If

        i = 1
        Grid.TextMatrix(rownum, i) = table!Name: i = i + 1
        Grid.TextMatrix(rownum, i) = table!region: i = i + 1
        Grid.TextMatrix(rownum, i) = table!erst: i = i + 1
        Grid.TextMatrix(rownum, i) = table!letzt: i = i + 1
        
        colShift = i + getPeriodShift(table!periodId) * multiplyCols
        Grid.TextMatrix(rownum, colShift) = table!orders_cnt
        Grid.TextMatrix(rownum, colShift + 1) = Format(table!paid, "# ###.00")
        Grid.TextMatrix(rownum, colShift + 2) = table!qty
        Grid.TextMatrix(rownum, colShift + 3) = Format(table!saled, "# ###.00")
        
        paid = table!paid + paid
        saled = table!saled + saled
        orders = table!orders_cnt + orders
        
        prevFirmId = table!firmId
        
        table.MoveNext
    Wend
    table.Close
    If rownum > 1 Then
        Grid.RemoveItem rownum
    End If
    
    activateTab 1
End Sub



Private Function getPeriodShift(periodId As Integer) As Integer
Dim i As Integer
Dim ln As Integer
    ln = UBound(columns)
    For i = 0 To ln
        If columns(i).periodId = periodId Then
            getPeriodShift = columns(i).index
            Exit Function
        End If
    Next i
End Function

Private Function setGridHeaders() As Boolean
Dim periodType As Variant
Dim index As Integer
Dim colInfo As ColumnInfo
Dim colIndex As Integer, i As Integer
Dim GridHeaderHead As String
Dim GridHeaderTail As String

    'Optimistic view
    setGridHeaders = True

    GridHeaderHead = "|<Название фирмы|<Регион|>1. Визит|>Посл.Визит"
    GridHeaderTail = ">Заказов|>На общую сумму|>Проплачено"
    periodType = getCurrentSetting("periodType", filterSettings)

    
    PreHeaderCount = analyzeHeader(GridHeaderHead)
    PostHeaderCount = analyzeHeader(GridHeaderTail)
    multiplyCols = 4

    sql = "call n_fill_periods( "
    
    If StartDate > "2000-01-01" Then
        sql = sql & "'" & Format(StartDate, "yyyymmdd") & "'"
    Else
        sql = sql & "null"
    End If
    sql = sql & ", "
    
    If endDate > "2000-01-01" Then
        sql = sql & "'" & Format(endDate, "yyyymmdd") & "'"
    Else
        sql = sql & "null"
    End If
    
    sql = sql & ", '" & periodType & "', 1)"
    
    
    Set table = myOpenRecordSet("##Results.2", sql, dbOpenDynaset)
    ReDim columns(0)
    index = 0
    If table.BOF Then
        setGridHeaders = False
    End If
    
    While Not table.EOF
        colInfo.label = table!label
        colInfo.year = table!year
        colInfo.stDate = table!st
        colInfo.enDate = table!EN
        colInfo.periodId = table!periodId
        colInfo.index = index
        colInfo.colWidth = getColumnWidth(index, table!label)
        
        
        ReDim Preserve columns(index)
        columns(index) = colInfo
        table.MoveNext
        index = index + 1
    Wend
    table.Close
    
    
    periodCount = UBound(columns) + 1
'    Grid.Cols = PreHeaderCount + periodCount * multiplyCols + PostHeaderCount
    Grid.row = 0
    
    
    For i = 0 To periodCount - 1
        For colIndex = 0 To multiplyCols - 1
            GridHeaderHead = GridHeaderHead & "|>" & columns(i).label
        Next colIndex
    Next i
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


Private Function getColumnWidth(i As Integer, label As String)
    getColumnWidth = 500
End Function


Private Sub setFilterParams()
Dim entry As MapEntry

    Grid.FormatString = "|Фирма|Регион"
    sql = "call n_boot_filter(" & filterId & ", '" & orders.cbM.Text & "')"
    
    Set table = myOpenRecordSet("##Results.3", sql, dbOpenDynaset)
    While Not table.EOF
        entry.key = table!pKey
        entry.value = table!pValue
        append filterSettings, entry
        table.MoveNext
    Wend
    table.Close
End Sub

Private Sub cleanTable()
    cleanSettings filterSettings
    clearGrid Me.Grid
    Me.Grid.Cols = 2

End Sub


Private Function analyzeHeader(header As String) As Integer
Dim i As Integer, ln As Integer

    ln = Len(header)
    If ln > 0 Then
        analyzeHeader = 1
    Else
        analyzeHeader = 0
        Exit Function
    End If

    For i = 1 To ln
        If Mid(header, i, 1) = "|" Then
            analyzeHeader = analyzeHeader + 1
        End If
    Next i
End Function


Private Sub activateTab(tabNumber As Integer)
Dim i As Integer, j As Integer, colIndex As Integer

    activeTab = tabNumber

    For i = 0 To periodCount - 1
        For j = 0 To multiplyCols - 1
            colIndex = PreHeaderCount + (i * multiplyCols) + j
            If j + 1 = tabNumber Then
                Grid.colWidth(colIndex) = columns(i).colWidth
            Else
                Grid.colWidth(colIndex) = 0
            End If
        Next j
    Next i
End Sub


Sub saveGridColWidth()
Dim i As Integer, colIndex As Integer
Dim ln As Integer

    For i = 0 To periodCount - 1
        colIndex = PreHeaderCount + (i * multiplyCols) + activeTab - 1
        columns(i).colWidth = Grid.colWidth(colIndex)
    Next i
    
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


