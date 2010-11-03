Attribute VB_Name = "AnalysTypes"
' Для определения заголовка таблицы Анализа продаж
Public Type columnDef
    columnId As Integer
    columnName As String
    nameRu As String
    align As String
    hidden As Integer
    inHead As Integer
    columnWidth As Integer
    columnFormat As String
    saved As Boolean
End Type

Public GridHeaderHeadDef() As columnDef
Public GridHeaderTailDef() As columnDef


Public Type PeriodDef
    periodId As Integer
    label As String
    year As Integer
    Index As Integer
    ColWidth As Integer
    stDate As Date
    enDate As Date
End Type


Public periods() As PeriodDef


Public Sub initColumns(ByRef headerList() As columnDef, headType As Integer, ManagId As String, _
    Optional filterId As Integer = 0, _
    Optional byRow As Integer = 0, Optional byColumn As Integer = 0 _
)
Dim filterStr As String

    sql = "call n_exec_result_columns_def ( " & headType & ", '" & ManagId & "'"
    If IsMissing(filterId) Then
        sql = sql & ", null"
    Else
        sql = sql & ", " & filterId
    End If
    If Not IsMissing(byRow) And Not IsMissing(byColumn) Then
        sql = sql & ", " & byRow & ", " & byColumn
    End If
    sql = sql & ")"
    
    Set Table = myOpenRecordSet("##initFilter.2", sql, dbOpenForwardOnly)
    If Table Is Nothing Then
        Exit Sub
    End If
    
    
    Dim Index As Integer, columnDefInfo As columnDef
    
    Index = 0
    While Not Table.EOF
        If IsNull(Table!ManagId) Then
            columnDefInfo.saved = True
        Else
            columnDefInfo.saved = False
        End If
        
        If columnDefInfo.saved Or headType = 0 Then
            columnDefInfo.columnId = Table!columnId
            columnDefInfo.columnName = Table!columnName
            columnDefInfo.nameRu = Table!nameRu
            columnDefInfo.align = Table!align
            columnDefInfo.hidden = Table!hidden
            columnDefInfo.inHead = Table!headType
            If Not IsNull(Table!columnWidth) Then
                columnDefInfo.columnWidth = Table!columnWidth
            End If
            If Not IsNull(Table!Format) Then
                columnDefInfo.columnFormat = Table!Format
            End If
    
            ReDim Preserve headerList(Index)
            headerList(Index) = columnDefInfo
            Index = Index + 1
        End If
        
        Table.MoveNext
    Wend
    Table.Close
End Sub


Public Sub AjustColumnWidths(ByRef Grid As MSFlexGrid, ByRef dummyLabel As label)

    Dim I As Integer, iPeriods As Integer, iTab As Integer, W As Long
    Dim startColToCheck As Long, numOfTabs As Integer
    
    For I = 0 To UBound(GridHeaderHeadDef)
        If Not GridHeaderHeadDef(I).hidden Then
            startColToCheck = startColToCheck + 1
        End If
    Next I
    
    For I = 0 To UBound(GridHeaderTailDef)
        If Not GridHeaderTailDef(I).hidden Then
            numOfTabs = numOfTabs + 1
        End If
    Next I
    
    For iPeriods = 0 To UBound(periods)
        For iTab = 1 To numOfTabs
            I = startColToCheck + iPeriods * numOfTabs + iTab - 1
            dummyLabel.Caption = Grid.TextMatrix(Grid.Rows - 1, I)
            W = dummyLabel.Width * 1.35
            If W > periods(iPeriods).ColWidth Then
                periods(iPeriods).ColWidth = W
            End If
        Next iTab
    Next iPeriods
    
    
    
End Sub
