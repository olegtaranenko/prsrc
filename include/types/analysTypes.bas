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
    index As Integer
    colWidth As Integer
    stDate As Date
    enDate As Date
End Type


Public periods() As PeriodDef


Public Sub initColumns(ByRef headerList() As columnDef, headType As Integer, managId As String, _
    Optional filterId As Integer = 0, _
    Optional byRow As Integer = 0, Optional byColumn As Integer = 0 _
)
Dim filterStr As String

    sql = "call n_exec_result_columns_def ( " & headType & ", '" & managId & "'"
    If IsMissing(filterId) Then
        sql = sql & ", null"
    Else
        sql = sql & ", " & filterId
    End If
    If Not IsMissing(byRow) And Not IsMissing(byColumn) Then
        sql = sql & ", " & byRow & ", " & byColumn
    End If
    sql = sql & ")"
    
    Set table = myOpenRecordSet("##initFilter.2", sql, dbOpenForwardOnly)
    If table Is Nothing Then
        Exit Sub
    End If
    
    
    Dim index As Integer, columnDefInfo As columnDef
    
    index = 0
    While Not table.EOF
        If IsNull(table!managId) Then
            columnDefInfo.saved = True
        Else
            columnDefInfo.saved = False
        End If
        
        If columnDefInfo.saved Or headType = 0 Then
            columnDefInfo.columnId = table!columnId
            columnDefInfo.columnName = table!columnName
            columnDefInfo.nameRu = table!nameRu
            columnDefInfo.align = table!align
            columnDefInfo.hidden = table!hidden
            columnDefInfo.inHead = table!headType
            If Not IsNull(table!columnWidth) Then
                columnDefInfo.columnWidth = table!columnWidth
            End If
            If Not IsNull(table!Format) Then
                columnDefInfo.columnFormat = table!Format
            End If
    
            ReDim Preserve headerList(index)
            headerList(index) = columnDefInfo
            index = index + 1
        End If
        
        table.MoveNext
    Wend
    table.Close
End Sub


Public Sub AjustColumnWidths(ByRef Grid As MSFlexGrid, ByRef dummyLabel As label)

    Dim I As Integer, iPeriods As Integer, iTab As Integer, w As Long
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
            w = dummyLabel.width * 1.35
            If w > periods(iPeriods).colWidth Then
                periods(iPeriods).colWidth = w
            End If
        Next iTab
    Next iPeriods
    
    
    
End Sub
