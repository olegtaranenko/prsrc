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
    If table Is Nothing Then myBase.Close: End
    
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


