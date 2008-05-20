Attribute VB_Name = "DbUtils"
Option Explicit
Public myBase As Database
Public wrkDefault As Workspace


Sub baseOpen()
Dim str As String, dburl As String
    dburl = getDbUrl
    
On Error GoTo ERRb

Set wrkDefault = DBEngine.CreateWorkspace("wrkDefault", "dba", "sql", dbUseODBC) ' для орг-ии транзакций

    Set myBase = wrkDefault.OpenDatabase("Connection1", _
       dbDriverNoPrompt, False, _
       "ODBC;UID=dba;PWD=sql;DSN=" & dburl)
    
    If myBase Is Nothing Then GoTo ERRb
    
    sql = "call bootstrap_blocking()"
    If myExecute("##bootstrap", sql, 0) = 0 Then GoTo ERRb
    
    Exit Sub
    
ERRb:
       
    If errorCodAndMsg("388", -100) Then '##388
        fatalError "Проблемы с доступом к серверу базы данных." & vbCr & "dbUrl = " & dburl
    End If
End Sub


Function getDbUrl() As String
    getDbUrl = getEffectiveSetting("dbUrl")
    If getDbUrl = "" Then
        fatalError "Необходимо исправить конфигурацию запуска программы." & vbCr & "Не установлено значение параметра dbUrl"
    End If
End Function
