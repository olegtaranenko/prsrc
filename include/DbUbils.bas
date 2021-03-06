Attribute VB_Name = "DbUtils"
Option Explicit
Public myBase As Database
Public wrkDefault As Workspace

Function setNullableParamInt(p As Variant) As String
    If IsNull(p) Or p = "" Then
        setNullableParamInt = "Null"
    Else
        setNullableParamInt = CStr(p)
    End If
End Function


Function setNullableParamStr(p As Variant) As String
    If IsNull(p) Then
        setNullableParamStr = "Null"
    Else
        setNullableParamStr = "'" & CStr(p) & "'"
    End If
End Function

Sub baseOpen()
Dim str As String, dburl As String
    dburl = getDbUrl
    
On Error GoTo ERRb

Set wrkDefault = DBEngine.CreateWorkspace("wrkDefault", "dba", "sql", dbUseODBC) ' ��� ���-�� ����������

    Set myBase = wrkDefault.OpenDatabase("Connection1", _
       dbDriverNoPrompt, False, _
       "ODBC;UID=dba;PWD=sql;DSN=" & dburl)
    
    If myBase Is Nothing Then GoTo ERRb
    
    sql = "call bootstrap_blocking()"
    If myExecute("##bootstrap", sql, 0) = 0 Then GoTo ERRb
    
    Exit Sub
    
ERRb:
       
    If errorCodAndMsg("388", -100) Then '##388
        fatalError "�������� � �������� � ������� ���� ������." & vbCr & "dbUrl = " & dburl
    End If
End Sub

Public Sub reconnectDB()
    Dim conn As String
    conn = myBase.Connect
    myBase.Connection.Close
    Set myBase = wrkDefault.OpenConnection("Connection1", dbDriverNoPrompt, False, conn)
    'myBase.Connect
    'myBase.Connection.Connect
End Sub


Function getSystemField(Field As String) As Variant
    getSystemField = Null
    Set tbSystem = myOpenRecordSet("##147", "System", dbOpenForwardOnly)
    If tbSystem Is Nothing Then myBase.Close: End
    getSystemField = tbSystem.fields(Field)
    tbSystem.Close
End Function


Function getDbUrl() As String
    getDbUrl = getEffectiveSetting("dbUrl")
    If getDbUrl = "" Then
        fatalError "���������� ��������� ������������ ������� ���������." & vbCr & "�� ����������� �������� ��������� dbUrl"
    End If
End Function


