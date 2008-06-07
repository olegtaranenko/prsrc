Attribute VB_Name = "LogUtils"
Private logFileName As String
Private logLevel As Integer
Private logLevelStr As String


Private Const LOG_PANIC As Integer = 0
Private Const LOG_ERROR As Integer = 10
Private Const LOG_WARN As Integer = 20
Private Const LOG_INFO As Integer = 30
Private Const LOG_DEBUG As Integer = 40
Private Const LOG_TRACE As Integer = 50

'

Private logEnabled As Boolean


'Howto как пользоваться системой журналирования
' в cfg.cfg установить параметры:
' log = DEBUG|TRACE|...
' logger = <file name>
'
'


Function panic(msg As String, Optional logger_id) As String
    If isPanicEnabled Then
        fileLogMsg msg, LOG_PANIC
    End If
End Function

Function erro(msg As String, Optional logger_id) As String
    If isErrorEnabled Then
        fileLogMsg msg, LOG_DEBUG
    End If
End Function
Function warn(msg As String, Optional logger_id) As String
    If isWarningEnabled Then
        fileLogMsg msg, LOG_WARN
    End If
End Function

Function info(msg As String, Optional logger_id) As String
    If isInfoEnabled Then
        fileLogMsg msg, LOG_INFO
    End If
End Function

Function dbg(msg As String, Optional logger_id) As String
    If isDebugEnabled Then
        fileLogMsg msg, LOG_DEBUG
    End If
End Function

Function trace(msg As String, Optional logger_id) As String
    If isTraceEnabled Then
        fileLogMsg msg, LOG_TRACE
    End If
End Function


Function isPanicEnabled() As Boolean
    If logEnabled Then
        isPanicEnabled = True
    Else
        isPanicEnabled = False
    End If
End Function

Function isErrorEnabled() As Boolean
    If logEnabled And logLevel > LOG_PANIC Then
        isErrorEnabled = True
    Else
        isErrorEnabled = False
    End If
End Function

Function isWarningEnabled() As Boolean
    If logEnabled And logLevel > LOG_ERROR Then
        isWarningEnabled = True
    Else
        isWarningEnabled = False
    End If
End Function

Function isInfoEnabled() As Boolean
    If logEnabled And logLevel > LOG_WARN Then
        isInfoEnabled = True
    Else
        isInfoEnabled = False
    End If
End Function

Function isDebugEnabled() As Boolean
    If logEnabled And logLevel > LOG_INFO Then
        isDebugEnabled = True
    Else
        isDebugEnabled = False
    End If
End Function

Function isTraceEnabled() As Boolean
    If logEnabled And logLevel > LOG_DEBUG Then
        isTraceEnabled = True
    Else
        isTraceEnabled = False
    End If
End Function


Sub fileLogMsg(msg, level As Integer)
Dim fileNumber
    If Not logEnabled Then
        Exit Sub
    End If
    fileNumber = FreeFile
    Open logFileName For Append As #fileNumber
    Print #fileNumber, formatted(msg, level)
    Close #fileNumber
    
End Sub


Function formatted(msg, level As Integer)
Dim levelStr As String

    levelStr = Level2String(level)
    formatted = Format(Now, "dd.mm.yyyy hh:m:ss") & "-" & levelStr & " - " & msg
End Function

Function Level2String(level As Integer)
    
    If level = 0 Then
        Level2String = "     "
    ElseIf level < LOG_ERROR Then
        Level2String = "PANIC"
    ElseIf level < LOG_WARN Then
        Level2String = "ERROR"
    ElseIf level < LOG_INFO Then
        Level2String = "WARN "
    ElseIf level < LOG_DEBUG Then
        Level2String = "INFO "
    ElseIf level < LOG_TRACE Then
        Level2String = "DEBUG"
    Else
        Level2String = "TRACE"
    End If
End Function


Sub initLogFileName()
Dim attr As Boolean

    logFileName = getEffectiveSetting("logger")

    If logFileName = "" Then
        logFileName = App.path & "\" & App.exeName & ".log"
    End If
On Error GoTo IOErr
    logEnabled = True
'    If Dir(logFileName) Then
'        attr = (GetAttr(logFileName) And vbDirectory) = vbDirectory
'    End If
    fileLogMsg " =========== Start new log session for " & App.exeName, 0
    logLevelStr = getEffectiveSetting("log")
    setLogLevel (logLevelStr)
    
IOErr:
    If attr Then
        MsgBox "Ошибка инициализации лог-файла: " & logFileName _
        & vbCr & "Функции протоколирования будут отключены", _
        vbExclamation, "Предупреждения"
        logEnabled = False
    Else
        logEnabled = True
    End If
    
End Sub

Public Sub setLogLevel(level)
Dim levelInt As Integer

    If IsNumeric(level) Then
        levelInt = CInt(level)
    Else
        levelInt = string2LogLevel(CStr(level))
    End If
    info ("Уровень протоколирования установлен в " & Level2String(levelInt))
    logLevel = levelInt
End Sub

Function string2LogLevel(ByVal level As String)
    If LCase(level) = "panic" Then
        string2LogLevel = LOG_PANIC
    ElseIf LCase(level) = "error" Then
        string2LogLevel = LOG_ERROR
    ElseIf LCase(level) = "warn" Then
        string2LogLevel = LOG_WARN
    ElseIf LCase(level) = "info" Then
        string2LogLevel = LOG_INFO
    ElseIf LCase(level) = "debug" Then
        string2LogLevel = LOG_DEBUG
    ElseIf LCase(level) = "trace" Then
        string2LogLevel = LOG_TRACE
    Else
        string2LogLevel = LOG_WARN
        'MsgBox "Неизвестный уровень протокола " & level _
        & vbCr & "Протоколирование будет отключено", vbExclamation, "Предупреждение"
        logEnabled = True
    End If
    
End Function

