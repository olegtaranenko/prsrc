Attribute VB_Name = "CmdUtils"
Option Explicit

Dim rawCmdArguments() As String
Public argumentSettings() As MapEntry
Public appSettings() As MapEntry
Public siteSettings() As MapEntry
Public settings() As MapEntry




Public appCfgFile As String
Public siteCfgFile As String


Sub dummy()
Dim IsEmpty, Value
End Sub

Sub parseCommandLine(Optional MaxArgs)
   'Declare variables.
   Dim c, CmdLine, CmdLnLen, InArg, I, NumArgs
   Dim inQuoted As Boolean
   
   'See if MaxArgs was provided.
   If IsMissing(MaxArgs) Then MaxArgs = 10
   'Make array of the correct size.
   ReDim rawCmdArguments(MaxArgs)
   NumArgs = 0: InArg = False
   'Get command line arguments.
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   'Go thru command line one character
   'at a time.
   For I = 1 To CmdLnLen
      c = Mid(CmdLine, I, 1)
      If c = """" Then
          If inQuoted Then
             InArg = False
             inQuoted = False
          Else
            If NumArgs = MaxArgs Then Exit For
            NumArgs = NumArgs + 1
            InArg = True
            inQuoted = True
          End If
      'Test for space or tab.
      ElseIf (c <> " " And c <> vbTab) Or inQuoted Then
         'Neither space nor tab.
         'Test if already in argument.
         If Not InArg Then
         'New argument begins.
         'Test for too many arguments.
            If NumArgs = MaxArgs Then Exit For
            NumArgs = NumArgs + 1
            InArg = True
         End If
         'Concatenate character to current argument.
         rawCmdArguments(NumArgs) = rawCmdArguments(NumArgs) & c
      Else
         'Found a space or tab.
         'Set InArg flag to False.
         InArg = False
      End If
      
   Next I
   'Resize array just enough to hold arguments.
   ReDim Preserve rawCmdArguments(NumArgs)
End Sub



Function isAbsolute(path As String) As Boolean
Dim firstChar As String, secondChar As String

    firstChar = Mid(path, 1, 1)
    secondChar = Mid(path, 2, 1)
    isAbsolute = False
    If Len(path) <= 2 Then
        Exit Function
    End If
    If firstChar = "\" Or secondChar = ":" Then
        isAbsolute = True
    End If
End Function

Function delegateArguments(Key As String, Optional separator As String = " ") As String

Dim I As Integer
Dim argTokens As MapEntry
Dim argSnippet As String
    
    For I = 1 To UBound(rawCmdArguments)
        If rawCmdArguments(I) = "-" & Key Then
            argTokens = tokenizeKeyValue(rawCmdArguments(I + 1))
            If argTokens.Value <> "" Then
                argSnippet = separator & "-" & argTokens.Key & " " & argTokens.Value
            Else
                argSnippet = separator & argTokens.Key
            End If
            delegateArguments = delegateArguments & argSnippet
        End If
    Next I
End Function

Function tokenizeKeyValue(ByVal match As String) As MapEntry
    tokenizeKeyValue = tokenize(match, "=")
End Function

' Исходим из предположения, что в кач-ве параметра используется строка типа
' url = prior. Содержание до и после первого знака delimiter попадает
' в key и value соответственно.
' Если в выражении нет знакак delimiter, то key получает значение match, а value - пустую строку.
' делаем тримминг и key и value
Function tokenize(match As String, delimiter As String) As MapEntry
Dim result As MapEntry
Dim equalPos As Long
Dim rawValue As String
Dim castValue As Variant
Dim booleanValue As Boolean

    equalPos = InStr(1, match, delimiter, vbTextCompare)
    If equalPos <> 0 Then
        result.Key = trimAll(Left(match, equalPos - 1))
        rawValue = trimAll(Mid(match, equalPos + 1))
    Else
        result.Key = match
        rawValue = ""
    End If
    
    If rawValue = "истина" Or rawValue = "true" Or rawValue = "Истина" Or rawValue = "True" Then
        castValue = True
    ElseIf rawValue = "Ложь" Or rawValue = "ложь" Or rawValue = "False" Or rawValue = "false" Then
        castValue = False
    Else
        castValue = rawValue
    End If
    result.Value = castValue
    tokenize = result
    
End Function

'кроме нач и кон пробелов удаляет и vbTab
Function trimAll(str As String) As String
    Dim I As Integer, ch As String ', lPoz As Integer, rPoz As Integer
    
    For I = 1 To Len(str)
        ch = Mid$(str, I, 1)
        If ch <> " " And ch <> vbTab Then GoTo AA
    Next I
    trimAll = ""
    Exit Function
AA:
    str = Mid$(str, I)
    For I = Len(str) To 1 Step -1
        ch = Mid$(str, I, 1)
        If ch <> " " And ch <> vbTab Then Exit For
    Next I
    trimAll = Left$(str, I)
    
End Function


Sub loadEffectiveSettings()

    parseCommandLine
    If Not loadCmdSettings(argumentSettings) Then
        fatalError "Ошибка аргумента в командной строке"
    End If
    appCfgFile = getMapEntry("appCfgFile", argumentSettings)
    If appCfgFile = "" Then
        appCfgFile = getAppCfgDefaultName
    End If
    If loadFileSettings(appCfgFile, appSettings) < 0 Then
        fatalError "Ошибка при загрузке файла конфигурации программы (appCfgFile)"
    End If
    siteCfgFile = getMapEntry("siteCfgFile", argumentSettings)
    If siteCfgFile = "" Then
        siteCfgFile = getMapEntry("siteCfgFile", appSettings)
    End If
    If siteCfgFile = "" Then
        siteCfgFile = getSiteCfgDefaultName
    End If
    If loadFileSettings(siteCfgFile, siteSettings) < 0 Then
        fatalError "Ошибка при загрузке файла конфигурации рабочего места (siteCfgFile)"
    End If
End Sub

Sub loadEffectiveSettingsCfg()
    loadEffectiveSettings
    buildEffectiveSettings
End Sub

Sub loadEffectiveSettingsApp()
    loadEffectiveSettings
    buildEffectiveSettings
End Sub

Function loadFileSettings(filePath As String, ByRef curSettings() As MapEntry) As Integer
Dim entry As MapEntry

    Dim str As String, str2 As String, I As Integer, J As Integer
    str = filePath
    ReDim curSettings(0)
    
    On Error GoTo EN1 'если сетевая папка недоступна, то Dir дает ERR
    If Dir$(str) = "" Then
        ' отсутствие файла конфигурации не является ошибкой
        loadFileSettings = 0
    Else
      Open str For Input As #1
      While Not EOF(1)
        Line Input #1, str
        entry = tokenizeKeyValue(str)
        append curSettings, entry
      Wend
      Close #1
      
      loadFileSettings = 1
    End If
    Exit Function
EN1:
    loadFileSettings = -1
End Function


Sub saveAppSettings()
    saveFileSettings appCfgFile, appSettings
End Sub

Sub saveSiteSettings()
    saveFileSettings siteCfgFile, siteSettings
End Sub


Sub saveFileSettings(filePath As String, ByRef curSettings() As MapEntry)
Dim I As Integer, str  As String
Dim doSave As Boolean

    str = filePath
    On Error GoTo EN1
    Open str For Output As #1
    For I = 1 To UBound(curSettings)
        Print #1, curSettings(I).Key & " = " & curSettings(I).Value
    Next I
EN1:
    On Error Resume Next
    Close #1
End Sub

Function getAppCfgDefaultName() As String
    getAppCfgDefaultName = App.path & "\" & App.ExeName & ".cfg"
End Function

Function getSiteCfgDefaultName() As String
    getSiteCfgDefaultName = App.path & "\site.cfg"
End Function

Function getMapEntry(Key As String, ByRef map() As MapEntry) As Variant
Dim I As Integer
    For I = 1 To UBound(map)
        If map(I).Key = Key Then
            getMapEntry = map(I).Value
            Exit Function
        End If
    Next I
'    getMapEntry = Null
End Function

' возвращает индекс key в массиве map
' Empty если не найдет такой параметр
Function getMapEntryIndex(ByRef map() As MapEntry, Key As String) As Integer
Dim I As Integer
    For I = 1 To UBound(map)
        If map(I).Key = Key Then
            getMapEntryIndex = I
            Exit Function
        End If
    Next I
    getMapEntryIndex = Empty
End Function



Sub buildEffectiveSettings()
Dim ln As Integer
Dim I As Integer

    ln = UBound(argumentSettings)
    ReDim settings(ln)
    For I = 1 To ln
        settings(I) = argumentSettings(I)
    Next I

    mergeWithPreference settings, appSettings
    
    mergeWithPreference settings, siteSettings

End Sub


Sub mergeWithPreference(ByRef mergeTo() As MapEntry, mergeFrom() As MapEntry)

Dim lnFrom As Integer
Dim I As Integer
Dim entry As MapEntry
Dim exists As Variant

    lnFrom = UBound(mergeFrom)
    For I = 1 To lnFrom
        entry = mergeFrom(I)
        exists = getMapEntry(entry.Key, mergeTo)
        If IsEmpty(exists) Then
           append mergeTo, entry
        End If
    Next I

End Sub
Public Sub cleanSettings(curSetting() As MapEntry)
    ReDim curSetting(0)
End Sub

Public Sub setSiteSetting(Key As String, Value)
    setCurrentSetting siteSettings, Key, Value
    setCurrentSetting settings, Key, Value
End Sub


Public Sub setAppSetting(Key As String, Value)
    setCurrentSetting appSettings, Key, Value
    setCurrentSetting settings, Key, Value
End Sub

Public Sub setCurrentSetting(curSettings() As MapEntry, Key As String, paramVal)
Dim I As Integer
    
    I = getMapEntryIndex(curSettings, Key)
    If I > 0 Then
        curSettings(I).Value = paramVal
    Else
        Dim entry As MapEntry
        entry.Key = Key
        entry.Value = paramVal
        append curSettings, entry
    End If
    
End Sub

Function getEffectiveSetting(Key As String, Optional defaultValue) As Variant
Dim entry As MapEntry, Value

    Value = getMapEntry(Key, settings)
    If Not IsEmpty(Value) Then
        getEffectiveSetting = Value
        Exit Function
    End If
    If Not IsMissing(defaultValue) Then
        getEffectiveSetting = defaultValue
    End If
    
End Function


Function loadCmdSettings(curSettings() As MapEntry) As Boolean
'Парсинг коммандной строки, можно аргумент задавать без значения.
'В этом случае его значение равно Null.
'Пример в строке stime -otlad -dostup a -devel параметр otlad имеет значение Null.
'Если же stime -dostup a -devel, то тогда параметр otlad имеет значение Empty.

Dim I As Integer
Dim entry As MapEntry, exists As Variant
Dim Value As Variant

    ReDim argumentSettings(0)
    For I = 1 To UBound(rawCmdArguments)
        entry.Value = Null
        If isKey(rawCmdArguments(I)) Then
            entry.Key = Mid(rawCmdArguments(I), 2)
            If isNotKey(I + 1) Then
                Value = Null
                If I + 1 <= UBound(rawCmdArguments) Then
                    Value = rawCmdArguments(I + 1)
                    I = I + 1
                End If
                
                exists = getMapEntry(entry.Key, curSettings)
                If Not IsEmpty(exists) Then
                    exists = exists & " " & Value
                    appendValue curSettings, entry.Key, Value, " "
                Else
                    entry.Value = Value
                    append argumentSettings, entry
                End If
            Else
                exists = getMapEntry(rawCmdArguments(I), curSettings)
                If IsNull(exists) Then
                    entry.Key = rawCmdArguments(I)
                    append argumentSettings, entry
                ElseIf Not IsEmpty(entry.Key) Then
                    append argumentSettings, entry
                End If
            End If
        End If
    Next I
    loadCmdSettings = True
End Function

Function isKey(arg As String) As Boolean
    If Left(arg, 1) = "-" Then
        isKey = True
    Else
        isKey = False
    End If
End Function

Function isNotKey(I As Integer) As Boolean
Dim arg As String
Dim sz As Integer

    sz = UBound(rawCmdArguments)
    If I > sz Then
        isNotKey = True
    Else
        isNotKey = Not isKey(rawCmdArguments(I))
    End If
End Function


Sub append(curSettings() As MapEntry, entry As MapEntry)
Dim ln As Integer
    
    ln = UBound(curSettings) + 1
    ReDim Preserve curSettings(ln)
    curSettings(ln) = entry
End Sub

Sub appendValue(curSettings() As MapEntry, Key As String, Value As Variant, separator As String)
Dim sz As Integer, I As Integer

    sz = UBound(curSettings)
    For I = 1 To sz
        If curSettings(I).Key = Key Then
            curSettings(I).Value = curSettings(I).Value & separator & Value
            Exit Sub
        End If
    Next I
End Sub

Sub setAndSave(scope As String, Key As String, ByRef Value As Variant)

    If scope = "app" Then
        setCurrentSetting appSettings, Key, Value
        saveFileSettings appCfgFile, appSettings
    End If

End Sub


Function stripPath(myFile As String) As String
Dim I As Integer
Dim sz As Integer
    sz = Len(myFile)
    For I = sz - 1 To 1 Step -1
        If Mid(myFile, I, 1) = "\" Then
            stripPath = Mid(myFile, I + 1)
            Exit Function
        End If
    Next I
    stripPath = myFile
End Function

Sub checkReloadCfg()
Dim reloadCfgSrc As String, reloadCfgDst As String

'    MsgBox App.EXEName & ""

    reloadCfgSrc = getMapEntry("reloadCfgSrc", argumentSettings)
    'trace "reloadCfgSrc = " & reloadCfgSrc
    reloadCfgDst = getMapEntry("reloadCfgDst", argumentSettings)
    'trace "reloadCfgDst = " & reloadCfgDst

    If reloadCfgSrc <> "" Then
        MsgBox "Обнаружена необходимость обновления управляющей программы." _
            & vbCr & "Это произойдет автоматически, после чего нужно будет запустить программу еще раз.", vbExclamation, "Предупреждение"
        On Error GoTo EN1

        FileCopy reloadCfgSrc, reloadCfgDst
        End
    End If
    Exit Sub
EN1:
    fatalError "Ошибка при копировании управляющей программы cfg.exe" & vbCr & Err.Description
    End
End Sub

