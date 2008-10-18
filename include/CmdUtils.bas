Attribute VB_Name = "CmdUtils"
Option Explicit

Dim rawCmdArguments() As String
Public argumentSettings() As MapEntry
Public appSettings() As MapEntry
Public siteSettings() As MapEntry
Public settings() As MapEntry




Public appCfgFile As String
Public siteCfgFile As String

Type MapEntry
    key As String
    value As Variant
End Type



Sub parseCommandLine(Optional MaxArgs)
   'Declare variables.
   Dim c, CmdLine, CmdLnLen, InArg, i, NumArgs
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
   For i = 1 To CmdLnLen
      c = Mid(CmdLine, i, 1)
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
      
   Next i
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

Function delegateArguments(key As String, Optional separator As String = " ") As String

Dim i As Integer
Dim argTokens As MapEntry
Dim argSnippet As String
    
    For i = 1 To UBound(rawCmdArguments)
        If rawCmdArguments(i) = "-" & key Then
            argTokens = tokenizeKeyValue(rawCmdArguments(i + 1))
            If argTokens.value <> "" Then
                argSnippet = separator & "-" & argTokens.key & " " & argTokens.value
            Else
                argSnippet = separator & argTokens.key
            End If
            delegateArguments = delegateArguments & argSnippet
        End If
    Next i
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

    equalPos = InStr(1, match, delimiter, vbTextCompare)
    If equalPos <> 0 Then
        result.key = trimAll(left(match, equalPos - 1))
        result.value = trimAll(Mid(match, equalPos + 1))
    Else
        result.key = match
        result.value = ""
    End If
    tokenize = result
    
End Function

'кроме нач и кон пробелов удаляет и vbTab
Function trimAll(str As String) As String
    Dim i As Integer, ch As String ', lPoz As Integer, rPoz As Integer
    
    For i = 1 To Len(str)
        ch = Mid$(str, i, 1)
        If ch <> " " And ch <> vbTab Then GoTo AA
    Next i
    trimAll = ""
    Exit Function
AA:
    str = Mid$(str, i)
    For i = Len(str) To 1 Step -1
        ch = Mid$(str, i, 1)
        If ch <> " " And ch <> vbTab Then Exit For
    Next i
    trimAll = left$(str, i)
    
End Function


Sub loadEffectiveSettings()

    parseCommandLine
    If Not loadCmdSettings(argumentSettings) Then
        fatalError "Ошибка аргумента в командной строке"
    End If
    appCfgFile = getCurrentSetting("appCfgFile", argumentSettings)
    If appCfgFile = "" Then
        appCfgFile = getAppCfgDefaultName
    End If
    If loadFileSettings(appCfgFile, appSettings) < 0 Then
        fatalError "Ошибка при загрузке файла конфигурации программы (appCfgFile)"
    End If
    siteCfgFile = getCurrentSetting("siteCfgFile", argumentSettings)
    If siteCfgFile = "" Then
        siteCfgFile = getCurrentSetting("siteCfgFile", appSettings)
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

    Dim str As String, str2 As String, i As Integer, j As Integer
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

Sub saveFileSettings(filePath As String, ByRef curSettings() As MapEntry)
Dim i As Integer, str  As String
Dim doSave As Boolean

    str = filePath
    On Error GoTo EN1
    Open str For Output As #1
    For i = 1 To UBound(curSettings)
        Print #1, curSettings(i).key & " = " & curSettings(i).value
    Next i
EN1:
    On Error Resume Next
    Close #1
End Sub

Function getAppCfgDefaultName() As String
    getAppCfgDefaultName = App.path & "\" & App.EXEName & ".cfg"
End Function

Function getSiteCfgDefaultName() As String
    getSiteCfgDefaultName = App.path & "\site.cfg"
End Function

Function getCurrentSetting(key As String, ByRef curSettings() As MapEntry) As Variant
Dim i As Integer
    For i = 1 To UBound(curSettings)
        If curSettings(i).key = key Then
            getCurrentSetting = curSettings(i).value
            Exit Function
        End If
    Next i
'    getCurrentSetting = Null
End Function

' возвращает индекс key в массиве curSettings
' Empty если не найдет такой параметр
Function getMapEntry(ByRef curSettings() As MapEntry, key As String) As Integer
Dim i As Integer
    For i = 1 To UBound(curSettings)
        If curSettings(i).key = key Then
            getMapEntry = i
            Exit Function
        End If
    Next i
    getMapEntry = Empty
End Function

Sub buildEffectiveSettings()
Dim ln As Integer
Dim i As Integer

    ln = UBound(argumentSettings)
    ReDim settings(ln)
    For i = 1 To ln
        settings(i) = argumentSettings(i)
    Next i

    mergeWithPreference settings, appSettings
    
    mergeWithPreference settings, siteSettings

End Sub

Sub mergeWithPreference(ByRef mergeTo() As MapEntry, mergeFrom() As MapEntry)

Dim lnFrom As Integer
Dim i As Integer
Dim entry As MapEntry
Dim exists As Variant

    lnFrom = UBound(mergeFrom)
    For i = 1 To lnFrom
        entry = mergeFrom(i)
        exists = getCurrentSetting(entry.key, mergeTo)
        If IsEmpty(exists) Then
           append mergeTo, entry
        End If
    Next i

End Sub
Public Sub cleanSettings(curSetting() As MapEntry)
    ReDim curSetting(0)
End Sub

Public Sub setCurrentSetting(curSettings() As MapEntry, key As String, paramVal)
Dim i As Integer
Dim entry As MapEntry, value

    value = getCurrentSetting(key, curSettings)
    If Not IsEmpty(value) Then
        entry.value = paramVal
    Else
        entry.key = key
        entry.value = paramVal
        append curSettings, entry
    End If
    
End Sub

Function getEffectiveSetting(key As String, Optional defaultValue) As Variant
Dim entry As MapEntry, value

    value = getCurrentSetting(key, settings)
    If Not IsEmpty(value) Then
        getEffectiveSetting = value
        Exit Function
    End If
    If Not IsMissing(defaultValue) Then
        getEffectiveSetting = defaultValue
    End If
    
End Function

Function loadCmdSettings(curSettings() As MapEntry) As Boolean
Dim i As Integer
Dim entry As MapEntry, exists As Variant
Dim value As Variant

    ReDim argumentSettings(0)
    For i = 1 To UBound(rawCmdArguments)
        entry.value = Null
        If isKey(rawCmdArguments(i)) Then
            entry.key = Mid(rawCmdArguments(i), 2)
            If isNotKey(i + 1) Then
                value = Null
                If i + 1 <= UBound(rawCmdArguments) Then
                    value = rawCmdArguments(i + 1)
                    i = i + 1
                End If
                
                exists = getCurrentSetting(entry.key, curSettings)
                If Not IsEmpty(exists) Then
                    exists = exists & " " & value
                    appendValue curSettings, entry.key, value, " "
                Else
                    entry.value = value
                    append argumentSettings, entry
                End If
            Else
                exists = getCurrentSetting(rawCmdArguments(i), curSettings)
                If IsNull(exists) Then
                    entry.key = rawCmdArguments(i)
                    append argumentSettings, entry
                ElseIf Not IsEmpty(entry.key) Then
                    append argumentSettings, entry
                End If
            End If
        End If
    Next i
    loadCmdSettings = True
End Function

Function isKey(arg As String) As Boolean
    If left(arg, 1) = "-" Then
        isKey = True
    Else
        isKey = False
    End If
End Function

Function isNotKey(i As Integer) As Boolean
Dim arg As String
Dim sz As Integer

    sz = UBound(rawCmdArguments)
    If i > sz Then
        isNotKey = True
    Else
        isNotKey = Not isKey(rawCmdArguments(i))
    End If
End Function


Sub append(curSettings() As MapEntry, entry As MapEntry)
Dim ln As Integer
    
    ln = UBound(curSettings) + 1
    ReDim Preserve curSettings(ln)
    curSettings(ln) = entry
End Sub

Sub appendValue(curSettings() As MapEntry, key As String, value As Variant, separator As String)
Dim sz As Integer, i As Integer

    sz = UBound(curSettings)
    For i = 1 To sz
        If curSettings(i).key = key Then
            curSettings(i).value = curSettings(i).value & separator & value
            Exit Sub
        End If
    Next i
End Sub

Sub setAndSave(scope As String, key As String, value As String)
Dim curSettings() As MapEntry, curCfgFile As String

    If scope = "app" Then
        curSettings = appSettings
        curCfgFile = appCfgFile
    End If

    setCurrentSetting curSettings, key, value
    saveFileSettings curCfgFile, curSettings
End Sub


Function stripPath(myFile As String) As String
Dim i As Integer
Dim sz As Integer
    sz = Len(myFile)
    For i = sz - 1 To 1 Step -1
        If Mid(myFile, i, 1) = "\" Then
            stripPath = Mid(myFile, i + 1)
            Exit Function
        End If
    Next i
    stripPath = myFile
End Function

Sub checkReloadCfg()
Dim reloadCfgSrc As String, reloadCfgDst As String

'    MsgBox App.EXEName & ""

    reloadCfgSrc = getCurrentSetting("reloadCfgSrc", argumentSettings)
    'trace "reloadCfgSrc = " & reloadCfgSrc
    reloadCfgDst = getCurrentSetting("reloadCfgDst", argumentSettings)
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

