Attribute VB_Name = "CmdUtils"
Option Explicit

Dim rawCmdArguments() As String
Public argumentSettings() As MapEntry
Public appSettings() As MapEntry
Public siteSettings() As MapEntry
Public settings() As MapEntry




Public appCfgFile As String
Public siteCfgFile As String


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

' ������� �� �������������, ��� � ���-�� ��������� ������������ ������ ����
' url = prior. ���������� �� � ����� ������� ����� delimiter ��������
' � key � value ��������������.
' ���� � ��������� ��� ������ delimiter, �� key �������� �������� match, � value - ������ ������.
' ������ �������� � key � value
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
    
    If rawValue = "������" Or rawValue = "true" Or rawValue = "������" Or rawValue = "True" Then
        castValue = True
    ElseIf rawValue = "����" Or rawValue = "����" Or rawValue = "False" Or rawValue = "false" Then
        castValue = False
    Else
        castValue = rawValue
    End If
    result.Value = castValue
    tokenize = result
    
End Function

'����� ��� � ��� �������� ������� � vbTab
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
        fatalError "������ ��������� � ��������� ������"
    End If
    appCfgFile = getMapEntry(argumentSettings, "appCfgFile")
    If appCfgFile = "" Then
        appCfgFile = getAppCfgDefaultName
    End If
    If loadFileSettings(appCfgFile, appSettings) < 0 Then
        fatalError "������ ��� �������� ����� ������������ ��������� (appCfgFile)"
    End If
    siteCfgFile = getMapEntry(argumentSettings, "siteCfgFile")
    If siteCfgFile = "" Then
        siteCfgFile = getMapEntry(appSettings, "siteCfgFile")
    End If
    If siteCfgFile = "" Then
        siteCfgFile = getSiteCfgDefaultName
    End If
    If loadFileSettings(siteCfgFile, siteSettings) < 0 Then
        fatalError "������ ��� �������� ����� ������������ �������� ����� (siteCfgFile)"
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
Dim Entry As MapEntry

    Dim str As String, str2 As String, I As Integer, J As Integer
    str = filePath
    ReDim curSettings(0)
    
    On Error GoTo EN1 '���� ������� ����� ����������, �� Dir ���� ERR
    If Dir$(str) = "" Then
        ' ���������� ����� ������������ �� �������� �������
        loadFileSettings = 0
    Else
      Open str For Input As #1
      While Not EOF(1)
        Line Input #1, str
        Entry = tokenizeKeyValue(str)
        append curSettings, Entry
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
Dim Entry As MapEntry
Dim exists As Variant

    lnFrom = UBound(mergeFrom)
    For I = 1 To lnFrom
        Entry = mergeFrom(I)
        exists = getMapEntry(mergeTo, Entry.Key)
        If IsEmpty(exists) Then
           append mergeTo, Entry
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
        Dim Entry As MapEntry
        Entry.Key = Key
        Entry.Value = paramVal
        append curSettings, Entry
    End If
    
End Sub

Function getEffectiveSetting(Key As String, Optional defaultValue) As Variant
Dim Entry As MapEntry, Value

    Value = getMapEntry(settings, Key)
    If Not IsEmpty(Value) Then
        getEffectiveSetting = Value
        Exit Function
    End If
    If Not IsMissing(defaultValue) Then
        getEffectiveSetting = defaultValue
    End If
    
End Function


Function loadCmdSettings(curSettings() As MapEntry) As Boolean
'������� ���������� ������, ����� �������� �������� ��� ��������.
'� ���� ������ ��� �������� ����� Null.
'������ � ������ stime -otlad -dostup a -devel �������� otlad ����� �������� Null.
'���� �� stime -dostup a -devel, �� ����� �������� otlad ����� �������� Empty.

Dim I As Integer
Dim Entry As MapEntry, exists As Variant
Dim Value As Variant

    ReDim argumentSettings(0)
    For I = 1 To UBound(rawCmdArguments)
        Entry.Value = Null
        If isKey(rawCmdArguments(I)) Then
            Entry.Key = Mid(rawCmdArguments(I), 2)
            If isNotKey(I + 1) Then
                Value = Null
                If I + 1 <= UBound(rawCmdArguments) Then
                    Value = rawCmdArguments(I + 1)
                    I = I + 1
                End If
                
                exists = getMapEntry(curSettings, Entry.Key)
                If Not IsEmpty(exists) Then
                    exists = exists & " " & Value
                    appendValue curSettings, Entry.Key, Value, " "
                Else
                    Entry.Value = Value
                    append argumentSettings, Entry
                End If
            Else
                exists = getMapEntry(curSettings, rawCmdArguments(I))
                If IsNull(exists) Then
                    Entry.Key = rawCmdArguments(I)
                    append argumentSettings, Entry
                ElseIf Not IsEmpty(Entry.Key) Then
                    append argumentSettings, Entry
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

    reloadCfgSrc = getMapEntry(argumentSettings, "reloadCfgSrc")
    'trace "reloadCfgSrc = " & reloadCfgSrc
    reloadCfgDst = getMapEntry(argumentSettings, "reloadCfgDst")
    'trace "reloadCfgDst = " & reloadCfgDst

    If reloadCfgSrc <> "" Then
        MsgBox "���������� ������������� ���������� ����������� ���������." _
            & vbCr & "��� ���������� �������������, ����� ���� ����� ����� ��������� ��������� ��� ���.", vbExclamation, "��������������"
        On Error GoTo EN1

        FileCopy reloadCfgSrc, reloadCfgDst
        End
    End If
    Exit Sub
EN1:
    fatalError "������ ��� ����������� ����������� ��������� cfg.exe" & vbCr & Err.Description
    End
End Sub

