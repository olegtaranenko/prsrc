Attribute VB_Name = "CmdUtils"
Option Explicit

Dim rawCmdArguments() As String
Public appCfgFile As String
Public siteCfgFile As String

Type MapEntry
    key As String
    value As Variant
End Type



Sub parseCommandLine(Optional MaxArgs)
   'Declare variables.
   Dim c, CmdLine, CmdLnLen, InArg, I, NumArgs
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
      'Test for space or tab.
      If (c <> " " And c <> vbTab) Then
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




Function getFullExeName() As String
Dim path As String, exe As String
    exe = getExeName()
    If Not isAbsolute(exe) Then
        getFullExeName = getExePath() & "\" & exe
    Else
        getFullExeName = exe
    End If
End Function

Function getExePath() As String
Dim absolute As Boolean

    getExePath = getCurrentSetting("path", argumentSettings)
    
    If getExePath <> "" Then
        ' check if path is relative or absolute
    End If
    If Not isAbsolute(getExePath) Then
        getExePath = App.path & "\" & getExePath
    End If
    If Len(getExePath) = 0 Then
        getExePath = App.path
    End If
End Function


Function getExeName() As String
Dim I As Integer
Dim hasExt As Boolean
Dim appAliases(3) As MapEntry
    appAliases(1).key = "prior": appAliases(1).value = "PriorN.exe"
    appAliases(2).key = "stime": appAliases(2).value = "stimeN.exe"
    appAliases(3).key = "rowmat": appAliases(3).value = "Rowmat_N.exe"
    
    getExeName = getCurrentSetting("exe", argumentSettings)
    For I = 1 To UBound(appAliases)
        If getExeName = appAliases(I).key Then
            getExeName = appAliases(I).value
            Exit Function
        End If
    Next I
    
    If Len(getExeName) > 0 Then
        If Not Right(getExeName, 4) = ".exe" Then
            getExeName = getExeName & ".exe"
        End If
    End If
End Function


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

Dim I As Integer
Dim argTokens As MapEntry
Dim argSnippet As String

    For I = 1 To UBound(rawCmdArguments)
        If rawCmdArguments(I) = "-" & key Then
            argTokens = tokenizeKeyValue(rawCmdArguments(I + 1))
            If argTokens.value <> "" Then
                argSnippet = separator & "-" & argTokens.key & " " & argTokens.value
            Else
                argSnippet = separator & argTokens.key
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

    equalPos = InStr(1, match, delimiter, vbTextCompare)
    If equalPos <> 0 Then
        result.key = trimAll(Left(match, equalPos - 1))
        result.value = trimAll(Mid(match, equalPos + 1))
    Else
        result.key = match
        result.value = ""
    End If
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
        MsgBox "������ ��������� ������� ���������", vbExclamation, "���������� � ��������������"
        End
    End If
    appCfgFile = getCurrentSetting("appCfgFile", argumentSettings)
    If appCfgFile = "" Then
        appCfgFile = getAppCfgDefaultName
    End If
    If Not loadFileSettings(appCfgFile, appSettings) Then
        MsgBox "������ ��� �������� ����� ������������ ��������� (appCfgFile)", vbExclamation, "���������� � ��������������"
        End
    End If
    siteCfgFile = getCurrentSetting("siteCfgFile", argumentSettings)
    If siteCfgFile = "" Then
        siteCfgFile = getCurrentSetting("siteCfgFile", appSettings)
    End If
    If siteCfgFile = "" Then
        siteCfgFile = getSiteCfgDefaultName
    End If
    If Not loadFileSettings(appCfgFile, siteSettings) Then
        MsgBox "������ ��� �������� ����� ������������ �������� ����� (siteCfgFile)", vbExclamation, "���������� � ��������������"
        End
    End If
    buildEffectiveSettings
End Sub

Function loadFileSettings(filePath As String, ByRef curSettings() As MapEntry) As Boolean
Dim entry As MapEntry

    Dim str As String, str2 As String, I As Integer, j As Integer
    str = filePath
    ReDim curSettings(0)
    
    On Error GoTo EN1 '���� ������� ����� ����������, �� Dir ���� ERR
    If Dir$(str) = "" Then
        loadFileSettings = False
    Else
      Open str For Input As #1
      While Not EOF(1)
        Line Input #1, str
        entry = tokenizeKeyValue(str)
        append curSettings, entry
      Wend
      Close #1
      
      loadFileSettings = True
    End If
    Exit Function
EN1:
    loadFileSettings = False
End Function

Sub saveFileSettings(filePath As String, ByRef curSettings() As MapEntry)
Dim I As Integer, str  As String
Dim doSave As Boolean

    str = filePath
    On Error GoTo EN1
    Open str For Output As #1
    For I = 1 To UBound(curSettings)
        Print #1, curSettings(I).key & " = " & curSettings(I).value
    Next I
EN1:
    On Error Resume Next
    Close #1
End Sub

Function getAppCfgDefaultName() As String
    getAppCfgDefaultName = App.path & "\" & App.EXEName & ".cfg"
End Function

Function getSiteCfgDefaultName() As String
    getSiteCfgDefaultName = App.path & "\..\site.cfg"
End Function

Function getCurrentSetting(key As String, ByRef curSettings() As MapEntry) As Variant
Dim I As Integer
    For I = 1 To UBound(curSettings)
        If curSettings(I).key = key Then
            getCurrentSetting = curSettings(I).value
            Exit Function
        End If
    Next I
'    getCurrentSetting = Null
End Function

' ���������� ������ key � ������� curSettings
' Empty ���� �� ������ ����� ��������
Function getMapEntry(ByRef curSettings() As MapEntry, key As String) As Integer
Dim I As Integer
    For I = 1 To UBound(curSettings)
        If curSettings(I).key = key Then
            getMapEntry = I
            Exit Function
        End If
    Next I
    getMapEntry = Empty
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
        exists = getCurrentSetting(entry.key, mergeTo)
        If IsEmpty(exists) Then
           append mergeTo, entry
        End If
    Next I

End Sub

Public Sub setCurrentSetting(curSettings() As MapEntry, key As String, paramVal)
Dim I As Integer
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
Dim I As Integer
Dim entry As MapEntry, exists
Dim value As String

    ReDim argumentSettings(0)
    For I = 1 To UBound(rawCmdArguments)
        If isKey(rawCmdArguments(I)) Then
            If isNotKey(I + 1) Then
                entry.key = Mid(rawCmdArguments(I), 2)
                value = rawCmdArguments(I + 1)
                exists = getCurrentSetting(entry.key, curSettings)
                If Not IsEmpty(exists) Then
                    exists.value = exists.value & " " & value
                Else
                    entry.value = value
                    append argumentSettings, entry
                End If
                I = I + 1
            Else
                exists = getCurrentSetting(rawCmdArguments(I), curSettings)
                If IsNull(exists) Then
                    entry.key = rawCmdArguments(I)
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
