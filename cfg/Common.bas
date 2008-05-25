Attribute VB_Name = "Common"
Option Explicit

Sub Main()
Dim CmdLine As String, path As String, exe As String
Dim exeHandle As Double
Dim existsFile As String

'-- Main application entry point.
Dim localExe As String
Dim failed As Boolean


    loadEffectiveSettingsCfg
    initLogFileName

    
    printWindowsVersion
    
    localExe = getFullExeName()

    existsFile = Dir(localExe)
    If existsFile = "" Or ((GetAttr(localExe) And vbDirectory) = vbDirectory) Then
        dbg "Exe filename = "
        erro "File " & localExe & " does not exists"
        fatalError "Файл " & localExe & " не обнаружен. " _
        & vbCr & "Необходимо исправить конфигурацию запуска приложения cfg.exe"
    End If
    
    Dim repositoryExe As String, currentExe As String
    failed = checkSelfVersion(localExe)
    If failed Then
        fatalError "При проверке версии управляющей программы произошла ошибка."
    End If
    
    
    failed = Not checkVersionExe(localExe)
    
    If failed Then
        fatalError "При проверке версии программы " & localExe & "произошла ошибка."
    End If

    CmdLine = localExe & getAppArguments
'    setAndSave "app", "lastRun", CmdLine
    
    exeHandle = Shell(CmdLine, vbNormalFocus)
    Debug.Print localExe & getAppArguments
End Sub


Function checkSelfVersion(ByVal localExe As String) As Boolean
Dim myInfo As VersionInfo, repositoryInfo As VersionInfo
Dim repositoryPath As String
Dim repositoryExe As String, appExe As String

    checkSelfVersion = True             ' optimistic view
    
On Error GoTo EN1
    repositoryPath = getEffectiveSetting("repositoryPath")
    repositoryExe = repositoryPath & "\" & App.exeName & ".exe"
    
    getAppInfo myInfo
    If GetDllVersion(repositoryExe, repositoryInfo) Then
        If compareVersion(myInfo, repositoryInfo) < 0 Then
            Dim cmd As String
            cmd = localExe & " -reloadCfgSrc " & repositoryExe & " -reloadCfgDst " & myInfo.path
            ' запустить обновление ...
            Shell cmd, vbHide
            '... и молча уйти
            End
        End If
    Else
        GoTo EN1
    End If
    
    Exit Function
EN1:
    checkSelfVersion = False
End Function


Function checkVersionExe(localExe As String) As Boolean
Dim myInfo As VersionInfo, repositoryInfo As VersionInfo
Dim m As Variant
Dim nameWithoutPath As String
Dim repositoryExe As String, repositoryPath As String
Dim check As Integer

    repositoryPath = getEffectiveSetting("repositoryPath")
    
    nameWithoutPath = stripPath(localExe)
    repositoryExe = repositoryPath & "\" & nameWithoutPath
    
    checkVersionExe = True
    If GetDllVersion(nameWithoutPath, myInfo) Then
    If GetDllVersion(repositoryExe, repositoryInfo) Then
        
        check = compareVersion(myInfo, repositoryInfo)
        If check < 0 Then
            If vbYes = MsgBox("Обнаружена новая версия программы " & nameWithoutPath & ". " _
                & vbCr & "Текущая версия программы - " & infoToString(myInfo) _
                & vbCr & "Новая версия программы - " & infoToString(repositoryInfo) _
                & vbCr & "Нажмите Да(Yes) чтобы обновить версию" _
                , vbYesNo, "Внимание") _
            Then
                'do copy
                ShellAndHold "xcopy /y " & repositoryExe & " " & localExe, vbHide
            End If
        End If
        End If
    Else
        MsgBox "Невозможно определить версию файла " & localExe, vbExclamation
        checkVersionExe = False
    End If
End Function

Function getAppArguments()
    getAppArguments = delegateArguments("arg")
End Function


Function getExePath() As String
Dim absolute As Boolean

    getExePath = getCurrentSetting("path", argumentSettings)
    
    If getExePath <> "" Then
        ' check if path is relative or absolute
    End If
    If Not isAbsolute(getExePath) Then
        getExePath = App.path & getExePath
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
        If Not right(getExeName, 4) = ".exe" Then
            getExeName = getExeName & ".exe"
        End If
    End If
End Function


Function getFullExeName() As String
Dim path As String, exe As String
    exe = getExeName()
    If Not isAbsolute(exe) Then
        getFullExeName = getExePath() & "\" & exe
    Else
        getFullExeName = exe
    End If
End Function

