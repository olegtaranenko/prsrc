Attribute VB_Name = "Common"
Option Explicit

'-- Main application entry point.
Sub Main()

Dim localExe As String

Dim success As Boolean

    loadEffectiveSettingsCfg
    initLogFileName

    printWindowsVersion
    
    If getEffectiveSetting("exe") <> "" Then
        localExe = getFullExeName(getEffectiveSetting("exe"))
        success = exeStart(localExe)
    ElseIf getCurrentSetting("deploy", argumentSettings) <> "" Then
        localExe = getFullExeName(getCurrentSetting("deploy", argumentSettings))
        success = exeDeploy(localExe)
        If success Then
            MsgBox "Файл " & localExe & " успешно выложен в репозиторий.", vbInformation, "Сообщение"
        Else
            MsgBox "Ошибка при попытке выложить файл " & localExe & " в репозиторий.", vbCritical, "Ошибка"
        End If
    End If
    
    If Not success Then
        fatalError "Unexpected error " & Err.Description
    End If
End Sub


' returns true if succes false otherwise
Private Function exeDeploy(localExe As String) As Boolean
Dim failed As Boolean
    exeDeploy = False
    On Error GoTo failed
    
    failed = checkVersionDeploy(localExe)
    
    If failed Then
        fatalError "Error by deploy " & localExe & "."
    End If
    exeDeploy = True
    Exit Function
failed:
End Function


' returns true if succes false otherwise
Private Function exeStart(localExe As String) As Boolean

Dim existsFile As String
Dim failed As Boolean
Dim CmdLine As String, path As String, exe As String
Dim exeHandle As Double


    exeStart = False
    
    On Error GoTo failed
    existsFile = Dir(localExe)
    If existsFile = "" Or ((GetAttr(localExe) And vbDirectory) = vbDirectory) Then
        trace "Exe filename = "
        erro "File " & localExe & " does not exists"
        fatalError "Файл " & localExe & " не обнаружен. " _
        & vbCr & "Строка запуска cfg.exe: " & Command() _
        & vbCr & "Необходимо исправить конфигурацию запуска приложения cfg.exe"
    End If
    
    failed = checkSelfVersion(localExe)
    If failed Then
        fatalError "Oшибка при проверке версии управляющей программы."
    End If
    
    ' проверяем сами себя. Если вдруг обнаружена более свежая версия
    ' то тогда через клиентскую программу обновляем cfg.exe
    failed = checkVersionExe(localExe)
    
    If failed Then
        fatalError "Oшибка при проверке версии программы " & localExe & "."
    End If

    CmdLine = localExe & getAppArguments
    
    exeHandle = Shell(CmdLine, vbNormalFocus)
'        Debug.Print localExe & getAppArguments
    exeStart = True
    Exit Function
failed:
    
End Function


Function checkSelfVersion(ByVal localExe As String) As Boolean
Dim myInfo As VersionInfo, repositoryInfo As VersionInfo
Dim repositoryPath As String
Dim repositoryExe As String, appExe As String

    trace "checkSelfVersion() start..."
    checkSelfVersion = False             ' optimistic view
    
On Error GoTo EN1
    repositoryPath = getEffectiveSetting("repositoryPath")
    repositoryExe = repositoryPath & "\" & App.exeName & ".exe"
    
    getAppInfo myInfo
    trace "myInfo: " & infoToString(myInfo)
    
    If GetDllVersion(repositoryExe, repositoryInfo) Then
        trace "repository info: " & infoToString(repositoryInfo)
        
        If compareVersion(myInfo, repositoryInfo) < 0 Then
            Dim cmd As String
            cmd = localExe & " -reloadCfgSrc " & repositoryExe & " -reloadCfgDst " & myInfo.path
            ' запустить обновление ...
            trace "cmd line " & cmd
            Dim handle As Variant

            handle = Shell(cmd, vbNormalFocus)

            trace "after executing shell. Exe handle = " & handle
            '... и молча уйти
            End
        End If
    Else
        GoTo EN1
    End If
    
    Exit Function
EN1:
    checkSelfVersion = True
End Function

Function checkVersionDeploy(localExe As String) As Boolean
Dim myInfo As VersionInfo, repositoryInfo As VersionInfo
'Dim m As Variant
Dim nameWithoutPath As String
Dim repositoryExe As String, repositoryPath As String
Dim check As Integer
Dim doCopy As Boolean
Dim remoteFound As Boolean

On Error GoTo EN1
    repositoryPath = getEffectiveSetting("repositoryPath")
    nameWithoutPath = stripPath(localExe)
    repositoryExe = repositoryPath & "\" & nameWithoutPath

    checkVersionDeploy = False
    doCopy = False
    If GetDllVersion(localExe, myInfo) Then
        If GetDllVersion(repositoryExe, repositoryInfo) Then
            
            check = compareVersion(myInfo, repositoryInfo)
            If vbYes = MsgBox("Deploy of " & localExe & ". " _
                & vbCr & "Local version - " & infoToString(myInfo) _
                & vbCr & "Remote version - " & infoToString(repositoryInfo) _
                & vbCr & "Deploy?" _
                , vbYesNo, "Внимание") _
            Then
                doCopy = True
            End If
            remoteFound = True
        End If
    Else
        MsgBox "Local file " & localExe & " not found!", vbExclamation
        checkVersionDeploy = True
    End If
    If doCopy Then
        If compareVersion(myInfo, repositoryInfo) <> 0 And remoteFound Then
            Dim archivePath As String, archiveExe As String
            archivePath = getArchevePath(repositoryPath, repositoryInfo)
            archiveExe = getArchiveName(nameWithoutPath, repositoryInfo)
            FileCopy repositoryExe, archivePath & "\" & archiveExe
        End If
        FileCopy localExe, repositoryExe
    End If
    Exit Function
EN1:
    checkVersionDeploy = True
End Function

Function getArchiveName(name As String, info As VersionInfo) As String
    getArchiveName = name & "." & infoToString(info)
End Function

Function getArchevePath(path As String, info As VersionInfo) As String
Dim archivePart As String
archivePart = getEffectiveSetting("archive", "archive")
    getArchevePath = path & "\" & archivePart
End Function

Function checkVersionExe(localExe As String) As Boolean
Dim myInfo As VersionInfo, repositoryInfo As VersionInfo
Dim m As Variant
Dim nameWithoutPath As String
Dim repositoryExe As String, repositoryPath As String
Dim check As Integer
Dim doCopy As Boolean

On Error GoTo EN1
    repositoryPath = getEffectiveSetting("repositoryPath")
    nameWithoutPath = stripPath(localExe)
    repositoryExe = repositoryPath & "\" & nameWithoutPath

    checkVersionExe = False
    doCopy = False
    If getEffectiveSetting("overwrite") = 1 Then
        doCopy = True
    ElseIf GetDllVersion(localExe, myInfo) Then
        If GetDllVersion(repositoryExe, repositoryInfo) Then
            
            check = compareVersion(myInfo, repositoryInfo)
            If check < 0 Then
                If vbYes = MsgBox("Обнаружена новая версия программы " & nameWithoutPath & ". " _
                    & vbCr & "Текущая версия программы - " & infoToString(myInfo) _
                    & vbCr & "Новая версия программы - " & infoToString(repositoryInfo) _
                    & vbCr & "Нажмите Да(Yes) чтобы обновить версию" _
                    , vbYesNo, "Внимание") _
                Then
                    doCopy = True
                End If
            End If
        Else
            checkVersionExe = True
        End If
    Else
        MsgBox "Невозможно определить версию файла " & localExe, vbExclamation
        checkVersionExe = True
    End If
    If doCopy Then
        'ShellAndHold "xcopy /y " & repositoryExe & " " & localExe, vbHide
        info "Found update for " & localExe & ", version: " & infoToString(myInfo)
        info "New file: " & repositoryExe & " Version: " & infoToString(repositoryInfo)
        FileCopy repositoryExe, localExe
        info "... update was successful"
    End If
    Exit Function
EN1:
    checkVersionExe = True
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
        getExePath = App.path & "\" & getExePath
    End If
    If Len(getExePath) = 0 Then
        getExePath = App.path
    End If
End Function


Function getExeName(alias As String) As String
Dim I As Integer
Dim hasExt As Boolean
Dim appAliases(3) As MapEntry
    appAliases(1).Key = "prior": appAliases(1).value = "PriorN.exe"
    appAliases(2).Key = "stime": appAliases(2).value = "stimeN.exe"
    appAliases(3).Key = "rowmat": appAliases(3).value = "Rowmat_N.exe"
    
    getExeName = alias
    For I = 1 To UBound(appAliases)
        If getExeName = appAliases(I).Key Then
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


Function getFullExeName(alias As String) As String
Dim path As String, exe As String
    exe = getExeName(alias)
    If Not isAbsolute(exe) Then
        getFullExeName = getExePath() & "\" & exe
    Else
        getFullExeName = exe
    End If
End Function

