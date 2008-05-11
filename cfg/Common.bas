Attribute VB_Name = "Common"
Option Explicit

Sub Main()
Dim CmdLine As String, path As String, exe As String
Dim exeHandle As Double
Dim existsFile As String

'-- Main application entry point.
Dim myFilename As String
Dim failed As Boolean


    loadEffectiveSettingsCfg

    myFilename = getFullExeName()

    existsFile = Dir(myFilename)
    If existsFile = "" Or ((GetAttr(myFilename) And vbDirectory) = vbDirectory) Then
        fatalError "Файл " & myFilename & " не обнаружен. " _
        & vbCr & "Необходимо исправить конфигурацию запуска приложения cfg.exe"
    End If
    
    
    failed = Not checkVersionExe(myFilename)
    
    If failed Then
        fatalError "При запуске программы произошла ошибка. Обратитесь к администратору."
    End If

    CmdLine = myFilename & getAppArguments
'    setAndSave "app", "lastRun", CmdLine
    
    exeHandle = Shell(CmdLine, vbNormalFocus)
    Debug.Print myFilename & getAppArguments
End Sub

Function checkVersionExe(myFilename As String) As Boolean
Dim myInfo As VersionInfo, repositoryInfo As VersionInfo
Dim m As Variant
Dim nameWithoutPath As String
Dim repositoryExe As String, repositoryPath As String
Dim check As Integer

    repositoryPath = getEffectiveSetting("repositoryPath")
    nameWithoutPath = stripPath(myFilename)
    repositoryExe = repositoryPath & "\" & nameWithoutPath
    
    checkVersionExe = True
    If GetDllVersion(myFilename, myInfo) And GetDllVersion(repositoryExe, repositoryInfo) Then
        check = compareVersion(myInfo, repositoryInfo)
        If check < 0 Then
            If vbYes = MsgBox("Обнаружена новая версия программы " & nameWithoutPath & ". " _
                & vbCr & "Текущая версия программы - " & infoToString(myInfo) _
                & vbCr & "Новая версия программы - " & infoToString(repositoryInfo) _
                & vbCr & "Нажмите Да(Yes) чтобы обновить версию" _
                , vbYesNo, "Внимание") _
            Then
                'do copy
                ShellAndHold "xcopy /y " & repositoryExe & " " & myFilename, vbHide
            End If
        End If
    Else
        MsgBox "Невозможно определить версию файла " & myFilename, vbExclamation
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

