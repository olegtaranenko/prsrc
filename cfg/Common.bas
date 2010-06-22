Attribute VB_Name = "Common"
Option Explicit
Dim ErrNumber As Integer, ErrDescription As String
Dim gMsg As String



'-- Main application entry point.
Sub Main()

    Dim localExe As String

    Dim msgStyle As VbMsgBoxStyle

    loadEffectiveSettingsCfg
    initLogFileName

    printWindowsVersion
    
    If getEffectiveSetting("exe") <> "" Then
        localExe = getFullExeName(getEffectiveSetting("exe"))
        If Not exeStart(localExe) Then
            fatalError "Unexpected error " & Err.Description
        End If
    ElseIf getCurrentSetting("help", argumentSettings) <> "" Then
        
    ElseIf getCurrentSetting("deploy", argumentSettings) <> "" Then
        localExe = getFullExeName(getCurrentSetting("deploy", argumentSettings))
        msgStyle = exeDeploy(localExe)
        MsgBox gMsg _
        & vbCr & vbCr & "���� :" & vbTab & localExe, msgStyle
        'MsgBox "������ ��� ������� �������� ���� " & localExe & " � �����������.", msgStyle
    End If
    
End Sub

Private Function isReporsitoryExists() As String

    isReporsitoryExists = getEffectiveSetting("repositoryPath")
    If isReporsitoryExists <> "" Then
        On Error GoTo failed
        If ((GetAttr(isReporsitoryExists) And vbDirectory) = vbDirectory) Then
            'ok
            Exit Function
        End If
    End If
    
failed:
    isReporsitoryExists = ""
    MsgBox "Unexpected error by checking the repository" _
        & vbCr & "Check network access or/and correct program settings" _
        & vbCr & "Setting to repository path: " & isReporsitoryExists, vbCritical, "Error"
        
End Function


Function exeDeploy(localExe As String) As VbMsgBoxStyle
Dim myInfo As VersionInfo, repositoryInfo As VersionInfo
'Dim m As Variant
Dim nameWithoutPath As String
Dim repositoryExe As String, repositoryPath As String
Dim check As Integer
Dim doCopy As Boolean
Dim remoteFound As Boolean
    
On Error GoTo failed
    exeDeploy = vbInformation

    repositoryPath = getEffectiveSetting("repositoryPath")
    nameWithoutPath = stripPath(localExe)
    repositoryExe = repositoryPath & "\" & nameWithoutPath

    gMsg = ""
    doCopy = False
    If GetDllVersion(localExe, myInfo) Then
        If GetDllVersion(repositoryExe, repositoryInfo) Then
            
            check = compareVersion(myInfo, repositoryInfo)
            If vbYes = MsgBox("File to deploy: " & vbTab & localExe & ". " _
                & vbCr & "Repository: " & vbTab & repositoryPath _
                & vbCr & "Local version: " & vbTab & infoToString(myInfo) _
                & vbCr & "Remote version: " & vbTab & infoToString(repositoryInfo) _
                & vbCr & "Do you want to deploy?" _
                , vbYesNo Or vbExclamation, "��������") _
            Then
                doCopy = True
            Else
                gMsg = "Deploing canceled..."
            End If
        Else
            gMsg = "Unexpected error by access to repository." _
            & vbCr & "Repository path: " & repositoryExe
            GoTo failed
        End If
    Else
        gMsg = "Local file " & localExe & " not found!"
        GoTo failed
    End If
    
    If doCopy Then
        'gMsg = "������ ��������� ������ ������ " & myInfo & " � " & repositoryInfo
        If compareVersion(myInfo, repositoryInfo) <> 0 Then
            Dim archivePath As String, archiveExe As String
            archivePath = getArchevePath(repositoryPath, repositoryInfo)
            archiveExe = archivePath & "\" & getArchiveName(nameWithoutPath, repositoryInfo)
            gMsg = "������ ������������� ����� " & repositoryExe & " � ����� " & archivePath
            FileCopy repositoryExe, archiveExe
        End If
        gMsg = "������ ����������� ����� " & localExe & " � ����� " & repositoryPath
        FileCopy localExe, repositoryExe
        gMsg = "Successful deploy!"
    End If
    Exit Function
failed:
    exeDeploy = vbCritical
End Function


' returns true if succes false otherwise
Private Function exeStart(localExe As String) As String

Dim existsFile As String, repositoryPath  As String
Dim failed As Boolean
Dim CmdLine As String, path As String, exe As String
Dim exeHandle As Double


    exeStart = False
    
    On Error GoTo failed
    existsFile = Dir(localExe)
    If existsFile = "" Or ((GetAttr(localExe) And vbDirectory) = vbDirectory) Then
        trace "Exe filename = "
        erro "File " & localExe & " does not exists"
        fatalError "���� " & localExe & " �� ���������. " _
        & vbCr & "������ ������� cfg.exe: " & Command() _
        & vbCr & "���������� ��������� ������������ ������� ���������� cfg.exe"
    End If
    
    
    repositoryPath = isReporsitoryExists
    If repositoryPath <> "" Then
        ' ��������� ���� ����. ���� ����� ���������� ����� ������ ������
        ' �� ����� ����� ���������� ��������� ��������� cfg.exe
        failed = checkSelfVersion(localExe, repositoryPath)
        If failed Then
            fatalError "O����� ��� �������� ������ ����������� ��������� cfg.exe"
        End If
        failed = checkVersionExe(localExe, repositoryPath)
        If failed Then
            If ErrNumber = 70 Then
                fatalError "O����� ��� ����������� ����� ������ ����� " & localExe _
                & vbCr & "�������� ����� ������ ��������� �� ������ ��� ���������� ������ ���������" _
                & vbCr & "���� ������ �� ����� -> "
            Else
                fatalError "O����� ��� �������� ������ ��������� " & localExe _
                & vbCr & "����� ������: " & ErrNumber & " - " & ErrDescription
            End If
        End If
    Else
        MsgBox "�� ��������� (��� �������� �����) ����������� ���������� ��������" _
        & vbCr & "������ ����� ���������� � ������� ������� ��� ����������� ���������� �� ����� ������", vbCritical, "���������� � ��������������"
    End If
    
    CmdLine = localExe & getAppArguments
    
    exeHandle = Shell(CmdLine, vbNormalFocus)
'        Debug.Print localExe & getAppArguments
    exeStart = True
    Exit Function
failed:
    
End Function


Function checkSelfVersion(ByVal localExe As String, repositoryPath As String) As Boolean
Dim myInfo As VersionInfo, repositoryInfo As VersionInfo
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
            ' ��������� ���������� ...
            trace "cmd line " & cmd
            Dim handle As Variant

            handle = Shell(cmd, vbNormalFocus)

            trace "after executing shell. Exe handle = " & handle
            '... � ����� ����
            End
        End If
    Else
        GoTo EN1
    End If
    
    Exit Function
EN1:
    checkSelfVersion = True
End Function



Function getArchiveName(name As String, info As VersionInfo) As String
    getArchiveName = name & "." & infoToString(info)
End Function

Function getArchevePath(path As String, info As VersionInfo) As String
Dim archivePart As String
archivePart = getEffectiveSetting("archive", "archive")
    getArchevePath = path & "\" & archivePart
End Function

Function checkVersionExe(localExe As String, repositoryPath As String) As Boolean
Dim myInfo As VersionInfo, repositoryInfo As VersionInfo
Dim m As Variant
Dim nameWithoutPath As String
Dim repositoryExe As String
Dim check As Integer
Dim doCopy As Boolean

On Error GoTo EN1
    nameWithoutPath = stripPath(localExe)
    repositoryExe = repositoryPath & "\" & nameWithoutPath

    checkVersionExe = False
    doCopy = False
    Dim localOk As Boolean, remoteOk As Boolean
On Error GoTo localFail
    localOk = GetDllVersion(localExe, myInfo)
    GoTo checkRemote
    
localFail:
    localOk = False
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    
checkRemote:
On Error GoTo remoteFail
    remoteOk = GetDllVersion(repositoryExe, repositoryInfo)
    GoTo mainCheck
remoteFail:
    remoteOk = False
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    
mainCheck:
On Error GoTo mainFail
        
    If localOk And remoteOk Then
        check = compareVersion(myInfo, repositoryInfo)
        If check < 0 Then
            If vbYes = MsgBox("���������� ����� ������ ��������� " & nameWithoutPath & ". " _
                & vbCr & "������� ������ ��������� - " & infoToString(myInfo) _
                & vbCr & "����� ������ ��������� - " & infoToString(repositoryInfo) _
                & vbCr & "������� ��(Yes) ����� �������� ������" _
                , vbYesNo, "��������") _
            Then
                doCopy = True
            End If
        End If
    ElseIf remoteOk Then
        ' ����������� ����������� �� ����������� ��������� ������ ����� � ������� ����������
        If vbYes = MsgBox( _
             " ������� ����������: " & vbTab & myInfo.path _
            & vbCr & "�����������: " & vbTab & repositoryInfo.path _
            & vbCr & "���������: " & vbTab & nameWithoutPath & " ������ [ " & infoToString(myInfo) & "]" _
            & vbCr & "������� ��(Yes) ����� ������� ���� �� ����." _
            , vbYesNo, "��������") _
        Then
            doCopy = True
        End If
    ElseIf localOk Then
        
    Else
        MsgBox "���������� ���������� ������ ����� " & localExe, vbExclamation
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
    
mainFail:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    checkVersionExe = False
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
    appAliases(1).Key = "prior": appAliases(1).Value = "PriorN.exe"
    appAliases(2).Key = "stime": appAliases(2).Value = "stimeN.exe"
    appAliases(3).Key = "rowmat": appAliases(3).Value = "Rowmat_N.exe"
    
    getExeName = alias
    For I = 1 To UBound(appAliases)
        If getExeName = appAliases(I).Key Then
            getExeName = appAliases(I).Value
            Exit Function
        End If
    Next I
    
    If Len(getExeName) > 0 Then
        If Not Right(getExeName, 4) = ".exe" Then
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

