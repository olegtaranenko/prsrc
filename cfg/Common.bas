Attribute VB_Name = "Common"
Option Explicit

Sub Main()
Dim CmdLine As String, path As String, exe As String
Dim exeHandle As Double
Dim existsFile As String

'-- Main application entry point.
Dim myFilename As String
Dim failed As Boolean


    parsedArgs = GetCommandLine
        
    myFilename = getFullExeName()
    
    existsFile = Dir(myFilename)
    If existsFile = "" Or ((GetAttr(myFilename) And vbDirectory) = vbDirectory) Then
        MsgBox "Файл " & myFilename & " не обнаружен. " _
        & vbCr & "Необходимо исправить конфигурацию запуска приложения cfg.exe", vbExclamation, "Ошибка, свяжитесь с администратором"
        End
    End If
    
    
    failed = Not checkVersionExe(myFilename)
    
    If failed Then
        MsgBox "При запуске программы произошла ошибка. Обратитесь к администратору.", vbExclamation, "Ошибка"
        End
    End If

    'exeHandle = Shell(myFilename & " " & getAppArguments, vbNormalFocus)
    Debug.Print myFilename & getAppArguments
End Sub

Function checkVersionExe(myFilename As String) As Boolean
Dim myLoadpath As String
Dim maj As Long, min As Long, rev As Long, bld As Long
Dim m As Variant

    If (GetDllVersion(myFilename, myLoadpath, maj, min, rev, bld)) Then
'        MsgBox "File " & myFilename & " is vesion " & CStr(maj) & _
               "." & CStr(min) & "." & CStr(rev) & "." & CStr(bld), _
               vbInformation, "VerRes"
        checkVersionExe = True
    Else
        MsgBox "Невозможно определить версию файла " & myFilename, vbExclamation
        checkVersionExe = False
    End If
End Function

Function getAppArguments()
    getAppArguments = delegateArguments("arg")
End Function
