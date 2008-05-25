Attribute VB_Name = "CommonUtils"
Option Explicit

Type VersionInfo
    path As String
    maj As Long
    min As Long
    rev As Long
    bld As Long
End Type


Sub fatalError(msg As String)
    MsgBox msg & vbCr & "Обратитесь к администратору", vbCritical, "Критическая ошибка"
    End
End Sub

Sub getAppInfo(ByRef version As VersionInfo)
    version.path = App.path & "\" & App.exeName & ".exe"
    version.maj = App.Major
    version.min = App.Minor
    version.bld = App.Revision
    version.rev = 0
End Sub

Function compareVersion(left As VersionInfo, right As VersionInfo) As Integer
Dim lessThen As Boolean
Dim greateThen As Boolean

    If left.maj < right.maj Then
        lessThen = True
    ElseIf left.maj > right.maj Then
        greateThen = True
    End If
    
    If Not lessThen And Not greateThen Then
        If left.min < right.min Then
            lessThen = True
        ElseIf left.min > right.min Then
            greateThen = True
        End If
    End If

    If Not lessThen And Not greateThen Then
        If left.rev < right.rev Then
            lessThen = True
        ElseIf left.rev > right.rev Then
            greateThen = True
        End If
    End If

    If Not lessThen And Not greateThen Then
        If left.bld < right.bld Then
            lessThen = True
        ElseIf left.bld > right.bld Then
            greateThen = True
        End If
    End If
    
    If lessThen Then
        compareVersion = -1
        Exit Function
    End If
    If greateThen Then
        compareVersion = 1
        Exit Function
    End If
    compareVersion = 0

End Function

Function infoToString(info As VersionInfo) As String
    infoToString = info.maj & "." & info.min & "." & info.bld
End Function


Function getMainTitle() As String
Dim version As VersionInfo

    getAppInfo version
    getMainTitle = " [версия " & infoToString(version) & "]"

End Function
