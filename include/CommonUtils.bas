Attribute VB_Name = "CommonUtils"
Option Explicit

Type MapEntry
    Key As String
    Value As Variant
End Type


Type VersionInfo
    path As String
    maj As Long
    min As Long
    rev As Long
    bld As Long
End Type


Sub fatalError(msg As String, Optional lookAdmin As String)
    Dim adminMsg As String
    If IsMissing(lookAdmin) Then
        adminMsg = "Обратитесь к администратору"
    Else
        adminMsg = lookAdmin
    End If
    MsgBox msg & vbCr & adminMsg, vbCritical, "Критическая ошибка"
    End
End Sub

Sub getAppInfo(ByRef version As VersionInfo)
    version.path = App.path & "\" & App.EXEName & ".exe"
    version.maj = App.Major
    version.min = App.Minor
    version.bld = App.Revision
    version.rev = 0
End Sub

Function compareVersion(Left As VersionInfo, Right As VersionInfo) As Integer
Dim lessThen As Boolean
Dim greateThen As Boolean

    If Left.maj < Right.maj Then
        lessThen = True
    ElseIf Left.maj > Right.maj Then
        greateThen = True
    End If
    
    If Not lessThen And Not greateThen Then
        If Left.min < Right.min Then
            lessThen = True
        ElseIf Left.min > Right.min Then
            greateThen = True
        End If
    End If

    If Not lessThen And Not greateThen Then
        If Left.rev < Right.rev Then
            lessThen = True
        ElseIf Left.rev > Right.rev Then
            greateThen = True
        End If
    End If

    If Not lessThen And Not greateThen Then
        If Left.bld < Right.bld Then
            lessThen = True
        ElseIf Left.bld > Right.bld Then
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

