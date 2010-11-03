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


Function getMapEntry(ByRef map() As MapEntry, Key As String) As Variant
Dim I As Integer
    For I = 1 To UBound(map)
        If map(I).Key = Key Then
            getMapEntry = map(I).Value
            Exit Function
        End If
    Next I
'    getMapEntry = Null
End Function

' возвращает индекс key в массиве map
' Empty если не найдет такой параметр
Function getMapEntryIndex(ByRef map() As MapEntry, Key As String) As Integer
Dim I As Integer
    For I = 1 To UBound(map)
        If map(I).Key = Key Then
            getMapEntryIndex = I
            Exit Function
        End If
    Next I
    getMapEntryIndex = -1
End Function


Sub append(map() As MapEntry, entry As MapEntry)
Dim ln As Integer
    
    ln = UBound(map) + 1
    ReDim Preserve map(ln)
    map(ln) = entry
End Sub

Sub appendKeyValue(map() As MapEntry, Key As String, ByRef Value)
Dim entry As MapEntry

    entry.Key = Key
    entry.Value = Value
    append map, entry
End Sub

Sub appendUnique(map() As MapEntry, Key As String, ByRef Value)
Dim Index As Integer

    Index = getMapEntryIndex(map, Key)
    If Index = -1 Then
        Dim entry As MapEntry
        entry.Key = Key
        Set entry.Value = Value
        append map, entry
    End If
End Sub


Function setMapEntry(ByRef map() As MapEntry, Key As String, Value As Variant) As Variant
Dim exists As MapEntry

    'Set exists = getMapEntry(map, Key)
    'If Not IsEmpty(exists) Then
    '    exists.Value = Value
    '    setMapEntry = exists
    'Else
    '    Dim Entry As MapEntry
    '    Entry = New MapEntry
    '    Entry.Key = Key
    '    Entry.Value = Value
    '    append map, Entry
    '    setMapEntry = Empty
    'End If

End Function


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
    version.path = App.path & "\" & App.ExeName & ".exe"
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
    getMainTitle = " [" & infoToString(version) & "]"

End Function

