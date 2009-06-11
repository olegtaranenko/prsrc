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
    version.path = App.path & "\" & App.EXEName & ".exe"
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

Function existsInTreeview(ByRef tTree As TreeView, Key As String) As Boolean
Dim I As Integer
    For I = 1 To tTree.Nodes.Count
        If tTree.Nodes(I).Key = Key Then
            existsInTreeview = True
            Exit Function
        End If
    Next I
    existsInTreeview = False
End Function


Sub GridToExcel(Grid As MSFlexGrid, Optional title As String = "")

Dim objExel As Excel.Application, c As Long, r As Long
Dim i As Integer, strA() As String, begRow As Integer, str As String

begRow = 3
If title = "" Then begRow = 1

Set objExel = New Excel.Application
objExel.Visible = True
objExel.SheetsInNewWorkbook = 1
objExel.Workbooks.Add
With objExel.ActiveSheet
.Cells(1, 2).value = title
ReDim Preserve strA(Grid.Cols + 1)
For r = 0 To Grid.Rows - 1
    Dim curColumn As Integer
    curColumn = 1
    For c = 1 To Grid.Cols - 1
        If Grid.colWidth(c) > 0 Then
            str = Grid.TextMatrix(r, c) '=' - наверно зарезервирован для ввода формул
            If left$(str, 1) = "=" Then str = "." & str
'иногда символы Cr и Lf (поле MEMO в базе) дают Err в Excel, поэтому из поля
            i = InStr(str, vbCr) 'MEMO берем только первую строчку
            If i > 0 Then str = left$(str, i - 1)
            i = InStr(str, vbLf) 'MEMO берем только первую строчку
            If i > 0 Then str = left$(str, i - 1)
            If IsNumeric(str) And r > 0 Then
                strA(curColumn - 1) = CStr(CDbl(str))
            Else
                strA(curColumn - 1) = str
            End If
            curColumn = curColumn + 1
        End If
    Next c
'    On Error Resume Next
   .Range(.Cells(begRow + r, 1), .Cells(begRow + r, Grid.Cols)).FormulaArray = strA
Next r

'objExel.ActiveSheet.Range("A" & begRow & ":U" & Grid.Rows + begRow).FormulaArray = strA
'.Range(.Cells(begRow, 1), .Cells(Grid.Rows + begRow, Grid.Rows)).FormulaArray = strA
End With
Set objExel = Nothing
End Sub
