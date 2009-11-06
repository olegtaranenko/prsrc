Attribute VB_Name = "AppUtils"
Option Explicit

' Ётот файл раздел€етс€ между prior, stime и rowmat.
' не использовать в cfg

Sub GridToExcel(Grid As MSFlexGrid, Optional title As String = "")

Dim objExel As Excel.Application, c As Long, r As Long
Dim I As Integer, strA() As String, begRow As Integer, str As String

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
            str = Grid.TextMatrix(r, c) '=' - наверно зарезервирован дл€ ввода формул
            Dim firstLetter As String
            firstLetter = left$(str, 1)
            If firstLetter = "=" Or firstLetter = "+" Then
                str = "'" & str
            End If
'иногда символы Cr и Lf (поле MEMO в базе) дают Err в Excel, поэтому из пол€
            I = InStr(str, vbCr) 'MEMO берем только первую строчку
            If I > 0 Then str = left$(str, I - 1)
            I = InStr(str, vbLf) 'MEMO берем только первую строчку
            If I > 0 Then str = left$(str, I - 1)
            If IsNumeric(str) And r > 0 Then
                strA(curColumn - 1) = CStr(CDbl(str))
            Else
                If Len(str) > 255 Then
                    str = left(str, 252) & "..."
                End If
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
