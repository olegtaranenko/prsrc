Attribute VB_Name = "AppUtils"
Option Explicit

Dim quantity As Long

'константы для whoReserved
Public Const rtNomZak = 1
Public Const rtReserv = 2
Public Const rtCeh = 3
Public Const rtData = 4
Public Const rtMen = 5
Public Const rtStatus = 6
Public Const rtFirma = 7
Public Const rtProduct = 8
Public Const rtZakazano = 9
Public Const rtOplacheno = 10



' Этот файл разделяется между prior, stime и rowmat.
' не использовать в cfg


Sub GridToExcel(Grid As MSFlexGrid, Optional title As String = "")
Dim ColWidth As String, Note As String


Dim objExel As Excel.Application, c As Long, r As Long
Dim I As Integer, strA() As String, begRow As Integer, str As String

begRow = 3
If title = "" Then begRow = 1

Set objExel = New Excel.Application
objExel.Visible = True
objExel.SheetsInNewWorkbook = 1
objExel.Workbooks.Add
With objExel.ActiveSheet
.Cells(1, 2).Value = title
ReDim Preserve strA(Grid.Cols + 1)
For r = 0 To Grid.Rows - 1
    Dim curColumn As Integer
    curColumn = 1
    For c = 1 To Grid.Cols - 1
        If Grid.ColWidth(c) > 0 Then
            str = Grid.TextMatrix(r, c) '=' - наверно зарезервирован для ввода формул
            Dim firstLetter As String
            firstLetter = Left$(str, 1)
            Dim doEscape As Boolean
            
            If firstLetter = "=" Or firstLetter = "+" Then
                doEscape = True
            End If
            
            If str = "--" Then
                doEscape = True
            End If
            
            If doEscape Then
                str = "'" & str
            End If
'иногда символы Cr и Lf (поле MEMO в базе) дают Err в Excel, поэтому из поля
            I = InStr(str, vbCr) 'MEMO берем только первую строчку
            If I > 0 Then str = Left$(str, I - 1)
            I = InStr(str, vbLf) 'MEMO берем только первую строчку
            If I > 0 Then str = Left$(str, I - 1)
            If IsNumeric(str) And r > 0 Then
                strA(curColumn - 1) = CStr(CDbl(str))
            Else
                If Len(str) > 255 Then
                    str = Left(str, 252) & "..."
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



Function newNumorder(Value As String) As Numorder
    Dim ret As Numorder
    Set ret = New Numorder
    ret.val = Value
    Set newNumorder = ret
End Function


Function getNextDocNum() As Long
Dim valueorder As Numorder

    Set valueorder = New Numorder
    valueorder.val = getSystemField("lastDocNum")
    If valueorder.IsEmpty Then
        valueorder.docs = True
    End If
    If Not valueorder.isCurrentDay Then
        Set valueorder = New Numorder
        valueorder.docs = True
    End If
    valueorder.nextNum
    sql = "UPDATE SYSTEM SET lastDocNum = " & valueorder.val
    'Debug.Print sql
    myBase.Execute (sql)
    numDoc = valueorder.val
    
    getNextDocNum = valueorder.val
    
End Function

Sub showRezerv(ByVal dostupOstatki As Double, ByVal factOstatki As Double, ByVal Edizm2 As String, ByRef f As Form)
    If Round(factOstatki, 2) > Round(dostupOstatki, 2) Then
        If MsgBox("Если Вы хотите просмотреть список всех заказов, под " & _
        "которые была зарезервирована эта номенклатура, нажмите <Да>.", _
        vbYesNo Or vbDefaultButton2, "Посмотреть, кто резервировал? '" & _
        gNomNom & "' ?") = vbYes Then
            Report.Edizm2 = Edizm2
            Report.Regim = "whoRezerved"
            Set Report.Caller = f
            Report.Sortable = True
            Report.Show vbModal
        End If
    Else
        MsgBox "Эта номенклатура никем не резервировалась.", , ""
    End If

End Sub



Function whoRezerved(ByRef Grid As MSFlexGrid, Optional p_term_index As Integer = 0) As Integer
Dim groupklassid As Integer, rowStr As String
Dim p_days_start As Integer, p_days_end As Integer
    

    Grid.Visible = False
    
    Grid.FormatString = "|<№ заказа|>кол-во|^Цех |^Дата |^ М|Статус" & _
    "|<Название Фирмы|<Изделия|>Заказано|>Согласовано"
    Grid.ColWidth(0) = 0
    'Grid.ColWidth(rtNomZak) =
    Grid.ColWidth(rtReserv) = 765
    Grid.ColWidth(rtCeh) = 765
    Grid.ColWidth(rtData) = 1600
    'Grid.ColWidth(rtMen) =
    Grid.ColWidth(rtStatus) = 930
    Grid.ColWidth(rtFirma) = 3270
    Grid.ColWidth(rtProduct) = 1950
    'Grid.ColWidth(rtZakazano) =
    Grid.ColWidth(rtOplacheno) = 810

    If p_term_index = 0 Then
        p_days_start = 10000
        p_days_end = 0
    ElseIf p_term_index = 1 Then
        p_days_start = 30
        p_days_end = 0
    ElseIf p_term_index = 2 Then
        p_days_start = 60
        p_days_end = 30
    ElseIf p_term_index = 3 Then
        p_days_start = 120
        p_days_end = 60
    ElseIf p_term_index = 4 Then
        p_days_start = 180
        p_days_end = 120
    ElseIf p_term_index = 5 Then
        p_days_start = 10000
        p_days_end = 120
    End If
    
    
    sql = "call wf_order_reserved ('" & gNomNom & "', " & p_days_start & ", " & p_days_end & ")"
    
    Set tbOrders = myOpenRecordSet("##350", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Function
    quantity = 0
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF

            quantity = quantity + 1
            Grid.TextMatrix(quantity, rtNomZak) = tbOrders!Numorder
            Grid.TextMatrix(quantity, rtReserv) = Format(tbOrders!quant, "# ##0.00")
            If Not IsNull(tbOrders!Ceh) Then _
                Grid.TextMatrix(quantity, rtCeh) = tbOrders!Ceh
            
            Grid.TextMatrix(quantity, rtData) = tbOrders!date1
            If Not IsNull(tbOrders!Manager) Then _
                Grid.TextMatrix(quantity, rtMen) = tbOrders!Manager
            
            If Not IsNull(tbOrders!Status) Then _
                Grid.TextMatrix(quantity, rtStatus) = tbOrders!Status
            
            If Not IsNull(tbOrders!client) Then _
                Grid.TextMatrix(quantity, rtFirma) = tbOrders!client
            
            If Not IsNull(tbOrders!Note) Then _
                Grid.TextMatrix(quantity, rtProduct) = tbOrders!Note
            
            If Not IsNull(tbOrders!sm_zakazano) Then _
                Grid.TextMatrix(quantity, rtZakazano) = Format(tbOrders!sm_zakazano, "# ##0.00")
                
            If Not IsNull(tbOrders!sm_paid) Then _
                Grid.TextMatrix(quantity, rtOplacheno) = Format(tbOrders!sm_paid, "# ##0.00")
                
            Grid.AddItem ""
            tbOrders.MoveNext
        Wend
    End If
  tbOrders.Close

'laRecSum.Caption = Round(sum, 2)
If quantity > 0 Then
    Grid.removeItem quantity + 1
End If
trigger = False
Grid.Visible = True
whoRezerved = quantity


End Function


Function RuDate2Date(ByVal dt As String) As Variant
Dim dotPos As Integer
Dim v_dd As Integer
Dim v_mm As Integer
Dim v_yyyy As Integer
    
    On Error GoTo catch
    dotPos = InStr(dt, ".")
    If IsNull(dotPos) Or dotPos = 0 Then
        v_dd = CInt(dt)
        dt = ""
    Else
        v_dd = CInt(Left(dt, dotPos))
        dt = Mid(dt, dotPos + 1)
    End If
    
    If Len(dt) > 0 Then
        dotPos = InStr(dt, ".")
        If IsNull(dotPos) Or dotPos = 0 Then
            v_mm = CInt(dt)
            dt = ""
        Else
            v_mm = CInt(Left(dt, dotPos))
            dt = Mid(dt, dotPos + 1)
        End If
    Else
        v_mm = Format(Now(), "mm")
    End If
    
    If Len(dt) <= 2 And Len(dt) > 0 Then
        v_yyyy = CInt("20" & Format(CInt(dt), "0#"))
    Else
        v_yyyy = Format(Now(), "yyyy")
    End If
    
    RuDate2Date = CDate(CStr(v_yyyy) & "-" & Format(v_mm, "0#") & "-" & Format(v_dd, "0#"))
    Exit Function
catch:
    RuDate2Date = False
End Function

Function myIsDate(ByVal str As String) As Variant
    Dim dt As Variant
    dt = RuDate2Date(str)
    If IsDate(dt) Then
        myIsDate = Format(dt, "dd.mm.yy")
    Else
        myIsDate = dt
    End If
End Function
