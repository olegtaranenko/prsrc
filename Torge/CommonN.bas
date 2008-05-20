Attribute VB_Name = "CommonN"
Option Explicit

Public gKlassId As String
Public gNomNom As String
Public gSeriaId As String
Public gProduct As String
Public gProductId As String
Public gSourceId As String
Public gDocDate As Date
Public wrkDefault As Workspace
Public tbNomenk As Recordset
Public tbProduct As Recordset
Public tbDocs As Recordset
Public tbDMC As Recordset
Public tbGuide As Recordset
Public mousRight As Integer
Public nodeKey As String
Public prevRow As Long
Public gridIzLoad As Boolean
'Public DocFromKarta As Boolean
Public DMCnomNom() As String ' номер(а), кот в загруж.карточке
Public DMCklass As String ' номер группы, кот в загруж.карточке
'Public nomNoms() As String ' список номенклатуры при multiSelect
Public tmpNum As Single ' временая в т.ч. для isNunericTbox()


'Public isLoadNomenklatura As Boolean
Public numDoc As Long, numExt As Integer
Public begDate As Date ' Дата вступительных остатков
Public NN() As String, QQ() As Single ' откатываемая номенклатура и кол-во
Public QQ2() As Single ' вспомагательое откатываемое кол-во

Sub backNomenk()
Dim q As Single, i As Integer, str As String, n As Integer, rr As Integer

wrkDefault.BeginTrans

For rr = 1 To UBound(QQ)
  Set tbDMC = myOpenRecordSet("##157", "sDMC", dbOpenTable)
  If tbDMC Is Nothing Then GoTo EN1
  tbDMC.Index = "nomDoc"
  tbDMC.Seek "=", numDoc, numExt, NN(rr)
  If tbDMC.NoMatch Then
      tbDMC.Close
      GoTo EN1
  End If
  q = Round(tbDMC!quantity - QQ(rr), 2)
  If q = 0 Then
     tbDMC.Delete
  Else
    tbDMC.Edit
    tbDMC!quantity = q
    tbDMC.Update
  End If
  tbDMC.Close
  
    Set tbNomenk = myOpenRecordSet("##156", "sGuideNomenk", dbOpenTable)
    If tbNomenk Is Nothing Then GoTo EN1
    tbNomenk.Index = "PrimaryKey"
    tbNomenk.Seek "=", NN(rr)
    If tbNomenk.NoMatch Then
        tbNomenk.Close
EN1:    wrkDefault.Rollback
        MsgBox "Не удалось отменить операцию по этому документу, поэтому " & _
        "проделайте необходимые изменения вручную.", , "Внимание!"
        GoTo EN2
    End If
    tbNomenk.Edit
    tbNomenk!nowOstatki = Round(tbNomenk!nowOstatki - QQ(rr), 2)
    tbNomenk.Update
    tbNomenk.Close

Next rr
wrkDefault.CommitTrans
EN2:
ReDim QQ(0)
End Sub

Function getStrNumDoc(N_Doc As Long, N_Ext As Integer) As String
    getStrNumDoc = N_Doc
    If N_Ext = 0 Then
        getStrNumDoc = getStrNumDoc & "/"
    ElseIf N_Ext < 255 Then
        getStrNumDoc = getStrNumDoc & "/" & N_Ext
    End If
End Function
    
Function isNumericTbox(tBox As TextBox, Optional minVal, Optional maxVal) As Boolean

If checkNumeric(tBox.Text, minVal, maxVal) Then
    isNumericTbox = True
Else
    isNumericTbox = False
    tBox.SetFocus
    tBox.SelStart = 0
    tBox.SelLength = Len(tBox.Text)
End If
End Function

Function checkNumeric(val As String, Optional minVal, Optional maxVal) As Boolean

checkNumeric = True
If IsNumeric(val) Then
    tmpNum = val
    If Not IsMissing(maxVal) Then
        If (minVal > tmpNum Or tmpNum > maxVal) Then
            MsgBox "значение должно быть в диапазоне от " & minVal & _
            "  до " & maxVal, , "Error"
            checkNumeric = False
        End If
    ElseIf Not IsMissing(minVal) Then
        If minVal > tmpNum Then
            MsgBox "значение должно быть больше " & minVal
            checkNumeric = False
        End If
    End If
Else
    MsgBox "Недопустимое значение", , "Error"
    checkNumeric = False
End If
End Function

Sub clearGrid(Grid As MSFlexGrid)
Dim il As Long
 For il = Grid.Rows To 3 Step -1
    Grid.RemoveItem (il)
 Next il
 clearGridRow Grid, 1
End Sub

Function byErrSqlGetValues(ParamArray val() As Variant) As Boolean
Dim tabl As Recordset, i As Integer, maxi As Integer

'getFieldBySql = False
byErrSqlGetValues = False
maxi = UBound(val())
If maxi < 1 Then
    MsgBox "мало параметров для п\п byErrSqlGetValues()"
    Exit Function
End If

Set tabl = myOpenRecordSet(CStr(val(0)), CStr(val(1)), dbOpenDynaset) 'dbOpenForwardOnly)
If tabl Is Nothing Then Exit Function
If tabl.BOF Then Exit Function
tabl.MoveFirst
For i = 2 To maxi
    If TypeName(val(i)) = "Single" And IsNull(tabl.Fields(i - 2)) Then
        val(i) = 0
    Else
        val(i) = tabl.Fields(i - 2)
    End If
Next i
'getFieldBySql = True
byErrSqlGetValues = True
tabl.Close
End Function
' фактически это и блокировка и соовт. записей в sDMC(rez)
' nnExt=0 исп-ся для разблок-ки оставшегося резерва при частичном списании
Function docLock(Optional unLok As String = "", Optional nnExt) As Boolean
Dim str As String
If IsMissing(nnExt) Then nnExt = numExt

Set tbDocs = myOpenRecordSet("##158", "sDocs", dbOpenTable) 'dbOpenForwardOnly)
If tbDocs Is Nothing Then Exit Function

docLock = False
tbDocs.Index = "PrimaryKey"
tbDocs.Seek "=", numDoc, nnExt
If tbDocs.NoMatch Then
    MsgBox "Похоже документ уже удалили", , "Error - 166"
Else
    tbDocs.Edit ' блокируем
    str = tbDocs!rowLock
    If str <> "" And str <> Documents.cbM.Text Then
       tbDocs.Update ' снимаем блокировку
       If unLok = "" Then _
       MsgBox "Документ '" & getStrNumDoc(tbDocs!numDoc, tbDocs!numExt) & _
       "' временно занят другим менеджером (" & str & ")"
       GoTo EN1
    End If
    If unLok = "" Then
        tbDocs!rowLock = Documents.cbM.Text
    Else
        tbDocs!rowLock = ""
    End If
    tbDocs.Update
    docLock = True
End If
EN1:
tbDocs.Close
End Function

'используется в 1 месте
Function strWhereByStEndDateBox(frm As Form) As String
Dim strWhere As String, addNullDate As String, stDate As String, enDate As String
  If frm.cbStartDate.value = 1 Then
    stDate = "(sDocs.xDate)>='" & Format(frm.tbStartDate.Text, "yyyy-mm-dd") & "'"
    addNullDate = ""
 Else
    stDate = ""
    addNullDate = " OR (sDocs.xDate) Is Null"
 End If

 If frm.cbEndDate.value = 1 Then
    enDate = "(sDocs.xDate)<='" & Format(frm.tbEndDate.Text, "yyyy-mm-dd") & " 11:59:59 PM'"
 Else
    enDate = ""
 End If
 If stDate <> "" And enDate <> "" Then
    strWhere = stDate & " AND " & enDate
 ElseIf stDate <> "" Or enDate <> "" Then
    strWhere = stDate & enDate
 Else
    addNullDate = ""
    strWhere = ""
 End If
 strWhereByStEndDateBox = strWhere & addNullDate

End Function

Sub textBoxInGridCell(tb As TextBox, Grid As MSFlexGrid)
    tb.Width = Grid.CellWidth + 50
'    tb.Text = Grid.TextMatrix(mousRow, mousCol)
    tb.Text = Grid.TextMatrix(Grid.row, Grid.col)
    tb.Left = Grid.CellLeft + Grid.Left
    tb.Top = Grid.CellTop + Grid.Top
    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)
    tb.Visible = True
    tb.SetFocus
    tb.ZOrder
    Grid.Enabled = False 'иначе курсор по ней бегает
End Sub

'не записыват неуникальное значение, для полей, где такие
'значения запрещены. А  генерит при этом error?
Function ValueToTableField(myErrCod As String, value As String, Table As String, _
field As String, Optional by As String = "") As Boolean
Dim sql As String, byStr As String  ', numOrd As String

ValueToTableField = False
If value = "" Then value = Chr(34) & Chr(34)
If by = "" Then
    byStr = ".numOrder)= " & gNzak
ElseIf by = "byFirmId" Then
    byStr = ".FirmId)= " & gFirmId
ElseIf by = "byKlassId" Then
    byStr = ".klassId)= " & gKlassId
ElseIf by = "byNomNom" Then
    byStr = ".nomNom)= " & "'" & gNomNom & "'"
ElseIf by = "bySeriaId" Then
    byStr = ".seriaId)= " & gSeriaId
ElseIf by = "byProductId" Then
    byStr = ".prId)= " & gProductId
'ElseIf by = "bySourceId" Then
'    byStr = ".sourceId)= " & gSourceId
ElseIf by = "byNumDoc" Then
    sql = "UPDATE " & Table & " SET " & Table & "." & field & "=" & value _
        & " WHERE (((" & Table & ".numDoc)=" & numDoc & " AND (" & Table & _
        ".numExt)=" & numExt & " ));"
    GoTo AA
Else
    Exit Function
End If
sql = "UPDATE " & Table & " SET " & Table & "." & field & _
" = " & value & " WHERE (((" & Table & byStr & " ));"
AA:
'MsgBox "sql = " & sql
If myExecute(myErrCod, sql) = 0 Then ValueToTableField = True
End Function




