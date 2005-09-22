VERSION 5.00
Begin VB.Form Pribil 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Реализация"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmStat1 
      Caption         =   "Статистика"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3600
      TabIndex        =   33
      Top             =   3540
      Width           =   1215
   End
   Begin VB.CommandButton cmStat2 
      Caption         =   "Статистика"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5040
      TabIndex        =   32
      Top             =   3540
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Caption         =   " Услуги"
      Height          =   2415
      Left            =   2100
      TabIndex        =   27
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton cmStat3 
         Caption         =   "Статистика"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   31
         Top             =   1980
         Width           =   1215
      End
      Begin VB.CommandButton cmDetail3 
         Caption         =   "Детализация"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   28
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label laRealiz1 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   30
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label laUslug 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   29
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Итого"
      Height          =   2415
      Left            =   6420
      TabIndex        =   23
      Top             =   1560
      Width           =   1275
      Begin VB.Label laRealiz 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Width           =   915
      End
      Begin VB.Label laClear 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   25
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label laMaterials 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   24
         Top             =   720
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Продажа"
      Height          =   2415
      Left            =   4980
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton cmDetail2 
         Caption         =   "Детализация"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   19
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label laMaterials2 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   22
         Top             =   720
         Width           =   915
      End
      Begin VB.Label laClear2 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   21
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label laRealiz2 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Материалы не под заказы:"
      Height          =   795
      Left            =   3540
      TabIndex        =   15
      Top             =   660
      Width           =   2715
      Begin VB.CommandButton cmDetail 
         Caption         =   "Детализация"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label laOther 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Товары"
      Height          =   2415
      Left            =   3540
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton cmDetail1 
         Caption         =   "Детализация"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label laProduct 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   915
      End
      Begin VB.Label laMaterials1 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   720
         Width           =   915
      End
      Begin VB.Label laClear1 
         Alignment       =   1  'Правая привязка
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Фиксировано один
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   1140
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   -60
      TabIndex        =   3
      Top             =   -120
      Width           =   8715
      Begin VB.TextBox tbStartDate 
         Height          =   285
         Left            =   960
         MaxLength       =   8
         TabIndex        =   7
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox tbEndDate 
         Height          =   285
         Left            =   1980
         MaxLength       =   8
         TabIndex        =   6
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton cmManag 
         Caption         =   "Применить"
         Height          =   315
         Left            =   2880
         TabIndex        =   5
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmSave 
         Caption         =   "Записать в журнал Х.О. "
         Enabled         =   0   'False
         Height          =   315
         Left            =   6300
         TabIndex        =   4
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label laPeriod 
         Caption         =   "Период  с "
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   795
      End
      Begin VB.Label laPo 
         Caption         =   "по"
         Height          =   195
         Left            =   1785
         TabIndex        =   8
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Реализац. - материалы:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label laHMaterials 
      Caption         =   "Материалы под заказы:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2340
      Width           =   1875
   End
   Begin VB.Label laHRealiz 
      Caption         =   "Реализация:"
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "Pribil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pDateWhere As String, nDateWhere As String, uDateWhere As String
Public bDateWhere As String, mDateWhere As String
Dim begDateHron As Date ' Начало ведения хронологии

Private Sub cmDetail_Click()
Me.MousePointer = flexHourglass
Report.param1 = laOther.Caption
Report.Regim = "mat"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmDetail1_Click()
Me.MousePointer = flexHourglass
'Report.param1 = laRealiz1.Caption
Report.param1 = laProduct.Caption
Report.param2 = laMaterials1.Caption
Report.Regim = ""
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmDetail2_Click()
Me.MousePointer = flexHourglass
Report.param1 = laRealiz2.Caption
Report.param2 = laMaterials2.Caption
Report.Regim = "bay"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmDetail3_Click()
Me.MousePointer = flexHourglass
'Report.param1 = laRealiz1.Caption
Report.param1 = laUslug.Caption
'Report.param2 = laMaterials1.Caption
Report.Regim = "uslug"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmManag_Click() 'кнопка применить из отч.Реализация
Dim oborot As Single, dohod As Single, s2 As Single, s As Single


Me.MousePointer = flexHourglass

strWhere = getWhereByDateBoxes(Me, "xPredmetyByIzdeliaOut.outDate", begDateHron) ' между
If strWhere = "error" Then GoTo EN1
If strWhere <> "" Then strWhere = " WHERE ((" & strWhere & "))"
pDateWhere = strWhere

'хронология по изделиям
'sql = "SELECT numOrder, prId, prExt, quant FROM xPredmetyByIzdeliaOut" & _
strWhere & ";"
sql = "SELECT xPredmetyByIzdeliaOut.*, xPredmetyByIzdelia.cenaEd " & _
"FROM xPredmetyByIzdelia INNER JOIN xPredmetyByIzdeliaOut ON " & _
"(xPredmetyByIzdelia.prExt = xPredmetyByIzdeliaOut.prExt) AND " & _
"(xPredmetyByIzdelia.prId = xPredmetyByIzdeliaOut.prId) AND " & _
"(xPredmetyByIzdelia.numOrder = xPredmetyByIzdeliaOut.numOrder)" & strWhere & ";"

'MsgBox sql
Set tbProduct = myOpenRecordSet("##306", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then GoTo EN1


oborot = 0: dohod = 0
If Not tbProduct.BOF Then
 While Not tbProduct.EOF
    gNzak = tbProduct!numOrder
    gProductId = tbProduct!prId
    prExt = tbProduct!prExt
    oborot = oborot + getProductNomenkSum * tbProduct!quant
    dohod = dohod + tbProduct!cenaEd * tbProduct!quant
'    s = getProductNomenkSum * tbProduct!quant
'    tmpDate = tbProduct!outDate
'If gNzak = 3120105 Then
'    gNzak = gNzak
'End If
    tbProduct.MoveNext
 Wend
End If
tbProduct.Close

tmpStr = strWhere ' если далее нужна будет детализация
'хронология по отд.номеклатурам
strWhere = getWhereByDateBoxes(Me, "xPredmetyByNomenkOut.outDate", begDateHron) ' между
If strWhere <> "" Then strWhere = " WHERE ((" & strWhere & "))"
nDateWhere = strWhere

sql = "SELECT xPredmetyByNomenkOut.quant, sGuideNomenk.cost, sGuideNomenk.perList, " & _
"xPredmetyByNomenk.cenaEd FROM (sGuideNomenk INNER JOIN xPredmetyByNomenk " & _
"ON sGuideNomenk.nomNom = xPredmetyByNomenk.nomNom) INNER JOIN " & _
"xPredmetyByNomenkOut ON (xPredmetyByNomenk.nomNom = " & _
"xPredmetyByNomenkOut.nomNom) AND (xPredmetyByNomenk.numOrder =" & _
"xPredmetyByNomenkOut.numOrder)" & strWhere & ";"
'" xPredmetyByNomenkOut.numOrder, " &

'MsgBox sql
Set tbNomenk = myOpenRecordSet("##307", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    oborot = oborot + tbNomenk!cost * tbNomenk!quant / tbNomenk!perList '!!!
'    gNzak = tbNomenk!numOrder:  s = tbNomenk!cost * tbNomenk!quant
    dohod = dohod + tbNomenk!cenaEd * tbNomenk!quant
    tbNomenk.MoveNext
  Wend
End If
tbNomenk.Close
    
dohod = Round(dohod, 2)
laProduct.Caption = Format(dohod, "0.00")

oborot = Round(oborot, 2)
laMaterials1.Caption = Format(oborot, "0.00")
laClear1.Caption = Format(dohod - oborot, "0.00")

'услуги
strWhere = getWhereByDateBoxes(Me, "xUslugOut.outDate", begDateHron) ' между
If strWhere <> "" Then strWhere = " WHERE ((" & strWhere & "))"
uDateWhere = strWhere
sql = "SELECT Sum(quant) AS Sum_quant from xUslugOut " & strWhere & ";"
'MsgBox sql
If byErrSqlGetValues("##380", sql, s2) Then
    s2 = Round(s2, 2)
    laUslug.Caption = Format(s2, "0.00")
    laRealiz1.Caption = laUslug.Caption
    dohod = Round(dohod + s2, 2)
End If
'laRealiz1.Caption = Format(dohod, "0.00")
'oborot = Round(oborot, 2)
'laMaterials1.Caption = Format(oborot, "0.00")
'laClear1.Caption = Format(dohod - oborot, "0.00")


'Отгружено по  заказам продаж
strWhere = getWhereByDateBoxes(Me, "sDocs.xDate", begDateHron) ' между
bDateWhere = strWhere
'sql = "SELECT Sum(BayOrders.shipped) AS Sum_shipped " & _
"FROM sDocs INNER JOIN BayOrders ON sDocs.numDoc = BayOrders.numOrder " & _
"WHERE (((sDocs.numExt)=1)"
'до и после 18.06.04 были и numExt=2 но почему работает для заказа с 2мя накладными
'If strWhere = "" Then
'    sql = sql & ");"
'Else
'    sql = sql & " AND (" & strWhere & "));"
'End If
If strWhere <> "" Then strWhere = " WHERE ((" & strWhere & "))"
sql = "SELECT Sum([sDMC].[quant]*[sDMCrez].[intQuant]/[sGuideNomenk].[perList]) AS bSum " & _
"FROM sGuideNomenk INNER JOIN ((BayOrders INNER JOIN sDocs ON BayOrders.numOrder = sDocs.numDoc) INNER JOIN (sDMC INNER JOIN sDMCrez ON sDMC.nomNom = sDMCrez.nomNom) ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc) AND (BayOrders.numOrder = sDMCrez.numDoc)) ON sGuideNomenk.nomNom = sDMC.nomNom " & _
strWhere & ";"

'MsgBox sql
s = s2 = 0
If byErrSqlGetValues("##431", sql, s) Then
    s = Round(s, 2)
    laRealiz2.Caption = Format(s, "0.00")
    dohod = Round(dohod + s, 2)
End If

'Материалы заказов продаж (дата списания считается датой отгрузки)
'If strWhere <> "" Then strWhere = " WHERE ((" & strWhere & "))"


sql = "SELECT Sum([sGuideNomenk].[cost]*[sDMC].[quant]/[sGuideNomenk]." & _
"[perList]) AS cena FROM sGuideNomenk INNER JOIN ((sDocs INNER JOIN " & _
"BayOrders ON sDocs.numDoc = BayOrders.numOrder) INNER JOIN sDMC ON " & _
"(sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc)) ON " & _
"sGuideNomenk.nomNom = sDMC.nomNom" & strWhere & ";"
'MsgBox sql
If byErrSqlGetValues("##430", sql, s2) Then
    s2 = Round(s2, 2)
    laMaterials2.Caption = Format(s2, "0.00")
    oborot = Round(oborot + s2, 2)
End If
laClear2.Caption = Format(s - s2, "0.00")

laRealiz.Caption = Format(dohod, "0.00")
laMaterials.Caption = Format(oborot, "0.00")
laClear.Caption = Format(dohod - oborot, "0.00")

'материалы не под заказ
strWhere = getWhereByDateBoxes(Me, "sDocs.xDate", begDateHron) ' между
If strWhere <> "" Then strWhere = "(" & strWhere & ") AND "
mDateWhere = strWhere & "((sDocs.numExt)=254) AND ((sDocs.destId)>-1000 And (sDocs.destId)<0)"
'!!!
sql = "SELECT Sum([sDMC].[quant]*[sGuideNomenk].[cost]/[sGuideNomenk].[perList]) AS sum " & _
"FROM sGuideSource INNER JOIN (sGuideNomenk INNER JOIN (sDocs INNER JOIN sDMC ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc)) ON sGuideNomenk.nomNom = sDMC.nomNom) ON sGuideSource.sourceId = sDocs.destId " & _
"WHERE (" & mDateWhere & ");"
'MsgBox sql
If byErrSqlGetValues("##404", sql, s) Then
    laOther.Caption = Format(s, "0.00")
End If

EN1:
Me.MousePointer = flexDefault
cmSave.Enabled = True
cmDetail.Enabled = True
cmDetail1.Enabled = True
cmDetail2.Enabled = True
cmDetail3.Enabled = True
cmStat1.Enabled = True
cmStat2.Enabled = True
cmStat3.Enabled = True

End Sub

Sub addRowToBook(sum As String)
Dim mig As Single, str As String

' As Integer,  As Integer,  As Integer,
'detailId As Integer, purposeId As Integer, KredDebitor
tbDocs.AddNew
mig = Timer
While (Timer - mig < 1#): DoEvents: Wend  ' Дата в сек. д.б. уникальной

tmpStr = Format(Now(), "dd/mm/yy hh:nn:ss")
tbDocs!xDate = CDate(tmpStr)
tbDocs!UEsumm = sum
tbDocs!m = AUTO.cbM.Text
tbDocs!debit = debit
tbDocs!subDebit = subDebit
tbDocs!kredit = kredit
tbDocs!subKredit = subKredit
str = tbStartDate.Text & " - " & tbEndDate.Text
tbDocs!note = str
tbDocs!purposeId = purposeId
'tbDocs!detailId = detailId

On Error GoTo ER1
tbDocs.Update
Journal.addRowToGrid sum, str
Exit Sub
ER1:
errorCodAndMsg "378" '##378

'MsgBox Error, , "Ошибка 378-" & Err & ":  " '##378

End Sub

Private Sub cmMaterials_Click()

End Sub

Private Sub cmRealiz_Click()

End Sub

Private Sub cmSave_Click() '$$4
Dim str As String, str2 As String, i As Integer
cmSave.Enabled = False
'If Not (IsNumeric(laRealiz.Caption) And IsNumeric(laRealiz.Caption)) Then

str = "В Настройках не заданы параметры записи для поля '"
str2 = "' в журнал Х.О., поэтому эта запись произведена НЕ будет."

Set tbDocs = myOpenRecordSet("##376", "yBook", dbOpenTable) 'dbOpenForwardOnly)
'If tbDocs Is Nothing Then Exit Sub

strWhere = "SELECT 1,Debit, subDebit, Kredit, subKredit, pId as purposeId " & _
"From yGuidePurpose WHERE (((auto)='"

sql = strWhere & "r'));"
If Not byErrSqlGetValues("W##401", sql, i, debit, subDebit, kredit, subKredit, _
purposeId) Then GoTo EN1
If i = 0 Then 'err WHERE
    MsgBox str & laHRealiz.Caption & str2, , "Предупреждение"
    GoTo EN1
Else
    addRowToBook laRealiz.Caption
End If

sql = strWhere & "z'));"
If Not byErrSqlGetValues("W##402", sql, i, debit, subDebit, kredit, subKredit, _
purposeId) Then GoTo EN1
If i = 0 Then 'err WHERE
    MsgBox str & laHMaterials.Caption & str2, , "Предупреждение"
    GoTo EN1
Else
    addRowToBook laMaterials.Caption
End If

sql = strWhere & "m'));"
If Not byErrSqlGetValues("W##403", sql, i, debit, subDebit, kredit, subKredit, _
purposeId) Then GoTo EN1
If i = 0 Then 'err WHERE
    MsgBox str & laOther.Caption & str2, , "Предупреждение"
Else
    addRowToBook laOther.Caption
End If
EN1:
tbDocs.Close
Unload Me
On Error Resume Next
Journal.Grid.SetFocus
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmStat1_Click()
Me.MousePointer = flexHourglass
'Report.param1 = laRealiz1.Caption
Report.param1 = laProduct.Caption
Report.param2 = laMaterials1.Caption
Report.Regim = "relizStatistic"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmStat2_Click()
Me.MousePointer = flexHourglass
Report.param1 = laRealiz2.Caption
Report.param2 = laMaterials2.Caption
Report.Regim = "bayStatistic"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmStat3_Click()

Me.MousePointer = flexHourglass
Report.param1 = laUslug.Caption
Report.Regim = "uslugStatistic"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub Form_Load()
tbEndDate.Text = Format(CurDate, "dd/mm/yy")
begDateHron = "01.09.03" '
tbStartDate.Text = "01." & Format(CurDate, "mm/yy")
'tbStartDate.Text = "01.09.03"
'tbStartDate.Text = "01.06.04"
End Sub
'для отчета Прибыль
Function getProductNomenkSum() As Variant
Dim i As Integer, j As Integer, gr() As String, sum As Single

getProductNomenkSum = Null
'вариантная ном-ра изделия

'sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup " & _
"FROM sProducts INNER JOIN xVariantNomenc ON (sProducts.nomNom = " & _
"xVariantNomenc.nomNom) AND (sProducts.ProductId = xVariantNomenc.prId) " & _
"WHERE (((xVariantNomenc.numOrder)=" & numDoc & ") AND (" & _
"(xVariantNomenc.prId)=" & gProductId & ") AND ((xVariantNomenc.prExt)=" & prExt & "));"
'!!!
sql = "SELECT sProducts.xgroup, [sGuideNomenk].[cost]*[sProducts].[quantity]" & _
"/[sGuideNomenk].[perList] AS sum " & _
"FROM (sGuideNomenk INNER JOIN sProducts ON sGuideNomenk.nomNom = " & _
"sProducts.nomNom) INNER JOIN xVariantNomenc ON (sProducts.nomNom = " & _
"xVariantNomenc.nomNom) AND (sProducts.ProductId = xVariantNomenc.prId) " & _
"WHERE (((xVariantNomenc.numOrder)=" & gNzak & ") AND (" & _
"(xVariantNomenc.prId)=" & gProductId & ") AND ((xVariantNomenc.prExt)=" & prExt & "));"

'MsgBox sql
Set tbNomenk = myOpenRecordSet("##192", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
ReDim gr(0): i = 0: sum = 0
While Not tbNomenk.EOF
    i = i + 1
    sum = sum + tbNomenk!sum
    ReDim Preserve gr(i): gr(i) = tbNomenk!xgroup
    tbNomenk.MoveNext
Wend
tbNomenk.Close
    
'НЕвариантная ном-ра изделия
'sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup " & _
"From sProducts WHERE (((sProducts.ProductId)=" & gProductId & "));"
'[cost]-цена фактическая !!!
sql = "SELECT sProducts.xgroup, [sGuideNomenk].[cost]*[sProducts].[quantity]" & _
"/[sGuideNomenk].[perList] AS sum " & _
"FROM sGuideNomenk INNER JOIN sProducts ON sGuideNomenk.nomNom = sProducts.nomNom " & _
"WHERE (((sProducts.ProductId)=" & gProductId & "));"

'MsgBox sql
Set tbNomenk = myOpenRecordSet("##177", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
While Not tbNomenk.EOF
    For j = 1 To UBound(gr) ' если группа состоит из одной ном-ры, то она
        If gr(j) = tbNomenk!xgroup Then GoTo NXT ' НЕвариантна, т.к. не
    Next j                                      ' не попала в xVariantNomenc
    i = i + 1
    sum = sum + tbNomenk!sum
NXT: tbNomenk.MoveNext
Wend
tbNomenk.Close
getProductNomenkSum = sum
End Function

