VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Otgruz 
   BackColor       =   &H8000000A&
   Caption         =   "Отгрузка"
   ClientHeight    =   3360
   ClientLeft      =   168
   ClientTop       =   456
   ClientWidth     =   11688
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   11688
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   300
      TabIndex        =   9
      Top             =   2940
      Width           =   855
   End
   Begin VB.CommandButton cmCancel 
      Caption         =   "Выход"
      Height          =   315
      Left            =   10620
      TabIndex        =   2
      Top             =   2940
      Width           =   1035
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   4260
      TabIndex        =   4
      Text            =   "tbMobile"
      Top             =   660
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lbDate 
      Height          =   1968
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2415
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   10035
      _ExtentX        =   17695
      _ExtentY        =   4255
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      AllowBigSelection=   0   'False
      MergeCells      =   2
      AllowUserResizing=   1
   End
   Begin VB.Label laFirm 
      BackColor       =   &H80000005&
      Caption         =   "laFirm"
      Height          =   255
      Left            =   4260
      TabIndex        =   8
      Top             =   120
      Width           =   5235
   End
   Begin VB.Label Label2 
      Caption         =   "Фирма:"
      Height          =   195
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   675
   End
   Begin VB.Label laZakaz 
      BackColor       =   &H80000005&
      Caption         =   "laZakaz"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Заказ №:"
      Height          =   195
      Left            =   1620
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label laNomer 
      Alignment       =   2  'Center
      Caption         =   "Накладная /  Отгрузка   от:"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1275
   End
End
Attribute VB_Name = "Otgruz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public quantity2 As Long
Public mousRow2 As Long, mousCol2 As Long
Public Regim As String
Public closeZakaz As Boolean
Public isLoad As Boolean
Public orderRate As Double

'Const usSumm = 1
'Const usOutSum = 2
'Const usNowSum = 3
'Const prId = 0
'Const prType = 1
Const prNomNom = 1
Const prDescript = 2
Const prEdIzm = 3
Const prCenaEd = 4
Const prQuant = 5
Const prSumm = 6
Const prOutQuant = 7
Const prOutSum = 8
Const prNowQuant = 9
Const prNowSum = 10

Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
'Dim beEdit As Boolean, outDate() As Date, outtDate As Date
Dim outDate() As Date, outLen As Integer, deltaHeight As Integer

Private Sub cmCancel_Click()
    Unload Me
End Sub

Private Function adjustGirdMoneyColWidth(inStartup As Boolean) As Long
Dim J As Integer
Dim ret As Long
' Расширить колонки для рублей
    For J = 1 To Grid2.Cols - 1
        If Grid2.colWidth(J) > 0 _
            And (Grid2.TextMatrix(1, J) = "Цена" Or Grid2.TextMatrix(1, J) = "Сумма") _
        Then
            If sessionCurrency = CC_RUBLE Then
                Grid2.colWidth(J) = Grid2.colWidth(J) * ColWidthForRuble
            ElseIf Not inStartup Then
                Grid2.colWidth(J) = Grid2.colWidth(J) / ColWidthForRuble
            End If
        End If
        ret = ret + Grid2.colWidth(J)
    Next J
    adjustGirdMoneyColWidth = ret
End Function



Private Sub cmExel_Click()

End Sub


Private Sub cmHelp_Click()

End Sub

Private Sub cmDel_Click()
Dim I As Integer, J As Integer

strWhere = "WHERE (((outDate)='" & Format(lbDate.Text, "yyyy-mm-dd hh:nn:ss") & _
"' AND (numOrder)=" & gNzak & "));"
'MsgBox strWhere

wrkDefault.BeginTrans

  sql = "DELETE From bayNomenkOut " & strWhere
  J = myExecute("##208", sql, 0)
  If J > 0 Then GoTo ER1
  
  cErr = 219 '##219
'  If Not IsNumeric(saveShipped) Then GoTo ER1

wrkDefault.CommitTrans
loadOtgruz
Orders.Grid.TextMatrix(Orders.mousRow, orOtgrugeno) = rated(getShipped(gNzak), orderRate)

'Orders.openOrdersRowToGrid "##218": tqOrders.Close

lbDate.SetFocus
Exit Sub
ER1:
wrkDefault.Rollback
MsgBox "Удаление не прошло", , ""

End Sub

Private Sub Form_Load()
Dim str As String
isLoad = False
'laInform.Caption = "Проставте количества в колонке 'К сборке'. После " & _
"этого эти количества можно будет списать по накладной. И только после списания " & _
"соответствующая сумма появится в поле 'Отгружено'."
oldHeight = Me.Height
oldWidth = Me.Width
deltaHeight = Me.Height - lbDate.Height
gridIsLoad = False
noClick = True
Me.Caption = "Отгрузка изделий по заказу № " & gNzak
Grid2.Rows = 3
Grid2.FixedRows = 2
Grid2.MergeRow(0) = True
str = "|Надлежит отгрузить"
tmpStr = str & str
str = "|Уже отгружено"
tmpStr = tmpStr & str & str
str = "|Текущая отгрузка"
tmpStr = tmpStr & str & str
Grid2.FormatString = "|<|||" & tmpStr

'Grid2.TextMatrix(1, prType) = "Тип"
Grid2.TextMatrix(1, prNomNom) = "Номер"
Grid2.TextMatrix(1, prDescript) = "Описание"
Grid2.TextMatrix(1, prEdIzm) = "Ед.измерения"
Grid2.TextMatrix(1, prCenaEd) = "Цена за ед."
str = "Кол-во"
Grid2.TextMatrix(1, prQuant) = str
Grid2.TextMatrix(1, prOutQuant) = str
Grid2.TextMatrix(1, prNowQuant) = str
str = "Сумма"
Grid2.TextMatrix(1, prSumm) = str
Grid2.TextMatrix(1, prOutSum) = str
Grid2.TextMatrix(1, prNowSum) = str

Grid2.colWidth(0) = 0
'Grid2.ColWidth(prType) = 380
Grid2.colWidth(prNomNom) = 1185
Grid2.colWidth(prDescript) = 3300 '1270
Grid2.colWidth(prEdIzm) = 420
Grid2.colWidth(prCenaEd) = 630
Grid2.colWidth(prQuant) = 630
Grid2.colWidth(prSumm) = 1035
Grid2.colWidth(prOutQuant) = 630
Grid2.colWidth(prOutSum) = 800
Grid2.colWidth(prNowQuant) = 630
Grid2.colWidth(prNowSum) = 800

Me.Width = adjustGirdMoneyColWidth(True) + lbDate.left + lbDate.Width + 400

loadOtgruz
fitFormToGrid Me, Grid2
noClick = False

isLoad = True
End Sub


Sub loadPredmeti() '
Dim s As Single, sum As Single, sum2 As Single, str As String

MousePointer = flexHourglass
Grid2.Visible = False
quantity2 = 0
clearGrid Grid2, 2


sql = "SELECT  sDMCrez.quantity, sDMCrez.curQuant, sDMCrez.intQuant, " & _
"sGuideNomenk.nomNom, sGuideNomenk.cod, sGuideNomenk.Size, sGuideNomenk.nomName, " & _
"sGuideNomenk.ed_Izmer2, sGuideNomenk.perList " & _
"FROM sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom " & _
"Where (((sDMCrez.numDoc) = " & gNzak & ")) ORDER BY sDMCrez.nomNom;"

'MsgBox sql

Set tbNomenk = myOpenRecordSet("##210", sql, dbOpenForwardOnly)
'If tbNomenk Is Nothing Then GoTo EN1


If Not tbNomenk.BOF Then
  sum = 0: sum2 = 0
  While Not tbNomenk.EOF
    quantity2 = quantity2 + 1
    Grid2.TextMatrix(quantity2 + 1, prDescript) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid2.TextMatrix(quantity2 + 1, prNomNom) = tbNomenk!nomNom
    Grid2.TextMatrix(quantity2 + 1, prEdIzm) = tbNomenk!ed_Izmer2
    Grid2.TextMatrix(quantity2 + 1, prCenaEd) = Round(rated(tbNomenk!intQuant, orderRate), 2)
    
    s = tbNomenk!quantity / tbNomenk!perList
    Grid2.TextMatrix(quantity2 + 1, prQuant) = Round(s, 2)
    s = s * tbNomenk!intQuant      'итоговая сумма по позиции (сколько вообще заказано этой позиции в заказе)
    Grid2.TextMatrix(quantity2 + 1, prSumm) = Round(rated(s, orderRate), 2)
    sum = sum + s
        
    
    'отпущено до даты
    strWhere = "'" & Format(lbDate.Text, "yyyy-mm-dd hh:nn:ss") & "'"
    

'далее quant из bayNomenkOut (отгрузка) - в целых
    sql = "SELECT Sum(quant) AS Sum_quant From bayNomenkOut " & _
    "WHERE (((numOrder)=" & gNzak & ") AND ((nomNom)='" & _
    tbNomenk!nomNom & "') AND ((outDate)"
    str = sql & "<" & strWhere & "));"
    'MsgBox str
    byErrSqlGetValues "W##203", str, s

    ' заказано ранее (до текущей отгрузки)
    Grid2.TextMatrix(quantity2 + 1, prOutQuant) = Round(s, 2)
    Grid2.TextMatrix(quantity2 + 1, prOutSum) = Round(rated(tbNomenk!intQuant * s, orderRate), 2)
    sum2 = sum2 + tbNomenk!intQuant * s
    
    If lbDate.ListIndex = lbDate.ListCount - 1 Then ' последний
        ReDim Preserve QQ(quantity2 + 1): QQ(quantity2 + 1) = s
    End If
    
    'отпущено на дату
    str = sql & "=" & strWhere & "));"
    
    byErrSqlGetValues "W##204", str, s
    Grid2.TextMatrix(quantity2 + 1, prNowQuant) = s
    Grid2.TextMatrix(quantity2 + 1, prNowSum) = Round(rated(tbNomenk!intQuant * s, orderRate), 2)
'  End If
    Grid2.AddItem ""
    tbNomenk.MoveNext
  Wend
  'Grid2.RemoveItem quantity2 + 1
End If
tbNomenk.Close
EN1:
Grid2.Visible = True
If quantity2 > 0 Then
'    If quantity2 > 23 Then cmPrint.Visible = False
    Grid2.TextMatrix(quantity2 + 2, prQuant) = "Итого:"
    Grid2.row = quantity2 + 2: Grid2.col = prSumm
    Grid2.Text = Round(rated(sum, orderRate), 2)
    Grid2.CellFontBold = True
    Grid2.col = prOutSum
    Grid2.Text = Round(rated(sum2, orderRate), 2)
    Grid2.CellFontBold = True
    Grid2.col = prNowSum
    Grid2.CellFontBold = True
    
    nowSummToGrid
    
    Grid2.row = 2: Grid2.col = 1
    On Error Resume Next
    If Me.isLoad Then
        Grid2.SetFocus
    Else
        Grid2.TabIndex = 0
    End If
End If
MousePointer = flexDefault


End Sub

Sub loadOtgruz()
Dim I As Integer

lbDate.Clear
loadOutDates
For I = 0 To outLen
    lbDate.AddItem Format(outDate(I), "dd.mm.yy hh:nn:ss")
Next I
lbDate.ListIndex = outLen: gridIsLoad = False

ReDim QQ(0)

loadPredmeti
Grid2.col = prNowQuant
mousCol2 = prNowQuant
'Grid2.row = 2
'mousRow2 = 2


gridIsLoad = True
End Sub

'Function loadOutDates() As Boolean

'loadOutDates = False
'даты отгрузки
'sql = "SELECT xDate From sDocs WHERE (((numDoc)=" & gNzak & ")) " & _
'"GROUP BY sDocs.xDate;" ' т.к. до 18.06.04 по списанию - до 2х накладных
'Set table = myOpenRecordSet("##222", sql, dbOpenForwardOnly)
'If table Is Nothing Then Exit Function


'ReDim outDate(0): outLen = 0
'If Not table.BOF Then
' loadOutDates = True
' While Not table.EOF
'    outDate(outLen) = table!xDate
'    outLen = outLen + 1: ReDim Preserve outDate(outLen)
'    table.MoveNext
' Wend
'End If
'table.Close
'outDate(outLen) = Now()

'End Function

Function loadOutDates() As Boolean


loadOutDates = False
'даты отгрузки
    sql = "SELECT outDate From BayNomenkOut " & _
    "WHERE (((numOrder)=" & gNzak & ")) GROUP BY outDate ORDER BY 1;"
Set tbProduct = myOpenRecordSet("##222", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then myBase.Close: End

ReDim outDate(0): outLen = 0
If Not tbProduct.BOF Then
 loadOutDates = True
 While Not tbProduct.EOF
    outDate(outLen) = tbProduct!outDate
    outLen = outLen + 1: ReDim Preserve outDate(outLen)
    tbProduct.MoveNext
 Wend
End If
tbProduct.Close
outDate(outLen) = Now()

End Function


Sub nowSummToGrid()
Dim il As Long, sum As Single
sum = 0
For il = 2 To Grid2.Rows - 2
    sum = sum + Grid2.TextMatrix(il, prNowSum)
Next il
Grid2.TextMatrix(il, prNowSum) = Round(sum, 2)


End Sub

''обновляет поле shipped в Orders
'Function saveBayShipped() As Variant
'Dim s As Single, s1 As Single, str As String
'saveBayShipped = Null
''sql = "SELECT Sum([sDMC].[quant]*[sDMC].[intQuant]/[sGuideNomenk].[perList]) AS cSum " & _
'"FROM sGuideNomenk INNER JOIN sDMC ON sGuideNomenk.nomNom = sDMCrez.nomNom " & _
'"WHERE (((sDMCrez.numDoc)=" & gNzak & "));"
'sql = "SELECT Sum([sDMC].[quant]*[sDMCrez].[intQuant]/[sGuideNomenk].[perList]) AS Выражение1 " & _
'"FROM (sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom) INNER JOIN sDMC ON (sDMCrez.nomNom = sDMC.nomNom) AND (sDMCrez.numDoc = sDMC.numDoc) " & _
'"WHERE (((sDMCrez.numDoc)=" & gNzak & "));"
'If Not byErrSqlGetValues("W##209", sql, s) Then Exit Function
's = Round(s, 2): str = s
'If s = 0 Then str = "Null"
'sql = "UPDATE bayOrders SET shipped = " & str & " WHERE (((numOrder)=" & gNzak & "));"
'If myExecute("##215", sql) = 0 Then
'    saveBayShipped = s
'    If s = 0 Then str = ""
'    Orders.Grid.TextMatrix(Orders.Grid.row, orOtgrugeno) = str
'End If
'End Function

Private Sub Form_Resize()

Dim h As Integer, w As Integer
If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid2.Height = Grid2.Height + h
lbDate.Height = Me.Height - deltaHeight
Grid2.Width = Grid2.Width + w

cmCancel.Top = cmCancel.Top + h
cmCancel.left = cmCancel.left + w
cmDel.Top = cmDel.Top + h

'cmPrint.Top = cmPrint.Top + h
'cmExel.Top = cmExel.Top + h
'cmHelp.Top = cmHelp.Top + h

End Sub

Private Sub Grid2_Click()
mousCol2 = Grid2.MouseCol
mousRow2 = Grid2.MouseRow
If mousRow2 > 1 Then Grid2_EnterCell
End Sub

Private Sub Grid2_DblClick()
If mousRow2 = 0 Then Exit Sub
If Grid2.CellBackColor = &H88FF88 Then
    gNomNom = Grid2.TextMatrix(Grid2.row, prNomNom)
    textBoxInGridCell tbMobile, Grid2
End If

End Sub

Sub lbHide2()
tbMobile.Visible = False
Grid2.Enabled = True
Grid2.SetFocus
Grid2_EnterCell
End Sub

Private Sub Grid2_EnterCell()
Dim str As String, I As Integer

If quantity2 = 0 Or Not gridIsLoad Or closeZakaz Or noClick Then Exit Sub

If 2 > Grid2.row Or Grid2.row > quantity2 + 1 Then Exit Sub

mousRow2 = Grid2.row
mousCol2 = Grid2.col
str = Grid2.TextMatrix(mousRow2, prNomNom) '
I = InStr(str, "/")
prExt = 0: If I > 1 Then prExt = left$(str, I - 1)   'номер поставки

gNomNom = Grid2.TextMatrix(mousRow2, prNomNom)

'If (Regim = "" And mousCol2 = prNowQuant) Then
If lbDate.ListIndex = lbDate.ListCount - 1 And mousCol2 = prNowQuant Then ' последний
    Grid2.CellBackColor = &H88FF88
Else
    Grid2.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid2_GotFocus()
Grid2_EnterCell
End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid2_DblClick
End Sub

Private Sub Grid2_LeaveCell()
If Grid2.row > 1 Then Grid2.CellBackColor = Grid2.BackColor
End Sub

Private Sub Grid2_LostFocus()
Grid2.CellBackColor = Grid2.BackColor
End Sub

Private Sub Grid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid2.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid2.colWidth(Grid2.MouseCol)

End Sub


Private Sub lbDate_Click()
Dim str As String

If noClick Then Exit Sub
'If lbDate.Index < outLen Then

cmDel.Enabled = ((lbDate.ListIndex < outLen) And Not closeZakaz)

gridIsLoad = False
'If lbDate.ListIndex = lbDate.ListCount - 1 Then ' последний
'    str = "К сборке"
'Else
    str = "Текущ.отгрузка"
'End If

Grid2.TextMatrix(0, prNowQuant) = str
Grid2.TextMatrix(0, prNowSum) = str



loadPredmeti
'nowSummToGrid
Grid2.row = 2
Grid2.col = prNowQuant

'gridIsLoad = False
mousCol2 = Grid2.col
mousRow2 = Grid2.row
gridIsLoad = True
End Sub

Private Sub tbMobile_DblClick()
lbHide2
End Sub

Private Sub tbMobileOld_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim pQuant  As Single, maxQ As Single

'If KeyCode = vbKeyReturn Then
    
  
'  maxQ = Grid2.TextMatrix(mousRow2, prQuant) 'надлежит отгрузить
'  maxQ = maxQ - Grid2.TextMatrix(mousRow2, prOutQuant) 'на столько можно увеличивать
'  If Not isNumericTbox(tbMobile, 0, maxQ) Then Exit Sub
'  pQuant = Round(tbMobile.Text, 0)
'  sql = "UPDATE sDMCrez SET curQuant = " & pQuant & _
'  " WHERE (((numDoc)=" & gNzak & ") AND ((nomNom)='" & gNomNom & "'));"
'
'  If myExecute("##201", sql) <> 0 Then GoTo ER1
  
  
  ''корректируем Уже отпущено на сегодня
'  Grid2.TextMatrix(mousRow2, prNowQuant) = pQuant
'  maxQ = Grid2.TextMatrix(mousRow2, prCenaEd)
'  Grid2.TextMatrix(mousRow2, prNowSum) = Round(pQuant * maxQ, 2)
  '
'  nowSummToGrid
'
'EN1: lbHide2
'ElseIf KeyCode = vbKeyEscape Then
'  lbHide2
'End If


Exit Sub
ER1:
wrkDefault.Rollback
ER2:
lbHide2
MsgBox "Отгрузка не прошла", , "Error-" & cErr & " Сообщите администратору"
End Sub

Private Function checkQuantInput() As Boolean
Dim pQuant As Long
    checkQuantInput = True
    
    If Not checkNumeric(tbMobile.Text) Then
        GoTo finally
    End If
    pQuant = Round(tbMobile.Text)
    If CStr(pQuant) <> tbMobile.Text Then
        MsgBox "Может быть только целым значением", , "Error"
        GoTo finally
    End If
    Exit Function
finally:
    checkQuantInput = False
    tbMobile.SetFocus
    tbMobile.SelStart = 1
    tbMobile.SelLength = Len(tbMobile.Text)
End Function

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim pQuant As Single, s As Single, maxQ As Single

If KeyCode = vbKeyReturn Then
  
  wrkDefault.BeginTrans
  
  maxQ = Grid2.TextMatrix(mousRow2, prQuant) 'надлежит отгрузить
  If Not checkQuantInput() Then
    wrkDefault.Rollback: Exit Sub
  End If
  maxQ = maxQ - QQ(mousRow2) 'на столько можно увеличивать
  If Not isNumericTbox(tbMobile, 0, maxQ) Then wrkDefault.Rollback: Exit Sub
    pQuant = tbMobile.Text

    On Error GoTo ER1
    sql = "SELECT * from bayNomenkOut " & _
    "WHERE (((outDate)='" & Format(outDate(lbDate.ListIndex), "yyyy-mm-dd hh:nn:ss") & _
    "') AND ((numOrder)=" & gNzak & ") AND ((nomNom)='" & gNomNom & "'));"
        
    Set tbNomenk = myOpenRecordSet("##201", sql, dbOpenForwardOnly)
    If tbNomenk.BOF Then
        If pQuant > 0 Then
            tbNomenk.AddNew
            tbNomenk!outDate = outDate(lbDate.ListIndex)
            tbNomenk!numorder = gNzak
            tbNomenk!nomNom = gNomNom
            tbNomenk!quant = pQuant
            tbNomenk.Update
        End If
    ElseIf pQuant = 0 Then
        tbNomenk.Delete
    Else
        tbNomenk.Edit
        tbNomenk!quant = pQuant
        tbNomenk.Update
    End If
    tbNomenk.Close
    
'  End If

  cErr = 216 '##216
  'If Not IsNumeric(saveShipped) Then GoTo ER1
  'Orders.openOrdersRowToGrid "##217":  tqOrders.Close
  wrkDefault.CommitTrans
  
  Orders.Grid.TextMatrix(Orders.mousRow, orOtgrugeno) = Round(rated(getShipped(gNzak), orderRate), 2)
  
  'корректируем Уже отпущено на сегодня
'  QQ(mousRow2) = QQ(mousRow2) + pQuant - Grid2.TextMatrix(mousRow2, prNowQuant)
  
  Grid2.TextMatrix(mousRow2, prNowQuant) = pQuant
  maxQ = Grid2.TextMatrix(mousRow2, prCenaEd)
  Grid2.TextMatrix(mousRow2, prNowSum) = Round(pQuant * maxQ, 2)
  
  OutNowSummToGrid2
  
EN1: lbHide2
ElseIf KeyCode = vbKeyEscape Then
  lbHide2
End If


Exit Sub
ER1:
wrkDefault.Rollback
errorCodAndMsg 201
lbHide2
'MsgBox "Отгрузка не прошла", , "Error-" & cErr & " Сообщите администратору"
End Sub


Sub OutNowSummToGrid2()
Dim il As Long, sum As Single, sum2 As Single
sum = 0: sum2 = 0
For il = 2 To Grid2.Rows - 2
    sum = sum + Grid2.TextMatrix(il, prOutSum)
    sum2 = sum2 + Grid2.TextMatrix(il, prNowSum)
Next il
Grid2.TextMatrix(il, prOutSum) = Round(sum, 2)
Grid2.TextMatrix(il, prNowSum) = Round(sum2, 2)


End Sub


