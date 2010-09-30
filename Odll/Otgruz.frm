VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Otgruz 
   BackColor       =   &H8000000A&
   Caption         =   "Отгрузка"
   ClientHeight    =   3024
   ClientLeft      =   168
   ClientTop       =   456
   ClientWidth     =   10128
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3024
   ScaleWidth      =   10128
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmOtgruzDate 
      Caption         =   "Сменить дату"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   300
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmCancel 
      Cancel          =   -1  'True
      Caption         =   "Выход"
      Height          =   315
      Left            =   8880
      TabIndex        =   3
      Top             =   2640
      Width           =   1035
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   4260
      TabIndex        =   5
      Text            =   "tbMobile"
      Top             =   660
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lbDate 
      Height          =   2160
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid5 
      Height          =   2535
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   8595
      _ExtentX        =   15155
      _ExtentY        =   4466
      _Version        =   393216
      AllowBigSelection=   0   'False
      MergeCells      =   2
      AllowUserResizing=   1
   End
   Begin VB.Label laNomer 
      Alignment       =   2  'Center
      Caption         =   "Даты отгрузки:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   1335
   End
   Begin VB.Menu mnOtruzDate 
      Caption         =   "Сменить дату отгрузки"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Otgruz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public quantity5 As Long
Public mousRow5 As Long, mousCol5 As Long
Public Regim As String
Public closeZakaz As Boolean
Public orderRate As Double
Public idWerk As Integer
Dim myAsWhole As Integer

Dim doOtgruzDateUpdate As Boolean
Dim uslugOrdered As Double

Const usSumm = 1
Const usOutSum = 2
Const usNowSum = 3

Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim Outdate() As Date, outLen As Integer, deltaHeight As Integer

Private Sub cmCancel_Click()
    Unload Me
End Sub

Private Sub cmDel_Click()
Dim I As Integer, J As Integer

strWhere = "WHERE outDate ='" & Format(lbDate.Text, "yyyy-mm-dd hh:nn:ss") & _
"' AND numOrder = " & gNzak
'MsgBox strWhere

wrkDefault.BeginTrans

If Regim = "uslug" Then
  sql = "DELETE From xUslugOut " & strWhere
'  MsgBox sql
  I = myExecute("##207", sql)
  If I <> 0 Then GoTo ER1
Else
  sql = "DELETE From xPredmetyByIzdeliaOut " & strWhere
'  MsgBox sql

  I = myExecute("##207", sql, 0) 'дает -1, если нет таких записей - а это возможно
  If I > 0 Then GoTo ER1

  sql = "DELETE From xPredmetyByNomenkOut " & strWhere
  J = myExecute("##208", sql, 0)
  If J > 0 Then GoTo ER1
End If
  cErr = 219 '##219
  If Not IsNumeric(saveShipped) Then GoTo ER1

wrkDefault.CommitTrans
loadOtgruz
Orders.openOrdersRowToGrid "##218": tqOrders.Close

lbDate.SetFocus
Exit Sub
ER1:
wrkDefault.Rollback
MsgBox "Удаление не прошло", , ""
End Sub

Private Sub updateOtgruzDate()
Dim strSet As String, I As Integer
    tbMobile.Visible = False
    If IsDate(tbMobile.Text) Then
        Dim newDate As Date, oldDate As Date
        Dim newDateStr As String
        newDate = CDate(tbMobile.Text)
        oldDate = CDate(lbDate.Text)
        lbDate.List(lbDate.ListIndex) = Format(newDate, "dd.mm.yy hh:mm:ss")
        newDateStr = "'" & Format(newDate, "yyyymmdd hh:mm:ss") & "'"
        strSet = " set outDate = "
        strWhere = " WHERE outDate ='" & Format(Format(oldDate, "yyyymmdd hh:mm:ss")) & "' AND numOrder = " & gNzak
        
wrkDefault.BeginTrans
        If Regim = "uslug" Then
            sql = sql & "update xUslugOut"
            I = myExecute("##0.209.1", sql & strSet & newDateStr & strWhere, 0)
            If I > 0 Then GoTo ER1
        Else
            sql = "update xPredmetyByIzdeliaOut"
            
            'Debug.Print sql & strSet & newDateStr & strWhere
            I = myExecute("##0.209.2", sql & strSet & newDateStr & strWhere, 0) 'дает -1, если нет таких записей - а это возможно
            If I > 0 Then GoTo ER1
            
            sql = "update xPredmetyByNomenkOut "
            I = myExecute("##0.209.3", sql & strSet & newDateStr & strWhere, 0)
            If I > 0 Then GoTo ER1
        End If
    Else
        MsgBox "ВВедена некорректная дата. Попробуйте еще раз", , "Ошибка ввода"
    End If
    wrkDefault.CommitTrans

    lbDate.SetFocus
    Exit Sub
ER1:
    wrkDefault.Rollback
    MsgBox "Непредвиденная ошибка при смене даты отгрузки", , "Сообщите администартору"
    
End Sub

Private Sub cmOtgruzDate_Click()
Dim strDate As String

    doOtgruzDateUpdate = tbMobile.Visible = True And tbMobile.Top = cmOtgruzDate.Top
    If doOtgruzDateUpdate Then
        updateOtgruzDate
        Exit Sub
    End If

    strDate = Format(lbDate.Text, "dd.mm.yy")
    If tbMobile.Visible = False Then
        tbMobile.Left = cmOtgruzDate.Left + cmOtgruzDate.Width + 100
        tbMobile.Top = cmOtgruzDate.Top
        tbMobile.Visible = True
        tbMobile.Text = strDate
        tbMobile.Width = 800
        tbMobile.SetFocus
        tbMobile.SelStart = 0
        tbMobile.SelLength = 10
    End If
    
    
End Sub

Private Sub Form_Load()
Dim str As String

oldHeight = Me.Height
oldWidth = Me.Width
deltaHeight = Me.Height - lbDate.Height

' продажи или пр-во?
If idWerk = 1 Then
    myAsWhole = 1
Else
    myAsWhole = 0
End If

gridIsLoad = False
noClick = True
If dostup = "a" Then
    cmOtgruzDate.Visible = True
End If
If Regim = "uslug" Then
    Me.Caption = "Отгрузка  по заказу № " & gNzak
    Me.Width = Me.Width - 4200
    Me.Height = Me.Height - 195 * 5
    Grid5.FormatString = "|Надлежит отгрузить|Уже отгружено|Текущ.отгрузка"
    Grid5.ColWidth(0) = 0
    Grid5.ColWidth(usSumm) = 1650
    Grid5.ColWidth(usOutSum) = 1300
    Grid5.ColWidth(usNowSum) = 1300
    loadOtgruz
    quantity5 = 1 ' для Grid5_EnterCell
    noClick = False
    Exit Sub
End If
Me.Caption = "Отгрузка изделий по заказу № " & gNzak
Grid5.Rows = 3
Grid5.FixedRows = 2
Grid5.MergeRow(0) = True
str = "|Надлежит отгрузить"
tmpStr = str & str
str = "|Уже отгружено"
tmpStr = tmpStr & str & str
str = "|Текущ.отгрузка"
tmpStr = tmpStr & str & str
Grid5.FormatString = "||<|||" & tmpStr

Grid5.TextMatrix(1, prType) = "Тип"
Grid5.TextMatrix(1, prName) = "Номер"
Grid5.TextMatrix(1, prDescript) = "Описание"
Grid5.TextMatrix(1, prEdizm) = "Ед.измерения"
Grid5.TextMatrix(1, prCenaEd) = "Цена за ед."
str = "Кол-во"
Grid5.TextMatrix(1, prQuant) = str
Grid5.TextMatrix(1, prOutQuant) = str
Grid5.TextMatrix(1, prNowQuant) = str
str = "Сумма"
Grid5.TextMatrix(1, prSumm) = str
Grid5.TextMatrix(1, prOutSum) = str
Grid5.TextMatrix(1, prNowSum) = str

Grid5.ColWidth(prId) = 0
Grid5.ColWidth(prType) = 0 '380
Grid5.ColWidth(prName) = 1185
Grid5.ColWidth(prDescript) = 1270 + 380
Grid5.ColWidth(prEdizm) = 420
Grid5.ColWidth(prCenaEd) = 495
Grid5.ColWidth(prQuant) = 630
Grid5.ColWidth(prSumm) = 1035
Grid5.ColWidth(prOutQuant) = 630
Grid5.ColWidth(prOutSum) = 800
Grid5.ColWidth(prNowQuant) = 630
Grid5.ColWidth(prNowSum) = 800


loadOtgruz
noClick = False

End Sub

Sub loadUslug()
Dim S As Double

sql = "SELECT ordered From Orders WHERE (((Orders.numOrder)=" & gNzak & "));"
If byErrSqlGetValues("##227", sql, S) Then _
    Grid5.TextMatrix(1, usSumm) = Round(rated(S, orderRate), 2)
uslugOrdered = S

getOtgrugeno 1, 1 ' usNowSum и usOutSum

End Sub

Sub loadOtgruz()
Dim I As Integer

lbDate.Clear
loadOutDates
For I = 0 To outLen
    lbDate.AddItem Format(Outdate(I), "dd.mm.yy hh:nn:ss")
Next I
lbDate.ListIndex = outLen: gridIsLoad = False

ReDim QQ(0)
ReDim QQ2(0)
ReDim QQ3(0) 'будут храниться цена за Единицу (cenaEd)

If Regim = "uslug" Then
    loadUslug
    Grid5.col = usNowSum
    mousCol5 = usNowSum
Else
    loadPredmeti Me, orderRate, idWerk, myAsWhole, "fromOtgruz", 0
    Grid5.col = prNowQuant
    mousCol5 = prNowQuant
    Grid5.row = 2
    mousRow5 = 2
End If


gridIsLoad = True
End Sub

Function loadOutDates() As Boolean


loadOutDates = False
'даты отгрузки
If Regim = "uslug" Then
    sql = "SELECT outDate From xUslugOut WHERE (((numOrder)=" & gNzak & _
    ")) GROUP BY outDate ORDER BY outDate;"
Else
    sql = "SELECT outDate From xPredmetyByIzdeliaOut " & _
    "WHERE (((numOrder)=" & gNzak & ")) GROUP BY outDate " & _
    "UNION SELECT outDate From xPredmetyByNomenkOut " & _
    "WHERE (((numOrder)=" & gNzak & ")) GROUP BY outDate ORDER BY 1;"
End If
Set tbProduct = myOpenRecordSet("##222", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then myBase.Close: End

ReDim Outdate(0): outLen = 0
If Not tbProduct.BOF Then
 loadOutDates = True
 While Not tbProduct.EOF
    Outdate(outLen) = tbProduct!Outdate
    outLen = outLen + 1: ReDim Preserve Outdate(outLen)
    tbProduct.MoveNext
 Wend
End If
tbProduct.Close
Outdate(outLen) = Now()

End Function

Sub getOtgrugeno(row As Long, myCenaEd As Double, Optional byNomenk As String = "")
Dim S As Double, str  As String
strWhere = "'" & Format(lbDate.Text, "yyyy-mm-dd hh:nn:ss") & "'"

Dim Nomnom1 As Nomnom
Dim isNomnom As Boolean
isNomnom = False


'отпущено до даты
If Regim = "uslug" Then
    sql = "SELECT Sum(quant) AS Sum_quant From xUslugOut " & _
    " WHERE numOrder = " & gNzak & " AND outDate < " & strWhere
ElseIf byNomenk = "" Then
    sql = "SELECT Sum(quant) AS Sum_quant From xPredmetyByIzdeliaOut " & _
    "WHERE numOrder = " & gNzak & " AND prId = " & tbNomenk!prId & _
    " AND prExt = " & tbNomenk!prExt & " AND outDate < " & strWhere
Else
    Set Nomnom1 = nomnomCache.getNomnom(tbNomenk!Nomnom)
    isNomnom = True
    sql = "SELECT Sum(quant) AS Sum_quant From xPredmetyByNomenkOut " & _
    "WHERE numOrder = " & gNzak & " AND nomNom = '" & tbNomenk!Nomnom & _
    "' AND outDate <" & strWhere
End If
'MsgBox sql
byErrSqlGetValues "W##203", sql, S
Dim itemSumma As Double

itemSumma = tbNomenk!cenaEd * S
If myAsWhole = 1 And isNomnom Then
    S = Nomnom1.getQuantity(S, myAsWhole)
End If

If Regim = "uslug" Then
    Grid5.TextMatrix(row, usOutSum) = Round(rated(S, orderRate), 2)
Else
    Grid5.TextMatrix(row, prOutQuant) = Round(S, 2)
    If IsNumeric(tbNomenk!cenaEd) Then
        Grid5.TextMatrix(row, prOutSum) = Round(rated(itemSumma, orderRate), 2)
    End If
End If

ReDim Preserve QQ(row): QQ(row) = S

'отпущено на дату
If Regim = "uslug" Then
    sql = "SELECT Sum(quant) AS Sum_quant From xUslugOut " & _
    "WHERE numOrder = " & gNzak & " AND outDate = " & strWhere
ElseIf byNomenk = "" Then
    sql = "SELECT quant From xPredmetyByIzdeliaOut " & _
    "WHERE numOrder = " & gNzak & " AND prId = " & tbNomenk!prId & _
    " AND prExt = " & tbNomenk!prExt & " AND outDate = " & strWhere
Else
    sql = "SELECT quant From xPredmetyByNomenkOut "
    sql = sql & " WHERE numOrder = " & gNzak & " AND nomNom = '" & tbNomenk!Nomnom & "' AND outDate = " & strWhere
End If
'MsgBox sql
byErrSqlGetValues "W##204", sql, S

itemSumma = tbNomenk!cenaEd * S
If myAsWhole = 1 And isNomnom Then
    S = Nomnom1.getQuantity(S, myAsWhole)
End If
ReDim Preserve QQ2(row): QQ2(row) = S

ReDim Preserve QQ3(row): QQ3(row) = myCenaEd

If Regim = "uslug" Then
    Grid5.TextMatrix(row, usNowSum) = Round(rated(S, orderRate), 2)
Else
    Grid5.TextMatrix(row, prNowQuant) = Round(S, 2)
    If IsNumeric(tbNomenk!cenaEd) Then
        Grid5.TextMatrix(row, prNowSum) = Round(rated(itemSumma, orderRate), 2)
    End If
End If
End Sub

'обновляет поле shipped в Orders
Function saveShipped(Optional doUpdate As Boolean = True) As Variant
Dim S As Double, s1 As Double

saveShipped = Null
If Regim = "" Then
    sql = "SELECT Sum(pi.cenaEd * pio.quant)" & _
    " FROM xPredmetyByIzdelia pi " & _
    " JOIN xPredmetyByIzdeliaOut pio ON (pi.prExt = pio.prExt)" & _
        " AND pi.prId = pio.prId " & _
        " AND pi.numOrder = pio.numOrder "
    sql = sql & " WHERE pi.numOrder =" & gNzak
    If Not byErrSqlGetValues("W##213", sql, S) Then Exit Function

    sql = "SELECT Sum(pn.cenaEd * pno.quant) " & _
    " FROM xPredmetyByNomenk pn" & _
    " JOIN xPredmetyByNomenkOut pno ON pn.nomNom = pno.nomNom" & _
        " AND pn.numOrder = pno.numOrder "
    If myAsWhole = 1 Then
        sql = sql & " JOIN sGuideNomenk n ON n.nomnom = pn.nomnom"
    End If
    sql = sql & " WHERE pn.numOrder =" & gNzak
    If Not byErrSqlGetValues("W##214", sql, s1) Then Exit Function

    S = S + s1
Else 'услуги
    sql = "SELECT Sum(quant) AS Sum_quant From xUslugOut " & _
    "WHERE numOrder = " & gNzak
    If Not byErrSqlGetValues("W##301", sql, S) Then Exit Function
End If
If S > 0 Then
    tmpStr = S
Else
    tmpStr = "Null"
End If

If doUpdate Then
    orderUpdate "##368", tmpStr, "Orders", "shipped"
End If
saveShipped = S
End Function

Private Sub Form_Resize()

Dim H As Integer, W As Integer
If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next
H = Me.Height - oldHeight
oldHeight = Me.Height
W = Me.Width - oldWidth
oldWidth = Me.Width
Grid5.Height = Grid5.Height + H
lbDate.Height = Me.Height - deltaHeight
Grid5.Width = Grid5.Width + W

cmDel.Top = cmDel.Top + H
cmCancel.Top = cmCancel.Top + H
cmCancel.Left = cmCancel.Left + W
cmOtgruzDate.Top = cmDel.Top
cmOtgruzDate.Left = cmDel.Left + cmDel.Width + 100

End Sub

Private Sub Grid5_DblClick()
If mousRow5 = 0 Then Exit Sub
If Grid5.CellBackColor = &H88FF88 Then
    textBoxInGridCell tbMobile, Grid5
End If

End Sub

Sub lbHide5()
tbMobile.Visible = False
Grid5.Enabled = True
Grid5.SetFocus
Grid5_EnterCell
End Sub

Private Sub Grid5_EnterCell()
Dim str As String, I As Integer

If quantity5 = 0 Or Not gridIsLoad Or closeZakaz Then Exit Sub

If Grid5.row > quantity5 + 1 Then Exit Sub

mousRow5 = Grid5.row
mousCol5 = Grid5.col

getIdFromGrid5Row Me 'gProductId gNomNom

If (Regim = "" And mousCol5 = prNowQuant) Or _
(Regim = "uslug" And mousCol5 = usNowSum) Then
    Grid5.CellBackColor = &H88FF88
Else
    Grid5.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid5_GotFocus()
Grid5_EnterCell
End Sub

Private Sub Grid5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid5_DblClick
End Sub

Private Sub Grid5_LeaveCell()
Grid5.CellBackColor = Grid5.BackColor
End Sub

Private Sub Grid5_LostFocus()
Grid5.CellBackColor = Grid5.BackColor
End Sub

Private Sub Grid5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Grid5.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid5.ColWidth(Grid5.MouseCol)

End Sub

Sub OutNowSummToGrid5()
Dim IL As Long, sum As Double, sum2 As Double
sum = 0: sum2 = 0
For IL = 2 To Grid5.Rows - 2
    sum = sum + Grid5.TextMatrix(IL, prOutSum)
    sum2 = sum2 + Grid5.TextMatrix(IL, prNowSum)
Next IL
Grid5.TextMatrix(IL, prOutSum) = Round(sum, 2)
Grid5.TextMatrix(IL, prNowSum) = Round(sum2, 2)


End Sub

Private Sub lbDate_Click()
If noClick Then Exit Sub
cmDel.Enabled = ((lbDate.ListIndex < outLen) And Not closeZakaz)
cmOtgruzDate.Enabled = lbDate.ListIndex < outLen

gridIsLoad = False
If Regim = "" Then
    loadPredmeti Me, orderRate, idWerk, myAsWhole, "fromOtgruz"
    OutNowSummToGrid5
    Grid5.row = 2
    Grid5.col = prNowQuant
Else
    loadUslug
    Grid5.row = 1
    Grid5.col = usOutSum
End If

'gridIsLoad = False
mousCol5 = Grid5.col
mousRow5 = Grid5.row
gridIsLoad = True
End Sub

Private Sub tbMobile_DblClick()
lbHide5
End Sub
'$odbc15$

' total - общая сумма по позиции (сколько максимально можно отпустить) - только в долларах
' before - уже отпущено - доллары
' nowBack и nowFore - на вход - значения от в введенной валюте, на выход - в долларах
' для nowFore - определяет, если позиция близка к закрытию - корректирует валютное значение с учетом квантования
Function adjustOtgruz(ByVal total As Double, ByVal before As Double, ByRef nowBack As Double, nowFore As Double) As Boolean
    
    adjustOtgruz = True
 
    Dim nowForeStr As String
    
    nowForeStr = tbMobile.Text
    If Not IsNumeric(nowForeStr) Then
        adjustOtgruz = False
        GoTo finally
    Else
        nowFore = CDbl(nowForeStr)
    End If
    
    Dim overflow As Double, ceilValue As Double
    If sessionCurrency = CC_RUBLE Then
        Dim maxInputRub As Double
        
        nowBack = nowBack / orderRate
        nowFore = nowFore / orderRate
        
        ceilValue = (total - before) * orderRate
        overflow = (total - (before + nowFore)) * orderRate
    Else
        overflow = (total - (before - nowBack + nowFore))
    
        
    End If
    
    If overflow < -0.01 Or nowFore < 0 Then
        MsgBox "значение должно быть в диапазоне от 0 " _
        & " до " & CStr(Round(ceilValue, 2)), , "Error"
        adjustOtgruz = False
        GoTo finally
    End If
    
finally:
    If Not adjustOtgruz Then
        tbMobile.SetFocus
        tbMobile.SelStart = 0
        tbMobile.SelLength = Len(tbMobile.Text)
    End If
    
    
End Function

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim pQuant As Double, S As Double, maxQ As Double, preciseCenaEd As Double
Dim Nomnom1 As Nomnom


If KeyCode = vbKeyReturn Then
    doOtgruzDateUpdate = tbMobile.Visible = True And tbMobile.Top = cmOtgruzDate.Top
    preciseCenaEd = QQ3(mousRow5)
    If doOtgruzDateUpdate Then
        updateOtgruzDate
        Exit Sub
    End If
  If Regim = "uslug" Then
    Dim nowBack As Double, nowFore As Double
    
    nowBack = Grid5.TextMatrix(mousRow5, usNowSum)
    If Not adjustOtgruz(uslugOrdered, QQ(1), nowBack, nowFore) Then
        ' ввели не число
        Exit Sub
    End If
    
    'maxQ = uslugOrdered
    'maxQ = maxQ - QQ(1) 'на столько можно увеличивать
    'maxQ = maxQ + currentOtgruz
    'If Not isNumericTbox(tbMobile, 0, rated(maxQ, orderRate)) Then Exit Sub
    'pquant = tuneCurencyAndGranularity(,,,)
    'pQuant = Round(tbMobile.Text, 2)
    'tbMobile.Text = rated(nowFore, orderRate)
    
    sql = "SELECT * from xUslugOut WHERE numOrder = " & gNzak & _
    " AND outDate = '" & Format(Outdate(lbDate.ListIndex), "yyyy-mm-dd  hh:nn:ss") & "'"
    Set tbProduct = myOpenRecordSet("##229", sql, dbOpenForwardOnly)
    
    If tbProduct.BOF Then
        If nowFore > 0 Then
            tbProduct.AddNew
            tbProduct!Outdate = Outdate(lbDate.ListIndex)
            tbProduct!Numorder = gNzak
            tbProduct!quant = nowFore
            tbProduct.update
        End If
    ElseIf nowFore = 0 Then
        tbProduct.Delete
    Else
        tbProduct.Edit
        tbProduct!quant = nowFore
        tbProduct.update
    End If
    tbProduct.Close
    
    cErr = 302 '##302
    If Not IsNumeric(saveShipped) Then GoTo ER2
    Orders.openOrdersRowToGrid "##228":  tqOrders.Close

  'корректируем Уже отпущено на сегодня
    'QQ(1) = QQ(1) + nowFore - Grid5.TextMatrix(1, usNowSum)
    Grid5.TextMatrix(1, usNowSum) = rated(nowFore, orderRate)
    lbHide5
    Exit Sub
  End If
  
  
    On Error GoTo ER1
  wrkDefault.BeginTrans
  
  
  maxQ = Grid5.TextMatrix(mousRow5, prQuant) 'надлежит отгрузить
  maxQ = maxQ - QQ(mousRow5) 'на столько можно увеличивать
  If Not isNumericTbox(tbMobile, 0, maxQ) Then wrkDefault.Rollback: Exit Sub
  If Grid5.TextMatrix(mousRow5, prType) = "изделие" Then
    pQuant = Round(tbMobile.Text)
    tbMobile.Text = pQuant
  
    cErr = "202" '##202
    'изд-е
    sql = "SELECT pio.* FROM xPredmetyByIzdeliaOut pio " & _
    " WHERE  pio.outDate = '" & Format(Outdate(lbDate.ListIndex), "yyyy-mm-dd  hh:nn:ss") & "'" & _
    " AND pio.numOrder = " & gNzak & " AND pio.prId = " & gProductId & " AND pio.prExt = " & prExt
    'Debug.Print sql
    Set tbProduct = myOpenRecordSet("##200", sql, dbOpenForwardOnly)
    
    If tbProduct.BOF Then
        If pQuant > 0 Then
            tbProduct.AddNew
            tbProduct!Outdate = Outdate(lbDate.ListIndex)
            tbProduct!Numorder = gNzak
            tbProduct!prId = gProductId
            tbProduct!prExt = prExt
            tbProduct!quant = pQuant
            tbProduct.update
        End If
    ElseIf pQuant = 0 Then
        tbProduct.Delete
    Else
        tbProduct.Edit
        tbProduct!quant = pQuant
        tbProduct.update
    End If
    tbProduct.Close
  Else 'отдельная ном-ра
    pQuant = Round(tbMobile.Text, 2)
    tbMobile.Text = pQuant

    Set Nomnom1 = nomnomCache.getNomnom(gNomNom)
    
    sql = "SELECT pno.* from xPredmetyByNomenkOut pno " & _
    " WHERE outDate = '" & Format(Outdate(lbDate.ListIndex), "yyyy-mm-dd  hh:nn:ss") & _
    "' AND numOrder = " & gNzak & "  AND nomNom = '" & gNomNom & "'"
    
    Set tbNomenk = myOpenRecordSet("##201", sql, dbOpenForwardOnly)
    If tbNomenk.BOF Then
        If pQuant > 0 Then
            tbNomenk.AddNew
            tbNomenk!Outdate = Outdate(lbDate.ListIndex)
            tbNomenk!Numorder = gNzak
            tbNomenk!Nomnom = gNomNom
            tbNomenk!quant = Nomnom1.getQuantityRevert(pQuant, myAsWhole)
            tbNomenk.update
        End If
    ElseIf pQuant = 0 Then
        tbNomenk.Delete
    Else
        tbNomenk.Edit
        tbNomenk!quant = Nomnom1.getQuantityRevert(pQuant, myAsWhole)
        tbNomenk.update
    End If
    tbNomenk.Close
    
    Grid5.TextMatrix(mousRow5, prNowSum) = Nomnom1.getQuantityRevert(pQuant * preciseCenaEd, myAsWhole)
    
  End If

  cErr = 216 '##216
  If Not IsNumeric(saveShipped) Then GoTo ER1
  
  Orders.openOrdersRowToGrid "##217":  tqOrders.Close

  wrkDefault.CommitTrans
  'корректируем Уже отпущено на сегодня
  QQ(mousRow5) = QQ(mousRow5) + pQuant - Grid5.TextMatrix(mousRow5, prNowQuant)
  
  Grid5.TextMatrix(mousRow5, prNowQuant) = pQuant
  
  OutNowSummToGrid5
  
EN1: lbHide5
ElseIf KeyCode = vbKeyEscape Then
  lbHide5
End If


Exit Sub
ER1:
errorCodAndMsg ("Отгрузка не прошла")
wrkDefault.Rollback
ER2:
lbHide5
MsgBox "Отгрузка не прошла", , "Error-" & cErr & " Сообщите администратору"
End Sub
