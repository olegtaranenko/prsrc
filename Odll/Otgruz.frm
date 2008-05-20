VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Otgruz 
   BackColor       =   &H8000000A&
   Caption         =   "��������"
   ClientHeight    =   3030
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmDel 
      Caption         =   "�������"
      Enabled         =   0   'False
      Height          =   315
      Left            =   300
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmCancel 
      Cancel          =   -1  'True
      Caption         =   "�����"
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
      Height          =   2205
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
      _ExtentX        =   15161
      _ExtentY        =   4471
      _Version        =   393216
      AllowBigSelection=   0   'False
      MergeCells      =   2
      AllowUserResizing=   1
   End
   Begin VB.Label laNomer 
      Alignment       =   2  '���������
      Caption         =   "���� ��������:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   1335
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

Const usSumm = 1
Const usOutSum = 2
Const usNowSum = 3

Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����
Dim outDate() As Date, outLen As Integer, deltaHeight As Integer

Private Sub cmCancel_Click()

Unload Me
End Sub

Private Sub cmDel_Click()
Dim I As Integer, j As Integer

strWhere = "WHERE (((outDate)='" & Format(lbDate.Text, "yyyy-mm-dd hh:nn:ss") & _
"' AND (numOrder)=" & gNzak & "));"
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

  I = myExecute("##207", sql, 0) '���� -1, ���� ��� ����� ������� - � ��� ��������
  If I > 0 Then GoTo ER1

  sql = "DELETE From xPredmetyByNomenkOut " & strWhere
  j = myExecute("##208", sql, 0)
  If j > 0 Then GoTo ER1
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
MsgBox "�������� �� ������", , ""
End Sub

Private Sub Form_Load()
Dim str As String

oldHeight = Me.Height
oldWidth = Me.Width
deltaHeight = Me.Height - lbDate.Height
gridIsLoad = False
noClick = True
If Regim = "uslug" Then
    Me.Caption = "��������  �� ������ � " & gNzak
    Me.Width = Me.Width - 4200
    Me.Height = Me.Height - 195 * 5
    Grid5.FormatString = "|�������� ���������|��� ���������|�����.��������"
    Grid5.ColWidth(0) = 0
    Grid5.ColWidth(usSumm) = 1650
    Grid5.ColWidth(usOutSum) = 1300
    Grid5.ColWidth(usNowSum) = 1300
    loadOtgruz
    quantity5 = 1 ' ��� Grid5_EnterCell
    noClick = False
    Exit Sub
End If
Me.Caption = "�������� ������� �� ������ � " & gNzak
Grid5.Rows = 3
Grid5.FixedRows = 2
Grid5.MergeRow(0) = True
str = "|�������� ���������"
tmpStr = str & str
str = "|��� ���������"
tmpStr = tmpStr & str & str
str = "|�����.��������"
tmpStr = tmpStr & str & str
Grid5.FormatString = "||<|||" & tmpStr

Grid5.TextMatrix(1, prType) = "���"
Grid5.TextMatrix(1, prName) = "�����"
Grid5.TextMatrix(1, prDescript) = "��������"
Grid5.TextMatrix(1, prEdIzm) = "��.���������"
Grid5.TextMatrix(1, prCenaEd) = "���� �� ��."
str = "���-��"
Grid5.TextMatrix(1, prQuant) = str
Grid5.TextMatrix(1, prOutQuant) = str
Grid5.TextMatrix(1, prNowQuant) = str
str = "�����"
Grid5.TextMatrix(1, prSumm) = str
Grid5.TextMatrix(1, prOutSum) = str
Grid5.TextMatrix(1, prNowSum) = str

Grid5.ColWidth(prId) = 0
Grid5.ColWidth(prType) = 0 '380
Grid5.ColWidth(prName) = 1185
Grid5.ColWidth(prDescript) = 1270 + 380
Grid5.ColWidth(prEdIzm) = 420
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
Dim s As Single

sql = "SELECT ordered From Orders WHERE (((Orders.numOrder)=" & gNzak & "));"
If byErrSqlGetValues("##227", sql, s) Then _
    Grid5.TextMatrix(1, usSumm) = Round(s, 2)
getOtgrugeno 1 ' usNowSum � usOutSum

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

If Regim = "uslug" Then
    loadUslug
    Grid5.col = usNowSum
    mousCol5 = usNowSum
Else
    loadPredmeti Me, "fromOtgruz"
    Grid5.col = prNowQuant
    mousCol5 = prNowQuant
    Grid5.row = 2
    mousRow5 = 2
End If


gridIsLoad = True
End Sub

Function loadOutDates() As Boolean


loadOutDates = False
'���� ��������
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

Sub getOtgrugeno(row As Long, Optional byNomenk As String = "")
Dim s As Single, str  As String
strWhere = "'" & Format(lbDate.Text, "yyyy-mm-dd hh:nn:ss") & "'"

'�������� �� ����

If Regim = "uslug" Then
    sql = "SELECT Sum(quant) AS Sum_quant From xUslugOut " & _
    "WHERE (((numOrder)=" & gNzak & ") AND ((outDate)<" & strWhere & "));"
ElseIf byNomenk = "" Then
    sql = "SELECT Sum(quant) AS Sum_quant From xPredmetyByIzdeliaOut " & _
    "WHERE (((numOrder)=" & gNzak & ") AND ((prId)= " & tbNomenk!prId & _
    ") AND ((prExt)=" & tbNomenk!prExt & ") AND ((outDate)<" & strWhere & "));"
Else
    sql = "SELECT Sum(quant) AS Sum_quant From xPredmetyByNomenkOut " & _
    "WHERE (((numOrder)=" & gNzak & ") AND ((nomNom)= '" & tbNomenk!nomNom & _
    "') AND ((outDate)<" & strWhere & "));"
End If
'MsgBox sql
byErrSqlGetValues "W##203", sql, s
If Regim = "uslug" Then
    Grid5.TextMatrix(row, usOutSum) = s
Else
    Grid5.TextMatrix(row, prOutQuant) = s
    If IsNumeric(tbNomenk!cenaEd) Then _
        Grid5.TextMatrix(row, prOutSum) = Round(tbNomenk!cenaEd * s, 2)
End If

'��������� ��� �������� �� �������
If lbDate.ListIndex = outLen Then ReDim Preserve QQ(row): QQ(row) = s
'�������� �� ����
If Regim = "uslug" Then
    sql = "SELECT Sum(quant) AS Sum_quant From xUslugOut " & _
    "WHERE (((numOrder)=" & gNzak & ") AND ((outDate)=" & strWhere & "));"
ElseIf byNomenk = "" Then
    sql = "SELECT quant From xPredmetyByIzdeliaOut " & _
    "WHERE (((numOrder)=" & gNzak & ") AND ((prId)= " & tbNomenk!prId & _
    ") AND ((prExt)=" & tbNomenk!prExt & ") AND ((outDate)=" & strWhere & "));"
Else
    sql = "SELECT quant  From xPredmetyByNomenkOut " & _
    "WHERE (((numOrder)=" & gNzak & ") AND ((nomNom)= '" & tbNomenk!nomNom & _
    "') AND ((outDate)=" & strWhere & "));"
End If
'MsgBox sql
byErrSqlGetValues "W##204", sql, s

If Regim = "uslug" Then
    Grid5.TextMatrix(row, usNowSum) = s
Else
    Grid5.TextMatrix(row, prNowQuant) = s
    If IsNumeric(tbNomenk!cenaEd) Then _
        Grid5.TextMatrix(row, prNowSum) = Round(tbNomenk!cenaEd * s, 2)
End If
End Sub

'��������� ���� shipped � Orders
Function saveShipped() As Variant
Dim s As Single, s1 As Single

saveShipped = Null
If Regim = "" Then
    sql = "SELECT Sum(xPredmetyByIzdelia.cenaEd*xPredmetyByIzdeliaOut.quant) " & _
    "FROM xPredmetyByIzdelia INNER JOIN xPredmetyByIzdeliaOut ON " & _
    "(xPredmetyByIzdelia.prExt = xPredmetyByIzdeliaOut.prExt) AND " & _
    "(xPredmetyByIzdelia.prId = xPredmetyByIzdeliaOut.prId) AND " & _
    "(xPredmetyByIzdelia.numOrder = xPredmetyByIzdeliaOut.numOrder) " & _
    "WHERE (((xPredmetyByIzdelia.numOrder)=" & gNzak & "));"
    If Not byErrSqlGetValues("W##213", sql, s) Then Exit Function

    sql = "SELECT Sum(xPredmetyByNomenk.cenaEd*xPredmetyByNomenkOut.quant) " & _
    "FROM xPredmetyByNomenk INNER JOIN xPredmetyByNomenkOut ON " & _
    "(xPredmetyByNomenk.nomNom = xPredmetyByNomenkOut.nomNom) AND " & _
    "(xPredmetyByNomenk.numOrder = xPredmetyByNomenkOut.numOrder) " & _
    "WHERE (((xPredmetyByNomenk.numOrder)=" & gNzak & "));"
    If Not byErrSqlGetValues("W##214", sql, s1) Then Exit Function

    s = Round(s + s1, 2)
Else '������
    sql = "SELECT Sum(quant) AS Sum_quant From xUslugOut " & _
    "WHERE (((numOrder)=" & gNzak & "));"
    If Not byErrSqlGetValues("W##301", sql, s) Then Exit Function
End If
If s > 0 Then
    tmpStr = s
Else
    tmpStr = "Null"
End If
'sql = "UPDATE Orders SET shipped = " & tmpStr & " WHERE (((numOrder)=" & gNzak & "));"
'If myExecute("##368", sql) = 0 Then saveShipped = s
orderUpdate "##368", tmpStr, "Orders", "shipped"
saveShipped = s
End Function

Private Sub Form_Resize()

Dim h As Integer, w As Integer
If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid5.Height = Grid5.Height + h
lbDate.Height = Me.Height - deltaHeight
Grid5.Width = Grid5.Width + w

cmDel.Top = cmDel.Top + h
cmCancel.Top = cmCancel.Top + h
cmCancel.Left = cmCancel.Left + w

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

Private Sub Grid5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid5.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid5.ColWidth(Grid5.MouseCol)

End Sub

Sub OutNowSummToGrid5()
Dim il As Long, sum As Single, sum2 As Single
sum = 0: sum2 = 0
For il = 2 To Grid5.Rows - 2
    sum = sum + Grid5.TextMatrix(il, prOutSum)
    sum2 = sum2 + Grid5.TextMatrix(il, prNowSum)
Next il
Grid5.TextMatrix(il, prOutSum) = Round(sum, 2)
Grid5.TextMatrix(il, prNowSum) = Round(sum2, 2)


End Sub

Private Sub lbDate_Click()
If noClick Then Exit Sub
cmDel.Enabled = ((lbDate.ListIndex < outLen) And Not closeZakaz)

gridIsLoad = False
If Regim = "" Then
    loadPredmeti Me, "fromOtgruz"
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
Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim pQuant As Single, s As Single, maxQ As Single

If KeyCode = vbKeyReturn Then
  If Regim = "uslug" Then
    maxQ = Grid5.TextMatrix(1, usSumm)
    maxQ = maxQ - QQ(1) '�� ������� ����� �����������
    maxQ = Round(maxQ + Grid5.TextMatrix(mousRow5, usNowSum), 2)
    If Not isNumericTbox(tbMobile, 0, maxQ) Then Exit Sub
    pQuant = Round(tbMobile.Text, 2)
    tbMobile.Text = pQuant
    
    sql = "SELECT * from xUslugOut WHERE (((numOrder)=" & gNzak & _
    ") AND ((outDate)='" & Format(outDate(lbDate.ListIndex), "yyyy-mm-dd  hh:nn:ss") & "'));"
    Set tbProduct = myOpenRecordSet("##229", sql, dbOpenForwardOnly)
'    If tbProduct Is Nothing Then GoTo ER2
'    tbProduct.index = "Key2"
'    tbProduct.Seek "=", outDate(lbDate.ListIndex), gNzak
'    If tbProduct.NoMatch Then
    If tbProduct.BOF Then
        If pQuant > 0 Then
            tbProduct.AddNew
            tbProduct!outDate = outDate(lbDate.ListIndex)
            tbProduct!numOrder = gNzak
            tbProduct!quant = pQuant
                tbProduct.Update
        End If
    ElseIf pQuant = 0 Then
        tbProduct.Delete
    Else
        tbProduct.Edit
        tbProduct!quant = pQuant
        tbProduct.Update
    End If
    tbProduct.Close
    
    cErr = 302 '##302
    If Not IsNumeric(saveShipped) Then GoTo ER2
    Orders.openOrdersRowToGrid "##228":  tqOrders.Close

  '������������ ��� �������� �� �������
    QQ(1) = QQ(1) + pQuant - Grid5.TextMatrix(1, usNowSum)
    Grid5.TextMatrix(1, usNowSum) = pQuant
    lbHide5
    Exit Sub
  End If
  
  
    On Error GoTo ER1
  wrkDefault.BeginTrans
  
  maxQ = Grid5.TextMatrix(mousRow5, prQuant) '�������� ���������
  maxQ = maxQ - QQ(mousRow5) '�� ������� ����� �����������
  maxQ = Round(maxQ + Grid5.TextMatrix(mousRow5, prNowQuant), 2)
  If Not isNumericTbox(tbMobile, 0, maxQ) Then wrkDefault.Rollback: Exit Sub
  If Grid5.TextMatrix(mousRow5, prType) = "�������" Then
    pQuant = Round(tbMobile.Text)
    tbMobile.Text = pQuant
  
    cErr = "202" '##202
    '���-�
    sql = "SELECT * FROM xPredmetyByIzdeliaOut " & _
    "WHERE (((outDate)='" & Format(outDate(lbDate.ListIndex), "yyyy-mm-dd  hh:nn:ss") & "'" & _
    ") AND ((numOrder)=" & gNzak & ") AND ((prId)=" & gProductId & _
    ") AND ((prExt)=" & prExt & "));"
    Debug.Print sql
    Set tbProduct = myOpenRecordSet("##200", sql, dbOpenForwardOnly)
'    If tbProduct Is Nothing Then GoTo ER1
'    tbProduct.index = "Key"
'    tbProduct.Seek "=", outDate(lbDate.ListIndex), gNzak, gProductId, prExt
'    If tbProduct.NoMatch Then
    If tbProduct.BOF Then
        If pQuant > 0 Then
            tbProduct.AddNew
            tbProduct!outDate = outDate(lbDate.ListIndex)
            tbProduct!numOrder = gNzak
            tbProduct!prId = gProductId
            tbProduct!prExt = prExt
            tbProduct!quant = pQuant
            tbProduct.Update
        End If
    ElseIf pQuant = 0 Then
        tbProduct.Delete
    Else
        tbProduct.Edit
        tbProduct!quant = pQuant
        tbProduct.Update
    End If
    tbProduct.Close
  Else '��������� ���-��
    pQuant = Round(tbMobile.Text, 2)
    tbMobile.Text = pQuant

    sql = "SELECT * from xPredmetyByNomenkOut " & _
    "WHERE (((outDate)='" & Format(outDate(lbDate.ListIndex), "yyyy-mm-dd  hh:nn:ss") & _
    "') AND ((numOrder)=" & gNzak & ") AND ((nomNom)='" & gNomNom & "'));"
'MsgBox sql
    Set tbNomenk = myOpenRecordSet("##201", sql, dbOpenForwardOnly)
'    If tbNomenk Is Nothing Then GoTo ER1
'    tbNomenk.index = "Key"
'    tbNomenk.Seek "=", outDate(lbDate.ListIndex), gNzak, gNomNom
'    If tbNomenk.NoMatch Then
    If tbNomenk.BOF Then
        If pQuant > 0 Then
            tbNomenk.AddNew
            tbNomenk!outDate = outDate(lbDate.ListIndex)
            tbNomenk!numOrder = gNzak
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
    
  End If

  cErr = 216 '##216
  If Not IsNumeric(saveShipped) Then GoTo ER1
  
  Orders.openOrdersRowToGrid "##217":  tqOrders.Close

  wrkDefault.CommitTrans
  '������������ ��� �������� �� �������
  QQ(mousRow5) = QQ(mousRow5) + pQuant - Grid5.TextMatrix(mousRow5, prNowQuant)
  
  Grid5.TextMatrix(mousRow5, prNowQuant) = pQuant
  maxQ = Grid5.TextMatrix(mousRow5, prCenaEd)
  Grid5.TextMatrix(mousRow5, prNowSum) = Round(pQuant * maxQ, 2)
  
  OutNowSummToGrid5
  
EN1: lbHide5
ElseIf KeyCode = vbKeyEscape Then
  lbHide5
End If


Exit Sub
ER1:
errorCodAndMsg ("�������� �� ������")
wrkDefault.Rollback
ER2:
lbHide5
MsgBox "�������� �� ������", , "Error-" & cErr & " �������� ��������������"
End Sub