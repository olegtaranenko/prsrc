VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CehOrders 
   BackColor       =   &H8000000A&
   Caption         =   " "
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "CehOrders.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmNaklad 
      Caption         =   "���������� ���������"
      Height          =   315
      Left            =   3120
      TabIndex        =   15
      Top             =   5340
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '���
      Caption         =   "Frame1"
      Height          =   435
      Left            =   5040
      TabIndex        =   13
      Top             =   2280
      Width           =   2295
      Begin VB.Label Label2 
         Alignment       =   2  '���������
         Appearance      =   0  '������
         BackColor       =   &H80000005&
         BorderStyle     =   1  '����������� ����
         Caption         =   "���� ��������..."
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   2175
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4320
      Top             =   5280
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "������"
      Height          =   255
      Left            =   10980
      TabIndex        =   12
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmZagruz 
      Caption         =   "��������"
      Height          =   315
      Left            =   8760
      TabIndex        =   11
      Top             =   5340
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmExAll 
      Caption         =   "�����"
      Height          =   315
      Left            =   10740
      TabIndex        =   10
      Top             =   5340
      Width           =   975
   End
   Begin VB.TextBox tbNomZak 
      Height          =   285
      Left            =   3780
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.CheckBox chSingl 
      Caption         =   "������ �����"
      Height          =   195
      Left            =   2460
      TabIndex        =   7
      Top             =   60
      Width           =   1335
   End
   Begin VB.CheckBox chDetail 
      Caption         =   "��������� <F2>"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox lbProblem 
      Height          =   1425
      Left            =   3300
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmRefresh 
      Caption         =   "��������"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   5340
      Width           =   915
   End
   Begin VB.ListBox lbStatus 
      Height          =   1815
      ItemData        =   "CehOrders.frx":030A
      Left            =   540
      List            =   "CehOrders.frx":0329
      TabIndex        =   3
      Top             =   2820
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ListBox lbObrazec 
      Height          =   450
      ItemData        =   "CehOrders.frx":0361
      Left            =   1560
      List            =   "CehOrders.frx":036B
      TabIndex        =   2
      Top             =   4140
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lbMaket 
      Height          =   450
      ItemData        =   "CehOrders.frx":0378
      Left            =   2460
      List            =   "CehOrders.frx":0382
      TabIndex        =   1
      Top             =   4140
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8705
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      Caption         =   "<F1>"
      Height          =   195
      Left            =   5040
      TabIndex        =   9
      Top             =   60
      Width           =   375
   End
   Begin VB.Menu mnNomZak 
      Caption         =   "����� ������"
      Visible         =   0   'False
      Begin VB.Menu mnFind 
         Caption         =   "����� � ������� ������"
      End
      Begin VB.Menu mnCancel 
         Caption         =   " "
      End
   End
End
Attribute VB_Name = "CehOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cehRows As Long, cehRowsOld As Long
Dim sum As Single
Dim marker As String ' ������ � 0 ������� ���������� ��� lb, ���-�� �� muose
Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����
Dim colWdth(20) As Integer
Public Regim As String ' ����� ����
Public mousRow As Long    '
Public mousCol As Long    '
Dim maxExt



Private Sub chDetail_Click()
cehBegin
gridIsLoad = True
Grid.col = chKey
Grid.col = 1
Grid.SetFocus
End Sub

Private Sub chSingl_Click()
If chSingl.value = 1 And Not IsNumeric(tbNomZak.Text) Then
    MsgBox "����� ������ ������ �������.", , "��������������:"
    chSingl.value = 0
    Exit Sub
End If
cehBegin
gridIsLoad = True
Grid.col = chKey
Grid.col = 1
Grid.SetFocus

End Sub

Private Sub cmExAll_Click()
Unload Me
End Sub

Private Sub cmNaklad_Click()
#If Not COMTEC = 1 Then '----------------------------------------------------
sDocs.Regim = "fromCeh"
sDocs.Show vbModal
#End If '--------------------------------------------------------------
End Sub

Private Sub cmPrint_Click()
Me.PrintForm
'Me.Height = 20000 ��� ���� Err384, ���� ����� ��� ���������������
End Sub

Private Sub cmRefresh_Click()
cehBegin
gridIsLoad = True
Grid.col = 1
End Sub

Private Sub cmZagruz_Click()
Zagruz.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then lbHide
If KeyCode = vbKeyF1 Then
    If chSingl.value = 1 Then
        chSingl.value = 0
    Else
        chSingl.value = 1
    End If
End If
End Sub

Sub cehBegin()
Dim str As String, I As Integer, j As Integer, il As Long, tmpTopRow As Long
tmpTopRow = Grid.TopRow

#If onErrorOtlad Then
    On Error GoTo errMsg
    GoTo START
errMsg:
    MsgBox Error, , "������  " & Err & " � �\� cehBegin" '
    End
START:
#End If

gridIsLoad = False
Screen.MousePointer = flexHourglass

#If Not COMTEC = 1 Then '----------------------------------------------
getNakladnieList "ceh"
#End If '--------------------------------------------------------------

' ���������� ��������� ��������
colWdth(chNomZak) = Grid.ColWidth(chNomZak)
colWdth(chIzdelia) = Grid.ColWidth(chIzdelia)
colWdth(chM) = Grid.ColWidth(chM)
colWdth(chVrVip) = Grid.ColWidth(chVrVip)
colWdth(chStatus) = Grid.ColWidth(chStatus)
colWdth(chProcVip) = Grid.ColWidth(chProcVip)
colWdth(chProblem) = Grid.ColWidth(chProblem)
colWdth(chDataVid) = Grid.ColWidth(chDataVid)
'colWdth(chDataRes) = Grid.ColWidth(chDataRes)
colWdth(chVrVid) = Grid.ColWidth(chVrVid)
colWdth(chFirma) = Grid.ColWidth(chFirma)
colWdth(chLogo) = Grid.ColWidth(chLogo) + Grid.ColWidth(chDataRes)

Grid.Visible = False
For il = Grid.Rows To 3 Step -1
    Grid.RemoveItem (il)
Next il
Grid.row = 1
For il = 0 To Grid.Cols - 1
    Grid.col = il
    Grid.CellBackColor = Grid.BackColor
    Grid.CellForeColor = vbBlack
    Grid.TextMatrix(1, il) = ""
Next il
' ��������������� ��������� ��������
Grid.ColWidth(chNomZak) = colWdth(chNomZak)
Grid.ColWidth(chIzdelia) = colWdth(chIzdelia)
Grid.ColWidth(chM) = colWdth(chM)
Grid.ColWidth(chVrVip) = colWdth(chVrVip)
Grid.ColWidth(chStatus) = colWdth(chStatus)
Grid.ColWidth(chProcVip) = colWdth(chProcVip)
Grid.ColWidth(chProblem) = colWdth(chProblem)
Grid.ColWidth(chDataVid) = colWdth(chDataVid)
If chDetail.value = 1 Then
    Grid.ColWidth(chDataRes) = 740
Else
    Grid.ColWidth(chDataRes) = 0
End If
Grid.ColWidth(chVrVid) = colWdth(chVrVid)
Grid.ColWidth(chFirma) = colWdth(chFirma)
Grid.ColWidth(chLogo) = colWdth(chLogo) - Grid.ColWidth(chDataRes)

Me.Caption = Ceh(cehId) & mainTitle
sql = "select * from w" & Ceh(cehId)
'Set myQuery = myBase.Connection.QueryDefs("w" & Ceh(cehId))
Set tbCeh = myOpenRecordSet("##34", sql, dbOpenDynaset)
If tbCeh Is Nothing Then myQuery.Close: myBase.Close: End

cehRows = 0
If Not tbCeh.BOF Then
  
  tbCeh.MoveFirst
  While Not tbCeh.EOF
    gNzak = tbCeh!numOrder
    
    If chSingl.value = 1 And gNzak <> tbNomZak.Text Then GoTo NXT
'If gNzak = 3103125 Then
'    gNzak = gNzak
'End If
    If IsDate(tbCeh!DateTimeMO) Then
      If tbCeh!DateTimeMO < CDate("01.01.2000") _
      Or tbCeh!DateTimeMO > CDate("01.01.2050") Then
            msgOfZakaz "##308", "������������ ���� ��. ���������� � ���������. "
            GoTo NXT
      End If
      If IsNull(tbCeh!workTimeMO) Then
            toCehFromStr "m" '�����
      Else  ' �������
        toCehFromStr "o" '�����
      End If ' �������
    End If 'MO
MN:
    toCehFromStr '************************************
NXT:
    tbCeh.MoveNext
  Wend
End If
tbCeh.Close
'myQuery.Close

Grid.col = chKey: Grid.Sort = 3 '�������� ����.
Grid.row = 1

If cehRows = cehRowsOld Then Grid.TopRow = tmpTopRow
cehRowsOld = cehRows

Grid.Visible = True
On Error Resume Next
Grid.SetFocus
Screen.MousePointer = flexDefault
Frame1.Visible = False
End Sub

Sub toCehFromStr(Optional isMO As String = "")
Dim str As String, I As Integer, j As Integer, k As Integer, s As Variant
Dim color As Long, str1 As String  ', is100 As Boolean

#If onErrorOtlad Then
    On Error GoTo errMsg
    GoTo START
errMsg:
    MsgBox Error, , "������  " & Err & " � �\� toCehFromStr" '
    End
START:
#End If

k = 0
marker = ""
color = vbBlack
'If sampl = "" Then
If isMO <> "o" Then
    str = ""
    If tbCeh!StatusId = 2 Then '������
        color = vbBlue
    ElseIf tbCeh!StatusId = 3 Or tbCeh!StatusId = 9 Then ' ��������
        color = &HAA00& ' �.���.
    ElseIf tbCeh!StatusId = 5 Then ' �������
        marker = "�"
        color = vbRed
    ElseIf tbCeh!StatusId = 1 Or tbCeh!StatusId = 8 Or tbCeh!StatusId = 4 Then ' � ������ � �����
        marker = "�"
    End If
Else
    marker = "�"
    str = "o"
End If

If isMO = "m" Then ' �����
    If cehRows > 0 Then Grid.AddItem ("")
    str = "�"
    cehRows = cehRows + 1
    If tbCeh!statM = "�����" Then
        Grid.TextMatrix(cehRows, chStatus) = tbCeh!statM
    Else
        Grid.TextMatrix(cehRows, chStatus) = ""
    End If
    marker = "�"
    LoadDateKey tbCeh!DateTimeMO, "##38"
    LoadDate Grid, cehRows, chVrVid, tbCeh!DateTimeMO, "hh"
    GoTo MN
End If

    If cehRows > 0 Then Grid.AddItem ("") '����� ��������� ��� ���.�����
    cehRows = cehRows + 1
    Grid.col = chNomZak
    Grid.row = cehRows
    Grid.CellForeColor = color
 
    If str = "" Then '���.����� ������
        s = Round(100 * (1 - tbCeh!nevip), 1)
        If s > 0 Then Grid.TextMatrix(cehRows, chProcVip) = s
        
        s = tbCeh!workTime
        LoadDateKey tbCeh!outDateTime, "##36"
        LoadDate Grid, cehRows, chVrVid, tbCeh!outDateTime, "hh"
    Else
        If tbCeh!statO = "�����" Then _
            Grid.TextMatrix(cehRows, chProcVip) = "100"
        s = tbCeh!workTimeMO
        If s < 0 Then s = -s
        LoadDateKey tbCeh!DateTimeMO, "##36"
        LoadDate Grid, cehRows, chVrVid, tbCeh!DateTimeMO, "hh"
    End If
    If IsNull(s) Then
        msgOfZakaz ("##36")
        Grid.TextMatrix(cehRows, chVrVip) = "(??) "
    Else
      If chDetail.value = 1 Then '
        Grid.TextMatrix(cehRows, chVrVip) = "(" & s & ")"
      Else
        Grid.TextMatrix(cehRows, chVrVip) = s
      End If
    End If
If isMO = "o" Then
   If tbCeh!statO = "�����" Then
     Grid.TextMatrix(cehRows, chStatus) = tbCeh!statO '�������
   Else
     Grid.TextMatrix(cehRows, chStatus) = "" '�������
   End If
ElseIf (tbCeh!StatusId = 1 Or tbCeh!StatusId = 8) And Not IsNumeric(tbCeh!stat) Then
    Grid.TextMatrix(cehRows, chStatus) = tbCeh!stat
ElseIf tbCeh!StatusId = 2 Then ' ������
    str1 = "�": GoTo AA
ElseIf tbCeh!StatusId = 3 Or tbCeh!StatusId = 9 Then  ' ��������
    str1 = "�"
AA: Grid.col = chStatus
    Grid.CellForeColor = color
    Grid.TextMatrix(cehRows, chStatus) = str1 & " �� " & Format(tbCeh!dateRS, "dd.mm.yy")
Else
    Grid.TextMatrix(cehRows, chStatus) = status(tbCeh!StatusId)
End If
MN:
#If Not COMTEC = 1 Then '----------------------------------------------
 For I = 1 To UBound(tmpL) '�������� ������ � ����������� ����������
    If tmpL(I) = gNzak Then
        Grid.col = chIzdelia
        Grid.row = cehRows
        Grid.CellForeColor = 200
        Exit For
    End If
 Next I
#End If '--------------------------------------------------------------
Grid.TextMatrix(cehRows, 0) = marker
Grid.TextMatrix(cehRows, chNomZak) = gNzak & str
If str <> "" Then colorGridRow Grid, cehRows, &HCCCCCC '��������� ��
Grid.TextMatrix(cehRows, chM) = tbCeh!Manag
Grid.TextMatrix(cehRows, chFirma) = tbCeh!name
Grid.TextMatrix(cehRows, chLogo) = tbCeh!Logo
Grid.TextMatrix(cehRows, chIzdelia) = tbCeh!Product
If tbCeh!StatusId = 5 Then ' �������
        Grid.TextMatrix(cehRows, chProblem) = Problems(tbCeh!ProblemId)
End If

End Sub

Sub LoadDateKey(val As Variant, myErr As String)
Dim I As Integer

If Not IsNull(val) Then
  If IsDate(val) Then
    Grid.TextMatrix(cehRows, chDataVid) = Format(val, "dd.mm.yy")
    I = DateDiff("d", curDate, val) + 1 '�����
    Grid.TextMatrix(cehRows, chKey) = I
'    If i = stDay Then
'        Grid.col = chDataVid
'        Grid.CellForeColor = &H8800&
'        Grid.CellFontBold = True
'    End If
    Exit Sub
  End If
End If
msgOfZakaz (myErr)
Grid.TextMatrix(cehRows, chDataRes) = "??"
Grid.TextMatrix(cehRows, chKey) = 0
End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width

#If Not COMTEC = 1 Then '----------------------------------------------
cmNaklad.Visible = True
#End If '--------------------------------------------------------------

'If dostup = "y" Or dostup = "c" Then cmZagruz.Visible = True
If Not (dostup = "a" Or dostup = "m" Or dostup = "" Or dostup = "b") Then
    cmZagruz.Visible = True
    Orders.managLoad "fromCeh" ' �������� Manag()
End If
If dostup = "" Then cmNaklad.Visible = False


Screen.MousePointer = flexHourglass

Dim I As Integer
For I = begCehProblemId To lenProblem
    lbProblem.AddItem Problems(I)
Next I

Grid.FormatString = "    |<� ������|^�|������ |>��.���|>%��|��������|" & _
"<���� ������|<��.���|<���� �������|<��������|<����|<�������|����"

Grid.ColWidth(chM) = 270
Grid.ColWidth(chVrVip) = 388
Grid.ColWidth(chStatus) = 870
Grid.ColWidth(chProcVip) = 420
Grid.ColWidth(chProblem) = 900
Grid.ColWidth(chDataRes) = 735
Grid.ColWidth(chVrVid) = 330
Grid.ColWidth(chDataVid) = 735
Grid.ColWidth(chFirma) = 2000
Grid.ColWidth(chLogo) = 1200
Grid.ColWidth(chKey) = 0 ' ��� ���������� �� ����
Grid.ColWidth(0) = 0
Grid.ColWidth(chNomZak) = 1000
Grid.ColWidth(chIzdelia) = 2450

Timer1.Interval = 500
Timer1.Enabled = True '����� cehBegin
End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next
lbHide
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w
cmRefresh.Top = cmRefresh.Top + h
cmExAll.Top = cmExAll.Top + h
cmExAll.Left = cmExAll.Left + w
cmZagruz.Top = cmZagruz.Top + h
cmZagruz.Left = cmZagruz.Left + w
cmPrint.Left = cmPrint.Left + w
cmNaklad.Top = cmNaklad.Top + h
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (dostup = "a" Or dostup = "m" Or dostup = "" Or dostup = "b") Then
    exitAll '��� �����
End If
isCehOrders = False
End Sub


Private Sub Grid_Click()
If Not gridIsLoad Then Exit Sub
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If mousRow = 0 Then
    Grid.CellBackColor = Grid.BackColor
    If mousCol = 0 Then Exit Sub
    If mousCol = chNomZak Then
        SortCol Grid, mousCol
    ElseIf mousCol = chDataRes Or mousCol = chDataVid Then
        SortCol Grid, mousCol, "date"
    Else
        SortCol Grid, mousCol
    End If
    Grid.row = 1    ' ������ ����� ����� ���������
    Grid_EnterCell

End If
End Sub

Private Sub Grid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Dim date1 As Date, date2 As Date ' � 2 � ������
Dim date1S, date2S As String

date1S = sortGrid.TextMatrix(Row1, mousCol)
date2S = sortGrid.TextMatrix(Row2, mousCol)

If Not IsDate(date1S) Then
    Cmp = -1
    GoTo CC:
ElseIf Not IsDate(date2S) Then
    Cmp = 1
    GoTo CC:
End If

date1 = date1S
date2 = date2S
If date1 > date2 Then
    Cmp = 1
ElseIf date1 < date2 Then
    Cmp = -1
Else
    Cmp = 0
End If
CC:
If trigger Then Cmp = -Cmp

End Sub

Private Sub Grid_DblClick()

'If mousCol = chNomZak And dostup <> "c" And dostup <> "y" Then
If mousCol = chNomZak And (dostup = "a" Or dostup = "m" Or dostup = "" Or _
dostup = "b") Then Me.PopupMenu mnNomZak

getNumFromStr (Grid.TextMatrix(mousRow, chNomZak))

#If Not COMTEC = 1 Then '----------------------------------------------
If mousCol = chIzdelia And Grid.CellForeColor = 200 Then
    numDoc = gNzak
    numExt = 0
    Nakladna.Regim = "predmeti"
    Nakladna.Show vbModal
End If
#End If '--------------------------------------------------------------

If dostup = "" Then Exit Sub
marker = Grid.TextMatrix(mousRow, 0)
If mousRow = 0 Or marker = "" Then Exit Sub

If mousCol = chStatus Then
    If marker = "�" Then '  "�������"
        listBoxInGridCell lbObrazec, Grid, "select"
    ElseIf marker = "�" Then '      "�����"
        listBoxInGridCell lbMaket, Grid, "select"
    ElseIf LCase$(marker) = "�" Then '  "� ������"
        listBoxInGridCell lbStatus, Grid, "select"
    End If
End If
End Sub

Private Sub Grid_EnterCell()
If Not gridIsLoad Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col
getNumFromStr (Grid.TextMatrix(mousRow, chNomZak))
tbNomZak.Text = gNzak
If dostup = "" Then Exit Sub
marker = Grid.TextMatrix(mousRow, 0)
oldCellColor = Grid.CellBackColor
If (mousCol = chStatus And marker <> "") Or _
(mousCol = chIzdelia And Grid.CellForeColor = 200) Then
    Grid.CellBackColor = &H88FF88
Else
    Grid.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid_DblClick

End Sub

Private Sub Grid_LeaveCell()
If Not gridIsLoad Then Exit Sub
Grid.CellBackColor = oldCellColor
End Sub

Private Sub lbMaket_DblClick()
Dim I As Integer

If noClick Then Exit Sub

sql = "SELECT StatM From OrdersMO WHERE (((numOrder)=" & gNzak & "));"
If Not byErrSqlGetValues("##312", sql, tmpStr) Then Exit Sub
If tmpStr = "���������" Then
    msgZakazDeleted "����� ��� ���������"
    GoTo EN1
ElseIf lbMaket.Text = "�����" Then
    I = ValueToTableField("W##37", "'�����'", "OrdersMO", "StatM")
Else
    I = ValueToTableField("W##37", "'� ������'", "OrdersMO", "StatM")
End If
If I = 0 Then
    Grid.TextMatrix(mousRow, chStatus) = lbMaket.Text
ElseIf I = -1 Then
    msgZakazDeleted
End If
EN1:
lbHide ' � �.�. ���������

End Sub

Private Sub lbMaket_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbMaket_DblClick
End Sub

Private Sub lbObrazec_DblClick()
Dim j As Integer, str As String, old As String, v As Variant
Dim proc As String, status As String
'sChr As String, dChr As String,
If noClick Then Exit Sub
old = Grid.TextMatrix(mousRow, chStatus)
If lbObrazec.Text = "�����" And lbObrazec.Text <> old Then
    proc = "100%": status = "'�����'"
ElseIf lbObrazec.Text <> old Then '              �������
    proc = "0%": status = "'� ������'"
Else
    lbHide
    Exit Sub
End If
lbObrazec.Visible = False

wrkDefault.BeginTrans
    
v = makeProcReady(proc, "obraz")
If IsNull(v) Then ' ������� ���������
    msgZakazDeleted "������� ��� ���������"
ElseIf v Then
    If ValueToTableField("##54", status, "OrdersMO", "StatO") = 0 Then
        wrkDefault.CommitTrans
        cehBegin
    Else
        wrkDefault.Rollback
    End If
Else ' ����� ��� ������ ����������
    wrkDefault.Rollback
    msgZakazDeleted
End If

lbHide ' � �.�. ���������
End Sub

Private Sub lbObrazec_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbObrazec_DblClick
End Sub

Private Sub lbProblem_DblClick()
Dim str As String, I As Integer

If noClick Then Exit Sub

wrkDefault.BeginTrans   ' ������ ����������

I = ValueToTableField("W##41", "'� ������'", "OrdersInCeh", "Stat") '�.� ���� �������� Stat=�����, �� �� ������ �� ���������
If I = 0 Then
    If ValueToTableField("##41", "5", "Orders", "StatusId") <> 0 Then GoTo ER1

    str = lbProblem.ListIndex + begCehProblemId
    If ValueToTableField("##41", str, "Orders", "ProblemId") = 0 Then
        wrkDefault.CommitTrans  ' ������������� ����������
        cehBegin
    Else
ER1:    wrkDefault.Rollback    ' ������� ����������
    End If
ElseIf I = -1 Then
    wrkDefault.Rollback
    msgZakazDeleted
End If

lbHide ' � �.�. ���������
End Sub

Private Sub lbProblem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbProblem_DblClick
ElseIf KeyCode = vbKeyEscape Then
'    wrkDefault.Rollback ' ������� ����������
End If
End Sub

Sub lbHide()
lbStatus.Visible = False
lbObrazec.Visible = False
lbMaket.Visible = False
lbProblem.Visible = False
Grid.Enabled = True
Grid.SetFocus
On Error GoTo ER1 ' �������������� ������ �.�. ��� �������
Grid.row = mousRow
gridIsLoad = True
Grid.col = mousCol
'Grid_EnterCell
Exit Sub
ER1:
gridIsLoad = True
End Sub

Function procReadyIs100() As Boolean
Dim str As String

    procReadyIs100 = False
    str = Grid.TextMatrix(mousRow, chProcVip)
    If Not IsNumeric(str) Then GoTo ERR100
    If str < 99.99 Then
ERR100:
        MsgBox "������ ����� ������ ���� �������� �� 100%", , _
               "������������ ������!"
        lbHide
        Exit Function
    End If
    procReadyIs100 = True

End Function

Private Sub lbStatus_DblClick()
Dim str As String, I As Integer

#If onErrorOtlad Then
    On Error GoTo errMsg
    GoTo START
errMsg:
    MsgBox Error, , "������  " & Err & " � �\� lbStatus_DblClick" '
    End
START:
#End If

If noClick Then Exit Sub
str = lbStatus.Text
If str = "�������" Then
    Grid.col = chProblem
    listBoxInGridCell lbProblem, Grid
    Exit Sub
ElseIf str = "�����" Then
    lbStatus.Visible = False
    If Not procReadyIs100() Then Exit Sub
#If Not COMTEC = 1 Then '----------------------------------------------------
    If Not predmetiIsClose("etap") Then '
        str = "�������� ����� "
        If QQ2(0) = 0 Then str = ""
        MsgBox "�� ����� ������ �������(��������) �� ��� �������� " & str & _
        "(��� ��������� �������� �� ������� �������)!", , _
        "������������ ������ ��� ������ � " & gNzak
'        Grid.SetFocus
    Else
#Else
    If 1 = 1 Then
#End If '-------------------------------------------------------------------
        wrkDefault.BeginTrans
        I = ValueToTableField("W##41", "'" & str & "'", "OrdersInCeh", "Stat")
        If I = 0 Then
            If ValueToTableField("##39", "4", "Orders", "StatusId") <> 0 Then GoTo ER1
            If ValueToTableField("##39", "0", "Orders", "ProblemId") <> 0 Then GoTo ER1
#If Not COMTEC = 1 Then '----------------------------------------------------
'��� ��� �������, ����������� �����.����, ��������, ��� ��� �. � ����� ���-��
            If Not newEtap("xEtapByIzdelia") Then GoTo ER1
            If Not newEtap("xEtapByNomenk") Then GoTo ER1
#End If
            wrkDefault.CommitTrans
            cehBegin
        ElseIf I = -1 Then
            GoTo ER2
        Else
            GoTo ER1
        End If
    End If
ElseIf str = "25%" Or str = "50%" Or str = "75%" Or str = "100%" Then
    lbStatus.Visible = False
    wrkDefault.BeginTrans
    If makeProcReady(str) Then '� � ��� ����� ��� ������� ����� �� ����
        If ValueToTableField("##39", "1", "Orders", "StatusId") <> 0 Then GoTo ER1 ' "� ������"
        str = "� ������"
        GoTo AA
    End If
    GoTo ER2
Else '  �����, "*" � "� ������"
    lbStatus.Visible = False
    wrkDefault.BeginTrans
    If makeProcReady("0%") Then
        If ValueToTableField("##41", "'" & str & "'", "OrdersInCeh", "Stat") <> 0 Then GoTo ER1
        If ValueToTableField("##39", "1", "Orders", "StatusId") <> 0 Then GoTo ER1
AA:     If ValueToTableField("##39", "0", "Orders", "ProblemId") = 0 Then
            wrkDefault.CommitTrans
            cehBegin
        Else
ER1:        wrkDefault.Rollback
        End If
    Else
ER2:    wrkDefault.Rollback
        msgZakazDeleted
    End If
End If
lbHide ' � �.�. ���������
End Sub

Sub msgZakazDeleted(Optional str As String = "")
    If str = "" Then str = "����� ��� ������"
    MsgBox "������ ���� " & str & " ���������� �� ����. ������� " & _
    "������ '��������'.", , "��������������"
End Sub
'$odbc14$
'��� �������, ���. ��������� ���������� Null
Function makeProcReady(stat As String, Optional obraz As String = "") As Variant
Dim s As Single, t As Single, n As Single, virabotka As Single, str As String
Dim statO As String

makeProcReady = False

If stat = "25%" Then
    s = 0.75 ' �����������
    GoTo AA
ElseIf stat = "50%" Then
    s = 0.5
    GoTo AA
ElseIf stat = "75%" Then
    s = 0.25
    GoTo AA
ElseIf stat = "100%" Then
    s = 0
    GoTo AA
Else
    s = 1
AA:
' sql = "SELECT OrdersInCeh.Stat, Orders.workTime, " & _
 "OrdersInCeh.Nevip, OrdersMO.workTimeMO, " & _
 "OrdersMO.StatO  FROM (Orders INNER JOIN OrdersInCeh ON Orders.numOrder = " & _
 "OrdersInCeh.numOrder) LEFT JOIN OrdersMO ON Orders.numOrder = OrdersMO.numOrder " & _
 "WHERE (((Orders.numOrder)=" & gNzak & ") AND ((Orders.CehId)=" & cehId & "));"

'Set table = myOpenRecordSet("##386", sql, dbOpenDynaset) 'dbOpenTable)
'If table Is Nothing Then Exit Function
'If table.BOF Then
'    table.Close: Exit Function
'End If
 
  If obraz <> "" Then
    obraz = "o"
'    If table!statO <> "�����" And lbObrazec.Text = "�����" Then ' 100%
    sql = "SELECT workTimeMO, StatO from OrdersMO WHERE (((numOrder)=" & gNzak & "));"
    If Not byErrSqlGetValues("##386", sql, virabotka, statO) Then Exit Function
    If s = 0 Then ' 100%
'        virabotka = table!workTimeMO
    Else
'        virabotka = -table!workTimeMO
        virabotka = -virabotka
    End If
  Else
    sql = "SELECT Orders.workTime, OrdersInCeh.Nevip " & _
    "FROM Orders INNER JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " & _
    "WHERE (((Orders.numOrder)=" & gNzak & "));"
    If Not byErrSqlGetValues("##421", sql, t, n) Then Exit Function
'    virabotka = Round((table!nevip - s) * table!workTime, 2)
    virabotka = Round((n - s) * t, 2)
  End If
'table.Close



'���-�� ����� ��������� � ������� � 75% �� 0%
  str = Format(curDate, "yy.mm.dd")

  sql = "SELECT xDate, Virabotka, numOrder, obrazec from Itogi_" & Ceh(cehId) & _
  " WHERE (((xDate)='" & str & "') AND ((numOrder)=" & gNzak & ") AND " & _
  "((obrazec)='" & obraz & "'));"
  'Set tbOrders = myOpenRecordSet("##374", "Itogi_" & Ceh(cehId), dbOpenTable)
  Set tbOrders = myOpenRecordSet("##374", sql, dbOpenTable)
'If Not tbOrders Is Nothing Then
'    tbOrders.index = "Key"
'    tbOrders.Seek "=", str, gNzak, obraz
    
    
'    If tbOrders.NoMatch Then
    If tbOrders.BOF Then
        tbOrders.AddNew
        tbOrders!xDate = str
        tbOrders!numOrder = gNzak
        tbOrders!obrazec = obraz
    Else
        virabotka = virabotka + tbOrders!virabotka
        tbOrders.Edit
    End If
    tbOrders!virabotka = virabotka
    tbOrders.Update
    tbOrders.Close
'End If
    
' table.Edit ''$odbc18!$ �������� ������ ��� ������
'   If obraz = "" Then table!nevip = s

   If obraz = "o" Then '          ��� �������
'    If IsNull(table!statO) Then msgOfZakaz "##311", "� ������� � ���� ��� �������"
    'If table!statO = "���������" Then
      If statO = "���������" Then
'        table.Close
        makeProcReady = Null
        Exit Function
      End If
'   End If
   Else 'obraz = ""
'  If obraz <> "o" Then
'     table!stat = "� ������"     '������ �� ��� MO
     sql = "UPDATE OrdersInCeh SET Stat = '� ������', " & _
     "Nevip = " & s & " WHERE (((numOrder)=" & gNzak & "));"
     If myExecute("##422", sql) <> 0 Then Exit Function
'     table!nevip = s
   End If
' table.Update
' table.Close
End If 'If stat
makeProcReady = True


End Function

Private Sub lbStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbStatus_DblClick
End Sub

Private Sub mnFind_Click()
Orders.Show
Orders.loadWithFiltr gNzak
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False

cehBegin
gridIsLoad = True
Grid.col = 1
isCehOrders = True
trigger = True

End Sub

#If Not COMTEC = 1 Then '----------------------------------------------------

Function newEtap(table As String) As Boolean
newEtap = False
sql = "UPDATE " & table & " SET prevQuant = [eQuant] " & _
"WHERE (((numOrder)=" & gNzak & "));"
If myExecute("##193", sql, 0) > 0 Then Exit Function
newEtap = True
End Function

#End If

