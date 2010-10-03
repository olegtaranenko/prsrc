VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form WerkOrders 
   BackColor       =   &H8000000A&
   Caption         =   " "
   ClientHeight    =   5784
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   Icon            =   "CehOrders.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5784
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmNakladZakaz 
      Caption         =   "��������� ��� �����"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3480
      TabIndex        =   24
      Top             =   5340
      Width           =   1932
   End
   Begin VB.Frame frmRemark 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3675
      Left            =   6840
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   4755
      Begin VB.TextBox tbType 
         Height          =   2835
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   240
         Width           =   4515
      End
      Begin VB.CommandButton cmCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   3600
         TabIndex        =   20
         Top             =   3240
         Width           =   795
      End
      Begin VB.CommandButton cmOk 
         Caption         =   "Ok"
         Height          =   315
         Left            =   660
         TabIndex        =   19
         Top             =   3240
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "���������� � ������"
         Height          =   252
         Left            =   360
         TabIndex        =   23
         Top             =   0
         Width           =   1872
      End
      Begin VB.Label laNumorderRemark 
         Height          =   252
         Left            =   2280
         TabIndex        =   22
         Top             =   0
         Width           =   912
      End
   End
   Begin VB.ComboBox cbEquips 
      Height          =   288
      Left            =   6840
      TabIndex        =   16
      Top             =   5400
      Width           =   2652
   End
   Begin VB.CommandButton cmNaklad 
      Caption         =   "��������� �� �� ������"
      Height          =   315
      Left            =   1200
      TabIndex        =   15
      Top             =   5340
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   435
      Left            =   5040
      TabIndex        =   13
      Top             =   2280
      Width           =   2295
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
      Left            =   7560
      Top             =   5400
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
      Left            =   9480
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
      Height          =   1200
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
      Height          =   1776
      ItemData        =   "CehOrders.frx":030A
      Left            =   540
      List            =   "CehOrders.frx":0329
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.ListBox lbObrazec 
      Height          =   432
      ItemData        =   "CehOrders.frx":0361
      Left            =   1560
      List            =   "CehOrders.frx":036B
      TabIndex        =   2
      Top             =   4140
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lbMaket 
      Height          =   432
      ItemData        =   "CehOrders.frx":0378
      Left            =   2460
      List            =   "CehOrders.frx":0382
      TabIndex        =   1
      Top             =   4140
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   8700
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label lbEquips 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Caption         =   "������������:"
      Height          =   252
      Left            =   5400
      TabIndex        =   17
      Top             =   5400
      Width           =   1332
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
Attribute VB_Name = "WerkOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public idWerk As Integer
Dim werkRows As Long, werkRowsOld As Long
Dim sum As Single
Dim marker As String ' ������ � 0 ������� ���������� ��� lb, ���-�� �� muose
Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����
Dim colWdth(20) As Integer
Public Regim As String ' ����� ����
Public mousRow As Long    '
Public mousCol As Long    '
'Public werkId As Integer
Dim maxExt

Dim tbCeh As Recordset
Dim idEquip As Integer


Private Sub cbEquips_Click()
    If noClick Then Exit Sub
    cmNakladZakaz.Enabled = False
    idEquip = cbEquips.ItemData(cbEquips.ListIndex)
    werkBegin
    gridIsLoad = True
End Sub

Private Sub chDetail_Click()
Dim StatusId As String, Worktime As String, Left As String, Numorder As String, Outdatetime As String, Rollback As String

werkBegin
gridIsLoad = True
Grid.col = chKey
Grid.col = 1
Grid.SetFocus
End Sub

Private Sub chSingl_Click()
If chSingl.Value = 1 And Not IsNumeric(tbNomZak.Text) Then
    MsgBox "����� ������ ������ �������.", , "��������������:"
    chSingl.Value = 0
    Exit Sub
End If
werkBegin
gridIsLoad = True
Grid.col = chKey
Grid.col = 1
Grid.SetFocus

End Sub

Private Sub cmEquip_Click(Index As Integer)
    idEquip = Index
    werkBegin
    gridIsLoad = True
End Sub

Private Sub cmExAll_Click()
Unload Me
End Sub

Private Sub cmNaklad_Click()
sDocs.Regim = "fromCeh"
sDocs.idWerk = idWerk
sDocs.Show vbModal
End Sub

Private Sub cmCancel_Click()
    frmRemark.Visible = False
    Grid.Enabled = True
    Grid.SetFocus
End Sub

Private Sub callNaklad()
    Dim myWerkId As Integer
    numDoc = gNzak
    numExt = 0
    Nakladna.Regim = "predmeti"
    
    Nakladna.idWerk = Grid.TextMatrix(mousRow, chWerkId)
    Nakladna.Show vbModal

End Sub


Private Sub cmNakladZakaz_Click()
    callNaklad
End Sub

Private Sub cmOk_Click()
'    orderUpdate "##19.3", "'" & tbType.Text & "'", "Orders", "Remark"
'    openOrdersRowToGrid "##activate", True
'    tqOrders.Close
'    frmRemark.Visible = False
'    Grid.Enabled = True
'    Grid.col = orRemark
'    Grid.SetFocus
End Sub


Private Sub cmPrint_Click()
Me.PrintForm
'Me.Height = 20000 ��� ���� Err384, ���� ����� ��� ���������������
End Sub

Private Sub cmRefresh_Click()
cmNakladZakaz.Enabled = False
werkBegin
gridIsLoad = True
Grid.col = 1
End Sub

Private Sub cmZagruz_Click()
Zagruz.Show
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then lbHide
If KeyCode = vbKeyF1 Then
    If chSingl.Value = 1 Then
        chSingl.Value = 0
    Else
        chSingl.Value = 1
    End If
End If
End Sub


Sub werkBegin(Optional doEquipToolbar As Boolean = False)
Dim str As String, I As Integer, J As Integer, IL As Long, tmpTopRow As Long
tmpTopRow = Grid.TopRow

#If onErrorOtlad Then
    On Error GoTo errMsg
    GoTo START
errMsg:
    MsgBox Error, , "������  " & Err & " � �\� werkBegin" '
    End
START:
#End If

gridIsLoad = False
Screen.MousePointer = flexHourglass

getNakladnieList "werk"

' ���������� ��������� ��������
colWdth(chNomZak) = Grid.ColWidth(chNomZak)
colWdth(chM) = Grid.ColWidth(chM)
colWdth(chEquip) = Grid.ColWidth(chEquip)
colWdth(chStatus) = Grid.ColWidth(chStatus)
colWdth(chVrVip) = Grid.ColWidth(chVrVip)
colWdth(chProcVip) = Grid.ColWidth(chProcVip)
colWdth(chProblem) = Grid.ColWidth(chProblem)
colWdth(chDataVid) = Grid.ColWidth(chDataVid)
'colWdth(chDataRes) = Grid.ColWidth(chDataRes)
colWdth(chVrVid) = Grid.ColWidth(chVrVid)
colWdth(chFirma) = Grid.ColWidth(chFirma)
colWdth(chIzdelia) = Grid.ColWidth(chIzdelia)

If chDetail.Value = 1 Then
    colWdth(chLogo) = Grid.ColWidth(chLogo) + Grid.ColWidth(chDataRes)
Else
    colWdth(chLogo) = Grid.ColWidth(chLogo)
End If


Grid.Visible = False
For IL = Grid.Rows To 3 Step -1
    Grid.RemoveItem (IL)
Next IL
Grid.row = 1
For IL = 0 To Grid.Cols - 1
    Grid.col = IL
    Grid.CellBackColor = Grid.BackColor
    Grid.CellForeColor = vbBlack
    Grid.TextMatrix(1, IL) = ""
Next IL

' ��������������� ��������� ��������
Grid.ColWidth(chNomZak) = colWdth(chNomZak)
Grid.ColWidth(chM) = colWdth(chM)
Grid.ColWidth(chEquip) = colWdth(chEquip)
Grid.ColWidth(chVrVip) = colWdth(chVrVip)
Grid.ColWidth(chStatus) = colWdth(chStatus)
Grid.ColWidth(chProcVip) = colWdth(chProcVip)
Grid.ColWidth(chProblem) = colWdth(chProblem)
Grid.ColWidth(chDataVid) = colWdth(chDataVid)
Grid.ColWidth(chIzdelia) = colWdth(chIzdelia)

If chDetail.Value = 1 Then
    Grid.ColWidth(chDataRes) = 740
Else
    Grid.ColWidth(chDataRes) = 0
End If
Grid.ColWidth(chVrVid) = colWdth(chVrVid)
Grid.ColWidth(chFirma) = colWdth(chFirma)
Grid.ColWidth(chLogo) = colWdth(chLogo) - Grid.ColWidth(chDataRes)

Dim RightLinie As Long
Dim equipIndex As Integer, HShift As Long



    'HShift = cmEquip(0).Width + 20
    'RightLinie = cmEquip(0).Left + HShift
    
If doEquipToolbar Then
    initEquipCombo Me.cbEquips, idWerk, idEquip
End If


Dim myTitle(1) As String, mysql(1) As String
myTitle(0) = Werk(idWerk)
If idWerk = 0 Then
    mysql(0) = ""
Else
    mysql(0) = "WerkId = " & idWerk
End If

myTitle(1) = cbEquips.Text
If idEquip = 0 Then
    mysql(1) = ""
Else
    mysql(1) = "equipId = " & idEquip
End If

Dim firstArg As Boolean, Where As String
firstArg = True
For I = 0 To 1
    If mysql(I) <> "" Then
        If firstArg Then
            Where = "WHERE"
            firstArg = False
        Else
            Where = Where & " AND"
        End If
        Where = Where & " " & mysql(I)
    End If
Next I

Me.Caption = myTitle(0) & " - " & myTitle(1) & "  " & mainTitle

sql = "select * from vw_Reestr " & Where


Set tbCeh = myOpenRecordSet("##34", sql, dbOpenDynaset)
If tbCeh Is Nothing Then myQuery.Close: myBase.Close: End

werkRows = 0
Dim MaketFlag As Boolean
Dim MaketNumorder As Long
MaketNumorder = 0

If Not tbCeh.BOF Then
  
  tbCeh.MoveFirst
  While Not tbCeh.EOF
    gNzak = tbCeh!Numorder
    If gNzak <> MaketNumorder Then
        MaketFlag = True
        MaketNumorder = gNzak
    End If
    
    If chSingl.Value = 1 And gNzak <> tbNomZak.Text Then GoTo NXT
    If IsDate(tbCeh!DateTimeMO) Then
      If tbCeh!DateTimeMO < CDate("01.01.2000") _
        Or tbCeh!DateTimeMO > CDate("01.01.2150") _
      Then
        msgOfZakaz "##308", "������������ ���� ��. ���������� � ���������. ", tbCeh!Manag
        GoTo NXT
      End If
      If IsNull(tbCeh!WorktimeMO) Then
        If MaketFlag Then
            toCehFromStr "m" '�����
            MaketFlag = False
        End If
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

Grid.col = chKey: Grid.Sort = 3 '�������� ����.
Grid.row = 1

If werkRows = werkRowsOld Then Grid.TopRow = tmpTopRow
werkRowsOld = werkRows

Grid.Visible = True
On Error Resume Next
Grid.SetFocus
Screen.MousePointer = flexDefault
Frame1.Visible = False
End Sub

Sub toCehFromStr(Optional isMO As String = "")
Dim str As String, I As Integer, J As Integer, K As Integer, S As Variant
Dim color As Long, str1 As String  ', is100 As Boolean

#If onErrorOtlad Then
    On Error GoTo errMsg
    GoTo START
errMsg:
    MsgBox Error, , "������  " & Err & " � �\� toCehFromStr" '
    End
START:
#End If

K = 0
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
    If werkRows > 0 Then Grid.AddItem ("")
    str = "�"
    werkRows = werkRows + 1
    If tbCeh!StatM = "�����" Then
        Grid.TextMatrix(werkRows, chStatus) = tbCeh!StatM
    Else
        Grid.TextMatrix(werkRows, chStatus) = ""
    End If
    marker = "�"
    LoadDateKey tbCeh!DateTimeMO, "##38"
    LoadDate Grid, werkRows, chVrVid, tbCeh!DateTimeMO, "hh"
    GoTo MN
End If

    If werkRows > 0 Then Grid.AddItem ("") '����� ��������� ��� ���.�����
    werkRows = werkRows + 1
    
    Grid.TextMatrix(werkRows, chEquip) = tbCeh!Equip
    Grid.TextMatrix(werkRows, chEquipId) = tbCeh!EquipId
    
    Grid.col = chNomZak
    Grid.row = werkRows
    Grid.CellForeColor = color
 
    If str = "" Then '���.����� ������
        S = Round(100 * (1 - tbCeh!nevip), 1)
        If S > 0 Then Grid.TextMatrix(werkRows, chProcVip) = S
        
        S = tbCeh!Worktime
        LoadDateKey tbCeh!Outdatetime, "##36"
        LoadDate Grid, werkRows, chVrVid, tbCeh!Outdatetime, "hh"
    Else
        If tbCeh!StatO = "�����" Then _
            Grid.TextMatrix(werkRows, chProcVip) = "100"
        S = tbCeh!WorktimeMO
        If S < 0 Then S = -S
        LoadDateKey tbCeh!DateTimeMO, "##36"
        LoadDate Grid, werkRows, chVrVid, tbCeh!DateTimeMO, "hh"
    End If
    If IsNull(S) Then
        msgOfZakaz ("##36"), , tbCeh!Manag
        Grid.TextMatrix(werkRows, chVrVip) = "(??) "
    Else
      If chDetail.Value = 1 Then '
        Grid.TextMatrix(werkRows, chVrVip) = "(" & S & ")"
      Else
        Grid.TextMatrix(werkRows, chVrVip) = Round(S, 2)
      End If
    End If
If isMO = "o" Then
   If tbCeh!StatO = "�����" Then
     Grid.TextMatrix(werkRows, chStatus) = tbCeh!StatO '�������
   Else
     Grid.TextMatrix(werkRows, chStatus) = "" '�������
   End If
ElseIf (tbCeh!StatusId = 1 Or tbCeh!StatusId = 8) And Not IsNumeric(tbCeh!Stat) Then
    If Not IsNull(tbCeh!Stat) Then Grid.TextMatrix(werkRows, chStatus) = tbCeh!Stat
ElseIf tbCeh!StatusId = 2 Then ' ������
    str1 = "�": GoTo AA
ElseIf tbCeh!StatusId = 3 Or tbCeh!StatusId = 9 Then  ' ��������
    str1 = "�"
AA: Grid.col = chStatus
    Grid.CellForeColor = color
    Grid.TextMatrix(werkRows, chStatus) = str1 & " �� " & Format(tbCeh!DateRS, "dd.mm.yy")
Else
    Grid.TextMatrix(werkRows, chStatus) = Status(tbCeh!StatusId)
End If
MN:
 For I = 1 To UBound(tmpL) '�������� ������ � ����������� ����������
    If tmpL(I) = gNzak Then
        Grid.col = chIzdelia
        Grid.row = werkRows
        Grid.CellForeColor = 200
        Exit For
    End If
 Next I
Grid.TextMatrix(werkRows, 0) = marker
Grid.TextMatrix(werkRows, chNomZak) = gNzak & str
If str <> "" Then colorGridRow Grid, werkRows, &HCCCCCC '��������� ��
Grid.TextMatrix(werkRows, chM) = tbCeh!Manag
Grid.TextMatrix(werkRows, chFirma) = tbCeh!Name
If idWerk = 1 Then
    Grid.TextMatrix(werkRows, chRemark) = tbCeh!Remark
Else
    Grid.TextMatrix(werkRows, chLogo) = tbCeh!Logo
    Grid.TextMatrix(werkRows, chIzdelia) = tbCeh!Product
End If
Grid.TextMatrix(werkRows, chWerkId) = tbCeh!WerkId

If tbCeh!StatusId = 5 Then ' �������
    Grid.TextMatrix(werkRows, chProblem) = Problems(tbCeh!ProblemId)
End If

End Sub

Sub LoadDateKey(val As Variant, myErr As String)
Dim I As Integer

If Not IsNull(val) Then
  If IsDate(val) Then
    Grid.TextMatrix(werkRows, chDataVid) = Format(val, "dd.mm.yy")
    I = DateDiff("d", curDate, val) + 1 '�����
    Grid.TextMatrix(werkRows, chKey) = I
'    If i = stDay Then
'        Grid.col = chDataVid
'        Grid.CellForeColor = &H8800&
'        Grid.CellFontBold = True
'    End If
    Exit Sub
  End If
End If
msgOfZakaz myErr, , tbCeh!Manag
Grid.TextMatrix(werkRows, chDataRes) = "??"
Grid.TextMatrix(werkRows, chKey) = 0
End Sub

Private Sub Form_Load()

Dim I As Integer


oldHeight = Me.Height
oldWidth = Me.Width

cmNaklad.Visible = True

If Not (dostup = "a" Or dostup = "m" Or dostup = "" Or dostup = "b") Then
    cmZagruz.Visible = True
    Orders.managLoad "fromCeh" ' �������� Manag()
End If
If dostup = "" Then cmNaklad.Visible = False


Screen.MousePointer = flexHourglass

For I = begWerkProblemId To lenProblem
    lbProblem.AddItem Problems(I)
Next I

Dim gridHeaderStr As String
gridHeaderStr = "    |<� ������|^�|������|������ |>��.���|>%��|��������|<���� ������|<��.���|<���� �������|<��������" _

If idWerk = 1 Then
    gridHeaderStr = gridHeaderStr _
        & "||<����������"
Else
    gridHeaderStr = gridHeaderStr _
        & "|<����|<�������"
End If

gridHeaderStr = gridHeaderStr _
    & "|����|equipid|werkid"
    
Grid.FormatString = gridHeaderStr

Grid.ColWidth(chM) = 270
Grid.ColWidth(chVrVip) = 388
Grid.ColWidth(chEquip) = 570
Grid.ColWidth(chStatus) = 870
Grid.ColWidth(chProcVip) = 420
Grid.ColWidth(chProblem) = 900
Grid.ColWidth(chDataRes) = 735
Grid.ColWidth(chVrVid) = 330
Grid.ColWidth(chDataVid) = 735
Grid.ColWidth(chFirma) = 2000
Grid.ColWidth(chKey) = 0 ' ��� ���������� �� ����
Grid.ColWidth(chEquipId) = 0
Grid.ColWidth(0) = 0
Grid.ColWidth(chNomZak) = 1000
Grid.ColWidth(chWerkId) = 0

If idWerk = 1 Then
    Grid.ColWidth(chLogo) = 0
    Grid.ColWidth(chRemark) = 3650
Else
    Grid.ColWidth(chLogo) = 1200
    Grid.ColWidth(chIzdelia) = 2450
End If



Timer1.Interval = 500
Timer1.Enabled = True '����� werkBegin

End Sub

Private Sub Form_Resize()
Dim H As Integer, W As Integer, I As Integer

If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next
lbHide
H = Me.Height - oldHeight
oldHeight = Me.Height
W = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + H
Grid.Width = Grid.Width + W
cmRefresh.Top = cmRefresh.Top + H
cmExAll.Top = cmExAll.Top + H
cmExAll.Left = cmExAll.Left + W
cmZagruz.Top = cmZagruz.Top + H
cmZagruz.Left = cmZagruz.Left + W
cmPrint.Left = cmPrint.Left + W
cmNaklad.Top = cmNaklad.Top + H
cmNakladZakaz.Top = cmNakladZakaz.Top + H

lbEquips.Top = lbEquips.Top + H
cbEquips.Top = cbEquips.Top + H
Dim RightLine As Integer

'For I = 0 To cmEquip.UBound
'    cmEquip(I).Top = cmEquip(I).Top + H
'    If RightLine < cmEquip(I).Left + cmEquip(I).Width Then
'        RightLine = cmEquip(I).Left + cmEquip(I).Width
'    End If
'Next I


End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (dostup = "a" Or dostup = "m" Or dostup = "" Or dostup = "b") Then
    exitAll '��� �����
End If
isWerkOrders = False
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

If Grid.TextMatrix(mousRow, chWerkId) = 1 And mousCol = chRemark And cmNakladZakaz.Enabled Then
    callNaklad
End If

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
If (mousCol = chStatus And marker <> "") Then
    Grid.CellBackColor = &H88FF88
Else
    Grid.CellBackColor = vbYellow
End If


cmNakladZakaz.Enabled = False
Dim I As Integer
If IsNumeric(gNzak) Then
    For I = 1 To UBound(tmpL) '�������� ������ � ����������� ����������
        If tmpL(I) = gNzak Then
            cmNakladZakaz.Enabled = True
            Exit For
        End If
    Next I
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

sql = "SELECT StatM From OrdersInCeh WHERE (((numOrder)=" & gNzak & "));"
If Not byErrSqlGetValues("##312", sql, tmpStr) Then Exit Sub
If tmpStr = "���������" Then
    msgZakazDeleted "����� ��� ���������"
    GoTo EN1
ElseIf lbMaket.Text = "�����" Then
    I = ValueToTableField("W##37", "'�����'", "OrdersInCeh", "StatM")
Else
    I = ValueToTableField("W##37", "'� ������'", "OrdersInCeh", "StatM")
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
Dim J As Integer, str As String, old As String, V As Variant
Dim proc As String, Status As String
'sChr As String, dChr As String,
If noClick Then Exit Sub
old = Grid.TextMatrix(mousRow, chStatus)
If lbObrazec.Text = "�����" And lbObrazec.Text <> old Then
    proc = "100%": Status = "'�����'"
ElseIf lbObrazec.Text <> old Then '              �������
    proc = "0%": Status = "'� ������'"
Else
    lbHide
    Exit Sub
End If
lbObrazec.Visible = False

wrkDefault.BeginTrans
    
gEquipId = Grid.TextMatrix(mousRow, chEquipId)
V = makeProcReady(proc, gEquipId, "obraz")
If IsNull(V) Then ' ������� ���������
    msgZakazDeleted "������� ��� ���������"
ElseIf V Then
    If ValueToTableField("##54", Status, "OrdersEquip", "StatO", "byEquipId") = 0 Then
        wrkDefault.CommitTrans
        werkBegin
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

gEquipId = Grid.TextMatrix(mousRow, chEquipId)
I = ValueToTableField("W##41", "'� ������'", "OrdersEquip", "Stat", "byEquipId") '�.� ���� �������� Stat=�����, �� �� ������ �� ���������
If I = 0 Then
    If ValueToTableField("##41", "5", "Orders", "StatusId") <> 0 Then GoTo ER1

    str = lbProblem.ListIndex + begWerkProblemId
    If ValueToTableField("##41", str, "Orders", "ProblemId") = 0 Then
        wrkDefault.CommitTrans  ' ������������� ����������
        werkBegin
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


gEquipId = Grid.TextMatrix(mousRow, chEquipId)

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
    If Not predmetiIsClose("etap") Then '
        str = "�������� ����� "
        If QQ2(0) = 0 Then str = ""
        MsgBox "�� ����� ������ �������(��������) �� ��� �������� " & str & _
        "(��� ��������� �������� �� ������� �������)!", , _
        "������������ ������ ��� ������ � " & gNzak
'        Grid.SetFocus
    Else
        wrkDefault.BeginTrans
        ' ������� ������ ������ �� ����� ��� ����� ������������
        I = ValueToTableField("W##41", "'" & str & "'", "OrdersEquip", "Stat", "byEquipId")
        
        Dim minId, maxId  'as variant
        sql = "select max(Stat) as maxId, min(Stat) as minId " _
        & " FROM OrdersEquip oe" _
        & " WHERE oe.numorder = " & gNzak
        
        'Debug.Print sql
        
        byErrSqlGetValues "##39.2", sql, maxId, minId
        
        If minId = maxId And minId = "�����" Then
            If ValueToTableField("##39.1", "4", "OrdersEquip", "StatusEquipId") <> 0 Then GoTo ER1
            If ValueToTableField("##39.3", "4", "Orders", "StatusId") <> 0 Then GoTo ER1
            If ValueToTableField("##39.4", "0", "Orders", "ProblemId") <> 0 Then GoTo ER1
            '��� ��� �������, ����������� �����.����, ��������, ��� ��� �. � ����� ���-��
            If Not newEtap("xEtapByIzdelia") Then GoTo ER1
            If Not newEtap("xEtapByNomenk") Then GoTo ER1
        End If
        wrkDefault.CommitTrans
        werkBegin
    End If
ElseIf str = "25%" Or str = "50%" Or str = "75%" Or str = "100%" Then
    lbStatus.Visible = False
    wrkDefault.BeginTrans
    If makeProcReady(str, Grid.TextMatrix(mousRow, chEquipId)) Then '� � ��� ����� ��� ������� ����� �� ����
        If ValueToTableField("##39", "1", "Orders", "StatusId") <> 0 Then GoTo ER1 ' "� ������"
        str = "� ������"
        GoTo AA
    End If
    GoTo ER2
Else '  �����, "*" � "� ������"
    lbStatus.Visible = False
    wrkDefault.BeginTrans
    If makeProcReady("0%", Grid.TextMatrix(mousRow, chEquipId)) Then
        If ValueToTableField("##41", "'" & str & "'", "OrdersEquip", "Stat", "byEquipId") <> 0 Then GoTo ER1
        If ValueToTableField("##39", "1", "Orders", "StatusId") <> 0 Then GoTo ER1
AA:     If ValueToTableField("##39", "0", "Orders", "ProblemId") = 0 Then
            wrkDefault.CommitTrans
            werkBegin
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
Function makeProcReady(Stat As String, EquipId As Integer, Optional obraz As String = "") As Variant
Dim S As Single, T As Single, N As Single, virabotka As Single, str As String
Dim StatO As String

makeProcReady = False
If Stat = "25%" Then
    S = 0.75 ' �����������
    GoTo AA
ElseIf Stat = "50%" Then
    S = 0.5
    GoTo AA
ElseIf Stat = "75%" Then
    S = 0.25
    GoTo AA
ElseIf Stat = "100%" Then
    S = 0
    GoTo AA
Else
    S = 1
AA:
 
  If obraz <> "" Then
    obraz = "o"
    ''??TODO
    sql = "SELECT oe.workTimeMO, oe.StatO " _
    & " FROM OrdersEquip oe " _
    & " WHERE oe.numOrder = " & gNzak & " AND equipId = " & EquipId
    If Not byErrSqlGetValues("##386", sql, virabotka, StatO) Then Exit Function
    If S = 0 Then ' 100%
    Else
        virabotka = -virabotka
    End If
  Else
    sql = "SELECT oe.workTime, isnull(oe.Nevip, 1) as nevip " _
    & " FROM OrdersEquip oe " _
    & " WHERE oe.numOrder = " & gNzak & " AND equipId = " & EquipId
    If Not byErrSqlGetValues("##421", sql, T, N) Then Exit Function
    
    virabotka = Round((N - S) * T, 2)
  End If


'���-�� ����� ��������� � ������� � 75% �� 0%
    str = Format(curDate, "yy.mm.dd")
    
    sql = "call putWerkOrderReady(" & gNzak & ", '" & str & "', '" & obraz & "', " & virabotka & ", " & EquipId & ", " & S & ")"
  
    myExecute "##374", sql
    
    If obraz = "o" Then '          ��� �������
        If StatO = "���������" Then
            makeProcReady = Null
            Exit Function
        End If
    Else 'obraz = ""
        gEquipId = EquipId
        ValueToTableField "##41", "'� ������'", "OrdersEquip", "Stat", "byEquipId"
    End If
    
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

werkBegin (True)
gridIsLoad = True
Grid.col = 1
isWerkOrders = True
trigger = True

End Sub

Function newEtap(Table As String) As Boolean
newEtap = False
sql = "UPDATE " & Table & " SET prevQuant = eQuant WHERE numOrder =" & gNzak
If myExecute("##193", sql, 0) > 0 Then Exit Function
newEtap = True
End Function

