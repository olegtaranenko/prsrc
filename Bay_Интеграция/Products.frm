VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form sProducts 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������������ ���������"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   1725
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmExel2 
      Caption         =   "������ � Exel"
      Height          =   315
      Left            =   8640
      TabIndex        =   18
      Top             =   5940
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   2700
      Top             =   4440
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   9900
      TabIndex        =   16
      Text            =   "tbMobile"
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "������ � Exel"
      Height          =   315
      Left            =   2340
      TabIndex        =   14
      Top             =   5940
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame gridFrame 
      BackColor       =   &H00800000&
      BorderStyle     =   0  '���
      Height          =   2055
      Left            =   3180
      TabIndex        =   10
      Top             =   3420
      Visible         =   0   'False
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid Grid4 
         Height          =   1455
         Left            =   60
         TabIndex        =   11
         Top             =   300
         Visible         =   0   'False
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2566
         _Version        =   393216
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Label laGrid4 
         Alignment       =   2  '���������
         BackColor       =   &H00800000&
         Caption         =   "laGrid4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   60
         Width           =   7215
      End
      Begin VB.Label Label2 
         Alignment       =   2  '���������
         Caption         =   "���� ������� ���������, ������� ����. ���-�� ������� � ������� <Enter>, ����� - <ESC>.."
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   1740
         Width           =   7215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5595
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   9869
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  '���
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Label laGrid 
         Caption         =   "Label1"
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   -15
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmHide 
      Caption         =   "������ ���."
      Enabled         =   0   'False
      Height          =   315
      Left            =   5700
      TabIndex        =   7
      Top             =   5940
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox tbQuant 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Top             =   5940
      Width           =   735
   End
   Begin VB.CommandButton cmSel 
      Caption         =   "��������"
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   5940
      Width           =   915
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "�����"
      Height          =   315
      Left            =   11040
      TabIndex        =   2
      Top             =   5940
      Width           =   795
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   5580
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   9843
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   5595
      Left            =   7200
      TabIndex        =   17
      Top             =   240
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   9869
      _Version        =   393216
      AllowBigSelection=   0   'False
      HighLight       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label laGrid1 
      Height          =   195
      Left            =   2760
      TabIndex        =   19
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label laBegin 
      Caption         =   "Label2"
      Height          =   4395
      Left            =   2760
      TabIndex        =   6
      Top             =   900
      Width           =   3795
   End
   Begin VB.Label laQuant 
      Caption         =   "�������"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4860
      TabIndex        =   5
      Top             =   5985
      Width           =   675
   End
   Begin VB.Label laGrid2 
      Caption         =   "�������������� ������ ���������:"
      Height          =   195
      Left            =   7380
      TabIndex        =   1
      Top             =   30
      Width           =   3495
   End
   Begin VB.Menu mnContext 
      Caption         =   "�� ������� ���������"
      Visible         =   0   'False
      Begin VB.Menu mnDel 
         Caption         =   "�������"
      End
      Begin VB.Menu mnCancel 
         Caption         =   "��������"
      End
   End
End
Attribute VB_Name = "sProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'#Const OTLAD = True '����������� ������  �����(����� ���������������� � ���������� � Main)

Public isLoad As Boolean
Public Regim As String
Public mousCol2 As Long
Public mousRow2 As Long
Public mousCol3 As Long
Public mousRow3 As Long
Public zakazano As Single
Public FO As Single ' ��

Dim mousCol4 As Long, mousRow4 As Long
Dim msgBilo As Boolean
Dim grColor As Long
Dim flag As Integer
'Dim maxNumExt As Integer, minNumExt As Integer

Dim mousCol As Long, mousRow As Long
Dim mousCol5 As Long, mousRow5 As Long
Dim Node As Node
Dim quantity  As Long, quantity2 As Long, quantity3 As Long
Public quantity5 As Long
Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����
Dim tvVes As Single, gridVes As Single, grid2Ves As Single '���� ��������. ��������

Dim tbSeries As Recordset
Dim tbKlass As Recordset
Dim typeId As Integer
Dim beShift As Boolean
'Dim QP() As Single
'Dim VN() As String

'Grid4
Const frNomNom = 1
Const frNomName = 2
Const frEdIzm = 3
Const frOstat = 4

'Grid2
Const dnNomNom = 1
Const dnNomName = 2
Const dnEdIzm = 3
Const dnVesEd = 4
Const dnCenaEd = 5
Const dnQuant = 6
Const dnSumm = 7
Const dnVes = 8

'������������ �� ������ ��� �������(Grid)
Const nkNomer = 1
Const nkName = 2
Const nkEdIzm = 3
'Const nkQuant = 5
Const nkCurOstat = 4
Const nkDostup = 5

Private Sub cmExel_Click()
'    GridToExcel Grid, laGrid1.Caption
    GridToExcel Grid, laGrid.Caption

End Sub

Private Sub cmExel2_Click()
    GridToExcel Grid2, laGrid2.Caption
End Sub

Private Sub cmExit_Click()
'If Regim = "ostat" Then
    Unload Me
'ElseIf checkRowsQuant Then
'    Unload Me
'End If
    
End Sub

Private Sub cmHide_Click()
Dim i As Integer
If quantity = 0 Then Exit Sub
For i = Grid.row To Grid.RowSel
    Grid.RemoveItem Grid.row
    quantity = quantity - 1
Next i
Grid.SetFocus
Grid_EnterCell
End Sub

Private Sub cmSel_Click() '<��������>
'Dim befColor As Long, il As Long, nl As Long, n As Integer, str As String

'laQuant.Visible = False
If beNaklads() Then Exit Sub

dostupOstatkiToGrid

tbQuant.Enabled = True
laQuant.Enabled = True
  tbQuant.Text = 1
  tbQuant.SelLength = 1
  tbQuant.SetFocus
  cmSel.Enabled = False
End Sub

Private Sub Form_Load()
Dim str As String, i As Integer, delta As Single

'oldHeight = Me.Height
'oldWidth = Me.Width

controlEnable False
laQuant.Visible = False
laQuant.Caption = "����"

Frame.Visible = False
    
laBegin = "� �������������� �������� (������ Mouse) ������, ��� ���� " & _
"��������� �������, ��� ����� ������������ ��� ������������ ���� ������." & _
vbCrLf & "      �������� � ���� ������� ��������� ������� � ������� <��������>." & _
vbCrLf & vbCrLf & "��� ������������� ��������� ��� �������� ��� " & _
"������ �������."

loadKlass

noClick = False
msgBilo = False
Grid.FormatString = "|<�����|<��������|<��.���|�.�������|�.�������"
'Grid2.FormatString = "|<�����|<��������|<��.���������|���� �� ��.|���-��|�����"Grid.FormatString = "|<�����|<��������|<��.���|�.�������|�.�������"
Grid2.FormatString = "|<�����|<��������|<��.���������|���.��|���� �� ��.|" & _
                     "���-��|�����|���"
Grid.ColWidth(0) = 0
Grid.ColWidth(nkNomer) = 900
'Grid.ColWidth(nkName) = 2820 � Resize
Grid.ColWidth(nkEdIzm) = 630 'ostat
'Grid.ColWidth(nkQuant) = 0
Grid.ColWidth(nkCurOstat) = 0
Grid.ColWidth(nkDostup) = 0
cmExel.Visible = False

If Regim = "ostat" Then
    Me.Caption = "��������� ��������"
    cmExel.Visible = True
    cmHide.Visible = True
    Grid.ColWidth(nkName) = 3510
'    Grid.ColWidth(nkQuant) = 0
    Grid.ColWidth(nkCurOstat) = 810
    Grid.ColWidth(nkDostup) = 800
    cmSel.Visible = False
    tbQuant.Visible = False
    laQuant.Visible = False
    laGrid2.Visible = False
    Grid2.Visible = False
    Grid.Width = 7000 '6230
    Me.Width = Grid.Width + 2527
    Frame.Width = Grid.Width
    cmExit.Left = Me.Width - cmExit.Width - 200
    Grid2.Width = 0 '��� Resize
    GoTo EN1
ElseIf Regim = "" Or Regim = "closeZakaz" Then
    cmExel2.Visible = True
End If

gSeriaId = 0 '���������  ��� ���������� ������

Grid2.ColWidth(0) = 0
Grid2.ColWidth(dnNomNom) = 0 '900
'Grid2.ColWidth(dnNomName) =  � Resize
Grid2.ColWidth(dnEdIzm) = 435
Grid2.ColWidth(dnCenaEd) = 510
Grid2.ColWidth(dnQuant) = 600
Grid2.ColWidth(dnSumm) = 660
Grid2.ColWidth(dnVesEd) = 495
Grid2.ColWidth(dnVes) = 600

quantity2 = 0
loadPredmeti ' ���� �������� ������ �� ��������� ������
If quantity2 > 0 Then
    str = "��������������"
Else
    str = "������������"
End If
Me.Caption = str & " ��������� � ������ � " & numDoc
If Regim = "closeZakaz" Then
    Me.Caption = "�������� � ������ � " & numDoc
    laBegin.Caption = "��� �������� �����. �������������� ��������� ����������."
'    opNomenk.Enabled = False
'    opProduct.Enabled = False
    tv.Enabled = False
    cmSel.Enabled = False
    laGrid2.Enabled = False
'   Grid2.Enabled = False
    cmExel2.Visible = False
End If
EN1:
oldHeight = Me.Height
oldWidth = Me.Width
tvVes = tv.Width / (tv.Width + Grid.Width + Grid2.Width)
gridVes = Grid.Width / (tv.Width + Grid.Width + Grid2.Width)
grid2Ves = Grid2.Width / (tv.Width + Grid.Width + Grid2.Width)
isLoad = True
End Sub

Sub loadPredmeti() '

Dim s As Single, sum As Single, sumVes As Single, quant As Single

MousePointer = flexHourglass
Grid2.Visible = False
quantity2 = 0
clearGrid Grid2
sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.ed_Izmer2, " & _
"sGuideNomenk.Size, sGuideNomenk.cod, sGuideNomenk.perList, sDMCrez.quantity, " & _
"sDMCrez.intQuant, sGuideNomenk.VES " & _
"FROM sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom " & _
"Where (((sDMCrez.numDoc) = " & numDoc & ")) ORDER BY sGuideNomenk.nomNom;"

'MsgBox sql

Set tbNomenk = myOpenRecordSet("##118", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  sum = 0: sumVes = 0
  While Not tbNomenk.EOF
    quantity2 = quantity2 + 1
    Grid2.TextMatrix(quantity2, dnNomName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid2.TextMatrix(quantity2, dnNomNom) = tbNomenk!nomNom
    Grid2.TextMatrix(quantity2, dnEdIzm) = tbNomenk!ed_Izmer2
    Grid2.TextMatrix(quantity2, dnVesEd) = tbNomenk!VES

    Grid2.TextMatrix(quantity2, dnCenaEd) = tbNomenk!intQuant
    quant = Round(tbNomenk!quantity / tbNomenk!perList, 2)
    Grid2.TextMatrix(quantity2, dnQuant) = quant
    s = Round(tbNomenk!VES * quant, 3)
    Grid2.TextMatrix(quantity2, dnVes) = s
    sumVes = sumVes + s
    
    s = Round(quant * tbNomenk!intQuant, 2)
    Grid2.TextMatrix(quantity2, dnSumm) = s
    sum = sum + s
    Grid2.AddItem ""
    tbNomenk.MoveNext
  Wend
  'Grid2.RemoveItem quantity2 + 1
End If
tbNomenk.Close
EN1:
Grid2.Visible = True
If quantity2 > 0 Then
    Grid2.TextMatrix(quantity2 + 1, dnQuant) = "�����:"
    Grid2.row = quantity2 + 1: Grid2.col = dnSumm
    Grid2.Text = Round(sum, 2)
    Grid2.CellFontBold = True
    
    Grid2.col = dnVes
    Grid2.Text = Round(sumVes, 3)
    Grid2.CellFontBold = True
    
    Grid2.row = 1: Grid2.col = 1
    On Error Resume Next
    If Me.isLoad Then
        Grid2.SetFocus
    Else
        Grid2.TabIndex = 0
    End If
End If
MousePointer = flexDefault


End Sub


Sub loadSeria()
Dim key As String, pKey As String, K() As String, pK()  As String
Dim i As Integer, iErr As Integer
bilo = False
sql = "SELECT sGuideSeries.*  From sGuideSeries ORDER BY sGuideSeries.seriaId;"
Set tbSeries = myOpenRecordSet("##110", sql, dbOpenForwardOnly)
If tbSeries Is Nothing Then myBase.Close: End
If Not tbSeries.BOF Then
 'Dim i As Integer
 'i = tbSeries.Fields("seriaName").Size
 tv.Nodes.Clear
 Set Node = tv.Nodes.Add(, , "k0", "���������� �� ������")
 
 ReDim K(0): ReDim pK(0): ReDim NN(0): iErr = 0
 While Not tbSeries.EOF
    If tbSeries!seriaId = 0 Then GoTo NXT1
    key = "k" & tbSeries!seriaId
    pKey = "k" & tbSeries!parentSeriaId
    On Error GoTo ERR1 ' ��������� ������ ������
    Set Node = tv.Nodes.Add(pKey, tvwChild, key, tbSeries!seriaName)
    On Error GoTo 0
NXT1:
    tbSeries.MoveNext
 Wend
End If
tbSeries.Close

While bilo ' ���������� ��� �������
  bilo = False
  For i = 1 To UBound(K())
    If K(i) <> "" Then
        On Error GoTo ERR2 ' ��������� ��� ������
        Set Node = tv.Nodes.Add(pK(i), tvwChild, K(i), NN(i))
        On Error GoTo 0
        K(i) = ""
        Node.Sorted = True
    End If
NXT:
  Next i
Wend
tv.Nodes.Item("k0").Expanded = True
Exit Sub
ERR1:
 iErr = iErr + 1: bilo = True
 ReDim Preserve K(iErr): ReDim Preserve pK(iErr): ReDim Preserve NN(iErr)
 K(iErr) = key: pK(iErr) = pKey: NN(iErr) = tbSeries!seriaName
 Resume Next

ERR2: bilo = True: Resume NXT

End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If Not isLoad Then Exit Sub
If Me.WindowState = vbMinimized Then Exit Sub
If Me.WindowState = vbMaximized And Me.Width > cDELLwidth Then '����� DELL
    Grid2.ColWidth(dnNomName) = 5220
    Grid.ColWidth(nkName) = 5670
Else
    Grid2.ColWidth(dnNomName) = 1230 '2340
    Grid.ColWidth(nkName) = 2820
End If
On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
tv.Height = tv.Height + h
tv.Width = tv.Width + w * tvVes

Grid.Left = Grid.Left + w * tvVes
laGrid1.Left = Grid.Left
laBegin.Left = Grid.Left
Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w * gridVes

Grid2.Left = Grid2.Left + w * (tvVes + gridVes)
laGrid2.Left = Grid2.Left
Grid2.Height = Grid2.Height + h
Grid2.Width = Grid2.Width + w * grid2Ves

cmSel.Top = cmSel.Top + h
cmSel.Left = cmSel.Left + w
tbQuant.Top = tbQuant.Top + h
tbQuant.Left = tbQuant.Left + w
laQuant.Top = laQuant.Top + h
laQuant.Left = laQuant.Left + w
cmExit.Top = cmExit.Top + h
cmExit.Left = cmExit.Left + w
cmExel2.Top = cmExel2.Top + h
cmExel2.Left = cmExel2.Left + w
cmExel.Top = cmExel.Top + h
cmHide.Top = cmHide.Top + h

End Sub

Private Sub Form_Unload(Cancel As Integer)
isLoad = False
If Regim = "" Then '�������� ������
    If beNaklads("noMsg") Then Exit Sub '�.�. � ���� ����� ������ �������� �� �����
    If quantity2 > 0 Then
        Orders.Grid.CellForeColor = 200
    Else
        Orders.Grid.CellForeColor = vbBlack
    End If
End If
End Sub

Sub dostupOstatkiToGrid(Optional reg As String)
Dim s As Single, sum As Single, rr As Long, il As Long

Me.MousePointer = flexHourglass
'If numExt = 254 Then
laGrid4.Caption = "��������� �������"

clearGrid Grid4
Grid4.FormatString = "|<�����|<��������|<��.���������|O������"
Grid4.ColWidth(0) = 0
Grid4.ColWidth(frNomNom) = 870
Grid4.ColWidth(frNomName) = 4485
Grid4.ColWidth(frEdIzm) = 645
Grid4.ColWidth(frOstat) = 885

nomencOstatkiToGrid 1

Grid4.Visible = True
EN1:
Me.MousePointer = flexDefault
gridFrame.Visible = True
gridFrame.ZOrder

End Sub

Public Function nomencOstatkiToGrid(row As Long) As Single
Dim s As Single, str As String, z As Single, str2 As String

'�.�������
sql = "SELECT nomName, Ed_Izmer2, perList From sGuideNomenk " & _
"WHERE (((nomNom)='" & gNomNom & "'));"
'MsgBox sql
byErrSqlGetValues "##144", sql, str, str2, tmpSng
If row > 0 Then
    Grid4.TextMatrix(row, frNomNom) = gNomNom
    Grid4.TextMatrix(row, frNomName) = str
    Grid4.TextMatrix(row, frEdIzm) = str2
End If


'AA: ��������� ��������� �������
FO = PrihodRashod("+", -1001) - PrihodRashod("-", -1001)
    
sql = "SELECT Sum(quantity) AS Sum_quantity, " & _
"Sum(Sum_quant) AS Sum_Sum_quant From wCloseNomenk" & _
" WHERE (((nomNom)='" & gNomNom & "'));"
If Not byErrSqlGetValues("##145", sql, z, s) Then myBase.Close: End
nomencOstatkiToGrid = FO - (z - s) ' �����, ��� ���������

nomencOstatkiToGrid = nomencOstatkiToGrid / tmpSng

If row > 0 Then _
    Grid4.TextMatrix(row, frOstat) = Round(nomencOstatkiToGrid, 2)

End Function

Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If quantity = 0 Then Exit Sub
'� Grid ����������� ���������� �.�. �������� �����������
End Sub

Private Sub Grid_DblClick()
Dim il As Long, curRow As Long

grColor = Grid.CellBackColor
If grColor = &H88FF88 Then
    If MsgBox("���� �� ������ ����������� ������ ���� �������, ��� " & _
    "������� ���� ��������������� ��� ������������, ������� <��>.", _
    vbYesNo Or vbDefaultButton2, "����������, ��� ������������? '" & _
    gNomNom & "' ?") = vbYes Then
        Report.Regim = "whoRezerved"
        Report.Show vbModal
    End If
Else
    cmSel_Click
End If
End Sub

Private Sub Grid_EnterCell()
Dim f As String, d As Single

If quantity = 0 Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col


'If opProduct.value Then
'    gProductId = Grid.TextMatrix(mousRow, gpId)
'    gProduct = Grid.TextMatrix(mousRow, gpName)
'Else
gNomNom = Grid.TextMatrix(mousRow, nkNomer)

'End If
 
Grid.CellBackColor = vbYellow
If mousCol = nkDostup Then
    f = Grid.TextMatrix(mousRow, nkCurOstat)
    d = Grid.TextMatrix(mousRow, nkDostup)
    If d < f Then Grid.CellBackColor = &H88FF88
End If

End Sub

Private Sub Grid_GotFocus()
'If opProduct.value Then rightORleft "l"
'rightORleft "l"
cmHide.Enabled = True
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid_DblClick
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyEscape Then Grid.CellBackColor = Grid.BackColor
If KeyCode = vbKeyEscape Then Grid_EnterCell
End Sub

Private Sub Grid_LeaveCell()
If Grid.col <> 0 Then Grid.CellBackColor = Grid.BackColor
End Sub


Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
End Sub

Private Sub Grid2_Click()
mousCol2 = Grid2.MouseCol
mousRow2 = Grid2.MouseRow

If mousRow2 = 0 Then
    Grid2.CellBackColor = Grid.BackColor
'    SortCol Grid2, mousCol
    trigger = Not trigger
    Grid2.Sort = 9
    
    Grid2.row = 1    ' ������ ����� ����� ���������
Else
    If quantity2 = 0 Then Exit Sub
End If
'Grid_EnterCell

End Sub


Private Sub Grid2_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    If Row1 = Grid2.Rows - 1 Then
        Cmp = 1: Exit Sub
    End If
    If Row2 = Grid2.Rows - 1 Then
        Cmp = -1: Exit Sub
    End If
    If Grid2.TextMatrix(Row1, mousCol2) < Grid2.TextMatrix(Row2, mousCol2) Then
        Cmp = -1
    ElseIf Grid2.TextMatrix(Row1, mousCol2) > Grid2.TextMatrix(Row2, mousCol2) Then
        Cmp = 1
    Else
        Cmp = 0
    End If
    If (trigger) Then Cmp = -Cmp
End Sub

Private Sub Grid2_DblClick()
If mousRow2 = 0 Then Exit Sub
If Grid2.CellBackColor = &H88FF88 Then textBoxInGridCell tbMobile, Grid2

End Sub

Private Sub Grid2_EnterCell()

mousRow2 = Grid2.row
mousCol2 = Grid2.col

If mousCol2 = dnSumm Or mousCol2 = dnCenaEd Then
    Grid2.CellBackColor = &H88FF88
Else
    Grid2.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid2_DblClick

End Sub

Private Sub Grid2_LeaveCell()
If Grid2.col <> 0 Then Grid2.CellBackColor = Grid2.BackColor

End Sub

Private Sub Grid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid2.MouseRow = 0 Then
    If Shift = 2 Then MsgBox "ColWidth = " & Grid2.ColWidth(Grid2.MouseCol)
ElseIf Button = 2 And quantity2 <> 0 Then
    Grid2.row = Grid2.MouseRow
    Grid2.col = dnNomNom
    gNomNom = Grid2.Text
    Grid2.SetFocus
    Grid2.CellBackColor = vbButtonFace
    Me.PopupMenu mnContext
'    noClick = True
End If
End Sub


Private Sub Grid4_GotFocus()
tbQuant.SetFocus
End Sub

Private Sub Grid4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid4.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid4.ColWidth(Grid4.MouseCol)

End Sub

Sub nomenkToNNQQ(pQuant As Single, eQuant As Single, prQuant As Single)
Dim j As Integer, leng As Integer

leng = UBound(NN)

    For j = 1 To leng
        If NN(j) = tbNomenk!nomNom Then
            QQ(j) = QQ(j) + pQuant * tbNomenk!quantity
            If eQuant > 0 Then _
                QQ2(j) = QQ2(j) + eQuant * tbNomenk!quantity
            If prQuant > 0 Then _
                QQ3(j) = QQ3(j) + prQuant * tbNomenk!quantity
            Exit Sub
        End If
    Next j
    leng = leng + 1
    ReDim Preserve NN(leng): NN(leng) = tbNomenk!nomNom
    ReDim Preserve QQ(leng): QQ(leng) = pQuant * tbNomenk!quantity
    ReDim Preserve QQ2(leng): QQ2(leng) = eQuant * tbNomenk!quantity
'    QQ2(leng) = 0: If eQuant > 0 Then QQ2(leng) = eQuant * tbNomenk!quantity
    ReDim Preserve QQ3(leng): QQ3(leng) = prQuant * tbNomenk!quantity
    

End Sub



Private Sub Grid5_Click()

End Sub

Private Sub mnCancel2_Click()

End Sub

Private Sub mnDel_Click()
Dim pQuant As Single, i As Integer ', str  As String

If beNaklads() Then Exit Sub

If MsgBox("�� ������ ������� ������� '" & gNomNom & _
"'", vbYesNo Or vbDefaultButton2, "����������� ��������") = vbNo Then Exit Sub

sql = "DELETE From sDMCrez WHERE (((numDoc)=" & gNzak & _
") AND ((nomNom)='" & gNomNom & "'));"
'MsgBox sql

If myExecute("##348", sql) = 0 Then
    loadPredmeti ' ���-��
    Orders.Grid.TextMatrix(Orders.Grid.row, orZakazano) = getOrdered(gNzak)
End If
'Grid2.SetFocus ' ����� �� �����������

End Sub

'
Sub controlEnable(EN As Boolean)
If Not EN Then ' ������ �����
    Grid.Visible = False
End If
cmSel.Enabled = EN
End Sub


Function nomenkToDMC(delta As Single, Optional noOpen As String = "") As Boolean
Dim s As Single

nomenkToDMC = False

If noOpen = "" Then
    
    If Not lockSklad Then Exit Function
    
'    s = nomencOstatkiToGrid(1) - delta ' ������������ ��������� �������
'AA: If s < -0.005 Then '� 2� ������
'        If MsgBox("������� ������ '" & gNomNom & "' �� ������������� '" & _
'        sDocs.getGridColSour() & "'�������� (" & s & "), ����������?", _
'        vbOKCancel Or vbDefaultButton2, "�����������") = vbCancel Then GoTo EN1
'    End If
    
    Set tbDMC = myOpenRecordSet("##123", "sDMC", dbOpenTable)
    If tbDMC Is Nothing Then GoTo EN1
    tbDMC.Index = "NomDoc"
End If

tbDMC.Seek "=", numDoc, numExt, gNomNom
If tbDMC.NoMatch Then
    tbDMC.AddNew
    tbDMC!numDoc = numDoc
    tbDMC!numExt = 254
    tbDMC!nomNom = gNomNom
    tbDMC!quant = Round(delta, 2)
Else
    tbDMC.Edit
    tbDMC!quant = Round(tbDMC!quant + delta, 2)
End If
tbDMC.Update
    
If noOpen = "" Then tbDMC.Close

'������������ �������(��� ������������ �� ������������)
If Not ostatCorr(delta) Then MsgBox "�� ������ ��������� ��������. " & _
     "�������� ��������������.", , "Error 83" '##83
nomenkToDMC = True

EN1:
If noOpen = "" Then lockSklad "un"
End Function

Private Sub tbMobile2_Change()

End Sub

Private Sub tbMobile5_Change()

End Sub

Private Sub mnDel2_Click()

End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim c As Single, s As Single, str As String

If KeyCode = vbKeyReturn Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If mousCol2 = dnSumm Then
        s = tbMobile.Text
        c = s / CSng(Grid2.TextMatrix(mousRow2, dnQuant)) '�� ���������
        GoTo BB
    Else 'dnCenaEd
        c = tbMobile.Text
        s = c * CSng(Grid2.TextMatrix(mousRow2, dnQuant))
BB:     sql = "UPDATE sDMCrez SET intQuant = " & c & " WHERE (((numDoc)=" & _
        gNzak & ") AND ((nomNom)='" & Grid2.TextMatrix(mousRow2, dnNomNom) & "'));"
        If myExecute("##205", sql) = 0 Then
            Grid2.TextMatrix(mousRow2, dnCenaEd) = Round(c, 2)
            Grid2.TextMatrix(mousRow2, dnSumm) = Round(s, 2)
            s = getOrdered(gNzak)
            Orders.Grid.TextMatrix(Orders.Grid.row, orZakazano) = s
            Orders.Grid.TextMatrix(Orders.Grid.row, orOtgrugeno) = getShipped(gNzak)
'            tmpVar = saveOrdered
'            If IsNumeric(tmpVar) Then
                Grid2.TextMatrix(Grid2.Rows - 1, dnSumm) = s 'tmpVar
'                Otgruz.saveBayShipped '���� ������ � �� ��������
'            End If
        End If
    End If
    
    lbHide
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Sub lbHide()
tbMobile.Visible = False
Grid2.Enabled = True
Grid2.SetFocus
Grid2_EnterCell

End Sub

Function deficitAndNoIgnore(delta As Single) As Boolean
Dim s As Single, il As Long


deficitAndNoIgnore = False
s = nomencOstatkiToGrid(il) - delta ' ������������ ��������� �������
If s < -0.005 Then
    If MsgBox("������� ������ '" & gNomNom & "' � ��������� ��������" & " �������� (" & _
    s & "), ����������?", vbOKCancel Or vbDefaultButton2, "�����������") _
    = vbOK Then Exit Function
    deficitAndNoIgnore = True
End If
End Function

Private Sub tbQuant_KeyDown(KeyCode As Integer, Shift As Integer)
Dim s As Single ', str As String
'Dim i As Integer, NN2() As String
Dim per As Single ', delta As Single

If KeyCode = vbKeyReturn Then

  If Not isNumericTbox(tbQuant, 0.01) Then Exit Sub

  If Not lockSklad Then Exit Sub
  
  sql = "SELECT perList From sGuideNomenk WHERE (((nomNom)='" & gNomNom & "'));"
  If Not byErrSqlGetValues("W##346", sql, per) Then GoTo AA
    
  s = nomencOstatkiToGrid(1) - tbQuant.Text ' ������������ ��������� �������
  If s < -0.005 Then '� 2� ������
    If MsgBox("������� ������ '" & gNomNom & "' � ��������� �������� " & _
    "�������� (" & s & "), ����������?", vbOKCancel Or vbDefaultButton2, _
    "�����������") = vbCancel Then GoTo AA
  End If
        
  per = per * tbQuant.Text
  If Not nomenkToDMCrez(per) Then
AA: lockSklad "un"
    Grid.SetFocus
    GoTo ESC
  End If
    
  lockSklad "un"

  loadPredmeti ' ���-��
  
  Grid.SetFocus
  GoTo ES2

ElseIf KeyCode = vbKeyEscape Then
ESC: tbQuant.Text = ""
    Grid.SetFocus
ES2: gridFrame.Visible = False
    tbQuant.Enabled = False
    laQuant.Enabled = False
    cmSel.Enabled = True
End If

End Sub
'��� delta < 0 - ����. ��������
Function nomenkToDMCrez(delta As Single, Optional mov As String = "") As Boolean
Dim s As Single, i As Integer

nomenkToDMCrez = False

' ���������� ������������ �� �����������
strWhere = " WHERE (((numDoc)=" & numDoc & ") AND ((nomNom)='" & gNomNom & "'));"

If mov = "" Then mov = "rez"

sql = "SELECT quantity FROM sDMC" & mov & strWhere
If Not byErrSqlGetValues("W##423", sql, s) Then Exit Function

If s = 0 Then ' ����� ���-�� ���
  If delta >= 0.01 Then
    sql = "INSERT INTO sDMC" & mov & " ( numDoc, nomNom, quantity ) " & _
    "SELECT " & numDoc & ", '" & gNomNom & "', " & delta & ";"
    If myExecute("##195", sql) <> 0 Then Exit Function
  End If
Else
    delta = s + delta

    If Round(delta, 2) > 0 Then
        sql = "UPDATE sDMC" & mov & " SET quantity = " & delta & strWhere
        If myExecute("##152", sql) <> 0 Then Exit Function
    Else
        sql = "DELETE FROM sDMC" & mov & strWhere
        If myExecute("##424", sql) <> 0 Then Exit Function
    End If
End If
        
'EN1:    If mov = "" Or mov = "mov" Then tbDMC.Close
nomenkToDMCrez = True
End Function

Private Sub Timer1_Timer()
'biloG3Enter_Cell = False
Timer1.Enabled = False
End Sub

Private Sub tv_AfterLabelEdit(Cancel As Integer, NewString As String)
' If Not flseriaAdd Then
'ValueToTableField "##115", "'" & NewString & "'", "sProducts", "seriaName", "bySeriaId"
gSeriaId = Mid$(tv.SelectedItem.key, 2)
ValueToTableField "##115", "'" & NewString & "'", "sGuideSeries", "seriaName", "bySeriaId"
End Sub

Sub loadKlass()
Dim key As String, pKey As String, K() As String, pK()  As String
Dim i As Integer, iErr As Integer
bilo = False
sql = "SELECT sGuideKlass.*  From sGuideKlass ORDER BY sGuideKlass.parentKlassId;"
Set tbKlass = myOpenRecordSet("##102", sql, dbOpenForwardOnly)
If tbKlass Is Nothing Then myBase.Close: End
If Not tbKlass.BOF Then
 tv.Nodes.Clear
 Set Node = tv.Nodes.Add(, , "k0", "�������������")
 Node.Sorted = True
 Set Node = tv.Nodes.Add("k0", tvwChild, "all", "              ")

 ReDim K(0): ReDim pK(0): ReDim NN(0): iErr = 0
 While Not tbKlass.EOF
    If tbKlass!klassId = 0 Then GoTo NXT1
    key = "k" & tbKlass!klassId
    pKey = "k" & tbKlass!parentKlassId
    On Error GoTo ERR1 ' ��������� ������ ������
    Set Node = tv.Nodes.Add(pKey, tvwChild, key, tbKlass!klassName)
    On Error GoTo 0
    Node.Sorted = True
NXT1:
    tbKlass.MoveNext
 Wend
  tv.Nodes.Item("all").Text = "���� ��������"
End If
tbKlass.Close

While bilo ' ���������� ��� �������
  bilo = False
  For i = 1 To UBound(K())
    If K(i) <> "" Then
        On Error GoTo ERR2 ' ��������� ��� ������
        Set Node = tv.Nodes.Add(pK(i), tvwChild, K(i), NN(i))
        On Error GoTo 0
        K(i) = ""
        Node.Sorted = True
    End If
NXT:
  Next i
Wend
tv.Nodes.Item("k0").Expanded = True
Exit Sub
ERR1:
 iErr = iErr + 1: bilo = True
 ReDim Preserve K(iErr): ReDim Preserve pK(iErr): ReDim Preserve NN(iErr)
 K(iErr) = key: pK(iErr) = pKey: NN(iErr) = tbKlass!klassName
 Resume Next

ERR2: bilo = True: Resume NXT

End Sub

Sub loadKlassNomenk()
Dim il As Long, r As Single, strWhere As String
Dim beg As Single, prih As Single, rash As Single, oldNow As Single



If tv.SelectedItem.key = "all" Then
    strWhere = ""
    quantity = 0
Else
  strWhere = "WHERE (((klassId)=" & gKlassId & "))"
End If

laGrid1.Caption = "������������ �� ������ '" & tv.SelectedItem.Text & "'"

Me.MousePointer = flexHourglass

If beShift Then
    Grid.AddItem ""
Else
    quantity = 0
    Grid.Visible = False
    clearGrid Grid
End If


sql = "SELECT nomNom, nomName, Size, cod, ed_Izmer, ed_Izmer2, nowOstatki " & _
"From sGuideNomenk " & strWhere & ";"
'"WHERE (((sGuideNomenk.klassId)=" & gKlassId & "));"

Set tbNomenk = myOpenRecordSet("##103", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
 tbNomenk.MoveFirst
 While Not tbNomenk.EOF
'    beg = tbNomenk!begOstatki
    gNomNom = tbNomenk!nomNom
    quantity = quantity + 1
    Grid.TextMatrix(quantity, nkNomer) = gNomNom
    Grid.TextMatrix(quantity, nkName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid.TextMatrix(quantity, nkEdIzm) = tbNomenk!ed_Izmer2
    
    If Regim = "ostat" Then
        '��������� ������� ��� ������� ����� (nomencOstatkiToGrid ������ perList � tmpSng):
        Grid.TextMatrix(quantity, nkDostup) = Round(nomencOstatkiToGrid(-1) - 0.4999, 0)
        Grid.TextMatrix(quantity, nkCurOstat) = Round(FO / tmpSng, 2) 'FO �� nomencOstatkiToGrid
    End If
    Grid.AddItem ""
    tbNomenk.MoveNext
 Wend
 If quantity > 0 Then Grid.RemoveItem quantity + 1
End If
tbNomenk.Close

Grid.Visible = True
EN1:
Me.MousePointer = flexDefault
End Sub

    
Private Sub tv_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, str As String
If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
'    tv_NodeClick tv.SelectedItem
End If
End Sub


Private Sub tv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
beShift = False
If Shift = 2 Then beShift = True

End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)

If tv.SelectedItem.key = "k0" Then
    controlEnable False
    quantity = 0
    laGrid.Caption = ""
    Exit Sub
End If

tbQuant.Enabled = False
laQuant.Enabled = False
controlEnable True
gKlassId = Mid$(tv.SelectedItem.key, 2)
loadKlassNomenk
'    Grid.Visible = True
Grid_EnterCell
On Error Resume Next
Grid.SetFocus

'biloG3Enter_Cell = False

End Sub


