VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
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
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbInside 
      Height          =   315
      Left            =   6180
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmExel2 
      Caption         =   "������ � Exel"
      Height          =   315
      Left            =   8640
      TabIndex        =   25
      Top             =   5940
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   5940
      Top             =   600
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   7620
      TabIndex        =   23
      Text            =   "tbMobile"
      Top             =   1380
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   315
      Left            =   1500
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "������ � Exel"
      Height          =   315
      Left            =   2340
      TabIndex        =   18
      Top             =   5940
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame gridFrame 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   3180
      TabIndex        =   14
      Top             =   3420
      Visible         =   0   'False
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid Grid4 
         Height          =   1455
         Left            =   60
         TabIndex        =   15
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
         Alignment       =   2  'Center
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
         TabIndex        =   17
         Top             =   60
         Width           =   7215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "���� ������� ���������, ������� ����. ���-�� ������� � ������� <Enter>, ����� - <ESC>.."
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   1740
         Width           =   7215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grid"
      Height          =   315
      Left            =   900
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grid3"
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2835
      Left            =   2400
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5001
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   2475
      Left            =   2400
      TabIndex        =   11
      Top             =   300
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4366
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmHide 
      Caption         =   "������ ���."
      Enabled         =   0   'False
      Height          =   315
      Left            =   5700
      TabIndex        =   10
      Top             =   5940
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.OptionButton opNomenk 
      Caption         =   "����� ������������"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   480
      Width           =   2175
   End
   Begin VB.OptionButton opProduct 
      Caption         =   "����� ������� �������"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   180
      Width           =   2235
   End
   Begin VB.TextBox tbQuant 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   5940
      Width           =   735
   End
   Begin VB.CommandButton cmSel 
      Caption         =   "��������"
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   5940
      Width           =   915
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "�����"
      Height          =   315
      Left            =   11040
      TabIndex        =   3
      Top             =   5940
      Width           =   795
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   4980
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   8784
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2895
      Left            =   7200
      TabIndex        =   24
      Top             =   2940
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   5106
      _Version        =   393216
      AllowBigSelection=   0   'False
      HighLight       =   0
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid5 
      Height          =   2415
      Left            =   7200
      TabIndex        =   20
      Top             =   300
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   4260
      _Version        =   393216
      AllowBigSelection=   0   'False
      MergeCells      =   2
      AllowUserResizing=   1
   End
   Begin VB.Label laGrid1 
      Caption         =   "laGrid1"
      Height          =   195
      Left            =   2400
      TabIndex        =   1
      Top             =   60
      Width           =   3735
   End
   Begin VB.Label laGrid 
      Caption         =   "laGrid"
      Height          =   195
      Left            =   2400
      TabIndex        =   28
      Top             =   2820
      Width           =   3795
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   540
      TabIndex        =   27
      Top             =   2820
      Width           =   3675
   End
   Begin VB.Label laGrid5 
      Caption         =   "������ ���������:"
      Height          =   195
      Left            =   7800
      TabIndex        =   21
      Top             =   45
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.Label laBegin 
      Caption         =   "Label2"
      Height          =   4395
      Left            =   2760
      TabIndex        =   7
      Top             =   900
      Width           =   3795
   End
   Begin VB.Label laQuant 
      Caption         =   "�������"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4860
      TabIndex        =   6
      Top             =   5985
      Width           =   675
   End
   Begin VB.Label laGrid2 
      Caption         =   "�������������� ������ ���������:"
      Height          =   195
      Left            =   7320
      TabIndex        =   2
      Top             =   2730
      Width           =   3495
   End
   Begin VB.Menu mnContext 
      Caption         =   "�� ������� ���������"
      Visible         =   0   'False
      Begin VB.Menu mnDel 
         Caption         =   "�������"
      End
      Begin VB.Menu mnOnfly 
         Caption         =   "������������� � �������"
      End
   End
   Begin VB.Menu mnContext2 
      Caption         =   "�� ��������� � �� ����"
      Visible         =   0   'False
      Begin VB.Menu mnDel2 
         Caption         =   "�������"
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

Private selectNomenkFlag As Boolean
Private isCtrlDown As Boolean

Dim mousCol4 As Long, mousRow4 As Long
Dim msgBilo As Boolean, biloG3Enter_Cell As Boolean

Const groupColor1 = &HBBFFBB ' ������ �� vbBottonFace
Const groupColor2 = &HBBBBFF '
Dim grColor As Long
Dim mousCol As Long, mousRow As Long
Dim mousCol5 As Long
Public mousRow5 As Long

Dim quantity  As Long, quantity2 As Long, quantity3 As Long
Public quantity5 As Long
Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����
Dim tvVes As Single, gridVes As Single, grid2Ves As Single '���� ��������. ��������

Dim tbKlass As Recordset
Dim typeId As Integer
Dim beShift As Boolean
'������ ������� ��� �����������(Grid3)
Const gpNN = 0
Const gpName = 1
Const gpSize = 2
Const gpDescript = 3
Const gpId = 4 ' �������

'������������ �� ������ ��� �������(Grid)
Const nkNomer = 1
Const nkName = 2
Const nkEdIzm = 3
Const nkQuant = 4
Const nkCurOstat = 5
Const nkDostup = 6

'Grid4
Const frNomNom = 1
Const frNomName = 2
Const frEdIzm = 3
Const frOstat = 4

Public convertToIzdelie As Boolean




'������������ � ��������� (Grid2) ��. Common
'������� Grid5'�� �����n

Private Sub cbInside_Click()
If isLoad And Grid.Visible Then loadKlassNomenk
End Sub

Private Sub cmExel_Click()
If opNomenk.value Then
    GridToExcel Grid, laGrid1.Caption
Else
    GridToExcel Grid, laGrid.Caption
End If

End Sub

Private Sub cmExel2_Click()
    GridToExcel Grid2, laGrid2.Caption
End Sub

Private Sub cmExit_Click()
    Unload Me
End Sub

Private Sub cmHide_Click()
Dim I As Integer
If quantity = 0 Then Exit Sub
For I = Grid.row To Grid.RowSel
    Grid.removeItem Grid.row
    quantity = quantity - 1
Next I
Grid.SetFocus
Grid_EnterCell
End Sub

Private Sub cmSel_Click() '<��������>
Dim befColor As Long, il As Long, nl As Long, n As Integer, str As String
Dim per As Single

sql = "SELECT perList From sGuideNomenk " & _
"WHERE (((nomNom)='" & gNomNom & "'));"
byErrSqlGetValues "##355", sql, per

If Regim = "fromDocs" And per = 1# Then   '���������� ���-�� �.�. ������ �� ������ -1001
    str = sDocs.lbInside.List(1) '-1002 - �������
    If sDocs.Grid.TextMatrix(sDocs.mousRow, dcSour) = str Or _
    sDocs.Grid.TextMatrix(sDocs.mousRow, dcDest) = str Then
        MsgBox "������� '" & gNomNom & "' �� ����� ���������� �� ������ '" & _
        str & "'.", , "��������������"
        Exit Sub
    End If
End If

If Not (Regim = "fromDocs" And sDocs.Regim = "fromCeh") Then _
    If beNaklads() Then Exit Sub

If Regim = "" Then
    If Otgruz.loadOutDates Then '�������� ��������
        MsgBox "�� ������ " & gNzak & " ��� �������� ��������.", , "�������������� ���������!"
        Exit Sub
    End If
End If

If opProduct.value Then
  '�������� ������������ ������� ***********
  befColor = 0: bilo = False
  Grid.col = nkQuant
  For il = Grid.Rows - 1 To 0 Step -1 '��� � ����� �. ������� ����� ���� ����� �.�. ������ ���� Row=0 c ������ ������
     Grid.row = il
     grColor = Grid.CellBackColor
     If grColor <> befColor Then
        If (befColor = groupColor1 Or befColor = groupColor2) And Not bilo Then
            MsgBox "����� ������ � ������� '���-��' ����������� �������, " & _
            "����� ������� ���� �������(������� ���� ��� <Enter>) ������ " & _
            "����, ������� ������ � �������.", , "������� '" & gProduct & _
            "' �� ��������������!"
            Exit Sub
        End If
        bilo = False
     End If
     If Grid.CellFontBold Then
        If grColor = groupColor1 Or grColor = groupColor2 Then bilo = True '�������
     End If
     befColor = grColor
  Next il '*********************************
   
  dostupOstatkiToGrid "multiN"
Else
  dostupOstatkiToGrid
End If
tbQuant.Enabled = True
laQuant.Enabled = True
  tbQuant.Text = 1
  tbQuant.SelLength = 1
  tbQuant.SetFocus
  cmSel.Enabled = False
End Sub

Private Sub Command1_Click()
gridOrGrid3Hide "grid"
End Sub

Private Sub Command2_Click()
gridOrGrid3Hide "grid3"
End Sub

Private Sub Command3_Click()
gridOrGrid3Hide
End Sub

Private Sub Form_Activate()
Dim I As Integer

End Sub

Private Sub Form_GotFocus()
Dim I As Integer

End Sub

Private Sub Form_Load()
Dim str As String, I As Integer, delta As Single
ReDim selectedItems(0)

If Regim = "fromDocs" And sDocs.Regim = "fromCeh" And skladId = -1002 Then _
        opProduct.Enabled = False
noClick = False
msgBilo = False
isLoad = False

Grid.FormatString = "|<�����|<��������|<��.���������|���-��|�.�������|�.�������"

Grid.ColWidth(0) = 0
Grid.ColWidth(nkNomer) = 0 '900
Grid.ColWidth(nkEdIzm) = 630 'ostat
Grid.ColWidth(nkCurOstat) = 0

Grid2.FormatString = "|<�����|<��������|<��.���������|���-��"
Grid2.ColWidth(0) = 0
Grid2.ColWidth(fnNomNom) = 0 '900
Grid2.ColWidth(fnEdIzm) = 435
Grid2.ColWidth(fnQuant) = 585

Grid3.FormatString = "|<�����|<������|<��������|id"

Grid5.FormatString = "|���|<���|<��������|<��.���������|���� �� ��." & _
"|���-��|�����|����� ��������|���-�� �� ���.�����"
Grid5.ColWidth(prId) = 0
Grid5.ColWidth(prName) = 1185
Grid5.ColWidth(prType) = 0
Grid5.ColWidth(prEdizm) = 420
Grid5.ColWidth(prCenaEd) = 495
Grid5.ColWidth(prEtap) = 660
Grid5.ColWidth(prEQuant) = 675

cmExel.Visible = False
If Regim = "ostatP" Then
    Regim = "ostat"
    opProduct.value = True: opProduct_Click
    GoTo AA
ElseIf Regim = "ostat" Then
    opNomenk.value = True: opNomenk_Click
AA: Me.Caption = "��������� ��������"
    cmExel.Visible = True
    Grid.ColWidth(nkName) = 3510 + 900 ' ����������� � ������� �� ���.��������
    Grid.ColWidth(nkQuant) = 0
    Grid.ColWidth(nkCurOstat) = 710
    Grid.ColWidth(nkDostup) = 700
    cmSel.Visible = False
    tbQuant.Visible = False
    laQuant.Visible = False
    laGrid2.Visible = False
    Grid2.Visible = False
    Grid5.Visible = False
    Grid.Width = 7700
    Me.Width = Grid.Width + 2527
    Grid3.Width = Grid.Width
    laGrid.Width = Grid.Width
    cmExit.Left = Me.Width - cmExit.Width - 200
    
    sql = "SELECT sourceId, sourceName From sGuideSource " & _
    "WHERE (((sourceId)<-1000)) ORDER BY sourceId DESC;"
    Set table = myOpenRecordSet("##359", sql, dbOpenDynaset)
    If table Is Nothing Then myBase.Close: End
    While Not table.EOF
        cbInside.AddItem table!SourceName
        table.MoveNext
    Wend
    table.Close
    cbInside.ListIndex = 0
    cbInside.Visible = True
    isLoad = True
    GoTo EN1 ' Exit Sub
ElseIf Regim = "fromDocs" Then
    laGrid5.Visible = False
    Grid5.Visible = False
    laGrid2.Top = laGrid5.Top
    delta = Grid2.Top - Grid5.Top
    Grid2.Top = Grid5.Top
    Grid2.Height = Grid2.Height + delta
ElseIf Regim = "" Or Regim = "closeZakaz" Then '�������� ������
BB: cmExel2.Visible = True
    laGrid5.Visible = True
End If

gSeriaId = 0 '���������  ��� ���������� ������

quantity2 = 0
loadProducts '�������������� ������
If Regim = "" Or Regim = "closeZakaz" Then loadPredmeti Me ' ������ ���������
If quantity2 > 0 Then
    str = "��������������"
Else
    str = "������������"
End If
If Regim = "fromDocs" Then
    Me.Caption = str & " ��������� � ��������� � " & numDoc
Else
    Me.Caption = str & " ��������� � ������ � " & numDoc
End If
If Regim = "closeZakaz" Then
    Me.Caption = "�������� � ������ � " & numDoc
    laBegin.Caption = "��� �������� �����. �������������� ��������� ����������."
    opNomenk.Enabled = False
    opProduct.Enabled = False
    cmSel.Enabled = False
    laGrid2.Enabled = False
    Grid2.Enabled = False
    cmExel2.Visible = False
    tv.Enabled = False
    laGrid1.Visible = False
    laGrid.Visible = False
Else
    opNomenk.value = True ': opNomenk_Click
End If

EN1:
oldHeight = Me.Height
oldWidth = Me.Width
tvVes = tv.Width / (tv.Width + Grid.Width + Grid2.Width)
gridVes = Grid.Width / (tv.Width + Grid.Width + Grid2.Width)
grid2Ves = Grid2.Width / (tv.Width + Grid.Width + Grid2.Width)
isLoad = True
End Sub

Sub loadProducts() ' ������������ ������ ��� ���������
MousePointer = flexHourglass
Grid2.Visible = False
quantity2 = 0
clearGrid Grid2
If numExt = 254 Then
    sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.perList, " & _
    "sGuideNomenk.ed_Izmer, sGuideNomenk.ed_Izmer2, sDMC.quant as quantity  " & _
    ",sGuideNomenk.Size, sGuideNomenk.cod " & _
    "FROM sGuideNomenk INNER JOIN sDMC ON sGuideNomenk.nomNom = sDMC.nomNom  " & _
    "WHERE (((sDMC.numDoc)=" & numDoc & " And (sDMC.numExt)=" & numExt & "));"
ElseIf Regim <> "fromDocs" Then GoTo AA
ElseIf numExt = 0 And sDocs.reservNoNeed Then ' ��� ��������������
    sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.perList,  " & _
    "sGuideNomenk.ed_Izmer, sGuideNomenk.ed_Izmer2, sDMCmov.quantity  " & _
    ",sGuideNomenk.Size, sGuideNomenk.cod " & _
    "FROM sGuideNomenk INNER JOIN sDMCmov ON sGuideNomenk.nomNom = sDMCmov.nomNom  " & _
    "WHERE (((sDMCmov.numDoc)=" & numDoc & "));"
Else ' ���-�� ������ ��� ���������� �� ���� � ������ ������
AA: sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.perList, " & _
    "sGuideNomenk.ed_Izmer, sGuideNomenk.ed_Izmer2, sDMCrez.quantity  " & _
    ",sGuideNomenk.Size, sGuideNomenk.cod " & _
    "FROM sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom  " & _
    "WHERE (((sDMCrez.numDoc)=" & numDoc & "));"
End If
'MsgBox sql

Set tbNomenk = myOpenRecordSet("##118", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    quantity2 = quantity2 + 1
    Grid2.TextMatrix(quantity2, fnNomNom) = tbNomenk!nomNom
'    Grid2.TextMatrix(quantity2, dnNomName) = tbNomenk!nomName
    Grid2.TextMatrix(quantity2, fnNomName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid2.TextMatrix(quantity2, fnEdIzm) = tbNomenk!ed_Izmer
    Grid2.TextMatrix(quantity2, fnQuant) = Round(tbNomenk!quantity, 2)
    If Regim = "fromDocs" Then
      If sDocs.isIntMove() Then
        Grid2.TextMatrix(quantity2, fnEdIzm) = tbNomenk!ed_Izmer2
        Grid2.TextMatrix(quantity2, fnQuant) = Round(tbNomenk!quantity / tbNomenk!perList, 2)
      End If
    End If
    
    Grid2.AddItem ""
    tbNomenk.MoveNext
  Wend
  Grid2.removeItem quantity2 + 1
End If
tbNomenk.Close
EN1:
Grid2.Visible = True
MousePointer = flexDefault


End Sub

'���������� ������ ������� �.�. ���� �����
Sub rightORleft(reg As String) ' reg =l ��� r
Static begWidth2 As Integer, begWidth As Integer, begLeft As Integer
Dim delta As Integer

If Regim = "ostatP" Or Regim = "ostat" Then Exit Sub
If begWidth = 0 Then ' �.�. ������ ���� ���
    begWidth = Grid.Width
    begWidth2 = Grid2.Width
    begLeft = Grid2.Left
End If
If opProduct.value Then
    delta = 2000 ' Product
Else
    delta = 1200 ' Nomenk
End If
 
If reg = "r" Then
    Grid.Width = begWidth
    Grid2.Width = begWidth2
    Grid2.Left = begLeft
ElseIf reg = "l" Then
    Grid.Width = begWidth + delta
    Grid2.Width = begWidth2 - delta
    Grid2.Left = begLeft + delta
End If

Grid3.Width = Grid.Width

End Sub


Private Sub Form_Resize()
Dim h As Integer, w As Integer, hh As Single, ww As Single

If Not isLoad Then Exit Sub
If Me.WindowState = vbMinimized Then Exit Sub
If Me.WindowState = vbMaximized And Me.Width > cDELLwidth Then '����� DELL
    Grid5.ColWidth(prDescript) = 2400 + 375
    Grid2.ColWidth(fnNomName) = 5250 + 930
Else
    Grid5.ColWidth(prDescript) = 840 + 375
    Grid2.ColWidth(fnNomName) = 2340 + 930
End If
setNameColWidth '� Grid3 � Grid

On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width

tv.Height = tv.Height + h
tv.Width = tv.Width + w * tvVes

Grid.Left = Grid.Left + w * tvVes
laGrid.Left = Grid.Left
cbInside.Left = laGrid1.Left + laGrid1.Width
Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w * gridVes
Grid3.Left = Grid.Left
laGrid1.Left = Grid3.Left
laBegin.Left = Grid3.Left
Grid3.Width = Grid.Width

Grid2.Left = Grid2.Left + w * (tvVes + gridVes)
laGrid2.Left = Grid2.Left
laGrid2.Top = laGrid2.Top + h / 2

Grid2.Top = Grid2.Top + h / 2
Grid2.Height = Grid2.Height + h / 2
Grid2.Width = Grid2.Width + w * grid2Ves
Grid5.Left = Grid2.Left
Grid5.Height = Grid5.Height + h / 2

laGrid5.Left = laGrid5.Left + w * (tvVes + gridVes)
Grid5.Width = Grid2.Width

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
laBegin.Top = laBegin.Top + h
laBegin.Left = laBegin.Left + w

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
ElseIf Regim = "fromDocs" Then
   If sDocs.isIntMove() Then sDocs.ckPerList.value = 1 Else sDocs.ckPerList.value = 0
End If
End Sub

Sub dostupOstatkiToGrid(Optional reg As String)
Dim s As Single, sum As Single, rr As Long, il As Long

Me.MousePointer = flexHourglass
If numExt = 254 Or (Regim = "fromDocs" And sDocs.Regim = "fromCeh") Then
    laGrid4.Caption = "����������� ������� �� ������������� '" & sDocs.getGridColSour() & "'"
Else
    laGrid4.Caption = "��������� �������"
End If
clearGrid Grid4
Grid4.FormatString = "|<�����|<��������|<��.���������|���-��"
Grid4.ColWidth(0) = 0
Grid4.ColWidth(frNomNom) = 870
Grid4.ColWidth(frNomName) = 4485
Grid4.ColWidth(frEdIzm) = 645
Grid4.ColWidth(frOstat) = 885

If reg = "multiN" Then
    Grid.col = nkQuant: il = 0
    For rr = 1 To Grid.Rows - 1
        Grid.row = rr
        If Grid.CellFontBold Then
            il = il + 1
            gNomNom = Grid.TextMatrix(rr, nkNomer)
            nomencOstatkiToGrid il
            Grid4.AddItem ""
        End If
    Next rr
    Grid4.removeItem Grid4.Rows - 1
Else
    nomencOstatkiToGrid 1
End If
Grid4.Visible = True
EN1:
Me.MousePointer = flexDefault
gridFrame.Visible = True
gridFrame.ZOrder

End Sub
'��.����� ������� �� �������
'����������� Grid4 ����������� � dostupOstatkiToGrid
Public Function nomencOstatkiToGrid(row As Long) As Single
Dim s As Single, str As String, str2 As String, str3 As String, z As Single

'�.�������
sql = "SELECT  nomName, Ed_Izmer, Ed_Izmer2, perList From sGuideNomenk " & _
"WHERE (((nomNom)='" & gNomNom & "'));"
'Debug.Print sql
byErrSqlGetValues "##144", sql, str, str2, str3, tmpSng
If row > 0 Then
    Grid4.TextMatrix(row, frNomNom) = gNomNom
    Grid4.TextMatrix(row, frNomName) = str
    Grid4.TextMatrix(row, frEdIzm) = str2
End If
If Regim = "fromDocs" Then
    nomencOstatkiToGrid = PrihodRashod("+", skladId) - PrihodRashod("-", skladId) '�. ������� �� ������
    If sDocs.isIntMove() Then
        If row > 0 Then Grid4.TextMatrix(row, frEdIzm) = str3
        nomencOstatkiToGrid = nomencOstatkiToGrid / tmpSng
    End If
Else
'��������� ��������� �������
AA:
If cbInside.Visible And cbInside.ListIndex = 1 Then
    FO = PrihodRashod("+", -1002) - PrihodRashod("-", -1002) ' �������
Else
    FO = PrihodRashod("+", -1001) - PrihodRashod("-", -1001) ' ������� �� ����������
    sql = "SELECT Sum(quantity) AS Sum_quantity, " & _
    "Sum(Sum_quant) AS Sum_Sum_quant From wCloseNomenk " & _
    "WHERE (((nomNom)='" & gNomNom & "'));"
'    Debug.Print sql
    If Not byErrSqlGetValues("##145", sql, z, s) Then myBase.Close: End
    nomencOstatkiToGrid = FO - (z - s) ' �����, ��� ���������
End If
End If
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
ElseIf grColor = groupColor1 Or grColor = groupColor2 Then
    curRow = Grid.row
    Grid.CellFontBold = True
'    Grid.col = nkQuant
    For il = curRow - 1 To 1 Step -1  '����� �� �����
        Grid.row = il
        If Grid.CellBackColor <> grColor Then Exit For
        Grid.CellFontBold = False
    Next il
    For il = curRow + 1 To Grid.Rows - 1 '���� �� �����
        Grid.row = il
        If Grid.CellBackColor <> grColor Then Exit For
        Grid.CellFontBold = False
    Next il
    Grid.row = curRow
End If
End Sub

Private Sub Grid_EnterCell()
Dim f As String, d As Single

If quantity = 0 Or Grid.col = nkQuant Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col

gNomNom = Grid.TextMatrix(mousRow, nkNomer)

Grid.CellBackColor = vbYellow
If mousCol = nkDostup And cbInside.ListIndex = 0 Then
    f = Grid.TextMatrix(mousRow, nkCurOstat)
    d = Grid.TextMatrix(mousRow, nkDostup)
    If d < f Then Grid.CellBackColor = &H88FF88
End If

End Sub

Private Sub Grid_GotFocus()
cmHide.Enabled = True
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid_DblClick
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Grid_EnterCell
End Sub

Private Sub Grid_LeaveCell()
If Grid.col <> 0 And Grid.col <> nkQuant Then Grid.CellBackColor = Grid.BackColor
End Sub


Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
End Sub

Private Sub Grid2_Click()
mousCol2 = Grid2.MouseCol
mousRow2 = Grid2.MouseRow
If quantity2 = 0 Then Exit Sub
If Grid2.MouseRow = 0 Then
    Grid2.CellBackColor = Grid2.BackColor
    If mousCol2 = fnQuant Then
        SortCol Grid2, mousCol2, "numeric"
    Else
        SortCol Grid2, mousCol2
    End If
    Grid2.row = 1    ' ������ ����� ����� ���������
    Grid_EnterCell
End If

End Sub


Private Sub Grid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid2.MouseRow = 0 And Shift = 2 Then
        MsgBox "ColWidth = " & Grid2.ColWidth(Grid2.MouseCol)
ElseIf Button = 2 And Regim = "fromDocs" And quantity2 > 0 Then
    Grid2.col = fnNomNom
    Grid2.row = Grid2.MouseRow
    mousRow2 = Grid2.row
    gNomNom = Grid2.Text
    Grid2.SetFocus
    Grid2.CellBackColor = vbButtonFace
    Me.PopupMenu mnContext2
    Grid2.CellBackColor = Grid2.BackColor
End If
End Sub

Sub gridOrGrid3Hide(Optional purpose As String)
Dim I As Integer, maxHeight As Integer

maxHeight = tv.Top + tv.Height - Grid3.Top + 30
If purpose = "grid3" Then ' ����
    Grid.Top = Grid3.Top
    Grid.Height = maxHeight
    Grid.ZOrder
ElseIf purpose = "grid" Then ' ���� ������ ���� Grid3
    Grid3.Height = maxHeight
    Grid3.ZOrder
Else '                ��� ������������
    I = Grid.CellHeight * max(3, (quantity + 3))  '����� ������ Grid

    If I > maxHeight / 2 Then I = maxHeight / 2
    Grid.Height = I
    Grid.Top = tv.Top + tv.Height - I + 30
    laGrid.Top = Grid.Top - laGrid.Height  '+90
    Grid3.Height = laGrid.Top - Grid3.Top
    If Not Grid3.RowIsVisible(mousRow3) Then rowViem mousRow3, Grid3
End If

End Sub


Sub newProductRow(row As Long)
    
    gProductId = Grid3.TextMatrix(row, gpId)
    gProduct = Grid3.TextMatrix(row, gpName)
    laGrid.Visible = True
    laGrid.Caption = "������ ������������ �� ������� '" & gProduct & "'"
    loadProductNomenk gProductId
    controlEnable True
    gridOrGrid3Hide ""
    Grid.TopRow = 1
    
End Sub
  
Private Sub Grid3_EnterCell()
'Static prevRow As Long

If quantity3 = 0 Or Grid3.MouseRow = 0 Then Exit Sub
If biloG3Enter_Cell Then Exit Sub '���� sub ��� ���������� � ����� ������

biloG3Enter_Cell = True
Timer1.Enabled = False

Timer1.Interval = 100
Timer1.Enabled = True

mousRow3 = Grid3.row
mousCol3 = Grid3.col

'If prevRow <> Grid3.row Then newProductRow
Grid3.CellBackColor = &HCCCCCC

'prevRow = Grid3.row

End Sub

Private Sub Grid3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        newProductRow Grid3.row
    End If
End Sub

Private Sub Grid3_LeaveCell()
Grid3.CellBackColor = Grid3.BackColor
End Sub

Private Sub Grid3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'�� �������� ��� � Grid3_Click

If Grid3.MouseRow = 0 Then
    Grid3.CellBackColor = Grid3.BackColor
    SortCol Grid3, mousCol3
    Grid3.row = 1    ' ������ ����� ����� ���������
    gridOrGrid3Hide "grid"
Else
    'If biloG3Enter_Cell Then Exit Sub
    'mousCol3 = Grid3.MouseCol
    'mousRow3 = Grid3.MouseRow
    'If quantity3 = 0 Then Exit Sub
    'Grid3.CellBackColor = &HCCCCCC
    'newProductRow
End If

End Sub

Private Sub Grid3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Grid3.MouseRow = 0 And Shift = 2 Then MsgBox "ColWidth = " & Grid3.ColWidth(Grid3.MouseCol)

On Error Resume Next ' ����� ����� �� tv_click
Grid3.row = Grid3.MouseRow '����� ����� ��������� ����.����� ������ �� gridOrGrid3Hide
Grid3.RowSel = Grid3.MouseRow '
biloG3Enter_Cell = False
newProductRow Grid3.MouseRow

End Sub

Private Sub Grid4_GotFocus()
    tbQuant.SetFocus
End Sub

Private Sub Grid4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid4.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid4.ColWidth(Grid4.MouseCol)

End Sub


Private Sub Grid5_Click()
If Not noClick Then Grid5_EnterCell
noClick = False
End Sub

Private Sub Grid5_DblClick()
Dim id As Integer

If mousRow5 = 0 Then Exit Sub
If mousCol5 = prEQuant Then
    If Not predmetiIsClose("prev") Then
        MsgBox "�� ����� ��� �������� ��������.", , "�������������� ����������!"
        Exit Sub
    End If
    sql = "SELECT StatusId from Orders WHERE (((numOrder)=" & gNzak & "));"
    If Not byErrSqlGetValues("##389", sql, id) Then Exit Sub

'    If id <> 0 And id <> 6 Then '�� ������ ��� ������
'        MsgBox "������� ���� ���������� ������ ����� ������, ����� ����� " & _
'        "������ ��� ������ �� ����������� �����.", , ""
    If id <> 0 And id <> 4 Then '�� ������ ��� �� �����
        MsgBox "������� ���� ���������� ������ ����� ������, ����� ����� " & _
        "������ ��� �����.", , ""
        Exit Sub
    End If
    If Not msgBilo And Grid5.TextMatrix(mousRow5, prEtap) = "" Then
        msgBilo = True
        MsgBox "���� �� ������ ������ ���� ���������� ������, ������� " & _
        "����������, ����������� ��� ���������� ����� �����.  ��� �������� " & _
        "���� ��������� ������������� ����������." & vbCrLf & "����� ������� " & _
        "����.", , "��������������"
    End If
End If
If Grid5.CellBackColor = &H88FF88 Then textBoxInGridCell tbMobile, Grid5
End Sub

Private Sub Grid5_EnterCell()
Static prevRow As Long, prevCol As Long
' EnterCell ����������� �� MouseDown

    If hasSelection(Grid5) Then Exit Sub
    
    If Not selectNomenkFlag And Not isCtrlDown Then
        If quantity5 = 0 Then Exit Sub
        If Grid5.row > quantity5 Then
            If prevRow <> Grid5.row Then
                laGrid2.Caption = "�������������� ������ ��������� ������:"
                loadProducts '�����-�� ������
            End If
            prevRow = Grid5.row
            Exit Sub
        End If
        
        mousRow5 = Grid5.row
        mousCol5 = Grid5.col
        getIdFromGrid5Row Me
        
        If mousCol5 = prSumm Or mousCol5 = prCenaEd Or mousCol5 = prEQuant Then
            Grid5.CellBackColor = &H88FF88
        Else
            Grid5.CellBackColor = vbYellow
        End If
         
         '�� ������� - ����� ��� �������� �� ���� ��������:
         'if Grid5.col = prEtap Then
         '   If IsNumeric(Grid5.TextMatrix(mousRow5, prEtap)) Then _
         '       productNomenkToGrid2 CInt(Grid5.TextMatrix(mousRow5, prEtap))
        'ElseIf prevRow <> Grid5.row Or prevCol = prEtap Then
        '    productNomenkToGrid2 CInt(Grid5.TextMatrix(mousRow5, prQuant))
        'End If
        
        If prevRow <> Grid5.row Then
             productNomenkToGrid2 CInt(Grid5.TextMatrix(mousRow5, prQuant))
        End If
        
        prevRow = Grid5.row
        prevCol = Grid5.col
    Else
        '����� ������ ������������ ��� �������� �������
        
    End If
    
End Sub
'$odbc14$
Sub productNomenkToGrid2(quant As Single)
Dim il As Long, str As String, str2 As String, str3 As String, str4 As String

If quantity5 = 0 Then Exit Sub

Grid2.Visible = False
clearGrid Grid2
quantity2 = 0
If Grid5.TextMatrix(mousRow5, prType) = "�������" Then
  ReDim NN(0): ReDim QQ(0)
  If productNomenkToNNQQ(quant, 0, 0) Then
    laGrid2.Caption = "������ �� �������� ������� '" & Grid5.TextMatrix(mousRow5, nkName) & "'"
'    Set tbNomenk = myOpenRecordSet("##193", "select * from sGuideNomenk", dbOpenForwardOnly)
'    If tbNomenk Is Nothing Then Exit Sub
'    tbNomenk.index = "PrimaryKey"
    For il = 1 To UBound(NN)
        quantity2 = quantity2 + 1
        Grid2.AddItem ""
        Grid2.TextMatrix(il, fnNomNom) = NN(il)
        sql = "SELECT Size, ed_Izmer, cod, nomName from sGuideNomenk " & _
        "WHERE (((nomNom)='" & NN(il) & "'));"
        byErrSqlGetValues "##413", sql, str, str2, str3, str4
        
'        tbNomenk.Seek "=", NN(il)
'        If tbNomenk.NoMatch Then msgOfEnd ("##194")
        Grid2.TextMatrix(il, fnNomName) = str3 & " " & str4 & " " & str
        Grid2.TextMatrix(il, fnEdIzm) = str2
        Grid2.TextMatrix(il, fnQuant) = QQ(il)
    Next il
'    tbNomenk.Close
    If quantity2 > 0 Then Grid2.removeItem Grid2.Rows - 1
  End If
Else
    laGrid2.Caption = "��������� ������������"
    
    quantity2 = 1
    Grid2.TextMatrix(1, fnNomNom) = Grid5.TextMatrix(mousRow5, prName)
    Grid2.TextMatrix(1, fnNomName) = Grid5.TextMatrix(mousRow5, prName) & _
        " " & Grid5.TextMatrix(mousRow5, prDescript)
    Grid2.TextMatrix(1, fnEdIzm) = Grid5.TextMatrix(mousRow5, prEdizm)
    Grid2.TextMatrix(1, fnQuant) = Grid5.TextMatrix(mousRow5, prQuant)
End If
Grid2.Visible = True

End Sub

Private Sub Grid5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 And Shift = vbCtrlMask Then
        ' �������� ��������� ����� ��������� ����� ������� �� ����� ������ ����
        isCtrlDown = True
    End If
    
    If KeyCode = vbKeyReturn Then Grid5_DblClick

End Sub

Private Sub Grid5_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Then
        isCtrlDown = False
    End If
    If KeyCode = vbKeyEscape Then Grid5_EnterCell
End Sub
Private Sub Grid5_LeaveCell()
    If Not isCtrlDown And Not hasSelection(Grid5) Then
        Grid5.CellBackColor = Grid5.BackColor
    End If

End Sub

Private Sub Grid5_LostFocus()
    Grid5_LeaveCell
End Sub




Private Sub cleanSelection(Grd As MSFlexGrid)
Dim I As Integer, j As Integer
Dim currentRow As Integer, currentCol As Integer

ReDim selectedItems(0)
currentCol = Grd.col
currentRow = Grd.row

For I = Grd.Rows - 1 To 1 Step -1
    Grd.row = I
    For j = Grd.Cols - 1 To 1 Step -1
        Grd.col = j
        Grd.CellBackColor = Grd.BackColor
        Grd.CellForeColor = Grd.ForeColor
    Next j
Next I
Grd.row = currentRow
Grd.col = currentCol

End Sub
Private Function hasSelection(Grd As MSFlexGrid) As Boolean
Dim I As Integer
Dim currentRow As Integer

    hasSelection = False
    If UBound(selectedItems) > 0 Then
        hasSelection = True
    End If
    
End Function
Private Function useMaxSelection() As Long
Dim I As Integer
Dim sz As Integer

    useMaxSelection = 0
    sz = UBound(selectedItems)
    For I = 1 To sz
        If selectedItems(I) > useMaxSelection Then
            useMaxSelection = selectedItems(I)
        End If
    Next I
    removeItem (CStr(useMaxSelection))
End Function

Private Sub appendItem(item As String)
Dim I As Integer
Dim found As Boolean: found = False
Dim sz As Integer

    sz = UBound(selectedItems)
    For I = 1 To sz
        If selectedItems(I) = item Then found = True: Exit For
    Next I
    If Not found Then
        ReDim Preserve selectedItems(sz + 1)
        selectedItems(sz + 1) = item
    End If
    
End Sub

Private Sub removeItem(item As String)
Dim I As Integer
Dim found As Boolean: found = False
Dim sz As Integer

    sz = UBound(selectedItems)
    For I = 1 To sz
        If found Then
            ' �������� �����
            selectedItems(I - 1) = selectedItems(I)
        ElseIf selectedItems(I) = item Then
            found = True
        End If
    Next I
    If found Then
        ReDim Preserve selectedItems(sz - 1)
    End If
    'selectedItems(sz) = item
    
End Sub

Private Sub mark(Grd As MSFlexGrid, setFlag As Boolean)
Dim fColorSel As Long
Dim bColorSel As Long
Dim I As Integer
Dim currentCol As Integer, currentLeft As Integer

    
'    If IsMissing(color) Then color = vbRed
    currentCol = Grd.col
    currentLeft = Grd.CellLeft
    
    If setFlag Then
        fColorSel = vbWhite
        bColorSel = vbRed
        appendItem (CStr(Grd.row))
    Else
        fColorSel = vbBlack
        bColorSel = Grd.BackColor
        removeItem (CStr(Grd.row))
    End If

    For I = 0 To Grd.Cols - 1
        Grd.col = I
        Grd.CellBackColor = bColorSel
        Grd.CellForeColor = fColorSel
    Next I
    'Grd.CellLeft = currentLeft
    Grd.col = currentCol
    
    
End Sub


Private Sub Grid5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim I As Integer

    If isCtrlDown And Button = 1 And Grid5.row <> 0 And Grid5.row <> Grid5.Rows - 1 Then
        '���������� ��� ������
        If Grid5.CellBackColor = vbRed Then
            mark Grid5, False
        Else
            mark Grid5, True
        End If
    End If
    
    If Shift = vbCtrlMask And Button = 1 Then
        'selectNomenkFlag = True
        'Grid5.SelectionMode = flexSelectionByRow
    End If
    If selectNomenkFlag And Button = 2 And Shift = 0 Then
        ' �������� ���� "�������� ������� �� ����"
        'selectNomenkFlag = selectNomenkFlag
    End If
    If selectNomenkFlag And Button = 1 And Shift <> 2 Then
        'selectNomenkFlag = False
        'Grid5.SelectionMode = flexSelectionFree
    End If
    
End Sub

Private Sub Grid5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If isCtrlDown Then
    Else
        If hasSelection(Grid5) And Button = 1 Then
            cleanSelection Grid5
        End If
        If Grid5.MouseRow = 0 Then
            If Shift = 2 Then MsgBox "ColWidth = " & Grid5.ColWidth(Grid5.MouseCol)
        ElseIf Button = 2 And 0 < Grid5.MouseRow And Grid5.MouseRow < Grid5.Rows - 1 _
        And quantity5 > 0 And Regim <> "closeZakaz" Then
'            Grid5.row = Grid5.MouseRow
'            Grid5.col = prName
'            Grid5.SetFocus
'            Grid5.CellBackColor = vbButtonFace
            getIdFromGrid5Row Me   ' gNomNom � gProductId
            If Not hasSelection(Grid5) Then
                isCtrlDown = True
                mark Grid5, True
                isCtrlDown = False
            End If
            
            Me.PopupMenu mnContext
            noClick = True '��� Grid5_Click
        End If
    End If
    If Button = 1 And Shift = 2 Then
        '�������� ��� ������
        
    End If
'If Button = 2 And frmMode = "" Then

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
    ReDim Preserve QQ3(leng): QQ3(leng) = prQuant * tbNomenk!quantity
    

End Sub

Function zakazNomenkToNNQQ() As Boolean
zakazNomenkToNNQQ = False

ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): QQ2(0) = 0: ReDim QQ3(0)

'���-�� �������� �������
sql = "SELECT xPredmetyByIzdelia.prId, " & _
"xPredmetyByIzdelia.prExt, " & _
"xPredmetyByIzdelia.quant, " & _
"xEtapByIzdelia.eQuant, " & _
"xEtapByIzdelia.prevQuant, " & _
"xEtapByIzdelia.prevQuant " & _
"FROM xPredmetyByIzdelia " & _
"LEFT JOIN xEtapByIzdelia ON (xPredmetyByIzdelia.prExt = xEtapByIzdelia.prExt) AND (xPredmetyByIzdelia.prId = xEtapByIzdelia.prId) AND (xPredmetyByIzdelia.numOrder = xEtapByIzdelia.numOrder)" & _
"WHERE (((xPredmetyByIzdelia.numOrder)=" & gNzak & "));"

Set tbProduct = myOpenRecordSet("##319", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Function

While Not tbProduct.EOF
  gProductId = tbProduct!prId
  prExt = tbProduct!prExt
  If IsNull(tbProduct!eQuant) Then
    productNomenkToNNQQ tbProduct!quant, 0, 0
  Else
    productNomenkToNNQQ tbProduct!quant, tbProduct!eQuant, tbProduct!prevQuant
    QQ2(0) = 1 ' ���� ����
  End If
  tbProduct.MoveNext
Wend
tbProduct.Close

'��������� ���-��
sql = "SELECT xPredmetyByNomenk.nomNom, xPredmetyByNomenk.quant as quantity, " & _
"xEtapByNomenk.eQuant, xEtapByNomenk.prevQuant FROM xPredmetyByNomenk " & _
"LEFT JOIN xEtapByNomenk ON (xPredmetyByNomenk.nomNom = xEtapByNomenk.nomNom) AND (xPredmetyByNomenk.numOrder = xEtapByNomenk.numOrder) " & _
"WHERE (((xPredmetyByNomenk.numOrder)=" & gNzak & "));"
Set tbNomenk = myOpenRecordSet("##320", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
While Not tbNomenk.EOF
  If IsNull(tbNomenk!eQuant) Then
    nomenkToNNQQ 1, 0, 0
  Else
    nomenkToNNQQ 1, (tbNomenk!eQuant / tbNomenk!quantity), (tbNomenk!prevQuant / tbNomenk!quantity)
    QQ2(0) = 1 ' ���� ����
  End If
  tbNomenk.MoveNext
Wend
tbNomenk.Close
zakazNomenkToNNQQ = True
End Function

'����� ���-�� ���� ReDim NN(0): ReDim QQ(0): ReDim QQ2(0) : ReDim QQ3(0):QQ2(0)=0 - �� �.�����
Function productNomenkToNNQQ(pQuant As Single, eQuant As Single, _
                                               prQuant As Single) As Boolean
Dim I As Integer, gr() As String

productNomenkToNNQQ = False

'���������� ���-�� �������
sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xGroup " & _
"FROM sProducts " & _
"INNER JOIN xVariantNomenc ON (sProducts.nomNom = xVariantNomenc.nomNom) AND (sProducts.ProductId = xVariantNomenc.prId) " & _
"WHERE (((xVariantNomenc.numOrder)=" & gNzak & ") AND (" & _
"(xVariantNomenc.prId)=" & gProductId & ") AND ((xVariantNomenc.prExt)=" & prExt & "));"
'Debug.Print sql
Set tbNomenk = myOpenRecordSet("##192", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
ReDim gr(0): I = 0
While Not tbNomenk.EOF
    nomenkToNNQQ pQuant, eQuant, prQuant
    I = I + 1
    ReDim Preserve gr(I): gr(I) = tbNomenk!xGroup
    tbNomenk.MoveNext
Wend
tbNomenk.Close
    
'������������ ���-�� �������
sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xGroup " & _
"From sProducts WHERE (((sProducts.ProductId)=" & gProductId & "));"

'MsgBox sql
Set tbNomenk = myOpenRecordSet("##177", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
While Not tbNomenk.EOF
    For I = 1 To UBound(gr) ' ���� ������ ������� �� ����� ���-��, �� ���
        If gr(I) = tbNomenk!xGroup Then GoTo NXT ' �����������, �.�. ��
    Next I                                      ' �� ������ � xVariantNomenc
    nomenkToNNQQ pQuant, eQuant, prQuant
NXT: tbNomenk.MoveNext
Wend
tbNomenk.Close

productNomenkToNNQQ = True
End Function


'$odbc15$
Public Sub mnDel_Click()
Dim pQuant As Single, I As Integer, str  As String, str2 As String
Dim comma As String
Dim hasEtap As Boolean
Dim gIndex As Integer
Dim j As Integer
Dim lIndex As Long

comma = ""
hasEtap = False
If beNaklads() Then Exit Sub

If Otgruz.loadOutDates Then '�������� ��������
    MsgBox "�� ������ " & gNzak & " ��� �������� ��������.", , "�������� ����������!"
    Exit Sub
End If
If hasSelection(Grid5) Then
    For I = 1 To UBound(selectedItems)
        str2 = str2 & comma & Grid5.TextMatrix(CInt(selectedItems(I)), prName)
        comma = ", "
        str = Grid5.TextMatrix(CInt(selectedItems(I)), prEQuant)
        If IsNumeric(str) Then
            If CSng(str) > 0 Then _
                hasEtap = True
        End If
    Next I
    
End If
If Not convertToIzdelie Then
    If MsgBox("�� ������ ������� �������(�) '" & str2 & _
    "'", vbYesNo Or vbDefaultButton2, "����������� ��������") = vbNo Then Exit Sub
'str = Grid5.TextMatrix(mousRow5, prEQuant)
End If

If hasEtap Then
  MsgBox "�� ����� ��� ���������� �������� ���� �������� �����. " _
    & "�������� (���� ����� �������������� ��������) ��� ��������, �������� ��� ��� �����������, ���� �� " _
    & "����������� ��� �� ������������ ��� ������� � ���� �� ���."
    Exit Sub
End If

wrkDefault.BeginTrans

deleteSelected

wrkDefault.CommitTrans

tmpVar = saveOrdered
If Not IsNumeric(tmpVar) Then GoTo ER1
wrkDefault.CommitTrans

Grid5.TextMatrix(Grid5.Rows - 1, prSumm) = tmpVar
Orders.openOrdersRowToGrid "##220":    tqOrders.Close
    
For j = 1 To UBound(selectedItems)
    lIndex = useMaxSelection()
    quantity5 = quantity5 - 1
    Grid5.removeItem lIndex
    If quantity5 <= 0 Then
        clearGridRow Grid5, lIndex
    End If
Next j

If Not convertToIzdelie Then
    loadProducts ' ���-�� ������
    Grid2.SetFocus ' ����� �� �����������
End If
Exit Sub
ER0:
tbDMC.Close
ER1:
wrkDefault.rollback
MsgBox "�������� �� ������", , "Error 196" '##196
End Sub

Private Sub deleteSelected()
Dim I As Integer, j As Integer
Dim pQuant As Single

    For j = 1 To UBound(selectedItems)
        mousRow5 = CInt(selectedItems(j))
        getIdFromGrid5Row Me
        If Grid5.TextMatrix(mousRow5, prType) = "�������" Then
            '�������� ���-� � ����.���������� ���-��� (�.�.��������� ��������)
        '    Set tbProduct = myOpenRecordSet("##138", "select * from xPredmetyByIzdelia", dbOpenForwardOnly)
            sql = "SELECT quant from xPredmetyByIzdelia " & _
            "WHERE (((numOrder)=" & numDoc & ") AND ((prId)=" & gProductId & _
            ") AND ((prExt)=" & prExt & "));"
            Set tbProduct = myOpenRecordSet("##138", sql, dbOpenForwardOnly)
        '    If tbProduct Is Nothing Then GoTo ER1
        '    tbProduct.index = "Key"
        '    tbProduct.Seek "=", numDoc, gProductId, prExt
        '    If tbProduct.NoMatch Then tbProduct.Close: GoTo ER1
            If tbProduct.BOF Then tbProduct.Close: Exit Sub
            pQuant = tbProduct!quant
            
            ReDim NN(0): ReDim QQ(0)
            productNomenkToNNQQ pQuant, 0, 0 ' �.�. ����� ���������
            
            tbProduct.Delete
            tbProduct.Close
            
            '�������� ���-�� ������� �� DMCrez
        '    Set tbDMC = myOpenRecordSet("##152", "select * from sDMCrez", dbOpenForwardOnly)
        '    If tbDMC Is Nothing Then Exit Sub
        '    tbDMC.index = "NomDoc"
            
            For I = 1 To UBound(NN)
                gNomNom = NN(I)
                If Not nomenkToDMCrez(-QQ(I)) Then Exit Sub
            Next I
         '   tbDMC.Close
        Else '��������� ���-��
        
            '�������� ����� �� �������
            sql = "DELETE From xEtapByNomenk WHERE (((numOrder)=" & gNzak & _
            ") AND ((nomNom)='" & gNomNom & "'));"
            myExecute "##336", sql, 0 '���� ����
            
            '�������� ���-�� �� �������
            sql = "SELECT quant from xPredmetyByNomenk " & _
            "WHERE (((numOrder)=" & gNzak & ") AND ((nomNom)='" & gNomNom & "'));"
        'MsgBox sql
        '    Set tbNomenk = myOpenRecordSet("##198", "select * from xPredmetyByNomenk", dbOpenForwardOnly)
            Set tbNomenk = myOpenRecordSet("##198", sql, dbOpenForwardOnly)
        '    If tbNomenk Is Nothing Then GoTo ER1
        '    tbNomenk.index = "Key"
        '    tbNomenk.Seek "=", numDoc, gNomNom
        '    If tbNomenk.NoMatch Then tbNomenk.Close: GoTo ER1
            If tbNomenk.BOF Then tbNomenk.Close: Exit Sub
            pQuant = tbNomenk!quant
            tbNomenk.Delete
            tbNomenk.Close
            
            '�������� ���-�� �� DMCrez
            If Not nomenkToDMCrez(-pQuant) Then Exit Sub
        End If
    Next j

End Sub

'$odbc15$
Private Sub mnDel2_Click()
Dim s As Single, str As String
 
If Not (Regim = "fromDocs" And sDocs.Regim = "fromCeh") Then _
    If beNaklads() Then Exit Sub

If MsgBox("������� ������� � '" & gNomNom & _
"', �� �������?", vbYesNo Or vbDefaultButton2, "����������� ��������") _
= vbNo Then GoTo EN1

If Regim = "fromDocs" And sDocs.Regim = "fromCeh" Then
    str = "rez":  If skladId = -1002 Then str = "mov"
    sql = "DELETE From sDMC" & str & " WHERE (((numDoc)=" & numDoc & _
    ") AND ((nomNom)='" & gNomNom & "'));"
    If myExecute("##341", sql) = 0 Then GoTo NX1
    Exit Sub
End If

wrkDefault.BeginTrans
sql = "SELECT quant from sDMC  WHERE (((numDoc)=" & numDoc & _
") AND ((numExt)=" & numExt & ") AND ((nomNom)='" & gNomNom & "'));"
'Set tbDMC = myOpenRecordSet("##123", "select * from sDMC", dbOpenForwardOnly)
Set tbDMC = myOpenRecordSet("##123", sql, dbOpenForwardOnly)
'If tbDMC Is Nothing Then GoTo ER1
'tbDMC.index = "NomDoc"
'tbDMC.Seek "=", numDoc, numExt, gNomNom
cErr = 179 '##179
'If tbDMC.NoMatch Then GoTo ER1
If tbDMC.BOF Then GoTo ER1
s = tbDMC!quant
tbDMC.Delete
tbDMC.Close

If ostatCorr(-s) Then
    wrkDefault.CommitTrans
Else
    cErr = 125 '##125
ER1: wrkDefault.rollback
    MsgBox "�� ������ ��������� ��������. " & _
    "�������� ��������������.", , "Error " & cErr
    GoTo EN1
End If
NX1:
quantity2 = quantity2 - 1
If quantity2 = 0 Then
    clearGridRow Grid2, 1
Else
    Grid2.removeItem mousRow2
End If

EN1: Grid2.SetFocus
End Sub

Private Sub mnOnfly_Click()
    OnFly.Show vbModal
    If convertToIzdelie Then
        loadPredmeti Me
        loadProducts
        convertToIzdelie = False
    End If
End Sub

Private Sub opNomenk_Click()

controlEnable False
laQuant.Visible = False
laQuant.Caption = ""

laGrid.Visible = False
gridOrGrid3Hide "grid3"

If Regim = "ostat" Then
    cmHide.Visible = True
Else
    Grid.ColWidth(nkName) = 2970
    Grid.ColWidth(nkQuant) = 0
End If
If Me.WindowState = vbMaximized And Me.Width > cDELLwidth Then '����� DELL
    Grid.ColWidth(nkName) = 4350
End If
    

laBegin = "� �������������� �������� (������ Mouse) ������, ��� ���� " & _
"��������� �������, ��� ����� ������������ ��� ������������ ���� ������."
If Regim = "" Then
    laBegin = laBegin & _
vbCrLf & "      �������� � ���� ������� ��������� ������� � ������� <��������>." & _
vbCrLf & vbCrLf & "��� ������������� ��������� ��� �������� ��� " & _
"������ �������."
Else
    laBegin = laBegin & vbCrLf & vbCrLf & "���� ��� ���� ���������� <Ctrl>, �� " & _
    "������� ������ ����� ��������� � ���������� ������." & vbCrLf & vbCrLf & _
    "��� �������� ���������� ������������ �� ���������(������) �������� � " & _
    "���������� ����(������ ���� Mouse) ������ '�������'"
End If
loadKlass

laGrid1.Caption = ""
cbInside.Enabled = True
End Sub

Sub controlEnable(EN As Boolean)
If Not EN Then ' ������ �����
    Grid.Visible = False
    Grid3.Visible = False
End If
If Regim <> "closeZakaz" Then cmSel.Enabled = EN
End Sub

Private Sub opProduct_Click()
cmHide.Visible = False

controlEnable False
laQuant.Visible = True
laQuant.Caption = "�������"

Grid3.ColWidth(gpNN) = 0
Grid3.ColWidth(gpId) = 0

Grid.ColWidth(0) = 0
Grid.ColWidth(nkNomer) = 0 '900
If Regim = "ostat" Then
    Grid3.ColWidth(gpName) = 2085
    Grid3.ColWidth(gpSize) = 1080
Else
    Grid3.ColWidth(gpName) = 1305
    Grid3.ColWidth(gpSize) = 840 '855
    Grid.ColWidth(nkQuant) = 700
End If
setNameColWidth
loadSeria tv
Dim str As String
str = ""

laBegin = "� ����� ������ �������� (������ Mouse) �����, ��� ���� ��������� " & _
"�������, ��� ����� ������������ ��� ������� ���� �����." & vbCrLf & _
"     �������� �� ������ �������,  ��� ��������� �������� � ���� " & _
" ������������."
If Regim = "ostat" Then
Else
    laBegin = laBegin & vbCrLf & "  ������� <��������>, ���������� ��������� " & _
    "���������� ������� � ������� <Enter>." & vbCrLf & vbCrLf & "��� " & _
    "������������� ��������� ��� �������� ��� ������ ������� ."
End If
laGrid1.Caption = ""

If cbInside.Visible Then
    cbInside.ListIndex = 0 '�����1
    cbInside.Enabled = False
End If
End Sub

Sub setNameColWidth()
If Me.WindowState = vbMaximized And Me.Width > cDELLwidth Then '����� DELL
    Grid3.ColWidth(gpDescript) = 5055
    Grid.ColWidth(nkName) = 4350 + 900
ElseIf Regim = "ostat" Then
    Grid3.ColWidth(gpDescript) = 3495
Else
    Grid.ColWidth(nkName) = 2100 + 900
    Grid3.ColWidth(gpDescript) = 2200
End If

End Sub
'$odbc15$
Function nomenkToDMC(delta As Single, Optional noLock As String = "") As Boolean
Dim s As Single, I As Integer

nomenkToDMC = False

If noLock = "" Then
    If Not lockSklad Then Exit Function
'    Set tbDMC = myOpenRecordSet("##123", "select * from sDMC", dbOpenForwardOnly)
'    If tbDMC Is Nothing Then GoTo EN1
'    tbDMC.index = "NomDoc"
End If

'tbDMC.Seek "=", numDoc, numExt, gNomNom
'If tbDMC.NoMatch Then
'    tbDMC.AddNew
'    tbDMC!numDoc = numDoc
'    tbDMC!numExt = numExt
'    tbDMC!nomNom = gNomNom
'    tbDMC!quant = Round(delta, 2)
'Else
'    tbDMC.Edit
'    tbDMC!quant = Round(tbDMC!quant + delta, 2)
'End If
'tbDMC.Update
sql = "UPDATE sDMC SET quant = quant + " & delta & _
" WHERE numDoc=" & numDoc & " AND numExt=" & numExt & _
" AND nomNom='" & gNomNom & "';"
I = myExecute("W##123", sql, 0)
If I > 0 Then
    GoTo EN1
ElseIf I < 0 Then ' ������ ���, ������� ���������
    sql = "INSERT INTO sDMC ( numDoc, numExt, nomNom, quant )" & _
    "SELECT " & numDoc & ", " & numExt & ", '" & gNomNom & "', " & delta & ";"
    Debug.Print sql
    If myExecute("##348", sql) <> 0 Then GoTo EN1
End If
'If noLock = "" Then tbDMC.Close

'������������ �������(��� ������������ �� ������������)
If Not ostatCorr(delta) Then MsgBox "�� ������ ��������� ��������. " & _
     "�������� ��������������.", , "Error 83" '##83
nomenkToDMC = True

EN1:
If noLock = "" Then lockSklad "un"
End Function

Sub nomenkToPredmeti()
Dim delta As Single, s As Single, quant As Single

  If Not lockSklad Then Exit Sub
  
  quant = tbQuant.Text
    sql = "select * from xPredmetyByNomenk where numOrder = " & numDoc & _
    " and nomNom = '" & gNomNom & "';"
    'Debug.Print sql
    Set tbNomenk = myOpenRecordSet("##117", sql, dbOpenForwardOnly)
On Error GoTo errr
    If tbNomenk Is Nothing Then GoTo EN1
'    tbNomenk.index = "Key"
'    tbNomenk.Seek "=", numDoc, gNomNom
'    If Not tbNomenk.NoMatch Then
    If Not tbNomenk.BOF Then
        MsgBox "������������ '" & gNomNom & "' ��� ���� � ��������� ������!", , "��������������"
        GoTo EN2
    End If

    
    s = nomencOstatkiToGrid(1) - quant ' ������������ ��������� �������
    If s < -0.005 Then '� 2� ������
        If MsgBox("������� ������ '" & gNomNom & "' � ��������� �������� " & _
        "�������� (" & s & "), ����������?", vbOKCancel Or vbDefaultButton2, _
        "�����������") = vbCancel Then GoTo EN2
    End If
        
    wrkDefault.BeginTrans
        
    If Not nomenkToDMCrez(quant) Then GoTo ER1
    
    tbNomenk.AddNew
    tbNomenk!numOrder = numDoc
    tbNomenk!nomNom = gNomNom
    tbNomenk!quant = quant
    tbNomenk.update
        
    wrkDefault.CommitTrans
  GoTo EN2
ER1: wrkDefault.rollback
EN2: tbNomenk.Close
EN1: lockSklad "un"
Exit Sub

errr:
Debug.Print sql: errorCodAndMsg ("���������� ������������ � ������")
End Sub


Private Sub tbMobile2_Change()

End Sub

Private Sub tbMobile5_Change()
    
End Sub

'$odbc15$
Public Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim c As Single, s As Single, str As String

If KeyCode = vbKeyReturn Then
    getIdFromGrid5Row Me
    str = Grid5.TextMatrix(mousRow5, prType)
    If str = "�������" Then
        strWhere = " WHERE (((xPredmetyByIzdelia.numOrder)=" & gNzak & _
        ") AND ((xPredmetyByIzdelia.prId)=" & gProductId & _
        ") AND ((xPredmetyByIzdelia.prExt)=" & prExt & "));"
    Else
        strWhere = " WHERE (((xPredmetyByNomenk.numOrder)=" & gNzak & _
        ") AND ((xPredmetyByNomenk.nomNom)='" & gNomNom & "'));"
    End If
    If mousCol5 = prEQuant Then 'prEtap
        If str = "�������" Then
            sql = "SELECT xEtapByIzdelia.prevQuant, xPredmetyByIzdelia.quant " & _
            "FROM xEtapByIzdelia RIGHT JOIN xPredmetyByIzdelia ON " & _
            "(xEtapByIzdelia.prExt = xPredmetyByIzdelia.prExt) AND " & _
            "(xEtapByIzdelia.prId = xPredmetyByIzdelia.prId) AND " & _
            "(xEtapByIzdelia.numOrder = xPredmetyByIzdelia.numOrder)" & strWhere
'            tmpStr = "xEtapByIzdelia"
        Else
            sql = "SELECT xEtapByNomenk.prevQuant, xPredmetyByNomenk.quant " & _
            "FROM xEtapByNomenk RIGHT JOIN xPredmetyByNomenk ON " & _
            "(xEtapByNomenk.nomNom = xPredmetyByNomenk.nomNom) AND " & _
            "(xEtapByNomenk.numOrder = xPredmetyByNomenk.numOrder) " & strWhere
'            tmpStr = "xEtapByNomenk"
        End If
        If Not byErrSqlGetValues("##315", sql, c, s) Then Exit Sub
        If Not isNumericTbox(tbMobile, 0, s - c) Then Exit Sub
        
        s = tbMobile.Text: s = Round(s + c, 2)
        
'        Set tbProduct = myOpenRecordSet("##316", tmpStr, dbOpenTable)
'        If tbProduct Is Nothing Then myBase.Close: End
'        tbProduct.index = "Key"
        If str = "�������" Then
            sql = "SELECT * from xEtapByIzdelia WHERE (((numOrder)=" & _
            gNzak & ") AND ((prId)=" & gProductId & ") AND ((prExt)=" & prExt & "));"
'            tbProduct.Seek "=", gNzak, gProductId, prExt
        Else
            sql = "SELECT * from xEtapByNomenk WHERE (((numOrder)=" & _
            gNzak & ") AND ((nomNom)='" & gNomNom & "'));"
'            tbProduct.Seek "=", gNzak, gNomNom
        End If
        Set tbProduct = myOpenRecordSet("##316", sql, dbOpenTable)
'        If Not tbProduct.NoMatch Then
        If Not tbProduct.BOF Then
            If s < 0.005 Then
                tbProduct.Delete
                Grid5.TextMatrix(mousRow5, prEtap) = ""
                Grid5.TextMatrix(mousRow5, prEQuant) = ""
            Else
                tbProduct.Edit
                GoTo AA
            End If
        ElseIf s > 0.005 Then
            tbProduct.AddNew
            tbProduct!numOrder = gNzak
            If str = "�������" Then
                tbProduct!prId = gProductId
                tbProduct!prExt = prExt
            Else
                tbProduct!nomNom = gNomNom
            End If
AA:         tbProduct!eQuant = s
            tbProduct.update
            Grid5.TextMatrix(mousRow5, prEtap) = s
            Grid5.TextMatrix(mousRow5, prEQuant) = tbMobile.Text
        End If
        tbProduct.Close
            
        lbHide
        Exit Sub
    End If
    
    If Not Me.convertToIzdelie Then
        If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    
        If mousCol5 = prSumm Then
            s = tbMobile.Text
            c = s / CSng(Grid5.TextMatrix(mousRow5, prQuant)) '�� ���������
            GoTo BB
        Else
            c = tbMobile.Text
            s = c * CSng(Grid5.TextMatrix(mousRow5, prQuant))
BB:         If str = "�������" Then
                sql = "UPDATE xPredmetyByIzdelia SET cenaEd = " & c & _
                "  WHERE (((numOrder)=" & gNzak & ") AND ((prId)=" & gProductId & _
                ") AND ((prExt)=" & prExt & "));"
            Else
                sql = "UPDATE xPredmetyByNomenk SET cenaEd = " & c & _
                " WHERE (((numOrder)=" & gNzak & ") AND ((nomNom)='" & gNomNom & "'));"
            End If
    '        MsgBox sql
            
            If myExecute("##205", sql) = 0 Then
                Grid5.TextMatrix(mousRow5, prCenaEd) = Round(c, 2)
                Grid5.TextMatrix(mousRow5, prSumm) = Round(s, 2)
                tmpVar = saveOrdered
                If IsNumeric(tmpVar) Then
                    Grid5.TextMatrix(Grid5.Rows - 1, prSumm) = tmpVar
                    Otgruz.saveShipped '���� ������ � �� ��������
                    Orders.openOrdersRowToGrid "##212"
                    tqOrders.Close
                End If
            End If
        End If
        lbHide
    End If
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Sub lbHide()
tbMobile.Visible = False
Grid5.Enabled = True
Grid5.SetFocus
Grid5_EnterCell
End Sub

'��.����� delta � �������� ������� �� ������� � ��������� ����� �������
'�.�. �������� ��� delta ����� �������� ������� �� tbQuant
Function deficitAndNoIgnore(delta As Single) As Boolean
Dim s As Single, il As Long


deficitAndNoIgnore = False
s = nomencOstatkiToGrid(il) - delta ' ������������ ��������� �������
If s < -0.005 Then
    If numExt = 254 Or numExt = 0 Then ' ��������� ��� �����. �� ����
        tmpStr = "' �� ������������� '" & sDocs.getGridColSour() & "'"
    Else
        tmpStr = "' � ��������� ��������"
    End If
    If MsgBox("������� ������ '" & gNomNom & tmpStr & " �������� (" & _
    s & "), ����������?", vbOKCancel Or vbDefaultButton2, "�����������") _
    = vbOK Then Exit Function
    deficitAndNoIgnore = True
End If
End Function

'$odbc15$
Private Sub tbQuant_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rr As Long, il As Long, pQuant As Single, s As Single, str As String
Dim I As Integer, NN2() As String

If KeyCode = vbKeyReturn Then
   
    
 
If opNomenk.value Then
  If Not isNumericTbox(tbQuant, 0.01) Then Exit Sub
  s = Round(tbQuant.Text, 2)
  If s <> tbQuant.Text Then
    MsgBox "����� ������ ����� ������� - �� ������ ����!", , "��������� ����"
    tbQuant.SetFocus
    Exit Sub
  End If
  If Regim = "" Then '�������� � ������ �.� 0<numExt<254
    nomenkToPredmeti
  Else
    s = Round(tbQuant.Text, 2)
    If sDocs.isIntMove() Then
        I = Round(s, 0)
        If s <> I Then
            MsgBox "���������� ������ ���� �����!", , "��������� ����"
            tbQuant.Text = "1": Exit Sub
        End If
  
        sql = "SELECT perList From sGuideNomenk " & _
        "WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
        byErrSqlGetValues "##343", sql, s
        s = Round(I * s, 2)
    End If
    If deficitAndNoIgnore(tbQuant.Text) Then Exit Sub
    If numExt = 0 And sDocs.reservNoNeed Then
        ' ���������� �� ���� �� ���.�������� � ������������ -  �� �����������
        nomenkToDMCrez s, "mov"
    Else '�� sDocs
'        If deficitAndNoIgnore(tbQuant.Text) Then Exit Sub
        If numExt = 254 Then
            nomenkToDMC s
        Else ' ���������� �� ���� (�� ���.�����)
            nomenkToDMCrez s
        End If
    End If
  End If
Else ' ����� ������ �������� ������, ������� ������ � �����(������) ed.izmer
  If Not isNumericTbox(tbQuant, 1) Then Exit Sub
  pQuant = Round(tbQuant.Text)
  tbQuant.Text = pQuant
  
  If Not lockSklad Then Exit Sub
  
  Grid.col = nkQuant: il = 0
  For rr = 1 To Grid.Rows - 1
    Grid.row = rr
    If Grid.CellFontBold Then
      il = il + 1
      gNomNom = Grid.TextMatrix(rr, nkNomer)
      s = CSng(Grid.TextMatrix(rr, nkQuant)) * pQuant
      If deficitAndNoIgnore(s) Then GoTo ER2
    End If
  Next rr
  
  wrkDefault.BeginTrans
  
'***** ��������� �������� � C����� �� ������������ (��� ����)*********************
  str = "sDMCrez"
  If numExt = 254 Then str = "sDMC" 'numExt = 254 ������������ Regim="fromDocs"
'  Set tbDMC = myOpenRecordSet("##152", str, dbOpenTable)
'  If tbDMC Is Nothing Then GoTo ER1
'  tbDMC.index = "NomDoc"
  
  Grid.col = nkQuant
  I = 0: ReDim NN(0)
  For rr = 1 To Grid.Rows - 1
    Grid.row = rr
    If Grid.CellFontBold Then
      gNomNom = Grid.TextMatrix(rr, nkNomer)
      If Grid.CellBackColor = groupColor1 Or Grid.CellBackColor = groupColor2 Then
'      If Grid.CellBackColor <> Grid.BackColor Then
        I = I + 1: ReDim Preserve NN(I): NN(I) = gNomNom '���������� ���-��
      End If
      s = CSng(Grid.TextMatrix(rr, nkQuant)) * pQuant
      If numExt = 254 Then ' ������������  ���������
        If Not nomenkToDMC(s, "noLock") Then GoTo ER0 '
      Else '   �������������� ������ ���� ��������� ��������� �� ��� ����(������������ ����� ����������)
        If Not nomenkToDMCrez(s) Then GoTo ER0
      End If
    End If
  Next rr
'  tbDMC.Close

If sDocs.Regim <> "fromCeh" Then
'******** ��������� �������� � C����� �� ��������  *************************
  If numExt <> 254 Then ' ���� ������ ������ ��� �������
      quickSort NN, 1
      If Not addToPredmetiTable(pQuant, getPrExtByNomenk()) Then GoTo ER1
  End If
'*************************************************************************
End If
  wrkDefault.CommitTrans
  
  lockSklad "un"
End If ' opNomenk.value
  
If Regim = "" Then loadPredmeti Me ' ������
loadProducts ' ���-�� ������
   
'  Grid2.col = dnQuant
'  Grid2.SetFocus
Grid.SetFocus
GoTo ES2

ER0:
'tbDMC.Close
ER1:
wrkDefault.rollback
ER2:
lockSklad "un"
GoTo ESC

ElseIf KeyCode = vbKeyEscape Then
ESC: tbQuant.Text = ""
ES2: gridFrame.Visible = False
    tbQuant.Enabled = False
    laQuant.Enabled = False
    cmSel.Enabled = True
End If

End Sub
'$odbc15$
'��� delta < 0 - ����. ��������
Function nomenkToDMCrez(ByVal delta As Single, Optional mov As String = "") As Boolean
Dim s As Single, I As Integer

nomenkToDMCrez = False
'    If mov = "mov" Then ' ���������� ������������ �� �����������
'        Set tbDMC = myOpenRecordSet("##152", "select * from sDMCmov", dbOpenTable)
'        GoTo AA
'    ElseIf mov = "" Then
'        Set tbDMC = myOpenRecordSet("##152", "select * from sDMCrez", dbOpenTable)
'AA:     If tbDMC Is Nothing Then Exit Function
'        tbDMC.index = "NomDoc"
'    End If
'        tbDMC.Seek "=", numDoc, gNomNom
'        If tbDMC.NoMatch Then
'            If delta < 0 Then '����� � ��������� ������(� sDMCrez ��� ��� ���)
'                GoTo EN1      '��� ��� ������������� ��������� ��������
'                msgOfEnd ("##195")
'            End If
'            tbDMC.AddNew
'            tbDMC!numDoc = numDoc
'            tbDMC!nomNom = gNomNom
'            tbDMC!quantity = Round(delta, 2)
'        Else
'            s = Round(tbDMC!quantity + delta, 2)
'            If s <= 0 Then tbDMC.Delete: GoTo EN1
'            tbDMC.Edit
'            tbDMC!quantity = s
'        End If
'        tbDMC.Update

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



'$odbc15!$
'lastExt=0 - ���� � ������� ������ ��� ��������� ��������(��� ��� �� ���������)
'���� ������� ��������, �������� � NN() ����, �� ���� ����� ��� ����������
'����� ���������� �����.  ����.����� �������� ��������
Function addToPredmetiTable(pQuant As Single, lastExt As Integer) As Boolean
Dim I As Integer

addToPredmetiTable = False

If UBound(NN) = 0 Then '������������ ���-�, � ���� numExt(� ������ � lastExt)=0
    If lastExt <> 0 Then msgOfEnd "##190", "lastExt=" & lastExt
    prExt = 0
ElseIf lastExt = 0 Then '���������� ���-�, �� ��� ���� ������ �� ����
    prExt = 1
ElseIf lastExt < 0 Then ' ���� �������� ����� ���-�, �� ������
    prExt = -lastExt + 1 ' ����� ������ ����.�������
Else ' ���� ������ ���� �������
    prExt = lastExt
End If


sql = "SELECT * from xPredmetyByIzdelia " & _
"WHERE (((numOrder)=" & numDoc & ") AND ((prId)=" & gProductId & _
") AND ((prExt)=" & prExt & "));"
'MsgBox sql
Set tbProduct = myOpenRecordSet("##185", sql, dbOpenForwardOnly)
On Error GoTo errr

wrkDefault.BeginTrans

'Debug.Print sql
If tbProduct.BOF Then
    If lastExt > 0 Then msgOfEnd "##317", "lastExt=" & lastExt
    tbProduct.AddNew
    tbProduct!numOrder = numDoc
    tbProduct!prId = gProductId
    tbProduct!prExt = prExt
    tbProduct!quant = pQuant
    tbProduct.update
' ���������� ���-�� �������
  If UBound(NN) > 0 Then
    Set tbNomenk = myOpenRecordSet("##191", "select * from xVariantNomenc", dbOpenForwardOnly)
    If tbNomenk Is Nothing Then End

    For I = 1 To UBound(NN)
        tbNomenk.AddNew
        tbNomenk!numOrder = numDoc
        tbNomenk!prId = gProductId
        tbNomenk!prExt = prExt
        tbNomenk!nomNom = NN(I)
        tbNomenk.update
    Next I
    tbNomenk.Close
  End If
  wrkDefault.CommitTrans
    
Else
    If lastExt < 0 Then msgOfEnd "##428", "lastExt=" & lastExt
    tbProduct.Edit
    tbProduct!quant = Round(tbProduct!quant + pQuant)
    tbProduct.update
End If
'EP:
tbProduct.Close
wrkDefault.CommitTrans
addToPredmetiTable = True
Exit Function

errr:
errorCodAndMsg ("���������� ��������")
End Function


'��������� ���� ordered � Orders
Function saveOrdered(Optional update As Boolean = True) As Variant
Dim s As Single, s1 As Single

saveOrdered = Null
sql = "SELECT Sum([quant]*[cenaEd]) From xPredmetyByIzdelia GROUP BY numOrder " & _
"HAVING (((numOrder)=" & gNzak & "));"
If Not byErrSqlGetValues("W##368", sql, s) Then Exit Function

sql = "SELECT Sum([quant]*[cenaEd]) From xPredmetyByNomenk GROUP BY numOrder " & _
"HAVING (((numOrder)=" & gNzak & "));"
'MsgBox sql
If Not byErrSqlGetValues("W##210", sql, s1) Then Exit Function

s = Round(s + s1, 2)
If update Then
    sql = "UPDATE Orders SET ordered = " & s & " WHERE (((numOrder)=" & gNzak & "));"
    If myExecute("##211", sql) = 0 Then saveOrdered = s
End If
saveOrdered = s
End Function

'=0 - ���� � ������� ������ ��� ��������� ��������(��� ��� �� ��������� ���
'������ ���-� ���) ���� ������� ��������, �������� � NN() ����, �� ���� �����
'��� ����������, ����� ���������� �����.  ����.����� �������� ��������
Function getPrExtByNomenk() As Integer
Dim I As Integer, j As Integer, prevExt As Integer

getPrExtByNomenk = 0 '
sql = "SELECT xVariantNomenc.prExt, xVariantNomenc.nomNom " & _
"From xVariantNomenc WHERE (((xVariantNomenc.numOrder)=" & numDoc & _
") AND ((xVariantNomenc.prId)=" & gProductId & ")) ORDER BY xVariantNomenc.prExt;"

Set tbNomenk = myOpenRecordSet("##187", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then myBase.Close: End
If Not tbNomenk.BOF Then
    j = 0: ReDim NN2(UBound(NN))

CC: j = j + 1
    If j <= UBound(NN) Then
        NN2(j) = tbNomenk!nomNom
    End If
    prevExt = tbNomenk!prExt
    tbNomenk.MoveNext
    If tbNomenk.EOF Then GoTo AA:
    If prevExt <> tbNomenk!prExt Then
AA:     If j = UBound(NN) Then ' ���������-�� ���-��(��� - ���� �������� ������)
            quickSort NN2, 1
            For I = 1 To UBound(NN)
                If NN(I) <> NN2(I) Then GoTo BB
            Next I
            getPrExtByNomenk = prevExt
            GoTo EN
        End If
BB:     j = 0
    End If
    If Not tbNomenk.EOF Then GoTo CC

    getPrExtByNomenk = -prevExt
End If
EN:
tbNomenk.Close
End Function

Private Sub Timer1_Timer()
biloG3Enter_Cell = False
Timer1.Enabled = False
End Sub

Private Sub tv_AfterLabelEdit(Cancel As Integer, NewString As String)
gSeriaId = Mid$(tv.SelectedItem.key, 2)
ValueToTableField "##115", "'" & NewString & "'", "sGuideSeries", "seriaName", "bySeriaId"
End Sub

Sub loadKlass()
Dim key As String, pKey As String, k() As String, pK()  As String
Dim I As Integer, iErr As Integer
bilo = False
sql = "SELECT sGuideKlass.*  From sGuideKlass ORDER BY sGuideKlass.parentKlassId;"
Set tbKlass = myOpenRecordSet("##102", sql, dbOpenForwardOnly)
If tbKlass Is Nothing Then myBase.Close: End
If Not tbKlass.BOF Then
 tv.Nodes.Clear
 Set Node = tv.Nodes.Add(, , "k0", "�������������")
 Node.Sorted = True
 Set Node = tv.Nodes.Add("k0", tvwChild, "all", "              ")

 ReDim k(0): ReDim pK(0): ReDim NN(0): iErr = 0
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
  tv.Nodes.item("all").Text = "���� ��������"
End If
tbKlass.Close

While bilo ' ���������� ��� �������
  bilo = False
  For I = 1 To UBound(k())
    If k(I) <> "" Then
        On Error GoTo ERR2 ' ��������� ��� ������
        Set Node = tv.Nodes.Add(pK(I), tvwChild, k(I), NN(I))
        On Error GoTo 0
        k(I) = ""
        Node.Sorted = True
    End If
NXT:
  Next I
Wend
tv.Nodes.item("k0").Expanded = True
Exit Sub
ERR1:
 iErr = iErr + 1: bilo = True
 ReDim Preserve k(iErr): ReDim Preserve pK(iErr): ReDim Preserve NN(iErr)
 k(iErr) = key: pK(iErr) = pKey: NN(iErr) = tbKlass!klassName
 Resume Next

ERR2: bilo = True: Resume NXT

End Sub

Sub loadProductNomenk(ByVal v_productId As Integer)
Dim s As Single, grBef As String

Me.MousePointer = flexHourglass

quantity = 0
Grid.Visible = False
clearGrid Grid

sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xGroup, " & _
"sGuideNomenk.Size, sGuideNomenk.cod, " & _
" sGuideNomenk.nomName, sGuideNomenk.ed_Izmer  " & _
"FROM sGuideNomenk INNER JOIN sProducts ON sGuideNomenk.nomNom = sProducts.nomNom " & _
"WHERE (((sProducts.ProductId)=" & v_productId & ")) ORDER BY sProducts.xGroup DESC;"
'MsgBox sql
Set tbNomenk = myOpenRecordSet("##108", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then Exit Sub
If Not tbNomenk.BOF Then
  Grid.col = nkQuant
   grBef = ""
  While Not tbNomenk.EOF
  
    Grid.row = quantity '�������� ����. ������ �.�. ��������
    Dim str As String: str = Grid.Text
    quantity = quantity + 1
    If grBef = tbNomenk!xGroup And grBef <> "" Then
        Grid.CellBackColor = grColor ' ����������
        Grid.CellFontBold = False
        Grid.row = quantity
        Grid.CellBackColor = grColor ' �������
        Grid.CellFontBold = False
        bilo = False ' �� ���� ������������ �����
    Else
        Grid.row = quantity
        Grid.CellFontBold = True
        If Not bilo Then '
            If grColor = groupColor1 Then
                grColor = groupColor2
            Else
                grColor = groupColor1
            End If
            bilo = True
        End If
    End If
    grBef = tbNomenk!xGroup
        
    gNomNom = tbNomenk!nomNom
    Grid.TextMatrix(quantity, nkNomer) = gNomNom
    Grid.TextMatrix(quantity, nkName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid.TextMatrix(quantity, nkEdIzm) = tbNomenk!ed_Izmer
'    ReDim Preserve QP(quantity): QP(quantity) = tbNomenk!quantity
        Grid.TextMatrix(quantity, nkQuant) = tbNomenk!quantity
    '��������� �������:
    Grid.TextMatrix(quantity, nkDostup) = Round(nomencOstatkiToGrid(-1), 2)
    If Regim = "ostat" Then
        Grid.TextMatrix(quantity, nkCurOstat) = Round(FO, 2)
    End If
    
    Grid.AddItem ""
    tbNomenk.MoveNext
  Wend
  Grid.removeItem quantity + 1
End If
tbNomenk.Close
Grid.Visible = True
Grid.ZOrder
Me.MousePointer = flexDefault

End Sub

Sub loadKlassNomenk()
Dim il As Long, s As Single, strWhere As String
Dim beg As Single, prih As Single, rash As Single, oldNow As Single



If tv.SelectedItem.key = "all" Then
    strWhere = ""
    quantity = 0
Else
  strWhere = "WHERE (((klassId)=" & gKlassId & "))"
End If

laGrid1.Caption = "������������ �� ������ '" & tv.SelectedItem.Text & "'"

If (Regim = "ostat" And cbInside.ListIndex = 1) Or Regim = "fromDocs" Then
    Grid.ColWidth(nkDostup) = 0 '��� ������ �������� �� ����������
Else
    Grid.ColWidth(nkDostup) = 700
End If

Me.MousePointer = flexHourglass

If beShift Then
    Grid.AddItem ""
Else
    quantity = 0
    Grid.Visible = False
    clearGrid Grid
End If


sql = "SELECT nomNom, nomName, Size, cod, ed_Izmer, ed_Izmer2, perList " & _
"From sGuideNomenk " & strWhere & " ORDER BY sGuideNomenk.nomNom ;"

Debug.Print sql

Set tbNomenk = myOpenRecordSet("##103", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
 tbNomenk.MoveFirst
 While Not tbNomenk.EOF
    gNomNom = tbNomenk!nomNom
    quantity = quantity + 1
    Grid.TextMatrix(quantity, nkNomer) = gNomNom
    Grid.TextMatrix(quantity, nkName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    
    s = nomencOstatkiToGrid(-1) '��������� �������
    Grid.TextMatrix(quantity, nkEdIzm) = tbNomenk!ed_Izmer
    If Regim = "fromDocs" Then
        If sDocs.isIntMove() Then GoTo AA
        GoTo NXT:
    End If
    Grid.TextMatrix(quantity, nkCurOstat) = Round(FO, 2)
    Grid.TextMatrix(quantity, nkDostup) = Round(s, 2)
    If Regim = "ostat" Then
      If cbInside.ListIndex = 0 Then '�����1
        Grid.TextMatrix(quantity, nkDostup) = Round(nomencOstatkiToGrid(-1) / tbNomenk!perList, 2)
        Grid.TextMatrix(quantity, nkCurOstat) = Round(FO / tbNomenk!perList, 2)
AA:     Grid.TextMatrix(quantity, nkEdIzm) = tbNomenk!ed_Izmer2
      End If
    End If
NXT:
    Grid.AddItem ""
    tbNomenk.MoveNext
 Wend
 If quantity > 0 Then Grid.removeItem quantity + 1
End If
tbNomenk.Close

Grid.Visible = True
EN1:
Me.MousePointer = flexDefault
End Sub

Sub loadSeriaProduct()
Dim il As Long, strWhere As String

If tv.SelectedItem.key = "k0" Then
    gSeriaId = 0
    Grid3.Visible = False
    Exit Sub
End If

Me.MousePointer = flexHourglass
laGrid1.Caption = "������ ������� ������� �� ����� '" & tv.SelectedItem.Text & "'"

quantity3 = 0
Grid3.Visible = False
clearGrid Grid3
il = 0

strWhere = " WHERE sGuideProducts.prSeriaId = " & gSeriaId _
    & " and exists (select 1 from sproducts where sproducts.productid = sGuideProducts.prId)"

sql = "SELECT sGuideProducts.prDescript, sGuideProducts.prName, " & _
"sGuideProducts.prId, sGuideProducts.prSize From sGuideProducts " & strWhere & _
"ORDER BY sGuideProducts.SortNom;"
Set tbProduct = myOpenRecordSet("##104", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then GoTo EN1
If Not tbProduct.BOF Then
 While Not tbProduct.EOF
    quantity3 = quantity3 + 1
    
    Grid3.TextMatrix(quantity3, gpNN) = quantity3
    Grid3.TextMatrix(quantity3, gpId) = tbProduct!prId
'    If Not IsNull(tbProduct!prName) Then
    Grid3.TextMatrix(quantity3, gpName) = tbProduct!prName
    Grid3.TextMatrix(quantity3, gpSize) = tbProduct!prSize
    Grid3.TextMatrix(quantity3, gpDescript) = tbProduct!prDescript

    Grid3.AddItem ""
    tbProduct.MoveNext
 Wend
 Grid3.removeItem quantity3 + 1
End If
tbProduct.Close
Grid3.Visible = True
EN1:
Me.MousePointer = flexDefault

End Sub
    
Private Sub tv_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer, str As String
If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
    tv_NodeClick tv.SelectedItem
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
If opProduct.value Then
    gSeriaId = Mid$(tv.SelectedItem.key, 2)
    loadSeriaProduct
    Grid3.Visible = True
    Grid.Visible = False
    laGrid.Visible = False
    cmSel.Enabled = False
    gridOrGrid3Hide "grid"
Else
    controlEnable True
    gKlassId = Mid$(tv.SelectedItem.key, 2)
    loadKlassNomenk
'    Grid.Visible = True
End If
Grid_EnterCell
On Error Resume Next
Grid.SetFocus

biloG3Enter_Cell = False

End Sub

