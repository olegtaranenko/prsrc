VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Nakladna 
   BackColor       =   &H8000000A&
   Caption         =   "�������� "
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmClose 
      Caption         =   "�������"
      Height          =   315
      Left            =   2820
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmSostav 
      Caption         =   "������ ���."
      Height          =   315
      Left            =   4380
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Frame gridFrame 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   2760
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin MSFlexGridLib.MSFlexGrid Grid4 
         Height          =   615
         Left            =   60
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1085
         _Version        =   393216
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Label laGrid4 
         Alignment       =   2  'Center
         Caption         =   "laGid4"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   5295
      End
   End
   Begin VB.TextBox tbMobile2 
      Height          =   315
      Left            =   780
      TabIndex        =   14
      Text            =   "tbMobile2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4095
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7223
      _Version        =   393216
      AllowBigSelection=   0   'False
      MergeCells      =   3
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "������"
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   5160
      Width           =   915
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "�����"
      Height          =   315
      Left            =   8880
      TabIndex        =   2
      Top             =   5160
      Width           =   795
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "������ � Exel"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1935
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3413
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Label laControl 
      Caption         =   "��������: "
      Height          =   195
      Left            =   6180
      TabIndex        =   29
      Top             =   5220
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label laControl2 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6975
      TabIndex        =   28
      Top             =   5160
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label laDest 
      Caption         =   "laDest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   26
      Top             =   2820
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Label laSours 
      Caption         =   "laSours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   25
      Top             =   2820
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label laDocNum 
      Caption         =   "laDocNum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   24
      Top             =   2460
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label laKomu 
      Caption         =   "����:"
      Height          =   195
      Index           =   1
      Left            =   2940
      TabIndex        =   23
      Top             =   2820
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label laOt 
      Caption         =   "��:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   22
      Top             =   480
      Width           =   255
   End
   Begin VB.Label laNakl 
      Caption         =   "��������� �"
      Height          =   195
      Left            =   180
      TabIndex        =   21
      Top             =   2460
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label laDate 
      Height          =   195
      Left            =   7020
      TabIndex        =   18
      Top             =   0
      Width           =   1155
   End
   Begin VB.Label laSignatura 
      BackColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   6780
      TabIndex        =   13
      Top             =   300
      Width           =   1395
   End
   Begin VB.Label laPerson 
      Caption         =   "�����������:"
      Height          =   195
      Left            =   5700
      TabIndex        =   12
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label laFirm 
      Caption         =   "laFirm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      TabIndex        =   11
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label laPlatel 
      Caption         =   "����������:"
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label laDest 
      Caption         =   "laDest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label laKomu 
      Caption         =   "����:"
      Height          =   195
      Index           =   0
      Left            =   3060
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Label laSours 
      Caption         =   "laSours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   540
      TabIndex        =   7
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label laOt 
      Caption         =   "��:"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   2820
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label laDocNum 
      Caption         =   "laDocNum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label laTitle 
      Caption         =   "��������� �"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   60
      Width           =   1095
   End
End
Attribute VB_Name = "Nakladna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����
Dim quantity2 As Long
Public mousCol2 As Long
Public mousRow2 As Long
Public Regim As String
Dim secondNaklad As String, beSUO As Boolean ' ���� �������� ���-��

Const nkNomNom = 1
Const nkNomName = 2
Const nkEdIzm = 3
Const nkTreb = 4
Const nkClos = 5
Const nkEtap = 6
Const nkEClos = 7
Const nkQuant = 8
Const nkIntEdIzm = 9
Const nkIntQuant = 10

'$odbc15$
Private Sub cmClose_Click()
Dim I As Integer, j As Integer, NN2() As String, k As Integer
Dim numExtO As Integer, id As Integer, l As Long, s As Single
Dim mov As Single, moveNum As String, per As Single, str As String, str2 As String

If Not lockSklad Then Exit Sub

ReDim NN(0): ReDim NN2(0): ReDim NN3(0): ReDim QQ(0): ReDim QQ2(0): ReDim QQ3(0)
I = 0: j = 0: moveNum = ""
For l = 1 To quantity2
  str = Grid2(0).TextMatrix(l, nkQuant)
  If IsNumeric(str) Then
    mov = 0
    gNomNom = Grid2(0).TextMatrix(l, nkNomNom)
    If Grid2(0).TextMatrix(l, 0) = "" Then '�������
        I = I + 1: ReDim Preserve NN(I): ReDim Preserve QQ(I)
        NN(I) = gNomNom: QQ(I) = str
        skladId = -1001: GoTo AA
    Else ' �������� ��������� �� ������ ��������
        j = j + 1: ReDim Preserve NN2(j)
        ReDim Preserve QQ2(j): ReDim Preserve QQ3(j)
        NN2(j) = gNomNom: QQ2(j) = str: QQ3(j) = 0
        skladId = -1002
        If IsNumeric(Grid2(0).TextMatrix(l, nkIntQuant)) Then '����� ������������
            sql = "SELECT perList, ed_izmer2 from sGuideNomenk " & _
            "WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
            If Not byErrSqlGetValues("##366", sql, per, str2) Then Exit Sub
            
            QQ3(j) = per * Grid2(0).TextMatrix(l, nkIntQuant)
            s = PrihodRashod("+", -1001) - PrihodRashod("-", -1001) '�. ������� �� ������
            s = Round(s - QQ3(j), 2)
            If s < 0 Then
              If MsgBox("������� ������ '" & gNomNom & "' � ����. �������� " & _
              "� ������������� '" & sDocs.lbInside.List(0) & _
              "' �������� (" & Round(s / per, 2) & " " & str2 & "), ����������?", _
              vbOKCancel Or vbDefaultButton2, "�����������") = vbCancel Then
                lockSklad "un"
                GoTo EN1
              End If
            End If
            mov = QQ3(j)
            moveNum = "yes"
        End If
        
AA:     s = PrihodRashod("+", skladId) - PrihodRashod("-", skladId) '�. ������� �� ������
        s = Round(mov + s - str, 2)
        If s < 0 Then
          If MsgBox("������� ������ '" & gNomNom & "' � ����. �������� " & _
          "� ������������� '" & sDocs.lbInside.List(-1001 - skladId) & _
          "' �������� (" & s & "), ����������?", _
          vbOKCancel Or vbDefaultButton2, "�����������") = vbCancel Then
            lockSklad "un"
            GoTo EN1
          End If
        End If
    End If
  End If
Next l
If I + j = 0 Then
  If Regim = "predmeti" Then
    MsgBox "��������� ���������� ��� ��� �������, ������� �� ������ �������.", , "������ ���������!"
  Else
    MsgBox ""
  End If
EN1: Grid2(0).SetFocus
    Exit Sub
End If

If cehId = 0 Then
  id = -6 '�������� ��� ������
ElseIf cehId = 3 Then           '$ceh$
  id = getStatiaId("��-�� SUB") '
Else
  id = -cehId
End If

wrkDefault.BeginTrans

Set tbDocs = myOpenRecordSet("##357", "select * from sDocs", dbOpenForwardOnly) 'dbOpenForwardOnly)
If tbDocs Is Nothing Then GoTo ER3

'Set tbDMC = myOpenRecordSet("##348", "select * from sDMC", dbOpenForwardOnly)
'If tbDMC Is Nothing Then GoTo ER1
'tbDMC.index = "NomDoc"

tmpDate = Now
numExtO = 0
If j > 0 Then numExtO = getNextNumExt()
If moveNum = "yes" Then
    If Not sDocs.getNextDocNum() Then GoTo ER1
    moveNum = numDoc
    numExt = 254
    tbDocs.AddNew
    tbDocs!numDoc = numDoc
    tbDocs!numExt = numExt
    tbDocs!xDate = tmpDate
    tbDocs!Note = gNzak & "/" & numExtO
    tbDocs!sourId = -1001
    tbDocs!destId = -1002
    tbDocs.update
    For k = 1 To j
      If QQ3(k) > 0 Then
        gNomNom = NN2(k)
        If Not sProducts.nomenkToDMC(QQ3(k), "noLock") Then GoTo ER2
      End If
    Next k
    tmpDate = DateAdd("S", 1, tmpDate)
End If

numDoc = gNzak
If j > 0 Then
  tbDocs.AddNew
  tbDocs!numDoc = numDoc
  tbDocs!numExt = numExtO
  tbDocs!xDate = tmpDate
  tbDocs!Note = moveNum
  tbDocs!sourId = -1002
  tbDocs!destId = id
  tbDocs.update
  For k = 1 To j
    gNomNom = NN2(k): numExt = numExtO
    If Not sProducts.nomenkToDMC(QQ2(k), "noLock") Then GoTo ER2
    If Not clrCehQuant Then GoTo ER2
  Next k
  tmpDate = DateAdd("S", 1, tmpDate)
End If

numExt = 0
If I > 0 Then
  numExt = getNextNumExt()
  tbDocs.AddNew
  tbDocs!numDoc = numDoc
  tbDocs!numExt = numExt
  tbDocs!xDate = tmpDate
  tbDocs!sourId = -1001
  tbDocs!destId = id
  tbDocs.update
  For k = 1 To I
    gNomNom = NN(k)
    If Not sProducts.nomenkToDMC(QQ(k), "noLock") Then GoTo ER2
    If Not clrCehQuant Then GoTo ER2
  Next k
End If
'tbDMC.Close
tbDocs.Close
wrkDefault.CommitTrans
lockSklad "un"
Unload Me
sDocs.loadDocs CStr(numExtO) & " " & moveNum ' �������� 1-3 ���������

Exit Sub

ER2:
'tbDMC.Close
ER1:
tbDocs.Close
ER3:
wrkDefault.rollback
lockSklad "un"
MsgBox "�������� �� ������. �������� ��������������.", , "Error - " & cErr

End Sub

Function getStatiaId(name As String) As Integer

sql = "SELECT sourceId from sGuideSource WHERE (((sourceId)<0 And " & _
"(sGuideSource.sourceId)>-1000) AND ((sGuideSource.sourceName)='" & name & "'));"
If byErrSqlGetValues("W##387", sql, getStatiaId) Then
    If getStatiaId = 0 Then GoTo AA
Else
AA: getStatiaId = -5 '���������� ����.
End If
End Function

Function clrCehQuant() As Boolean
clrCehQuant = False
sql = "UPDATE sDMCrez SET curQuant = 0, intQuant = 0 " & _
"WHERE (((sDMCrez.numDoc)=" & numDoc & ") AND ((sDMCrez.nomNom)='" & gNomNom & "'));"
If myExecute("##365", sql) = 0 Then clrCehQuant = True
End Function

Private Sub cmExel_Click()
Dim str As String
str = laDocNum(0).Caption
GridToExcel Grid2(0), "��������� � " & str
End Sub

Private Sub cmExit_Click()

Unload Me

End Sub

Private Sub cmPrint_Click()
laDate.Caption = Format(Now(), "dd.mm.yy hh:nn")
Me.PrintForm
End Sub

Private Sub cmSostav_Click()
Me.MousePointer = flexHourglass
sql = "SELECT xPredmetyByIzdelia.prId, xPredmetyByIzdelia.prExt, " & _
"sGuideProducts.prName, sGuideProducts.prDescript FROM sGuideProducts INNER JOIN xPredmetyByIzdelia " & _
"ON sGuideProducts.prId = xPredmetyByIzdelia.prId " & _
"WHERE (((xPredmetyByIzdelia.numOrder)=" & gNzak & "));"
Set tbProduct = myOpenRecordSet("##332", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Sub

If tbProduct.BOF Then
    MsgBox "� ����� ������ ��� ������� �������.", , ""
    GoTo EN1
End If

Report.Regim = "fromCehNaklad"
Report.Show vbModal

EN1:
tbProduct.Close
Me.MousePointer = flexDefault
End Sub

Private Sub Command1_Click()
MsgBox "ColWidth(nkTreb)=" & Grid2(1).ColWidth(nkTreb)
End Sub

Private Sub Form_Load()
Dim str As String
Dim notBay As Long, I As Integer, delta As Integer

oldHeight = Me.Height
oldWidth = Me.Width

laTitle.Caption = "    ����� �"
Grid2(0).MergeRow(0) = True
Grid2(0).FormatString = "|<�����|<��������|<��.���������|����������� �� ������|��������|����������� �� �����|�������� �� �����|���-��|�����������|�����������"
Grid2(0).ColWidth(0) = 0
Grid2(0).ColWidth(nkNomNom) = 945
Grid2(0).ColWidth(nkNomName) = 4500 '5265
Grid2(0).ColWidth(nkEdIzm) = 645
Grid2(0).ColWidth(nkQuant) = 735
'������� �����. ������� ���������� ����� � �� ������������ �������� (��. loadToGrid)
If Regim = "" Then '������ ���������
    Grid2(0).ColWidth(nkTreb) = 0
    Grid2(0).ColWidth(nkClos) = 0
Else
    Grid2(0).ColWidth(nkTreb) = 630
    Grid2(0).ColWidth(nkClos) = 855
End If
Grid2(0).ColWidth(nkEtap) = 780
Grid2(0).ColWidth(nkEClos) = 765
Grid2(0).ColWidth(nkIntEdIzm) = 700
Grid2(0).ColWidth(nkIntQuant) = 700
Grid2(1).FormatString = "|<�����|<��������|<��.���������|����������� �� " & _
"������|��������|����������� �� �����|�������� �� �����|���-��|�����������|�����������"
Grid2(1).ColWidth(0) = 0
Grid2(1).ColWidth(nkNomNom) = 945
Grid2(1).ColWidth(nkNomName) = 5265
Grid2(1).ColWidth(nkEdIzm) = 645
Grid2(1).ColWidth(nkQuant) = 735
Grid2(1).ColWidth(nkEtap) = 0 '780
Grid2(1).ColWidth(nkEClos) = 0 '765
Grid2(1).ColWidth(nkTreb) = 0
Grid2(1).ColWidth(nkClos) = 0
Grid2(1).ColWidth(nkIntEdIzm) = 0
Grid2(1).ColWidth(nkIntQuant) = 0

cmExit.Caption = "�����"
secondNaklad = ""
If Regim = "" Then ' ��� ����������
    laTitle.Caption = "��������� �"
    laDocNum(0).Caption = getStrDocExtNum(numDoc, numExt)
    If sDocs.Grid.RowSel - sDocs.Grid.row > 0 Then
        secondNaklad = sDocs.Grid.TextMatrix(sDocs.Grid.RowSel, dcNumDoc)
        laDocNum(1).Caption = secondNaklad
        cmExel.Visible = False
    End If
    laControl.Visible = True
    laControl2.Visible = True
ElseIf Regim = "predmeti" Then ' � ����
    Me.Caption = "�������� � ������."
    cmSostav.Visible = True
    GoTo BB
ElseIf Regim = "toNaklad" Then
    cmClose.Visible = True
    cmPrint.Visible = False
    cmExel.Visible = False
    laPerson.Visible = False
    laSignatura.Visible = False
    Me.Caption = "�������� ��������� ������."
    cmExit.Caption = "������"
BB: laDocNum(0).Caption = numDoc
    laOt(0).Visible = False
    laSours(0).Visible = False
    laKomu(0).Visible = False
End If

MousePointer = flexHourglass

laPlatel.Visible = False
laFirm.Visible = False
If Regim = "" And numExt = 0 Then
        laFirm.Visible = True
        laFirm.Caption = "(����������� �� " & Ceh(cehId) & ")"
ElseIf numExt <> 254 Then  '� ������
    sql = "SELECT Orders.numOrder, GuideFirms.Name " & _
    "FROM GuideFirms INNER JOIN Orders ON GuideFirms.FirmId = Orders.FirmId " & _
    "WHERE (((Orders.numOrder)=" & numDoc & "));"
    notBay = 0
    byErrSqlGetValues "W##170", sql, notBay, str
    If notBay > 0 Then GoTo AA ' ����� �� �� �������
    
    sql = "SELECT BayGuideFirms.Name " & _
    "FROM BayGuideFirms INNER JOIN BayOrders ON BayGuideFirms.FirmId = " & _
    "BayOrders.FirmId WHERE (((BayOrders.numOrder)=" & numDoc & "));"
    If byErrSqlGetValues("##353", sql, str) Then
AA:     laPlatel.Visible = True
        laFirm.Visible = True
        laFirm.Caption = str
    End If
End If

loadToGrid 0
I = 350 + (Grid2(0).CellHeight + 17) * quantity2
If secondNaklad = "" Then
    delta = I - Grid2(0).Height
    Me.Height = Me.Height + delta
Else ' � ����� ����� ����� 2 ���������
'laDocNum(1) � ��� 2� �������� ���� �������� ������� ���� Grid2(0)
    Grid2(0).Height = I
    delta = Grid2(0).Top + I - laDocNum(1).Top + 200
    laDocNum(1).Top = laDocNum(1).Top + delta
    laDocNum(1).Visible = True
    Grid2(1).Top = Grid2(1).Top + delta
    Grid2(1).Visible = True
    laNakl.Top = laNakl.Top + delta
    laNakl.Visible = True
    laOt(1).Top = laOt(1).Top + delta
    laOt(1).Visible = True
    laSours(1).Top = laSours(1).Top + delta
    laSours(1).Visible = True
    laKomu(1).Top = laKomu(1).Top + delta
    laKomu(1).Visible = True
    laDest(1).Top = laDest(1).Top + delta
    laDest(1).Visible = True
    
    sDocs.getDocExtNomFromStr secondNaklad: loadToGrid 1
    I = 350 + (Grid2(1).CellHeight + 17) * quantity2 '2� �������
    delta = delta + I - Grid2(1).Height ' ��������� 1� ������� + ���-� 2��
    Grid2(1).Height = I
    oldHeight = Me.Height + delta ' Me.Height=oldHeight � Resize
    
    cmPrint.Top = cmPrint.Top + delta
    cmExit.Top = cmExit.Top + delta
End If

MousePointer = flexDefault
End Sub
'ind=1 �.�. ������ ��� Regim = ""
Sub loadToGrid(ind As Integer)
Dim I As Integer, s As Single, s2 As Single, str As String, str2 As String


ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): QQ2(0) = 0: ReDim QQ3(0)

If Regim = "toNaklad" Then
  laSours(ind).Caption = ""
  If cehId = 3 Then                    '$ceh$
    laDest(ind).Caption = "��-�� SUB"  '
  Else
    laDest(ind).Caption = sDocs.lbStatia.List(cehId - 1)
  End If
  If Not sProducts.zakazNomenkToNNQQ Then Exit Sub
ElseIf Regim = "" Then
  sql = "SELECT sGuideSource.sourceName, sGuideDest.sourceName AS destName " & _
  "FROM sGuideSource AS sGuideDest INNER JOIN (sGuideSource INNER JOIN " & _
  "sDocs ON sGuideSource.sourceId = sDocs.sourId) ON sGuideDest.sourceId = sDocs.destId " & _
  "WHERE (((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & "));"
  Debug.Print sql
  If byErrSqlGetValues("##172", sql, str, str2) Then
      laSours(ind).Caption = str
      laDest(ind).Caption = str2
  End If
  If sDocs.reservNoNeed Then str = "mov" Else str = "rez"
  If numExt = 0 Then ' ���������� ��������� �� ����
    sql = "SELECT nomNom, quantity as quant  FROM sDMC" & str & _
    " WHERE (((numDoc)=" & numDoc & "));"
  Else
    sql = "SELECT nomNom, quant FROM sDMC " & _
    "WHERE (((numDoc)=" & numDoc & ") AND ((numExt)=" & numExt & "));"
  End If
  Set tbDMC = myOpenRecordSet("##318", sql, dbOpenForwardOnly)
  If Not tbDMC Is Nothing Then
    I = 0
    While Not tbDMC.EOF
        I = I + 1
        ReDim Preserve NN(I): NN(I) = tbDMC!nomNom
        ReDim Preserve QQ(I): QQ(I) = tbDMC!quant
        tbDMC.MoveNext
    Wend
    tbDMC.Close
  End If
ElseIf Regim = "predmeti" Then
  laSours(0).Caption = "�����1"
  If cehId = 1 Then
      laDest(ind).Caption = "��-�� YAG"
  ElseIf cehId = 2 Then
      laDest(ind).Caption = "��-�� CO2"
  ElseIf cehId = 3 Then                 '$$ceh
      laDest(ind).Caption = "��-�� SUB" '
  End If
  If Not sProducts.zakazNomenkToNNQQ Then GoTo EN1
End If



Grid2(ind).Visible = False
quantity2 = 0
clearGrid Grid2(ind)
beSUO = False
'Set tbNomenk = myOpenRecordSet("##129", "select * from sGuideNomenk", dbOpenForwardOnly)
'If tbNomenk Is Nothing Then GoTo EN1
'tbNomenk.index = "PrimaryKey"
For I = 1 To UBound(NN)
'    tbNomenk.Seek "=", NN(i)
    sql = "SELECT nomName, ed_Izmer, perList, Size, ed_Izmer2, cod " & _
    "from sGuideNomenk WHERE (((nomNom)='" & NN(I) & "'));"
    Set tbNomenk = myOpenRecordSet("##129", sql, dbOpenForwardOnly)
'    If Not tbNomenk.NoMatch Then
    If Not tbNomenk.BOF Then
        quantity2 = quantity2 + 1
        'Grid2(ind).TextMatrix(quantity2, 0) = tbNomenk!obrez
        If tbNomenk!perList > 1 Then Grid2(ind).TextMatrix(quantity2, 0) = "��" '��������
        Grid2(ind).TextMatrix(quantity2, nkNomNom) = NN(I)
        Grid2(ind).TextMatrix(quantity2, nkNomName) = tbNomenk!cod & " " & _
            tbNomenk!nomName & " " & tbNomenk!Size
        Grid2(ind).TextMatrix(quantity2, nkEdIzm) = tbNomenk!ed_Izmer
        If Regim = "" Then
            If laDest(ind).Caption = "�������" Then
              Grid2(ind).TextMatrix(quantity2, nkEdIzm) = tbNomenk!ed_Izmer2
              Grid2(ind).TextMatrix(quantity2, nkQuant) = Round(QQ(I) / tbNomenk!perList, 2)
            Else
              Grid2(ind).TextMatrix(quantity2, nkQuant) = Round(QQ(I), 2)
            End If
        Else ' "toNaklad"
            Grid2(ind).TextMatrix(quantity2, nkTreb) = Round(QQ(I), 2)
            Grid2(ind).TextMatrix(quantity2, nkEtap) = Round(QQ2(I) - QQ3(I), 2)
            
            sql = "SELECT Sum(quant) AS Sum_quant From sDMC WHERE " & _
            "(((sDMC.numDoc)=" & numDoc & ") AND ((sDMC.nomNom)='" & NN(I) & "'));"
            If byErrSqlGetValues("##194", sql, s) Then
                Grid2(ind).TextMatrix(quantity2, nkClos) = Round(s, 2)
                Grid2(ind).TextMatrix(quantity2, nkEClos) = Round(s - QQ3(I), 2)
            End If
            If Regim <> "" Then
              If tbNomenk!perList <> 1 Then '��� �������� ���. ������� ��� �����
                beSUO = True
                Grid2(ind).TextMatrix(quantity2, nkIntEdIzm) = tbNomenk!ed_Izmer2
              End If
              s = 0: s2 = 0
              sql = "SELECT curQuant, intQuant from sDMCrez " & _
              "WHERE (((numDoc)=" & gNzak & ") AND ((nomNom)='" & NN(I) & "'));"
              byErrSqlGetValues "##362", sql, s, s2
              If s > 0 Then _
                Grid2(ind).TextMatrix(quantity2, nkQuant) = s
              If s2 > 0 Then _
                Grid2(ind).TextMatrix(quantity2, nkIntQuant) = s2
            End If
        End If
        Grid2(ind).AddItem ""
    End If
Next I
tbNomenk.Close
If quantity2 > 0 Then
    Grid2(ind).removeItem quantity2 + 1
End If
If ind = 0 Then
  If QQ2(0) = 0 Then  '���� �� ������� ������� �������
    Grid2(0).ColWidth(nkEtap) = 0
    Grid2(0).ColWidth(nkEClos) = 0
  ElseIf dostup <> "" Then ' ��� ���������� ��������� � ����� ������
    Grid2(0).ColWidth(nkTreb) = 0
    Grid2(0).ColWidth(nkClos) = 0
  End If
  If Not beSUO Then
    Grid2(0).ColWidth(nkIntEdIzm) = 0
    Grid2(0).ColWidth(nkIntQuant) = 0
  End If
End If
Dim sum  As Long
sum = 0
For I = 0 To Grid2(ind).Cols - 1
    sum = sum + Grid2(ind).ColWidth(I)
Next I
sum = sum + 680 '650
If sum < 8300 Then sum = 8300
Me.Width = sum
Grid2(ind).col = nkQuant
EN1:
Grid2(ind).Visible = True
End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer
If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next

If secondNaklad <> "" Then Me.Height = oldHeight: Me.Top = 0

h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid2(0).Height = Grid2(0).Height + h
Grid2(0).Width = Grid2(0).Width + w
Grid2(1).Width = Grid2(1).Width + w

cmPrint.Top = cmPrint.Top + h
cmExel.Top = cmExel.Top + h
cmSostav.Top = cmSostav.Top + h
cmClose.Top = cmClose.Top + h
cmClose.Left = cmClose.Left + w
laControl.Top = laControl.Top + h
laControl2.Top = laControl2.Top + h
laControl.Left = laControl.Left + w
laControl2.Left = laControl2.Left + w
cmExit.Top = cmExit.Top + h
cmExit.Left = cmExit.Left + w
End Sub

Private Sub Form_Unload(Cancel As Integer)
Regim = "" '����� ��� lbInside_LostFocus
End Sub

Private Sub Grid2_DblClick(Index As Integer)
Dim str As String, per As Single, ed_Izmer As String

If Grid2(Index).CellBackColor = &H88FF88 Then '****************************
 
If mousCol2 = nkIntQuant Then
    str = Grid2(Index).TextMatrix(mousRow2, nkQuant)
    If Not IsNumeric(str) Then GoTo AA
    If CSng(str) <= 0 Then
AA:     MsgBox "������� ��������� �������� � ������� '���-��'", , "��������������"
        Exit Sub
    End If
End If

Me.MousePointer = flexHourglass
 
tmpStr = "����������� ������� �� ������������� '"
gNomNom = Grid2(Index).TextMatrix(mousRow2, fnNomNom)
 
If Grid2(Index).TextMatrix(mousRow2, 0) = "" Or mousCol2 = nkIntQuant Then '�������
    sql = "SELECT perList, ed_Izmer2 From sGuideNomenk " & _
    "WHERE (((nomNom)='" & gNomNom & "'));"
    byErrSqlGetValues "##364", sql, per, ed_Izmer
    If per < 0.01 Then per = 1
    
    laGrid4.Caption = tmpStr & sDocs.lbInside.List(0) & "'"
    skladId = -1001
Else ' ��������
    per = 1
    ed_Izmer = Grid2(Index).TextMatrix(mousRow2, fnEdIzm)
    
    laGrid4.Caption = tmpStr & sDocs.lbInside.List(1) & "'"
    skladId = -1002
End If
 clearGrid Grid4
 Grid4.FormatString = "|<�����|<��������|<��.���������|�������"
 Grid4.ColWidth(0) = 0
 Grid4.ColWidth(1) = 870
 Grid4.ColWidth(2) = 2745
 Grid4.ColWidth(3) = 645
 Grid4.ColWidth(4) = 900

 Grid4.TextMatrix(1, 1) = Grid2(Index).TextMatrix(mousRow2, fnNomNom)
 Grid4.TextMatrix(1, 2) = Grid2(Index).TextMatrix(mousRow2, fnNomName)
 Grid4.TextMatrix(1, 3) = ed_Izmer
 Grid4.TextMatrix(1, 4) = Round((PrihodRashod("+", skladId) - _
                    PrihodRashod("-", skladId)) / per, 2) '�. ������� �� ������

 Grid4.Visible = True
EN1:
 Me.MousePointer = flexDefault
 gridFrame.Visible = True
     textBoxInGridCell tbMobile2, Grid2(0)
End If '*************************************************************
End Sub

Private Sub Grid2_EnterCell(Index As Integer)
Dim t As Single, s As Single
If Index > 0 Then Exit Sub
mousRow2 = Grid2(Index).row
mousCol2 = Grid2(Index).col
If quantity2 = 0 Or Regim <> "predmeti" Or dostup = "" Then Exit Sub

If mousCol2 = nkQuant Or (mousCol2 = nkIntQuant And _
Grid2(Index).TextMatrix(mousRow2, 0) <> "") Then
    Grid2(Index).CellBackColor = &H88FF88
Else
    Grid2(Index).CellBackColor = vbYellow
End If

End Sub

Private Sub Grid2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid2_DblClick (Index)
End Sub

Private Sub Grid2_LeaveCell(Index As Integer)
Grid2(Index).CellBackColor = Grid2(Index).BackColor

End Sub

Private Sub Grid2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid2(Index).MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid2(Index).ColWidth(Grid2(Index).MouseCol)

End Sub

Sub lbHide2()
tbMobile2.Visible = False
gridFrame.Visible = False
Grid2(0).Enabled = True
Grid2(0).SetFocus
Grid2_EnterCell 0
End Sub

Private Sub Grid4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid4.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid4.ColWidth(Grid4.MouseCol)

End Sub

Private Sub tbMobile2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim delta As Single, quant As Single, s As Single, str As String

If KeyCode = vbKeyReturn Then
  
  If mousCol2 = nkQuant Then
    If QQ2(0) = 0 Then '��� �����
        quant = Grid2(0).TextMatrix(mousRow2, nkTreb)
        quant = Round(quant - Grid2(0).TextMatrix(mousRow2, nkClos), 2)
    Else
        quant = Grid2(0).TextMatrix(mousRow2, nkEtap)
        quant = Round(quant - Grid2(0).TextMatrix(mousRow2, nkEClos), 2)
    End If
    
    If Not isNumericTbox(tbMobile2, 0, quant) Then Exit Sub
    
    quant = Round(tbMobile2.Text, 2)
    str = "cur"
Else 'nkIntQuant
    If Not isNumericTbox(tbMobile2, 0) Then Exit Sub
    quant = Round(tbMobile2.Text, 0)
    If quant <> CSng(tbMobile2.Text) Then
        MsgBox "���������� ������ ���� �����!", , "��������������"
        Exit Sub
    End If
    str = "int"
End If
sql = "UPDATE sDMCrez SET " & str & "Quant = " & quant & _
" WHERE (((numDoc)=" & numDoc & ") AND ((nomNom)='" & _
Grid2(0).TextMatrix(mousRow2, nkNomNom) & "'));"
If myExecute("##363", sql) = 0 Then
    If quant = 0 Then
        Grid2(0).TextMatrix(mousRow2, mousCol2) = ""
    Else
        Grid2(0).TextMatrix(mousRow2, mousCol2) = quant
    End If
End If
lbHide2
Grid2(0).SetFocus

ElseIf KeyCode = vbKeyEscape Then
NN:  lbHide2
End If

End Sub
