VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Documents 
   BackColor       =   &H8000000A&
   Caption         =   "�����"
   ClientHeight    =   6192
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   11748
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6192
   ScaleWidth      =   11748
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lbVenture 
      Appearance      =   0  'Flat
      Height          =   600
      Left            =   5500
      TabIndex        =   31
      Top             =   1000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����"
      Height          =   315
      Left            =   3840
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox tbEnable 
      BackColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   11280
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.ListBox lbInside 
      Height          =   240
      ItemData        =   "Documents.frx":0000
      Left            =   9000
      List            =   "Documents.frx":0002
      TabIndex        =   26
      Top             =   2520
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox lbGroup 
      Height          =   432
      ItemData        =   "Documents.frx":0004
      Left            =   9000
      List            =   "Documents.frx":000E
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ListBox lbSource 
      Height          =   2352
      Left            =   6300
      TabIndex        =   27
      Top             =   780
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.CommandButton cmAdd2 
      Caption         =   "��������"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6300
      TabIndex        =   16
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   1260
      Top             =   5640
   End
   Begin VB.CommandButton cmProduct 
      Caption         =   "�������"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6600
      TabIndex        =   25
      Top             =   5100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame frBad 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   600
      TabIndex        =   19
      Top             =   1500
      Visible         =   0   'False
      Width           =   4515
      Begin VB.CommandButton cmExit 
         Caption         =   "�����"
         Height          =   315
         Left            =   3480
         TabIndex        =   24
         Top             =   3300
         Width           =   915
      End
      Begin VB.CommandButton cmKarta 
         Caption         =   "��������"
         Height          =   315
         Left            =   180
         TabIndex        =   23
         Top             =   3300
         Width           =   915
      End
      Begin VB.ListBox lbBad 
         Height          =   2544
         ItemData        =   "Documents.frx":002F
         Left            =   60
         List            =   "Documents.frx":0031
         TabIndex        =   20
         Top             =   540
         Width           =   4395
      End
      Begin VB.Label laFrame 
         Alignment       =   2  'Center
         Caption         =   "laFrame"
         Height          =   435
         Left            =   180
         TabIndex        =   22
         Top             =   60
         Width           =   4155
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   4515
      End
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "�������"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3960
      TabIndex        =   18
      Top             =   5760
      Width           =   915
   End
   Begin VB.CommandButton cmDel2 
      Caption         =   "�������"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7740
      TabIndex        =   17
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox tbMobile2 
      Height          =   315
      Left            =   10320
      TabIndex        =   15
      Text            =   "tbMobile2"
      Top             =   4500
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   4680
      TabIndex        =   14
      Text            =   "tbMobile"
      Top             =   1020
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "��������"
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   5760
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4815
      Left            =   5880
      TabIndex        =   10
      Top             =   735
      Visible         =   0   'False
      Width           =   5835
      _ExtentX        =   10287
      _ExtentY        =   8509
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4815
      Left            =   60
      TabIndex        =   9
      Top             =   735
      Width           =   5775
      _ExtentX        =   10181
      _ExtentY        =   8509
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmLoad 
      Caption         =   "���������"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   -60
      Width           =   11655
      Begin VB.TextBox tbEndDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Top             =   180
         Width           =   795
      End
      Begin VB.CheckBox ckEndDate 
         Caption         =   " "
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox tbStartDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1260
         TabIndex        =   2
         Text            =   "01.11.02"
         Top             =   180
         Width           =   795
      End
      Begin VB.CheckBox ckStartDate 
         Caption         =   " "
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   180
         Width           =   315
      End
      Begin VB.Label laFiltr 
         Caption         =   "�������� �� ����� ���!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7020
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label laPo 
         Caption         =   "���"
         Height          =   195
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   195
      End
      Begin VB.Label laPeriod 
         Caption         =   "������ �  "
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Label laGrid2 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   6120
      TabIndex        =   12
      Top             =   540
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "������ ����������"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   540
      Width           =   5535
   End
   Begin VB.Menu mnReestr 
      Caption         =   "���������"
      Begin VB.Menu mnDocFind 
         Caption         =   "����� �� ������       F7"
      End
   End
   Begin VB.Menu mnMeassure 
      Caption         =   "���������"
      Begin VB.Menu mnVentureIncomeSetting 
         Caption         =   "������ �� ������������"
      End
   End
   Begin VB.Menu mnReports 
      Caption         =   "������"
      Begin VB.Menu mnOstVed 
         Caption         =   "���. �������� �� ����"
      End
      Begin VB.Menu mnOborot 
         Caption         =   "��������� ���������"
         Visible         =   0   'False
      End
      Begin VB.Menu sourOborot 
         Caption         =   "��������� �� �����������"
         Visible         =   0   'False
      End
      Begin VB.Menu ventureOborot 
         Caption         =   "��������� �� ������������"
      End
      Begin VB.Menu VentureRest 
         Caption         =   "������� �� ������������"
      End
      Begin VB.Menu mnFiltrOborot 
         Caption         =   "������������ ��� �������"
         Visible         =   0   'False
      End
      Begin VB.Menu mnReservedAll 
         Caption         =   "����������������� ���-��"
      End
      Begin VB.Menu mnSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnSkladStand 
         Caption         =   "��������� ������"
      End
      Begin VB.Menu mnKarta 
         Caption         =   "�������� ��������"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnServic 
      Caption         =   "������"
      Begin VB.Menu mnOstat 
         Caption         =   "�������� �� �����.�������"
      End
      Begin VB.Menu mnViewOst 
         Caption         =   "���������� ����. ��������"
      End
      Begin VB.Menu mnCurOstat 
         Caption         =   "������ ������� ��������"
      End
      Begin VB.Menu mnVentureOrder 
         Caption         =   "��������� ����� �������������"
      End
      Begin VB.Menu mnSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnWebs 
         Caption         =   "����� ��� Web"
      End
      Begin VB.Menu mnWeb 
         Caption         =   "���� �������� ��� WEB"
         Visible         =   0   'False
      End
      Begin VB.Menu mnToExcel 
         Caption         =   "Web ����� � Excel"
      End
      Begin VB.Menu mnPriceToExcel 
         Caption         =   "Web ����� � Excel"
      End
      Begin VB.Menu mnFilters 
         Caption         =   "WEB �������"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnGuides 
      Caption         =   "�����������"
      Begin VB.Menu mnNomenc 
         Caption         =   "������������"
      End
      Begin VB.Menu mnProducts 
         Caption         =   "������� �������"
      End
      Begin VB.Menu mnSource 
         Caption         =   "����������"
      End
      Begin VB.Menu mnInside 
         Caption         =   "����.�������-�"
      End
      Begin VB.Menu mnStatia 
         Caption         =   "������ ������"
      End
      Begin VB.Menu mnGuideFormuls 
         Caption         =   "������� ��� ������"
      End
      Begin VB.Menu mnManag 
         Caption         =   "���������"
      End
   End
   Begin VB.Menu mnContext5 
      Caption         =   "�������� ������� ������-��"
      Visible         =   0   'False
      Begin VB.Menu mnAdd5 
         Caption         =   "�������� �� �����-��"
      End
      Begin VB.Menu mnDel5 
         Caption         =   "�������"
      End
   End
End
Attribute VB_Name = "Documents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isLoad As Boolean
Dim objExel As Excel.Application, exRow As Long


Dim quantity  As Long
Dim sum As Single
'Dim guideDist(10) As String
Dim mousCol As Long, mousRow As Long
Dim mousCol2 As Long, mousRow2 As Long
Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����

Dim quantity2 As Long
Dim minut As Integer
Dim partial As Boolean
Dim destId As String

Const dcDate = 1
Const dcNumDoc = 2
'Const dcM = 3
Const dcSour = 3
Const dcDest = 4
Const dcNote = 6
Const dcVenture = 5
'Grid2
Const dnNomNom = 1
Const dnNomName = 2
Const dnQuant2 = 3
Const dnEdIzm2 = 4
Const dnQuant = 5
Const dnEdIzm = 6

'reg ="","single","add"
Sub loadDocs(Optional reg As String = "")
Dim strWhere As String, I As Integer, str As String
 prevRow = -1
 Grid.Visible = False
 numExt = 255 '��������� ���������
 If reg = "" Then
    strWhere = getWhereByDateBoxes(Me, "sDocs.xDate", CDate("01.11.2000"))
 ElseIf reg = "docsFind" Then
    strWhere = "sDocs.numDoc=" & numDoc
 Else
    strWhere = "sDocs.numDoc=" & numDoc & " AND sDocs.numExt=" & numExt
 End If
 If strWhere <> "" Then strWhere = " AND " & strWhere
 
 Me.MousePointer = flexHourglass
 If reg <> "add" Then
    gridIsLoad = False
    quantity = 0
    clearGrid Grid
 End If
 
 sql = "SELECT sDocs.xDate, sDocs.Note, sDocs.numDoc, sDocs.numExt, " & _
 "sDocs.destId, sGuideSource.sourceName, GS.sourceName AS destName " & _
 ", v.ventureName as venture_name " & _
 "FROM sDocs INNER JOIN sGuideSource ON sDocs.sourId = sGuideSource.sourceId " & _
 "JOIN sGuideSource AS GS ON sDocs.destId = GS.sourceId " & _
 "left JOIN guideVenture v ON v.ventureId = sDocs.ventureId " & _
 "WHERE sDocs.numExt =255 " & strWhere & "  ORDER BY sDocs.xDate; "
' "WHERE ((" & str & " AND (GuideManag.Manag)='" & cbM.Text & "' " & strWhere & "));"
' Debug.Print sql
 Set tbDocs = myOpenRecordSet("##176", sql, dbOpenForwardOnly)
 If tbDocs Is Nothing Then End
 If Not tbDocs.BOF Then
  While Not tbDocs.EOF
    Grid.AddItem ""
    quantity = quantity + 1
    LoadDate Grid, quantity, dcDate, tbDocs!xDate, "dd.mm.yy"
'   str = tbDocs!numDoc & "/" & tbDocs!numExt
    str = tbDocs!numDoc
    Grid.TextMatrix(quantity, dcNumDoc) = str
    Grid.TextMatrix(quantity, 0) = tbDocs!destId
    Grid.TextMatrix(quantity, dcSour) = tbDocs!SourceName
    Grid.TextMatrix(quantity, dcDest) = tbDocs!destName
    Grid.TextMatrix(quantity, dcNote) = tbDocs!note
    If Not IsNull(tbDocs!venture_name) Then _
        Grid.TextMatrix(quantity, dcVenture) = tbDocs!venture_name

    tbDocs.MoveNext
  Wend
End If
'Debug.Print sql
tbDocs.Close
rowViem quantity, Grid
Grid.Visible = True
If quantity > 0 Then
    If reg <> "add" Or quantity = 1 Then Grid.RemoveItem quantity + 1
    Grid.col = 1
    gridIsLoad = True '
    Grid.col = 2      '����� loadDocNomenk
    Grid.row = quantity
'    loadDocNomenk
    On Error Resume Next
    Grid.SetFocus
    cmDel.Enabled = True
    Grid2.Visible = True
    laGrid2.Visible = True
    cmAdd2.Enabled = True
Else
    cmDel.Enabled = False
    Grid2.Visible = False
    laGrid2.Visible = False
    cmAdd2.Enabled = False
End If
gridIsLoad = True

Me.MousePointer = flexDefault
    
End Sub

Private Sub cbM_Change()

End Sub

Private Sub ckEndDate_Click()
If ckEndDate.value = 1 Then
    tbEndDate.Enabled = True
Else
    tbEndDate.Enabled = False
End If

End Sub

Private Sub ckStartDate_Click()
If ckStartDate.value = 1 Then
    tbStartDate.Enabled = True
Else
    tbStartDate.Enabled = False
End If

End Sub

Private Sub cmAdd_Click()
Dim str As String, intNum As Integer, l As Long, il As Long
Dim strNow As String, DateFromNum As String, dNow As Date
 
il = right$(Format(Now, "yymmdd\0\0"), 7) + 200001  ' ����� �� �������� � ��������

Set tbSystem = myOpenRecordSet("##149", "System", dbOpenTable)
If tbSystem Is Nothing Then Exit Sub
tbSystem.Edit
l = tbSystem!lastDocNum + 1
If l < il Then l = il
tbSystem!lastDocNum = l
tbSystem.Update
tbSystem.Close
numDoc = l
numExt = 255
addDoc
Grid.col = dcSour
End Sub

Sub addDoc()
Set tbDocs = myOpenRecordSet("##129", "sDocs", dbOpenTable) 'dbOpenForwardOnly)
If tbDocs Is Nothing Then Exit Sub

On Error GoTo ERR1
tbDocs.AddNew
'If mnIn.Checked Then
    tbDocs!sourId = 0
    tbDocs!destId = -1001
'Else
'    tbDocs!sourId = -1001
'    tbDocs!destId = -6
'End If

tbDocs!numDoc = numDoc
tbDocs!numExt = numExt
tbDocs!xDate = Now
'tbDocs!ManagId = manId(cbM.ListIndex)
tbDocs.Update
tbDocs.Close

loadDocs "add" ' �� ��������� ��� ���-��
Exit Sub
ERR1:
errorCodAndMsg "##tDocs update"
End Sub

Private Sub cmAdd2_Click()
If Grid.TextMatrix(mousRow, dcSour) = "" Then
    MsgBox "��������� ���� '������' � ������� ����������", , "��������������"
    Grid.col = dcSour
    On Error Resume Next
    Grid.SetFocus
    Exit Sub
End If
If Nomenklatura.isRegimLoad Then Unload Nomenklatura
Nomenklatura.Regim = "fromDocuments"
Nomenklatura.setRegim
Nomenklatura.Show vbModal
loadDocNomenk

Grid2.row = max(quantity2, 1)
Grid2_EnterCell
On Error Resume Next
Grid2.SetFocus

End Sub

Function backOstatki(strWhere) As Boolean
Dim str  As String

backOstatki = False
sql = "UPDATE sGuideNomenk INNER JOIN sDMC " & _
"ON sGuideNomenk.nomNom = sDMC.nomNom " & "SET sGuideNomenk.nowOstatki = " & _
"[sGuideNomenk].[nowOstatki]-[sDMC].[quant] " & strWhere

'��������� �. � �� ���� - ������� ����
If myExecute("##122", sql, 0) <= 0 Then backOstatki = True
End Function

Private Sub cmDel_Click()
'Dim strWhere As String


sql = "SELECT numDoc from sDMC WHERE (((numDoc)=" & numDoc & "));"
If Not byErrSqlGetValues("W##426", sql, tmpSng) Then GoTo EN2
If tmpSng <> 0 Then
    MsgBox "���� �� ������ ������� ���������, �� ������� ������� " & _
    "�� ��� ��� ��������.", , "�������� ����������!"
    GoTo EN2
End If

If MsgBox("������� �������� � '" & Grid.TextMatrix(mousRow, dcNumDoc) & _
"', �� �������?", vbYesNo Or vbDefaultButton2, "����������� ��������") _
= vbNo Then GoTo EN1

wrkDefault.BeginTrans

'strWhere = "WHERE (((sDMC.numDoc)=" & numDoc & ") AND ((sDMC.numExt)=" & numExt & "));"
''����������� ������� �� ����� - �� ��������� �. � �� ����
'If Not backOstatki(strWhere) Then
'    wrkDefault.Rollback
'    GoTo EN1
'End If

'�������� ���-�� (� ����� �����. ������� �� ��� - �.�. ��������� ��������� ��������)
sql = "DELETE  From sDocs WHERE (((sDocs.numDoc)=" & numDoc & _
      ") AND ((sDocs.numExt)=" & numExt & "));"
'MsgBox sql
If myExecute("##121", sql) = 0 Then
    quantity = quantity - 1
    If quantity = 0 Then
        clearGridRow Grid, mousRow
    Else
        Grid.RemoveItem mousRow
    End If
    wrkDefault.CommitTrans
Else
    wrkDefault.Rollback
End If
EN1:
Grid2.Visible = False
'cmProduct.Enabled = False
laGrid2.Visible = False
EN2:
Grid_EnterCell
On Error Resume Next
Grid.SetFocus
End Sub

Private Sub cmDel2_Click()
Dim strWhere As String

strWhere = "WHERE (((sDMC.numDoc)=" & numDoc & ") AND ((sDMC.numExt)=" & _
numExt & ") AND ((sDMC.nomNom)='" & gNomNom & "'));"

If MsgBox("������� ������� � '" & Grid2.TextMatrix(mousRow2, dnNomNom) & _
"', �� �������?", vbYesNo Or vbDefaultButton2, "����������� ��������") _
= vbNo Then GoTo EN1

wrkDefault.BeginTrans

'����������� ������� �� �����
If Not backOstatki(strWhere) Then
    wrkDefault.Rollback
    GoTo EN1
End If


sql = "DELETE  From sDMC  " & strWhere
If myExecute("##125", sql) = 0 Then
    quantity2 = quantity2 - 1
    If quantity2 = 0 Then
        clearGridRow Grid2, mousRow2
    Else
        Grid2.RemoveItem mousRow2
        Grid2_EnterCell
    End If
    wrkDefault.CommitTrans
Else
    wrkDefault.Rollback
End If
EN1:
On Error Resume Next
If quantity2 = 0 Then
    Grid.SetFocus
Else
    Grid2.SetFocus
End If

End Sub

Private Sub cmExit_Click()
frBad.Visible = False
End Sub

Private Sub cmKarta_Click()
lbBad_DblClick
End Sub

Private Sub cmLoad_Click()
laFiltr.Visible = False
loadDocs

End Sub

Function getNextNumExt() As Integer
Dim v As Variant

getNextNumExt = 0
sql = "SELECT Max(sDocs.numExt) AS Max_numExt From sDocs " & _
"WHERE (((sDocs.numDoc)=" & numDoc & "));"

If Not byErrSqlGetValues("##128", sql, v) Then Exit Function
If IsNumeric(v) Then
    getNextNumExt = v + 1
Else
    getNextNumExt = 1
End If

End Function


Private Sub cmProduct_Click()
'If Not docLock() Then Exit Sub
ReDim QQ(0)

Products.Regim = "select"
Products.Show vbModal
If UBound(QQ) > 0 Then ' ���-�� ��������
    If Not loadDocNomenk("check") Then
        backNomenk '�����
        loadDocNomenk
    End If
Else
    loadDocNomenk
End If
'docLock "un"

End Sub

Private Sub Command1_Click()
Dim str As String, str2 As String

'str = strWhereByStEndDateBox(Me)
str2 = getWhereByDateBoxes(Me, "sDocs.xDate", CDate("01.11.2000"))
If str = str2 Then
    str = str & "   - ���������"
Else
    str = str & "   - �� ��������� �  Where = '" & str2 & "'"
End If
MsgBox "Where = '" & str & "'"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF7 Then
    mnDocFind_Click
End If
End Sub

Private Sub Form_Load()
Dim str As String ', i As Integer
oldHeight = Me.Height
oldWidth = Me.Width
isLoad = True
If otlad = "otlaD" Then
    mnFilters.Visible = True
    Me.BackColor = otladColor
End If

Me.Caption = "�����. ��������� ���������.     " & mainTitle

If dostup = "a" Then
    mnOborot.Visible = True
    sourOborot.Visible = True
    ventureOborot.Visible = True
    mnFiltrOborot.Visible = True
    mnSep1.Visible = True
    mnKarta.Visible = True
End If

sql = "SELECT sGuideSource.sourceName From sGuideSource " & _
"WHERE (((sGuideSource.sourceId)>0)) ORDER BY sGuideSource.sourceName;"
Set table = myOpenRecordSet("##144", sql, dbOpenForwardOnly)
If table Is Nothing Then End
While Not table.EOF
    lbSource.AddItem table!SourceName
    table.MoveNext
Wend
table.Close
'lbSource.Height = 195 * lbSource.ListCount + 100

loadLbInside
initVentureLB

'Set wrkDefault = DBEngine.Workspaces(0) ' ��� ���-�� ����������

tbStartDate.Text = Format(DateAdd("d", -7, CurDate), "dd/mm/yy")
tbEndDate.Text = Format(CurDate, "dd/mm/yy")

Grid.FormatString = "|<����|<� ���-��|<������|<����|<������|<����������"
Grid.colWidth(0) = 0
Grid.colWidth(dcDate) = 800
Grid.colWidth(dcNumDoc) = 915
'Grid.ColWidth(dcM) = 300
Grid.colWidth(dcSour) = 1100
Grid.colWidth(dcDest) = 1100
Grid.colWidth(dcNote) = 1530
Grid.colWidth(dcVenture) = 800

Grid2.FormatString = "|<�����|<��������|���-��|<��.���������|���-��|<��.���.������������"
Grid2.colWidth(0) = 0
Grid2.colWidth(dnNomNom) = 0 '945
Grid2.colWidth(dnNomName) = 2400 + 430 + 650 + 945
Grid2.colWidth(dnEdIzm) = 0 '435
Grid2.colWidth(dnQuant) = 0 '660
Grid2.colWidth(dnEdIzm2) = 435
Grid2.colWidth(dnQuant2) = 660

End Sub

Sub loadLbInside()
Dim I As Integer

sql = "SELECT sGuideSource.sourceId, sGuideSource.sourceName From sGuideSource " & _
"WHERE (((sGuideSource.sourceId)<-1000)) ORDER BY sGuideSource.sourceId DESC;"
Set table = myOpenRecordSet("##95", sql, dbOpenDynaset)
If table Is Nothing Then myBase.Close: End
ReDim insideId(0): I = 0: ' ReDim statiaId(0): j = 0
lbInside.Clear
While Not table.EOF
    lbInside.AddItem table!SourceName
    ReDim Preserve insideId(I)
    insideId(I) = table!sourceId
    I = I + 1
    table.MoveNext
Wend
table.Close
lbInside.Height = lbInside.Height + 195 * (lbInside.ListCount - 1)
End Sub


Private Sub Form_Resize()
Dim h As Integer, w As Integer

If WindowState = vbMinimized Then Exit Sub
On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w / 2

Grid2.Height = Grid2.Height + h
Grid2.Width = Grid2.Width + w / 2
Grid2.left = Grid2.left + w / 2
cmLoad.Top = cmLoad.Top + h
cmAdd.Top = cmAdd.Top + h
cmDel.Top = cmDel.Top + h
cmAdd2.Top = cmAdd2.Top + h
cmProduct.Top = cmProduct.Top + h
cmDel2.Top = cmDel2.Top + h

End Sub

Private Sub Form_Unload(Cancel As Integer)
'tbSystem.Close
isLoad = False
If GuideSource.isLoad Then Unload GuideSource
If KartaDMC.isLoad Then Unload KartaDMC
If Nomenklatura.isRegimLoad Then Unload Nomenklatura
If Products.isLoad Then Unload Products
If VentureOrder.isLoad Then Unload VentureOrder

'myBase.Close
End Sub

Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If Grid.TextMatrix(mousRow, dcVenture) = "" Then
    cmAdd2.Enabled = False
    cmDel2.Enabled = False
Else
    cmAdd2.Enabled = True
    cmDel2.Enabled = True
End If
End Sub

Private Sub Grid_DblClick()
If mousRow = 0 Then Exit Sub
If Grid.CellBackColor = &H88FF88 Then
    If mousCol = dcSour Then
'        listBoxInGridCell lbGroup, Grid
        listBoxInGridCell lbSource, Grid, "select"
    ElseIf mousCol = dcDest Then
        listBoxInGridCell lbInside, Grid, "select"
    ElseIf mousCol = dcDate Then
        If MsgBox("��������� ���� ��������� ������ ������ ����� �������� " & _
        "���������� �������. ���� �� ������� � ������������� ��������� ���� " & _
        "������� <��>.", vbYesNo Or vbDefaultButton2, "����������� ��������� " & _
        "����!") = vbYes Then textBoxInGridCell tbMobile, Grid
    ElseIf mousCol = dcVenture Then
        listBoxInGridCell lbVenture, Grid, Grid.TextMatrix(mousRow, mousCol)
    Else
        tbMobile.MaxLength = 50
        textBoxInGridCell tbMobile, Grid
    End If
End If

End Sub

Function loadDocNomenk(Optional reg As String = "") As Boolean
Dim il As Long, str As String, s As Single, I As Integer ', str2 As String
Dim msgOst As String, r As Single, o As Single

loadDocNomenk = True ' �� ���� ������ - ����
msgOst = ""
Me.MousePointer = flexHourglass
Grid2.Visible = False

gDocDate = Grid.TextMatrix(mousRow, dcDate)
laGrid2.Caption = "������������ �� ��������� '" & numDoc & "'"
'Grid2.Clear

quantity2 = 0
clearGrid Grid2

sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.cod, " & _
"sGuideNomenk.Size, sGuideNomenk.ed_Izmer, sGuideNomenk.perList, " & _
"sGuideNomenk.ed_Izmer2,  sDMC.quant FROM sGuideNomenk INNER JOIN " & _
"(sDocs INNER JOIN sDMC ON (sDocs.numExt = sDMC.numExt) AND " & _
"(sDocs.numDoc = sDMC.numDoc)) ON sGuideNomenk.nomNom = sDMC.nomNom " & _
"WHERE (((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & ")) " & _
"ORDER BY sGuideNomenk.nomNom;"
'MsgBox sql
sum = 0
Set tbNomenk = myOpenRecordSet("##118", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then Exit Function
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    quantity2 = quantity2 + 1
    Grid2.TextMatrix(quantity2, dnNomNom) = tbNomenk!nomnom
    
    str = gNomNom
    If KartaDMC.DMCnomNomCur = tbNomenk!nomnom Then Grid2.row = quantity2  ' ���� ��������� ����� ���. ������������ - �� ��������� ��
    gNomNom = str                                                          ' ��� �������� Enter_Cell �.�. ������ gNomNom
    Grid2.TextMatrix(quantity2, dnNomName) = tbNomenk!cod & " " & _
            tbNomenk!nomName & " " & tbNomenk!Size
'    s = Round(tbNomenk!quant, 2)
'    Grid2.TextMatrix(quantity2, dnEdIzm) = tbNomenk!ed_Izmer
'    Grid2.TextMatrix(quantity2, dnQuant) = s
    If Grid.TextMatrix(Grid.row, 0) = -1002 Then
        Grid2.TextMatrix(quantity2, dnEdIzm2) = tbNomenk!ed_izmer
        Grid2.TextMatrix(quantity2, dnQuant2) = Round(tbNomenk!quant, 2)
    Else
        Grid2.TextMatrix(quantity2, dnEdIzm2) = tbNomenk!ed_Izmer2
        Grid2.TextMatrix(quantity2, dnQuant2) = Round(tbNomenk!quant / tbNomenk!perList, 2)
    End If
'    Grid2.TextMatrix(quantity2, dnQuant2) = Round(s / tbNomenk!perList, 2)
    Grid2.AddItem ""
    tbNomenk.MoveNext
  Wend
  Grid2.RemoveItem quantity2 + 1
End If
tbNomenk.Close
Grid2.Visible = True
Me.MousePointer = flexDefault
End Function

Private Sub Grid_EnterCell()
If quantity = 0 Or Not gridIsLoad Then Exit Sub
 mousRow = Grid.row
 mousCol = Grid.col

numDoc = Grid.TextMatrix(mousRow, dcNumDoc)
destId = Grid.TextMatrix(mousRow, 0)
If prevRow <> mousRow And gridIsLoad Then
    prevRow = mousRow
    loadDocNomenk
End If
If mousCol = 0 Then Exit Sub

If mousCol = dcDest And quantity2 <> 0 Then GoTo AA
If mousCol = dcSour And destId = -1002 Then GoTo AA

If mousCol = dcDate Or mousCol > dcNumDoc Then
    Grid.CellBackColor = &H88FF88
Else
AA:  Grid.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid_GotFocus()
    cmProduct.Visible = False
    cmDel2.Enabled = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Grid_DblClick
'ElseIf KeyCode = vbKeyEscape Then
'    lbHide
End If

End Sub

Sub lbHide()
tbMobile.Visible = False
lbGroup.Visible = False
lbSource.Visible = False
lbInside.Visible = False
lbVenture.Visible = False
Grid.Enabled = True
On Error Resume Next
Grid.SetFocus
Grid_EnterCell
End Sub

Sub lbHide2()
tbMobile2.Visible = False
Grid2.Enabled = True
On Error Resume Next
Grid2.SetFocus
Grid2_EnterCell
End Sub

Public Sub initVentureLB()
' ������� ������� ������ ��������
While lbVenture.ListCount
    lbVenture.RemoveItem (0)
Wend

sql = "select * from GuideVenture where standalone = 0 and id_analytic is not null"

Set table = myOpenRecordSet("##72", sql, dbOpenForwardOnly)
If table Is Nothing Then myBase.Close: End

'lbVenture.AddItem "", 0
While Not table.EOF
    lbVenture.AddItem "" & table!ventureName & ""
    lbVenture.ItemData(lbVenture.ListCount - 1) = table!ventureId
    table.MoveNext
Wend
table.Close
lbVenture.Height = 225 * lbVenture.ListCount

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
        MsgBox "ColWidth = " & Grid.colWidth(Grid.MouseCol)
End Sub

Private Sub Grid2_Click()
mousCol2 = Grid2.MouseCol
mousRow2 = Grid2.MouseRow
If quantity2 = 0 Then Exit Sub
If Grid2.MouseRow = 0 Then
    Grid2.CellBackColor = Grid2.BackColor
    If mousCol2 = dnQuant Then
        SortCol Grid2, mousCol2, "numeric"
    Else
        SortCol Grid2, mousCol2
    End If
    SortCol Grid2, mousCol2
    Grid2.row = 1    ' ������ ����� ����� ���������
'    Grid2_EnterCell
End If
Grid2_EnterCell
End Sub

Private Sub Grid2_DblClick()
If mousRow2 = 0 Then Exit Sub
If Grid2.CellBackColor = &H88FF88 Then
    textBoxInGridCell tbMobile2, Grid2
End If

End Sub

Private Sub Grid2_EnterCell()
If quantity2 = 0 Then Exit Sub
mousRow2 = Grid2.row
mousCol2 = Grid2.col

gNomNom = Grid2.TextMatrix(mousRow2, dnNomNom)

If mousCol2 = dnQuant2 Then
    Grid2.CellBackColor = &H88FF88
Else
    Grid2.CellBackColor = vbYellow
End If


End Sub

Private Sub Grid2_GotFocus()
'    cmAdd2.Visible = True
'    cmDel2.Visible = True
cmDel2.Enabled = (quantity2 > 0)
'If quantity2 > 0 Then
'    cmDel2.Enabled = True
'Else
'    cmDel2.Enabled = False
'End If
    
End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid2_DblClick

End Sub

Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyEscape Then Grid2.CellBackColor = Grid2.BackColor
If KeyCode = vbKeyEscape Then Grid2_EnterCell

End Sub

Private Sub Grid2_LeaveCell()
Grid2.CellBackColor = Grid2.BackColor

End Sub

Private Sub Grid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid2.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid2.colWidth(Grid2.MouseCol)
End Sub

Private Sub lbBad_DblClick()
Dim I As Integer
I = InStr(lbBad.Text, "  ")
gNomNom = left$(lbBad.Text, I - 1)
ReDim DMCnomNom(1)
DMCnomNom(1) = gNomNom
KartaDMC.Grid.Visible = False
KartaDMC.nomenkName = Mid$(lbBad.Text, I + 2)
KartaDMC.Show
'frBad.Visible = False
End Sub

Private Sub lbBad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbBad_DblClick
End Sub

Private Sub lbGroup_DblClick()
If lbGroup.ListIndex = 0 Then
    listBoxInGridCell lbSource, Grid
Else
    listBoxInGridCell lbInside, Grid
End If
End Sub

Private Sub lbGroup_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbGroup_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub lbInside_DblClick()
Dim id As String
Const INVENT = "��������������." ' ����� ��� ������� ������� ������ ������

id = insideId(lbInside.ListIndex)
If mousCol = dcSour Then
    If valueToDocsField("##96", id, "sourId") Then GoTo AA
Else
  If id = -1002 Then ' �� ����� �������� ������ �� ������� ��������������
     If Grid.TextMatrix(mousRow, dcSour) = "���������" Then
        GoTo BB '      � ���������
     Else
        sql = "UPDATE sDocs, sGuideSource SET sDocs.sourId = " & _
        "[sGuideSource].[sourceId], sDocs.destId = " & id & _
        " WHERE (((sGuideSource.sourceName)='" & INVENT & "') AND " & _
        "((numDoc)=" & numDoc & ") AND ((numExt)=" & numExt & "));"
        If myExecute("##96", sql) = 0 Then
            Grid.TextMatrix(mousRow, dcSour) = INVENT
            GoTo AA
        End If
    End If
  Else
BB: If valueToDocsField("##96", id, "destId") Then
AA:     Grid.Text = lbInside.Text
        Grid.TextMatrix(mousRow, 0) = id
    End If
  End If
End If
'If lbInside.Text = str2 Then
'    MsgBox "� �������� '������' � '����' ����������� ���������� ��������", , "��������������"
'    Exit Sub
'End If
lbHide


End Sub

Function valueToDocsField(myErrCod As String, value As String, field As String) As Boolean
sql = "UPDATE sDocs  SET sDocs." & field & "=" & value & _
" WHERE (((sDocs.numDoc)=" & numDoc & " AND (sDocs.numExt)=" & numExt & "));"
'MsgBox sql
valueToDocsField = False
If myExecute(myErrCod, sql) = 0 Then valueToDocsField = True
End Function

Private Sub lbInside_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbInside_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If


End Sub

Private Sub lbSource_DblClick()
sql = "UPDATE sDocs, sGuideSource SET sDocs.sourId = [sGuideSource].[sourceId] " & _
"WHERE (((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & _
") AND (([sGuideSource].[sourceName])='" & lbSource.Text & "'));"
'sql = "UPDATE sDocs, sGuideSource SET sDocs.sourId = sGuideSource.sourceId " & _
"WHERE (((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & _
") AND ((sGuideSource.sourceName)='" & lbSource.Text & "'));"
'MsgBox sql
If myExecute("##126", sql) = 0 Then Grid.Text = lbSource.Text
lbHide

End Sub

Private Sub lbSource_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbSource_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub


Private Sub lbVenture_DblClick()
Dim newNote As String, nCount As Integer

If lbVenture.Visible = False Then Exit Sub
sql = "select wf_make_venture_income(" & Grid.TextMatrix(mousRow, dcNumDoc) & ", " & lbVenture.ItemData(lbVenture.ListIndex) & ")"

'i = orderUpdate("##72", lbVenture.ItemData(lbVenture.ListIndex), "Orders", "ventureId")
byErrSqlGetValues "##126.1", sql, nCount
If nCount > 0 Then
    Grid.Text = lbVenture.Text
    newNote = getValueFromTable("sDocs", "Note", "numDoc = " & Grid.TextMatrix(mousRow, dcNumDoc))
    If IsNull(newNote) Then newNote = ""
    Grid.TextMatrix(mousRow, dcNote) = newNote
Else
    MsgBox "��������� �� ���������. ��������, ���� ������� ���������� ������, ������� ��� ������ �� ������ ������ �����������", , "���������������"
End If

lbHide


End Sub

Private Sub lbVenture_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lbVenture_DblClick
    ElseIf KeyCode = vbKeyEscape Then
        lbHide
    End If
End Sub

Private Sub mnCurOstat_Click()
Nomenklatura.Regim = "checkCurOstat"
Nomenklatura.Show
Nomenklatura.setRegim

End Sub

Private Sub mnDocFind_Click()
Static value

AA:     value = InputBox("������� ����� ��������� (������)", "�����", value)
        If value = "" Then Exit Sub
        If Not IsNumeric(value) Then
            MsgBox "����� ������ ���� ������"
            GoTo AA
        End If
laFiltr.Visible = False
numDoc = value
loadDocs "docsFind"
End Sub

Private Sub mnFiltrOborot_Click()
Nomenklatura.Regim = "fltOborot"
Nomenklatura.Show
Nomenklatura.setRegim

End Sub

Private Sub mnGuideFormuls_Click()
GuideFormuls.Regim = ""
GuideFormuls.Show
End Sub


Private Sub mnInside_Click()
GuideInside.Show vbModal
End Sub


Private Sub mnKarta_Click()
Nomenklatura.Regim = "forKartaDMC"
Nomenklatura.Show
Nomenklatura.setRegim
End Sub

Private Sub mnManag_Click()
GuideManag.Show vbModal
End Sub

Private Sub mnNomenc_Click()
Dim n1 As Nomenklatura
    Set n1 = New Nomenklatura
    n1.Regim = ""
    n1.Show
    n1.setRegim
End Sub

Private Sub mnOborot_Click()
Dim n1 As Nomenklatura
    Set n1 = New Nomenklatura
    n1.Regim = "asOborot"
    n1.Show
    n1.setRegim
End Sub
'�������� �� �������.�������
Private Sub mnOstat_Click()
Dim ost As Single, bef As Integer, I As Integer

frBad.Visible = False
Me.MousePointer = flexHourglass
'sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.ostCheck, " & _
      "sGuideNomenk.begOstatki  FROM sGuideNomenk;"
sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.ostCheck FROM sGuideNomenk;"
Set tbNomenk = myOpenRecordSet("##134", sql, dbOpenDynaset)
'Set tbNomenk = myOpenRecordSet("##134", "GuideNomenk", dbOpenTable)
If tbNomenk Is Nothing Then GoTo EN1
If tbNomenk.BOF Then GoTo EN1
'quantity = 0
tbNomenk.MoveFirst
bilo = False
While Not tbNomenk.EOF
  ost = 0 'tbNomenk!begOstatki
'  If ost < 0 Then GoTo BIL0
  tbNomenk.Edit
  tbNomenk!ostCheck = ""
  tbNomenk.Update
  
  sql = "SELECT sDMC.quant, sDocs.xDate, sDocs.sourId, sDocs.destId " & _
  "FROM sDocs INNER JOIN sDMC ON (sDocs.numExt = sDMC.numExt) AND " & _
  "(sDocs.numDoc = sDMC.numDoc) WHERE (((sDMC.nomNom)='" & tbNomenk!nomnom & "')) " & _
  " ORDER BY sDocs.xDate;"
  Set tbDMC = myOpenRecordSet("##135", sql, dbOpenForwardOnly)
  If tbDMC Is Nothing Then GoTo EN1
  If tbDMC.BOF Then GoTo NXT
'  tbDMC.MoveFirst
  bef = 0

  While Not tbDMC.EOF
    I = DateDiff("d", begDate, tbDMC!xDate)
    If ost <= -0.01 And I <> bef Then GoTo NXT2
    bef = I
    If tbDMC!sourId < -1000 Then _
        ost = ost - tbDMC!quant
    If tbDMC!destId < -1000 Then _
        ost = ost + tbDMC!quant
    tbDMC.MoveNext
  Wend
NXT:
  tbDMC.Close
NXT2:
  If ost <= -0.01 Then
'        tbDMC.Close
BIL0:   bilo = True
        tbNomenk.Edit
        tbNomenk!ostCheck = "m"
        tbNomenk.Update
  End If
  tbNomenk.MoveNext
Wend
tbNomenk.Close
valueToSystemField Now(), "checkOstDate"
'tbSystem.Edit
'tbSystem!checkOstDate = Now()
'tbSystem.Update
mnViewOst_Click
EN1:
Me.MousePointer = flexDefault

End Sub

Private Sub mnOstVed_Click()
Dim n1 As Nomenklatura
    Set n1 = New Nomenklatura
    n1.Regim = "asOstat"
    n1.Show
    n1.setRegim
End Sub

Private Sub mnPriceToExcel_Click()

    Products.PriceToExcel
End Sub

Private Sub mnProducts_Click()
Products.Regim = "" ' ������ ����������
Products.Show vbModal
End Sub

Private Sub mnReservedAll_Click()
    'Report.param1 = laOther.Caption
    Report.emptyColIndex = 1
    Report.groupIdColIndex = 0
    Report.subtitleColIndex = 2
    Report.numSortSecondColIndex = 0 ' �� ������ ������
    Report.numSortThirdColIndex = 2 ' �� �������� ������������
    Report.Subtitle = True
    
    Report.Regim = "reservedAll"
    Report.Sortable = True
    Set Report.Caller = Me
    Report.Show vbModal

End Sub

Private Sub mnSkladStand_Click()

    ReDim sqlRowDetail(1)
    ReDim aRowText(1)
    ReDim rowFormatting(1)
    ReDim aRowSortable(1)
    ReDim arowSubtitle(1)
    
    
    sqlRowDetail(1) = "call wf_nomenk_areport"
    aRowText(1) = " ������� ��������� ������"
    rowFormatting(1) = "#|<����� ���.|<��������|�� ���.|>����|>�-�� ����|>�-�� ����|>�����.����|>����� ����."
    aRowSortable(1) = True
    arowSubtitle(1) = True
    Set Report.Caller = Me
    Report.Regim = "aReportDetail"
    Report.param1 = 1
    
    Report.Show vbModal
    
    
End Sub

Private Sub mnSource_Click()
GuideSource.Show vbModal
End Sub

Private Sub mnStatia_Click()
GuideStatia.Show vbModal
End Sub

Private Sub mnToExcel_Click()
    ostatToWeb "toExcel"
End Sub

Private Sub mnVentureOrder_Click()
    VentureOrder.Show
End Sub

Private Sub mnViewOst_Click()

Me.MousePointer = flexHourglass
lbBad.Clear
laFrame.Caption = "������ ������������ � �������������� ���������" & _
vbCrLf & "(�������� �� " & Format(getSystemField("checkOstDate"), "dd.mm.yy hh:nn") & ")"

sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName From sGuideNomenk " & _
"WHERE (((sGuideNomenk.ostCheck)='m')) ORDER BY sGuideNomenk.nomNom;"
Set tbNomenk = myOpenRecordSet("##136", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If tbNomenk.BOF Then
    MsgBox "������� � �������������� ��������� �� ����������.", , "���������� ��������"
Else
    While Not tbNomenk.EOF
        lbBad.AddItem tbNomenk!nomnom & "  " & tbNomenk!nomName
        tbNomenk.MoveNext
    Wend
    tbNomenk.Close
    lbBad.ListIndex = 0
    frBad.Visible = True
End If
EN1:
Me.MousePointer = flexDefault
End Sub


Private Sub mnWeb_Click()
Me.MousePointer = flexHourglass
    
If MsgBox("�� ������ '��' ����� ����������� ���� �����c��� �������� ��� WEB." _
, vbDefaultButton2 Or vbYesNo, "����������� ������") = vbNo Then Exit Sub

ostatToWeb

Me.MousePointer = flexDefault
End Sub

Function setVertBorders(lineWeight As Long) As Integer
On Error GoTo ERR1
objExel.ActiveSheet.Cells(exRow, 1).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 2).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 3).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 4).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 5).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 6).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 7).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 8).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 9).Borders(xlEdgeRight).Weight = xlMedium
Exit Function

ERR1:
setVertBorders = Err

End Function

'������ Nomenks - !!! ��� ��������� �\� ������� ��� � � Prior
'��� ���� Nomenklatura.nomencDostupOstatki �������� �� sProducts.nomencOstatkiToGrid(-1)

'��� �������� ������ � ���� ��� � MS Excel ������� �� ������ �� ����
'������������.
'��� ������������ � ��� ������������� �� �������, ������� ��������
'����������� ��������� (��. �������������(�����)� ����������� ������������
'� ��������� stime)
'������������� ���������� �� ����. sGuideKlass � ���������� �� sGuideNomenk.
'�� � �� ���� klassId parentKlassId � klassName ���� �������� ��������� ��
'���� Comtec(����������� ������� ����� ��������) $comtec$
Sub ostatToWeb(Optional toExel As String = "")
Dim tmpFile As String, I As Integer, findId As Integer, str As String
Dim minusQuant   As Integer
minusQuant = 0

'�� ����������� ������������ �������� ������ Id ���� �����(�������),
' � ������� ���� ���� �� ���� ������������.
sql = "SELECT klassId from sGuideNomenk GROUP BY klassId;"
Set tbNomenk = myOpenRecordSet("##408", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Sub

'� ���� ����� �������� ����� ���� ����� ������(�� Id)====================
'Set tbGuide = myOpenRecordSet("##407", "select * from sGuideKlass", dbOpenForwardOnly)
'If tbGuide Is Nothing Then Exit Sub
'tbGuide.index = "PrimaryKey"

ReDim NN(0): I = 0
While Not tbNomenk.EOF
    I = I + 1
    ReDim Preserve NN(I): NN(I) = Format(tbNomenk!KlassId, "0000")
    findId = tbNomenk!KlassId

AA: 'tbGuide.Seek "=", findId
'    If tbGuide.NoMatch Then msgOfEnd ("##409")
    sql = "SELECT klassName, parentKlassId from sGuideKlass " & _
    "WHERE (((klassId)=" & findId & "));"
    If Not byErrSqlGetValues("##417", sql, str, findId) Then tbNomenk.Close: Exit Sub
            
'    NN(i) = tbGuide!klassName & " / " & NN(i) ' � ����� ��������� Id
    NN(I) = str & " / " & NN(I) ' � ����� ��������� Id
'    findId = tbGuide!parentKlassId
  
    If findId > 0 Then GoTo AA '� ����� ������� ������ ������� �������������
                               '����� ���� ����� ������, � ������� ��� ������
    tbNomenk.MoveNext
Wend
'tbGuide.Close
tbNomenk.Close
'=========================================================================

'���� ���� �� ������� ��������� -------------------------------------------

quickSort NN, 1 ' ��������� ����� �����

If toExel = "" Then
    On Error GoTo ERR1
    tmpFile = webNomenks & "tmp"
    Open tmpFile For Output As #1
Else
    On Error GoTo ERR2
    Set objExel = New Excel.Application
    objExel.Visible = True
    objExel.SheetsInNewWorkbook = 1
    objExel.Workbooks.Add
    objExel.ActiveSheet.Cells(1, 2).value = "������� �� ������ �� " & Format(Now(), "dd.mm.yy")
    objExel.ActiveSheet.Cells(1, 2).Font.Bold = True
    exRow = 4
    objExel.ActiveSheet.Cells(exRow - 1, 3).value = RateAsString
    objExel.ActiveSheet.Cells(exRow - 1, 5).value = "���� �������� ���"


    objExel.ActiveSheet.Columns(1).columnWidth = 12.57
    objExel.ActiveSheet.Columns(2).columnWidth = 39.71
    objExel.ActiveSheet.Columns(3).columnWidth = 10
    objExel.ActiveSheet.Columns(4).columnWidth = 6.2
    objExel.ActiveSheet.Columns(5).columnWidth = 6.2
    objExel.ActiveSheet.Columns(6).columnWidth = 7: objExel.ActiveSheet.Columns(6).HorizontalAlignment = xlHAlignRight
    objExel.ActiveSheet.Columns(7).columnWidth = 7: objExel.ActiveSheet.Columns(7).HorizontalAlignment = xlHAlignRight
    objExel.ActiveSheet.Columns(8).columnWidth = 7: objExel.ActiveSheet.Columns(8).HorizontalAlignment = xlHAlignRight
    objExel.ActiveSheet.Columns(9).columnWidth = 7: objExel.ActiveSheet.Columns(9).HorizontalAlignment = xlHAlignRight
    
    'cErr = setVertBorders(xlMedium)
'xlColumnDataType
    'If cErr <> 0 Then GoTo ERR2
'xlDiagonalDown, xlDiagonalUp, xlEdgeBottom, xlEdgeLeft, xlEdgeRight
'xlEdgeTop, xlInsideHorizontal, or xlInsideVertical.
    With objExel.ActiveSheet.Range("A" & exRow & ":I" & exRow)
        '.Borders(xlEdgeBottom).Weight = xlMedium ' xlThin
        '.Borders(xlEdgeTop).Weight = xlMedium
    End With
    'exRow = exRow + 1
End If
'------------------------------------------------------------------------


For I = 1 To UBound(NN) ' ������� ���� �����
  str = NN(I)
  findId = right$(str, 4) ' ��������� �� ���� ������ id ������
  
'$comtec$  ����� ������ �� ����.sGuideNomenk � �� �� ���� ���� �������� ��
'����������� �� ���� Comtec ������ �� ����.������������ � ���������
'����������� ����������� �� ��������� stime:
'�����  ���  �������� ������  ��.���������  ����.������������  CenaSale web
'nomNom cod  nomName  Size    ed_Izmer2     perList            CENA_W   Web
  sql = "SELECT n.nomNom, n.nomName, n.ed_Izmer2, n.CENA_W, n.perList" _
  & ", n.cod, n.Size, n.kodel, n.margin, n.kolonok" _
  & ", n.cenaOpt2, n.cenaOpt3, n.cenaOpt4" _
  & ", k.kolon1, k.kolon2, k.kolon3, k.kolon4" _
  & " From sGuideNomenk n " _
  & " join sguideklass k on k.klassId = n.klassId" _
  & " Where n.web = 'web'  AND n.klassId=" & findId _
  & " ORDER BY n.nomNom"

  Set tbNomenk = myOpenRecordSet("##331", sql, dbOpenDynaset)
  If tbNomenk Is Nothing Then GoTo EN1
  If Not tbNomenk.BOF Then
      bilo = False
      While Not tbNomenk.EOF
        gNomNom = tbNomenk!nomnom
        tmpSng = Nomenklatura.nomencDostupOstatki
        
'���� ���� �� ������� ��������� (����� �������� ��������� �����)------------
        If Not bilo Then
            bilo = True
            str = left$(str, Len(str) - 6)
            If toExel = "" Then
                str = "<b>" & str & "</b>"
'� ����� ���� ��� � ������ �� ��������� ������, �� � Web ��� �� ������ � �������
                Print #1, vbTab & str & vbTab & "<b>��.���</b>" & _
                vbTab & "<b>���-��</b>" & vbTab & "<b>����</b>" & _
                vbTab & "<b>���</b>" & vbTab & "<b>������</b>"
            Else
                objExel.ActiveSheet.Cells(exRow, 2).value = str
                objExel.ActiveSheet.Cells(exRow, 2).Font.Bold = True
                With objExel.ActiveSheet.Range("A" & exRow & ":I" & exRow)
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlEdgeBottom).Weight = xlThin
                    .Borders(xlEdgeRight).Weight = xlMedium
                End With
                
                exRow = exRow + 1
                cErr = setVertBorders(xlThin)
                'If cErr <> 0 Then GoTo ERR2
                
                objExel.ActiveSheet.Cells(exRow, 1).value = "���"
                objExel.ActiveSheet.Cells(exRow, 2).value = "��������"
                objExel.ActiveSheet.Cells(exRow, 3).value = "������"
                objExel.ActiveSheet.Cells(exRow, 4).value = "��.���."
                objExel.ActiveSheet.Cells(exRow, 5).value = "���-��"
                'objExel.ActiveSheet.Cells(exRow, 6).value = "���� ��"
                With objExel.ActiveSheet.Range("A" & exRow & ":I" & exRow)
                    .Borders(xlEdgeBottom).Weight = xlThin
                    .Font.Italic = True
                    .HorizontalAlignment = xlHAlignCenter
                End With
                If Not IsNull(tbNomenk!Kolon1) Then
                    objExel.ActiveSheet.Cells(exRow, 6).value = Chr(160) & tbNomenk!Kolon1
                    objExel.ActiveSheet.Cells(exRow, 6).Font.Bold = True
                End If
                If Not IsNull(tbNomenk!Kolon2) Then
                    objExel.ActiveSheet.Cells(exRow, 7).value = Chr(160) & tbNomenk!Kolon2
                    objExel.ActiveSheet.Cells(exRow, 7).Font.Bold = True
                End If
                If Not IsNull(tbNomenk!Kolon3) Then
                    objExel.ActiveSheet.Cells(exRow, 8).value = Chr(160) & tbNomenk!Kolon3
                    objExel.ActiveSheet.Cells(exRow, 8).Font.Bold = True
                End If
                If Not IsNull(tbNomenk!Kolon4) Then
                    objExel.ActiveSheet.Cells(exRow, 9).value = Chr(160) & tbNomenk!Kolon4
                    objExel.ActiveSheet.Cells(exRow, 9).Font.Bold = True
                End If
                    
                cErr = setVertBorders(xlThin)
                If cErr <> 0 Then GoTo ERR2
                exRow = exRow + 1
            End If
        End If
'---------------------------------------------------------------------------
'����� �������� ��������� �� ������ ������������ ������
        str = tbNomenk!ed_Izmer2
'        If str = "����" Or str = "�����" Then
'            tmpSng = tmpSng / tbNomenk!perList
'        End If
        tmpSng = Round(tmpSng - 0.4999, 0)
        Dim cena2W As String
        cena2W = Chr(160) & Format(tbNomenk!CENA_W, "0.00") ' ������� ��� �����, �.�. "3.00" ��� ����������� "3"
        If toExel = "" Then
            Print #1, tbNomenk!nomnom & vbTab & tbNomenk!nomName & vbTab & _
            str & vbTab & Round(tmpSng, 2) & vbTab & cena2W & _
            vbTab & tbNomenk!cod & vbTab & tbNomenk!Size
        Else
If tmpSng < -0.01 Then
    minusQuant = minusQuant + 1 '************************
End If
            objExel.ActiveSheet.Cells(exRow, 1).value = tbNomenk!cod
            objExel.ActiveSheet.Cells(exRow, 2).value = tbNomenk!nomName
            objExel.ActiveSheet.Cells(exRow, 3).value = tbNomenk!Size
            objExel.ActiveSheet.Cells(exRow, 4).value = str
            objExel.ActiveSheet.Cells(exRow, 5).value = Round(tmpSng, 2)
            objExel.ActiveSheet.Cells(exRow, 6).value = cena2W 'Round(tbNomenk!CENA_W, 2)
            'if not isnull
            
            Dim kolonok As Integer, optBasePrice As Double, margin As Double, iKolon As Integer, manualOpt As Boolean
            kolonok = tbNomenk!kolonok
            margin = tbNomenk!margin
            optBasePrice = cena2W * (1 - margin / 100)
            
            If kolonok > 0 Then
                manualOpt = False
            Else
                manualOpt = True
            End If
            
            For iKolon = 2 To Abs(kolonok)
                If manualOpt Then
                    objExel.ActiveSheet.Cells(exRow, 5 + iKolon).value = _
                        Chr(160) & Format(tbNomenk("CenaOpt" & CStr(iKolon)), "0.00")
                Else
                    objExel.ActiveSheet.Cells(exRow, 5 + iKolon).value = _
                        Chr(160) & Format(calcKolonValue(optBasePrice, margin, tbNomenk!kodel, Abs(kolonok), iKolon), "0.00")
                End If
            Next iKolon
            
            cErr = setVertBorders(xlThin)
            If cErr <> 0 Then GoTo ERR2
            exRow = exRow + 1:
        End If

        tbNomenk.MoveNext
      Wend
  End If
    tbNomenk.Close
'  End If
Next I
EN1:
If toExel = "" Then
    Close #1
    Kill webNomenks
    Name tmpFile As webNomenks
Else
    With objExel.ActiveSheet.Range("A" & exRow & ":I" & exRow)
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    Set objExel = Nothing
End If

If (minusQuant > 0) Then MsgBox "���������� " & minusQuant & _
" ������� � �������������� ���������.", , "��������������"
'End With
Exit Sub

ERR2:
Set objExel = Nothing
If cErr <> 424 And Err <> 424 Then GoTo ERR3 ' 424 - �� ��������� ����� ������ ������� ���-�
Exit Sub

ERR1:
If Err = 76 Then
    MsgBox "���������� ������� ���� " & tmpFile, , "Error: �� ��������� �� ��� ���� � �����"
ElseIf Err = 53 Then
    Resume Next ' ����� �.�� ����
ElseIf Err = 47 Then
    MsgBox "���������� ������� ���� " & tmpFile, , "Error: ��� ������� �� ������."
ElseIf cErr <> 424 Then
    cErr = Err
ERR3: MsgBox Error, , "������ 429-" & cErr '##429
    'End
End If

End Sub

Private Sub mnWebs_Click()
Dim str As String, ch As String, slen As Integer, oper As String, I As Integer
Dim tmpFile As String ', filtrList As String

If MsgBox("�� ������ '��' ����� ������������ ����� ��� WEB: ���� �����c��� " & _
"�������� � ���� ������������ ������� �������." _
, vbDefaultButton2 Or vbYesNo, "����������� ������") = vbNo Then Exit Sub


Me.MousePointer = flexHourglass

sql = "UPDATE sGuideNomenk SET sGuideNomenk.web2 = '';"
If myExecute("##405", sql) <> 0 Then GoTo EN2


sql = "SELECT sProducts.nomNom, sGuideProducts.prName, sGuideProducts.prId " & _
"FROM sGuideProducts INNER JOIN sProducts ON sGuideProducts.prId = " & _
"sProducts.ProductId WHERE (((sGuideProducts.web)<>''));"
'MsgBox sql
Set tbProduct = myOpenRecordSet("##373", sql, dbOpenDynaset)

If Not tbProduct Is Nothing Then
  
  If tbProduct.BOF Then
    MsgBox "�� ���� ������� �� �������� ��� Web", , "���� ����������� �� ������!"
  Else
    tmpFile = webProducts & "tmp"
    On Error GoTo ERR1
    Open tmpFile For Output As #1
'    On Error GoTo 0
    While Not tbProduct.EOF
'      sql = "UPDATE sGuideNomenk INNER JOIN sProducts ON sGuideNomenk.nomNom " & _
'      "= sProducts.nomNom   SET sGuideNomenk.web2 = 'web' " & _
'      "WHERE (((sProducts.ProductId)=" & tbProduct!prId & "));"
'      myExecute "##372", sql
    
      Print #1, tbProduct!prName & vbTab & tbProduct!nomnom
      tbProduct.MoveNext
    Wend
    Close #1
'    On Error Resume Next ' ����� �.�� ����
    Kill webProducts
'    On Error GoTo 0
    Name tmpFile As webProducts
  End If
  tbProduct.Close
End If

Documents.ostatToWeb '������ � �����
    
GoTo EN2
ERR1:
If Err = 76 Then
    MsgBox "���������� ������� ���� " & tmpFile, , "Error: �� ��������� �� ��� ���� � �����"
ElseIf Err = 53 Then
    Resume Next ' ����� �.�� ����
ElseIf Err = 47 Then
    MsgBox "���������� ������� ���� " & tmpFile, , "Error: ��� ������� �� ������."
Else
    MsgBox Error, , "������ 47-" & Err '##47
    'End
End If
EN2:
On Error Resume Next '�����, ���� ����� ����� ������� �������� ������� ����������
On Error Resume Next
Grid.SetFocus
Me.MousePointer = flexDefault

End Sub

Private Sub sourOborot_Click()
Dim n1 As Nomenklatura
    Set n1 = New Nomenklatura
    n1.Regim = "sourOborot"
    n1.Show
    n1.setRegim

End Sub

Private Sub tbMobile_DblClick()
lbHide
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, I As Integer

If KeyCode = vbKeyReturn Then
 
 
 If mousCol = dcDate Then
     If Not isDateTbox(tbMobile, "fry") Then Exit Sub
     str = "'" & Format(tmpDate, "yyyy-mm-dd") & "'"
     If tmpDate > CurDate Then
        MsgBox "���� ��������� �� ����� ���� � �������", , "������������ ��������!"
        GoTo EN1
     End If
     sql = "UPDATE sDocs  SET  sDocs.[xDate] = " & str & " WHERE " & _
     "(((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & "));"
     If myExecute("##119", sql) <> 0 Then GoTo EN1
 ElseIf mousCol = dcNote Then
    If Not valueToDocsField("##119", "'" & tbMobile.Text & "'", "Note") _
            Then GoTo EN1
 End If
 Grid.TextMatrix(mousRow, mousCol) = tbMobile.Text
 lbHide
ElseIf KeyCode = vbKeyEscape Then
    KeyCode = 0
' If mousCol = gpName And frmMode = "productAdd" Then
'    frmMode = ""
'    Grid.RemoveItem Grid.Rows - 1
' End If
EN1:
 lbHide
End If

End Sub

Private Sub tbMobile2_DblClick()
lbHide2

End Sub
Sub msgOfLateEdit(delta As Single)

   If delta < 0 And DateDiff("d", gDocDate, CurDate) <> 0 Then
        MsgBox "��������� ���������� ������ ������ ����� " & _
        "�������� � ��������� ������������� ��������.    " & _
        "������������� ��������� ����������� �������� ���� �� ������� " & _
        "'�������� �� �������' �� ���� '������'.", , "��������������"
    End If
End Sub

Private Sub tbMobile2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nowOst As Single, rezerv As Single, quant As Single, delta As Single
Dim I As Integer, j As Integer, tmp As Long

If KeyCode = vbKeyReturn Then
    If Not isNumericTbox(tbMobile2, 0) Then Exit Sub
    
'    sql = "SELECT sGuideNomenk.nowOstatki, sGuideNomenk.perList, " & _
    "sGuideNomenk.ed_Izmer, sGuideNomenk.ed_Izmer2, sDMC.quant " & _
    "FROM sGuideNomenk INNER JOIN sDMC ON sGuideNomenk.nomNom = sDMC.nomNom  " & _
    "WHERE (((sDMC.numDoc)=" & numDoc & ") AND ((sDMC.numExt)=" & numExt & _
    ") AND ((sGuideNomenk.nomNom)='" & gNomNom & "'));"
    sql = "SELECT nowOstatki, perList FROM sGuideNomenk " & _
    "WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
    'MsgBox sql
    Set tbNomenk = myOpenRecordSet("##123", sql, dbOpenForwardOnly)
    
    
    quant = tbMobile2.Text
    delta = Round(quant, 0)
    If Grid.TextMatrix(Grid.row, 0) <> -1002 Then
        If delta <> quant Then
            MsgBox "���������� ������ ���� �����", , "������"
            Exit Sub
        End If
        quant = Round(quant * tbNomenk!perList)
    End If
    sql = "SELECT quant FROM  sDMC  WHERE (((nomNom)='" & gNomNom & "') AND " & _
    "((numDoc)=" & numDoc & " AND (numExt)=" & numExt & "))"
'    MsgBox sql
    Set tbDMC = myOpenRecordSet("##458", sql, dbOpenForwardOnly)
    
    wrkDefault.BeginTrans
    
    delta = tbDMC!quant ' ��� �� ���
    delta = Round(quant - delta, 2)

    tbDMC.Edit
    tbDMC!quant = quant
    tbDMC.Update
    tbNomenk.Edit
    tbNomenk!nowOstatki = tbNomenk!nowOstatki + delta
    tbNomenk.Update
    
    wrkDefault.CommitTrans
    
    tbDMC.Close
    tbNomenk.Close
    
    msgOfLateEdit (delta)
    lbHide2
    tmp = Grid2.row
  
   loadDocNomenk


 '��������, ���� � loadDocNomenk ��� ����� �� ����-�� If Else �� ����
  If laFiltr.Visible Then   ' ���� ����� �� �����
'    If KartaDMC.DMCnomNomCur = gNomNom Then
    For I = 1 To UBound(DMCnomNom)
        If DMCnomNom(I) = gNomNom Then  ' � ���� ��������������� ���-��
            Timer1.Interval = 10        ' �� �����
            Timer1.Enabled = True       ' �� ���������� �����
            Exit For
        End If
    Next I
  Else '���� ��� �� ����� �� ����� �� �� ���������  ����� ��� �� ����������
    If KartaDMC.isLoad Then Unload KartaDMC '      ������������� ����������
  End If ' ���� ����� ��������� � ���� ���-�� ��� � ����� �� � �� ���� ���������
 
 Grid2.row = tmp
 Grid2.col = dnQuant2
EN1:
On Error Resume Next
 Grid2.SetFocus
ElseIf KeyCode = vbKeyEscape Then
    KeyCode = 0
    lbHide2
End If
End Sub

Private Sub Timer1_Timer()
Dim I As Integer
    Timer1.Enabled = False
    Me.MousePointer = flexHourglass
'    KartaDMC.Grid.Visible = False
    KartaDMC.quantity = 0
    For I = 1 To UBound(DMCnomNom)
        KartaDMC.getKartaDMC DMCnomNom(I)
    Next I
'    KartaDMC.Grid.Visible = True
    KartaDMC.ZOrder
    Me.MousePointer = flexDefault
End Sub

Private Sub ventureOborot_Click()
    Analityc.applicationType = "stime"
    Analityc.managId = AUTO.cbM.Text
    Analityc.Show
End Sub
