VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Orders 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   Caption         =   "��������������� ��������� �������"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "Orders.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lbVenture 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   5500
      TabIndex        =   41
      Top             =   1000
      Width           =   1095
   End
   Begin VB.CommandButton cmZagrSUB 
      Caption         =   "SUB"
      Height          =   315
      Left            =   6540
      TabIndex        =   40
      Top             =   5700
      Width           =   495
   End
   Begin VB.CommandButton cmCehSUB 
      Caption         =   "SUB"
      Height          =   315
      Left            =   6540
      TabIndex        =   39
      Top             =   5400
      Width           =   495
   End
   Begin VB.ListBox lbAnnul 
      Height          =   840
      ItemData        =   "Orders.frx":030A
      Left            =   240
      List            =   "Orders.frx":031A
      TabIndex        =   37
      Top             =   1980
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   10560
      Top             =   5340
   End
   Begin VB.TextBox tbEnable 
      BackColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   11460
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5460
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox tbInform 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   9015
   End
   Begin VB.ListBox lbClose 
      Height          =   450
      ItemData        =   "Orders.frx":0342
      Left            =   240
      List            =   "Orders.frx":034C
      TabIndex        =   35
      Top             =   3180
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lbTema 
      Height          =   2400
      Left            =   3720
      TabIndex        =   34
      Top             =   1020
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   0
      TabIndex        =   30
      Top             =   -75
      Width           =   11835
      Begin VB.CheckBox cbStartDate 
         Caption         =   " "
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox tbStartDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1260
         TabIndex        =   5
         Text            =   "01.09.02"
         Top             =   180
         Width           =   795
      End
      Begin VB.CheckBox cbEndDate 
         Caption         =   " "
         Height          =   315
         Left            =   2460
         TabIndex        =   6
         Top             =   180
         Width           =   315
      End
      Begin VB.CheckBox cbClose 
         Caption         =   "  "
         Height          =   195
         Left            =   5040
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cbM 
         Height          =   315
         ItemData        =   "Orders.frx":0365
         Left            =   11160
         List            =   "Orders.frx":0367
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   615
      End
      Begin VB.CheckBox chConflict 
         Caption         =   "  "
         Height          =   315
         Left            =   9240
         TabIndex        =   9
         Top             =   180
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox tbEndDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   180
         Width           =   795
      End
      Begin VB.Label laPeriod 
         Caption         =   "������ �  "
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   240
         Width           =   795
      End
      Begin VB.Label laPo 
         Caption         =   "���"
         Height          =   195
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   195
      End
      Begin VB.Label laClos 
         Caption         =   ",  � �. �. ��������"
         Height          =   195
         Left            =   3600
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label laConflict 
         Caption         =   "������������"
         Height          =   195
         Left            =   8040
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label laFiltr 
         Caption         =   "������� ������ !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5880
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
   End
   Begin VB.ListBox lbType 
      Height          =   1035
      ItemData        =   "Orders.frx":0369
      Left            =   1560
      List            =   "Orders.frx":037C
      TabIndex        =   29
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lbDel 
      Height          =   450
      ItemData        =   "Orders.frx":0391
      Left            =   240
      List            =   "Orders.frx":039B
      TabIndex        =   28
      Top             =   3900
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmExvel 
      Caption         =   "������ � Excel"
      Height          =   315
      Left            =   9660
      TabIndex        =   16
      Top             =   5580
      Width           =   1515
   End
   Begin VB.ListBox lbM 
      Height          =   255
      Left            =   1500
      TabIndex        =   27
      Top             =   1020
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmToWeb 
      Caption         =   "����� ��� WEB"
      Height          =   315
      Left            =   7920
      TabIndex        =   15
      Top             =   5580
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton cmZagrCO2 
      Caption         =   "CO2"
      Height          =   315
      Left            =   6060
      TabIndex        =   14
      Top             =   5700
      Width           =   495
   End
   Begin VB.CommandButton cmCehCO2 
      Caption         =   "CO2"
      Height          =   315
      Left            =   6060
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmZagrYAG 
      Caption         =   "YAG"
      Height          =   315
      Left            =   5580
      TabIndex        =   13
      Top             =   5700
      Width           =   495
   End
   Begin VB.CommandButton cmCehYAG 
      Caption         =   "YAG"
      Height          =   315
      Left            =   5580
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lbStat 
      Height          =   645
      ItemData        =   "Orders.frx":03B5
      Left            =   240
      List            =   "Orders.frx":03C2
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   240
      TabIndex        =   23
      Text            =   "tbMobile"
      Top             =   1620
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lbProblem 
      Height          =   2595
      Left            =   2460
      TabIndex        =   22
      Top             =   1020
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lbCeh 
      Height          =   645
      ItemData        =   "Orders.frx":03DD
      Left            =   2100
      List            =   "Orders.frx":03EA
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4455
      Left            =   0
      TabIndex        =   2
      Top             =   780
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7858
      _Version        =   393216
      BackColor       =   16777215
      ForeColorFixed  =   0
      BackColorSel    =   65535
      ForeColorSel    =   -2147483630
      GridLines       =   2
      AllowUserResizing=   1
      FormatString    =   $"Orders.frx":03FD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "��������"
      Height          =   315
      Left            =   3120
      TabIndex        =   10
      Top             =   5580
      Width           =   1275
   End
   Begin VB.CommandButton cmRefr 
      Caption         =   "���������"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   5580
      Width           =   975
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Label laZagruz 
      Caption         =   "��������:"
      Height          =   195
      Left            =   4680
      TabIndex        =   26
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label laCeh 
      Caption         =   "���.������:"
      Height          =   195
      Left            =   4680
      TabIndex        =   25
      Top             =   5460
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "��������:"
      Height          =   195
      Left            =   10320
      TabIndex        =   20
      Top             =   120
      Width           =   855
   End
   Begin VB.Label laInform 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1260
      TabIndex        =   19
      Top             =   5580
      Width           =   1575
   End
   Begin VB.Menu mnMenu 
      Caption         =   "����"
      Begin VB.Menu mnSetkaY 
         Caption         =   "����� ������� YAG                            F1"
      End
      Begin VB.Menu mnSetkaC 
         Caption         =   "����� ������� CO2                            F2"
      End
      Begin VB.Menu mnSetkaS 
         Caption         =   "����� ������� SUB                             F3"
      End
      Begin VB.Menu mnArhZone 
         Caption         =   "������ ��������� � ���������      F6"
      End
      Begin VB.Menu mnGuideFirms 
         Caption         =   "���������� ��������� ����������� F11"
      End
      Begin VB.Menu mnFirmFind 
         Caption         =   "����� ����� �� ��������               F12"
      End
      Begin VB.Menu mnReports 
         Caption         =   "������"
         Visible         =   0   'False
      End
      Begin VB.Menu mnSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "����� �� ���������                Alt F4"
      End
   End
   Begin VB.Menu mnMeassure 
      Caption         =   "���������"
      Begin VB.Menu mnPathSet 
         Caption         =   "��������� �����"
      End
      Begin VB.Menu mnComtexAdmin 
         Caption         =   "���������� � ������"
      End
   End
   Begin VB.Menu mnServic 
      Caption         =   "������"
      Begin VB.Menu mnWebs 
         Caption         =   "����� ��� Web"
      End
      Begin VB.Menu mnToExcel 
         Caption         =   "Web ����� � Excel"
      End
      Begin VB.Menu mnPriceToExcel 
         Caption         =   "Web ����� � Excel"
      End
   End
   Begin VB.Menu mnSklad 
      Caption         =   "�����"
      Begin VB.Menu mnNomenk 
         Caption         =   "������� �� ���-��    F4"
      End
      Begin VB.Menu mnProduct 
         Caption         =   "�� ���.  ��������"
      End
      Begin VB.Menu mnNaklad 
         Caption         =   "���������"
      End
   End
   Begin VB.Menu mnContext 
      Caption         =   "aa"
      Visible         =   0   'False
      Begin VB.Menu mnFirmsGuide 
         Caption         =   "���� � ���������� �����������"
      End
      Begin VB.Menu mnNoArhivFiltr 
         Caption         =   "������ ""������ � ���������"""
      End
      Begin VB.Menu mnSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnNoCloseFiltr 
         Caption         =   "������ ""���������� ������"""
         Visible         =   0   'False
      End
      Begin VB.Menu mnNoClose 
         Caption         =   "����� ""���������� ������"""
         Visible         =   0   'False
      End
      Begin VB.Menu mnAllOrders 
         Caption         =   "����� ""��� ������ �����"""
         Visible         =   0   'False
      End
      Begin VB.Menu mnBillFirma 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnQuickBill 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'$odbcXX$ - �������� �\�, ���. �������� (�� �������� OpenRecordSet, Index �
'           Seek) � ����������. (XX - ��� ���� ������)
'$odbc18!$- � �������� ���������

'$odbs?$  - �������� ������������ �����, ��� ���������� ����������
'$odbsE$  - ������� ���, ��� � ������ �������� ����� ���������� ��������� �������
'$NOodbc$ - �������� �\�, ���. �� ������� ��������� � ��� ����� �� �����

'$comtec$ - �������� �����, ���� ���� ������ x`�������� ���, ������� ���������
'���������(��. �����. �����) ���������������� ���������� ���� ������
Option Explicit

Public mousRow As Long
Public mousCol As Long
Public mousRow4 As Long
Public mousCol4 As Long
Public g_id_bill As String
Private loadBaseTimestamp As Date

Dim quantity4 As Long
Dim outDate() As Date
Dim tbUslug As Recordset
Dim strToWeb As String
Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����
Dim beClick As Boolean
Dim flDelRowInMobile As Boolean
Dim minut As Integer
Dim objExel As Excel.Application, exRow As Long
Dim head1 As String, head2 As String, head3 As String, head4 As String
Dim gain2 As Single, gain3 As Single, gain4 As Single



Const AddCaption = "��������"
Const t17_00 = 61200 ' � ��������

Const rowFromOrdersSQL = "select " & _
"    Orders.numOrder, GuideCeh.Ceh, Orders.inDate" & _
"   ,GuideManag.Manag, GuideStatus.Status, Orders.StatusId, GuideProblem.Problem" & _
"   ,Orders.DateRS, GuideFirms.Name, Orders.outDateTime, Orders.Type" & _
"   ,Orders.workTime, Orders.Logo, Orders.Product, Orders.ordered" & _
"   ,Orders.temaId, Orders.paid, Orders.shipped,  Orders.Invoice" & _
"   ,OrdersMO.DateTimeMO, OrdersMO.workTimeMO, OrdersMO.StatM, OrdersMO.StatO" & _
"   ,GuideManag_1.Manag AS lastManag, OrdersInCeh.urgent" & _
"   ,guideventure.venturename as venture" & _
"   ,lastModified" & _
"   ,id_bill" & _
"   ,GuideFirms.id_voc_names as id_voc_names" & _
"   ,guideventure.sysname as servername" & _
"   ,Orders.zalog, Orders.nal" & _
" from orders " & _
" JOIN GuideStatus ON GuideStatus.StatusId = Orders.StatusId " & _
" JOIN GuideProblem ON GuideProblem.ProblemId = Orders.ProblemId " & _
" JOIN GuideManag ON GuideManag.ManagId = Orders.ManagId " & _
" JOIN GuideFirms ON GuideFirms.FirmId = Orders.FirmId " & _
" JOIN GuideCeh ON GuideCeh.CehId = Orders.CehId " & _
" LEFT JOIN GuideManag AS GuideManag_1 ON Orders.lastManagId = GuideManag_1.ManagId " & _
" LEFT JOIN OrdersMO ON Orders.numOrder = OrdersMO.numOrder " & _
" LEFT JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder " & _
" left join guideventure on guideventure.ventureId = orders.ventureid "
    
    
Private Sub cbClose_Click()
cmRefr.Caption = "���������"
End Sub

Private Sub cbEndDate_Click()
cmRefr.Caption = "���������"
tbEndDate.Enabled = Not tbEndDate.Enabled
End Sub

Private Sub cbM_Click()
If zakazNum = 0 Then
    On Error Resume Next ' �.�. ������������� cbM �� Load
    cmRefr.SetFocus
Else
'If cbM.ListIndex > -1 Then cmAdd.Enabled = True
    lbHide
End If
cbM.TabStop = False
End Sub

Private Sub cbM_LostFocus()
If cbM.ListIndex < 0 Then
    MsgBox "��������� ���� '��������'", , "��������������"
    On Error Resume Next
    cbM.SetFocus
End If

End Sub

Private Sub cbStartDate_Click()
cmRefr.Caption = "���������"
tbStartDate.Enabled = Not tbStartDate.Enabled
End Sub
    
Sub begFiltrDisable()
    laPeriod.Enabled = False
    laPo.Enabled = False
    laClos.Enabled = False
    cbStartDate.Enabled = False
    tbStartDate.Enabled = False
    cbEndDate.Enabled = False
    tbEndDate.Enabled = False
    cbClose.Enabled = False
End Sub

Sub begFiltrEnable()
    laPeriod.Enabled = True
    laPo.Enabled = True
    laClos.Enabled = True
    cbStartDate.Enabled = True
    If cbStartDate.value = 1 Then tbStartDate.Enabled = True
    cbEndDate.Enabled = True
    If cbEndDate.value = 1 Then tbEndDate.Enabled = True
    cbClose.Enabled = True
End Sub

Private Sub chConflict_Click()
cmRefr.Caption = "���������"
If chConflict.value = 1 Then
    laConflict.ForeColor = vbRed
    begFiltrDisable
Else
    laConflict.ForeColor = vbBlack
    begFiltrEnable
End If
End Sub

Private Sub cmAdd_Click() ' �� ����� nextDayDetect()
Dim str As String, intNum As Integer, l As Long, numOrder As String
Dim strNow As String, DateFromNum As String, dNow As Date
 
 strNow = Format(Now, "dd.mm.yyyy")
 dNow = strNow
 strNow = right$(Format(Now, "yymmdd"), 5)
 
 wrkDefault.BeginTrans 'lock01
 sql = "update system set resursLock = resursLock" 'lock02
 myBase.Execute (sql) 'lock03

str = getSystemField("lastPrivatNum")
DateFromNum = Mid$(str, 4, 2) & "." & Mid$(str, 2, 2) & ".200" & left$(str, 1)
tmpDate = DateFromNum
DateFromNum = left$(str, 5)
intNum = right$(str, Len(str) - 5)

'���� ��� ��������� ����������� ����� �� �������, ��� ��������� ���� ���
    If tmpDate >= dNow Then
            intNum = intNum + 1
            If intNum < 100 Then
                str = Format(intNum, "00")
            Else
                str = Format(intNum, "000")
            End If
            l = DateFromNum & str
            'tbSystem!lastPrivatNum = DateFromNum & str
            myBase.Execute ("update system set lastPrivatNum = " & DateFromNum & str)
    Else        ' �������� ����. ����
        l = strNow & "01"
        'tbSystem!lastPrivatNum = strNow & "01"
        myBase.Execute ("update System set lastPrivatNum = " & strNow & "01")
        'tbSystem.Update
        befDays = DateDiff("d", tmpDate, Now)
        nextDay
        GoTo BB
    End If
'MsgBox " ����� ����� = " & strNow & str
 'tbSystem.Update
BB:
wrkDefault.CommitTrans
'tbSystem.Close
Dim baseCehId As Integer, baseCeh As String, isBaseOrder As Boolean
Dim baseFirmId As Integer, baseFirm As String
Dim baseProblemId As Integer, baseProblem As String, begPubNum As Long

gNzak = Grid.TextMatrix(Orders.mousRow, orNomZak)
If InStr(Orders.cmAdd.Caption, "+") > 0 Then
  sql = "SELECT Orders.CehId, Orders.ProblemId, Orders.FirmId, " & _
        "GuideCeh.Ceh, GuideProblem.Problem, GuideFirms.Name " & _
        "FROM GuideProblem INNER JOIN (GuideFirms INNER JOIN " & _
        "(GuideCeh INNER JOIN Orders ON GuideCeh.CehId = Orders.CehId) " & _
        "ON GuideFirms.FirmId = Orders.FirmId) ON GuideProblem.ProblemId " & _
        "= Orders.ProblemId WHERE (((Orders.numOrder)=" & gNzak & "));"
'  On Error GoTo NXT1
  Set tbOrders = myBase.OpenRecordset(sql, dbOpenForwardOnly)
  baseCehId = tbOrders!cehId
  baseFirmId = tbOrders!firmId
  baseProblemId = tbOrders!ProblemId
  baseCeh = tbOrders!Ceh
  baseFirm = tbOrders!name
  baseProblem = tbOrders!problem
  isBaseOrder = True
  tbOrders.Close
Else
  isBaseOrder = False
End If
NXT1:
cmAdd.Caption = AddCaption

'wrkDefault.BeginTrans
sql = "select * from Orders where numOrder = " & l
Set tbOrders = myOpenRecordSet("##07", sql, dbOpenForwardOnly) 'dbOpenForwardOnly
'If tbOrders Is Nothing Then Exit Sub

'If Not uniqOrderNum(tbOrders, l) Then
If Not tbOrders.BOF Then
    MsgBox "����� " & l & " �� �������� (��. ����� �� " _
    & table!inDate & ").  ��������� ������� ��� ���������� � ��������������!", , ""
    tbOrders.Close
    Exit Sub
End If

On Error GoTo ERR1
tbOrders.AddNew
tbOrders!StatusId = 0
tbOrders!numOrder = l
tbOrders!inDate = Now
tbOrders!ManagId = manId(Orders.cbM.ListIndex)
If isBaseOrder Then
  tbOrders!cehId = baseCehId
  tbOrders!firmId = baseFirmId
  tbOrders!ProblemId = baseProblemId
End If
tbOrders.update

If zakazNum > 0 Then Grid.AddItem ""
zakazNum = zakazNum + 1
Grid.TextMatrix(zakazNum, 0) = zakazNum
Grid.TextMatrix(zakazNum, orInvoice) = "���� ?"
Grid.TextMatrix(zakazNum, orNomZak) = l
Grid.TextMatrix(zakazNum, orData) = Format(Now, "dd.mm.yy")
Grid.TextMatrix(zakazNum, orMen) = Orders.cbM.Text
Grid.TextMatrix(zakazNum, orStatus) = status(0)
If isBaseOrder Then
  Grid.TextMatrix(zakazNum, orCeh) = baseCeh
  Grid.TextMatrix(zakazNum, orProblem) = baseProblem
  Grid.TextMatrix(zakazNum, orFirma) = baseFirm
End If
rowViem Grid.rows - 1, Grid
tbOrders.Close
Grid.row = zakazNum
Grid.col = orCeh
Grid.LeftCol = orNomZak
Grid.SetFocus
'wrkDefault.CommitTrans

Exit Sub
ERR1:
errorCodAndMsg "##419"

End Sub

'Function uniqOrderNum(table As Recordset, OrderNum As Long) As Boolean
'Dim str As String
    '
'    sql = "SELECT Orders.numOrder from Orders WHERE (((Orders.numOrder)= " & OrderNum & "));"
'    Set tbOrders = myOpenRecordSet("##402", sql, dbOpenForwardOnly)
    
    
'    If tbOrders.BOF Then
'    table.Seek "=", OrderNum
'    If table.NoMatch Then
'        uniqOrderNum = True
'    Else
'        uniqOrderNum = False
'            str = "���������� � ��������������!"
'        MsgBox "����� " & table!numOrder & " �� �������� (��. ����� �� " _
'                & table!inDate & "). " & str
'    End If
'    tbOrders.Close
'End Function

Private Sub cmCehCO2_Click()
If cehId <> 2 And isCehOrders Then Unload CehOrders
cehId = 2
CehOrders.Show 'vbModal

End Sub

Private Sub cmCehSUB_Click()
If cehId <> 3 And isCehOrders Then Unload CehOrders
cehId = 3
CehOrders.Show
End Sub

Private Sub cmCehYAG_Click()
If cehId <> 1 And isCehOrders Then Unload CehOrders
cehId = 1
CehOrders.Show 'vbModal
End Sub

Private Sub cmExvel_Click()
GridToExcel Grid
End Sub

Private Sub cmRefr_Click()
Dim minDate As Date, maxDate As Date

If chConflict.value = 0 Then
  begFiltrEnable
  If cbStartDate.value = 1 And cbEndDate.value = 1 Then
    minDate = tbStartDate.Text
    maxDate = tbEndDate.Text
    If minDate > maxDate Then
        MsgBox "������ ������� ������ ���� ������ �����", , "ERROR"
        Exit Sub
    End If
  End If
End If
beClick = False
Me.MousePointer = flexHourglass
begFiltr
LoadBase

Me.MousePointer = flexDefault
If chConflict.value = 1 And zakazNum = 0 Then _
    MsgBox "������������ ���", , "����������"
cmRefr.Caption = "��������"
laFiltr.Visible = False

End Sub

Sub valToWeb(val As Variant, Optional formatStr As String = "")
Dim chTab As String ', str As String

chTab = vbTab
If strToWeb = "" Then chTab = ""
If IsNull(val) Then
    strToWeb = strToWeb & chTab & Chr(160)
ElseIf val = "" Then
    strToWeb = strToWeb & chTab & Chr(160)
ElseIf formatStr <> "" Then
    strToWeb = strToWeb & chTab & Format(val, formatStr)
Else
    strToWeb = strToWeb & chTab & val
End If
End Sub

Sub openOrdersRowToGrid(myErr As String)

gNzak = Grid.TextMatrix(mousRow, orNomZak)
sql = rowFromOrdersSQL & " WHERE (((Orders.numOrder)= " & gNzak & " ));"
Set tqOrders = myOpenRecordSet("##57", sql, dbOpenForwardOnly)
If tqOrders Is Nothing Then myBase.Close: End
If tqOrders.BOF Then myBase.Close: End

copyRowToGrid mousRow
End Sub

Function isConflict(Optional msg As String = "") As Boolean
Dim problem As String, ordered, paid, shipped, stat As String, dateRS As Variant
Dim toClos As Boolean, titl As String, statM As String, statO As String

isConflict = False

Const ukagite = " ������� ��������� ��������!"
titl = "����� � " & gNzak & " � ��������������!"
  
problem = tqOrders!problem
ordered = tqOrders!ordered
paid = tqOrders!paid
shipped = tqOrders!shipped
stat = status(tqOrders!StatusId)

toClos = False
If msg = "toClose" Then msg = "": toClos = True

If stat = "������" Or stat = "��������" Then
  If Timer > t17_00 Then
    If DateDiff("d", tqOrders!dateRS, Now()) >= 0 Then
        isConflict = True
        If msg <> "" Then MsgBox "���������� ���� ��", , "����� � " & gNzak
    End If
  End If
ElseIf stat = "�����" Or toClos Then
    If msg = "msg" Then msg = "����� '�����' ��"
    GoTo EE
ElseIf stat = "�����������" And msg = "msg" Then
    msg = "�����"
EE:
  If IsNull(ordered) Then GoTo AA
  If Not IsNumeric(ordered) Then GoTo AA
  If ordered < 0.01 Then
AA: isConflict = True
    If msg <> "" Then MsgBox msg & " �� �������.", , titl
    Exit Function
  End If

  If IsNull(paid) Then GoTo BB
  If Not IsNumeric(paid) Then GoTo BB
  If ordered - paid > 0.01 Then
BB:
  If problem <> Problems(1) Then '������
    isConflict = True
'    If msg <> "" Then MsgBox "����� '�����' �� ����������." & ukagite, , titl
    If msg <> "" Then MsgBox msg & " ����������." & ukagite, , titl
  End If
  Exit Function
End If
    
If IsNull(shipped) Then GoTo ��
If Not IsNumeric(shipped) Then GoTo ��
If ordered - shipped > 0.01 Then
��:
  If problem <> Problems(4) Then '��������
    isConflict = True
    If msg <> "" Then MsgBox msg & " �� ��������� ��������." & ukagite, , titl
  End If
  Exit Function
End If
    
If paid - ordered > 0.01 Then
  If problem <> Problems(5) Then '���������
    isConflict = True
    If msg <> "" Then MsgBox msg & " �� ��������� ��������." & ukagite, , titl
  End If
  Exit Function
End If

If toClos Then Exit Function

If problem <> Problems(2) And problem <> Problems(16) Then '���-�� ��� ������
    isConflict = True
    If msg <> "" Then
       If problem = "" Then
            MsgBox "����� �� ������ ������� ���� �����. " & _
            "���� ��� �� ��� - " & ukagite, , titl
       Else
            MsgBox "��������� ���� ����� ������ � �� ����� �������� �� " & _
            "���.�����, ��� ���� ��������� ������ ��������: '" & _
            Problems(2) & "' � '" & Problems(16) & "'", , titl
       End If
    End If
End If

End If
End Function

Private Sub cmToWeb_Click()
Dim outDate As String, outTime As String, nbsp As String, tmpFile As String
Dim v As Variant

Me.MousePointer = flexHourglass
'Set myQuery = myBase.QueryDefs("wEbSvodka")
sql = "select * from wEbSvodka ORDER BY outDateTime"
'Set tqOrders = myOpenRecordSet("##46", myQuery.name, dbOpenForwardOnly) 'dbOpenDynaset)
Set tqOrders = myOpenRecordSet("##46", sql, dbOpenDynaset)
If tqOrders Is Nothing Then GoTo ENs
If Not tqOrders.BOF Then
  tmpFile = webSvodkaPath & "tmp"
  On Error GoTo ERR1
  Open tmpFile For Output As #1
  nbsp = "&" & "nbsp"
  tmpDate = Now
  outDate = Format(tmpDate, "dd.mm.yy")
  outTime = Format(tmpDate, "hh:nn")
  Print #1, outDate & nbsp & nbsp & nbsp & nbsp & nbsp & outTime
  Print #1, ""
  While Not tqOrders.EOF
      If isConflict() Then
        '�������� �����
        MsgBox "��������� ���� ���������� ������������, � ������ ����� " & _
        "�������� ������ ������ � ��������������. ����� ������������ �� " & _
        "����������� ������ ����� �������� �������� <Ctrl>+<I>.", , "���� �� �������!"
        chConflict.value = 1
        cmRefr_Click
        Close #1
        Kill tmpFile
        Exit Sub
      End If
    strToWeb = ""
    valToWeb tqOrders!xLogin
    valToWeb tqOrders!numOrder
    valToWeb status(tqOrders!StatusId)
    valToWeb tqOrders!outDateTime, "dd.mm.yy"
    valToWeb tqOrders!outDateTime, "hh"
    valToWeb tqOrders!problem
    valToWeb tqOrders!Logo
    valToWeb tqOrders!Product
    valToWeb tqOrders!ordered
    valToWeb tqOrders!paid
    valToWeb tqOrders!shipped
    valToWeb tqOrders!name
    valToWeb tqOrders!Manag
    valToWeb tqOrders!dateRS
    Print #1, strToWeb
    tqOrders.MoveNext
  Wend
  Close #1
End If
tqOrders.Close

'On Error Resume Next ' ����� �.�� ����
Kill webSvodkaPath
'On Error GoTo 0
Name tmpFile As webSvodkaPath

If chConflict.value = 1 Then
    MsgBox "������������ ���. ���� ������ ������.", , "����������:"
    chConflict.value = 0
End If

sql = "SELECT GuideFirms.xLogin, GuideFirms.Pass From GuideFirms " & _
"Where (((GuideFirms.xLogin) <> '')) ORDER BY GuideFirms.xLogin;"
'MsgBox sql
Set tbFirms = myOpenRecordSet("##80", sql, dbOpenDynaset)
If Not tbFirms Is Nothing Then
  tbFirms.MoveFirst
  If Not tbFirms.BOF Then
    tmpFile = webLoginsPath & "tmp"
    On Error GoTo ERR1
    Open tmpFile For Output As #1
'    On Error GoTo 0
    bilo = False
    While Not tbFirms.EOF
        If tbFirms!PASS = "" Then bilo = True
        Print #1, tbFirms!xLogin & vbTab & tbFirms!PASS & Chr(10); ';' - �������� ������� ����� ������
        tbFirms.MoveNext
    Wend
    Close #1
    If bilo Then
        MsgBox "� ����������� ��������� ����������� ��������� ������ ��� " & _
        "�������. ������� ���� �������-������� ��� WEB �� ����� ��������.", , "��������������"
    Else
'        On Error Resume Next ' ����� �.�� ����
        Kill webLoginsPath
'        On Error GoTo 0
        Name tmpFile As webLoginsPath
    End If
  End If
  tbFirms.Close
End If
ENs:
Me.MousePointer = flexDefault
Exit Sub

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
GoTo ENs
End Sub

Private Sub cmZagrCO2_Click()
cehId = 2
Zagruz.Show
End Sub

Private Sub cmZagrSUB_Click()
cehId = 3
Zagruz.Show

End Sub

Private Sub cmZagrYAG_Click()
cehId = 1
Zagruz.Show
End Sub

Sub lbHide(Optional noFocus As String = "")
tbMobile.Visible = False
lbCeh.Visible = False
lbStat.Visible = False
lbProblem.Visible = False
lbM.Visible = False
lbDel.Visible = False
lbType.Visible = False
lbTema.Visible = False
lbClose.Visible = False
lbAnnul.Visible = False
lbVenture.Visible = False

 

Grid.Enabled = True
If noFocus = "" Then
    Grid.SetFocus
    Grid_EnterCell
End If
End Sub

Private Sub Form_Activate()
Static beStart As Boolean

On Error Resume Next '�.�. ������-�� ���������� �� ����� �������� ��
'FindFirm �  GuideFirms
If beStart Then Orders.Grid.SetFocus
beStart = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, value As String, I As Integer, il As Long

If cbM.ListIndex < 0 Then
    'cbM_LostFocus
    Exit Sub
End If

If LCase(tbEnable.Text) <> "arh" And LCase(tbEnable.Text) <> "���" _
And tbEnable.Visible Then Exit Sub
If KeyCode = vbKeyEscape Then
    cmAdd.Caption = AddCaption
    lbHide
ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL Then
    tbEnable.Text = ""
    tbEnable.Visible = True
    tbEnable.SetFocus
ElseIf KeyCode = vbKeyF1 Then
    mnSetkaY_Click
ElseIf KeyCode = vbKeyF2 Then
    mnSetkaC_Click
ElseIf KeyCode = vbKeyF3 Then 'ceh$$
    mnSetkaS_Click
ElseIf KeyCode = vbKeyF6 And tbEnable.Visible Then
    mnArhZone_Click
ElseIf KeyCode = vbKeyF4 Then
    mnNomenk_Click '�� ����������� hotkey � ����, �.�. cbM_LostFocus
ElseIf KeyCode = vbKeyF5 Then
    cmAdd_Click
ElseIf KeyCode = vbKeyF7 Then
    If mousCol = orNomZak Then
        value = ""
AA:     value = InputBox("������� ����� ������", "�����", value)
        If value = "" Then Exit Sub
        If Not IsNumeric(value) Then
            MsgBox "����� ������ ���� ������"
            GoTo AA
        End If
        If findValInCol(Grid, value, orNomZak) Then Exit Sub
        If MsgBox("��������� ����� ������ �� ���� ����?", vbYesNo, _
        "����� ����������� ����� �� ������!") = vbNo Then Exit Sub
        For I = 1 To orColNumber
            orSqlWhere(I) = ""
        Next I
        loadWithFiltr value
        Grid_EnterCell '��������� ���� �������
    ElseIf mousCol = orFirma Then
        value = Grid.TextMatrix(mousRow, orFirma)
        value = InputBox("������� ������ �������� ��� ��������.", "����� � ������� '�������� �����'", value)
        If value = "" Then Exit Sub
        If findExValInCol(Grid, value, orFirma) > 0 Then Exit Sub
        If MsgBox("��������� ����������� ����� ����� '" & value & "' ?", vbYesNo, _
        "����� ����������� ����� ���� ����� �� ������!") = vbNo Then Exit Sub
        If tbEnable.Visible Then
            FindFirm.cmAllOrders.Visible = True
            FindFirm.cmNoClose.Visible = True
            FindFirm.cmNoCloseFiltr.Visible = True
        End If
        FindFirm.tb.Text = value
        FindFirm.Show vbModal
'    ElseIf mousCol = orIzdelia Or mousCol = orLogo Then
    Else
        value = Grid.TextMatrix(mousRow, mousCol)
        value = InputBox("������� ������� ������.", "�����", value)
        If findExValInCol(Grid, value, CInt(mousCol)) > 0 Then Exit Sub
        MsgBox "�������� �� ������"
'    Else
'        MsgBox "�� ����� ���� ����� �� ������������", , "��������������"
    End If
ElseIf KeyCode = vbKeyF11 Then
    mnGuideFirms_Click '�� ����������� hotkey � ����, �.�. cbM_LostFocus
ElseIf KeyCode = vbKeyF12 Then
    mnFirmFind_Click
ElseIf KeyCode = vbKeyMenu Then
    If cmAdd.Enabled And beClick And cmAdd.Caption = AddCaption Then _
                    cmAdd.Caption = AddCaption & " +"
ElseIf KeyCode = vbKeyI And Shift = vbCtrlMask Then
    If zakazNum < 1 Then Exit Sub
    openOrdersRowToGrid "##55"
    bilo = isConflict("msg")
    tqOrders.Close
    If bilo Then Exit Sub
    MsgBox "� ���� ������ ������������ ���", , "����� � " & gNzak
ElseIf KeyCode = vbKeyF And Shift = vbCtrlMask Then
    Filtr.cmReset_Click
    GoTo BB
ElseIf KeyCode = vbKeyB And Shift = vbCtrlMask Then
    Filtr.cmReset_Click
    listBoxSelectByText Filtr.lbM, Grid.TextMatrix(mousRow, orMen)
    listBoxSelectByText Filtr.lbStatus, Grid.TextMatrix(mousRow, orStatus)
    str = Grid.TextMatrix(mousRow, orFirma)
    If str <> "" Then
        Filtr.lbFirm.AddItem str, 0
        Filtr.lbFirm.Selected(0) = True
    End If
BB:
    If left$(Filtr.cmAdvan.Caption, 1) = "�" Then Filtr.cmAdvan_Click
    Filtr.lbStatus.Clear
    For I = 0 To 7 ' ������� �. ����������
       If tbEnable.Visible Or I <> 6 Then Filtr.lbStatus.AddItem status(I)
    Next I
    Filtr.laEnable.Visible = tbEnable.Visible
    Filtr.Show
End If
End Sub

Sub loadWithFiltr(Optional nomZak As String = "")
If IsNumeric(nomZak) Then ' ����� ������ �� ����
    orSqlWhere(0) = "" '���-�� ������ � �������� �������
    orSqlWhere(orNomZak) = strWhereByValCol(nomZak, orNomZak)
ElseIf nomZak = "" Then
    orSqlWhere(0) = ""
    orSqlWhere(mousCol) = strWhereByValCol(Grid.Text, CInt(mousCol))
    If orSqlWhere(mousCol) = "" Then Exit Sub ' � ���� ���� �� ������������ ������
End If
Me.MousePointer = flexHourglass
laFiltr.Visible = True
LoadBase
cmRefr.Caption = "���������"
Me.MousePointer = flexDefault
orSqlWhere(0) = "" '���-�� ���������� (��� �������� �������)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyMenu Then cmAdd.Caption = AddCaption

End Sub

Private Sub Form_Load()
Dim I As Integer, str As String

If tbEnable.Visible Then mnAllOrders.Visible = True

oldHeight = Me.Height
oldWidth = Me.Width


If Not IsEmpty(otlad) Then
    Frame1.BackColor = otladColor
    Me.BackColor = otladColor

    mnReports.Visible = True
    tbEnable.Visible = True
    tbEnable.Text = "arh"
    cmToWeb.Visible = True
End If
If dostup <> "a" Then mnPathSet.Visible = False '$$2
If dostup = "a" Or dostup = "b" Then
    mnNaklad.Visible = True
Else
    mnNaklad.Visible = False
End If

#If Not COMTEC = 1 Then '---------------------------------------------------
    mnServic.Visible = False
#Else
    mnNaklad.Visible = False
#End If '-------------------------------------------------------------------

beClick = False
flDelRowInMobile = False
Me.Caption = Me.Caption & mainTitle
orColNumber = 0
mousCol = 1
initOrCol orNomZak, "nOrders.numOrder"
initOrCol orInvoice, "sOrders.Invoice"
initOrCol orVenture, "sOrders.ventureName"
initOrCol orCeh, "sGuideCeh.Ceh"
initOrCol orData, "dOrders.inDate"
initOrCol orMen, "sGuideManag.Manag"
initOrCol orStatus, "sGuideStatus.Status"
initOrCol orProblem, "sGuideProblem.Problem"
initOrCol orDataRS, "dOrders.DateRS"
initOrCol orFirma, "sGuideFirms.Name"
initOrCol orDataVid, "dOrders.outDateTime"
initOrCol orVrVid
initOrCol orVrVip, "nOrders.workTime"
initOrCol orM
initOrCol orO
initOrCol orMOData, "dOrdersMO.DateTimeMO"
initOrCol orMOVrVid
initOrCol orOVrVip, "dOrdersMO.workTimeMO"
initOrCol orLogo, "sOrders.Logo"
initOrCol orIzdelia, "sOrders.Product"
initOrCol orType, "sOrders.Type"
initOrCol orTema, "nOrders.temaId"
initOrCol orZalog, "nOrders.zalog"
initOrCol orNal, "nOrders.nal"
initOrCol orZakazano, "nOrders.ordered"
initOrCol orOplacheno, "nOrders.paid"
initOrCol orOtgrugeno, "nOrders.shipped"
initOrCol orLastMen, "sGuideManag_1.Manag"
initOrCol orlastModified, "dOrders.lastModified"
initOrCol orBillId, "nOrders.id_bill"
initOrCol orVocnameId, "nOrders.id_voc_names"
initOrCol orServername, "sOrders.servername"

ReDim Preserve orSqlWhere(orColNumber)

laCeh.Visible = True
cmCehYAG.Visible = True
cmCehCO2.Visible = True
cmCehSUB.Visible = True '$$ceh

zakazNum = 0
tbStartDate.Text = Format(DateAdd("d", -7, curDate), "dd/mm/yy")
tbEndDate.Text = Format(curDate, "dd/mm/yy")

Grid.FormatString = "|>� ������|>� �����|<������| ��� |^���� |^ �|<������ |<��������|" & _
"<������|<�������� �����|<���� ������|��.������|��.����������|�����|�������|" & _
"<���� ������ MO|<��.������ MO|O �.����������|<����|<�������|" & _
"���������|<����|�����|���.���.|��������|�����������|���������|^ M"
Grid.Cols = Grid.Cols + 4 ' lastModified, id_bill, id_voc_names, servername
Grid.ColWidth(0) = 0
Grid.ColWidth(orData) = 840
Grid.ColWidth(orDataVid) = 975
Grid.ColWidth(orVrVid) = 330
Grid.ColWidth(orVrVip) = 750
Grid.ColWidth(orO) = 720
Grid.ColWidth(orMOData) = 795 + 50
Grid.ColWidth(orMOVrVid) = 570 + 50
Grid.ColWidth(orOVrVip) = 810
Grid.ColWidth(orZalog) = 540
Grid.ColWidth(orNal) = 540
Grid.ColWidth(orZakazano) = 540
Grid.ColWidth(orOplacheno) = 540
Grid.ColWidth(orOtgrugeno) = 615
Grid.ColWidth(orType) = 450
'Grid.ColWidth(orVenture) = 650
Grid.ColWidth(orlastModified) = 0
Grid.ColWidth(orBillId) = 0
Grid.ColWidth(orVocnameId) = 0
Grid.ColWidth(orServername) = 0

'*********************************************************************$$7
managLoad '�������� Manag() cbM lbM � Filtr.lbM

lbM.Height = lbM.Height + 195 * (lbM.ListCount - 1)
Filtr.lbM.Height = Filtr.lbM.Height + 195 * (Filtr.lbM.ListCount - 1)

If Not IsEmpty(otlad) Then cbM.ListIndex = cbM.ListCount - 1



Set table = myOpenRecordSet("##72", "GuideTema", dbOpenForwardOnly)
If table Is Nothing Then myBase.Close: End

I = 0
While Not table.EOF
    lbTema.AddItem table!Tema, table!temaId
    Filtr.lbTema.AddItem table!Tema, table!temaId
    table.MoveNext
Wend
table.Close

For I = 0 To lenProblem
    If Problems(I) <> "no" Then lbProblem.AddItem Problems(I)
Next I

isOrders = True
trigger = True
initVentureLB

End Sub
Public Sub initVentureLB()
' ������� ������� ������ ��������
While lbVenture.ListCount
    lbVenture.removeItem (0)
Wend

sql = "select * from GuideVenture where standalone = 0"

Set table = myOpenRecordSet("##72", sql, dbOpenForwardOnly)
If table Is Nothing Then myBase.Close: End

lbVenture.AddItem "", 0
While Not table.EOF
    lbVenture.AddItem "" & table!ventureName & ""
    lbVenture.ItemData(lbVenture.ListCount - 1) = table!ventureId
    table.MoveNext
Wend
table.Close
lbVenture.Height = 225 * lbVenture.ListCount

End Sub
 
Public Sub managLoad(Optional fromCeh As String = "")
Dim I As Integer, str As String

sql = "SELECT * From GuideManag  ORDER BY forSort;"
Set table = myOpenRecordSet("##03", sql, dbOpenForwardOnly)
If table Is Nothing Then myBase.Close: End
I = 0: ReDim manId(0):
Dim imax As Integer: imax = 0: ReDim Manag(0)
While Not table.EOF
    str = table!Manag
    If str = "not" Then
        GoTo AA
    ElseIf LCase(table!forSort) <> "unused" Then
        If fromCeh = "" Then
          If table!ManagId <> 0 Then cbM.AddItem str
          lbM.AddItem str
          Filtr.lbM.AddItem str
        End If
        manId(I) = table!ManagId
        I = I + 1
        ReDim Preserve manId(I):
AA:     If imax < table!ManagId Then
            imax = table!ManagId
            ReDim Preserve Manag(imax)
        End If
        Manag(table!ManagId) = str
    End If
    table.MoveNext
Wend
table.Close

End Sub
 

Sub begFiltr() '******* ��������� ������
Dim stDate As String, enDate As String, I As Integer
Dim addNullDate As String, strWhere As String
 
 For I = 1 To orColNumber
    orSqlWhere(I) = ""
 Next I
 
If chConflict.value = 1 Then '  ******************************
    orSqlWhere(orStatus) = "(Orders.StatusId)=4" '�����
    If Timer > t17_00 Then
       orSqlWhere(orStatus) = orSqlWhere(orStatus) & ") OR (" & _
       "(Orders.StatusId)=2) OR ((Orders.StatusId)=3" '������ ��������
    End If
Else                         '********************************
 
 If cbStartDate.value = 1 Then
    stDate = "(Orders.inDate)>='" & _
             Format(Orders.tbStartDate.Text, "yyyy-mm-dd") & "'"
    addNullDate = ""
 Else
    stDate = ""
    addNullDate = " OR (Orders.inDate) Is Null"
 End If

 If cbEndDate.value = 1 Then
    enDate = "(Orders.inDate)<='" & _
            Format(Orders.tbEndDate.Text, "yyyy-mm-dd") & " 11:59:59 PM'"
 Else
    enDate = ""
 End If
 If stDate <> "" And enDate <> "" Then
    strWhere = stDate & ") AND( " & enDate
 ElseIf stDate <> "" Or enDate <> "" Then
    strWhere = stDate & enDate
 Else
    addNullDate = ""
    strWhere = ""
 End If
 orSqlWhere(orData) = strWhere & addNullDate
 
 If cbClose.value = 0 Or Not tbEnable.Visible Then
    orSqlWhere(orStatus) = "(Orders.StatusId)<>6" '������
 Else
    orSqlWhere(orStatus) = ""
 End If
 
 getWhereInvoice

End If 'chConflict.value      ********************************
 
End Sub
Sub getWhereInvoice()
 If Not tbEnable.Visible Then
    orSqlWhere(orInvoice) = "isNumeric(Orders.Invoice)=1 OR (Orders.shipped) Is Null"
 End If
End Sub
Private Sub Form_Resize()
Dim h As Integer, w As Integer
lbHide "noFocus"


If Me.WindowState = vbMinimized Then Exit Sub

On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w
cmRefr.Top = cmRefr.Top + h
laInform.Top = laInform.Top + h
cmAdd.Top = cmAdd.Top + h
cmToWeb.Top = cmToWeb.Top + h
laCeh.Top = laCeh.Top + h
cmCehYAG.Top = cmCehYAG.Top + h
cmCehCO2.Top = cmCehCO2.Top + h
cmCehSUB.Top = cmCehSUB.Top + h '$$ceh
laZagruz.Top = laZagruz.Top + h
cmZagrYAG.Top = cmZagrYAG.Top + h
cmZagrCO2.Top = cmZagrCO2.Top + h
cmZagrSUB.Top = cmZagrSUB.Top + h '$$ceh
cmExvel.Top = cmExvel.Top + h
tbEnable.Top = tbEnable.Top + h
tbEnable.left = tbEnable.left + w
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Filtr
isOrders = False
exitAll
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

Private Sub Grid_Click()
If noClick Then Exit Sub
'laInform.Caption = laInform.Caption & "   cRow=" & Grid.row & "  cCol=" & Grid.col
If zakazNum = 0 Then Exit Sub
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow

If mousRow = 0 Then
    Grid.CellBackColor = Grid.BackColor
    If mousCol = 0 Then Exit Sub
    If mousCol = orNomZak Or mousCol = orZakazano Or mousCol = orOplacheno _
    Or mousCol = orOtgrugeno Or mousCol = orVrVip Or mousCol = orOVrVip _
    Or mousCol = orZalog Or mousCol = orNal Then
        SortCol Grid, mousCol, "numeric"
    ElseIf mousCol = orData Or mousCol = orDataRS Or mousCol = orDataVid Then
        SortCol Grid, mousCol, "date"
    Else
        SortCol Grid, mousCol
    End If
    Grid.row = 1    ' ������ ����� ����� ���������
End If
Grid_EnterCell
End Sub
    
Sub GuideFirmOnOff()
Dim tmpRow As Long, tmpCol As Long
    GuideFirms.Show vbModal
    Orders.SetFocus
End Sub

Function haveUslugi() As Boolean
Dim s As Single

End Function
Function stopOrderAtVenture() As Boolean
'    If ((mousCol <> orZakazano And mousCol <> orVenture And Grid.TextMatrix(mousRow, orZakazano) = "") Or Not isVentureGreen) Then
    stopOrderAtVenture = False
    If Not isVentureGreen Or Grid.TextMatrix(mousRow, orVenture) <> "" Or mousCol = orVenture Then Exit Function
    If mousCol <> orFirma And Grid.TextMatrix(mousRow, orZakazano) <> "" And (mousCol <> orZakazano) Then
        stopOrderAtVenture = True
    End If
End Function




Function checkInvoiceMerge(p_numOrder As String, p_newInvoice As String) As Integer
Dim ret As Integer

    sql = "select wf_check_jscet_merge (" & p_numOrder & ", '" & p_newInvoice & "')"
On Error GoTo sqle
    byErrSqlGetValues "##100.2", sql, checkInvoiceMerge
    If checkInvoiceMerge < 0 Then
        MsgBox "��� ����������� ������� � ���� ���� ���������, ����� �����-�������� � ����������� ���� ����������" _
        & vbCr & "��������� ��� ���� � ���������� ��� ���", , "������ ������������ �����"
        wrkDefault.rollback
    End If
    
    Exit Function
sqle:
    wrkDefault.rollback
    errorCodAndMsg "checkInvoiceMerge"
End Function


Function checkInvoiceSplit(p_numOrder As String, p_newInvoice As String) As Integer
    sql = "select wf_check_jscet_split (" & p_numOrder & ")"
On Error GoTo sqle
    byErrSqlGetValues "##100.1", sql, checkInvoiceSplit
    Exit Function
sqle:
    errorCodAndMsg "checkInvoiceSplit"
End Function


Function tryInvoiceMove(p_numOrder As String, p_Invoice As String, id_jscet_new As Integer, p_newInvoice As String) As Boolean
Dim mText As String
    tryInvoiceMove = True
On Error GoTo sqle
    mText = "�����������, ��� �� ������ " _
        & "��������� ����� �� ����� " & p_Invoice & " � ���� " & p_newInvoice
    sql = "call wf_move_jscet (" & p_numOrder & ", " & CStr(id_jscet_new) & ")"
    Debug.Print sql
    If MsgBox(mText, vbOKCancel, "�� �������?") = vbOK Then
        myBase.Execute sql
    Else
        wrkDefault.rollback
        tryInvoiceMove = False
    End If
    Exit Function
sqle:
    wrkDefault.rollback
    errorCodAndMsg "tryInvoiceMove"
    tryInvoiceMove = False
End Function


Function tryInvoiceSplit(p_numOrder As String, p_Invoice As String) As Boolean
Dim mText As String
    
    tryInvoiceSplit = True
On Error GoTo sqle
    mText = "�����������, ��� �� ������ " _
        & "�������� ����� �� ����� " & p_Invoice & " � ��������� ����"
    If MsgBox(mText, vbOKCancel, "�� �������?") = vbOK Then
        sql = "call wf_split_jscet (" & p_numOrder & ")"
        myBase.Execute sql
    Else
        wrkDefault.rollback
        tryInvoiceSplit = False
    End If
    Exit Function
sqle:
    wrkDefault.rollback
    errorCodAndMsg "tryInvoiceSplit"
    tryInvoiceSplit = False
End Function


Function tryInvoiceMerge(p_numOrder As String, id_jscet_new As Integer, p_newInvoice As String) As Boolean
Dim mText As String
    
    tryInvoiceMerge = True
On Error GoTo sqle
    If id_jscet_new > 0 Then
        If MsgBox("�����������, ��� �� ������ ������������ �������� ������ � ����� " & p_newInvoice, vbOKCancel, "�� �������?") = vbOK Then
            sql = "call wf_merge_jscet (" & p_numOrder & ", " & CStr(id_jscet_new) & ", " & p_newInvoice & ")"
            Debug.Print sql
            myBase.Execute sql
        Else
            wrkDefault.rollback
            tryInvoiceMerge = False
        End If
    End If
    Exit Function
sqle:
    wrkDefault.rollback
    errorCodAndMsg "tryInvoiceSplit"
    tryInvoiceMerge = False
    
End Function
Function OrderIsMerged() As Boolean
Dim exists As Integer

    OrderIsMerged = False
    sql = "select count(*) from orders o" _
        & " join guideventure v on o.ventureid = v.ventureid" _
        & " where statusid < 6 " _
        & " and v.venturename = '" & Grid.TextMatrix(mousRow, orVenture) & "'" _
        & " and invoice = '" & Grid.TextMatrix(mousRow, orInvoice) & "'" _
        & " and numorder != " & Grid.TextMatrix(mousRow, orNomZak)
'        Debug.Print sql
        
    byErrSqlGetValues "##OrderIsMerged", sql, exists
    If exists > 0 Then
        OrderIsMerged = True
    End If
    
End Function


'$odbc08!$
Private Sub Grid_DblClick()
Dim str As String, statId As Integer, s As Single
Dim orderTimestamp As Date
Dim lastManager As String
Dim strDate As String
Dim billCompany As String
Dim I As Integer


If zakazNum = 0 Then Exit Sub
If mousRow = 0 Then Exit Sub

gNzak = Grid.TextMatrix(mousRow, orNomZak)
sql = "SELECT O.CehId, O.StatusId, O.lastModified, m.Manag From Orders o " _
& " join GuideManag m on o.lastManagId = m.managid " _
& " WHERE (((O.numOrder)=" & gNzak & "));"
'Debug.Print (sql)
If Not byErrSqlGetValues("##174", sql, cehId, statId, orderTimestamp, lastManager) Then Exit Sub

If mousCol = orVrVip Then
    If dostup = "a" And statId = 4 Then
      If MsgBox("��� ������������������� ��������� ������� ����������! " & _
      " ���� �� ������� ������� '��'.", vbYesNo Or vbDefaultButton2, _
      "����� � " & gNzak) = vbYes Then textBoxInGridCell tbMobile, Grid
    End If
ElseIf mousCol = orNomZak Then
#If COMTEC = 1 Then '---------------------------------------------------
    MsgBox "�������� ������ ������ � �����x", , ""
#Else
  If statId = 7 Then
    MsgBox "� ������ � ������ �������� �� ����� ���� ���������!", , "��������������"
    Exit Sub
  End If
  
'  If Grid.CellForeColor = 200 Or Grid.CellForeColor = vbBlue Then
  tmpStr = ""
  If havePredmetiNew Then
    str = "����������"
  Else
    If statId > 3 Then
        MsgBox "� ����� ������ ��� ���������!", , ""
        Exit Sub
    End If
    str = Grid.TextMatrix(mousRow, orZakazano)
    If Not IsNumeric(str) Then GoTo AA
    If CSng(str) > 0 Then
        str = "(=" & str & ") "
    Else
AA:     str = ""
    End If
    
    tmpStr = Grid.TextMatrix(mousRow, orOtgrugeno)
    If Not IsNumeric(tmpStr) Then GoTo BB
    If CSng(tmpStr) > 0 Then
        tmpStr = "(=" & tmpStr & ") "
    Else
BB:     tmpStr = ""
    End If
    
    If str <> "" Or tmpStr <> "" Then
       tmpStr = vbCrLf & vbCrLf & "��������! � ���������� ���� �������� " & _
       "����� ������������� �������� ��� ������� '��������'" & str & _
       " � '���������'" & tmpStr & ".  ����� ����� ������� " & _
       "������ ���������� ��������."
    End If
    str = "������������"
  End If
  If MsgBox("�� ������ " & str & " �������� � ������? " & tmpStr, _
  vbYesNo Or vbDefaultButton2, "����� � " & gNzak) = vbYes Then
     sql = "DELETE From xUslugOut WHERE (((numOrder)=" & gNzak & "));"
     'Debug.Print sql
     myExecute "##304", sql, 0 '������� ���� ����
        
    If statId = 6 Then
      sProducts.Regim = "closeZakaz"
    Else
      sProducts.Regim = ""
    End If
    numDoc = gNzak
    numExt = 0 ' ��� ���� ��� �����. �\�, ��� ����� ������� ������ ��������� �������
    sProducts.Show vbModal
  End If

#End If '-------------------------------------------------------------------
  Exit Sub
End If

If Grid.CellBackColor = vbYellow Then Exit Sub

If stopOrderAtVenture Then
    MsgBox "����� ���, ��� ���-�� ������� � �������, ����� ������� �����������, ����� ������� �� ����� �����������", , "����"
    Exit Sub
End If
strDate = Grid.TextMatrix(mousRow, orlastModified)
If strDate <> "" Then
    loadBaseTimestamp = CDate(Grid.TextMatrix(mousRow, orlastModified))
Else
    loadBaseTimestamp = CDate(0)
End If
    

If orderTimestamp > loadBaseTimestamp And lastManager <> cbM.Text Then
    MsgBox "����� ����, ��� �� ��������� ���������� � ������, �� ��� ������� ���������� " _
    & lastManager & " � " & orderTimestamp & "." _
    & vbCr & "�������� ������ � ���������� ��������� �������� �����." _
     , , "����"
    Exit Sub
End If
If mousCol = orVenture Then
    If Grid.TextMatrix(mousRow, orVenture) <> "" Then
        ' ���������, ���� ����� ������ � ���� ������ � ������, �� �� ��������� ���� ������� ����
        If OrderIsMerged Then
            MsgBox "����� ������ � ������ �����, � ������� ������ ������ ������" _
                & vbCr & "���������� ������� �������� ����� � ��������� ���� � ��� ����� ������ ����� ������ �����������" _
                , , "������ �������� �����������"
            Exit Sub
        End If
    End If
     listBoxInGridCell lbVenture, Grid, Grid.TextMatrix(mousRow, orVenture)
ElseIf mousCol = orFirma Then
    
    If Grid.TextMatrix(mousRow, orVenture) <> "" Then
        ' ���������, ���� ����� ������ � ���� ������ � ������, �� �� ��������� ���� ������� ����
        
        billCompany = "����������"
    
        If Grid.CellForeColor = vbRed Then
            sql = "select wf_retrieve_bill_company(" + Grid.TextMatrix(mousRow, orBillId) + ", '" + Grid.TextMatrix(mousRow, orVenture) + "')"
'            Debug.Print sql
            If byErrSqlGetValues("W##102.1", sql, billCompany) Then
                mnBillFirma.Tag = Grid.TextMatrix(mousRow, orBillId)
            End If
            If billCompany = "" Then
                billCompany = "Id = [" & Grid.TextMatrix(mousRow, orBillId) & "]"
            End If
        Else
            mnBillFirma.Tag = ""
        End If
        
        mnBillFirma.Visible = True
        mnBillFirma.Caption = "����������: " + billCompany
        
        For I = mnQuickBill.UBound To 1 Step -1
            Unload mnQuickBill(I)
        Next I
        
        If serverIsAccessible(Grid.TextMatrix(mousRow, orVenture)) Then
        
            sql = _
                 " select o.id_bill, max(o.inDate) as lastDate " _
                & " from orders o" _
                & " join orders z on z.firmid = o.firmid and z.ventureid = o.ventureid and z.numorder = " & gNzak _
                & " where " _
                & "     o.id_bill is not null " _
                & " group by o.id_bill" _
                & " order by lastDate desc"
                  
            
            Set tbOrders = myOpenRecordSet("##102.2", sql, 0)
            If Not tbOrders.BOF Then
    '            Load mnQuickBill(0)
    '            mnQuickBill(0).Caption = "-"
                I = 0
                While Not tbOrders.EOF
                    If CStr(tbOrders!id_bill) <> Grid.TextMatrix(mousRow, orBillId) Then
                        mnQuickBill(0).Visible = True
                        Load mnQuickBill(1 + I)
                        mnQuickBill(I + 1).Tag = tbOrders!id_bill
                        sql = "select wf_retrieve_bill_company(" + CStr(tbOrders!id_bill) + ", '" + Grid.TextMatrix(mousRow, orVenture) + "')"
                        byErrSqlGetValues "W##102.1", sql, billCompany
                        mnQuickBill(I + 1).Caption = billCompany
                        I = I + 1
                    End If
                    tbOrders.MoveNext
                Wend
                tbOrders.Close
            End If
        End If
        If I = 0 Then
            mnQuickBill(0).Visible = False
        End If
        
'        success = byErrSqlGetValues("##102.2", sql, lastBillCompany)
        
    Else
        mnBillFirma.Visible = False
        mnQuickBill(0).Visible = False
        For I = mnQuickBill.UBound To 1 Step -1
            Unload mnQuickBill(I)
        Next I
    End If
    
    Me.PopupMenu mnContext
ElseIf mousCol = orCeh Then
    ' ���� �� ���������
#If Not COMTEC = 1 Then '---------------------------------------------------
    sql = "SELECT sDocs.numDoc From sDocs WHERE (((sDocs.numDoc)=" & gNzak & "));"
    If Not byErrSqlGetValues("W##175", sql, numDoc) Then Exit Sub
    If numDoc > 0 Then
        MsgBox "�� ����� ������ �������� ���������.", , "��������� ���� �����������!"
        Exit Sub
    End If
#End If '-------------------------------------------------------------------
    listBoxInGridCell lbCeh, Grid
ElseIf mousCol = orStatus Then
'$odbs?$ � ���� ����� �� ������ ����� �����.������, =========================
'������������� �� � ������� ��������� �� �����.
'(��������� ����������� �������� MsgBox)

    wrkDefault.BeginTrans 'lock01
'    sql = "update system set resursLock = resursLock" 'lock02
    sql = "UPDATE Orders set rowLock = rowLock where numOrder = " & gNzak 'lock02
    myBase.Execute (sql) 'lock03 ���������
    
    sql = "SELECT rowLock, StatusId, workTime FROM Orders WHERE numOrder = " & gNzak
    Set tbOrders = myOpenRecordSet("##29", sql, dbOpenForwardOnly)
'    If tbOrders Is Nothing Then Exit Sub
    'If Not findZakazInTable("Orders", "msg") Then GoTo ExCl
    If tbOrders.BOF Then
'       tbOrders.Update ' ������� ����������
       wrkDefault.CommitTrans ' ������� ����������
       tbOrders.Close
       MsgBox "�������� �� ��� ������. �������� ������", , "����� �� ������!!!"
       Exit Sub
    End If
    str = tbOrders!rowLock
    If str <> "" And str <> Orders.cbM.Text Then
'       tbOrders.Update ' ������� ����������
       wrkDefault.CommitTrans ' ������� ����������
       tbOrders.Close
       MsgBox "����� " & gNzak & " �������� ����� ������ ���������� (" & str & ")"
       Exit Sub
    End If
    tbOrders.Edit
    tbOrders!rowLock = Orders.cbM.Text
    tbOrders.update ' ������� ����������
    statId = tbOrders!StatusId
    If Not IsNull(tbOrders!workTime) Then _
        neVipolnen = tbOrders!workTime
    wrkDefault.CommitTrans ' ������� ����������
    tbOrders.Close
    
 ' ����� ����� ==============================================================
   
   If statId = 4 Then ' "�����"
     If dostup = "a" Then GoTo ALL
     listBoxInGridCell lbStat, Grid, "select"
   ElseIf statId = 6 Then ' "������"
     GoTo ALL '���� ������ ��� dostup='a', �.�. ��� ������ - ���� ������
   ElseIf statId = 7 Then ' "�����������"
     listBoxInGridCell lbDel, Grid, "select"
   ElseIf Grid.TextMatrix(mousRow, orCeh) <> "" Then
        If statId = 1 Then '� ������                                 $$1
          sql = "SELECT Nevip from OrdersInCeh WHERE (((numOrder)=" & gNzak & "));"
'          Set tbCeh = myOpenRecordSet("##373", "select * from OrdersInCeh", dbOpenForwardOnly)
          Set tbCeh = myOpenRecordSet("##373", sql, dbOpenForwardOnly)
'          If tbCeh Is Nothing Then Exit Sub '$$1
'          tbCeh.index = "PrimaryKey"        '
'          tbCeh.Seek "=", gNzak             '
'          If tbCeh.NoMatch Then             '
           If tbCeh.BOF Then
            neVipolnen = 0
            tbCeh.Close
            MsgBox "������ � " & gNzak & " ��� � ������� ������ ������� " & _
            "������� ������ ��������� �������. �������� ������ ���� " & _
            "���������� � ��������������.", , "Error"             '
            GoTo ALL                              '
          Else
              neVipolnen = Round(neVipolnen * tbCeh!nevip, 2)   '$$1
              tbCeh.Close
          End If                                  '
        End If                                    '
        If startParams Then
            Zakaz.Show vbModal
            refreshTimestamp gNzak
        Else
            msgOfZakaz ("##19")
        End If
   Else
     If dostup <> "a" Then
        listBoxInGridCell lbClose, Grid, "select"
     Else
ALL:    listBoxInGridCell lbAnnul, Grid, "select"
     End If
   End If
   ValueToTableField "##19", "", "Orders", "rowLock"
   Exit Sub

ElseIf mousCol = orMen Then
    listBoxInGridCell lbM, Grid
ElseIf mousCol = orProblem Then
    listBoxInGridCell lbProblem, Grid ', Grid.TextMatrix(mousRow, mousCol)
ElseIf mousCol = orType Then
    listBoxInGridCell lbType, Grid
ElseIf mousCol = orTema Then
    listBoxInGridCell lbTema, Grid
ElseIf mousCol = orVrVid Or mousCol = orMOVrVid Or mousCol = orLogo _
Or mousCol = orIzdelia Or mousCol = orType Or mousCol = orInvoice Then
    textBoxInGridCell tbMobile, Grid
#If Not COMTEC = 1 Then '----------------------------------------------------
ElseIf mousCol = orOplacheno Or mousCol = orZalog Or mousCol = orNal Then
    textBoxInGridCell tbMobile, Grid
ElseIf mousCol = orZakazano Then
  If havePredmetiNew Then
    MsgBox "�������� � ���� ���� �� �������������, �.�. � ������ ���� " & _
    "�������� (��� ��������� �������� �� ���� '� ������')", , "��������������"
    Exit Sub
  Else
    textBoxInGridCell tbMobile, Grid
  End If
ElseIf mousCol = orOtgrugeno Then
    If IsNumeric(Grid.TextMatrix(mousRow, orInvoice)) Or _
    Grid.TextMatrix(mousRow, orStatus) = "������" Then
        textBoxOrOtgruzFrm
    ElseIf MsgBox("", vbYesNo, "���� ?") = vbYes Then
        Grid.col = orInvoice
        Grid.LeftCol = orInvoice
        Grid.SetFocus
    Else
        If (Grid.TextMatrix(mousRow, orCeh) = "" Or _
        Grid.TextMatrix(mousRow, orStatus) = "�����") And _
        Grid.TextMatrix(mousRow, orInvoice) = "���� ?" Then ' � 2� ������
            flDelRowInMobile = Not tbEnable.Visible '�������� �����, ���� �� �� � ���. ����
            textBoxOrOtgruzFrm
        Else
            MsgBox "��� ��������� ������� ����� ������ ��������� ��� �����", , "������"
        End If
    End If
#End If '-------------------------------------------------------------------
End If


End Sub
Private Function isVentureGreen() As Boolean
Dim item_exists As Boolean, I As Integer

    isVentureGreen = False
    If Grid.TextMatrix(mousRow, orInvoice) <> "���� ?" And Grid.TextMatrix(mousRow, orVenture) = "" Then Exit Function
    If Grid.TextMatrix(mousRow, orOtgrugeno) <> "" Then Exit Function
    item_exists = False
    For I = 1 To lbVenture.ListCount - 1
        If (lbVenture.List(I) = Grid.TextMatrix(mousRow, orVenture)) Then
            item_exists = True
        End If
    Next I
    If Not item_exists And Grid.TextMatrix(mousRow, orVenture) <> "" Then Exit Function

    isVentureGreen = True
    
End Function
Public Sub Grid_EnterCell()
If noClick Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col
If mousCol = orFirma And Grid.CellForeColor = vbRed Then
    mousCol = mousCol
End If

flDelRowInMobile = False
If zakazNum = 0 Then Exit Sub
beClick = True
tbInform.Text = Grid.TextMatrix(mousRow, mousCol)

bilo = (mousCol = orZakazano Or mousCol = orOplacheno Or mousCol = orOtgrugeno Or mousCol = orZalog Or mousCol = orNal)
If (dostup = "a" Or Grid.TextMatrix(mousRow, orStatus) <> "������") _
   And ( _
       mousCol = orFirma _
       Or mousCol = orProblem _
       Or mousCol = orType _
       Or (mousCol = orCeh And Grid.TextMatrix(mousRow, orStatus) = "������") _
       Or mousCol = orMen _
       Or mousCol = orVrVid _
       Or mousCol = orStatus _
       Or (mousCol = orMOVrVid And (Grid.TextMatrix(mousRow, orM) <> "" Or Grid.TextMatrix(mousRow, orO) <> "")) _
       Or mousCol = orLogo _
       Or mousCol = orIzdelia _
       Or bilo _
       Or (mousCol = orTema And Grid.TextMatrix(mousRow, orType) = "�") _
       Or (mousCol = orInvoice And (dostup = "b" Or Grid.TextMatrix(mousRow, orVenture) = "" Or Grid.TextMatrix(mousRow, orMen) = cbM.Text)) _
       Or (mousCol = orVenture And isVentureGreen) _
   ) _
Then
        Grid.CellBackColor = &H88FF88
    If mousCol = orVrVid Or mousCol = orMOVrVid _
    Or mousCol = orLogo Or mousCol = orIzdelia Or mousCol = orOplacheno Then
        tbInform.Locked = False
    Else
        tbInform.Locked = True
    End If
Else
    Grid.CellBackColor = vbYellow
    tbInform.Locked = True
End If

End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    If mousCol = orFirma Then
        If Grid.TextMatrix(mousRow, orVenture) <> "" Then
            ' ���������, ���� ����� ������ � ���� ������ � ������, �� �� ��������� ���� ������� ����
            If OrderIsMerged Then
                MsgBox "����� ������ � ������ �����, � ������� ������ ������ ������" _
                    & vbCr & "���������� ������� �������� ����� � ��������� ���� � ��� ����� ������ ����� ������ �����" _
                    , , "������ �������� �����-��������"
                Exit Sub
            End If
        End If
        gNzak = Grid.TextMatrix(mousRow, orNomZak)
    
        If zakazNum = 0 Then Exit Sub
        FindFirm.Regim = "edit"
        FindFirm.cmSelect.Visible = True
        FindFirm.tb.Text = Grid.TextMatrix(mousRow, orFirma)
        FindFirm.Show vbModal
    Else
        Grid_DblClick
    End If
End If
End Sub

Private Sub Grid_LeaveCell()
If noClick Then Exit Sub
Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
End If
End Sub

Private Sub lbAnnul_DblClick()
Dim str As String, id As String

If noClick Then Exit Sub
' ����� ���-�� ������ "������" � "�����������"
str = Grid.TextMatrix(mousRow, mousCol) ' ������ ��������
If lbAnnul.Text = str Then GoTo EN1 '  �������� ��  ����������
If lbAnnul.Text = "�����������" Then
    do_Annul
ElseIf lbAnnul.Text = "������" Then
        If orderClose Then
            visits "+"    ' ��������� ��������� ������
            Grid.TextMatrix(mousRow, mousCol) = lbAnnul.Text
        End If
ElseIf lbAnnul.Text = "������" Then
    id = 0
#If Not COMTEC = 1 Then '---------------------------------------------------
    '"�����" --> "������" - ��� ���������, ���� ������ ����
    If str = "�����" And isNewEtap And Not predmetiIsClose Then GoTo BB
#End If '-------------------------------------------------------------------
    GoTo AA
ElseIf lbAnnul.Text = "�����" Then
    id = 4
AA: If MsgBox("����� ��������� ������� ����� ��������� ������ � ��������� " & _
    "��������� � ������ ��������. ���� �� ������� , ������� <��>, ����� ����������� " & _
    "����������� ��� ���� ������ �� ������������ ������ �������. ���� " & _
    "� ������ ���� �������� � �� ��� �����, �� ��������� �������� � ���������� ����� ����������!" _
    , vbDefaultButton2 Or vbYesNo, "��������!!") = vbNo Then GoTo EN1
BB: wrkDefault.BeginTrans
    str = manId(cbM.ListIndex)
    orderUpdate "##50", str, "Orders", "lastManagId"
    If orderUpdate("##50", id, "Orders", "StatusId") = 0 Then
        Grid.TextMatrix(mousRow, mousCol) = lbAnnul.Text
'        If lbAnnul.Text = "������" Then - !!! ���� ������ ���� ����� �� ����
'            orderUpdate "##329", 0, "Orders", "CehId" '����� ��� ������
'            Grid.TextMatrix(mousRow, orCeh) = "" ' ��� ��������� ������� ������
'        End If
        wrkDefault.CommitTrans
    Else
        wrkDefault.rollback
    End If
End If
EN1:
lbHide
End Sub

Private Sub lbAnnul_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbAnnul_DblClick
End Sub

Private Sub lbCeh_DblClick()
If noClick Then Exit Sub
If lbCeh.Visible = False Then Exit Sub

Grid.Text = lbCeh.Text
If orderUpdate("##21", lbCeh.ListIndex + 1, "Orders", "CehId") Then _
    Grid.Text = lbCeh.Text
lbHide
End Sub

Private Sub lbCeh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbCeh_DblClick
End Sub

Private Sub lbClose_DblClick()
Dim str As String

If noClick Then Exit Sub
If lbClose.Visible = False Then Exit Sub
' ����� ���-�� ������ "������"
If lbClose.Text = "�����������" Then
  do_Annul "no_visit"
End If
lbHide
    
End Sub

Private Sub lbClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbClose_DblClick

End Sub
 
Function do_Annul(Optional txt As String = "") As Boolean
Dim str As String
    do_Annul = False
    numDoc = gNzak
#If Not COMTEC = 1 Then '----------------------------------------------------
'    If beNaklads("noMsg") Then
'        MsgBox "� ����� ������ ���� ���������. ������� ������� ��.", , "������������� ����������!"
'        Exit Function
'    End If
    If havePredmetiNew Then
        MsgBox "� ����� ������ ���� ��������. ������� ������� ��.", , "������������� ����������!"
        Exit Function
    End If
#End If '--------------------------------------------------------------------
    do_Annul = True
    If txt = "no_Do" Then Exit Function
    
    wrkDefault.BeginTrans
    delZakazFromReplaceRS ' ���� �� ��� ����
    If txt = "" Then visits "-"    ' ��������� ��������� ������
    str = manId(cbM.ListIndex)
    orderUpdate "##369", str, "Orders", "lastManagId"
    If orderUpdate("##369", 7, "Orders", "StatusId") = 0 Then
        Grid.TextMatrix(mousRow, mousCol) = "�����������"
        wrkDefault.CommitTrans
    Else
        wrkDefault.rollback
    End If

End Function

Sub do_Del()
  If MsgBox("�� ������ <��> ��� ���������� �� ������ ����� ������������ " & _
  "������� �� ����!", vbDefaultButton2 Or vbYesNo, "������� ����� " & _
  gNzak & " ?") = vbYes Then
    wrkDefault.BeginTrans
    
    '������ ����-�� ��������� (��������)
    
#If Not COMTEC = 1 Then '------------------------------------------------
    sql = "DELETE From sDMCrez WHERE numDoc =" & gNzak & ";"
    myExecute "##305", sql, 0
#End If '------------------------------------------------------------------
'$comtec$ �������� ������������� ������ ���� ���������, ���� � ������� ������
'��� �������� �������� ���������� ������� ������ �� ������ �� ������� Orders
'    sql = "DELETE OrdersMO.*, Orders.*, OrdersInCeh.* " & _
    "FROM (Orders LEFT JOIN OrdersInCeh ON Orders.numOrder = OrdersInCeh.numOrder) " & _
    "LEFT JOIN OrdersMO ON Orders.numOrder = OrdersMO.numOrder " & _
    "WHERE Orders.numOrder = " & gNzak & ";"
'� ���� ���� ��������� ��������
    sql = "DELETE FROM Orders WHERE numOrder=" & gNzak & ";"
'    myBase.Execute sql
    If myExecute("##136", sql) = 0 Then
        delZakazFromGrid
        wrkDefault.CommitTrans
    Else
ERR1:   wrkDefault.rollback
    End If
  End If

End Sub

Private Sub lbDel_DblClick()
If noClick Then Exit Sub
If lbDel.Visible = False Then Exit Sub
If lbDel.Text = "�������" Then
  do_Del
End If
lbHide

End Sub

Sub delZakazFromGrid()
    zakazNum = zakazNum - 1 '
    If zakazNum = 0 Then
        clearGridRow Grid, mousRow
    Else
        Grid.removeItem mousRow
    End If
    Grid.col = orNomZak

End Sub

Private Sub lbDel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbDel_DblClick
End Sub

Private Sub lbM_DblClick()
Dim str As String, I As Integer

If noClick Then Exit Sub
If lbM.Visible = False Then Exit Sub
Grid.Text = lbM.Text
str = manId(lbM.ListIndex)
orderUpdate "##22", str, "Orders", "ManagId"
str = manId(cbM.ListIndex)
orderUpdate "##49", str, "Orders", "lastManagId"

lbHide
End Sub

Private Sub lbM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbM_DblClick
End Sub

Private Sub lbProblem_DblClick()
Dim str As String, I As Integer, DNM As String

If noClick Then Exit Sub
If lbProblem.Visible = False Then Exit Sub


Grid.Text = lbProblem.Text
str = lbProblem.ListIndex
If lbProblem.ListIndex > 5 Then str = lbProblem.ListIndex + 4
orderUpdate "##22", str, "Orders", "ProblemId"
str = manId(cbM.ListIndex)
orderUpdate "##49", str, "Orders", "lastManagId"

lbHide

DNM = Format(Now(), "dd.mm.yy hh:nn") & vbTab & Orders.cbM.Text & " " & gNzak ' ������ vbTab
On Error Resume Next ' � �����.��������� ���� �� Open logFile ���� Err: ���� ��� ������
Open logFile For Append As #2
Print #2, DNM & " ��������=" & lbProblem.Text & _
"   ���=" & Grid.TextMatrix(mousRow, orZakazano) & _
" �����=" & Grid.TextMatrix(mousRow, orZalog) & _
" ���=" & Grid.TextMatrix(mousRow, orNal) & _
" ���=" & Grid.TextMatrix(mousRow, orOplacheno) & _
" ���=" & Grid.TextMatrix(mousRow, orOtgrugeno)
Close #2
End Sub

Private Sub lbProblem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbProblem_DblClick
End Sub

Function orderClose() As Boolean
Dim sql2 As String, str As String, account_is_closed As Integer

orderClose = False

openOrdersRowToGrid "##56"
bilo = isConflict("toClose")
str = tqOrders!Type
tqOrders.Close
'lbHide 24 ���������� bilo 23.11.04

If str = "" Then
    MsgBox "����� ���������  ���������� ������� ��������� ������.", , "�������� ����������!"
    Exit Function
End If
    
If Not bilo Then
    If Grid.TextMatrix(mousRow, orProblem) = "" Then
        If Not predmetiIsClose Then ' ��� �������� ����� ��� ������� ��� ������
            MsgBox "� ����� ������ ���� ����������� ��������.", , "�������� ����������!"
            Exit Function
        End If
        sql = "select wf_order_closed_comtex (" & gNzak & ", '" & Grid.TextMatrix(mousRow, orServername) & "')"
        byErrSqlGetValues "##45.1", sql, account_is_closed
        If account_is_closed <> 1 Then
            MsgBox "������ ������� �����, �� ��� ���, ���� �� �� ������ � �����������.", , "�������� ����������!"
            Exit Function
        End If
        
        wrkDefault.BeginTrans   ' ������ ����������
        str = manId(cbM.ListIndex)
        orderUpdate "##45", 6, "Orders", "StatusId"
        orderUpdate "##48", str, "Orders", "lastManagId"
        delZakazFromCeh
        sql = "DELETE From sDMCrez WHERE (((numDoc)=" & gNzak & "));"
        myExecute "##326", sql, 0
        sql = "DELETE From xEtapByIzdelia WHERE (((numOrder)=" & gNzak & "));"
        myExecute "##327", sql, 0
        sql = "DELETE From xEtapByNomenk WHERE (((numOrder)=" & gNzak & "));"
        myExecute "##328", sql, 0
        
        wrkDefault.CommitTrans  ' ������������� ����������
        orderClose = True
    Else
        MsgBox "���������� ������� ����� ��������� � ���� ����������� " & _
        "��������", , "����� � ����������!"
    End If
    Exit Function
End If
  If Grid.TextMatrix(mousRow, mousCol) = "������" Then
    MsgBox "���������� ������� ����� ��������� �� ����� �������� � �������" _
    , , "����� � ����������!"
  Else
    MsgBox "���������� ������� ����� ��������� �� ����� ������������ (<Ctrl> " & _
       "+ <I> - ��� ���������) ��� ��������.", , "����� � ����������!"
  End If
End Function

Sub delZakazFromCeh()
        
        sql = "DELETE From OrdersInCeh WHERE (((" & _
              "OrdersInCeh.numOrder)=" & gNzak & "));"
        tmpStr = "DELETE From OrdersMO WHERE (((" & _
              "OrdersMO.numOrder)=" & gNzak & "));"
        On Error Resume Next '���� ���� ����� �� ������� �� ������ ��� ���
        myBase.Execute sql
        myBase.Execute tmpStr
        delZakazFromReplaceRS ' ���� �� ��� ����
        On Error GoTo 0
End Sub
'$odbc15$
Private Sub lbStat_DblClick()
Dim v As Variant

If noClick Then Exit Sub
        
If lbStat.Text = "������" Then
  If orderClose Then Grid.TextMatrix(mousRow, mousCol) = lbStat.Text
ElseIf lbStat.Text = "������" Then
#If Not COMTEC = 1 Then '----------------------------------------------------
    v = isNewEtap
    If IsNull(v) Then
        MsgBox "������ ��������� ������� ����� ����� � '������', ��������� " & _
        " �  ��� ��������� �� ��� ������ ���� ��������.", , "������������ ������!"
    ElseIf Not v Then
        MsgBox "��� �������� ������ ����� ���������� � ��������� ������ " & _
        "������ �������� � �������  '���-�� �� �������� �����'"
    ElseIf predmetiIsClose Then '
        MsgBox "� ����� ������ ��� �������� �������. �������� ������ ����� " & _
        "����������!", , "������������ ������!"
    Else
#Else
    If 1 = 1 Then
#End If '--------------------------------------------------------------------
        wrkDefault.BeginTrans
        delZakazFromCeh
        
'        Set tbOrders = myOpenRecordSet("##325", "select * from Orders", dbOpenForwardOnly)
'        If tbOrders Is Nothing Then GoTo ER1
'        tbOrders.index = "PrimaryKey"
'        tbOrders.Seek "=", gNzak
'        If tbOrders.NoMatch Then GoTo ER0
'        tbOrders.Edit
'        tbOrders!StatusId = 0
'        tbOrders!dateRS = Null
'        tbOrders!outDateTime = Null
'        tbOrders!workTime = Null
'        tbOrders!lastManagId = manId(cbM.ListIndex)
'        tbOrders.Update
'        tbOrders.Close
        
        sql = "UPDATE Orders SET StatusId = 0, DateRS = Null, " & _
        "outDateTime = Null, workTime = 0, lastManagId = '" & _
        manId(cbM.ListIndex) & "' WHERE (((numOrder)=" & gNzak & "));"
        If myExecute("##325", sql) <> 0 Then GoTo ER1
        
        wrkDefault.CommitTrans
        Grid.TextMatrix(mousRow, orStatus) = "������"
        Grid.TextMatrix(mousRow, orDataVid) = ""
        Grid.TextMatrix(mousRow, orVrVid) = ""
        Grid.TextMatrix(mousRow, orVrVip) = ""
        Grid.TextMatrix(mousRow, orDataRS) = ""
        Grid.TextMatrix(mousRow, orM) = ""
        Grid.TextMatrix(mousRow, orO) = ""
        Grid.TextMatrix(mousRow, orMOData) = ""
        Grid.TextMatrix(mousRow, orMOVrVid) = ""
        Grid.TextMatrix(mousRow, orOVrVip) = ""
        Grid.TextMatrix(mousRow, orLastMen) = ""
    End If
End If
lbHide
'Exit Sub
'ER0: tbOrders.Close
ER1:
' wrkDefault.Rollback:
lbHide
End Sub

Private Sub lbStat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbStat_DblClick
End Sub

Private Sub lbTema_DblClick()
Dim str As String, I As Integer, DNM As String

If noClick Then Exit Sub
If lbTema.Visible = False Then Exit Sub

Grid.Text = lbTema.Text
str = lbTema.ListIndex

orderUpdate "##73", str, "Orders", "TemaId"

lbHide


End Sub

Private Sub lbTema_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbTema_DblClick
End Sub

Private Sub lbType_DblClick()
Dim str As String, I As Integer

If noClick Then Exit Sub
If lbType.Visible = False Then Exit Sub
Grid.Text = lbType.Text
orderUpdate "##71", "'" & lbType.Text & "'", "Orders", "Type"
If Grid.Text <> "�" Then
    orderUpdate "##73", 0, "Orders", "TemaId"
    Grid.TextMatrix(mousRow, orTema) = ""
End If
lbHide

End Sub

Private Sub lbType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbType_DblClick
End Sub

Private Sub lbVenture_DblClick()
Dim str As Variant, I As Integer, newInv As String

If noClick Then Exit Sub
If lbVenture.Visible = False Then Exit Sub
I = orderUpdate("##72", lbVenture.ItemData(lbVenture.ListIndex), "Orders", "ventureId")
If I = 0 Then
    Grid.Text = lbVenture.Text
    If (lbVenture.ListIndex = 0) Then Grid.Text = ""
    newInv = getValueFromTable("Orders", "invoice", "numOrder = " & gNzak)
    Grid.TextMatrix(mousRow, orInvoice) = newInv
    str = getValueFromTable("GuideVenture", "sysname", "ventureId = " & lbVenture.ListIndex)
    If IsNull(str) Then str = ""
    Grid.TextMatrix(mousRow, orServername) = str
End If

lbHide

End Sub

Private Sub lbVenture_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbVenture_DblClick

End Sub

Private Sub mnAllOrders_Click()
Me.MousePointer = flexHourglass
Report.Regim = "allOrdersByFirmName"
Report.Show vbModal
Grid.SetFocus
Me.MousePointer = flexDefault

End Sub

Private Sub mnArhZone_Click()
loadArhinOrders
End Sub

Private Sub mnBillFirma_Click()
Dim ventureName As String

    ventureName = Grid.TextMatrix(mousRow, orVenture)
    If serverIsAccessible(ventureName) Then
        g_id_bill = mnBillFirma.Tag
        FirmComtex.Show vbModal
    Else
        MsgBox "������ " & ventureName & " �� �������� ", , "��������������"
    End If
    
End Sub

Private Sub mnComtexAdmin_Click()
cfg.Regim = "comtexAdmin"
cfg.setRegim
cfg.Show vbModal
End Sub

Private Sub mnExit_Click()
    exitAll
End Sub

Private Sub mnFirmFind_Click()
    If tbEnable.Visible Then
        FindFirm.cmAllOrders.Visible = True
        FindFirm.cmNoClose.Visible = True
        FindFirm.cmNoCloseFiltr.Visible = True
    End If
    FindFirm.Show vbModal

End Sub

Private Sub mnFirmsGuide_Click()
    If Grid.TextMatrix(mousRow, orVenture) <> "" Then
        ' ���������, ���� ����� ������ � ���� ������ � ������, �� �� ��������� ���� ������� ����
        If OrderIsMerged Then
            MsgBox "����� ������ � ������ �����, � ������� ������ ������ ������" _
                & vbCr & "���������� ������� �������� ����� � ��������� ���� � ��� ����� ������ ����� ������ �����" _
                , , "������ �������� �����-��������"
            Exit Sub
        End If
    End If
    Me.MousePointer = flexHourglass
    GuideFirms.Regim = "fromContext"
    
    GuideFirms.Show vbModal
    Me.MousePointer = flexDefault

End Sub

Private Sub mnGuideFirms_Click()
Me.MousePointer = flexHourglass
GuideFirms.Regim = "fromMenu"
GuideFirms.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub mnMeassure_Click()
cbM_LostFocus
End Sub

Private Sub mnMenu_Click()
cbM_LostFocus
End Sub

Private Sub mnNaklad_Click()
#If Not COMTEC = 1 Then '---------------------------------------------------

sDocs.Regim = ""
sDocs.Show
#End If '----------------------------------------------------------------
End Sub

Private Sub mnNoArhivFiltr_Click()
loadFirmOrders "noArhiv"
End Sub

Private Sub mnNoClose_Click()
Me.MousePointer = flexHourglass
Report.Regim = "OrdersByFirmName"
Report.Show vbModal
Grid.SetFocus
Me.MousePointer = flexDefault
End Sub

Private Sub mnNoCloseFiltr_Click()
loadFirmOrders ""
End Sub

Private Sub mnNomenk_Click()
#If Not COMTEC = 1 Then '----------------------------------------------------
    sProducts.Regim = "ostat"
    sProducts.Show vbModal
#Else
    MsgBox "����� ���������� ������ ����� � ������������ ��������� " & _
    "��������� �������� �� ������ �����������.", , "" '$comtec$
#End If '--------------------------------------------------------------------
End Sub

Private Sub mnPathSet_Click()
cfg.loadFileConfiguration ' ��������� ���������� �� ������ ������
cfg.Regim = "pathSet"
cfg.setRegim
cfg.Show vbModal
webSvodkaPath = SvodkaPath          '$$2
webLoginsPath = loginsPath          '

End Sub

Private Sub mnProduct_Click()
#If Not COMTEC = 1 Then '----------------------------------------------------
    sProducts.Regim = "ostatP"
    sProducts.Show vbModal
#Else
    MsgBox "����� ���������� ������ ����� � ������������ ��������� " & _
    "��������� �������� �� ������������, �������� � �������.", , "" '$comtec$
#End If '------------------------------------------------------------------
End Sub

Private Sub mnQuickBill_Click(Index As Integer)
    If Index = 0 Then Exit Sub
    FirmComtex.makeBillChoice mnQuickBill(Index).Tag, Grid.TextMatrix(mousRow, orServername)
End Sub

Private Sub mnReports_Click()
Reports.Show vbModal
End Sub

Private Sub mnServic_Click()
cbM_LostFocus
End Sub

Private Sub mnSetkaC_Click()
    startParams (2)
    Zakaz.Show vbModal
End Sub

Private Sub mnSetkaS_Click()
    startParams (3)
    Zakaz.Show vbModal
End Sub

Private Sub mnSetkaY_Click()
    startParams (1)
    Zakaz.Show vbModal
End Sub

Private Sub mnSklad_Click()
cbM_LostFocus
End Sub

Private Sub tbEnable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If LCase(tbEnable.Text) <> "arh" And LCase(tbEnable.Text) <> "���" Then
        tbEnable.Text = ""
        tbEnable.Visible = False
    End If
    Grid.SetFocus
ElseIf KeyCode = vbKeyDelete Then
    minut = 0
    Timer1_Timer
End If
End Sub

Private Sub tbEnable_LostFocus()
If LCase(tbEnable.Text) = "arh" Or LCase(tbEnable.Text) = "���" Then ' �� ��� � onKeyDown
    laClos.Visible = True
    cbClose.Visible = True
    mnAllOrders.Visible = True
    mnSep2.Visible = True
    mnNoCloseFiltr.Visible = True
    mnNoClose.Visible = True
    laConflict.Visible = True
    chConflict.Visible = True
    cmToWeb.Visible = True
    mnReports.Visible = True
    If dostup = "a" Then
    Else
        minut = 5
        Timer1.Interval = 60000 ' 1 ������
        Timer1.Enabled = True
    End If
Else
    tbEnable.Visible = False
End If
Grid.SetFocus
End Sub

Private Sub tbEndDate_Change()
cmRefr.Caption = "���������"

End Sub

Function DateFromMobileVrVid(col As Integer) As String
Dim n As Integer

If checkNumeric(tbMobile.Text, 9, 21) Then
    n = tbMobile.Text
    DateFromMobileVrVid = Grid.TextMatrix(mousRow, col)
    If DateFromMobileVrVid = "" Then
        MsgBox "����� ����� ����� ������ ����� ����, ��� ����� ��������� ����!", , ""
        lbHide
        Exit Function
    End If
    DateFromMobileVrVid = "'" & Format(DateFromMobileVrVid & " " & _
                          n & ":00:00", "yyyy-mm-dd hh:nn:ss") & "'"
    Grid.TextMatrix(mousRow, mousCol) = n
Else
    tbMobile.SelStart = 0
    tbMobile.SelLength = Len(tbMobile.Text)
    DateFromMobileVrVid = ""
End If

End Function

Function isFloatFromMobile(field As String) As Boolean

        If checkNumeric(tbMobile.Text, 0) Then
            orderUpdate "##23", tbMobile.Text, "Orders", field
            Grid.TextMatrix(mousRow, mousCol) = tbMobile.Text
            isFloatFromMobile = True
        Else
            tbMobile.SelStart = 0
            tbMobile.SelLength = Len(tbMobile.Text)
            isFloatFromMobile = False
        End If
End Function

Private Sub tbInform_GotFocus()
    tbInform.SelStart = Len(tbInform.Text)

End Sub

Private Sub tbInform_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    gNzak = Grid.TextMatrix(mousRow, orNomZak)
    tbMobile.Text = tbInform.Text
    tbMobile_KeyDown vbKeyReturn, 0
ElseIf KeyCode = vbKeyEscape Then
    Grid.SetFocus
End If
End Sub

Private Sub tbMobile_Change()
tbInform.Text = tbMobile.Text
End Sub

Private Sub tbMobile_DblClick()
lbHide
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, DNM As String, s As Single
Dim id_jscet_split As Integer
Dim id_jscet_merge As Integer
Dim mFault As String
Dim bFault As Boolean
Dim p_newInvoice As String, p_Invoice As String
Dim next_nu As String

If KeyCode = vbKeyReturn Then
DNM = Format(Now(), "dd.mm.yy hh:nn") & vbTab & cbM.Text & " " & gNzak ' ������ vbTab
   
    If mousCol = orVrVip Then
        If Not isFloatFromMobile("workTime") Then Exit Sub
'        orderUpdate "##24", tbMobile.Text, "Orders", "workTime"
    ElseIf mousCol = orVrVid Then
        str = DateFromMobileVrVid(orDataVid)
        If str = "" Then Exit Sub
        orderUpdate "##24", str, "Orders", "outDateTime"
    ElseIf mousCol = orMOVrVid Then
        str = DateFromMobileVrVid(orMOData)
        If str = "" Then Exit Sub
        orderUpdate "##25", str, "OrdersMO", "DateTimeMO"
    ElseIf mousCol = orLogo Then
        orderUpdate "##26", "'" & tbMobile.Text & "'", "Orders", "Logo"
        Grid.TextMatrix(mousRow, mousCol) = tbMobile.Text
        On Error Resume Next ' � �����.��������� ���� �� Open logFile ���� Err: ���� ��� ������
        Open logFile For Append As #2
        Print #2, DNM & " ����=" & tbMobile.Text
        Close #2
    ElseIf mousCol = orIzdelia Then
        orderUpdate "##27", "'" & tbMobile.Text & "'", "Orders", "Product"
        Grid.TextMatrix(mousRow, mousCol) = tbMobile.Text
        On Error Resume Next ' � �����.��������� ���� �� Open logFile ���� Err: ���� ��� ������
        Open logFile For Append As #2
        Print #2, DNM & " �������=" & tbMobile.Text
        Close #2
    ElseIf mousCol = orZakazano Then
        If Not isFloatFromMobile("ordered") Then Exit Sub
    ElseIf mousCol = orOplacheno Then
        If Not isFloatFromMobile("paid") Then Exit Sub
    ElseIf mousCol = orZalog Then
        If Not isFloatFromMobile("zalog") Then Exit Sub
    ElseIf mousCol = orNal Then
        If Not isFloatFromMobile("nal") Then Exit Sub
    ElseIf mousCol = orOtgrugeno Then
        If Not isFloatFromMobile("shipped") Then Exit Sub
        s = Round(tbMobile.Text, 2)
        If s = 0 Then
#If Not COMTEC = 1 Then '---------------------------------------------------
            orderUpdate "##78", "Null", "Orders", "shipped"
            Grid.TextMatrix(mousRow, orOtgrugeno) = ""
#End If '-------------------------------------------------------------------
        ElseIf flDelRowInMobile Then
            flDelRowInMobile = False
            delZakazFromGrid
        End If
    ElseIf mousCol = orInvoice Then
        If Grid.TextMatrix(mousRow, orVenture) <> "" Then
            sql = "select nextnu_remote( '" & Grid.TextMatrix(mousRow, orServername) & "', 'jscet')"
            byErrSqlGetValues "##78.1", sql, next_nu
            If tbMobile.Text <> next_nu Then
                If vbYes <> MsgBox("��������� ����� �� ����������� ������ ���� ����� " _
                    & next_nu & ". ������� ��, ���� �� ������������� ������ �������� ����� ������ �� " _
                    & tbMobile.Text, vbYesNo, "��������") _
                Then
                    GoTo AA
                End If
            End If
        End If
        
        If InStr(tbMobile.Text, "����") > 0 Or tbMobile.Text = "0" Then
            str = Grid.TextMatrix(mousRow, orOtgrugeno)
            If IsNumeric(str) And dostup <> "a" Then
              If Grid.TextMatrix(mousRow, orCeh) = "" Or _
              Grid.TextMatrix(mousRow, orStatus) = "�����" Then ' � 2� ������
                delZakazFromGrid
              Else
                MsgBox "��� ��������� �� ����������� ������� ����� ������ ����� ����", , "������"
                GoTo AA
              End If
            Else '���� � "��������� ������ ���"
                Grid.TextMatrix(mousRow, mousCol) = "���� ?"
            End If
            orderUpdate "##77", "'" & "���� ?" & "'", "Orders", "Invoice"
        Else
            If Grid.TextMatrix(mousRow, orVenture) <> "" Then
        
                id_jscet_split = checkInvoiceSplit(gNzak, tbMobile.Text)
                id_jscet_merge = checkInvoiceMerge(gNzak, tbMobile.Text)
                
                p_newInvoice = tbMobile.Text
                p_Invoice = Grid.TextMatrix(mousRow, orInvoice)
                mFault = ""
                bFault = False
                
                If id_jscet_merge < 0 Then
                    mFault = "����� " & gNzak & " �� ��� ����������� � ����� " & p_newInvoice
                ElseIf id_jscet_split > 0 And id_jscet_merge > 0 Then
                    bFault = tryInvoiceMove(gNzak, p_Invoice, id_jscet_merge, p_newInvoice)
                    mFault = mFault = "����� " & gNzak & " �� ��� ��������� �� ����� " & gNzak & " � ���� " & p_newInvoice
                ElseIf id_jscet_split > 0 Then
                    bFault = tryInvoiceSplit(gNzak, p_Invoice)
                    mFault = "����� " & gNzak & " �� ��� ������� � ��������� ����"
                ElseIf id_jscet_merge > 0 Then
                    bFault = tryInvoiceMerge(gNzak, id_jscet_merge, p_newInvoice)
                    mFault = "����� " & gNzak & " �� ��� ����������� � ����� " & p_newInvoice
                End If
                
                If Not bFault And mFault <> "" Then
                    MsgBox "��������� ������" & vbCr & mFault, , "�������� ��������������"
                    Exit Sub
                End If
            End If
            
            If Not isFloatFromMobile("Invoice") Then Exit Sub
        End If
    End If
    str = manId(cbM.ListIndex)
    orderUpdate "##48", str, "Orders", "lastManagId"
    
    GoTo AA
ElseIf KeyCode = vbKeyEscape Then
AA:
lbHide
End If

End Sub

Private Sub tbStartDate_Change()
cmRefr.Caption = "���������"
End Sub

Private Sub tbStartDate_GotFocus()
oldValue = tbStartDate.Text
End Sub

Private Sub tbStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tbStartDate_LostFocus
End If
End Sub

Private Sub tbStartDate_LostFocus()
isDateTbox tbStartDate
End Sub

Sub LoadBase(Optional reg As String = "")
Dim numZak As Long, I As Integer

laInform.Caption = ""
Grid.Visible = False
clearGrid Grid

#If Not COMTEC = 1 Then '----------------------------------------------------
   getNakladnieList
#End If
zakazNum = 0
'LoadOrders********************************************************
sql = rowFromOrdersSQL & getSqlWhere & " ORDER BY Orders.inDate" & ";"
'MsgBox getSqlWhere
'Debug.Print "sql4 = " & sql
Set tqOrders = myOpenRecordSet("##08", sql, dbOpenForwardOnly)
If tqOrders Is Nothing Then myBase.Close: End
If Not tqOrders.BOF Then
While Not tqOrders.EOF
 
 numZak = tqOrders!numOrder
  
 If chConflict.value = 1 Then If Not isConflict() Then GoTo NXT
 
' On Error GoTo ERR1
 If zakazNum > 0 Then Grid.AddItem ""
 
 zakazNum = zakazNum + 1
 
 Grid.TextMatrix(zakazNum, orNomZak) = numZak
 noClick = True
    If Not IsNull(tqOrders!id_bill) Then '�������
         Grid.col = orFirma
         Grid.row = zakazNum
         Grid.CellForeColor = vbRed
    End If
 If tqOrders!StatusId < 6 Then '***************
#If Not COMTEC = 1 Then '----------------------------------------------------
   For I = 1 To UBound(tmpL)
     If tmpL(I) = numZak Then
        Grid.col = orNomZak
        Grid.row = zakazNum
        Grid.CellForeColor = 200
        Exit For
     ElseIf tmpL(I) = -numZak Then '��� ��������� �������
        Grid.col = orNomZak
        Grid.row = zakazNum
        Grid.CellForeColor = vbBlue
        Exit For
     End If
   Next I
#End If '--------------------------------------------------------------
   If tqOrders!urgent = "y" Then '�������
        Grid.col = orCeh
        Grid.row = zakazNum
        Grid.CellForeColor = 200
   End If
#If Not COMTEC = 1 Then '----------------------------------------------------
 ElseIf tqOrders!StatusId = 6 Then
    sql = "SELECT xPredmetyByIzdelia.numOrder from xPredmetyByIzdelia " & _
    "Where (((xPredmetyByIzdelia.numOrder) = " & numZak & ")) " & _
    "UNION SELECT xPredmetyByNomenk.numOrder from xPredmetyByNomenk " & _
    "WHERE (((xPredmetyByNomenk.numOrder)=" & numZak & "));"
    numZak = 0
    byErrSqlGetValues "W##360", sql, numZak
    If numZak > 0 Then
        Grid.col = orNomZak
        Grid.row = zakazNum
        Grid.CellForeColor = &H8800& ' �.���.
    End If
#End If '--------------------------------------------------------------
 End If '*************************************
 noClick = False
 
 copyRowToGrid zakazNum

NXT:
 tqOrders.MoveNext
Wend

End If 'Not tqOrders.BOF
loadBaseTimestamp = Now()
NXT2:
tqOrders.Close '*********************************************

laInform.Caption = " ���-�� ���.: " & zakazNum
rowViem zakazNum, Grid
Grid.Visible = True
If zakazNum > 0 Then
    Grid.col = 1
    Grid.row = zakazNum
    
    On Error Resume Next
    Grid_EnterCell
    Grid.SetFocus
End If
Exit Sub

ERR1:
If Err = 7 Then
    MsgBox "��������� ��������� ������ ��� ��������� ���� �������. " & _
    "���������� ������� ������ ���������!", , "������ 351"
Else
    MsgBox Error, , "������ 351-" & Err & ":  " '##351
End If
On Error Resume Next
GoTo NXT2
End Sub

Function getSqlWhere() As String
Dim I As Integer

getSqlWhere = ""
For I = 0 To orColNumber
  If orSqlWhere(I) <> "" Then
    If getSqlWhere = "" Then
        getSqlWhere = "(" & orSqlWhere(I) & ")"
    Else
        getSqlWhere = getSqlWhere & " AND " & "(" & orSqlWhere(I) & ")"
    End If
'    MsgBox "orSqlWhere=" & orSqlWhere(I) & "  getSqlWhere=" & getSqlWhere
  End If
Next I
If getSqlWhere <> "" Then getSqlWhere = " WHERE (" & getSqlWhere & ")"
'MsgBox "Where = " & getSqlWhere
    
End Function

Function strWhereByValCol(value As String, col As Integer, Optional _
operator As String = "=") As String
Dim str As String, typ As String, oper As String

oper = " " & operator & " "
strWhereByValCol = ""
str = orSqlFields(col)
If str = "" Then
    MsgBox "�� ����� ���� ������ �� ������������"
    Exit Function
End If
typ = left$(str, 1)
str = Mid$(str, 2)
If typ = "d" Then
    If value = "" Then
        value = " Is Null"
    Else
        If operator = "=" Then
            value = left$(value, 6) & "20" & Mid$(value, 7, 2) '��� ����� ���� � Win98 ���������� "����" - ������ ����
            value = " Like '" & value & "%'"
        ElseIf operator = "<" Then
            value = " <= '" & Format(value, "yyyy-mm-dd") & " 11:59:59 PM'"
        Else
            value = " >= '" & Format(value, "yyyy-mm-dd") & "'"
        End If
    End If
ElseIf typ = "s" Then
    value = " = '" & value & "'"
Else
    If value = "" Then
        value = " Is Null"
    Else
        value = oper & value
    End If
End If
strWhereByValCol = "(" & str & ")" & value

End Function


Sub loadArhinOrders()
Dim I As Integer

For I = 1 To orColNumber
    orSqlWhere(I) = ""
Next I

orSqlWhere(orInvoice) = "(Orders.Invoice) Like '����%'"
orSqlWhere(orStatus) = "(GuideStatus.Status) <> '������'"
orSqlWhere(orOtgrugeno) = "Not(Orders.shipped) Is Null"
Orders.MousePointer = flexHourglass
Orders.LoadBase
Orders.MousePointer = flexDefault
Orders.laFiltr.Visible = True
Orders.begFiltrDisable

End Sub

Sub loadFirmOrders(stat As String, Optional ordNom As String = "")
Dim I As Integer

For I = 1 To orColNumber
    orSqlWhere(I) = ""
Next I
If stat = "noArhiv" Then
    stat = ""
    orSqlWhere(orInvoice) = "isNumeric(Orders.Invoice) =1 OR " & _
    "(Orders.Invoice) Is Null OR (Orders.shipped) Is Null"
End If
If stat <> "all" And stat <> "" Then
    orSqlWhere(orFirma) = "(GuideFirms.Name) = '" & stat & "'"
Else
    orSqlWhere(orFirma) = "(GuideFirms.Name) = '" & Grid.Text & "'"
End If
If stat <> "all" Then _
    orSqlWhere(orStatus) = "(GuideStatus.Status) <> '������'"

MousePointer = flexHourglass
LoadBase
If ordNom <> "" Then findValInCol Grid, ordNom, orNomZak
MousePointer = flexDefault
laFiltr.Visible = True
begFiltrDisable

End Sub

Sub copyRowToGrid(row As Long)

 If Not IsNull(tqOrders!Invoice) Then _
     Grid.TextMatrix(row, orInvoice) = tqOrders!Invoice
 Grid.TextMatrix(row, orCeh) = tqOrders!Ceh
 Grid.TextMatrix(row, orMen) = tqOrders!Manag
 Grid.TextMatrix(row, orFirma) = tqOrders!name
 LoadDate Grid, row, orData, tqOrders!inDate, "dd.mm.yy"
 
 StatParamsLoad row
 
 Grid.TextMatrix(row, orLogo) = tqOrders!Logo
 Grid.TextMatrix(row, orIzdelia) = tqOrders!Product
 If Not IsNull(tqOrders!Type) Then _
    Grid.TextMatrix(row, orType) = tqOrders!Type
 If Not IsNull(tqOrders!temaId) Then _
    Grid.TextMatrix(row, orTema) = lbTema.List(tqOrders!temaId)
 LoadNumeric Grid, row, orZakazano, tqOrders!ordered
 LoadNumeric Grid, row, orOplacheno, tqOrders!paid
 LoadNumeric Grid, row, orZalog, tqOrders!zalog
 LoadNumeric Grid, row, orNal, tqOrders!nal
 LoadNumeric Grid, row, orOtgrugeno, tqOrders!shipped
 If Not IsNull(tqOrders!lastManag) Then _
    Grid.TextMatrix(row, orLastMen) = tqOrders!lastManag
 If Not IsNull(tqOrders!Venture) Then _
    Grid.TextMatrix(row, orVenture) = tqOrders!Venture
 If Not IsNull(tqOrders!LastModified) Then
    Grid.TextMatrix(row, orlastModified) = CStr(tqOrders!LastModified)
 End If
 If Not IsNull(tqOrders!id_bill) Then
    Grid.TextMatrix(row, orBillId) = CStr(tqOrders!id_bill)
 End If
 If Not IsNull(tqOrders!id_voc_names) Then
    Grid.TextMatrix(row, orVocnameId) = CStr(tqOrders!id_voc_names)
 End If
 If Not IsNull(tqOrders!serverName) Then
    Grid.TextMatrix(row, orServername) = CStr(tqOrders!serverName)
 End If
End Sub

Sub addToCbStatus(id, Optional begin As String = "")

Static I As Integer
If begin <> "" Then I = 0
If id > lenStatus Then
    MsgBox "Err � Orders\addToCbStatus"
End If

Zakaz.cbStatus.AddItem status(id)
statId(I) = id
I = I + 1

End Sub
    
Function cbMOsetByText(cb As ComboBox, stat As Variant) As Boolean
    cbMOsetByText = False
Dim I As Integer, txt As String
    txt = ""
    If Not IsNull(stat) Then txt = CStr(stat)
    If txt = "�����" Then
        If cb.List(3) <> "�����" Then cb.AddItem "�����", 3
        If cb.List(4) <> "���������" Then cb.AddItem "���������", 4
        cb.ListIndex = 3
        cbMOsetByText = True
    ElseIf txt = "���������" Then
        If cb.List(3) = "�����" Then
            I = 4
        Else
            I = 3
        End If
        If cb.List(I) <> "���������" Then cb.AddItem "���������", I
        cb.ListIndex = I
    ElseIf txt = "� ������" Then
        cb.ListIndex = 2
        cbMOsetByText = True
    ElseIf txt = "�����" Or txt = "�������" Then
        cb.ListIndex = 1
    Else
        cb.ListIndex = 0
    End If

End Function
'$odbc08!$
Function startParams(Optional idCeh As Integer = 0) As Boolean
Dim I As Integer, str As String, j As Integer ', sumSroch As Single
Dim item As ListItem, id As Integer, v As Variant, s As Single

startParams = False


'Set tbOrders = myOpenRecordSet("##28", "Orders", dbOpenTable)
'If tbOrders Is Nothing Then Exit Function
maxDay = 0

If idCeh > 0 Then ' ����� � ������ ����� �������
    Zakaz.cmAdd.Visible = False
    Zakaz.cmRepit.Visible = False
    cehId = idCeh
    gNzak = ""
    id = 0
    Zakaz.urgent = ""
Else
    sql = "SELECT urgent from OrdersInCeh WHERE (((numOrder)=" & gNzak & "));"
    byErrSqlGetValues "W##381", sql, str
    Zakaz.urgent = str
    
    Zakaz.laNomZak.Caption = gNzak
    Zakaz.cmAdd.Visible = True
    Zakaz.cmRepit.Visible = True
    
'    If Not findZakazInTable_("Orders") Then Exit Function '$#$
    sql = "SELECT StatusId, outDateTime from Orders " & _
    "WHERE (((Orders.numOrder)=" & gNzak & "));"
    Set tbOrders = myOpenRecordSet("##402", sql, dbOpenForwardOnly)
    If tbOrders.BOF Then Exit Function
    
    If IsDate(tbOrders!outDateTime) Then
        I = DateDiff("d", curDate, tbOrders!outDateTime) + 1
        addDays I '��������� ���, �.�. ���� ��� ���.������ ����� ���������
                  '������ ��� ���� ������, ���� ��� stDay � rMaxDay
    End If
    id = tbOrders!StatusId
    tbOrders.Close
End If
    
    zagruzFromCeh gNzak '              1| � delta(), Ostatki() !!! ����� ��������
    getResurs
    Zakaz.lvAddDays  ' ��������� ������� � ����
    For I = 1 To maxDay
        Zakaz.lv.ListItems("k" & I).SubItems(zkPrinato) = Round(getNevip(I), 1)
        Zakaz.lv.ListItems("k" & I).SubItems(zkResurs) = Round(nomRes(I) * kpd * Nstan, 1)
    Next I
    Zakaz.lv.ListItems("k1").SubItems(zkResurs) = Round(nr * Nstan * kpd, 1)

If id = 0 Or id = 7 Then '������ ��� �������
    neVipolnen = 0
    neVipolnen_O = 0
    If idCeh > 0 Then
        Zakaz.Caption = "����� ������� " & Ceh(cehId)
    ElseIf id = 0 Then
        Zakaz.Caption = "����������� ������ � ��� " & Ceh(cehId)
    End If
    
    Zakaz.tbWorkTime = ""
    Zakaz.tbReadyDate = ""
Else
    Zakaz.Caption = "�������������� ������"
    Zakaz.tbDateRS = Grid.TextMatrix(mousRow, orDataRS)
    Zakaz.tbReadyDate = Grid.TextMatrix(mousRow, orDataVid)
          
    Zakaz.tbWorkTime = neVipolnen
    
    v = getTableField("OrdersMO", "StatM")
    If cbMOsetByText(Zakaz.cbM, v) Then
        Zakaz.tbDateMO = Grid.TextMatrix(mousRow, orMOData)
    End If
    v = getTableField("OrdersMO", "StatO")
    If cbMOsetByText(Zakaz.cbO, v) Then
        Zakaz.tbDateMO = Grid.TextMatrix(mousRow, orMOData)
        If Zakaz.cbO.Text = "�����" Then
            Zakaz.tbVrVipO = Grid.TextMatrix(mousRow, orOVrVip)
            Zakaz.tbVrVipO.Enabled = False
            Zakaz.tbDateMO.Enabled = False
        Else 'AS nevipO
            sql = "SELECT workTimeMO  FROM OrdersMO " & _
            "WHERE (((numOrder)=" & gNzak & "));"
'MsgBox sql
            byErrSqlGetValues "##384", sql, s '$odbc18!$
            neVipolnen_O = Round(s, 2)
        
            Zakaz.tbVrVipO = neVipolnen_O
        End If
    End If
End If
'tbOrders.Close
I = getNextDay(1)
v = Zakaz.lv.ListItems("k1").SubItems(zkMost)
If Not IsNumeric(v) Then v = 0
Zakaz.laZapas.Caption = Round(nomRes(I) * kpd * Nstan + v, 1)

Zakaz.cmZapros.Enabled = False

'��������� ���� �� ���� ������
For I = 1 To maxDay
    delta(I) = 0
Next I
str = "DateDiff(day,'" & Format(curDate, "yyyy-mm-dd") & "',outDateTime)"
sql = "SELECT " & str & " AS day, Orders.FirmId From Orders " & _
"Where (((Orders.cehId) = " & cehId & ") And ((Orders.StatusId) < 4)) " & _
"GROUP BY " & str & ", Orders.FirmId  HAVING (((" & str & ")>=0));"
'MsgBox str & Chr(13) & Chr(13) & sql
'Debug.Print sql

Set tbOrders = myOpenRecordSet("##76", sql, dbOpenForwardOnly)
If Not tbOrders Is Nothing Then
 If Not tbOrders.BOF Then
 While Not tbOrders.EOF
    I = tbOrders!day + 1
    delta(I) = delta(I) + 1
    tbOrders.MoveNext
 Wend
 End If
 tbOrders.Close
End If
For I = 1 To maxDay
    Zakaz.lv.ListItems("k" & I).SubItems(zkFirmKolvo) = Round(delta(I), 1)
Next I

Zakaz.cbStatus.Clear
addToCbStatus 7, "b" '"�������."
If id = 5 Then
    addToCbStatus 5   '"�������"
ElseIf id = 8 Then
    id = 1
    addToCbStatus 1 '"� ������"
Else
    addToCbStatus 0 '"������"  '�� ��������� � �.�. ���
    addToCbStatus 1 '"� ������"
    addToCbStatus 2 '"������"  '������-� � ������� ��������
    addToCbStatus 3 '"��������."
End If

For I = 0 To Zakaz.cbStatus.ListCount
    If statId(I) = id Then
        Zakaz.cbStatus.ListIndex = I
        GoTo NN
    End If
Next I
MsgBox "Err in Orders\startParams"
NN:

Zakaz.lv.ListItems("k" & stDay).ForeColor = &HBB00&
Zakaz.lv.ListItems("k" & stDay).Bold = True

If idCeh = 0 Then
    Zakaz.Regim = ""
    Zakaz.newZagruz
Else
    Zakaz.Regim = "setka"
    Zakaz.newZagruz "setka" '������ ������ ���� ���
End If

startParams = True
End Function

Private Sub Timer1_Timer()
minut = minut - 1
If minut <= 0 Then
    cbClose.value = 0
    chConflict.value = 0
    
    Timer1.Enabled = False
    tbEnable.Visible = False
    laClos.Visible = False
    cbClose.Visible = False
    mnAllOrders.Visible = False
    mnSep2.Visible = False
    mnNoCloseFiltr.Visible = False
    mnNoClose.Visible = False
    mnReports.Visible = False
    laConflict.Visible = False
    chConflict.Visible = False
    cmToWeb.Visible = False
    mnQuickBill(0).Visible = False
    mnBillFirma.Visible = False
End If
End Sub

#If COMTEC = 1 Then  '----------------------------------------------------

Private Sub mnToExcel_Click()
MsgBox "����� �������������� ���������� ���� ������� ������ ���� ����������� " & _
"�� ���� ������. ��. ����������� $comtec$ � ���� ���������.", , ""
Exit Sub

ostatToWeb "toExcel"
End Sub


'��� ��������� ��������� ����� Products � Nomenks, ����������� ��� ���������
'�������� �� ������ �� ����� �����.
'���������� ������� ������� � ��� ���������� �� �������� sGuideProducts
'(���������� ����������) � sProducts (�������� �������� ������������).
'�� � �� ����  prName("���" ������� �� ����� ������ ����������� �������
'������� � ���������� stime),web("web") � nomNom("�����" ������������ �
'������ ������) ���� �������� �� ����������� �� ���� Comtec. $comtec$
Private Sub mnWebs_Click()
Dim str As String, ch As String, slen As Integer, oper As String, I As Integer
Dim tmpFile As String

MsgBox "����� �������������� ���������� ���� ������� ������ ���� ����������� " & _
"�� ���� ������. ��. ����������� $comtec$ � ���� ���������.", , ""
Exit Sub




If MsgBox("�� ������ '��' ����� ������������ ����� ��� WEB: ���� �����c��� " & _
"�������� � ���� ������������ ������� �������." _
, vbDefaultButton2 Or vbYesNo, "����������� ������") = vbNo Then Exit Sub

Me.MousePointer = flexHourglass

sql = "SELECT sProducts.nomNom, sGuideProducts.prName " & _
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
    While Not tbProduct.EOF
      Print #1, tbProduct!prName & vbTab & tbProduct!nomNom
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

ostatToWeb '������ � �����
    
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

Function setVertBorders(lineWeight As Long) As Integer
On Error GoTo ERR1

objExel.ActiveSheet.Cells(exRow, 1).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 2).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 3).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 4).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 5).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 6).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 7).Borders(xlEdgeRight).Weight = lineWeight
objExel.ActiveSheet.Cells(exRow, 8).Borders(xlEdgeRight).Weight = xlMedium
Exit Function

ERR1:
setVertBorders = Err

End Function

Function nomencDostupOstatki(nomer) As Single
'$comtec$ ����� ���� ��������� ���������(�.�. � ������ ������������������)
'         ������� ��� ������������ � ���������� ������� � nomer

End Function

'$odbc15$
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
    ReDim Preserve NN(I): NN(I) = Format(tbNomenk!klassId, "0000")
    findId = tbNomenk!klassId

AA: 'tbGuide.Seek "=", findId
'    If tbGuide.NoMatch Then msgOfEnd ("##409")
    sql = "SELECT klassName, parentKlassId from sGuideKlass " & _
    "WHERE (((klassId)=" & findId & "));"
    If Not byErrSqlGetValues("##417", sql, str, findId) Then tbNomenk.Close: Exit Sub
            
'    NN(I) = tbGuide!klassName & " / " & NN(I) ' � ����� ��������� Id
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
    exRow = 3
    objExel.ActiveSheet.Cells(exRow, 1).value = "���"
    objExel.ActiveSheet.Cells(exRow, 2).value = "��������"
    objExel.ActiveSheet.Cells(exRow, 3).value = "������"
    objExel.ActiveSheet.Cells(exRow, 4).value = "��.���������"
    objExel.ActiveSheet.Cells(exRow, 5).value = "���-��"
    objExel.ActiveSheet.Cells(exRow, 6).value = "����"
    objExel.ActiveSheet.Columns(1).ColumnWidth = 12.57
    objExel.ActiveSheet.Columns(2).ColumnWidth = 39.71
    objExel.ActiveSheet.Columns(3).ColumnWidth = 10
    objExel.ActiveSheet.Columns(4).ColumnWidth = 6.2
    objExel.ActiveSheet.Columns(5).ColumnWidth = 6.2
    objExel.ActiveSheet.Columns(6).ColumnWidth = 6
    objExel.ActiveSheet.Columns(6).HorizontalAlignment = xlHAlignRight
    cErr = setVertBorders(xlMedium)
'xlColumnDataType
    If cErr <> 0 Then GoTo ERR2
'xlDiagonalDown, xlDiagonalUp, xlEdgeBottom, xlEdgeLeft, xlEdgeRight
'xlEdgeTop, xlInsideHorizontal, or xlInsideVertical.
    With objExel.ActiveSheet.Range("A" & exRow & ":F" & exRow)
        .Borders(xlEdgeBottom).Weight = xlMedium ' xlThin
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    exRow = 4
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
  sql = "SELECT nomNom, nomName, ed_Izmer2, CENA_W, perList, " & _
  "cod, Size From sGuideNomenk " & _
  "Where (((web) = 'web' ) AND ((klassId)=" & findId & ")) ORDER BY nomNom;"

  Set tbNomenk = myOpenRecordSet("##331", sql, dbOpenDynaset)
  If tbNomenk Is Nothing Then GoTo EN1
  If Not tbNomenk.BOF Then
      bilo = False
      While Not tbNomenk.EOF
        
        tmpSng = nomencDostupOstatki(tbNomenk!nomNom)
        
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
                cErr = setVertBorders(xlThin)
                If cErr <> 0 Then GoTo ERR2
                exRow = exRow + 1
            End If
        End If
'---------------------------------------------------------------------------
'����� �������� ��������� �� ������ ������������ ������
        str = tbNomenk!ed_Izmer2
'        If str = "����" Or str = "�����" Then
            tmpSng = tmpSng / tbNomenk!perList
'        End If
        tmpSng = Round(tmpSng - 0.4999, 0)
        Dim cena2w As String
        cena2w = Chr(160) & Format(tbNomenk!CENA_W, "0.00") ' ������� ��� �����, �.�. "3.00" ��� ����������� "3"
        If toExel = "" Then
            Print #1, tbNomenk!nomNom & vbTab & tbNomenk!nomName & vbTab & _
            str & vbTab & Round(tmpSng, 2) & vbTab & cena2w & _
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
            objExel.ActiveSheet.Cells(exRow, 6).value = cena2w 'Round(tbNomenk!CENA_W, 2)
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
    With objExel.ActiveSheet.Range("A" & exRow & ":F" & exRow)
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

'$odbc15$
'��� ��������� ������ � Excel ����� ���� �� ������� ��������, ���������� �
'������� "web" (��. ���������� ������� ������� '� ��������� stime).
'��� ������� � ��� ������������� �� ������(����� � �����������), �������
'�������� ����������� ���������.
'���������� ����� ���������� �� ����. sGuideSeries � ���������� �������
'������� �� sGuideProducts. �� � �� ���� prSeriaId parentKlassId � klassName
'���� �������� ��������� �� ���� Comtec $comtec$

Private Sub mnPriceToExcel_Click()
Dim I As Integer, findId As Integer, str As String

MsgBox "����� �������������� ���������� ���� ������� ������ ���� ����������� " & _
"�� ���� ������. ��. ����������� $comtec$ � ���� ���������.", , ""
Exit Sub



'�� ����������� ������� ������� �������� ������ Id ���� �����(�����),
'� ������� ���� ���� �� ���� �������.
sql = "SELECT prSeriaId from sGuideProducts Where(((web) = 'web')) GROUP BY prSeriaId;"
Set tbProduct = myOpenRecordSet("##412", sql, dbOpenDynaset)
If tbProduct Is Nothing Then Exit Sub

'� ���� ����� �������� ����� ���� ����� ������(�� Id)====================
'Set tbGuide = myOpenRecordSet("##413", "select * from sGuideSeries", dbOpenForwardOnly)
'If tbGuide Is Nothing Then Exit Sub
'tbGuide.index = "PrimaryKey"

ReDim NN(0): I = 0
While Not tbProduct.EOF
    I = I + 1
    ReDim Preserve NN(I): NN(I) = Format(tbProduct!prSeriaId, "0000")
    findId = tbProduct!prSeriaId

AA: ' tbGuide.Seek "=", findId
'    If tbGuide.NoMatch Then msgOfEnd ("##414")
    sql = "SELECT seriaName, parentSeriaId from sGuideSeries " & _
    "WHERE (((seriaId)=" & findId & "));"
    If Not byErrSqlGetValues("##414", sql, str, findId) Then tbProduct.Close: Exit Sub
    
'    NN(I) = tbGuide!seriaName & " / " & NN(I) ' � ����� ��������� Id
    NN(I) = str & " / " & NN(I) ' � ����� ��������� Id
'    findId = tbGuide!parentSeriaId
    If findId > 0 Then GoTo AA '� ����� ������� ������ ������� �������������
                               '����� ���� ����� ������, � ������� ��� ������
    tbProduct.MoveNext
Wend
'tbGuide.Close
tbProduct.Close
'=========================================================================


'���� ���� �� ������� ��������� -------------------------------------------

quickSort NN, 1

On Error GoTo ERR2

    Set objExel = New Excel.Application
    objExel.Visible = True
    objExel.SheetsInNewWorkbook = 1
    objExel.Workbooks.Add
    objExel.ActiveSheet.Cells(1, 2).value = "�����-���� �� " & Format(Now(), "dd.mm.yy")
    objExel.ActiveSheet.Cells(1, 2).Font.Bold = True
    exRow = 3
    objExel.ActiveSheet.Columns(1).ColumnWidth = 10
    objExel.ActiveSheet.Columns(2).ColumnWidth = 10
    objExel.ActiveSheet.Columns(3).ColumnWidth = 50
    objExel.ActiveSheet.Columns(4).HorizontalAlignment = xlHAlignRight
    objExel.ActiveSheet.Columns(5).HorizontalAlignment = xlHAlignRight
    objExel.ActiveSheet.Columns(6).HorizontalAlignment = xlHAlignRight
    objExel.ActiveSheet.Columns(7).HorizontalAlignment = xlHAlignRight

'------------------------------------------------------------------------

For I = 1 To UBound(NN) ' ������� ���� �����
  str = NN(I)
  findId = right$(str, 4) ' ��������� �� ���� ������ id ������

'$comtec$  ����� ������ �� ����.sGuideProducts � �� �� ���� ���� �������� ��
'����������� �� ���� Comtec ������ �� ����.������������ � ���������
'�����-�� ������� ������� �� ��������� stime:
'"�����"    "���"   "web"   "��������"    ������   "1-5"   "���."
'SortNom   prName    web    prDescript    prSize   Cena4    page
  
  sql = "SELECT prName, prDescript, prSize, Cena4, page From sGuideProducts " & _
  "Where (((prSeriaId)=" & findId & " AND ((web) = 'web'))) ORDER BY SortNom;" 'prName;"

  Set tbProduct = myOpenRecordSet("##415", sql, dbOpenDynaset)
  If Not tbProduct Is Nothing Then
    If Not tbProduct.BOF Then
      bilo = False
      While Not tbProduct.EOF

'���� ���� �� ������� ��������� (����� �������� ��������� �����)------------
        If Not bilo Then
            bilo = True
            
            With objExel.ActiveSheet.Range("A" & exRow & ":H" & exRow)
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).Weight = xlThin
            End With
            
            str = left$(str, Len(str) - 6)
            objExel.ActiveSheet.Cells(exRow, 2).value = str
            objExel.ActiveSheet.Cells(exRow, 2).Font.Bold = True
            objExel.ActiveSheet.Cells(exRow, 8).Borders(xlEdgeRight). _
            Weight = xlMedium
            
            exRow = exRow + 1
            objExel.ActiveSheet.Range("A" & exRow & ":H" & exRow). _
            Borders(xlEdgeBottom).Weight = xlThin
            
            objExel.ActiveSheet.Cells(exRow, 1).value = "���"
            objExel.ActiveSheet.Cells(exRow, 2).value = "������[��]"
            objExel.ActiveSheet.Cells(exRow, 3).value = "��������"
            
            gain2 = 0
            gSeriaId = findId
            If getGainAndHead Then
                objExel.ActiveSheet.Cells(exRow, 4).value = " " & head1
                objExel.ActiveSheet.Cells(exRow, 5).value = " " & head2
                objExel.ActiveSheet.Cells(exRow, 6).value = " " & head3
                objExel.ActiveSheet.Cells(exRow, 7).value = " " & head4
                objExel.ActiveSheet.Cells(exRow, 8).value = "    ���."
            End If
            cErr = setVertBorders(xlThin)
            If cErr <> 0 Then GoTo ERR2
            exRow = exRow + 1
        End If
'---------------------------------------------------------------------------
'����� �������� ��������� �� ������� ������� ������
        
        objExel.ActiveSheet.Cells(exRow, 1).value = tbProduct!prName
        objExel.ActiveSheet.Cells(exRow, 2).value = tbProduct!prSize
        objExel.ActiveSheet.Cells(exRow, 3).value = tbProduct!prDescript
        objExel.ActiveSheet.Cells(exRow, 4).value = Chr(160) & Format(tbProduct!Cena4, "0.00")
        If gain2 > 0 Then
            objExel.ActiveSheet.Cells(exRow, 5).value = Chr(160) & Format(Round(tbProduct!Cena4 * gain2, 1), "0.00")
            objExel.ActiveSheet.Cells(exRow, 6).value = Chr(160) & Format(Round(tbProduct!Cena4 * gain3, 1), "0.00")
            objExel.ActiveSheet.Cells(exRow, 7).value = Chr(160) & Format(Round(tbProduct!Cena4 * gain4, 1), "0.00")
        End If
        objExel.ActiveSheet.Cells(exRow, 8).value = " " & tbProduct!Page
        cErr = setVertBorders(xlThin)
        If cErr <> 0 Then GoTo ERR2
        exRow = exRow + 1:

        tbProduct.MoveNext
      Wend
    End If
    tbProduct.Close
  End If
Next I
With objExel.ActiveSheet.Range("A" & exRow & ":H" & exRow)
    .Borders(xlEdgeTop).Weight = xlMedium
End With

Set objExel = Nothing
Exit Sub

ERR2:
If cErr <> 424 And Err <> 424 Then  ' 424 - �� ��������� ����� ������ ������� ���-�
    MsgBox Error, , "������ 421-" & cErr '##421
End If
Set objExel = Nothing

End Sub

'$comtec$ ���������� �����  sGuideSeries � ��� ��� ����, ������������ �����
'���� �������� ��������� �� ���� Comtec(����������� ������� ����� ��������)
Function getGainAndHead() As Boolean
getGainAndHead = False
sql = "SELECT head1, head2, head3, head4, gain2, gain3, gain4 " & _
"from sGuideSeries WHERE (((sGuideSeries.seriaId)=" & gSeriaId & "));"
If Not byErrSqlGetValues("##416", sql, head1, head2, head3, head4, gain2, _
gain3, gain4) Then Exit Function
getGainAndHead = True
End Function

#Else '----------------------------------------------------------------------
Sub textBoxOrOtgruzFrm()
        If havePredmetiNew Then
            Otgruz.Regim = ""
            GoTo AA
        ElseIf oldUslug Then ' ������ ����� ��� ����� �������� ������
            textBoxInGridCell tbMobile, Grid
        Else
            Otgruz.Regim = "uslug"
AA:         Otgruz.closeZakaz = (Grid.TextMatrix(mousRow, orStatus) = "������")
            Otgruz.Show vbModal
            If IsNumeric(Grid.TextMatrix(mousRow, orOtgrugeno)) And _
            flDelRowInMobile Then delZakazFromGrid
        End If
End Sub
'$odbc15$
Function oldUslug() As Boolean
Dim s As Single, o

oldUslug = False

sql = "SELECT ordered, shipped From Orders WHERE (((numOrder)=" & gNzak & "));"
If Not byErrSqlGetValues("##303", sql, o, s) Then myBase.Close: End

sql = "SELECT outDate, quant from xUslugOut WHERE (((numOrder)=" & gNzak & "));"
'Set tbProduct = myOpenRecordSet("##229", "select * from xUslugOut", dbOpenForwardOnly)
Set tbProduct = myOpenRecordSet("##229", sql, dbOpenForwardOnly)
'If tbProduct Is Nothing Then myBase.Close: End
'tbProduct.index = "Key"
'tbProduct.Seek "=", gNzak
'If tbProduct.NoMatch Then '�.�. �������� �������� �� ������� � �� �����������
If tbProduct.BOF Then '�.�. �������� �������� �� ������� � �� �����������
    If o - s < 0.005 Then
        oldUslug = True
    ElseIf s > 0.005 Then
'���� ���� �������, ����� �� ������ ������� 0< ��������� < �������� � ���. ��� � xUslugOut
'�� 15,12,04 ����� ���� 75 �� ������ "������ ��� ���� ��������"
        tbProduct.AddNew
        tmpDate = "31.08.2003 10:00:00"
        tbProduct!outDate = tmpDate
        tbProduct!numOrder = gNzak
        tbProduct!quant = s
        tbProduct.update
    End If
End If
tbProduct.Close

End Function

Function isNewEtap() As Variant
Dim s As Single

#If onErrorOtlad Then
    On Error GoTo errMsg
    GoTo START
errMsg:
    MsgBox Error, , "������  " & Err & " � �\� isNewEtap"
    End
START:
#End If

isNewEtap = Null

sql = "SELECT Max([eQuant]-[prevQuant]) as max From xEtapByIzdelia " & _
"WHERE ((numOrder)=" & gNzak & ")  " & _
"UNION SELECT Max([eQuant]-[prevQuant]) as max From xEtapByNomenk " & _
"WHERE ((numOrder)=" & gNzak & ");"
'Debug.Print sql
 Set table = myOpenRecordSet("##385", sql, dbOpenDynaset) 'dbOpenTable)
 If table Is Nothing Then Exit Function
 s = -1
 While Not table.EOF ' ������ 2 �����
    s = max(s, table!max)
    table.MoveNext
 Wend
 table.Close
 If s > 0.005 Then
    isNewEtap = True
 ElseIf s <> -1 Then
    isNewEtap = False
 End If
 
End Function

Function havePredmetiNew() As Boolean
Dim s As Single

havePredmetiNew = False
sql = "SELECT quant From xPredmetyByIzdelia " & _
"WHERE numOrder=" & gNzak & " " & _
"UNION SELECT quant From xPredmetyByNomenk " & _
"WHERE numOrder=" & gNzak & ";"
'Debug.Print sql
If Not byErrSqlGetValues("W##221", sql, s) Then myBase.Close: End
If s > 0 Then havePredmetiNew = True

End Function


#End If '---------------------------------------------------------------------




