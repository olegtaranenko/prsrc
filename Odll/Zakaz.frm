VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Zakaz 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����������� ������ � ������� ������"
   ClientHeight    =   5904
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9468
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5904
   ScaleWidth      =   9468
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmCeh 
      Caption         =   "YAG"
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   30
      Top             =   5520
      Width           =   510
   End
   Begin VB.CheckBox ckCeh 
      BackColor       =   &H8000000A&
      Caption         =   "���"
      Height          =   255
      Left            =   6480
      TabIndex        =   29
      Top             =   5280
      Visible         =   0   'False
      Width           =   675
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4572
      Left            =   120
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   6312
      _ExtentX        =   11134
      _ExtentY        =   8065
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmNewUklad 
      BackColor       =   &H8000000A&
      Caption         =   "����� �������"
      Height          =   375
      Left            =   6480
      TabIndex        =   26
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cbO 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Zakaz.frx":0000
      Left            =   8160
      List            =   "Zakaz.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2940
      Width           =   1035
   End
   Begin VB.ComboBox cbM 
      Enabled         =   0   'False
      Height          =   288
      ItemData        =   "Zakaz.frx":0026
      Left            =   6660
      List            =   "Zakaz.frx":0033
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2940
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Left            =   7080
      Top             =   4140
   End
   Begin VB.TextBox tbDateMO 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   7
      Top             =   3360
      Width           =   915
   End
   Begin VB.TextBox tbDateRS 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   4
      Top             =   2220
      Width           =   915
   End
   Begin VB.TextBox tbReadyDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   3
      Top             =   1740
      Width           =   915
   End
   Begin VB.TextBox tbVrVipO 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   8
      Top             =   3780
      Width           =   915
   End
   Begin VB.ComboBox cbStatus 
      Height          =   288
      Left            =   8220
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton cmExit 
      BackColor       =   &H8000000A&
      Cancel          =   -1  'True
      Caption         =   "�����"
      Height          =   315
      Left            =   8340
      TabIndex        =   12
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmRepit 
      BackColor       =   &H8000000A&
      Caption         =   "Cancel"
      Height          =   315
      Left            =   8340
      TabIndex        =   11
      Top             =   4740
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmZapros 
      BackColor       =   &H8000000A&
      Caption         =   "���������"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7560
      TabIndex        =   9
      Top             =   4260
      Width           =   975
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4512
      Left            =   60
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   360
      Width           =   6372
      _ExtentX        =   11240
      _ExtentY        =   7959
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   1535
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "������"
         Object.Width           =   1296
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1164
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "������"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "������"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�������"
         Object.Width           =   1429
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "���.���"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "���.���"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "�.���"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "�.���"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "�.���"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "��.���"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "��.���"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "�����"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "��.���"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmAdd 
      BackColor       =   &H8000000A&
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6780
      TabIndex        =   10
      Top             =   4740
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox tbWorktime 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      TabIndex        =   2
      Top             =   1320
      Width           =   915
   End
   Begin VB.CheckBox ckCehDone 
      BackColor       =   &H8000000A&
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   5520
      Width           =   252
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Label1"
      Height          =   195
      Left            =   8340
      TabIndex        =   28
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label laMO 
      BackColor       =   &H8000000A&
      Caption         =   "�����                    �������"
      Height          =   195
      Left            =   6840
      TabIndex        =   25
      Top             =   2700
      Width           =   2115
   End
   Begin VB.Label laZapas 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "�����:"
      Height          =   195
      Left            =   4680
      TabIndex        =   23
      Top             =   60
      Width           =   675
   End
   Begin VB.Label laError 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1080
      TabIndex        =   22
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label laNomZak 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   288
      Left            =   8220
      TabIndex        =   0
      Top             =   456
      Width           =   912
   End
   Begin VB.Label laVrVipO 
      BackColor       =   &H8000000A&
      Caption         =   "��.���. �������"
      Enabled         =   0   'False
      Height          =   195
      Left            =   6540
      TabIndex        =   21
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label laDateMO 
      BackColor       =   &H8000000A&
      Caption         =   "���� ���.\���."
      Enabled         =   0   'False
      Height          =   195
      Left            =   6540
      TabIndex        =   20
      Top             =   3420
      Width           =   1275
   End
   Begin VB.Label laHnomZak 
      BackColor       =   &H8000000A&
      Caption         =   "� ������"
      Height          =   192
      Left            =   6540
      TabIndex        =   19
      Top             =   480
      Width           =   852
   End
   Begin VB.Label laStatus 
      BackColor       =   &H8000000A&
      Caption         =   "������:"
      Height          =   192
      Left            =   6540
      TabIndex        =   18
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label laDateRS 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "���� �\� (�� �����)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   6540
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label laReadyDate 
      BackColor       =   &H8000000A&
      Caption         =   "���� ������"
      Height          =   195
      Left            =   6540
      TabIndex        =   15
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label laWorkTime 
      BackColor       =   &H8000000A&
      Caption         =   "����� ����������"
      Height          =   255
      Left            =   6540
      TabIndex        =   14
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Label laMess 
      BackColor       =   &H8000000A&
      Height          =   432
      Left            =   420
      TabIndex        =   16
      Top             =   4980
      Width           =   5832
   End
End
Attribute VB_Name = "Zakaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public urgent As Variant ' As String '"y" - ������� �����
Public Regim As String
Public idEquip As Integer
Public idWerk As Integer
Public isUpdated As Boolean  ' ��������� � true ���� ���-�� � ������� ������ ������������� ����������

' M125 - ����� ����� ������ ������
Public festStatusId As Integer


Dim isTimeZakaz As Boolean ' ������������ "����� ���������� � ���"
Dim oldHeight As Integer, oldWidth As Integer
Dim ZbDay As Integer, ZeDay As Integer, ObDay As Integer, OeDay As Integer

Dim FormIsActiv As Boolean
Dim perekr As Double  '������������� ����� MO
Dim parts As String
Dim be_cmRepit As Boolean
Dim tmpMaxDay As Integer
Dim perenos As Integer ' ����� ��������
Dim workChange As Boolean
Dim isMzagruz As Boolean
Dim quantity As Integer
Dim statusIdNew As Integer, statusIdOld As Integer
Dim zakazBean As ZakazVO



Const zgNomZak = 1
Const zgStatus = 2
Const zgVrVip = 3
Const zgNevip = 4
Const zgInDate = 5
Const zgOutDate = 6
Const zgOtlad = 7



Sub lvAddDay(I As Integer)
Dim Left As String, Rollback As String, Worktime As String, Outdatetime As String

Dim item As ListItem, str As String
    str = Format(DateAdd("d", I - 1, curDate), "dd/mm/yy")
    Set item = Zakaz.lv.ListItems.Add(, "k" & I, str)
    day = Weekday(DateAdd("d", I - 1, curDate))
    If day = vbSunday Or day = vbSaturday Then item.ForeColor = &HFF
End Sub

Sub lvAddDays(Optional newLen As Integer = -1)
Dim I As Integer, J As Integer

If newLen < 0 Then newLen = maxDay

J = lv.ListItems.Count
If newLen > J Then ' j=0 ����� startParams �����-�� ������ ���
    For I = J + 1 To newLen
        lvAddDay I
    Next I
ElseIf newLen < J Then
    For I = newLen + 1 To J
        lv.ListItems.Remove "k" & I
    Next I
    End If
End Sub

Private Sub cbM_Click()
cmZapros.Enabled = True
If cbM.Text = "� ������" Or cbM.Text = "�����" Then
    If FormIsActiv Then Zakaz.cmZapros.Enabled = True
    laDateMO.Enabled = True
    tbDateMO.Enabled = True
ElseIf Not (cbO.Text = "� ������" Or cbO.Text = "�����") Then
    laDateMO.Enabled = False
    tbDateMO.Enabled = False
    tbDateMO.Text = ""
End If
End Sub

Private Sub cbO_Click()
cmZapros.Enabled = True
If cbO.Text = "� ������" Or cbO.Text = "�����" Then
    If FormIsActiv Then Zakaz.cmZapros.Enabled = True
    laDateMO.Enabled = True
    tbDateMO.Enabled = True
    laVrVipO.Enabled = True
    tbVrVipO.Enabled = True
    tbVrVipO.Text = zakazBean.WorktimeMO
Else
    If Not (cbM.Text = "� ������" Or cbM.Text = "�����") Then
        laDateMO.Enabled = False
        tbDateMO.Enabled = False
        tbDateMO.Text = ""
    End If
    laVrVipO.Enabled = False
    tbVrVipO.Enabled = False
    tbVrVipO.Text = ""
End If
End Sub

Private Sub cmCeh_Click(Index As Integer)
    idEquip = Index + 1
    'statusIdOld = statusIdNew
    startParams
    'newZagruz ' ���������� � startParams (!?)
End Sub

Private Sub cmNewUklad_Click()
    newZagruz
End Sub

Sub getBegEndDays(Optional Stat As String = "")
Dim priemData As String

If Regim = "setka" Then ' �� F1 F2
    priemData = Format(curDate, "dd.mm.yy")
Else
    priemData = Orders.Grid.TextMatrix(Orders.Grid.row, orData)
End If
If Stat = "�������" Then
    Grid.TextMatrix(quantity, zgInDate) = priemData
    Grid.TextMatrix(quantity, zgOutDate) = tbDateMO.Text
    Grid.TextMatrix(quantity, zgOtlad) = ObDay & " " & OeDay
    Exit Sub
ElseIf Stat <> "" Then
    Grid.TextMatrix(quantity, zgOtlad) = ZbDay & " " & ZeDay
    If IsDate(tbDateRS.Text) Then
        Grid.TextMatrix(quantity, zgInDate) = tbDateRS.Text
    Else
        Grid.TextMatrix(quantity, zgInDate) = priemData
    End If
    Grid.TextMatrix(quantity, zgOutDate) = tbReadyDate.Text
    Exit Sub
End If

'ZbDay = DateDiff("d", curDate, priemData)
'ObDay = ZbDay

If IsDate(tbReadyDate.Text) Then
    ZbDay = DateDiff("d", curDate, priemData)
    ZeDay = DateDiff("d", curDate, tbReadyDate.Text)
End If
If IsDate(tbDateRS.Text) Then
    ZbDay = DateDiff("d", curDate, tbDateRS.Text)
End If
If tbVrVipO.Enabled Then
    ObDay = DateDiff("d", curDate, priemData)
    OeDay = DateDiff("d", curDate, tbDateMO.Text)
End If

End Sub
' ����� � ������ �������
Function zakazToGrid(reg As String, Stat As String, nevip As String) As Boolean
zakazToGrid = False

If reg = "" Then
    quantity = quantity + 1
    Grid.TextMatrix(quantity, zgNomZak) = laNomZak.Caption
    Grid.TextMatrix(quantity, zgStatus) = Stat
    If Stat = "�������" Then
        Grid.TextMatrix(quantity, zgVrVip) = tbVrVipO.Text
    Else
        Grid.TextMatrix(quantity, zgVrVip) = tbWorktime.Text
    End If
    Grid.TextMatrix(quantity, zgNevip) = nevip
    
    getBegEndDays Stat
    
    Grid.AddItem ""
Else
    If laNomZak.Caption = Grid.TextMatrix(Grid.row, zgNomZak) Then
        If Grid.TextMatrix(Grid.row, zgStatus) = "�������" Then
            If Stat = "�������" Then Exit Function
        ElseIf Stat <> "�������" Then
            Exit Function ' ������������ ������ ��(�������-��) �����.������
        End If
    End If
End If
zakazToGrid = True
End Function

'��� ����� ��������� �������
'reg="tail" ��� ����������� ���������(������ ��� ������ �� ���� �����)
'reg = "fromCehZagruz"
'reg = "setka" �� F1,F2 - �.�. ������ ���
'reg = "" - double click at status cell

Sub newZagruz(Optional reg As String = "", Optional equipId As Integer)
Dim S As Double, nevip As Double, I As Integer
Dim bDay As Integer, eDay As Integer, Stat As String, vEquipId As Integer

If equipId <> 0 Then
    vEquipId = equipId
Else
    vEquipId = Me.idEquip
End If

'isMzagruz - true: ���� ������� �������� �� �� ����, �� ���� ��������.
If reg = "fromCehZagruz" Then
    isMzagruz = False
Else
    isMzagruz = (ckCeh.value = 0)
End If

'ZeDay = maxDay + 1  '����� �� ��������� !!��� �������� ���-�� ����� 4 ����
'OeDay = ZeDay
ObDay = -32000: ZbDay = ObDay ' ����, ��� �����.���� �� ��������� (� 4� ������)
If isMzagruz Then getBegEndDays 'ZbDay,ObDay,ZbDay,ObDay (���� ���������)

If reg = "" Then ' �� �� Enter_cell
    Grid.Clear
    Grid.Rows = 2
    Grid.FormatString = "|<� ������|<������|��.���|���-��|���.����|���.����|"
    Grid.ColWidth(0) = 0
    Grid.ColWidth(zgStatus) = 765
    Grid.ColWidth(zgOutDate) = 825
    Grid.ColWidth(zgOtlad) = 525
    quantity = 0
End If


ReDim ost(maxDay): ReDim befOst(maxDay)
Dim firstRes As Double
'firstRes = Round(nr * Nstan * kpd, 1)  '23.11.04
firstRes = nr * Nstan * KPD             '
ost(1) = firstRes
befOst(1) = firstRes
For I = 2 To maxDay
    ost(I) = nomRes(I) * KPD * Nstan                '23.11.04
    befOst(I) = nomRes(I) * KPD * Nstan             '
'    ost(i) = Round(nomRes(i) * kpd * Nstan, 1)     '
'    befOst(i) = Round(nomRes(i) * kpd * Nstan, 1)  '
Next I

'VrVipParts ��������� �� Nevip
sql = "SELECT o.numOrder, oe.workTime, " & _
" DateDiff(day,Now(),oe.outDateTime) AS endDay, " & _
" DateDiff(day,Now(),o.inDate) AS begDay, dateadd(hour, isnull(o.outtime, 12), oe.outDateTime) as outdatetime, " & _
" o.inDate, o.StatusId, oe.Nevip, oc.urgent " & _
vbCr & " FROM Orders o " & _
" JOIN OrdersEquip oe ON oe.numorder = o.numorder AND oe.equipId = " & vEquipId & _
" JOIN OrdersInCeh oc ON o.numOrder = oc.numOrder " & _
" Where (o.StatusId = 1 Or o.StatusId = 5) " & _
vbCr & " UNION ALL " _
& vbCr & " SELECT o.numOrder, oe.workTime, DateDiff(day,Now(),oe.outDateTime) AS endDay, " & _
" DateDiff(day,Now(),o.DateRS) AS begDay, dateadd(hour, isnull(o.outtime, 12), oe.outDateTime) as outdatetime, " & _
" o.DateRS, o.StatusId, oe.Nevip, oc.urgent " & _
vbCr & " FROM Orders o " & _
" JOIN OrdersEquip oe ON oe.numorder = o.numorder AND oe.equipId = " & vEquipId & _
" JOIN OrdersInCeh oc ON o.numOrder = oc.numOrder " & _
" Where (o.StatusId = 2 Or o.StatusId = 3) " & _
vbCr & " UNION ALL " _
& vbCr & " SELECT o.numOrder, oe.workTimeMO, DateDiff(day,Now(),oc.DateTimeMO) AS endDay, " & _
" DateDiff(day,Now(),o.inDate) AS begDay, dateadd(hour, isnull(o.outtime, 12), oc.DateTimeMO) as outdatetime, " & _
" o.inDate, 1 AS StatusId, -1 AS Nevip, '' AS urgent " & _
vbCr & " FROM Orders o " & _
" JOIN OrdersInCeh oc ON o.numOrder = oc.numOrder " & _
" JOIN OrdersEquip oe ON oe.numorder = o.numorder AND oe.equipId = " & vEquipId & _
" Where oc.statO = '� ������' " & " ORDER BY "

If isMzagruz Then
    sql = sql & "4 DESC" ' � ������� ���������� ���� ������
Else
    sql = sql & "3" ' � ������� ����������  ���� �����
End If
'Debug.Print sql
Set tbOrders = myOpenRecordSet("##370", sql, dbOpenForwardOnly) ', dbOpenDynaset)
If tbOrders Is Nothing Then Exit Sub
While Not tbOrders.EOF
    bDay = tbOrders!begDay '  ���. Now()
    If Not IsNull(tbOrders!endDay) Then
        eDay = tbOrders!endDay '  ���. Now()
    Else
        eDay = 0
    End If
'    If eDay > maxDay Then msgOfEnd ("##371")
    
    If isMzagruz Then '� ���������� ���.����� �.������� ������ �� �����, ���-
        If tbOrders!Numorder = laNomZak.Caption Then GoTo NXT ' ���� � ����
        '��� ����������, �.�. � ���� ����� ����� ����� ����� ������
        
'��������� ���.����� ���, ����� �� �������� ������� ���������� ���� ������
'        If eDay > OeDay Or (eDay = OeDay And bDay <= ObDay) Then ' �� �������� ����������
        If bDay <= ObDay Then  ' �� �������� ����������
            dayCorrect ObDay, OeDay
            ukladka ost, OeDay, ObDay, tbVrVipO.Text '�������� �������  (� bef �� ��������)
            If Not zakazToGrid(reg, "�������", tbVrVipO.Text) Then GoTo EN1
'            OeDay = maxDay + 1 ' ����� ����� �� ����������
            ObDay = -32000 ' ����� ����� �� ����������
        End If
'        If eDay > ZeDay Or (eDay = ZeDay And bDay <= ZbDay) Then ' �� �������� ����������
        If bDay <= ZbDay Then ' �� �������� ����������
            dayCorrect ZbDay, ZeDay, CStr(Me.urgent)
            ukladka ost, ZeDay, ZbDay, tbWorktime.Text '�������� ������� (� bef �� ��������)
            If Not zakazToGrid(reg, cbStatus.Text, tbWorktime.Text) Then GoTo EN1
            'ZeDay = maxDay + 1 ' ����� ����� �� ����������
            ZbDay = -32000 ' ����� ����� �� ����������
        End If
'        If tbOrders!numOrder = laNomZak.Caption Then GoTo NXT '� �� �� ����
    End If
    
    If eDay > maxDay Then msgOfEnd ("##371")
    
'    dayCorrect bDay, eDay, tbOrders!urgent ����.������������� �����. ������� �������� � ����, ��� �.�. ������ ������ � � � � ����
    dayCorrect bDay, eDay, ""
    
    If tbOrders!nevip = -1 Then '"�������"
        nevip = tbOrders!Worktime
    Else
        nevip = Round(tbOrders!Worktime * tbOrders!nevip, 2)
    End If
    If isMzagruz Then
        ukladka ost, eDay, bDay, nevip '�������� �������
        ukladka befOst, eDay, bDay, nevip '�������� �������
    Else
        ukladka ost, bDay, eDay, nevip
        If tbOrders!StatusId = 1 Or tbOrders!StatusId = -1 Then _
            ukladka befOst, bDay, eDay, nevip ' �����(� ���. � �������)
    End If
    
    If reg = "" Then
      quantity = quantity + 1
      Grid.TextMatrix(quantity, zgNomZak) = tbOrders!Numorder
'      If tbOrders!StatusId = -1 Then
      If tbOrders!nevip = -1 Then '"�������"
        Grid.TextMatrix(quantity, zgStatus) = "�������"
      Else
        Grid.TextMatrix(quantity, zgStatus) = status(tbOrders!StatusId)
      End If
      Grid.TextMatrix(quantity, zgVrVip) = tbOrders!Worktime
      Grid.TextMatrix(quantity, zgNevip) = nevip
      Grid.TextMatrix(quantity, zgInDate) = Format(tbOrders!inDate, "dd.mm.yy")
      Grid.TextMatrix(quantity, zgOutDate) = Format(tbOrders!Outdatetime, "dd.mm.yy")
      Grid.TextMatrix(quantity, zgOtlad) = bDay & " " & eDay
      Grid.AddItem ""
    End If
    
    If reg = "tail" Then ' �� Enter_cell
      If tbOrders!Numorder = Grid.TextMatrix(Grid.row, zgNomZak) Then
        If Grid.TextMatrix(Grid.row, zgStatus) = "�������" Then
          If tbOrders!nevip = -1 Then GoTo EN1 '"�������"
        ElseIf tbOrders!nevip <> -1 Then
          GoTo EN1 ' ������������ ������ ��(�������-��) �����.������
        End If
      End If
    End If
NXT:
    tbOrders.MoveNext
Wend

If isMzagruz Then
'���� � ���� ��� ������� ��� ����� ��� ������� �� ZeDay � �.�. � OeDay ��
'�� ���������, �.�. ���� ��������� � ���������
'  If OzDay < maxDay + 1 Then  '
  If ObDay > -32000 Then  '
    dayCorrect ObDay, OeDay
    ukladka ost, OeDay, ObDay, tbVrVipO.Text '�������� �������  (� bef �� ��������)
    zakazToGrid reg, "�������", tbVrVipO.Text
  End If
'  If ZeDay < maxDay + 1 Then
  If ZbDay > -32000 Then
    dayCorrect ZbDay, ZeDay
    ukladka ost, ZeDay, ZbDay, tbWorktime.Text '�������� ������� (� bef �� ��������)
    zakazToGrid reg, cbStatus.Text, tbWorktime.Text
  End If
End If

If reg = "" And quantity > 0 Then Grid.removeItem Grid.Rows - 1
EN1:
tbOrders.Close

If reg = "fromCehZagruz" Then Exit Sub

If ckCeh.value = 0 Then
  For I = 1 To maxDay
    lv.ListItems("k" & I).SubItems(zkMbef) = Round(befOst(I), 1) '23.11.04
    lv.ListItems("k" & I).SubItems(zkMzagr) = _
               Round(nomRes(I) * KPD * Nstan - befOst(I), 1)
    lv.ListItems("k" & I).ListSubItems(zkMbef).Bold = False
    lv.ListItems("k" & I).ListSubItems(zkMbef).ForeColor = 0
    If reg = "setka" Then
        If befOst(I) < 0 Then
            lv.ListItems("k" & I).ListSubItems(zkMbef).Bold = True
            lv.ListItems("k" & I).ListSubItems(zkMbef).ForeColor = 200
        End If
    Else
        lv.ListItems("k" & I).SubItems(zkMost) = Round(ost(I), 1) '23.11.04
   
        lv.ListItems("k" & I).ListSubItems(zkMost).Bold = False
        lv.ListItems("k" & I).ListSubItems(zkMost).ForeColor = 0
        If befOst(I) < 0 Then
            lv.ListItems("k" & I).ListSubItems(zkMbef).Bold = True
            lv.ListItems("k" & I).ListSubItems(zkMbef).ForeColor = 200
            If ost(I) < befOst(I) Then GoTo AA
        ElseIf ost(I) < 0 Then
AA:         lv.ListItems("k" & I).ListSubItems(zkMost).Bold = True
            lv.ListItems("k" & I).ListSubItems(zkMost).ForeColor = 200
        ElseIf ost(I) <> befOst(I) Then
            lv.ListItems("k" & I).ListSubItems(zkMost).Bold = True
        End If
    End If
  Next I
  lv.ListItems("k1").SubItems(zkMzagr) = Round(firstRes - befOst(1), 1) '23.11.04
Else
  For I = 1 To maxDay
   lv.ListItems("k" & I).SubItems(zkCost) = Round(ost(I), 1) '23.11.04
   lv.ListItems("k" & I).SubItems(zkCliv) = _
            Round(nomRes(I) * KPD * Nstan - befOst(I), 1)
   lv.ListItems("k" & I).SubItems(zkCzagr) = _
            Round(nomRes(I) * KPD * Nstan - ost(I), 1)
  Next I
  lv.ListItems("k1").SubItems(zkCzagr) = Round(firstRes - ost(1), 1) '23.11.04
  lv.ListItems("k1").SubItems(zkCliv) = Round(firstRes - befOst(1), 1) '23.11.04
End If

End Sub
    
Sub dayCorrect(bDay As Integer, eDay As Integer, Optional urgen As String = "")
    bDay = bDay + 1: eDay = eDay + 1 '����-� ���-�� DateDiff(,now())
    If bDay < 1 Then bDay = 1
    If urgen = "" Then ' �� �������
        eDay = getPrev2DayRes_(eDay) '�� 2 ���
    End If
    If bDay > eDay Then bDay = eDay
End Sub

Sub ukladka(ost() As Double, bDay As Integer, eDay As Integer, ByVal nevip As Double)
Dim I As Integer, stp As Integer

stp = 1
If bDay > eDay Then stp = -1
For I = bDay To eDay Step stp
    If ost(I) > 0 Then ' �� ����� ������ �� ������������
        ost(I) = Round(ost(I) - nevip, 2)
        If ost(I) >= 0 Then
            nevip = 0
            Exit Sub
        End If
        nevip = -ost(I)
        ost(I) = 0
    End If
Next I
If nevip > 0 Then
    I = max(bDay, eDay)
    ost(I) = ost(I) - nevip
End If
End Sub

Sub formMaximize()
Dim oldWidth As Integer
    Me.WindowState = vbMaximized
     cmNewUklad.Visible = True
    Grid.Visible = True
    ckCeh.Visible = True
    Label1.Visible = True
    lv.ColumnHeaders(zkCzagr + 1).Width = 680
    lv.ColumnHeaders(zkCost + 1).Width = 680
    lv.ColumnHeaders(zkCliv + 1).Width = 680
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If Shift = vbCtrlMask And KeyCode = vbKeyO Then
    formMaximize
End If
End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If Me.WindowState = vbMinimized Then Exit Sub
'lbHide

On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width

'lv.Height = lv.Height + h
lv.Width = lv.Width + w
laHnomZak.Left = laHnomZak.Left + w
laNomZak.Left = laNomZak.Left + w
laStatus.Left = laStatus.Left + w
cbStatus.Left = cbStatus.Left + w
laWorkTime.Left = laWorkTime.Left + w
tbWorktime.Left = tbWorktime.Left + w
laReadyDate.Left = laReadyDate.Left + w
tbReadyDate.Left = tbReadyDate.Left + w
laDateRS.Left = laDateRS.Left + w
tbDateRS.Left = tbDateRS.Left + w
laMO.Left = laMO.Left + w
cbM.Left = cbM.Left + w
cbO.Left = cbO.Left + w
laDateMO.Left = laDateMO.Left + w
tbDateMO.Left = tbDateMO.Left + w
laVrVipO.Left = laVrVipO.Left + w
tbVrVipO.Left = tbVrVipO.Left + w
cmZapros.Left = cmZapros.Left + w
cmAdd.Left = cmAdd.Left + w
cmRepit.Left = cmRepit.Left + w
cmExit.Left = cmExit.Left + w
cmExit.Top = cmExit.Top + h

End Sub

Private Sub Form_Unload(Cancel As Integer)
    '���� ������ �� �����������:
    If getSystemField("resursLock") = Orders.cbM.Text Then unLockBase
    Orders.Grid_EnterCell ' ��������� ������
    
    Unload Equipment
    'synchOrderRow

End Sub

Private Sub Grid_EnterCell()
Static I As Integer

Grid.CellBackColor = vbButtonFace
If quantity > 0 Then
    I = I + 1
    Label1.Caption = I
    newZagruz "tail"
End If
End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_LostFocus()
Grid_LeaveCell
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)

End Sub

Private Sub laNomZak_Click()
    Dim Left As String, Worktime As String, tbWorktime As String, Rollback As String
End Sub

Private Sub tbDateMO_GotFocus()
'If FormIsActiv Then Zakaz.cmZapros.Enabled = True
If tbDateMO.Text = "" Then
    tbDateMO.Text = Format(curDate, "dd/mm/yy")
End If
tbDateMO.SelStart = 0
tbDateMO.SelLength = 2

End Sub

Private Sub cbStatus_Click()
Dim noClick As Boolean

'If FormIsActiv Then Zakaz.cmZapros.Enabled = True
Dim I As Integer

statusIdNew = statId(cbStatus.ListIndex)

cmZapros.Enabled = statusIdOld <> statusIdNew

For I = 0 To UBound(Equip) - 1
    If ckCehDone(I).Tag = CStr(statusIdNew) Then
        ckCehDone(I).value = 1
    Else
        ckCehDone(I).value = 0
    End If
Next I

tbWorktime.Text = zakazBean.Worktime
If Not IsNull(zakazBean.DateRS) Then
    tbDateRS.Text = Format(zakazBean.DateRS, "dd.mm.yy")
End If

If Not IsNull(zakazBean.Outdatetime) Then
    tbReadyDate.Text = Format(zakazBean.Outdatetime, "dd.mm.yy")
End If


If cbStatus.Text = "� ������" Then
    laMO.Enabled = False
    cbM.Enabled = False
    cbO.Enabled = False
    tbDateMO.Enabled = False
    laVrVipO.Enabled = False
    tbVrVipO.Enabled = False
    'tbWorktime.SetFocus
ElseIf cbStatus.Text = "��������" Then
    cbM.Enabled = True
    cbO.Enabled = True
    laMO.Enabled = True
Else
    laMO.Enabled = False
    cbM.Enabled = False
    cbO.Enabled = False
    tbDateMO.Enabled = False
    laVrVipO.Enabled = False
    tbVrVipO.Enabled = False
End If

If cbStatus.Text = "��������" Or cbStatus.Text = "������" Then
    tbDateRS.Enabled = True             ' ������ ������������
    laDateRS.Enabled = True
    Zakaz.laWorkTime.Enabled = True
    Zakaz.laReadyDate.Enabled = True
    Zakaz.tbReadyDate.Enabled = True
    Zakaz.tbWorktime.Enabled = True
ElseIf cbStatus.Text = "� ������" Or cbStatus.Text = "�������" Then
    tbDateRS.Enabled = False
    laDateRS.Enabled = False
    Zakaz.laWorkTime.Enabled = True
    Zakaz.laReadyDate.Enabled = True
    Zakaz.tbReadyDate.Enabled = True
    Zakaz.tbWorktime.Enabled = True
    If tbWorktime.Enabled And tbWorktime.Visible Then tbWorktime.SetFocus
Else
    laWorkTime.Enabled = False
    laReadyDate.Enabled = False
    tbReadyDate.Text = ""
    tbReadyDate.Enabled = False
    tbWorktime.Text = ""
    tbWorktime.Enabled = False
    tbDateRS.Text = ""
    tbDateRS.Enabled = False
    laDateRS.Enabled = False
End If
End Sub


'$odbc08$
Private Sub cmAdd_Click()
Dim I As Integer, str As String, item As ListItem, S As Double, T As Double
Dim VrVip As String, VrVipO As String, editWorkTime As Boolean
Dim Worktime As String

'MaxDay = tmpMaxDay ������� ��� ��� ������ �� �����

Timer1.Enabled = False

sql = "SELECT o.statusId, oe.worktime " _
& " from Orders o" _
& " left join OrdersEquip oe on oe.numorder = o.numorder and oe.equipId = " & idEquip _
& " WHERE o.numOrder = " & gNzak
Set tbOrders = myOpenRecordSet("##30", sql, dbOpenForwardOnly) '$#$

If tbOrders.BOF Then
    tbOrders.Close
AA: If getSystemField("resursLock") = Orders.cbM.Text Then unLockBase '���� ������ �� ���������
    MsgBox "�������� �� ��� ������. �������� ������", , "����� �� ������!!!"
    Exit Sub
End If

Dim workTimeOld As Double
workTimeOld = 0
If Not IsNull(tbOrders!Worktime) Then
    workTimeOld = tbOrders!Worktime
End If
'statusIdOld = tbOrders!StatusId
tbOrders.Close

wrkDefault.BeginTrans

' ��������� ��������� ������
If (statusIdOld = 0 Or statusIdOld = 7) And statusIdNew <> 0 And statusIdNew <> 7 Then
    visits "+"
ElseIf Not (statusIdOld = 0 Or statusIdOld = 7) And (statusIdNew = 0 Or statusIdNew = 7) Then
    visits "-"
End If

If statusIdNew = 7 Then delZakazFromReplaceRS ' ���� ������������ ����� ��� ����

If statusIdNew <> statusIdOld Or (neVipolnen_O = 0 And neVipolnen = 0) Then
    editWorkTime = False    '
Else                        '���� ���-�� �������-�� � ������ �� ��������
    editWorkTime = True     '�� ��� �����. ������� ��������� ��.���.
End If

str = tbReadyDate.Text
If str = "" Then
'    tbOrders!outDateTime = Null
    str = "Null"
Else
    str = "'" & "20" & Mid$(str, 7, 2) & "-" & Mid$(str, 4, 2) & "-" & Left$(str, 2)
    sql = Orders.Grid.TextMatrix(Orders.mousRow, orVrVid)
    If sql = "" Then
'        tbOrders!outDateTime = tbReadyDate.Text
        str = str & "'"
    Else
        str = str & " " & sql & ":00'"
'       tbOrders!outDateTime = str
    End If
End If

Dim v_outDateTime As String
v_outDateTime = str


If tbDateRS.Enabled = True Then
'    tbOrders!dateRS = tbDateRS.Text
    str = tbDateRS.Text
    str = "'" & "20" & Mid$(str, 7, 2) & "-" & Mid$(str, 4, 2) & _
    "-" & Left$(str, 2) & "'"
Else
'    tbOrders!dateRS = Null
    str = "Null"
End If
sql = "UPDATE Orders SET dateRS = " & str & " WHERE Orders.numOrder = " & gNzak
'MsgBox sql
If myExecute("##392", sql) <> 0 Then GoTo ER1

sql = "SELECT * from OrdersInCeh WHERE numOrder = " & gNzak
Set tbOrders = myOpenRecordSet("##01", sql, dbOpenForwardOnly)

Worktime = workTimeOld ' ��� ������, ���� �� ��������
If Not tbOrders.BOF Then
    If isTimeZakaz Then
       If workChange Then
         If (statusIdNew = 1 Or statusIdNew = 5) And editWorkTime Then '�������� � ������ ��� �������
            Worktime = Round(workTimeOld + tbWorktime.Text - neVipolnen, 1) '����� � ������ ���������
            sql = "UPDATE OrdersEquip SET Nevip = " & tbWorktime.Text / Worktime _
             & " WHERE numOrder =" & gNzak & " AND equipId = " & idEquip
            If myExecute("##393", sql) <> 0 Then GoTo ER1
         Else
            Worktime = tbWorktime.Text
         End If
       End If
       sql = "UPDATE OrdersInCeh SET urgent = '" & urgent & _
       "' WHERE OrdersInCeh.numOrder = " & gNzak
       If myExecute("##403", sql) <> 0 Then GoTo ER1
       GoTo DD
    Else
        sql = "DELETE from OrdersInCeh WHERE numOrder = " & gNzak
        If myExecute("##394", sql) <> 0 Then GoTo ER1
        Worktime = 0
    End If
Else
    If isTimeZakaz Then
        Worktime = tbWorktime.Text
        sql = "INSERT INTO OrdersInCeh ( numOrder, urgent)" & _
        "SELECT " & gNzak & ",'" & urgent & "'"
        If myExecute("##395", sql) <> 0 Then GoTo ER1
DD:     noClick = True
        Orders.Grid.col = orWerk
        If urgent = "y" Then
            Orders.Grid.CellForeColor = 200
        Else
            Orders.Grid.CellForeColor = vbBlack
        End If
        Orders.Grid.col = orStatus
        noClick = False
    End If
End If

sql = "UPDATE OrdersEquip SET outDateTime = " & v_outDateTime _
    & ", workTime = " & Worktime _
    & ", statusEquipId = " & statusIdNew _
    & " WHERE numOrder = " & gNzak & " and equipId =" & idEquip
'Debug.Print sql
If myExecute("##391", sql) <> 0 Then GoTo ER1


sql = "UPDATE Orders SET statusId = " & statusIdNew & " WHERE Orders.numOrder =" & gNzak
If myExecute("##396", sql) <> 0 Then GoTo ER1


' ������������ ��� �� ������������ � ������
sql = "SELECT * from OrdersInCeh WHERE numOrder =" & gNzak
Set table = myOpenRecordSet("##02", sql, dbOpenForwardOnly)
bilo = Not table.BOF
table.Close

 If statusIdNew = 3 Then ' ��������
  If cbM.Text = "� ������" Or cbM.Text = "�����" Or _
    cbO.Text = "� ������" Or cbO.Text = "�����" Then
    str = tbDateMO.Text
    str = "'" & "20" & Mid$(str, 7, 2) & Mid$(str, 4, 2) & Left$(str, 2)
    sql = Orders.Grid.TextMatrix(Orders.mousRow, orMOVrVid)
    If sql = "" Then
        str = str & "'"
    Else
        str = str & " " & sql & ":00'"
    End If
  Else
    str = "Null"
  End If
  If cbO.Text = "� ������" Or cbO.Text = "�����" Then
    Worktime = tbVrVipO.Text
  Else
    Worktime = "Null"
  End If
  If bilo Then      '
    sql = "UPDATE OrdersInCeh SET StatM = '" & cbM.Text & "', StatO = '" & cbO.Text & _
    "', DateTimeMO = " & str & _
    " WHERE numOrder = " & gNzak
  Else
    sql = "INSERT INTO OrdersInCeh ( numOrder, StatM, StatO, DatetimeMO ) " & _
    "SELECT " & gNzak & ", '" & _
    cbM.Text & "', '" & cbO.Text & "', " & str
  End If
  'Debug.Print sql
  If myExecute("##397", sql) <> 0 Then GoTo ER1
    
  sql = "UPDATE OrdersEquip SET workTimeMO = " & Worktime & _
    " WHERE numOrder = " & gNzak & " and equipId = " & idEquip
    
  If myExecute("##397.2", sql) <> 0 Then GoTo ER1
    
 End If ' ��������
tbOrders.Close

    
'******** ������� ���� RS ***********************************
If perenos = 1 Then ' ��� ����������� ������� ��
    sql = "INSERT INTO ReplaceRS ( numOrder, newDateIn, newDateRS, newDateOut) " & _
    "SELECT " & gNzak & ", '" & _
    yymmdd(Orders.Grid.TextMatrix(Orders.mousRow, orData)) & "', '" & _
    yymmdd(Orders.Grid.TextMatrix(Orders.mousRow, orDataRS)) & "', '" & _
    yymmdd(Orders.Grid.TextMatrix(Orders.mousRow, orDataVid)) & "';"
'    MsgBox sql
    If myExecute("##399", sql) <> 0 Then GoTo ER1
    GoTo ��
  ElseIf perenos = 2 Then ' ��� ����������� ������� ��
��: ' table.AddNew
    sql = "INSERT INTO ReplaceRS ( numOrder, newDateIn, newDateRS, newDateOut) " & _
    "SELECT " & gNzak & ", '" & Format(Now(), "yyyy-mm-dd") & "', '" & _
    yymmdd(tbDateRS.Text) & "', '" & yymmdd(tbReadyDate.Text) & "';"
    If myExecute("##400", sql) <> 0 Then GoTo ER1
'    GoTo BB
  ElseIf perenos = 3 Then ' ��� ����������� ������� ��
    sql = "SELECT Max(newDateIn) AS MaxDate from ReplaceRS " & _
    "WHERE (((numOrder)=" & gNzak & "));"
    If byErrSqlGetValues("##22", sql, str) Then
      If str <> "" Then
        sql = "UPDATE ReplaceRS SET newDateIn = '" & Format(Now(), "yyyy-mm-dd") & _
        "', newDateRS = '" & yymmdd(tbDateRS.Text) & _
        "', newDateOut = '" & yymmdd(tbReadyDate.Text) & _
        "' WHERE (((numOrder)=" & gNzak & ") AND (newDateIn)= '" & str & "');"
        If myExecute("##401", sql) <> 0 Then GoTo ER1
      End If
    End If
End If
'******************************************************************

If getSystemField("resursLock") = Orders.cbM.Text Then unLockBase '���� ������ �� ���������

wrkDefault.CommitTrans

str = Orders.openOrdersRowToGrid("##16")

On Error Resume Next ' � �����.��������� ���� �� Open logFile ���� Err: ���� ��� ������
Open logFile For Append As #2
Print #2, str
Close #2

Dim nextEquipId As Integer
ckCehDone(idEquip - 1).Tag = statusIdNew
 
If Not chooseTheEquipment(statusIdNew, nextEquipId) Then
    ' refresh the Orders.Grid row
    
    Unload Me
Else
    idEquip = nextEquipId
    startParams
End If

Exit Sub


ER1:
wrkDefault.Rollback
On Error Resume Next
'table.Close
'tbOrders.Close
'tqOrders.Close
tbOrders.Close
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmRepit_Click()
workChange = False
    cmAdd.Enabled = False
    tbReadyDate.Enabled = True
    tbWorktime.Enabled = True
    tbReadyDate.SetFocus
    cmZapros.Enabled = False
    startParams
    Timer1.Enabled = False
    If getSystemField("resursLock") = Orders.cbM.Text Then unLockBase '���� ������ �� ���������
    be_cmRepit = True
    laMess.Caption = ""
End Sub

Private Sub cmStatus_Click()

End Sub

Function getNextDayRes(tmpDay As Integer) As Integer
Dim I As Integer

getNextDayRes = maxDay
If tmpDay = maxDay Then Exit Function
I = tmpDay + 1
While nomRes(I) = 0
    If I = maxDay Then Exit Function
    I = I + 1
Wend
If I = maxDay Then Exit Function
getNextDayRes = I
End Function

Function getPrevDayRes(ByVal iDay As Integer) As Integer
Dim I As Integer


If iDay < 2 Then GoTo EN1

While iDay > 1
    iDay = iDay - 1
    If nomRes(iDay) > 0 Then GoTo EN2
Wend
If iDay > 0 Then GoTo EN2
EN1:
For iDay = 1 To maxDay '
    If nomRes(iDay) > 0 Then Exit For
Next iDay
EN2:
getPrevDayRes = iDay
End Function
Function getPrev2DayRes_(ByVal iDay As Integer) As Integer
Dim I As Integer


If iDay < 3 Then GoTo EN1

While iDay > 1
    iDay = iDay - 1
    If nomRes(iDay) > 0 Then GoTo EN0
Wend
EN0:
While iDay > 1
    iDay = iDay - 1
    If nomRes(iDay) > 0 Then GoTo EN2
Wend
If iDay > 0 Then GoTo EN2
EN1:
For iDay = 1 To maxDay '
    If nomRes(iDay) > 0 Then Exit For
Next iDay
EN2:
getPrev2DayRes_ = iDay
End Function

Function getPrev2DayRes(tmpDay As Integer) As Integer
Dim I As Integer
getPrev2DayRes = 1
If tmpDay < 2 Then Exit Function

I = tmpDay - 1
While nomRes(I) = 0
    If I < 2 Then Exit Function
    I = I - 1
Wend
If I < 2 Then Exit Function
I = I - 1
While nomRes(I) = 0
    If I < 2 Then Exit Function
    I = I - 1
Wend
If I < 2 Then Exit Function
getPrev2DayRes = I
End Function

Function getPrev2Day(tmpDay As Integer) As Integer
getPrev2Day = tmpDay - 1
day = Weekday(DateAdd("d", getPrev2Day - 1, curDate))
While day = vbSaturday Or day = vbSunday
    getPrev2Day = getPrev2Day - 1
    day = Weekday(DateAdd("d", getPrev2Day - 1, curDate))
Wend

getPrev2Day = getPrev2Day - 1
day = Weekday(DateAdd("d", getPrev2Day - 1, curDate))
While day = vbSaturday Or day = vbSunday
    getPrev2Day = getPrev2Day - 1
    day = Weekday(DateAdd("d", getPrev2Day - 1, curDate))
Wend

If getPrev2Day < 1 Then getPrev2Day = 1
End Function

Private Sub cmZapros_Click() ' zagruzFromCeh ���������� ���������� end(beg)Day(MO)
Dim I As Integer, str As String, num As Integer, v As Variant
Dim begDay As Integer, endDay As Integer, begDayMO As Integer, endDayMO As Integer
Dim begDay_ As Integer, endDay_ As Integer ', begDayMO_ As Integer, endDayMO_ As Integer
Dim title As String, msg As String

'cmZapros.Enabled = True
cmAdd.Enabled = False
laMess.Caption = ""
isTimeZakaz = True
perenos = 0
I = statId(cbStatus.ListIndex)

If I = 7 Then ' �������.
    If Not Orders.do_Annul("no_Do") Then Exit Sub
    GoTo BB
ElseIf I = 0 Then  ' ������  (����� � ������ ����� �� �.�.)
    ' �� �������� ������
BB: isTimeZakaz = False
    
    For I = 1 To lv.ListItems.Count
        lv.ListItems("k" & I).SubItems(zkMost) = lv.ListItems("k" & I).SubItems(zkMbef)
       lv.ListItems("k" & I).ListSubItems(zkMost).Bold = False
        lv.ListItems("k" & I).ListSubItems(zkMost).ForeColor = 0
    Next I
    cmAdd.Enabled = True
    Exit Sub
End If

If Not isNumericTbox(tbWorktime, 0, 2000) Then Exit Sub
tbWorktime.Text = Round(tbWorktime.Text, 1)
If Not isDateTbox(tbReadyDate, "fri") Then Exit Sub

tmpDate = CDate(tbReadyDate.Text)
endDay = DateDiff("d", curDate, tmpDate) + 1

maxDay = 0     '��������� ���, �.�. ���� ��� ���.������ ����� ���������
addDays endDay '1: ������ ��� ���� ������, ���� ��� stDay � rMaxDay

If endDay < 1 Then
ErrDate: MsgBox "���� �� ��� ��� � �������.", , "������������ ��������"
        Exit Sub
End If
If endDay > 100 Then _
    If MsgBox("���������� ���� ������  ����� " & endDay & " ����. " & _
        "�������������?", vbYesNo, "��������!!!") = vbNo Then Exit Sub
        
If tbDateRS.Enabled = True Then
    If Not isDateTbox(tbDateRS, "fri") Then Exit Sub
    tmpDate = CDate(tbDateRS.Text)
    begDay = DateDiff("d", curDate, tmpDate) + 1
    If begDay < 1 Then GoTo ErrDate
                
    If begDay > endDay Then
        MsgBox "���� �\� �� ����� ���� ����� ���� ������", , "������������ ����"
        Exit Sub
    End If
                
    str = "����� ����� �\� � ����� ������ ������ ���� �� ������� " & _
        "���� ��� ������� ���!" & Chr(13) & "����� ��� ���������� �������� ������ " & _
        "� ������, �� ������ �������." & Chr(13) & Chr(13) & "���� �� �������, ��� " & _
            "��� ��������� �� ��������� ���������� �������, ������� - <��>"
    sql = "������� �������:"
Else ' " � ������
    begDay = 1
    str = "�� ������ ������� �����. �������������?"
    sql = "��������!!!"
End If
begDay_ = begDay
endDay_ = getPrev2Day(endDay)
endDay = getPrev2DayRes(endDay)

urgent = ""
If endDay_ <= begDay_ Then
    If MsgBox(str, vbYesNo, sql) = vbNo Then Exit Sub
    urgent = "y"
End If
begDay = getNextDayRes(begDay)
If endDay < begDay Then begDay = begDay_ '������� ���������� begDay
If endDay < begDay Then endDay = begDay  '���� �� �������, �� ����� endDay

'******** ������� ���� RS ***********************************
If tbDateRS.Enabled = True Then
If IsDate(Orders.Grid.TextMatrix(Orders.mousRow, orDataRS)) Then ' ���� ���� ��
tmpDate = Orders.Grid.TextMatrix(Orders.mousRow, orDataRS)       ' ����������
If DateDiff("d", tmpDate, tbDateRS.Text) <> 0 Then        '

tmpDate = Orders.Grid.TextMatrix(Orders.mousRow, orData)         ' � ������� ��
If DateDiff("d", tmpDate, curDate) > 0 Then               ' ���� ������ ������
  title = "������� � 1  ������������?"
  str = "����� ��������� ������ 2 �������� ���� �� (� ���� ������)." & _
  Chr(10) & "�� 3-� ��� ���������� ������������ �����!" & Chr(10)
  msg = str & Chr(10) & "���� ������� ��� �������� ������� <��>"
  
  sql = "SELECT ReplaceRS.newDateIn, ReplaceRS.newDateRS, ReplaceRS.newDateOut " & _
  "From ReplaceRS  Where (((ReplaceRS.numOrder) = " & gNzak & ")) " & _
  "ORDER BY ReplaceRS.newDateIn;"
  
  Set table = myOpenRecordSet("##22", sql, dbOpenDynaset) 'dbOpenTable)
  If Not table Is Nothing Then
    If table.BOF Then ' ������ ���� ��� � ReplaceRS
      If MsgBox(msg, vbYesNo, title) = vbNo Then Exit Sub
         perenos = 1
    Else
      table.MoveFirst: I = 0
      While Not table.EOF
        I = I + 1
        table.MoveNext
      Wend
      table.MoveLast
      If DateDiff("d", table!newDateIn, curDate) > 0 Then ' ���� �� ����������
         str = I                                      ' ������ ��� ��
         Mid(title, 11) = str                             ' �� �������
         If MsgBox(msg, vbYesNo, title) = vbNo Then Exit Sub
         perenos = 2
      Else
         title = "������� � " & I - 1
         MsgBox str, , title
         perenos = 3
      End If
    End If 'Table.BOF
    table.Close
  End If 'Not Table Is Nothing
End If ' � ������� �� ���� ������ ������

End If ' ���� ���� ��
End If ' ����������
End If 'tbDateRS.Enabled = True
'*********************************************************
If cbStatus.Text = "��������" Then
    title = "������������ ������ ��"
    If (cbM.Text = "� ������" Or cbM.Text = "�����") And _
    (cbO.Text = "� ������" Or cbO.Text = "�����") Then
        MsgBox "����� � ������� �� ����� ������������ ���� �������� � ���", , title
        Exit Sub
    ElseIf cbM.Text = "" And cbO.Text = "" Then
        MsgBox "��� ������ '������������' ���������� ���������� ������ ������ �(���) �������", , title
        Exit Sub
    End If
ElseIf cbStatus.Text = "�������" Then
    GoTo EE
ElseIf cbStatus.Text = "� ������" Then
    If ((cbO.Text <> "" And cbO.Text <> "���������") _
    Or (cbM.Text <> "" And cbM.Text <> "���������")) And FormIsActiv Then
        MsgBox "��� �������� ������ � ������ ����������, ����� ����� " & _
        "�(���) ������� ���� ����������.", , "������������ ������!"
        cbStatus.Text = "��������"
        Exit Sub
    Else
EE:     tbDateRS.Text = ""
    GoTo DD
    End If
Else
DD: cbM.ListIndex = 0
    cbO.ListIndex = 0
    tbVrVipO.Text = ""
    tbDateMO.Text = ""
End If

endDayMO = 0 ' ����� ��� MO
begDayMO = 0
If cbM.Text = "� ������" Then GoTo AA  '�����
If cbO.Text = "� ������" Then          '�������
    If Not isNumericTbox(tbVrVipO, 0.1, 2000) Then Exit Sub
    tbVrVipO.Text = Round(tbVrVipO.Text, 1)
AA:
    If Not isDateTbox(tbDateMO, "fri") Then Exit Sub
    tmpDate = CDate(tbDateMO.Text)
    endDayMO = DateDiff("d", curDate, tmpDate) + 1
    If endDayMO < 1 Then GoTo ErrDate
    If endDayMO > begDay_ Then ' �� �������������
        MsgBox "���� M��.\���. �� ����� ���� ����� ���� �\�"
        Exit Sub
    End If
    endDayMO = getPrev2DayRes(endDayMO)
    begDayMO = 1
    I = getNextDayRes(begDayMO)
    If I <= endDayMO Then begDayMO = I
    If endDayMO < begDayMO Then endDayMO = begDayMO
End If

If endDayMO - begDayMO + endDay - begDay > 40 Then
    MsgBox "����� ������ ��������, ��� ��������� ����������� �������. " & _
    "���� ����� �������� ������������� ���������, �������� ��������������!" _
      , , "������� �� ����� ���������� ���� �����!"
    Exit Sub
End If

wrkDefault.BeginTrans
myBase.Execute ("update system set resursLock = resursLock")

sql = "select * from System"
'Set tbSystem = myOpenRecordSet("##94", sql, dbOpenForwardOnly)
'If tbSystem Is Nothing Then myBase.Close: End
'tbSystem.Edit
I = 0
     be_cmRepit = False
      str = getSystemField("resursLock")
'     str = tbSystem!resursLock
     If str = "nextDay" Then
'        tbSystem.Update
        wrkDefault.Rollback
        MsgBox "����������, ��� ��� ���� ��� �������� ���� �� ����� ����. " & _
        "�������� �������������� ��� ������� ����, ����� �� �������� ����� � " & _
        "������������� �������� � �����.", , _
        "������ � �������� ������������!"
        GoTo CC
     End If
     While str <> "" And str <> Orders.cbM.Text
'        tbSystem.Update
        wrkDefault.Rollback
        cmZapros.Enabled = False
        laMess.ForeColor = 200
        laMess.Caption = I & " ���: ������ � �������� �������� ����� " & _
        "���������� " & Chr(34) & str & Chr(34) & Chr(13) _
        & Chr(10) & ". �����."
        delay (1)
        I = I + 1
        If be_cmRepit Then
            'cmZapros.Enabled = True
CC:         'tbSystem.Close
            Exit Sub
        End If
        wrkDefault.BeginTrans
        myBase.Execute ("update system set resursLock = resursLock")
'        tbSystem.Edit
        str = getSystemField("resursLock")
        'str = tbSystem!resursLock
     Wend
     'cmZapros.Enabled = True
     myBase.Execute ("update system set resursLock = '" & Orders.cbM.Text & "'")
'tbSystem!resursLock = Orders.cbM.Text
'tbSystem.Update
wrkDefault.CommitTrans
'tbSystem.Close
laMess.Caption = ""

zagruzFromCeh idEquip, gNzak ' � delta(), Ostatki()  !!!����� gNzak

tmpMaxDay = getResurs(idEquip)  ' ���-� nomRes()
Zakaz.lvAddDays tmpMaxDay '������� ��� ��������� ��������� ������(���) �
'������� �������� �.�. �������� �. ��������� ������ ���� ������
    
For I = 1 To tmpMaxDay
    lv.ListItems("k" & I).SubItems(zkResurs) = Round(nomRes(I) * KPD * Nstan, 1)
Next I

newZagruz

v = lv.ListItems("k1").SubItems(zkMost)
If Not IsNumeric(v) Then v = 0
I = getNextDay(1)
laZapas.Caption = Round(nomRes(I) * KPD * Nstan + v, 1)

If cmRepit.Visible Then '  �� �� <F1> <F2>
    tiki = 11
    cmAdd.Enabled = True
    Timer1.Interval = 1 ' ������ ���� �����
    Timer1.Enabled = True
End If

End Sub

Private Sub Form_Activate()
FormIsActiv = True
End Sub


Private Sub cehSelectorAccess(cehIndex As Integer, action As Boolean, syncStatus As Boolean)
    ckCehDone(cehIndex).Visible = action
    ckCehDone(cehIndex).value = IIf(syncStatus, 1, 0)
    cmCeh(cehIndex).Visible = action
    cmCeh(cehIndex).Enabled = action
End Sub


Private Sub cehSelectorsInit(action As Boolean)
Dim I As Integer
    For I = 0 To UBound(Equip) - 1
        ckCehDone(I).Visible = False
        'ckCehDone(I).Enabled = False
        cmCeh(I).Visible = False
        cmCeh(I).Enabled = False
    Next I
    
End Sub


Private Function chooseTheEquipment(orderStatusId As Integer, ByRef suggestedCehId As Integer) As Boolean

    Dim I As Integer
    Dim firstVisibleId As Integer
    firstVisibleId = -1
    chooseTheEquipment = True
    For I = 0 To UBound(Equip) - 1
        If ckCehDone(I).Tag <> CStr(orderStatusId) And ckCehDone(I).Tag <> "" Then
            suggestedCehId = I + 1
            Exit Function
        End If
        If firstVisibleId <> 0 And ckCehDone(I).Tag <> "" Then
            firstVisibleId = I
        End If
    Next I
    chooseTheEquipment = False
    suggestedCehId = firstVisibleId + 1
    
End Function

' returns -1 ���� ��� �� ������ ������������
' � ��������� ������ - statusId ������
Private Function InitZagruz() As Integer

    Dim myCehId As Integer, cehCtlIndex As Integer, statusIsSync As Boolean
    sql = "select oe.*, o.statusId  " _
    & " from OrdersEquip oe " _
    & " join orders o on o.numorder = oe.numorder" _
    & " where oe.numorder = " & gNzak
    
    
    Dim atLeastOne As Boolean
    atLeastOne = False
    Set tbOrders = myOpenRecordSet("##273", sql, dbOpenForwardOnly) ', dbOpenDynaset)
    statusIdOld = 0
    If tbOrders Is Nothing Then Exit Function
    While Not tbOrders.EOF
        myCehId = tbOrders("equipId")
        cehCtlIndex = myCehId - 1
        statusIdOld = tbOrders!StatusId
        If Not IsNull(tbOrders!statusEquipID) Then
            statusIsSync = tbOrders!statusEquipID = tbOrders!StatusId
            ' � ��� �������� ckCehDone() ������ ������ �� ������������ �� ������ �������� �����.
            ' ����� ��� ����� ������� ������ ����� �������� ��������� �� �� �� �������� �� ��������������
            ckCehDone(cehCtlIndex).Tag = tbOrders!statusEquipID
        Else
            statusIsSync = tbOrders!StatusId = 0
            ckCehDone(cehCtlIndex).Tag = 0
        End If
        cehSelectorAccess cehCtlIndex, True, statusIsSync
        atLeastOne = True
        tbOrders.MoveNext
    Wend
    tbOrders.Close
    
    If Not atLeastOne Then
        ' warning: no ceh assigned
    Else
        Dim newEquipId As Integer
        chooseTheEquipment statusIdOld, newEquipId
        idEquip = newEquipId
        
    End If
    
    InitZagruz = atLeastOne
    
End Function

Private Sub Form_Load()
Dim I As Integer, str As String
FormIsActiv = False
be_cmRepit = True
workChange = False
oldHeight = Me.Height
oldWidth = Me.Width

lv.ColumnHeaders(zkHide + 1).Width = 0

    For I = 1 To UBound(Equip) - 1
        Const HShift = 960
        Load ckCehDone(I)
        Load cmCeh(I)
        ckCehDone(I).Left = ckCehDone(I - 1).Left + HShift
        cmCeh(I).Left = cmCeh(I - 1).Left + HShift
        cmCeh(I).Caption = Equip(I + 1)
    Next I


cehSelectorsInit False

statusIdNew = -1

If festStatusId > 0 Then
    statusIdNew = festStatusId
End If


If Regim = "" Then
    If InitZagruz Then
    End If
End If

startParams

End Sub

Private Sub tbDateRS_GotFocus()
If FormIsActiv Then Zakaz.cmZapros.Enabled = True
tbDateRS.SelStart = 0
tbDateRS.SelLength = 2

End Sub

Private Sub tbReadyDate_GotFocus()
'If FormIsActiv Then Zakaz.cmZapros.Enabled = True
tbReadyDate.SelStart = 0
tbReadyDate.SelLength = 2

Me.cmZapros.Enabled = IsNumeric(tbWorktime.Text) And IsDate(tbReadyDate)

End Sub

Private Sub tbReadyDate_KeyDown(KeyCode As Integer, Shift As Integer)
Dim S As Double, I As Integer
If KeyCode = vbKeyReturn Then

If tbDateRS.Enabled Then
  If isDateTbox(tbReadyDate, "fri") Then
    S = Round(CDbl(tbWorktime.Text), 1)
    I = -(Int((CDbl(S) - 0.05) / 3) + 1 + 2) ' + 2 - ���� ��� �� ����. �����
    getWorkDay I, tbReadyDate.Text ' ���� tmpDate
    If tmpDate < curDate Then tmpDate = curDate
    tbDateRS.Text = Format(tmpDate, "dd.mm.yy")
  End If
End If

End If
Me.cmZapros.Enabled = IsNumeric(tbWorktime.Text) And IsDate(tbReadyDate)

End Sub

Private Sub tbVrVipO_Change()
If FormIsActiv Then
    Zakaz.cmZapros.Enabled = True
End If
End Sub

Private Sub tbWorktime_Change()
If FormIsActiv Then
    Zakaz.cmZapros.Enabled = True
    workChange = True
End If
End Sub

Private Sub tbWorktime_KeyDown(KeyCode As Integer, Shift As Integer)
Dim S As Double, I As Integer

If KeyCode = vbKeyReturn Then

  If isNumericTbox(tbWorktime, 0, 2000) Then
     If cbStatus.Text = "� ������" Then
        S = Round(CDbl(tbWorktime.Text), 1)
        tbWorktime.Text = S
        I = Int((CDbl(S) - 0.05) / 3)
        getWorkDay 3 + I ' ���� tmpDate
        tbReadyDate.Text = Format(tmpDate, "dd.mm.yy")
     Else
        tbReadyDate.Text = "00." & Format(tmpDate, "mm.yy")
     End If
  End If
End If

End Sub

Private Sub Timer1_Timer()
tiki = tiki - 1
If tiki > 0 Then
    laMess.ForeColor = 0
    laMess.Caption = "��� ������� �� ������ <Ok>" & Chr(13) & Chr(10) & _
    "� ��� �������� ��������� ������: " & tiki
    Timer1.Interval = 1000 ' 1c
Else
    Timer1.Enabled = False
    laMess.Caption = ""
    cmAdd.Enabled = False
    unLockBase
End If
End Sub



Public Function startParams() As Boolean
Dim I As Integer, str As String, J As Integer ', sumSroch As Double
Dim item As ListItem, v As Variant, S As Double
startParams = False

maxDay = 0

If idWerk > 0 Then ' ����� � ������ ����� �������
    Me.cmAdd.Visible = False
    Me.cmRepit.Visible = False
    gNzak = ""
    statusIdOld = 0
    Me.urgent = ""
Else
    
    Me.laNomZak.Caption = gNzak
    Me.cmAdd.Visible = True
    Me.cmRepit.Visible = True
    
    sql = "SELECT o.numorder, o.StatusId, o.DateRS, o.outTime, o.werkId" _
    & ", oe.outDateTime, oe.statusEquipId, oe.equipId, oe.worktime, oe.workTimeMO" _
    & ", oc.DateTimeMO, oc.StatM, oc.StatO" _
    & ", oc.stat as statusInCeh, oe.nevip, oc.urgent" _
    & ", o.lastModified, o.lastManagId, 0 as presentationFormat" _
    & " from Orders o" _
    & " JOIN OrdersEquip oe on oe.numorder = o.numorder" _
    & " LEFT JOIN OrdersInCeh oc on oc.numorder = o.numorder" _
    & " WHERE o.numOrder =" & gNzak & " AND oe.equipId = " & CStr(idEquip)
    Set tbOrders = myOpenRecordSet("##402", sql, dbOpenForwardOnly)
    
    Set zakazBean = New ZakazVO
    zakazBean.initFromDb
    
    tbOrders.Close
    
    If Not zakazBean.inited Then
        Exit Function
    End If
    Me.urgent = zakazBean.urgent
    
    If IsDate(zakazBean.Outdatetime) Then
        I = DateDiff("d", curDate, zakazBean.Outdatetime) + 1
        addDays I '��������� ���, �.�. ���� ��� ���.������ ����� ���������
                  '������ ��� ���� ������, ���� ��� stDay � rMaxDay
    End If
    
    If statusIdNew <> -1 And statusIdOld <> statusIdNew Then
        statusIdOld = statusIdNew
    Else
        statusIdOld = zakazBean.StatusId
    End If
    
    
End If
    
    zagruzFromCeh idEquip, gNzak '              1| � delta(), Ostatki() !!! ����� ��������
    getResurs idEquip
    Me.lvAddDays  ' ��������� ������� � ����
    For I = 1 To maxDay
        Me.lv.ListItems("k" & I).SubItems(zkPrinato) = Round(getNevip(I, idEquip), 1)
        Me.lv.ListItems("k" & I).SubItems(zkResurs) = Round(nomRes(I) * KPD * Nstan, 1)
    Next I
    Me.lv.ListItems("k1").SubItems(zkResurs) = Round(nr * Nstan * KPD, 1)

   
If statusIdOld = 0 Or statusIdOld = 7 Then '������ ��� �������
    neVipolnen = 0
    neVipolnen_O = 0
    Me.Caption = "����� �� ������������ " & Equip(idEquip)
    
    'tbWorktime.Text = ""
    'tbReadyDate.Text = ""
Else
    Me.Caption = "�������������� ������ [" & Equip(idEquip) & "]"
    If Not IsNull(zakazBean.DateRS) Then
        Me.tbDateRS.Text = Format(zakazBean.DateRS, "dd.mm.yy")
    End If
    Me.tbReadyDate.Text = Format(zakazBean.Outdatetime, "dd.mm.yy")
          
    v = zakazBean.StatM
    If cbMOsetByText(Me.cbM, v) Then
        Me.tbDateMO.Text = zakazBean.DateTimeMO
    End If
    
    Me.tbWorktime.Text = zakazBean.Worktime
     
    v = zakazBean.StatO
    If cbMOsetByText(Me.cbO, v) Then
        Me.tbDateMO = zakazBean.DateTimeMO
        If Me.cbO.Text = "�����" Then
            'Me.tbVrVipO.Text = Orders.Grid.TextMatrix(Orders.mousRow, orOVrVip)
            Me.tbVrVipO.Text = zakazBean.Worktime
            Me.tbVrVipO.Enabled = False
            Me.tbDateMO.Enabled = False
        Else 'AS nevipO
            neVipolnen_O = zakazBean.WorktimeMO
            Me.tbVrVipO.Text = neVipolnen_O
            'Me.tbVrVipO.Text = zakazBean.workTimeMO
        End If
    End If
End If

I = getNextDay(1)
v = Me.lv.ListItems("k1").SubItems(zkMost)
If Not IsNumeric(v) Then v = 0
Me.laZapas.Caption = Round(nomRes(I) * KPD * Nstan + v, 1)

'��������� ���� �� ���� ������
For I = 1 To maxDay
    delta(I) = 0
Next I
str = "DateDiff(day, now(), oe.outDateTime)"
sql = "SELECT " & str & " AS day, o.FirmId" _
& " From Orders o" _
& " join OrdersEquip oe on oe.numorder = o.numorder and oe.equipId = " & idEquip _
& " Where o.StatusId < 4" _
& " GROUP BY " & str & ", o.FirmId" _
& " HAVING " & str & " >= 0"

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
    Me.lv.ListItems("k" & I).SubItems(zkFirmKolvo) = Round(delta(I), 1)
Next I

cbBuildStatuses Me.cbStatus, statusIdOld

For I = 0 To Me.cbStatus.ListCount
    If statId(I) = statusIdOld Then
        Me.cbStatus.ListIndex = I
        GoTo NN
    End If
Next I

MsgBox "Err in Zakaz\startParams"
NN:

Me.cmZapros.Enabled = IsNumeric(tbWorktime.Text) And IsDate(tbReadyDate)

Me.lv.ListItems("k" & stDay).ForeColor = &HBB00&
Me.lv.ListItems("k" & stDay).Bold = True

Me.newZagruz Me.Regim  '������ ������ ���� ���

startParams = True
End Function

