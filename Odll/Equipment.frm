VERSION 5.00
Begin VB.Form Equipment 
   Caption         =   "������������ ������"
   ClientHeight    =   4524
   ClientLeft      =   48
   ClientTop       =   588
   ClientWidth     =   9156
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4524
   ScaleWidth      =   9156
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame OrderFrame 
      Caption         =   "�����"
      Height          =   1932
      Left            =   240
      TabIndex        =   12
      Top             =   0
      Width           =   8772
      Begin VB.CheckBox cbUrgent 
         Enabled         =   0   'False
         Height          =   252
         Left            =   5160
         TabIndex        =   27
         Top             =   720
         Width           =   252
      End
      Begin VB.ComboBox cbStatus 
         Enabled         =   0   'False
         Height          =   288
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   720
         Width           =   1452
      End
      Begin VB.ComboBox cbO 
         Enabled         =   0   'False
         Height          =   288
         ItemData        =   "Equipment.frx":0000
         Left            =   1380
         List            =   "Equipment.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1440
         Width           =   1035
      End
      Begin VB.ComboBox cbM 
         Enabled         =   0   'False
         Height          =   288
         ItemData        =   "Equipment.frx":0032
         Left            =   120
         List            =   "Equipment.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1440
         Width           =   1035
      End
      Begin VB.TextBox tbDateMO 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   14
         Top             =   1440
         Width           =   1152
      End
      Begin VB.TextBox tbDateRS 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   13
         Top             =   1440
         Width           =   1152
      End
      Begin VB.Label Label7 
         Caption         =   "������"
         Height          =   252
         Left            =   5160
         TabIndex        =   26
         Top             =   360
         Width           =   612
      End
      Begin VB.Label lbZakazDateOut 
         Caption         =   "�/�"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1440
         TabIndex        =   25
         Top             =   720
         Width           =   1812
      End
      Begin VB.Label Label6 
         Caption         =   "���� ���."
         Height          =   252
         Left            =   1440
         TabIndex        =   24
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "����� ������"
         Height          =   252
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label lbNumorder 
         Caption         =   "����� ������"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "������"
         Height          =   252
         Left            =   3360
         TabIndex        =   21
         Top             =   360
         Width           =   852
      End
      Begin VB.Label laMO 
         Caption         =   "�����                    �������"
         Height          =   252
         Left            =   180
         TabIndex        =   20
         Top             =   1080
         Width           =   2112
      End
      Begin VB.Label laDateMO 
         Caption         =   "���� ���.\���."
         Enabled         =   0   'False
         Height          =   252
         Left            =   4680
         TabIndex        =   19
         Top             =   1128
         Width           =   1272
      End
      Begin VB.Label laDateRS 
         Caption         =   "���� �\� (�� �����)"
         Enabled         =   0   'False
         Height          =   192
         Left            =   2760
         TabIndex        =   18
         Top             =   1128
         Width           =   1692
      End
   End
   Begin VB.CommandButton cmApply 
      Caption         =   "���������"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmExit 
      Cancel          =   -1  'True
      Caption         =   "������"
      Height          =   315
      Left            =   7800
      TabIndex        =   0
      Top             =   4080
      Width           =   1152
   End
   Begin VB.Frame EquipFrame 
      Caption         =   "�� ������������"
      Height          =   1932
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   8772
      Begin VB.CheckBox cbEquipment 
         Caption         =   " YAG"
         Height          =   372
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   1692
      End
      Begin VB.TextBox tbWorktime 
         Height          =   288
         Index           =   0
         Left            =   2040
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   528
      End
      Begin VB.TextBox tbWorktimeO 
         Height          =   285
         Index           =   0
         Left            =   2640
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   552
      End
      Begin VB.Label laVrVipO 
         Caption         =   "���-��"
         Height          =   252
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   612
      End
      Begin VB.Label Label11 
         Caption         =   "�����."
         Height          =   252
         Left            =   2040
         TabIndex        =   34
         Top             =   360
         Width           =   612
      End
      Begin VB.Label Label10 
         Caption         =   "� ����"
         Height          =   252
         Left            =   6120
         TabIndex        =   33
         Top             =   360
         Width           =   612
      End
      Begin VB.Label lbEquipStat 
         Height          =   252
         Index           =   0
         Left            =   6000
         TabIndex        =   32
         Top             =   720
         Width           =   996
      End
      Begin VB.Label lbNevip 
         Height          =   252
         Index           =   0
         Left            =   8160
         TabIndex        =   31
         Top             =   720
         Width           =   516
      End
      Begin VB.Label Label9 
         Caption         =   "%���."
         Height          =   252
         Left            =   8160
         TabIndex        =   30
         Top             =   360
         Width           =   492
      End
      Begin VB.Label lbEquipStatusO 
         Height          =   252
         Index           =   0
         Left            =   7080
         TabIndex        =   29
         Top             =   720
         Width           =   996
      End
      Begin VB.Label Label8 
         Caption         =   "���-��"
         Height          =   252
         Left            =   7200
         TabIndex        =   28
         Top             =   360
         Width           =   612
      End
      Begin VB.Label Label1 
         Caption         =   "����� ������������"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1812
      End
      Begin VB.Label lbDateOut 
         Height          =   252
         Index           =   0
         Left            =   3240
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label2 
         Caption         =   "� ����"
         Height          =   252
         Left            =   3360
         TabIndex        =   9
         Top             =   360
         Width           =   732
      End
      Begin VB.Label Label5 
         Caption         =   "������"
         Height          =   252
         Left            =   5040
         TabIndex        =   7
         Top             =   360
         Width           =   612
      End
      Begin VB.Label lbEquipStatus 
         Height          =   252
         Index           =   0
         Left            =   5040
         TabIndex        =   6
         Top             =   720
         Width           =   876
      End
   End
End
Attribute VB_Name = "Equipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public orderStatusStr As String
Public readonlyFlag As Boolean
Public originalStatusId As Integer

Dim Err As String ' ����� �� ������ �������
Dim currStatusId As Integer, urgent As String
Private idWerk 'As Integer




Private Function setVisibleByEquipment(Index As Integer, visibleFlag As Boolean) As Boolean
    tbWorktime(Index).Visible = visibleFlag
    'cmSetOutDate(Index).Visible = visibleFlag
    lbDateOut(Index).Visible = visibleFlag
    tbWorktimeO(Index).Visible = visibleFlag
    lbEquipStatus(Index).Visible = visibleFlag
    lbEquipStat(Index).Visible = visibleFlag
    lbEquipStatusO(Index).Visible = visibleFlag
    lbNevip(Index).Visible = visibleFlag
    setVisibleByEquipment = visibleFlag
    
    
End Function


Private Sub cbEquipment_Click(Index As Integer)
    
    setVisibleByEquipment Index, cbEquipment(Index).Visible
    
    If tbWorktime(Index).Visible Then
        tbWorktime(Index).SetFocus
    End If

End Sub


Private Sub cbM_Change()
Dim I As Integer
End Sub


Private Sub cmApply_Click()
    Dim I As Integer
    For I = 0 To cbEquipment.UBound
        If cbEquipment(I) Then
            ' insert & update the row
            putOrderEquip (I)
        Else
            ' delete the row
            deleteOrderEquip (I)
        End If
    Next I
    
    If currStatusId <> originalStatusId Then
        sql = "update orders set statusId = " & currStatusId & " where numorder = " & gNzak
        myExecute "##eq05", sql
        'sql = "update ordersEquip set statusEquipId = " & currStatusId & " where numorder = " & gNzak
        'myExecute "##eq06", sql
    End If
    
    Orders.refreshCurrentRow = True
    
    Unload Me
End Sub


Private Sub cmExit_Click()
    Unload Me
End Sub


Private Sub putOrderEquip(Index As Integer)
    Dim Worktime As Double, WorktimeMO As Double
    Dim DateOut As String
    Dim EquipId As Integer
    EquipId = Index + 1
    If IsNumeric(tbWorktime(Index).Text) Then
        Worktime = tbWorktime(Index).Text
    Else
        Worktime = 0
    End If
    
    If IsDate(lbDateOut(Index).Caption) Then
        DateOut = Format(CDate(lbDateOut(Index).Caption), "'yyyymmdd hh:nn'")
    Else
        DateOut = "null"
    End If
    
    If IsNumeric(tbWorktimeO(Index).Text) Then
        WorktimeMO = tbWorktimeO(Index).Text
    Else
        WorktimeMO = 0
    End If
    
    
    sql = "call putOrderEquip (" & gNzak & "," & EquipId & "," & Worktime & "," & DateOut & "," & WorktimeMO & ")"
    'Debug.Print sql
    myExecute "W#eq.2", sql
    
End Sub

Private Sub deleteOrderEquip(Index As Integer)
    Dim EquipId As Integer
    EquipId = Index + 1
    
    sql = "call deleteOrderEquip (" & gNzak & "," & EquipId & ")"
    myExecute "W#eq.3", sql, -1
    
End Sub


Private Sub cmSetOutDate_Click(Index As Integer)
    Dim EquipId As Integer
    EquipId = Index + 1
    
End Sub


Private Sub setReadonly()
    Dim I As Integer
    For I = 0 To cbEquipment.UBound
        cbEquipment(I).Enabled = Not Me.readonlyFlag
        tbWorktime(I).Enabled = Not Me.readonlyFlag
        tbWorktimeO(I).Enabled = Not Me.readonlyFlag
        
    Next I
    cmApply.Enabled = Not readonlyFlag
    
End Sub


Private Sub Form_Load()
    lbNumorder.Caption = gNzak
    'lbStatus.Caption = orderStatusStr
    setReadonly
    
    loadEnv
    
    loadEquipment
    
    tuneEnv
End Sub


Private Sub loadEnv()

Dim I As Integer, VShift As Integer, LowLinie As Long

    For I = 1 To UBound(Equip) - 1
    
        VShift = LowLinie + 15
        Load cbEquipment(I)
        Load tbWorktime(I)
        Load lbDateOut(I)
        Load lbEquipStatus(I)
        Load lbEquipStat(I)
        Load lbEquipStatusO(I)
        Load lbNevip(I)
        Load tbWorktimeO(I)
        
        
        cbEquipment(I).Caption = EquipFullName(I + 1)
        cbEquipment(I).Visible = True
        LowLinie = tbWorktime(I).Top + tbWorktime(I).Height
    Next I
    
    cbBuildStatuses Me.cbStatus, Me.originalStatusId
    cbStatus.Text = Status(Me.originalStatusId)
    
End Sub

Private Sub tuneEnv()
Dim I As Integer, LowLinie As Long, VShift As Long
Dim equipIndex As Integer
    
    If idWerk > 0 Then
        VShift = tbWorktime(0).Height + 15
        LowLinie = tbWorktime(0).Top
        hideEquipAll
        sql = "select * " _
            & " from WerkEquip we " _
            & " where we.werkId = " & idWerk _
            & " order by we.werkId"
    
        Set tbOrders = myOpenRecordSet("##eq04", sql, dbOpenForwardOnly)
        If Not tbOrders Is Nothing Then
            While Not tbOrders.EOF
                equipIndex = tbOrders!EquipId - 1
                cbEquipment(equipIndex).Visible = True
                setVisibleByEquipment equipIndex, True
                AlignEquipmentControls equipIndex, LowLinie
                tbOrders.MoveNext
                LowLinie = LowLinie + VShift
            Wend
            tbOrders.Close
        End If
    End If
    
    EquipFrame.Height = LowLinie + 150
    LowLinie = EquipFrame.Top + EquipFrame.Height + 100
    cmExit.Top = LowLinie
    cmApply.Top = LowLinie
    LowLinie = cmApply.Top + cmApply.Height + 100
    Me.Height = LowLinie + 700
    
End Sub

Private Sub AlignEquipmentControls(equipCtlIndex As Integer, VShift As Long)
    cbEquipment(equipCtlIndex).Top = VShift
    tbWorktime(equipCtlIndex).Top = VShift
    lbDateOut(equipCtlIndex).Top = VShift
    lbEquipStatus(equipCtlIndex).Top = VShift
    lbEquipStat(equipCtlIndex).Top = VShift
    lbEquipStatusO(equipCtlIndex).Top = VShift
    lbNevip(equipCtlIndex).Top = VShift
    tbWorktimeO(equipCtlIndex).Top = VShift
End Sub


Private Sub loadEquipment()
    If gNzak = "" Then Exit Sub
    
    Dim Outdate As Variant, Outtime, StatO, StatM, Stat, DateTimeMO, DateRS, str
    
    sql = "select o.StatusId, o.Outdatetime, o.outTime" _
    & ", oc.urgent, oc.StatM, oc.DateTimeMO" _
    & ", o.DateRS, o.werkId" _
    & " from orders o " _
    & " left join OrdersInCeh oc on oc.numorder = o.numorder " _
    & " where o.numorder = " & gNzak
    
    byErrSqlGetValues "w#eq01", sql, currStatusId, Outdate, Outtime, urgent _
        , StatM, DateTimeMO, DateRS, idWerk
    
    
    str = Format(Outdate, "dd.mm.yyyy")
    If Not IsNull(Outtime) Then
        str = str & " " & Outtime & ":00"
    End If
    lbZakazDateOut.Caption = str
    'Debug.Print sql
    
    If urgent <> "" Then
        cbUrgent.Value = 1
    Else
        cbUrgent.Value = 0
    End If
    
    If IsNull(StatO) Or StatO = "" Then
        cbO.ListIndex = 0
    Else
        cbO.Text = StatO
    End If

    cbMOsetByText cbO, StatO, 0
    
    If IsNull(StatM) Or StatM = "" Then
        cbM.ListIndex = 0
    Else
        cbM.Text = StatM
    End If
    
    cbMOsetByText cbM, StatM, 0
    
    If Not IsNull(DateRS) Then
        tbDateRS = Format(DateRS, "dd.mm.yy")
    Else
        tbDateRS = ""
    End If
    
    If Not IsNull(DateTimeMO) Then
        tbDateMO = Format(DateRS, "dd.mm.yy")
    Else
        tbDateMO = ""
    End If
    
    sql = "select oe.worktime, oe.worktimeMO, oe.Stat, oe.StatO" _
    & " , oe.EquipId, oe.Outdatetime, (1 - isnull(nevip, 1)) * 100 as nevip " _
    & " , s.status as statusEquip " _
    & " , isnull(oc.urgent, '') as urgent" _
    & " FROM OrdersEquip oe " _
    & " LEFT JOIN OrdersInCeh oc on oe.numorder = oc.numorder" _
    & " LEFT JOIN StatusGuide  s ON s.statusId = oe.statusEquipId" _
    & " WHERE oe.numorder = " & gNzak

    Set tbOrders = myOpenRecordSet("##eq02", sql, dbOpenForwardOnly)
    If Not tbOrders Is Nothing Then
        If Not tbOrders.BOF Then
            '
            While Not tbOrders.EOF
                If Not tbOrders("equipId") Is Nothing Then
                    Dim EquipId As Integer
                    EquipId = tbOrders!EquipId - 1
                    cbEquipment(EquipId).Value = 1
                    
                    If Not IsNull(tbOrders!Worktime) Then
                        tbWorktime(EquipId).Text = tbOrders!Worktime
                    End If
                    
                    If Not IsNull(tbOrders!WorktimeMO) Then
                        tbWorktimeO(EquipId).Text = tbOrders!WorktimeMO
                    Else
                        tbWorktimeO(EquipId).Text = ""
                    End If
                    
                    If Not IsNull(tbOrders!Outdatetime) Then
                        lbDateOut(EquipId).Caption = tbOrders!Outdatetime
                    Else
                        lbDateOut(EquipId).Caption = ""
                    End If
                    
                    If Not IsNull(tbOrders!statusEquip) Then
                        lbEquipStatus(EquipId).Caption = tbOrders!statusEquip
                    Else
                        lbEquipStatus(EquipId).Caption = ""
                    End If
                    If Not IsNull(tbOrders!Stat) Then
                        lbEquipStat(EquipId).Caption = tbOrders!Stat
                    Else
                        lbEquipStat(EquipId).Caption = ""
                    End If
                    
                    If Not IsNull(tbOrders!StatO) Then
                        lbEquipStatusO(EquipId).Caption = tbOrders!StatO
                    Else
                        lbEquipStatusO(EquipId).Caption = ""
                    End If
                    If Not IsNull(tbOrders!nevip) Then
                        lbNevip(EquipId).Caption = tbOrders!nevip
                    Else
                        lbNevip(EquipId).Caption = ""
                    End If
                End If
                tbOrders.MoveNext
            Wend
        End If
        tbOrders.Close
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CleanupEquip
End Sub

Private Sub hideEquipAll()
Dim I As Integer
    For I = 0 To UBound(Equip) - 1
        HideEquip I, False
    Next I
End Sub

Private Sub HideEquip(EquipId As Integer, Show As Boolean)
    cbEquipment(EquipId).Visible = Show
    tbWorktime(EquipId).Visible = Show
    tbWorktimeO(EquipId).Visible = Show
    lbDateOut(EquipId).Visible = Show
    lbEquipStatus(EquipId).Visible = Show
    lbEquipStat(EquipId).Visible = Show
    lbEquipStatusO(EquipId).Visible = Show
    lbNevip(EquipId).Visible = Show
End Sub


Private Sub CleanupEquip()
Dim I As Integer
    For I = 0 To UBound(Equip) - 1
        'cbEquipment(I).value = 0
        tbWorktime(I).Visible = False
        tbWorktimeO(I).Visible = False
        lbDateOut(I).Visible = False
    Next I
    tbDateMO.Text = ""
    tbDateRS.Text = ""
    cbM.ListIndex = 0
    cbO.ListIndex = 0
    If cbStatus.ListCount > 0 Then
        cbStatus.ListIndex = 0
    End If
End Sub

