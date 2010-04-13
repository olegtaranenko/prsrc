VERSION 5.00
Begin VB.Form Equipment 
   Caption         =   "Оборудование заказа"
   ClientHeight    =   4668
   ClientLeft      =   48
   ClientTop       =   588
   ClientWidth     =   6336
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4668
   ScaleWidth      =   6336
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Заказ"
      Height          =   2052
      Left            =   240
      TabIndex        =   23
      Top             =   0
      Width           =   5892
      Begin VB.ComboBox cbStatus 
         Height          =   288
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   720
         Width           =   1452
      End
      Begin VB.ComboBox cbO 
         Height          =   288
         ItemData        =   "Equipment.frx":0000
         Left            =   4380
         List            =   "Equipment.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   720
         Width           =   1035
      End
      Begin VB.ComboBox cbM 
         Height          =   288
         ItemData        =   "Equipment.frx":0026
         Left            =   3120
         List            =   "Equipment.frx":0033
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox tbDateMO 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   25
         Top             =   1440
         Width           =   1152
      End
      Begin VB.TextBox tbDateRS 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   24
         Top             =   1440
         Width           =   1152
      End
      Begin VB.Label lbZakazDateOut 
         Caption         =   "Н/А"
         Height          =   252
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   1692
      End
      Begin VB.Label Label6 
         Caption         =   "Дата выд."
         Height          =   252
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "Номер заказа"
         Height          =   252
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label lbNumorder 
         Caption         =   "Номер заказа"
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
         TabIndex        =   33
         Top             =   720
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "Статус"
         Height          =   252
         Left            =   1560
         TabIndex        =   32
         Top             =   360
         Width           =   852
      End
      Begin VB.Label laMO 
         Caption         =   "Макет                    Образец"
         Height          =   252
         Left            =   3180
         TabIndex        =   31
         Top             =   360
         Width           =   2112
      End
      Begin VB.Label laDateMO 
         Caption         =   "Дата Мак.\Обр."
         Enabled         =   0   'False
         Height          =   252
         Left            =   3960
         TabIndex        =   30
         Top             =   1128
         Width           =   1272
      End
      Begin VB.Label laDateRS 
         Caption         =   "Дата Р\С (не позже)"
         Enabled         =   0   'False
         Height          =   192
         Left            =   2040
         TabIndex        =   29
         Top             =   1128
         Width           =   1692
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "По умолчанию"
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   4200
      Width           =   1332
   End
   Begin VB.CommandButton cmApply 
      Caption         =   "Применить"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmExit 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   315
      Left            =   4560
      TabIndex        =   0
      Top             =   4200
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "По оборудованию"
      Height          =   1932
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   5892
      Begin VB.CheckBox cbUrgent 
         Enabled         =   0   'False
         Height          =   252
         Index           =   2
         Left            =   5280
         TabIndex        =   40
         Top             =   1440
         Width           =   252
      End
      Begin VB.CheckBox cbUrgent 
         Enabled         =   0   'False
         Height          =   252
         Index           =   1
         Left            =   5280
         TabIndex        =   39
         Top             =   1080
         Width           =   252
      End
      Begin VB.CheckBox cbUrgent 
         Enabled         =   0   'False
         Height          =   252
         Index           =   0
         Left            =   5280
         TabIndex        =   38
         Top             =   720
         Width           =   252
      End
      Begin VB.CheckBox cbEquipment 
         Caption         =   " YAG"
         Height          =   372
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   732
      End
      Begin VB.CheckBox cbEquipment 
         Caption         =   " CO2"
         Height          =   372
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   732
      End
      Begin VB.CheckBox cbEquipment 
         Caption         =   " SUB"
         Height          =   372
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   732
      End
      Begin VB.TextBox tbWorktime 
         Height          =   288
         Index           =   0
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   492
      End
      Begin VB.TextBox tbWorktime 
         Height          =   288
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   492
      End
      Begin VB.TextBox tbWorktime 
         Height          =   288
         Index           =   2
         Left            =   1200
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   492
      End
      Begin VB.TextBox tbVrVipO 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   552
      End
      Begin VB.TextBox tbVrVipO 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   552
      End
      Begin VB.TextBox tbVrVipO 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   552
      End
      Begin VB.Label Label7 
         Caption         =   "Сроч-ть"
         Height          =   252
         Left            =   5160
         TabIndex        =   37
         Top             =   360
         Width           =   612
      End
      Begin VB.Label Label1 
         Caption         =   "Вр. вып."
         Height          =   252
         Left            =   1200
         TabIndex        =   22
         Top             =   360
         Width           =   732
      End
      Begin VB.Label lbDateOut 
         Caption         =   "Н/А"
         Height          =   252
         Index           =   0
         Left            =   2520
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label lbDateOut 
         Caption         =   "Н/А"
         Height          =   252
         Index           =   1
         Left            =   2520
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label lbDateOut 
         Caption         =   "Н/А"
         Height          =   252
         Index           =   2
         Left            =   2520
         TabIndex        =   19
         Top             =   1440
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label2 
         Caption         =   "Дата Изгот."
         Height          =   252
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label laVrVipO 
         Caption         =   "MO"
         Height          =   252
         Left            =   2040
         TabIndex        =   17
         Top             =   360
         Width           =   252
      End
      Begin VB.Label Label5 
         Caption         =   "Статус"
         Height          =   252
         Left            =   4200
         TabIndex        =   16
         Top             =   360
         Width           =   972
      End
      Begin VB.Label lbEquipStatus 
         Height          =   252
         Index           =   2
         Left            =   4200
         TabIndex        =   15
         Top             =   1440
         Width           =   996
      End
      Begin VB.Label lbEquipStatus 
         Height          =   252
         Index           =   1
         Left            =   4200
         TabIndex        =   14
         Top             =   1080
         Width           =   996
      End
      Begin VB.Label lbEquipStatus 
         Height          =   252
         Index           =   0
         Left            =   4200
         TabIndex        =   13
         Top             =   720
         Width           =   996
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
Dim Err As String ' чтобы не прыгал регистр
Dim currStatusId As Integer




Private Function setVisibleByEquipment(Index As Integer) As Boolean
    Dim visibleFlag As Boolean
    visibleFlag = cbEquipment(Index).value = 1
    tbWorktime(Index).Visible = visibleFlag
    'cmSetOutDate(Index).Visible = visibleFlag
    lbDateOut(Index).Visible = visibleFlag
    tbVrVipO(Index).Visible = False
    lbEquipStatus(Index).Visible = visibleFlag
    
    If visibleFlag Then
        If currStatusId = 3 Then '"согласов"
            tbVrVipO(Index).Visible = True
        Else
            tbVrVipO(Index).Visible = False
        End If
    
    End If
    setVisibleByEquipment = visibleFlag
    
    
End Function


Private Sub cbEquipment_Click(Index As Integer)
    
    setVisibleByEquipment Index
    
    If tbWorktime(Index).Visible Then
        tbWorktime(Index).SetFocus
    End If

End Sub


Private Sub cbM_Change()
Dim I As Integer
End Sub

Private Sub cbStatus_Click()
Dim I As Integer

    
    currStatusId = statId(cbStatus.ListIndex)

    laMO.Enabled = False
    cbM.Enabled = False
    cbO.Enabled = False
    tbDateMO.Enabled = False
    tbDateRS.Enabled = False
    laDateRS.Enabled = False
    If currStatusId = 1 Then ' "в работе"
    
    ElseIf currStatusId = 2 Then ' "резерв"
        laDateRS.Enabled = True
        tbDateRS.Enabled = True
    ElseIf currStatusId = 3 Then ' "согласов"
        cbM.Enabled = True
        cbO.Enabled = True
        laMO.Enabled = True
        laDateRS.Enabled = True
        tbDateRS.Enabled = True
    Else
        laMO.Enabled = False
        cbM.Enabled = False
        cbO.Enabled = False
        tbDateMO.Enabled = False
    End If
    
    Dim hasVisible As Boolean, isVisible As Boolean
    For I = 0 To cbEquipment.UBound
        isVisible = setVisibleByEquipment(I)
        If Not hasVisible Then
            hasVisible = isVisible
        End If
    Next I
    
    'If hasVisible Then
    '    laVrVipO.Visible = True
    'Else
    '    laVrVipO.Visible = False
    'End If
    
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
    Orders.refreshCurrentRow = True
    Unload Me
End Sub


Private Sub cmExit_Click()
    Unload Me
End Sub


Private Sub putOrderEquip(Index As Integer)
    Dim Worktime As Single
    Dim DateOut As String
    Dim cehId As Integer
    cehId = Index + 1
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
    sql = "call putOrderEquip (" & gNzak & "," & cehId & "," & Worktime & "," & DateOut & ")"
    myExecute "W#eq.2", sql
    
End Sub

Private Sub deleteOrderEquip(Index As Integer)
    Dim cehId As Integer
    cehId = Index + 1
    
    sql = "call deleteOrderEquip (" & gNzak & "," & cehId & ")"
    myExecute "W#eq.3", sql, -1
    
End Sub

Private Sub cmSetOutDate_Click(Index As Integer)
    Dim cehId As Integer
    cehId = Index + 1
    
End Sub

Private Sub setReadonly()
    Dim I As Integer
    For I = 0 To cbEquipment.UBound
        cbEquipment(I).Enabled = Not Me.readonlyFlag
        tbWorktime(I).Enabled = Not Me.readonlyFlag
        tbVrVipO(I).Enabled = Not Me.readonlyFlag
        
    Next I
    cmApply.Enabled = Not readonlyFlag
    
End Sub

Private Sub Form_Load()
    lbNumorder.Caption = gNzak
    'lbStatus.Caption = orderStatusStr
    setReadonly
    
    loadEquipment
    
End Sub

Private Sub loadEquipment()
    If gNzak = "" Then Exit Sub
    
    Dim Outdate As Date
    sql = "select o.StatusId, oe.Outdatetime" _
    & " from orders o " _
    & " join vw_OrdersEquipSummary oe on o.numorder = oe.numorder " _
    & " where o.numorder = " & gNzak
    
    byErrSqlGetValues "W#eq01", sql, currStatusId, Outdate
    lbZakazDateOut.Caption = Format(Outdate, "dd.mm.yyyy")
    
    
    sql = "select oe.* " _
    & " , s.status " _
    & " , isnull(oc.urgent, '') as urgent " _
    & " from OrdersEquip oe " _
    & " join guideStatus s on s.statusId = oe.statusEquipId " _
    & " left join OrdersInCeh oc on oe.numorder = oc.numorder and oe.cehId = oc.cehId " _
    & " where oe.numorder = " & gNzak

    Set tbOrders = myOpenRecordSet("##eq02", sql, dbOpenForwardOnly)
    If Not tbOrders Is Nothing Then
        If Not tbOrders.BOF Then
            '
            While Not tbOrders.EOF
                If Not tbOrders("cehId") Is Nothing Then
                    Dim cehId As Integer
                    cehId = tbOrders!cehId - 1
                    cbEquipment(cehId).value = 1
                    If tbOrders!urgent <> "" Then
                        cbUrgent(cehId).value = 1
                    End If
                    
                    If Not IsNull(tbOrders!Worktime) Then
                        tbWorktime(cehId).Text = tbOrders!Worktime
                    End If
                    
                    If Not IsNull(tbOrders!Outdatetime) Then
                        lbDateOut(cehId).Caption = tbOrders!Outdatetime
                    Else
                        lbDateOut(cehId).Caption = "N/A"
                    End If
                    
                    If Not IsNull(tbOrders!status) Then
                        lbEquipStatus(cehId).Caption = tbOrders!status
                    Else
                        lbDateOut(cehId).Caption = " "
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

Private Sub CleanupEquip()
Dim I As Integer
    For I = 0 To lenCeh - 1
        cbEquipment(I).value = 0
        tbWorktime(I).Visible = False
        tbVrVipO(I).Visible = False
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

