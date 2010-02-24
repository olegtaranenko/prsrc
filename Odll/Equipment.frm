VERSION 5.00
Begin VB.Form Equipment 
   Caption         =   "Оборудование заказа"
   ClientHeight    =   2724
   ClientLeft      =   48
   ClientTop       =   588
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2724
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tbVrVipO 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   30
      Top             =   1920
      Visible         =   0   'False
      Width           =   552
   End
   Begin VB.TextBox tbVrVipO 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   29
      Top             =   1560
      Visible         =   0   'False
      Width           =   552
   End
   Begin VB.TextBox tbVrVipO 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   552
   End
   Begin VB.TextBox tbDateRS 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   26
      Top             =   1800
      Width           =   915
   End
   Begin VB.TextBox tbDateMO 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   24
      Top             =   1200
      Width           =   915
   End
   Begin VB.ComboBox cbM 
      Height          =   288
      ItemData        =   "Equipment.frx":0000
      Left            =   3120
      List            =   "Equipment.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   480
      Width           =   1035
   End
   Begin VB.ComboBox cbO 
      Height          =   288
      ItemData        =   "Equipment.frx":0024
      Left            =   4380
      List            =   "Equipment.frx":0031
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   480
      Width           =   1035
   End
   Begin VB.ComboBox cbStatus 
      Height          =   288
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   480
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "По умолчанию"
      Height          =   315
      Left            =   2520
      TabIndex        =   18
      Top             =   2280
      Width           =   1332
   End
   Begin VB.CommandButton cmSetOutDate 
      Caption         =   "..."
      Height          =   252
      Index           =   2
      Left            =   4200
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CommandButton cmSetOutDate 
      Caption         =   "..."
      Height          =   252
      Index           =   1
      Left            =   4200
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CommandButton cmSetOutDate 
      Caption         =   "..."
      Height          =   252
      Index           =   0
      Left            =   4200
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox tbWorktime 
      Height          =   288
      Index           =   2
      Left            =   1080
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox tbWorktime 
      Height          =   288
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox tbWorktime 
      Height          =   288
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CheckBox cbEquipment 
      Caption         =   " SUB"
      Height          =   372
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   732
   End
   Begin VB.CheckBox cbEquipment 
      Caption         =   " CO2"
      Height          =   372
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   732
   End
   Begin VB.CheckBox cbEquipment 
      Caption         =   " YAG"
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   732
   End
   Begin VB.CommandButton cmApply 
      Caption         =   "Применить"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmExit 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   315
      Left            =   5640
      TabIndex        =   0
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label laVrVipO 
      Caption         =   "MO"
      Height          =   252
      Left            =   1920
      TabIndex        =   31
      Top             =   840
      Width           =   252
   End
   Begin VB.Label laDateRS 
      Caption         =   "Дата Р\С (не позже)"
      Enabled         =   0   'False
      Height          =   192
      Left            =   4800
      TabIndex        =   27
      Top             =   1500
      Width           =   1692
   End
   Begin VB.Label laDateMO 
      Caption         =   "Дата Мак.\Обр."
      Enabled         =   0   'False
      Height          =   252
      Left            =   4800
      TabIndex        =   25
      Top             =   840
      Width           =   1272
   End
   Begin VB.Label laMO 
      Caption         =   "Макет                    Образец"
      Height          =   252
      Left            =   3180
      TabIndex        =   23
      Top             =   120
      Width           =   2112
   End
   Begin VB.Label Label4 
      Caption         =   "Статус"
      Height          =   252
      Left            =   1560
      TabIndex        =   19
      Top             =   120
      Width           =   852
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
      TabIndex        =   17
      Top             =   480
      Width           =   1212
   End
   Begin VB.Label Label3 
      Caption         =   "Номер заказа"
      Height          =   252
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "Дата выдачи"
      Height          =   252
      Left            =   2520
      TabIndex        =   12
      Top             =   840
      Width           =   1572
   End
   Begin VB.Label lbDateOut 
      Caption         =   "Н/А"
      Height          =   252
      Index           =   2
      Left            =   2400
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Label lbDateOut 
      Caption         =   "Н/А"
      Height          =   252
      Index           =   1
      Left            =   2400
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Label lbDateOut 
      Caption         =   "Н/А"
      Height          =   252
      Index           =   0
      Left            =   2400
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Label Label1 
      Caption         =   "Вр. вып."
      Height          =   252
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   732
   End
End
Attribute VB_Name = "Equipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public orderStatusStr As String
Dim Err As String ' чтобы не прыгал регистр
Dim currStatusId As Integer



Private Function setVisibleByEquipment(Index As Integer) As Boolean
    Dim visibleFlag As Boolean
    visibleFlag = cbEquipment(Index).value = 1
    tbWorktime(Index).Visible = visibleFlag
    cmSetOutDate(Index).Visible = visibleFlag
    lbDateOut(Index).Visible = visibleFlag
    tbVrVipO(Index).Visible = False
    
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
    For I = 0 To cbEquipment.Ubound
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
    For I = 0 To cbEquipment.Ubound
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

Private Sub Form_Load()
    lbNumorder.Caption = gNzak
    'lbStatus.Caption = orderStatusStr
    
    loadEquipment
    
End Sub

Private Sub loadEquipment()
    sql = "select o.statusId, oe.* from OrdersEquip oe join orders o on o.numorder = oe.numorder where oe.numorder = " & gNzak
    Set tbOrders = myOpenRecordSet("##eq.1", sql, dbOpenForwardOnly)
    If Not tbOrders Is Nothing Then
        If Not tbOrders.BOF Then
            currStatusId = tbOrders!statusId
            '
            While Not tbOrders.EOF
                If Not tbOrders("cehId") Is Nothing Then
                    Dim cehId As Integer
                    cehId = tbOrders!cehId - 1
                    cbEquipment(cehId).value = 1
                    If Not IsNull(tbOrders!Worktime) Then
                        tbWorktime(cehId).Text = tbOrders!Worktime
                    End If
                    
                    If Not IsNull(tbOrders!outDateTime) Then
                        lbDateOut(cehId).Caption = tbOrders!outDateTime
                    Else
                        lbDateOut(cehId).Caption = "N/A"
                    End If
                End If
                tbOrders.MoveNext
            Wend
        End If
        tbOrders.Close
    End If
    
    
End Sub

