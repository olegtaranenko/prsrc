VERSION 5.00
Begin VB.Form Equipment 
   Caption         =   "Оборудование заказа"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   588
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2496
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "По умолчанию"
      Height          =   315
      Left            =   2520
      TabIndex        =   18
      Top             =   2040
      Width           =   1332
   End
   Begin VB.CommandButton cmSetOutDate 
      Caption         =   "..."
      Height          =   252
      Index           =   2
      Left            =   3720
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CommandButton cmSetOutDate 
      Caption         =   "..."
      Height          =   252
      Index           =   1
      Left            =   3720
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CommandButton cmSetOutDate 
      Caption         =   "..."
      Height          =   252
      Index           =   0
      Left            =   3720
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox tbWorktime 
      Height          =   288
      Index           =   2
      Left            =   1200
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox tbWorktime 
      Height          =   288
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox tbWorktime 
      Height          =   288
      Index           =   0
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CheckBox cbEquipment 
      Caption         =   " SUB"
      Height          =   372
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   972
   End
   Begin VB.CheckBox cbEquipment 
      Caption         =   " CO2"
      Height          =   372
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   972
   End
   Begin VB.CheckBox cbEquipment 
      Caption         =   " YAG"
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton cmApply 
      Caption         =   "Применить"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmExit 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   315
      Left            =   4980
      TabIndex        =   0
      Top             =   2040
      Width           =   795
   End
   Begin VB.Label lbStatus 
      Caption         =   "Статус заказа"
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
      Left            =   3600
      TabIndex        =   20
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label4 
      Caption         =   "Статус"
      Height          =   252
      Left            =   2640
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
      Left            =   1440
      TabIndex        =   17
      Top             =   120
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
      Top             =   600
      Width           =   1572
   End
   Begin VB.Label lbDateOut 
      Caption         =   "Н/А"
      Height          =   252
      Index           =   2
      Left            =   2160
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label lbDateOut 
      Caption         =   "Н/А"
      Height          =   252
      Index           =   1
      Left            =   2160
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label lbDateOut 
      Caption         =   "Н/А"
      Height          =   252
      Index           =   0
      Left            =   2160
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "Время вып."
      Height          =   252
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Width           =   972
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


Private Sub cbEquipment_Click(Index As Integer)
    Dim visibleFlag As Boolean
    visibleFlag = cbEquipment(Index)
    tbWorktime(Index).Visible = visibleFlag
    cmSetOutDate(Index).Visible = visibleFlag
    lbDateOut(Index).Visible = visibleFlag
    If tbWorktime(Index).Visible Then
        tbWorktime(Index).SetFocus
    End If
    If tbWorktime(Index).Locked Then
        Dim I As Integer
        
    End If
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
    Dim worktime As Single
    Dim DateOut As String
    Dim cehId As Integer
    cehId = Index + 1
    If IsNumeric(tbWorktime(Index).Text) Then
        worktime = tbWorktime(Index).Text
    Else
        worktime = 0
    End If
    
    If IsDate(lbDateOut(Index).Caption) Then
        DateOut = Format(CDate(lbDateOut(Index).Caption), "'yyyymmdd hh:nn'")
    Else
        DateOut = "null"
    End If
    sql = "call putOrderEquip (" & gNzak & "," & cehId & "," & worktime & "," & DateOut & ")"
    myExecute "W#eq.2", sql
    
End Sub

Private Sub deleteOrderEquip(Index As Integer)
    Dim cehId As Integer
    cehId = Index + 1
    
    sql = "call deleteOrderEquip (" & gNzak & "," & cehId & ")"
    myExecute "W#eq.3", sql, -1
    
End Sub

Private Sub Form_Load()
    lbNumorder.Caption = gNzak
    lbStatus.Caption = orderStatusStr
    
    loadEquipment
    
End Sub

Private Sub loadEquipment()
    sql = "select * from OrdersEquip where numorder = " & gNzak
    Set tbOrders = myOpenRecordSet("##eq.1", sql, dbOpenForwardOnly)
    If Not tbOrders Is Nothing Then
        If Not tbOrders.BOF Then
            While Not tbOrders.EOF
                If Not tbOrders("cehId") Is Nothing Then
                    Dim cehId As Integer
                    cehId = tbOrders!cehId - 1
                    cbEquipment(cehId).value = 1
                    If Not IsNull(tbOrders!worktime) Then
                        tbWorktime(cehId).Text = tbOrders!worktime
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

