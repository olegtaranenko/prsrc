VERSION 5.00
Begin VB.Form Filtr 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Сложный  фильтр"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   8520
      TabIndex        =   27
      Top             =   360
      Width           =   1215
      Begin VB.OptionButton opAnys 
         Caption         =   "не важно"
         Height          =   495
         Left            =   0
         TabIndex        =   30
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opNotAlls 
         Caption         =   "не полност"
         Height          =   495
         Left            =   0
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton opAlls 
         Caption         =   "полностью"
         Height          =   495
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmAdvan 
      Caption         =   "Свернуть  "
      Height          =   315
      Left            =   4380
      TabIndex        =   25
      Top             =   5700
      Width           =   1635
   End
   Begin VB.ListBox lbTema 
      Enabled         =   0   'False
      Height          =   2400
      Left            =   1800
      MultiSelect     =   2  'Расширенно
      TabIndex        =   22
      Top             =   3180
      Width           =   1575
   End
   Begin VB.ListBox lbType 
      Height          =   1035
      ItemData        =   "Filtr.frx":0000
      Left            =   1080
      List            =   "Filtr.frx":0013
      TabIndex        =   21
      Top             =   3180
      Width           =   375
   End
   Begin VB.OptionButton opAny 
      Caption         =   "не важно"
      Height          =   255
      Left            =   7200
      TabIndex        =   20
      Top             =   1200
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton opNotAll 
      Caption         =   "не полност"
      Height          =   255
      Left            =   7200
      TabIndex        =   19
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton opAll 
      Caption         =   "полностью"
      Height          =   255
      Left            =   7200
      TabIndex        =   18
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmAddFirm 
      Caption         =   "Добавить"
      Height          =   330
      Left            =   5040
      TabIndex        =   1
      Top             =   1980
      Width           =   915
   End
   Begin VB.CommandButton cmExit 
      Cancel          =   -1  'True
      Caption         =   "Выход"
      Height          =   315
      Left            =   8880
      TabIndex        =   0
      Top             =   5700
      Width           =   795
   End
   Begin VB.CommandButton cmFiltrGo 
      Caption         =   "Применить"
      Height          =   315
      Left            =   6900
      TabIndex        =   6
      Top             =   5700
      Width           =   1095
   End
   Begin VB.CommandButton cmReset 
      Caption         =   "Сброс"
      Height          =   315
      Left            =   2400
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmBeg 
      Caption         =   "Нач. фильтр"
      Height          =   315
      Left            =   240
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5700
      Width           =   1095
   End
   Begin VB.ListBox lbFirm 
      Height          =   1425
      ItemData        =   "Filtr.frx":0025
      Left            =   3780
      List            =   "Filtr.frx":0027
      MultiSelect     =   2  'Расширенно
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   420
      Width           =   3315
   End
   Begin VB.ListBox lbStatus 
      Height          =   1620
      Left            =   2400
      MultiSelect     =   2  'Расширенно
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   420
      Width           =   1155
   End
   Begin VB.CheckBox cbStartDate 
      Caption         =   "  "
      Height          =   315
      Left            =   540
      TabIndex        =   2
      Top             =   420
      Width           =   315
   End
   Begin VB.TextBox tbStartDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   420
      Width           =   795
   End
   Begin VB.CheckBox cbEndDate 
      Caption         =   "  "
      Height          =   315
      Left            =   540
      TabIndex        =   4
      Top             =   840
      Width           =   315
   End
   Begin VB.TextBox tbEndDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   795
   End
   Begin VB.ListBox lbM 
      Height          =   255
      Left            =   1740
      MultiSelect     =   2  'Расширенно
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   420
      Width           =   555
   End
   Begin VB.Label laEnable 
      BorderStyle     =   1  'Фиксировано один
      Height          =   1035
      Left            =   540
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "Отгружено"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8400
      TabIndex        =   26
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label laTema 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "Тема"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1740
      TabIndex        =   24
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label laKategor 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "Категория"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "Оплачено"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   17
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "Название Фирмы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3660
      TabIndex        =   13
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "Статус"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "М"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   60
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Центровка
      Appearance      =   0  'Плоска
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "Дата"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Width           =   1695
   End
   Begin VB.Label laS_Po 
      Caption         =   "с    по"
      Height          =   555
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   195
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Filtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbEndDate_Click()
tbEndDate.Enabled = Not tbEndDate.Enabled

End Sub

Private Sub cbStartDate_Click()
tbStartDate.Enabled = Not tbStartDate.Enabled

End Sub

Private Sub cmAddFirm_Click()
    FindFirm.Regim = "fromFiltr"
    FindFirm.cmSelect.Visible = True
    FindFirm.Show vbModal
    lbFirm.SetFocus
End Sub

Public Sub cmAdvan_Click()
Dim delta As Integer
delta = 3000

If Left$(cmAdvan.Caption, 1) = "Д" Then
    Me.Height = Me.Height + delta
    cmBeg.Top = cmBeg.Top + delta
    cmReset.Top = cmReset.Top + delta
    cmFiltrGo.Top = cmFiltrGo.Top + delta
    cmAdvan.Top = cmAdvan.Top + delta
    cmExit.Top = cmExit.Top + delta
    cmAdvan.Caption = "Свернуть <<"
    lbTema.Visible = True
    lbType.Visible = True
    laKategor.Visible = True
    laTema.Visible = True
    lbDeSelectAll lbTema
    Me.Top = Me.Top - delta / 2
Else
    Me.Height = Me.Height - delta
    Me.Top = Me.Top + delta / 2
    cmBeg.Top = cmBeg.Top - delta
    cmReset.Top = cmReset.Top - delta
    cmFiltrGo.Top = cmFiltrGo.Top - delta
    cmAdvan.Top = cmAdvan.Top - delta
    cmExit.Top = cmExit.Top - delta
    lbTema.Visible = False
    lbType.Visible = False
    laKategor.Visible = False
    laTema.Visible = False
    cmAdvan.Caption = "Дополнительно >>"
End If
End Sub

Private Sub cmBeg_Click()
cmReset_Click
tbStartDate.Text = Orders.tbStartDate.Text
cbStartDate.value = Orders.cbStartDate.value
tbEndDate.Text = Orders.tbEndDate.Text
cbEndDate.value = Orders.cbEndDate.value
If Orders.cbClose.value = 1 Then
    lbSelectAll lbStatus
Else
    lbSelectAll lbStatus, "закрыт"
End If

End Sub

Private Sub cmExit_Click()
Me.Visible = False
End Sub

Private Sub cmFiltrGo_Click()
Dim str As String, strWhere As String, I As Integer

Orders.chConflict.value = 0
For I = 1 To orColNumber
    orSqlWhere(I) = ""
Next I
' ********************** поле Дата
strWhere = ""
If cbStartDate.value = 1 Then
    If Not isDateTbox(tbStartDate) Then Exit Sub
    strWhere = Orders.strWhereByValCol(tbStartDate.Text, orData, ">")
End If
If cbEndDate.value = 1 Then
    If Not isDateTbox(tbEndDate) Then Exit Sub
    If strWhere = "" Then
        strWhere = Orders.strWhereByValCol(tbEndDate.Text, orData, "<")
    Else
        strWhere = strWhere & " AND " & Orders.strWhereByValCol(tbEndDate.Text, orData, "<")
    End If
End If

If cbStartDate.value <> 1 And cbEndDate.value = 1 Then
    strWhere = strWhere & " OR " & Orders.strWhereByValCol("", orData) ' + незаполненные даты
End If
orSqlWhere(orData) = strWhere
'********************
lbToOrSqlWhere lbM, orMen
lbToOrSqlWhere lbStatus, orStatus
lbToOrSqlWhere lbFirm, orFirma, "notAll"
If lbType.Visible Then
    lbToOrSqlWhere lbType, orType
    If lbTema.Enabled Then
        lbToOrSqlWhere lbTema, orTema, "byId"
    End If
End If
If opAll Then
    orSqlWhere(0) = "(Orders.paid)>=[Orders].[ordered]"
ElseIf opNotAll Then
    orSqlWhere(0) = "((Orders.paid)<[Orders].[ordered]) OR ((Orders.ordered)>0) AND ((Orders.paid)) Is Null"
Else
    orSqlWhere(0) = ""
End If
If opAlls Then
    str = "(Orders.shipped)>=[Orders].[ordered]"
ElseIf opNotAlls Then
    str = "((Orders.shipped)<[Orders].[ordered]) OR ((Orders.ordered)>0) AND ((Orders.shipped)) Is Null"
Else
    str = ""
End If
If orSqlWhere(0) = "" Or str = "" Then
    orSqlWhere(0) = orSqlWhere(0) & str
Else
    orSqlWhere(0) = orSqlWhere(0) & ") AND (" & str
End If

Orders.getWhereInvoice

Filtr.Visible = False
Orders.begFiltrDisable
Orders.loadWithFiltr ("multi")
End Sub

Sub lbSelectAll(listBox As listBox, Optional except As String = "##")
Dim I As Integer

For I = 0 To listBox.ListCount - 1
    If listBox.List(I) = except Then
        listBox.Selected(I) = False
    Else
        listBox.Selected(I) = True
    End If
Next I
End Sub

Public Sub cmReset_Click()
cbStartDate.value = 0
cbEndDate.value = 0
lbFirm.Clear
opAny.value = True
opAnys.value = True
End Sub

Private Sub lbType_Click()
If lbType.Text = "Н" Then
    lbTema.Enabled = True
Else
    lbDeSelectAll lbTema
    lbTema.Enabled = False
End If
End Sub

Private Sub Option2_Click()

End Sub
