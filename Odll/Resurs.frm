VERSION 5.00
Begin VB.Form dopResurs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Дополнительный  ресурс  на  "
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "Resurs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Top             =   1800
      Width           =   915
   End
   Begin VB.CommandButton cmChan 
      Caption         =   "Изменить"
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   540
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox cb 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   900
      Width           =   1275
   End
   Begin VB.TextBox tb 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   300
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "будет распределен"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   360
      Width           =   1515
   End
   Begin VB.Label lbSince 
      Caption         =   "на период с  "
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Дополнит.ресурс [час]"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1755
   End
End
Attribute VB_Name = "dopResurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmAdd_Click()
cmChan_Click
End Sub

Private Sub cmChan_Click()
Dim s As Double, i As Integer, j As Integer

    s = Zagruz.lv.SelectedItem.SubItems(zgNomRes)
    If isNumericTbox(tb, 0, 22 - s) Then
        s = 0: j = Mid$(Zagruz.lv.SelectedItem.key, 2)
        For i = j To j + cb.ListIndex
            s = s + nomRes(i)
        Next i
        If s = 0 Then
            MsgBox "Для ввода доп.ресурса на этом интервале требуется " & _
        "хотя бы один день с ненулевой Номинальный ресурс." _
        , , "Доп.ресурс невозможен!"
        Else
            Zagruz.ZagruzLoad ' с учетом новых значений
        End If
        Unload Me
    End If
End Sub

Private Sub cmDel_Click()
tb.Text = "0"
cmChan_Click
End Sub

Private Sub Form_Load()
Dim key As String, i As Integer, str As String, line As Integer

key = ClickItem.key

line = Mid$(ClickItem.key, 2)
For i = line To maxDay
    tmpDate = DateAdd("d", i - 1, CurDate)
'    tmpDate = CurDate + i - 1
    str = Format(tmpDate, "dd.mm.yy")
    If i = line Then
        Me.Caption = Me.Caption & str
        lbSince.Caption = lbSince.Caption & str & " по"
    End If
    cb.AddItem str
Next

If CDbl(Zagruz.lv.ListItems(key).SubItems(zgDopRes)) = 0 Then
    cmChan.Enabled = False
    cmDel.Enabled = False
    cb.ListIndex = 0
Else
    tb.Text = Zagruz.lv.ListItems(key).SubItems(zgDopRes)
    cmAdd.Enabled = False
'    tmpDate = Zagruz.lv.ListItems(key).SubItems(ZagruzDataIsp)
    cb.ListIndex = endRes(line) - line
End If
flEdit = "dop"
End Sub

Private Sub mDelmmand1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
flEdit = ""
End Sub
