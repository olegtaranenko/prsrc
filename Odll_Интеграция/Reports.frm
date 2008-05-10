VERSION 5.00
Begin VB.Form Reports 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Отчеты"
   ClientHeight    =   5340
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4308
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4308
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Пересчет статистики для Справочника"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   1140
      Visible         =   0   'False
      Width           =   4035
      Begin VB.CommandButton cmYear 
         Caption         =   "Применить"
         Height          =   315
         Left            =   2820
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Статистика посещений по фирмам "
      Height          =   2835
      Left            =   120
      TabIndex        =   6
      Top             =   1860
      Visible         =   0   'False
      Width           =   4035
      Begin VB.ListBox lbType 
         Height          =   1200
         ItemData        =   "Reports.frx":0000
         Left            =   1500
         List            =   "Reports.frx":0016
         TabIndex        =   17
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmFirms 
         Caption         =   "Применить"
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton opRA 
         Caption         =   "Рекламщики"
         Height          =   255
         Left            =   300
         TabIndex        =   12
         Top             =   1920
         Width           =   1515
      End
      Begin VB.OptionButton opKK 
         Caption         =   "Конечники"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   1620
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.ListBox lbTema 
         Enabled         =   0   'False
         Height          =   2352
         Left            =   2100
         MultiSelect     =   2  'Extended
         TabIndex        =   10
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox tbStartDate 
         Height          =   285
         Left            =   480
         MaxLength       =   7
         TabIndex        =   8
         Text            =   "11.2002"
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox tbEndDate 
         Height          =   285
         Left            =   480
         MaxLength       =   7
         TabIndex        =   7
         Text            =   "11.2002"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label laPo 
         Caption         =   "с    по"
         Height          =   675
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Статистика по менеджерам"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4035
      Begin VB.TextBox tbEndDate2 
         Height          =   285
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   4
         Text            =   "11.2002"
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox tbStartDate2 
         Height          =   285
         Left            =   480
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "11.2000"
         Top             =   420
         Width           =   735
      End
      Begin VB.CommandButton cmManag 
         Caption         =   "Применить"
         Height          =   315
         Left            =   2820
         TabIndex        =   2
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "по"
         Height          =   195
         Left            =   1260
         TabIndex        =   14
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label1 
         Caption         =   "с"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   135
      End
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      Top             =   4860
      Width           =   855
   End
End
Attribute VB_Name = "Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmFirms_Click()
Dim strWhere As String
Me.MousePointer = flexHourglass

If Not periodValid(tbStartDate, tbEndDate) Then Exit Sub

If opKK.value Then
    Report.Regim = "KK"
Else
    Report.Regim = "RA"
End If
Report.Show vbModal
Me.MousePointer = flexDefault
End Sub

Private Sub cmKK_Click()
End Sub

Private Sub cmManag_Click()
If Not periodValid(tbStartDate2, tbEndDate2) Then Exit Sub

Me.MousePointer = flexHourglass
Report.Regim = "Manag"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmRA_Click()

End Sub
 
Function periodValid(tbStartDate As TextBox, tbEndDate As TextBox) As Boolean

periodValid = False
If Not textBoxDateValid(tbStartDate) Then Exit Function
If tmpDate < CDate("01.11.2000") Then tbStartDate.Text = "11.2000"

If Not textBoxDateValid(tbEndDate) Then Exit Function
If tmpDate > curDate Then tbEndDate.Text = Format(curDate, "mm.yyyy")

periodValid = True
End Function

Function textBoxDateValid(tb As TextBox) As Boolean
Dim str As String

textBoxDateValid = False
str = trimAll(tb.Text)
tb.Text = str
If Len(str) <> 7 Then GoTo AA
str = "01." & str
If IsDate(str) Then
    tmpDate = str
    textBoxDateValid = True
Else
AA: MsgBox "Неверный формат даты", , "Error"
    tb.SetFocus
End If
End Function


Private Sub cmYear_Click()

Me.MousePointer = flexHourglass
statistic "all"
Me.MousePointer = flexDefault

End Sub

Private Sub Form_Load()
Dim I As Integer
    Frame3.Visible = True
    Frame2.Visible = True
tbEndDate.Text = Format(curDate, "mm.yyyy")
tbEndDate2.Text = Format(curDate, "mm.yyyy")
tbStartDate.Text = tbEndDate.Text
tbStartDate2.Text = tbEndDate2.Text

For I = 0 To Filtr.lbTema.ListCount - 1
   lbTema.List(I) = Filtr.lbTema.List(I)
Next I
lbType.Selected(0) = True
End Sub

Private Sub lbType_Click()
If lbType.Text = "Н" Then
    lbTema.Enabled = True
Else
    lbDeSelectAll lbTema
    lbTema.Enabled = False
End If

End Sub

