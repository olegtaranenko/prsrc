VERSION 5.00
Begin VB.Form Reports 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Отчеты"
   ClientHeight    =   1956
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4308
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1956
   ScaleWidth      =   4308
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Статистика посещений по фирмам "
      Height          =   852
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   4035
      Begin VB.CommandButton cmFirms 
         BackColor       =   &H8000000A&
         Caption         =   "Применить"
         Height          =   315
         Left            =   2820
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox tbStartDate 
         Height          =   285
         Left            =   360
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "11.2002"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox tbEndDate 
         Height          =   285
         Left            =   1800
         MaxLength       =   7
         TabIndex        =   2
         Text            =   "11.2002"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "по"
         Height          =   192
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   192
      End
      Begin VB.Label laPo 
         BackColor       =   &H8000000A&
         Caption         =   "с"
         Height          =   192
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   192
      End
   End
   Begin VB.CommandButton cmExit 
      BackColor       =   &H8000000A&
      Caption         =   "Выход"
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      Top             =   1320
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

Report.Regim = "stat"
Report.Show vbModal
Me.MousePointer = flexDefault
End Sub

 
Function periodValid(tbStartDate As TextBox, tbEndDate As TextBox) As Boolean

periodValid = False
If Not textBoxDateValid(tbStartDate) Then Exit Function
If tmpDate < CDate("01.11.2000") Then tbStartDate.Text = "11.2000"

If Not textBoxDateValid(tbEndDate) Then Exit Function
If tmpDate > CurDate Then tbEndDate.Text = Format(CurDate, "mm.yyyy")

periodValid = True
End Function

Function textBoxDateValid(tb As TextBox) As Boolean
Dim str As String

textBoxDateValid = False
str = Trim(tb.Text)
tb.Text = str
If Len(str) <> 7 Then GoTo AA
str = "01." & str
If IsDate(str) Then
    tmpDate = str
    textBoxDateValid = True
Else
AA: MsgBox "Неверный формат даты", , "Error"
'    tb.SelStart = 0
  '  tb.SelLength = 2
    tb.SetFocus
End If
End Function


Private Sub cmYear_Click()


End Sub

Private Sub Form_Load()
Dim I As Integer
Frame3.Visible = True
Frame2.Visible = True
tbEndDate.Text = Format(CurDate, "mm.yyyy")
tbStartDate.Text = tbEndDate.Text

End Sub

