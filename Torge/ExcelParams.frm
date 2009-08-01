VERSION 5.00
Begin VB.Form ExcelParamDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Параметры вывода в Excel"
   ClientHeight    =   3192
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   6036
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   6036
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tbRate 
      Height          =   288
      Left            =   720
      TabIndex        =   4
      Text            =   "30.1"
      Top             =   1080
      Width           =   1572
   End
   Begin VB.OptionButton rbUE 
      Caption         =   "В условных единицах"
      Height          =   372
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   2172
   End
   Begin VB.OptionButton rbRub 
      Caption         =   "В рублях по курсу"
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   2172
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "ExcelParamDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public exitCode As Integer
Public outputUE As Boolean
Public rubRate As Double


Private Sub CancelButton_Click()
    exitCode = vbCancel
    Unload Me
End Sub

Private Sub Form_Load()
    rbUE.value = True
    tbRate.Text = getCurrentRate()
End Sub

Private Sub OKButton_Click()
    exitCode = vbOK
    If IsNumeric(tbRate.Text) Then
        rubRate = tbRate.Text
        Unload Me
    Else
        MsgBox "Некорректное значение курса", , "Ошибка ввода"
        tbRate.SetFocus
        tbRate.SelStart = 0
        tbRate.SelLength = Len(tbRate.Text)
    End If
End Sub

Private Sub rbRub_Click()
    tbRate.Enabled = True
    outputUE = False
End Sub

Private Sub rbUE_Click()
    tbRate.Enabled = False
    outputUE = True
End Sub
