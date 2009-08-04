VERSION 5.00
Begin VB.Form ExcelParamDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Параметры вывода в Excel"
   ClientHeight    =   3012
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   7056
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3012
   ScaleWidth      =   7056
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tbKegl 
      Height          =   288
      Left            =   5040
      TabIndex        =   8
      Text            =   "8"
      Top             =   240
      Width           =   492
   End
   Begin VB.TextBox tbMainTitle 
      Height          =   288
      Left            =   600
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1560
      Width           =   5172
   End
   Begin VB.TextBox tbRate 
      Height          =   288
      Left            =   600
      TabIndex        =   4
      Text            =   "30.1"
      Top             =   960
      Width           =   732
   End
   Begin VB.OptionButton rbUE 
      Caption         =   "В условных единицах"
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   240
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
      Caption         =   "Отмена"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbKegl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Кегль отчета"
      Height          =   192
      Left            =   3744
      TabIndex        =   7
      Top             =   240
      Width           =   1116
   End
   Begin VB.Label lbMainTitle 
      AutoSize        =   -1  'True
      Caption         =   "Основной заголовок отчета"
      Height          =   192
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   2316
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
Public RubRate As Double
Public mainReportTitle As String
Public kegl As Integer


Dim doUnload As Boolean


Private Sub CancelButton_Click()
    exitCode = vbCancel
    Unload Me
End Sub

Private Sub Form_Load()
    rbUE.value = True
    tbRate.Text = getCurrentRate()
    If mainTitle <> "" Then
        tbMainTitle.Text = mainReportTitle
    End If
    If kegl <> 0 Then
        tbKegl.Text = kegl
    End If
End Sub

Private Sub OKButton_Click()
    exitCode = vbOK
    doUnload = True
    If IsNumeric(tbRate.Text) Then
        RubRate = tbRate.Text
    Else
        MsgBox "Некорректное значение курса", , "Ошибка ввода"
        tbRate.SetFocus
        tbRate.SelStart = 0
        tbRate.SelLength = Len(tbRate.Text)
        doUnload = False
    End If
    
    mainReportTitle = tbMainTitle.Text
    If IsNumeric(tbKegl.Text) Then
        kegl = tbKegl.Text
    Else
        MsgBox "Некорректное значение кегля отчета", , "Ошибка ввода"
        tbKegl.SetFocus
        tbKegl.SelStart = 0
        tbKegl.SelLength = Len(tbKegl.Text)
        doUnload = False
    End If
    
    If doUnload Then
        Unload Me
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
