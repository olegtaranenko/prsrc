VERSION 5.00
Begin VB.Form ExcelParamDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Параметры вывода в Excel"
   ClientHeight    =   2388
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   7488
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2388
   ScaleWidth      =   7488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tbCommonRabbat 
      Height          =   288
      Left            =   4296
      TabIndex        =   11
      Text            =   "8"
      Top             =   1440
      Width           =   492
   End
   Begin VB.ComboBox cbProdCategory 
      Height          =   288
      Left            =   4320
      TabIndex        =   10
      Text            =   "web"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.TextBox tbKegl 
      Height          =   288
      Left            =   4320
      TabIndex        =   8
      Text            =   "8"
      Top             =   720
      Width           =   492
   End
   Begin VB.TextBox tbMainTitle 
      Height          =   288
      Left            =   600
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   360
      Width           =   6612
   End
   Begin VB.TextBox tbRate 
      Height          =   288
      Left            =   600
      TabIndex        =   4
      Text            =   "30.1"
      Top             =   1440
      Width           =   732
   End
   Begin VB.OptionButton rbUE 
      Caption         =   "В условных единицах"
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   2172
   End
   Begin VB.OptionButton rbRub 
      Caption         =   "В рублях по курсу"
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2172
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label laCommonRabbat 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Общая скидка (%)"
      Height          =   192
      Left            =   2820
      TabIndex        =   12
      Top             =   1440
      Width           =   1416
   End
   Begin VB.Label lProdCategory 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Категория изделий"
      Height          =   192
      Left            =   2592
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1644
   End
   Begin VB.Label lbKegl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Кегль отчета "
      Height          =   192
      Left            =   3108
      TabIndex        =   7
      Top             =   720
      Width           =   1152
   End
   Begin VB.Label lbMainTitle 
      AutoSize        =   -1  'True
      Caption         =   "Основной заголовок отчета"
      Height          =   192
      Left            =   360
      TabIndex        =   5
      Top             =   120
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
Public commonRabbat As Single
Public doProdCategory As Boolean
Public prodCategoryId As Integer
Public showRabbat As Boolean
Public Regim As String
Public withPrice As Boolean


Dim doUnload As Boolean


Private Sub CancelButton_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    exitCode = vbCancel
    If Not withPrice Then
        rbUE.Visible = False
        rbRub.Visible = False
        tbRate.Visible = False
    Else
        rbUE.Visible = True
        rbRub.Visible = True
        tbRate.Visible = True
    End If
    If outputUE Then
        rbUE.Value = True
    Else
        rbRub.Value = True
    End If
    tbRate.Text = getCurrentRate()
    If mainTitle <> "" Then
        tbMainTitle.Text = mainReportTitle
    End If
    If kegl <> 0 Then
        tbKegl.Text = kegl
    End If
    
    initProdCategoryBox cbProdCategory
    If doProdCategory Then
        lProdCategory.Visible = True
        cbProdCategory.Visible = True
        If cbProdCategory.ListCount > 0 Then
            cbProdCategory.ListIndex = 1
        Else
            cbProdCategory.ListIndex = 0
        End If
    Else
        lProdCategory.Visible = False
        cbProdCategory.Visible = False
    End If
    
    If showRabbat Then
        tbCommonRabbat.Visible = True
        laCommonRabbat.Visible = True
        tbCommonRabbat.Text = CStr(commonRabbat)
    Else
        tbCommonRabbat.Visible = False
        laCommonRabbat.Visible = False
        tbCommonRabbat.Text = "0"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If exitCode = vbOK Then
        setAppSetting Regim & ".kegl", kegl
        setAppSetting Regim & ".title", mainReportTitle
        setAppSetting Regim & ".ue", outputUE
        setAppSetting Regim & ".rabbat", commonRabbat
    
        saveFileSettings getAppCfgDefaultName, appSettings
    End If
    Regim = ""
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
    
    If doProdCategory Then
        doProdCategory = False
        prodCategoryId = cbProdCategory.ItemData(cbProdCategory.ListIndex)
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

Private Sub tbCommonRabbat_Change()
    If IsNumeric(tbCommonRabbat.Text) Then
        commonRabbat = CSng(tbCommonRabbat.Text)
    End If
End Sub
