VERSION 5.00
Begin VB.Form ExcelParamDialog 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Параметры вывода в Excel"
   ClientHeight    =   2760
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   8244
   FillColor       =   &H8000000A&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   8244
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbPriceType 
      Height          =   288
      ItemData        =   "ExcelParams.frx":0000
      Left            =   6960
      List            =   "ExcelParams.frx":000A
      TabIndex        =   13
      Text            =   "дилеров"
      Top             =   840
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.TextBox tbCommonRabbat 
      Height          =   288
      Left            =   6960
      TabIndex        =   11
      Text            =   "40"
      Top             =   1200
      Width           =   612
   End
   Begin VB.ComboBox cbProdCategory 
      Height          =   288
      ItemData        =   "ExcelParams.frx":001C
      Left            =   4320
      List            =   "ExcelParams.frx":001E
      TabIndex        =   10
      Text            =   "web"
      Top             =   1200
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox tbKegl 
      Height          =   288
      Left            =   4320
      TabIndex        =   8
      Text            =   "8"
      Top             =   840
      Width           =   492
   End
   Begin VB.TextBox tbMainTitle 
      Height          =   288
      Left            =   600
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   360
      Width           =   7452
   End
   Begin VB.TextBox tbRate 
      Height          =   288
      Left            =   600
      TabIndex        =   4
      Text            =   "30.1"
      Top             =   1560
      Width           =   732
   End
   Begin VB.OptionButton rbUE 
      BackColor       =   &H8000000A&
      Caption         =   "В условных единицах"
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   2052
   End
   Begin VB.OptionButton rbRub 
      BackColor       =   &H8000000A&
      Caption         =   "В рублях по курсу"
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   2172
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H8000000A&
      Caption         =   "Отмена"
      Height          =   375
      Left            =   6840
      MaskColor       =   &H8000000A&
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H8000000A&
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lbPriceType 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Прайс для:"
      Height          =   192
      Left            =   5808
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.Label laCommonRabbat 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Общая скидка (%)"
      Height          =   192
      Left            =   5460
      TabIndex        =   12
      Top             =   1200
      Width           =   1416
   End
   Begin VB.Label lProdCategory 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Категория изделий"
      Height          =   192
      Left            =   2592
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   1644
   End
   Begin VB.Label lbKegl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Кегль отчета "
      Height          =   192
      Left            =   3108
      TabIndex        =   7
      Top             =   840
      Width           =   1152
   End
   Begin VB.Label lbMainTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
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
Public priceType As Integer

Public Regim As String
Public withPrice As Boolean


Dim doUnload As Boolean


Private Sub CancelButton_Click()
    Dim Value As String
    Unload Me
End Sub


Private Sub cbPriceType_Click()
    If cbPriceType.ListIndex = 1 Then
        priceType = 1
        tbCommonRabbat.Visible = True
        laCommonRabbat.Visible = True
    Else
        priceType = 0
        tbCommonRabbat.Visible = False
        laCommonRabbat.Visible = False
    End If
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
    
    lbPriceType.Visible = False
    cbPriceType.Visible = False
    laCommonRabbat.Visible = False
    tbCommonRabbat.Visible = False
    tbCommonRabbat.Text = CStr(commonRabbat)
    
    If Regim = "awards" Then
        lbPriceType.Visible = True
        cbPriceType.Visible = True
        
        If priceType = 1 Then
            cbPriceType.ListIndex = 1
        Else
            cbPriceType.ListIndex = 0
        End If
        
        If cbPriceType.ListIndex = 1 Then
            showRabbat = True
        Else
            showRabbat = False
        End If
    End If

    If showRabbat Then
        tbCommonRabbat.Visible = True
        laCommonRabbat.Visible = True
    End If
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If exitCode = vbOK Then
        setAppSetting Regim & ".kegl", kegl
        setAppSetting Regim & ".title", mainReportTitle
        setAppSetting Regim & ".ue", outputUE
        setAppSetting Regim & ".rabbat", commonRabbat
        setAppSetting Regim & ".pricetype", cbPriceType.ListIndex
    
        saveFileSettings getAppCfgDefaultName, appSettings
    End If
    showRabbat = False
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
