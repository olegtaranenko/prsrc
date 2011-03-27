VERSION 5.00
Begin VB.Form ExcelParamDialog 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Параметры вывода в Excel"
   ClientHeight    =   3252
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   9744
   FillColor       =   &H8000000A&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3252
   ScaleWidth      =   9744
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ckHeaders 
      BackColor       =   &H8000000A&
      Caption         =   "Включая файл заголовков"
      Height          =   252
      Left            =   360
      TabIndex        =   21
      Top             =   1920
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3012
   End
   Begin VB.TextBox tbOutputPath 
      Height          =   288
      Left            =   600
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   7452
   End
   Begin VB.TextBox tbContact2 
      Height          =   288
      Left            =   600
      TabIndex        =   17
      Top             =   2760
      Width           =   7452
   End
   Begin VB.TextBox tbContact1 
      Height          =   288
      Left            =   600
      TabIndex        =   15
      Top             =   2160
      Width           =   7452
   End
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
      Left            =   8400
      MaskColor       =   &H8000000A&
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H8000000A&
      Caption         =   "OK"
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lbOutputPath 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Вывод файла в директорию"
      Height          =   192
      Left            =   360
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   2304
   End
   Begin VB.Label lbContact2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Подзаголовок (телефоны)"
      Height          =   192
      Left            =   360
      TabIndex        =   18
      Top             =   2520
      Width           =   2172
   End
   Begin VB.Label lbContact1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Подзаголовок (контакты)"
      Height          =   192
      Left            =   360
      TabIndex        =   16
      Top             =   1920
      Width           =   2100
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
Public contact1 As String
Public contact2 As String
Public kegl As Integer
Public commonRabbat As Single
Public doProdCategory As Boolean
Public prodCategoryId As Integer
Public showRabbat As Boolean
Public priceType As Integer
Public includeHeaders As Boolean

Public Regim As String
Public withPrice As Boolean
Public CsvAsOutput As Boolean
Public csvFileName As String

Dim doUnload As Boolean


Private Sub CancelButton_Click()
    Dim Value As String
    Unload Me
End Sub


Private Sub cbPriceType_Click()
    If cbPriceType.ListIndex = 1 Then
        priceType = 1
        tbCommonRabbat.Visible = False
        laCommonRabbat.Visible = False
    Else
        priceType = 0
        tbCommonRabbat.Visible = True
        laCommonRabbat.Visible = True
    End If
End Sub

Private Sub Form_Load()
    exitCode = vbCancel
    If csvFileName <> "" Then
        Me.Caption = csvFileName
    End If
    
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
    
    If CsvAsOutput Then
        lbOutputPath.Visible = True
        tbOutputPath.Visible = True
        tbOutputPath.Text = getEffectiveSetting("ProductsPath", "")
    Else
        lbOutputPath.Visible = False
        tbOutputPath.Visible = False
    End If

    If kegl < 0 Then
        tbKegl.Visible = False
        lbKegl.Visible = False
    ElseIf kegl > 0 Then
        tbKegl.Text = kegl
        tbKegl.Visible = True
        lbKegl.Visible = True
    End If
    
    If Regim = "pricePM" Then
        tbContact1.Visible = False
        tbContact2.Visible = False
        lbContact1.Visible = False
        lbContact2.Visible = False
    Else
        tbContact1.Visible = True
        tbContact2.Visible = True
        lbContact1.Visible = True
        lbContact2.Visible = True
    End If
    
    
    If mainReportTitle = "-" Then
        tbMainTitle.Visible = False
        lbMainTitle.Visible = False
    ElseIf mainReportTitle <> "" Then
        tbMainTitle.Text = mainReportTitle
    End If
    
    If contact1 = "-" Then
        tbContact1.Visible = False
        lbContact1.Visible = False
    ElseIf contact1 <> "" Then
        tbContact1.Text = contact1
    End If
    
    If contact2 = "-" Then
        tbContact2.Visible = False
        lbContact2.Visible = False
    ElseIf contact2 <> "" Then
        tbContact2.Text = contact2
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
            showRabbat = False
        Else
            showRabbat = True
        End If
    End If

    If showRabbat Then
        tbCommonRabbat.Visible = True
        laCommonRabbat.Visible = True
    Else
        tbCommonRabbat.Visible = False
        laCommonRabbat.Visible = False
    End If
    
    If includeHeaders Then
        ckHeaders.Visible = True
        ckHeaders.Value = getEffectiveSetting("includeHeaders", 1)
    Else
        ckHeaders.Visible = False
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If exitCode = vbOK Then
        saveTheParams
    End If
    showRabbat = False
    'includeHeaders = False
    Regim = ""
    csvFileName = ""
End Sub

Private Sub saveTheParams()

    setAppSetting Regim & ".kegl", kegl
    setAppSetting Regim & ".title", mainReportTitle
    setAppSetting Regim & ".ue", outputUE
    setAppSetting Regim & ".rabbat", commonRabbat
    setAppSetting Regim & ".pricetype", cbPriceType.ListIndex
    setAppSetting Regim & ".pricetype", cbPriceType.ListIndex
    setAppSetting Regim & ".includeHeaders", ckHeaders.Value
    If tbContact1.Visible Then
        setAppSetting ".contact1", contact1
    End If
    If tbContact2.Visible Then
        setAppSetting ".contact2", contact2
    End If

    saveFileSettings getAppCfgDefaultName, appSettings
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
    
    If doProdCategory And cbProdCategory.ListIndex >= 0 Then
        doProdCategory = False
        prodCategoryId = cbProdCategory.ItemData(cbProdCategory.ListIndex)
    End If
    
    If CsvAsOutput Then
        Dim csvPath As String
        csvPath = tbOutputPath.Text
        If csvPath <> "" And Right(csvPath, 1) <> "\" Then
            csvPath = csvPath & "\"
        End If
        
        If Dir$(csvPath) = "" Then
            MsgBox tbOutputPath.Text & ": Не существует такая директория или доступ к ней заблокирован.", , "Проверьте путь"
            Exit Sub
        End If
        'Fileexists
        If tbOutputPath.Text <> getEffectiveSetting("ProductsPath", "") Then
            setSiteSetting "ProductsPath", tbOutputPath.Text
            saveSiteSettings
        End If
    End If
    contact1 = tbContact1.Text
    contact2 = tbContact2.Text
    includeHeaders = ckHeaders.Visible And ckHeaders.Value = 1
    
    
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

