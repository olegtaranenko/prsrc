@VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form cfg 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   8700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   7740
      TabIndex        =   16
      Top             =   3780
      Width           =   795
   End
   Begin VB.ListBox lbActive 
      Height          =   432
      ItemData        =   "cfg.frx":0000
      Left            =   1500
      List            =   "cfg.frx":000A
      TabIndex        =   17
      Top             =   780
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   3240
      TabIndex        =   14
      Text            =   "tbMobile"
      Top             =   2100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmProducts 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   12
      Top             =   1056
      Width           =   675
   End
   Begin VB.CommandButton cmNomenks 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   10
      Top             =   756
      Width           =   675
   End
   Begin VB.CommandButton cmSvodka 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   8
      Top             =   456
      Width           =   675
   End
   Begin VB.CommandButton cmLogins 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   6
      Top             =   156
      Width           =   675
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   2100
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1992
      Left            =   120
      TabIndex        =   0
      Top             =   1620
      Width           =   8472
      _ExtentX        =   14944
      _ExtentY        =   3514
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label laInform 
      Height          =   312
      Left            =   4680
      TabIndex        =   15
      Top             =   3780
      Width           =   2412
   End
   Begin VB.Label laGrid 
      Caption         =   "Список доступных баз"
      Height          =   252
      Left            =   180
      TabIndex        =   13
      Top             =   1440
      Width           =   2172
   End
   Begin VB.Label laProducts 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   252
      Left            =   2280
      TabIndex        =   11
      Top             =   1056
      Width           =   5532
   End
   Begin VB.Label laNomenks 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   252
      Left            =   2280
      TabIndex        =   9
      Top             =   756
      Width           =   5532
   End
   Begin VB.Label laSvodka 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   252
      Left            =   2280
      TabIndex        =   7
      Top             =   456
      Width           =   5532
   End
   Begin VB.Label laLogins 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   252
      Left            =   2280
      TabIndex        =   5
      Top             =   156
      Width           =   5532
   End
   Begin VB.Label Label6 
      Caption         =   "Файл составных изделий:"
      Height          =   192
      Left            =   180
      TabIndex        =   4
      Top             =   1080
      Width           =   2052
   End
   Begin VB.Label Label5 
      Caption         =   "Файл простых изделий:"
      Height          =   192
      Left            =   180
      TabIndex        =   3
      Top             =   780
      Width           =   1872
   End
   Begin VB.Label Label4 
      Caption         =   "Файл Сводки:"
      Height          =   192
      Left            =   180
      TabIndex        =   2
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "Файл логинов:"
      Height          =   192
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   1152
   End
End
Attribute VB_Name = "cfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'не забыть выгрузку (см '$$$$)
'Public gCfgFilePath As String
'Public curBaseInd As Integer
'Public workBasePath As String
Public loginsPath As String
Public SvodkaPath As String
Public NomenksPath As String
Public ProductsPath As String
Public isLoad As Boolean
Private clickedRow As Integer

Public Regim As String
'Dim key() As String
'Dim val() As String
'Dim glb() As Boolean

Const bsDbName = 1
Const bsServer = 2
Const bsActive = 3
Const bsPrefix = 4



Function loadFileConfiguration() As Boolean
Dim str As String, I As Integer

loadFileConfiguration = False
'ReDim key(0): ReDim val(0)

'loadFileSettings "local" 'лок cfg.файл

'gCfgFilePath = getParam("gCfgFilePath")
'If gCfgFilePath <> "" Then ' путь к глоб.ф-лу был уже определен
'    If Not loadFileSettings(gCfgFilePath) Then
'        MsgBox "Повторите запуск позже или сообщите Администратору!", , _
        "Не найден путь '" & gCfgFilePath & "'."
'        End
'    End If
'Else
'    gCfgFilePath = App.path & "\" & "global.cfg"
'    loadFileSettings gCfgFilePath далее присвоятся значения по умолчанию
'End If
'loginsPath = getParamOrDefault("loginsPath", "файл логинов")
loginsPath = getEffectiveSetting("loginsPath")
If loginsPath = "" Then loginsPath = _
    "\\Server\C\WebServers\home\petmas.ru\mirror\files\logins."

'SvodkaPath = getParamOrDefault("SvodkaPath", "файл Сводки")
SvodkaPath = getEffectiveSetting("SvodkaPath")
If SvodkaPath = "" Then SvodkaPath = _
    "\\Server\C\WebServers\home\petmas.ru\mirror\files\svodkaW."

'NomenksPath = getParamOrDefault("NomenksPath", "файл простых изделий")
NomenksPath = getEffectiveSetting("NomenksPath")
If NomenksPath = "" Then NomenksPath = _
    "\\Server\C\WebServers\home\petmas.ru\mirror\files\Nomenks."

'ProductsPath = getParamOrDefault("ProductsPath", "файл составных изделий")
ProductsPath = getEffectiveSetting("ProductsPath")
If ProductsPath = "" Then ProductsPath = _
    "\\Server\C\WebServers\home\petmas.ru\mirror\files\Products."

DD:
    loadFileConfiguration = True
EE: 'запись параметров

End Function

Private Sub cmExit_Click()
    Unload Me
End Sub


Private Sub cmLogins_Click()
cdOpen.DialogTitle = "Выберите файл логинов."
cdOpen.FileName = ""
cdOpen.ShowOpen
If cdOpen.FileName = "" Then Exit Sub
loginsPath = cdOpen.FileName
saveFileSettings appCfgFile, appSettings
laLogins.Caption = cdOpen.FileName

End Sub

Private Sub cmNomenks_Click()
cdOpen.DialogTitle = "Выберите файл простых изделий."
cdOpen.FileName = ""
cdOpen.ShowOpen
If cdOpen.FileName = "" Then Exit Sub
NomenksPath = cdOpen.FileName
saveFileSettings appCfgFile, appSettings
laNomenks.Caption = cdOpen.FileName
End Sub

Private Sub cmProducts_Click()
cdOpen.DialogTitle = "Выберите файл составных изделий."
cdOpen.FileName = ""
cdOpen.ShowOpen
If cdOpen.FileName = "" Then Exit Sub
ProductsPath = cdOpen.FileName
saveFileSettings appCfgFile, appSettings
laProducts.Caption = cdOpen.FileName
End Sub

Private Sub cmSvodka_Click()
cdOpen.DialogTitle = "Выберите файл Сводки."
cdOpen.FileName = ""
cdOpen.ShowOpen
If cdOpen.FileName = "" Then Exit Sub
SvodkaPath = cdOpen.FileName
saveFileSettings appCfgFile, appSettings
laSvodka.Caption = cdOpen.FileName
End Sub

Private Sub Form_Load()
isLoad = True
End Sub

Sub setRegim()
Dim I As Integer

If Regim = "comtexAdmin" Then
    Grid.rows = 2: Grid.Cols = 5: Grid.Clear
    Grid.FormatString = "|<Бухгалтерская база|<Cервер|<Совместная работа|<Префикс"
    MsgBox "Будьте уверены, что вы знаете, что вы делаете. В противном случае изменения сделанные в открывающеммся окне могут повлечь за собой проблемы в режиме совместной работы Prior и Comtex", , "Предупреждение"
Else
    Grid.rows = 2: Grid.Cols = 2: Grid.Clear
    Grid.FormatString = "|<Усл.название|<Полный путь к файлу|Рабочая|Текущая"
End If

If Regim = "pathSet" Then
    Me.Caption = "Установка путей"
    Grid.ColWidth(0) = 0
'    laGlobal.Caption = gCfgFilePath
    laLogins.Caption = loginsPath
    laSvodka.Caption = SvodkaPath
    laNomenks.Caption = NomenksPath
    laProducts.Caption = ProductsPath
ElseIf Regim = "comtexAdmin" Then
    Me.Caption = "Выбор базы"
'    laGlobal.Visible = False
    laLogins.Visible = False
    laSvodka.Visible = False
    laNomenks.Visible = False
    laProducts.Visible = False
'    cmGlobal.Visible = False
    
    I = laGrid.Top
    laGrid.Top = Me.Top + 100
    I = laGrid.Top - I
    Grid.Top = Grid.Top + I
    'Me.Height = Me.Height + i - 500 '(место под кнопки не нужно)
    Me.Height = Me.Height + I - 200
    laGrid.ZOrder
    Grid.ZOrder
    Grid.ColWidth(0) = 0
    Grid.ColWidth(bsDbName) = 2000
    Grid.ColWidth(bsServer) = 1000
    Grid.ColWidth(bsActive) = 800
    Grid.ColWidth(bsPrefix) = 800
    I = Grid.Width
    Grid.Width = Grid.ColWidth(bsDbName) + Grid.ColWidth(bsServer) + Grid.ColWidth(bsActive) + Grid.ColWidth(bsPrefix) + 350
    I = Grid.Width - I
    Me.Width = Me.Width + I
    cmExit.Left = cmExit.Left + I
    cmExit.Top = Grid.Top + Grid.Height + 100
    sql = "GuideVenture"
    
    Set table = myOpenRecordSet("##72", sql, dbOpenForwardOnly)
    If table Is Nothing Then myBase.Close: End
    I = 0
    While Not table.EOF
        Grid.TextMatrix(I + 1, bsDbName) = table!ventureName
        Grid.TextMatrix(I + 1, bsServer) = table!sysname
        If table!standalone = 0 Then
            Grid.TextMatrix(I + 1, bsActive) = "Да"
        Else
            Grid.TextMatrix(I + 1, bsActive) = "Нет"
        End If
        
        Grid.TextMatrix(I + 1, bsPrefix) = table!invCode

        table.MoveNext
        Grid.AddItem ""
        I = I + 1
    Wend
    table.Close
    Grid.removeItem Grid.rows - 1
End If
  

Grid_EnterCell
End Sub


Sub baseOpen()
Dim str As String, dburl As String
    dburl = getEffectiveSetting("dbUrl")
    
    If dburl = "" Then
        fatalError "Необходимо исправить конфигурацию запуска программы." & vbCr & "Не установлено значение параметра dbUrl"
    End If
    
On Error GoTo ERRb

    Set myBase = wrkDefault.OpenDatabase("Connection1", _
       dbDriverNoPrompt, False, _
       "ODBC;UID=dba;PWD=sql;DSN=" & dburl)
    If myBase Is Nothing Then End
    
    sql = "call bootstrap_blocking()"
    If myExecute("##bootstrap", sql, 0) = 0 Then End
    
    Exit Sub
    
ERRb:
       
    If errorCodAndMsg("388", -100) Then '##388
        fatalError "Проблемы с доступом к серверу базы данных." & vbCr & "dbUrl = " & dburl
    End If
End Sub


Private Sub Grid_Click()
    clickedRow = Grid.MouseRow
    If Grid.MouseRow = 0 Then
    '        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
    End If
End Sub


Sub changeMsg()
    MsgBox "Вы изменили текущую базу! Чтобы измененния вошли в силу " & _
    "перезапустите программу?"
End Sub

Private Sub Grid_DblClick()
Dim ind As Integer, I As Integer, str As String

    If Regim = "comtexAdmin" Then
        If Grid.col = bsActive Then
            listBoxInGridCell lbActive, Grid, Grid.TextMatrix(Grid.MouseRow, Grid.MouseCol)
        ElseIf Grid.col = bsPrefix Then
            textBoxInGridCell tbMobile, Grid
        End If
        
        Exit Sub
    End If

End Sub
Sub lbHide()
    tbMobile.Visible = False
    Grid.Enabled = True
    On Error Resume Next
    Grid.SetFocus
    Grid_EnterCell
    lbActive.Visible = False
End Sub

Private Sub Grid_EnterCell()

    If Regim = "comtexAdmin" And Grid.col <= 2 Then
        Grid.CellBackColor = vbYellow
    Else
        Grid.CellBackColor = &H88FF88
    End If
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Grid_DblClick
End Sub

Private Sub Grid_LeaveCell()
    Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub lbActive_DblClick()
Dim success As Integer
    If noClick Then Exit Sub
    If lbActive.Visible = False Then Exit Sub
    
    'Grid.Text = lbActive.Text
    sql = "select slave_set_standalone(" & lbActive.ListIndex _
        & ", '" & Grid.TextMatrix(clickedRow, bsServer) & "'" _
        & ", 1)"
    If byErrSqlGetValues("##1.1", sql, success) Then
        If success = 0 Then
            MsgBox "При изменении параметра произошла ошибка." _
            & vbCr & "Возможно недоступен сервер, в котором также нужно установить настройку ""Работать Независимо""." _
            & " Некорректная работа серверов в режиме интеграции может повлечь за собой разрушение целостности интегрированных данных, которое потом будет трудно исправить." _
            & vbCr & "Если вы не уверены в том, что сделали все правильно, обязательно сообщите администратору" _
            , , "Предупреждение"
        End If

        Grid.Text = lbActive.Text
        Orders.initVentureLB
        clickedRow = -1
    End If
    lbHide

End Sub

Private Sub lbActive_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lbActive_DblClick
    If KeyCode = vbKeyEscape Then lbHide

End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim I As Integer

    If KeyCode = vbKeyReturn Then
        If Regim = "comtexAdmin" Then
            sql = "update guideVenture set invCode = " & tbMobile.Text & " where sysname = '" & Grid.TextMatrix(clickedRow, bsServer) & "'"
            I = myExecute("##1.2", sql)
        End If
        Grid.Text = tbMobile.Text
        saveFileSettings appCfgFile, appSettings
        lbHide
    ElseIf KeyCode = vbKeyEscape Then
        lbHide
    End If

End Sub
