VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form cfg 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   4524
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4524
   ScaleWidth      =   8700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   7740
      TabIndex        =   19
      Top             =   4020
      Width           =   795
   End
   Begin VB.ListBox lbActive 
      Height          =   432
      ItemData        =   "cfg.frx":0000
      Left            =   1500
      List            =   "cfg.frx":000A
      TabIndex        =   20
      Top             =   1020
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   3240
      TabIndex        =   17
      Text            =   "tbMobile"
      Top             =   2340
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmProducts 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   15
      Top             =   1290
      Width           =   675
   End
   Begin VB.CommandButton cmNomenks 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   13
      Top             =   990
      Width           =   675
   End
   Begin VB.CommandButton cmSvodka 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   11
      Top             =   690
      Width           =   675
   End
   Begin VB.CommandButton cmLogins 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   9
      Top             =   390
      Width           =   675
   End
   Begin VB.CommandButton cmGlobal 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   3
      Top             =   90
      Width           =   675
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   2100
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   1860
      Width           =   8475
      _ExtentX        =   14944
      _ExtentY        =   3514
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label laInform 
      Height          =   315
      Left            =   4680
      TabIndex        =   18
      Top             =   4020
      Width           =   2415
   End
   Begin VB.Label laGrid 
      Caption         =   "Список доступных баз"
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label laProducts 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1290
      Width           =   5535
   End
   Begin VB.Label laNomenks 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   990
      Width           =   5535
   End
   Begin VB.Label laSvodka 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   690
      Width           =   5535
   End
   Begin VB.Label laLogins 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   390
      Width           =   5535
   End
   Begin VB.Label Label6 
      Caption         =   "Файл составных изделий:"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Файл простых изделий:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1020
      Width           =   1875
   End
   Begin VB.Label Label4 
      Caption         =   "Файл Сводки:"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Файл логинов:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label laGlobal 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   90
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Глобальный конфиг. файл:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "cfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'не забыть выгрузку (см '$$$$)
Public gCfgFilePath As String
Public curBaseInd As Integer
'Public workBasePath As String
Public loginsPath As String
Public SvodkaPath As String
Public NomenksPath As String
Public ProductsPath As String
Public isLoad As Boolean
Private clickedRow As Integer

Public Regim As String
Dim key() As String
Dim val() As String
Dim glb() As Boolean

Const bsDbName = 1
Const bsServer = 2
Const bsActive = 3
Const bsPrefix = 4



'загружает параметры из cfg.файлов, если чего не хватает, то запрашивает их
'через диалог и тут же сохраняет
Function loadCfg() As Boolean
Dim str As String, I As Integer

loadCfg = False
ReDim key(0): ReDim val(0): ReDim glb(0)

loadParams "cfg" 'лок cfg.файл

gCfgFilePath = getParam("gCfgFilePath")
If gCfgFilePath <> "" Then ' путь к глоб.ф-лу был уже определен
    If Not loadParams(gCfgFilePath) Then
        MsgBox "Повторите запуск позже или сообщите Администратору!", , _
        "Не найден путь '" & gCfgFilePath & "'."
        End
    End If
Else
    gCfgFilePath = App.Path & "\" & "global.cfg"
'    loadParams gCfgFilePath далее присвоятся значения по умолчанию
End If
'loginsPath = getParamOrDefault("loginsPath", "файл логинов")
loginsPath = getParam("loginsPath")
If loginsPath = "" Then loginsPath = _
    "\\Server\C\WebServers\home\petmas.ru\mirror\files\logins."

'SvodkaPath = getParamOrDefault("SvodkaPath", "файл Сводки")
SvodkaPath = getParam("SvodkaPath")
If SvodkaPath = "" Then SvodkaPath = _
    "\\Server\C\WebServers\home\petmas.ru\mirror\files\svodkaW."

'NomenksPath = getParamOrDefault("NomenksPath", "файл простых изделий")
NomenksPath = getParam("NomenksPath")
If NomenksPath = "" Then NomenksPath = _
    "\\Server\C\WebServers\home\petmas.ru\mirror\files\Nomenks."

'ProductsPath = getParamOrDefault("ProductsPath", "файл составных изделий")
ProductsPath = getParam("ProductsPath")
If ProductsPath = "" Then ProductsPath = _
    "\\Server\C\WebServers\home\petmas.ru\mirror\files\Products."

DD:
    loadCfg = True
EE: 'запись параметров

End Function

Sub saveCfg(Optional onlyLocal As String = "")

    If Regim = "comtexAdmin" Then
        Exit Sub
    End If
    saveParams "cfg"
    If onlyLocal = "" Then saveParams gCfgFilePath
    
End Sub

Private Sub cmExit_Click()
    Unload Me
End Sub

Private Sub cmGlobal_Click()
Dim old As String

old = gCfgFilePath
cdOpen.DialogTitle = "Выберите Глобальный конфигурационный файл."
cdOpen.FileName = ""
cdOpen.ShowOpen
If cdOpen.FileName = "" Then Exit Sub
gCfgFilePath = cdOpen.FileName

If Dir$(gCfgFilePath) = "" Then
  saveCfg 'если файл НЕ существует то создаем его с параметрами по умолчанию
Else ' если есть, то создаем только с
  saveCfg "localOnly"
End If
loadCfg
setRegim ' новые значения на экран
If gCfgFilePath <> old Then changeMsg
On Error Resume Next
Grid.SetFocus
'laGlobal.Caption = cdOpen.FileName

End Sub

Private Sub cmLogins_Click()
cdOpen.DialogTitle = "Выберите файл логинов."
cdOpen.FileName = ""
cdOpen.ShowOpen
If cdOpen.FileName = "" Then Exit Sub
loginsPath = cdOpen.FileName
saveCfg
laLogins.Caption = cdOpen.FileName

End Sub

Private Sub cmNomenks_Click()
cdOpen.DialogTitle = "Выберите файл простых изделий."
cdOpen.FileName = ""
cdOpen.ShowOpen
If cdOpen.FileName = "" Then Exit Sub
NomenksPath = cdOpen.FileName
saveCfg
laNomenks.Caption = cdOpen.FileName
End Sub

Private Sub cmProducts_Click()
cdOpen.DialogTitle = "Выберите файл составных изделий."
cdOpen.FileName = ""
cdOpen.ShowOpen
If cdOpen.FileName = "" Then Exit Sub
ProductsPath = cdOpen.FileName
saveCfg
laProducts.Caption = cdOpen.FileName
End Sub

Private Sub cmSvodka_Click()
cdOpen.DialogTitle = "Выберите файл Сводки."
cdOpen.FileName = ""
cdOpen.ShowOpen
If cdOpen.FileName = "" Then Exit Sub
SvodkaPath = cdOpen.FileName
saveCfg
laSvodka.Caption = cdOpen.FileName
End Sub

Private Sub Form_Load()
isLoad = True
End Sub

Sub setRegim()
Dim I As Integer

If Regim = "comtexAdmin" Then
    Grid.Rows = 2: Grid.Cols = 5: Grid.Clear
    Grid.FormatString = "|<Бухгалтерская база|<Cервер|<Совместная работа|<Префикс"
    MsgBox "Будьте уверены, что вы знаете, что вы делаете. В противном случае изменения сделанные в открывающеммся окне могут повлечь за собой проблемы в режиме совместной работы Prior и Comtex", , "Предупреждение"
Else
    Grid.Rows = 2: Grid.Cols = 2: Grid.Clear
    Grid.FormatString = "|<Усл.название|<Полный путь к файлу|Рабочая|Текущая"
End If

If Regim = "pathSet" Then
    Me.Caption = "Установка путей"
    Grid.ColWidth(0) = 0
    laGlobal.Caption = gCfgFilePath
    laLogins.Caption = loginsPath
    laSvodka.Caption = SvodkaPath
    laNomenks.Caption = NomenksPath
    laProducts.Caption = ProductsPath
ElseIf Regim = "comtexAdmin" Then
    Me.Caption = "Выбор базы"
    laGlobal.Visible = False
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
    Grid.removeItem Grid.Rows - 1
End If
  

Grid_EnterCell
End Sub

Public Sub setParam(paramKey As String, paramVal, Optional p_glb As Boolean = False)
Dim I As Integer
    
    For I = 1 To UBound(key)
        If paramKey = key(I) Then GoTo AA
    Next I
    
    I = UBound(key) + 1
    ReDim Preserve key(I): ReDim Preserve val(I): ReDim Preserve glb(I)
    key(I) = paramKey
    glb(I) = p_glb
AA:
    val(I) = paramVal
End Sub

Public Function getParam(paramKey As String) As String
    Dim I As Integer
    
    For I = 1 To UBound(key)
        If paramKey = key(I) Then
            getParam = val(I)
            Exit Function
        End If
    Next I
    getParam = ""
End Function

Sub baseOpen(Optional baseIndex As Integer = -1)
Dim str As String, dburl As String
    dburl = getParam("dbUrl")
    
    If otlad = "otlaD" Then
'        dburl = "dev_prior"
        mainTitle = "    otlad"
    Else
'        dburl = "prior"
        mainTitle = "Prior, build [324] от 18.04.2008"
    End If
    
    Set myBase = wrkDefault.OpenDatabase("Connection1", _
       dbDriverNoPrompt, False, _
       "ODBC;UID=dba;PWD=sql;DSN=" & dburl)
    If myBase Is Nothing Then End
    
    sql = "call bootstrap_blocking()"
    If myExecute("##bootstrap", sql, 0) = 0 Then End
    
    Exit Sub
    
ERRb:
       
    If errorCodAndMsg("388", -100) Then '##388
        MsgBox "Не обнаружен Сервер базы", , "Сообщите Администратору!"
        End
    End If
    
    End
    
End Sub


Sub saveParams(filePath As String)
Dim I As Integer, str  As String
Dim doSave As Boolean

    If filePath = "cfg" Then
        str = App.Path & "\" & App.EXEName & ".cfg"
    Else
        str = filePath
    End If
    On Error GoTo EN1
    Open str For Output As #1
    For I = 1 To UBound(key)
        If filePath = "cfg" And Not glb(I) Then
            doSave = True
        Else
            doSave = False
        End If
        If doSave Then
            Print #1, key(I) & " = " & val(I)
        End If
    Next I
EN1:
    On Error Resume Next
    Close #1
End Sub

Function loadParams(filePath As String) As Boolean
    Dim str As String, str2 As String, I As Integer, j As Integer
    If filePath = "cfg" Then
        str = App.Path & "\" & App.EXEName & ".cfg"
    Else
        str = filePath
    End If
    
    On Error GoTo EN1 'если сетевая папка недоступна, то Dir дает ERR
    If Dir$(str) = "" Then
        loadParams = False
    Else
      Open str For Input As #1
      While Not EOF(1)
        Line Input #1, str
        I = InStr(str, "=")
        If I > 0 Then
            str2 = myTrim(Mid$(str, I + 1))
            str = myTrim(Left$(str, I - 1))
            I = UBound(key) + 1
            ReDim Preserve key(I): ReDim Preserve val(I): ReDim Preserve glb(I)
            key(I) = str
            val(I) = str2
            If filePath = "cfg" Then
                glb(I) = False
            Else
                glb(I) = True
            End If
        End If
      Wend
      Close #1
      
      loadParams = True
    End If
    Exit Function
EN1:
    loadParams = False
End Function



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
        saveCfg
        lbHide
    ElseIf KeyCode = vbKeyEscape Then
        lbHide
    End If

End Sub
'кроме нач и кон пробелов удаляет и vbTab
Function myTrim(str As String) As String
    Dim I As Integer, ch As String ', lPoz As Integer, rPoz As Integer
    
    For I = 1 To Len(str)
        ch = Mid$(str, I, 1)
        If ch <> " " And ch <> vbTab Then GoTo AA
    Next I
    myTrim = ""
    Exit Function
AA:
    str = Mid$(str, I)
    For I = Len(str) To 1 Step -1
        ch = Mid$(str, I, 1)
        If ch <> " " And ch <> vbTab Then Exit For
    Next I
    myTrim = Left$(str, I)
    
End Function

