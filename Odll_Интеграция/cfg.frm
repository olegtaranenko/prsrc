VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
      TabIndex        =   21
      Top             =   4020
      Width           =   795
   End
   Begin VB.ListBox lbActive 
      Height          =   432
      ItemData        =   "cfg.frx":0000
      Left            =   1500
      List            =   "cfg.frx":000A
      TabIndex        =   22
      Top             =   1020
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   3240
      TabIndex        =   19
      Text            =   "tbMobile"
      Top             =   2340
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmProducts 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   17
      Top             =   1290
      Width           =   675
   End
   Begin VB.CommandButton cmNomenks 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   15
      Top             =   990
      Width           =   675
   End
   Begin VB.CommandButton cmSvodka 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   13
      Top             =   690
      Width           =   675
   End
   Begin VB.CommandButton cmLogins 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   11
      Top             =   390
      Width           =   675
   End
   Begin VB.CommandButton cmGlobal 
      Caption         =   "Выбор"
      Height          =   255
      Left            =   7860
      TabIndex        =   5
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
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   4020
      Width           =   855
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   4020
      Width           =   915
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
      TabIndex        =   20
      Top             =   4020
      Width           =   2415
   End
   Begin VB.Label laGrid 
      Caption         =   "Список доступных баз"
      Height          =   255
      Left            =   180
      TabIndex        =   18
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label laProducts 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   1290
      Width           =   5535
   End
   Begin VB.Label laNomenks 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   990
      Width           =   5535
   End
   Begin VB.Label laSvodka 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   690
      Width           =   5535
   End
   Begin VB.Label laLogins 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   390
      Width           =   5535
   End
   Begin VB.Label Label6 
      Caption         =   "Файл составных изделий:"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Файл простых изделий:"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   1020
      Width           =   1875
   End
   Begin VB.Label Label4 
      Caption         =   "Файл Сводки:"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Файл логинов:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label laGlobal 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   90
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Глобальный конфиг. файл:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
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

Const bsName = 1
Const bsPath = 2
Const bsWork = 3
Const bsCurr = 4

Const bsDbName = 1
Const bsServer = 2
Const bsActive = 3
Const bsPrefix = 4



'загружает параметры из cfg.файлов, если чего не хватает, то запрашивает их
'через диалог и тут же сохраняет
Function loadCfg() As Boolean
Dim str As String, curBaseName As String, I As Integer

loadCfg = False
ReDim key(0): ReDim val(0)

loadParams "cfg" 'лок cfg.файл
curBaseName = getParam("curBaseName")
'gCfgFilePath = getParamOrDefault("gCfgFilePath", "Глобальный конфигурационный файл")
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

If curBaseName <> "" Then ' ищем тек.базу
    For I = 0 To UBound(base)
        If base(I) = curBaseName Then
            curBaseInd = I
            GoTo DD 'на запись параметров
        End If
    Next I
End If
'при первом запуске по кр.мере д.б. указана рабочая база - остальные потом в меню
If base(0) = "" Or basePath(0) = "" Then
  basePath(0) = App.Path & "\" & "empty.mdb"
  base(0) = "Пустая"
End If

'curBaseName = base(0)
curBaseInd = 0
DD:
loadCfg = True
EE: 'запись параметров
saveParams gCfgFilePath

ReDim key(0): ReDim val(0)
setParam "curBaseName", base(curBaseInd)
setParam "gCfgFilePath", gCfgFilePath
saveParams "cfg"

End Function

Sub saveCfg(Optional only As String = "")

If Regim = "comtexAdmin" Then
    Exit Sub
End If

ReDim key(0): ReDim val(0)
setParam "curBaseName", base(curBaseInd)
setParam "gCfgFilePath", gCfgFilePath
saveParams "cfg"

If only <> "" Then Exit Sub

ReDim key(0): ReDim val(0)
setParam "loginsPath", loginsPath
setParam "SvodkaPath", SvodkaPath
setParam "NomenksPath", NomenksPath
setParam "ProductsPath", ProductsPath
saveParams gCfgFilePath
End Sub

Private Sub cmAdd_Click()
Dim I As Integer
cdOpen.DialogTitle = "Выберите файл базы."
cdOpen.FileName = ""
'cdOpen.Filter = "(*.hex) | *.hex"
cdOpen.ShowOpen
If cdOpen.FileName <> "" Then
    If Grid.TextMatrix(1, 1) <> "" Then Grid.AddItem ""
    I = UBound(base) + 1
    ReDim Preserve base(I): base(I) = ""
    ReDim Preserve basePath(I): basePath(I) = cdOpen.FileName
    saveCfg
    Grid.TextMatrix(Grid.Rows - 1, 0) = I
    Grid.TextMatrix(Grid.Rows - 1, bsPath) = cdOpen.FileName
    Grid.row = Grid.Rows - 1
    Grid.col = bsName
End If

On Error Resume Next
Grid.SetFocus

End Sub

Private Sub cmEdit_Click()

End Sub

Private Sub cmDel_Click()
If Grid.TextMatrix(Grid.row, bsWork) <> "" Or Grid.TextMatrix(Grid.row, bsCurr) <> "" Then
    MsgBox "Нельзя удалять ссылку на рабочую или текущую базу.", , "Предупрежедние"
Else
    base(Grid.TextMatrix(Grid.row, 0)) = ""
    saveCfg
    Grid.removeItem Grid.row
    Grid.row = 1
    Grid_EnterCell
End If
On Error Resume Next

Grid.SetFocus
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
    Grid.ColWidth(bsPath) = 5715
'    Grid.ColWidth(bsName) = 1000
    Grid.ColWidth(bsWork) = 540
    Grid.ColWidth(bsCurr) = 585
'Public curBaseInd As Integer
'Public workBasePath As String
    laGlobal.Caption = gCfgFilePath
    laLogins.Caption = loginsPath
    laSvodka.Caption = SvodkaPath
    laNomenks.Caption = NomenksPath
    laProducts.Caption = ProductsPath
    If base(0) <> "" Then
      For I = 0 To UBound(base)
        Grid.TextMatrix(I + 1, 0) = I
        Grid.TextMatrix(I + 1, bsName) = base(I)
        Grid.TextMatrix(I + 1, bsPath) = basePath(I)
        If I = 0 Then Grid.TextMatrix(I + 1, bsWork) = "Да"
        If I = curBaseInd Then Grid.TextMatrix(I + 1, bsCurr) = "Да"
        Grid.AddItem ""
      Next I
      Grid.removeItem Grid.Rows - 1
    End If
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
ElseIf Regim = "baseChoise" Then
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
    Grid.ColWidth(bsPath) = 0
    Grid.ColWidth(bsName) = 1400
    Grid.ColWidth(bsWork) = 0
    Grid.ColWidth(bsCurr) = 765
    I = Grid.Width
    Grid.Width = Grid.ColWidth(bsName) + Grid.ColWidth(bsCurr) + 350
    I = Grid.Width - I
    Me.Width = Me.Width + I
    cmExit.Left = cmExit.Left + I
    cmExit.Top = Grid.Top + Grid.Height + 100
    For I = 0 To UBound(base)
        Grid.TextMatrix(I + 1, 0) = I
        Grid.TextMatrix(I + 1, bsName) = base(I)
        If I = curBaseInd Then Grid.TextMatrix(I + 1, bsCurr) = "Да"
        Grid.AddItem ""
    Next I
    Grid.removeItem Grid.Rows - 1
End If
  

Grid_EnterCell
End Sub
Sub setParam(paramKey As String, paramVal)
Dim I As Integer

For I = 1 To UBound(key)
    If paramKey = key(I) Then GoTo AA
Next I
I = UBound(key) + 1
ReDim Preserve key(I)
ReDim Preserve val(I)
key(I) = paramKey
AA:
val(I) = paramVal
End Sub

Function getParam(paramKey As String) As String
Dim I As Integer

For I = 1 To UBound(key)
    If paramKey = key(I) Then
        getParam = val(I)
        Exit Function
    End If
Next I
getParam = ""
End Function
'$odbc15!$
Sub baseOpen(Optional baseIndex As Integer = -1)
Dim str As String

On Error GoTo ERRb
RETR:
If baseIndex < 0 Then
    str = "C:\VB_DIMA\dlsricN.mdb"
Else
    str = basePath(baseIndex)
End If
'Set myBase = OpenDatabase(str, False, False, ";PWD=play")

'Set wrkDefault = DBEngine.CreateWorkspace("wrkDefault", "dba", "sql", dbUseODBC) ' для орг-ии транзакций

'On Error GoTo ERRb
If otlad = "otlaD" Then
   Set myBase = wrkDefault.OpenDatabase("Connection1", _
      dbDriverNoPrompt, False, _
      "ODBC;UID=dba;PWD=sql;DSN=prior")
      mainTitle = "    otlad"
Else
   Set myBase = wrkDefault.OpenDatabase("Connection1", _
      dbDriverNoPrompt, False, _
      "ODBC;UID=dba;PWD=sql;DSN=prior")
      mainTitle = "    New"
End If
If myBase Is Nothing Then End
Exit Sub

ERRb:
   
If errorCodAndMsg("388", -100) Then '##388
    MsgBox "Не обнаружен Сервер базы", , "Сообщите Администратору!"
    End
End If

sql = "call bootstrap_blocking"
If myExecute("##0.1", sql, 0) = 0 Then End

'   Dim strError As String
'   Dim errLoop
'   For Each errLoop In Errors
'      With errLoop
'         strError = _
'            "Error  : '" & .Number & "'" & vbCr
'         strError = strError & _
'            "   " & .Description & vbCr
'         strError = strError & _
'            "(Source:   " & .Source & ")"
'      End With
'      MsgBox strError
'   Next


'str = "Не удалось подключиться к базе '" & str & "':"

End

End Sub

Function getParamOrDefault(paramKey As String, defPath As String) As String

getParamOrDefault = getParam(paramKey)
If getParamOrDefault = "" Then
    
End If


End Function

Sub saveParams(filePath As String)
Dim I As Integer, str  As String

If filePath = "cfg" Then
    str = App.Path & "\" & App.EXEName & ".cfg"
Else
    str = filePath
End If
On Error GoTo EN1
Open str For Output As #1
For I = 1 To UBound(key)
    Print #1, key(I) & " = " & val(I)
Next I
If filePath <> "cfg" Then 'только для глобального
  If Trim(base(0)) <> "" Then Print #1, "workBase_" & base(0) & " = " & basePath(0)
  For I = 1 To UBound(base)
    If base(I) <> "" Then ' т.к. могли удалить в cfg.форме
        Print #1, "base_" & base(I) & " = " & basePath(I)
    End If
  Next I
End If
EN1:
On Error Resume Next
Close #1
End Sub

Function loadParams(filePath As String) As Boolean
Dim str As String, str2 As String, I As Integer, j As Integer  ', ind As Integer  ', key As String, val As String

ReDim key(0): ReDim val(0):
ReDim base(0): ReDim basePath(0): base(0) = ""
If filePath = "cfg" Then
    str = App.Path & "\" & App.EXEName & ".cfg"
Else
    str = filePath
End If
'str = "C:\aa\" поиск папки
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
        If InStr(LCase(str), "base_") > 0 Then '----------------------------!!
            If InStr(LCase(str), "workbase_") > 0 Then '=============!
                base(0) = Mid$(str, 10) 'len("workBbase_")+1=====!
                basePath(0) = str2
            Else
                I = UBound(base) + 1
                ReDim Preserve base(I)
                ReDim Preserve basePath(I)
                base(I) = Mid$(str, 6) 'len("base_")+1 --------------!!
                basePath(I) = str2
            End If
        Else
            I = UBound(key) + 1
            ReDim Preserve key(I): ReDim Preserve val(I)
            key(I) = str
            val(I) = str2
'        laInform.Caption = laInform.Caption & vbCrLf & "'" & _
        key & "'   '" & val & "'"
        End If
    End If
  Wend
  Close #1
  I = UBound(base)
  If base(0) = "" And I > 0 Then 'если нет рабочей базы то ею будет последняя
        base(0) = base(I)
        basePath(0) = basePath(I)
        ReDim Preserve base(I - 1)
        ReDim Preserve basePath(I - 1)
  End If
  
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

'If Grid.MouseRow = 0 Then Exit Sub так нельзя
If Regim = "comtexAdmin" Then
    If Grid.col = bsActive Then
        listBoxInGridCell lbActive, Grid, Grid.TextMatrix(Grid.MouseRow, Grid.MouseCol)
    ElseIf Grid.col = bsPrefix Then
        textBoxInGridCell tbMobile, Grid
    End If
    
    Exit Sub
End If

If Regim = "baseChoise" Then
    If Grid.col < 2 Then Exit Sub
    GoTo BB
End If
If Grid.col = bsName Then
    textBoxInGridCell tbMobile, Grid
ElseIf Grid.col = bsPath Then
    cdOpen.DialogTitle = "Выбор новой базы."
    cdOpen.FileName = ""
    cdOpen.ShowOpen
    If cdOpen.FileName = "" Then Exit Sub
    If Grid.TextMatrix(Grid.row, bsCurr) <> "" Then changeMsg
    basePath(Grid.TextMatrix(Grid.row, 0)) = cdOpen.FileName
    saveCfg
    Grid.Text = cdOpen.FileName
ElseIf Grid.col = bsWork Then
    If Trim(Grid.TextMatrix(Grid.row, bsName)) = "" Then GoTo AA
    'текущ.строка(база) и та, где был индекс раб.базы обмениваются индексами
    ind = Grid.TextMatrix(Grid.row, 0)
    str = base(ind)
    base(ind) = base(0)
    base(0) = str
    str = basePath(ind)
    basePath(ind) = basePath(0)
    basePath(0) = str
    For I = 1 To Grid.Rows - 1
        Grid.TextMatrix(I, bsWork) = ""
        If Grid.TextMatrix(I, 0) = 0 Then 'там где был индекс раб.базы
            Grid.TextMatrix(I, 0) = ind
            If Grid.TextMatrix(I, bsCurr) <> "" Then _
                curBaseInd = ind 'если она еще была и текущей
        End If
    Next I
    If Grid.TextMatrix(Grid.row, bsCurr) <> "" Then curBaseInd = 0 'если текущ. строка была текущей базой
    saveCfg
    Grid.Text = "Да"
    Grid.TextMatrix(Grid.row, 0) = 0 'в тек строку - индекс раб.базы
Else ' bsCurr
    If Trim(Grid.TextMatrix(Grid.row, bsName)) = "" Then
AA:     MsgBox "Сначала введите название базы.", , ""
        Grid.col = bsName
        On Error Resume Next
        Grid.SetFocus
    Else
BB:     curBaseInd = Grid.TextMatrix(Grid.row, 0)
        saveCfg
        If Grid.Text <> "Да" Then changeMsg
        For I = 1 To Grid.Rows - 1
            Grid.TextMatrix(I, bsCurr) = ""
        Next I
        Grid.Text = "Да"
    End If
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

If Regim = "baseChoise" And Grid.col < 2 _
    Or Regim = "comtexAdmin" And Grid.col <= 2 _
Then
    Grid.CellBackColor = vbYellow
Else
    Grid.CellBackColor = &H88FF88
End If
' laInform.Caption = "row=" & Grid.Row & "  col=" & Grid.Col
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
'Debug.Print sql
    
'    If Not myExecute(0, sql) Then
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
    Else
        base(Grid.TextMatrix(Grid.row, 0)) = tbMobile.Text
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

