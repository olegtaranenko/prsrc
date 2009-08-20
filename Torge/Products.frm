VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Products 
   BackColor       =   &H8000000A&
   Caption         =   "Справочник готовых изделий"
   ClientHeight    =   6396
   ClientLeft      =   60
   ClientTop       =   1740
   ClientWidth     =   11880
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6396
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmSostavExcel 
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   9240
      TabIndex        =   42
      Top             =   5940
      Width           =   1335
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   10860
      TabIndex        =   8
      Top             =   5940
      Width           =   795
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   10720
      TabIndex        =   19
      Top             =   300
      Width           =   1095
      Begin VB.CommandButton cmCancel 
         Caption         =   "Отменить"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   41
         Top             =   5100
         Width           =   1035
      End
      Begin VB.CommandButton cmApple 
         Caption         =   "Применить"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   40
         Top             =   4620
         Width           =   1035
      End
      Begin VB.TextBox tbGain 
         Height          =   285
         Index           =   2
         Left            =   540
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   4200
         Width           =   495
      End
      Begin VB.TextBox tbGain 
         Height          =   285
         Index           =   1
         Left            =   540
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   3060
         Width           =   495
      End
      Begin VB.TextBox tbGain 
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox tbCol 
         Height          =   285
         Index           =   3
         Left            =   540
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   3900
         Width           =   495
      End
      Begin VB.TextBox tbCol 
         Height          =   285
         Index           =   2
         Left            =   540
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox tbCol 
         Height          =   285
         Index           =   1
         Left            =   540
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1620
         Width           =   495
      End
      Begin VB.TextBox tbCol 
         Height          =   285
         Index           =   0
         Left            =   540
         MaxLength       =   10
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   285
         Left            =   540
         TabIndex        =   33
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Коэф."
         Height          =   195
         Left            =   0
         TabIndex        =   32
         Top             =   4260
         Width           =   435
      End
      Begin VB.Label Label12 
         Caption         =   "Кол-во"
         Height          =   195
         Left            =   0
         TabIndex        =   31
         Top             =   3960
         Width           =   555
      End
      Begin VB.Label Label11 
         Caption         =   "Колонка4"
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   3660
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Коэф."
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   3120
         Width           =   435
      End
      Begin VB.Label Label9 
         Caption         =   "Кол-во"
         Height          =   195
         Left            =   0
         TabIndex        =   27
         Top             =   2820
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "Колонка3"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Коэф."
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label Label6 
         Caption         =   "Кол-во"
         Height          =   315
         Left            =   0
         TabIndex        =   24
         Top             =   1620
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "Колонка2"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Коэф."
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Кол-во"
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Колонка1"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.ListBox lbPrWeb 
      Height          =   432
      ItemData        =   "Products.frx":0000
      Left            =   3900
      List            =   "Products.frx":000A
      TabIndex        =   18
      Top             =   1740
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.ListBox lbWeb 
      Height          =   432
      ItemData        =   "Products.frx":0015
      Left            =   7560
      List            =   "Products.frx":001F
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame frTitle 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   5100
      TabIndex        =   15
      Top             =   3300
      Visible         =   0   'False
      Width           =   5475
      Begin VB.Label laTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   "laTitle"
         Height          =   255
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmExcel 
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   4800
      TabIndex        =   14
      Top             =   5940
      Width           =   1335
   End
   Begin VB.CommandButton cmNomenk 
      Caption         =   "Справочник номенклатуры"
      Height          =   315
      Left            =   180
      TabIndex        =   13
      Top             =   5940
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox tbQuant 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8220
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "1"
      Top             =   5940
      Width           =   375
   End
   Begin VB.CommandButton cmSel 
      Caption         =   "Выбрать"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7260
      TabIndex        =   9
      Top             =   5940
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Left            =   3360
      Top             =   5760
   End
   Begin VB.Timer Timer1 
      Left            =   2940
      Top             =   5820
   End
   Begin VB.TextBox tbMobile2 
      Height          =   315
      Left            =   7920
      TabIndex        =   6
      Text            =   "tbMobile2"
      Top             =   4260
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Text            =   "tbMobile"
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   2580
      TabIndex        =   1
      Top             =   240
      Width           =   3675
      _ExtentX        =   6477
      _ExtentY        =   9758
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      _ExtentX        =   4255
      _ExtentY        =   9758
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   5535
      Left            =   6360
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7641
      _ExtentY        =   9758
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Label laBegin 
      Caption         =   "Label2"
      Height          =   3615
      Left            =   2760
      TabIndex        =   12
      Top             =   2100
      Width           =   3435
   End
   Begin VB.Label laQuant 
      Caption         =   "комплектов"
      Enabled         =   0   'False
      Height          =   195
      Left            =   8640
      TabIndex        =   11
      Top             =   5980
      Width           =   915
   End
   Begin VB.Label laNomenk 
      Caption         =   " "
      Height          =   195
      Left            =   7200
      TabIndex        =   7
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label laProduct 
      Height          =   195
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Список Серий"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2355
   End
   Begin VB.Menu mnContext 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu mnAdd 
         Caption         =   "Добавить"
      End
      Begin VB.Menu mnRen 
         Caption         =   "Преименовать"
         Visible         =   0   'False
      End
      Begin VB.Menu mnDel 
         Caption         =   "Удалить"
         Visible         =   0   'False
      End
      Begin VB.Menu mnSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnRepl 
         Caption         =   "Переместить"
         Visible         =   0   'False
      End
      Begin VB.Menu mnCancel 
         Caption         =   "Отменить"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnContext2 
      Caption         =   "Context2"
      Visible         =   0   'False
      Begin VB.Menu mnAdd2 
         Caption         =   "Добавить"
      End
      Begin VB.Menu mnCopy2 
         Caption         =   "Добавить по обр."
      End
      Begin VB.Menu mnDel2 
         Caption         =   "Удалить"
      End
      Begin VB.Menu mnSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnRepl2 
         Caption         =   "Переместить"
      End
   End
   Begin VB.Menu mnContext3 
      Caption         =   "Перемещение гот.изд"
      Visible         =   0   'False
      Begin VB.Menu mnInsert 
         Caption         =   "Вставить"
      End
      Begin VB.Menu mnCancel3 
         Caption         =   "Отменить"
      End
   End
   Begin VB.Menu mnContext4 
      Caption         =   "Отмена перем-я"
      Visible         =   0   'False
      Begin VB.Menu mnCancel4 
         Caption         =   "Отмена"
      End
   End
   Begin VB.Menu mnContext5 
      Caption         =   "Добавить удалить номенк-ру"
      Visible         =   0   'False
      Begin VB.Menu mnAdd5 
         Caption         =   "Добавить"
      End
      Begin VB.Menu mnEdit5 
         Caption         =   "Изменить"
      End
      Begin VB.Menu mnDel5 
         Caption         =   "Удалить"
      End
   End
End
Attribute VB_Name = "Products"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Const OTLADproduct = "" '"Шарики" ' отладка Прайса
Const msgQuant = 10 ' мах число Err сообщений, после кот. программа выходит


Public isLoad As Boolean
Public Regim As String
Public mousCol2 As Long
Public mousRow2 As Long
Public SumCenaFreight As String  ', VremObr As Single
Public SumCenaSale As String
Public soursNom As String

Dim mousCol As Long, mousRow As Long
Dim Node As Node
Dim tbSeries As Recordset
Dim quantity  As Long
Dim quantity2  As Long
Dim frmMode As String
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim sumGridsWidth As Integer ' суммарная ширина Grid и Grid2
Dim wesGrid As Single ' относительный размер Grid
Const between = 50 'помежность между Grid_ами

Const gpNomenk = 0
Const gpId = 1 '*************** должны совпадать
Const gpSortNom = 2
Const gpName = 3
Const gpPrWeb = 4
Const gpRabbat = 5
Const gpDescript = 6
Const gpSize = 7
Const gpVremObr = 8
Const gpSumCenaFreight = 9
Const gpSumCenaSale = 10
Const gpFormulaNom = 11
Const gpFormula = 12 'скрыт
Const gpCena3 = 13
Const gpCol1 = 14
Const gpCol2 = 15
Const gpCol3 = 16
Const gpCol4 = 17
Const gpPage = 18 ' страница в прайсе
Const gpUsed = 19

''для прайса
'Const prHideName = 0
'Const prId = 1 '*************** должны совпадать
'Const prName = 2
'Const prDescript = 3
'Const prCena3 = 4
'Const prCena4 = 5

Const gpNomNom = 1
Const gpWeb = 2
Const gpNomName = 3
Const gpCenaFreight = 4
Const gpCENA_W = 5
Const gpQuant = 6
Const gpEdIzm = 7
Const gpGroup = 8


Private Sub cmApple_Click()

If Not isNumericTbox(tbGain(0), 0.01) Then Exit Sub
If Not isNumericTbox(tbGain(1), 0.01) Then Exit Sub
If Not isNumericTbox(tbGain(2), 0.01) Then Exit Sub

sql = "UPDATE sGuideSeries SET head1 = '" & tbCol(0).Text & "', head2 = '" & _
tbCol(1).Text & "', head3 = '" & tbCol(2).Text & "', head4 = '" & _
tbCol(3).Text & "', gain2 = " & tbGain(0).Text & ", gain3 = " & _
tbGain(1).Text & ", gain4 = " & tbGain(2).Text & _
" WHERE (((seriaId)=" & gSeriaId & "));"
'MsgBox sql
If myExecute("##418", sql) = 0 Then
    loadSeriaProduct
    cmApple.Enabled = False
    cmCancel.Enabled = False
Else
    cmCancel_Click
End If

End Sub

Private Sub cmCancel_Click()
If getGainAndHead Then
    tbCol(0).Text = head1
    tbCol(1).Text = head2
    tbCol(2).Text = head3
    tbCol(3).Text = head4
    tbGain(0).Text = gain2
    tbGain(1).Text = gain3
    tbGain(2).Text = gain4
End If
cmApple.Enabled = False
cmCancel.Enabled = False
End Sub

Private Sub cmExcel_Click()
    GridToExcel Grid, laProduct.Caption
End Sub

Private Sub cmExit_Click()
'If Grid2.Visible Then
If Regim <> "" Then
    Unload Me
ElseIf checkRowsByCol Then
    Unload Me
Else
    controlGridsWidth
    Grid2.Visible = True
    On Error Resume Next
    Grid2.SetFocus
End If
End Sub

Private Sub cmNomenk_Click()
Nomenklatura.Regim = "" 'new
Nomenklatura.setRegim
Nomenklatura.Show vbModal           '
'loadProductNomenk                   '
'Grid2.row = max(quantity2, 1)       '
'Grid2.SetFocus                      '

End Sub

Sub setGridsWidth()
    If Not Grid2.Visible Then
        Grid.Width = sumGridsWidth + between ' + помежность между Grid
    Else
        Grid.Width = sumGridsWidth * wesGrid
        Grid2.Width = sumGridsWidth - Grid.Width
        Grid2.left = Grid.left + Grid.Width + between ' + помежность между Grid
    End If

End Sub

'reg="left"     - все занимает Grid
'иначе(reg="")  - Grid соизмерима с Grid2
Sub controlGridsWidth(Optional reg As String = "")
'Static oldReg As String

'If oldWidth = 0 Then ' только один раз в сам. начале
''    oldReg = "###"
'    sumGridsWidth = Grid2.Left - Grid.Left + Grid2.Width - between '
'    wesGrid = Grid.Width / sumGridsWidth
'End If

Grid.MergeCells = flexMergeNever
If reg = "left" Then
    Grid2.Visible = False
    Grid.colWidth(gpCol1) = 700
    Grid.colWidth(gpCol2) = 700
    Grid.colWidth(gpCol3) = 700
    Grid.colWidth(gpCol4) = 700
    Grid.colWidth(gpPage) = 405
    Grid.colWidth(gpVremObr) = 630
    Grid.colWidth(gpFormulaNom) = 420
Else
    Grid2.Visible = True
    Grid.colWidth(gpCol1) = 0
    Grid.colWidth(gpCol2) = 0
    Grid.colWidth(gpCol3) = 0
    Grid.colWidth(gpCol4) = 0
    Grid.colWidth(gpPage) = 0
    Grid.colWidth(gpVremObr) = 0
    Grid.colWidth(gpFormulaNom) = 0
End If
setGridsWidth ' в завис-ти от Grid2.Visible
End Sub



Private Sub cmSel_Click()
Dim q As Single, I As Integer, str As String, n As Integer, rr As Integer

If Not isNumericTbox(tbQuant, 1) Then Exit Sub

n = tbQuant.Text
wrkDefault.BeginTrans


sql = "SELECT sProducts.nomNom, sProducts.quantity From sProducts " & _
"WHERE (((sProducts.ProductId)=" & gProductId & "));"
Set tbProduct = myOpenRecordSet("##139", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Sub
'If tbProduct.BOF Then Exit Sub
rr = 0
While Not tbProduct.EOF
    q = Round(tbProduct!quantity * n, 2)
    If q = 0 Then GoTo NXT
    rr = rr + 1: ReDim Preserve QQ(rr): ReDim Preserve NN(rr)
    QQ(rr) = -q: NN(rr) = tbProduct!nomnom ' если придется делать откат
    
    Set tbNomenk = myOpenRecordSet("##150", "sGuideNomenk", dbOpenTable)
    If tbNomenk Is Nothing Then GoTo EN1
    tbNomenk.index = "PrimaryKey"
    tbNomenk.Seek "=", tbProduct!nomnom
    If tbNomenk.NoMatch Then
        tbNomenk.Close
        GoTo EN1
    End If
    tbNomenk.Edit
    tbNomenk!nowOstatki = Round(tbNomenk!nowOstatki - q, 2)
    tbNomenk.Update
    tbNomenk.Close

    '!!! нужно бы лучше сначала проверить на предмет наличия, а потом уж добавлять
    ' а не наоборот
    str = Format(gDocDate, "yyyy-mm-dd") 'сначала пробуем добавить запись
    sql = "INSERT INTO sDMC ([xDate], nomNom, quantity, numDoc, numExt, lastM ) " & _
    "SELECT DateDiff('d',[System].[begOstatDate],'" & str & "'), '" & _
    tbProduct!nomnom & "', -" & q & ", " & numDoc & ", " & numExt & ", '" & _
    AUTO.cbM.Text & "' " & " From System;"
    'MsgBox sql
    I = myExecute("##151", sql, -196)
    If I = -2 Then 'если эта позиция уже есть, то обновляем сущ.запись
    
        Set tbDMC = myOpenRecordSet("##152", "sDMC", dbOpenTable)
        If tbDMC Is Nothing Then GoTo EN1
        tbDMC.index = "nomDoc"
        tbDMC.Seek "=", numDoc, numExt, tbProduct!nomnom
        If tbDMC.NoMatch Then
            tbDMC.Close
            GoTo EN1
        End If
        tbDMC.Edit
        q = Round(tbDMC!quantity - q, 2)
        tbDMC!quantity = q
        tbDMC!lastM = AUTO.cbM.Text
        tbDMC.Update
        tbDMC.Close
    
    ElseIf I <> 0 Then
EN1:    wrkDefault.Rollback
        ReDim QQ(0)
        GoTo EN2
    End If
NXT: tbProduct.MoveNext
Wend
wrkDefault.CommitTrans
EN2:
tbProduct.Close

'tbQuant.Text = "1"
Unload Me

End Sub


Private Sub cmSostavExcel_Click()
    GridToExcel Grid2, laNomenk.Caption
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim value As String, str As String

If Shift = vbCtrlMask And KeyCode = vbKeyF Then
    If mousCol = gpDescript Then
        str = "Описание"
    Else
        str = "Номер"
    End If
    value = InputBox("Укажите " & str & " или фрагмент.", _
    "Поиск готового изделия", value)
    If value = "" Then Exit Sub
    controlGridsWidth "left"
    loadSeriaProduct value
    On Error Resume Next
    Grid.SetFocus 'чтобы избежать vbKeyEnter(от InputBox) на tv
End If

End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width
'If oldWidth = 0 Then ' только один раз в сам. начале
''    oldReg = "###"
    sumGridsWidth = Grid2.left - Grid.left + Grid2.Width - between '
    wesGrid = Grid.Width / sumGridsWidth
'End If

initProdCategoryBox lbPrWeb

frmMode = ""
gridIsLoad = False
gSeriaId = 0 'необходим  для добавления класса
'If baseNamePath = "" Then ' стартуем с этой формы и создаем prodGuide.exe
'    Regim = "onlyGuide"
'End If
    
    Grid.FormatString = "см.Входящие|id|<Номер|<Код|web|Скид.|<Описание|<Размер|Время обработки" & _
    "|SumCenaFreight|SumCenaSale|№ формулы|Формула|Цена 3|кол1|кол2|кол3|кол4|Стр." _
    & "|"
    Grid.colWidth(gpNomenk) = 0 '300
    Grid.colWidth(gpId) = 0
    Grid.colWidth(gpUsed) = 0
    Grid.colWidth(gpSortNom) = 700
    Grid.colWidth(gpName) = 1065
    Grid.colWidth(gpDescript) = 3720
    Grid.colWidth(gpPrWeb) = 405
    Grid.colWidth(gpRabbat) = 405
    Grid.colWidth(gpFormula) = 0
    Grid.colWidth(gpSumCenaFreight) = 1275
    Grid.colWidth(gpSumCenaSale) = 1275
    Grid.colWidth(gpCena3) = 1275
    


'Grid2.FormatString = "|<Номер|Web|<Название|Ц.доставка|Ц.продажа|Кол-во|<Ед.измерения|<Группа"
Grid2.FormatString = "|<Номер|Web|<Название|CenaFreight|CenaSale|Кол-во|<Ед.измерения|<Группа"
Grid2.colWidth(0) = 0
Grid2.colWidth(gpNomNom) = 870
Grid2.colWidth(gpNomName) = 4005 '1800 '2370
Grid2.colWidth(gpEdIzm) = 435
Grid2.colWidth(gpCenaFreight) = 615
Grid2.colWidth(gpCENA_W) = 615
Grid2.colWidth(gpGroup) = 540


Grid.Visible = False
Frame1.Visible = False
cmSostavExcel.Visible = False

loadSeria
isLoad = True

laBegin = "В левом списке найдите (кликом Mouse) серию, к которой относится " & _
"искомое изделие, при этом откроется таблица, где будут все изделия этой " & _
"серии." & vbCrLf & vbCrLf & "При двойном клике Mouse по серым клеткам " & _
"таблицы откроется вторая таблица, где будет вся номенклатура, входящая " & _
"в соответствующее изделие." & vbCrLf & vbCrLf
If Regim = "select" Then
    cmSel.Visible = True
    tbQuant.Visible = True
    laQuant.Visible = True
    laBegin = laBegin & "Установите требуемое количество выбранного " & _
    "изделия (комплектов) и нажмите <Выбрать>."

Else
    laBegin = laBegin & "Контексные меню в списке " & _
   "и в таблицах вызываются правым кликом Mouse (но не на серых клетках)."
    cmSel.Visible = False
    tbQuant.Visible = False
    laQuant.Visible = False
End If
gridIsLoad = True
End Sub

Sub loadSeria()
Dim Key As String, pKey As String, k() As String, pK()  As String
Dim I As Integer, iErr As Integer
bilo = False
sql = "SELECT sGuideSeries.*  From sGuideSeries ORDER BY sGuideSeries.seriaId;"
Set tbSeries = myOpenRecordSet("##110", sql, dbOpenForwardOnly)
If tbSeries Is Nothing Then End
If Not tbSeries.BOF Then
 'Dim i As Integer
 'i = tbSeries.Fields("seriaName").Size
 tv.Nodes.Clear
 Set Node = tv.Nodes.Add(, , "k0", "Справочник по сериям")
 Node.Sorted = True
 
 ReDim k(0): ReDim pK(0): ReDim NN(0): iErr = 0
 While Not tbSeries.EOF
    If tbSeries!seriaId = 0 Then GoTo NXT1
    Key = "k" & tbSeries!seriaId
    pKey = "k" & tbSeries!parentSeriaId
    On Error GoTo ERR1 ' назначить второй проход
    Set Node = tv.Nodes.Add(pKey, tvwChild, Key, tbSeries!seriaName)
    On Error GoTo 0
    Node.Sorted = True
NXT1:
    tbSeries.MoveNext
 Wend
End If
tbSeries.Close

While bilo ' необходимы еще проходы
  bilo = False
  For I = 1 To UBound(k())
    If k(I) <> "" Then
        On Error GoTo ERR2 ' назначить еще проход
        Set Node = tv.Nodes.Add(pK(I), tvwChild, k(I), NN(I))
        On Error GoTo 0
        k(I) = ""
        Node.Sorted = True
    End If
NXT:
  Next I
Wend
tv.Nodes.Item("k0").Expanded = True
Exit Sub
ERR1:
 iErr = iErr + 1: bilo = True
 ReDim Preserve k(iErr): ReDim Preserve pK(iErr): ReDim Preserve NN(iErr)
 k(iErr) = Key: pK(iErr) = pKey: NN(iErr) = tbSeries!seriaName
 Resume Next

ERR2: bilo = True: Resume NXT
End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If WindowState = vbMinimized Then Exit Sub
On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width

sumGridsWidth = sumGridsWidth + w
setGridsWidth

Grid.Height = Grid.Height + h
'Grid.Width = Grid.Width + w / 2
Grid2.Height = Grid2.Height + h
'Grid2.Width = Grid2.Width + w / 2
'Grid2.Left = Grid2.Left + w / 2
tv.Height = tv.Height + h

cmSel.Top = cmSel.Top + h
cmSel.left = cmSel.left + w
tbQuant.Top = tbQuant.Top + h
tbQuant.left = tbQuant.left + w
laQuant.Top = laQuant.Top + h
laQuant.left = laQuant.left + w
cmExit.Top = cmExit.Top + h
cmExit.left = cmExit.left + w
cmExcel.Top = cmExcel.Top + h
cmExcel.left = Grid.left + Grid.Width - cmExcel.Width
Frame1.left = Frame1.left + w
cmSostavExcel.left = Grid2.left + Grid2.Width - cmSostavExcel.Width
cmSostavExcel.Top = cmSostavExcel.Top + h

End Sub

Private Sub Form_Unload(Cancel As Integer)
isLoad = False

End Sub

Private Sub Grid_Click()
Static prevRow As Long, trg As Boolean

mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If quantity = 0 Then Exit Sub

If Grid.MouseRow = 0 And mousCol <> 0 Then
    gridIsLoad = False
    Grid.CellBackColor = Grid.BackColor
    SortCol Grid, mousCol
    Grid.row = 1    ' только чтобы снять выделение
    Grid.col = gpName
    gridIsLoad = True
    Grid.CellBackColor = &H88FF88 'т.к. Grid_EnterCell здесь нельзя
End If

End Sub
Sub refrProductCenaToGrid()
Dim str As String
        str = productFormula
        Grid.TextMatrix(mousRow, gpCena3) = str
        Grid.TextMatrix(mousRow, gpSumCenaFreight) = Format(SumCenaFreight, "0.00") ' только после
        Grid.TextMatrix(mousRow, gpSumCenaSale) = SumCenaSale 'Format(SumCenaSale, "0.00")
        Grid.TextMatrix(mousRow, gpFormula) = tmpStr    ' productFormula
'        If Not IsNumeric(SumCenaFreight) Then
'            MsgBox SumCenaFreight, , ""
'        ElseIf Not IsNumeric(str) Then
'            MsgBox str, , ""
'        ElseIf Not IsNumeric(SumCenaSale) Then
'            MsgBox SumCenaSale, , ""
'        End If

End Sub

Private Sub Grid_DblClick()
Dim str As String

If mousRow = 0 Or mousCol = 0 Then Exit Sub

If mousCol = gpFormulaNom Then
    If GuideFormuls.isLoad Then Unload GuideFormuls
    GuideFormuls.Regim = "fromProduct"
    GuideFormuls.Show vbModal
    If tmpStr = "" Then Exit Sub
    If ValueToTableField("##312", "'" & tmpStr & "'", "sGuideProducts", _
    "formulaNom", "byProductId") Then
        Grid.TextMatrix(mousRow, gpFormulaNom) = tmpStr
         refrProductCenaToGrid
    End If
ElseIf mousCol = gpPrWeb Then
    listBoxInGridCell lbPrWeb, Grid, "select"
ElseIf Grid.CellBackColor = &H88FF88 Then
    If frmMode <> "price" Then frmMode = ""
    textBoxInGridCell tbMobile, Grid
End If

End Sub

Private Sub Grid_EnterCell()
Static prevCol As Long

If quantity = 0 Or Not gridIsLoad Then
    prevCol = 0
    Exit Sub
End If
mousRow = Grid.row
mousCol = Grid.col
If mousRow = 0 Then Exit Sub

If mousCol = gpSumCenaFreight Then GoTo YY

If mousCol = gpCena3 Then
    laTitle.Caption = "Цена3 = " & Grid.TextMatrix(mousRow, gpFormula) & " "
    frTitle.Top = Grid.CellTop + 2 * Grid.CellHeight + 50
    frTitle.Visible = True
    frTitle.ZOrder
YY: Grid.CellBackColor = vbYellow
    Exit Sub
Else
    frTitle.Visible = False
End If

'If frmMode = "" Then
    gProductId = Grid.TextMatrix(mousRow, gpId)
'Else
'    gProductId = Grid.TextMatrix(mousRow, prId)
'End If

cmSostavExcel.Visible = False
If mousCol = gpName Then
    If prevCol <> gpName Then controlGridsWidth ""
    loadProductNomenk
    cmSostavExcel.Visible = True
ElseIf prevCol = gpName Then
    controlGridsWidth "left"
End If
prevCol = mousCol
If (frmMode <> "") Or Regim = "select" Then Exit Sub

If mousCol = gpDescript Then
    tbMobile.MaxLength = 100
Else
    tbMobile.MaxLength = 30
End If
' tbInform.MaxLength =tbMobile.MaxLength

If mousCol = 0 Then Exit Sub

Grid.CellBackColor = &H88FF88

End Sub


Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Grid_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If


End Sub

Function checkRowsByCol() As Boolean
Dim il As Long, str As String
checkRowsByCol = True
If quantity2 = 0 Then Exit Function
For il = 1 To Grid2.Rows - 1
    str = Grid2.TextMatrix(il, gpQuant)
    If Not IsNumeric(str) Then Exit Function ' это возм. при закрытой Grid2
    If CSng(str) < 0.01 Then
BB:
'       Grid2.SetFocus
        Grid2.row = il
        Grid2.col = gpQuant
        MsgBox "Для оной из номенклатур, входящей в изделие, указано " & _
        "нулевое кол-во. Проставте количество либо удалите номенклатуру.", , _
        "Недопустимое значение!"
        checkRowsByCol = False
        Exit Function
    End If
Next il

End Function

Sub lbHide()
tbMobile.Visible = False
lbPrWeb.Visible = False
Grid.Enabled = True
On Error Resume Next
Grid.SetFocus
Grid_EnterCell
End Sub

Sub lbHide2()
tbMobile2.Visible = False
lbWeb.Visible = False
Grid2.Enabled = True
On Error Resume Next
Grid2.SetFocus
Grid2_EnterCell
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyEscape Then Grid.CellBackColor = Grid.BackColor
If KeyCode = vbKeyEscape Then Grid_EnterCell
End Sub

Private Sub Grid_LeaveCell()
If gridIsLoad Then
    prevRow = Grid.row
    If Grid.col <> 0 Then Grid.CellBackColor = Grid.BackColor
End If
End Sub


Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim I As Integer

If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.colWidth(Grid.MouseCol)

If Regim = "select" Or Grid.MouseRow = 0 Then Exit Sub
mousCol = Grid.MouseCol
If Button = 2 And frmMode = "" And mousCol > 0 Then

'    If quantity > 0 And Grid.row <> Grid.RowSel Then !!!Зачем, это не дает переместить одно изделие
    If quantity > 0 Then
        ReDim NN(Grid.RowSel - Grid.row + 1)
        For I = Grid.row To Grid.RowSel
            NN(I - Grid.row + 1) = Grid.TextMatrix(I, gpId) 'только для перемещения
        Next I
    End If

    Grid.col = mousCol
    Grid.row = Grid.MouseRow
    
    On Error Resume Next
    Grid.SetFocus
    Grid.CellBackColor = vbButtonFace
    gProductId = Grid.TextMatrix(Grid.row, gpId)
    
    If quantity = 0 Then
        mnRepl2.Visible = False
        mnDel2.Visible = False
        mnSep2.Visible = False
    Else
        mnRepl2.Visible = True
        mnDel2.Visible = True
        mnSep2.Visible = True
    End If
    Timer1.Interval = 10
    Timer1.Enabled = True
ElseIf frmMode = "productReplace" Then
    Me.PopupMenu mnContext4
End If
        
End Sub

Private Sub Grid2_Click()
mousCol2 = Grid2.MouseCol
mousRow2 = Grid2.MouseRow
If quantity2 = 0 Then Exit Sub

If Grid2.MouseRow = 0 Then
    Grid2.CellBackColor = Grid2.BackColor
    If mousCol2 = gpQuant Then
        SortCol Grid2, mousCol2, "numeric"
    Else
        SortCol Grid2, mousCol2
    End If
    SortCol Grid2, mousCol2
    Grid2.row = 1    ' только чтобы снять выделение
    Grid2_EnterCell
End If

End Sub

Private Sub Grid2_DblClick()
If mousRow2 = 0 Or Grid2.CellBackColor <> &H88FF88 Then Exit Sub

If mousCol2 = gpWeb Then
    listBoxInGridCell lbWeb, Grid2, "select"
Else
    tmpStr = productUsedIn(gProductId)
    If tmpStr <> "" Then
        MsgBox "Это изделие используется в заказах: " & tmpStr, , _
        "Редактирование невозможно!"
        Exit Sub
    End If
    textBoxInGridCell tbMobile2, Grid2
End If
End Sub

Private Sub Grid2_EnterCell()
mousRow2 = Grid2.row
mousCol2 = Grid2.col

If Not gridIsLoad Then Exit Sub

gNomNom = Grid2.TextMatrix(mousRow2, gpNomNom)
If quantity2 = 0 Or frmMode <> "" Or Regim = "select" Then Exit Sub
If mousCol2 = gpGroup Then
    tbMobile2.MaxLength = 1
Else
    tbMobile2.MaxLength = 10
End If

If mousCol2 >= gpQuant And Grid.TextMatrix(mousRow, gpUsed) = "" Then
    Grid2.CellBackColor = &H88FF88
Else
    Grid2.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid2_GotFocus()
        controlGridsWidth

End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid2_DblClick

End Sub

Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Grid2_EnterCell

End Sub

Private Sub Grid2_LeaveCell()
Grid2.CellBackColor = Grid2.BackColor

End Sub
Private Function productUsedIn(ByVal productId As Integer) As String

    sql = "SELECT o.numorder FROM Orders o" _
    & " JOIN xPredmetyByIzdelia i ON O.numOrder = i.numOrder and i.prId = " & productId _
    & " GROUP BY o.numOrder"
    
    Set tbProduct = myOpenRecordSet("##320", sql, dbOpenForwardOnly)
    If tbProduct Is Nothing Then Exit Function
    If Not tbProduct.BOF Then
      While Not tbProduct.EOF
        productUsedIn = productUsedIn & "  " & tbProduct!numorder
        tbProduct.MoveNext
      Wend
    End If
    tbProduct.Close

End Function

Private Sub Grid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid2.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid2.colWidth(Grid2.MouseCol)
If Regim = "select" Then Exit Sub

If Button = 2 And frmMode = "" Then

    tmpStr = productUsedIn(gProductId)
    If tmpStr <> "" Then
        MsgBox "Это изделие используется в заказах: " & tmpStr, , _
        "Редактирование невозможно!"
        Exit Sub
    End If

    Grid2.CellBackColor = vbButtonFace

Dim startRow As Integer, stopRow As Integer, curRow As Integer

    If Grid2.row >= Grid2.RowSel Then
        startRow = Grid2.RowSel
        stopRow = Grid2.row
    Else
        startRow = Grid2.row
        stopRow = Grid2.RowSel
    End If
    
    If startRow <> stopRow Then
        mnEdit5.Visible = False
        mnAdd5.Visible = False
    Else
        mnEdit5.Visible = True
        mnAdd5.Visible = True
    End If
    
    
    If quantity2 = 0 Then
        mnDel5.Visible = False
    Else
        mnDel5.Visible = True
    End If
    If Grid.TextMatrix(mousRow, gpUsed) = "" Then
        Me.PopupMenu mnContext5
    Else
        'MsgBox ""
    End If
End If

End Sub

Sub controlVisible()
Grid.Visible = False
Frame1.Visible = False
Grid2.Visible = False
cmSel.Visible = False
tbQuant.Visible = False
laQuant.Visible = False
laProduct.Caption = ""
laNomenk.Caption = ""
End Sub

Private Sub lbPrWeb_DblClick()
Dim success As Boolean, prodCategoryId As Integer, val As String

prodCategoryId = lbPrWeb.ItemData(lbPrWeb.ListIndex)
If prodCategoryId = 0 Then
    val = "null"
Else
    val = CStr(prodCategoryId)
End If
success = ValueToTableField("##411", val, "sGuideProducts", "prodCategoryId", "byProductId")
If success Then
    Grid.TextMatrix(mousRow, gpPrWeb) = lbPrWeb.Text
End If
lbHide

End Sub

Private Sub lbPrWeb_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbPrWeb_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

'setWebFlags
Private Sub lbWeb_DblClick()

If Nomenklatura.setWebFlags(Grid2.TextMatrix(mousRow2, gpWeb), lbWeb.Text) _
Then Grid2.TextMatrix(mousRow2, gpWeb) = lbWeb.Text

lbHide2

End Sub

Private Sub lbWeb_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbWeb_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide2
End If

End Sub

Private Sub mnAdd_Click()
Static I As Integer
Dim str  As String, id As Integer
controlVisible

I = I + 1
str = "новый " & I
'cmClassAdd.Enabled = False
wrkDefault.BeginTrans
sql = "UPDATE sGuideSeries SET seriaId = seriaId WHERE seriaId=0"
myBase.Execute sql

sql = "SELECT max(seriaId) FROM sGuideSeries"
If Not byErrSqlGetValues("##461", sql, id) Then GoTo ERR1
id = id + 1

'sql = "SELECT sGuideSeries.seriaId, sGuideSeries.seriaName, " & _
'"sGuideSeries.parentseriaId From sGuideSeries ORDER BY sGuideSeries.seriaId;"
'Set tbSeries = myOpenRecordSet("##106", sql, dbOpenDynaset)
'If tbSeries Is Nothing Then Exit Sub
'If tbSeries.BOF Then
'    id = 1
'Else
'    tbSeries.MoveLast
'    id = tbSeries!seriaId + 1
'End If
'tbSeries.AddNew
'tbSeries!seriaId = id
'tbSeries!seriaName = str
'tbSeries!parentSeriaId = gSeriaId
'tbSeries.Update
sql = "INSERT INTO sGuideSeries (seriaId, seriaName, parentSeriaId) " & _
" values (" & id & ", '" & str & "', " & gSeriaId & ")"
'MsgBox sql
If myExecute("##462", sql) <> 0 Then GoTo ERR1

wrkDefault.CommitTrans

EN1:
'tbSeries.Close
Set Node = tv.Nodes.Add(tv.SelectedItem.Key, tvwChild, "k" & id, str)
tv.Nodes("k" & id).EnsureVisible
tv.Nodes("k" & id).Selected = True
tv.StartLabelEdit
Exit Sub

ERR1:
errorCodAndMsg ("##461")


End Sub

Function productFormula(Optional noOpen As String = "")
Dim str As String

If noOpen = "" Then
    sql = "SELECT sGuideProducts.*, sGuideFormuls.Formula FROM sGuideFormuls " & _
    "INNER JOIN sGuideProducts ON sGuideFormuls.nomer = sGuideProducts.formulaNom " & _
    "WHERE (((sGuideProducts.prId)=" & gProductId & "));"
    'MsgBox sql
    Set tbProduct = myOpenRecordSet("##316", sql, dbOpenDynaset)
    If tbProduct Is Nothing Then Exit Function
    If tbProduct.BOF Then tbProduct.Close: Exit Function
End If

SumCenaFreight = getSumCena(tbProduct!prId)
If InStr(tbProduct!formula, "SumCenaFreight") > 0 Then
  If IsNumeric(SumCenaFreight) Then
    sc.ExecuteStatement "SumCenaFreight=" & SumCenaFreight
    SumCenaFreight = Round(CSng(SumCenaFreight), 2)
  Else
    productFormula = "error СумЦ.доставка" 'текст ошибки
    tbProduct.Close
    GoTo EN1
  End If
End If

SumCenaSale = getSumCena(tbProduct!prId, "Sale")
If InStr(tbProduct!formula, "SumCenaSale") > 0 Then
  If IsNumeric(SumCenaSale) Then
    sc.ExecuteStatement "SumCenaSale=" & SumCenaSale
    SumCenaSale = Round(CSng(SumCenaSale), 2)
  Else
    tbProduct.Close
    productFormula = "error СумЦоПродажа" 'текст ошибки
    GoTo EN1
  End If
End If

On Error GoTo ERR2
sc.ExecuteStatement "VremObr = " & tbProduct!VremObr
productFormula = Round(sc.Eval(tbProduct!formula), 2)
GoTo EN1
ERR2:
    productFormula = "error: " & Error
'    MsgBox Error & " - при выполнении формулы '" & tbProduct!formula & _
'    "' для изделия '" & tbProduct!prName & "' (" & tmpStr & ")", , _
'    "Ошибка 316 - " & Err & ":  " '##316
EN1:
tmpStr = tbProduct!formula
If noOpen = "" Then tbProduct.Close
End Function



Sub productAdd(Optional obraz As String = "")
Dim str As String

'Grid.CellBackColor = vbWhite
If Grid.TextMatrix(Grid.Rows - 1, gpName) = "" Then 'если последняя строка(код) пуста, т.е. если таблица пуста
    frmMode = "productAdd"
ElseIf obraz = "" Then
    frmMode = "productAdd"
    Grid.AddItem ""
Else
    Grid.AddItem ""
    frmMode = "productCopy"
End If
'If quantity > 0 Then Grid.AddItem (str)
'    Grid.TextMatrix(Grid.Rows - 1, gpSize) = str
'    Grid.TextMatrix(Grid.Rows - 1, gpDescript) = sql
If frmMode = "productCopy" Then
    Grid.TextMatrix(Grid.Rows - 1, gpSumCenaFreight) = Grid.TextMatrix(mousRow, gpSumCenaFreight)
    Grid.TextMatrix(Grid.Rows - 1, gpSumCenaSale) = Grid.TextMatrix(mousRow, gpSumCenaSale)
    Grid.TextMatrix(Grid.Rows - 1, gpCena3) = Grid.TextMatrix(mousRow, gpCena3)
    Grid.TextMatrix(Grid.Rows - 1, gpFormula) = Grid.TextMatrix(mousRow, gpFormula)
    Grid.TextMatrix(Grid.Rows - 1, gpCol1) = Grid.TextMatrix(mousRow, gpCol1)
    Grid.TextMatrix(Grid.Rows - 1, gpCol2) = Grid.TextMatrix(mousRow, gpCol2)
    Grid.TextMatrix(Grid.Rows - 1, gpCol3) = Grid.TextMatrix(mousRow, gpCol3)
    Grid.TextMatrix(Grid.Rows - 1, gpCol4) = Grid.TextMatrix(mousRow, gpCol4)
    Grid.TextMatrix(Grid.Rows - 1, gpSortNom) = Grid.TextMatrix(mousRow, gpSortNom)
    Grid.TextMatrix(Grid.Rows - 1, gpVremObr) = Grid.TextMatrix(mousRow, gpVremObr)
    Grid.TextMatrix(Grid.Rows - 1, gpFormulaNom) = Grid.TextMatrix(mousRow, gpFormulaNom)
    Grid.TextMatrix(Grid.Rows - 1, gpPage) = Grid.TextMatrix(mousRow, gpPage)
    Grid.TextMatrix(Grid.Rows - 1, gpSize) = Grid.TextMatrix(mousRow, gpSize)
    Grid.TextMatrix(Grid.Rows - 1, gpDescript) = Grid.TextMatrix(mousRow, gpDescript)
    Grid.TextMatrix(Grid.Rows - 1, gpPrWeb) = Grid.TextMatrix(mousRow, gpPrWeb)
    Grid.TextMatrix(Grid.Rows - 1, gpRabbat) = Grid.TextMatrix(mousRow, gpRabbat)
Else
    Grid.TextMatrix(Grid.Rows - 1, gpSumCenaFreight) = "Error: Не обнаружены комплектующие"
    Grid.TextMatrix(Grid.Rows - 1, gpSumCenaSale) = "Error: Не обнаружены комплектующие"
End If

str = Grid.TextMatrix(mousRow, gpName)
gridIsLoad = False
Grid.row = Grid.Rows - 1
mousRow = Grid.Rows - 1
Grid.col = gpName
mousCol = gpName
gridIsLoad = True
'Grid.SetFocus
textBoxInGridCell tbMobile, Grid
If obraz <> "" Then
    tbMobile.Text = str
    tbMobile.SelStart = Len(str)
End If

End Sub

Private Sub mnAdd2_Click()
Grid.CellBackColor = Grid.BackColor

productAdd
End Sub

Private Sub mnAdd5_Click()
    Nomenklatura.Regim = "nomenkSelect" 'new
    Timer2.Interval = 10 'new
    Timer2.Enabled = True
End Sub

Private Sub mnCancel_Click()
    frmMode = ""
    mnRepl.Caption = "Переместить"
    mnAdd.Visible = True
    mnRen.Visible = True
    mnDel.Visible = True
    mnSep.Visible = True
    mnCancel.Visible = False
    Me.MousePointer = flexDefault

End Sub

Private Sub mnCancel3_Click()
frmMode = ""
Me.MousePointer = flexDefault
Grid.CellBackColor = Grid.BackColor
On Error Resume Next
tv.SetFocus

End Sub

Private Sub mnCancel4_Click()
frmMode = ""
Me.MousePointer = flexDefault
Grid.CellBackColor = Grid.BackColor
On Error Resume Next
tv.SetFocus

End Sub

Private Sub mnCopy2_Click()
productAdd "obraz"
End Sub

Private Sub mnDel_Click()
Dim I As Integer

If MsgBox("Для удаления класса  нажмите <Да>." & Chr(13) & Chr(13) & _
"Удаление возможно, если класс не содержит элементов и других подклассов", _
vbYesNo Or vbDefaultButton2, "Удалить '" & tv.SelectedItem.Text & _
"'. Вы уверены?") = vbNo Then GoTo EN1


sql = "DELETE  From sGuideSeries WHERE seriaId =" & gSeriaId
I = myExecute("##107", sql, -198)
If I = 0 Then
    tv.Nodes.Remove tv.SelectedItem.Key
    controlVisible
Else
    Exit Sub

End If
EN1:
On Error Resume Next
tv.SetFocus

End Sub

Private Sub mnDel2_Click()
Dim I As Integer
If frmMode = "productReplace" Then
    On Error Resume Next
    tv.SetFocus
ElseIf frmMode = "" Then
  gProductId = Grid.TextMatrix(mousRow, gpId)
  If MsgBox("После нажатия <Да> данный элемент будет удален из Справочника", _
  vbDefaultButton2 Or vbYesNo, "Удалить '" & Grid.TextMatrix(mousRow, gpName) & "'. Вы уверены?") _
  = vbYes Then
    sql = "DELETE From sGuideProducts " & _
    "WHERE (((sGuideProducts.prId)=" & gProductId & "));"
'    MsgBox sql
    I = myExecute("##114", sql, -198)
    If I = 0 Then
        quantity = quantity - 1
        If quantity = 0 Then
            clearGridRow Grid, mousRow
        Else
            Grid.RemoveItem mousRow
        End If
    ElseIf I = -2 Then
        MsgBox "Нельзя удалять непустое изделие, сначала удалите входящие " & _
        "в него элементы.", , "Удаление невозможно !"
    End If
'  Else
'    Grid.CellBackColor = Grid.BackColor
  End If
 Grid.CellBackColor = Grid.BackColor
 Grid_EnterCell
 On Error Resume Next
 Grid.SetFocus
End If

End Sub

Private Sub mnDel5_Click() 'см. mnEdit5_Click
Dim startRow As Integer, stopRow As Integer, curRow As Integer

    If Grid2.row >= Grid2.RowSel Then
        startRow = Grid2.RowSel
        stopRow = Grid2.row
    Else
        startRow = Grid2.row
        stopRow = Grid2.RowSel
    End If
    
    If stopRow <> startRow Then
        If MsgBox("Вы уверены, что хотите удалить " & stopRow - startRow + 1 & " компонент(a)(ов)?", vbYesNo, "Требуется подтверждение") = vbNo Then Exit Sub
    End If
    For curRow = startRow To stopRow
        sql = "DELETE From sProducts WHERE (((sProducts.ProductId)=" & gProductId & ") " & _
        "AND ((sProducts.nomNom)='" & Grid2.TextMatrix(startRow, gpNomNom) & "'));"
        'Debug.Print Grid2.TextMatrix(startRow, gpNomNom)
        myBase.Execute sql
        quantity2 = quantity2 - 1
        If quantity2 = 0 Then
            clearGridRow Grid2, 1
        Else
            Grid2.RemoveItem startRow
        End If
    Next curRow
    
    On Error Resume Next
    Grid2.SetFocus
    Grid2.col = 0


End Sub

Private Sub mnEdit5_Click() ' см. mnDel5_Click
soursNom = Grid2.TextMatrix(Grid2.row, gpNomNom)

Nomenklatura.Regim = "singleSelect"
Timer2.Interval = 10 ' добавляем ном-ру
Timer2.Enabled = True

End Sub

Private Sub mnInsert_Click()
Dim str As String, I As Integer

frmMode = ""
Grid.CellBackColor = Grid.BackColor
    
Me.MousePointer = flexDefault
str = Mid$(tv.SelectedItem.Key, 2)
For I = 1 To UBound(NN)
    gProductId = NN(I)
    ValueToTableField "##112", str, "sGuideProducts", "prSeriaId", "byProductId"
Next I
tv_NodeClick tv.SelectedItem
On Error Resume Next
tv.SetFocus

End Sub

Private Sub mnRen_Click()
gSeriaId = gSeriaId
tv.StartLabelEdit

End Sub

Private Sub mnRepl_Click()
Dim str As String
str = tv.SelectedItem.Key
If frmMode = "" Then
    If str = "k0" Then Exit Sub
    frmMode = "seriaReplace"
    mnRepl.Caption = "Вставить"
    mnAdd.Visible = False
    mnRen.Visible = False
    mnDel.Visible = False
    mnSep.Visible = False
    mnCancel.Visible = True
    Me.MousePointer = flexUpArrow
    nodeKey = str
ElseIf frmMode = "seriaReplace" Then
    frmMode = ""
    mnRepl.Caption = "Переместить"
    mnAdd.Visible = True
    mnRen.Visible = True
    mnDel.Visible = True
    mnSep.Visible = True
    mnCancel.Visible = False
    Me.MousePointer = flexDefault
    controlVisible
    If str = nodeKey Then
        MsgBox "Нельзя переместить серию саму в себя", , "Недопустимая операция!"
    Else
        sql = "UPDATE sGuideSeries SET sGuideSeries.parentSeriaId = " & _
        Mid$(str, 2) & " WHERE (((sGuideSeries.seriaId)=" & Mid$(nodeKey, 2) & "));"
'        MsgBox sql
        myBase.Execute sql
        loadSeria
    End If
    
ElseIf frmMode = "produktReplace" Then
MsgBox "неиспользуемый алгоритм", , "Err ##888"
End
End If

End Sub

Private Sub mnRepl2_Click()
Me.MousePointer = flexUpArrow
On Error Resume Next
tv.SetFocus
frmMode = "productReplace"

End Sub

Private Sub tbCol_Change(index As Integer)
cmApple.Enabled = True
cmCancel.Enabled = True

End Sub

Private Sub tbGain_Change(index As Integer)
cmApple.Enabled = True
cmCancel.Enabled = True

End Sub

Private Sub tbMobile_DblClick()
lbHide
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, I As Integer ', str2 As String

If KeyCode = vbKeyReturn Then
 
 str = tbMobile.Text
 tmpStr = str
 
 If (frmMode = "" And mousCol = gpCol1) Then  ' это проверка д.б. вначале
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not ValueToTableField("##111", str, "sGuideProducts", _
    "Cena4", "byProductId") Then GoTo EN1
 ElseIf mousCol = gpName Then
  If InStr(str, "/") > 0 Then
    MsgBox "Код изделия не должен содержать символ '/', так как он " & _
    "используется как разделитель между вариантом и кодом  Готового " & _
    "изделия в предметах заказа.", , ""
    Exit Sub
  End If
  If frmMode = "productAdd" Or frmMode = "productCopy" Then
    
    wrkDefault.BeginTrans 'lock01
    sql = "update system set resursLock = resursLock" 'lock02
    myBase.Execute (sql) 'lock03


    sql = "SELECT max(prId) From sGuideProducts"
    If Not byErrSqlGetValues("##463", sql, I) Then Exit Sub
    I = I + 1
    Dim flds As String, vals As String
    flds = "prId, prName, prSeriaId"
    vals = I & ", '" & str & "', " & gSeriaId
    If frmMode = "productCopy" Then
        On Error GoTo Rollback ' т.к. исходная м.б. почти пустая
        str = Grid.TextMatrix(mousRow, gpSortNom)
        If str <> "" Then flds = flds & ", SortNom": vals = vals & ", '" & str & "'"
        str = Grid.TextMatrix(mousRow, gpVremObr)
        If str <> "" Then flds = flds & ", VremObr": vals = vals & ", '" & str & "'"
        str = Grid.TextMatrix(mousRow, gpFormulaNom)
        If str <> "" Then flds = flds & ", FormulaNom": vals = vals & ", '" & str & "'"
        str = Grid.TextMatrix(mousRow, gpCol1)
        If str <> "" Then flds = flds & ", Cena4": vals = vals & ", '" & str & "'"
        str = Grid.TextMatrix(mousRow, gpPage)
        If str <> "" Then flds = flds & ", Page": vals = vals & ", '" & str & "'"
        str = Grid.TextMatrix(mousRow, gpSize)
        If str <> "" Then flds = flds & ", prSize": vals = vals & ", '" & str & "'"
        str = Grid.TextMatrix(mousRow, gpRabbat)
        If str <> "" Then flds = flds & ", rabbat": vals = vals & ", '" & str & "'"
        str = Grid.TextMatrix(mousRow, gpPrWeb)
        If str <> "" Then
            sql = "select * from GuideProdCategory where sysname = '" & str & "'"
            byErrSqlGetValues "##1001", sql, str
            If str <> "" Then
                flds = flds & ", prodCategoryId"
                vals = vals & ", " & str
            End If
        End If
        str = Grid.TextMatrix(mousRow, gpDescript)
        If str <> "" Then flds = flds & ", prDescript": vals = vals & ", '" & str & "'"
'        On Error GoTo 0
    End If
    sql = "INSERT INTO sGuideProducts (" & flds & ") VALUES (" & vals & ")"
    If myExecute("##111", sql) <> 0 Then GoTo Rollback
    wrkDefault.CommitTrans
    
    Grid.TextMatrix(mousRow, gpName) = tmpStr
    Grid.TextMatrix(mousRow, gpId) = I
    quantity = quantity + 1
    Grid.TextMatrix(mousRow, gpNomenk) = quantity
    If frmMode = "productCopy" Then
        sql = "INSERT INTO sProducts ( ProductId, nomNom, quantity, xGroup ) " & _
        "SELECT " & I & ", sProducts.nomNom, sProducts.quantity, sProducts.xGroup " & _
        "From sProducts WHERE (((sProducts.ProductId)=" & gProductId & "));"
        myExecute "##155", sql, 0 'предметов м. и не быть
    End If
    frmMode = ""
    GoTo EN1
  ElseIf frmMode = "" Then
    ValueToTableField "##111", "'" & str & "'", "sGuideProducts", "prName", "byProductId"
  End If
 ElseIf mousCol = gpSortNom Then
    ValueToTableField "##111", "'" & str & "'", "sGuideProducts", "SortNom", "byProductId"
 ElseIf mousCol = gpSize Then
    ValueToTableField "##111", "'" & str & "'", "sGuideProducts", "prSize", "byProductId"
 ElseIf mousCol = gpDescript Then
    ValueToTableField "##111", "'" & str & "'", "sGuideProducts", "prDescript", "byProductId"
 ElseIf mousCol = gpRabbat Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not ValueToTableField("##111", str, "sGuideProducts", "rabbat", "byProductId") Then GoTo EN1
 ElseIf mousCol = gpVremObr Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not ValueToTableField("##111", str, "sGuideProducts", "vremObr", _
    "byProductId") Then GoTo EN1
    refrProductCenaToGrid
 ElseIf mousCol = gpPage Then
'    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not ValueToTableField("##111", "'" & str & "'", "sGuideProducts", "page", _
    "byProductId") Then GoTo EN1
 End If

 Grid.TextMatrix(mousRow, mousCol) = str
 lbHide
ElseIf KeyCode = vbKeyEscape Then
 If mousCol = gpName And (frmMode = "productAdd" Or frmMode = "productCopy") Then
    frmMode = ""
    If quantity > 0 Then Grid.RemoveItem Grid.Rows - 1
 End If
Rollback:
  wrkDefault.Rollback
EN1:
  lbHide
End If
End Sub

Private Sub tbMobile2_DblClick()
lbHide
End Sub

Private Sub tbMobile2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String

If KeyCode = vbKeyReturn Then
    str = Trim(tbMobile2.Text)
    If mousCol2 = gpGroup Then
'        Grid2.TextMatrix(mousRow2, mousCol2) = str
        sql = "xgroup ='" & str & "'"
    Else
        If Not isNumericTbox(tbMobile2, 0) Then Exit Sub
        sql = "quantity =" & tmpNum
'        str = tmpNum возможно понадобится
    End If
    sql = "UPDATE sProducts SET sProducts." & sql & " WHERE (((sProducts" & _
    ".ProductId)=" & gProductId & ") AND ((sProducts.nomNom)='" & gNomNom & "'));"
'    MsgBox sql
    If myExecute("##154", sql) = 0 Then
            Grid2.TextMatrix(mousRow2, mousCol2) = str
'            If mousCol2 = gpQuant Then refrProductCenaToGrid
            refrProductCenaToGrid
    End If
    GoTo EN1
ElseIf KeyCode = vbKeyEscape Then
EN1:
  lbHide2
  On Error Resume Next
  Grid2.SetFocus
End If

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Me.PopupMenu mnContext2

End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False

'Nomenklatura.Regim = "nomenkSelect" 'new
Nomenklatura.setRegim
Nomenklatura.Show vbModal           '
loadProductNomenk gpQuant ' эту кол-ку подсветить в строке где gNomNom
refrProductCenaToGrid
End Sub

Private Sub tv_AfterLabelEdit(Cancel As Integer, NewString As String)
' If Not flseriaAdd Then
'ValueToTableField "##115", "'" & NewString & "'", "sProducts", "seriaName", "bySeriaId"
gSeriaId = Mid$(tv.SelectedItem.Key, 2)
ValueToTableField "##115", "'" & NewString & "'", "sGuideSeries", "seriaName", "bySeriaId"
End Sub

Sub loadProductNomenk(Optional markCol As Integer = 0)
Dim il As Long, str As String

Me.MousePointer = flexHourglass
Grid2.Visible = False
gridIsLoad = False
'laProduct.Caption = "Список готовых изделий по серии '" & tv.SelectedItem.Text & "'"
For il = Grid2.Rows To 3 Step -1
    Grid2.RemoveItem (il)
Next il
clearGridRow Grid2, 1
quantity2 = 0

sql = "SELECT p.*, n.nomName, n.ed_Izmer, n.Size, n.cod, n.perList, n.CENA1, n.VES, n.STAVKA" _
    & ", n.CENA_W, f.Formula, pc.sysname as web " _
    & " FROM sProducts p " _
    & " JOIN sGuideNomenk n on n.nomNom = p.nomNom " _
    & " JOIN sGuideFormuls f ON f.nomer = n.formulaNom " _
    & " JOIN sGuideProducts gp on gp.prId = p.productId " _
    & " left join GuideProdCategory pc on pc.prodCategoryId = gp.prodCategoryId " _
    & " WHERE p.ProductId = " & gProductId

Set tbNomenk = myOpenRecordSet("##108", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then Exit Sub
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    quantity2 = quantity2 + 1
       
    Grid2.TextMatrix(quantity2, gpNomNom) = tbNomenk!nomnom
    Grid2.TextMatrix(quantity2, gpNomName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid2.TextMatrix(quantity2, gpEdIzm) = tbNomenk!ed_izmer
    If Not IsNull(tbNomenk!quantity) Then _
            Grid2.TextMatrix(quantity2, gpQuant) = tbNomenk!quantity
    str = nomenkFormula("noOpen")
    If IsNumeric(str) Then
        Grid2.TextMatrix(quantity2, gpCenaFreight) = Round(str / tbNomenk!perList, 3)
    Else
        Grid2.TextMatrix(quantity2, gpCenaFreight) = str
    End If
'    str = tbNomenk!CENA_W
'    If IsNumeric(str) Then
'        Grid2.TextMatrix(quantity2, gpCENA_W) = Round(str / tbNomenk!perList, 3)
        Grid2.TextMatrix(quantity2, gpCENA_W) = Round(tbNomenk!CENA_W / tbNomenk!perList, 3)
'    Else
'        Grid2.TextMatrix(quantity2, gpCENA_W) = str
'    End If
    Grid2.TextMatrix(quantity2, gpGroup) = tbNomenk!xgroup
    If Not IsNull(tbNomenk!web) Then
        Grid2.TextMatrix(quantity2, gpWeb) = tbNomenk!web
    End If
    
    Grid2.AddItem ""
    tbNomenk.MoveNext
  Wend
  Grid2.RemoveItem quantity2 + 1
End If
tbNomenk.Close
If quantity2 = 0 Then
    laNomenk.Caption = ""
Else
    laNomenk.Caption = "Список номенклатуры по изделию '" & _
    Grid.TextMatrix(mousRow, gpName) & "'"
End If

'    Grid2.CellBackColor = Grid2.BackColor
    trigger = True
    SortCol Grid2, gpNomNom
    Grid2.row = 1    ' только чтобы снять выделение
'    Grid2_EnterCell

If markCol > 0 Then
  For il = 1 To quantity2
    If Grid2.TextMatrix(il, gpNomNom) = gNomNom Then
        Grid2.row = il
        gridIsLoad = True
        Grid2.col = markCol
        Grid2.Visible = True
        On Error Resume Next
        Grid2.SetFocus
        gridIsLoad = False
     End If
  Next il
Else
    Grid2.row = 1    ' только чтобы снять выделение
End If

Grid2.Visible = True
gridIsLoad = True
Me.MousePointer = flexDefault
End Sub


'Function speChaRemov(str As String) As String
'Dim chr As String, i As Integer

''speChaRemov = "за " & str & " шт.": Exit Function
'speChaRemov = " " & str: Exit Function


'почемуто строка "1-3" в Excel воспринимается как 3.января
'speChaRemov = ""
'For i = 1 To Len(str)
'    chr = Mid$(str, i, 1)
'    If chr = "-" Then
'        speChaRemov = speChaRemov & "~"
'    Else
'        speChaRemov = speChaRemov & chr
'    End If
    
'Next i
'End Function


Sub loadSeriaProduct(Optional filtr As String = "")
Dim il As Long, strWhere As String, str  As String

Grid.Visible = False
Frame1.Visible = False
gridIsLoad = False

If tv.SelectedItem.Key = "k0" And filtr = "" Then
    gSeriaId = 0
    Grid.Visible = False
    Frame1.Visible = False
    If frmMode <> "" Then GoTo EN1
    laProduct.Caption = ""
    Exit Sub
End If

Me.MousePointer = flexHourglass

clearGrid Grid
'For il = Grid.Rows To 3 Step -1
'    Grid.RemoveItem (il)
'Next il
'clearGridRow Grid, 1


If Not getGainAndHead Then GoTo EN1

Grid.TextMatrix(0, gpCol1) = head1
Grid.TextMatrix(0, gpCol2) = head2
Grid.TextMatrix(0, gpCol3) = head3
Grid.TextMatrix(0, gpCol4) = head4
tbCol(0).Text = head1
tbCol(1).Text = head2
tbCol(2).Text = head3
tbCol(3).Text = head4
tbGain(0).Text = gain2
tbGain(1).Text = gain3
tbGain(2).Text = gain4

cmApple.Enabled = False ' именно здесь
cmCancel.Enabled = False


il = 0
quantity = 0


If filtr = "" Then
    laProduct.Caption = "Список готовых изделий по серии '" & tv.SelectedItem.Text & "'"
    strWhere = "WHERE p.prSeriaId = " & gSeriaId
Else
    If mousCol = gpDescript Then
        str = "Описание"
        strWhere = "WHERE p.prDescript Like '%" & filtr & "%'"
    Else
        str = "Номер"
        strWhere = "WHERE p.prName Like '%" & filtr & "%'"
    End If
    laProduct.Caption = "Список готовых изделий по фильтру '" & filtr & _
    "' в колонке '" & str & "'"
End If

sql = "SELECT p.prId, p.prName, p.prSize, p.SortNom, p.VremObr, p.FormulaNom, p.prDescript, p.cena4, p.page, pc.sysname as web, f.Formula, p.rabbat " _
    & ", max(i.prid) as used" _
    & " FROM sGuideProducts p " _
    & " LEFT JOIN sGuideFormuls f ON f.nomer = p.formulaNom " _
    & " left join xPredmetyByIzdelia i on i.prId = p.prId " _
    & " left join GuideProdCategory pc on pc.prodCategoryId = p.prodCategoryId " _
    & strWhere _
    & " GROUP BY p.prId, p.prName, p.prSize, p.SortNom, p.VremObr, p.FormulaNom, p.prDescript, p.cena4, p.page, pc.sysname, f.Formula, p.rabbat " _
    & " ORDER BY p.SortNom"

'Debug.Print sql


Set tbProduct = myOpenRecordSet("##103", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then GoTo EN1
If Not tbProduct.BOF Then
 While Not tbProduct.EOF
    quantity = quantity + 1
    Grid.TextMatrix(quantity, gpId) = tbProduct!prId
    Grid.TextMatrix(quantity, gpNomenk) = quantity
    Grid.TextMatrix(quantity, gpRabbat) = tbProduct!rabbat
    If Not IsNull(tbProduct!prName) Then _
            Grid.TextMatrix(quantity, gpName) = tbProduct!prName
    Grid.TextMatrix(quantity, gpSortNom) = tbProduct!SortNom
    If Not IsNull(tbProduct!prSize) Then _
            Grid.TextMatrix(quantity, gpSize) = tbProduct!prSize
    Grid.TextMatrix(quantity, gpVremObr) = tbProduct!VremObr
    Grid.TextMatrix(quantity, gpFormulaNom) = tbProduct!FormulaNom
    If Not IsNull(tbProduct!prDescript) Then _
            Grid.TextMatrix(quantity, gpDescript) = tbProduct!prDescript
'If tbProduct!prName = "штучки22" Then
    Grid.TextMatrix(quantity, gpCena3) = Format(productFormula("noOpen"), "0.00")
    Grid.TextMatrix(quantity, gpSumCenaFreight) = Format(SumCenaFreight, "0.00") ' только после
    Grid.TextMatrix(quantity, gpSumCenaSale) = Format(SumCenaSale, "0.00") ' только после
    Grid.TextMatrix(quantity, gpFormula) = tmpStr                    ' productFormula
'End If
    Grid.TextMatrix(quantity, gpCol1) = Format(tbProduct!Cena4, "0.00")
    
    Grid.TextMatrix(quantity, gpCol2) = Format(Round(tbProduct!Cena4 * gain2, 1), "0.00")
    Grid.TextMatrix(quantity, gpCol3) = Format(Round(tbProduct!Cena4 * gain3, 1), "0.00")
    Grid.TextMatrix(quantity, gpCol4) = Format(Round(tbProduct!Cena4 * gain4, 1), "0.00")
    Grid.TextMatrix(quantity, gpPage) = tbProduct!Page
    If Not IsNull(tbProduct!web) Then
        Grid.TextMatrix(quantity, gpPrWeb) = tbProduct!web
    End If
    If Not IsNull(tbProduct!used) Then
        Grid.TextMatrix(quantity, gpUsed) = tbProduct!used
    End If
    

    Grid.AddItem ""
    tbProduct.MoveNext
 Wend
 Grid.RemoveItem quantity + 1
End If
tbProduct.Close

'Grid.col = gpName
'Grid.row = 1
'Grid.CellBackColor = &H88FF88 'Grid_EnterCell нельзя
Grid.Visible = True
Frame1.Visible = True
'Grid.SetFocus
EN1:
If frmMode = "productReplace" Then
    Me.MousePointer = flexUpArrow
Else
    Me.MousePointer = flexDefault
End If
gridIsLoad = True
End Sub
    
Private Sub tv_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer, str As String
If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
    tv_NodeClick tv.SelectedItem
End If
End Sub

Private Sub tv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 2 And Regim = "" Then
If Button = 2 And Regim <> "select" Then
    mousRight = 1
Else
    mousRight = 0
End If

End Sub

Private Sub tv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim str As String
If mousRight = 2 Then
  If frmMode = "productReplace" Then
    Me.PopupMenu mnContext3
  Else
    str = tv.SelectedItem.Key
'    If str = "all" Then Exit Sub
    If str = "k0" Then
        mnRen.Visible = False
        mnDel.Visible = False
        If frmMode <> "seriaReplace" Then mnRepl.Visible = False
        mnSep.Visible = False
    ElseIf frmMode = "" Then
        mnRen.Visible = True
        mnDel.Visible = True
        mnRepl.Visible = True
        mnSep.Visible = True
    End If
    Me.PopupMenu mnContext
  End If
'    Timer1.Interval = 10
'    Timer1.Enabled = True
End If

End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)

gSeriaId = Mid$(tv.SelectedItem.Key, 2)
controlGridsWidth "left"

If mousRight = 1 Then
    mousRight = 2 ' правый клик был именно из Node
ElseIf frmMode = "" Then
    Grid2.Visible = False
    laNomenk.Caption = ""

    loadSeriaProduct
    Grid_EnterCell 'чтобы  prevCol = 0

End If

End Sub


