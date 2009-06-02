VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Nomenklatura 
   BackColor       =   &H8000000A&
   Caption         =   "Справочник по номенклатуре"
   ClientHeight    =   6396
   ClientLeft      =   60
   ClientTop       =   1740
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6396
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox tbPostav 
      Height          =   285
      Left            =   8400
      TabIndex        =   42
      Text            =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   10740
      TabIndex        =   8
      Top             =   5940
      Width           =   855
   End
   Begin VB.CheckBox ckUnUsed 
      Caption         =   "Unused"
      Height          =   315
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5940
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3435
      Left            =   4860
      TabIndex        =   30
      Top             =   900
      Visible         =   0   'False
      Width           =   4095
      Begin VB.ListBox lbEdIzm 
         Height          =   816
         ItemData        =   "Nomenclatura.frx":0000
         Left            =   1860
         List            =   "Nomenclatura.frx":0010
         TabIndex        =   39
         Top             =   1800
         Width           =   495
      End
      Begin VB.ListBox lbEdIzm2 
         Height          =   624
         ItemData        =   "Nomenclatura.frx":0024
         Left            =   480
         List            =   "Nomenclatura.frx":0031
         TabIndex        =   38
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox tbPerList 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2940
         TabIndex        =   40
         Text            =   "tbPerList"
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   2340
         TabIndex        =   33
         Top             =   2940
         Width           =   1095
      End
      Begin VB.CommandButton cmOk 
         Caption         =   "Ok"
         Height          =   315
         Left            =   780
         TabIndex        =   32
         Top             =   2940
         Width           =   1035
      End
      Begin VB.Label laEdIzm 
         Caption         =   "Ед.изм.пр-ва"
         Height          =   195
         Left            =   1680
         TabIndex        =   37
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Осн.ед.изм."
         Height          =   315
         Left            =   300
         TabIndex        =   36
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label laHeader 
         Caption         =   "laHeader"
         Height          =   1335
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   3915
      End
      Begin VB.Label laPerList 
         Caption         =   "К-т произ-ва"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2880
         TabIndex        =   34
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lab1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   3435
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmObrez 
      Caption         =   "Только обрезн."
      Height          =   315
      Left            =   1560
      TabIndex        =   29
      Top             =   5940
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lbWeb 
      Height          =   432
      ItemData        =   "Nomenclatura.frx":0047
      Left            =   3240
      List            =   "Nomenclatura.frx":0051
      TabIndex        =   28
      Top             =   1260
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CheckBox chPerList 
      Caption         =   "В целых"
      Height          =   195
      Left            =   7500
      TabIndex        =   27
      Top             =   40
      Width           =   1035
   End
   Begin VB.ListBox lbSource 
      Height          =   2736
      Left            =   8640
      TabIndex        =   26
      Top             =   360
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Frame frTitle 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   2880
      TabIndex        =   24
      Top             =   3060
      Visible         =   0   'False
      Width           =   7335
      Begin VB.Label laTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   "laTitle"
         Height          =   255
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   7275
      End
   End
   Begin VB.ListBox lbMark 
      Height          =   432
      ItemData        =   "Nomenclatura.frx":005C
      Left            =   3000
      List            =   "Nomenclatura.frx":0066
      TabIndex        =   23
      Top             =   3660
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.ListBox lbYesNo 
      Height          =   432
      ItemData        =   "Nomenclatura.frx":0078
      Left            =   3060
      List            =   "Nomenclatura.frx":0082
      TabIndex        =   22
      Top             =   1920
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "До"
      Height          =   315
      Left            =   1260
      TabIndex        =   21
      Top             =   5460
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "между"
      Height          =   315
      Left            =   2100
      TabIndex        =   20
      Top             =   5460
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox ckStartDate 
      Caption         =   " "
      Height          =   315
      Left            =   3720
      TabIndex        =   19
      Top             =   0
      Width           =   195
   End
   Begin VB.CheckBox ckEndDate 
      Caption         =   " "
      Height          =   315
      Left            =   5040
      TabIndex        =   1
      Top             =   0
      Width           =   195
   End
   Begin VB.ComboBox cbInside 
      Height          =   315
      Left            =   10320
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chGain 
      Caption         =   "в У.Е."
      Height          =   195
      Left            =   6360
      TabIndex        =   16
      Top             =   40
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmKlassLoad 
      Caption         =   "Загрузить"
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   5940
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmExcel 
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   7020
      TabIndex        =   14
      Top             =   5940
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmHide 
      Caption         =   "Скрыть выд."
      Height          =   315
      Left            =   5100
      TabIndex        =   13
      Top             =   5940
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Timer Timer2 
      Left            =   2160
      Top             =   4500
   End
   Begin VB.TextBox tbEndDate 
      Height          =   285
      Left            =   5280
      TabIndex        =   3
      Text            =   "23.12.02"
      Top             =   0
      Width           =   795
   End
   Begin VB.TextBox tbStartDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   2
      Text            =   "23.12.02"
      Top             =   0
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Timer Timer1 
      Left            =   1740
      Top             =   4500
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   3420
      TabIndex        =   7
      Text            =   "tbMobile"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   2880
      TabIndex        =   4
      Top             =   300
      Width           =   8895
      _ExtentX        =   15685
      _ExtentY        =   9758
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2715
      _ExtentX        =   4784
      _ExtentY        =   10181
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label lbPostav 
      Alignment       =   1  'Right Justify
      Caption         =   "Срок постав."
      Height          =   192
      Left            =   7200
      TabIndex        =   43
      Top             =   60
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label lbInside 
      Caption         =   "Внут.подразд-е:"
      Height          =   255
      Left            =   8940
      TabIndex        =   17
      Top             =   60
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label laInform 
      Caption         =   "sdsds"
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   300
      TabIndex        =   12
      Top             =   600
      Width           =   2355
   End
   Begin VB.Label laPo 
      Caption         =   "На:"
      Height          =   195
      Left            =   4740
      TabIndex        =   11
      Top             =   60
      Width           =   255
   End
   Begin VB.Label laPeriod 
      Caption         =   "Период c"
      Height          =   195
      Left            =   2940
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label laBegin 
      Height          =   1335
      Left            =   4020
      TabIndex        =   9
      Top             =   2400
      Width           =   6615
   End
   Begin VB.Label laKolvo 
      Caption         =   "Число записей:"
      Height          =   195
      Left            =   3000
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label laQuant 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4260
      TabIndex        =   5
      Top             =   5940
      Visible         =   0   'False
      Width           =   555
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
      Begin VB.Menu mnSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnCost 
         Caption         =   "Себестоимость (из Komtex)"
      End
   End
   Begin VB.Menu mnContext2 
      Caption         =   "Context2"
      Visible         =   0   'False
      Begin VB.Menu mnAdd2 
         Caption         =   "Добавть"
      End
      Begin VB.Menu mnCopy 
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
      Begin VB.Menu mnSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnKarta 
         Caption         =   "Карточка движения"
      End
      Begin VB.Menu mnKartaAdd 
         Caption         =   "Добавить к Карточке"
         Visible         =   0   'False
      End
      Begin VB.Menu mnKartaVenture 
         Caption         =   "Движения по предприятиям"
      End
      Begin VB.Menu mnKartaVentureAdd 
         Caption         =   "Добавить по предприятиям"
      End
      Begin VB.Menu mnSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnPriceHistory 
         Caption         =   "История изменения цены"
      End
   End
   Begin VB.Menu mnContext3 
      Caption         =   "Перемещение ном-ры"
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
      Caption         =   "Номенклатура"
      Visible         =   0   'False
      Begin VB.Menu mnToDoc 
         Caption         =   "Добавить к документу"
      End
      Begin VB.Menu mnToProduct 
         Caption         =   "Добавить к изделию"
      End
   End
   Begin VB.Menu mnContext6 
      Caption         =   "Заменить ном-ру"
      Visible         =   0   'False
      Begin VB.Menu mnEdit5 
         Caption         =   "Выбрать"
      End
   End
End
Attribute VB_Name = "Nomenklatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Regim As String
Public isRegimLoad As Boolean
Public isLoad As Boolean

Dim Node As Node
Dim tbKlass As Recordset
Dim mousCol As Long, mousRow As Long
Dim quantity  As Long
Dim frmMode As String
Dim flKlassAdd As Boolean
Dim beShift As Boolean
Dim oldRegim As String
Public FO As Single ' ФО
Dim dOst As Single
Dim oldCellColor As Long
Dim tbmobile_readonly As Boolean
Dim gSrokPostav As Boolean ' переменная испольуется для определения,нужно ли пересчитывать запасы и к заявке, если срок доставки изменился

'Dim tbDateisVisible As Boolean ' видны ли поля для ввода дат.
'Dim replNomNom As String ' перемещаемая номенклатура
Dim NN() As String ' перемещаемая номенклатура
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim gainC As Single, gain As Single
'Dim colHeads(10) As String

'!!!При добавлении колонки добавить переменную и в initCol -99, "", 0
Dim nkNomer As Integer
Dim nkName As Integer
Dim nkEdIzm As Integer
Dim nkEdIzm2 As Integer
Dim nkZakup As Integer
Dim nkZapas As Integer
Dim nkDeficit As Integer
Dim nkBegOstat As Integer
Dim nkCurOstat As Integer
Dim nkCheckOst As Integer
Dim nkDostup As Integer
Dim nkPrihod As Integer
Dim nkRashod As Integer
Dim nkRashodBay As Integer
Dim nkSaledProcent As Integer
Dim nkAvgOutcome As Integer
Dim nkEndOstat As Integer
Dim nkCena As Integer
Dim nkPrevCost As Integer
Dim nkSkladOst As Integer
Dim nkCENA1 As Integer
Dim nkCenaFreight As Integer
Dim nkVES As Integer
Dim nkSTAVKA As Integer
Dim nkFormulaNom As Integer
Dim nkSource As Integer
'Dim nkFormula As Integer
Dim nkPerList As Integer
Dim nkWeb As Integer
Dim nkSize As Integer
Dim nkPack As Integer
Dim nkCod As Integer
'Dim nkObrez As Integer
Dim nkMargin As Integer
Dim nkKodel As Integer
Dim nkKolonok As Integer
Dim nkCena1W As Integer
Dim nkCena2W As Integer
Dim nkKolon2 As Integer
Dim nkKolon3 As Integer
Dim nkKolon4 As Integer
Dim nkYesNo As Integer
Dim nkMark As Integer
Dim nkZakupBax As Integer
Dim nkZakupWeight As Integer


Private Sub setMnPriceHistoryStatus()
Dim cnt As String
    If nkPrevCost <> -1 Then
        cnt = Grid.TextMatrix(mousRow, nkPrevCost)
    End If
    
'    sql = "select count(*) from sPriceHistory where nomnom = '" & Grid.TextMatrix(mousRow, nkNomer) & "'"
'    byErrSqlGetValues "##05.05", sql, cnt
    If nkPrevCost <> -1 And cnt <> "--" Then
        mnPriceHistory.Visible = True
        mnSep4.Visible = True
        If mousCol = nkPrevCost Then
            mnPriceHistory.Caption = "Вернуть предыдущую цену"
        Else
            mnPriceHistory.Caption = "Просмотреть полную историю"
        End If
    Else
        mnPriceHistory.Visible = False
        mnSep4.Visible = False
    End If
        
    
End Sub

Private Sub cbInside_Click()
'If isRegimLoad Then loadKlassNomenk
If Grid.Visible Then
    gridColControl
    loadKlassNomenk
End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chGain_Click()
If noClick Then Exit Sub

If chGain.value = 1 Then
'    chPerList.Enabled = False
    chPerList.value = 0
    loadKlassNomenk
ElseIf chPerList.value = 0 Then
    loadKlassNomenk ' т.е. если сбросилась не по установке chGain
End If
On Error Resume Next
Grid.SetFocus
'Grid_EnterCell

End Sub

Private Sub chPerList_Click()
If noClick Then Exit Sub
If chPerList.value = 1 Then
'    chGain.Enabled = False
    chGain.value = 0
    loadKlassNomenk
ElseIf chGain.value = 0 Then
    loadKlassNomenk ' т.е. если сбросилась не по установке chGain
End If
    
On Error Resume Next
Grid.SetFocus

End Sub

Private Sub ckEndDate_Click()
If noClick Then Exit Sub

tbEndDate.Enabled = Not tbEndDate.Enabled
gridColControl
End Sub

Sub gridColControl()
Dim delta As Integer

delta = 825
Grid.colWidth(nkName) = 2085 '2520
If Regim = "asOborot" Or Regim = "sourOborot" Or Regim = "asOstat" Then
    If ckEndDate.value = 1 Or cbInside.ListIndex <> 1 Then
        Grid.colWidth(nkDostup) = 0
        If Regim = "asOborot" Or Regim = "sourOborot" Then
            Grid.colWidth(nkZapas) = 0
            Grid.colWidth(nkZakup) = 0
            Grid.colWidth(nkDeficit) = 0
            Grid.colWidth(nkMark) = 0
            Grid.colWidth(nkName) = 3420
        Else
            Grid.colWidth(nkName) = Grid.colWidth(nkName) + delta
        End If
        Grid.TextMatrix(0, nkEndOstat) = "Кон.Остатки"
    Else 'т.е. когда кон.дата не отмечена и установлен Склад1
        Grid.colWidth(nkDostup) = delta
        If Regim = "asOborot" Or Regim = "sourOborot" Then
            Grid.colWidth(nkZapas) = 645
            Grid.colWidth(nkZakup) = 630
            Grid.colWidth(nkDeficit) = 780
            Grid.colWidth(nkMark) = 705
        End If
        Grid.TextMatrix(0, nkEndOstat) = "Ф.Остатки"
    End If
    Grid.Visible = False
    ckUnUsed.Visible = False
End If

End Sub

Private Sub ckStartDate_Click()
If noClick Then Exit Sub
If ckStartDate.value = 0 Then
    tbStartDate.Enabled = False
Else
    tbStartDate.Enabled = True
End If
End Sub

Private Sub ckUnUsed_Click()
Dim l As Long, bColor As Long

If Not Grid.Visible Then
    ckUnUsed.value = 0
    Exit Sub
End If
bColor = &HCCCCCC
If ckUnUsed.value = 0 Then
    bColor = 0
Else
    Grid.CellBackColor = Grid.BackColor 'поскольку LeaveCell будет заблокирована
End If
For l = 1 To Grid.Rows - 1
    If Grid.TextMatrix(l, nkMark) = "Unused" Then _
                    colorGridRow Grid, l, bColor
Next l
Grid_EnterCell
On Error Resume Next
Grid.SetFocus

End Sub

Private Sub cmCancel_Click()
Frame1.Visible = False
On Error Resume Next
Grid.SetFocus
End Sub

Private Sub cmExcel_Click()
If Regim = "asOborot" Or Regim = "fltOborot" Or Regim = "sourOborot" Then
    GridToExcel Grid, "Оборотная ведомость на период с " & tbStartDate & _
    " по " & tbEndDate
ElseIf Regim = "asOstat" Then
    GridToExcel Grid, "Ведомость остатков на " & tbEndDate
ElseIf Regim = "checkCurOstat" Then
    GridToExcel Grid, "Позиции у кот. текущие остатки не совпадали с " & _
    "вычисленными на основе начальных остатков и записей из ДМЦ."
Else
    GridToExcel Grid, "Состояние номенклатуры на " & Format(Now(), "dd.mm.yy")
End If
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmHide_Click()
Dim i As Integer

For i = Grid.row To Grid.RowSel
    Grid.RemoveItem Grid.row
    quantity = quantity - 1
Next i

End Sub

Private Sub cmKlassLoad_Click()
If Regim = "fltOborot" Then ' ном-ра для закупки
    loadKlassNomenk
Else
    controlVisible False
    loadKlass
End If
End Sub


Sub initCol(curCol As Integer, colName As String, colWdth As Integer, _
Optional align As String = "")
Static i As Integer

If curCol = -99 Then
    nkNomer = -1
    nkName = -1
    nkEdIzm = -1
    nkEdIzm2 = -1
    nkZakup = -1
    nkZapas = -1
    nkDeficit = -1
    nkBegOstat = -1
    nkCurOstat = -1
    nkCheckOst = -1
    nkDostup = -1
    nkPrihod = -1
    nkRashod = -1
    nkEndOstat = -1
    nkCena = -1
    nkPrevCost = -1
    nkSkladOst = -1
    nkCENA1 = -1
    nkCenaFreight = -1
    nkVES = -1
    nkSTAVKA = -1
    nkFormulaNom = -1
    nkSource = -1
    nkPerList = -1
    nkWeb = -1
    nkSize = -1
    nkPack = -1
    nkCod = -1
'    nkObrez = -1
    nkMargin = -1
    nkKodel = -1
    nkKolonok = -1
    nkCena1W = -1
    nkCena2W = -1
    nkKolon2 = -1
    nkKolon3 = -1
    nkKolon4 = -1
    nkYesNo = -1
    nkMark = -1

    Grid.Cols = 2
    i = 0
Else
    i = i + 1
    If i > 1 Then Grid.Cols = Grid.Cols + 1
    curCol = i
End If

Grid.colWidth(i) = colWdth
If align <> "" Then Grid.ColAlignment(i) = align
Grid.TextMatrix(0, i) = colName
End Sub

Sub controlGridHight(Optional max As String = "")
Static oldTop As Integer, oldHeight As Integer

If oldTop = 0 Then
    oldTop = Grid.Top
    oldHeight = Grid.Height
End If
If max = "" Then
    Grid.Top = oldTop
    Grid.Height = oldHeight
Else
    Grid.Top = tv.Top
    Grid.Height = tv.Height
End If
End Sub

Sub cotnrolTopElementsVisible(en As Boolean)
tbEndDate.Visible = en
ckEndDate.Visible = en
If Regim = "fltOborot" Then ' ном-ра для закупки
    chGain.Visible = True
'    chPerList.Visible = True
Else
    chGain.Visible = en
'    chPerList.Visible = en
End If
chPerList.Visible = False
chGain.Enabled = False
chPerList.Enabled = False
If Regim = "asOstat" Then
    chPerList.Visible = True
    laPeriod.Visible = False
    laPo.Visible = False
    tbStartDate.Visible = False
    ckStartDate.Visible = False
    lbInside.Visible = False
'    cbInside.Visible = False
'    chGain.Visible = False
'    chPerList.Visible = False
Else
    laPeriod.Visible = en
    laPo.Visible = en
    tbStartDate.Visible = en
    ckStartDate.Visible = en
    lbInside.Visible = en
    cbInside.Visible = en
End If
End Sub

Function setWebFlags(oldVal As String, newVal As String) As Boolean

setWebFlags = False
If oldVal = newVal Then Exit Function

sql = "SELECT * FROM sGuideNomenk WHERE nomNom = '" & gNomNom & "'"
Set tbNomenk = myOpenRecordSet("##407", sql, dbOpenForwardOnly)
'Set tbNomenk = myOpenRecordSet("##407", "sGuideNomenk", dbOpenTable)
'If tbNomenk Is Nothing Then Exit Function
'tbNomenk.Index = "PrimaryKey"
'tbNomenk.Seek "=", gNomNom
'If tbNomenk.NoMatch Then msgOfEnd ("##406")
If tbNomenk.BOF Then msgOfEnd ("##406")
    
tbNomenk.Edit
tbNomenk!web = newVal
tbNomenk.Update

setWebFlags = True

EN1:
tbNomenk.Close
End Function

Sub setRegim()
Dim delta As Integer, str As String, str2 As String, i As Integer, j As Integer
frmMode = ""
flKlassAdd = False
gKlassId = 0 'необходим  для добавления класса
laQuant.Visible = False ' здесь происходит Load_Form
laKolvo.Visible = False
Grid.Visible = False
ckUnUsed.Visible = False
laInform.Caption = ""

tv.Visible = True
initCol -99, "", 0 '-99 указывает на столбец №0, т.к 0 nkXxxx м.=0
initCol nkNomer, "Номер", 960, flexAlignLeftCenter
initCol nkCod, "Код", 1050, flexAlignLeftCenter '700
initCol nkName, "Описание", 4155, flexAlignLeftCenter
initCol nkSize, "Размер", 675, flexAlignLeftCenter
If Regim <> "" Then initCol nkEdIzm, "Ед.изм.производства", 435, flexAlignLeftCenter
controlGridHight ' ниже tv

If oldRegim <> Regim Or oldRegim = "##undef##" Then
    ckStartDate.value = 1 '    переключает tbStartDate.Enabled
    noClick = True
    ckEndDate.value = 0
    chGain.value = 0
    chPerList.value = 1
    cbInside.ListIndex = 1
    noClick = False

    tbEndDate.Enabled = True
End If
chPerList.Visible = False
If Regim = "fltOborot" Then ' ном-ра для закупки
    cotnrolTopElementsVisible False
    initCol nkCena, "Цена факт.", 660
'    initCol nkEndOstat, "Ф.остатки", 870  'парам-ры уст-ся в ckEndDate_Click
    initCol nkEndOstat, "", 870  'парам-ры уст-ся в ckEndDate_Click
    initCol nkDostup, "Д.остатки", 700 '
    initCol nkZapas, "Мин.запас", 645
    initCol nkZakup, "Макс.запас", 630
    initCol nkDeficit, "К.заявке", 780
    initCol nkMark, "Маркер", 705
ElseIf Regim = "asOborot" Or Regim = "sourOborot" Then
    cotnrolTopElementsVisible True
    initCol nkCena, "Цена факт.", 660
    initCol nkBegOstat, "Нач.остатки", 675
    initCol nkPrihod, "Приход", 800
    initCol nkRashod, "Расход", 650
    initCol nkRashodBay, "Продано", 650
'    initCol nkEndOstat, "Кон.Остатки", 700  'парам-ры уст-ся в ckEndDate_Click
    initCol nkEndOstat, "", 700  'парам-ры уст-ся в ckEndDate_Click
    initCol nkDostup, "Д.остатки", 0 '
    initCol nkAvgOutcome, "Ср.расход", 500, flexAlignRightTop
    initCol nkZapas, "Мин.запас", 0
    initCol nkZakup, "Макс.запас", 0 '
    initCol nkDeficit, "К.заявке", 0 '
    initCol nkSaledProcent, "% Продаж", 500
    initCol nkMark, "Маркер", 0      '
    If Regim = "asOborot" Then
        initCol nkWeb, "Web", 450
        lbPostav.Visible = True
        tbPostav.Visible = True
        initCol nkZakupBax, "Заяв.сумма", 650
        initCol nkZakupWeight, "Заяв.вес", 650
    Else
        lbPostav.Visible = False
        tbPostav.Visible = False
    End If
    
    ckEndDate_Click ' меняет размер кол.nkName
ElseIf Regim = "asOstat" Then
    cmObrez.Visible = True
    cotnrolTopElementsVisible True
'    Grid.ColWidth(nkName) = 3435
    initCol nkPerList, "Коэф.производства", 735
    initCol nkCena, "Цена факт.", 660
    initCol nkBegOstat, "Нач.остатки", 0
    For i = 0 To Documents.lbInside.ListCount - 1
        initCol nkSkladOst, Documents.lbInside.List(i), 750
    Next i
    nkSkladOst = nkSkladOst - i + 1
    initCol nkEndOstat, "Ф.остатки", 870      'парам-ры уст-ся в ckEndDate_Click
    initCol nkDostup, "Д.остатки", 0 '
    initCol nkMark, "Маркер", 0      'только для подсветки
    ckEndDate_Click ' меняет размер кол.nkName
ElseIf Regim = "checkCurOstat" Then
    cotnrolTopElementsVisible False
    tv.Visible = False
    initCol nkBegOstat, "Нач.остатки", 675
    initCol nkCurOstat, "Ф.остатки", 930
    initCol nkCheckOst, "Провер.остатки", 660
Else
    controlGridHight "max" 'равно tv
    Grid.RowHeight(0) = 320
    cotnrolTopElementsVisible False
    Grid.colWidth(nkName) = 3105
    'initCol nkSize, "Размер", 675
    initCol nkEdIzm2, "Ед.измерения", 525, flexAlignLeftCenter
    initCol nkPack, "В упаковке", 600, flexAlignLeftCenter
    initCol nkPerList, "Коэф.производства", 735
    initCol nkEdIzm, "Ед.изм.производства", 435
    initCol nkEndOstat, "Ф.остатки", 700  'парам-ры уст-ся в ckEndDate_Click
    initCol nkDostup, "Д.остатки", 700 '
    initCol nkPrevCost, "Пред.Ф.Цена", 735
    initCol nkCena, "Цена фактическая(cenaFact)", 735
    initCol nkCENA1, "Цена поставщика(CENA1)", 795
    initCol nkVES, "Bec", 555
    initCol nkSTAVKA, "Ставка тамож.пошлины", 630
    initCol nkFormulaNom, "№ Формулы", 480
    initCol nkCenaFreight, "Цена с доставкой(CenaFreight)", 792
    initCol nkSource, "Поставшик", 945, flexAlignLeftCenter
    
    initCol nkMargin, "Маржа", 450
    initCol nkKodel, "К/Дел", 400
    initCol nkKolonok, "Колонок", 400
    initCol nkCena1W, "Ц_Продажи", 700
    initCol nkCena2W, "CenaSale", 700
    initCol nkKolon2, "Кол2", 740
    initCol nkKolon3, "Кол3", 740
    initCol nkKolon4, "Кол4", 740

    initCol nkWeb, "Web", 450
'    initCol nkObrez, "обрезков учет", 555
    initCol nkYesNo, "Сверено", 645
    initCol nkMark, "Маркер", 0      'только для подсветки
End If
If nkName > -1 Then Grid.ColAlignment(nkName) = flexAlignLeftCenter
Me.Caption = "Cправочник по номенклатуре    "
laBegin.Caption = ""

str = "Найдите требуемую номенклатуру в справочнике, затем " & _
    "в ее контектном меню (правый клик мышки) выберите команду "
str2 = vbCrLf & vbCrLf & "При необходимости повторите это для других " & _
       "позиций." & vbCrLf & vbCrLf & "В конце нажмите кнопку <Выход>."

If cmKlassLoad.Visible Then 'после режима "fltOborot"
    cmKlassLoad.Visible = False
    tv.Visible = True
    Grid.Width = Grid.Width - tv.Width
    Grid.left = Grid.left + tv.Width
End If
If Regim = "nomenkSelect" Or Regim = "singleSelect" Then 'из справ-ка гот.изделий
'    cmSel.Visible = True
    Me.Caption = "Выбор номенклатуры из Справочника для изделия '" & _
         gProduct & "'"
    laBegin.Caption = str & "'Добавить к изделию'." & str2
ElseIf Regim = "fromDocuments" Then
    Me.Caption = "Выбор номенклатуры из Справочника для документа №" & _
        numDoc & "/" & numExt
    laBegin.Caption = str & "'Добавить к документу'." & str2
ElseIf Regim = "forKartaDMC" Then
    laBegin.Caption = str & "'Карточка Движения'." & "     " & _
    "Эту команду можно применить и для выделенной(мышкой) группы позиций." & _
     vbCrLf & vbCrLf & "После загрузки Карточки в Справочнике  станет " & _
     "доступна команда 'Добавить к Карточке' (в т.ч. и для группы)."
ElseIf Regim = "asOborot" Or Regim = "sourOborot" Then
'    tbStartDate.Visible = True
'    laPeriod.Visible = True
'    laPo.Caption = " по"
    Me.Caption = "Оборотная ведомость с разбивкой по "
    If Regim = "sourOborot" Then
        Me.Caption = Me.Caption & "поставщикам"
    Else
        Me.Caption = Me.Caption & "группам"
    End If
    laBegin.Caption = "Задайте период Оборотной ведомости и Подразделения, " & _
    "затем  в левой панели выберите группу с требуемой номенклатурой."
ElseIf Regim = "fltOborot" Then
    tv.Visible = False
    Grid.left = Grid.left - tv.Width
    Grid.Width = Grid.Width + tv.Width
    cmKlassLoad.Visible = True
    Me.Caption = "Номенклатура для закупки"
    cmKlassLoad_Click
    GoTo EN1
ElseIf Regim = "asOstat" Then
    Me.Caption = "Ведомость остатков"
'    laPeriod.Visible = False
    laPo.Caption = "На:"
    laBegin.Caption = "Задайте дату Ведомости остатков и Подразделения, " & _
    "затем в левой панели выберите группу с требуемой номенклатурой."
'    tbStartDate.Visible = False
'    ckStartDate.Visible = False
ElseIf Regim = "checkCurOstat" Then
    Me.Caption = "Сверка текущих остатков по всей номенклатуре"
End If

If otlad <> "otlaD" Then tbEndDate.Text = Format(CurDate, "dd/mm/yy")
'If otlad <> "otlad" Or Regim = "asOstat" Then _
                             tbStartDate.Text = Format(begDate, "dd/mm/yy")
If otlad <> "otlaD" Or Regim = "asOstat" Then _
                             tbStartDate.Text = "01.01." & Format(CurDate, "yy")
    
'Grid.Visible = False


If Regim = "checkCurOstat" Then
    Timer2.Interval = 10 '  чтобы сначала высветилась пустая форма, не
    Timer2.Enabled = True ' дожидаясь загрузки из базы
ElseIf oldRegim <> Regim Or oldRegim = "##undef##" Then '"fltOborot" сюда не попадает
    loadKlass
End If
EN1:
oldRegim = Regim
isRegimLoad = True

End Sub

Function ostatCorr(myErr As String, delta As Single) As Boolean
'корректируем остатки
ostatCorr = False
Set tbNomenk = myOpenRecordSet(myErr, "sGuideNomenk", dbOpenTable)
If tbNomenk Is Nothing Then Exit Function
tbNomenk.index = "PrimaryKey"
tbNomenk.Seek "=", gNomNom
If Not tbNomenk.NoMatch Then
    tbNomenk.Edit
    tmpSingle = Round(tbNomenk!nowOstatki - delta, 2)
    tbNomenk!nowOstatki = tmpSingle
    tbNomenk.Update
    ostatCorr = True
End If
tbNomenk.Close
End Function


Function valueToNomencField(myErr As String, val As Variant, field As String) As Boolean
valueToNomencField = False

On Error GoTo SqlFail
sql = "SELECT * FROM sGuideNomenk WHERE nomNom = '" & gNomNom & "'"
Set tbNomenk = myOpenRecordSet(myErr, sql, dbOpenForwardOnly)
'Set tbNomenk = myOpenRecordSet(myErr, "sGuideNomenk", dbOpenTable)
'If tbNomenk Is Nothing Then Exit Function
'tbNomenk.Index = "PrimaryKey"
'tbNomenk.Seek "=", gNomNom
'If Not tbNomenk.NoMatch Then
If Not tbNomenk.BOF Then
    tbNomenk.Edit
    If TypeName(val) = "Single" And IsNumeric(tbNomenk.fields(field)) Then
        tmpSingle = Round(val - tbNomenk.fields(field), 2) 'delta
    End If
'    If field = "begOstatki" Then
'        tbNomenk!nowOstatki = Round(tbNomenk!nowOstatki + tmpSingle, 2)
    If field = "mark" And val = lbMark.List(0) Then
        tbNomenk!normZapas = Round(nomencDostupOstatki, 2) 'Д.О -> Min запас
        tbNomenk!zakup = FO 'Ф.О -> Max запас
    End If
    tbNomenk.fields(field) = val
    tbNomenk.Update
    valueToNomencField = True
End If
tbNomenk.Close
Exit Function
SqlFail:
errorCodAndMsg "##valueToNomenkField"
End Function



Private Sub cmMark_Click()

End Sub

Private Sub cmObrez_Click()

loadKlassNomenk "obrez"
End Sub

Private Sub cmOk_Click()
Dim str As String, oldPer As Single, newPer As Single, str2 As String
    
If Not isNumericTbox(tbPerList, 1) Then Exit Sub

str = lbEdIzm.Text
If str = "" Then
    str = lbEdIzm2.Text
Else
    If CSng(tbPerList.Text) < 1.01 Then
        MsgBox "Значение Коэффициента производства должно быть не менее " & _
        "1.01", , "Предупреждение"
        Exit Sub
    End If
End If
oldPer = Grid.TextMatrix(mousRow, nkPerList)
newPer = Round(tbPerList.Text, 2)
If oldPer = 1 And newPer <> 1 Then
    str2 = "'   МОЖНО": GoTo AA
ElseIf oldPer <> 1 And newPer = 1 Then
    str2 = "'   НЕЛЬЗЯ"
AA: If MsgBox("Внимание!!!  Теперь номенклатуру  '" & gNomNom & str2 & _
    " будет списывать со Склада обрезков." & vbCrLf & "В связи с этим рекомендуется " & _
    "внимательно просмотреть ее Карточку движения.", vbYesNo Or vbDefaultButton2, "Продолжить!") = vbNo Then Exit Sub
End If
sql = "UPDATE sGuideNomenk SET ed_Izmer = '" & str & _
"', perList = " & tbPerList.Text & ", ed_Izmer2 = '" & _
lbEdIzm2.Text & "' WHERE (((nomNom)='" & gNomNom & "'));"
If myExecute("##437", sql) = 0 Then
    Grid.TextMatrix(mousRow, nkPerList) = newPer
    Grid.TextMatrix(mousRow, nkEdIzm2) = lbEdIzm2.Text
    Grid.TextMatrix(mousRow, nkEdIzm) = str
End If
On Error Resume Next
Grid.SetFocus
Frame1.Visible = False
End Sub

'Private Sub cmSel_Click()

'End Sub

Private Sub Command1_Click()
Dim str As String, str2 As String

'str = getWhereDocsByDateBoxes(Me)
str2 = getWhereByDateBoxes(Me, "sDocs.xDate", begDate)
If str = str2 Then
    str = str & "   - совпадает"
Else
    str = str & "   - не совпадает с  Where = '" & str2 & "'"
End If
MsgBox "Where = '" & str & "'"

End Sub

Private Sub Command2_Click()
Dim str As String, str2 As String

'str = getWhereDocsByDateBoxes(Me, "befo")
str2 = getWhereByDateBoxes(Me, "sDocs.xDate", begDate, "befo")
If str = str2 Then
    str = str & "   - совпадает"
Else
    str = str & "   - не совпадает с  Where = '" & str2 & "'"
End If
MsgBox "Where = '" & str & "'"

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim value As String

If KeyCode = vbKeyF7 Then
    If mousCol = nkName Then
        value = InputBox("Укажите полное название или фрагмент.", _
        "Поиск в колонке 'Описание'", value)
        If value = "" Then Exit Sub
        If findExValInCol(Grid, value, nkName) > 0 Then Exit Sub
        MsgBox "В текущем списке фрагмент не найден.", , "Результат поиска"
    End If
ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF Then
    value = InputBox("Укажите номер или его фрагмент.", "Поиск номенклатуры", value)
    If value = "" Then Exit Sub
    loadKlassNomenk "F" & value
ElseIf Shift = vbCtrlMask And KeyCode = vbKeyG Then
    value = InputBox("Укажите полное название(описание) или фрагмент.", _
    "Поиск номенклатуры", value)
    If value = "" Then Exit Sub
    loadKlassNomenk "G" & value
ElseIf KeyCode = vbKeyEscape Then
    Frame1.Visible = False
    On Error Resume Next
    Grid.SetFocus
End If

End Sub

Private Sub Form_Load()
Dim i As Integer

oldHeight = Me.Height
oldWidth = Me.Width
isRegimLoad = False
oldRegim = "##undef##"
quantity = 0
cbInside.AddItem "все"
For i = 0 To Documents.lbInside.ListCount - 1
    cbInside.AddItem Documents.lbInside.List(i)
Next i

For i = 0 To Documents.lbSource.ListCount - 1
    lbSource.AddItem Documents.lbSource.List(i)
Next i
'lbSource.Height = 195 * lbSource.ListCount + 100

cbInside.ListIndex = 0
isLoad = True
End Sub

Sub loadKlass()
Dim Key As String, pKey As String, k() As String, pK()  As String
Dim i As Integer, iErr As Integer, groupText As String

If Regim = "sourOborot" Then
'    i = i
    tv.Nodes.Clear
    sql = "SELECT sGuideSource.sourceName, sGuideSource.sourceId " & _
    "From sGuideSource WHERE (((sGuideSource.sourceId)>=0));"
    Set table = myOpenRecordSet("##144", sql, dbOpenForwardOnly)
    If table Is Nothing Then Exit Sub
    While Not table.EOF
        Key = "k" & table!sourceId
        If table!sourceId = 0 Then
            Set Node = tv.Nodes.Add(, , Key, "(незаданные)")
        Else
            Set Node = tv.Nodes.Add(, , Key, table!SourceName)
        End If
            
        table.MoveNext
    Wend
    table.Close
  Exit Sub
End If

bilo = False
sql = "SELECT sGuideKlass.*  From sGuideKlass ORDER BY sGuideKlass.parentKlassId;"
Set tbKlass = myOpenRecordSet("##102", sql, dbOpenForwardOnly)
If tbKlass Is Nothing Then End
If Not tbKlass.BOF Then
'    i = i
 tv.Nodes.Clear
 Set Node = tv.Nodes.Add(, , "k0", "Классификатор")
 Node.Sorted = True
 Set Node = tv.Nodes.Add("k0", tvwChild, "all", "             ")
 Node.Sorted = True
 
 ReDim k(0): ReDim pK(0): ReDim NN(0): iErr = 0
 While Not tbKlass.EOF
    If tbKlass!KlassId = 0 Then GoTo NXT1
    Key = "k" & tbKlass!KlassId
    pKey = "k" & tbKlass!parentKlassId
    On Error GoTo ERR1 ' назначить второй проход
    Set Node = tv.Nodes.Add(pKey, tvwChild, Key, tbKlass!klassName)
    On Error GoTo 0
    Node.Sorted = True
NXT1:
    tbKlass.MoveNext
 Wend
 tv.Nodes.Item("all").Text = "00 Весь перечень"
End If
tbKlass.Close

While bilo ' необходимы еще проходы
  bilo = False
  For i = 1 To UBound(k())
    If k(i) <> "" Then
        On Error GoTo ERR2 ' назначить еще проход
        Set Node = tv.Nodes.Add(pK(i), tvwChild, k(i), NN(i))
        On Error GoTo 0
        k(i) = ""
        Node.Sorted = True
    End If
NXT:
  Next i
Wend
tv.Nodes.Item("k0").Expanded = True

'tv.SetFocus
If Regim = "" Then
    Set Node = tv.Nodes.Add("all", tvwLast, "p0", "Пересчет себестоимости номенклатуры")
    
    sql = "SELECT pbc.*, k.klassName from sPriceBulkChange pbc " _
    & " join sGuideKlass k on k.klassId = pbc.guide_klass_id " _
    & " order by xDate"
    
    Set tbKlass = myOpenRecordSet("##102.1", sql, dbOpenForwardOnly)
    If tbKlass Is Nothing Then End
    If Not tbKlass.BOF Then
     While Not tbKlass.EOF
        If tbKlass!guide_klass_id <> 0 Then
            groupText = " по группе " + tbKlass!klassName
        Else
            groupText = " по всей номенклатуре"
        End If
        
        Set Node = tv.Nodes.Add("p0", tvwChild, "p" + CStr(tbKlass!id), Format(tbKlass!xDate, "dd.mm.yyyy hh:mm") + groupText)
        tbKlass.MoveNext
     Wend
    End If
    tbKlass.Close
End If ' пересчет себестоимости

Exit Sub
ERR1:
 iErr = iErr + 1: bilo = True
 ReDim Preserve k(iErr): ReDim Preserve pK(iErr): ReDim Preserve NN(iErr)
 k(iErr) = Key: pK(iErr) = pKey: NN(iErr) = tbKlass!klassName
 Resume Next

ERR2: bilo = True: Resume NXT
End Sub


Private Sub showNomenkPath(nomnom As String)
Dim klassName As String, ret As String, KlassId As Integer, parentKlassId As Integer

    bilo = True
    sql = "select KlassId, nomnom from sguidenomenk where nomnom = '" & nomnom & "'"

    byErrSqlGetValues "##102.2", sql, KlassId
    
    While bilo
        sql = "select klassName, parentKlassId from sguideklass where klassid = " & KlassId
        If byErrSqlGetValues("##102.2", sql, klassName, parentKlassId) Then
            If KlassId = 0 Then
                bilo = False
            Else
                If Len(ret) > 0 Then ret = " / " & ret
                ret = klassName & ret
            End If
            KlassId = parentKlassId
        End If
'        Set tbKlass = myOpenRecordSet("##102.x", sql, dbOpenForwardOnly)
'        If tbKlass Is Nothing Then End
'        If Not tbKlass.BOF Then
'            If Len(ret) > 0 Then ret = " / " & ret
'            ret = tbKlass!klassName & ret
'            klassid = tbKlass!parentKlassId
'            If klassid = 0 Then bilo = False
'        End If
'        tbKlass.Close
    Wend
    textBoxInGridCell tbMobile, Grid, ret
'    tbMobile.Width = 1000

End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If WindowState = vbMinimized Then Exit Sub
On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w
tv.Height = tv.Height + h

cmKlassLoad.Top = cmKlassLoad.Top + h
laKolvo.Top = laKolvo.Top + h
laQuant.Top = laQuant.Top + h
cmHide.Top = cmHide.Top + h
cmExcel.Top = cmExcel.Top + h
'cmSel.Top = cmSel.Top + h
cmExit.Top = cmExit.Top + h
cmExit.left = cmExit.left + w
cmObrez.Top = cmObrez.Top + h
ckUnUsed.Top = ckUnUsed.Top + h
'.Left = .Left + w

End Sub

Private Sub Form_Unload(Cancel As Integer)
isRegimLoad = False
'oldRegim = Empty
End Sub

Private Sub Grid_Click()
'mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If quantity = 0 Then Exit Sub
If mousRow = 0 Then
    Grid.CellBackColor = Grid.BackColor
    
    'If mousCol > 3 And ((mousCol < nkSource And Regim = "") Or _
    '(mousCol < nkMark And Regim <> "")) Then
    '    SortCol Grid, mousCol, "numeric"
    'Else
    '    SortCol Grid, mousCol
    'End If
    'Grid.row = 1    ' только чтобы снять выделение
    'Grid_EnterCell
End If

End Sub


Private Sub Grid_DblClick()
Dim str As String, i As Integer

If Grid.CellBackColor <> &H88FF88 _
 And Not (Regim = "" And mousRow = 0 And mousCol >= nkCena2W And mousCol <= nkKolon4) _
Then
    Exit Sub
End If

If Regim = "" Then
    If mousRow = 0 Then
        textBoxInGridCell tbMobile, Grid, , 0
        Exit Sub
    End If
    If mousCol = nkEdIzm2 Or mousCol = nkEdIzm Or mousCol = nkPerList Then
        noClick = True
        gNomNom = Grid.TextMatrix(mousRow, nkNomer)
        laHeader.Caption = "Если для номенклатуры      '" & gNomNom & _
        "'   предполагается списывание со Cклада обрезков, " & _
        "кроме основной единицы измерения необходимо выбрать дополнительную (Ед.изм.производства),  а " & _
        "также коэффициент пересчета между этими единицами (К-т производства)."
        
        tbPerList.Text = Grid.TextMatrix(mousRow, nkPerList)
        If Not IsNumeric(tbPerList.Text) Then tbPerList.Text = 1
        lbEdIzm2.ListIndex = 0
        For i = 1 To lbEdIzm2.ListCount - 1
            If lbEdIzm2.List(i) = Grid.TextMatrix(mousRow, nkEdIzm2) Then _
                lbEdIzm2.ListIndex = i
        Next i
        lbEdIzm.ListIndex = 0
        For i = 1 To lbEdIzm.ListCount - 1
            If lbEdIzm.List(i) = Grid.TextMatrix(mousRow, nkEdIzm) Then _
                lbEdIzm.ListIndex = i
        Next i
        Frame1.Visible = True
        Frame1.ZOrder
        noClick = False
        lbEdIzm2.SetFocus
     'ElseIf mousCol = nkMargin Then
        'GuideFormuls.Regim = "fromNomenkW"
     '   GoTo BB
     ElseIf mousCol = nkFormulaNom Then
        GuideFormuls.Regim = "fromNomenk"
BB:     If GuideFormuls.isLoad Then Unload GuideFormuls
        GuideFormuls.Show vbModal
        If tmpStr = "" Then Exit Sub
        If mousCol = nkFormulaNom Then
          If valueToNomencField("##311", tmpStr, "formulaNom") Then
            Grid.TextMatrix(mousRow, nkFormulaNom) = tmpStr
            cenaFreight = nomenkFormula
            Grid.TextMatrix(mousRow, nkCenaFreight) = cenaFreight
            Grid.TextMatrix(mousRow, 0) = tmpStr 'теперь это сама формула
            'GoTo DD ' может влиять и на WEB значения
          End If
        End If
     ElseIf mousCol = nkYesNo Then
        listBoxInGridCell lbYesNo, Grid, "select"
     ElseIf mousCol = nkSource Then
        listBoxInGridCell lbSource, Grid, "select"
     ElseIf mousCol = nkWeb Then
        listBoxInGridCell lbWeb, Grid, "select"
     Else
     'ElseIf Products.Regim <> "onlyGuide" Then !!! Зачем это было(оно мешает - если открывалось окно справочника Гот.изделий
        textBoxInGridCell tbMobile, Grid
     End If
    'ElseIf Regim = "asOborot" Or Regim = "sourOborot" Then
    ElseIf Regim = "asOborot" Then
        If mousCol <> nkWeb Then GoTo AA
        listBoxInGridCell lbWeb, Grid, "select"
    ElseIf Regim = "sourOborot" Then
AA:
     If mousCol = nkMark Then
        listBoxInGridCell lbMark, Grid, "select"
     ElseIf mousCol = nkDostup Then
        dOst = Round(nomencDostupOstatki, 2)
        If Round(FO, 2) > dOst Then
            If MsgBox("Если Вы хотите просмотреть список всех заказов, под " & _
            "которые была зарезервирована эта номенклатура, нажмите <Да>.", _
            vbYesNo Or vbDefaultButton2, "Посмотреть, кто резервировал? '" & _
            gNomNom & "' ?") = vbYes Then
                Report.Regim = "whoRezerved"
                Set Report.Caller = Me
                Report.Sortable = True
                Report.Show vbModal
            End If
        Else
            MsgBox "Эта номенклатура никем не резервировалась.", , ""
        End If
     Else
        textBoxInGridCell tbMobile, Grid
     End If
End If
End Sub


Private Function inKolonYellow() As Boolean
    Dim kolonok As Integer
    kolonok = CInt(Grid.TextMatrix(mousRow, nkKolonok))
    inKolonYellow = False
    If kolonok > 0 Then
        Dim greenKolon As Integer
        greenKolon = nkKolon2 + kolonok - 2
        If greenKolon <> mousCol Then
            inKolonYellow = True
        End If
    Else
        'ручной ввод оптовых цен
        If mousCol - nkKolon2 >= Abs(kolonok) - 1 Then
            inKolonYellow = True
        End If
    End If
End Function


Private Sub Grid_EnterCell()
'If quantity = 0 Or Not cmAdd.Visible Or frmMode <> "" Then Exit Sub

If noClick Then Exit Sub
If quantity > 0 And frmMode = "" Then
 mousRow = Grid.row
 mousCol = Grid.col
 setMnPriceHistoryStatus
 gNomNom = Grid.TextMatrix(mousRow, nkNomer) 'нужно и для правой Mouse
 
 oldCellColor = Grid.CellBackColor
 
 If ((chGain.Visible And chGain.value > 0) _
 Or (chPerList.Visible And chPerList.value > 0)) Then
    Exit Sub
 End If
 If Regim = "" Then
    If mousCol = nkPrevCost Or mousCol = nkCena1W _
        Or (mousCol >= nkKolon2 And mousCol <= nkKolon4 And inKolonYellow) _
    Then
        Grid.CellBackColor = vbYellow
        frTitle.Visible = False
        Exit Sub
    Else
        If mousCol = nkCenaFreight Then
            laTitle.Caption = "CenaFreight = " & Grid.TextMatrix(mousRow, 0) & " "
            frTitle.Top = Grid.CellTop + Grid.CellHeight + 50
            frTitle.Visible = True
            frTitle.ZOrder
            Grid.CellBackColor = vbYellow
            Exit Sub
        Else
            frTitle.Visible = False
        End If
    End If
 End If
 
 If Regim = "forKartaDMC" Or Regim = "nomenkSelect" Or Regim = "singleSelect" Or Regim = "fromDocuments" Then
    Grid.CellBackColor = vbButtonFace
    Exit Sub
 ElseIf Regim = "" Then
    Grid.CellBackColor = &H88FF88
    GoTo BB
 ElseIf Regim = "asOstat" Or Regim = "fltOborot" Or Regim = "checkCurOstat" Then
    Exit Sub
 ElseIf (Regim = "asOborot" Or Regim = "sourOborot") And (mousCol = nkDostup _
 Or mousCol = nkMark Or ((mousCol = nkZakup Or mousCol = nkZapas) And _
 Grid.TextMatrix(mousRow, nkMark) = lbMark.List(0)) Or mousCol = nkWeb) Then
    Grid.CellBackColor = &H88FF88
 Else
    Grid.CellBackColor = vbYellow
 End If
BB:
 If mousCol = nkNomer Or mousCol = nkCod Then
    tbMobile.MaxLength = 20
 ElseIf mousCol = nkName Then
    tbMobile.MaxLength = 50
 Else
    tbMobile.MaxLength = 10
 End If
' tbInform.MaxLength =tbMobile.MaxLength
End If
End Sub

Private Sub Grid_GotFocus()
'    cmSel.Enabled = True
    cmHide.Enabled = True
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyReturn Then
    If Shift = vbAltMask Then
        showNomenkPath (Grid.TextMatrix(mousRow, nkNomer))
        tbmobile_readonly = True
    Else
        Grid_DblClick
    End If
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub
Sub lbHide()
tbMobile.Visible = False
Frame1.Visible = False
lbYesNo.Visible = False
lbMark.Visible = False
lbSource.Visible = False
lbWeb.Visible = False
Grid.Enabled = True
On Error Resume Next
Grid.SetFocus
Grid_EnterCell
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then Grid_EnterCell


End Sub

Private Sub Grid_LeaveCell()
'If Not noClick Then Grid.CellBackColor = Grid.BackColor
If noClick Then Exit Sub
If ckUnUsed.value = 0 Then
    Grid.CellBackColor = Grid.BackColor
Else
    Grid.CellBackColor = oldCellColor
End If
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

If Grid.MouseRow = 0 And Shift = 2 Then
        MsgBox "ColWidth = " & Grid.colWidth(Grid.MouseCol)

ElseIf frmMode = "nomenkReplace" Then
    Me.PopupMenu mnContext4
'If Button = 2 And frmMode = "" Then
ElseIf Button = 2 And frmMode = "" Then
    mnKarta.Visible = True
    mnKartaAdd.Visible = False
    If KartaDMC.isLoad Then
        If KartaDMC.Grid.Visible Then mnKartaAdd.Visible = True
    End If
    If quantity > 0 And Grid.row <> Grid.RowSel Then
        ReDim NN(Abs(Grid.RowSel - Grid.row) + 1)
        For i = Grid.row To Grid.RowSel
            NN(i - Grid.row + 1) = Grid.TextMatrix(i, nkNomer) 'только для перемещения
        Next i
'        mnKarta.Visible = False
'        mnKartaAdd.Visible = False
        mnAdd2.Visible = False
        GoTo CC
    ElseIf Regim = "forKartaDMC" And quantity > 0 Then
        mnAdd2.Visible = False
        graySelect "multi"
        GoTo AA
'    ElseIf Regim = "" Then
'    ElseIf Regim = "" Or Regim = "asOstat" Then
    ElseIf Regim = "singleSelect" And quantity > 0 Then
        mnToProduct.Visible = True
        mnToDoc.Visible = False
        graySelect
        Me.PopupMenu mnContext6
    ElseIf Regim = "nomenkSelect" And quantity > 0 Then
        mnToProduct.Visible = True
        mnToDoc.Visible = False
        graySelect
        Me.PopupMenu mnContext5
    ElseIf Regim = "fromDocuments" And quantity > 0 Then
        mnToDoc.Visible = True
        mnToProduct.Visible = False
        graySelect
        Me.PopupMenu mnContext5
    ElseIf Regim <> "checkCurOstat" Then
        mnAdd2.Visible = True
        If quantity > 0 Then
         If IsNumeric(gKlassId) Then  '$$4 добавлять и прочее ко всему переченю нельзя
            mnDel2.Visible = True
            mnSep2.Visible = True
            mnRepl2.Visible = True
            mnSep3.Visible = True
            mnCopy.Visible = True
          Else                      '$$4
            mnDel2.Visible = False  '
            mnSep2.Visible = False  '
            mnRepl2.Visible = False '
            mnSep3.Visible = False  '
            mnCopy.Visible = False  '
            mnAdd2.Visible = False  '
          End If
          graySelect
          GoTo BB
        Else
            If Not IsNumeric(gKlassId) Then Exit Sub '$$4добавлять и прочее ко всему переченю нельзя
            mnKarta.Visible = False
            mnKartaAdd.Visible = False
AA:         mnRepl2.Visible = False
CC:         mnDel2.Visible = False
            mnSep2.Visible = False
            mnSep3.Visible = False
            mnCopy.Visible = False
BB:         'graySelect
            Timer1.Interval = 10    'mnContext2
            Timer1.Enabled = True   '
        End If
    End If
End If

End Sub

Private Sub graySelect(Optional Regim As String = "")
If Regim = "" Then
    Grid.col = Grid.MouseCol
    Grid.row = Grid.MouseRow
Else
End If
Grid.CellBackColor = vbButtonFace
gNomNom = Grid.TextMatrix(Grid.row, nkNomer)
ReDim NN(1): NN(1) = gNomNom
On Error Resume Next
Grid.SetFocus
End Sub


'Sub setMinMaxZapas(row As Long)
'Dim s As Single
'
'     Grid.TextMatrix(row, nkZapas) = Grid.TextMatrix(row, nkDostup)
'     If ckEndDate.value = 1 And DateDiff("d", tbEndDate.Text, Now()) > 0 Then
'        sql = "SELECT sGuideNomenk.nowOstatki From sGuideNomenk " & _
'        "WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
'        If byErrSqlGetValues("##318", sql, s) Then _
'            Grid.TextMatrix(row, nkZakup) = Round(s, 2) 'Макс.запас
'     Else
'        Grid.TextMatrix(row, nkZakup) = Grid.TextMatrix(row, nkEndOstat)
'     End If
'End Sub




Private Sub lbEdIzm2_Click()
If noClick Then Exit Sub
lbEdIzm.ListIndex = 0
tbPerList.Text = 1
End Sub

Private Sub lbEdIzm_Click()
If lbEdIzm.ListIndex = 0 Then
    laPerList.Enabled = False
    tbPerList.Enabled = False
    If Not noClick Then tbPerList.Text = 1
Else
    laPerList.Enabled = True
    tbPerList.Enabled = True
End If
End Sub

Private Sub lbMark_DblClick()

If lbMark.Text = lbMark.List(1) Then
    'If MsgBox("При выборе значения '" & lbMark.List(1) & "' будут затерты " & _
    "старые значения  колонок 'Мин.' и 'Макс.запас'!", vbYesNo Or _
    vbDefaultButton2, "Продолжить ?") = vbNo Then GoTo EN1
End If
If valueToNomencField("##153", lbMark.Text, "mark") Then
    Grid.TextMatrix(mousRow, nkMark) = lbMark.Text
    If Regim = "asOborot" Then
        If lbMark.Text = lbMark.List(0) Then
            'used
            
            Dim ves As Variant, normZapas As Variant, zakup As Variant, Cena1 As Variant
            sql = "select ves, normZapas, zakup, cena1 from sguidenomenk where nomnom = '" & Grid.TextMatrix(mousRow, nkNomer) & "'"
            byErrSqlGetValues "##102.2", sql, ves, normZapas, zakup, Cena1

            Dim outcome As Single, strOutcome As String
            strOutcome = Grid.TextMatrix(mousRow, nkAvgOutcome)
            If right(strOutcome, 1) = "*" Then
                strOutcome = left(strOutcome, Len(strOutcome) - 1)
            End If
            If IsNumeric(strOutcome) Then
                outcome = CSng(strOutcome)
            Else
                outcome = 0
            End If
            
            recaluculateZakup mousRow, outcome, CSng(Grid.TextMatrix(mousRow, nkDostup)) _
                , Cena1, CVar(lbMark.Text), ves, normZapas, zakup
            
        Else
            Grid.TextMatrix(mousRow, nkZakupBax) = ""
            Grid.TextMatrix(mousRow, nkZakupWeight) = ""
            'Grid.TextMatrix(mousRow, nkZakup) = "0"
            'Grid.TextMatrix(mousRow, nkZapas) = "0"
            Grid.TextMatrix(mousRow, nkDeficit) = "0"
        End If
    Else
        calcZacup mousRow
    End If
End If
EN1:
lbHide

End Sub

Private Sub lbMark_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbMark_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub



Private Sub lbObrez_Click()

End Sub

Private Sub lbSource_DblClick()
sql = " UPDATE sGuideNomenk, sGuideSource SET sGuideNomenk.sourId = " & _
"[sGuideSource].[sourceId] WHERE (((sGuideNomenk.nomNom)='" & gNomNom & _
"') AND (([sGuideSource].[sourceName])='" & lbSource.Text & "'));"
'MsgBox sql
If myExecute("##319", sql) = 0 Then Grid.Text = lbSource.Text
lbHide

End Sub

Private Sub lbSource_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbSource_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub lbWeb_DblClick()
'If ValueToTableField("##359", "'" & lbWeb & "'", "sGuideNomenk", _
"web", "byNomNom") Then Grid.TextMatrix(mousRow, nkWeb) = lbWeb.Text

If setWebFlags(Grid.TextMatrix(mousRow, nkWeb), lbWeb.Text) Then _
    Grid.TextMatrix(mousRow, nkWeb) = lbWeb.Text
lbHide

End Sub

Private Sub lbWeb_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbWeb_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub lbYesNo_DblClick()
    If ValueToTableField("##153", "'" & lbYesNo & "'", "sGuideNomenk", _
    "YesNo", "byNomNom") Then Grid.TextMatrix(mousRow, nkYesNo) = lbYesNo.Text
lbHide

End Sub

Private Sub lbYesNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbYesNo_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub


Private Sub mnAdd_Click()
Static i As Integer
Dim str  As String, id As Integer
controlVisible False
i = i + 1
str = "новый " & i
'cmClassAdd.Enabled = False
'flKlassAdd = True

wrkDefault.BeginTrans
sql = "UPDATE sGuideKlass SET klassId = klassId WHERE klassId=0"
myBase.Execute sql

sql = "SELECT max(klassId) FROM sGuideKlass"
If Not byErrSqlGetValues("##460", sql, id) Then GoTo ERR1
id = id + 1
'sql = "SELECT sGuideKlass.klassId, sGuideKlass.klassName, " & _
'"sGuideKlass.parentKlassId From sGuideKlass ORDER BY sGuideKlass.klassId;"
'Set tbKlass = myOpenRecordSet("##106", sql, dbOpenDynaset)
'If tbKlass Is Nothing Then Exit Sub
'If tbKlass.BOF Then
'    id = 1
'Else
'    tbKlass.MoveLast
'    id = tbKlass!klassId + 1
'End If
'On Error GoTo ERR1
sql = "INSERT INTO sGuideKlass (klassId, klassName, parentKlassId) " & _
" values (" & id & ", '" & str & "', " & gKlassId & ")"
'MsgBox sql
If myExecute("##106", sql) <> 0 Then GoTo ERR1

wrkDefault.CommitTrans
'tbKlass.AddNew
'tbKlass!klassId = id
'tbKlass!klassName = str
'tbKlass! = gKlassId
'tbKlass.Update

'tbKlass.Close
Set Node = tv.Nodes.Add(tv.SelectedItem.Key, tvwChild, "k" & id, str)
tv.Nodes("k" & id).EnsureVisible
tv.Nodes("k" & id).Selected = True
tv.StartLabelEdit
Exit Sub

ERR1:
errorCodAndMsg ("##106")

End Sub

Sub nomenkAdd(Optional obraz As String = "")
Dim str As String, obrazRow As Long, choise As Integer  ', strNom As String

'choise = MsgBox("Если вы нажмете 'Да', то в колонке 'Обрезков учет' эта " & _
'"номенклатура будет помечена 'Да'. Это означает, что на производстве она будет " & _
'"списываться со Склада обрезков.", vbYesNoCancel + vbDefaultButton3, "Нужен ли учет обрезков?")
'If choise = vbCancel Then Exit Sub
frmMode = "nomenkAdd"
If obraz <> "" Then frmMode = "nomenkAddObraz"

Grid.CellBackColor = vbWhite

obrazRow = mousRow
If quantity > 0 Then Grid.AddItem ""

Grid.row = Grid.Rows - 1
mousRow = Grid.Rows - 1
Grid.col = nkYesNo
Grid.Text = lbYesNo.List(0)
Grid.col = nkNomer
mousCol = nkNomer
textBoxInGridCell tbMobile, Grid
'If choise = vbYes Then Grid.TextMatrix(mousRow, nkObrez) = "Да"
Grid.TextMatrix(mousRow, nkPerList) = 1
Grid.TextMatrix(mousRow, nkCena1W) = "error: Формула не задана"
Grid.TextMatrix(mousRow, nkCenaFreight) = "error: Формула не задана"
If obraz <> "" Then
    Grid.TextMatrix(mousRow, nkName) = Grid.TextMatrix(obrazRow, nkName)
    Grid.TextMatrix(mousRow, nkEdIzm) = Grid.TextMatrix(obrazRow, nkEdIzm)
    Grid.TextMatrix(mousRow, nkEdIzm2) = Grid.TextMatrix(obrazRow, nkEdIzm2)
    Grid.TextMatrix(mousRow, nkPerList) = Grid.TextMatrix(obrazRow, nkPerList)
    Grid.TextMatrix(mousRow, nkPack) = Grid.TextMatrix(obrazRow, nkPack)
    Grid.TextMatrix(mousRow, nkCena) = Grid.TextMatrix(obrazRow, nkCena)
    Grid.TextMatrix(mousRow, nkCENA1) = Grid.TextMatrix(obrazRow, nkCENA1)
    Grid.TextMatrix(mousRow, nkVES) = Grid.TextMatrix(obrazRow, nkVES)
    Grid.TextMatrix(mousRow, nkSTAVKA) = Grid.TextMatrix(obrazRow, nkSTAVKA)
    Grid.TextMatrix(mousRow, 0) = Grid.TextMatrix(obrazRow, 0) 'сама формула
    Grid.TextMatrix(mousRow, nkFormulaNom) = Grid.TextMatrix(obrazRow, nkFormulaNom)
    Grid.TextMatrix(mousRow, nkCenaFreight) = Grid.TextMatrix(obrazRow, nkCenaFreight)
'    Grid.TextMatrix(mousRow, nkWebFormula) = Grid.TextMatrix(obrazRow, nkWebFormula)
    Grid.TextMatrix(mousRow, nkMargin) = Grid.TextMatrix(obrazRow, nkMargin)
    Grid.TextMatrix(mousRow, nkKodel) = Grid.TextMatrix(obrazRow, nkKodel)
    Grid.TextMatrix(mousRow, nkKolonok) = Grid.TextMatrix(obrazRow, nkKolonok)
    Grid.TextMatrix(mousRow, nkCena1W) = Grid.TextMatrix(obrazRow, nkCena1W)
    Grid.TextMatrix(mousRow, nkCena2W) = Grid.TextMatrix(obrazRow, nkCena2W)
    Grid.TextMatrix(mousRow, nkKolon2) = Grid.TextMatrix(obrazRow, nkKolon2)
    Grid.TextMatrix(mousRow, nkKolon3) = Grid.TextMatrix(obrazRow, nkKolon3)
    Grid.TextMatrix(mousRow, nkKolon4) = Grid.TextMatrix(obrazRow, nkKolon4)
    Grid.TextMatrix(mousRow, nkSource) = Grid.TextMatrix(obrazRow, nkSource)
    Grid.TextMatrix(mousRow, nkSize) = Grid.TextMatrix(obrazRow, nkSize)
    Grid.TextMatrix(mousRow, nkCenaFreight) = Grid.TextMatrix(obrazRow, nkCenaFreight)
    tbMobile.Text = gNomNom
    tbMobile.SelStart = Len(gNomNom)
End If
On Error Resume Next
Grid.SetFocus

End Sub

Private Sub mnAdd2_Click()
    nomenkAdd
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
'    Grid.CellForeColor = vbBlack
    Grid.CellBackColor = Grid.BackColor
    On Error Resume Next
    tv.SetFocus

End Sub

Private Sub mnCancel4_Click()
    frmMode = ""
    Me.MousePointer = flexDefault
'    Grid.CellForeColor = vbBlack
    Grid.CellBackColor = Grid.BackColor
    On Error Resume Next
    tv.SetFocus

End Sub

Private Sub mnCopy_Click()
nomenkAdd "obraz"
End Sub

Private Sub mnCost_Click()
Dim msgOk As VbMsgBoxResult
Dim tvKlass As String
Dim tvNode As Node
Dim KlassId As String
Dim newKey As String, xDate As String, groupText As String
Dim queryTimeout As Variant


    msgOk = MsgBox("Вы уверены, что хотите пересчитать себестоимость?" _
        , vbOKCancel, "Предупреждение")
    If msgOk <> vbOK Then
        Exit Sub
    End If
    
    MousePointer = flexHourglass
    Set tvNode = tv.Nodes(tv.SelectedItem.index)
    KlassId = Mid(tvNode.Key, 2)
    
    sql = "select wf_cost_bulk_change( " & KlassId & ")"
    queryTimeout = myBase.queryTimeout
    myBase.queryTimeout = 6000
    
    byErrSqlGetValues "##nomenklatura.1", sql, newKey
    myBase.queryTimeout = queryTimeout
    If newKey > 0 Then
        sql = "select xDate from sPriceBulkChange where id = " & newKey
        byErrSqlGetValues "##nomenklatura.3", sql, xDate
        If KlassId <> "0" Then
            groupText = " по группе " + tvNode.Text
        Else
            groupText = " по всей номенклатуре группе "
        End If
        Set tvNode = tv.Nodes.Add("p0", tvwChild, "p" + newKey, CStr(xDate) + groupText)
        
        MsgBox "Себестоимоть номенклатуры по категории '" & tvNode.Text & "' успешно пересчитана." _
            , vbOKOnly, "Сообщение"
        tvNode.EnsureVisible
        tv.SelectedItem = tvNode
        tv_NodeClick tvNode
    ElseIf newKey < 0 Then
        MsgBox "Цена фактическая (и история) по некоторым или всем позициям " & _
        vbCr & "в категории '" & tvNode.Text & "' обнулились из-за отсутсвия движения." _
            , vbOKOnly, "Предупреждение"
    Else
        MsgBox "Себестоимость номенклатуры по категории '" & tvNode.Text & "' не изменилась." _
            , vbOKOnly, "Сообщение"
    End If
    
    MousePointer = flexDefault
End Sub

Private Sub mnDel_Click()
Dim i As Integer
If MsgBox("Для удаления класса  нажмите <Да>." & Chr(13) & Chr(13) & _
"Удаление возможно, если класс не содержит элементов и других подклассов", _
vbYesNo Or vbDefaultButton2, "Удалить '" & tv.SelectedItem.Text & _
"'. Вы уверены?") = vbNo Then GoTo EN1

sql = "DELETE from sGuideKlass " & _
      "WHERE (((sGuideKlass.klassId)=" & gKlassId & "));"
i = myExecute("##114", sql, -198)
If i = -2 Then
     MsgBox "Нельзя удалять непустой класс, сначала удалите входящие в " & _
     "него элементы.", , "Удаление невозможно !"
    tv.SetFocus
Else
    tv.Nodes.Remove tv.SelectedItem.Key
    controlVisible False
End If
EN1:
End Sub

Private Sub mnDel2_Click()
Grid.CellForeColor = vbBlack
Grid.CellBackColor = vbWhite
  If MsgBox("После нажатия <Да> данный элемент будет удален из Справочника", _
  vbDefaultButton2 Or vbYesNo, "Удалить '" & gNomNom & "'. Вы уверены?") = vbYes Then
    sql = "delete from  sGuideNomenk WHERE nomNom ='" & gNomNom & "'"
Debug.Print sql

    myExecute "##delete nomnom", sql
    quantity = quantity - 1
    If quantity = 0 Then
        clearGridRow Grid, mousRow
    Else
        Grid.RemoveItem mousRow
    End If
  End If
GoTo EN1

ERR1:
If errorCodAndMsg("##164", -198) Then
    MsgBox "Номенклатура задействована либо в каком-то документе либо " & _
    "в каком-то изделии из Cправочника изделий, поэтому сначала " & _
    "необходимо удалить ее оттуда.", , "Позиция '" & gNomNom & _
    "' задействована!"
End If
EN1:
Grid_EnterCell
On Error Resume Next
Grid.SetFocus
End Sub

Private Sub mnEdit5_Click()
wrkDefault.BeginTrans


sql = "DELETE From sProducts WHERE (((sProducts.ProductId)=" & gProductId & ") " & _
"AND ((sProducts.nomNom)='" & Products.soursNom & "'));"
'MsgBox sql
If myExecute("##315", sql) <> 0 Then GoTo ER:

'myBase.Execute sql

If toProduct Then
    wrkDefault.CommitTrans
Else
ER: wrkDefault.Rollback
End If
Unload Me
End Sub

Private Sub mnInsert_Click()
Dim str As String, i As Integer
    
    frmMode = ""
    Grid.CellForeColor = vbBlack
    Grid.CellBackColor = vbWhite
'    gNomNom = replNomNom
    Me.MousePointer = flexDefault
    str = Mid$(tv.SelectedItem.Key, 2)
For i = 1 To UBound(NN)
    gNomNom = NN(i)
    ValueToTableField "##112", str, "sGuideNomenk", "klassId", "byNomNom"
Next i
    tv_NodeClick tv.SelectedItem
    On Error Resume Next
    tv.SetFocus

End Sub

Private Sub mnKarta_Click()
Dim i As Integer, lenght As Integer

Grid.CellBackColor = vbWhite
KartaDMC.Grid.Visible = False
KartaDMC.quantity = 0
lenght = Grid.RowSel - Grid.row + 1
If lenght = 1 Then
'    KartaDMC.cmCheck.Visible = True
    KartaDMC.controlVisible True
    KartaDMC.nomenkName = Grid.TextMatrix(Grid.row, nkName)
Else
    KartaDMC.controlVisible False
End If
ReDim DMCnomNom(lenght)
For i = 1 To lenght
    DMCnomNom(i) = Grid.TextMatrix(Grid.row + i - 1, nkNomer)
Next i
i = UBound(DMCnomNom)

KartaDMC.Show
End Sub

Private Sub mnKartaAdd_Click()
Dim i As Integer, str As String, lenght As Integer, newLen As Integer
Dim j As Integer, l As Long
Grid.CellBackColor = vbWhite
If KartaDMC.cmCheck.Visible Then KartaDMC.removeHead
KartaDMC.controlVisible False

'Добавляем новые эл-ты
lenght = UBound(DMCnomNom)
newLen = lenght
Me.MousePointer = flexHourglass
'KartaDMC.Grid.Visible = False
For i = Grid.row To Grid.RowSel
    str = Grid.TextMatrix(i, nkNomer)
    For j = 1 To lenght ' может этот эл-т был уже добавлен
        If DMCnomNom(j) = str Then GoTo NXT
    Next j
    newLen = newLen + 1
    ReDim Preserve DMCnomNom(newLen)
    DMCnomNom(newLen) = str ' чтобы корректно работал перерасчет Карты после правки в документе
    KartaDMC.getKartaDMC str
NXT:
Next i
KartaDMC.Show
'KartaDMC.Grid.Visible = True
Me.MousePointer = flexDefault
End Sub

Private Sub mnKartaVenture_Click()
Dim selectedRows As Integer
Dim i As Integer
Dim curRow As Integer, startRow As Integer, stopRow As Integer
    
    selectedRows = Abs(Grid.row - Grid.RowSel) + 1
    ReDim DMCnomNom(selectedRows)
    
    If Grid.row >= Grid.RowSel Then
        startRow = Grid.RowSel
        stopRow = Grid.row
    Else
        startRow = Grid.row
        stopRow = Grid.RowSel
    End If
    
    i = 0
    curRow = Grid.row
    For curRow = startRow To stopRow
        DMCnomNom(i + 1) = Grid.TextMatrix(curRow, nkNomer)
        i = i + 1
    Next curRow
    
    Me.MousePointer = flexHourglass
    If VentureHistory.ckPerList.value <> 1 Then
        VentureHistory.ckPerList.value = 1
    Else
        VentureHistory.fillGrid
    End If
    VentureHistory.Show
    Me.MousePointer = flexDefault
    
End Sub

Private Sub mnPriceHistory_Click()
Dim stRow As Long, enRow As Long, i As Integer
Dim lNomnom As String, oldPrevCost As String, newPrevCost As Variant
    If mousCol <> nkPrevCost Then
    ' История изменения цены
        PriceHistory.Show
    Else
        If Grid.RowSel < mousRow Then
            stRow = Grid.RowSel
            enRow = mousRow
        Else
            stRow = mousRow
            enRow = Grid.RowSel
        End If
        If MsgBox("Вы уверены, что хотите вернуть предыдущую цену?", vbYesNo Or vbDefaultButton2, "Подтверждение") <> vbYes Then Exit Sub
        For i = stRow To enRow
            lNomnom = Grid.TextMatrix(i, nkNomer)
            oldPrevCost = Grid.TextMatrix(i, nkPrevCost)
            newPrevCost = Null
            If oldPrevCost <> "--" Then
                sql = "select wf_price_revert ( '" & lNomnom & "', " & oldPrevCost & ")"
                byErrSqlGetValues "##price_revert", sql, newPrevCost
                If IsNull(newPrevCost) Then
                    newPrevCost = "--"
                End If
                Grid.TextMatrix(i, nkCena) = oldPrevCost
                Grid.TextMatrix(i, nkPrevCost) = newPrevCost
            End If
        Next i
    End If
End Sub

Private Sub mnRen_Click()
tv.StartLabelEdit
End Sub

Private Sub mnRepl_Click()
Dim str As String
str = tv.SelectedItem.Key
If frmMode = "" Then
    If str = "all" Or str = "k0" Then Exit Sub
    frmMode = "klassReplace"
    mnRepl.Caption = "Вставить"
    mnAdd.Visible = False
    mnRen.Visible = False
    mnDel.Visible = False
    mnSep.Visible = False
    mnCancel.Visible = True
    Me.MousePointer = flexUpArrow
    nodeKey = str
ElseIf frmMode = "klassReplace" Then
    frmMode = ""
    mnRepl.Caption = "Переместить"
    mnAdd.Visible = True
    mnRen.Visible = True
    mnDel.Visible = True
    mnSep.Visible = True
    mnCancel.Visible = False
    Me.MousePointer = flexDefault
    controlVisible False

    If str = nodeKey Then
        MsgBox "Нельзя переместить класс сам в себя", , "Недопустимая операция!"
    Else
        sql = "UPDATE sGuideKlass SET sGuideKlass.parentKlassId = " & _
        Mid$(str, 2) & " WHERE (((sGuideKlass.klassId)=" & Mid$(nodeKey, 2) & "));"
'        MsgBox sql
        myBase.Execute sql
        loadKlass
    End If
    
ElseIf frmMode = "nomenkReplace" Then
MsgBox "неиспользуемый алгоритм", , "Err ##999"
End
'    frmMode = ""
'    mnRepl.Caption = "Переместить"
'    mnAdd.Visible = True
'    mnRen.Visible = True
'    mnDel.Visible = True
 '   mnSep.Visible = True
'    mnCancel.Visible = False
'    Me.MousePointer = flexDefault
'    str = Mid$(tv.SelectedItem.key, 2)
'    ValueToTableField "##112", str, "sGuideNomenk", "klassId", "byNomNom"
'    tv_NodeClick tv.SelectedItem
'    tv.SetFocus
End If
End Sub

'Private Sub mnReplace2_Click()
'Grid.CellForeColor = vbBlack
'Grid.CellBackColor = vbWhite
'mousRight = 0

'End Sub

Private Sub mnRepl2_Click()
Dim str As String, id As Integer

    Me.MousePointer = flexUpArrow
 '   replNomNom = gNomNom
    On Error Resume Next
    tv.SetFocus
    frmMode = "nomenkReplace"

End Sub

Private Sub mnToDoc_Click()
Dim id As Integer, per As Single

gNomNom = Grid.TextMatrix(mousRow, nkNomer)
sql = "SELECT sGuideNomenk.perList, sDocs.destId, sGuideSource.sourceName " & _
"FROM sGuideNomenk, sGuideSource INNER JOIN sDocs ON sGuideSource.sourceId = sDocs.destId " & _
"WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "') AND ((sDocs.numDoc)=" & _
numDoc & ") AND ((sDocs.numExt)= " & numExt & "));"
If Not byErrSqlGetValues("##424", sql, per, id, tmpStr) Then Exit Sub
'If (id = -1002 And str = "") Then
If (id = -1002 And per = 1#) Then
   MsgBox "Эта номенклатура не должна приходоваться на склад '" & tmpStr & "'", , "Error"
   Exit Sub
End If

Set tbDMC = myOpenRecordSet("##117", "sDMC", dbOpenTable) '
If tbDMC Is Nothing Then Exit Sub
tbDMC.AddNew
tbDMC!nomnom = gNomNom
tbDMC!numDoc = numDoc
tbDMC!numExt = numExt
On Error GoTo ERR1
tbDMC.Update
tbDMC.Close
Exit Sub

ERR1:
tbDMC.Close
If errorCodAndMsg("##138", -193) Then
'If Err = 3022 Then
    MsgBox "Одна и та же номенклатура может присутствовать в одном " & _
    "и том же документе не более одного раза.", , "Позиция '" & gNomNom & _
    "' уже есть!"
'Else
'    MsgBox Error, , "Ошибка 138-" & Err & ":  " '##138
End If
'Grid.SetFocus

End Sub

Private Sub mnToProduct_Click()
toProduct
End Sub

Function toProduct() As Boolean

toProduct = False
Set tbProduct = myOpenRecordSet("##117", "sProducts", dbOpenTable) '
If tbProduct Is Nothing Then Exit Function
tbProduct.AddNew
tbProduct!productId = gProductId
tbProduct!nomnom = gNomNom
On Error GoTo ERR1
tbProduct.Update
tbProduct.Close
'Products.loadProductNomenk
toProduct = True
Exit Function

ERR1:
If errorCodAndMsg("##425", -196) Then
'If Err = 3022 Then
    MsgBox "Одна и та же номенклатура может присутствовать в одном " & _
    "и том же изделии не более одного раза.", , "Позиция '" & gNomNom & _
    "' уже есть!"
'Else
'    MsgBox Error, , "Ошибка 425-" & Err & "(gNomNom=" & gNomNom & ":)  " '##425
End If
tbProduct.Close

End Function

Private Sub tbEndDate_Change()
controlVisible False
End Sub

Private Sub tbMobile_DblClick()
lbHide
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, i As Integer, old As String, row As Long, col As Integer, newPrevCost As String
Dim s As Single, result As String 'field() As Variant


If tbmobile_readonly And KeyCode = vbKeyReturn Then
    tbmobile_readonly = False
    KeyCode = vbKeyEscape
End If

If KeyCode = vbKeyReturn Then
 

 str = Trim(tbMobile.Text)
 old = Trim(Grid.TextMatrix(mousRow, nkNomer))
 
 If mousCol = nkCod Then
    If Not valueToNomencField("##420", str, "cod") Then GoTo EN1
 ElseIf mousCol = nkSize Then
    If Not valueToNomencField("##420", str, "Size") Then GoTo EN1
 ElseIf mousCol = nkPack Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not valueToNomencField("##420", CSng(str), "Pack") Then GoTo EN1
 ElseIf mousCol = nkNomer Then
    sql = "SELECT * FROM sGuideNomenk WHERE nomNom = '" & old & "'"
    Set tbNomenk = myOpenRecordSet("##104", sql, dbOpenForwardOnly)
'    If tbNomenk Is Nothing Then GoTo EN1
'   On Error GoTo ERR1
    If frmMode = "nomenkAddObraz" Then
        tbNomenk.AddNew
        On Error Resume Next 'некоторые ячейки м.б.пустыми
        tbNomenk!nomName = Grid.TextMatrix(mousRow, nkName)
        tbNomenk!ed_izmer = Grid.TextMatrix(mousRow, nkEdIzm)
        tbNomenk!ed_Izmer2 = Grid.TextMatrix(mousRow, nkEdIzm2)
        tbNomenk!perList = Grid.TextMatrix(mousRow, nkPerList)
        tbNomenk!Pack = Grid.TextMatrix(mousRow, nkPack)
        tbNomenk!cost = Grid.TextMatrix(mousRow, nkCena)
        tbNomenk!Cena1 = Grid.TextMatrix(mousRow, nkCENA1)
        tbNomenk!ves = Grid.TextMatrix(mousRow, nkVES)
        tbNomenk!STAVKA = Grid.TextMatrix(mousRow, nkSTAVKA)
        tbNomenk!FormulaNom = Grid.TextMatrix(mousRow, nkFormulaNom)
'        tbNomenk! = Grid.TextMatrix(mousRow, nkCenaFreight)
'        tbNomenk! = Grid.TextMatrix(mousRow, nkWebFormula)
        tbNomenk!margin = Grid.TextMatrix(mousRow, nkMargin)
        tbNomenk!kodel = Grid.TextMatrix(mousRow, nkKodel)
        tbNomenk!kolonok = Grid.TextMatrix(mousRow, nkKolonok)
        ' не добавлят вычисляемые kolon1/2/3/4
        tbNomenk!CENA_W = Grid.TextMatrix(mousRow, nkCena2W)
        sql = "SELECT sourceId from sGuideSource WHERE (((sourceName)='" & _
        Grid.TextMatrix(mousRow, nkSource) & "'));"
        If byErrSqlGetValues("##438", sql, i) Then tbNomenk!sourId = i
        tbNomenk!Size = Grid.TextMatrix(mousRow, nkSize)
    ElseIf frmMode = "nomenkAdd" Then
        tbNomenk.AddNew
    Else
'        tbNomenk.Index = "PrimaryKey"
'        tbNomenk.Seek "=", old
        tbNomenk.Edit
    End If
    On Error GoTo ERR1
    tbNomenk!nomnom = str
    tbNomenk!KlassId = gKlassId
    tbNomenk.Update
    tbNomenk.Close
    Grid.TextMatrix(mousRow, nkNomer) = str
    quantity = quantity + 1
    On Error GoTo 0
 ElseIf mousCol = nkName Then
    If Not valueToNomencField("##104", str, "nomName") Then GoTo EN1
 ElseIf mousCol = nkPerList Then
    If Not isNumericTbox(tbMobile, 1) Then Exit Sub
    If Not valueToNomencField("##104", CSng(str), "perList") Then GoTo EN1
 ElseIf mousCol = nkZakup Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not valueToNomencField("##104", CSng(str), "Zakup") Then GoTo EN1
    GoTo AA
 ElseIf mousCol = nkZapas Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not valueToNomencField("##104", CSng(str), "normZapas") Then GoTo EN1
AA: Grid.TextMatrix(mousRow, mousCol) = str
    calcZacup mousRow
    GoTo EN1
 ElseIf mousCol = nkCena Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    newPrevCost = Grid.TextMatrix(mousRow, nkCena)
    If Not valueToNomencField("##104", CSng(str), "cost") Then
        GoTo EN1
    End If
    Grid.TextMatrix(mousRow, nkPrevCost) = newPrevCost
    cenaFact = str
    GoTo CC
' ElseIf mousCol = nkBegOstat Then ' при этом на столько же изм-ся тек.ост-ки
'    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
'    If Not valueToNomencField("##104", CSng(str), "begOstatki") Then GoTo EN1 'это корректирует и остатки
'    If (Regim = "asOborot" Or Regim = "asOstat" Or Regim = "fltOborot" _
'    Or Regim = "sourOborot") And (ckEndDate.value = 0) Then
'        s = Grid.TextMatrix(mousRow, nkEndOstat) + tmpSingle
'        Grid.TextMatrix(mousRow, nkCurOstat) = Round(s, 2)
'        s = Grid.TextMatrix(mousRow, nkDostup) + tmpSingle
'        Grid.TextMatrix(mousRow, nkDostup) = Round(s, 2)
'    End If
 ElseIf mousCol = nkCENA1 Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not valueToNomencField("##104", CSng(str), "CENA1") Then GoTo EN1
    GoTo BB
 ElseIf mousCol = nkVES Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not valueToNomencField("##104", CSng(str), "VES") Then GoTo EN1
    GoTo BB
 ElseIf mousCol = nkSTAVKA Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not valueToNomencField("##104", CSng(str), "STAVKA") Then GoTo EN1
BB: result = nomenkFormula 'tmpStr
    Grid.TextMatrix(mousRow, nkCenaFreight) = result
    Grid.TextMatrix(mousRow, 0) = tmpStr 'освежаем формулу
    cenaFact = Grid.TextMatrix(mousRow, nkCena)
CC: cenaFreight = Grid.TextMatrix(mousRow, nkCenaFreight)
'    result = nomenkFormula("", "W") ' ,берем номер WEB формулы из sGuideNomenk
'    Grid.TextMatrix(mousRow, nkCena1W) = result
'    Grid.TextMatrix(mousRow, nkWebFormula) = tmpStr 'освежаем формулу
'    If Not IsNumeric(result) Then MsgBox result, , "Колонка '" & _
    Grid.TextMatrix(0, nkWebFormulaNom) & "'"
 ElseIf Regim = "" And mousRow = 0 Then
    Dim iKolon As Integer
    iKolon = mousCol - nkKolon2 + 2
    ValueToTableField "##mr0", "'" & str & "'", "sGuideKlass", "kolon" & CStr(iKolon), "byKlassId"
 
 ElseIf Regim = "" And (mousCol >= nkMargin And mousCol <= nkKolon4) Then
    If Not isNumericTbox(tbMobile) Then Exit Sub
    If Not checkNumeric(str, getMinValue(mousCol), getMaxValue(mousCol)) Then
        GoTo EN1
    End If
    
    Dim refreshGridCell As Long
    refreshGridCell = getRefreshIndex(mousCol)
    
    Dim margin As Double, baseCena As Double, cena2W As Double, manualOpt As Boolean
    If mousCol = nkCena2W Then
        cena2W = CDbl(str)
    Else
        cena2W = CDbl(Grid.TextMatrix(mousRow, nkCena2W))
    End If
    
    If mousCol = nkMargin Then
        margin = CDbl(str)
    Else
        margin = CDbl(Grid.TextMatrix(mousRow, nkMargin))
    End If
    
    If refreshGridCell = nkCena1W Then
        Grid.TextMatrix(mousRow, nkCena1W) = Format(CDbl(Grid.TextMatrix(mousRow, nkCenaFreight)) / (1 - margin / 100), "0.00")
    End If
    
    Dim kolonok As Integer, kodel As Double
    If mousCol = nkKolonok Then
        kolonok = CInt(str)
    Else
        kolonok = CInt(Grid.TextMatrix(mousRow, nkKolonok))
    End If
    If kolonok > 0 Then
        manualOpt = False
    Else
        manualOpt = True
    End If
    
    kolonok = Abs(kolonok)
    If mousCol = nkKodel Then
        kodel = CDbl(str)
    Else
        kodel = CDbl(Grid.TextMatrix(mousRow, nkKodel))
    End If
    
    baseCena = cena2W * (1 - margin / 100)
    
    If mousCol >= nkKolon2 And mousCol <= nkKolon4 Then
        If manualOpt Then
            If Not checkNumeric(str, 0, cena2W) Then
                Exit Sub
            End If
            Dim Nkol As Integer
            Nkol = mousCol - nkKolon2 + 2
            If Not valueToNomencField("##104", CSng(str), "CenaOpt" & Nkol) Then
                GoTo EN1
            End If
            GoTo EN2
        Else
            If Not checkNumeric(str, baseCena, cena2W) Then
                Exit Sub
            End If
            Dim kolonVal As Double
            kolonVal = CDbl(str)
            kodel = (kolonVal - baseCena) / (cena2W - baseCena)
            If Not valueToNomencField("##kodel", kodel, "kodel") Then
                GoTo EN1
            Else
                Grid.TextMatrix(mousRow, nkKodel) = Format(kodel, "0.0#")
            End If
        End If
        
    Else
        If Not valueToNomencField("##104", CSng(str), getChangedField(mousCol)) Then GoTo EN1
    End If
    
    Dim cenaOpt(3) As Double
    If manualOpt Then
        sql = "select cenaOpt2, cenaOpt3, cenaOpt4 from sguidenomenk where nomnom = '" & gNomNom & "'"
        byErrSqlGetValues "##438", sql, cenaOpt(1), cenaOpt(2), cenaOpt(3)
    End If
    
    For i = 1 To 3
        Grid.TextMatrix(mousRow, nkKolon2 + i - 1) = ""
        If kolonok > i Then
            If manualOpt Then
                Grid.TextMatrix(mousRow, nkKolon2 + i - 1) = Format(cenaOpt(i), "0.00")
            Else
                Grid.TextMatrix(mousRow, nkKolon2 + i - 1) = Format(calcKolonValue(baseCena, margin, kodel, kolonok, i + 1), "0.00")
            End If
        End If
    Next i
 End If
EN2:
 Grid.TextMatrix(mousRow, mousCol) = str
EN1:
 frmMode = ""
 lbHide
 
ElseIf KeyCode = vbKeyEscape Then

 If mousCol = nkNomer And (frmMode = "nomenkAdd" Or frmMode = "nomenkAddObraz") Then
    frmMode = ""
    If quantity > 0 Then
        Grid.RemoveItem quantity + 1 ' ту, которую зря добавили
    End If
 End If
 lbHide
 
End If
Exit Sub

ERR1:
If errorCodAndMsg("##105", -193) Then _
    MsgBox "Такой номер уже есть", , "Ошибка-105"
tbMobile.SetFocus
End Sub



Function getChangedField(iCol As Long) As String
    If iCol = nkMargin Then
        getChangedField = "margin"
    ElseIf iCol = nkKodel Then
        getChangedField = "kodel"
    ElseIf iCol = nkKolonok Then
        getChangedField = "Kolonok"
    ElseIf iCol = nkCena2W Then
        getChangedField = "CENA_W"
    End If
    
End Function

Function getMinValue(iCol As Long) As Double
    If iCol = nkMargin Then
        getMinValue = 1
    ElseIf iCol = nkKodel Then
        getMinValue = 0
    ElseIf iCol = nkKolonok Then
        getMinValue = -4
    ElseIf iCol = nkCena2W Then
        getMinValue = 0
    ElseIf iCol >= nkKolon2 And iCol <= nkKolon4 Then
        getMinValue = 0
    End If
    
End Function

Function getMaxValue(iCol As Long) As Double
    If iCol = nkMargin Then
        getMaxValue = 99
    ElseIf iCol = nkKodel Then
        getMaxValue = 1
    ElseIf iCol = nkKolonok Then
        getMaxValue = 4
    ElseIf iCol = nkCena2W Then
        getMaxValue = 1000000
    ElseIf iCol >= nkKolon2 And iCol <= nkKolon4 Then
        getMaxValue = 1000000
    End If
End Function

Function getRefreshIndex(iCol As Long) As Long
    If iCol = nkMargin Then
        getRefreshIndex = nkCena1W
    ElseIf iCol = nkKodel Then
        getRefreshIndex = nkKolon2
    ElseIf iCol = nkKolonok Then
        getRefreshIndex = nkKolon2
    ElseIf iCol = nkCena2W Then
    End If
End Function

'возвращает True если Regim = "fltOborot" и не надо закупать
'а при reg <> ""  вычисляет ДО и ФО
Function calcZacup(row As Long, Optional reg As String = "") As Boolean
Dim maxZap As Single, minzap As Single, zakup As Single, str As String
     
calcZacup = False
If reg = "" Then ' при редактированиии  и только для used
     maxZap = Grid.TextMatrix(row, nkZakup)
     minzap = Grid.TextMatrix(row, nkZapas)
     str = Grid.TextMatrix(row, nkMark)
     gain = 1 ' редактирование запрещено при установленном флаге "в У.Е."
Else
    maxZap = Round(tbNomenk!zakup * gainC, 2) 'Макс.запас в базе в целых!
    minzap = Round(tbNomenk!normZapas * gainC, 2) 'maxZap
    str = tbNomenk!mark
End If
'dOst = Round(nomencDostupOstatki * gain, 2)  'доступные остатки (и FO)
dOst = Round(nomencDostupOstatki("int"), 2)  'доступные остатки (и FO) в целых


If dOst >= minzap Or str = lbMark.List(1) Then  'если ДО достаточны или Unused
    If Regim = "fltOborot" Then calcZacup = True: Exit Function
    Grid.TextMatrix(row, nkDeficit) = 0 ' к закупке
Else
    Grid.TextMatrix(row, nkDeficit) = Round(maxZap - dOst, 2) 'к закупке
End If
    
'If reg = "" Then Exit Function

If str = lbMark.List(0) Then 'used
    Grid.TextMatrix(row, nkZapas) = minzap
    Grid.TextMatrix(row, nkZakup) = maxZap
Else
    Grid.TextMatrix(row, nkZapas) = dOst 'minZap
'    Grid.TextMatrix(row, nkZakup) = Round(FO * gain, 2) 'maxZap
    Grid.TextMatrix(row, nkZakup) = Round(FO, 2) 'maxZap
End If

End Function

Private Sub tbPostav_GotFocus()
    If IsNumeric(tbPostav.Text) Then
        gSrokPostav = CSng(tbPostav.Text)
    End If
End Sub

Private Sub tbPostav_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        tbPostav_LostFocus
    End If
End Sub

Private Sub tbPostav_LostFocus()
    If IsNumeric(tbPostav.Text) Then
        Dim newSrok As Single: newSrok = CSng(tbPostav.Text)
        
        If newSrok <> gSrokPostav Then
            loadKlassNomenk
        End If
        gSrokPostav = newSrok
    End If
End Sub

Private Sub tbStartDate_Change()
controlVisible False
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Me.PopupMenu mnContext2
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    loadKlassNomenk
    Grid.Visible = True
    ckUnUsed.Visible = True
    If quantity = 0 Then
        laInform.Caption = "Все текущие остатки корректны."
    Else
        laInform.Caption = "Показаны позиции, у которых текущие остатки " & _
        "не совпадали с вычисленными на основе начальных остатков и записей " & _
        "из ДМЦ." & vbCrLf & vbCrLf & "Все несовпадения уже устранены."
    End If

End Sub


Private Sub tv_AfterLabelEdit(Cancel As Integer, NewString As String)
If Regim = "sourOborot" Then Exit Sub
' If Not flKlassAdd Then
gKlassId = Mid$(tv.SelectedItem.Key, 2)
ValueToTableField "##101", "'" & NewString & "'", "sGuideKlass", "klassName", "byKlassId"
End Sub


Function getBaySaledQty(p_nomnom As String, p_startDate As String, p_endDate As String) As Single
    sql = "select wf_sale_nomenk_qty ('" & p_nomnom & "', convert(datetime, " & p_startDate & "), convert(datetime, " & p_endDate & "))"
    byErrSqlGetValues "##getBaySaledQty", sql, getBaySaledQty
End Function


Private Sub getAvgOutcome(p_nomnom As String, ByVal p_startDate As String, ByVal p_endDate As String, _
ByRef avgOutcome As Double, ByRef missedDays As Integer, ByRef saledQty As Double, ByRef incomeQty As Double, ByRef outcomeQty As Double)
Dim csvResult As String
    If p_startDate = "" Then
        p_startDate = "null"
    End If
    If p_endDate = "" Then
        p_endDate = "null"
    End If
    sql = "select wf_sale_turnover_metrics ('" & p_nomnom & "', convert(datetime, " & p_startDate & "), convert(datetime, " & p_endDate & "))"
    byErrSqlGetValues "##getAvgOutcome", sql, csvResult
    parseCsvOutcome csvResult, avgOutcome, missedDays, saledQty, incomeQty, outcomeQty
End Sub


Private Sub parseCsvOutcome(ByVal csv As String, ByRef avgOutcome As Double, ByRef missedDays As Integer, ByRef saledQty As Double, ByRef incomeQty As Double, ByRef outcomeQty As Double)
Dim done As Boolean, sepIndex As Long, token As String
Dim restCsv As String, currentOrder As Integer

    done = False
    restCsv = csv
    currentOrder = 0
    
    While Not done
        sepIndex = InStr(1, restCsv, ";")
        If restCsv = "" Then
            done = True
        Else
            If sepIndex > 0 Then
                token = left(restCsv, sepIndex - 1)
                restCsv = Mid(restCsv, sepIndex + 1)
            Else
                token = restCsv
                restCsv = ""
            End If
        End If
        If token = "" Then
            token = "0"
        End If
        If IsNumeric(token) Then
            If currentOrder = 0 Then
                avgOutcome = CDbl(token)
            ElseIf currentOrder = 1 Then
                missedDays = CInt(token)
            ElseIf currentOrder = 2 Then
                saledQty = CDbl(token)
            ElseIf currentOrder = 3 Then
                incomeQty = CDbl(token)
            ElseIf currentOrder = 4 Then
                outcomeQty = CDbl(token)
            End If
        End If
        currentOrder = currentOrder + 1
        If restCsv = "" Then
            done = True
        End If
    Wend

End Sub

Sub loadKlassNomenk(Optional filtr As String = "")
Dim il As Long, strWhere As String, befWhere  As String
Dim insWhere As String, strN As String, i As Integer, s As Single
Dim beg As Single, prih As Double, rash As Double, oldNow As Single
Dim Cena1 As Double

'
' Regim = "" - справочник по номенклатуре
'       asOborot - оборотная ведомость
'       asOstat - ведомость остатков
'       sourOborot - оборотная по поставщикам
'       ventureOborot - оборотная по предприятиям
'       forKartaDMC - карточка движения
'       fltOborot - номенклатура для закупки
'       obrez - подрежим "только обрезные" в ведомости остатков. Работает не правильно для Всей Номенклатуре(не делает фильтрацию)
'       checkCurOstat - Сверка текущих остатков
'
'
'

ckUnUsed.value = 0
If Regim = "asOborot" Or Regim = "sourOborot" Or Regim = "fltOborot" Then
    ' Set dates for parameters to stored procs (i.e.
    setStartEndDates tbStartDate, tbEndDate
    befWhere = getWhereByDateBoxes(Me, "sDocs.xDate", begDate, "befo")
    If befWhere = "error" Then Exit Sub
End If
If Regim = "asOborot" Or Regim = "sourOborot" Or Regim = "fltOborot" Or _
Regim = "asOstat" Or Regim = "obrez" Then
     insWhere = getWhereByDateBoxes(Me, "sDocs.xDate", begDate)
     If insWhere = "error" Then Exit Sub
End If
Me.MousePointer = flexHourglass

'MsgBox "insWhere=" & insWhere & "  befWhere=" & befWhere

If beShift Then
  Grid.AddItem ""
Else
  quantity = 0
  Grid.Visible = False
  ckUnUsed.Visible = False
  clearGrid Grid
End If

controlVisible True
Grid.Visible = False
If filtr = "obrez" Then
    strWhere = "WHERE sGuideNomenk.perList > 1.0 "
ElseIf filtr <> "" Then
    If left$(filtr, 1) = "F" Then
        strWhere = "WHERE sGuideNomenk.nomNom Like '*" & Mid$(filtr, 2) & "*'"
    ElseIf left$(filtr, 1) = "G" Then
        strWhere = "WHERE sGuideNomenk.nomName Like '*" & Mid$(filtr, 2) & "*'"
    End If
ElseIf Regim = "checkCurOstat" Or Regim = "fltOborot" Then
    strWhere = ""
ElseIf tv.SelectedItem.Key = "all" Then
    If frmMode <> "" Then GoTo EN1
    strWhere = ""
    sql = "SELECT sGuideNomenk.* From sGuideNomenk"
    quantity = 0
ElseIf Regim = "sourOborot" Then
    strWhere = "WHERE (((sGuideNomenk.sourId)=" & gKlassId & "))"
ElseIf gKlassType = "p" Then
    strWhere = "join sPriceBulkChange pbc on pbc.id = " & gKlassId _
    & " join sPriceHistory sph on sph.bulk_id = pbc.id and sGuideNomenk.nomnom = sph.nomnom "
Else
    strWhere = "WHERE sGuideNomenk.klassId = " & gKlassId
End If

If IsNumeric(gKlassId) And Regim = "" Then
    adjustKolonHeaders gKlassId, gKlassType
End If


sql = "SELECT ph.prev_cost, sGuideNomenk.*, f.Formula, " _
& vbCr & "sGuideSource.sourceName, ph.nomnom as priceChanged " _
& vbCr & "FROM sGuideNomenk " _
& vbCr & " JOIN sGuideSource ON sGuideNomenk.sourId = sGuideSource.sourceId " _
& vbCr & " JOIN sGuideFormuls f ON sGuideNomenk.formulaNom = f.nomer " _
& vbCr & " left join (select h.cost as prev_cost, h.nomnom from spricehistory h join (select max(change_date) as change_date, nomnom from spricehistory m group by nomnom) mx on mx.nomnom = h.nomnom and mx.change_date = h.change_date ) ph on ph.nomnom = sguidenomenk.nomnom  " _
& vbCr & strWhere & " ORDER BY sGuideNomenk.nomNom ;"
'Debug.Print sql
'MsgBox sql
Set tbNomenk = myOpenRecordSet("##165", sql, dbOpenForwardOnly) ' dbOpenDynaset)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
 tbNomenk.MoveFirst
 While Not tbNomenk.EOF
    gNomNom = tbNomenk!nomnom
'    beg = tbNomenk!begOstatki
    cenaFact = tbNomenk!cost 'цена фактическая
    cenaFreight = nomenkFormula("noOpen")
    
'    If cbInside.ListIndex > 1 Then beg = 0
    oldNow = Round(tbNomenk!nowOstatki, 2)
    
    strN = "(sDMC.nomNom) = '" & gNomNom & "'"
    gainC = 1
    If chGain.Visible And chGain.value = 1 Then _
        gainC = tbNomenk!cost 'цена фактическая
    gain = gainC
    If ((chPerList.Visible And chPerList.value = 1) Or Regim <> "asOstat") And Regim <> "" Then
        gain = gain / tbNomenk!perList 'оборотные всегда выдаем в целых
        Grid.TextMatrix(quantity + 1, nkEdIzm) = tbNomenk!ed_Izmer2
    Else
        Grid.TextMatrix(quantity + 1, nkEdIzm) = tbNomenk!ed_izmer
    End If
    
    chGain.Enabled = True
    chPerList.Enabled = True
    
    If Regim = "checkCurOstat" Then
        prih = PrihodRashod2("+", strN) - PrihodRashod2("-", strN)
        prih = Round(prih, 2)
        If Abs(oldNow - prih) < 0.001 Then GoTo NXT
        sql = "UPDATE sGuideNomenk SET nowOstatki = " & prih & _
        " WHERE  nomNom = '" & gNomNom & "'"
        If myExecute("##473", sql) <> 0 Then myBase.Close: End
'        tbNomenk.Edit
'        tbNomenk!nowOstatki =prih
'        tbNomenk.Update
        Grid.TextMatrix(quantity + 1, nkCheckOst) = prih * gain
    ElseIf Regim = "sourOborot" Or _
    Regim = "fltOborot" Or Regim = "asOstat" Or Regim = "" Then
        strWhere = strN
        If Regim = "asOstat" Or Regim = "" Then
          dOst = Round(nomencDostupOstatki * gain, 2)  'доступные остатки (и FO)
        Else
          If calcZacup(quantity + 1, "load") Then GoTo NXT 'пропускаем, если не требует закупки а Regim = "fltOborot"
'            Grid.TextMatrix(quantity + 1, nkMark) = tbNomenk!mark
        End If
        Grid.TextMatrix(quantity + 1, nkDostup) = dOst
        Grid.TextMatrix(quantity + 1, nkCena) = cenaFact
    End If
    quantity = quantity + 1
    Grid.TextMatrix(quantity, nkNomer) = gNomNom
    Grid.TextMatrix(quantity, nkCod) = tbNomenk!cod
    Grid.TextMatrix(quantity, nkName) = tbNomenk!nomName
    Grid.TextMatrix(quantity, nkSize) = tbNomenk!Size
            
     If Regim <> "checkCurOstat" Then _
        Grid.TextMatrix(quantity, nkMark) = tbNomenk!mark
    
    
    If Regim = "asOborot" Or Regim = "sourOborot" Or Regim = "asOstat" Then
'      Grid.TextMatrix(quantity, nkBegOstat) = Round(beg * gain, 2) 'времено по кнопке 'обрезная'
      If insWhere <> "" Then strWhere = insWhere & ") And (" & strWhere
    End If
    
    If Regim = "fltOborot" Then ' ном-ра для закупки
'        Grid.TextMatrix(quantity, nkEndOstat) = Round(gain * oldNow, 2)
        Grid.TextMatrix(quantity, nkEndOstat) = Round(gain * FO, 2) ' строка д.б после вызова nomencDostupOstatki
'        GoTo AA
    ElseIf Regim = "asOborot" Or Regim = "sourOborot" Then
        Dim saled As Double, saledProcent As Double
        Dim avgOutcome As Double, missedDays As Integer
        
        getAvgOutcome tbNomenk!nomnom, startDate, endDate, avgOutcome, missedDays, saled, prih, rash
'        prih = PrihodRashod2("+", strWhere)
'        rash = PrihodRashod2("-", strWhere)
        If befWhere = "" Then
            beg = 0
        Else
            strWhere = befWhere & ") And (" & strN
'            beg = beg + PrihodRashod2("+", strWhere) - PrihodRashod2("-", strWhere)
            beg = PrihodRashod2("+", strWhere) - PrihodRashod2("-", strWhere)
        End If
    
        Grid.TextMatrix(quantity, nkBegOstat) = Round(beg * gain, 2) ' ост на начало
        Grid.TextMatrix(quantity, nkPrihod) = Round(prih, 2)
        Grid.TextMatrix(quantity, nkRashod) = Round(rash, 2)
        
        dOst = Round(nomencDostupOstatki("int"), 2)  'доступные остатки (и FO) в целых
        
        Cena1 = tbNomenk!Cena1
        
        recaluculateZakup quantity, avgOutcome, dOst, Cena1, _
                tbNomenk!mark, tbNomenk!ves, tbNomenk!normZapas, tbNomenk!zakup
                
        'Dim z As Boolean: z = calcZacup(quantity, "load")
        Grid.TextMatrix(quantity, nkDostup) = dOst
        Grid.TextMatrix(quantity, nkCena) = cenaFact
        
        
'        = getBaySaledQty(tbNomenk!nomnom, startDate, endDate)
        Grid.TextMatrix(quantity, nkRashodBay) = Round(saled, 2)
        If rash > 0 And saled > 0 Then
            Grid.TextMatrix(quantity, nkSaledProcent) = Format((saled / (rash)) * 100, "##0")
        Else
            If saled > 0 Then
                Grid.TextMatrix(quantity, nkSaledProcent) = "100"
            End If
        End If
        If avgOutcome <> 0 Then
            Grid.TextMatrix(quantity, nkAvgOutcome) = Format(avgOutcome, "# ##0.0")
            If missedDays <> 0 Then
                Grid.TextMatrix(quantity, nkAvgOutcome) = Grid.TextMatrix(quantity, nkAvgOutcome) & "*"
            End If
        Else
            Grid.TextMatrix(quantity, nkAvgOutcome) = "-"
        End If
        Grid.TextMatrix(quantity, nkEndOstat) = Round(beg * gain + prih - rash, 2)  ' остаток на конец
        If Regim = "asOborot" Then GoTo BB
    ElseIf Regim = "asOstat" Then
    
        Grid.TextMatrix(quantity, nkPerList) = tbNomenk!perList      '
 '       beg = beg * gain:
'        rash = 0 ' это б. сумма по всем складам
        For i = 1 To Documents.lbInside.ListCount
          prih = PrihodRashod2("+", strWhere, i) - PrihodRashod2("-", strWhere, i)
'          prih = beg + gain * prih
          prih = gain * prih
'          rash = rash + prih
          Grid.TextMatrix(quantity, nkSkladOst + i - 1) = Round(prih, 2)
'          beg = 0 ' на остальных складах
        Next i
'        Grid.TextMatrix(quantity, nkEndOstat) = Round(rash, 2)
        Grid.TextMatrix(quantity, nkEndOstat) = Round(gain * FO, 2) ' строка д.б после вызова nomencDostupOstatki
    Else
'        Grid.TextMatrix(quantity, nkBegOstat) = Round(gain * tbNomenk!begOstatki, 2)
        If Regim = "checkCurOstat" Then
            Grid.TextMatrix(quantity, nkCurOstat) = Round(gain * oldNow, 2)
        Else
            If Regim = "" Then
                Grid.TextMatrix(quantity, nkEndOstat) = Round(FO, 2) ' строка д.б после вызова nomencDostupOstatki
                Grid.TextMatrix(quantity, nkDostup) = Round(dOst, 2)
            End If
            Grid.TextMatrix(quantity, nkCena) = cenaFact
            If Not IsNull(tbNomenk!prev_cost) Then
                Grid.TextMatrix(quantity, nkPrevCost) = tbNomenk!prev_cost
            Else
                Grid.TextMatrix(quantity, nkPrevCost) = "--"
            End If
            Cena1 = tbNomenk!Cena1
            Grid.TextMatrix(quantity, nkCENA1) = Format(Cena1, "0.00")
            Grid.TextMatrix(quantity, nkVES) = tbNomenk!ves
            Grid.TextMatrix(quantity, nkSTAVKA) = tbNomenk!STAVKA
            Grid.TextMatrix(quantity, 0) = tbNomenk!formula
            Grid.TextMatrix(quantity, nkCenaFreight) = Format(cenaFreight, "0.00")
            Grid.TextMatrix(quantity, nkFormulaNom) = tbNomenk!FormulaNom
            Grid.TextMatrix(quantity, nkYesNo) = tbNomenk!YesNo
            If Not IsNull(tbNomenk!SourceName) Then _
                Grid.TextMatrix(quantity, nkSource) = tbNomenk!SourceName
            If Not IsNull(tbNomenk!perList) Then _
                Grid.TextMatrix(quantity, nkPerList) = tbNomenk!perList

            If IsNumeric(cenaFreight) Then
                Grid.TextMatrix(quantity, nkCena1W) = Format(cenaFreight / (1 - tbNomenk!margin / 100), "0.00")
            End If
            
            Grid.TextMatrix(quantity, nkCena2W) = Format(tbNomenk!CENA_W, "0.00")
            Dim optBasePrice As Double
            optBasePrice = tbNomenk!CENA_W * (1 - tbNomenk!margin / 100)
            Grid.TextMatrix(quantity, nkMargin) = tbNomenk!margin
            Grid.TextMatrix(quantity, nkKodel) = Format(tbNomenk!kodel, "0.0#")
            Grid.TextMatrix(quantity, nkKolonok) = tbNomenk!kolonok
            Dim kolonok As Integer, manualOpt As Boolean
            kolonok = tbNomenk!kolonok
            If kolonok > 0 Then
                manualOpt = False
            Else
                manualOpt = True
            End If
            
            For i = 1 To Abs(kolonok) - 1
                Grid.TextMatrix(quantity, nkKolon2 + i - 1) = ""
                If manualOpt Then
                    Grid.TextMatrix(quantity, nkKolon2 + i - 1) = Format(tbNomenk("CenaOpt" & CStr(i + 1)), "0.00")
                Else
                    Grid.TextMatrix(quantity, nkKolon2 + i - 1) = Format(calcKolonValue(optBasePrice, tbNomenk!margin, tbNomenk!kodel, Abs(kolonok), i + 1), "0.00")
                End If
            Next i
            

'            Grid.TextMatrix(quantity, nkSize) = tbNomenk!Size
            Grid.TextMatrix(quantity, nkEdIzm2) = tbNomenk!ed_Izmer2
            If IsNumeric(tbNomenk!Pack) Then _
                Grid.TextMatrix(quantity, nkPack) = tbNomenk!Pack

BB:
            Grid.TextMatrix(quantity, nkWeb) = tbNomenk!web
        End If
    End If
    Grid.AddItem ""
NXT:
    tbNomenk.MoveNext
 Wend
NXT2:
 If quantity > 0 Then Grid.RemoveItem quantity + 1
End If
tbNomenk.Close
laQuant = quantity
Grid.Visible = True
ckUnUsed.Visible = True
On Error Resume Next
Grid.Visible = True
Grid.SetFocus
Grid_EnterCell

EN1:
If frmMode = "" Then
    Me.MousePointer = flexDefault
Else
    Me.MousePointer = flexUpArrow
End If
End Sub

Sub adjustKolonHeaders(ByVal KlassId As Integer, ByVal KlassType)
Dim Kolon1 As String
Dim Kolon2 As String
Dim Kolon3 As String
Dim Kolon4 As String

    If KlassType <> "p" Then
        sql = "SELECT kolon1, kolon2, kolon3, kolon4 from sGuideKlass where klassId = " & KlassId
        byErrSqlGetValues "##ACH", sql, Kolon1, Kolon2, Kolon3, Kolon4
        If "" <> IIf(IsNull(Kolon1), "", Kolon1) Then
            Grid.TextMatrix(0, nkCena2W) = Kolon1
        Else
            Grid.TextMatrix(0, nkCena2W) = ""
        End If
        
        If "" <> IIf(IsNull(Kolon2), "", Kolon2) Then
            Grid.TextMatrix(0, nkKolon2) = Kolon2
        Else
            Grid.TextMatrix(0, nkKolon2) = ""
        End If
        
        If "" <> IIf(IsNull(Kolon3), "", Kolon3) Then
            Grid.TextMatrix(0, nkKolon3) = Kolon3
        Else
            Grid.TextMatrix(0, nkKolon3) = ""
        End If
        
        If "" <> IIf(IsNull(Kolon4), "", Kolon4) Then
            Grid.TextMatrix(0, nkKolon4) = Kolon4
        Else
            Grid.TextMatrix(0, nkKolon4) = ""
        End If
    Else
        Grid.TextMatrix(0, nkCena2W) = "CenaSale"
        Grid.TextMatrix(0, nkKolon2) = ""
        Grid.TextMatrix(0, nkKolon3) = ""
        Grid.TextMatrix(0, nkKolon4) = ""
    End If
End Sub

Private Sub recaluculateZakup(ByVal row As Long, ByVal avgOutcome As Single, ByVal dOst As Single, ByVal cenaFact As String _
        , ByVal mark As Variant, ByVal ves As Variant, ByVal normZapas As Variant, ByVal zakup As Variant _
)
        
    If avgOutcome > 0 Then
        ' вычислить мин/макс запасы/к заявке по новой формуле
        If IsNumeric(tbPostav.Text) Then
            Dim srokPostav As Single
            srokPostav = CSng(tbPostav.Text)
            
            Dim minzap As Single: minzap = srokPostav * avgOutcome
            Grid.TextMatrix(row, nkZapas) = Round(minzap, 0)
            Grid.TextMatrix(row, nkZakup) = Round(minzap * 2, 0)
            Dim kZajav As Single
            kZajav = avgOutcome * (srokPostav * 2 + 0.5) - dOst
            If kZajav < 0 Then
                kZajav = 0
            ElseIf dOst >= minzap Then
                ' доступные остатки больше, чем мин. запас
                'kZajav = 0
            ElseIf mark = lbMark.List(1) Then
                'unused
                kZajav = 0
            End If
            Grid.TextMatrix(row, nkDeficit) = Round(kZajav, 0)
            If kZajav > 0 Then
                kZajav = Round(kZajav, 0)
                Grid.TextMatrix(row, nkZakupBax) = Round(kZajav * cenaFact, 2)
                Grid.TextMatrix(row, nkZakupWeight) = Round(kZajav * ves, 1)
            End If
        End If
    Else
        Grid.TextMatrix(row, nkZapas) = Round(normZapas * gainC, 2) 'maxZap
        Grid.TextMatrix(row, nkZakup) = Round(zakup * gainC, 2) 'Макс.запас в базе в целых!
        Grid.TextMatrix(row, nkDeficit) = "0"
    End If

End Sub



Public Function nomencDostupOstatki(Optional intQuant As String = "") As Single
Dim s As Single, z As Single
' Ф.остатки
FO = PrihodRashod("+", -1001) - PrihodRashod("-", -1001)

'зарезервировано для производства и  для продажи
sql = "SELECT Sum(quantity) AS Sum_quantity, " & _
"Sum(Sum_quant) AS Sum_Sum_quant From wCloseNomenk " & _
"WHERE (((nomNom)='" & gNomNom & "'));"

'Debug.Print sql
If Not byErrSqlGetValues("##469", sql, z, s) Then myBase.Close: End
nomencDostupOstatki = FO - (z - s) ' минус, что несписано
If intQuant <> "" Or Regim = "" Then
    sql = "SELECT perList from sGuideNomenk WHERE (((nomNom)='" & gNomNom & "'));"
    If Not byErrSqlGetValues("##436", sql, s) Then myBase.Close: End
    nomencDostupOstatki = nomencDostupOstatki / s
    FO = FO / s
End If
End Function

'skladId=0 - cуммарно по всем складам
'skladId=2 - cуммарно по 1 и 2му складам
Function PrihodRashod(reg As String, skladId As Integer) As Single
Dim qWhere As String, s As Single

PrihodRashod = 0

If reg = "+" Then
'    If skladId = 0 Then
'        qWhere = ") AND ((sDocs.destId) < -1000)"
'    ElseIf skladId = 2 Then
'        qWhere = ") AND ((sDocs.destId) = -1001 Or (sDocs.destId) = -1002)"
'    Else
        qWhere = ") AND ((sDocs.destId) =" & skladId & ")"
'    End If
ElseIf reg = "-" Then
'    If skladId = 0 Then
'        qWhere = ") AND ((sDocs.sourId) < -1000)"
'    ElseIf skladId = 2 Then
'        qWhere = ") AND ((sDocs.sourId) = -1001 Or (sDocs.sourId) = -1002)"
'    Else
        qWhere = ") AND ((sDocs.sourId) =" & skladId & ")"
'    End If
End If
sql = "SELECT Sum(sDMC.quant) AS Sum_quantity FROM sDocs INNER JOIN " & _
"sDMC ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc) " & _
"WHERE (((sDMC.nomNom) = '" & gNomNom & "' " & qWhere & ");"
'MsgBox sql
byErrSqlGetValues "##157", sql, PrihodRashod

'If skladId >= -1001 And reg = "+" Then
'    sql = "SELECT sGuideNomenk.begOstatki From sGuideNomenk " & _
'    "WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
'    If Not byErrSqlGetValues("##161", sql, s) Then Exit Function
'    PrihodRashod = PrihodRashod + s
'End If
End Function

Function PrihodRashod2(reg As String, sWhere As String, Optional sklad As Integer = 0) As Single
Dim qWhere, str As String

PrihodRashod2 = 0
If sklad > 0 Then
    str = "= " & -1000 - sklad ' по кажд.складу (вед-ть остатков)
    GoTo AA
ElseIf Regim = "checkCurOstat" Then
    str = "< -1000"
    GoTo AA
ElseIf cbInside.ListIndex = 0 Then 'по всем, исключая межскладские
    If reg = "+" Then
        qWhere = ") AND ((sDocs.sourId)> -1000 AND (sDocs.destId)< -1000)"
    ElseIf reg = "-" Then
        qWhere = ") AND ((sDocs.sourId)< -1000 AND (sDocs.destId)> -1000)"
    End If
Else
    str = "= " & -1000 - cbInside.ListIndex 'по кажд.складу (оборот-я вед-ть)
AA: If reg = "+" Then
        qWhere = ") AND ((sDocs.destId)" & str & ")"
    ElseIf reg = "-" Then
        qWhere = ") AND ((sDocs.sourId)" & str & ")"
    End If
End If

sql = "SELECT Sum(sDMC.quant) AS Sum_quantity FROM sDocs INNER JOIN " & _
"sDMC ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc) " & _
"WHERE ((" & sWhere & qWhere & ");"
'MsgBox sql
If Not byErrSqlGetValues("##173", sql, PrihodRashod2) Then myBase.Close: End

End Function

Sub controlVisible(enabl As Boolean)
Grid.Visible = enabl
ckUnUsed.Visible = enabl
laKolvo.Visible = enabl
laQuant.Visible = enabl
cmHide.Visible = enabl
cmExcel.Visible = enabl
'chGain.Visible = enabl
End Sub



Private Sub tv_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, str As String
    If Regim = "sourOborot" Then Exit Sub
    
    'If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
    If KeyCode = vbKeyReturn Then
        tv_NodeClick tv.SelectedItem
    End If
End Sub

Private Sub tv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Regim = "sourOborot" Then Exit Sub

If Button = 2 Then
    mousRight = 1
Else
    mousRight = 0
End If
beShift = False
If Shift = 2 Then beShift = True

End Sub

Private Sub tv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim str As String
If Regim = "sourOborot" Then Exit Sub

If mousRight = 2 Then
  If frmMode = "nomenkReplace" Then
    Me.PopupMenu mnContext3
  ElseIf Regim = "" Then
    str = tv.SelectedItem.Key
    If str = "all" Then
        mnSep11.Visible = False
        mnCost.Visible = False
        Exit Sub
    End If
    If str = "k0" Then
        mnRen.Visible = False
        mnDel.Visible = False
        If frmMode <> "klassReplace" Then mnRepl.Visible = False
        mnSep.Visible = False
    ElseIf frmMode = "" Then
        mnRen.Visible = True
        mnDel.Visible = True
        mnRepl.Visible = True
        mnSep.Visible = True
    End If
    
    If bulkChangEnabled And Mid(str, 1, 1) <> "p" Then
        mnSep11.Visible = True
        mnCost.Visible = True
    Else
        mnSep11.Visible = False
        mnCost.Visible = False
    End If
    
    Me.PopupMenu mnContext
  End If
'    Timer1.Interval = 10
'    Timer1.Enabled = True
End If
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    If Not IsNull(tv.SelectedItem.Key) Then
        gKlassId = Mid$(tv.SelectedItem.Key, 2)
        gKlassType = Mid$(tv.SelectedItem.Key, 1, 1)
        gSourceId = gKlassId
        If mousRight = 1 Then
            mousRight = 2 ' правый клик был именно из Node
            Exit Sub
        End If
        
        If Regim <> "sourOborot" And tv.SelectedItem.Key = "k0" Then
            If frmMode = "" Then controlVisible False
            quantity = 0
            Exit Sub
        End If
        
        loadKlassNomenk
    End If
'Grid.SetFocus
'Grid_EnterCell

End Sub
