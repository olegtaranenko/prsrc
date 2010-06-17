VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Documents 
   BackColor       =   &H8000000A&
   Caption         =   "Склад"
   ClientHeight    =   6192
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   11748
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6192
   ScaleWidth      =   11748
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lbVenture 
      Appearance      =   0  'Flat
      Height          =   600
      Left            =   5500
      TabIndex        =   31
      Top             =   1000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "между"
      Height          =   315
      Left            =   3840
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox tbEnable 
      BackColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   11280
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.ListBox lbInside 
      Height          =   240
      ItemData        =   "Documents.frx":0000
      Left            =   9000
      List            =   "Documents.frx":0002
      TabIndex        =   26
      Top             =   2520
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox lbGroup 
      Height          =   432
      ItemData        =   "Documents.frx":0004
      Left            =   9000
      List            =   "Documents.frx":000E
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ListBox lbSource 
      Height          =   2352
      Left            =   6300
      TabIndex        =   27
      Top             =   780
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.CommandButton cmAdd2 
      Caption         =   "Добавить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6300
      TabIndex        =   16
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   1260
      Top             =   5640
   End
   Begin VB.CommandButton cmProduct 
      Caption         =   "Изделие"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6600
      TabIndex        =   25
      Top             =   5100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame frBad 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   600
      TabIndex        =   19
      Top             =   1500
      Visible         =   0   'False
      Width           =   4515
      Begin VB.CommandButton cmExit 
         Caption         =   "Выход"
         Height          =   315
         Left            =   3480
         TabIndex        =   24
         Top             =   3300
         Width           =   915
      End
      Begin VB.CommandButton cmKarta 
         Caption         =   "Карточка"
         Height          =   315
         Left            =   180
         TabIndex        =   23
         Top             =   3300
         Width           =   915
      End
      Begin VB.ListBox lbBad 
         Height          =   2544
         ItemData        =   "Documents.frx":002F
         Left            =   60
         List            =   "Documents.frx":0031
         TabIndex        =   20
         Top             =   540
         Width           =   4395
      End
      Begin VB.Label laFrame 
         Alignment       =   2  'Center
         Caption         =   "laFrame"
         Height          =   435
         Left            =   180
         TabIndex        =   22
         Top             =   60
         Width           =   4155
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   4515
      End
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3960
      TabIndex        =   18
      Top             =   5760
      Width           =   915
   End
   Begin VB.CommandButton cmDel2 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7740
      TabIndex        =   17
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox tbMobile2 
      Height          =   315
      Left            =   10320
      TabIndex        =   15
      Text            =   "tbMobile2"
      Top             =   4500
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   4680
      TabIndex        =   14
      Text            =   "tbMobile"
      Top             =   1020
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   5760
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4815
      Left            =   5880
      TabIndex        =   10
      Top             =   735
      Visible         =   0   'False
      Width           =   5835
      _ExtentX        =   10287
      _ExtentY        =   8509
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4815
      Left            =   60
      TabIndex        =   9
      Top             =   735
      Width           =   5775
      _ExtentX        =   10181
      _ExtentY        =   8509
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmLoad 
      Caption         =   "Загрузить"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   -60
      Width           =   11655
      Begin VB.TextBox tbEndDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Top             =   180
         Width           =   795
      End
      Begin VB.CheckBox ckEndDate 
         Caption         =   " "
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox tbStartDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1260
         TabIndex        =   2
         Text            =   "01.11.02"
         Top             =   180
         Width           =   795
      End
      Begin VB.CheckBox ckStartDate 
         Caption         =   " "
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   180
         Width           =   315
      End
      Begin VB.Label laFiltr 
         Caption         =   "Документ из Карты ДМЦ!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7020
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label laPo 
         Caption         =   "пос"
         Height          =   195
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   195
      End
      Begin VB.Label laPeriod 
         Caption         =   "Период с  "
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Label laGrid2 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   6120
      TabIndex        =   12
      Top             =   540
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Реестр документов"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   540
      Width           =   5535
   End
   Begin VB.Menu mnReestr 
      Caption         =   "Документы"
      Begin VB.Menu mnDocFind 
         Caption         =   "Поиск по номеру       F7"
      End
   End
   Begin VB.Menu mnMeassure 
      Caption         =   "Настройка"
      Begin VB.Menu mnVentureIncomeSetting 
         Caption         =   "Приход по предприятиям"
      End
   End
   Begin VB.Menu mnReports 
      Caption         =   "Отчеты"
      Begin VB.Menu mnOstVed 
         Caption         =   "Вед. остатков на дату"
      End
      Begin VB.Menu mnOborot 
         Caption         =   "Оборотная ведомость"
         Visible         =   0   'False
      End
      Begin VB.Menu sourOborot 
         Caption         =   "Оборотная по поставщикам"
         Visible         =   0   'False
      End
      Begin VB.Menu ventureOborot 
         Caption         =   "Оборотная по предприятиям"
      End
      Begin VB.Menu VentureRest 
         Caption         =   "Остатки по предприятиям"
      End
      Begin VB.Menu mnFiltrOborot 
         Caption         =   "Номенклатура для закупки"
         Visible         =   0   'False
      End
      Begin VB.Menu mnReservedAll 
         Caption         =   "Зарезервированная ном-ра"
      End
      Begin VB.Menu mnSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnSkladStand 
         Caption         =   "Состояние склада"
      End
      Begin VB.Menu mnKarta 
         Caption         =   "Карточка движения"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnServic 
      Caption         =   "Сервис"
      Begin VB.Menu mnOstat 
         Caption         =   "Проверка на отриц.остатки"
      End
      Begin VB.Menu mnViewOst 
         Caption         =   "Результаты посл. проверки"
      End
      Begin VB.Menu mnCurOstat 
         Caption         =   "Сверка текущих остатков"
      End
      Begin VB.Menu mnVentureOrder 
         Caption         =   "Накладные между предприятиями"
      End
      Begin VB.Menu mnSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnWebs 
         Caption         =   "Файлы для Web"
      End
      Begin VB.Menu mnWeb 
         Caption         =   "Файл остатков для WEB"
         Visible         =   0   'False
      End
      Begin VB.Menu mnSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnToExcelWeb 
         Caption         =   "Остатки по складу материалов"
      End
      Begin VB.Menu mnBrightAwards 
         Caption         =   "Остатки по складу Bright Awards"
      End
      Begin VB.Menu mnSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnToExcel 
         Caption         =   "Прайс-лист материалов"
      End
      Begin VB.Menu mnBrightAwardsPrice 
         Caption         =   "Прайс-лист Bright Awards технический"
      End
      Begin VB.Menu mnSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnPriceToExcel 
         Caption         =   "Прайс-лист Bright Awards"
      End
      Begin VB.Menu mnPricePM 
         Caption         =   "Прайс-лист Петровских Мастерских"
      End
      Begin VB.Menu mnPriceDealer 
         Caption         =   "Прайс-лист дилеров"
      End
      Begin VB.Menu mnPriceRA 
         Caption         =   "Прайс-лист РПФ"
      End
      Begin VB.Menu mnFilters 
         Caption         =   "WEB фильтры"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnGuides 
      Caption         =   "Справочники"
      Begin VB.Menu mnNomenc 
         Caption         =   "Номенклатура"
      End
      Begin VB.Menu mnProducts 
         Caption         =   "Готовые изделия"
      End
      Begin VB.Menu mnSource 
         Caption         =   "Поставщики"
      End
      Begin VB.Menu mnInside 
         Caption         =   "Внут.подразд-я"
      End
      Begin VB.Menu mnStatia 
         Caption         =   "Статьи затрат"
      End
      Begin VB.Menu mnGuideFormuls 
         Caption         =   "Формулы для прайса"
      End
      Begin VB.Menu mnConstants 
         Caption         =   "Константы для формул"
      End
      Begin VB.Menu mnProdCategory 
         Caption         =   "Категории Гот.Изделий"
      End
      Begin VB.Menu mnManag 
         Caption         =   "Менеджеры"
      End
   End
   Begin VB.Menu mnContext5 
      Caption         =   "Добавить удалить номенк-ру"
      Visible         =   0   'False
      Begin VB.Menu mnAdd5 
         Caption         =   "Добавить из Справ-ка"
      End
      Begin VB.Menu mnDel5 
         Caption         =   "Удалить"
      End
   End
End
Attribute VB_Name = "Documents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isLoad As Boolean


Dim quantity  As Long
Dim sum As Single
'Dim guideDist(10) As String
Dim mousCol As Long, mousRow As Long
Dim mousCol2 As Long, mousRow2 As Long
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы

Dim quantity2 As Long
Dim minut As Integer
Dim partial As Boolean
Dim destId As String

Const dcDate = 1
Const dcNumDoc = 2
'Const dcM = 3
Const dcSour = 3
Const dcDest = 4
Const dcNote = 6
Const dcVenture = 5
'Grid2
Const dnNomNom = 1
Const dnNomName = 2
Const dnQuant2 = 3
Const dnEdIzm2 = 4
Const dnQuant = 5
Const dnEdIzm = 6

Sub dummy()
Dim Note, ColWidth, Nomname, Value
End Sub

'reg ="","single","add"
Sub loadDocs(Optional reg As String = "")
Dim strWhere As String, I As Integer, str As String
 prevRow = -1
 Grid.Visible = False
 numExt = 255 'приходные накладные
 If reg = "" Then
    strWhere = getWhereByDateBoxes(Me, "sDocs.xDate", CDate("01.11.2000"))
 ElseIf reg = "docsFind" Then
    strWhere = "sDocs.numDoc=" & numDoc
 Else
    strWhere = "sDocs.numDoc=" & numDoc & " AND sDocs.numExt=" & numExt
 End If
 If strWhere <> "" Then strWhere = " AND " & strWhere
 
 Me.MousePointer = flexHourglass
 If reg <> "add" Then
    gridIsLoad = False
    quantity = 0
    clearGrid Grid
 End If
 
 sql = "SELECT sDocs.xDate, sDocs.Note, sDocs.numDoc, sDocs.numExt, " & _
 "sDocs.destId, sGuideSource.sourceName, GS.sourceName AS destName " & _
 ", v.ventureName as venture_name " & _
 "FROM sDocs INNER JOIN sGuideSource ON sDocs.sourId = sGuideSource.sourceId " & _
 "JOIN sGuideSource AS GS ON sDocs.destId = GS.sourceId " & _
 "left JOIN guideVenture v ON v.ventureId = sDocs.ventureId " & _
 "WHERE sDocs.numExt =255 " & strWhere & "  ORDER BY sDocs.xDate; "
' "WHERE ((" & str & " AND (GuideManag.Manag)='" & cbM.Text & "' " & strWhere & "));"
' Debug.Print sql
 Set tbDocs = myOpenRecordSet("##176", sql, dbOpenForwardOnly)
 If tbDocs Is Nothing Then End
 If Not tbDocs.BOF Then
  While Not tbDocs.EOF
    Grid.AddItem ""
    quantity = quantity + 1
    LoadDate Grid, quantity, dcDate, tbDocs!xDate, "dd.mm.yy"
'   str = tbDocs!numDoc & "/" & tbDocs!numExt
    str = tbDocs!numDoc
    Grid.TextMatrix(quantity, dcNumDoc) = str
    Grid.TextMatrix(quantity, 0) = tbDocs!destId
    Grid.TextMatrix(quantity, dcSour) = tbDocs!SourceName
    Grid.TextMatrix(quantity, dcDest) = tbDocs!destName
    Grid.TextMatrix(quantity, dcNote) = tbDocs!Note
    If Not IsNull(tbDocs!venture_name) Then _
        Grid.TextMatrix(quantity, dcVenture) = tbDocs!venture_name

    tbDocs.MoveNext
  Wend
End If
'Debug.Print sql
tbDocs.Close
rowViem quantity, Grid
Grid.Visible = True
If quantity > 0 Then
    If reg <> "add" Or quantity = 1 Then Grid.RemoveItem quantity + 1
    Grid.col = 1
    gridIsLoad = True '
    Grid.col = 2      'вызов loadDocNomenk
    Grid.row = quantity
'    loadDocNomenk
    On Error Resume Next
    Grid.SetFocus
    cmDel.Enabled = True
    Grid2.Visible = True
    laGrid2.Visible = True
    cmAdd2.Enabled = True
Else
    cmDel.Enabled = False
    Grid2.Visible = False
    laGrid2.Visible = False
    cmAdd2.Enabled = False
End If
gridIsLoad = True

Me.MousePointer = flexDefault
    
End Sub

Private Sub cbM_Change()

End Sub

Private Sub ckEndDate_Click()
If ckEndDate.Value = 1 Then
    tbEndDate.Enabled = True
Else
    tbEndDate.Enabled = False
End If

End Sub

Private Sub ckStartDate_Click()
If ckStartDate.Value = 1 Then
    tbStartDate.Enabled = True
Else
    tbStartDate.Enabled = False
End If

End Sub

Private Sub cmAdd_Click()
Dim str As String ', intNum As Integer, l As Long, il As Long
Dim strNow As String, DateFromNum As String, dNow As Date
 
numDoc = getNextDocNum()
numExt = 255

addDoc
loadDocs "add" ' не загружать все док-ты

Grid.col = dcSour
End Sub

Sub addDoc()

On Error GoTo ERR1

    sql = "insert into sdocs (numdoc, numExt, xDate, sourId, destId) values (" _
    & numDoc & ", " & numExt & ", convert(datetime, '" & Format(Now, "yyyy-mm-dd hh:mm:ss.000") & "', 121), 0, -1001, " _
    & ")"
        
      'Debug.Print sql
    myExecute "##sDocs.cmAdd", sql

    Exit Sub
ERR1:
    errorCodAndMsg "##tDocs update"
End Sub

Private Sub cmAdd2_Click()
If Grid.TextMatrix(mousRow, dcSour) = "" Then
    MsgBox "Заполните поле 'Откуда' в реестре документов", , "Предупреждение"
    Grid.col = dcSour
    On Error Resume Next
    Grid.SetFocus
    Exit Sub
End If
If Nomenklatura.isRegimLoad Then Unload Nomenklatura
Nomenklatura.Regim = "fromDocuments"
Nomenklatura.setRegim
Nomenklatura.Show vbModal
loadDocNomenk

Grid2.row = max(quantity2, 1)
Grid2_EnterCell
On Error Resume Next
Grid2.SetFocus

End Sub

Function backOstatki(strWhere) As Boolean
Dim str  As String

backOstatki = False
sql = "UPDATE sGuideNomenk INNER JOIN sDMC " & _
"ON sGuideNomenk.nomNom = sDMC.nomNom " & "SET sGuideNomenk.nowOstatki = " & _
"[sGuideNomenk].[nowOstatki]-[sDMC].[quant] " & strWhere

'предметов м. и не быть - поэтому ноль
If myExecute("##122", sql, 0) <= 0 Then backOstatki = True
End Function

Private Sub cmDel_Click()
'Dim strWhere As String


sql = "SELECT numDoc from sDMC WHERE (((numDoc)=" & numDoc & "));"
If Not byErrSqlGetValues("W##426", sql, tmpSng) Then GoTo EN2
If tmpSng <> 0 Then
    MsgBox "Если вы хотите удалить накладную, то сначала удалите " & _
    "из нее все предметы.", , "Удаление невозможно!"
    GoTo EN2
End If

If MsgBox("Удалить документ № '" & Grid.TextMatrix(mousRow, dcNumDoc) & _
"', Вы уверены?", vbYesNo Or vbDefaultButton2, "Подтвердите удаление") _
= vbNo Then GoTo EN1

wrkDefault.BeginTrans

'strWhere = "WHERE (((sDMC.numDoc)=" & numDoc & ") AND ((sDMC.numExt)=" & numExt & "));"
''увеличиваем остатки на конец - но предметов м. и не быть
'If Not backOstatki(strWhere) Then
'    wrkDefault.Rollback
'    GoTo EN1
'End If

'удаление док-та (а также соотв. записей из ДМЦ - т.к. разрешено каскадное удаление)
sql = "DELETE  From sDocs WHERE (((sDocs.numDoc)=" & numDoc & _
      ") AND ((sDocs.numExt)=" & numExt & "));"
'MsgBox sql
If myExecute("##121", sql) = 0 Then
    quantity = quantity - 1
    If quantity = 0 Then
        clearGridRow Grid, mousRow
    Else
        Grid.RemoveItem mousRow
    End If
    wrkDefault.CommitTrans
Else
    wrkDefault.Rollback
End If
EN1:
Grid2.Visible = False
'cmProduct.Enabled = False
laGrid2.Visible = False
EN2:
Grid_EnterCell
On Error Resume Next
Grid.SetFocus
End Sub

Private Sub cmDel2_Click()
Dim strWhere As String

strWhere = "WHERE (((sDMC.numDoc)=" & numDoc & ") AND ((sDMC.numExt)=" & _
numExt & ") AND ((sDMC.nomNom)='" & gNomNom & "'));"

If MsgBox("Удалить позицию № '" & Grid2.TextMatrix(mousRow2, dnNomNom) & _
"', Вы уверены?", vbYesNo Or vbDefaultButton2, "Подтвердите удаление") _
= vbNo Then GoTo EN1

wrkDefault.BeginTrans

'увеличиваем остатки на конец
If Not backOstatki(strWhere) Then
    wrkDefault.Rollback
    GoTo EN1
End If


sql = "DELETE  From sDMC  " & strWhere
If myExecute("##125", sql) = 0 Then
    quantity2 = quantity2 - 1
    If quantity2 = 0 Then
        clearGridRow Grid2, mousRow2
    Else
        Grid2.RemoveItem mousRow2
        Grid2_EnterCell
    End If
    wrkDefault.CommitTrans
Else
    wrkDefault.Rollback
End If
EN1:
On Error Resume Next
If quantity2 = 0 Then
    Grid.SetFocus
Else
    Grid2.SetFocus
End If

End Sub

Private Sub cmExit_Click()
frBad.Visible = False
End Sub

Private Sub cmKarta_Click()
lbBad_DblClick
End Sub

Private Sub cmLoad_Click()
laFiltr.Visible = False
loadDocs

End Sub

Function getNextNumExt() As Integer
Dim v As Variant

getNextNumExt = 0
sql = "SELECT Max(sDocs.numExt) AS Max_numExt From sDocs " & _
"WHERE (((sDocs.numDoc)=" & numDoc & "));"

If Not byErrSqlGetValues("##128", sql, v) Then Exit Function
If IsNumeric(v) Then
    getNextNumExt = v + 1
Else
    getNextNumExt = 1
End If

End Function


Private Sub cmProduct_Click()
'If Not docLock() Then Exit Sub
ReDim QQ(0)

Products.Regim = "select"
Products.Show vbModal
If UBound(QQ) > 0 Then ' что-то добавили
    If Not loadDocNomenk("check") Then
        backNomenk 'откат
        loadDocNomenk
    End If
Else
    loadDocNomenk
End If
'docLock "un"

End Sub

Private Sub Command1_Click()
Dim str As String, str2 As String

'str = strWhereByStEndDateBox(Me)
str2 = getWhereByDateBoxes(Me, "sDocs.xDate", CDate("01.11.2000"))
If str = str2 Then
    str = str & "   - совпадает"
Else
    str = str & "   - не совпадает с  Where = '" & str2 & "'"
End If
MsgBox "Where = '" & str & "'"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF7 Then
    mnDocFind_Click
End If
End Sub

Private Sub Form_Load()
Dim str As String ', i As Integer
oldHeight = Me.Height
oldWidth = Me.Width
isLoad = True
If otlad = "otlaD" Then
    mnFilters.Visible = True
    Me.BackColor = otladColor
End If

Me.Caption = "Склад. Приходные накладные.     " & mainTitle

If dostup = "a" Then
    mnOborot.Visible = True
    sourOborot.Visible = True
    ventureOborot.Visible = True
    mnFiltrOborot.Visible = True
    mnSep1.Visible = True
    mnKarta.Visible = True
End If

sql = "SELECT sGuideSource.sourceName From sGuideSource " & _
"WHERE (((sGuideSource.sourceId)>0)) ORDER BY sGuideSource.sourceName;"
Set table = myOpenRecordSet("##144", sql, dbOpenForwardOnly)
If table Is Nothing Then End
While Not table.EOF
    lbSource.AddItem table!SourceName
    table.MoveNext
Wend
table.Close
'lbSource.Height = 195 * lbSource.ListCount + 100

loadLbInside
initVentureLB

'Set wrkDefault = DBEngine.Workspaces(0) ' для орг-ии транзакций

tbStartDate.Text = Format(DateAdd("d", -7, CurDate), "dd/mm/yy")
tbEndDate.Text = Format(CurDate, "dd/mm/yy")

Grid.FormatString = "|<Дата|<№ Док-та|<Откуда|<Куда|<Предпр|<Примечание"
Grid.ColWidth(0) = 0
Grid.ColWidth(dcDate) = 800
Grid.ColWidth(dcNumDoc) = 915
'Grid.ColWidth(dcM) = 300
Grid.ColWidth(dcSour) = 1100
Grid.ColWidth(dcDest) = 1100
Grid.ColWidth(dcNote) = 1530
Grid.ColWidth(dcVenture) = 800

Grid2.FormatString = "|<Номер|<Название|кол-во|<Ед.измерения|кол-во|<Ед.изм.производства"
Grid2.ColWidth(0) = 0
Grid2.ColWidth(dnNomNom) = 0 '945
Grid2.ColWidth(dnNomName) = 2400 + 430 + 650 + 945
Grid2.ColWidth(dnEdIzm) = 0 '435
Grid2.ColWidth(dnQuant) = 0 '660
Grid2.ColWidth(dnEdIzm2) = 435
Grid2.ColWidth(dnQuant2) = 660

End Sub

Sub loadLbInside()
Dim I As Integer

sql = "SELECT sGuideSource.sourceId, sGuideSource.sourceName From sGuideSource " & _
"WHERE (((sGuideSource.sourceId)<-1000)) ORDER BY sGuideSource.sourceId DESC;"
Set table = myOpenRecordSet("##95", sql, dbOpenDynaset)
If table Is Nothing Then myBase.Close: End
ReDim insideId(0): I = 0: ' ReDim statiaId(0): j = 0
lbInside.Clear
While Not table.EOF
    lbInside.AddItem table!SourceName
    ReDim Preserve insideId(I)
    insideId(I) = table!sourceId
    I = I + 1
    table.MoveNext
Wend
table.Close
lbInside.Height = lbInside.Height + 195 * (lbInside.ListCount - 1)
End Sub


Private Sub Form_Resize()
Dim H As Integer, W As Integer

If WindowState = vbMinimized Then Exit Sub
On Error Resume Next
H = Me.Height - oldHeight
oldHeight = Me.Height
W = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + H
Grid.Width = Grid.Width + W / 2

Grid2.Height = Grid2.Height + H
Grid2.Width = Grid2.Width + W / 2
Grid2.Left = Grid2.Left + W / 2
cmLoad.Top = cmLoad.Top + H
cmAdd.Top = cmAdd.Top + H
cmDel.Top = cmDel.Top + H
cmAdd2.Top = cmAdd2.Top + H
cmProduct.Top = cmProduct.Top + H
cmDel2.Top = cmDel2.Top + H

End Sub

Private Sub Form_Unload(Cancel As Integer)
'tbSystem.Close
isLoad = False
If GuideSource.isLoad Then Unload GuideSource
If KartaDMC.isLoad Then Unload KartaDMC
If Nomenklatura.isRegimLoad Then Unload Nomenklatura
If Products.isLoad Then Unload Products
If VentureOrder.isLoad Then Unload VentureOrder

'myBase.Close
End Sub

Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If Grid.TextMatrix(mousRow, dcVenture) = "" Then
    cmAdd2.Enabled = False
    cmDel2.Enabled = False
Else
    cmAdd2.Enabled = True
    cmDel2.Enabled = True
End If
End Sub

Private Sub Grid_DblClick()
If mousRow = 0 Then Exit Sub
If Grid.CellBackColor = &H88FF88 Then
    If mousCol = dcSour Then
'        listBoxInGridCell lbGroup, Grid
        listBoxInGridCell lbSource, Grid, "select"
    ElseIf mousCol = dcDest Then
        listBoxInGridCell lbInside, Grid, "select"
    ElseIf mousCol = dcDate Then
        If MsgBox("Изменение даты документа задним числом может изменить " & _
        "результаты Отчетов. Если вы уверены в необходимости изменения даты " & _
        "нажмите <Да>.", vbYesNo Or vbDefaultButton2, "Подтвердите изменение " & _
        "Даты!") = vbYes Then textBoxInGridCell tbMobile, Grid
    ElseIf mousCol = dcVenture Then
        listBoxInGridCell lbVenture, Grid, Grid.TextMatrix(mousRow, mousCol)
    Else
        tbMobile.MaxLength = 50
        textBoxInGridCell tbMobile, Grid
    End If
End If

End Sub

Function loadDocNomenk(Optional reg As String = "") As Boolean
Dim il As Long, str As String, s As Single, I As Integer ', str2 As String
Dim msgOst As String, r As Single, o As Single

loadDocNomenk = True ' не надо отката - пока
msgOst = ""
Me.MousePointer = flexHourglass
Grid2.Visible = False

gDocDate = Grid.TextMatrix(mousRow, dcDate)
laGrid2.Caption = "Номенклатура по документу '" & numDoc & "'"
'Grid2.Clear

quantity2 = 0
clearGrid Grid2

sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.cod, " & _
"sGuideNomenk.Size, sGuideNomenk.ed_Izmer, sGuideNomenk.perList, " & _
"sGuideNomenk.ed_Izmer2,  sDMC.quant FROM sGuideNomenk INNER JOIN " & _
"(sDocs INNER JOIN sDMC ON (sDocs.numExt = sDMC.numExt) AND " & _
"(sDocs.numDoc = sDMC.numDoc)) ON sGuideNomenk.nomNom = sDMC.nomNom " & _
"WHERE (((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & ")) " & _
"ORDER BY sGuideNomenk.nomNom;"
'MsgBox sql
sum = 0
Set tbNomenk = myOpenRecordSet("##118", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then Exit Function
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    quantity2 = quantity2 + 1
    Grid2.TextMatrix(quantity2, dnNomNom) = tbNomenk!nomnom
    
    str = gNomNom
    If KartaDMC.DMCnomNomCur = tbNomenk!nomnom Then Grid2.row = quantity2  ' если загружена карта тек. номенклатуры - то подсветим ее
    gNomNom = str                                                          ' это вызывает Enter_Cell т.е. меняет gNomNom
    Grid2.TextMatrix(quantity2, dnNomName) = tbNomenk!cod & " " & _
            tbNomenk!Nomname & " " & tbNomenk!Size
'    s = Round(tbNomenk!quant, 2)
'    Grid2.TextMatrix(quantity2, dnEdIzm) = tbNomenk!ed_Izmer
'    Grid2.TextMatrix(quantity2, dnQuant) = s
    If Grid.TextMatrix(Grid.row, 0) = -1002 Then
        Grid2.TextMatrix(quantity2, dnEdIzm2) = tbNomenk!ed_izmer
        Grid2.TextMatrix(quantity2, dnQuant2) = Round(tbNomenk!quant, 2)
    Else
        Grid2.TextMatrix(quantity2, dnEdIzm2) = tbNomenk!ed_Izmer2
        Grid2.TextMatrix(quantity2, dnQuant2) = Round(tbNomenk!quant / tbNomenk!perList, 2)
    End If
'    Grid2.TextMatrix(quantity2, dnQuant2) = Round(s / tbNomenk!perList, 2)
    Grid2.AddItem ""
    tbNomenk.MoveNext
  Wend
  Grid2.RemoveItem quantity2 + 1
End If
tbNomenk.Close
Grid2.Visible = True
Me.MousePointer = flexDefault
End Function

Private Sub Grid_EnterCell()
If quantity = 0 Or Not gridIsLoad Then Exit Sub
 mousRow = Grid.row
 mousCol = Grid.col

numDoc = Grid.TextMatrix(mousRow, dcNumDoc)
destId = Grid.TextMatrix(mousRow, 0)
If prevRow <> mousRow And gridIsLoad Then
    prevRow = mousRow
    loadDocNomenk
End If
If mousCol = 0 Then Exit Sub

If mousCol = dcDest And quantity2 <> 0 Then GoTo AA
If mousCol = dcSour And destId = -1002 Then GoTo AA

If mousCol = dcDate Or mousCol > dcNumDoc Then
    Grid.CellBackColor = &H88FF88
Else
AA:  Grid.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid_GotFocus()
    cmProduct.Visible = False
    cmDel2.Enabled = False
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Grid_DblClick
'ElseIf KeyCode = vbKeyEscape Then
'    lbHide
End If

End Sub

Sub lbHide()
tbMobile.Visible = False
lbGroup.Visible = False
lbSource.Visible = False
lbInside.Visible = False
lbVenture.Visible = False
Grid.Enabled = True
On Error Resume Next
Grid.SetFocus
Grid_EnterCell
End Sub

Sub lbHide2()
tbMobile2.Visible = False
Grid2.Enabled = True
On Error Resume Next
Grid2.SetFocus
Grid2_EnterCell
End Sub

Public Sub initVentureLB()
' Сначала удаляем старые значения
While lbVenture.ListCount
    lbVenture.RemoveItem (0)
Wend

sql = "select * from GuideVenture where standalone = 0 and id_analytic is not null"

Set table = myOpenRecordSet("##72", sql, dbOpenForwardOnly)
If table Is Nothing Then myBase.Close: End

'lbVenture.AddItem "", 0
While Not table.EOF
    lbVenture.AddItem "" & table!ventureName & ""
    lbVenture.ItemData(lbVenture.ListCount - 1) = table!ventureId
    table.MoveNext
Wend
table.Close
lbVenture.Height = 225 * lbVenture.ListCount

End Sub


Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyEscape Then Grid.CellBackColor = Grid.BackColor
If KeyCode = vbKeyEscape Then Grid_EnterCell

End Sub

Private Sub Grid_LeaveCell()
If Grid.col <> 0 Then Grid.CellBackColor = Grid.BackColor

End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
End Sub

Private Sub Grid2_Click()
mousCol2 = Grid2.MouseCol
mousRow2 = Grid2.MouseRow
If quantity2 = 0 Then Exit Sub
If Grid2.MouseRow = 0 Then
    Grid2.CellBackColor = Grid2.BackColor
    If mousCol2 = dnQuant Then
        SortCol Grid2, mousCol2, "numeric"
    Else
        SortCol Grid2, mousCol2
    End If
    SortCol Grid2, mousCol2
    Grid2.row = 1    ' только чтобы снять выделение
'    Grid2_EnterCell
End If
Grid2_EnterCell
End Sub

Private Sub Grid2_DblClick()
If mousRow2 = 0 Then Exit Sub
If Grid2.CellBackColor = &H88FF88 Then
    textBoxInGridCell tbMobile2, Grid2
End If

End Sub

Private Sub Grid2_EnterCell()
If quantity2 = 0 Then Exit Sub
mousRow2 = Grid2.row
mousCol2 = Grid2.col

gNomNom = Grid2.TextMatrix(mousRow2, dnNomNom)

If mousCol2 = dnQuant2 Then
    Grid2.CellBackColor = &H88FF88
Else
    Grid2.CellBackColor = vbYellow
End If


End Sub

Private Sub Grid2_GotFocus()
'    cmAdd2.Visible = True
'    cmDel2.Visible = True
cmDel2.Enabled = (quantity2 > 0)
'If quantity2 > 0 Then
'    cmDel2.Enabled = True
'Else
'    cmDel2.Enabled = False
'End If
    
End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid2_DblClick

End Sub

Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyEscape Then Grid2.CellBackColor = Grid2.BackColor
If KeyCode = vbKeyEscape Then Grid2_EnterCell

End Sub

Private Sub Grid2_LeaveCell()
Grid2.CellBackColor = Grid2.BackColor

End Sub

Private Sub Grid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid2.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid2.ColWidth(Grid2.MouseCol)
End Sub

Private Sub lbBad_DblClick()
Dim I As Integer
I = InStr(lbBad.Text, "  ")
gNomNom = Left$(lbBad.Text, I - 1)
ReDim DMCnomNom(1)
DMCnomNom(1) = gNomNom
KartaDMC.Grid.Visible = False
KartaDMC.nomenkName = Mid$(lbBad.Text, I + 2)
KartaDMC.Show
'frBad.Visible = False
End Sub

Private Sub lbBad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbBad_DblClick
End Sub

Private Sub lbGroup_DblClick()
If lbGroup.ListIndex = 0 Then
    listBoxInGridCell lbSource, Grid
Else
    listBoxInGridCell lbInside, Grid
End If
End Sub

Private Sub lbGroup_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbGroup_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub lbInside_DblClick()
Dim id As String
Const INVENT = "Инвентаризация." ' точка для отличия такойже статьи затрат

id = insideId(lbInside.ListIndex)
If mousCol = dcSour Then
    If valueToDocsField("##96", id, "sourId") Then GoTo AA
Else
  If id = -1002 Then ' на склад Обрезков только по статьям Инвентаризации
     If Grid.TextMatrix(mousRow, dcSour) = "Поставщик" Then
        GoTo BB '      и Поставщик
     Else
        sql = "UPDATE sDocs, sGuideSource SET sDocs.sourId = " & _
        "[sGuideSource].[sourceId], sDocs.destId = " & id & _
        " WHERE (((sGuideSource.sourceName)='" & INVENT & "') AND " & _
        "((numDoc)=" & numDoc & ") AND ((numExt)=" & numExt & "));"
        If myExecute("##96", sql) = 0 Then
            Grid.TextMatrix(mousRow, dcSour) = INVENT
            GoTo AA
        End If
    End If
  Else
BB: If valueToDocsField("##96", id, "destId") Then
AA:     Grid.Text = lbInside.Text
        Grid.TextMatrix(mousRow, 0) = id
    End If
  End If
End If
'If lbInside.Text = str2 Then
'    MsgBox "В колонках 'Откуда' и 'Куда' недопустимы одинаковые значения", , "Предупреждение"
'    Exit Sub
'End If
lbHide


End Sub

Function valueToDocsField(myErrCod As String, Value As String, field As String) As Boolean
sql = "UPDATE sDocs  SET sDocs." & field & "=" & Value & _
" WHERE (((sDocs.numDoc)=" & numDoc & " AND (sDocs.numExt)=" & numExt & "));"
'MsgBox sql
valueToDocsField = False
If myExecute(myErrCod, sql) = 0 Then valueToDocsField = True
End Function

Private Sub lbInside_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbInside_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If


End Sub

Private Sub lbSource_DblClick()
sql = "UPDATE sDocs, sGuideSource SET sDocs.sourId = [sGuideSource].[sourceId] " & _
"WHERE (((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & _
") AND (([sGuideSource].[sourceName])='" & lbSource.Text & "'));"
'sql = "UPDATE sDocs, sGuideSource SET sDocs.sourId = sGuideSource.sourceId " & _
"WHERE (((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & _
") AND ((sGuideSource.sourceName)='" & lbSource.Text & "'));"
'MsgBox sql
If myExecute("##126", sql) = 0 Then Grid.Text = lbSource.Text
lbHide

End Sub

Private Sub lbSource_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbSource_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub


Private Sub lbVenture_DblClick()
Dim newNote As String, nCount As Integer

If lbVenture.Visible = False Then Exit Sub
sql = "select wf_make_venture_income(" & Grid.TextMatrix(mousRow, dcNumDoc) & ", " & lbVenture.ItemData(lbVenture.ListIndex) & ")"

'i = orderUpdate("##72", lbVenture.ItemData(lbVenture.ListIndex), "Orders", "ventureId")
byErrSqlGetValues "##126.1", sql, nCount
If nCount > 0 Then
    Grid.Text = lbVenture.Text
    newNote = getValueFromTable("sDocs", "Note", "numDoc = " & Grid.TextMatrix(mousRow, dcNumDoc))
    If IsNull(newNote) Then newNote = ""
    Grid.TextMatrix(mousRow, dcNote) = newNote
Else
    MsgBox "Изменение не произошло. Вероятно, была попытка измененить приход, который был сделан до начала работы предприятия", , "Передупреждение"
End If

lbHide


End Sub

Private Sub lbVenture_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lbVenture_DblClick
    ElseIf KeyCode = vbKeyEscape Then
        lbHide
    End If
End Sub



Private Sub mnBrightAwards_Click()

    Dim myRegim As String
    myRegim = "awardsWeb"
    ExcelParamDialog.mainReportTitle = getEffectiveSetting(myRegim & ".title", "Остатки по складу Bright Awards (без цен)")
    ExcelParamDialog.kegl = getEffectiveSetting(myRegim & ".kegl", 8)
    ExcelParamDialog.outputUE = getEffectiveSetting(myRegim & ".ue", True)
    ExcelParamDialog.Regim = myRegim
    ExcelParamDialog.withPrice = False
    ExcelParamDialog.showRabbat = False
    
    ExcelParamDialog.Show vbModal, Me
    If Not ExcelParamDialog.exitCode = vbOK Then
        Exit Sub
    End If
    BrightAwardsRestToExcel myRegim, , ExcelParamDialog.mainReportTitle, _
        ExcelParamDialog.kegl
    
End Sub

Private Sub mnBrightAwardsPrice_Click()
    Dim myRegim As String
    myRegim = "awards"
    ExcelParamDialog.mainReportTitle = getEffectiveSetting(myRegim & ".title", "Остатки по складу Bright Awards")
    ExcelParamDialog.kegl = getEffectiveSetting(myRegim & ".kegl", 8)
    ExcelParamDialog.outputUE = getEffectiveSetting(myRegim & ".ue", True)
    ExcelParamDialog.priceType = getEffectiveSetting(myRegim & ".pricetype", 0)
    ExcelParamDialog.commonRabbat = getEffectiveSetting(myRegim & ".rabbat", 0)
    ExcelParamDialog.Regim = myRegim
    ExcelParamDialog.withPrice = True
    
    ExcelParamDialog.Show vbModal, Me
    If Not ExcelParamDialog.exitCode = vbOK Then
        Exit Sub
    End If
    If ExcelParamDialog.outputUE Then
        BrightAwardsRestToExcel myRegim, , ExcelParamDialog.mainReportTitle, ExcelParamDialog.kegl, _
            ExcelParamDialog.priceType, ExcelParamDialog.commonRabbat
    Else
        BrightAwardsRestToExcel myRegim, ExcelParamDialog.RubRate, ExcelParamDialog.mainReportTitle, _
            ExcelParamDialog.kegl, ExcelParamDialog.priceType, ExcelParamDialog.commonRabbat
    End If
End Sub

Private Sub mnConstants_Click()
    GuideConstants.Show
End Sub

Private Sub mnCurOstat_Click()
Nomenklatura.Regim = "checkCurOstat"
Nomenklatura.Show
Nomenklatura.setRegim

End Sub

Private Sub mnDocFind_Click()
Static Value

AA:     Value = InputBox("Введите номер документа (заказа)", "Поиск", Value)
        If Value = "" Then Exit Sub
        If Not IsNumeric(Value) Then
            MsgBox "Номер должен быть числом"
            GoTo AA
        End If
laFiltr.Visible = False
numDoc = Value
loadDocs "docsFind"
End Sub

Private Sub mnFiltrOborot_Click()
Nomenklatura.Regim = "fltOborot"
Nomenklatura.Show
Nomenklatura.setRegim

End Sub

Private Sub mnGuideFormuls_Click()
GuideFormuls.Regim = ""
GuideFormuls.Show
End Sub


Private Sub mnInside_Click()
GuideInside.Show vbModal
End Sub


Private Sub mnKarta_Click()
Nomenklatura.Regim = "forKartaDMC"
Nomenklatura.Show
Nomenklatura.setRegim
End Sub

Private Sub mnManag_Click()
GuideManag.Show vbModal
End Sub

Private Sub mnNomenc_Click()
Dim n1 As Nomenklatura
    Set n1 = New Nomenklatura
    n1.Regim = ""
    n1.Show vbModeless
    n1.setRegim
End Sub

Private Sub mnOborot_Click()
Dim n1 As Nomenklatura
    Set n1 = New Nomenklatura
    n1.Regim = "asOborot"
    n1.Show vbModeless
    n1.setRegim
End Sub
'проверка на отрицат.остатки
Private Sub mnOstat_Click()
Dim ost As Single, bef As Integer, I As Integer

frBad.Visible = False
Me.MousePointer = flexHourglass
'sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.ostCheck, " & _
      "sGuideNomenk.begOstatki  FROM sGuideNomenk;"
sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.ostCheck FROM sGuideNomenk;"
Set tbNomenk = myOpenRecordSet("##134", sql, dbOpenDynaset)
'Set tbNomenk = myOpenRecordSet("##134", "GuideNomenk", dbOpenTable)
If tbNomenk Is Nothing Then GoTo EN1
If tbNomenk.BOF Then GoTo EN1
'quantity = 0
tbNomenk.MoveFirst
bilo = False
While Not tbNomenk.EOF
  ost = 0 'tbNomenk!begOstatki
'  If ost < 0 Then GoTo BIL0
  tbNomenk.Edit
  tbNomenk!ostCheck = ""
  tbNomenk.Update
  
  sql = "SELECT sDMC.quant, sDocs.xDate, sDocs.sourId, sDocs.destId " & _
  "FROM sDocs INNER JOIN sDMC ON (sDocs.numExt = sDMC.numExt) AND " & _
  "(sDocs.numDoc = sDMC.numDoc) WHERE (((sDMC.nomNom)='" & tbNomenk!nomnom & "')) " & _
  " ORDER BY sDocs.xDate;"
  Set tbDMC = myOpenRecordSet("##135", sql, dbOpenForwardOnly)
  If tbDMC Is Nothing Then GoTo EN1
  If tbDMC.BOF Then GoTo NXT
'  tbDMC.MoveFirst
  bef = 0

  While Not tbDMC.EOF
    I = DateDiff("d", begDate, tbDMC!xDate)
    If ost <= -0.01 And I <> bef Then GoTo NXT2
    bef = I
    If tbDMC!sourId < -1000 Then _
        ost = ost - tbDMC!quant
    If tbDMC!destId < -1000 Then _
        ost = ost + tbDMC!quant
    tbDMC.MoveNext
  Wend
NXT:
  tbDMC.Close
NXT2:
  If ost <= -0.01 Then
'        tbDMC.Close
BIL0:   bilo = True
        tbNomenk.Edit
        tbNomenk!ostCheck = "m"
        tbNomenk.Update
  End If
  tbNomenk.MoveNext
Wend
tbNomenk.Close
valueToSystemField Now(), "checkOstDate"
'tbSystem.Edit
'tbSystem!checkOstDate = Now()
'tbSystem.Update
mnViewOst_Click
EN1:
Me.MousePointer = flexDefault

End Sub

Private Sub mnOstVed_Click()
Dim n1 As Nomenklatura
    Set n1 = New Nomenklatura
    n1.Regim = "asOstat"
    n1.Show
    n1.setRegim
End Sub

Private Sub mnPricePM_Click()
    Dim myRegim As String
    myRegim = "pricePM"
    ExcelParamDialog.mainReportTitle = getEffectiveSetting(myRegim & ".title", "КОРПОРАТИВНЫЕ ПРИЗЫ И НАГРАДЫ (Каталог 2008-2009 Выпуск 5)")
    ExcelParamDialog.kegl = getEffectiveSetting(myRegim & ".kegl", 9)
    ExcelParamDialog.outputUE = getEffectiveSetting(myRegim & ".ue", True)
    ExcelParamDialog.Regim = myRegim
    
    ExcelParamDialog.doProdCategory = True
    ExcelParamDialog.withPrice = True
    ExcelParamDialog.Show vbModal, Me
    If Not ExcelParamDialog.exitCode = vbOK Then
        Exit Sub
    End If

    Dim reportRate As Double
    If ExcelParamDialog.outputUE Then
        reportRate = 1
    Else
        reportRate = ExcelParamDialog.RubRate
    End If
    PriceToExcel myRegim, reportRate, ExcelParamDialog.mainReportTitle, ExcelParamDialog.kegl, ExcelParamDialog.prodCategoryId

End Sub

Private Sub mnPriceToExcel_Click()

    Dim myRegim As String
    myRegim = "default"
    ExcelParamDialog.mainReportTitle = getEffectiveSetting(myRegim & ".title", "КОРПОРАТИВНЫЕ ПРИЗЫ И НАГРАДЫ (Каталог 2008-2009 Выпуск 5)")
    ExcelParamDialog.kegl = getEffectiveSetting(myRegim & ".kegl", 9)
    ExcelParamDialog.outputUE = getEffectiveSetting(myRegim & ".ue", True)
    ExcelParamDialog.Regim = myRegim
    
    ExcelParamDialog.doProdCategory = True
    ExcelParamDialog.withPrice = True
    ExcelParamDialog.Show vbModal, Me
    If Not ExcelParamDialog.exitCode = vbOK Then
        Exit Sub
    End If

    Dim reportRate As Double
    If ExcelParamDialog.outputUE Then
        reportRate = 1
    Else
        reportRate = ExcelParamDialog.RubRate
    End If
    PriceToExcel myRegim, reportRate, ExcelParamDialog.mainReportTitle, ExcelParamDialog.kegl, ExcelParamDialog.prodCategoryId

End Sub

Private Sub mnPriceDealer_Click()
    
    Dim myRegim As String
    myRegim = "dealer"
    ExcelParamDialog.mainReportTitle = getEffectiveSetting(myRegim & ".title", "Прайс-лист Bright Awards для Дилеров")
    ExcelParamDialog.kegl = getEffectiveSetting(myRegim & ".kegl", 9)
    ExcelParamDialog.outputUE = getEffectiveSetting(myRegim & ".ue", True)
    ExcelParamDialog.doProdCategory = True
    ExcelParamDialog.Regim = myRegim
    ExcelParamDialog.withPrice = True
    
    ExcelParamDialog.Show vbModal, Me
    If Not ExcelParamDialog.exitCode = vbOK Then
        Exit Sub
    End If
    
    Dim reportRate As Double
    If ExcelParamDialog.outputUE Then
        reportRate = 1
    Else
        reportRate = ExcelParamDialog.RubRate
    End If
    
    PriceToExcel myRegim, reportRate, ExcelParamDialog.mainReportTitle, ExcelParamDialog.kegl, ExcelParamDialog.prodCategoryId

End Sub

Private Sub mnPriceRA_Click()
    Dim myRegim As String
    myRegim = "agency"
    ExcelParamDialog.mainReportTitle = getEffectiveSetting(myRegim & ".title", "Прайс-лист Bright Awards для Рекламных Агенств")
    ExcelParamDialog.kegl = getEffectiveSetting(myRegim & ".kegl", 9)
    ExcelParamDialog.outputUE = getEffectiveSetting(myRegim & ".ue", True)
    ExcelParamDialog.doProdCategory = True
    ExcelParamDialog.showRabbat = True
    ExcelParamDialog.Regim = myRegim
    ExcelParamDialog.withPrice = True
    
    ExcelParamDialog.Show vbModal, Me
    If Not ExcelParamDialog.exitCode = vbOK Then
        Exit Sub
    End If

    Dim reportRate As Double
    If ExcelParamDialog.outputUE Then
        reportRate = 1
    Else
        reportRate = ExcelParamDialog.RubRate
    End If
    PriceToExcel myRegim, reportRate, ExcelParamDialog.mainReportTitle, ExcelParamDialog.kegl, ExcelParamDialog.prodCategoryId, ExcelParamDialog.commonRabbat
End Sub


Private Sub mnProdCategory_Click()
    GuideProdCategory.Show
End Sub

Private Sub mnProducts_Click()
    Products.Regim = "" ' Просто Справочник
    Products.Show vbModeless
End Sub

Private Sub mnReservedAll_Click()
    'Report.param1 = laOther.Caption
    Report.emptyColIndex = 1
    Report.groupIdColIndex = 0
    Report.subtitleColIndex = 2
    Report.numSortSecondColIndex = 0 ' по номеру группы
    Report.numSortThirdColIndex = 2 ' по названию номенклатуры
    Report.Subtitle = True
    
    Report.Regim = "reservedAll"
    Report.Sortable = True
    Set Report.Caller = Me
    Report.Show vbModal

End Sub

Private Sub mnSkladStand_Click()

    ReDim sqlRowDetail(1)
    ReDim aRowText(1)
    ReDim rowFormatting(1)
    ReDim aRowSortable(1)
    ReDim arowSubtitle(1)
    
    
    sqlRowDetail(1) = "call wf_nomenk_areport"
    aRowText(1) = " Текущее состояние склада"
    rowFormatting(1) = "#|<Номер ном.|<Название|Ед изм.|>Цена|>К-во Факт|>К-во Макс|>Сумма.Факт|>Сумма макс."
    aRowSortable(1) = True
    arowSubtitle(1) = True
    Set Report.Caller = Me
    Report.Regim = "aReportDetail"
    Report.param1 = 1
    
    Report.Show vbModal
    
    
End Sub

Private Sub mnSource_Click()
    GuideSource.Show vbModal
End Sub

Private Sub mnStatia_Click()
    GuideStatia.Show vbModal
End Sub

Private Sub mnToExcel_Click()
        
    Dim myRegim As String
    myRegim = "toExcel"
    ExcelParamDialog.mainReportTitle = getEffectiveSetting(myRegim & ".title", "ОСТАТКИ ПО СКЛАДУ РАСХОДНЫХ МАТЕРИАЛОВ И КОМПЛЕКТУЮЩИХ")
    ExcelParamDialog.kegl = getEffectiveSetting(myRegim & ".kegl", 8)
    ExcelParamDialog.outputUE = getEffectiveSetting(myRegim & ".ue", True)
    ExcelParamDialog.Regim = myRegim
    ExcelParamDialog.withPrice = True
    
    ExcelParamDialog.Show vbModal, Me
    If Not ExcelParamDialog.exitCode = vbOK Then
        Exit Sub
    End If
    If ExcelParamDialog.outputUE Then
        ostatToWeb myRegim, , ExcelParamDialog.mainReportTitle, ExcelParamDialog.kegl
    Else
        ostatToWeb myRegim, ExcelParamDialog.RubRate, ExcelParamDialog.mainReportTitle, ExcelParamDialog.kegl
    End If
End Sub

Private Sub mnToExcelWeb_Click()
    Dim myRegim As String
    myRegim = "toExcelWeb"
    ExcelParamDialog.mainReportTitle = getEffectiveSetting(myRegim & ".title", "ОСТАТКИ ПО СКЛАДУ РАСХОДНЫХ МАТЕРИАЛОВ И КОМПЛЕКТУЮЩИХ")
    ExcelParamDialog.kegl = getEffectiveSetting(myRegim & ".kegl", 8)
    ExcelParamDialog.Regim = myRegim
    ExcelParamDialog.withPrice = False
    
    ExcelParamDialog.Show vbModal, Me
    If Not ExcelParamDialog.exitCode = vbOK Then
        Exit Sub
    End If
    ostatToWeb myRegim, , ExcelParamDialog.mainReportTitle, ExcelParamDialog.kegl
End Sub

Private Sub mnVentureOrder_Click()
    VentureOrder.Show
End Sub

Private Sub mnViewOst_Click()

Me.MousePointer = flexHourglass
lbBad.Clear
laFrame.Caption = "Список номенклатуры с отрицательными остатками" & _
vbCrLf & "(проверка от " & Format(getSystemField("checkOstDate"), "dd.mm.yy hh:nn") & ")"

sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName From sGuideNomenk " & _
"WHERE (((sGuideNomenk.ostCheck)='m')) ORDER BY sGuideNomenk.nomNom;"
Set tbNomenk = myOpenRecordSet("##136", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If tbNomenk.BOF Then
    MsgBox "Записей с отрицательными остатками не обнаружено.", , "Результаты проверки"
Else
    While Not tbNomenk.EOF
        lbBad.AddItem tbNomenk!nomnom & "  " & tbNomenk!Nomname
        tbNomenk.MoveNext
    Wend
    tbNomenk.Close
    lbBad.ListIndex = 0
    frBad.Visible = True
End If
EN1:
Me.MousePointer = flexDefault
End Sub


Private Sub mnWeb_Click()
Me.MousePointer = flexHourglass
    
If MsgBox("По кнопке 'ДА' будет перезаписан файл складcких остатков для WEB." _
, vbDefaultButton2 Or vbYesNo, "Подтвердите запись") = vbNo Then Exit Sub

ostatToWeb

Me.MousePointer = flexDefault
End Sub


Private Sub BrightAwardsRestToExcel(Optional Regim As String = "", Optional RubRate As Double = 1, Optional mainReportTitle As String, Optional reportKegl As Integer = 10, Optional priceType = 0, Optional commonRabbat As Single = 0)
    Dim lastCol As String, lastColInt As Integer, I As Integer
    Dim currentSeriaId As Integer
    Dim currentProductId As Integer
    Dim RPF_Rate As Single
    
    If Regim = "awards" Then
        lastCol = "I"
    Else
        ' awardsWeb
        lastCol = "E"
    End If
    lastColInt = 5 + (Asc(lastCol) - Asc("E"))
    Dim priceRegim As String
    
    If priceType = 0 Then ' dealer
        priceRegim = "dealer"
    ElseIf priceType = 1 Then ' RPF
        priceRegim = "agency"
    Else
        priceRegim = "default"
    End If

    On Error GoTo ERR2
    Set objExel = New Excel.Application
    objExel.Visible = True
    objExel.SheetsInNewWorkbook = 1
    objExel.Workbooks.Add
    With objExel.ActiveSheet
        .Cells.Font.Size = reportKegl
        .Columns(1).columnWidth = 12.57
        .Columns(2).columnWidth = 39.71
        .Columns(3).columnWidth = 10
        .Columns(4).columnWidth = 6.2
        .Columns(5).columnWidth = 6.2
        Const priceColWidth = 6
        If Regim = "awards" Then
            .Columns(6).columnWidth = priceColWidth
            .Columns(7).columnWidth = priceColWidth
            .Columns(8).columnWidth = priceColWidth
            .Columns(9).columnWidth = priceColWidth
        End If

        ' печать стандартной шапки
        excelStdSchapka objExel, RubRate, mainReportTitle, lastCol, "МАРКМАСТЕР"
    
        exRow = 6
        sql = "call wf_report_bright_ostat"
        Set tbProduct = myOpenRecordSet("##331", sql, dbOpenDynaset)
        If tbProduct Is Nothing Then GoTo EN1
        If Not tbProduct.BOF Then
            While Not tbProduct.EOF
                If tbProduct!prSeriaId <> currentSeriaId Then
                    .Cells(exRow, 2).Value = tbProduct!serianame
                    .Cells(exRow, 2).Font.Bold = True
                    With .Range("A" & exRow & ":" & lastCol & exRow)
                        .Borders(xlEdgeTop).Weight = xlMedium
                        .Borders(xlEdgeBottom).Weight = xlThin
                        .Borders(xlEdgeRight).Weight = xlMedium
                    End With
                    
                    exRow = exRow + 1
                    cErr = setVertBorders(objExel, xlThin, lastColInt)
                    
                    .Cells(exRow, 1).Value = "Код"
                    .Cells(exRow, 2).Value = "Описание"
                    .Cells(exRow, 3).Value = "Размер"
                    .Cells(exRow, 4).Value = "Ед.изм."
                    .Cells(exRow, 5).Value = "Кол-во"
                    If Regim = "awards" Then
                        .Cells(exRow, 6).Value = "Цена 1"
                        .Cells(exRow, 7).Value = "Цена 2"
                        .Cells(exRow, 8).Value = "Цена 3"
                        .Cells(exRow, 9).Value = "Цена 4"
                    End If
                    exRow = exRow + 1
                End If

                If currentProductId <> tbProduct!prId Then
                    cErr = setVertBorders(objExel, xlThin, lastColInt)
                    For I = 1 To 4
                        .Cells(exRow, I).Font.Bold = True
                    Next I
                    .Cells(exRow, 1).Value = tbProduct!prName
                    .Cells(exRow, 2).Value = tbProduct!prDescript
                    .Cells(exRow, 3).Value = tbProduct!prSize
                    .Cells(exRow, 4).Value = "стр " & tbProduct!Page & "."
                    If Regim = "awards" Then
                        For I = 6 To 9
                            .Cells(exRow, I).Font.Bold = True
                        Next I
                        gain2 = tbProduct!gain2
                        gain3 = tbProduct!gain3
                        gain4 = tbProduct!gain4
                        ExcelProductPrices RPF_Rate, priceRegim, RubRate, exRow, 6, commonRabbat
                    End If
                    exRow = exRow + 1
                End If
                
                
                cErr = setVertBorders(objExel, xlThin, lastColInt)

                .Cells(exRow, 1).Value = tbProduct!Ncod
                .Cells(exRow, 1).HorizontalAlignment = xlHAlignRight
                .Cells(exRow, 2).Value = tbProduct!Nomname
                .Cells(exRow, 3).Value = tbProduct!Nsize
                .Cells(exRow, 4).Value = tbProduct!ed_Izmer2
                .Cells(exRow, 5).Value = tbProduct!qty_dost
                If Regim = "awards" Then
                    ExcelKolonPrices exRow, 6, RubRate, RPF_Rate
                End If
                
                currentSeriaId = tbProduct!prSeriaId
                currentProductId = tbProduct!prId
                
                exRow = exRow + 1
                tbProduct.MoveNext
            Wend
        End If
        tbProduct.Close
        With .Range("A" & exRow & ":" & lastCol & exRow)
            .Borders(xlEdgeTop).Weight = xlMedium
        End With
    
    End With
EN1:
    Set objExel = Nothing
    
    Exit Sub
ERR2:
    If cErr <> "424" And Err.number <> "424" Then GoTo ERR3  ' 424 - не дождались конца вывода закрыли док-т
    objExel.Quit
    Set objExel = Nothing
    Exit Sub
ERR3: MsgBox Error, , "Ошибка 429 - " & cErr '##429

End Sub

'запись Nomenks - !!! при изменении п\п зделать это и в Prior
'при этом Nomenklatura.nomencDostupOstatki заменить на sProducts.nomencOstatkiToGrid(-1)

'Эта програма выдает в файл или в MS Excel остатки по Складу по всей
'номенклатуре.
'Вся номенклатура у нас сгруппирована по классам, которые образуют
'древовидную структуру (см. Классификатор(слева)в Справочнике номенклатуры
'в программе stime)
'Классификатор реализован на табл. sGuideKlass а Справочник на sGuideNomenk.
'Их и их поля klassId parentKlassId и klassName надо заменить аналогами из
'базы Comtec(недостающие колонки можно добавить) $comtec$
Sub ostatToWeb(Optional toExel As String = "", Optional RubRate As Double = 1, Optional mainReportTitle As String, Optional reportKegl As Integer = 10)
Dim tmpFile As String, I As Integer, findId As Integer, str As String
Dim lastCol As String, lastColInt As Integer
Dim minusQuant   As Integer
minusQuant = 0

    On Error GoTo ERR2
    Set objExel = New Excel.Application
    objExel.Visible = True
    objExel.SheetsInNewWorkbook = 1
    objExel.Workbooks.Add
With objExel.ActiveSheet
    .Cells.Font.Size = 8

    If toExel = "toExcelWeb" Then
        lastCol = "E"
    Else
        lastCol = "H"
    End If
    lastColInt = 5 + (Asc(lastCol) - Asc("E"))
    
    ' печать стандартной шапки
    excelStdSchapka objExel, RubRate, mainReportTitle, lastCol, "МАРКМАСТЕР"

    exRow = 6
    Dim currentKlassId As Integer
    Dim withOstat As Integer, exCol As Long

    .Columns(1).columnWidth = 12.57
    .Columns(2).columnWidth = 39.71
    .Columns(3).columnWidth = 10
    .Columns(4).columnWidth = 6.2
    If Not toExel = "toExcelWeb" Then
        .Columns(5).columnWidth = 7: .Columns(5).HorizontalAlignment = xlHAlignRight
        .Columns(6).columnWidth = 7: .Columns(6).HorizontalAlignment = xlHAlignRight
        .Columns(7).columnWidth = 7: .Columns(7).HorizontalAlignment = xlHAlignRight
        .Columns(8).columnWidth = 7: .Columns(8).HorizontalAlignment = xlHAlignRight
    Else
        .Columns(5).columnWidth = 6.2
    End If
    
    'cErr = setVertBorders(objExel, xlMedium, lastColInt)
'xlColumnDataType
    'If cErr <> 0 Then GoTo ERR2
'xlDiagonalDown, xlDiagonalUp, xlEdgeBottom, xlEdgeLeft, xlEdgeRight
'xlEdgeTop, xlInsideHorizontal, or xlInsideVertical.
    With .Range("A" & exRow & ":" & lastCol & exRow)
        '.Borders(xlEdgeBottom).Weight = xlMedium ' xlThin
        '.Borders(xlEdgeTop).Weight = xlMedium
    End With
    'exRow = exRow + 1
'------------------------------------------------------------------------


    
  If toExel = "toExcelWeb" Then
    withOstat = 1
  End If
  sql = "call wf_report_mat_ost(" & withOstat & ")"
  
  Set tbProduct = myOpenRecordSet("##331", sql, dbOpenDynaset)
  If tbProduct Is Nothing Then GoTo EN1
  If Not tbProduct.BOF Then
      While Not tbProduct.EOF
        
        If tbProduct!KlassId <> currentKlassId Then
            str = tbProduct!klassName
            
            .Cells(exRow, 2).Value = str
            .Cells(exRow, 2).Font.Bold = True
            With .Range("A" & exRow & ":" & lastCol & exRow)
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeRight).Weight = xlMedium
            End With
            
            exRow = exRow + 1
            cErr = setVertBorders(objExel, xlThin, lastColInt)
            'If cErr <> 0 Then GoTo ERR2
            
            .Cells(exRow, 1).Value = "Код"
            .Cells(exRow, 2).Value = "Описание"
            .Cells(exRow, 3).Value = "Размер"
            .Cells(exRow, 4).Value = "Ед.изм."
            If toExel = "toExcelWeb" Then
                .Cells(exRow, 5).Value = "Кол-во"
                exCol = 6
            Else
                exCol = 5
            End If
            If Not toExel = "toExcelWeb" Then
                '.Cells(exRow, exCol).Value = "Цена УЕ"
                With .Range("A" & exRow & ":" & lastCol & exRow)
                    .Borders(xlEdgeBottom).Weight = xlThin
                    .Font.Italic = True
                    .HorizontalAlignment = xlHAlignCenter
                End With
                If Not IsNull(tbProduct!Kolon1) Then
                    .Cells(exRow, exCol).Value = Chr(160) & tbProduct!Kolon1
                    .Cells(exRow, exCol).Font.Bold = True
                End If
                If Not IsNull(tbProduct!Kolon2) Then
                    .Cells(exRow, exCol + 1).Value = Chr(160) & tbProduct!Kolon2
                    .Cells(exRow, exCol + 1).Font.Bold = True
                End If
                If Not IsNull(tbProduct!Kolon3) Then
                    .Cells(exRow, exCol + 2).Value = Chr(160) & tbProduct!Kolon3
                    .Cells(exRow, exCol + 2).Font.Bold = True
                End If
                If Not IsNull(tbProduct!Kolon4) Then
                    .Cells(exRow, exCol + 3).Value = Chr(160) & tbProduct!Kolon4
                    .Cells(exRow, exCol + 3).Font.Bold = True
                End If
            End If
            cErr = setVertBorders(objExel, xlThin, lastColInt)
            If cErr <> 0 Then GoTo ERR2
            exRow = exRow + 1
        End If
'---------------------------------------------------------------------------
'Далее выдаются параметры по каждой номенклатуре группы
        str = tbProduct!ed_Izmer2
        .Cells(exRow, 1).Value = tbProduct!cod
        .Cells(exRow, 2).Value = tbProduct!Nomname
        .Cells(exRow, 3).Value = tbProduct!Size
        .Cells(exRow, 4).Value = str
        If Not toExel = "toExcelWeb" Then
            ExcelKolonPrices exRow, exCol, RubRate
        Else
            tmpSng = Round(tbProduct!qty_dost - 0.499)
            If tmpSng < -0.01 Then
                minusQuant = minusQuant + 1 '************************
            End If
            .Cells(exRow, 5).Value = Round(tmpSng, 2)
        End If
        cErr = setVertBorders(objExel, xlThin, lastColInt)
        If cErr <> 0 Then GoTo ERR2
        
        currentKlassId = tbProduct!KlassId
        exRow = exRow + 1

        tbProduct.MoveNext
      Wend
    End If
    tbProduct.Close
EN1:
    With .Range("A" & exRow & ":" & lastCol & exRow)
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
End With

Set objExel = Nothing

'If (minusQuant > 0) Then MsgBox "Обнаружено " & minusQuant & " позиций с отрицательными остатками.", , "Предупреждение"
Exit Sub

ERR2:
Set objExel = Nothing
If cErr <> 424 And Err <> 424 Then GoTo ERR3 ' 424 - не дождались конца вывода закрыли док-т
Exit Sub

ERR1:
If Err = 76 Then
    MsgBox "Невозможно создать файл " & tmpFile, , "Error: Не обнаружен ПК или Путь к файлу"
ElseIf Err = 53 Then
    Resume Next ' файла м.не быть
ElseIf Err = 47 Then
    MsgBox "Невозможно создать файл " & tmpFile, , "Error: Нет доступа на запись."
ElseIf cErr <> 424 Then
    cErr = Err
ERR3: MsgBox Error, , "Ошибка 429-" & cErr '##429
    'End
End If

End Sub

Private Sub ExcelKolonPrices(exRow As Long, exCol As Long, RubRate As Double, Optional RPF_Rate As Single = 1)

    Dim cena2W As String
    cena2W = Chr(160) & Format(RPF_Rate * tbProduct!CENA_W * RubRate, "0.00") ' выводим как текст, т.к. "3.00" все равностанет "3"
    objExel.ActiveSheet.Cells(exRow, exCol).Value = cena2W
    
    Dim kolonok As Integer, optBasePrice As Double, margin As Double, iKolon As Integer, manualOpt As Boolean
    kolonok = tbProduct!kolonok
    margin = tbProduct!margin
    optBasePrice = tbProduct!CENA_W
    
    If kolonok > 0 Then
        manualOpt = False
    Else
        manualOpt = True
    End If
    
    For iKolon = 2 To Abs(kolonok)
        If manualOpt Then
            objExel.ActiveSheet.Cells(exRow, exCol - 1 + iKolon).Value = _
                Chr(160) & Format(RPF_Rate * tbProduct("CenaOpt" & CStr(iKolon)) * RubRate, "0.00")
        Else
            objExel.ActiveSheet.Cells(exRow, exCol - 1 + iKolon).Value = _
                Chr(160) & Format(RPF_Rate * RubRate * calcKolonValue(optBasePrice, margin, tbProduct!rabbat, Abs(kolonok), iKolon), "0.00")
        End If
    Next iKolon

End Sub

Private Sub mnWebs_Click()
Dim str As String, ch As String, slen As Integer, oper As String, I As Integer
Dim tmpFile As String ', filtrList As String

If MsgBox("По кнопке 'ДА' будет перезаписаны файлы для WEB: файл складcких " & _
"остатков и файл комплектации готовых изделий." _
, vbDefaultButton2 Or vbYesNo, "Подтвердите запись") = vbNo Then Exit Sub


Me.MousePointer = flexHourglass

sql = "UPDATE sGuideNomenk SET sGuideNomenk.web2 = '';"
If myExecute("##405", sql) <> 0 Then GoTo EN2


sql = "SELECT sProducts.nomNom, sGuideProducts.prName, sGuideProducts.prId " & _
"FROM sGuideProducts INNER JOIN sProducts ON sGuideProducts.prId = " & _
"sProducts.ProductId WHERE (((sGuideProducts.web)<>''));"
'MsgBox sql
Set tbProduct = myOpenRecordSet("##373", sql, dbOpenDynaset)

If Not tbProduct Is Nothing Then
  
  If tbProduct.BOF Then
    MsgBox "Ни одно изделие не помечено для Web", , "Файл комлектации не создан!"
  Else
    tmpFile = webProducts & "tmp"
    On Error GoTo ERR1
    Open tmpFile For Output As #1
'    On Error GoTo 0
    While Not tbProduct.EOF
'      sql = "UPDATE sGuideNomenk INNER JOIN sProducts ON sGuideNomenk.nomNom " & _
'      "= sProducts.nomNom   SET sGuideNomenk.web2 = 'web' " & _
'      "WHERE (((sProducts.ProductId)=" & tbProduct!prId & "));"
'      myExecute "##372", sql
    
      Print #1, tbProduct!prName & vbTab & tbProduct!nomnom
      tbProduct.MoveNext
    Wend
    Close #1
'    On Error Resume Next ' файла м.не быть
    Kill webProducts
'    On Error GoTo 0
    Name tmpFile As webProducts
  End If
  tbProduct.Close
End If

Documents.ostatToWeb 'именно в конце
    
GoTo EN2
ERR1:
If Err = 76 Then
    MsgBox "Невозможно создать файл " & tmpFile, , "Error: Не обнаружен ПК или Путь к файлу"
ElseIf Err = 53 Then
    Resume Next ' файла м.не быть
ElseIf Err = 47 Then
    MsgBox "Невозможно создать файл " & tmpFile, , "Error: Нет доступа на запись."
Else
    MsgBox Error, , "Ошибка 47-" & Err '##47
    'End
End If
EN2:
On Error Resume Next 'нужен, если фокус после нажатия передали другому приложению
On Error Resume Next
Grid.SetFocus
Me.MousePointer = flexDefault

End Sub

Private Sub sourOborot_Click()
Dim n1 As Nomenklatura
    Set n1 = New Nomenklatura
    n1.Regim = "sourOborot"
    n1.Show
    n1.setRegim

End Sub

Private Sub tbMobile_DblClick()
lbHide
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, I As Integer

If KeyCode = vbKeyReturn Then
 
 
 If mousCol = dcDate Then
     If Not isDateTbox(tbMobile, "fry") Then Exit Sub
     str = "'" & Format(tmpDate, "yyyy-mm-dd") & "'"
     If tmpDate > CurDate Then
        MsgBox "Дата документа не может быть в будущем", , "Недопустимое значение!"
        GoTo EN1
     End If
     sql = "UPDATE sDocs  SET  sDocs.[xDate] = " & str & " WHERE " & _
     "(((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & "));"
     If myExecute("##119", sql) <> 0 Then GoTo EN1
 ElseIf mousCol = dcNote Then
    If Not valueToDocsField("##119", "'" & tbMobile.Text & "'", "Note") _
            Then GoTo EN1
 End If
 Grid.TextMatrix(mousRow, mousCol) = tbMobile.Text
 lbHide
ElseIf KeyCode = vbKeyEscape Then
    KeyCode = 0
' If mousCol = gpName And frmMode = "productAdd" Then
'    frmMode = ""
'    Grid.RemoveItem Grid.Rows - 1
' End If
EN1:
 lbHide
End If

End Sub

Private Sub tbMobile2_DblClick()
lbHide2

End Sub
Sub msgOfLateEdit(delta As Single)

   If delta < 0 And DateDiff("d", gDocDate, CurDate) <> 0 Then
        MsgBox "Изменение количестра задним числом может " & _
        "привести к появлению отрицательных остатков.    " & _
        "Рекомендуется выполнить последующую проверку базы по команде " & _
        "'Проверка на Остатки' из меню 'Сервис'.", , "Предупреждение"
    End If
End Sub

Private Sub tbMobile2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nowOst As Single, rezerv As Single, quant As Single, delta As Single
Dim I As Integer, J As Integer, tmp As Long

If KeyCode = vbKeyReturn Then
    If Not isNumericTbox(tbMobile2, 0) Then Exit Sub
    
'    sql = "SELECT sGuideNomenk.nowOstatki, sGuideNomenk.perList, " & _
    "sGuideNomenk.ed_Izmer, sGuideNomenk.ed_Izmer2, sDMC.quant " & _
    "FROM sGuideNomenk INNER JOIN sDMC ON sGuideNomenk.nomNom = sDMC.nomNom  " & _
    "WHERE (((sDMC.numDoc)=" & numDoc & ") AND ((sDMC.numExt)=" & numExt & _
    ") AND ((sGuideNomenk.nomNom)='" & gNomNom & "'));"
    sql = "SELECT nowOstatki, perList FROM sGuideNomenk " & _
    "WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
    'MsgBox sql
    Set tbNomenk = myOpenRecordSet("##123", sql, dbOpenForwardOnly)
    
    
    quant = tbMobile2.Text
    delta = Round(quant, 0)
    If Grid.TextMatrix(Grid.row, 0) <> -1002 Then
        If delta <> quant Then
            MsgBox "Количество должно быть целым", , "Ошибка"
            Exit Sub
        End If
        quant = Round(quant * tbNomenk!perList)
    End If
    sql = "SELECT quant FROM  sDMC  WHERE (((nomNom)='" & gNomNom & "') AND " & _
    "((numDoc)=" & numDoc & " AND (numExt)=" & numExt & "))"
'    MsgBox sql
    Set tbDMC = myOpenRecordSet("##458", sql, dbOpenForwardOnly)
    
    wrkDefault.BeginTrans
    
    delta = tbDMC!quant ' это из ДМЦ
    delta = Round(quant - delta, 2)

    tbDMC.Edit
    tbDMC!quant = quant
    tbDMC.Update
    tbNomenk.Edit
    tbNomenk!nowOstatki = tbNomenk!nowOstatki + delta
    tbNomenk.Update
    
    wrkDefault.CommitTrans
    
    tbDMC.Close
    tbNomenk.Close
    
    msgOfLateEdit (delta)
    lbHide2
    tmp = Grid2.row
  
   loadDocNomenk


 'возможно, если в loadDocNomenk был откат то след-го If Else не надо
  If laFiltr.Visible Then   ' если вызов из карты
'    If KartaDMC.DMCnomNomCur = gNomNom Then
    For I = 1 To UBound(DMCnomNom)
        If DMCnomNom(I) = gNomNom Then  ' и если редактировалась ном-ра
            Timer1.Interval = 10        ' из карты
            Timer1.Enabled = True       ' то перерасчет карты
            Exit For
        End If
    Next I
  Else 'если это не вызов из карты то ее выгружаем  чтобы там не оставалась
    If KartaDMC.isLoad Then Unload KartaDMC '      необновленная информация
  End If ' хотя можно проверить и если ном-ры нет в карте то и не надо выгружать
 
 Grid2.row = tmp
 Grid2.col = dnQuant2
EN1:
On Error Resume Next
 Grid2.SetFocus
ElseIf KeyCode = vbKeyEscape Then
    KeyCode = 0
    lbHide2
End If
End Sub

Private Sub Timer1_Timer()
Dim I As Integer
    Timer1.Enabled = False
    Me.MousePointer = flexHourglass
'    KartaDMC.Grid.Visible = False
    KartaDMC.quantity = 0
    For I = 1 To UBound(DMCnomNom)
        KartaDMC.getKartaDMC DMCnomNom(I)
    Next I
'    KartaDMC.Grid.Visible = True
    KartaDMC.ZOrder
    Me.MousePointer = flexDefault
End Sub

Private Sub ventureOborot_Click()
    Analityc.applicationType = "stime"
    Analityc.managId = AUTO.cbM.Text
    Analityc.Show
End Sub
