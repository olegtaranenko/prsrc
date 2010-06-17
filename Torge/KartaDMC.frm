VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form KartaDMC 
   BackColor       =   &H8000000A&
   Caption         =   "Карточка движения"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   8235
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmExcel 
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   3240
      TabIndex        =   19
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CheckBox ckPerList 
      Caption         =   "В целых"
      Height          =   255
      Left            =   3780
      TabIndex        =   18
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "Печать"
      Height          =   315
      Left            =   7260
      TabIndex        =   17
      Top             =   5880
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "До"
      Height          =   315
      Left            =   3240
      TabIndex        =   16
      Top             =   5220
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "между"
      Height          =   315
      Left            =   5340
      TabIndex        =   15
      Top             =   5220
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmDocs 
      Caption         =   "Документ"
      Height          =   315
      Left            =   6840
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmCheck 
      Caption         =   "Проверить на минус.остатки"
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.CommandButton cmLoad 
      Caption         =   "Загрузить"
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox tbStartDate 
      Height          =   285
      Left            =   1620
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "01.01.01"
      Top             =   60
      Width           =   795
   End
   Begin VB.TextBox tbEndDate 
      Height          =   285
      Left            =   2700
      MaxLength       =   8
      TabIndex        =   1
      Top             =   60
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9340
      _Version        =   393216
      AllowBigSelection=   0   'False
      MergeCells      =   2
      AllowUserResizing=   1
   End
   Begin VB.Label Label5 
      Caption         =   "пос"
      Height          =   195
      Left            =   6720
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label laEnd 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      Height          =   285
      Left            =   6960
      TabIndex        =   13
      Top             =   330
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label laStart 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   285
      Left            =   5880
      TabIndex        =   12
      Top             =   330
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Период Проверки с"
      Height          =   195
      Left            =   4320
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label laQuant 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2460
      TabIndex        =   10
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Число записей:"
      Height          =   195
      Left            =   1140
      TabIndex        =   9
      Top             =   5940
      Width           =   1215
   End
   Begin VB.Label laBegin 
      Caption         =   "Установите Период загрузки и нажмите <Загрузить>"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   2880
      Width           =   4155
   End
   Begin VB.Label laPeriod 
      Caption         =   "Период Загрузки  с  "
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label laPo 
      Caption         =   "пос"
      Height          =   195
      Left            =   2460
      TabIndex        =   6
      Top             =   120
      Width           =   195
   End
   Begin VB.Menu mnContext 
      Caption         =   "Контекстное"
      Visible         =   0   'False
      Begin VB.Menu nmToBuff 
         Caption         =   "Копировать в буфер"
      End
   End
End
Attribute VB_Name = "KartaDMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isLoad As Boolean
Public nomenkName As String
Public Regim As String
Public quantity  As Long
Public DMCnomNomCur As String ' текущая номенклатура в групповой карте

Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim mousCol As Long, mousRow As Long

'Const ktDocType = 10 'скрыт
Const ktDate = 1
Const ktNomenk = 2
Const ktNomName = 3
Const ktDocNum = 4
Const ktSour = 5
Const ktDest = 6
Const ktEdIzm = 7
Const ktIn = 8
Const ktOut = 9
Const ktOstat = 10
'Const ktLastM = 10


Private Sub ckPerList_Click()
cmLoad_Click
End Sub


Private Sub cmCheck_Click()
    Me.MousePointer = flexHourglass
    Grid.Visible = False
    quantity = 0
    getKartaDMC DMCnomNom(1), "check"
    Grid.Visible = True
    Me.MousePointer = flexDefault
End Sub

Private Sub cmDocs_Click()
Exit Sub
'Кнопка отключена поскольку Расходные накладные убрали их этой программы
Dim I As Integer, str As String

DMCnomNomCur = Grid.TextMatrix(Grid.row, ktNomenk)
'numDoc = Documents.Grid.TextMatrix(Grid.row, ktDocNum)
getDocExtNomFromStr Grid.TextMatrix(Grid.row, ktDocNum)
numDoc = numDoc
Documents.laFiltr.Visible = True
Documents.Show
Documents.loadDocs "single"
End Sub

Private Sub cmExcel_Click()
GridToExcel Grid, "Карта движения по номенклатуре '" & gNomNom & "'"
End Sub

Private Sub cmLoad_Click()
Dim I As Integer

Me.MousePointer = flexHourglass
Grid.Visible = False
quantity = 0
For I = 1 To UBound(DMCnomNom())
    getKartaDMC DMCnomNom(I)
Next I
Grid.Visible = True
Me.MousePointer = flexDefault
End Sub

Public Sub controlVisible(en As Boolean)
    KartaDMC.cmCheck.Visible = en
'    KartaDMC.Label2.Visible = en
'    KartaDMC.laStart.Visible = en
'    KartaDMC.Label5.Visible = en
'    KartaDMC.laEnd.Visible = en
End Sub

'reg='check' - проверка по всему диапазону
Public Sub getKartaDMC(nNom As String, Optional reg As String = "")
Dim str2 As String, I As Integer, str As String, per As Single
Dim firstBad As Long, head As Integer, ed_izm As String, ed_izm2 As String
Dim prev As Integer, ost As Single, bOst As Single, iBef As Integer
Dim ost_outcome As Single, ost_income As Single

'нач. остатки
'sql = "SELECT sGuideNomenk.begOstatki, sGuideNomenk.nomName, "
sql = "SELECT sGuideNomenk.nomName, " & _
"sGuideNomenk.ed_Izmer, sGuideNomenk.ed_Izmer2, sGuideNomenk.perList " & _
"From sGuideNomenk WHERE (((sGuideNomenk.nomNom)='" & nNom & "'));"
bOst = 0
byErrSqlGetValues "##132", sql, nomenkName, ed_izm, ed_izm2, per ' $$4
If ckPerList.Value = 1 Then
    ed_izm = ed_izm2
Else
    per = 1
End If

strWhere = "(sDMC.nomNom) = '" & nNom & "'"
'изменения на начало периода
ost = 0
str = getWhereByDateBoxes(Me, "sDocs.xDate", begDate, "befo") 'до startDate
str2 = getWhereByDateBoxes(Me, "sDocs.xDate", begDate) ' между
If str = "error" Or str2 = "error" Then Exit Sub
If str <> "" Then
    sql = "SELECT " _
        & "Sum(if sourId <= -1000 then sDMC.quant else 0 endif) AS outcome" _
        & ", Sum(if destId <= -1000 then sDMC.quant else 0 endif) AS income" _
        & " FROM sDocs INNER JOIN " & _
        " sDMC ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc) " & _
        " WHERE (sourId >= -1000 or destId >= -1000) and " & str & " and " & strWhere
'    MsgBox "sql1=" & sql
'Debug.Print sql
    byErrSqlGetValues "##132", sql, ost_outcome, ost_income
    ost = ost_income - ost_outcome
    bOst = bOst + ost
End If

'изменения внутри периода
ost = bOst
If str2 <> "" Then strWhere = str2 & ") AND (" & strWhere
sql = "SELECT sDocs.numDoc, sDocs.numExt, sDocs.xDate, sDocs.sourId, sDocs.destId, " & _
"GS.sourceName as destName, sGuideSource.sourceName, sDMC.nomNom, sDMC.quant " & _
"FROM (sGuideSource INNER JOIN (sGuideSource AS GS INNER JOIN sDocs " & _
"ON GS.sourceId = sDocs.destId) ON sGuideSource.sourceId = sDocs.sourId) " & _
"INNER JOIN sDMC ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = " & _
"sDMC.numDoc) WHERE((" & strWhere & ")) ORDER BY sDocs.xDate;"
'MsgBox "sql2=" & sql
Set tbDMC = myOpenRecordSet("##130", sql, dbOpenForwardOnly)
If tbDMC Is Nothing Then Exit Sub
If quantity = 0 Then clearGrid Grid
'Grid.Visible = False
If UBound(DMCnomNom) = 1 Then
    head = 1
'    clearGrid Grid
    str = "         Остатки на начало периода ="
    Grid.MergeRow(1) = True
    Grid.TextMatrix(1, ktDest) = str
    Grid.TextMatrix(1, ktSour) = str
    Grid.TextMatrix(1, ktEdIzm) = ed_izm
'    Grid.TextMatrix(1, ktIn) = ""
'    Grid.TextMatrix(1, ktOut) = ""
    Grid.row = 1
    Grid.col = ktOstat
    Grid.CellFontBold = True
    'If ost < 0 Then Grid.CellForeColor = 200
    Grid.TextMatrix(1, ktOstat) = Round(bOst / per, 2)
Else
    head = 0
End If
prev = -1: iBef = -1: firstBad = 0
While Not tbDMC.EOF
    If quantity > 0 Or head = 1 Then Grid.AddItem ""
    I = DateDiff("d", begDate, tbDMC!xDate)
'    If i > iBef And ost < 0 And head = 1 Then ' если на посл. запись предыдущ дня
    If I > iBef And ost < 0 Then ' если на посл. запись предыдущ дня
        Grid.row = quantity + 1  ' были "-" остатки то предыдущ строку
        Grid.col = ktOstat       ' делаем красной
        Grid.CellForeColor = 200 '
        Grid.CellFontBold = True 'в 2х местах
        If firstBad = 0 Then firstBad = Grid.row ' первая строка с "-"
    End If
    iBef = I
    quantity = quantity + 1
    Grid.TextMatrix(quantity + head, ktDate) = Format(begDate + I, "dd.mm.yy")
    
    str = tbDMC!numDoc
    If tbDMC!numExt < 254 Then str = str & "/" & tbDMC!numExt
    Grid.TextMatrix(quantity + head, ktDocNum) = str
        
    Grid.TextMatrix(quantity + head, ktIn) = 0
    If tbDMC!destId < -1000 Then ' приход на склад
        Grid.TextMatrix(quantity + head, ktIn) = Round(tbDMC!quant / per, 2)
        ost = ost + tbDMC!quant
    End If
    Grid.TextMatrix(quantity + head, ktOut) = 0
    If tbDMC!sourId < -1000 Then ' уход со склада
        Grid.TextMatrix(quantity + head, ktOut) = Round(tbDMC!quant / per, 2)
        ost = ost - tbDMC!quant
    End If
    Grid.TextMatrix(quantity + head, ktEdIzm) = ed_izm
    Grid.TextMatrix(quantity + head, ktSour) = tbDMC!SourceName
    Grid.TextMatrix(quantity + head, ktDest) = tbDMC!destName
    
    Grid.TextMatrix(quantity + head, ktOstat) = Round(ost / per, 2)
'    Grid.TextMatrix(quantity + head, ktLastM) = tbDMC!lastM
'    Grid.TextMatrix(quantity + head, ktDocType) = tbDMC!docTypeId
    Grid.TextMatrix(quantity + head, ktNomenk) = nNom ' для определения DMCnomNomCur
    Grid.TextMatrix(quantity + head, ktNomName) = nomenkName
    tbDMC.MoveNext
Wend
tbDMC.Close
'Grid.Visible = True
laQuant.Caption = quantity
cmDocs.Enabled = False
If quantity > 0 Then cmDocs.Enabled = True
'If quantity > 0 Then 'коммент. т.к. м.б. '-' нач остатки
    Grid.row = quantity + head
    Grid.col = ktOstat
    Grid.CellFontBold = True
    If ost < 0 Then Grid.CellForeColor = 200    'в 2х местах
    If reg <> "" Then
        If firstBad = 0 And ost >= 0 Then
            MsgBox "Дней с '-' остатками НЕТ!", , "Результаты проверки"
            ValueToTableField "##137", "", "sGuideNomenk", "ostCheck", "byNomNom"
        Else
            Grid.row = firstBad
            ValueToTableField "##137", "'m'", "sGuideNomenk", "ostCheck", "byNomNom"
        End If
    End If
'End If

End Sub

Sub removeHead()
Dim l As Long

For l = 1 To Grid.Rows - 1
    If Grid.TextMatrix(l, ktDate) = "" Then
        Grid.RemoveItem l 'удаляем заголовок
        Exit For
    End If
Next l
End Sub

Private Sub cmPrint_Click()
Me.PrintForm

End Sub

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

Private Sub Form_Activate()

If UBound(DMCnomNom) = 1 Then
    Me.Caption = "Карточка движения по позиции № " & DMCnomNom(1) & _
                   " (" & nomenkName & ")"
    Grid.ColWidth(ktNomenk) = 0
    Grid.ColWidth(ktNomName) = 0
Else
    Grid.ColWidth(ktNomenk) = 930
    Grid.ColWidth(ktNomName) = 480
    Me.Caption = "Карточка движения по группе позиций"
End If
End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width


tbStartDate.Text = Format(begDate, "dd/mm/yy")
'KartaDMC.Caption = "Карточка движения по позиции № " & DMCnomNom & _
                   " (" & nomenkName & ")"
laStart.Caption = Format(begDate, "dd/mm/yy")
tbEndDate.Text = Format(CurDate, "dd/mm/yy")
laEnd.Caption = tbEndDate.Text

Grid.FormatString = "|<Дата|<Номенклатура|<Номер номенклатуры|<Документ|" & _
"<Откуда|<Куда|<Ед.измерения|Приход|Расход|Остаток"
Grid.ColWidth(0) = 0
Grid.ColWidth(ktDate) = 765
Grid.ColWidth(ktDest) = 700
Grid.ColWidth(ktDocNum) = 930
Grid.ColWidth(ktSour) = 1300
Grid.ColWidth(ktDest) = 1700
Grid.ColWidth(ktEdIzm) = 630
Grid.ColWidth(ktIn) = 800
Grid.ColWidth(ktOut) = 800
Grid.ColWidth(ktOstat) = 800

'Grid.ColWidth(ktDocType) = 0

isLoad = True
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
Grid.Width = Grid.Width + W
cmLoad.Top = cmLoad.Top + H
Label1.Top = Label1.Top + H
laQuant.Top = laQuant.Top + H
cmCheck.Top = cmCheck.Top + H
cmDocs.Top = cmDocs.Top + H
cmPrint.Top = cmPrint.Top + H
cmExcel.Top = cmExcel.Top + H
'.Left = .Left + w

End Sub

Private Sub Form_Unload(Cancel As Integer)
'DocFromKarta = False
'DMCnomNom = ""
DMCklass = gKlassId 'возможно это лишнее
isLoad = False
End Sub

Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If quantity = 0 Then Exit Sub
If mousRow = 0 Then
    Grid.CellBackColor = Grid.BackColor
    If mousCol = ktDate Then
        SortCol Grid, mousCol, "date"
    ElseIf mousCol < ktIn Then
'        SortCol Grid, mousCol, "numeric"
        SortCol Grid, mousCol
    End If
    Grid.row = 1    ' только чтобы снять выделение
'    Grid_EnterCell
End If

End Sub

Private Sub Grid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Dim date1 As Date, date2 As Date ' в 2 х местах
Dim date1S, date2S As String

date1S = sortGrid.TextMatrix(Row1, mousCol)
date2S = sortGrid.TextMatrix(Row2, mousCol)

'If Not IsDate(date1S) = "" And date2S = "" Then
'    Cmp = 0
'    Exit Sub
If Not IsDate(date1S) Then
    Cmp = -1
    GoTo CC:
ElseIf Not IsDate(date2S) Then
    Cmp = 1
    GoTo CC:
End If

date1 = date1S
date2 = date2S
If date1 > date2 Then
    Cmp = 1
ElseIf date1 < date2 Then
    Cmp = -1
Else
    Cmp = 0
End If
CC:
If trigger Then Cmp = -Cmp


End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
ElseIf Button = 2 Then
    Grid.col = Grid.MouseCol
    Grid.row = Grid.MouseRow
    On Error Resume Next
    Grid.SetFocus
    Grid.CellBackColor = vbButtonFace
    Me.PopupMenu mnContext

End If
End Sub

Private Sub nmToBuff_Click()
Grid.CellBackColor = Grid.BackColor
Clipboard.SetText Grid.Text
End Sub
