VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form WerkOrders 
   BackColor       =   &H8000000A&
   Caption         =   " "
   ClientHeight    =   5784
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   Icon            =   "CehOrders.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5784
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmEquip 
      Caption         =   "All"
      Height          =   315
      Index           =   0
      Left            =   3960
      TabIndex        =   16
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmNaklad 
      Caption         =   "Выписанные накладные"
      Height          =   315
      Left            =   1440
      TabIndex        =   15
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   435
      Left            =   5040
      TabIndex        =   13
      Top             =   2280
      Width           =   2295
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Идет загрузка..."
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   2175
      End
   End
   Begin VB.Timer Timer1 
      Left            =   7560
      Top             =   5400
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "Печать"
      Height          =   255
      Left            =   10980
      TabIndex        =   12
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmZagruz 
      Caption         =   "Загрузка"
      Height          =   315
      Left            =   8760
      TabIndex        =   11
      Top             =   5340
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmExAll 
      Caption         =   "Выход"
      Height          =   315
      Left            =   10740
      TabIndex        =   10
      Top             =   5340
      Width           =   975
   End
   Begin VB.TextBox tbNomZak 
      Height          =   285
      Left            =   3780
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.CheckBox chSingl 
      Caption         =   "только заказ"
      Height          =   195
      Left            =   2460
      TabIndex        =   7
      Top             =   60
      Width           =   1335
   End
   Begin VB.CheckBox chDetail 
      Caption         =   "Детальный <F2>"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox lbProblem 
      Height          =   1200
      Left            =   3300
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmRefresh 
      Caption         =   "Обновить"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   5400
      Width           =   915
   End
   Begin VB.ListBox lbStatus 
      Height          =   1776
      ItemData        =   "CehOrders.frx":030A
      Left            =   540
      List            =   "CehOrders.frx":0329
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.ListBox lbObrazec 
      Height          =   432
      ItemData        =   "CehOrders.frx":0361
      Left            =   1560
      List            =   "CehOrders.frx":036B
      TabIndex        =   2
      Top             =   4140
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lbMaket 
      Height          =   432
      ItemData        =   "CehOrders.frx":0378
      Left            =   2460
      List            =   "CehOrders.frx":0382
      TabIndex        =   1
      Top             =   4140
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   8700
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      Caption         =   "<F1>"
      Height          =   195
      Left            =   5040
      TabIndex        =   9
      Top             =   60
      Width           =   375
   End
   Begin VB.Menu mnNomZak 
      Caption         =   "Номер заказа"
      Visible         =   0   'False
      Begin VB.Menu mnFind 
         Caption         =   "Найти в Реестре приема"
      End
      Begin VB.Menu mnCancel 
         Caption         =   " "
      End
   End
End
Attribute VB_Name = "WerkOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public idWerk As Integer
Dim werkRows As Long, werkRowsOld As Long
Dim sum As Single
Dim marker As String ' символ в 0 колонке определяет тип lb, выз-го по muose
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim colWdth(20) As Integer
Public Regim As String ' режим окна
Public mousRow As Long    '
Public mousCol As Long    '
'Public werkId As Integer
Dim maxExt

Dim tbCeh As Recordset
Dim idEquip As Integer


Private Sub chDetail_Click()
Dim StatusId As String, Worktime As String, Left As String, Numorder As String, Outdatetime As String, Rollback As String

werkBegin
gridIsLoad = True
Grid.col = chKey
Grid.col = 1
Grid.SetFocus
End Sub

Private Sub chSingl_Click()
If chSingl.value = 1 And Not IsNumeric(tbNomZak.Text) Then
    MsgBox "Номер заказа выбран неверно.", , "Предупреждение:"
    chSingl.value = 0
    Exit Sub
End If
werkBegin
gridIsLoad = True
Grid.col = chKey
Grid.col = 1
Grid.SetFocus

End Sub

Private Sub cmEquip_Click(Index As Integer)
    idEquip = Index
    werkBegin
End Sub

Private Sub cmExAll_Click()
Unload Me
End Sub

Private Sub cmNaklad_Click()
sDocs.Regim = "fromCeh"
sDocs.idWerk = idWerk
sDocs.Show vbModal
End Sub

Private Sub cmPrint_Click()
Me.PrintForm
'Me.Height = 20000 это дает Err384, если форма уже максимизирована
End Sub

Private Sub cmRefresh_Click()
werkBegin
gridIsLoad = True
Grid.col = 1
End Sub

Private Sub cmZagruz_Click()
Zagruz.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then lbHide
If KeyCode = vbKeyF1 Then
    If chSingl.value = 1 Then
        chSingl.value = 0
    Else
        chSingl.value = 1
    End If
End If
End Sub

Sub werkBegin()
Dim str As String, I As Integer, J As Integer, IL As Long, tmpTopRow As Long
tmpTopRow = Grid.TopRow

#If onErrorOtlad Then
    On Error GoTo errMsg
    GoTo START
errMsg:
    MsgBox Error, , "Ошибка  " & Err & " в п\п werkBegin" '
    End
START:
#End If

gridIsLoad = False
Screen.MousePointer = flexHourglass

getNakladnieList "werk"

' запоминаем настройки столбцов
colWdth(chNomZak) = Grid.ColWidth(chNomZak)
colWdth(chM) = Grid.ColWidth(chM)
colWdth(chEquip) = Grid.ColWidth(chEquip)
colWdth(chStatus) = Grid.ColWidth(chStatus)
colWdth(chVrVip) = Grid.ColWidth(chVrVip)
colWdth(chProcVip) = Grid.ColWidth(chProcVip)
colWdth(chProblem) = Grid.ColWidth(chProblem)
colWdth(chDataVid) = Grid.ColWidth(chDataVid)
'colWdth(chDataRes) = Grid.ColWidth(chDataRes)
colWdth(chVrVid) = Grid.ColWidth(chVrVid)
colWdth(chFirma) = Grid.ColWidth(chFirma)
colWdth(chIzdelia) = Grid.ColWidth(chIzdelia)
colWdth(chLogo) = Grid.ColWidth(chLogo) + Grid.ColWidth(chDataRes)

Grid.Visible = False
For IL = Grid.Rows To 3 Step -1
    Grid.removeItem (IL)
Next IL
Grid.row = 1
For IL = 0 To Grid.Cols - 1
    Grid.col = IL
    Grid.CellBackColor = Grid.BackColor
    Grid.CellForeColor = vbBlack
    Grid.TextMatrix(1, IL) = ""
Next IL

' восстанавливаем настройки столбцов
Grid.ColWidth(chNomZak) = colWdth(chNomZak)
Grid.ColWidth(chM) = colWdth(chM)
Grid.ColWidth(chEquip) = colWdth(chEquip)
Grid.ColWidth(chVrVip) = colWdth(chVrVip)
Grid.ColWidth(chStatus) = colWdth(chStatus)
Grid.ColWidth(chProcVip) = colWdth(chProcVip)
Grid.ColWidth(chProblem) = colWdth(chProblem)
Grid.ColWidth(chDataVid) = colWdth(chDataVid)
Grid.ColWidth(chIzdelia) = colWdth(chIzdelia)

If chDetail.value = 1 Then
    Grid.ColWidth(chDataRes) = 740
Else
    Grid.ColWidth(chDataRes) = 0
End If
Grid.ColWidth(chVrVid) = colWdth(chVrVid)
Grid.ColWidth(chFirma) = colWdth(chFirma)
Grid.ColWidth(chLogo) = colWdth(chLogo) - Grid.ColWidth(chDataRes)

Dim EquipTitle As String, EquipSql As String
If idEquip = 0 Then
    EquipTitle = "All"
    EquipSql = ""
Else
    EquipTitle = Equip(idEquip)
    EquipSql = " AND equipId = " & idEquip
End If

Me.Caption = Werk(idWerk) & " - " & EquipTitle & "  " & mainTitle

' Сортируем, чтобы макет появился только один раз
sql = "select * from vw_Reestr where werkId = " & idWerk & EquipSql

'& " ORDER BY Numorder "

Set tbCeh = myOpenRecordSet("##34", sql, dbOpenDynaset)
If tbCeh Is Nothing Then myQuery.Close: myBase.Close: End

werkRows = 0
Dim MaketFlag As Boolean
Dim MaketNumorder As Long
MaketNumorder = 0

If Not tbCeh.BOF Then
  
  tbCeh.MoveFirst
  While Not tbCeh.EOF
    gNzak = tbCeh!Numorder
    If gNzak <> MaketNumorder Then
        MaketFlag = True
        MaketNumorder = gNzak
    End If
    
    If chSingl.value = 1 And gNzak <> tbNomZak.Text Then GoTo NXT
    If IsDate(tbCeh!DateTimeMO) Then
      If tbCeh!DateTimeMO < CDate("01.01.2000") _
        Or tbCeh!DateTimeMO > CDate("01.01.2050") _
      Then
        msgOfZakaz "##308", "Недопустимая дата МО. Обратитесь к менеджеру. ", tbCeh!Manag
        GoTo NXT
      End If
      If IsNull(tbCeh!workTimeMO) Then
        If MaketFlag Then
            toCehFromStr "m" 'макет
            MaketFlag = False
        End If
      Else  ' образец
        toCehFromStr "o" 'макет
      End If ' образец
    End If 'MO
MN:
    toCehFromStr '************************************
NXT:
    tbCeh.MoveNext
  Wend
End If
tbCeh.Close

Grid.col = chKey: Grid.Sort = 3 'числовое возр.
Grid.row = 1

If werkRows = werkRowsOld Then Grid.TopRow = tmpTopRow
werkRowsOld = werkRows

Grid.Visible = True
On Error Resume Next
Grid.SetFocus
Screen.MousePointer = flexDefault
Frame1.Visible = False
End Sub

Sub toCehFromStr(Optional isMO As String = "")
Dim str As String, I As Integer, J As Integer, K As Integer, S As Variant
Dim color As Long, str1 As String  ', is100 As Boolean

#If onErrorOtlad Then
    On Error GoTo errMsg
    GoTo START
errMsg:
    MsgBox Error, , "Ошибка  " & Err & " в п\п toCehFromStr" '
    End
START:
#End If

K = 0
marker = ""
color = vbBlack
'If sampl = "" Then
If isMO <> "o" Then
    str = ""
    If tbCeh!StatusId = 2 Then 'резерв
        color = vbBlue
    ElseIf tbCeh!StatusId = 3 Or tbCeh!StatusId = 9 Then ' согласов
        color = &HAA00& ' т.зел.
    ElseIf tbCeh!StatusId = 5 Then ' отложен
        marker = "р"
        color = vbRed
    ElseIf tbCeh!StatusId = 1 Or tbCeh!StatusId = 8 Or tbCeh!StatusId = 4 Then ' в работе и готов
        marker = "р"
    End If
Else
    marker = "о"
    str = "o"
End If

If isMO = "m" Then ' макет
    If werkRows > 0 Then Grid.AddItem ("")
    str = "м"
    werkRows = werkRows + 1
    If tbCeh!StatM = "готов" Then
        Grid.TextMatrix(werkRows, chStatus) = tbCeh!StatM
    Else
        Grid.TextMatrix(werkRows, chStatus) = ""
    End If
    marker = "м"
    LoadDateKey tbCeh!DateTimeMO, "##38"
    LoadDate Grid, werkRows, chVrVid, tbCeh!DateTimeMO, "hh"
    GoTo MN
End If

    If werkRows > 0 Then Grid.AddItem ("") 'кусок оформляем как осн.часть
    werkRows = werkRows + 1
    
    Grid.TextMatrix(werkRows, chEquip) = tbCeh!Equip
    Grid.TextMatrix(werkRows, chEquipId) = tbCeh!equipId
    
    Grid.col = chNomZak
    Grid.row = werkRows
    Grid.CellForeColor = color
 
    If str = "" Then 'осн.часть заказа
        S = Round(100 * (1 - tbCeh!nevip), 1)
        If S > 0 Then Grid.TextMatrix(werkRows, chProcVip) = S
        
        S = tbCeh!Worktime
        LoadDateKey tbCeh!Outdatetime, "##36"
        LoadDate Grid, werkRows, chVrVid, tbCeh!Outdatetime, "hh"
    Else
        If tbCeh!StatO = "готов" Then _
            Grid.TextMatrix(werkRows, chProcVip) = "100"
        S = tbCeh!workTimeMO
        If S < 0 Then S = -S
        LoadDateKey tbCeh!DateTimeMO, "##36"
        LoadDate Grid, werkRows, chVrVid, tbCeh!DateTimeMO, "hh"
    End If
    If IsNull(S) Then
        msgOfZakaz ("##36"), , tbCeh!Manag
        Grid.TextMatrix(werkRows, chVrVip) = "(??) "
    Else
      If chDetail.value = 1 Then '
        Grid.TextMatrix(werkRows, chVrVip) = "(" & S & ")"
      Else
        Grid.TextMatrix(werkRows, chVrVip) = Round(S, 2)
      End If
    End If
If isMO = "o" Then
   If tbCeh!StatO = "готов" Then
     Grid.TextMatrix(werkRows, chStatus) = tbCeh!StatO 'образец
   Else
     Grid.TextMatrix(werkRows, chStatus) = "" 'образец
   End If
ElseIf (tbCeh!StatusId = 1 Or tbCeh!StatusId = 8) And Not IsNumeric(tbCeh!Stat) Then
    If Not IsNull(tbCeh!Stat) Then Grid.TextMatrix(werkRows, chStatus) = tbCeh!Stat
ElseIf tbCeh!StatusId = 2 Then ' резерв
    str1 = "Р": GoTo AA
ElseIf tbCeh!StatusId = 3 Or tbCeh!StatusId = 9 Then  ' согласов
    str1 = "С"
AA: Grid.col = chStatus
    Grid.CellForeColor = color
    Grid.TextMatrix(werkRows, chStatus) = str1 & " на " & Format(tbCeh!DateRS, "dd.mm.yy")
Else
    Grid.TextMatrix(werkRows, chStatus) = Status(tbCeh!StatusId)
End If
MN:
#If Not COMTEC = 1 Then '----------------------------------------------
 For I = 1 To UBound(tmpL) 'отмечаем заказы с выписанными накладными
    If tmpL(I) = gNzak Then
        Grid.col = chIzdelia
        Grid.row = werkRows
        Grid.CellForeColor = 200
        Exit For
    End If
 Next I
#End If '--------------------------------------------------------------
Grid.TextMatrix(werkRows, 0) = marker
Grid.TextMatrix(werkRows, chNomZak) = gNzak & str
If str <> "" Then colorGridRow Grid, werkRows, &HCCCCCC 'маркируем МО
Grid.TextMatrix(werkRows, chM) = tbCeh!Manag
Grid.TextMatrix(werkRows, chFirma) = tbCeh!name
Grid.TextMatrix(werkRows, chLogo) = tbCeh!Logo
Grid.TextMatrix(werkRows, chIzdelia) = tbCeh!Product
If tbCeh!StatusId = 5 Then ' отложен
        Grid.TextMatrix(werkRows, chProblem) = Problems(tbCeh!ProblemId)
End If

End Sub

Sub LoadDateKey(val As Variant, myErr As String)
Dim I As Integer

If Not IsNull(val) Then
  If IsDate(val) Then
    Grid.TextMatrix(werkRows, chDataVid) = Format(val, "dd.mm.yy")
    I = DateDiff("d", curDate, val) + 1 'здесь
    Grid.TextMatrix(werkRows, chKey) = I
'    If i = stDay Then
'        Grid.col = chDataVid
'        Grid.CellForeColor = &H8800&
'        Grid.CellFontBold = True
'    End If
    Exit Sub
  End If
End If
msgOfZakaz myErr, , tbCeh!Manag
Grid.TextMatrix(werkRows, chDataRes) = "??"
Grid.TextMatrix(werkRows, chKey) = 0
End Sub

Private Sub Form_Load()

Dim I As Integer


oldHeight = Me.Height
oldWidth = Me.Width

#If Not COMTEC = 1 Then '----------------------------------------------
cmNaklad.Visible = True
#End If '--------------------------------------------------------------

'If dostup = "y" Or dostup = "c" Then cmZagruz.Visible = True
If Not (dostup = "a" Or dostup = "m" Or dostup = "" Or dostup = "b") Then
    cmZagruz.Visible = True
    Orders.managLoad "fromCeh" ' загрузка Manag()
End If
If dostup = "" Then cmNaklad.Visible = False


Screen.MousePointer = flexHourglass

For I = begWerkProblemId To lenProblem
    lbProblem.AddItem Problems(I)
Next I

Grid.FormatString = "    |<№ заказа|^М|Оборуд|Статус |>Вр.вып|>%вы|Проблемы|" & _
"<Дата выдачи|<Вр.выд|<дата ресурса|<Заказчик|<Лого|<Изделия|№Дня|equipid"

Grid.ColWidth(chM) = 270
Grid.ColWidth(chVrVip) = 388
Grid.ColWidth(chEquip) = 570
Grid.ColWidth(chStatus) = 870
Grid.ColWidth(chProcVip) = 420
Grid.ColWidth(chProblem) = 900
Grid.ColWidth(chDataRes) = 735
Grid.ColWidth(chVrVid) = 330
Grid.ColWidth(chDataVid) = 735
Grid.ColWidth(chFirma) = 2000
Grid.ColWidth(chLogo) = 1200
Grid.ColWidth(chKey) = 0 ' ДЛЯ СОРТИРОВКИ по дате
Grid.ColWidth(chEquipId) = 0
Grid.ColWidth(0) = 0
Grid.ColWidth(chNomZak) = 1000
Grid.ColWidth(chIzdelia) = 2450

Dim RightLinie As Long, HShift As Long
Dim equipIndex As Integer
    
    HShift = cmEquip(0).Width + 20
    RightLinie = cmEquip(0).Left + HShift
    
    sql = "select e.equipId, we.werkId, e.equipName, we.equipId as IsPresent " _
        & " from GuideEquip e " _
        & " LEFT JOIN WerkEquip we ON we.equipId = e.equipId AND we.werkId = " & idWerk _
        & "WHERE e.equipId > 0" _
        & " order by e.equipId"

    Set tbOrders = myOpenRecordSet("##we.01", sql, dbOpenForwardOnly)
    If Not tbOrders Is Nothing Then
        While Not tbOrders.EOF
            equipIndex = tbOrders!equipId
            Load cmEquip(equipIndex)
            If Not IsNull(tbOrders!IsPresent) Then
                cmEquip(equipIndex).Caption = tbOrders!equipName
                cmEquip(equipIndex).Visible = True
                cmEquip(equipIndex).Left = RightLinie
                RightLinie = RightLinie + HShift
            Else
                cmEquip(equipIndex).Visible = False
            End If
            tbOrders.MoveNext
        Wend
        tbOrders.Close
    End If


Timer1.Interval = 500
Timer1.Enabled = True 'вызов werkBegin

End Sub

Private Sub Form_Resize()
Dim H As Integer, W As Integer, I As Integer

If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next
lbHide
H = Me.Height - oldHeight
oldHeight = Me.Height
W = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + H
Grid.Width = Grid.Width + W
cmRefresh.Top = cmRefresh.Top + H
cmExAll.Top = cmExAll.Top + H
cmExAll.Left = cmExAll.Left + W
cmZagruz.Top = cmZagruz.Top + H
cmZagruz.Left = cmZagruz.Left + W
cmPrint.Left = cmPrint.Left + W
cmNaklad.Top = cmNaklad.Top + H

Dim RightLine As Integer

For I = 0 To cmEquip.UBound
    cmEquip(I).Top = cmEquip(I).Top + H
    If RightLine < cmEquip(I).Left + cmEquip(I).Width Then
        RightLine = cmEquip(I).Left + cmEquip(I).Width
    End If
Next I


End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (dostup = "a" Or dostup = "m" Or dostup = "" Or dostup = "b") Then
    exitAll 'для цехов
End If
isWerkOrders = False
End Sub


Private Sub Grid_Click()
If Not gridIsLoad Then Exit Sub
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If mousRow = 0 Then
    Grid.CellBackColor = Grid.BackColor
    If mousCol = 0 Then Exit Sub
    If mousCol = chNomZak Then
        SortCol Grid, mousCol
    ElseIf mousCol = chDataRes Or mousCol = chDataVid Then
        SortCol Grid, mousCol, "date"
    Else
        SortCol Grid, mousCol
    End If
    Grid.row = 1    ' только чтобы снять выделение
    Grid_EnterCell

End If
End Sub

Private Sub Grid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Dim date1 As Date, date2 As Date ' в 2 х местах
Dim date1S, date2S As String

date1S = sortGrid.TextMatrix(Row1, mousCol)
date2S = sortGrid.TextMatrix(Row2, mousCol)

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

Private Sub Grid_DblClick()

'If mousCol = chNomZak And dostup <> "c" And dostup <> "y" Then
If mousCol = chNomZak And (dostup = "a" Or dostup = "m" Or dostup = "" Or _
dostup = "b") Then Me.PopupMenu mnNomZak

getNumFromStr (Grid.TextMatrix(mousRow, chNomZak))

#If Not COMTEC = 1 Then '----------------------------------------------
If mousCol = chIzdelia And Grid.CellForeColor = 200 Then
    numDoc = gNzak
    numExt = 0
    Nakladna.Regim = "predmeti"
    Nakladna.idWerk = Me.idWerk
    Nakladna.Show vbModal
End If
#End If '--------------------------------------------------------------

If dostup = "" Then Exit Sub
marker = Grid.TextMatrix(mousRow, 0)
If mousRow = 0 Or marker = "" Then Exit Sub

If mousCol = chStatus Then
    If marker = "о" Then '  "образец"
        listBoxInGridCell lbObrazec, Grid, "select"
    ElseIf marker = "м" Then '      "макет"
        listBoxInGridCell lbMaket, Grid, "select"
    ElseIf LCase$(marker) = "р" Then '  "в работу"
        listBoxInGridCell lbStatus, Grid, "select"
    End If
End If
End Sub

Private Sub Grid_EnterCell()
If Not gridIsLoad Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col
getNumFromStr (Grid.TextMatrix(mousRow, chNomZak))
tbNomZak.Text = gNzak
If dostup = "" Then Exit Sub
marker = Grid.TextMatrix(mousRow, 0)
oldCellColor = Grid.CellBackColor
If (mousCol = chStatus And marker <> "") Or _
(mousCol = chIzdelia And Grid.CellForeColor = 200) Then
    Grid.CellBackColor = &H88FF88
Else
    Grid.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid_DblClick

End Sub

Private Sub Grid_LeaveCell()
If Not gridIsLoad Then Exit Sub
Grid.CellBackColor = oldCellColor
End Sub

Private Sub lbMaket_DblClick()
Dim I As Integer

If noClick Then Exit Sub

sql = "SELECT StatM From OrdersInCeh WHERE (((numOrder)=" & gNzak & "));"
If Not byErrSqlGetValues("##312", sql, tmpStr) Then Exit Sub
If tmpStr = "утвержден" Then
    msgZakazDeleted "макет уже утвержден"
    GoTo EN1
ElseIf lbMaket.Text = "готов" Then
    I = ValueToTableField("W##37", "'готов'", "OrdersInCeh", "StatM")
Else
    I = ValueToTableField("W##37", "'в работе'", "OrdersInCeh", "StatM")
End If
If I = 0 Then
    Grid.TextMatrix(mousRow, chStatus) = lbMaket.Text
ElseIf I = -1 Then
    msgZakazDeleted
End If
EN1:
lbHide ' в т.ч. подсветка

End Sub

Private Sub lbMaket_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbMaket_DblClick
End Sub

Private Sub lbObrazec_DblClick()
Dim J As Integer, str As String, old As String, V As Variant
Dim proc As String, Status As String
'sChr As String, dChr As String,
If noClick Then Exit Sub
old = Grid.TextMatrix(mousRow, chStatus)
If lbObrazec.Text = "готов" And lbObrazec.Text <> old Then
    proc = "100%": Status = "'готов'"
ElseIf lbObrazec.Text <> old Then '              образец
    proc = "0%": Status = "'в работе'"
Else
    lbHide
    Exit Sub
End If
lbObrazec.Visible = False

wrkDefault.BeginTrans
    
gEquipId = Grid.TextMatrix(mousRow, chEquipId)
V = makeProcReady(proc, gEquipId, "obraz")
If IsNull(V) Then ' образец утвержден
    msgZakazDeleted "образец уже утвержден"
ElseIf V Then
    If ValueToTableField("##54", Status, "OrdersEquip", "StatO", "byEquipId") = 0 Then
        wrkDefault.CommitTrans
        werkBegin
    Else
        wrkDefault.Rollback
    End If
Else ' заказ уже удален Менеджером
    wrkDefault.Rollback
    msgZakazDeleted
End If

lbHide ' в т.ч. подсветка
End Sub

Private Sub lbObrazec_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbObrazec_DblClick
End Sub

Private Sub lbProblem_DblClick()
Dim str As String, I As Integer

If noClick Then Exit Sub

wrkDefault.BeginTrans   ' начало транзакции

gEquipId = Grid.TextMatrix(mousRow, chEquipId)
I = ValueToTableField("W##41", "'в работе'", "OrdersEquip", "Stat", "byEquipId") 'т.к если оставить Stat=готов, то на завтра он удалиться
If I = 0 Then
    If ValueToTableField("##41", "5", "Orders", "StatusId") <> 0 Then GoTo ER1

    str = lbProblem.ListIndex + begWerkProblemId
    If ValueToTableField("##41", str, "Orders", "ProblemId") = 0 Then
        wrkDefault.CommitTrans  ' подтверждение транзакции
        werkBegin
    Else
ER1:    wrkDefault.Rollback    ' отммена транзакции
    End If
ElseIf I = -1 Then
    wrkDefault.Rollback
    msgZakazDeleted
End If

lbHide ' в т.ч. подсветка
End Sub

Private Sub lbProblem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbProblem_DblClick
ElseIf KeyCode = vbKeyEscape Then
'    wrkDefault.Rollback ' отммена транзакции
End If
End Sub

Sub lbHide()
lbStatus.Visible = False
lbObrazec.Visible = False
lbMaket.Visible = False
lbProblem.Visible = False
Grid.Enabled = True
Grid.SetFocus
On Error GoTo ER1 ' подсвечиваемая строка м.б. уже удалена
Grid.row = mousRow
gridIsLoad = True
Grid.col = mousCol
'Grid_EnterCell
Exit Sub
ER1:
gridIsLoad = True
End Sub

Function procReadyIs100() As Boolean
Dim str As String

    procReadyIs100 = False
    str = Grid.TextMatrix(mousRow, chProcVip)
    If Not IsNumeric(str) Then GoTo ERR100
    If str < 99.99 Then
ERR100:
        MsgBox "прежде заказ должен быть выполнен на 100%", , _
               "Недопустимый статус!"
        lbHide
        Exit Function
    End If
    procReadyIs100 = True

End Function

Private Sub lbStatus_DblClick()
Dim str As String, I As Integer


gEquipId = Grid.TextMatrix(mousRow, chEquipId)

#If onErrorOtlad Then
    On Error GoTo errMsg
    GoTo START
errMsg:
    MsgBox Error, , "Ошибка  " & Err & " в п\п lbStatus_DblClick" '
    End
START:
#End If

If noClick Then Exit Sub
str = lbStatus.Text
If str = "отложен" Then
    Grid.col = chProblem
    listBoxInGridCell lbProblem, Grid
    Exit Sub
ElseIf str = "готов" Then
    lbStatus.Visible = False
    If Not procReadyIs100() Then Exit Sub
    If Not predmetiIsClose("etap") Then '
        str = "текущего этапа "
        If QQ2(0) = 0 Then str = ""
        MsgBox "По этому заказу списаны(отпущены) не все предметы " & str & _
        "(для просмотра кликнете по колонке Изделия)!", , _
        "Недопустимый статус для Заказа № " & gNzak
'        Grid.SetFocus
    Else
        wrkDefault.BeginTrans
        ' Сначала меняем статус на Готов для этого оборудования
        I = ValueToTableField("W##41", "'" & str & "'", "OrdersEquip", "Stat", "byEquipId")
        
        Dim minId, maxId  'as variant
        sql = "select max(Stat) as maxId, min(Stat) as minId " _
        & " FROM OrdersEquip oe" _
        & " WHERE oe.numorder = " & gNzak
        
        'Debug.Print sql
        
        byErrSqlGetValues "##39.2", sql, maxId, minId
        
        If minId = maxId And minId = "готов" Then
            If ValueToTableField("##39.1", "4", "OrdersEquip", "StatusEquipId") <> 0 Then GoTo ER1
            If ValueToTableField("##39.3", "4", "Orders", "StatusId") <> 0 Then GoTo ER1
            If ValueToTableField("##39.4", "0", "Orders", "ProblemId") <> 0 Then GoTo ER1
            'раз все списано, отстегиваем текущ.этап, несмотря, что цех м. и снять гот-ть
            If Not newEtap("xEtapByIzdelia") Then GoTo ER1
            If Not newEtap("xEtapByNomenk") Then GoTo ER1
        End If
        wrkDefault.CommitTrans
        werkBegin
    End If
ElseIf str = "25%" Or str = "50%" Or str = "75%" Or str = "100%" Then
    lbStatus.Visible = False
    wrkDefault.BeginTrans
    If makeProcReady(str, Grid.TextMatrix(mousRow, chEquipId)) Then 'М в это время мог удалить заказ из цеха
        If ValueToTableField("##39", "1", "Orders", "StatusId") <> 0 Then GoTo ER1 ' "в работе"
        str = "в работе"
        GoTo AA
    End If
    GoTo ER2
Else '  пусто, "*" и "в работе"
    lbStatus.Visible = False
    wrkDefault.BeginTrans
    If makeProcReady("0%", Grid.TextMatrix(mousRow, chEquipId)) Then
        If ValueToTableField("##41", "'" & str & "'", "OrdersEquip", "Stat", "byEquipId") <> 0 Then GoTo ER1
        If ValueToTableField("##39", "1", "Orders", "StatusId") <> 0 Then GoTo ER1
AA:     If ValueToTableField("##39", "0", "Orders", "ProblemId") = 0 Then
            wrkDefault.CommitTrans
            werkBegin
        Else
ER1:        wrkDefault.Rollback
        End If
    Else
ER2:    wrkDefault.Rollback
        msgZakazDeleted
    End If
End If
lbHide ' в т.ч. подсветка
End Sub

Sub msgZakazDeleted(Optional str As String = "")
    If str = "" Then str = "заказ уже удален"
    MsgBox "Похоже этот " & str & " менеджером из цеха. Нажмите " & _
    "кнопку 'Обновить'.", , "Предупреждение"
End Sub


'$odbc14$
'Для образца, кот. Утвержден возвращает Null
Function makeProcReady(Stat As String, equipId As Integer, Optional obraz As String = "") As Variant
Dim S As Single, T As Single, N As Single, virabotka As Single, str As String
Dim StatO As String

makeProcReady = False
If Stat = "25%" Then
    S = 0.75 ' невыполнено
    GoTo AA
ElseIf Stat = "50%" Then
    S = 0.5
    GoTo AA
ElseIf Stat = "75%" Then
    S = 0.25
    GoTo AA
ElseIf Stat = "100%" Then
    S = 0
    GoTo AA
Else
    S = 1
AA:
 
  If obraz <> "" Then
    obraz = "o"
    ''??TODO
    sql = "SELECT oe.workTimeMO, oe.StatO " _
    & " FROM OrdersEquip oe " _
    & " WHERE oe.numOrder = " & gNzak & " AND equipId = " & equipId
    If Not byErrSqlGetValues("##386", sql, virabotka, StatO) Then Exit Function
    If S = 0 Then ' 100%
    Else
        virabotka = -virabotka
    End If
  Else
    sql = "SELECT oe.workTime, isnull(oe.Nevip, 1) as nevip " _
    & " FROM OrdersEquip oe " _
    & " WHERE oe.numOrder = " & gNzak & " AND equipId = " & equipId
    If Not byErrSqlGetValues("##421", sql, T, N) Then Exit Function
    
    virabotka = Round((N - S) * T, 2)
  End If


'гот-ть может изменится к примеру с 75% до 0%
    str = Format(curDate, "yy.mm.dd")
    
    sql = "call putWerkOrderReady(" & gNzak & ", '" & str & "', '" & obraz & "', " & virabotka & ", " & equipId & ", " & S & ")"
  
    myExecute "##374", sql
    
    If obraz = "o" Then '          это образец
        If StatO = "утвержден" Then
            makeProcReady = Null
            Exit Function
        End If
    Else 'obraz = ""
        gEquipId = equipId
        ValueToTableField "##41", "'в работе'", "OrdersEquip", "Stat", "byEquipId"
    End If
    
End If 'If stat
makeProcReady = True


End Function

Private Sub lbStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbStatus_DblClick
End Sub

Private Sub mnFind_Click()
Orders.Show
Orders.loadWithFiltr gNzak
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False

werkBegin
gridIsLoad = True
Grid.col = 1
isWerkOrders = True
trigger = True

End Sub

Function newEtap(table As String) As Boolean
newEtap = False
sql = "UPDATE " & table & " SET prevQuant = eQuant WHERE numOrder =" & gNzak
If myExecute("##193", sql, 0) > 0 Then Exit Function
newEtap = True
End Function

