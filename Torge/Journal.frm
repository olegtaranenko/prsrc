VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Journal 
   BackColor       =   &H8000000A&
   Caption         =   "Журнал хозяйственных операций"
   ClientHeight    =   6216
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   11796
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6216
   ScaleWidth      =   11796
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmExcel 
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   5700
      TabIndex        =   18
      Top             =   5760
      Width           =   1275
   End
   Begin VB.TextBox tbInform 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   200
      Locked          =   -1  'True
      MaxLength       =   150
      TabIndex        =   16
      Top             =   300
      Width           =   11355
   End
   Begin VB.CheckBox ckStartDate 
      Caption         =   " "
      Height          =   315
      Left            =   1140
      TabIndex        =   13
      Top             =   0
      Width           =   195
   End
   Begin VB.TextBox tbStartDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   0
      Width           =   795
   End
   Begin VB.CheckBox ckEndDate 
      Caption         =   " "
      Height          =   315
      Left            =   2700
      TabIndex        =   11
      Top             =   0
      Width           =   195
   End
   Begin VB.TextBox tbEndDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3180
      TabIndex        =   9
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmLoad 
      Caption         =   "Загрузить"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   2820
      Top             =   5640
   End
   Begin VB.ListBox lbGuids 
      Height          =   432
      ItemData        =   "Journal.frx":0000
      Left            =   2280
      List            =   "Journal.frx":000A
      TabIndex        =   7
      Top             =   1260
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   780
      TabIndex        =   6
      Text            =   "tbMobile"
      Top             =   1740
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   5760
      Width           =   1095
   End
   Begin VB.ListBox lbDebKreditor 
      Height          =   240
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox tbKurs 
      Height          =   285
      Left            =   6600
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   10860
      TabIndex        =   1
      Top             =   5760
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4995
      Left            =   180
      TabIndex        =   0
      Top             =   600
      Width           =   11415
      _ExtentX        =   20130
      _ExtentY        =   8805
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label laFiltr 
      Caption         =   "Записи, дающие Кредит = 23412"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   7560
      TabIndex        =   21
      Top             =   60
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label laSum 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4800
      TabIndex        =   20
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Сумма:"
      Height          =   195
      Left            =   4140
      TabIndex        =   19
      Top             =   60
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "Договор"
      Height          =   195
      Left            =   10860
      TabIndex        =   17
      Top             =   120
      Width           =   675
   End
   Begin VB.Label laPeriod 
      Caption         =   "Период с  "
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   60
      Width           =   795
   End
   Begin VB.Label laPo 
      Caption         =   "пос"
      Height          =   195
      Left            =   2340
      TabIndex        =   14
      Top             =   60
      Width           =   195
   End
   Begin VB.Label laKurs 
      Caption         =   "Курс:"
      Height          =   255
      Left            =   6180
      TabIndex        =   2
      Top             =   60
      Width           =   435
   End
   Begin VB.Menu mnReport 
      Caption         =   "Отчеты"
      Begin VB.Menu mnSOborot 
         Caption         =   "Оборот по счету"
      End
      Begin VB.Menu mnDohod 
         Caption         =   "Реализация"
      End
      Begin VB.Menu mnAreport 
         Caption         =   "Отчет А"
      End
   End
   Begin VB.Menu mnMyGuide 
      Caption         =   "Справочники"
      Begin VB.Menu mnGuide 
         Caption         =   "Справочник счетов"
      End
      Begin VB.Menu mnPurpose 
         Caption         =   "Справочник Операций"
      End
      Begin VB.Menu mnPurpDet 
         Caption         =   "Назначения"
      End
      Begin VB.Menu mnKreditors 
         Caption         =   "Дебиторы \ Кредиторы"
      End
      Begin VB.Menu mnShiz 
         Caption         =   "Шифры затрат"
      End
   End
   Begin VB.Menu mnNasroy 
      Caption         =   "Настройка"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Journal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isLoad As Boolean
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim quantity As Long
Dim mousCol As Long, mousRow As Long

Const cNewLbLine = "-new-"

Const jnDate = 1
Const jnM = 2
Const jnRate = 3
Const jnVal = 4
Const jnDebit = 5
Const jnSubDebit = 6
Const jnKredit = 7
Const jnSubKredit = 8
Const jnFirm = 9 ' временно
Const jnDebKreditor = 10
Const jnOrdersNum = 11
Const jnNote = 12
Const jnPurpose = 13
Const jnDetail = 14
Const jnVenture = 15
Const jnShiz = 16
Const jnId = 17


Private Sub mnAreport_Click()
    Set ReportA.Caller = Me
    ReportA.Show vbModal
End Sub

Private Sub mnDohod_Click()
    Pribil.Show
End Sub

Private Sub mnGuide_Click()
    jGuideSchets.Show vbModal
End Sub

Private Sub mnKreditors_Click()
    GuideDebKreditor.Show vbModal

End Sub

Private Sub mnNasroy_Click()
Nastroy.Show vbModal
End Sub

Private Sub mnPurpDet_Click()
jGuidePurpDet.Regim = ""
jGuidePurpDet.Show vbModal
End Sub

Private Sub mnPurpose_Click()
jGuidePurpose.Show vbModal
End Sub

Private Sub mnShiz_Click()
GuideShiz.Show vbModal
End Sub

Private Sub mnSOborot_Click()
jKassaReport.Show
End Sub

Private Sub ckEndDate_Click()
If ckEndDate.value = 1 Then
    tbEndDate.Enabled = True
Else
    tbEndDate.Enabled = False
End If
cmLoad.Caption = "Загрузить"

End Sub

Private Sub ckStartDate_Click()
If ckStartDate.value = 1 Then
    tbStartDate.Enabled = True
Else
    tbStartDate.Enabled = False
End If
cmLoad.Caption = "Загрузить"
End Sub

Private Sub cmAdd_Click()
Dim i As Integer
Dim str As String
Dim rate As Variant

cmAdd.Enabled = False 'нельзя жать чаще 1с, т.к. дата это ключ
'frmMode = "sourceAdd"

On Error GoTo adderr
sql = "SELECT Kurs FROM System"
If byErrSqlGetValues("##321", sql, str) Then
    rate = Abs(CDbl(str))
End If

Set tbDocs = myOpenRecordSet("##324", "yBook", dbOpenTable) 'dbOpenForwardOnly)
If tbDocs Is Nothing Then Exit Sub
i = InStr(cmAdd.Caption, "+")
If i > 0 Then
    If Not IsDate(Grid.TextMatrix(mousRow, jnDate)) Then
        MsgBox "Образец не определен"
        GoTo AA
    End If
    tbDocs.index = "Key"
    tbDocs.Seek "=", Grid.TextMatrix(mousRow, jnDate)
    If tbDocs.NoMatch Then
        MsgBox "Образец не найден в базе", , ""
AA:     cmAdd.Caption = left$(cmAdd.Caption, i - 2)
        tbDocs.Close
        Exit Sub
    End If
    purposeId = tbDocs!purposeId
'    detailId = tbDocs!detailId
    KredDebitor = tbDocs!KredDebitor
End If



tbDocs.AddNew
tmpDate = Format(Now(), "dd/mm/yy hh:nn:ss")
tbDocs!xDate = tmpDate

tbDocs!m = AUTO.cbM.Text


If i > 0 Then
    tbDocs!debit = Grid.TextMatrix(mousRow, jnDebit)
    str = Grid.TextMatrix(mousRow, jnSubDebit)
'    If Not IsNumeric(str) Then str = 0
    tbDocs!subDebit = str
    tbDocs!kredit = Grid.TextMatrix(mousRow, jnKredit)
    str = Grid.TextMatrix(mousRow, jnSubKredit)
'    If Not IsNumeric(str) Then str = 0
    tbDocs!subKredit = str
    tbDocs!ordersNum = Grid.TextMatrix(mousRow, jnOrdersNum)
    tbDocs!note = Grid.TextMatrix(mousRow, jnNote)
    tbDocs!purposeId = purposeId
'    tbDocs!detailId = detailId
    tbDocs!KredDebitor = KredDebitor
End If
If Not IsNull(rate) Then _
    tbDocs!rate = rate

tbDocs.Update
tbDocs.Close
GoTo aaa

adderr:
If errorCodAndMsg(cErr) Then
    Exit Sub
End If

aaa:
If quantity > 0 Then Grid.AddItem ("")
quantity = quantity + 1
noClick = True
Grid.row = Grid.Rows - 1
Grid.col = jnRate
noClick = False

Grid.TextMatrix(Grid.row, jnDate) = Format(tmpDate, "dd/mm/yy hh:nn:ss")
Grid.TextMatrix(Grid.row, jnM) = AUTO.cbM.Text
If Not IsNull(rate) Then _
    Grid.TextMatrix(Grid.row, jnRate) = rate
If i > 0 Then
    Grid.TextMatrix(Grid.row, jnDebit) = Grid.TextMatrix(mousRow, jnDebit)
    Grid.TextMatrix(Grid.row, jnSubDebit) = Grid.TextMatrix(mousRow, jnSubDebit)
    Grid.TextMatrix(Grid.row, jnKredit) = Grid.TextMatrix(mousRow, jnKredit)
    Grid.TextMatrix(Grid.row, jnSubKredit) = Grid.TextMatrix(mousRow, jnSubKredit)
    Grid.TextMatrix(Grid.row, jnDebKreditor) = Grid.TextMatrix(mousRow, jnDebKreditor)
    Grid.TextMatrix(Grid.row, jnOrdersNum) = Grid.TextMatrix(mousRow, jnOrdersNum)
    Grid.TextMatrix(Grid.row, jnNote) = Grid.TextMatrix(mousRow, jnNote)
    Grid.TextMatrix(Grid.row, jnPurpose) = Grid.TextMatrix(mousRow, jnPurpose)
    Grid.TextMatrix(Grid.row, jnDetail) = Grid.TextMatrix(mousRow, jnDetail)
    Grid.TextMatrix(Grid.row, jnVenture) = Grid.TextMatrix(mousRow, jnVenture)
    cmAdd.Caption = left$(cmAdd.Caption, i - 2)
End If
mousRow = Grid.Rows - 1
mousCol = jnRate
rowViem quantity, Grid
On Error Resume Next
Grid.SetFocus
Grid_EnterCell
'textBoxInGridCell tbMobile, Grid
Timer1.Interval = 1100
Timer1.Enabled = True ' разблокировка cmAdd
End Sub


Private Sub cmDel_Click()
If quantity <= 0 Then Exit Sub
If MsgBox("После нажатия <Да> текущая запись будет удалена.", _
vbDefaultButton2 Or vbYesNo, "Удалить, Вы уверены?") = vbNo Then Exit Sub

If valueToBookField("##358", 0, "delete") Then
    quantity = quantity - 1
    If quantity > 0 Then
        Grid.RemoveItem mousRow
    Else
        clearGridRow Grid, 1
    End If
End If
On Error Resume Next
Grid.SetFocus
Grid_EnterCell

End Sub

Private Sub cmExcel_Click()
GridToExcel Grid, "Журнал хоз.операций с " & tbStartDate.Text & _
"  по  " & tbEndDate.Text
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmLoad_Click()
laFiltr.Visible = False
loadBook
cmLoad.Caption = "Обновить"
Grid_EnterCell
On Error Resume Next
Grid.SetFocus

End Sub

Sub loadLbFromDebKreditor()
Dim i As Long

Set table = myOpenRecordSet("##353", "select * from yDebKreditor order by Name", dbOpenTable)
If table Is Nothing Then myBase.Close: End
lbDebKreditor.Clear
'Table.Index = "Name"
While Not table.EOF
    lbDebKreditor.AddItem table!name
    table.MoveNext
Wend
table.Close
'i = 195 * lbDebKreditor.ListCount + 100
'If i > Grid.Height - 1000 Then i = Grid.Height - 1000
'If i > 2000 Then i = 2000
'lbDebKreditor.Height = i


End Sub


Sub loadLbFromSchets(lb As ListBox, Optional selNumStr As String = "-1", _
Optional selSubNumStr As String = "-1")
Dim str As String, table As Recordset, heig As Integer
Dim selNum As Integer, selSubNum As Integer

selNum = CInt(selNumStr)
selSubNum = CInt(selSubNumStr)


If selNum = -2 Then ' без суб.счетов
    sql = "SELECT number From yGuideSchets GROUP BY number ORDER BY number;"
    Set table = myOpenRecordSet("##325", sql, dbOpenForwardOnly)
'    If Table Is Nothing Then Exit Sub
Else
'    Set Table = myOpenRecordSet("##325", "yGuideSchets", dbOpenTable)
    sql = "SELECT * FROM yGuideSchets order by number + subnumber;"
    Set table = myOpenRecordSet("##325", sql, dbOpenForwardOnly)
'    If Table Is Nothing Then Exit Sub
'    Table.Index = "Key"
End If
lb.Clear
While Not table.EOF
  If table!number < 255 Then
    'str = Format(Table!number, "00")
    str = table!number
    If selNum <> -2 Then
      If table!subNumber > 0 Then str = str & " " & table!subNumber
    End If
    lb.AddItem str
    If selNum <> -2 Then
        If selNum = table!number And selSubNum = table!subNumber _
            Then lb.ListIndex = lb.ListCount - 1
    End If
  End If
  table.MoveNext
Wend
table.Close

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, row As Long, str As String
Static prevRow As Long, value As String

If KeyCode = vbKeyMenu Then
    If InStr(cmAdd.Caption, "+") = 0 Then cmAdd.Caption = cmAdd.Caption & " +"
ElseIf KeyCode = vbKeyF7 Then
    If quantity = 0 Then
        MsgBox "Сначала произведите загрузку. При необходимости установите " & _
        "даты загрузки.", , ""
        Exit Sub
    End If
    If Shift = 1 And prevRow > 0 Then
'        value = InputBox("Введите образец для поиска или фрагмент.", _
'        "Продолжение поиска в поле '" & Grid.TextMatrix(0, mousCol) & "'", value)
        row = prevRow
        str = "Больше"
    Else
        value = InputBox("Введите образец для поиска или фрагмент. " & vbCrLf & _
        vbCrLf & "Далее для проджения поиска этого образца со следующей " & _
        "позииции Вы можете нажимать <Shift><F7>.", "Поиск в поле '" & _
        Grid.TextMatrix(0, mousCol) & "'", value)
        row = -1
        str = "Среди загруженных"
    End If
    If value = "" Then Exit Sub
    prevRow = findExValInCol(Grid, value, CInt(mousCol), row) + 1
    If prevRow < 1 Then _
        MsgBox str & " образец '" & value & "' не найден!", , ""
    
 End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If KeyCode = vbKeyMenu Then
    i = InStr(cmAdd.Caption, "+")
    If i > 0 Then cmAdd.Caption = left$(cmAdd.Caption, i - 2)
End If


End Sub

Private Sub Form_Load()
Dim s As Single

oldHeight = Me.Height
oldWidth = Me.Width

isLoad = True
ckStartDate.value = 1

If dostup = "a" Then mnNasroy.Visible = True

Me.Caption = Me.Caption & "      " & mainTitle
tbStartDate.Text = "01." & Format(CurDate, "mm/yy")
'tbStartDate.Text = Format(DateAdd("d", 0, begDate), "dd/mm/yy")
tbEndDate.Text = Format(CurDate, "dd/mm/yy")
If otlad = "otlaD" Then
    ckStartDate.value = 1
    Me.BackColor = otladColor
End If

sql = "SELECT Kurs FROM System;"
If byErrSqlGetValues("##321", sql, s) Then
    tbKurs.Text = Abs(s)
    'If s > 0.01 Then tbKurs.Text = s ' с "-" - это вчерашний курс
End If

Grid.FormatString = "|<Дата|М|Курс|Валюта|Дб|Сс|Кр|Сс|<Заказчик(временная)" & _
"|<Кредитор\Дебитор|<Договор|<Примечание|<Назначение|<Уточнение|<Предприятие|<Шифр затрат|id_xoz"
Grid.colWidth(0) = 0
Grid.colWidth(jnDate) = 780
Grid.colWidth(jnDebKreditor) = 2580
Grid.colWidth(jnOrdersNum) = 1200
Grid.colWidth(jnFirm) = 1395
Grid.colWidth(jnDetail) = 1500
Grid.colWidth(jnVenture) = 1600
Grid.colWidth(jnId) = 0
'jnNote


loadLbFromDebKreditor
quantity = 0
End Sub

Sub loadBook()
 Dim i As Integer, str As String

Grid.Visible = False
quantity = 0
clearGrid Grid
    
If laFiltr.Visible = True Then
    strWhere = "WHERE (" & jKassaReport.filtrWhere & ")"
Else
    strWhere = getWhereByDateBoxes(Me, "yBook.xDate", CDate("01.01.2003"))
    If strWhere = "error" Then GoTo EN1
    If strWhere <> "" Then strWhere = "WHERE ((" & strWhere & "))"
End If
 
sql = "SELECT yBook.xDate , yBook.UEsumm , yBook.Debit , yBook.subDebit , yBook.Kredit , yBook.subKredit , yBook.KredDebitor " & _
", isnull(yBook.ordersNum, '') as ordersNum , isnull(yBook.Note, '') Note , isnull(yGuidePurpose.pDescript, '') pDescript, isnull(yBook.descript, '') descript, yBook.M , yBook.firm , isnull(v.ventureName, '') as ventureName" & _
", s.nm as shiz_text, ybook.rate, ybook.id " & _
" from ybook " & _
" left join yGuidePurpose " & _
"   ON (yGuidePurpose.pId = yBook.purposeId)  " & _
"   AND (yGuidePurpose.subKredit = yBook.subKredit)  " & _
"   AND (yGuidePurpose.Kredit = yBook.Kredit)  " & _
"   AND (yGuidePurpose.subDebit = yBook.subDebit)  " & _
"   AND (yGuidePurpose.Debit = yBook.Debit)  " & _
" left join GuideVenture v on v.ventureId = yBook.ventureId" & _
" left join Shiz s on s.id = yBook.id_shiz" & _
" " & strWhere & _
" ORDER BY yBook.xDate; "


Set tbDocs = myOpenRecordSet("##323", sql, dbOpenForwardOnly)
If tbDocs Is Nothing Then GoTo EN1
Me.MousePointer = flexHourglass

If Not tbDocs.BOF Then
 While Not tbDocs.EOF
    quantity = quantity + 1
    i = tbDocs!KredDebitor
    Grid.TextMatrix(quantity, 0) = i
    Grid.TextMatrix(quantity, jnDate) = Format(tbDocs!xDate, "dd/mm/yy hh:nn:ss")
    If Not IsNull(tbDocs!rate) Then
        Grid.TextMatrix(quantity, jnRate) = Round(tbDocs!rate, 2)
    End If
    Grid.TextMatrix(quantity, jnVal) = Round(tbDocs!uesumm, 2)
    Grid.TextMatrix(quantity, jnDebit) = schType(tbDocs!debit, 255)
    Grid.TextMatrix(quantity, jnSubDebit) = schType(tbDocs!subDebit)
    Grid.TextMatrix(quantity, jnKredit) = schType(tbDocs!kredit, 255)
    Grid.TextMatrix(quantity, jnSubKredit) = schType(tbDocs!subKredit)
    If i > 0 Then
        sql = "SELECT Name From GuideFirms WHERE (((FirmId)=" & i & "));"
        GoTo AA
    ElseIf i < 0 Then
        sql = "SELECT Name From yDebKreditor WHERE (((id)=" & i & "));"
'"W##.." - Если фирму или дебкредитора удалили, то поле б пустым
AA:     If byErrSqlGetValues("W##428", sql, str) Then _
            Grid.TextMatrix(quantity, jnDebKreditor) = str
    End If

    If Not IsNull(tbDocs!firm) Then _
        Grid.TextMatrix(quantity, jnFirm) = tbDocs!firm
    Grid.TextMatrix(quantity, jnOrdersNum) = tbDocs!ordersNum
    Grid.TextMatrix(quantity, jnNote) = tbDocs!note
    Grid.TextMatrix(quantity, jnPurpose) = tbDocs!pDescript
    Grid.TextMatrix(quantity, jnDetail) = tbDocs!descript
    Grid.TextMatrix(quantity, jnVenture) = tbDocs!ventureName
    Grid.TextMatrix(quantity, jnM) = tbDocs!m
    If Not IsNull(tbDocs!shiz_text) Then _
        Grid.TextMatrix(quantity, jnShiz) = tbDocs!shiz_text
    If Not IsNull(tbDocs!id) Then _
        Grid.TextMatrix(quantity, jnId) = tbDocs!id
    Grid.AddItem ""
    tbDocs.MoveNext
 Wend
End If
tbDocs.Close
Grid.Visible = True
rowViem quantity, Grid

'laQuant.Caption = quantity
If quantity > 0 Then
    Grid.RemoveItem quantity + 1
    Grid.row = quantity
    Grid.col = 1
    tbInform.Locked = False
'    Grid.SetFocus
End If
EN1:
Grid.Visible = True
Me.MousePointer = flexDefault

End Sub

Function schType(number As Variant, Optional pass As Integer = 0) As String
schType = number
If number = pass Then
    schType = ""
Else
    If IsNumeric(number) Then
'        schType = Format(number, "00")
'               schType = number
    End If
    
End If
End Function

Sub setSchetsAndPurpose(purpose As String) ', detail As String)
Grid.TextMatrix(mousRow, jnDebit) = schType(debit, 255)
Grid.TextMatrix(mousRow, jnSubDebit) = schType(subDebit)
Grid.TextMatrix(mousRow, jnKredit) = schType(kredit, 255)
Grid.TextMatrix(mousRow, jnSubKredit) = schType(subKredit)
Grid.TextMatrix(mousRow, jnPurpose) = purpose
'Grid.TextMatrix(mousRow, jnDetail) = detail
End Sub


Function valueToBookField(myErrCod As String, value As String, field As String) As Boolean
Dim pId As Integer, dId As Integer, str As String

    valueToBookField = False
    
    wrkDefault.BeginTrans 'lock01
    sql = "update system set resursLock = resursLock" 'lock02
    myBase.Execute (sql) 'lock03
    
    
    strWhere = " WHERE id = " & Grid.TextMatrix(mousRow, jnId)
    
    If field = "delete" Then
        sql = "DELETE FROM yBook " & strWhere
        GoTo AA
    ElseIf field = "schets" Then
        pId = getPurposeIdByDescript(jGuidePurpose.lbPurpose.Text)
        
        sql = "UPDATE yBook set debit = '" & debit & "', subDebit = '" & subDebit & _
        "', kredit = '" & kredit & "', subKredit = '" & subKredit & "', purposeId = " & _
        pId & strWhere
        GoTo AA
    Else
        sql = "UPDATE yBook set " & field & " = " & value & strWhere
AA:
        If myExecute(myErrCod, sql) <> 0 Then Exit Function
        'Debug.Print sql
        
    End If
    
    wrkDefault.CommitTrans
    
EN1:
    valueToBookField = True
End Function


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

laKurs.left = laKurs.left + w
tbKurs.left = tbKurs.left + w
cmLoad.Top = cmLoad.Top + h
cmAdd.Top = cmAdd.Top + h
cmDel.Top = cmDel.Top + h
cmExit.Top = cmExit.Top + h
cmExcel.Top = cmExcel.Top + h
'cmExit.Left = cmExit.Left + w

'Grid_EnterCell
'Grid.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
isLoad = False
If beChange Then KursUpdate
If jKassaReport.isLoad Then Unload jKassaReport
End Sub

Private Sub Grid_Click()
If noClick Then Exit Sub
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If quantity = 0 Then Exit Sub
If mousRow = 0 Then
    Grid.CellBackColor = Grid.BackColor
    
    If mousCol > jnM And mousCol <= jnSubKredit Or mousCol = jnOrdersNum Then
        SortCol Grid, mousCol, "numeric"
    ElseIf mousCol = jnDate Then
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


Private Sub Grid_DblClick()
If Grid.CellBackColor <> &H88FF88 Then Exit Sub

If mousCol = jnSubKredit Or mousCol = jnSubDebit _
Or mousCol = jnDebit Or mousCol = jnKredit Or mousCol = jnPurpose Then
    jGuidePurpose.Regim = "select"
    getSchetsFromGrid Grid, jnDebit
    jGuidePurpose.purpose = Grid.TextMatrix(mousRow, jnPurpose)
'    jGuidePurpose.detail = Grid.TextMatrix(mousRow, jnDetail)
    jGuidePurpose.Show vbModal
ElseIf mousCol = jnDate Then
    textBoxInGridCell tbMobile, Grid, left$(Grid.TextMatrix(mousRow, mousCol), 8)
ElseIf mousCol = jnDebKreditor Then
    listBoxInGridCell lbGuids, Grid, "select" 'lbDebKreditor
ElseIf mousCol = jnRate Then
    If Not IsNumeric(tbKurs.Text) Then GoTo EE
    If CSng(tbKurs.Text) < 0.01 Then
EE:     tbKurs.SelStart = 0: tbKurs.SelLength = Len(tbKurs.Text)
        MsgBox "Введите курс!", , "Предупреждение"
        tbKurs.SetFocus
        Exit Sub
    End If
    textBoxInGridCell tbMobile, Grid
Else
    textBoxInGridCell tbMobile, Grid
End If
End Sub

Private Sub Grid_EnterCell()
If noClick Then Exit Sub
If quantity = 0 Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col
tbInform.Text = Grid.TextMatrix(mousRow, jnOrdersNum)
If Grid.TextMatrix(mousRow, jnVenture) <> "" Then
    cmDel.Enabled = False
Else
    cmDel.Enabled = True
End If
If (mousCol <> jnM And mousCol <> jnVenture _
  And (Grid.TextMatrix(mousRow, jnVenture) = "")) _
  Or (mousCol = jnRate And dostup = "a") _
Then  ' чтобы м.б. копировать из jnFirm
'If mousCol > 2 Then  ' чтобы м.б. копировать из jnFirm
   Grid.CellBackColor = &H88FF88
Else
   Grid.CellBackColor = vbYellow
End If

'If mousCol = nkNomer Then
'   tbMobile.MaxLength = 20
'Else
'   tbMobile.MaxLength = 10
'End If

End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Grid_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Sub lbHide()
tbMobile.Visible = False
lbDebKreditor.Visible = False
lbGuids.Visible = False
Grid.Enabled = True
On Error Resume Next
Grid.SetFocus
Grid_EnterCell

End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Grid_EnterCell
End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Grid.MouseRow = 0 And Shift = 2 Then
        MsgBox "ColWidth = " & Grid.colWidth(Grid.MouseCol)
'ElseIf quantity > 0 And Grid.row <> Grid.RowSel Then
ElseIf mousCol = jnVal Then
    laSum.Caption = Round(sumInGridCol(Grid, mousCol), 2)
End If
End Sub

Private Sub lbDebKreditor_DblClick()
Dim id As String

id = getIdFromTableByLb("yDebKreditor", lbDebKreditor.Text)
If IsNumeric(id) Then
    If valueToBookField("##354", id, "KredDebitor") Then
        Grid.Text = lbDebKreditor.Text
        If jKassaReport.isLoad Then jKassaReport.laInform.Visible = True
    End If
End If
lbHide

End Sub

Private Sub lbDebKreditor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbDebKreditor_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub lbGuids_DblClick()
lbHide
If lbGuids.ListIndex = 0 Then
    listBoxInGridCell lbDebKreditor, Grid
Else
    FindFirm.Regim = "edit"
    FindFirm.cmSelect.Visible = True
    FindFirm.tb.Text = Grid.TextMatrix(mousRow, jnDebKreditor)
    FindFirm.Show vbModal
End If


End Sub

Private Sub lbGuids_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbGuids_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub tbEndDate_Change()
cmLoad.Caption = "Загрузить"

End Sub

Private Sub tbInform_GotFocus()
    tbInform.SelStart = Len(tbInform.Text)
End Sub

Private Sub tbInform_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If valueToBookField("##326", tbInform.Text, "ordersNum") Then
        Grid.TextMatrix(mousRow, jnOrdersNum) = tbInform.Text
        Grid.col = jnOrdersNum
        On Error Resume Next
        Grid.SetFocus
    End If
End If

End Sub

Private Sub tbKurs_Change()
beChange = True
End Sub

Private Sub tbKurs_GotFocus()
beChange = False
End Sub
Sub addRowToGrid(sum As String, note As String)

If quantity > 0 Then Grid.AddItem ("")
quantity = quantity + 1
noClick = True
Grid.row = Grid.Rows - 1
Grid.col = jnRate
noClick = False
Grid.TextMatrix(Grid.row, jnDate) = tmpStr
Grid.TextMatrix(Grid.row, jnVal) = sum
Grid.TextMatrix(Grid.row, jnM) = AUTO.cbM.Text
Grid.TextMatrix(Grid.row, jnDebit) = schType(debit)
Grid.TextMatrix(Grid.row, jnSubDebit) = schType(subDebit)
Grid.TextMatrix(Grid.row, jnKredit) = schType(kredit)
Grid.TextMatrix(Grid.row, jnSubKredit) = schType(subKredit)
Grid.TextMatrix(Grid.row, jnNote) = note 'tbDocs!note


sql = "SELECT pDescript from yGuidePurpose WHERE (((Debit)='" & debit & _
"') AND ((subDebit)='" & subDebit & "') AND ((Kredit)='" & kredit & _
"') AND ((subKredit)='" & subKredit & "') AND ((pId)=" & purposeId & "));"

If byErrSqlGetValues("##377", sql, tmpStr) Then _
    Grid.TextMatrix(Grid.row, jnPurpose) = tmpStr
    
'sql = "SELECT descript FROM yGuideDetail WHERE (((Debit)=" & debit & _
'") AND ((subDebit)=" & subDebit & ") AND ((Kredit)=" & kredit & _
'") AND ((subKredit)=" & subKredit & ") AND ((purposeId)=" & purposeId & _
'") AND ((id)=" & detailId & "));"
''MsgBox sql
'If byErrSqlGetValues("##378", sql, tmpStr) Then _
'    Grid.TextMatrix(Grid.row, jnDetail) = tmpStr

'sql = "select ventureName from guideVenture where ventureid = " & ventureId & ";"
Grid_EnterCell
End Sub

Function getFreeDate(dayDate As String)
Dim sek As Long, str As String, l As Long


getFreeDate = Null

sql = "SELECT xDate from yBook Where (((xDate) Like '" & dayDate & "%')) ORDER BY xDate;"
'MsgBox sql
Set table = myOpenRecordSet("##374", sql, dbOpenForwardOnly)
If table Is Nothing Then Exit Function
sek = 0
While Not table.EOF
'  If Table!xDate <> DateAdd("s", sek, dayDate) Then
  If DateDiff("s", dayDate, table!xDate) <> sek Then GoTo EN1
  sek = sek + 1
  table.MoveNext
Wend

EN1:

If sek < 86400 Then 'число секунд в сутках
    getFreeDate = DateAdd("s", sek, dayDate)
End If
table.Close


End Function

Function getIdFromTableByLb(table As String, lbText As String) As String
Dim id As Integer

getIdFromTableByLb = ""

sql = "SELECT id From " & table & "  WHERE (((Name)='" & lbText & "'));"
'MsgBox sql
If Not byErrSqlGetValues("##330", sql, id) Then Exit Function

getIdFromTableByLb = id
End Function

'Sub getSchetsFromGrid()
Sub getSchetsFromGrid(Grid As MSFlexGrid, begCol As Integer)
Dim str As String

str = Grid.TextMatrix(Grid.row, begCol)
debit = 255
If str <> "" Then debit = str

str = Grid.TextMatrix(Grid.row, begCol + 1)
subDebit = "00"
If str <> "" Then subDebit = str


str = Grid.TextMatrix(Grid.row, begCol + 2)
kredit = 255
If str <> "" Then kredit = str

str = Grid.TextMatrix(Grid.row, begCol + 3)
subKredit = "00"
If str <> "" Then subKredit = str

End Sub

Sub KursUpdate()
    If Not isNumericTbox(tbKurs, 0.01) Then Exit Sub
    If beChange Then
        sql = "UPDATE System SET System.Kurs = " & tbKurs.Text & ";"
        myExecute "##322", sql
        beChange = False
    End If
End Sub

Private Sub tbKurs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    KursUpdate
    On Error Resume Next
    Grid.SetFocus
End If
End Sub

Private Sub tbKurs_LostFocus()
If beChange Then KursUpdate
End Sub

Private Sub tbMobile_DblClick()
lbHide
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String

If KeyCode = vbKeyReturn Then
  If mousCol = jnDate Then
    
     If Not isDateTbox(tbMobile) Then Exit Sub
'     str = getFreeDate(Format(tmpDate, "dd.mm.yyyy"))
     tmpDate = getFreeDate(Format(tmpDate, "yyyy-mm-dd"))
     If IsNull(tmpDate) Then GoTo EN1  '
     str = "'" & Format(tmpDate, "yyyy-mm-dd hh:nn:ss") & "'"
     If valueToBookField("##375", str, "xDate") Then _
        Grid.TextMatrix(mousRow, mousCol) = Format(tmpDate, "dd.mm.yy hh:nn:ss")
     GoTo EN1
  ElseIf mousCol = jnRate Then
    If Not isNumericTbox(tbMobile, 0.1) Then Exit Sub
    
    Grid.TextMatrix(mousRow, jnRate) = tbMobile.Text
    If Not valueToBookField("##375.0", tbMobile.Text, "rate") Then GoTo EN1
    If Grid.TextMatrix(mousRow, jnVenture) <> "" Then
        sql = "select retrieve_xoz_rub ( " & Grid.TextMatrix(mousRow, jnId) & ")"
        If Not byErrSqlGetValues("##330", sql, str) Then Exit Sub
        str = Round(CSng(str) / CSng(tbMobile.Text), 2)
        If Not valueToBookField("##375.1", str, "UEsumm") Then GoTo EN1
        Grid.TextMatrix(mousRow, jnVal) = str
    End If
  ElseIf mousCol = jnVal Then
    If Not isNumericTbox(tbMobile, 0.1) Then Exit Sub
    If Not valueToBookField("##326", tbMobile.Text, "UEsumm") Then GoTo EN1
  ElseIf mousCol = jnOrdersNum Then
    If Not valueToBookField("##326", "'" & tbMobile.Text & "'", "ordersNum") Then GoTo EN1
  ElseIf mousCol = jnNote Then
    If Not valueToBookField("##326", "'" & tbMobile.Text & "'", "Note") Then GoTo EN1
  ElseIf mousCol = jnDetail Then
    If Not valueToBookField("##326", "'" & tbMobile.Text & "'", "descript") Then GoTo EN1
  End If
  Grid.TextMatrix(mousRow, mousCol) = tbMobile.Text
  If jKassaReport.isLoad Then jKassaReport.laInform.Visible = True
EN1: lbHide
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub tbStartDate_Change()
cmLoad.Caption = "Загрузить"

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
cmAdd.Enabled = True
End Sub
