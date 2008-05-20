VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form jKassaReport 
   BackColor       =   &H8000000A&
   Caption         =   "Обороты по счетам"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   11250
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ckVenture 
      Caption         =   "Предприят."
      Height          =   315
      Left            =   9900
      TabIndex        =   27
      Top             =   3000
      Width           =   1275
   End
   Begin VB.TextBox tbMidDebit 
      Height          =   285
      Left            =   4860
      TabIndex        =   24
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox tbMidKredit 
      Height          =   285
      Left            =   4860
      TabIndex        =   23
      Top             =   660
      Width           =   1095
   End
   Begin VB.CheckBox ckDate 
      Caption         =   "Дате"
      Height          =   195
      Left            =   9900
      TabIndex        =   21
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CheckBox ckDKreditor 
      Caption         =   "Д\Кредитору"
      Height          =   195
      Left            =   9900
      TabIndex        =   20
      Top             =   3660
      Width           =   1335
   End
   Begin VB.CheckBox ckNumOrder 
      Caption         =   "№ договора"
      Height          =   195
      Left            =   9900
      TabIndex        =   19
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox ckPurpose 
      Caption         =   "Назначению"
      Height          =   195
      Left            =   9900
      TabIndex        =   18
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox tbEndKredit 
      Height          =   285
      Left            =   8580
      TabIndex        =   17
      Top             =   660
      Width           =   1155
   End
   Begin VB.TextBox tbEndDebet 
      Height          =   285
      Left            =   8580
      TabIndex        =   16
      Top             =   360
      Width           =   1155
   End
   Begin VB.TextBox tbBegKredit 
      Height          =   285
      Left            =   1500
      TabIndex        =   13
      Top             =   660
      Width           =   1095
   End
   Begin VB.TextBox tbBegDebet 
      Height          =   285
      Left            =   1500
      TabIndex        =   12
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   7440
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3735
      Left            =   60
      TabIndex        =   7
      Top             =   960
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   6588
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.ListBox lbSchets 
      Height          =   255
      Left            =   10200
      MultiSelect     =   2  'Расширенно
      TabIndex        =   6
      Top             =   240
      Width           =   795
   End
   Begin VB.CheckBox ckBySub 
      Caption         =   "С   учетом субсчетов"
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   0
      Value           =   1  'Отмечено
      Width           =   1935
   End
   Begin VB.TextBox tbEndDate 
      Height          =   285
      Left            =   1980
      MaxLength       =   8
      TabIndex        =   2
      Top             =   0
      Width           =   795
   End
   Begin VB.TextBox tbStartDate 
      Height          =   285
      Left            =   900
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "01.01.03"
      Top             =   0
      Width           =   795
   End
   Begin VB.CommandButton cmLoad 
      Caption         =   "Загрузить"
      Height          =   315
      Left            =   10020
      TabIndex        =   0
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label laInform 
      Caption         =   "Нажмите  <Обновить> !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   3360
      TabIndex        =   28
      Top             =   60
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Дебет оборот:"
      Height          =   195
      Left            =   3660
      TabIndex        =   26
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label Label6 
      Caption         =   "Кредит оборот:"
      Height          =   195
      Left            =   3660
      TabIndex        =   25
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label laGroup 
      Caption         =   "С разбивкой по:"
      Height          =   255
      Left            =   9900
      TabIndex        =   22
      Top             =   2460
      Width           =   1275
   End
   Begin VB.Label Label5 
      Caption         =   "Кредит на конец:"
      Height          =   195
      Left            =   7200
      TabIndex        =   15
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Дебет на конец:"
      Height          =   195
      Left            =   7200
      TabIndex        =   14
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Кредит на начало:"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "Дебет на начало:"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   420
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "По счетам:"
      Height          =   195
      Left            =   10200
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.Label laPo 
      Caption         =   "пос"
      Height          =   195
      Left            =   1740
      TabIndex        =   4
      Top             =   45
      Width           =   195
   End
   Begin VB.Label laPeriod 
      Caption         =   "Период  с"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   45
      Width           =   795
   End
End
Attribute VB_Name = "jKassaReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim reportDateWhere As String
Public filtrWhere As String
Public isLoad As Boolean
Dim mousCol As Long, mousRow As Long
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim quantity As Long, debitSql As String, kreditSql As String
Const byPurpose = 1
Const byVenture = 2
Const byOrdersNum = 3
Const byDebKreditor = 4
Const byDate = 5
Const byDebit = 6
Const byKredit = 7
Const widthPurpose = 2000
Const widthVenture = 1800
Const widthNumOrder = 1100
Const widthDKreditor = 1800
Const widthDate = 735

Private Sub ckBySub_Click()
cmLoad.Caption = "Загрузить"

If ckBySub.value = 0 Then
    Journal.loadLbFromSchets lbSchets, -2  ' без субсчетов
Else
    Journal.loadLbFromSchets lbSchets
End If
lbResize
End Sub

Private Sub ckDate_Click()
cmLoad.Caption = "Загрузить"
'If ckDate.value = 0 Then
'    Grid.ColWidth(byDate) = 0
'Else
'End If

End Sub

Private Sub ckDKreditor_Click()
cmLoad.Caption = "Загрузить"
'If ckDKreditor.value = 0 Then
'Else

End Sub

Private Sub ckNumOrder_Click()
cmLoad.Caption = "Загрузить"
End Sub

Private Sub ckPurpose_Click()
cmLoad.Caption = "Загрузить"
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmLoad_Click()
Dim dSql As String, kSql As String, dateWhere As String, dateWhereBef As String
Dim fields As String, from As String, xgroup  As String, str As String
Dim sum As Double, i As Integer, d As Double, k As Double
Dim OrdNums As Integer ', byOrdNums As String
    
cmLoad.Caption = "Обновить"
laInform.Visible = False
Me.MousePointer = flexHourglass

fields = "": xgroup = "": from = " FROM yBook ": sum = 0
OrdNums = 0 ': byOrdNums = ""
If ckPurpose.value = 1 Then
    OrdNums = 1
'    byOrdNums = "1, "
    sum = widthPurpose
    fields = "yGuidePurpose.pDescript"
    from = " FROM yBook LEFT JOIN yGuidePurpose ON (yGuidePurpose.pId = " & _
    "yBook.purposeId) AND (yGuidePurpose.subKredit = yBook.subKredit) AND " & _
    "(yGuidePurpose.Kredit = yBook.Kredit) AND (yGuidePurpose.subDebit = " & _
    "yBook.subDebit) AND (yGuidePurpose.Debit = yBook.Debit)"
End If
If ckVenture.value = 1 Then
    OrdNums = OrdNums + 1
'    byOrdNums = byOrdNums & OrdNums & ", "
    sum = sum + widthVenture
    If fields <> "" Then fields = fields & ", "
    fields = fields & "v.ventureName"
    If from = "" Then from = " FROM yBook "
    from = from & " Left JOIN GuideVenture v on v.ventureId = yBook.ventureId "
End If
If ckNumOrder.value = 1 Then
    OrdNums = OrdNums + 1
'    byOrdNums = byOrdNums & OrdNums & ", "
    sum = sum + widthNumOrder
    If fields <> "" Then fields = fields & ", "
    fields = fields & "yBook.ordersNum"
End If
If ckDKreditor.value = 1 Then
    OrdNums = OrdNums + 1
'    byOrdNums = byOrdNums & OrdNums & ", "
    sum = sum + widthDKreditor
    If fields <> "" Then fields = fields & ", "
    fields = fields & "yBook.KredDebitor"
End If
If ckDate.value = 1 Then
    OrdNums = OrdNums + 1
'    byOrdNums = byOrdNums & OrdNums & ", "
'    sum = sum + widthDate
    If fields <> "" Then fields = fields & ", "
'    fields = fields & "CDate(Format$([yBook].[xDate],'dd/mm/yy'))" '$CDATE$
    fields = fields & "DATEFORMAT([yBook].[xDate],'yy.mm.dd')" '$CDATE$

    xgroup = " GROUP BY " & fields
    fields = fields & " as cDate, "
ElseIf fields <> "" Then
    xgroup = " GROUP BY " & fields
    fields = fields & ", "
End If
    
'делим всю ширину проп-но номинальной(т.е. когда все присутствуют)
If sum > 0 Then
    sum = (widthPurpose + widthVenture + widthNumOrder + widthDKreditor) / sum
Else
    sum = 1
End If
Grid.ColWidth(byPurpose) = sum * widthPurpose * ckPurpose.value
Grid.ColWidth(byVenture) = sum * widthVenture * ckVenture.value
Grid.ColWidth(byOrdersNum) = sum * widthNumOrder * ckNumOrder.value
Grid.ColWidth(byDebKreditor) = sum * widthDKreditor * ckDKreditor.value
Grid.ColWidth(byDate) = widthDate * ckDate.value

dateWhereBef = getWhereByDateBoxes(Me, "yBook.xDate", begDate, "befo")
dateWhere = getWhereByDateBoxes(Me, "yBook.xDate", begDate)
reportDateWhere = dateWhere
If dateWhereBef = "error" Or dateWhere = "error" Then GoTo EN1
If dateWhere <> "" Then dateWhere = " WHERE ( " & dateWhere & " ) "


debitSql = lbSchetsToSql("Debit")
If debitSql = "" Then
    MsgBox "Выберите один или несколько счетов!", , "Предупреждение"
    GoTo EN1
End If
kreditSql = lbSchetsToSql("Kredit")

'sum (if yBook.debit = 50 and ybook.subdebit = 01 then ybook.uesumm else
'    0 endif) as debit,
'sum (if yBook.krebit = 50 and ybook.subkredit = 01 then ybook.uesumm
'    else 0 endif) as debit
'where ...

'dSql = " Sum((" & debitSql & ")*(-[yBook].[UEsumm]))"
dSql = " Sum(IF" & debitSql & " THEN [yBook].[UEsumm] ELSE 0 ENDIF)"

'kSql = " Sum((" & kreditSql & ")*(-[yBook].[UEsumm]))"
kSql = " Sum(IF" & kreditSql & " THEN [yBook].[UEsumm] ELSE 0 ENDIF)"
'MsgBox debitSql & vbCrLf & kreditSql

sql = "SELECT Sum(begDebit), Sum(begKredit) from yGuideSchets " & _
"WHERE " & lbSchetsToSql & ";"
'Debug.Print sql
If Not byErrSqlGetValues("##379", sql, d, k) Then GoTo EN1
sum = d - k

d = 0: k = 0
If dateWhereBef <> "" Then
    sql = "SELECT " & dSql & " AS Debit," & kSql & " AS Kredit " & _
    " FROM yBook  WHERE ( " & dateWhereBef & " );"

'MsgBox sql
'Debug.Print sql
    If Not byErrSqlGetValues("##363", sql, d, k) Then GoTo EN1
End If


sum = Round(sum + d - k, 2)
If sum < 0 Then
    tbBegDebet.Text = 0
    tbBegKredit.Text = -sum
Else
    tbBegDebet.Text = sum
    tbBegKredit.Text = 0
End If
    
str = "": For i = 1 To OrdNums: str = str & i & ", ": Next i
If str <> "" Then str = " ORDER BY " & Left$(str, Len(str) - 2)
'Label8.Caption = "'" & str & "'"

sql = "SELECT " & fields & dSql & " AS Debit," & kSql & " AS Kredit" & from & _
dateWhere & xgroup & " HAVING (" & dSql & ">0 OR " & kSql & ">0) " & str
    

'MsgBox sql
Debug.Print sql
Set Table = myOpenRecordSet("##361", sql, dbOpenForwardOnly) ' dbOpenDynaset)
If Table Is Nothing Then GoTo EN1
clearGrid Grid
quantity = 0
d = 0: k = 0
While Not Table.EOF
    quantity = quantity + 1
    If ckVenture.value = 1 Then _
        If Not IsNull(Table!ventureName) Then Grid.TextMatrix(quantity, byVenture) = Table!ventureName
    If ckPurpose.value = 1 Then _
        If Not IsNull(Table!pDescript) Then Grid.TextMatrix(quantity, byPurpose) = Table!pDescript
    If ckNumOrder.value = 1 Then _
        If Not IsNull(Table!ordersNum) Then Grid.TextMatrix(quantity, byOrdersNum) = Table!ordersNum
    If ckDKreditor.value = 1 Then
        i = Table!KredDebitor
        Grid.TextMatrix(quantity, 0) = i ' нужен для reportDateWhere
        If i > 0 Then
            sql = "SELECT Name From GuideFirms WHERE (((FirmId)=" & i & "));"
            GoTo AA
        ElseIf i < 0 Then
            sql = "SELECT Name From yDebKreditor WHERE (((id)=" & i & "));"
AA:         If byErrSqlGetValues("W##362", sql, str) Then _
            Grid.TextMatrix(quantity, byDebKreditor) = str
        End If
    End If
    If ckDate.value = 1 Then _
        Grid.TextMatrix(quantity, byDate) = Right$(Table!CDate, 2) & _
        Mid$(Table!CDate, 3, 4) & Left$(Table!CDate, 2)
    d = d + Table!debit
    Grid.TextMatrix(quantity, byDebit) = Round(Table!debit, 2)
    k = k + Table!kredit
    Grid.TextMatrix(quantity, byKredit) = Round(Table!kredit, 2)
    Grid.AddItem ""
    Table.MoveNext
Wend
Table.Close
Grid.Visible = True

d = Round(d, 2)
k = Round(k, 2)
tbMidDebit.Text = d
tbMidKredit.Text = k
sum = Round(d + tbBegDebet.Text - k - tbBegKredit.Text, 2)
If sum < 0 Then
    tbEndDebet.Text = 0
    tbEndKredit.Text = -sum
Else
    tbEndDebet.Text = sum
    tbEndKredit.Text = 0
End If

If quantity > 0 Then
    Grid.RemoveItem quantity + 1
    Grid.row = 1
    Grid.col = 1
    On Error Resume Next
    Grid_EnterCell
    Grid.SetFocus
End If
EN1:
Me.MousePointer = flexDefault
End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width

tbStartDate.Text = "01." & Format(CurDate, "mm/yy")
tbEndDate.Text = Format(CurDate, "dd/mm/yy")

ckBySub.value = 0

Grid.FormatString = "|<Назначение|<Предприятие|<Договор|Дебитор\Кредитор|<Дата|Дебет|Кредит"
Grid.ColWidth(0) = 0
Grid.ColWidth(byPurpose) = 0
Grid.ColWidth(byOrdersNum) = 0
Grid.ColWidth(byDebKreditor) = 0
Grid.ColWidth(byDate) = 0
Grid.ColWidth(byDebit) = 1000
Grid.ColWidth(byKredit) = 1000
isLoad = True
End Sub


Sub lbResize()
Dim heig As Integer, i As Integer

'i = Int((cmLoad.Top - lbSchets.Top) / 195)
i = Int((laGroup.Top - lbSchets.Top) / 195)
If i > lbSchets.ListCount Then i = lbSchets.ListCount
lbSchets.Height = 195 * i + 100
End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer
If WindowState = vbMinimized Then Exit Sub
On Error Resume Next

h = Me.Height - oldHeight
oldHeight = Me.Height
'w = Me.Width - oldWidth
'oldWidth = Me.Width
Me.Width = oldWidth
Grid.Height = Grid.Height + h
'Grid.Width = Grid.Width + w

cmLoad.Top = cmLoad.Top + h
cmExit.Top = cmExit.Top + h
laGroup.Top = laGroup.Top + h
ckPurpose.Top = ckPurpose.Top + h
ckVenture.Top = ckVenture.Top + h
ckNumOrder.Top = ckNumOrder.Top + h
ckDKreditor.Top = ckDKreditor.Top + h
ckDate.Top = ckDate.Top + h
'cmExit.Left = cmExit.Left + w

lbResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
isLoad = False
End Sub

Private Sub Grid_DblClick()
Dim str As String

If Grid.CellBackColor <> &H88FF88 Then Exit Sub

If MsgBox("Вы хотите посмотреть записи, которые образуют сумму " & _
Grid.TextMatrix(mousRow, mousCol), vbDefaultButton2 Or vbYesNo, _
"Продолжить?") = vbNo Then Exit Sub
 
'строим SQL исходя из установок на момент нажатия <Обновить>
If Grid.ColWidth(byDate) > 0 Then ' тогда даты периода не учитываем
    str = Grid.TextMatrix(mousRow, byDate)
    str = Left$(str, 6) & "20" & Mid$(str, 7, 2)
    filtrWhere = "((yBook.xDate) Like '" & str & "*')"
Else
    filtrWhere = reportDateWhere
End If
If mousCol = byDebit Then
    str = debitSql
    tmpStr = "Дебет = " & Grid.TextMatrix(mousRow, byDebit)
Else
    str = kreditSql
    tmpStr = "Кредит = " & Grid.TextMatrix(mousRow, byKredit)
End If
Journal.laFiltr.Caption = "Записи, дающие " & tmpStr
Journal.laFiltr.Visible = True
str = "(" & str & ")"
If filtrWhere = "" Then
    filtrWhere = str
Else
    filtrWhere = str & " AND (" & filtrWhere & ")"
End If

'по ширине узнаем, какие галки б. установлены
If Grid.ColWidth(byPurpose) > 0 Then
    If filtrWhere <> "" Then filtrWhere = filtrWhere & " AND "
    filtrWhere = filtrWhere & "((yGuidePurpose.pDescript)='" & _
    Grid.TextMatrix(mousRow, byPurpose) & "')"
End If
If Grid.ColWidth(byVenture) > 0 Then
    If filtrWhere <> "" Then filtrWhere = filtrWhere & " AND "
    filtrWhere = filtrWhere & "(isnull(v.ventureName, '')='" & _
    Grid.TextMatrix(mousRow, byVenture) & "')"
End If
If Grid.ColWidth(byOrdersNum) > 0 Then
    If filtrWhere <> "" Then filtrWhere = filtrWhere & " AND "
    filtrWhere = filtrWhere & "((yBook.OrdersNum)='" & _
    Grid.TextMatrix(mousRow, byOrdersNum) & "')"
End If
If Grid.ColWidth(byDebKreditor) > 0 Then
    If filtrWhere <> "" Then filtrWhere = filtrWhere & " AND "
    filtrWhere = filtrWhere & "((yBook.KredDebitor)=" & _
    Grid.TextMatrix(mousRow, 0) & ")"
End If

'MsgBox filtrWhere
Journal.ZOrder
Journal.loadBook
End Sub

Private Sub Grid_EnterCell()
If quantity = 0 Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col

If (mousCol = byDebit Or mousCol = byKredit) And _
Grid.TextMatrix(mousRow, mousCol) <> "0" Then
   Grid.CellBackColor = &H88FF88
Else
   Grid.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor

End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)

End Sub


Private Sub lbSchets_Click()
cmLoad.Caption = "Загрузить"

End Sub

Private Sub tbEndDate_Change()
cmLoad.Caption = "Загрузить"

End Sub

Private Sub tbStartDate_Change()
cmLoad.Caption = "Загрузить"
End Sub

Function lbSchetsToSql(Optional DKredit As String = "") As String
Dim i As Integer, schetSql As String, subSchet As String
Dim sTable As String

lbSchetsToSql = ""
If DKredit = "" Then
    sTable = "[yGuideSchets].[": DKredit = "Number"
Else
    sTable = "[yBook].["
End If
For i = 0 To lbSchets.ListCount - 1
    If lbSchets.Selected(i) Then
'        schetSql = "[yBook].[" & DKredit & "]=" & Left$(lbSchets.List(i), 2)
        schetSql = sTable & DKredit & "]=" & "'" & Left$(lbSchets.List(i), 2) & "'"
        If ckBySub.value = 1 Then
            subSchet = "'" & Mid$(lbSchets.List(i), 4) & "'"
            If subSchet = "''" Then subSchet = "'00'"
'            If Not IsNumeric(subSchet) Then
'                subSchet = "'" & subSchet & "'"
'            End If
            
'            schetSql = schetSql & " AND [yBook].[sub" & DKredit & "]=" & subSchet
            schetSql = schetSql & " AND " & sTable & "sub" & DKredit & "]=" & subSchet
        End If
        
        If lbSchetsToSql = "" Then
            lbSchetsToSql = "(" & schetSql & ")"
        Else
'            lbSchetsToSql = "(" & lbSchetsToSql & ") OR (" & schetSql & ")"
            lbSchetsToSql = lbSchetsToSql & " OR (" & schetSql & ")"
        End If
    End If
Next i
End Function



