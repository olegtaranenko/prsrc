VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GuideInside 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Справочник внутренних подразделений"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   3060
      Width           =   915
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Height          =   315
      Left            =   2580
      TabIndex        =   3
      Top             =   3060
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   5040
      TabIndex        =   2
      Top             =   3060
      Width           =   855
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "tbMobile"
      Top             =   1140
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   4895
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "GuideInside"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isLoad As Boolean
Public mousRow As Long    '
Public mousCol As Long    '
Dim quantity As Integer 'количество найденных фирм
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim frmMode As String

Const giSourceId = 0 ' скрытый
Const giName = 1
Const giNote = 2

Const sysQuant = 2 'кол-во системных подразделений



Private Sub Command3_Click()

End Sub

Private Sub cmAdd_Click()
frmMode = "sourceAdd"
If quantity > 0 Then Grid.AddItem ("")
Grid.row = Grid.Rows - 1
mousRow = Grid.Rows - 1
Grid.col = giName
mousCol = giName
On Error Resume Next
Grid.SetFocus
textBoxInGridCell tbMobile, Grid

End Sub

Private Sub cmDel_Click()
Dim id As Integer, i As Integer

'Set tbGuide = myOpenRecordSet("##145", "sGuideSource", dbOpenTable)
'If tbGuide Is Nothing Then Exit Sub
'tbGuide.Index = "PrimaryKey"
'tbGuide.Seek "=", gSourceId
'If tbGuide.NoMatch Then
'    MsgBox "Не найден склад с id=" & gSourceId, , ""
'    tbGuide.Close
'    Exit Sub
'End If
    
wrkDefault.BeginTrans   ' начало транзакции
sql = "DELETE FROM sGuideSource WHERE sourceId = " & gSourceId
i = myExecute("##145", sql, -198)
If i <> 0 Then GoTo ERR0

sql = "SELECT min(sourceId) FROM sGuideSource" ' макс.по модулю id
If Not byErrSqlGetValues("##467", sql, id) Then Exit Sub

If id < gSourceId Then 'если удалялся максимальный
    sql = "UPDATE sGuideSource SET sourceId = " & gSourceId & _
    " WHERE sourceId = " & id
    If myExecute("##468", sql) <> 0 Then Exit Sub
End If
'On Error GoTo ERR1
'tbGuide.Delete
''максимальный по модулю sourceId заменять на удаленный, чтобы не было дырок
'tbGuide.MoveFirst ' макс.по модулю id
'If tbGuide!sourceId > gSourceId Then 'если удалялся не максимальный
'    tbGuide.Edit
'    tbGuide!sourceId = gSourceId
'    tbGuide.Update
'End If

wrkDefault.CommitTrans  ' подтверждение транзакции
'tbGuide.Close


loadInside
If quantity <= sysQuant Then cmDel.Visible = False
Documents.loadLbInside
If Nomenklatura.isLoad Then Unload Nomenklatura 'чтобы перегрузить там список складов


Exit Sub

'ERR1:
'wrkDefault.Rollback ' отммена транзакции
'tbGuide.Close

'If Err = 3200 Then
'    MsgBox "Это подразделение используется в некоторых документах.", , _
'    "Удаление невозможно!"
'Else
'    MsgBox Error, , "Ошибка 457-" & Err & ":  " '##457
'End If

ERR0:
If i = -2 Then
    MsgBox "Это подразделение используется в некоторых документах.", , _
    "Удаление невозможно!"
End If
Grid.SetFocus

End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width

loadInside
isLoad = True
End Sub

Sub loadInside()

Grid.Rows = 2
Grid.FormatString = "|< Название |<Примечание"
Grid.ColWidth(giSourceId) = 0
Grid.ColWidth(giName) = 1545
Grid.ColWidth(giNote) = 4170
sql = "SELECT * From sGuideSource " & _
"Where (((sourceId) < -1000)) ORDER BY sourceId  DESC;"
Set tbGuide = myOpenRecordSet("##140", sql, dbOpenForwardOnly)
If tbGuide Is Nothing Then Exit Sub

quantity = 0
While Not tbGuide.EOF
    quantity = quantity + 1
    Grid.TextMatrix(quantity, giName) = tbGuide!SourceName
    Grid.TextMatrix(quantity, giNote) = tbGuide!FIO
    Grid.TextMatrix(quantity, giSourceId) = tbGuide!sourceId
    Grid.AddItem ""

    tbGuide.MoveNext
Wend
tbGuide.Close

If quantity > sysQuant Then cmDel.Visible = True

Grid.RemoveItem quantity + 1
Grid_EnterCell

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

cmAdd.Top = cmAdd.Top + h
cmDel.Top = cmDel.Top + h
cmExit.Top = cmExit.Top + h
cmExit.Left = cmExit.Left + w

End Sub

Private Sub Form_Unload(Cancel As Integer)
isLoad = False
End Sub

Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If quantity = 0 Then Exit Sub
If mousRow = 0 Then
    Grid.CellBackColor = Grid.BackColor
    
'    If mousCol > 3 Then
'        SortCol Grid, mousCol, "numeric"
'    Else
        SortCol Grid, mousCol
'    End If
    Grid.row = 1    ' только чтобы снять выделение
    Grid_EnterCell
End If

End Sub

Private Sub Grid_DblClick()
If Grid.CellBackColor = &H88FF88 Then
        tbMobile.MaxLength = 50
        textBoxInGridCell tbMobile, Grid
End If

End Sub

Private Sub Grid_EnterCell()
If quantity > 0 Then
 mousRow = Grid.row
 mousCol = Grid.col
 gSourceId = Grid.TextMatrix(mousRow, giSourceId)
 
  If quantity > sysQuant Then
     If mousRow > sysQuant Then  'это системное подразделение
        cmDel.Enabled = True
     Else
        cmDel.Enabled = False
    End If
 End If


 If mousCol > 0 Then
    Grid.CellBackColor = &H88FF88
 Else
    Grid.CellBackColor = vbYellow
 End If
End If

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

Grid.Enabled = True
On Error Resume Next
Grid.SetFocus
Grid_EnterCell
End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor

End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)

End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, i As Integer

If KeyCode = vbKeyReturn Then
  str = Trim(tbMobile.Text)
  If mousCol = giName Then
    If str = "" Then
        MsgBox "Недпустимое значение", , "Предупрежение"
        Exit Sub
    End If
    If frmMode = "sourceAdd" Then
      
      wrkDefault.BeginTrans 'lock01
      sql = "UPDATE sGuideSource SET sourceId = sourceId WHERE sourceId=0" 'lock02
      myBase.Execute (sql) 'lock03
      
'      sql = "SELECT sGuideSource.sourceId, sGuideSource.sourceName " & _
'      "From sGuideSource WHERE (((sGuideSource.sourceId)<1000)) " & _
'      "ORDER BY sGuideSource.sourceId DESC;"
'      Set tbGuide = myOpenRecordSet("##142", sql, dbOpenDynaset)
'      If tbGuide Is Nothing Then Exit Sub
'      tbGuide.MoveLast
'      gSourceId = tbGuide!sourceId - 1
'      tbGuide.AddNew
'      tbGuide!sourceId = gSourceId
'      tbGuide!SourceName = str
'      On Error GoTo ERR1
'      tbGuide.Update
'      tbGuide.Close
        
      sql = "SELECT min(sourceId) FROM sGuideSource"
      If Not byErrSqlGetValues("##466", sql, gSourceId) Then GoTo EN1
      gSourceId = gSourceId - 1
      
      sql = "INSERT INTO sGuideSource (sourceId,SourceName) " & _
      "VALUES (" & gSourceId & ", '" & str & "')"
      i = myExecute("##467", sql, -196)
      If i <> 0 Then GoTo ERR0
      
      wrkDefault.CommitTrans
      
      Grid.TextMatrix(mousRow, giSourceId) = gSourceId
      quantity = quantity + 1
      If quantity > sysQuant Then cmDel.Visible = True
      Documents.loadLbInside
      If Nomenklatura.isLoad Then Unload Nomenklatura 'чтобы перегрузить там список складов
    Else
      i = ValueToGuideSourceField("##142", str, "sourceName", -196)
      If i <> 0 Then GoTo ERR0
'       i = ValueToGuideSourceField("##142", str, "sourceName", 3022)
'       If i = 3022 Then
'            MsgBox "Это название уже есть", , "Ошибка!"
'            Exit Sub
'       ElseIf i <> 0 Then
'            GoTo EN1
'       End If
    End If
  ElseIf mousCol = giNote Then
       If ValueToGuideSourceField("##142", str, "FIO") <> 0 Then GoTo EN1
  End If
  Grid.TextMatrix(mousRow, mousCol) = str
  GoTo EN1
ElseIf KeyCode = vbKeyEscape Then
CNC:
 If mousCol = giName And frmMode = "sourceAdd" Then
    If quantity > 0 Then
        Grid.RemoveItem quantity + 1 ' ту, которую зря добавили
    End If
 End If
EN1:
 frmMode = ""
 lbHide
End If
Exit Sub

'ERR1:
'tbGuide.Close
'errorCodAndMsg "##143"

'If Err = 3022 Then
'    MsgBox "Это название уже есть", , "Ошибка"
'    GoTo CNC
'Else
'    MsgBox Error, , "Ошибка 143-" & Err & ":  " '##143
'    End
'End If

ERR0:
If i = -2 Then
    MsgBox "Это название уже есть (возможно в Cправочнике статей расхода " & _
    "или Справочнике поставщиков - что тоже не допускается).", , "Ошибка-" & cErr
    tbMobile.SetFocus
Else
    GoTo EN1
End If


End Sub


