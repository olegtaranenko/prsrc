VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form GuideStatia 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Справочник статей затрат"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Text            =   "tbMobile"
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   3060
      Width           =   855
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   3060
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   3060
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   4895
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "GuideStatia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public isLoad As Boolean
Public mousRow As Long    '
Public mousCol As Long    '
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim quantity As Integer 'количество найденных фирм
Dim frmMode As String

Const giSourceId = 0 ' скрытый
Const giName = 1
Const giNote = 2

Const sysQuant = 9 'кол-во системных статей


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
Dim i As Integer

sql = "DELETE From sGuideSource " & _
"WHERE (((sGuideSource.sourceId)=" & gSourceId & "));"
i = myExecute("##471", sql, -198)
If i = 0 Then
    quantity = quantity - 1
    If quantity > 0 Then Grid.RemoveItem mousRow
    If quantity <= sysQuant Then cmDel.Visible = False
ElseIf i = -2 Then
    MsgBox "Этот поставщик используется в некоторых документах.", , _
    "Удаление невозможно!"
End If

On Error Resume Next
Grid_EnterCell
Grid.SetFocus
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width

Grid.FormatString = "|< Название |<Примечание"
Grid.ColWidth(giSourceId) = 0
Grid.ColWidth(giName) = 1545
Grid.ColWidth(giNote) = 1980
sql = "SELECT sGuideSource.sourceName, sGuideSource.FIO, sGuideSource.Phone, " & _
"sGuideSource.Fax, sGuideSource.Email, sGuideSource.sourceId " & _
"From sGuideSource Where (((sGuideSource.sourceId) < 0 AND " & _
"(sGuideSource.sourceId)> -1001)) ORDER BY sGuideSource.sourceId DESC;"
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
Grid.RemoveItem quantity + 1
Grid_EnterCell
If quantity > sysQuant Then cmDel.Visible = True

isLoad = True
End Sub

Private Sub MSFlexGrid1_Click()

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
     If mousRow > sysQuant Then  'это системные статьи
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
  If str = "" Then
    MsgBox "Недпустимое значение", , "Предупрежение"
    Exit Sub
  End If
  If mousCol = giName Then
    If frmMode = "sourceAdd" Then
'      sql = "SELECT Max(sGuideSource.sourceId) AS Max FROM sGuideSource;"
      
      wrkDefault.BeginTrans 'lock01
      sql = "UPDATE sGuideSource SET sourceId = sourceId WHERE sourceId=0" 'lock02
      myBase.Execute (sql) 'lock03
      
      sql = "SELECT Min(sourceId) FROM sGuideSource WHERE sourceId > -1001"
      If Not byErrSqlGetValues("##380", sql, gSourceId) Then GoTo EN1
      gSourceId = gSourceId - 1
      
'      sql = "SELECT sGuideSource.sourceId, sGuideSource.sourceName " & _
'      "From sGuideSource WHERE (((sGuideSource.sourceId)<0 AND " & _
'      "(sGuideSource.sourceId)>-1001)) ORDER BY sGuideSource.sourceId DESC;"
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
      sql = "INSERT INTO sGuideSource (sourceId, SourceName) " & _
      "VALUES (" & gSourceId & ", '" & str & "')"
      i = myExecute("##457", sql, -196)
      If i <> 0 Then GoTo ERR0
      
      wrkDefault.CommitTrans
      
      Grid.TextMatrix(mousRow, giSourceId) = gSourceId
      quantity = quantity + 1
      If quantity > sysQuant Then cmDel.Visible = True
    Else
       i = ValueToGuideSourceField("##142", str, "sourceName", -196)
       If i <> 0 Then GoTo ERR0
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
'If Err = 3022 Then
'    MsgBox "Это название уже есть", , "Ошибка"
'    GoTo CNC
'Else
'    MsgBox Error, , "Ошибка 143-" & Err & ":  " '##143
'    End
'End If
ERR0:
If i = -2 Then
    MsgBox "Это название уже есть (возможно в Cправочнике поставщиков " & _
    "или Справочнике внутренних подразделений - что тоже не допускается).", , "Ошибка - " & cErr
    tbMobile.SetFocus
Else
    GoTo EN1
End If

End Sub






