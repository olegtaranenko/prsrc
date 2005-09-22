VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form GuideSource 
   BackColor       =   &H8000000A&
   Caption         =   "Справочник поставщиков"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   1740
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   9195
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   900
      TabIndex        =   4
      Text            =   "tbMobile"
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   8280
      TabIndex        =   3
      Top             =   5700
      Width           =   855
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   5700
      Width           =   855
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   5700
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   9340
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "GuideSource"
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

Const gpNazwFirm = 1
Const gpFIO = 2
Const gpTlf = 3
Const gpFax = 4
Const gpEmail = 5
Const gpSourceId = 6 ' скрытый


Private Sub Command3_Click()

End Sub

Private Sub cmAdd_Click()
frmMode = "sourceAdd"
If quantity > 0 Then Grid.AddItem ("")
Grid.row = Grid.Rows - 1
mousRow = Grid.Rows - 1
Grid.col = gpNazwFirm
mousCol = gpNazwFirm
On Error Resume Next
Grid.SetFocus
textBoxInGridCell tbMobile, Grid

End Sub

Private Sub cmDel_Click()
Dim i As Integer
If gSourceId = 34 Or gSourceId = 40 Then 'Инвентаризация и Коррекция
    MsgBox "Это не поставщик, а системная статья прихода.", , "Удаление невозможно!"
    Exit Sub
End If
sql = "DELETE  From sGuideSource " & _
"WHERE (((sGuideSource.sourceId)=" & gSourceId & "));"
i = myExecute("##470", sql, -198)
If i = 0 Then
    quantity = quantity - 1
    If quantity > 0 Then Grid.RemoveItem mousRow
ElseIf i = -2 Then
    MsgBox "Этот поставщик используется в некоторых документах.", , _
    "Удаление невозможно!"
End If
On Error Resume Next
Grid.SetFocus
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width

Grid.FormatString = "|< Название поставщиака|<Конт.лицо|<Телефон|<Факс|<e-mail|"
Grid.ColWidth(0) = 0
Grid.ColWidth(gpNazwFirm) = 2025
Grid.ColWidth(gpFIO) = 2880
Grid.ColWidth(gpTlf) = 1440
Grid.ColWidth(gpFax) = 1260
Grid.ColWidth(gpEmail) = 1005
Grid.ColWidth(gpSourceId) = 0
sql = "SELECT sGuideSource.sourceName, sGuideSource.FIO, sGuideSource.Phone, " & _
"sGuideSource.Fax, sGuideSource.Email, sGuideSource.sourceId " & _
"From sGuideSource " & _
"Where (((sGuideSource.sourceId) > 0)) ORDER BY sGuideSource.sourceName;"
Set tbGuide = myOpenRecordSet("##140", sql, dbOpenForwardOnly)
If tbGuide Is Nothing Then Exit Sub

quantity = 0
While Not tbGuide.EOF
    quantity = quantity + 1
    Grid.TextMatrix(quantity, gpNazwFirm) = tbGuide!SourceName
    Grid.TextMatrix(quantity, gpFIO) = tbGuide!FIO
    Grid.TextMatrix(quantity, gpTlf) = tbGuide!Phone
    Grid.TextMatrix(quantity, gpFax) = tbGuide!Fax
    Grid.TextMatrix(quantity, gpEmail) = tbGuide!Email
    Grid.TextMatrix(quantity, gpSourceId) = tbGuide!sourceId
    Grid.AddItem ""

    tbGuide.MoveNext
Wend
tbGuide.Close
If quantity > 0 Then Grid.RemoveItem quantity + 1
Grid_EnterCell
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
        If gSourceId = 34 Or gSourceId = 40 Then 'Инвентаризация и Коррекция
            MsgBox "Это не поставщик, а системная статья прихода. " & _
            "Изменение названия не изменит ее суть!", , "Внимание!"
        End If
        textBoxInGridCell tbMobile, Grid
End If

End Sub

Private Sub Grid_EnterCell()
If quantity > 0 Then
 mousRow = Grid.row
 mousCol = Grid.col
 gSourceId = Grid.TextMatrix(mousRow, gpSourceId)
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
  If mousCol = gpNazwFirm Then
    If frmMode = "sourceAdd" Then

      wrkDefault.BeginTrans 'lock01
      sql = "UPDATE sGuideSource SET sourceId = sourceId WHERE sourceId=0" 'lock02
      myBase.Execute (sql) 'lock03

      sql = "SELECT Max(sourceId) FROM sGuideSource;"
      If Not byErrSqlGetValues("##380", sql, gSourceId) Then GoTo EN1
      gSourceId = gSourceId + 1

'      sql = "SELECT sGuideSource.sourceId, sGuideSource.sourceName " & _
'      "From sGuideSource ORDER BY sGuideSource.sourceId;"
'      Set tbGuide = myOpenRecordSet("##142", sql, dbOpenDynaset)
'      If tbGuide Is Nothing Then Exit Sub
'      tbGuide.MoveLast
'      gSourceId = tbGuide!sourceId + 1
'      tbGuide.AddNew
'      tbGuide!sourceId = gSourceId
'      tbGuide!SourceName = str
'      On Error GoTo ERR1
'      tbGuide.Update
'      tbGuide.Close
      sql = "INSERT INTO sGuideSource (sourceId,SourceName) " & _
      "VALUES (" & gSourceId & ", '" & str & "')"
      i = myExecute("##464", sql, -196)
      If i <> 0 Then GoTo ERR0
      
      wrkDefault.CommitTrans
              
      Grid.TextMatrix(mousRow, gpSourceId) = gSourceId
      quantity = quantity + 1
    Else
      i = ValueToGuideSourceField("##142", str, "sourceName", -196)
      If i <> 0 Then GoTo ERR0
'       If i = 3022 Then
'            existMsg
'            Exit Sub
'       ElseIf i <> 0 Then
'            GoTo EN1
'       End If
    End If
  ElseIf mousCol = gpFIO Then
       If ValueToGuideSourceField("##142", str, "FIO") <> 0 Then GoTo EN1
  ElseIf mousCol = gpTlf Then
       If ValueToGuideSourceField("##142", str, "Phone") <> 0 Then GoTo EN1
  ElseIf mousCol = gpFax Then
       If ValueToGuideSourceField("##142", str, "Fax") <> 0 Then GoTo EN1
  ElseIf mousCol = gpEmail Then
       If ValueToGuideSourceField("##142", str, "Email") <> 0 Then GoTo EN1
  End If
  
  Grid.TextMatrix(mousRow, mousCol) = str
  GoTo EN1
ElseIf KeyCode = vbKeyEscape Then
CNC:
 If mousCol = gpNazwFirm And frmMode = "sourceAdd" Then
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
'If errorCodAndMsg("##105", -193) Then
'    existMsg
'    MsgBox "Это название уже есть.", , "Ошибка"
'    GoTo CNC
'End If

ERR0:
If i = -2 Then
    MsgBox "Это название уже есть (возможно в Cправочнике статей расхода " & _
    "или Справочнике внутренних подразделений - что тоже не допускается).", , "Ошибка" & cErr
    tbMobile.SetFocus
Else
    GoTo EN1
End If

End Sub
Sub existMsg()
End Sub
'Function ValueToGuideSourceField(myErrCod As String, value As String, _
'field As String, Optional passErr As Integer = -1) As Integer
'Dim i As Integer'

'ValueToGuideSourceField = False
'sql = "UPDATE sGuideSource SET sGuideSource." & field & _
'" = '" & value & "' WHERE (((sGuideSource.sourceId)=" & gSourceId & "));"
''MsgBox "sql = " & sql

'ValueToGuideSourceField = myExecute(myErrCod, sql, passErr)
'End Function
