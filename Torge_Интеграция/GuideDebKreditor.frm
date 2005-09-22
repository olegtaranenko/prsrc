VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form GuideDebKreditor 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Справочник дебиторов-кредиторов"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   3930
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
      Left            =   1560
      TabIndex        =   3
      Top             =   3060
      Width           =   855
   End
   Begin VB.CommandButton cmExit 
      Cancel          =   -1  'True
      Caption         =   "Выход"
      Height          =   315
      Left            =   2880
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
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   4895
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "GuideDebKreditor"
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
Dim gDKid As String

Const giDKid = 0 ' скрытый
Const giName = 1
Const giNote = 2


Private Sub Command3_Click()

End Sub

Private Sub cmAdd_Click()
frmMode = "Add"
If quantity > 0 Then Grid.AddItem ("")
Grid.row = Grid.Rows - 1
mousRow = Grid.Rows - 1
Grid.col = giName
mousCol = giName
Grid_EnterCell
On Error Resume Next
Grid.SetFocus
textBoxInGridCell tbMobile, Grid

End Sub

Private Sub cmDel_Click()
Dim i As Integer

sql = "DELETE  From yDebKreditor " & _
"WHERE (((yDebKreditor.id)=" & gDKid & "));"
'MsgBox sql
i = myExecute("##352", sql, -198) '3200)
If i = 0 Then
    quantity = quantity - 1
    If quantity > 0 Then
        Grid.RemoveItem mousRow
    Else '$$4
        clearGridRow Grid, 1
    End If
ElseIf i = -2 Then
    MsgBox "Этот Дебитор\Кредитор используется в некоторых документах.", , _
    "Удаление невозможно!"
End If
On Error Resume Next
Grid.SetFocus
Grid_EnterCell

End Sub

Private Sub cmExit_Click()
Journal.loadLbFromDebKreditor
Unload Me
End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width

Grid.FormatString = "|< Название |<Примечание"
Grid.ColWidth(giDKid) = 0
Grid.ColWidth(giName) = 1545
Grid.ColWidth(giNote) = 1980

sql = "SELECT * FROM yDebKreditor ORDER BY  name"
Set tbGuide = myOpenRecordSet("##348", sql, dbOpenForwardOnly)
'If tbGuide Is Nothing Then Exit Sub

'tbGuide.Index = "Name"
quantity = 0
While Not tbGuide.EOF
  If tbGuide!id <> 0 Then
    quantity = quantity + 1
    Grid.TextMatrix(quantity, giName) = tbGuide!Name
    Grid.TextMatrix(quantity, giNote) = tbGuide!note
    Grid.TextMatrix(quantity, giDKid) = tbGuide!id
    Grid.AddItem ""
  End If
  tbGuide.MoveNext
Wend
tbGuide.Close


If quantity > 0 Then Grid.RemoveItem quantity + 1
Grid_EnterCell
isLoad = True
End Sub

Function ValueToGuideDKrField(myErrCod As String, value As String, _
field As String, Optional passErr As Integer = -1) As Integer
Dim i As Integer

ValueToGuideDKrField = False
sql = "UPDATE yDebKreditor SET [" & field & _
"] = '" & value & "' WHERE (((id)=" & gDKid & "));"
'MsgBox "sql = " & sql

ValueToGuideDKrField = myExecute(myErrCod, sql, passErr)
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
 gDKid = Grid.TextMatrix(mousRow, giDKid)
 
' If mousRow < 2 Then  'это системное подразделение
'    cmDel.Enabled = False
' Else
'    cmDel.Enabled = True
' End If

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
    If frmMode = "Add" Then
        
      wrkDefault.BeginTrans 'lock01
      sql = "UPDATE yDebKreditor SET id = id WHERE id=0" 'lock02
      myBase.Execute (sql) 'lock03
        
    
'      Set tbGuide = myOpenRecordSet("##349", "yDebKreditor", dbOpenTable)
'      If tbGuide Is Nothing Then Exit Sub
'      tbGuide.Index = "Key"
'      gDKid = tbGuide!id - 1
'      tbGuide.AddNew
'      tbGuide!id = gDKid
'      tbGuide!Name = str
'      On Error GoTo ERR1
'      tbGuide.Update
'      tbGuide.Close
      
      sql = "SELECT min(id) FROM yDebKreditor"
      If Not byErrSqlGetValues("##349", sql, gDKid) Then GoTo EN1
      gDKid = gDKid - 1
      
      sql = "INSERT INTO yDebKreditor (id,Name) " & _
      "VALUES (" & gDKid & ", '" & str & "')"
      i = myExecute("##351", sql, -196)
      If i <> 0 Then GoTo ERR0
      
      wrkDefault.CommitTrans
      
      Grid.TextMatrix(mousRow, giDKid) = gDKid
      quantity = quantity + 1
    Else
       i = ValueToGuideDKrField("##350", str, "Name", -196)
       If i <> 0 Then GoTo ERR0
'       If i = -2 Then
'            MsgBox "Это название уже есть", , "Ошибка!"
'            Exit Sub
'       ElseIf i <> 0 Then
'            GoTo EN1
'       End If
    End If
  ElseIf mousCol = giNote Then
       If ValueToGuideDKrField("##350", str, "Note") <> 0 Then GoTo EN1
  End If
  Grid.TextMatrix(mousRow, mousCol) = str
  GoTo EN1
ElseIf KeyCode = vbKeyEscape Then
CNC:
 If mousCol = giName And frmMode = "Add" Then
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
ERR0:
'tbGuide.Close
'If Err = 3022 Then
If i = -2 Then
    MsgBox "Это название уже есть", , "Ошибка = " & cErr
    tbMobile.SetFocus
'    GoTo CNC
Else
    GoTo EN1
'    MsgBox Error, , "Ошибка 351-" & Err & ":  " '##351
'    End
End If

End Sub


