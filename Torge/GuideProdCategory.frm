VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GuideProdCategory 
   BackColor       =   &H8000000A&
   Caption         =   "Категории Готовых изделий"
   ClientHeight    =   5148
   ClientLeft      =   60
   ClientTop       =   1740
   ClientWidth     =   6756
   LinkTopic       =   "Form1"
   ScaleHeight     =   5148
   ScaleWidth      =   6756
   StartUpPosition =   1  'CenterOwner
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
      Left            =   5700
      TabIndex        =   3
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   4680
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   6495
      _ExtentX        =   11451
      _ExtentY        =   7430
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "GuideProdCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isLoad As Boolean
Public mousRow As Long    '
Public mousCol As Long    '
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim quantity As Integer 'количество найденных констант
Dim frmMode As String
Dim gProdCategoryId As String
Dim sourId As Integer, destId As Integer

Const gmProdCategoryId = 0 ' скрытый
Const gmProdCategory = 1
Const gmValue = 2



Private Sub cmAdd_Click()
frmMode = "sourceAdd"
If quantity > 0 Then Grid.AddItem ("")
Grid.row = Grid.Rows - 1
mousRow = Grid.Rows - 1
Grid.col = gmProdCategory
mousCol = gmProdCategory
cmAdd.Enabled = False
cmDel.Enabled = False
On Error Resume Next
Grid.SetFocus
textBoxInGridCell tbMobile, Grid

End Sub

Private Sub cmDel_Click()
Dim I As Integer
sql = "DELETE  From GuideProdCategory WHERE (((ProdCategoryId)=" & gProdCategoryId & "));"
I = myExecute("##440", sql, -198)
If I = 0 Then
    quantity = quantity - 1
    If quantity > 0 Then
        Grid.RemoveItem mousRow
    Else
        clearGridRow Grid, 1
    End If
ElseIf I = -2 Then
    MsgBox "У этого Менеджера есть заказы либо он задействовон в справочниках " & _
    "Фирм.", , "Удаление невозможно!"
End If

On Error Resume Next
Grid.SetFocus
Grid_EnterCell
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width

Grid.FormatString = "|<Название|<Значение"
Grid.colWidth(gmProdCategoryId) = 0
Grid.colWidth(gmProdCategory) = 585
Grid.colWidth(gmValue) = 4545
sql = "SELECT ProdCategoryId, Sysname, nameRu From GuideProdCategory "
Set tbGuide = myOpenRecordSet("##441", sql, dbOpenForwardOnly)
If tbGuide Is Nothing Then Exit Sub

quantity = 0
While Not tbGuide.EOF
    quantity = quantity + 1
    Grid.TextMatrix(quantity, gmProdCategoryId) = tbGuide!prodCategoryId
    Grid.TextMatrix(quantity, gmProdCategory) = tbGuide!sysname
    Grid.TextMatrix(quantity, gmValue) = tbGuide!nameRu
    Grid.AddItem ""

    tbGuide.MoveNext
Wend
tbGuide.Close
If quantity > 0 Then
    Grid.RemoveItem quantity + 1
    Grid.col = 1
    Grid.row = 1
    mousRow = 1
    mousCol = 1
    Grid_EnterCell
End If

isLoad = True
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
cmExit.left = cmExit.left + w

End Sub

Private Sub Form_Unload(Cancel As Integer)
    isLoad = False
End Sub

Private Sub Grid_Click()
Static sourDest As Boolean

mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If mousRow = 0 Then Exit Sub

End Sub

Private Sub Grid_DblClick()
If mousRow = 0 Then Exit Sub
If Grid.CellBackColor = &H88FF88 Then
        textBoxInGridCell tbMobile, Grid
End If

End Sub

Private Sub Grid_EnterCell()
 
If quantity > 0 Then
 mousRow = Grid.row
 mousCol = Grid.col
 gProdCategoryId = Grid.TextMatrix(mousRow, gmProdCategoryId)
 

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
cmAdd.Enabled = True
cmDel.Enabled = True

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
        MsgBox "ColWidth = " & Grid.colWidth(Grid.MouseCol)

End Sub


Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, I As Integer
Dim initStr As String

If KeyCode = vbKeyReturn Then
  str = Trim(tbMobile.Text)
  If str = "" Then
    MsgBox "Недпустимое значение", , "Предупрежение"
    Exit Sub
  End If
  If mousCol = gmProdCategory Then
    If frmMode = "sourceAdd" Then
        sql = "INSERT INTO GuideProdCategory (sysname) VALUES ( '" & str & "')"
        If myExecute("##465", sql) <> 0 Then GoTo EN1
        
        sql = "select ProdCategoryID from GuideProdCategory where sysname = '" & str & "'"
        byErrSqlGetValues "##465.2", sql, gProdCategoryId
        
        Grid.TextMatrix(mousRow, gmProdCategoryId) = gProdCategoryId
        quantity = quantity + 1
      
    Else
       If ValueToGuideProdCategoryField("##443", str, "sysname") <> 0 Then GoTo EN1
    End If
  ElseIf mousCol = gmValue Then
       If ValueToGuideProdCategoryField("##443", str, "nameRu") <> 0 Then GoTo EN1
  End If
  
  Grid.TextMatrix(mousRow, mousCol) = str
  GoTo EN1
ElseIf KeyCode = vbKeyEscape Then
CNC:
 If mousCol = gmProdCategory And frmMode = "sourceAdd" Then
    If quantity > 0 Then
        Grid.RemoveItem quantity + 1 ' ту, которую зря добавили
    End If
 End If
EN1:
 frmMode = ""
 lbHide
End If

End Sub

Function ValueToGuideProdCategoryField(myErrCod As String, value As String, _
field As String, Optional passErr As Integer = -11111) As Integer
'Dim i As Integer

ValueToGuideProdCategoryField = False
sql = "UPDATE GuideProdCategory SET [" & field & "] = '" & value & "' WHERE ProdCategoryId =" & gProdCategoryId

ValueToGuideProdCategoryField = myExecute(myErrCod, sql, passErr)
End Function


