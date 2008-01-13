VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GuideShiz 
   Caption         =   "Справочник шифров затрат"
   ClientHeight    =   3348
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5964
   LinkTopic       =   "Form1"
   ScaleHeight     =   3348
   ScaleWidth      =   5964
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lbActive 
      Height          =   624
      ItemData        =   "GuideShiz.frx":0000
      Left            =   2760
      List            =   "GuideShiz.frx":000F
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Text            =   "tbMobile"
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   4920
      TabIndex        =   2
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Height          =   315
      Left            =   2460
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   2880
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2775
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5835
      _ExtentX        =   10287
      _ExtentY        =   4890
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "GuideShiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isLoad As Boolean

Dim mousRow As Long    '
Dim mousCol As Long    '
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim quantity As Integer 'количество найденных фирм

Const shShizId = 0 ' скрытый
Const shText = 1
Const shMainCosts = 2

Private Sub cmAdd_Click()
    If quantity > 0 Then Grid.AddItem ("")
    Grid.row = Grid.Rows - 1
    mousRow = Grid.Rows - 1
    Grid.col = shText
    mousCol = shText
    On Error Resume Next
    Grid.SetFocus
    textBoxInGridCell tbMobile, Grid
End Sub

Private Sub cmExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width

Grid.FormatString = "|<Шифр|Основные?"
Grid.ColWidth(shShizId) = 0
Grid.ColWidth(shText) = 2585
Grid.ColWidth(shMainCosts) = 1005
sql = "SELECT id, nm, is_main_costs From Shiz " & _
"Where id > 0 ORDER BY nm"
Set tbGuide = myOpenRecordSet("##441", sql, dbOpenForwardOnly)
If tbGuide Is Nothing Then Exit Sub

quantity = 0
While Not tbGuide.EOF
    quantity = quantity + 1
    Grid.TextMatrix(quantity, shShizId) = tbGuide!id
    Grid.TextMatrix(quantity, shText) = tbGuide!nm
    If IsNull(tbGuide!is_main_costs) Then
        Grid.TextMatrix(quantity, shMainCosts) = ""
    ElseIf tbGuide!is_main_costs = 0 Then
        Grid.TextMatrix(quantity, shMainCosts) = "Нет"
    Else
        Grid.TextMatrix(quantity, shMainCosts) = "Да"
    End If
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
    'Grid_EnterCell
End If

isLoad = True
End Sub

Private Sub Grid_Click()
    mousCol = Grid.MouseCol
    mousRow = Grid.MouseRow
    If quantity = 0 Then Exit Sub
    If mousRow = 0 Then
        Grid.CellBackColor = Grid.BackColor
        Grid.row = 1    ' только чтобы снять выделение
        Grid_EnterCell
    End If
End Sub

Private Sub Grid_DblClick()
    If Grid.col = shMainCosts Then
        listBoxInGridCell lbActive, Grid, Grid.TextMatrix(Grid.MouseRow, Grid.MouseCol)
    ElseIf Grid.col = shText Then
        textBoxInGridCell tbMobile, Grid
    End If

End Sub

Private Sub Grid_EnterCell()
    If quantity > 0 Then
        mousRow = Grid.row
        mousCol = Grid.col
        'gManagId = Grid.TextMatrix(mousRow, gmManagId)
        
        
        If mousCol > 0 Then
           Grid.CellBackColor = &H88FF88
        Else
           Grid.CellBackColor = vbYellow
        End If
    End If


End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Grid_DblClick

End Sub

Private Sub Grid_LeaveCell()
    Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub lbActive_DblClick()
Dim success As Integer
    If lbActive.Visible = False Then Exit Sub
    sql = "update shiz set is_main_costs = "
    If lbActive.Text = "Да" Then
        sql = sql & "1"
    ElseIf lbActive.Text = "Нет" Then
        sql = sql & "0"
    Else
        sql = sql & "null"
    End If
    sql = sql & " where id = " & Grid.TextMatrix(mousRow, shShizId)
    myExecute "##shiz_update", sql
    Grid.Text = lbActive.Text
        
        
    
    lbHide
End Sub

Private Sub lbActive_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lbActive_DblClick
    If KeyCode = vbKeyEscape Then lbHide
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim new_id As Integer

    If KeyCode = vbKeyReturn Then
        If Grid.TextMatrix(mousRow, shShizId) = "" Then
            'add new shiz
            sql = "select wf_add_shiz ('" & tbMobile.Text & "') as new_id"
            byErrSqlGetValues "##insert shiz", sql, new_id
            If new_id > 0 Then
                'quantity = quantity + 1
                Grid.TextMatrix(quantity + 1, shShizId) = new_id
                Grid.TextMatrix(quantity + 1, shText) = tbMobile.Text
            ElseIf new_id = -1 Then
                MsgBox "Некорректное название шифра затрат. Такое название уже есть или оно пустое.", vbOKOnly, "Ошибка ввода"
                Grid.RemoveItem (quantity + 1)
            End If
        Else
            sql = "select id from shiz where nm = '" & tbMobile.Text & "' and id != " & Grid.TextMatrix(mousRow, shShizId)
            byErrSqlGetValues "W#insert shiz", sql, new_id
            If new_id <> 0 Or tbMobile.Text = "" Then
                MsgBox "Некорректное название шифра затрат. Такое название уже есть или оно пустое.", vbOKOnly, "Ошибка ввода"
            Else
                sql = "update shiz set nm = '" & tbMobile.Text _
                    & "' where id = " & Grid.TextMatrix(mousRow, shShizId)
                myExecute "##update shiz", sql
                Grid.Text = tbMobile.Text
            End If
        End If
        lbHide
    ElseIf KeyCode = vbKeyEscape Then
        Grid.RemoveItem (quantity + 1)
        lbHide
    End If
End Sub

Sub lbHide()
    tbMobile.Visible = False
    Grid.Enabled = True
    On Error Resume Next
    Grid.SetFocus
    Grid_EnterCell
    lbActive.Visible = False
End Sub

