VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form jGuideSchets 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Справочник счетов"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmLoad 
      Caption         =   "Обновить"
      Height          =   315
      Left            =   300
      TabIndex        =   5
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   540
      TabIndex        =   4
      Text            =   "tbMobile"
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   9780
      TabIndex        =   3
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3180
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   4560
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   7435
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "jGuideSchets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim quantity As Long
Dim mousCol As Long, mousRow As Long

Const gsNumber = 1
Const gsNote = 2
Const gsSubNumber = 3
Const gsSubNote = 4
Const gsBegDebit = 5
Const gsBegKredit = 6

Private Sub cmAdd_Click()
If quantity > 0 Then Grid.AddItem ("")
Grid.row = Grid.Rows - 1
mousRow = Grid.Rows - 1
Grid.col = gsNumber
mousCol = gsNumber

textBoxInGridCell tbMobile, Grid

End Sub

Private Sub cmDel_Click()
Dim I As String, J As String

I = Grid.TextMatrix(mousRow, gsNumber)
If Grid.TextMatrix(mousRow, gsSubNumber) = "" Then
    J = "00"
Else
    J = Grid.TextMatrix(mousRow, gsSubNumber)
End If
sql = "DELETE yGuideSchets WHERE (((number)='" & I & _
"') AND ((subNumber)='" & J & "'));"

I = myExecute("##475", sql, -198)
If I = 0 Then
    quantity = quantity - 1
    If quantity > 0 Then
        Grid.RemoveItem mousRow
    Else
        clearGridRow Grid, 1
        cmDel.Enabled = False
    End If
ElseIf I = -2 Then
    MsgBox "Этот счет используется.", , _
    "Удаление невозможно!"
End If

Grid_EnterCell
On Error Resume Next
Grid.SetFocus

End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmLoad_Click()
loadGuide
Grid_EnterCell
On Error Resume Next
Grid.SetFocus

End Sub

Private Sub Form_Load()
Grid.FormatString = "|Счет|<Наименование cчета|Субсчет|<Наименование cубcчета|Дебит.вступ. остатки|Кредит.вступ. остатки"
Grid.ColWidth(0) = 0
Grid.ColWidth(gsNumber) = 450
Grid.ColWidth(gsNote) = 3750
Grid.ColWidth(gsSubNumber) = 450
Grid.ColWidth(gsSubNote) = 3750
Grid.ColWidth(gsBegDebit) = 800
Grid.ColWidth(gsBegKredit) = 840

loadGuide

End Sub

Sub loadGuide()

sql = "SELECT * FROM yGuideSchets ORDER BY number,subNumber"
Set tbGuide = myOpenRecordSet("##329", sql, dbOpenForwardOnly)
If tbGuide Is Nothing Then Exit Sub
Me.MousePointer = flexHourglass
quantity = 0
clearGrid Grid
'tbGuide.Index = "Key"
If Not tbGuide.BOF Then
 While Not tbGuide.EOF
   If tbGuide!number <> "255" And tbGuide!number <> "" Then
    quantity = quantity + 1
    Grid.TextMatrix(quantity, gsNumber) = Journal.schType(tbGuide!number)
    Grid.TextMatrix(quantity, gsNote) = tbGuide!Note
    If tbGuide!subNumber <> 0 Then _
        Grid.TextMatrix(quantity, gsSubNumber) = Journal.schType(tbGuide!subNumber)
    Grid.TextMatrix(quantity, gsSubNote) = tbGuide!subNote
    Grid.TextMatrix(quantity, gsBegDebit) = tbGuide!begDebit
    Grid.TextMatrix(quantity, gsBegKredit) = tbGuide!begKredit
    Grid.AddItem ""
   End If
   tbGuide.MoveNext
 Wend
End If
tbGuide.Close
Grid.Visible = True

If quantity > 0 Then
    Grid.RemoveItem quantity + 1
    Grid.row = quantity
    Grid.col = 1
    cmDel.Enabled = True
End If
Me.MousePointer = flexDefault

End Sub

Private Sub Form_Resize()
On Error Resume Next
Grid_EnterCell
Grid.SetFocus

End Sub

Sub lbHide()
tbMobile.Visible = False
Grid.Enabled = True
On Error Resume Next
Grid.SetFocus
Grid_EnterCell
End Sub

Private Sub Grid_DblClick()
If Grid.CellBackColor = &H88FF88 Then textBoxInGridCell tbMobile, Grid

End Sub

Private Sub Grid_EnterCell()
If quantity = 0 Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col



If mousCol = gsNumber Or mousCol = gsSubNumber Then
'If mousCol = gsNote Or mousCol = gsSubNote Then
   Grid.CellBackColor = vbYellow
Else
   Grid.CellBackColor = &H88FF88
End If
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Grid_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor

End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)

End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim I As String, J As String, k As Integer

If KeyCode = vbKeyReturn Then

  If mousCol = gsNumber Then
    'If Not IsNumeric(tbMobile.Text) Then GoTo ER1
    I = Len(tbMobile.Text)
    If I >= 4 Then
        I = Left$(tbMobile.Text, 2)
        J = Mid$(tbMobile.Text, 4)
        GoTo AA
    ElseIf I = 2 Then
        I = tbMobile.Text
        J = "00"
AA: '    Set tbGuide = myOpenRecordSet("##355", "yGuideSchets", dbOpenTable)
'        If tbGuide Is Nothing Then Exit Sub
'        tbGuide.Index = "Key"
'        tbGuide.Seek "=", i, j
'        If tbGuide.NoMatch Then
'            tbGuide.AddNew
'            tbGuide!number = i
'            tbGuide!subNumber = j
'            tbGuide.Update
'            tbGuide.Close
            
            sql = "INSERT INTO yGuideSchets (number, subNumber) " & _
            "VALUES ('" & I & "', '" & J & "')"
            k = myExecute("##355", sql, -193)
            If k = -2 Then
                MsgBox "Счет " & tbMobile.Text & " уже есть!", , "Добавление невозможно"
                GoTo EN1
            ElseIf k <> 0 Then
                GoTo EN2
            End If
            Grid.TextMatrix(mousRow, gsNumber) = Journal.schType(I)
            Grid.TextMatrix(mousRow, gsSubNumber) = Journal.schType(J)
            quantity = quantity + 1 '$$4
            cmDel.Enabled = True    '
            GoTo EN2
'        Else
'            tbGuide.Close
'            MsgBox "Счет " & tbMobile.Text & " уже есть!", , "Добавление невозможно"
'            GoTo EN1
'        End If
    Else
ER1:    MsgBox "Введите номер счета и субсчета: первые две - номер счета, далее через пробел - " & _
        "номер субсчета. Пример 99 или 19 НДС. Номера счета и субсчета могут быть текстовой строкой", _
        , "Недопустимый формат счета!"
EN1:    tbMobile.SelStart = 0
        tbMobile.SelLength = I
        tbMobile.SetFocus
        Exit Sub
    End If
  ElseIf mousCol = gsNote Then
'        sql = "UPDATE yGuideSchets SET [note] = '" & tbMobile.Text & _
        "' WHERE (((number)=" & Grid.TextMatrix(mousRow, gsNumber) & "));"
    If Not ValueToGuideSchetField("##332", "'" & tbMobile.Text & "'", "note") Then GoTo EN2
  ElseIf mousCol = gsSubNote Then
'        sql = "UPDATE yGuideSchets SET subNote = '" & tbMobile.Text & _
        "' WHERE (((number)=" & i & ") AND ((subNumber)=" & j & "));"
    If Not ValueToGuideSchetField("##332", "'" & tbMobile.Text & "'", "subNote") Then GoTo EN2
  ElseIf mousCol = gsBegDebit Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not ValueToGuideSchetField("##332", tbMobile.Text, "begDebit") Then GoTo EN2
  ElseIf mousCol = gsBegKredit Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If Not ValueToGuideSchetField("##332", tbMobile.Text, "begKredit") Then GoTo EN2
  End If
  Grid.TextMatrix(mousRow, mousCol) = tbMobile.Text

  GoTo EN2
ElseIf KeyCode = vbKeyEscape Then
CNC:
 If mousCol = gsNumber Then
    If quantity > 0 Then
        Grid.RemoveItem quantity + 1 ' ту, которую зря добавили
    End If
 End If

EN2: lbHide
End If
End Sub

Function ValueToGuideSchetField(myErrCod As String, Value As String, _
field As String) As Boolean
Dim I As String, J As String
        
ValueToGuideSchetField = False
I = Grid.TextMatrix(mousRow, gsNumber)
If I = "" Then I = "00"
J = Grid.TextMatrix(mousRow, gsSubNumber)
If J = "" Then J = "00"

sql = "UPDATE yGuideSchets SET [" & field & "] = " & Value & _
" WHERE (((number)='" & I & "') AND ((subNumber)='" & J & "'));"

Debug.Print sql
If myExecute(myErrCod, sql) = 0 Then ValueToGuideSchetField = True

End Function


