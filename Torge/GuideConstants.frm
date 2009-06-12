VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GuideConstants 
   BackColor       =   &H8000000A&
   Caption         =   "��������� ��� ������"
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
      Caption         =   "�����"
      Height          =   315
      Left            =   5700
      TabIndex        =   3
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "�������"
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "��������"
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
Attribute VB_Name = "GuideConstants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isLoad As Boolean
Public mousRow As Long    '
Public mousCol As Long    '
Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����
Dim quantity As Integer '���������� ��������� ��������
Dim frmMode As String
Dim gConstantsId As String
Dim sourId As Integer, destId As Integer

Const gmConstantsId = 0 ' �������
Const gmConstants = 1
Const gmValue = 2
Const gmNote = 3



Private Sub cmAdd_Click()
frmMode = "sourceAdd"
If quantity > 0 Then Grid.AddItem ("")
Grid.row = Grid.Rows - 1
mousRow = Grid.Rows - 1
Grid.col = gmConstants
mousCol = gmConstants
cmAdd.Enabled = False
cmDel.Enabled = False
On Error Resume Next
Grid.SetFocus
textBoxInGridCell tbMobile, Grid

End Sub

Private Sub cmDel_Click()
Dim I As Integer
sql = "DELETE  From GuideConstants WHERE (((ConstantsId)=" & gConstantsId & "));"
I = myExecute("##440", sql, -198)
If I = 0 Then
    quantity = quantity - 1
    If quantity > 0 Then
        Grid.RemoveItem mousRow
    Else
        clearGridRow Grid, 1
    End If
ElseIf I = -2 Then
    MsgBox "� ����� ��������� ���� ������ ���� �� ������������ � ������������ " & _
    "����.", , "�������� ����������!"
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

Grid.FormatString = "|<��������|<��������|<����������"
Grid.colWidth(gmConstantsId) = 0
Grid.colWidth(gmConstants) = 585
Grid.colWidth(gmValue) = 1005
Grid.colWidth(gmNote) = 4545
sql = "SELECT ConstantsId, Constants, Value, Note From GuideConstants "
Set tbGuide = myOpenRecordSet("##441", sql, dbOpenForwardOnly)
If tbGuide Is Nothing Then Exit Sub

quantity = 0
While Not tbGuide.EOF
    quantity = quantity + 1
    Grid.TextMatrix(quantity, gmConstantsId) = tbGuide!ConstantsId
    Grid.TextMatrix(quantity, gmConstants) = tbGuide!Constants
    Grid.TextMatrix(quantity, gmValue) = tbGuide!value
    If Not IsNull(tbGuide!note) Then Grid.TextMatrix(quantity, gmNote) = tbGuide!note
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
 gConstantsId = Grid.TextMatrix(mousRow, gmConstantsId)
 

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

Private Function validateConstant(name As String, Optional ByVal value As String = "0") As Boolean
Dim initStr As String
    validateConstant = True
    If Not IsNumeric(value) Then
        MsgBox "�������� �������� ���������", vbOKOnly Or vbCritical, "��������� ����"
        GoTo er
    Else
        value = CStr(CDbl(value))
    End If
    
    On Error GoTo invalid
    
    initStr = name & "=" & value
    sc.ExecuteStatement (initStr)
    Exit Function
invalid:
    MsgBox "�������� �������� ��� ��������� ���������", vbOKOnly Or vbCritical, "��������� ����"
er:
    validateConstant = False
End Function

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, I As Integer
Dim initStr As String

If KeyCode = vbKeyReturn Then
  str = Trim(tbMobile.Text)
  If str = "" Then
    MsgBox "����������� ��������", , "�������������"
    Exit Sub
  End If
  If mousCol = gmConstants Then
    If frmMode = "sourceAdd" Then
        If Not validateConstant(str) Then GoTo CNC
        sql = "INSERT INTO GuideConstants (Constants) VALUES ( '" & str & "')"
        If myExecute("##465", sql) <> 0 Then GoTo EN1
        
        sql = "select constantsID from GuideConstants where Constants = '" & str & "'"
        byErrSqlGetValues "##465.2", sql, gConstantsId
        
        Grid.TextMatrix(mousRow, gmConstantsId) = gConstantsId
        quantity = quantity + 1
      
    Else
       If Not validateConstant(str, Grid.TextMatrix(mousRow, gmValue)) Then GoTo CNC
       If ValueToGuideConstantsField("##443", str, "Constants") <> 0 Then GoTo EN1
    End If
  ElseIf mousCol = gmValue Then
    If validateConstant(Grid.TextMatrix(mousRow, gmConstants), str) Then
        If ValueToGuideConstantsField("##443", str, "Value") <> 0 Then
            GoTo EN1
        End If
    Else
        Exit Sub
    End If
  ElseIf mousCol = gmNote Then
       If ValueToGuideConstantsField("##443", str, "Note") <> 0 Then GoTo EN1
  End If
  
  Grid.TextMatrix(mousRow, mousCol) = str
  GoTo EN1
ElseIf KeyCode = vbKeyEscape Then
CNC:
 If mousCol = gmConstants And frmMode = "sourceAdd" Then
    If quantity > 0 Then
        Grid.RemoveItem quantity + 1 ' ��, ������� ��� ��������
    End If
 End If
EN1:
 frmMode = ""
 lbHide
End If

End Sub

Function ValueToGuideConstantsField(myErrCod As String, value As String, _
field As String, Optional passErr As Integer = -11111) As Integer
'Dim i As Integer

ValueToGuideConstantsField = False
sql = "UPDATE GuideConstants SET [" & field & _
"] = '" & value & "' WHERE (((ConstantsId)=" & gConstantsId & "));"
'MsgBox "sql = " & sql

ValueToGuideConstantsField = myExecute(myErrCod, sql, passErr)
End Function


