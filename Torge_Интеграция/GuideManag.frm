VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form GuideManag 
   BackColor       =   &H8000000A&
   Caption         =   "Справочник менеджеров"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   1740
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6750
   Begin VB.CommandButton cmRepl 
      Caption         =   "--->"
      Enabled         =   0   'False
      Height          =   290
      Left            =   3180
      TabIndex        =   6
      Top             =   4680
      Width           =   375
   End
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
      _ExtentX        =   11456
      _ExtentY        =   7435
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label laDestM 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Фиксировано один
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   4680
      Width           =   315
   End
   Begin VB.Label laSourM 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Фиксировано один
      Height          =   255
      Left            =   2820
      TabIndex        =   5
      Top             =   4680
      Width           =   315
   End
End
Attribute VB_Name = "GuideManag"
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
Dim gManagId As String
Dim sourId As Integer, destId As Integer

Const gmManagId = 0 ' скрытый
Const gmManag = 1
Const gmForSort = 2
Const gmNote = 3


Private Sub Command3_Click()

End Sub

Private Sub cmAdd_Click()
frmMode = "sourceAdd"
If quantity > 0 Then Grid.AddItem ("")
Grid.row = Grid.Rows - 1
mousRow = Grid.Rows - 1
Grid.col = gmManag
mousCol = gmManag
On Error Resume Next
Grid.SetFocus
textBoxInGridCell tbMobile, Grid

End Sub

Private Sub cmDel_Click()
Dim i As Integer
sql = "DELETE  From GuideManag WHERE (((ManagId)=" & gManagId & "));"
i = myExecute("##440", sql, -198)
If i = 0 Then
    quantity = quantity - 1
    If quantity > 0 Then
        Grid.RemoveItem mousRow
    Else
        clearGridRow Grid, 1
    End If
    laSourM.Caption = ""
    laDestM.Caption = ""
ElseIf i = -2 Then
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

Private Sub cmRepl_Click()
If MsgBox("По кнопке <Да> все ссылки на менеджера '" & laSourM & _
"' в Реестрах заказов и в Справочниках по фирмам будут заменены на менеджера '" & _
laDestM & "' !", vbYesNo Or vbDefaultButton2, "Вы уверены?") = vbNo Then Exit Sub

wrkDefault.BeginTrans

sql = "UPDATE BayOrders SET ManagId = " & destId & _
" WHERE (((ManagId)=" & sourId & "));"
If myExecute("##445", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE Orders SET ManagId = " & destId & _
" WHERE (((ManagId)=" & sourId & "));"
If myExecute("##446", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE BayOrders SET lastManagId = " & destId & _
" WHERE (((lastManagId)=" & sourId & "));"
If myExecute("##447", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE Orders SET lastManagId = " & destId & _
" WHERE (((lastManagId)=" & sourId & "));"
If myExecute("##448", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE BayGuideFirms SET ManagId = " & destId & _
" WHERE (((ManagId)=" & sourId & "));"
If myExecute("##449", sql, 0) > 0 Then GoTo ER1

sql = "UPDATE GuideFirms SET ManagId = " & destId & _
" WHERE (((ManagId)=" & sourId & "));"
If myExecute("##450", sql, 0) <= 0 Then
    wrkDefault.CommitTrans
    MsgBox "Замена прошла успешно!", , ""
Else
ER1: wrkDefault.Rollback
    MsgBox "Замена НЕ прошла!", , ""
End If
Grid.SetFocus
End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width

Grid.FormatString = "|<Буква|<Порядок|<Примечание"
Grid.ColWidth(gmManagId) = 0
Grid.ColWidth(gmManag) = 585
Grid.ColWidth(gmForSort) = 1005
Grid.ColWidth(gmNote) = 4545
sql = "SELECT ManagId, Manag, forSort, Note From GuideManag " & _
"Where (((ManagId) > 0 AND (ManagId) <> 14 )) ORDER BY GuideManag.forSort;"
Set tbGuide = myOpenRecordSet("##441", sql, dbOpenForwardOnly)
If tbGuide Is Nothing Then Exit Sub

quantity = 0
While Not tbGuide.EOF
    quantity = quantity + 1
    Grid.TextMatrix(quantity, gmManagId) = tbGuide!ManagId
    Grid.TextMatrix(quantity, gmManag) = tbGuide!Manag
    Grid.TextMatrix(quantity, gmForSort) = tbGuide!ForSort
    Grid.TextMatrix(quantity, gmNote) = tbGuide!note
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
Static sourDest As Boolean

mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If mousRow = 0 Then Exit Sub

If mousCol = gmManag Then
  If laSourM.Caption = "" Then sourDest = False 'сразу после загрузки
  
  If sourDest Then
    laDestM.Caption = Grid.Text
    destId = Grid.TextMatrix(mousRow, gmManagId)
    cmRepl.Enabled = True
  Else
    laSourM.Caption = Grid.Text
    sourId = Grid.TextMatrix(mousRow, gmManagId)
  End If
  sourDest = Not sourDest
End If
End Sub

Private Sub Grid_DblClick()
If mousRow = 0 Then Exit Sub
If Grid.CellBackColor = &H88FF88 Then
        If gManagId = 34 Or gManagId = 40 Then 'Инвентаризация и Коррекция
            MsgBox "Это не поставщик, а системная статья прихода. " & _
            "Изменение названия не изменит ее суть!", , "Внимание!"
        End If
        textBoxInGridCell tbMobile, Grid

End If

End Sub

Private Sub Grid_EnterCell()
 If mousCol = gmManag Then
    tbMobile.MaxLength = 1
 ElseIf mousCol = gmForSort Then
    tbMobile.MaxLength = 6
 Else 'gmNote
    tbMobile.MaxLength = 50
 End If

If quantity > 0 Then
 mousRow = Grid.row
 mousCol = Grid.col
 gManagId = Grid.TextMatrix(mousRow, gmManagId)
 

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

Private Sub tbMobile_Change()
If mousCol = gmForSort And LCase(tbMobile.Text) = "u" Then
    tbMobile.Text = "unUsed"
End If
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, i As Integer

If KeyCode = vbKeyReturn Then
  str = Trim(tbMobile.Text)
  If str = "" Then
    MsgBox "Недпустимое значение", , "Предупрежение"
    Exit Sub
  End If
  If mousCol = gmManag Then
    If frmMode = "sourceAdd" Then
      
      wrkDefault.BeginTrans 'lock01
      sql = "UPDATE GuideManag SET ManagId = ManagId WHERE ManagId=0" 'lock02
      myBase.Execute (sql) 'lock03
    
      sql = "SELECT ManagId, Manag FROM GuideManag ORDER BY ManagId"
      Set tbGuide = myOpenRecordSet("##442", sql, dbOpenDynaset)
'      If tbGuide Is Nothing Then Exit Sub
'      tbGuide.Index = "Key"
      i = 0
      While Not tbGuide.EOF ' сначала исп-ем удаленные номера
        If tbGuide!ManagId > i Then GoTo AA
        tbGuide.MoveNext
        i = i + 1
      Wend
      tbGuide.Close
      If i > 255 Then msgOfEnd "##451", "переполнение GuideManag"
      
AA:   gManagId = i
'      tbGuide.AddNew
'      tbGuide!ManagId = gManagId
'      tbGuide!Manag = str
'      On Error GoTo ERR1
'      tbGuide.Update
'      tbGuide.Close

      sql = "INSERT INTO GuideManag (ManagId, Manag) VALUES (" & _
      gManagId & ", '" & str & "')"
      If myExecute("##465", sql) <> 0 Then GoTo EN1
      wrkDefault.CommitTrans
      
      Grid.TextMatrix(mousRow, gmManagId) = gManagId
      quantity = quantity + 1
    Else
       If ValueToGuideManagField("##443", str, "Manag") <> 0 Then GoTo EN1
'       If i = 3022 Then
'            'existMsg
'            Exit Sub
'       ElseIf i <> 0 Then
'            GoTo EN1
'       End If
    End If
  ElseIf mousCol = gmForSort Then
       If ValueToGuideManagField("##443", str, "ForSort") <> 0 Then GoTo EN1
  ElseIf mousCol = gmNote Then
       If ValueToGuideManagField("##443", str, "Note") <> 0 Then GoTo EN1
  End If
  
  Grid.TextMatrix(mousRow, mousCol) = str
  GoTo EN1
ElseIf KeyCode = vbKeyEscape Then
CNC:
 If mousCol = gmManag And frmMode = "sourceAdd" Then
    If quantity > 0 Then
                Grid.RemoveItem quantity + 1 ' ту, которую зря добавили
    End If
 End If
EN1:
 frmMode = ""
 lbHide
End If
'Exit Sub'

'ERR1:
'tbGuide.Close
'errorCodAndMsg "##444"
'If Err = 3022 Then
'    MsgBox "Это название уже есть", , "Ошибка!"
'    MsgBox "Это название уже есть.", , "Ошибка"'
'    GoTo CNC
'Else
'    MsgBox Error, , "Ошибка 444-" & Err & ":  " '##444
'    End
'End If

End Sub

Function ValueToGuideManagField(myErrCod As String, value As String, _
field As String, Optional passErr As Integer = -11111) As Integer
'Dim i As Integer

ValueToGuideManagField = False
sql = "UPDATE GuideManag SET [" & field & _
"] = '" & value & "' WHERE (((ManagId)=" & gManagId & "));"
'MsgBox "sql = " & sql

ValueToGuideManagField = myExecute(myErrCod, sql, passErr)
End Function


