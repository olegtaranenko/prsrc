VERSION 5.00
Begin VB.Form jGuidePurpDet 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Справочники  Назначений"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   6780
      TabIndex        =   0
      Top             =   5340
      Width           =   915
   End
   Begin VB.CommandButton cmSel 
      Caption         =   "Выбрать"
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   5340
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox tbPurpose 
      Height          =   285
      Left            =   150
      TabIndex        =   3
      Text            =   " "
      Top             =   4920
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.ListBox lbPurpose 
      Height          =   4935
      ItemData        =   "jGuidePurpDet.frx":0000
      Left            =   120
      List            =   "jGuidePurpDet.frx":0002
      TabIndex        =   4
      Top             =   240
      Width           =   3675
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   5340
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   5340
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label laPurpose 
      Caption         =   "Назначения"
      Height          =   195
      Left            =   1380
      TabIndex        =   6
      Top             =   60
      Width           =   1035
   End
End
Attribute VB_Name = "jGuidePurpDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Regim As String

Private Sub cmAdd_Click()
cmAdd.Enabled = False
tbPurpose.Text = ""
tbPurpose.Visible = True
tbPurpose.ZOrder
tbPurpose.SetFocus
End Sub

'Private Sub cmAdd2_Click()
'cmAdd2.Enabled = False
'tbDetail.Text = ""
'tbDetail.Visible = True
'tbDetail.ZOrder
'tbDetail.SetFocus
'End Sub

Private Sub cmDel_Click()
Dim i As Integer

'Set Table = myOpenRecordSet("##453", "yGuidePurpose", dbOpenSnapshot)
'If Table Is Nothing Then Exit Sub
'Table.FindFirst "pDescript='" & lbPurpose.Text & "'"
'If Not Table.NoMatch Then
'    Table.Close
'    MsgBox "Значение '" & lbPurpose.Text & "' используется в Справочнике " & _
'    "операций.", , "Удаление невозможно!"
'    Exit Sub
'End If
'Table.Close


sql = "DELETE  From yGuidePurp " & _
"WHERE (((descript)='" & lbPurpose.Text & "'));"
i = myExecute("##454", sql, -198)
If i = 0 Then
    loadLbFromGuide lbPurpose, "yGuidePurp"
    If lbPurpose.ListCount > 0 Then
'        lbPurpose.SetFocus
        lbPurpose.ListIndex = 0
    Else
        cmDel.Enabled = False
    End If
ElseIf i = -2 Then
    MsgBox "Это название используется в справочнике операций. " & _
    "Удаление невозможно! ", , "Error-454"
End If
lbPurpose.SetFocus
End Sub


Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmSel_Click()
Dim i As Integer

If Regim = "purpose" Then
  tmpStr = lbPurpose.Text
'Else
'  tmpStr = lbDetail.Text
End If
Unload Me
End Sub

Private Sub Form_Load()
If Regim = "purpose" Then
    loadLbFromGuide lbPurpose, "yGuidePurp"
'    lbDetail.Visible = False
    GoTo AA
'ElseIf Regim = "detail" Then
'    loadLbFromGuide lbDetail, "yGuideDet"
'    lbPurpose.Visible = False
'    laPurpose.Visible = False
'    laDetail.Left = laPurpose.Left
AA: Me.Caption = "Выбор из справочника"
'    Me.Width = Me.Width - lbDetail.Left + lbPurpose.Left
'    cmExit.Left = cmExit.Left - lbDetail.Left + lbPurpose.Left
    cmSel.Visible = True
'    lbDetail.Left = lbPurpose.Left
Else
    loadLbFromGuide lbPurpose, "yGuidePurp"
'    loadLbFromGuide lbDetail, "yGuideDet"
    lbPurpose.Visible = True
'    lbDetail.Visible = True
    cmAdd.Visible = True
    cmDel.Visible = True
'    cmAdd2.Visible = True
'    cmDel2.Visible = True
End If
End Sub

Sub loadLbFromGuide(lb As ListBox, tableName As String, Optional seeek As String = "")
Dim i As Integer

strWhere = " WHERE descript <> ''"
If Regim <> "" Then strWhere = ""

sql = "SELECT * FROM " & tableName & strWhere & " ORDER BY descript"
Set Table = myOpenRecordSet("##452", sql, dbOpenForwardOnly)
'Debug.Print sql
'If Table Is Nothing Then myBase.Close: End
'Table.Index = "Key"
lb.Clear
While Not Table.EOF
'  If Trim(Table.fields(0)) = "" Then
    lb.AddItem Table.fields(0)
    If seeek = Table.fields(0) Then lb.ListIndex = lb.ListCount - 1
'  End If
    Table.MoveNext
Wend
Table.Close

End Sub


'Private Sub lbDetail_Click()
'cmDel2.Enabled = True
'End Sub

Private Sub lbPurpose_Click()
cmDel.Enabled = True
End Sub

'Private Sub tbDetail_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim i As Integer
'If KeyCode = vbKeyReturn Then
'  tbDetail.Text = Trim(tbDetail.Text)
'  If tbDetail.Text = "" Then
'    MsgBox "Недопустимое значение", , ""
'    Exit Sub
'  End If
'  cmAdd2.Enabled = True
'  sql = "INSERT INTO yGuideDet (descript) VALUES ('" & tbDetail.Text & "')"
'  i = myExecute("##334", sql, -193)
'  If i = -2 Then
'
''  Set tbGuide = myOpenRecordSet("##477", "yGuideDet", dbOpenTable) ' dbOpenTable,dbOpenForwardOnly
''  If tbGuide Is Nothing Then Exit Sub
''  tbGuide.Index = "Key"
''  tbGuide.Seek "=", tbDetail.Text
''  If tbGuide.NoMatch Then
''    tbGuide.AddNew
''    tbGuide!descript = tbDetail.Text
''    tbGuide.Update
''  Else
'    MsgBox "Уточнение '" & tbDetail.Text & "' уже есть!", , "Предупреждение"
'  End If
'  loadLbFromGuide lbDetail, "yGuideDet", tbDetail.Text
''  tbGuide.Close
'
'  tbDetail.Visible = False
'ElseIf KeyCode = vbKeyEscape Then
'    cmAdd2.Enabled = True
'    tbDetail.Visible = False
'End If
'
'End Sub

Private Sub tbPurpose_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If KeyCode = vbKeyReturn Then
  tbPurpose.Text = Trim(tbPurpose.Text)
  If tbPurpose.Text = "" Then
    MsgBox "Недопустимое значение", , ""
    Exit Sub
  End If
  cmAdd.Enabled = True
  sql = "INSERT INTO yGuidePurp (descript) VALUES ('" & tbPurpose.Text & "')"
  i = myExecute("##334", sql, -193)
  If i = -2 Then
'  Set tbGuide = myOpenRecordSet("##334", "yGuidePurp", dbOpenTable) ' dbOpenTable,dbOpenForwardOnly
'  If tbGuide Is Nothing Then Exit Sub
'  tbGuide.Index = "Key"
'  tbGuide.Seek "=", tbPurpose.Text
'  If tbGuide.NoMatch Then
'    tbGuide.AddNew
'    tbGuide!descript = tbPurpose.Text
'    tbGuide.Update
'  Else
    MsgBox "Назначение '" & tbPurpose.Text & "' уже есть!", , "Предупреждение"
    tbPurpose.SetFocus
    Exit Sub
  ElseIf i <> 0 Then
    GoTo AA
  End If
  loadLbFromGuide lbPurpose, "yGuidePurp", tbPurpose.Text
'  tbGuide.Close
  
  tbPurpose.Visible = False
ElseIf KeyCode = vbKeyEscape Then
AA: cmAdd.Enabled = True
    tbPurpose.Visible = False
End If

End Sub
