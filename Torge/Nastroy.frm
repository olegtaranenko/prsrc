VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Nastroy 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройки"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmExit 
      Cancel          =   -1  'True
      Caption         =   "Выход"
      Height          =   315
      Left            =   7920
      TabIndex        =   2
      Top             =   1800
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1826
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label laTitle 
      Alignment       =   2  'Центровка
      Caption         =   "Label1"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8715
   End
End
Attribute VB_Name = "Nastroy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mousCol As Long, mousRow As Long

Const jnField = 1
Const jnDebit = 2
Const jnSubDebit = 3
Const jnKredit = 4
Const jnSubKredit = 5
Const jnPurpose = 6
'Const jnDetail = 7

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
laTitle.Caption = "Параметры записи полей из отчета 'Реализация'  " & _
" в журнал Х.О." & vbCrLf & "(для их изменения нажмите " & _
"<Enter> на зеленой подсветке)"

Grid.FormatString = "|<Поле из отчета|Дб|Сс|Кр|Сс|<Назначение"
Grid.ColWidth(0) = 0
Grid.ColWidth(jnField) = 2130
Grid.ColWidth(jnPurpose) = 3435
''Grid.ColWidth(jnDetail) = 1935

Grid.TextMatrix(1, jnField) = "Реализация"
Grid.AddItem Chr(9) & "Материалы под заказы"
Grid.AddItem Chr(9) & "Остальные материалы"


paramLoad

Grid_EnterCell

End Sub

Sub paramLoad()
Dim I As Integer
sql = "SELECT yGuidePurpose.pDescript, yGuidePurpose.Debit, " & _
"yGuidePurpose.Kredit, yGuidePurpose.subDebit, yGuidePurpose.subKredit, " & _
"yGuidePurpose.auto " & _
"FROM yGuidePurpose " & _
"WHERE (((yGuidePurpose.auto)<>''));"
'Debug.Print sql
    
Set tbDocs = myOpenRecordSet("##395", sql, dbOpenForwardOnly)
While Not tbDocs.EOF
  If tbDocs!AUTO = "r" Then
    I = 1: GoTo AA:
  ElseIf tbDocs!AUTO = "z" Then
    I = 2: GoTo AA:
  ElseIf tbDocs!AUTO = "m" Then
    I = 3
AA: Grid.TextMatrix(I, jnDebit) = Journal.schType(tbDocs!debit, 255)
    Grid.TextMatrix(I, jnSubDebit) = Journal.schType(tbDocs!subDebit)
    Grid.TextMatrix(I, jnKredit) = Journal.schType(tbDocs!kredit, 255)
    Grid.TextMatrix(I, jnSubKredit) = Journal.schType(tbDocs!subKredit)
    Grid.TextMatrix(I, jnPurpose) = tbDocs!pDescript
'    Grid.TextMatrix(i, jnDetail) = tbDocs!descript
  End If
  tbDocs.MoveNext
Wend
tbDocs.Close




End Sub


Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow

End Sub

Private Sub Grid_DblClick()
If Grid.CellBackColor <> &H88FF88 Then Exit Sub

If mousRow = 1 Then
    jGuidePurpose.Regim = "selectNr": GoTo AA
ElseIf mousRow = 2 Then
    jGuidePurpose.Regim = "selectNz": GoTo AA
ElseIf mousRow = 3 Then
    jGuidePurpose.Regim = "selectNm"
AA:
    Journal.getSchetsFromGrid Grid, jnDebit
    ReDim QQ(4): QQ(1) = debit ' т.к. затираются в jGuidePurpose
    QQ(2) = subDebit: QQ(3) = kredit: QQ(4) = subKredit
    
    jGuidePurpose.purpose = Grid.TextMatrix(mousRow, jnPurpose)
'    jGuidePurpose.detail = Grid.TextMatrix(mousRow, jnDetail)
    purposeId = getPurposeIdByDescript(jGuidePurpose.purpose)
'    detailId = getDetailIdByDescript(jGuidePurpose.purposeId,detail)
    jGuidePurpose.Show vbModal
End If

End Sub

Private Sub Grid_EnterCell()
mousRow = Grid.row
mousCol = Grid.col

If mousCol > 1 Then  ' чтобы м.б. копировать из jnFirm
   Grid.CellBackColor = &H88FF88
Else
   Grid.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid_DblClick

End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor

End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)

End Sub
