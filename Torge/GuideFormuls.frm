VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form GuideFormuls 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Справочник формул для прайса"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   10470
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lbPrams 
      Height          =   645
      Index           =   2
      ItemData        =   "GuideFormuls.frx":0000
      Left            =   1980
      List            =   "GuideFormuls.frx":000D
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ListBox lbPrams 
      Height          =   450
      Index           =   1
      ItemData        =   "GuideFormuls.frx":0037
      Left            =   3780
      List            =   "GuideFormuls.frx":0041
      TabIndex        =   8
      Top             =   2100
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmSel 
      Caption         =   "Выбрать"
      Height          =   315
      Left            =   5340
      TabIndex        =   4
      Top             =   5040
      Width           =   975
   End
   Begin VB.ListBox lbPrams 
      Height          =   645
      Index           =   0
      ItemData        =   "GuideFormuls.frx":005C
      Left            =   1980
      List            =   "GuideFormuls.frx":0069
      TabIndex        =   7
      Top             =   1500
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ListBox lbForWho 
      Height          =   645
      ItemData        =   "GuideFormuls.frx":0081
      Left            =   3300
      List            =   "GuideFormuls.frx":008E
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   9420
      TabIndex        =   0
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   840
      TabIndex        =   5
      Text            =   "tbMobile"
      Top             =   1140
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   8281
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "GuideFormuls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isLoad As Boolean
Public mousRow As Long    '
Public mousCol As Long    '
Public Regim As String
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim quantity As Integer 'количество найденных фирм
Dim frmMode As String
'Dim gFormulaId As String
'Dim notValidKey As Boolean '

Const gfNomer = 1
Const gfForWho = 2
Const gfFormula = 3
Const gfNote = 4


Private Sub Command3_Click()

End Sub

Private Sub cmAdd_Click()
frmMode = "rowAdd"
If quantity > 0 Then Grid.AddItem ("")
Grid.row = Grid.Rows - 1
mousRow = Grid.Rows - 1
Grid.col = gfNomer
mousCol = gfNomer
On Error Resume Next
Grid.SetFocus
textBoxInGridCell tbMobile, Grid

End Sub

Private Sub cmDel_Click()
Dim i As Integer, str As String

str = Grid.TextMatrix(mousRow, gfNomer)
If MsgBox("Удалить формулу №" & str & " из Справочника", _
 vbDefaultButton2 Or vbYesNo, "Удалить '" & gNomNom & "'. Вы уверены?") _
  = vbNo Then Exit Sub

sql = "DELETE From sGuideFormuls " & _
"WHERE (((sGuideFormuls.nomer)=" & str & "));"
i = myExecute("##304", sql, -198)
If i = 0 Then
    quantity = quantity - 1
    If quantity > 0 Then Grid.RemoveItem mousRow
ElseIf i = -2 Then
    MsgBox "Эта формула используется в справочнике ном-ры или изделий.", , _
    "Удаление невозможно!"
End If
On Error Resume Next
Grid.SetFocus
End Sub

Private Sub cmExit_Click()
tmpStr = "" ' формула не б.выбрана
Unload Me
End Sub

Private Sub cmSel_Click()

tmpStr = Grid.TextMatrix(mousRow, gfNomer)
Unload Me
End Sub

Private Sub Form_Activate()
If Regim = "" Then
    cmAdd.Visible = True
    cmDel.Visible = True
    cmSel.Visible = False
Else
    cmAdd.Visible = False
    cmDel.Visible = False
    cmSel.Visible = True
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim i As Integer

If tbMobile.Visible And mousCol = gfFormula And KeyAscii > 57 Then
    If KeyAscii = 222 Or KeyAscii = 254 Then 'точка на Рус.
        KeyAscii = 46                   'точка на Лат.
    Else
        KeyAscii = 0
        If Grid.TextMatrix(mousRow, gfForWho) = lbForWho.List(0) Then
            i = 0
'            listBoxInGridCell lbPrams(0), Grid
'            lbPrams(0).Top = lbPrams(0).Top + Grid.CellHeight
        ElseIf Grid.TextMatrix(mousRow, gfForWho) = lbForWho.List(2) Then
            i = 1
'            listBoxInGridCell lbPrams(1), Grid
'            lbPrams(1).Top = lbPrams(1).Top + Grid.CellHeight
        Else
            i = 2
'            listBoxInGridCell lbPrams2, Grid
'            lbPrams2.Top = lbPrams2.Top + Grid.CellHeight
        End If
        listBoxInGridCell lbPrams(i), Grid
        lbPrams(i).Top = lbPrams(i).Top + Grid.CellHeight
    End If
End If

End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width

Grid.FormatString = "|№|<Принадлежность|<Формула|<Примечание"
Grid.colWidth(0) = 0
Grid.colWidth(gfNomer) = 450
Grid.colWidth(gfFormula) = 6120
Grid.colWidth(gfForWho) = 1395
Grid.colWidth(gfNote) = 1980
sql = "SELECT sGuideFormuls.Formula, sGuideFormuls.Note, " & _
"sGuideFormuls.forWho, sGuideFormuls.nomer " & _
"From sGuideFormuls  ORDER BY sGuideFormuls.nomer;"
'"From sGuideFormuls  ORDER BY sGuideFormuls.forWho, sGuideFormuls.nomer;"
Set tbGuide = myOpenRecordSet("##305", sql, dbOpenForwardOnly)
'Set tbGuide = myOpenRecordSet("##305", "sGuideFormuls", dbOpenTable)
If tbGuide Is Nothing Then Exit Sub

'tbGuide.Index = "nomer"
quantity = 0
While Not tbGuide.EOF
    If tbGuide!nomer = 0 Then GoTo NXT1
    If Regim = "fromNomenkW" And tbGuide!forWho <> lbForWho.List(2) Then GoTo NXT1
    If Regim = "fromNomenk" And tbGuide!forWho <> lbForWho.List(0) Then GoTo NXT1
    If Regim = "fromProduct" And tbGuide!forWho <> lbForWho.List(1) Then GoTo NXT1
    quantity = quantity + 1
    Grid.TextMatrix(quantity, gfNomer) = tbGuide!nomer
    Grid.TextMatrix(quantity, gfFormula) = tbGuide!formula
    Grid.TextMatrix(quantity, gfForWho) = tbGuide!forWho
    Grid.TextMatrix(quantity, gfNote) = tbGuide!note
    Grid.AddItem ""
NXT1:
    tbGuide.MoveNext
Wend
tbGuide.Close

    ' init the Global constants to use its in formulas
    sql = "select * from GuideConstants"
    Set tbGuide = myOpenRecordSet("##0.3", sql, dbOpenForwardOnly)
    If tbGuide Is Nothing Then Exit Sub
    While Not tbGuide.EOF
        Dim initStr As String, paramListsNo As Integer
        For paramListsNo = 0 To 2
            Dim lb As ListBox
            Set lb = Me.lbPrams(paramListsNo)
            lb.AddItem tbGuide!Constants
        Next paramListsNo
        tbGuide.MoveNext
    Wend
    tbGuide.Close


If quantity > 0 Then Grid.RemoveItem quantity + 1
If Regim = "" Then Grid.TabIndex = 0: Grid_EnterCell
    
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
cmExit.left = cmExit.left + w

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
    
    If mousCol = gfNomer Then
        SortCol Grid, mousCol, "numeric"
    Else
        SortCol Grid, mousCol
    End If
    Grid.row = 1    ' только чтобы снять выделение
    Grid_EnterCell
End If

End Sub

Private Sub Grid_DblClick()

If Regim <> "" Then
    cmSel_Click
ElseIf mousCol = gfForWho Then
    listBoxInGridCell lbForWho, Grid
ElseIf Grid.CellBackColor = &H88FF88 Then
    tbMobile.MaxLength = 100
    textBoxInGridCell tbMobile, Grid
    If mousCol = gfFormula Then tbMobile.SelLength = 0
End If

End Sub

Private Sub Grid_EnterCell()
 mousRow = Grid.row
 mousCol = Grid.col
 
 If Regim <> "" Then
    Grid.CellBackColor = &HFFFFAA
 ElseIf mousCol > 0 Then
    Grid.CellBackColor = &H88FF88
 Else
    Grid.CellBackColor = vbYellow
 End If
'End If

End Sub

Private Sub Grid_GotFocus()
Grid_EnterCell
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
lbForWho.Visible = False
lbPrams(0).Visible = False
lbPrams(1).Visible = False
lbPrams(2).Visible = False
'lbPrams2.Visible = False
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

Private Sub lbForWho_DblClick()
If Grid.TextMatrix(mousRow, gfNomer) = "" Then
    Grid.col = gfNomer
    mousCol = gfNomer
    On Error Resume Next
    Grid.SetFocus
    textBoxInGridCell tbMobile, Grid
Else
    If ValueToFormulsField("##308", lbForWho.Text, "forWho") = 0 Then _
                Grid.Text = lbForWho.Text
End If
lbHide

End Sub

Private Sub lbForWho_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbForWho_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub lbPrams_DblClick(index As Integer)
Dim str As String, i As Integer '

str = left$(tbMobile.Text, tbMobile.SelStart)
str = str & lbPrams(index).Text
i = Len(str)
str = str & Mid$(tbMobile.Text, tbMobile.SelStart + 1)
tbMobile.Text = str
tbMobile.SelStart = i

tbMobile.SetFocus
lbPrams(index).Visible = False

End Sub


' возвращает 0 если не было ошибок. если ошибка имела место быть, то тогда
' возвращаем id формулы с ошибкой. -1 если ошибка для текущей формулы.

Function checkFormul(formula As String) As Boolean
Dim sclocal, scode As String
Dim dimStr As String, assignStr As String

checkFormul = False
Set sclocal = CreateObject("ScriptControl")
sclocal.Language = "VBScript"

    sql = "select * from GuideConstants"
    Set tbGuide = myOpenRecordSet("##0.3", sql, dbOpenForwardOnly)
    If tbGuide Is Nothing Then Exit Function
    While Not tbGuide.EOF
        dimStr = dimStr & ", " & tbGuide!Constants
        assignStr = assignStr & ": " & tbGuide!Constants & "=1"
        tbGuide.MoveNext
    Wend
    tbGuide.Close


scode = "Option Explicit" & vbCrLf & _
"Private Function Calc()" & vbCrLf & _
"Dim CENA1, VES, STAVKA, SumCenaFreight, VremObr, CenaFreight, cenaFact, SumCenaSale " & dimStr
scode = scode & vbCrLf & "CENA1=1: VES=1: STAVKA=1: SumCenaFreight=1: " & vbCrLf & _
"VremObr=1: CenaFreight=1: CenaFact=1: SumCenaSale=1" & assignStr
scode = scode & vbCrLf & "Calc = " & formula & vbCrLf & _
"End Function" & vbCrLf

On Error GoTo ERR1
sclocal.AddCode scode 'проверяется синтаксис
sclocal.Eval "Calc()" 'проверяются переменные

Set sclocal = Nothing
checkFormul = True
Exit Function

ERR1:
If Err <> 11 Then ' пропускаем деление на 0 - нам нужен только синтаксис
    MsgBox "Ошибка в формуле:   " & Error, , "Error 309 - " & Err & ":  "  '##309
End If
End Function


Private Sub lbPrams_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbPrams_DblClick index
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, i As Integer

'notValidKey = False
If KeyCode = vbKeyReturn Then
  str = Trim(tbMobile.Text)
  If mousCol = gfNomer Then
    If Not isNumericTbox(tbMobile, 0, 255) Then Exit Sub
    If frmMode = "rowAdd" Then
'      sql = "SELECT sGuideFormuls.nomer, sGuideFormuls.Formula, " & _
'      "sGuideFormuls.forWho, sGuideFormuls.Note " & _
'      "From sGuideFormuls ORDER BY sGuideFormuls.nomer;"
'      Set tbGuide = myOpenRecordSet("##306", sql, dbOpenDynaset)
'      If tbGuide Is Nothing Then Exit Sub
'      tbGuide.AddNew
'      tbGuide!nomer = str
'      tbGuide!formula = "1"
'      tbGuide!forWho = lbForWho.List(0)
'      On Error GoTo ERR1
'      tbGuide.Update
'      tbGuide.Close
             
      sql = "INSERT INTO sGuideFormuls (nomer, formula, forWho) " & _
      "VALUES (" & str & ", '1' ,'" & lbForWho.List(0) & "')"
      i = myExecute("##472", sql, -193)
      If i <> 0 Then GoTo ERR0
      
      
      Grid.TextMatrix(mousRow, gfNomer) = str
      Grid.TextMatrix(mousRow, gfFormula) = "1"
      Grid.TextMatrix(mousRow, gfForWho) = lbForWho.List(0)
    Else
       i = ValueToFormulsField("##307", str, "nomer", -193)
       If i <> 0 Then GoTo ERR0
'       If i = 3022 Then
'            MsgBox "Этот номер уже задействован.", , "Ошибка!"
'            Exit Sub
'       ElseIf i <> 0 Then
'            GoTo EN1
'       End If
    End If
  ElseIf mousCol = gfFormula Then
       If Not checkFormul(str) Then Exit Sub
       i = ValueToFormulsField("##303", str, "Formula")
  ElseIf mousCol = gfNote Then
       If ValueToFormulsField("##303", str, "Note") <> 0 Then GoTo EN1
  End If
  Grid.TextMatrix(mousRow, mousCol) = str
  GoTo EN1
ElseIf KeyCode = vbKeyEscape Then
CNC:
    If mousCol = gfNomer And frmMode = "rowAdd" Then
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
'    MsgBox "Этот номер уже задействован", , "Ошибка"
'    GoTo CNC
'Else
'    MsgBox Error, , "Ошибка 143-" & Err & ":  " '##143
'    End
'End If

ERR0:
If i = -2 Then
    MsgBox "Этот номер уже задействован", , "Ошибка - " & cErr
    tbMobile.SetFocus
Else
    GoTo EN1
End If

End Sub

Function ValueToFormulsField(myErrCod As String, value As String, _
field As String, Optional passErr As Integer = -1) As Integer
Dim i As Integer

ValueToFormulsField = False

sql = "UPDATE sGuideFormuls SET sGuideFormuls." & field & " = '" & value & _
"' WHERE (((sGuideFormuls.nomer)=" & Grid.TextMatrix(mousRow, gfNomer) & "));"
'MsgBox "sql = " & sql

ValueToFormulsField = myExecute(myErrCod, sql, passErr)
End Function


