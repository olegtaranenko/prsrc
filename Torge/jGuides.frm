VERSION 5.00
Begin VB.Form jGuidePurpose 
   BackColor       =   &H8000000A&
   Caption         =   "Справочник операций"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmLoad 
      Caption         =   "Обновить"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   915
   End
   Begin VB.CommandButton cmExit 
      Cancel          =   -1  'True
      Caption         =   "Выход"
      Height          =   315
      Left            =   5100
      TabIndex        =   6
      Top             =   4080
      Width           =   915
   End
   Begin VB.CommandButton cmSel 
      Caption         =   "Выбрать"
      Height          =   315
      Left            =   2460
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   1020
      Visible         =   0   'False
      Width           =   2235
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<-- Недопустимая операция"
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   60
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmSchet 
      Caption         =   "Справ-к счетов"
      Height          =   315
      Left            =   600
      TabIndex        =   10
      Top             =   3540
      Width           =   1335
   End
   Begin VB.ListBox lbKredit 
      Height          =   2985
      Left            =   1440
      TabIndex        =   1
      Top             =   420
      Width           =   1275
   End
   Begin VB.ListBox lbDebit 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   1095
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   3540
      Width           =   975
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   3540
      Width           =   975
   End
   Begin VB.ListBox lbPurpose 
      Height          =   2985
      Left            =   3000
      TabIndex        =   2
      Top             =   420
      Width           =   2835
   End
   Begin VB.Label laPurpose 
      Alignment       =   2  'Center
      Caption         =   "Назначение"
      Height          =   195
      Left            =   1740
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Кредит"
      Height          =   195
      Left            =   960
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Дебет"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "jGuidePurpose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Regim As String
Public purpose As String ', detail As String

'Const sgSchet = 1
'Const sgSubSchet = 2

Private Sub cmAdd_Click()
Dim i As Integer

jGuidePurpDet.Regim = "purpose"
tmpStr = ""
jGuidePurpDet.Show vbModal
If tmpStr = "" Then Exit Sub

wrkDefault.BeginTrans 'lock01
sql = "UPDATE yGuidePurpose SET pDescript = pDescript WHERE (((Debit)='255') " & _
"AND ((subDebit)='00') AND ((Kredit)='255') AND ((subKredit)='00') AND ((pId)=0));"
myBase.Execute (sql) 'lock03


sql = "SELECT pId From yGuidePurpose " & _
"WHERE (((Debit)='" & debit & "') AND ((subDebit)='" & subDebit & _
"') AND ((Kredit)='" & kredit & "') AND ((subKredit)='" & subKredit & _
"')) ORDER BY pId"
Set tbGuide = myOpenRecordSet("##476", sql, dbOpenForwardOnly)
'If tbGuide Is Nothing Then Exit Sub
'tbGuide.Index = "Key"

'tbGuide.Seek "=", debit, subDebit, kredit, subKredit, 0
i = 0
'If Not tbGuide.NoMatch Then
If Not tbGuide.BOF Then
'    tbGuide.MoveNext: i = 1
    While Not tbGuide.EOF ' сначала исп-ем удаленные номера
        If tbGuide!pId > i Then GoTo AA
        tbGuide.MoveNext
        i = i + 1
    Wend
    If i > 255 Then msgOfEnd "##335", "переполнение yGuidePurpose"
End If
AA:
tbGuide.Close

sql = "INSERT INTO yGuidePurpose (Debit, subDebit, Kredit, subKredit, " & _
"pId, pDescript ) VALUES ('" & debit & "', '" & subDebit & "', '" & kredit & _
"', '" & subKredit & "', " & i & ", '" & tmpStr & "')"
i = myExecute("##336", sql, -196)
If i = -2 Then
    MsgBox "Назначение '" & tmpStr & "' уже есть!", , "Предупреждение"
ElseIf i <> 0 Then
    Exit Sub
End If
wrkDefault.CommitTrans

'tbGuide.AddNew
'tbGuide!debit = debit
'tbGuide!subDebit = subDebit
'tbGuide!kredit = kredit
'tbGuide!subKredit = subKredit
'tbGuide!pId = i
'tbGuide!pDescript = tmpStr
tmpStr = lbPurpose.Text

'On Error GoTo ERR1
'tbGuide.Update
'tbGuide.Close

loadLbFromPurpose lbPurpose, tmpStr
'Exit Sub

'ERR1:
'If Err = 3022 Then
'    MsgBox "Назначение '" & tmpStr & "' уже есть!", , "Предупреждение"
'Else
'    MsgBox Error, , "Ошибка 336-" & Err & ":  " '##336
'End If

End Sub




Private Sub cmDel_Click()
Dim i As Integer

sql = "DELETE From yGuidePurpose " & _
"WHERE (((Debit)='" & debit & "') AND ((subDebit)='" & subDebit & _
"') AND ((Kredit)='" & kredit & "') AND ((subKredit)='" & subKredit & _
"') AND ((pDescript)='" & lbPurpose.Text & "'));"
i = myExecute("##337", sql, -198)
If i = 0 Then
    loadLbFromPurpose lbPurpose
ElseIf i = -2 Then
    MsgBox "Для удаления Назначения '" & lbPurpose.Text & "', сначала " & _
    "удалите входящие в него Уточнения.", , "Удаление непустого Назначения невозможно!"
    Exit Sub
End If
If lbPurpose.ListCount > 0 Then
    lbPurpose.SetFocus
    lbPurpose.ListIndex = 0
Else
    cmDel.Enabled = False
End If
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmLoad_Click()
loadListboxes
lbDebit.SetFocus
End Sub

Private Sub cmSchet_Click()
jGuideSchets.Show vbModal
End Sub

Private Sub cmSel_Click()

'If lbPurpose.ListIndex = -1 Or lbDetail.ListIndex = -1 Then
'    MsgBox "Обязательно установите Назначение операции  и Уточнение.", , "Предупреждение"
'    Exit Sub
'End If
If Regim = "select" Then '  из Journal
    If Journal.valueToBookField("##346", 0, "schets") Then
        Journal.setSchetsAndPurpose lbPurpose.Text ', lbDetail.Text
        If jKassaReport.isLoad Then jKassaReport.laInform.Visible = True
        Unload Me
    End If
Else 'из Nastroy

' сбрасываем флаг auto в старой операции
    debit = QQ(1): subDebit = QQ(2): kredit = QQ(3): subKredit = QQ(4)
    
    wrkDefault.BeginTrans
    
    If purpose <> "" Then ' старой пока не было
        purposeId = getPurposeIdByDescript(purpose)
        cErr = 399 '##399
        If purposeId < 0 Then 'GoTo ERR1
            wrkDefault.Rollback
            MsgBox "Запись не обнаружена!", , "Err " & cErr
            Exit Sub
        End If
            
        sql = "UPDATE yGuidePurpose SET AUTO=''  " & _
        "WHERE (((Debit)='" & debit & "') AND ((subDebit)='" & subDebit & _
        "') AND ((Kredit)='" & kredit & "') AND ((subKredit)='" & subKredit & _
        "') AND ((pId)=" & purposeId & "))"
        If myExecute("##397", sql) <> 0 Then Exit Sub
    End If
' устанавливаем флаг auto в новой операции
    getSchetsFromLb
    purposeId = getPurposeIdByDescript(lbPurpose.Text)
        
    'AUTO= r, z или m
    sql = "UPDATE yGuidePurpose SET AUTO='" & Right$(Regim, 1) & "'  " & _
    "WHERE (((Debit)='" & debit & "') AND ((subDebit)='" & subDebit & _
    "') AND ((Kredit)='" & kredit & "') AND ((subKredit)='" & subKredit & _
    "') AND ((pId)=" & purposeId & "));"
    If myExecute("##398", sql) <> 0 Then Exit Sub
    
    wrkDefault.CommitTrans
    
    Nastroy.paramLoad
    Unload Me
'EN1: tbGuide.Close
End If
End Sub

Private Sub Form_Load()
'loadLbFromTableField lbPurpose, "yGuidePurpose", "pDescript"
'If Regim = "select" Then
If Left$(Regim, 6) = "select" Then
    cmSel.Visible = True
    Me.Caption = "Выбор операции и назначения"
Else
    cmSel.Visible = False
    Me.Caption = "Справочник назначений"
End If
loadListboxes
End Sub

Sub loadListboxes()

'If Regim = "select" And debit < 255 Then
If Left$(Regim, 6) = "select" And debit <> "255" Then
    Journal.loadLbFromSchets lbDebit, debit, subDebit
    Journal.loadLbFromSchets lbKredit, kredit, subKredit
    loadLbFromPurpose lbPurpose, purpose
'    loadLbFromDetail Me, detail
Else
  Journal.loadLbFromSchets lbDebit
  Journal.loadLbFromSchets lbKredit
  If lbDebit.ListCount > 0 Then
        lbDebit.ListIndex = 0
        cmAdd.Enabled = True
  End If
  If lbKredit.ListCount > 1 Then
    lbKredit.ListIndex = 1
  ElseIf lbKredit.ListCount > 0 Then
    lbKredit.ListIndex = 0
  End If
End If
End Sub

Private Sub Form_Resize()
lbDebit.SetFocus
End Sub


Function debKreditInit() As Boolean
Dim str  As String

debKreditInit = False
If lbDebit.ListIndex = -1 Or lbKredit.ListIndex = -1 Then Exit Function
cmSel.Enabled = True
'lbDetail.Clear
cmAdd.Enabled = True
'cmAdd2.Enabled = False
cmDel.Enabled = False
'cmDel2.Enabled = False

Frame1.Visible = False
laPurpose.Enabled = True
'laDetail.Enabled = True

getSchetsFromLb

debKreditInit = True
End Function

Sub getSchetsFromLb()
Dim str As String

debit = Left$(lbDebit.Text, 2)
str = Mid$(lbDebit.Text, 4, 2)
subDebit = "00"
If str <> "" Then subDebit = str

kredit = Left$(lbKredit.Text, 2)
str = Mid$(lbKredit.Text, 4, 2)
subKredit = "00"
If str <> "" Then subKredit = str
End Sub

Sub loadLbFromPurpose(lb As ListBox, Optional fit As String = "")
Dim line As Integer


sql = "SELECT pDescript From yGuidePurpose WHERE (((Debit)='" & debit & _
"') AND ((subDebit)='" & subDebit & "') AND ((Kredit)='" & kredit & _
"') AND ((subKredit)='" & subKredit & "'))  ORDER BY pDescript;"

Debug.Print sql
Set Table = myOpenRecordSet("##333", sql, dbOpenDynaset)
If Table Is Nothing Then Exit Sub
lb.Clear
While Not Table.EOF
  lb.AddItem Table!pDescript
  If fit = Table!pDescript Then lb.ListIndex = lb.ListCount - 1 'здесь обновляется lbDetail, т.к. срабатывает lbPurpose_Click
  Table.MoveNext
Wend
Table.Close

If fit = "height" Then lb.Height = 195 * lb.ListCount + 100
End Sub

'Private Sub lbDetail_Click()
''If noClick Then Exit Sub
'If lbDetail.ListCount > 0 Then cmDel2.Enabled = True
'
'End Sub
'
Private Sub lbKredit_Click()
If debKreditInit() Then loadLbFromPurpose lbPurpose
End Sub

Private Sub lbPurpose_Click()
'If noClick Then Exit Sub
If lbPurpose.ListCount > 0 Then
    cmDel.Enabled = True
'    loadLbFromDetail Me
'    cmAdd2.Enabled = True
'Else
'    cmAdd2.Enabled = False
End If
'cmDel2.Enabled = False

End Sub

Private Sub lbDebit_Click()
If debKreditInit() Then loadLbFromPurpose lbPurpose
End Sub

Private Sub tbPurpose_Change()

End Sub
