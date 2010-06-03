VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Zagruz 
   BackColor       =   &H8000000A&
   Caption         =   "Загрузка"
   ClientHeight    =   6096
   ClientLeft      =   4668
   ClientTop       =   1656
   ClientWidth     =   7848
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6096
   ScaleWidth      =   7848
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   5640
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmHistory 
      Caption         =   "Журнал"
      Height          =   315
      Left            =   3540
      TabIndex        =   20
      Top             =   5640
      Width           =   795
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Сброс"
      Height          =   315
      Left            =   5820
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmExAll 
      Caption         =   "Выход"
      Height          =   315
      Left            =   6960
      TabIndex        =   10
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox tbMobile 
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Text            =   " tbMobile"
      Top             =   780
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmRefr 
      Caption         =   "Обновить"
      Height          =   315
      Left            =   4620
      TabIndex        =   8
      Top             =   5640
      Width           =   915
   End
   Begin VB.TextBox tbKPD 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Text            =   "1,0"
      Top             =   4900
      Width           =   435
   End
   Begin VB.TextBox tbStanki 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Text            =   "2"
      Top             =   5625
      Width           =   435
   End
   Begin VB.TextBox tbNomRes 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Text            =   "8"
      Top             =   5250
      Width           =   435
   End
   Begin VB.CheckBox chDopView 
      Caption         =   "см. Распред."
      Height          =   315
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4515
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      _ExtentX        =   13145
      _ExtentY        =   7959
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Дата"
         Object.Width           =   1614
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Принято заказов"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "nomRes"
         Text            =   "Смена"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Ресурс"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Загрузка"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Остатки"
         Object.Width           =   1517
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Живые"
         Object.Width           =   1455
      EndProperty
   End
   Begin VB.Label laVirab 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3660
      TabIndex        =   19
      Top             =   5250
      Width           =   660
   End
   Begin VB.Label laUsed 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3660
      TabIndex        =   18
      Top             =   4900
      Width           =   660
   End
   Begin VB.Label Label2 
      Caption         =   "Выработка"
      Height          =   195
      Left            =   2760
      TabIndex        =   17
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Исп.ресурс"
      Height          =   195
      Left            =   2760
      TabIndex        =   16
      Top             =   4920
      Width           =   915
   End
   Begin VB.Label laZagLive 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   4740
      Width           =   825
   End
   Begin VB.Label laZagAll 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4610
      TabIndex        =   13
      Top             =   4740
      Width           =   900
   End
   Begin VB.Label laZap 
      Caption         =   "Запас:"
      Height          =   195
      Left            =   5280
      TabIndex        =   12
      Top             =   0
      Width           =   555
   End
   Begin VB.Label laZapas 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   0
      Width           =   855
   End
   Begin VB.Label laKPD 
      Caption         =   "Эффект-ть производства:"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   4920
      Width           =   1995
   End
   Begin VB.Label laStanki 
      Caption         =   "Число станков:"
      Height          =   195
      Left            =   900
      TabIndex        =   4
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label laNomRes 
      Caption         =   "Смена по умолчанию:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   5280
      Width           =   1695
   End
End
Attribute VB_Name = "Zagruz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public idEquip As Integer
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы



Private Sub cmDel_Click()
If MsgBox("Если нажать <Да>, то на все дни будет установлен ресурс по " & _
"умолчанию.", vbYesNo Or vbDefaultButton2, "Сбросить ресурсы?") = vbYes Then
    sql = "DELETE FROM Resurs where equipId = " & idEquip
'MsgBox sql
    myExecute "##137", sql, 0
    valueToSystemField "##361", "", "resursLock"
End If
End Sub

Private Sub cmExAll_Click()
Unload Me
'    exitAll
End Sub


Private Sub cmHistory_Click()
Report.Regim = "Virabotka"
Report.idEquip = idEquip
Report.Show vbModal
End Sub

Private Sub Command1_Click()
'MsgBox "left=" & laZagLive.Left 'laZagAll.Left
MsgBox "width=" & laZagLive.Width
End Sub

Private Sub Form_Activate()
If dostup = "m" Or dostup = "a" Then
    tbKPD.Locked = False
    tbNomRes.Locked = False
    tbStanki.Locked = False
    chDopView.Enabled = True
Else
    tbKPD.Locked = True
    tbNomRes.Locked = True
    tbStanki.Locked = True
    chDopView.Enabled = False
End If
Me.Caption = "загрузка цеха " & Equip(idEquip)

If dostup <> "" Then cmDel.Visible = True
ZagruzLoad

End Sub

Private Sub Form_Load()
oldHeight = Me.Height
oldWidth = Me.Width

isZagruz = True
End Sub

Sub ZagruzLoad() ' бывшая begZagruz
 Dim I As Integer, Key As String, tekDate As String, S As Double

maxDay = 0
zagruzFromCeh idEquip ' в delta
getResurs idEquip  ' выч-е nomRes()

tbKPD.Text = KPD
tbStanki.Text = Nstan
tbNomRes.Text = newRes
cmRefr.Caption = "Обновить"
lv.ListItems.Clear

Zakaz.newZagruz "fromCehZagruz"

For I = 1 To maxDay
    tekDate = Format(DateAdd("d", I - 1, curDate), "dd/mm/yy")
    Key = "k" & I
    lv.ListItems.Add , Key, tekDate

    day = Weekday(DateAdd("d", I - 1, curDate))
    If day = vbSunday Or day = vbSaturday Then
        lv.ListItems(Key).ForeColor = &HFF&
    End If
  
    If I = stDay Then
        lv.ListItems(Key).ForeColor = &HBB00&
        lv.ListItems(Key).Bold = True
    End If
    
    lv.ListItems(Key).SubItems(zgPrinato) = Round(getNevip(I), 1)

    lv.ListItems(Key).SubItems(zgNomRes) = nomRes(I)
    
    S = Round(nomRes(I) * KPD * Nstan, 1)
    lv.ListItems(Key).SubItems(zgResurs) = S
    lv.ListItems(Key).SubItems(zgZagruz) = Round(S - ost(I), 1)
    lv.ListItems(Key).SubItems(zgOstatki) = Round(ost(I), 1)
    lv.ListItems(Key).SubItems(zgLive) = Round(S - befOst(I), 1)
    If ost(I) < 0 Then
        lv.ListItems(Key).ListSubItems(zgOstatki).Bold = True
        lv.ListItems(Key).ListSubItems(zgOstatki).ForeColor = 200
    End If
Next I

S = Round(nr * Nstan * KPD, 1)
lv.ListItems("k1").SubItems(zgResurs) = S
lv.ListItems("k1").SubItems(zgZagruz) = Round(S - ost(1), 1)
lv.ListItems("k1").SubItems(zgLive) = Round(S - befOst(1), 1)

lv.ListItems("k" & stDay).ForeColor = &HBB00&
lv.ListItems("k" & stDay).Bold = True
I = getNextDay(1)
laZapas.Caption = Round(nomRes(I) * KPD * Nstan + ost(1), 1)

zagAll = 0
zagLive = 0
For I = 1 To maxDay
    Key = "k" & I
    zagAll = zagAll + lv.ListItems(Key).SubItems(zgZagruz)
    zagLive = zagLive + lv.ListItems(Key).SubItems(zgLive)
Next I

laZagAll.Caption = Round(zagAll, 1) & "  "
laZagLive.Caption = Round(zagLive, 1) & "  "
laUsed.Caption = Round((nomRes(1) - nr) * Nstan * KPD, 2)

sql = "SELECT Sum(Virabotka) AS Sum_V From Itogi" _
& " WHERE numOrder >10  AND xDate ='" & Format(curDate, "yy.mm.dd") & "' AND equipId = " & idEquip
'Debug.Print sql
If byErrSqlGetValues("##375", sql, S) Then laVirab.Caption = Round(S, 2)

End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next
tbMobile.Visible = False

h = Me.Height - oldHeight
oldHeight = Me.Height
Me.Width = oldWidth

lv.Height = lv.Height + h
laKPD.Top = laKPD.Top + h
tbKPD.Top = tbKPD.Top + h
laNomRes.Top = laNomRes.Top + h
tbNomRes.Top = tbNomRes.Top + h
laStanki.Top = laStanki.Top + h
tbStanki.Top = tbStanki.Top + h
cmRefr.Top = cmRefr.Top + h
chDopView.Top = chDopView.Top + h
cmExAll.Top = cmExAll.Top + h
laZagAll.Top = laZagAll.Top + h
laZagLive.Top = laZagLive.Top + h
cmDel.Top = cmDel.Top + h
End Sub

Private Sub Form_Unload(Cancel As Integer)
isZagruz = False
End Sub

Private Sub lv_Click()
chDopView.value = 0
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
laZagLive.Width = laZagLive.Width + 20
End Sub

Private Sub lv_DblClick()
flClickDouble = True
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not (dostup = "m" Or dostup = "a") Then Exit Sub

Set ClickItem = lv.HitTest(X, Y)
If ClickItem Is Nothing Then Exit Sub
    
    If Not flClickDouble Then Exit Sub

    If X < (lv.ColumnHeaders("nomRes").Left + lv.ColumnHeaders("nomRes").Width) _
    And X > lv.ColumnHeaders("nomRes").Left Then

    tbMobile.Left = lv.ColumnHeaders("nomRes").Left + lv.Left + 20
    tbMobile.Top = ClickItem.Top + lv.Top + 50
    tbMobile.Visible = True
    tbMobile.Text = ClickItem.SubItems(zgNomRes)
    tbMobile.SetFocus
    tbMobile.SelLength = Len(tbMobile.Text)
    tbMobile.ZOrder
    flEdit = "mobile"
    End If
    
flClickDouble = False
End Sub

Private Sub cmRefr_Click()
cmRefr.Caption = "Обновить"

sql = "select * from GuideResurs where equipId = " & idEquip
Set tbSystem = myOpenRecordSet("##182", sql, dbOpenForwardOnly)
If tbSystem Is Nothing Then myBase.Close: End
 tbSystem.Edit
 ' сохраняем параметры по умолчанию
    tbSystem!KPD = tbKPD.Text
    tbSystem!Nstan = tbStanki.Text
    tbSystem!newRes = tbNomRes.Text
tbSystem.update
tbSystem.Close
ZagruzLoad
End Sub

Private Sub tbKPD_Change()
cmRefr.Caption = "Сохранить"
End Sub

Private Sub tbKPD_Click()
chDopView.value = 0
End Sub

'$odbc14$
Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim S As Double, dayMax As Integer, str As String, I As Integer

If KeyCode = vbKeyReturn Then
  If isNumericTbox(tbMobile, 0, 22) Then
        
    
    day = Mid$(Zagruz.lv.SelectedItem.Key, 2)
    nomRes(day) = tbMobile.Text
    
    ' макс дата в таблице ресурса
'    sql = "SELECT Max(xDate) AS MD from Resurs" & Equip(idEquip) & ";"
    sql = "SELECT Count(xDate) AS Count_Date FROM Resurs where equipId = " & idEquip
'    MsgBox sql
    If Not byErrSqlGetValues("##411", sql, dayMax) Then Exit Sub
'    If dayMax = 0 Then dayMax = 1
    
    wrkDefault.BeginTrans
    
    If day <= dayMax Then ' если день есть в табл.ресурса
        sql = "UPDATE Resurs SET nomRes = " & tbMobile.Text & _
        " WHERE xDate ='" & yymmdd(lv.SelectedItem.Text) & "' and equipId = " & idEquip
'Debug.Print sql

'        MsgBox sql
        If myExecute("##66", sql) <> 0 Then Exit Sub
    Else ' иначе обавляем дни
        For I = dayMax + 1 To day
            sql = "INSERT INTO Resurs (equipId, xDate, nomRes ) " & _
            "SELECT " & idEquip & ", '" & yymmdd(lv.ListItems("k" & I).Text) & "', " & nomRes(I) & ";"
'            MsgBox sql
            If myExecute("##413", sql) <> 0 Then Exit Sub
        Next I
    End If
    
    wrkDefault.CommitTrans
    
        
    ZagruzLoad  ' с учетом новых значений
        
    tbMobile.Visible = False
'        flEdit = ""
  End If
    
ElseIf KeyCode = vbKeyEscape Then
    tbMobile.Visible = False
End If

End Sub


Private Sub tbNomRes_Change()
cmRefr.Caption = "Сохранить"

End Sub

Private Sub tbNomRes_Click()
chDopView.value = 0
End Sub

Private Sub tbStanki_Change()
cmRefr.Caption = "Сохранить"

End Sub

Private Sub tbStanki_Click()
chDopView.value = 0
End Sub

