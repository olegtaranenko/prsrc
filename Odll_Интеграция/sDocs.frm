VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form sDocs 
   BackColor       =   &H8000000A&
   Caption         =   "Расходные накладные"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ckPerList 
      Caption         =   "В целых"
      Height          =   195
      Left            =   10800
      TabIndex        =   30
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmBay 
      Caption         =   "продаж"
      Height          =   315
      Left            =   3300
      TabIndex        =   29
      Top             =   5460
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   5280
      Top             =   4860
   End
   Begin VB.CommandButton cmClose 
      Caption         =   "Списать"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4920
      TabIndex        =   28
      Top             =   5460
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox ckCeh 
      Caption         =   "Из цеха"
      Height          =   315
      Left            =   4080
      TabIndex        =   27
      Top             =   0
      Width           =   915
   End
   Begin VB.ListBox lbGroup 
      Height          =   450
      ItemData        =   "sDocs.frx":0000
      Left            =   3660
      List            =   "sDocs.frx":000A
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ListBox lbStatia 
      Height          =   255
      Left            =   3660
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Frame frZakaz 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5115
      Left            =   1560
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ListBox lbZakaz 
         Height          =   4155
         Left            =   180
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmOk 
         Caption         =   "Ok"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   4620
         Width           =   855
      End
      Begin VB.CommandButton cmCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   315
         Left            =   1260
         TabIndex        =   21
         Top             =   4620
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Выберите заказ."
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   5115
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label laFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Выберите склад:"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "Печать накладной"
      Height          =   315
      Left            =   8760
      TabIndex        =   19
      Top             =   5460
      Width           =   1695
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   10920
      TabIndex        =   18
      Top             =   5460
      Width           =   795
   End
   Begin VB.CommandButton cmAdd2 
      Caption         =   "Добавить\Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6120
      TabIndex        =   17
      Top             =   5460
      Width           =   1635
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   360
      TabIndex        =   16
      Text            =   "tbMobile"
      Top             =   1020
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lbInside 
      Height          =   255
      Left            =   3660
      TabIndex        =   13
      Top             =   1500
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CheckBox ckStartDate 
      Caption         =   " "
      Height          =   315
      Left            =   1140
      TabIndex        =   9
      Top             =   0
      Width           =   315
   End
   Begin VB.TextBox tbStartDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Text            =   "01.11.02"
      Top             =   0
      Width           =   795
   End
   Begin VB.CheckBox ckEndDate 
      Caption         =   " "
      Height          =   315
      Left            =   2640
      TabIndex        =   7
      Top             =   0
      Width           =   315
   End
   Begin VB.TextBox tbEndDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2940
      TabIndex        =   6
      Top             =   0
      Width           =   795
   End
   Begin VB.CommandButton cmLoad 
      Caption         =   "Загрузить"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   5460
      Width           =   1095
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   5460
      Width           =   975
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   5460
      Width           =   915
   End
   Begin VB.CommandButton cmOrder 
      Caption         =   "к Заказу"
      Height          =   315
      Left            =   2340
      TabIndex        =   2
      Top             =   5460
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5055
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8916
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   5055
      Left            =   5940
      TabIndex        =   1
      Top             =   300
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8916
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Label laGrid2 
      Height          =   195
      Left            =   6000
      TabIndex        =   15
      Top             =   60
      Width           =   4755
   End
   Begin VB.Label laPeriod 
      Caption         =   "Период с  "
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   60
      Width           =   795
   End
   Begin VB.Label laPo 
      Caption         =   "пос"
      Height          =   195
      Left            =   2340
      TabIndex        =   10
      Top             =   60
      Width           =   195
   End
End
Attribute VB_Name = "sDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim quantity2 As Long
Dim quantity  As Long
Public mousCol As Long, mousRow As Long
Dim mousCol2 As Long, mousRow2 As Long
Dim prevRow As Long
Public gridIzLoad As Boolean
Public Regim As String
Dim statiaId() As String
Dim destId As Integer
Dim byClick As Boolean
Public isLoad As Boolean

Dim typeId As Integer
Const dnNomNom = 1
Const dnNomName = 2
Const dnEdIzm = 3
Const dnQuant = 4
Const dnLists = 5
' изменения для зарезервинных заказов
Const dnQuantRez = 4
Const dnQntClose = 5
Const dnNowOstatRez = 6

Private Sub cbEndDate_Click()
tbEndDate.Enabled = Not tbEndDate.Enabled
End Sub

Private Sub cbStartDate_Click()
tbStartDate.Enabled = Not tbStartDate.Enabled
End Sub

Private Sub ckCeh_Click()
If ckCeh.value = 0 Then
    Me.Caption = "Расходные накладные"
    cmAdd.Visible = True
    cmOrder.Visible = True
    cmBay.Visible = True
    cmDel.Visible = True
    cmAdd2.Visible = True
    cmPrint.Visible = True
    laPeriod.Visible = True
    ckStartDate.Visible = True
    tbStartDate.Visible = True
    laPo.Visible = True
    ckEndDate.Visible = True
    tbEndDate.Visible = True
    
    cmClose.Visible = False
    
Else
    Me.Caption = "Выписанные накладные не под заказы"
    cmAdd.Visible = False
    cmOrder.Visible = False
    cmBay.Visible = False
    cmDel.Visible = False
    cmAdd2.Visible = False
    cmPrint.Visible = False
    laPeriod.Visible = False
    ckStartDate.Visible = False
    tbStartDate.Visible = False
    laPo.Visible = False
    ckEndDate.Visible = False
    tbEndDate.Visible = False
    
    cmClose.Visible = True
    
End If
Grid2.Visible = False
quantity = 0
clearGrid Grid


End Sub

Private Sub ckEndDate_Click()
If ckEndDate.value = 1 Then
    tbEndDate.Enabled = True
Else
    tbEndDate.Enabled = False
End If

End Sub

Private Sub ckPerList_Click()
If noClick Then Exit Sub
byClick = True
loadDocNomenk
End Sub

Private Sub ckStartDate_Click()
If ckStartDate.value = 1 Then
    tbStartDate.Enabled = True
Else
    tbStartDate.Enabled = False
End If

End Sub

Function getNextDocNum() As Boolean
Dim il As Long, l As Long
getNextDocNum = False

il = Right$(Format(Now, "yymmdd\0\0"), 7) + 200001  ' чтобы не путались с заказами
sql = "select * from System"
Set tbSystem = myOpenRecordSet("##149", sql, dbOpenForwardOnly)
If tbSystem Is Nothing Then Exit Function
tbSystem.Edit
l = tbSystem!lastDocNum + 1
If l < il Then l = il
tbSystem!lastDocNum = l
tbSystem.update
tbSystem.Close
numDoc = l

getNextDocNum = True
End Function

Private Sub cmAdd_Click()
Dim str As String ', intNum As Integer, il As Long, l As Long
Dim strNow As String, DateFromNum As String, dNow As Date
 
On Error GoTo ER1
wrkDefault.BeginTrans   ' начало транзакции

If Not getNextDocNum() Then GoTo ER1

Set tbDocs = myOpenRecordSet("##141", "select * from sDocs", dbOpenForwardOnly) 'dbOpenForwardOnly)
If tbDocs Is Nothing Then GoTo ER1

tbDocs.AddNew
tbDocs!numDoc = numDoc
tbDocs!numExt = numExt
tbDocs!xDate = Now
tbDocs!sourId = -1002
tbDocs!destId = -7 'инвентаризация
If Regim = "fromCeh" Then
    numExt = 0 ' виртуальные накладные(зарезервир-е предметы)
    tbDocs!Note = Ceh(cehId)
Else
    numExt = 254
End If
tbDocs.update
tbDocs.Close

wrkDefault.CommitTrans  ' подтверждение транзакции

loadDocs "add" ' не загружать все док-ты
Exit Sub
ER1:
errorCodAndMsg ("Добавить расход")
wrkDefault.rollback
End Sub

Private Sub cmAdd2_Click()
sProducts.Regim = "fromDocs"
sProducts.Show vbModal
loadDocNomenk
On Error Resume Next
Grid.SetFocus
End Sub

Private Sub cmBay_Click()
lbZakaz.Clear

sql = "SELECT sDMCrez.numDoc " & _
"FROM sDMCrez INNER JOIN BayOrders ON sDMCrez.numDoc = BayOrders.numOrder " & _
"GROUP BY sDMCrez.numDoc HAVING (((Max(sDMCrez.curQuant))>0));"
Set tbDocs = myOpenRecordSet("##346", sql, dbOpenForwardOnly)
If tbDocs Is Nothing Then Exit Sub

While Not tbDocs.EOF
    lbZakaz.AddItem tbDocs!numDoc
    cmBay.Enabled = False
    tbDocs.MoveNext
Wend
tbDocs.Close

cmOk.Caption = "Списать"
viewZakazi

End Sub

Private Sub cmCancel_Click()
frZakaz.Visible = False
cmOk.Enabled = False
cmOrder.Enabled = True
cmBay.Enabled = True
End Sub

Function sqlDeficitToNNQQ(sklad As String, Optional reg As String = "") As Boolean
Static I As Integer, s As Single

sqlDeficitToNNQQ = False
    
Set tbDMC = myOpenRecordSet("##142", sql, dbOpenDynaset)
If tbDMC Is Nothing Then Exit Function

ReDim NN(0): ReDim QQ(0): I = 0

While Not tbDMC.EOF
    I = I + 1
    gNomNom = tbDMC!nomNom
    s = PrihodRashod("+", skladId) - PrihodRashod("-", skladId) 'Ф. остатки по складу
    ReDim Preserve NN(I): NN(I) = gNomNom
    ReDim Preserve QQ(I)
    If reg = "bay" Then
        QQ(I) = tbDMC!curQuant * tbDMC!perList
        s = Round((s - QQ(I)) / tbDMC!perList, 2)
    Else
        QQ(I) = tbDMC!quantity
        s = Round(s - QQ(I), 2)
    End If
    
    If s < 0 Then
      If MsgBox("Дефицит товара '" & tbDMC!nomNom & "' по подразделению '" & _
      sklad & "'" & " составит (" & s & "), продолжить?", _
      vbOKCancel Or vbDefaultButton2, "Подтвердите") = vbCancel Then Exit Function
    End If
    tbDMC.MoveNext
Wend
tbDMC.Close


'If i = 0 And (reg = "" Or reg = "end") Then
If I = 0 Then
    MsgBox "У этой накладной нет Предметов", , "Cписание невозможно!"
Else
    sqlDeficitToNNQQ = True
End If

End Function

Private Sub cmClose_Click()
Dim s  As Single, str  As String, I As Integer

If Not lockSklad Then Exit Sub
  
If reservNoNeed Then str = "mov" Else str = "rez"
sql = "SELECT nomNom, quantity FROM sDMC" & str & " WHERE (((numDoc)=" & numDoc & "));"
If Not sqlDeficitToNNQQ(Grid.TextMatrix(mousRow, dcSour)) Then GoTo EN1

wrkDefault.BeginTrans
  
'удаляем предметы из sDMCrez/mov
sql = "DELETE From sDMC" & str & " WHERE (((numDoc)=" & numDoc & "));"
If myExecute("##340", sql) <> 0 Then GoTo ER1

'виртуальная накладная становится настоящей

sql = "UPDATE sDocs SET [xDate] = '" & Format(Now(), "yyyy-mm-dd hh:nn:ss") & _
"', numExt = 254 WHERE (((numExt)=0) AND ((numDoc)=" & numDoc & "));"
'MsgBox sql
If myExecute("##341", sql) <> 0 Then GoTo ER1

' добавляем предметы в sDMC
'Set tbDMC = myOpenRecordSet("##142", "select * from sDMC", dbOpenForwardOnly)
'If tbDMC Is Nothing Then GoTo ER1
'tbDMC.index = "NomDoc"
numExt = 254
For I = 1 To UBound(NN)
    gNomNom = NN(I)
    If Not sProducts.nomenkToDMC(QQ(I), "noLock") Then GoTo ER2
Next I
'tbDMC.Close
  
wrkDefault.CommitTrans

ckCeh.value = 0

loadDocs "single"
MsgBox "Списание прошло успешно!", , ""

GoTo EN1

ER2: tbDMC.Close
ER1: wrkDefault.rollback
EN1: lockSklad "un"

End Sub
'$odbc15$
Private Sub cmDel_Click()
Dim str As String, isZakaz As Integer, count As Integer
Dim s As Single, sId As Integer, dId As Integer, I  As Integer

If MsgBox("Удалить накладную № '" & getStrDocExtNum(numDoc, numExt) & _
"', Вы уверены?", vbYesNo Or vbDefaultButton2, "Подтвердите удаление") _
= vbNo Then Grid.SetFocus: Exit Sub


If numExt = 0 Then
  
  wrkDefault.BeginTrans   ' начало транзакции
  
  sql = "DELETE FROM sDocs WHERE (((numDoc)=" & numDoc & ") AND ((numExt)=0));"
  If myExecute("##337", sql) <> 0 Then GoTo ER1
   
  If reservNoNeed Then ' межскладская выписанная из цеха -  не резервировалась
    sql = "DELETE FROM sDMCmov WHERE (((numDoc)=" & numDoc & "));"
  Else
    sql = "DELETE FROM sDMCrez WHERE (((numDoc)=" & numDoc & "));"
  End If
  I = myExecute("##338", sql, 0)
  If I = 0 Or I = -1 Then GoTo CC ' документ м.б. и пустым
ER1:
  wrkDefault.rollback ' отммена транзакции
  Exit Sub
End If

wrkDefault.BeginTrans   ' начало транзакции

'слить все этапы по ВСЕМ предметам в текущий
sql = "UPDATE xEtapByIzdelia SET prevQuant = 0 WHERE (((numOrder)=" & numDoc & "));"
myExecute "##334", sql, 0 'если есть
sql = "UPDATE xEtapByNomenk SET prevQuant = 0 WHERE (((numOrder)=" & numDoc & "));"
myExecute "##335", sql, 0 'если есть


sql = "SELECT sDocs.sourId, sDocs.destId, sDocs.numDoc, sDocs.numExt " & _
"From sDocs " & _
"WHERE (((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & "));"
If Not byErrSqlGetValues("W##180", sql, sId, dId) Then myBase.Close: End

If Not (sId < -1000 And dId < -1000) Then ' для межскладских не корректируем
    
'   корректируем остатки
    sql = "SELECT sDMC.nomNom, sDMC.quant  FROM sDMC " & _
    "WHERE (((sDMC.numDoc)=" & numDoc & ") AND ((sDMC.numExt)=" & numExt & "));"
    Set tbDMC = myOpenRecordSet("##109", sql, dbOpenDynaset)
    If tbDMC Is Nothing Then GoTo ERR1 '


    If Not tbDMC.BOF Then 'м.не быть для накл-й не под заказ******************
'        Set tbNomenk = myOpenRecordSet("##163", "select * from sGuideNomenk", dbOpenForwardOnly)
'        If tbNomenk Is Nothing Then GoTo ERR1
    
        While Not tbDMC.EOF
 '           tbNomenk.index = "PrimaryKey"
 '           tbNomenk.Seek "=", tbDMC!nomNom
            cErr = "116" '##116
            sql = "UPDATE sGuideNomenk SET nowOstatki = [nowOstatki]+" & _
            tbDMC!quant & " WHERE (((sGuideNomenk.nomNom)='" & tbDMC!nomNom & "'));"
            If myExecute("##418", sql) <> 0 Then GoTo ERR1
  '          If tbNomenk.NoMatch Then GoTo ERR1
  '          tbNomenk.Edit
  '          tbNomenk!nowOstatki = Round(tbNomenk!nowOstatki + tbDMC!quant)
  '          tbNomenk.Update
            tbDMC.MoveNext
        Wend
'        tbNomenk.Close
    End If '  **********************************************************
    tbDMC.Close
End If
    
'удаление док-та (а также соотв. записей из ДМЦ - т.к. разрешено каскадное удаление)
sql = "DELETE From sDocs WHERE (((sDocs.numDoc)=" & numDoc & _
") AND ((sDocs.numExt)=" & numExt & "));"
'MsgBox sql
If Not myExecute("##121.1", sql) = 0 Then GoTo ERR1

sql = "DELETE From sDMC WHERE (((numDoc)=" & numDoc & _
") AND ((numExt)=" & numExt & "));"
'MsgBox sql
I = myExecute("##121.2", sql, 0) '
If I <> 0 And I <> -1 Then GoTo ERR1 ' документ м.б. и пустым

CC:
wrkDefault.CommitTrans  ' подтверждение транзакции
gridRowDel
mousRow = quantity
Grid_EnterCell
If quantity > 0 Then
    loadDocNomenk
Else
    Grid2.Visible = False
    laGrid2.Visible = False
End If
EN2:
Grid.SetFocus
Exit Sub

ERR1:
wrkDefault.rollback ' отммена транзакции
On Error Resume Next
tbDMC.Close
tbDocs.Close
tbNomenk.Close
MsgBox "Удаление не прошло. Сообщите администратору", , "Error - " & cErr
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmLoad_Click()
loadDocs
If quantity > 0 Then cmClose.Enabled = True
End Sub

Private Sub cmOk_Click()
Dim id As Integer

cmOk.Enabled = False
numDoc = lbZakaz.Text
gNzak = numDoc
frZakaz.Visible = False

If cmOrder.Enabled Then ' кто погашен тот и вызвал
    bayZakazi
Else
    workZakazi
End If
End Sub

'$odbc15$
Sub bayZakazi()
Dim I As Integer, middle As Integer
cmBay.Enabled = True

If Not lockSklad Then Exit Sub
    sql = "SELECT sDMCrez.nomNom, sDMCrez.quantity, sDMCrez.curQuant, " & _
"sGuideNomenk.perList " & _
"FROM sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom " & _
"WHERE (((sDMCrez.numDoc)=" & numDoc & " AND (sDMCrez.curQuant)>0 ));"

skladId = -1001
If Not sqlDeficitToNNQQ(lbInside.List(0), "bay") Then GoTo EN1

wrkDefault.BeginTrans

numExt = getNextNumExt()
sql = "insert into sDocs (numDoc, numExt, xDate, sourId, destId) values (" _
    & numDoc & ", " & numExt & ", now(), -1001, -8) "

Debug.Print sql
If myExecute("##367.0", sql, 0) <> 0 Then GoTo ER3

For I = 1 To UBound(NN)
    gNomNom = NN(I)
    If Not sProducts.nomenkToDMC(QQ(I), "noLock") Then GoTo ER2
Next I

'tbDMC.Close
'tbDocs.Close

sql = "UPDATE sDMCrez SET curQuant = 0 WHERE (((numDoc)=" & numDoc & "));"
If myExecute("##367", sql) <> 0 Then GoTo ER3

wrkDefault.CommitTrans
EN1: lockSklad "un"

loadDocs "bayZakaz"

Exit Sub

ER2:
'tbDMC.Close
ER1:
'tbDocs.Close
ER3:
wrkDefault.rollback
lockSklad "un"
End Sub

Sub workZakazi()
'сделать, чтобы по "Ok" сразу не появлялись пока  накладные(пустые), а только
'по закрытию Nakladna.frm
cmOrder.Enabled = True
sql = "SELECT Orders.CehId From Orders WHERE (((Orders.numOrder)=" & numDoc & "));"
If Not byErrSqlGetValues("##98", sql, cehId) Then Exit Sub
 
cmDel.Enabled = True

Nakladna.Regim = "toNaklad"
Nakladna.Show vbModal

End Sub

Private Sub cmOrder_Click()
Dim I As Integer

lbZakaz.Clear
getNakladnieList "Buh"
For I = 1 To UBound(tmpL)
    If tmpL(I) > 0 Then ' незакрытые предметы
        lbZakaz.AddItem tmpL(I)
        cmOrder.Enabled = False
'        bilo = True
    End If
Next I
cmOk.Caption = "Ok"
viewZakazi
End Sub

Sub viewZakazi()
If cmOrder.Enabled And cmBay.Enabled Then
    MsgBox "Нет заказов с предметами для создания накладной!", , "Предупреждение"
Else
    frZakaz.Visible = True
    lbZakaz.SetFocus
    frZakaz.ZOrder
End If
End Sub


Private Sub cmPrint_Click()
Nakladna.Regim = ""
Nakladna.Show vbModal

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static value

If KeyCode = vbKeyF7 Then
AA: value = InputBox("Введите номер накладной или заказа.", "Поиск", value)
    If value = "" Then Exit Sub
    If Not IsNumeric(value) Then
        MsgBox "Номер должен быть числом"
        GoTo AA
    End If
    numDoc = value
    loadDocs "docsFind"
End If
End Sub

Private Sub Form_Load()
Dim I As Integer, j As Integer
oldHeight = Me.Height
oldWidth = Me.Width

If Regim = "fromCeh" Then
    Me.Caption = "Выписанные накладные не под заказы"
    laPeriod.Visible = False
    ckStartDate.Visible = False
    tbStartDate.Visible = False
    laPo.Visible = False
    ckEndDate.Visible = False
    tbEndDate.Visible = False
    
    cmOrder.Visible = False
    cmBay.Visible = False
    ckCeh.Visible = False
End If

tbStartDate.Text = Format(DateAdd("d", -14, curDate), "dd/mm/yy")
tbEndDate.Text = Format(curDate, "dd/mm/yy")
If otlad = "otlaD" Then ckStartDate.value = 1

Grid.FormatString = "|<Дата|<№ Док-та|<Окуда|<Куда|<Примечание"
Grid.ColWidth(dcSourId) = 0
Grid.ColWidth(dcDate) = 800
Grid.ColWidth(dcNumDoc) = 915
Grid.ColWidth(dcSour) = 1100
Grid.ColWidth(dcDest) = 1520
Grid.ColWidth(dcNote) = 1100

sql = "SELECT sGuideSource.sourceId, sGuideSource.sourceName From sGuideSource " & _
"WHERE (((sGuideSource.sourceId)<0)) ORDER BY sGuideSource.sourceId DESC;"
Set table = myOpenRecordSet("##95", sql, dbOpenDynaset)
If table Is Nothing Then myBase.Close: End
ReDim insideId(0): ReDim statiaId(0): I = 0: j = 0
While Not table.EOF
    If table!sourceId < -1000 Then 'внутр подр-я
        If Regim = "fromCeh" And table!sourceId < -1002 Then GoTo NX1
        lbInside.AddItem table!SourceName
        ReDim Preserve insideId(I)
        insideId(I) = table!sourceId
        I = I + 1
    Else
        If Regim = "fromCeh" And j > 4 Then GoTo NX1
        lbStatia.AddItem table!SourceName
        ReDim Preserve statiaId(j)
        statiaId(j) = table!sourceId
        j = j + 1
    End If
NX1: table.MoveNext
Wend

table.Close
lbInside.Height = lbInside.Height + 195 * (lbInside.ListCount - 1)
lbStatia.Height = lbStatia.Height + 195 * (lbStatia.ListCount - 1)
quantity = 0

If Regim = "fromCeh" Then
    Timer1.Interval = 300
    Timer1.Enabled = True ' loadDocs
End If



isLoad = True
End Sub

'reg ="","single","add","docsFind","bayZakaz",numExtO(№ накл. для орезных ном-р)
Sub loadDocs(Optional reg As String = "")
Dim strWhere As String, moveWhere As String, I As Integer, str As String

 prevRow = -1
 Grid.Visible = False
 If reg = "" Then
'    str = strWhereByStEndDateBox(Me)
    str = getWhereByDateBoxes(Me, "sDocs.xDate", begDate)
    If Regim = "fromCeh" Then
        strWhere = "((sDocs.numExt) = 0) AND ((sDocs.Note)='" & Ceh(cehId) & "')" 'вирт. накладные
    ElseIf ckCeh.value = 1 Then
        strWhere = "((sDocs.numExt) = 0)" 'вирт. накладные
    Else
        strWhere = "((sDocs.numExt) > 0 AND (sDocs.numExt)< 255)" 'расходные накладные
    End If
    If str <> "" Then strWhere = "(" & str & ") AND " & strWhere
 ElseIf reg = "docsFind" Then ' поиск расходных накладных
    strWhere = "((sDocs.numExt) >0 AND  (sDocs.numExt) < 255 AND " & _
    "(sDocs.numDoc)=" & numDoc & ") OR (sDocs.Note) Like '%" & numDoc & "%'"
 ElseIf reg = "bayZakaz" Then ' к заказу продаж
    strWhere = "(sDocs.numDoc)=" & numDoc
 ElseIf reg = "single" Or reg = "add" Then '
    strWhere = "(sDocs.numDoc)=" & numDoc & " AND (sDocs.numExt)=" & numExt
 ElseIf IsNumeric(reg) Then ' до 3х накладных к Prior заказу
    I = InStr(reg, " ")
    strWhere = "": str = ""
    moveWhere = Mid$(reg, I + 1)
    If moveWhere <> "" Then moveWhere = "((sDocs.numDoc)=" & moveWhere & ")  OR "
    I = Left$(reg, I - 1)
    If numExt > 0 Then strWhere = "(sDocs.numExt)=" & numExt
    If I > 0 Then
        str = "(sDocs.numExt)=" & I
        If strWhere <> "" Then str = " OR " & str
    End If
    strWhere = "(sDocs.numDoc)=" & numDoc & " AND (" & strWhere & str & ")"
 Else
    Exit Sub
 End If
 
 Me.MousePointer = flexHourglass
 If reg <> "add" Then
    gridIzLoad = False
    quantity = 0
    clearGrid Grid
 End If
  sql = "SELECT sDocs.xDate, sDocs.Note, sDocs.numDoc, sDocs.numExt, " & _
 "sGuideSource.sourceName, GS.sourceName AS destName, sDocs.sourId, sDocs.destId " & _
 "FROM (sDocs INNER JOIN sGuideSource ON sDocs.sourId = sGuideSource.sourceId) " & _
 "INNER JOIN sGuideSource AS GS ON sDocs.destId = GS.sourceId " & _
 "WHERE (" & moveWhere & "(" & strWhere & ")) ORDER BY sDocs.xDate;"
 Debug.Print sql
 
 'MsgBox sql
 
 Set tbDocs = myOpenRecordSet("##176", sql, dbOpenForwardOnly)
 If tbDocs Is Nothing Then End
 If Not tbDocs.BOF Then
 While Not tbDocs.EOF
    Grid.AddItem ""
    quantity = quantity + 1
    Grid.TextMatrix(quantity, dcSourId) = tbDocs!sourId
    LoadDate Grid, quantity, dcDate, tbDocs!xDate, "dd.mm.yy"
    Grid.TextMatrix(quantity, dcNumDoc) = getStrDocExtNum(tbDocs!numDoc, tbDocs!numExt)
    Grid.TextMatrix(quantity, dcNote) = tbDocs!Note
    If tbDocs!Note = "toCeh" Then ' не активирована
            Grid.row = quantity
            Grid.col = dcNumDoc
            Grid.CellForeColor = 200
    End If
    Grid.TextMatrix(quantity, dcSour) = tbDocs!SourceName
    Grid.TextMatrix(quantity, dcDest) = tbDocs!destName

    tbDocs.MoveNext
  Wend
End If
tbDocs.Close
rowViem quantity, Grid
Grid.Visible = True
If quantity > 0 Then
    If reg <> "add" Or quantity = 1 Then Grid.removeItem quantity + 1
    Grid.row = quantity
    Grid.col = 1
    gridIzLoad = True '
    Grid.col = 2      'вызов loadDocNomenk
    Grid.SetFocus
    cmDel.Enabled = True
Else
    cmDel.Enabled = False
    Grid2.Visible = False
    laGrid2.Visible = False
End If

gridIzLoad = True

Me.MousePointer = flexDefault
    
End Sub

Function getGridColSour() As String
getGridColSour = Grid.TextMatrix(mousRow, dcSour)
End Function


Sub getDocExtNomFromStr(nom As String)
Dim I As Integer
I = InStr(nom, "/")
If I = 0 Then
    numDoc = nom
    If Regim = "fromCeh" Or ckCeh.value = 1 Then
        numExt = 0
    Else
        numExt = 254
    End If
Else
    numDoc = Left$(nom, I - 1)
    numExt = Mid$(nom, I + 1)
End If
End Sub

Sub gridRowDel()
    quantity = quantity - 1
    If quantity = 0 Then
        clearGridRow Grid, mousRow
    Else
        Grid.removeItem mousRow
    End If

End Sub

Sub lbHide()
lbGroup.Visible = False
lbInside.Visible = False
lbStatia.Visible = False
tbMobile.Visible = False
Grid.Enabled = True
Grid.SetFocus
Grid_EnterCell
End Sub

'Optional reg As String = ""
Function loadDocNomenk() As Boolean
Dim il As Long, str As String, str2 As String, q As Single, I As Integer
Dim msgOst As String, r As Single, b As Single

loadDocNomenk = True ' не надо отката - пока
msgOst = ""
Me.MousePointer = flexHourglass
Grid2.Visible = False

gDocDate = Grid.TextMatrix(mousRow, dcDate)
laGrid2.Caption = "Предметы по накладной №" & getStrDocExtNum(numDoc, numExt)

Grid2.FormatString = "|<Номер|<Описание|<Ед.измерения|Кол-во|Целых"
Grid2.ColWidth(0) = 0
Grid2.ColWidth(dnNomNom) = 0 '960
Grid2.ColWidth(dnEdIzm) = 630
Grid2.ColWidth(dnQuant) = 645 '810
    Grid2.ColWidth(dnNomName) = 3435 + 960
    Grid2.ColWidth(dnLists) = 0

quantity2 = 0
clearGrid Grid2

If numExt = 0 And sDocs.reservNoNeed Then ' межскладская выписанная из цеха -  не резервируем
  sql = "SELECT sGuideNomenk.ed_Izmer, sGuideNomenk.nomName, sDMCmov.nomNom, " & _
  "sGuideNomenk.Size, sGuideNomenk.cod, sGuideNomenk.ed_Izmer2, " & _
  "sGuideNomenk.perList, sDMCmov.numDoc, sDMCmov.quantity as quant " & _
  "FROM sGuideNomenk INNER JOIN sDMCmov ON sGuideNomenk.nomNom = sDMCmov.nomNom  " & _
  "WHERE (((sDMCmov.numDoc) = " & numDoc & ")) ORDER BY sDMCmov.nomNom;"
ElseIf numExt = 0 Then ' из цеха со Склад1
  sql = "SELECT sGuideNomenk.ed_Izmer, sGuideNomenk.nomName, sDMCrez.nomNom, " & _
  "sGuideNomenk.Size, sGuideNomenk.cod, sGuideNomenk.ed_Izmer2, " & _
  "sGuideNomenk.perList, sDMCrez.numDoc, sDMCrez.quantity as quant " & _
  "FROM sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom  " & _
  "WHERE (((sDMCrez.numDoc) = " & numDoc & ")) ORDER BY sDMCrez.nomNom;"
Else
  sql = "SELECT sGuideNomenk.ed_Izmer, sGuideNomenk.nomName, sDMC.nomNom, " & _
  "sGuideNomenk.Size, sGuideNomenk.cod, sGuideNomenk.ed_Izmer2, " & _
  "sGuideNomenk.perList, sDMC.numDoc, sDMC.numExt, sDMC.quant " & _
  "FROM sGuideNomenk INNER JOIN sDMC ON sGuideNomenk.nomNom = sDMC.nomNom  " & _
  "WHERE (((sDMC.numDoc) = " & numDoc & " AND (sDMC.numExt) = " & numExt & _
  ")) ORDER BY sGuideNomenk.nomNom;"
End If
'MsgBox sql
   
   If byClick Then
        bilo = (ckPerList.value = 1)
   Else
        noClick = True
        bilo = isIntMove() ' если слева или справа целый склад
   End If


Set tbNomenk = myOpenRecordSet("##118", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then Exit Function
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    quantity2 = quantity2 + 1
    Grid2.TextMatrix(quantity2, dnNomNom) = tbNomenk!nomNom
    Grid2.TextMatrix(quantity2, dnNomName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid2.TextMatrix(quantity2, dnQuant) = Round(tbNomenk!quant, 2)
    
   If bilo Then
        If Not byClick Then ckPerList.value = 1
        Grid2.TextMatrix(quantity2, dnEdIzm) = tbNomenk!ed_Izmer2
        If IsNumeric(tbNomenk!perList) Then ' dnLists
          If tbNomenk!perList > 0.01 Then Grid2.TextMatrix(quantity2, dnQuant) _
                                = Round(tbNomenk!quant / tbNomenk!perList, 2)
        End If
    Else
        If Not byClick Then ckPerList.value = 0
        Grid2.TextMatrix(quantity2, dnEdIzm) = tbNomenk!ed_Izmer
    End If
   
    Grid2.AddItem ""
    tbNomenk.MoveNext
  Wend
End If
tbNomenk.Close
    
noClick = False
byClick = False

If quantity2 > 0 Then
    Grid2.removeItem quantity2 + 1
End If

Grid2.Visible = True
Me.MousePointer = flexDefault
End Function
'дает True, если слева или справа целый склад - т.е. целое перемещение
Function isIntMove(Optional msg As String = "") As Boolean
Dim str As String

isIntMove = (Grid.TextMatrix(mousRow, dcDest) = lbInside.List(0) Or _
Grid.TextMatrix(mousRow, dcSour) = lbInside.List(0))

End Function
'дает True, если не нужно резервирование - т.е. со скл. Обрезков или межскладская
Function reservNoNeed() As Boolean

reservNoNeed = (skladId = -1002 Or (skladId = -1001 And _
              Grid.TextMatrix(mousRow, dcDest) = lbInside.List(1)))
End Function

Function valueToDocsField(myErrCod As String, value As String, field As String) As Boolean
sql = "UPDATE sDocs  SET sDocs." & field & "=" & value & _
" WHERE (((sDocs.numDoc)=" & numDoc & " AND (sDocs.numExt)=" & numExt & "));"
'MsgBox sql
valueToDocsField = False
If myExecute(myErrCod, sql) = 0 Then valueToDocsField = True
End Function

Private Sub Form_Resize()
Dim h As Integer, w As Integer
If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + h

Grid2.Height = Grid2.Height + h

cmLoad.Top = cmLoad.Top + h
cmAdd.Top = cmAdd.Top + h
cmOrder.Top = cmOrder.Top + h
cmDel.Top = cmDel.Top + h
cmAdd2.Top = cmAdd2.Top + h
cmExit.Top = cmExit.Top + h
cmPrint.Top = cmPrint.Top + h
End Sub

Private Sub Form_Unload(Cancel As Integer)
isLoad = False
End Sub

Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If mousRow = 0 Then
    If mousCol = dcDate Then
        SortCol Grid, mousCol, "date"
    Else
        SortCol Grid, mousCol
    End If
    Grid.row = 1    ' только чтобы снять выделение
End If
Grid_EnterCell

End Sub

Private Sub Grid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Dim date1 As Date, date2 As Date ' в 2 х местах
Dim date1S, date2S As String

date1S = sortGrid.TextMatrix(Row1, mousCol)
date2S = sortGrid.TextMatrix(Row2, mousCol)

If Not IsDate(date1S) Then
    Cmp = -1
    GoTo CC:
ElseIf Not IsDate(date2S) Then
    Cmp = 1
    GoTo CC:
End If

date1 = date1S
date2 = date2S
If date1 > date2 Then
    Cmp = 1
ElseIf date1 < date2 Then
    Cmp = -1
Else
    Cmp = 0
End If
CC:
If trigger Then Cmp = -Cmp


End Sub

Private Sub Grid_DblClick()
Dim id As Integer

If mousRow = 0 Then Exit Sub
If Grid.CellBackColor <> &H88FF88 Then Exit Sub
If mousCol = dcSour Then
    listBoxInGridCell lbInside, Grid, "select"
ElseIf mousCol = dcDest Then
     If quantity2 > 0 Or skladId = -1002 Then
        listBoxInGridCell lbStatia, Grid, "select"
    Else
        listBoxInGridCell lbGroup, Grid
    End If
Else
    tbMobile.MaxLength = 50
    textBoxInGridCell tbMobile, Grid
End If
    
End Sub

Private Sub Grid_EnterCell()
Dim str As String

If quantity = 0 Or Not gridIzLoad Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col
getDocExtNomFromStr (Grid.TextMatrix(mousRow, dcNumDoc)) ' numDoc numExt
If numExt = 254 Or numExt = 0 Then skladId = Grid.TextMatrix(mousRow, dcSourId)

str = numDoc

cmAdd2.Enabled = False
If numExt = 254 Or numExt = 0 Then cmAdd2.Enabled = True

If prevRow <> mousRow And gridIzLoad Then
    prevRow = mousRow
    loadDocNomenk
End If
If mousCol = 0 Then Exit Sub
bilo = False
If ckCeh.value = 1 Then
    Grid.CellBackColor = vbButtonFace
    Exit Sub
End If
If mousCol = dcNote And numExt = 0 Then GoTo AA
If (mousCol > dcNumDoc And (numExt = 254 Or numExt = 0) And _
(quantity2 = 0 Or bilo)) Or mousCol = dcNote Then
    Grid.CellBackColor = &H88FF88
Else
AA: Grid.CellBackColor = vbYellow
End If
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid_DblClick

End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
Else
    If Grid.RowSel - Grid.row > 0 Then
        cmPrint.Caption = "Печать накладных"
    Else
        cmPrint.Caption = "Печать накладной"
    End If
End If

End Sub


Private Sub Grid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid2.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid2.ColWidth(Grid2.MouseCol)

End Sub

Private Sub lbGroup_DblClick()
If lbGroup.ListIndex = 0 Then
    listBoxInGridCell lbStatia, Grid, "select"
Else
    listBoxInGridCell lbInside, Grid, "select"
End If
    
End Sub

Private Sub lbGroup_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbGroup_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub lbInside_DblClick()
Dim str As String, str2 As String

If mousCol = dcSour Then
    str = insideId(lbInside.ListIndex)
    If numExt = 0 Then ' из цеха0
      If str = -1001 Then ' Целые
      Else ' -1002 Обрезки
        sql = "UPDATE sDocs SET sourId = -1002, destId = -3 " & _
        "WHERE (((numDoc)=" & numDoc & ") AND ((numExt)=0));"
        If myExecute("##339", sql) = 0 Then
            Grid.TextMatrix(mousRow, dcDest) = lbStatia.List(2)
            GoTo EN1
        End If
        GoTo EN2
      End If
    End If
    str = "sourId"
    str2 = Grid.TextMatrix(mousRow, dcDest)
Else
    str = "destId"
    str2 = Grid.TextMatrix(mousRow, dcSour)
    If lbInside.Text = lbInside.List(0) And str2 = lbInside.List(1) Then
        MsgBox "Недопустимое перемещение!", , "Предупреждение"
        Exit Sub
    End If
End If
If lbInside.Text = str2 Then
    MsgBox "В колонках 'Откуда' и 'Куда' недопустимы одинаковые значения", , "Предупреждение"
    Exit Sub
End If

If valueToDocsField("##96", insideId(lbInside.ListIndex), str) Then
EN1: Grid.Text = lbInside.Text
    If mousCol = dcSour Then Grid.TextMatrix(mousRow, dcSourId) = -1001 - lbInside.ListIndex
End If
EN2:   lbHide
End Sub


Private Sub lbInside_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbInside_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub lbInside_LostFocus()
'If Regim = "addForZakaz" And lbInside.Visible Then lbInside.SetFocus
If lbInside.Visible And Nakladna.Regim = "" Then lbInside.SetFocus
End Sub

Private Sub lbStatia_DblClick()
Dim str As String

If mousCol = dcSour Then
    str = "sourId"
Else
    str = "destId"
End If
If valueToDocsField("##168", statiaId(lbStatia.ListIndex), str) Then _
                Grid.Text = lbStatia.Text
lbHide

End Sub

Private Sub lbStatia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbStatia_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Private Sub lbZakaz_Click()
cmOk.Enabled = True
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If valueToDocsField("##178", "'" & tbMobile.Text & "'", "Note") Then _
                Grid.Text = tbMobile.Text
    lbHide
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
loadDocs
End Sub
