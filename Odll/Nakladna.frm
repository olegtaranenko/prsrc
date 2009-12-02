VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Nakladna 
   BackColor       =   &H8000000A&
   Caption         =   "Предметы "
   ClientHeight    =   5532
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5532
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmCheckout 
      Caption         =   "Выписать все"
      Height          =   315
      Left            =   2760
      TabIndex        =   31
      Top             =   5160
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.TextBox tbPageSize 
      Height          =   288
      Left            =   8280
      TabIndex        =   30
      Text            =   "30"
      Top             =   5160
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CommandButton cmClose 
      Caption         =   "Списать"
      Height          =   315
      Left            =   2820
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmSostav 
      Caption         =   "Состав изд."
      Height          =   315
      Left            =   4380
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Frame gridFrame 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   2760
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin MSFlexGridLib.MSFlexGrid Grid4 
         Height          =   615
         Left            =   60
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9335
         _ExtentY        =   1080
         _Version        =   393216
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Label laGrid4 
         Alignment       =   2  'Center
         Caption         =   "laGid4"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   5295
      End
   End
   Begin VB.TextBox tbMobile2 
      Height          =   315
      Left            =   780
      TabIndex        =   14
      Text            =   "tbMobile2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4095
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   9615
      _ExtentX        =   16955
      _ExtentY        =   7218
      _Version        =   393216
      AllowBigSelection=   0   'False
      MergeCells      =   3
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "Печать"
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   5160
      Width           =   915
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   8880
      TabIndex        =   2
      Top             =   5160
      Width           =   795
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Exel"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1935
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   9615
      _ExtentX        =   16955
      _ExtentY        =   3408
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Label laPageSize 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Caption         =   "Размер страницы"
      Height          =   252
      Left            =   6600
      TabIndex        =   29
      Top             =   5160
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label laPageOf 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Страница 1 из 3"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   6888
      TabIndex        =   28
      Top             =   60
      Visible         =   0   'False
      Width           =   1284
   End
   Begin VB.Label laDest 
      Caption         =   "laDest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   26
      Top             =   2820
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Label laSours 
      Caption         =   "laSours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   25
      Top             =   2820
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label laDocNum 
      Caption         =   "laDocNum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   24
      Top             =   2460
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label laKomu 
      Caption         =   "Кому:"
      Height          =   195
      Index           =   1
      Left            =   2940
      TabIndex        =   23
      Top             =   2820
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label laOt 
      BackColor       =   &H8000000A&
      Caption         =   "От:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   22
      Top             =   480
      Width           =   255
   End
   Begin VB.Label laNakl 
      Caption         =   "Накладная №"
      Height          =   195
      Left            =   180
      TabIndex        =   21
      Top             =   2460
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label laDate 
      BackColor       =   &H8000000A&
      Height          =   195
      Left            =   8340
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label laSignatura 
      BackColor       =   &H8000000A&
      Height          =   200
      Left            =   7020
      TabIndex        =   13
      Top             =   360
      Width           =   2472
   End
   Begin VB.Label laPerson 
      BackColor       =   &H8000000A&
      Caption         =   "Исполнитель:"
      Height          =   195
      Left            =   5700
      TabIndex        =   12
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label laFirm 
      BackColor       =   &H8000000A&
      Caption         =   "laFirm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      TabIndex        =   11
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label laPlatel 
      BackColor       =   &H8000000A&
      Caption         =   "Плательщик:"
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label laDest 
      BackColor       =   &H8000000A&
      Caption         =   "laDest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label laKomu 
      BackColor       =   &H8000000A&
      Caption         =   "Кому:"
      Height          =   195
      Index           =   0
      Left            =   3060
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Label laSours 
      BackColor       =   &H8000000A&
      Caption         =   "laSours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   540
      TabIndex        =   7
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label laOt 
      Caption         =   "От:"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   2820
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label laDocNum 
      BackColor       =   &H8000000A&
      Caption         =   "laDocNum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label laTitle 
      BackColor       =   &H8000000A&
      Caption         =   "Накладная №"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   60
      Width           =   1095
   End
End
Attribute VB_Name = "Nakladna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim quantity2 As Long
Public mousCol2 As Long
Public mousRow2 As Long
Public Regim As String
Public docDate As Date

Dim secondNaklad As String, beSUO As Boolean ' была листовая ном-ра

Dim lastPageSizePx As Long ' размер последней страницы в пикселях
Dim pageSizePx As Long ' размер страницы в пикселях
Dim pageNum As Integer ' количество страниц
Dim pageRows As Integer ' к-во строк на странице
Dim lastPageRows As Integer ' к-во строк на последней странице


Const nkNomNom = 1
Const nkNomName = 2
Const nkEdIzm = 3
Const nkTreb = 4
Const nkClos = 5
Const nkEtap = 6
Const nkEClos = 7
Const nkQuant = 8
Const nkIntEdIzm = 9
Const nkIntQuant = 10

Private Sub cmCheckout_Click()
Dim I As Long

    If MsgBox("Вы действительно хотите выписать все оставшиеся позиции?", _
        vbOKCancel Or vbDefaultButton2, "Подтвердите") = vbCancel Then
        Exit Sub
    End If

    For I = 1 To Grid2(0).rows - 1
        Dim nomRequest As Double: nomRequest = 0
        Dim nomCheckouted As Double: nomCheckouted = 0
        Dim treb As Integer, already As Integer
        
        If QQ2(0) = 0 Then 'нет этапа
            treb = nkTreb
            already = nkClos
        Else
            treb = nkEtap
            already = nkEClos
        End If
        
        If IsNumeric(Grid2(0).TextMatrix(I, treb)) Then _
            nomRequest = CDbl(Grid2(0).TextMatrix(I, treb))
            
        If IsNumeric(Grid2(0).TextMatrix(I, already)) Then _
            nomCheckouted = Grid2(0).TextMatrix(I, already)
            
        If nomRequest > nomCheckouted Then
            Dim quant As Double
            quant = nomRequest - nomCheckouted
            Grid2(0).TextMatrix(I, nkQuant) = CStr(quant)
            sql = "UPDATE sDMCrez SET curQuant = " & quant & _
                " WHERE numDoc = " & numDoc & " AND nomNom = '" & _
                Grid2(0).TextMatrix(I, nkNomNom) & "'"
            myExecute "##363.1", sql
                

            
        End If
    Next I
End Sub

'$odbc15$
Private Sub cmClose_Click()
Dim I As Integer, j As Integer, NN2() As String, k As Integer
Dim numExtO As Integer, id As Integer, l As Long, s As Double
Dim mov As Double, moveNum As String, per As Double, str As String, str2 As String

If Not lockSklad Then Exit Sub

ReDim NN(0): ReDim NN2(0): ReDim NN3(0): ReDim QQ(0): ReDim QQ2(0): ReDim QQ3(0)
I = 0: j = 0: moveNum = ""
For l = 1 To quantity2
  str = Grid2(0).TextMatrix(l, nkQuant)
  If IsNumeric(str) Then
    mov = 0
    gNomNom = Grid2(0).TextMatrix(l, nkNomNom)
    If Grid2(0).TextMatrix(l, 0) = "" Then 'штучная
        I = I + 1: ReDim Preserve NN(I): ReDim Preserve QQ(I)
        NN(I) = gNomNom: QQ(I) = str
        skladId = -1001: GoTo AA
    Else ' обрезную списываем со склада обрезков
        j = j + 1: ReDim Preserve NN2(j)
        ReDim Preserve QQ2(j): ReDim Preserve QQ3(j)
        NN2(j) = gNomNom: QQ2(j) = str: QQ3(j) = 0
        skladId = -1002
        If IsNumeric(Grid2(0).TextMatrix(l, nkIntQuant)) Then 'нужна межскладская
            sql = "SELECT perList, ed_izmer2 from sGuideNomenk " & _
            "WHERE (((sGuideNomenk.nomNom)='" & gNomNom & "'));"
            If Not byErrSqlGetValues("##366", sql, per, str2) Then Exit Sub
            
            QQ3(j) = per * Grid2(0).TextMatrix(l, nkIntQuant)
            s = PrihodRashod("+", -1001) - PrihodRashod("-", -1001) 'Ф. остатки по складу
            s = Round(s - QQ3(j), 2)
            If s < 0 Then
              If MsgBox("Дефицит товара '" & gNomNom & "' в факт. остатках " & _
              "в подразделении '" & sDocs.lbInside.List(0) & _
              "' составит (" & Round(s / per, 2) & " " & str2 & "), продолжить?", _
              vbOKCancel Or vbDefaultButton2, "Подтвердите") = vbCancel Then
                lockSklad "un"
                GoTo EN1
              End If
            End If
            mov = QQ3(j)
            moveNum = "yes"
        End If
        
AA:     s = PrihodRashod("+", skladId) - PrihodRashod("-", skladId) 'Ф. остатки по складу
        s = Round(mov + s - str, 2)
        If s < 0 Then
          If MsgBox("Дефицит товара '" & gNomNom & "' в факт. остатках " & _
          "в подразделении '" & sDocs.lbInside.List(-1001 - skladId) & _
          "' составит (" & s & "), продолжить?", _
          vbOKCancel Or vbDefaultButton2, "Подтвердите") = vbCancel Then
            lockSklad "un"
            GoTo EN1
          End If
        End If
    End If
  End If
Next l
If I + j = 0 Then
  If Regim = "predmeti" Then
    MsgBox "Проставте количества для тех позиций, которые Вы хотите списать.", , "Нечего списывать!"
  Else
    MsgBox ""
  End If
EN1: Grid2(0).SetFocus
    Exit Sub
End If

If cehId = 0 Then
  id = -6 'заказчик без работы
ElseIf cehId = 3 Then           '$ceh$
  id = getStatiaId("Пр-во SUB") '
Else
  id = -cehId
End If

wrkDefault.BeginTrans

Set tbDocs = myOpenRecordSet("##357", "select * from sDocs", dbOpenForwardOnly) 'dbOpenForwardOnly)
If tbDocs Is Nothing Then GoTo ER3

'Set tbDMC = myOpenRecordSet("##348", "select * from sDMC", dbOpenForwardOnly)
'If tbDMC Is Nothing Then GoTo ER1
'tbDMC.index = "NomDoc"
If docDate <> #12:00:00 AM# Then
    tmpDate = docDate
Else
    tmpDate = Now
End If

numExtO = 0
If j > 0 Then numExtO = getNextNumExt()
If moveNum = "yes" Then
    If Not sDocs.getNextDocNum() Then GoTo ER1
    moveNum = numDoc
    numExt = 254
    tbDocs.AddNew
    tbDocs!numDoc = numDoc
    tbDocs!numExt = numExt
    tbDocs!xDate = tmpDate
    tbDocs!Note = gNzak & "/" & numExtO
    tbDocs!sourId = -1001
    tbDocs!destId = -1002
    tbDocs.update
    For k = 1 To j
      If QQ3(k) > 0 Then
        gNomNom = NN2(k)
        If Not sProducts.nomenkToDMC(QQ3(k), "noLock") Then GoTo ER2
      End If
    Next k
    tmpDate = DateAdd("S", 1, tmpDate)
End If

numDoc = gNzak
If j > 0 Then
  tbDocs.AddNew
  tbDocs!numDoc = numDoc
  tbDocs!numExt = numExtO
  tbDocs!xDate = tmpDate
  tbDocs!Note = moveNum
  tbDocs!sourId = -1002
  tbDocs!destId = id
  tbDocs.update
  For k = 1 To j
    gNomNom = NN2(k): numExt = numExtO
    If Not sProducts.nomenkToDMC(QQ2(k), "noLock") Then GoTo ER2
    If Not clrCehQuant Then GoTo ER2
  Next k
  tmpDate = DateAdd("S", 1, tmpDate)
End If

numExt = 0
If I > 0 Then
  numExt = getNextNumExt()
  tbDocs.AddNew
  tbDocs!numDoc = numDoc
  tbDocs!numExt = numExt
  tbDocs!xDate = tmpDate
  tbDocs!sourId = -1001
  tbDocs!destId = id
  tbDocs.update
  For k = 1 To I
    gNomNom = NN(k)
    If Not sProducts.nomenkToDMC(QQ(k), "noLock") Then GoTo ER2
    If Not clrCehQuant Then GoTo ER2
  Next k
End If
'tbDMC.Close
tbDocs.Close
wrkDefault.CommitTrans
lockSklad "un"
Unload Me
sDocs.loadDocs CStr(numExtO) & " " & moveNum ' показать 1-3 накладных

Exit Sub

ER2:
'tbDMC.Close
ER1:
tbDocs.Close
ER3:
wrkDefault.rollback
lockSklad "un"
MsgBox "Списание не прошло. Сообщите администратору.", , "Error - " & cErr

End Sub

Function getStatiaId(name As String) As Integer

sql = "SELECT sourceId from sGuideSource WHERE (((sourceId)<0 And " & _
"(sGuideSource.sourceId)>-1000) AND ((sGuideSource.sourceName)='" & name & "'));"
If byErrSqlGetValues("W##387", sql, getStatiaId) Then
    If getStatiaId = 0 Then GoTo AA
Else
AA: getStatiaId = -5 'Петровские маст.
End If
End Function

Function clrCehQuant() As Boolean
clrCehQuant = False
sql = "UPDATE sDMCrez SET curQuant = 0, intQuant = 0 " & _
"WHERE (((sDMCrez.numDoc)=" & numDoc & ") AND ((sDMCrez.nomNom)='" & gNomNom & "'));"
If myExecute("##365", sql) = 0 Then clrCehQuant = True
End Function

Private Sub cmExel_Click()
Dim str As String
str = laDocNum(0).Caption
GridToExcel Grid2(0), "Накладная № " & str
End Sub

Private Sub cmExit_Click()

Unload Me

End Sub

Private Sub cmPrint_Click()
Dim I As Integer

laDate.Visible = True
laDate.Caption = Format(Now(), "dd.mm.yy hh:nn")

For I = 1 To pageNum
    setPage (I)
    Me.PrintForm
Next I

End Sub

Private Sub cmSostav_Click()
Me.MousePointer = flexHourglass
sql = "SELECT xPredmetyByIzdelia.prId, xPredmetyByIzdelia.prExt, " & _
"sGuideProducts.prName, sGuideProducts.prDescript FROM sGuideProducts INNER JOIN xPredmetyByIzdelia " & _
"ON sGuideProducts.prId = xPredmetyByIzdelia.prId " & _
"WHERE (((xPredmetyByIzdelia.numOrder)=" & gNzak & "));"
Set tbProduct = myOpenRecordSet("##332", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Sub

If tbProduct.BOF Then
    MsgBox "У этого заказа нет готовых изделий.", , ""
    GoTo EN1
End If

Report.Regim = "fromCehNaklad"
Report.Show vbModal

EN1:
tbProduct.Close
Me.MousePointer = flexDefault
End Sub

Private Sub Command1_Click()
MsgBox "ColWidth(nkTreb)=" & Grid2(1).ColWidth(nkTreb)
End Sub

Private Sub Form_Load()
Dim str As String
Dim notBay As Long, I As Long, delta As Long

oldHeight = Me.Height
oldWidth = Me.Width

laTitle.Caption = "    Заказ №"
Grid2(0).MergeRow(0) = True
Grid2(0).FormatString = "|<Номер|<Описание|<Ед.измерения|Затребовано по заказу|Отпущено|Затребовано по этапу|Отпущено по этапу|кол-во|Перемещение|Перемещение"
Grid2(0).ColWidth(0) = 0
Grid2(0).ColWidth(nkNomNom) = 945
Grid2(0).ColWidth(nkNomName) = 4500 '5265
Grid2(0).ColWidth(nkEdIzm) = 645
Grid2(0).ColWidth(nkQuant) = 735
'размеры некот. колонок обнуляются также и по результататм загрузки (см. loadToGrid)
If Regim = "" Then 'печать накладной
    Grid2(0).ColWidth(nkTreb) = 0
    Grid2(0).ColWidth(nkClos) = 0
Else
    Grid2(0).ColWidth(nkTreb) = 630
    Grid2(0).ColWidth(nkClos) = 855
End If
Grid2(0).ColWidth(nkEtap) = 780
Grid2(0).ColWidth(nkEClos) = 765
Grid2(0).ColWidth(nkIntEdIzm) = 700
Grid2(0).ColWidth(nkIntQuant) = 700
Grid2(1).FormatString = "|<Номер|<Описание|<Ед.измерения|Затребовано по " & _
"заказу|Отпущено|Затребовано по этапу|Отпущено по этапу|кол-во|Перемещение|Перемещение"
Grid2(1).ColWidth(0) = 0
Grid2(1).ColWidth(nkNomNom) = 945
Grid2(1).ColWidth(nkNomName) = 5265
Grid2(1).ColWidth(nkEdIzm) = 645
Grid2(1).ColWidth(nkQuant) = 735
Grid2(1).ColWidth(nkEtap) = 0 '780
Grid2(1).ColWidth(nkEClos) = 0 '765
Grid2(1).ColWidth(nkTreb) = 0
Grid2(1).ColWidth(nkClos) = 0
Grid2(1).ColWidth(nkIntEdIzm) = 0
Grid2(1).ColWidth(nkIntQuant) = 0

cmExit.Caption = "Выход"
secondNaklad = ""
If Regim = "" Then ' для распечатки
    laTitle.Caption = "Накладная №"
    laDocNum(0).Caption = getStrDocExtNum(numDoc, numExt)
    If sDocs.Grid.RowSel - sDocs.Grid.row > 0 Then
        secondNaklad = sDocs.Grid.TextMatrix(sDocs.Grid.RowSel, dcNumDoc)
        laDocNum(1).Caption = secondNaklad
        cmExel.Visible = False
    End If
    laPageSize.Visible = True
    laPageOf.Visible = True
    tbPageSize.Visible = True
    tbPageSize.Text = getEffectiveSetting("gCfgOrderPageSize", 35)
    
ElseIf Regim = "predmeti" Then ' в цеху
    Me.Caption = "Предметы к заказу."
    cmSostav.Visible = True
    cmCheckout.Visible = True
    GoTo BB
ElseIf Regim = "toNaklad" Then
    cmClose.Visible = True
    cmPrint.Visible = False
    cmExel.Visible = False
    laPerson.Visible = False
    laSignatura.Visible = False
    Me.Caption = "Списание предметов заказа."
    cmExit.Caption = "Отмена"
BB: laDocNum(0).Caption = numDoc
    laOt(0).Visible = False
    laSours(0).Visible = False
    laKomu(0).Visible = False
End If

MousePointer = flexHourglass

laPlatel.Visible = False
laFirm.Visible = False
If Regim = "" And numExt = 0 Then
        laFirm.Visible = True
        laFirm.Caption = "(несписанная из " & Ceh(cehId) & ")"
ElseIf numExt <> 254 Then  'к заказу
    sql = "SELECT Orders.numOrder, GuideFirms.Name " & _
    "FROM GuideFirms INNER JOIN Orders ON GuideFirms.FirmId = Orders.FirmId " & _
    "WHERE (((Orders.numOrder)=" & numDoc & "));"
    notBay = 0
    byErrSqlGetValues "W##170", sql, notBay, str
    If notBay > 0 Then GoTo AA ' заказ не на продажу
    
    sql = "SELECT BayGuideFirms.Name " & _
    "FROM BayGuideFirms INNER JOIN BayOrders ON BayGuideFirms.FirmId = " & _
    "BayOrders.FirmId WHERE (((BayOrders.numOrder)=" & numDoc & "));"
    If byErrSqlGetValues("##353", sql, str) Then
AA:     laPlatel.Visible = True
        laFirm.Visible = True
        laFirm.Caption = str
    End If
End If

loadToGrid 0

    paginateResult
setPage (1)

MousePointer = flexDefault
End Sub



Sub paginateResult()
Dim stdPageRows As Integer

    stdPageRows = CInt(getEffectiveSetting("gCfgOrderPageSize", 35))
    
    If quantity2 < stdPageRows Then
        pageNum = 1
        lastPageRows = quantity2
    Else
        pageRows = stdPageRows
        pageNum = quantity2 \ stdPageRows
        lastPageRows = quantity2 - pageNum * pageRows
        If lastPageRows > 0 Then
            pageNum = pageNum + 1
        End If
    End If
    
    If pageNum = 1 Then
        pageSizePx = getPageSize(quantity2)
        lastPageSizePx = pageSizePx
    Else
        pageSizePx = getPageSize(pageRows)
        lastPageSizePx = getPageSize(lastPageRows)
    End If
End Sub

Function getPageSize(ByVal rows As Integer) As Long
    getPageSize = 350 + (Grid2(0).CellHeight + 25) * rows
End Function

Sub setPage(pageNo As Integer)
Dim I As Long, delta As Long
    If pageNo = pageNum Then
        I = lastPageSizePx
    Else
        I = pageSizePx
    End If
    If secondNaklad = "" Then
        delta = I - Grid2(0).Height
        Me.Height = Me.Height + delta
    Else ' в одной форме сразу 2 накладных
    'laDocNum(1) и всю 2ю половину надо сдвинуть немного ниже Grid2(0)
        Grid2(0).Height = I
        delta = Grid2(0).Top + I - laDocNum(1).Top + 200
        laDocNum(1).Top = laDocNum(1).Top + delta
        laDocNum(1).Visible = True
        Grid2(1).Top = Grid2(1).Top + delta
        Grid2(1).Visible = True
        laNakl.Top = laNakl.Top + delta
        laNakl.Visible = True
        laOt(1).Top = laOt(1).Top + delta
        laOt(1).Visible = True
        laSours(1).Top = laSours(1).Top + delta
        laSours(1).Visible = True
        laKomu(1).Top = laKomu(1).Top + delta
        laKomu(1).Visible = True
        laDest(1).Top = laDest(1).Top + delta
        laDest(1).Visible = True
        
        sDocs.getDocExtNomFromStr secondNaklad: loadToGrid 1
        I = getPageSize(quantity2) '2я таблица
        delta = delta + I - Grid2(1).Height ' изменений 1й таблицы + изм-е 2ой
        Grid2(1).Height = I
        oldHeight = Me.Height + delta ' Me.Height=oldHeight в Resize
        
        cmPrint.Top = cmPrint.Top + delta
        cmExit.Top = cmExit.Top + delta
    End If
    
    
    Grid2(0).TopRow = (pageNo - 1) * pageRows + 1
    
    If pageNum > 1 Then
        laPageOf.Caption = "Страница " & pageNo & " из " & pageNum
        laPageOf.Visible = True
    Else
        laPageOf.Visible = False
    End If
    
End Sub

'ind=1 м.б. только при Regim = ""
Sub loadToGrid(ind As Integer)
Dim I As Integer, s As Double, s2 As Double, str As String, str2 As String


ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): QQ2(0) = 0: ReDim QQ3(0)

If Regim = "toNaklad" Then
  laSours(ind).Caption = ""
  If cehId = 3 Then                    '$ceh$
    laDest(ind).Caption = "Пр-во SUB"  '
  Else
    laDest(ind).Caption = sDocs.lbStatia.List(cehId - 1)
  End If
  If Not sProducts.zakazNomenkToNNQQ Then Exit Sub
ElseIf Regim = "" Then
  sql = "SELECT sGuideSource.sourceName, sGuideDest.sourceName AS destName " & _
  "FROM sGuideSource AS sGuideDest INNER JOIN (sGuideSource INNER JOIN " & _
  "sDocs ON sGuideSource.sourceId = sDocs.sourId) ON sGuideDest.sourceId = sDocs.destId " & _
  "WHERE (((sDocs.numDoc)=" & numDoc & ") AND ((sDocs.numExt)=" & numExt & "));"
  'Debug.Print sql
  If byErrSqlGetValues("##172", sql, str, str2) Then
      laSours(ind).Caption = str
      laDest(ind).Caption = str2
  End If
  If sDocs.reservNoNeed Then str = "mov" Else str = "rez"
  If numExt = 0 Then ' витуальные накладные из цеха
    sql = "SELECT nomNom, quantity as quant  FROM sDMC" & str & _
    " WHERE (((numDoc)=" & numDoc & "));"
  Else
    sql = "SELECT nomNom, quant FROM sDMC " & _
    "WHERE (((numDoc)=" & numDoc & ") AND ((numExt)=" & numExt & "));"
  End If
  Set tbDMC = myOpenRecordSet("##318", sql, dbOpenForwardOnly)
  If Not tbDMC Is Nothing Then
    I = 0
    While Not tbDMC.EOF
        I = I + 1
        ReDim Preserve NN(I): NN(I) = tbDMC!nomNom
        ReDim Preserve QQ(I): QQ(I) = tbDMC!quant
        tbDMC.MoveNext
    Wend
    tbDMC.Close
  End If
ElseIf Regim = "predmeti" Then
  laSours(0).Caption = "Склад1"
  If cehId = 1 Then
      laDest(ind).Caption = "Пр-во YAG"
  ElseIf cehId = 2 Then
      laDest(ind).Caption = "Пр-во CO2"
  ElseIf cehId = 3 Then                 '$$ceh
      laDest(ind).Caption = "Пр-во SUB" '
  End If
  If Not sProducts.zakazNomenkToNNQQ Then GoTo EN1
End If



Grid2(ind).Visible = False
quantity2 = 0
clearGrid Grid2(ind)
beSUO = False
'Set tbNomenk = myOpenRecordSet("##129", "select * from sGuideNomenk", dbOpenForwardOnly)
'If tbNomenk Is Nothing Then GoTo EN1
'tbNomenk.index = "PrimaryKey"
For I = 1 To UBound(NN)
'    tbNomenk.Seek "=", NN(i)
    sql = "SELECT nomName, ed_Izmer, perList, Size, ed_Izmer2, cod " & _
    "from sGuideNomenk WHERE (((nomNom)='" & NN(I) & "'));"
    Set tbNomenk = myOpenRecordSet("##129", sql, dbOpenForwardOnly)
'    If Not tbNomenk.NoMatch Then
    If Not tbNomenk.BOF Then
        quantity2 = quantity2 + 1
        'Grid2(ind).TextMatrix(quantity2, 0) = tbNomenk!obrez
        If tbNomenk!perlist > 1 Then Grid2(ind).TextMatrix(quantity2, 0) = "Да" 'обрезная
        Grid2(ind).TextMatrix(quantity2, nkNomNom) = NN(I)
        Grid2(ind).TextMatrix(quantity2, nkNomName) = tbNomenk!cod & " " & _
            tbNomenk!nomName & " " & tbNomenk!Size
        Grid2(ind).TextMatrix(quantity2, nkEdIzm) = tbNomenk!ed_Izmer
        If Regim = "" Then
            If laDest(ind).Caption = "Продажа" Then
              Grid2(ind).TextMatrix(quantity2, nkEdIzm) = tbNomenk!ed_Izmer2
              Grid2(ind).TextMatrix(quantity2, nkQuant) = Round(QQ(I) / tbNomenk!perlist, 2)
            Else
              Grid2(ind).TextMatrix(quantity2, nkQuant) = Round(QQ(I), 2)
            End If
        Else ' "toNaklad"
            Grid2(ind).TextMatrix(quantity2, nkTreb) = Round(QQ(I), 2)
            Grid2(ind).TextMatrix(quantity2, nkEtap) = Round(QQ2(I) - QQ3(I), 2)
            
            sql = "SELECT Sum(quant) AS Sum_quant From sDMC WHERE " & _
            "(((sDMC.numDoc)=" & numDoc & ") AND ((sDMC.nomNom)='" & NN(I) & "'));"
            If byErrSqlGetValues("##194", sql, s) Then
                Grid2(ind).TextMatrix(quantity2, nkClos) = Round(s, 2)
                Grid2(ind).TextMatrix(quantity2, nkEClos) = Round(s - QQ3(I), 2)
            End If
            If Regim <> "" Then
              If tbNomenk!perlist <> 1 Then 'для обрезной доп. колонка для целых
                beSUO = True
                Grid2(ind).TextMatrix(quantity2, nkIntEdIzm) = tbNomenk!ed_Izmer2
              End If
              s = 0: s2 = 0
              sql = "SELECT curQuant, intQuant from sDMCrez " & _
              "WHERE (((numDoc)=" & gNzak & ") AND ((nomNom)='" & NN(I) & "'));"
              byErrSqlGetValues "##362", sql, s, s2
              If s > 0 Then _
                Grid2(ind).TextMatrix(quantity2, nkQuant) = Round(s, 2)
              If s2 > 0 Then _
                Grid2(ind).TextMatrix(quantity2, nkIntQuant) = s2
            End If
        End If
        Grid2(ind).AddItem ""
    End If
Next I
tbNomenk.Close
If quantity2 > 0 Then
    Grid2(ind).removeItem quantity2 + 1
End If
If ind = 0 Then
  If QQ2(0) = 0 Then  'если не этапный убираем колонки
    Grid2(0).ColWidth(nkEtap) = 0
    Grid2(0).ColWidth(nkEClos) = 0
  ElseIf dostup <> "" Then ' для менеджеров оставляем в любом случае
    Grid2(0).ColWidth(nkTreb) = 0
    Grid2(0).ColWidth(nkClos) = 0
  End If
  If Not beSUO Then
    Grid2(0).ColWidth(nkIntEdIzm) = 0
    Grid2(0).ColWidth(nkIntQuant) = 0
  End If
End If
Dim sum  As Long
sum = 0
For I = 0 To Grid2(ind).Cols - 1
    sum = sum + Grid2(ind).ColWidth(I)
Next I
sum = sum + 680 '650
If sum < 8300 Then sum = 8300
Me.Width = sum
Grid2(ind).col = nkQuant
EN1:
Grid2(ind).Visible = True


End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer
    If Me.WindowState = vbMinimized Then Exit Sub
    On Error Resume Next
    
    If secondNaklad <> "" Then Me.Height = oldHeight: Me.Top = 0
    
    h = Me.Height - oldHeight
    oldHeight = Me.Height
    w = Me.Width - oldWidth
    oldWidth = Me.Width
    Grid2(0).Height = Grid2(0).Height + h
    Grid2(0).Width = Grid2(0).Width + w
    Grid2(1).Width = Grid2(1).Width + w
    
    cmPrint.Top = cmPrint.Top + h
    cmExel.Top = cmExel.Top + h
    cmSostav.Top = cmSostav.Top + h
    cmClose.Top = cmClose.Top + h
    cmClose.left = cmClose.left + w
    cmCheckout.Top = cmPrint.Top
    cmCheckout.left = cmPrint.left + cmPrint.Width + 150
    cmExit.Top = cmExit.Top + h
    cmExit.left = cmExit.left + w
    laDate.left = laDate.left + w
    tbPageSize.left = cmExit.left - tbPageSize.Width - 150
    tbPageSize.Top = cmExit.Top
    laPageSize.left = tbPageSize.left - laPageSize.Width - 50
    laPageSize.Top = cmExit.Top
    laPageOf.left = laDate.left - 100 - laPageOf.Width
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Regim = "" 'нужно для lbInside_LostFocus
End Sub

Private Sub Grid2_DblClick(Index As Integer)
Dim str As String, per As Double, ed_Izmer As String

If Grid2(Index).CellBackColor = &H88FF88 Then '****************************
 
If mousCol2 = nkIntQuant Then
    str = Grid2(Index).TextMatrix(mousRow2, nkQuant)
    If Not IsNumeric(str) Then GoTo AA
    If CDbl(str) <= 0 Then
AA:     MsgBox "Сначала проставте значение в колонке 'кол-во'", , "Предупреждение"
        Exit Sub
    End If
End If

Me.MousePointer = flexHourglass
 
tmpStr = "Фактические остатки по подразделению '"
gNomNom = Grid2(Index).TextMatrix(mousRow2, fnNomNom)
 
If Grid2(Index).TextMatrix(mousRow2, 0) = "" Or mousCol2 = nkIntQuant Then 'штучная
    sql = "SELECT perList, ed_Izmer2 From sGuideNomenk " & _
    "WHERE (((nomNom)='" & gNomNom & "'));"
    byErrSqlGetValues "##364", sql, per, ed_Izmer
    If per < 0.01 Then per = 1
    
    laGrid4.Caption = tmpStr & sDocs.lbInside.List(0) & "'"
    skladId = -1001
Else ' обрезная
    per = 1
    ed_Izmer = Grid2(Index).TextMatrix(mousRow2, fnEdIzm)
    
    laGrid4.Caption = tmpStr & sDocs.lbInside.List(1) & "'"
    skladId = -1002
End If
 clearGrid Grid4
 Grid4.FormatString = "|<Номер|<Описание|<Ед.измерения|Остатки"
 Grid4.ColWidth(0) = 0
 Grid4.ColWidth(1) = 870
 Grid4.ColWidth(2) = 2745
 Grid4.ColWidth(3) = 645
 Grid4.ColWidth(4) = 900

 Grid4.TextMatrix(1, 1) = Grid2(Index).TextMatrix(mousRow2, fnNomNom)
 Grid4.TextMatrix(1, 2) = Grid2(Index).TextMatrix(mousRow2, fnNomName)
 Grid4.TextMatrix(1, 3) = ed_Izmer
 Grid4.TextMatrix(1, 4) = Round((PrihodRashod("+", skladId) - _
                    PrihodRashod("-", skladId)) / per, 2) 'Ф. остатки по складу

 Grid4.Visible = True
EN1:
 Me.MousePointer = flexDefault
 gridFrame.Visible = True
     textBoxInGridCell tbMobile2, Grid2(0)
End If '*************************************************************
End Sub

Private Sub Grid2_EnterCell(Index As Integer)
Dim t As Double, s As Double
If Index > 0 Then Exit Sub
mousRow2 = Grid2(Index).row
mousCol2 = Grid2(Index).col
If quantity2 = 0 Or Regim <> "predmeti" Or dostup = "" Then Exit Sub

If mousCol2 = nkQuant Or (mousCol2 = nkIntQuant And _
Grid2(Index).TextMatrix(mousRow2, 0) <> "") Then
    Grid2(Index).CellBackColor = &H88FF88
Else
    Grid2(Index).CellBackColor = vbYellow
End If

End Sub

Private Sub Grid2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid2_DblClick (Index)
End Sub

Private Sub Grid2_LeaveCell(Index As Integer)
Grid2(Index).CellBackColor = Grid2(Index).BackColor

End Sub

Private Sub Grid2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid2(Index).MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid2(Index).ColWidth(Grid2(Index).MouseCol)

End Sub

Sub lbHide2()
tbMobile2.Visible = False
gridFrame.Visible = False
Grid2(0).Enabled = True
Grid2(0).SetFocus
Grid2_EnterCell 0
End Sub

Private Sub Grid4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid4.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid4.ColWidth(Grid4.MouseCol)

End Sub

Private Sub tbMobile2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim delta As Double, quant As Double, s As Double, str As String

If KeyCode = vbKeyReturn Then
  
  If mousCol2 = nkQuant Then
    If QQ2(0) = 0 Then 'нет этапа
        quant = Grid2(0).TextMatrix(mousRow2, nkTreb)
        quant = Round(quant - Grid2(0).TextMatrix(mousRow2, nkClos), 2)
    Else
        quant = Grid2(0).TextMatrix(mousRow2, nkEtap)
        quant = Round(quant - Grid2(0).TextMatrix(mousRow2, nkEClos), 2)
    End If
    
    If Not isNumericTbox(tbMobile2, 0, quant) Then Exit Sub
    
    quant = Round(tbMobile2.Text, 2)
    str = "cur"
Else 'nkIntQuant
    If Not isNumericTbox(tbMobile2, 0) Then Exit Sub
    quant = Round(tbMobile2.Text, 0)
    If quant <> CDbl(tbMobile2.Text) Then
        MsgBox "Количество должно быть целым!", , "Предупреждение"
        Exit Sub
    End If
    str = "int"
End If
sql = "UPDATE sDMCrez SET " & str & "Quant = " & quant & _
" WHERE (((numDoc)=" & numDoc & ") AND ((nomNom)='" & _
Grid2(0).TextMatrix(mousRow2, nkNomNom) & "'));"
If myExecute("##363", sql) = 0 Then
    If quant = 0 Then
        Grid2(0).TextMatrix(mousRow2, mousCol2) = ""
    Else
        Grid2(0).TextMatrix(mousRow2, mousCol2) = quant
    End If
End If
lbHide2
Grid2(0).SetFocus

ElseIf KeyCode = vbKeyEscape Then
NN:  lbHide2
End If

End Sub

Private Sub tbPageSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tbPageSize.Text) Then
            getEffectiveSetting "gCfgOrderPageSize", tbPageSize.Text
            saveFileSettings appCfgFile, appSettings
            paginateResult
            setPage (1)
        End If
    End If
End Sub
