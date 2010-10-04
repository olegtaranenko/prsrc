VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form sProducts 
   BackColor       =   &H8000000A&
   Caption         =   "Формирование накладной"
   ClientHeight    =   6384
   ClientLeft      =   60
   ClientTop       =   1740
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6384
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ckList 
      BackColor       =   &H8000000A&
      Caption         =   "в целых"
      Height          =   252
      Left            =   120
      TabIndex        =   29
      Top             =   5970
      Width           =   972
   End
   Begin VB.Frame splLeftH 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H8000000A&
      Height          =   80
      Left            =   2400
      MousePointer    =   7  'Size N S
      TabIndex        =   27
      Top             =   2760
      Width           =   3480
   End
   Begin VB.Frame splRightV 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H8000000A&
      Height          =   1764
      Left            =   6840
      MousePointer    =   9  'Size W E
      TabIndex        =   26
      Top             =   0
      Width           =   84
   End
   Begin VB.ComboBox cbInside 
      Height          =   315
      Left            =   6180
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmExel2 
      Caption         =   "Печать в Exel"
      Height          =   315
      Left            =   8640
      TabIndex        =   22
      Top             =   5940
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   5940
      Top             =   600
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   7620
      TabIndex        =   20
      Text            =   "tbMobile"
      Top             =   1380
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Exel"
      Height          =   315
      Left            =   2520
      TabIndex        =   16
      Top             =   5940
      Visible         =   0   'False
      Width           =   1284
   End
   Begin VB.Frame gridFrame 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   3180
      TabIndex        =   12
      Top             =   3420
      Visible         =   0   'False
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid Grid4 
         Height          =   1455
         Left            =   60
         TabIndex        =   13
         Top             =   300
         Visible         =   0   'False
         Width           =   7215
         _ExtentX        =   12721
         _ExtentY        =   2561
         _Version        =   393216
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Label laGrid4 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "laGrid4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   120
         TabIndex        =   15
         Top             =   60
         Width           =   7212
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Если остатки позволяют, введите треб. кол-во изделий и нажмите <Enter>, иначе - <ESC>.."
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   1740
         Width           =   7215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2835
      Left            =   2400
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   4464
      _ExtentX        =   7874
      _ExtentY        =   4995
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   2475
      Left            =   2400
      TabIndex        =   11
      Top             =   300
      Visible         =   0   'False
      Width           =   4464
      _ExtentX        =   7874
      _ExtentY        =   4360
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmHide 
      Caption         =   "Скрыть выд."
      Enabled         =   0   'False
      Height          =   315
      Left            =   5700
      TabIndex        =   10
      Top             =   5940
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.OptionButton opNomenk 
      BackColor       =   &H8000000A&
      Caption         =   "Выбор номенклатуры"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   480
      Width           =   2175
   End
   Begin VB.OptionButton opProduct 
      BackColor       =   &H8000000A&
      Caption         =   "Выбор готовых изделий"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   180
      Width           =   2235
   End
   Begin VB.TextBox tbQuant 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   5
      Top             =   5940
      Width           =   735
   End
   Begin VB.CommandButton cmSel 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   5940
      Width           =   915
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   11040
      TabIndex        =   3
      Top             =   5940
      Width           =   795
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   4980
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   8784
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2895
      Left            =   6960
      TabIndex        =   21
      Top             =   2940
      Width           =   4872
      _ExtentX        =   8594
      _ExtentY        =   5101
      _Version        =   393216
      AllowBigSelection=   0   'False
      HighLight       =   0
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid5 
      Height          =   2412
      Left            =   6960
      TabIndex        =   17
      Top             =   300
      Width           =   4872
      _ExtentX        =   8594
      _ExtentY        =   4255
      _Version        =   393216
      AllowBigSelection=   0   'False
      MergeCells      =   2
      AllowUserResizing=   1
   End
   Begin VB.Frame splLeftV 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H8000000A&
      Height          =   5892
      Left            =   2300
      MousePointer    =   9  'Size W E
      TabIndex        =   25
      Top             =   0
      Width           =   80
   End
   Begin VB.Frame splRightH 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H8000000A&
      Height          =   80
      Left            =   6000
      MousePointer    =   7  'Size N S
      TabIndex        =   28
      Top             =   2640
      Width           =   5640
   End
   Begin VB.Label laGrid2 
      BackColor       =   &H8000000A&
      Caption         =   "Номенклатурный состав предметов:"
      Height          =   195
      Left            =   7320
      TabIndex        =   2
      Top             =   2730
      Width           =   3495
   End
   Begin VB.Label laGrid1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "laGrid1"
      Height          =   192
      Left            =   2400
      TabIndex        =   1
      Top             =   60
      Width           =   516
   End
   Begin VB.Label laGrid 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "laGrid"
      Height          =   192
      Left            =   2400
      TabIndex        =   24
      Top             =   2820
      Width           =   432
   End
   Begin VB.Label laGrid5 
      BackColor       =   &H8000000A&
      Caption         =   "Состав предметов:"
      Height          =   195
      Left            =   7800
      TabIndex        =   18
      Top             =   45
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.Label laBegin 
      BackColor       =   &H8000000A&
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   4512
      Left            =   2400
      TabIndex        =   7
      Top             =   780
      Width           =   3552
   End
   Begin VB.Label laQuant 
      BackColor       =   &H8000000A&
      Caption         =   "изделий"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4860
      TabIndex        =   6
      Top             =   5985
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Menu mnContext 
      Caption         =   "Из состава предметов"
      Visible         =   0   'False
      Begin VB.Menu mnDel 
         Caption         =   "Удалить"
      End
      Begin VB.Menu mnOnfly 
         Caption         =   "Преобразовать в изделия"
      End
   End
   Begin VB.Menu mnContext2 
      Caption         =   "Из накладной и из цеха"
      Visible         =   0   'False
      Begin VB.Menu mnDel2 
         Caption         =   "Удалить"
      End
   End
End
Attribute VB_Name = "sProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'#Const OTLAD = True 'независимая работа  формы(иначе закомментировать и стартовать с Main)

Public isLoad As Boolean
Public Regim As String
Public mousCol2 As Long
Public mousRow2 As Long
Public mousCol3 As Long
Public mousRow3 As Long
Public zakazano As Double
Public FO As Double ' ФО
Public convertToIzdelie As Boolean
Public orderRate As Double
Public asWhole As Integer
Public idWerk As Integer


Private selectNomenkFlag As Boolean
Private isCtrlDown As Boolean

Dim mousCol4 As Long, mousRow4 As Long
Dim msgBilo As Boolean, biloG3Enter_Cell As Boolean

Const groupColor1 = &HBBFFBB ' только не vbBottonFace
Const groupColor2 = &HBBBBFF '
Dim grColor As Long
Dim mousCol As Long, mousRow As Long
Dim mousCol5 As Long
Public mousRow5 As Long

Dim quantity  As Long, quantity2 As Long, quantity3 As Long
Public quantity5 As Long
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim tvVes As Double, gridVes As Double, grid2Ves As Double 'веса горизонт. размеров

Dim tbKlass As Recordset
Dim typeId As Integer
Dim beShift As Boolean
'список изделий или номенклатур(Grid3)
Const gpNN = 0
Const gpName = 1
Const gpSize = 2
Const gpDescript = 3
Const gpId = 4 ' спрятан

'номенклатура по группе или изделию(Grid)
Const nkNomer = 1
Const nkName = 2
Const nkEdIzm = 3
Const nkQuant = 4
Const nkCurOstat = 5
Const nkDostup = 6
Const nkQuantPricise = 7

'Grid4
Const frNomNom = 1
Const frNomName = 2
Const frEdIzm = 3
Const frOstat = 4

Dim buntColumn As Integer
Dim Dragging As Boolean
Dim DraggingX As Single, DraggingY As Single

Sub nomenkToNNQQ(pQuant As Double, eQuant As Double, prQuantity As Double)
Dim J As Integer, leng As Integer

    leng = UBound(NN)

    For J = 1 To leng
        If NN(J) = tbNomenk!Nomnom Then
            QQ(J) = QQ(J) + pQuant * tbNomenk!quantity
            If eQuant > 0 Then _
                QQ2(J) = QQ2(J) + eQuant * tbNomenk!quantity
            If prQuantity > 0 Then _
                QQ3(J) = QQ3(J) + prQuantity * tbNomenk!quantity
            Exit Sub
        End If
    Next J
    leng = leng + 1
    ReDim Preserve NN(leng): NN(leng) = tbNomenk!Nomnom
    ReDim Preserve QQ(leng): QQ(leng) = pQuant * tbNomenk!quantity
    ReDim Preserve QQ2(leng): QQ2(leng) = eQuant * tbNomenk!quantity
    ReDim Preserve QQ3(leng): QQ3(leng) = prQuantity * tbNomenk!quantity
'    ReDim Preserve equip(leng)
'    If Not IsNull(tbNomenk!Equip) Then
'        equip(leng) = tbNomenk!Equip
'    End If

End Sub

Function zakazNomenkToNNQQ() As Boolean
zakazNomenkToNNQQ = False

ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): QQ2(0) = 0: ReDim QQ3(0)

'ном-ра входящих изделий
sql = "SELECT pi.prId, " & _
" pi.prExt, " & _
" pi.quant, " & _
" ei.eQuant, " & _
" ei.prevQuant, " & _
" ei.prevQuant " & _
" FROM xPredmetyByIzdelia pi" & _
" LEFT JOIN xEtapByIzdelia ei ON pi.prExt = ei.prExt AND pi.prId = ei.prId AND pi.numOrder = ei.numOrder" & _
" WHERE pi.numOrder = " & gNzak

Set tbProduct = myOpenRecordSet("##319", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Function

While Not tbProduct.EOF
  gProductId = tbProduct!prId
  prExt = tbProduct!prExt
  If IsNull(tbProduct!eQuant) Then
    productNomenkToNNQQ tbProduct!quant, 0, 0
  Else
    productNomenkToNNQQ tbProduct!quant, tbProduct!eQuant, tbProduct!prevQuant
    QQ2(0) = 1 ' есть этап
  End If
  tbProduct.MoveNext
Wend
tbProduct.Close

'отдельная ном-ра
sql = "SELECT pn.nomNom, pn.quant as quantity, " & _
"en.eQuant, en.prevQuant FROM xPredmetyByNomenk pn " & _
"LEFT JOIN xEtapByNomenk en ON pn.nomNom = en.nomNom AND pn.numOrder = en.numOrder " & _
" WHERE pn.numOrder =" & gNzak
Set tbNomenk = myOpenRecordSet("##320", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
While Not tbNomenk.EOF
  If IsNull(tbNomenk!eQuant) Then
    nomenkToNNQQ 1, 0, 0
  Else
    nomenkToNNQQ 1, (tbNomenk!eQuant / tbNomenk!quantity), (tbNomenk!prevQuant / tbNomenk!quantity)
    QQ2(0) = 1 ' есть этап
  End If
  tbNomenk.MoveNext
Wend
tbNomenk.Close
zakazNomenkToNNQQ = True
End Function

'перед исп-ем надо ReDim NN(0): ReDim QQ(0): ReDim QQ2(0) : ReDim QQ3(0):QQ2(0)=0 - не б.этапа
Function productNomenkToNNQQ(pQuant As Double, eQuant As Double, _
                                               prQuantity As Double) As Boolean
Dim I As Integer, gr() As String

productNomenkToNNQQ = False

'вариантная ном-ра изделия
sql = "SELECT p.nomNom, p.quantity, p.xGroup " & _
"FROM sProducts p" _
& " JOIN xVariantNomenc x ON p.nomNom = x.nomNom AND p.ProductId = x.prId " _
& " WHERE x.numOrder =" & gNzak & " AND " _
& " x.prId =" & gProductId & " AND x.prExt=" & prExt
'Debug.Print sql
Set tbNomenk = myOpenRecordSet("##192", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
ReDim gr(0): I = 0
While Not tbNomenk.EOF
    nomenkToNNQQ pQuant, eQuant, prQuantity
    I = I + 1
    ReDim Preserve gr(I): gr(I) = tbNomenk!xGroup
    tbNomenk.MoveNext
Wend
tbNomenk.Close
    
'НЕвариантная ном-ра изделия
sql = "SELECT p.nomNom, p.quantity, p.xGroup" _
& " From sProducts p " _
& " WHERE p.ProductId =" & gProductId

'MsgBox sql
'Debug.Print sql
Set tbNomenk = myOpenRecordSet("##177", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
While Not tbNomenk.EOF
    For I = 1 To UBound(gr) ' если группа состоит из одной ном-ры, то она
        If gr(I) = tbNomenk!xGroup Then GoTo NXT ' НЕвариантна, т.к. не
    Next I                                      ' не попала в xVariantNomenc
    nomenkToNNQQ pQuant, eQuant, prQuantity
NXT: tbNomenk.MoveNext
Wend
tbNomenk.Close

productNomenkToNNQQ = True
End Function




'номенклатура в документе (Grid2) см. Common
'колонки Grid5'см соммоn

Private Sub cbInside_Click()
If isLoad And Grid.Visible Then loadKlassNomenk
End Sub


Private Function FlipCurrency(inp As String, N As Nomnom) As String
    Dim flipCurrencyRound As Double
    
    'asWhole - куда нужно переключиться
    If Not IsNumeric(inp) Then Exit Function
    If asWhole = 0 Then
        FlipCurrency = inp / N.perlist
    Else
        FlipCurrency = inp * N.perlist
    End If
    flipCurrencyRound = Round(FlipCurrency, 2)
    If Abs(flipCurrencyRound) < 0.005 Then
        flipCurrencyRound = Round(FlipCurrency, 3)
    End If
    FlipCurrency = flipCurrencyRound
End Function

Private Function FlipQuantity(inp As String, N As Nomnom) As String
    Dim flipQuantityRound As Double
    
    'asWhole - куда нужно переключиться
    If Not IsNumeric(inp) Then Exit Function
    If asWhole = 0 Then
        FlipQuantity = inp * N.perlist
    Else
        FlipQuantity = inp / N.perlist
    End If
    flipQuantityRound = Round(FlipQuantity, 2)
    If Abs(flipQuantityRound) < 0.005 Then
        flipQuantityRound = Round(FlipQuantity, 3)
    End If
    FlipQuantity = flipQuantityRound
End Function

Private Function FlipEdizm(N As Nomnom) As String
    'asWhole - куда нужно переключиться
    If asWhole = 0 Then
        FlipEdizm = N.edizm1
    Else
        FlipEdizm = N.Edizm2
    End If
End Function



Private Sub FlipGrid(ByRef G As MSFlexGrid, ByRef nomnomColIdx As Long, onlyInCache As Boolean)
Dim I As Integer, J As Integer, ColTitle As String, ColIsVisible As Boolean
Dim Nomnom1 As Nomnom

    If Not G.Visible Then
        Exit Sub
    End If
    
    With G
        For J = 1 To .Rows - 1
            Set Nomnom1 = New Nomnom
            If onlyInCache Then
                Set Nomnom1 = nomnomCache.onlyInCache(.TextMatrix(J, nomnomColIdx))
            Else
                Set Nomnom1 = nomnomCache.getNomnom(.TextMatrix(J, nomnomColIdx))
            End If
            If Not Nomnom1.IsInited Then
                For I = 1 To .Cols - 1
                    ColTitle = .TextMatrix(0, I)
                    ColIsVisible = Not .ColWidth(I) = 0
                    If ColIsVisible Then
                        If Left(ColTitle, 4) = "Цена" Then
                            .TextMatrix(J, I) = FlipCurrency(.TextMatrix(J, I), Nomnom1)
                        ElseIf Left(ColTitle, 4) = "Ед.и" Then
                            .TextMatrix(J, I) = FlipEdizm(Nomnom1)
                        ElseIf Mid(ColTitle, 2, 3) = "ол-" Then
                            .TextMatrix(J, I) = FlipQuantity(.TextMatrix(J, I), Nomnom1)
                        ElseIf Mid(ColTitle, 2, 4) = ".ост" Then
                            .TextMatrix(J, I) = FlipQuantity(.TextMatrix(J, I), Nomnom1)
                        End If
                    End If
                Next I
                
            End If
        Next J
    End With
End Sub

Private Sub ckList_Click()
    If noClick Then Exit Sub
    asWhole = ckList.Value
    gAsWhole = asWhole
    
    FlipGrid Grid, nkNomer, False
    FlipGrid Grid2, 1, False
    FlipGrid Grid5, prId, True
    'Form_Load
End Sub

Private Sub cmExel_Click()
If opNomenk.Value Then
    GridToExcel Grid, laGrid1.Caption
Else
    GridToExcel Grid, laGrid.Caption
End If

End Sub

Private Sub cmExel2_Click()
    GridToExcel Grid2, laGrid2.Caption
End Sub

Private Sub cmExit_Click()
    Unload Me
End Sub

Private Sub cmHide_Click()
Dim I As Integer
If quantity = 0 Then Exit Sub
For I = Grid.row To Grid.RowSel
    Grid.RemoveItem Grid.row
    quantity = quantity - 1
Next I
Grid.SetFocus
Grid_EnterCell
End Sub

Private Sub cmSel_Click() '<Добавить>
Dim befColor As Long, IL As Long, NL As Long, N As Integer, str As String
Dim per As Double
Dim Nomnom1 As Nomnom

Set Nomnom1 = nomnomCache.getNomnom(gNomNom)
per = Nomnom1.perlist

If Regim = "fromDocs" And per = 1 Then   'необрезная ном-ра м.б. только на складе -1001
    str = sDocs.lbInside.List(1) '-1002 - Обрезки
    If sDocs.Grid.TextMatrix(sDocs.mousRow, dcSour) = str Or _
    sDocs.Grid.TextMatrix(sDocs.mousRow, dcDest) = str Then
        MsgBox "Позиция '" & gNomNom & "' не может находиться на складе '" & _
        str & "'.", , "Предупреждение"
        Exit Sub
    End If
End If

If Not (Regim = "fromDocs" And sDocs.Regim = "fromCeh") Then _
    If beNaklads() Then Exit Sub

If Regim = "" Then
    If Otgruz.loadOutDates Then 'началась отгрузка
        MsgBox "По заказу " & gNzak & " уже началась отгрузка.", , "Редактирование запрещено!"
        Exit Sub
    End If
End If

If opProduct.Value Then
    Dim hasDrobnProc As Integer
    If idWerk = 1 Then
        sql = "select count(*) from wf_izdeliaForSell where prId = " & gProductId
        byErrSqlGetValues "##aaa", sql, hasDrobnProc
        If hasDrobnProc = 0 Then
            MsgBox "Нельзя продавать изделия, в которые может входить мерная номенклатура." _
            & vbCr & "Продавайте их как набор отдельных позиций номенклатуры.", , _
            "Добавить позицию нельзя!"
            Exit Sub
        End If
        
    End If
    
    'проверка комплектации изделия ***********
    befColor = 0: bilo = False
    Grid.col = nkQuant
    For IL = Grid.Rows - 1 To 0 Step -1 'так в цикле м. выявить концы всех групп т.к. всегда есть Row=0 c другим цветом
       Grid.row = IL
       grColor = Grid.CellBackColor
       If grColor <> befColor Then
          If (befColor = groupColor1 Or befColor = groupColor2) And Not bilo Then
              MsgBox "Одним цветом в колонке 'Кол-во' подсвеченны позиции, " & _
              "среди которых надо выбрать(двойной клик или <Enter>) только " & _
              "одну, которая войдет в изделие.", , "Изделие '" & gProduct & _
              "' не укомплектовано!"
              Exit Sub
          End If
          bilo = False
       End If
       If Grid.CellFontBold Then
          If grColor = groupColor1 Or grColor = groupColor2 Then bilo = True 'выбрана
       End If
       befColor = grColor
    Next IL '*********************************
     
    dostupOstatkiToGrid "multiN"
Else
    If Regim = "" And idWerk = 1 Then
        sql = "select count(*) from sGuideNomenk " _
        & " where web = 'vmt' AND nomnom = '" & gNomNom & "'"
        byErrSqlGetValues "##p.1", sql, hasDrobnProc
        If hasDrobnProc = 1 Then
            MsgBox "Вспомогательные материалы продавать нельзя", , "Добавить позицию нельзя!"
            Exit Sub
        End If
    End If
    
    dostupOstatkiToGrid
End If
tbQuant.Enabled = True
laQuant.Enabled = True
  tbQuant.Text = 1
  tbQuant.SelLength = 1
  tbQuant.SetFocus
  cmSel.Enabled = False
End Sub

Private Sub Command1_Click()
gridOrGrid3Hide "grid"
End Sub

Private Sub Command2_Click()
gridOrGrid3Hide "grid3"
End Sub

Private Sub Command3_Click()
gridOrGrid3Hide
End Sub


Private Sub Form_Load()
Dim str As String, I As Integer, delta As Double
ReDim selectedItems(0)

If Regim = "fromDocs" And sDocs.Regim = "fromCeh" And skladId = -1002 Then _
        opProduct.Enabled = False
noClick = False
msgBilo = False
isLoad = False

noClick = True
If idWerk = 1 Then
    ckList.Enabled = True
Else
    ckList.Enabled = False
End If

If asWhole = 1 Then
    ckList.Value = 1
Else
    ckList.Value = 0
End If
noClick = False

Grid.FormatString = "|<Номер|<Описание|<Ед.измерения|Кол-во|Ф.остатки|Д.остатки|kvop"

Grid.ColWidth(0) = 0
Grid.ColWidth(nkNomer) = 0 '900
Grid.ColWidth(nkName) = 2230  'ostat
Grid.ColWidth(nkEdIzm) = 630 'ostat
Grid.ColWidth(nkCurOstat) = 0
Grid.ColWidth(nkQuantPricise) = 0

Grid3.FormatString = "|<Номер|<Размер|<Описание|id"

Grid2.FormatString = "|<Номер|<Описание|<Ед.измерения|кол-во"
Grid2.ColWidth(0) = 0
Grid2.ColWidth(fnNomNom) = 0 '900
Grid2.ColWidth(fnNomName) = 2200  '900
Grid2.ColWidth(fnEdIzm) = 435
Grid2.ColWidth(fnQuant) = 585


If idWerk <> 1 Then ' мерж продаж в приор
    Grid5.FormatString = "|Тип|<Код|<Описание|<Ед.измерения|Цена за ед." & _
    "|Кол-во|Сумма|Этапы суммарно|Кол-во по тек.этапу"
    Grid5.ColWidth(prId) = 0
    Grid5.ColWidth(prName) = 1185
    Grid5.ColWidth(prType) = 0
    Grid5.ColWidth(prEdizm) = 420
    Grid5.ColWidth(prCenaEd) = 495
    Grid5.ColWidth(prEtap) = 660
    Grid5.ColWidth(prEQuant) = 675
Else
'   Grid2.Visible = False
'   laGrid2.Visible = False
    Grid5.FormatString = "|Тип|<Код|<Описание|<Ед.измерения|Цена за ед." & _
    "|Кол-во|Сумма|Вес"
    Grid5.ColWidth(prId) = 0
    Grid5.ColWidth(prName) = 1185
    Grid5.ColWidth(prType) = 0
    Grid5.ColWidth(prEdizm) = 420
    Grid5.ColWidth(prCenaEd) = 495
    Grid5.ColWidth(prVes) = 500
End If
cmExel.Visible = False
If Regim = "ostatP" Then
    Regim = "ostat"
    opProduct.Value = True: opProduct_Click
    GoTo AA
ElseIf Regim = "ostat" Then
    opNomenk.Value = True: opNomenk_Click
AA: Me.Caption = "Ведомость остатков"
    cmExel.Visible = True
    laGrid.Visible = False
    cbInside.Visible = False
    Grid.ColWidth(nkName) = 3510 + 900 ' тестировать в Остатки по гот.изделиям
    Grid.ColWidth(nkQuant) = 700
    Grid.ColWidth(nkCurOstat) = 710
    Grid.ColWidth(nkDostup) = 700
    cmSel.Visible = False
    tbQuant.Visible = False
    laQuant.Visible = False
    laGrid2.Visible = False
    Grid2.Visible = False
    Grid5.Visible = False
    Grid.Width = 7700
    Me.Width = Grid.Width + 2527
    Grid3.Width = Grid.Width
    laGrid.Width = Grid.Width
    cmExit.Left = Me.Width - cmExit.Width - 200
    
    sql = "SELECT sourceId, sourceName From sGuideSource " & _
    "WHERE sourceId <-1000 ORDER BY sourceId DESC"
    Set Table = myOpenRecordSet("##359", sql, dbOpenDynaset)
    If Table Is Nothing Then myBase.Close: End
    While Not Table.EOF
        cbInside.AddItem Table!SourceName
        Table.MoveNext
    Wend
    Table.Close
    cbInside.ListIndex = 0
    'cbInside.Visible = True
    isLoad = True
    GoTo EN1 ' Exit Sub
ElseIf Regim = "fromDocs" Then
    laGrid5.Visible = False
    Grid5.Visible = False
    laGrid2.Top = laGrid5.Top
    delta = Grid2.Top - Grid5.Top
    Grid2.Top = Grid5.Top
    Grid2.Height = Grid2.Height + delta
ElseIf Regim = "" Or Regim = "closeZakaz" Then 'предметы заказа
BB: cmExel2.Visible = True
    laGrid5.Visible = True
End If

gSeriaId = 0 'необходим  для добавления класса

quantity2 = 0
loadProducts 'номенклатурный состав
If Regim = "" Or Regim = "closeZakaz" Then
    loadPredmeti Me, orderRate, idWerk, Me.asWhole  ' состав предметов
End If

If quantity2 > 0 Then
    str = "Редактирование"
Else
    str = "Формирование"
End If
If Regim = "fromDocs" Then
    Me.Caption = str & " предметов к накладной № " & numDoc
Else
    Me.Caption = str & " предметов к заказу № " & numDoc
End If
If Regim = "closeZakaz" Then
    Me.Caption = "Предметы к заказу № " & numDoc
    laBegin.Caption = "Это закрытый заказ. Редактирование предметов невозможно."
    opNomenk.Enabled = False
    opProduct.Enabled = False
    cmSel.Enabled = False
    laGrid2.Enabled = False
    Grid2.Enabled = False
    cmExel2.Visible = False
    tv.Enabled = False
    laGrid1.Visible = False
    laGrid.Visible = False
Else
    opProduct.Value = True ': opProduct_Click
End If

EN1:
splLeftH.Visible = False
oldHeight = Me.Height
oldWidth = Me.Width
tvVes = tv.Width / (tv.Width + Grid.Width + Grid5.Width)
gridVes = Grid.Width / (tv.Width + Grid.Width + Grid5.Width)
grid2Ves = Grid5.Width / (tv.Width + Grid.Width + Grid5.Width)
isLoad = True
End Sub


Sub loadProducts() ' номенклатура заказа или накладной


MousePointer = flexHourglass
Grid2.Visible = False

Dragging = False
quantity2 = 0
clearGrid Grid2
If numExt = 254 Then
    sql = "SELECT n.nomNom, n.nomName, n.perList" _
    & " ,n.ed_Izmer, n.ed_Izmer2, d.quant as quantity  " _
    & " ,n.Size, n.cod, n.ves" _
    & " FROM sDMC d" _
    & " JOIN sGuideNomenk n ON n.nomNom = d.nomNom" _
    & " WHERE d.numDoc = " & numDoc & " And d.numExt =" & numExt
ElseIf Regim <> "fromDocs" Then GoTo AA
ElseIf numExt = 0 And sDocs.reservNoNeed Then ' без резервирования
    sql = "SELECT n.nomNom, n.nomName, n.perList" _
    & " ,n.ed_Izmer, n.ed_Izmer2, d.quantity  " _
    & " ,n.Size, n.cod" _
    & " FROM sGuideNomenk n" _
    & " JOIN sDMCmov d ON n.nomNom = d.nomNom" _
    & " WHERE d.numDoc =" & numDoc
Else ' ном-ра заказа или выписанная из Цеха с Целого склада
AA: sql = "SELECT n.nomNom, n.nomName, n.perList " _
    & " ,n.ed_Izmer, n.ed_Izmer2, d.quantity" _
    & " ,n.Size, n.cod, ves" _
    & " FROM sDMCrez d " _
    & " JOIN sGuideNomenk n ON n.nomNom = d.nomNom" _
    & " WHERE d.numDoc =" & numDoc
End If
'MsgBox sql

Dim Nomnom1 As Nomnom
Set tbNomenk = myOpenRecordSet("##118", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    Set Nomnom1 = nomnomCache.getNomnom(tbNomenk!Nomnom, True)
    
    quantity2 = quantity2 + 1
    Grid2.TextMatrix(quantity2, fnNomNom) = tbNomenk!Nomnom
'    Grid2.TextMatrix(quantity2, dnNomName) = tbNomenk!nomName
    Grid2.TextMatrix(quantity2, fnNomName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid2.TextMatrix(quantity2, fnEdIzm) = Nomnom1.getEdizm(asWhole)
    Grid2.TextMatrix(quantity2, fnQuant) = Nomnom1.getQuantity(tbNomenk!quantity, asWhole)
    If Regim = "fromDocs" Then
      If sDocs.isIntMove() Then
        Grid2.TextMatrix(quantity2, fnEdIzm) = tbNomenk!ed_Izmer2
        Grid2.TextMatrix(quantity2, fnQuant) = Round(tbNomenk!quantity / tbNomenk!perlist, 2)
      End If
    End If
    
    Grid2.AddItem ""
    tbNomenk.MoveNext
  Wend
  Grid2.RemoveItem quantity2 + 1
End If
tbNomenk.Close
EN1:
Grid2.Visible = True
MousePointer = flexDefault


End Sub

'изначально правая таблица д.б. шире левых
Sub rightORleft(reg As String) ' reg =l или r
Static begWidth2 As Integer, begWidth As Integer, begLeft As Integer
Dim delta As Integer

If Regim = "ostatP" Or Regim = "ostat" Then Exit Sub
If begWidth = 0 Then ' т.е. только один раз
    begWidth = Grid.Width
    begWidth2 = Grid2.Width
    begLeft = Grid2.Left
End If
If opProduct.Value Then
    delta = 2000 ' Product
Else
    delta = 1200 ' Nomenk
End If
 
If reg = "r" Then
    Grid.Width = begWidth
    Grid2.Width = begWidth2
    Grid2.Left = begLeft
ElseIf reg = "l" Then
    Grid.Width = begWidth + delta
    Grid2.Width = begWidth2 - delta
    Grid2.Left = begLeft + delta
End If

Grid3.Width = Grid.Width

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging Then
        Dragging = True
    End If

End Sub

Private Sub Form_Resize()
Dim H As Integer, W As Integer, hh As Double, ww As Double, Left As Long

If Not isLoad Then Exit Sub
If Me.WindowState = vbMinimized Then Exit Sub
If Me.WindowState = vbMaximized And Me.Width > cDELLwidth Then 'экран DELL
    Grid5.ColWidth(prDescript) = 2400 + 375
    Grid2.ColWidth(fnNomName) = 5250 + 930
Else
    Grid5.ColWidth(prDescript) = 840 + 375
    Grid2.ColWidth(fnNomName) = 2340 + 930
End If
setNameColWidth 'в Grid3 и Grid

On Error Resume Next
H = Me.Height - oldHeight
oldHeight = Me.Height
W = Me.Width - oldWidth
oldWidth = Me.Width

tv.Height = tv.Height + H
tv.Width = tv.Width + W * tvVes

Grid.Left = Grid.Left + W * tvVes
laGrid.Left = Grid.Left
cbInside.Left = laGrid1.Left + laGrid1.Width
Grid.Height = Grid.Height + H
Grid.Width = Grid.Width + W * gridVes
Grid3.Left = Grid.Left
laGrid1.Left = Grid3.Left
laBegin.Left = tv.Left + tv.Width + 100
Grid3.Width = Grid.Width

Grid2.Left = Grid2.Left + W * (tvVes + gridVes)
laGrid2.Left = Grid2.Left
laGrid2.Top = laGrid2.Top + H / 2

Grid2.Top = Grid2.Top + H / 2
Grid2.Height = Grid2.Height + H / 2
Grid2.Width = Grid2.Width + W * grid2Ves
Grid5.Left = Grid2.Left
Grid5.Height = Grid5.Height + H / 2

laGrid5.Left = laGrid5.Left + W * (tvVes + gridVes)
Grid5.Width = Grid2.Width

splLeftV.Top = Grid5.Top
splLeftV.Left = tv.Left + tv.Width + 15
splLeftV.Height = tv.Top + tv.Height

splRightV.Top = splLeftV.Top
splRightV.Left = Grid5.Left - 15 - splRightV.Width
splRightV.Height = splLeftV.Height

splLeftH.Top = Grid5.Top + Grid5.Height + 5
splLeftH.Left = splLeftV.Left + splLeftV.Width
splLeftH.Width = Grid3.Width

splRightH.Top = Grid5.Top + Grid5.Height + 5
splRightH.Left = splRightV.Left + splRightV.Width
splRightH.Width = Grid5.Width

'' по заданию шефа таблицы выравнивать так, чтобы между сплиттерами не было свободных мест.
'Dim frameWidth As Long ' ширина пространства между сплиттерами
'Dim CIndex As Long     ' индекс колонки, которая будет меняться
'Dim WidthDelta As Long ' На сколько изменить ширину колонки'

'frameWidth = splRightV.Left - splLeftV.Left + splLeftV.Width
'With Grid
'    CIndex = nkName
'    WidthDelta = .Width - frameWidth
'    .Width = .Width - WidthDelta
'End With

'Grid.ColWidth(CIndex) = Grid.ColWidth(CIndex) - WidthDelta
'With Grid3
'    CIndex = fnNomName
'    WidthDelta = .Width - frameWidth
'    .Width = .Width - WidthDelta
'    '.ColWidth(CIndex) = .ColWidth(CIndex) + WidthDelta
'End With
'Grid3.ColWidth(CIndex) = Grid3.ColWidth(CIndex) - WidthDelta

'frameWidth = Me.Width - splRightV.Left + splRightV.Width
'With Grid2
'    CIndex = fnNomName
'    WidthDelta = .Width - frameWidth
'    .Width = .Width - WidthDelta
'    '.ColWidth(CIndex) = .ColWidth(CIndex) + WidthDelta
'End With
'Grid2.ColWidth(CIndex) = Grid2.ColWidth(CIndex) - WidthDelta
'
'With Grid5
'    CIndex = prName
'    WidthDelta = .Width - frameWidth
'    .Width = .Width - WidthDelta
'    '.ColWidth(CIndex) = .ColWidth(CIndex) + WidthDelta
'End With
'Grid5.ColWidth(CIndex) = Grid5.ColWidth(CIndex) - WidthDelta

laQuant.Top = laQuant.Top + H
laQuant.Left = laQuant.Left + W
cmExit.Top = cmExit.Top + H
cmExit.Left = cmExit.Left + W
cmExel2.Top = cmExel2.Top + H
cmExel2.Left = cmExel2.Left + W
cmExel.Top = cmExel.Top + H
cmHide.Top = cmHide.Top + H
laBegin.Top = tv.Top
ckList.Top = ckList.Top + H

cmSel.Top = cmSel.Top + H
tbQuant.Top = tbQuant.Top + H

cmSel.Left = Grid3.Left + Grid3.Width / 2 - cmSel.Width - 50
tbQuant.Left = cmSel.Left + cmSel.Width + 50

End Sub

Private Sub Form_Unload(Cancel As Integer)
isLoad = False
If Regim = "" Then 'предметы заказа
    If beNaklads("noMsg") Then Exit Sub 'т.к. в этой форме ничего поменять не могли
    If quantity2 > 0 Then
        Orders.Grid.CellForeColor = 200
    Else
        Orders.Grid.CellForeColor = vbBlack
    End If
ElseIf Regim = "fromDocs" Then
   If sDocs.isIntMove() Then sDocs.ckPerList.Value = 1 Else sDocs.ckPerList.Value = 0
End If
End Sub

Sub dostupOstatkiToGrid(Optional reg As String)
Dim S As Double, sum As Double, rr As Long, IL As Long

Me.MousePointer = flexHourglass
If numExt = 254 Or (Regim = "fromDocs" And sDocs.Regim = "fromCeh") Then
    laGrid4.Caption = "Фактические остатки по подразделению '" & sDocs.getGridColSour() & "'"
Else
    laGrid4.Caption = "Доступные остатки"
End If
clearGrid Grid4
Grid4.FormatString = "|<Номер|<Описание|<Ед.измерения|Кол-во"
Grid4.ColWidth(0) = 0
Grid4.ColWidth(frNomNom) = 870
Grid4.ColWidth(frNomName) = 4485
Grid4.ColWidth(frEdIzm) = 645
Grid4.ColWidth(frOstat) = 885

If reg = "multiN" Then
    Grid.col = nkQuant: IL = 0
    For rr = 1 To Grid.Rows - 1
        Grid.row = rr
        If Grid.CellFontBold Then
            IL = IL + 1
            gNomNom = Grid.TextMatrix(rr, nkNomer)
            nomencOstatkiToGrid IL
            Grid4.AddItem ""
        End If
    Next rr
    Grid4.RemoveItem Grid4.Rows - 1
Else
    nomencOstatkiToGrid 1
End If
Grid4.Visible = True
EN1:
Me.MousePointer = flexDefault
gridFrame.Visible = True
gridFrame.ZOrder

End Sub
'ед.измер зависит от складов
'заголововок Grid4 формируется в dostupOstatkiToGrid
Public Function nomencOstatkiToGrid(row As Long, Optional hasLiveRowset As Boolean = False) As Double
Dim S As Double, str As String, str2 As String, str3 As String, z As Double

'Ф.остатки
Dim Nomnom1 As Nomnom

If row > 0 Then
    Grid4.TextMatrix(row, frNomNom) = gNomNom
    Set Nomnom1 = nomnomCache.getNomnom(gNomNom, hasLiveRowset)
    Grid4.TextMatrix(row, frNomName) = Nomnom1.nomName
    Grid4.TextMatrix(row, frEdIzm) = Nomnom1.getEdizm(asWhole)
End If

If Regim = "fromDocs" Then
    nomencOstatkiToGrid = PrihodRashod("+", skladId) - PrihodRashod("-", skladId) 'Ф. остатки по складу
    If sDocs.isIntMove() Then
        If row > 0 Then Grid4.TextMatrix(row, frEdIzm) = str3
        Set Nomnom1 = nomnomCache.getNomnom(gNomNom, hasLiveRowset)
        nomencOstatkiToGrid = nomencOstatkiToGrid / Nomnom1.perlist
    End If
Else
'вычисляем доступные остатки
AA:
If cbInside.Visible And cbInside.ListIndex = 1 Then
    FO = PrihodRashod("+", -1002) - PrihodRashod("-", -1002) ' обрезки
Else
    FO = PrihodRashod("+", -1001) - PrihodRashod("-", -1001) ' обрезки не показываем
    sql = "SELECT Sum(quantity) AS Sum_quantity, " & _
    "Sum(Sum_quant) AS Sum_Sum_quant From wCloseNomenk " & _
    "WHERE nomNom ='" & gNomNom & "'"
'    Debug.Print sql
    If Not byErrSqlGetValues("##145", sql, z, S) Then myBase.Close: End
        nomencOstatkiToGrid = FO - (z - S) ' минус, что несписано
    End If
End If
If row > 0 Then _
    Grid4.TextMatrix(row, frOstat) = Nomnom1.getQuantity(nomencOstatkiToGrid, asWhole)

End Function

Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If quantity = 0 Then Exit Sub
'в Grid недопустима сортировка т.к. цветовая группировка
End Sub

Private Sub Grid_DblClick()
Dim IL As Long, curRow As Long

grColor = Grid.CellBackColor
If grColor = &H88FF88 Then
    showRezerv 1, 2, Grid.TextMatrix(mousRow, nkEdIzm), Me
ElseIf grColor = groupColor1 Or grColor = groupColor2 Then
    If Me.Regim <> "ostat" Then
        curRow = Grid.row
        Grid.CellFontBold = True
    '    Grid.col = nkQuant
        For IL = curRow - 1 To 1 Step -1  'вверх от клика
            Grid.row = IL
            If Grid.CellBackColor <> grColor Then Exit For
            Grid.CellFontBold = False
        Next IL
        For IL = curRow + 1 To Grid.Rows - 1 'вниз от клика
            Grid.row = IL
            If Grid.CellBackColor <> grColor Then Exit For
            Grid.CellFontBold = False
        Next IL
        Grid.row = curRow
    Else
        
    End If
End If
End Sub

Function checkFactDost(ByRef factCol As Integer, ByRef dostCol As Integer) As Boolean
Dim fact As Double, dost As Double
Dim factStr As String, dostStr As String
    factStr = Grid.TextMatrix(mousRow, factCol)
    If Not IsNumeric(factStr) Then
        fact = 0
    Else
        fact = CDbl(factStr)
    End If
    
    dostStr = Grid.TextMatrix(mousRow, dostCol)
    If Not IsNumeric(dostStr) Then
        dost = 0
    Else
        dost = CDbl(dostStr)
    End If
    checkFactDost = dost < fact
End Function

Private Sub Grid_EnterCell()

mousRow = Grid.row
mousCol = Grid.col
gNomNom = Grid.TextMatrix(mousRow, nkNomer)

If quantity = 0 Or Grid.col = buntColumn Then Exit Sub


Grid.CellBackColor = vbYellow
If mousCol = nkDostup And cbInside.ListIndex = 0 And checkFactDost(nkCurOstat, nkDostup) Then
    Grid.CellBackColor = &H88FF88
End If

End Sub

Private Sub Grid_GotFocus()
cmHide.Enabled = True
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid_DblClick
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Grid_EnterCell
End Sub

Private Sub Grid_LeaveCell()
If Grid.col <> 0 And Grid.col <> buntColumn Then Grid.CellBackColor = Grid.BackColor
End Sub


Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
End Sub

Private Sub Grid2_Click()
mousCol2 = Grid2.MouseCol
mousRow2 = Grid2.MouseRow
If quantity2 = 0 Then Exit Sub
If Grid2.MouseRow = 0 Then
    Grid2.CellBackColor = Grid2.BackColor
    If mousCol2 = fnQuant Then
        SortCol Grid2, mousCol2, "numeric"
    Else
        SortCol Grid2, mousCol2
    End If
    Grid2.row = 1    ' только чтобы снять выделение
    Grid_EnterCell
End If

End Sub


Private Sub Grid2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Grid2.MouseRow = 0 And Shift = 2 Then
        MsgBox "ColWidth = " & Grid2.ColWidth(Grid2.MouseCol)
ElseIf Button = 2 And Regim = "fromDocs" And quantity2 > 0 Then
    Grid2.col = fnNomNom
    Grid2.row = Grid2.MouseRow
    mousRow2 = Grid2.row
    gNomNom = Grid2.Text
    Grid2.SetFocus
    Grid2.CellBackColor = vbButtonFace
    Me.PopupMenu mnContext2
    Grid2.CellBackColor = Grid2.BackColor
End If
End Sub

Sub gridOrGrid3Hide(Optional purpose As String)
Dim formHeight As Double

formHeight = tv.Top + tv.Height - Grid3.Top
If purpose = "grid3" Then ' есть
    Grid.Top = Grid3.Top
    Grid.Height = formHeight
    Grid.ZOrder
ElseIf purpose = "grid" Then ' есть только одна Grid3
    Grid3.Height = formHeight
    splLeftH.Top = Grid3.Top + Grid3.Height
    Grid3.ZOrder
Else '                обе присутствуют
    Dim gHeightMax As Double, g3HeightMax As Double
    Dim fullHeight As Double
    
    gHeightMax = Grid.Rows * (Grid.CellHeight + 13) + 95
    g3HeightMax = Grid3.Rows * (Grid3.CellHeight + 13) + 95
    fullHeight = gHeightMax + g3HeightMax + laGrid.Height
    
    If fullHeight > formHeight Then
        Dim normalGrid As Single, normalGird3 As Single
        Dim baseIsGrid As Boolean, baseIsGrid3 As Boolean
        
        If gHeightMax < formHeight * 2 / 3 Then
            baseIsGrid = True
        ElseIf g3HeightMax < formHeight * 2 / 3 Then
            baseIsGrid3 = True
        End If
        
        If Not baseIsGrid And Not baseIsGrid3 Then
            normalGrid = gHeightMax / fullHeight
            normalGird3 = g3HeightMax / fullHeight
            gHeightMax = CInt(normalGrid * formHeight)
            baseIsGrid = True
        End If
        
        If baseIsGrid Then
            Grid.Height = gHeightMax
            Grid.Top = tv.Top + tv.Height - Grid.Height
            laGrid.Top = Grid.Top - laGrid.Height  '+90
            Grid3.Height = laGrid.Top - Grid3.Top
            splLeftH.Top = Grid3.Top + Grid3.Height
        Else
            Grid3.Height = g3HeightMax
            splLeftH.Top = Grid3.Top + Grid3.Height
            laGrid.Top = Grid3.Top + Grid3.Height
            Grid.Top = laGrid.Top + laGrid.Height
            Grid.Height = tv.Top + tv.Height - Grid.Top
        End If
        
        
    Else
        Grid3.Height = g3HeightMax
        splLeftH.Top = Grid3.Top + g3HeightMax
        laGrid.Top = splLeftH.Top + splLeftH.Height
        Grid.Height = gHeightMax
        Grid.Top = laGrid.Top + laGrid.Height
        Grid.Height = tv.Top + tv.Height - Grid.Top
    End If
    If mousRow3 < Grid3.Rows Then
        If Not Grid3.RowIsVisible(mousRow3) Then rowViem mousRow3, Grid3
    End If
    If Not Grid.ColIsVisible(1) Then
        Grid.LeftCol = 1
    End If
    
End If

End Sub


Sub newProductRow(row As Long)
    
    gProductId = Grid3.TextMatrix(row, gpId)
    If IsNumeric(gProductId) Then
        gProduct = Grid3.TextMatrix(row, gpName)
        laGrid.Visible = True
        laGrid.Caption = "Список номенклатуры по изделию '" & gProduct & "'"
        loadProductNomenk gProductId
        controlEnable True
        gridOrGrid3Hide ""
        Grid.TopRow = 1
    End If
    
End Sub
  
Private Sub Grid3_EnterCell()
'Static prevRow As Long

If quantity3 = 0 Or Grid3.MouseRow = 0 Then Exit Sub
If biloG3Enter_Cell Then Exit Sub 'если sub уже вызывалась в текущ сеансе

biloG3Enter_Cell = True
Timer1.Enabled = False

Timer1.Interval = 100
Timer1.Enabled = True

mousRow3 = Grid3.row
mousCol3 = Grid3.col

'If prevRow <> Grid3.row Then newProductRow
Grid3.CellBackColor = &HCCCCCC

'prevRow = Grid3.row

End Sub

Private Sub Grid3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        newProductRow Grid3.row
    End If
End Sub

Private Sub Grid3_LeaveCell()
Grid3.CellBackColor = Grid3.BackColor
End Sub

Private Sub Grid3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'не выносить это в Grid3_Click

If Grid3.MouseRow = 0 Then
    Grid3.CellBackColor = Grid3.BackColor
    SortCol Grid3, mousCol3
    Grid3.row = 1    ' только чтобы снять выделение
    gridOrGrid3Hide "grid"
Else
    'If biloG3Enter_Cell Then Exit Sub
    'mousCol3 = Grid3.MouseCol
    'mousRow3 = Grid3.MouseRow
    'If quantity3 = 0 Then Exit Sub
    'Grid3.CellBackColor = &HCCCCCC
    'newProductRow
End If

End Sub

Private Sub Grid3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Grid3.MouseRow = 0 And Shift = 2 Then MsgBox "ColWidth = " & Grid3.ColWidth(Grid3.MouseCol)

'On Error Resume Next ' когда снова по tv_click
Grid3.row = Grid3.MouseRow 'чтобы снять выделение неск.строк возник по gridOrGrid3Hide
Grid3.RowSel = Grid3.MouseRow '
biloG3Enter_Cell = False
newProductRow Grid3.MouseRow

End Sub

Private Sub Grid4_GotFocus()
    tbQuant.SetFocus
End Sub

Private Sub Grid4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Grid4.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid4.ColWidth(Grid4.MouseCol)

End Sub


Private Sub Grid5_Click()
If Not noClick Then Grid5_EnterCell
noClick = False
End Sub

Private Sub Grid5_DblClick()
Dim id As Integer

If mousRow5 = 0 Then Exit Sub
If mousCol5 = prEQuant Then
    If Not predmetiIsClose("prev") Then
        MsgBox "По этапу уже началось списание.", , "Редактирование невозможно!"
        Exit Sub
    End If
    sql = "SELECT StatusId from Orders WHERE (((numOrder)=" & gNzak & "));"
    If Not byErrSqlGetValues("##389", sql, id) Then Exit Sub

'    If id <> 0 And id <> 6 Then 'НЕ закрыт или принят
'        MsgBox "Открыть этап выполнения заказа можно только, когда заказ " & _
'        "Принят или Закрыт по предыдущему этапу.", , ""
    If id <> 0 And id <> 4 Then 'НЕ принят или не готов
        MsgBox "Открыть этап выполнения заказа можно только, когда заказ " & _
        "Принят или Готов.", , ""
        Exit Sub
    End If
    If Not msgBilo And Grid5.TextMatrix(mousRow5, prEtap) = "" Then
        msgBilo = True
        MsgBox "Если Вы хотите задать Этап выполнения заказа, введите " & _
        "количества, необходимые для выполнения этого Этапа.  Это позволит " & _
        "цеху поставить промежуточную Готовность." & vbCrLf & "Иначе введите " & _
        "ноль.", , "Предупреждение"
    End If
End If
If Grid5.CellBackColor = &H88FF88 Then textBoxInGridCell tbMobile, Grid5
End Sub

Private Sub Grid5_EnterCell()
Static prevRow As Long, prevCol As Long
' EnterCell срабатывает до MouseDown
    If Not Grid5.Visible Then
        Exit Sub
    End If
    If hasSelection(Grid5) Then Exit Sub
    
    If Not selectNomenkFlag And Not isCtrlDown Then
        If quantity5 = 0 Then Exit Sub
        If Grid5.row > quantity5 Then
            If prevRow <> Grid5.row Then
                laGrid2.Caption = "Номенклатурный состав предметов заказа:"
                loadProducts 'номен-ра заказа
            End If
            prevRow = Grid5.row
            Exit Sub
        End If
        
        mousRow5 = Grid5.row
        mousCol5 = Grid5.col
        getIdFromGrid5Row Me
        
        If mousCol5 = prSumm Or mousCol5 = prCenaEd Or mousCol5 = prEQuant Then
            Grid5.CellBackColor = &H88FF88
        Else
            Grid5.CellBackColor = vbYellow
        End If
         
         'не стирать - здесь нет суммарно по всем изделиям:
         'if Grid5.col = prEtap Then
         '   If IsNumeric(Grid5.TextMatrix(mousRow5, prEtap)) Then _
         '       productNomenkToGrid2 CInt(Grid5.TextMatrix(mousRow5, prEtap))
        'ElseIf prevRow <> Grid5.row Or prevCol = prEtap Then
        '    productNomenkToGrid2 CInt(Grid5.TextMatrix(mousRow5, prQuant))
        'End If
        
        If prevRow <> Grid5.row Then
             productNomenkToGrid2 CInt(Grid5.TextMatrix(mousRow5, prQuant))
        End If
        
        prevRow = Grid5.row
        prevCol = Grid5.col
    Else
        'Режим выбора номенклатуры для сложного изделия
        
    End If
    
End Sub
'$odbc14$
Sub productNomenkToGrid2(quant As Double)
Dim IL As Long, str As String, str2 As String, str3 As String, str4 As String

If quantity5 = 0 Then Exit Sub


Grid2.Visible = False
clearGrid Grid2
quantity2 = 0
If Grid5.TextMatrix(mousRow5, prType) = "изделие" Then
  ReDim NN(0): ReDim QQ(0)
  If productNomenkToNNQQ(quant, 0, 0) Then
    laGrid2.Caption = "Состав по готовому изделию '" & Grid5.TextMatrix(mousRow5, nkName) & "'"
'    Set tbNomenk = myOpenRecordSet("##193", "select * from sGuideNomenk", dbOpenForwardOnly)
'    If tbNomenk Is Nothing Then Exit Sub
'    tbNomenk.index = "PrimaryKey"
    Dim Nomnom1 As Nomnom
    For IL = 1 To UBound(NN)
        quantity2 = quantity2 + 1
        Grid2.AddItem ""
        Grid2.TextMatrix(IL, fnNomNom) = NN(IL)
        Set Nomnom1 = New Nomnom
        Set Nomnom1 = nomnomCache.getNomnom(NN(IL))
        
        Grid2.TextMatrix(IL, fnNomName) = Nomnom1.cod & " " & Nomnom1.nomName & " " & Nomnom1.Size
        Grid2.TextMatrix(IL, fnEdIzm) = Nomnom1.getEdizm(asWhole)
        Grid2.TextMatrix(IL, fnQuant) = Nomnom1.getQuantity(QQ(IL), asWhole)
    Next IL
    
    If quantity2 > 0 Then Grid2.RemoveItem Grid2.Rows - 1
  End If
Else
    laGrid2.Caption = "Отдельная номенклатура"
    
    quantity2 = 1
    Grid2.TextMatrix(1, fnNomNom) = Grid5.TextMatrix(mousRow5, prName)
    Grid2.TextMatrix(1, fnNomName) = Grid5.TextMatrix(mousRow5, prName) & _
        " " & Grid5.TextMatrix(mousRow5, prDescript)
    Grid2.TextMatrix(1, fnEdIzm) = Grid5.TextMatrix(mousRow5, prEdizm)
    Grid2.TextMatrix(1, fnQuant) = Grid5.TextMatrix(mousRow5, prQuant)
End If
Grid2.Visible = True

End Sub

Private Sub Grid5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 And Shift = vbCtrlMask Then
        ' Возможно следующим нашим действием будет нажатие на левую педаль мыши
        isCtrlDown = True
    End If
    
    If KeyCode = vbKeyReturn Then Grid5_DblClick

End Sub

Private Sub Grid5_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Then
        isCtrlDown = False
    End If
    If KeyCode = vbKeyEscape Then Grid5_EnterCell
End Sub
Private Sub Grid5_LeaveCell()
    If Not isCtrlDown And Not hasSelection(Grid5) Then
        Grid5.CellBackColor = Grid5.BackColor
    End If

End Sub

Private Sub Grid5_LostFocus()
    Grid5_LeaveCell
End Sub




Private Sub cleanSelection(Grd As MSFlexGrid)
Dim I As Integer, J As Integer
Dim currentRow As Integer, currentCol As Integer

ReDim selectedItems(0)
currentCol = Grd.col
currentRow = Grd.row

For I = Grd.Rows - 1 To 1 Step -1
    Grd.row = I
    For J = Grd.Cols - 1 To 1 Step -1
        Grd.col = J
        Grd.CellBackColor = Grd.BackColor
        Grd.CellForeColor = Grd.ForeColor
    Next J
Next I
Grd.row = currentRow
Grd.col = currentCol

End Sub
Private Function hasSelection(Grd As MSFlexGrid) As Boolean
Dim I As Integer
Dim currentRow As Integer

    hasSelection = False
    If UBound(selectedItems) > 0 Then
        hasSelection = True
    End If
    
End Function
Private Function useMaxSelection() As Long
Dim I As Integer
Dim sz As Integer

    useMaxSelection = 0
    sz = UBound(selectedItems)
    For I = 1 To sz
        If selectedItems(I) > useMaxSelection Then
            useMaxSelection = selectedItems(I)
        End If
    Next I
    RemoveItem (CStr(useMaxSelection))
End Function

Private Sub appendItem(Item As String)
Dim I As Integer
Dim found As Boolean: found = False
Dim sz As Integer

    sz = UBound(selectedItems)
    For I = 1 To sz
        If selectedItems(I) = Item Then found = True: Exit For
    Next I
    If Not found Then
        ReDim Preserve selectedItems(sz + 1)
        selectedItems(sz + 1) = Item
    End If
    
End Sub

Private Sub RemoveItem(Item As String)
Dim I As Integer
Dim found As Boolean: found = False
Dim sz As Integer

    sz = UBound(selectedItems)
    For I = 1 To sz
        If found Then
            ' сдвинуть хвост
            selectedItems(I - 1) = selectedItems(I)
        ElseIf selectedItems(I) = Item Then
            found = True
        End If
    Next I
    If found Then
        ReDim Preserve selectedItems(sz - 1)
    End If
    'selectedItems(sz) = item
    
End Sub

Private Sub mark(Grd As MSFlexGrid, setFlag As Boolean)
Dim fColorSel As Long
Dim bColorSel As Long
Dim I As Integer
Dim currentCol As Integer, currentLeft As Integer

    
'    If IsMissing(color) Then color = vbRed
    currentCol = Grd.col
    currentLeft = Grd.CellLeft
    
    If setFlag Then
        fColorSel = vbWhite
        bColorSel = vbRed
        appendItem (CStr(Grd.row))
    Else
        fColorSel = vbBlack
        bColorSel = Grd.BackColor
        RemoveItem (CStr(Grd.row))
    End If

    For I = 0 To Grd.Cols - 1
        Grd.col = I
        Grd.CellBackColor = bColorSel
        Grd.CellForeColor = fColorSel
    Next I
    'Grd.CellLeft = currentLeft
    Grd.col = currentCol
    
    
End Sub


Private Sub Grid5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer

    If isCtrlDown And Button = 1 And Grid5.row <> 0 And Grid5.row <> Grid5.Rows - 1 Then
        'Подсветить всю строку
        If Grid5.CellBackColor = vbRed Then
            mark Grid5, False
        Else
            mark Grid5, True
        End If
    End If
    
    If Shift = vbCtrlMask And Button = 1 Then
        'selectNomenkFlag = True
        'Grid5.SelectionMode = flexSelectionByRow
    End If
    If selectNomenkFlag And Button = 2 And Shift = 0 Then
        ' Показать меню "Создания Изделия на лету"
        'selectNomenkFlag = selectNomenkFlag
    End If
    If selectNomenkFlag And Button = 1 And Shift <> 2 Then
        'selectNomenkFlag = False
        'Grid5.SelectionMode = flexSelectionFree
    End If
    
End Sub

Private Sub Grid5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If isCtrlDown Then
    Else
        If hasSelection(Grid5) And Button = 1 Then
            cleanSelection Grid5
        End If
        If Grid5.MouseRow = 0 Then
            If Shift = 2 Then MsgBox "ColWidth = " & Grid5.ColWidth(Grid5.MouseCol)
        ElseIf Button = 2 And 0 < Grid5.MouseRow And Grid5.MouseRow < Grid5.Rows - 1 _
        And quantity5 > 0 And Regim <> "closeZakaz" Then
'            Grid5.row = Grid5.MouseRow
'            Grid5.col = prName
'            Grid5.SetFocus
'            Grid5.CellBackColor = vbButtonFace
            getIdFromGrid5Row Me   ' gNomNom и gProductId
            If Not hasSelection(Grid5) Then
                isCtrlDown = True
                mark Grid5, True
                isCtrlDown = False
            End If
            
            Me.PopupMenu mnContext
        End If
    End If
    If Button = 1 And Shift = 2 Then
        'Выделить всю строку
        
    End If
'If Button = 2 And frmMode = "" Then

End Sub




'$odbc15$
Public Sub mnDel_Click()
Dim pQuant As Double, I As Integer, str  As String, str2 As String
Dim comma As String
Dim hasEtap As Boolean
Dim gIndex As Integer
Dim J As Integer
Dim lIndex As Long

comma = ""
hasEtap = False
If beNaklads() Then Exit Sub

If Otgruz.loadOutDates Then 'началась отгрузка
    MsgBox "По заказу " & gNzak & " уже началась отгрузка.", , "Удаление невозможно!"
    Exit Sub
End If
If hasSelection(Grid5) Then
    For I = 1 To UBound(selectedItems)
        str2 = str2 & comma & Grid5.TextMatrix(CInt(selectedItems(I)), prName)
        comma = ", "
        If idWerk <> 1 Then
            str = Grid5.TextMatrix(CInt(selectedItems(I)), prEQuant)
            If IsNumeric(str) Then
                If CDbl(str) > 0 Then _
                    hasEtap = True
            End If
        End If
    Next I
    
End If
If Not convertToIzdelie Then
    If MsgBox("Вы хотите удалить позицию(и) '" & str2 & _
    "'", vbYesNo Or vbDefaultButton2, "Подтвердите удаление") = vbNo Then Exit Sub
'str = Grid5.TextMatrix(mousRow5, prEQuant)
End If

If hasEtap Then
  MsgBox "По одной или нескольким позициям были заведены этапы. " _
    & "Обнулите (если нужно предварительно запомнив) эти значения, возможно оно Вам понадобится, если Вы " _
    & "собираетесь тут же восстановить эти позиции и этап по ней."
    Exit Sub
End If

wrkDefault.BeginTrans

deleteSelected

wrkDefault.CommitTrans

tmpVar = saveOrdered(orderRate)
If Not IsNumeric(tmpVar) Then GoTo ER1
wrkDefault.CommitTrans

Grid5.TextMatrix(Grid5.Rows - 1, prSumm) = tmpVar
Orders.openOrdersRowToGrid "##220":    tqOrders.Close
    
For J = 1 To UBound(selectedItems)
    lIndex = useMaxSelection()
    quantity5 = quantity5 - 1
    Grid5.RemoveItem lIndex
    If quantity5 <= 0 Then
        clearGridRow Grid5, lIndex
    End If
Next J

If Not convertToIzdelie Then
    loadProducts ' ном-ра заказа
    Grid2.SetFocus ' здесь не срабатывает
End If
Exit Sub
ER0:
tbDMC.Close
ER1:
wrkDefault.Rollback
MsgBox "Удаление не прошло", , "Error 196" '##196
End Sub

Private Sub deleteSelected()
Dim I As Integer, J As Integer
Dim pQuant As Double

    For J = 1 To UBound(selectedItems)
        mousRow5 = CInt(selectedItems(J))
        getIdFromGrid5Row Me
        If Grid5.TextMatrix(mousRow5, prType) = "изделие" Then
            'удаление изд-я с возм.вариантной ном-рой (т.к.каскадное удаление)
        '    Set tbProduct = myOpenRecordSet("##138", "select * from xPredmetyByIzdelia", dbOpenForwardOnly)
            sql = "SELECT quant from xPredmetyByIzdelia " & _
            "WHERE (((numOrder)=" & numDoc & ") AND ((prId)=" & gProductId & _
            ") AND ((prExt)=" & prExt & "));"
            Set tbProduct = myOpenRecordSet("##138", sql, dbOpenForwardOnly)
        '    If tbProduct Is Nothing Then GoTo ER1
        '    tbProduct.index = "Key"
        '    tbProduct.Seek "=", numDoc, gProductId, prExt
        '    If tbProduct.NoMatch Then tbProduct.Close: GoTo ER1
            If tbProduct.BOF Then tbProduct.Close: Exit Sub
            pQuant = tbProduct!quant
            
            ReDim NN(0): ReDim QQ(0)
            productNomenkToNNQQ pQuant, 0, 0 ' д.б. перед удалением
            
            tbProduct.Delete
            tbProduct.Close
            
            'удаления ном-ры изделия из DMCrez
        '    Set tbDMC = myOpenRecordSet("##152", "select * from sDMCrez", dbOpenForwardOnly)
        '    If tbDMC Is Nothing Then Exit Sub
        '    tbDMC.index = "NomDoc"
            
            For I = 1 To UBound(NN)
                gNomNom = NN(I)
                If Not nomenkToDMCrez(-QQ(I)) Then Exit Sub
            Next I
         '   tbDMC.Close
        Else 'отдельная ном-ра
        
            'удаление этапа по позиции
            sql = "DELETE From xEtapByNomenk WHERE (((numOrder)=" & gNzak & _
            ") AND ((nomNom)='" & gNomNom & "'));"
            myExecute "##336", sql, 0 'если есть
            
            'удаления ном-ры из состава
            sql = "SELECT quant from xPredmetyByNomenk " & _
            "WHERE numOrder = " & gNzak & " AND nomNom = '" & gNomNom & "'"
            Set tbNomenk = myOpenRecordSet("##198", sql, dbOpenForwardOnly)
            
            If tbNomenk.BOF Then tbNomenk.Close: Exit Sub
            pQuant = tbNomenk!quant
            tbNomenk.Delete
            tbNomenk.Close
            
            'удаления ном-ры из DMCrez
            If Not nomenkToDMCrez(-pQuant) Then Exit Sub
        End If
    Next J

End Sub

'$odbc15$
Private Sub mnDel2_Click()
Dim S As Double, str As String
 
If Not (Regim = "fromDocs" And sDocs.Regim = "fromCeh") Then _
    If beNaklads() Then Exit Sub

If MsgBox("Удалить позицию № '" & gNomNom & _
"', Вы уверены?", vbYesNo Or vbDefaultButton2, "Подтвердите удаление") _
= vbNo Then GoTo EN1

If Regim = "fromDocs" And sDocs.Regim = "fromCeh" Then
    str = "rez":  If skladId = -1002 Then str = "mov"
    sql = "DELETE From sDMC" & str & " WHERE (((numDoc)=" & numDoc & _
    ") AND ((nomNom)='" & gNomNom & "'));"
    If myExecute("##341", sql) = 0 Then GoTo NX1
    Exit Sub
End If

wrkDefault.BeginTrans
sql = "SELECT quant from sDMC  WHERE (((numDoc)=" & numDoc & _
") AND ((numExt)=" & numExt & ") AND ((nomNom)='" & gNomNom & "'));"
'Set tbDMC = myOpenRecordSet("##123", "select * from sDMC", dbOpenForwardOnly)
Set tbDMC = myOpenRecordSet("##123", sql, dbOpenForwardOnly)
'If tbDMC Is Nothing Then GoTo ER1
'tbDMC.index = "NomDoc"
'tbDMC.Seek "=", numDoc, numExt, gNomNom
cErr = 179 '##179
'If tbDMC.NoMatch Then GoTo ER1
If tbDMC.BOF Then GoTo ER1
S = tbDMC!quant
tbDMC.Delete
tbDMC.Close

If ostatCorr(-S) Then
    wrkDefault.CommitTrans
Else
    cErr = 125 '##125
ER1: wrkDefault.Rollback
    MsgBox "Не прошла коррекция остатков. " & _
    "Сообщите администратору.", , "Error " & cErr
    GoTo EN1
End If
NX1:
quantity2 = quantity2 - 1
If quantity2 = 0 Then
    clearGridRow Grid2, 1
Else
    Grid2.RemoveItem mousRow2
End If

EN1: Grid2.SetFocus
End Sub

Private Sub mnOnfly_Click()
    OnFly.Show vbModal
    If convertToIzdelie Then
        loadPredmeti Me, orderRate, idWerk, asWhole
        loadProducts
        convertToIzdelie = False
    End If
End Sub

Private Sub opNomenk_Click()

controlEnable False
laQuant.Visible = False
laQuant.Caption = ""

laGrid.Visible = False
gridOrGrid3Hide "grid3"
laBegin.Visible = True

If Regim = "ostat" Then
    cmHide.Visible = True
    Grid.ColWidth(nkQuant) = 0
Else
    Grid.ColWidth(nkName) = 2970
    Grid.ColWidth(nkQuant) = 0
End If
If Me.WindowState = vbMaximized And Me.Width > cDELLwidth Then 'экран DELL
    Grid.ColWidth(nkName) = 4350
End If
    

laBegin = "В классификаторе выберите (кликом Mouse) группу, при этом " & _
"откроется таблица, где будет представлена вся номенклатура этой группы."
If Regim = "" Then
    laBegin = laBegin & _
vbCrLf & "      Выберите в этой таблице требуемую позицию и нажмите <Добавить>." & _
vbCrLf & vbCrLf & "При необходимости повторите эти действия для " & _
"других позиций."
Else
    laBegin = laBegin & vbCrLf & vbCrLf & "Если при этом удерживать <Ctrl>, то " & _
    "текущая группа будет добавлена к предыдущей группе." & vbCrLf & vbCrLf & _
    "Для удаления конкретной номенклатуры из предметов(справа) выберите в " & _
    "контексном меню(правый клик Mouse) коману 'удалить'"
End If
loadKlass

laGrid1.Caption = ""
cbInside.Enabled = True
End Sub

Sub controlEnable(en As Boolean)
If Not en Then ' только гасим
    Grid.Visible = False
    Grid3.Visible = False
    splLeftH.Visible = False
End If
If Regim <> "closeZakaz" Then cmSel.Enabled = en
End Sub

Private Sub opProduct_Click()
cmHide.Visible = False

controlEnable False
laQuant.Visible = True
laQuant.Caption = "изделий"
laBegin.Visible = True

Grid3.ColWidth(gpNN) = 0
Grid3.ColWidth(gpId) = 0

Grid.ColWidth(0) = 0
Grid.ColWidth(nkNomer) = 0 '900
If Regim = "ostat" Then
    Grid3.ColWidth(gpName) = 2085
    Grid3.ColWidth(gpSize) = 1080
    Grid.ColWidth(nkQuant) = 700
Else
    Grid3.ColWidth(gpName) = 1305
    Grid3.ColWidth(gpSize) = 840 '855
    Grid.ColWidth(nkQuant) = 700
End If
setNameColWidth
loadSeria tv
Dim str As String
str = ""

laBegin = "В левом списке выберите (кликом Mouse) серию, при этом откроется " & _
"таблица, где будут представлены все изделия этой серии." & vbCrLf & _
"     Кликните по строке изделия,  для просмотра входящей в него " & _
" номенклатуры."
If Regim = "ostat" Then
Else
    laBegin = laBegin & vbCrLf & "  Нажмите <Добавить>, установите требуемое " & _
    "количество изделия и нажмите <Enter>." & vbCrLf & vbCrLf & "При " & _
    "необходимости повторите эти действия для других изделий ."
End If
laGrid1.Caption = ""

If cbInside.Visible Then
    cbInside.ListIndex = 0 'Склад1
    cbInside.Enabled = False
End If
End Sub

Sub setNameColWidth()
If Me.WindowState = vbMaximized And Me.Width > cDELLwidth Then 'экран DELL
    Grid3.ColWidth(gpDescript) = 5055
    Grid.ColWidth(nkName) = 4200
ElseIf Regim = "ostat" Then
    Grid3.ColWidth(gpDescript) = 3495
Else
    Grid.ColWidth(nkName) = 2100
    Grid3.ColWidth(gpDescript) = 2200
End If

End Sub
'$odbc15$
Function nomenkToDMC(delta As Double, Optional noLock As String = "") As Boolean
Dim S As Double, I As Integer

nomenkToDMC = False

If noLock = "" Then
    If Not lockSklad Then Exit Function
End If

sql = "UPDATE sDMC SET quant = quant + " & delta & _
" WHERE numDoc=" & numDoc & " AND numExt=" & numExt & _
" AND nomNom='" & gNomNom & "';"
I = myExecute("W##123", sql, 0)
If I > 0 Then
    GoTo EN1
ElseIf I < 0 Then ' записи нет, поэтому добавляем
    sql = "INSERT INTO sDMC ( numDoc, numExt, nomNom, quant )" & _
    "SELECT " & numDoc & ", " & numExt & ", '" & gNomNom & "', " & delta & ";"
    Debug.Print sql
    If myExecute("##348", sql) <> 0 Then GoTo EN1
End If
'If noLock = "" Then tbDMC.Close

'корректируем остатки(для межскладских не корректирует)
If Not ostatCorr(delta) Then MsgBox "Не прошла коррекция остатков. " & _
     "Сообщите администратору.", , "Error 83" '##83
nomenkToDMC = True

EN1:
If noLock = "" Then lockSklad "un"
End Function

Sub nomenkToPredmeti(ByRef needToRefresh As Boolean)
Dim delta As Double, S As Double, quant As Double

    quant = tbQuant.Text
    Dim Nomnom1 As Nomnom
    Set Nomnom1 = nomnomCache.getNomnom(gNomNom)
    quant = Nomnom1.getQuantityRevert(quant, asWhole)
    S = nomencOstatkiToGrid(1) - quant ' одновременно обновляем таблицу
  
    If S < -0.005 Then 'в 2х местах
        If MsgBox("Дефицит товара '" & gNomNom & "' в доступных остатках " & _
        "составит (" & S & "), продолжить?", vbOKCancel Or vbDefaultButton2, _
        "Подтвердите") = vbCancel Then
            Exit Sub
        End If
    End If
    
On Error GoTo errr

    wrkDefault.BeginTrans
    sql = "select * from xPredmetyByNomenk where numOrder = " & numDoc & _
            " and nomNom = '" & gNomNom & "'"
    'Debug.Print sql
    Set tbNomenk = myOpenRecordSet("##117", sql, dbOpenForwardOnly)
    If tbNomenk Is Nothing Then End
    
    If Not tbNomenk.BOF Then
        tbNomenk.Edit
        tbNomenk!quant = tbNomenk!quant + quant
    Else
        tbNomenk.AddNew
        tbNomenk!Numorder = numDoc
        tbNomenk!Nomnom = gNomNom
        tbNomenk!quant = quant
    End If
    tbNomenk.update
    needToRefresh = True

    If Not nomenkToDMCrez(quant) Then GoTo ER1

    wrkDefault.CommitTrans
  GoTo EN2
ER1: wrkDefault.Rollback
EN2: tbNomenk.Close

Exit Sub

errr:

    tbNomenk.Close
    errorCodAndMsg ("Добавление номенклатуры к заказу")
End Sub



Private Sub splLeftH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    DraggingY = Y
End Sub


Private Sub splLeftH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging Then
        Dim DraggingShift As Single
        DraggingShift = Y
        If Not (Grid3.Height + DraggingShift > 100) Then
            Exit Sub
        End If
        If Not (Grid.Height - DraggingShift > 100) Then
            Exit Sub
        End If
        splLeftH.Top = splLeftH.Top + DraggingShift
        Grid3.Height = Grid3.Height + DraggingShift
        laGrid.Top = splLeftH.Top + splLeftH.Height
        Grid.Top = laGrid.Top + laGrid.Height
        Grid.Height = Grid.Height - DraggingShift
    End If
End Sub

Private Sub splLeftH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub



Private Sub splLeftV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    DraggingX = X
End Sub

Private Sub splLeftV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging Then
        Dim DraggingShift As Single
        DraggingShift = X
        If tv.Width + DraggingShift > 100 Then
        Else
            Exit Sub
        End If
        
        If Grid3.Width - DraggingShift > 100 Then
        Else
            Exit Sub
        End If
            Grid3.Width = Grid3.Width - DraggingShift
            tv.Width = tv.Width + DraggingShift
        splLeftV.Left = splLeftV.Left + DraggingShift
        Grid3.Left = Grid3.Left + DraggingShift
        laGrid.Left = Grid3.Left
        laGrid.Width = Grid3.Width
        Grid.Left = Grid3.Left
        Grid.Width = Grid3.Width
        laBegin.Left = Grid3.Left
        If laBegin.Width > DraggingShift Then _
            laBegin.Width = Grid3.Width
        splLeftH.Left = Grid3.Left
        splLeftH.Width = Grid3.Width
    End If
End Sub

Private Sub splLeftV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub splRightH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    DraggingY = Y
End Sub

Private Sub splRightH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging Then
        Dim DraggingShift As Single
        DraggingShift = Y
        If Not (Grid5.Height + DraggingShift > 100) Then
            Exit Sub
        End If
        If Not (Grid2.Height - DraggingShift > 100) Then
            Exit Sub
        End If
        splRightH.Top = splRightH.Top + DraggingShift
        Grid5.Height = Grid5.Height + DraggingShift
        laGrid2.Top = splRightH.Top + splRightH.Height
        Grid2.Top = laGrid2.Top + laGrid2.Height
        Grid2.Height = Grid2.Height - DraggingShift
    End If
End Sub

Private Sub splRightH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub


Private Sub splRightV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    DraggingX = X
End Sub

Private Sub splRightV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging Then
        Dim DraggingShift As Single
        DraggingShift = X
        If Not (Grid3.Width + DraggingShift > 100) Then
            Exit Sub
        End If
        If Not (Grid5.Width - DraggingShift > 100) Then
            Exit Sub
        End If
            
        Grid3.Width = Grid3.Width + DraggingShift
        Grid.Width = Grid3.Width
        splRightV.Left = splRightV.Left + DraggingShift
        splLeftH.Width = Grid3.Width
        Grid5.Left = Grid5.Left + DraggingShift
        Grid5.Width = Grid5.Width - DraggingShift
        laGrid2.Left = Grid5.Left
        laGrid2.Width = Grid5.Width
        Grid2.Left = Grid5.Left
        Grid2.Width = Grid5.Width
        splRightH.Left = Grid5.Left
        splRightH.Width = Grid5.Width
    End If
End Sub

Private Sub splRightV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

'$odbc15$
Public Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim c As Double, S As Double, str As String

If KeyCode = vbKeyReturn Then
    getIdFromGrid5Row Me
    str = Grid5.TextMatrix(mousRow5, prType)
    If str = "изделие" Then
        strWhere = " WHERE xPredmetyByIzdelia.numOrder = " & gNzak & _
        " AND xPredmetyByIzdelia.prId = " & gProductId & _
        " AND xPredmetyByIzdelia.prExt = " & prExt
    Else
        strWhere = " WHERE xPredmetyByNomenk.numOrder = " & gNzak & _
        " AND xPredmetyByNomenk.nomNom = '" & gNomNom & "'"
    End If
    If mousCol5 = prEQuant Then 'prEtap
        If str = "изделие" Then
            sql = "SELECT xEtapByIzdelia.prevQuant, xPredmetyByIzdelia.quant " & _
            "FROM xEtapByIzdelia RIGHT JOIN xPredmetyByIzdelia ON " & _
            "(xEtapByIzdelia.prExt = xPredmetyByIzdelia.prExt) AND " & _
            "(xEtapByIzdelia.prId = xPredmetyByIzdelia.prId) AND " & _
            "(xEtapByIzdelia.numOrder = xPredmetyByIzdelia.numOrder)" & strWhere
'            tmpStr = "xEtapByIzdelia"
        Else
            sql = "SELECT xEtapByNomenk.prevQuant, xPredmetyByNomenk.quant " & _
            "FROM xEtapByNomenk RIGHT JOIN xPredmetyByNomenk ON " & _
            "(xEtapByNomenk.nomNom = xPredmetyByNomenk.nomNom) AND " & _
            "(xEtapByNomenk.numOrder = xPredmetyByNomenk.numOrder) " & strWhere
'            tmpStr = "xEtapByNomenk"
        End If
        If Not byErrSqlGetValues("##315", sql, c, S) Then Exit Sub
        If Not isNumericTbox(tbMobile, 0, Round(S - c, 2)) Then Exit Sub
        
        S = tbMobile.Text: S = Round(S + c, 2)
        
'        Set tbProduct = myOpenRecordSet("##316", tmpStr, dbOpenTable)
'        If tbProduct Is Nothing Then myBase.Close: End
'        tbProduct.index = "Key"
        If str = "изделие" Then
            sql = "SELECT * from xEtapByIzdelia WHERE numOrder = " & _
            gNzak & " AND prId = " & gProductId & " AND prExt = " & prExt
'            tbProduct.Seek "=", gNzak, gProductId, prExt
        Else
            sql = "SELECT * from xEtapByNomenk WHERE numOrder = " & _
            gNzak & " AND nomNom = '" & gNomNom & "'"
'            tbProduct.Seek "=", gNzak, gNomNom
        End If
        Set tbProduct = myOpenRecordSet("##316", sql, dbOpenTable)
'        If Not tbProduct.NoMatch Then
        If Not tbProduct.BOF Then
            If S < 0.005 Then
                tbProduct.Delete
                Grid5.TextMatrix(mousRow5, prEtap) = ""
                Grid5.TextMatrix(mousRow5, prEQuant) = ""
            Else
                tbProduct.Edit
                GoTo AA
            End If
        ElseIf S > 0.005 Then
            tbProduct.AddNew
            tbProduct!Numorder = gNzak
            If str = "изделие" Then
                tbProduct!prId = gProductId
                tbProduct!prExt = prExt
            Else
                tbProduct!Nomnom = gNomNom
            End If
AA:         tbProduct!eQuant = S
            tbProduct.update
            Grid5.TextMatrix(mousRow5, prEtap) = S
            Grid5.TextMatrix(mousRow5, prEQuant) = tbMobile.Text
        End If
        tbProduct.Close
            
        lbHide
        Exit Sub
    End If
    
    If Not Me.convertToIzdelie Then
        If Not isNumericTbox(tbMobile, 0) Then Exit Sub

        Dim tunedCenaEd As Double, preciseQuantity As Double
        If Grid5.TextMatrix(mousRow5, prType) = "изделие" Then
            preciseQuantity = Grid5.TextMatrix(mousRow5, prQuant)
        Else
            If asWhole = 1 Then
                sql = "select pn.quant / n.perlist from xPredmetyByNomenk pn join sguidenomenk n on n.nomnom = pn.nomnom where pn.numorder = " & gNzak & " AND pn.nomnom = '" & gNomNom & "'"
            Else
                sql = "select pn.quant from xPredmetyByNomenk pn where pn.numorder = " & gNzak & " AND pn.nomnom = '" & gNomNom & "'"
            End If
            byErrSqlGetValues "##500", sql, preciseQuantity
        End If

        If mousCol5 = prSumm Then
            S = tuneCurencyAndGranularity(tbMobile.Text, orderRate, sessionCurrency, preciseQuantity)
            tunedCenaEd = S / preciseQuantity 'не округлять
        Else
            tunedCenaEd = tuneCurencyAndGranularity(tbMobile.Text, orderRate, sessionCurrency, 1)
            S = tunedCenaEd * preciseQuantity
        End If
        
        If str = "изделие" Then
            sql = "UPDATE xPredmetyByIzdelia SET cenaEd = " & tunedCenaEd & _
            "  WHERE numOrder = " & gNzak & " AND prId = " & gProductId & _
            " AND prExt = " & prExt
        Else
            If idWerk <> 1 Or asWhole = 0 Then
                sql = "UPDATE xPredmetyByNomenk SET cenaEd = " & tunedCenaEd & _
                " WHERE numOrder = " & gNzak & " AND nomNom = '" & gNomNom & "'"
            Else
                sql = "UPDATE xPredmetyByNomenk SET cenaEd = " & tunedCenaEd & "/ n.perList" _
                & " FROM sGuideNomenk n " _
                & " WHERE n.nomnom = xPredmetyByNomenk.nomnom and xPredmetyByNomenk.numOrder = " & gNzak & " AND xPredmetyByNomenk.nomNom = '" & gNomNom & "'"
            End If
        End If
'        MsgBox sql
        
        If myExecute("##205", sql) = 0 Then
            Grid5.TextMatrix(mousRow5, prCenaEd) = Round(rated(tunedCenaEd, orderRate), 2)
            Grid5.TextMatrix(mousRow5, prSumm) = Round(rated(S, orderRate), 2)
            tmpVar = saveOrdered(orderRate)
            If IsNumeric(tmpVar) Then
                Grid5.TextMatrix(Grid5.Rows - 1, prSumm) = Round(rated(tmpVar, orderRate), 2)
                Otgruz.saveShipped 'цена влияет и на отгрузку
                Orders.openOrdersRowToGrid "##212"
                tqOrders.Close
            End If
        End If
        lbHide
    End If
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Sub lbHide()
tbMobile.Visible = False
Grid5.Enabled = True
Grid5.SetFocus
Grid5_EnterCell
End Sub

'ед.измер delta и Дефицита зависит от наличия в накладной целых складов
'т.о. значение для delta может напрямую браться из tbQuant
Function deficitAndNoIgnore(delta As Double) As Boolean
Dim S As Double, IL As Long


deficitAndNoIgnore = False
S = nomencOstatkiToGrid(IL) - delta ' одновременно обновляем таблицу
If S < -0.005 Then
    If numExt = 254 Or numExt = 0 Then ' накладная или выпис. из Цеха
        tmpStr = "' по подразделению '" & sDocs.getGridColSour() & "'"
    Else
        tmpStr = "' в доступных остатках"
    End If
    If MsgBox("Дефицит товара '" & gNomNom & tmpStr & " составит (" & _
    S & "), продолжить?", vbOKCancel Or vbDefaultButton2, "Подтвердите") _
    = vbOK Then Exit Function
    deficitAndNoIgnore = True
End If
End Function

'$odbc15$
Private Sub tbQuant_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rr As Long, IL As Long, pQuant As Double, S As Double, str As String
Dim I As Integer, NN2() As String
Dim needToRefresh As Boolean ' показывает, нужно ли обновить сумму в поле ordered



If KeyCode = vbKeyReturn Then
Dim Nomnom1 As Nomnom


    
 
If opNomenk.Value Then
  If Not isNumericTbox(tbQuant, 0.01) Then Exit Sub
  S = Round(tbQuant.Text, 2)
  If S <> tbQuant.Text Then
    MsgBox "Число знаков после запятой - не больше двух!", , "Повторите ввод"
    tbQuant.SetFocus
    Exit Sub
  End If
  If Regim = "" Then 'предметы к заказу т.е 0<numExt<254
    nomenkToPredmeti needToRefresh
  Else
    S = Round(tbQuant.Text, 2)
    If sDocs.isIntMove() Then
        I = Round(S, 0)
        If S <> I Then
            MsgBox "Количество должно быть целым!", , "Повторите ввод"
            tbQuant.Text = "1": Exit Sub
        End If
        
        Set Nomnom1 = nomnomCache.getNomnom(gNomNom)
        S = Nomnom1.perlist
        S = Round(I * S, 2)
    End If
    If deficitAndNoIgnore(tbQuant.Text) Then Exit Sub
    If numExt = 0 And sDocs.reservNoNeed Then
        ' выписанные из цеха со скл.Обрезков и межскладские -  не резервируем
        nomenkToDMCrez S, "mov"
    Else 'из sDocs
'        If deficitAndNoIgnore(tbQuant.Text) Then Exit Sub
        If numExt = 254 Then
            nomenkToDMC S
        Else ' выписанная из цеха (со Скл.Целых)
            nomenkToDMCrez S
        End If
    End If
  End If
Else ' здесь только предметы заказы, поэтому всегда в одной(мелкой) ed.izmer
  If Not isNumericTbox(tbQuant, 1) Then Exit Sub
  pQuant = Round(tbQuant.Text)
  tbQuant.Text = pQuant
  
  If Not lockSklad Then Exit Sub
  
  Grid.col = nkQuant: IL = 0
  For rr = 1 To Grid.Rows - 1
    Grid.row = rr
    If Grid.CellFontBold Then
      IL = IL + 1
      gNomNom = Grid.TextMatrix(rr, nkNomer)
      
      S = Grid.TextMatrix(rr, nkQuantPricise) * pQuant
      If deficitAndNoIgnore(S) Then GoTo ER2
    End If
  Next rr
  
  wrkDefault.BeginTrans
  
'***** добавляем предметы в Cостав по номенклатуре (для цеха)*********************
  str = "sDMCrez"
  If numExt = 254 Then str = "sDMC" 'numExt = 254 эквивалентно Regim="fromDocs"
'  Set tbDMC = myOpenRecordSet("##152", str, dbOpenTable)
'  If tbDMC Is Nothing Then GoTo ER1
'  tbDMC.index = "NomDoc"
  
  Grid.col = nkQuant
  I = 0: ReDim NN(0)
  For rr = 1 To Grid.Rows - 1
    Grid.row = rr
    If Grid.CellFontBold Then
      gNomNom = Grid.TextMatrix(rr, nkNomer)
      If Grid.CellBackColor = groupColor1 Or Grid.CellBackColor = groupColor2 Then
        I = I + 1: ReDim Preserve NN(I): NN(I) = gNomNom 'вариантная ном-ра
      End If
      S = Grid.TextMatrix(rr, nkQuantPricise) * pQuant
      If numExt = 254 Then ' номенклатура  накладной
        If Not nomenkToDMC(S, "noLock") Then GoTo ER0 '
      Else '   резервирование заказа либо предметов накладной из для цеха(межскладские здесь невозможны)
        If Not nomenkToDMCrez(S) Then GoTo ER0
      End If
    End If
  Next rr
'  tbDMC.Close

If sDocs.Regim <> "fromCeh" Then
'******** добавляем предметы в Cостав по изделиям  *************************
  If numExt <> 254 Then ' этот состав только для заказов
      quickSort NN, 1
      If Not addToPredmetiTable(pQuant, getPrExtByNomenk(), needToRefresh) Then GoTo ER1
  End If
'*************************************************************************
End If
  wrkDefault.CommitTrans
  
  lockSklad "un"
End If ' opNomenk.value

If Regim = "" Then
    loadPredmeti Me, orderRate, idWerk, asWhole, , needToRefresh ' состав
End If
    
loadProducts ' ном-ра заказа
   
'  Grid2.col = dnQuant
'  Grid2.SetFocus
Grid.SetFocus
GoTo ES2

ER0:
'tbDMC.Close
ER1:
wrkDefault.Rollback
ER2:
lockSklad "un"
GoTo ESC

ElseIf KeyCode = vbKeyEscape Then
ESC: tbQuant.Text = ""
ES2: gridFrame.Visible = False
    tbQuant.Enabled = False
    laQuant.Enabled = False
    cmSel.Enabled = True
End If

End Sub
'$odbc15$
'при delta < 0 - возм. удаление
Function nomenkToDMCrez(ByVal delta As Double, Optional mov As String = "") As Boolean
Dim S As Double, I As Integer

nomenkToDMCrez = False
'    If mov = "mov" Then ' выписанные межскладские не резервируем
'        Set tbDMC = myOpenRecordSet("##152", "select * from sDMCmov", dbOpenTable)
'        GoTo AA
'    ElseIf mov = "" Then
'        Set tbDMC = myOpenRecordSet("##152", "select * from sDMCrez", dbOpenTable)
'AA:     If tbDMC Is Nothing Then Exit Function
'        tbDMC.index = "NomDoc"
'    End If
'        tbDMC.Seek "=", numDoc, gNomNom
'        If tbDMC.NoMatch Then
'            If delta < 0 Then 'когда у закрытого заказа(в sDMCrez его уже нет)
'                GoTo EN1      'при его аннулировании удаляются предметы
'                msgOfEnd ("##195")
'            End If
'            tbDMC.AddNew
'            tbDMC!numDoc = numDoc
'            tbDMC!nomNom = gNomNom
'            tbDMC!quantity = Round(delta, 2)
'        Else
'            s = Round(tbDMC!quantity + delta, 2)
'            If s <= 0 Then tbDMC.Delete: GoTo EN1
'            tbDMC.Edit
'            tbDMC!quantity = s
'        End If
'        tbDMC.Update

' выписанные межскладские не резервируем
strWhere = " WHERE numDoc = " & numDoc & " AND nomNom='" & gNomNom & "'"

If mov = "" Then mov = "rez"

sql = "SELECT quantity FROM sDMC" & mov & strWhere
If Not byErrSqlGetValues("W##423", sql, S) Then Exit Function

If S < 0.01 Then ' такой ном-ры еще в таблице sDMCRez нет
  If delta >= 0.01 Then
    sql = "INSERT INTO sDMC" & mov & " ( numDoc, nomNom, quantity ) " & _
    "SELECT " & numDoc & ", '" & gNomNom & "', " & delta & ";"
    If myExecute("##195", sql) <> 0 Then Exit Function
  End If
Else
    delta = S + delta

    If Round(delta, 3) >= 0.01 Then
        sql = "UPDATE sDMC" & mov & " SET quantity = " & delta & strWhere
        If myExecute("##152", sql) <> 0 Then Exit Function
    Else
        sql = "DELETE FROM sDMC" & mov & strWhere
        If myExecute("##424", sql) <> 0 Then Exit Function
    End If
End If
        
'EN1:    If mov = "" Or mov = "mov" Then tbDMC.Close
nomenkToDMCrez = True
End Function



'$odbc15!$
'lastExt=0 - если у изделия вообще нет вариантов поставки(или оно не вариантно)
'если вариант поставки, заданный в NN() есть, то дает номер его расширения
'иначе возвращает отриц.  макс.номер варианта поставки
Function addToPredmetiTable(pQuant As Double, lastExt As Integer, ByRef needToRefresh As Boolean) As Boolean
Dim I As Integer

addToPredmetiTable = False

If UBound(NN) = 0 Then 'невариантное изд-е, у него numExt(а потому и lastExt)=0
    If lastExt <> 0 Then msgOfEnd "##190", "lastExt=" & lastExt
    prExt = 0
ElseIf lastExt = 0 Then 'вариантное изд-е, но его пока вообще не было
    prExt = 1
ElseIf lastExt < 0 Then ' есть варианты этого изд-я, но другие
    prExt = -lastExt + 1 ' тогда делаем след.вариант
Else ' есть именно этот вариант
    prExt = lastExt
End If


sql = "SELECT * from xPredmetyByIzdelia " & _
"WHERE numOrder =" & numDoc & " AND prId = " & gProductId & _
" AND prExt =" & prExt
'MsgBox sql
Set tbProduct = myOpenRecordSet("##185", sql, dbOpenForwardOnly)
On Error GoTo errr


wrkDefault.BeginTrans

'Debug.Print sql
If tbProduct.BOF Then
    If lastExt > 0 Then msgOfEnd "##317", "lastExt=" & lastExt
    tbProduct.AddNew
    tbProduct!Numorder = numDoc
    tbProduct!prId = gProductId
    tbProduct!prExt = prExt
    tbProduct!quant = pQuant
    tbProduct.update
' вариантная ном-ра изделий
  If UBound(NN) > 0 Then
    Set tbNomenk = myOpenRecordSet("##191", "select * from xVariantNomenc", dbOpenForwardOnly)
    If tbNomenk Is Nothing Then End

    For I = 1 To UBound(NN)
        tbNomenk.AddNew
        tbNomenk!Numorder = numDoc
        tbNomenk!prId = gProductId
        tbNomenk!prExt = prExt
        tbNomenk!Nomnom = NN(I)
        tbNomenk.update
    Next I
    tbNomenk.Close
  End If
  wrkDefault.CommitTrans
    
Else
    If lastExt < 0 Then msgOfEnd "##428", "lastExt=" & lastExt
    tbProduct.Edit
    tbProduct!quant = Round(tbProduct!quant + pQuant)
    tbProduct.update
    needToRefresh = True
End If
'EP:
tbProduct.Close



wrkDefault.CommitTrans
addToPredmetiTable = True
Exit Function

errr:
errorCodAndMsg ("Добавление варианта")
End Function


'обновляет поле ordered в Orders
Function saveOrdered(orderRate As Double, Optional update As Boolean = True) As Variant
Dim S As Double, s1 As Double

    saveOrdered = Null
    sql = "SELECT Sum([quant]*[cenaEd]) From xPredmetyByIzdelia GROUP BY numOrder " & _
    "HAVING (((numOrder)=" & gNzak & "));"
    If Not byErrSqlGetValues("W##368", sql, S) Then Exit Function
    
    sql = "SELECT Sum([quant]*[cenaEd]) From xPredmetyByNomenk GROUP BY numOrder " & _
    "HAVING (((numOrder)=" & gNzak & "));"
    'MsgBox sql
    If Not byErrSqlGetValues("W##210", sql, s1) Then Exit Function
    
    S = S + s1
    If update Then
        orderUpdate "##211", CStr(S), "orders", "ordered"
        Orders.Grid.TextMatrix(Orders.mousRow, orZakazano) = Round(S, 2)
    End If
    saveOrdered = S

End Function

'=0 - если у изделия вообще нет вариантов поставки(или оно не вариантно или
'такого изд-я нет) если вариант поставки, заданный в NN() есть, то дает номер
'его расширения, иначе возвращает отриц.  макс.номер варианта поставки
Function getPrExtByNomenk() As Integer
Dim I As Integer, J As Integer, prevExt As Integer

getPrExtByNomenk = 0 '
sql = "SELECT xVariantNomenc.prExt, xVariantNomenc.nomNom " & _
"From xVariantNomenc WHERE (((xVariantNomenc.numOrder)=" & numDoc & _
") AND ((xVariantNomenc.prId)=" & gProductId & ")) ORDER BY xVariantNomenc.prExt;"

Set tbNomenk = myOpenRecordSet("##187", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then myBase.Close: End
If Not tbNomenk.BOF Then
    J = 0: ReDim NN2(UBound(NN))

CC: J = J + 1
    If J <= UBound(NN) Then
        NN2(J) = tbNomenk!Nomnom
    End If
    prevExt = tbNomenk!prExt
    tbNomenk.MoveNext
    If tbNomenk.EOF Then GoTo AA:
    If prevExt <> tbNomenk!prExt Then
AA:     If J = UBound(NN) Then ' совпадает-ли кол-во(Нет - если поменяли состав)
            quickSort NN2, 1
            For I = 1 To UBound(NN)
                If NN(I) <> NN2(I) Then GoTo BB
            Next I
            getPrExtByNomenk = prevExt
            GoTo en
        End If
BB:     J = 0
    End If
    If Not tbNomenk.EOF Then GoTo CC

    getPrExtByNomenk = -prevExt
End If
en:
tbNomenk.Close
End Function

Private Sub Timer1_Timer()
biloG3Enter_Cell = False
Timer1.Enabled = False
End Sub

Private Sub tv_AfterLabelEdit(Cancel As Integer, NewString As String)
gSeriaId = Mid$(tv.SelectedItem.Key, 2)
ValueToTableField "##115", "'" & NewString & "'", "sGuideSeries", "seriaName", "bySeriaId"
End Sub

Sub loadKlass()
Dim Key As String, pKey As String, K() As String, pK()  As String
Dim I As Integer, iErr As Integer
bilo = False
sql = "SELECT sGuideKlass.*  From sGuideKlass ORDER BY sGuideKlass.parentKlassId;"
Set tbKlass = myOpenRecordSet("##102", sql, dbOpenForwardOnly)
If tbKlass Is Nothing Then myBase.Close: End
If Not tbKlass.BOF Then
 tv.Nodes.Clear
 Set Node = tv.Nodes.Add(, , "k0", "Классификатор")
 Node.Sorted = True
 Set Node = tv.Nodes.Add("k0", tvwChild, "all", "              ")

 ReDim K(0): ReDim pK(0): ReDim NN(0): iErr = 0
 While Not tbKlass.EOF
    If tbKlass!klassId = 0 Then GoTo NXT1
    Key = "k" & tbKlass!klassId
    pKey = "k" & tbKlass!parentKlassId
    On Error GoTo ERR1 ' назначить второй проход
    Set Node = tv.Nodes.Add(pKey, tvwChild, Key, tbKlass!klassName)
    On Error GoTo 0
    Node.Sorted = True
NXT1:
    tbKlass.MoveNext
 Wend
  tv.Nodes.Item("all").Text = "Весь перечень"
End If
tbKlass.Close

While bilo ' необходимы еще проходы
  bilo = False
  For I = 1 To UBound(K())
    If K(I) <> "" Then
        On Error GoTo ERR2 ' назначить еще проход
        Set Node = tv.Nodes.Add(pK(I), tvwChild, K(I), NN(I))
        On Error GoTo 0
        K(I) = ""
        Node.Sorted = True
    End If
NXT:
  Next I
Wend
tv.Nodes.Item("k0").Expanded = True
Exit Sub
ERR1:
 iErr = iErr + 1: bilo = True
 ReDim Preserve K(iErr): ReDim Preserve pK(iErr): ReDim Preserve NN(iErr)
 K(iErr) = Key: pK(iErr) = pKey: NN(iErr) = tbKlass!klassName
 Resume Next

ERR2: bilo = True: Resume NXT

End Sub

Sub loadProductNomenk(ByVal v_productId As Integer)
Dim S As Double, grBef As String
If Me.Regim = "" Then
    buntColumn = nkQuant
ElseIf Me.Regim = "fromDocs" Then
    buntColumn = nkDostup
ElseIf Me.Regim = "ostat" Then
    buntColumn = nkQuant
Else
    buntColumn = 0
End If
Dragging = False

Me.MousePointer = flexHourglass

quantity = 0
Grid.Visible = False
clearGrid Grid

sql = "SELECT p.nomNom, p.quantity, p.xGroup, " & _
" n.Size, n.cod, n.nomName, n.ed_Izmer, n.ed_Izmer2, n.perList, n.cod, n.size, n.ves" & _
" FROM sGuideNomenk n " & _
" JOIN sProducts p ON n.nomNom = p.nomNom " & _
" WHERE p.ProductId =" & v_productId & _
" ORDER BY p.xGroup DESC"

Dim Nomnom1 As Nomnom

Set tbNomenk = myOpenRecordSet("##108", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then Exit Sub
If Not tbNomenk.BOF Then
  Grid.col = buntColumn
   grBef = ""
  While Not tbNomenk.EOF
    
    Grid.row = quantity 'возможно пред. строку н.б. окрасить
    Dim str As String: str = Grid.Text
    quantity = quantity + 1
    If grBef = tbNomenk!xGroup And grBef <> "" Then
        Grid.CellBackColor = grColor ' предыдущая
        Grid.CellFontBold = False
        Grid.row = quantity
        Grid.CellBackColor = grColor ' текущая
        Grid.CellFontBold = False
        bilo = False ' не было переключение цвета
    Else
        Grid.row = quantity
        Grid.CellFontBold = True
        If Not bilo Then '
            If grColor = groupColor1 Then
                grColor = groupColor2
            Else
                grColor = groupColor1
            End If
            bilo = True
        End If
    End If
    grBef = tbNomenk!xGroup
        
    gNomNom = tbNomenk!Nomnom
    nomnomCache.evict (tbNomenk!Nomnom)
    Set Nomnom1 = nomnomCache.getNomnom(tbNomenk!Nomnom, True)
    
    Grid.TextMatrix(quantity, nkNomer) = gNomNom
    Grid.TextMatrix(quantity, nkName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid.TextMatrix(quantity, nkEdIzm) = Nomnom1.getEdizm(asWhole)
    Grid.TextMatrix(quantity, nkQuant) = Nomnom1.getQuantity(tbNomenk!quantity, asWhole)
    Grid.TextMatrix(quantity, nkQuantPricise) = Round(tbNomenk!quantity, 5)
    
    Nomnom1.quantInProduct = tbNomenk!quantity
    'доступные остатки:
    Grid.TextMatrix(quantity, nkDostup) = Nomnom1.getQuantity(nomencOstatkiToGrid(-1, True), asWhole)
    If Regim = "ostat" Then
        Grid.TextMatrix(quantity, nkCurOstat) = Round(FO, 2)
    End If
    
    Grid.AddItem ""
    tbNomenk.MoveNext
  Wend
  Grid.RemoveItem quantity + 1
End If
tbNomenk.Close
Grid.Visible = True
Grid.ZOrder
Me.MousePointer = flexDefault

End Sub




Sub loadKlassNomenk()
Dim IL As Long, S As Double, strWhere As String
Dim beg As Double, prih As Double, rash As Double, oldNow As Double
buntColumn = 0



If tv.SelectedItem.Key = "all" Then
    strWhere = ""
    quantity = 0
Else
    strWhere = "WHERE klassId = " & gKlassId
End If

laGrid1.Caption = "Номенклатура из группы '" & tv.SelectedItem.Text & "'"

If (Regim = "ostat" And cbInside.ListIndex = 1) Or Regim = "fromDocs" Then
    Grid.ColWidth(nkDostup) = 0 'для склада Обрезков не показываем
Else
    Grid.ColWidth(nkDostup) = 700
End If

Me.MousePointer = flexHourglass

If beShift Then
    Grid.AddItem ""
Else
    quantity = 0
    Grid.Visible = False
    clearGrid Grid
End If


sql = "SELECT nomNom, nomName, Size, cod, ed_Izmer, ed_Izmer2, perList, ves " & _
"From sGuideNomenk " & strWhere & " ORDER BY sGuideNomenk.nomNom"

'Debug.Print sql
Dim Nomnom1 As Nomnom

Set tbNomenk = myOpenRecordSet("##103", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  tbNomenk.MoveFirst
  While Not tbNomenk.EOF

    gNomNom = tbNomenk!Nomnom
    Set Nomnom1 = nomnomCache.getNomnom(gNomNom, True)
    
    quantity = quantity + 1
    Grid.TextMatrix(quantity, nkNomer) = gNomNom
    Grid.TextMatrix(quantity, nkName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    
    S = nomencOstatkiToGrid(-1, True)  'доступные остатки
    Grid.TextMatrix(quantity, nkEdIzm) = Nomnom1.getEdizm(asWhole)
    
    If Regim = "fromDocs" Then
        If sDocs.isIntMove() Then GoTo AA
        GoTo NXT:
    End If
    Grid.TextMatrix(quantity, nkCurOstat) = Nomnom1.getQuantity(FO, asWhole)
    Grid.TextMatrix(quantity, nkDostup) = Nomnom1.getQuantity(S, asWhole)
    If Regim = "ostat" Then
      If cbInside.ListIndex = 0 Then 'склад1
        Grid.TextMatrix(quantity, nkDostup) = Round(nomencOstatkiToGrid(-1, True) / tbNomenk!perlist, 2)
        Grid.TextMatrix(quantity, nkCurOstat) = Round(FO / tbNomenk!perlist, 2)
AA:     Grid.TextMatrix(quantity, nkEdIzm) = tbNomenk!ed_Izmer2
      End If
    End If
NXT:
    Grid.AddItem ""
    tbNomenk.MoveNext
    'Unload Nomnom1
 Wend
 If quantity > 0 Then Grid.RemoveItem quantity + 1
End If
tbNomenk.Close

Grid.Visible = True
EN1:
Me.MousePointer = flexDefault
End Sub

Sub loadSeriaProduct()
Dim IL As Long, strWhere As String

If tv.SelectedItem.Key = "k0" Then
    gSeriaId = 0
    Grid3.Visible = False
    splLeftH.Visible = False
    Exit Sub
End If

Me.MousePointer = flexHourglass
laGrid1.Caption = "Список готовых изделий по серии '" & tv.SelectedItem.Text & "'"

quantity3 = 0
Grid3.Visible = False
splLeftH.Visible = False
clearGrid Grid3
IL = 0

strWhere = " WHERE sGuideProducts.prSeriaId = " & gSeriaId _
    & " and exists (select 1 from sproducts where sproducts.productid = sGuideProducts.prId)"

sql = "SELECT sGuideProducts.prDescript, sGuideProducts.prName, " & _
"sGuideProducts.prId, sGuideProducts.prSize From sGuideProducts " & strWhere & _
"ORDER BY sGuideProducts.SortNom;"
Set tbProduct = myOpenRecordSet("##104", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then GoTo EN1
If Not tbProduct.BOF Then
 While Not tbProduct.EOF
    quantity3 = quantity3 + 1
    
    Grid3.TextMatrix(quantity3, gpNN) = quantity3
    Grid3.TextMatrix(quantity3, gpId) = tbProduct!prId
'    If Not IsNull(tbProduct!prName) Then
    Grid3.TextMatrix(quantity3, gpName) = tbProduct!prName
    Grid3.TextMatrix(quantity3, gpSize) = tbProduct!prSize
    Grid3.TextMatrix(quantity3, gpDescript) = tbProduct!prDescript

    Grid3.AddItem ""
    tbProduct.MoveNext
 Wend
 Grid3.RemoveItem quantity3 + 1
End If
tbProduct.Close
Grid3.Visible = True
splLeftH.Visible = True

EN1:
Me.MousePointer = flexDefault

End Sub
    
Private Sub tv_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer, str As String
If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
    tv_NodeClick tv.SelectedItem
End If
End Sub


Private Sub tv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
beShift = False
If Shift = 2 Then beShift = True

End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)

If tv.SelectedItem.Key = "k0" Then
    controlEnable False
    quantity = 0
    laGrid.Caption = ""
    Exit Sub
End If

tbQuant.Enabled = False
laQuant.Enabled = False
laBegin.Visible = False
If opProduct.Value Then
    gSeriaId = Mid$(tv.SelectedItem.Key, 2)
    loadSeriaProduct
    Grid3.Visible = True
    splLeftH.Visible = True
    Grid.Visible = False
    laGrid.Visible = False
    cmSel.Enabled = False
    gridOrGrid3Hide "grid"
Else
    controlEnable True
    gKlassId = Mid$(tv.SelectedItem.Key, 2)
    loadKlassNomenk
'    Grid.Visible = True
End If
Grid_EnterCell
On Error Resume Next
Grid.SetFocus

biloG3Enter_Cell = False

End Sub


