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
   Begin VB.Frame splLeftH 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H8000000A&
      Height          =   80
      Left            =   2400
      MousePointer    =   7  'Size N S
      TabIndex        =   22
      Top             =   3360
      Width           =   3480
   End
   Begin VB.CommandButton cmExel2 
      Caption         =   "Печать в Exel"
      Height          =   315
      Left            =   8640
      TabIndex        =   17
      Top             =   5940
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   2700
      Top             =   4440
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   9900
      TabIndex        =   15
      Text            =   "tbMobile"
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Exel"
      Height          =   315
      Left            =   900
      TabIndex        =   13
      Top             =   5940
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame gridFrame 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   3180
      TabIndex        =   9
      Top             =   3420
      Visible         =   0   'False
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid Grid4 
         Height          =   1455
         Left            =   60
         TabIndex        =   10
         Top             =   300
         Visible         =   0   'False
         Width           =   7215
         _ExtentX        =   12721
         _ExtentY        =   2561
         _Version        =   393216
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Label laGrid 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   192
         Left            =   2400
         TabIndex        =   20
         Top             =   3360
         Width           =   492
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
         Left            =   60
         TabIndex        =   12
         Top             =   0
         Width           =   7212
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Если остатки позволяют, введите треб. кол-во изделий и нажмите <Enter>, иначе - <ESC>.."
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   1740
         Width           =   7215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2472
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   4692
      _ExtentX        =   8276
      _ExtentY        =   4360
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmHide 
      Caption         =   "Скрыть выд."
      Enabled         =   0   'False
      Height          =   315
      Left            =   5700
      TabIndex        =   7
      Top             =   5940
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox tbQuant 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Top             =   5940
      Width           =   735
   End
   Begin VB.CommandButton cmSel 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   5940
      Width           =   915
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   11040
      TabIndex        =   2
      Top             =   5940
      Width           =   795
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   5580
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   9843
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   5595
      Left            =   7200
      TabIndex        =   16
      Top             =   240
      Width           =   4635
      _ExtentX        =   8170
      _ExtentY        =   9864
      _Version        =   393216
      AllowBigSelection=   0   'False
      HighLight       =   0
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   3072
      Left            =   2400
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   4692
      _ExtentX        =   8276
      _ExtentY        =   5419
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Frame splRightV 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H8000000A&
      Height          =   5892
      Left            =   7080
      MousePointer    =   9  'Size W E
      TabIndex        =   21
      Top             =   0
      Width           =   80
   End
   Begin VB.Frame splLeftV 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H8000000A&
      Height          =   5892
      Left            =   2280
      MousePointer    =   9  'Size W E
      TabIndex        =   23
      Top             =   0
      Width           =   80
   End
   Begin VB.Label laGrid1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Height          =   192
      Left            =   2760
      TabIndex        =   18
      Top             =   0
      Width           =   36
   End
   Begin VB.Label laBegin 
      BackColor       =   &H8000000A&
      Caption         =   "Label2"
      Height          =   4395
      Left            =   2760
      TabIndex        =   6
      Top             =   900
      Width           =   3795
   End
   Begin VB.Label laQuant 
      BackColor       =   &H8000000A&
      Caption         =   "изделий"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4860
      TabIndex        =   5
      Top             =   5985
      Width           =   675
   End
   Begin VB.Label laGrid2 
      BackColor       =   &H8000000A&
      Caption         =   "Номенклатурный состав предметов:"
      Height          =   195
      Left            =   7380
      TabIndex        =   1
      Top             =   30
      Width           =   3495
   End
   Begin VB.Menu mnContext 
      Caption         =   "Из состава предметов"
      Visible         =   0   'False
      Begin VB.Menu mnDel 
         Caption         =   "Удалить"
      End
      Begin VB.Menu mnCancel 
         Caption         =   "Отменить"
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
Public orderRate As Double


Dim mousCol4 As Long, mousRow4 As Long
Dim msgBilo As Boolean, biloG3Enter_Cell As Boolean
Const groupColor1 = &HBBFFBB ' только не vbBottonFace
Const groupColor2 = &HBBBBFF '

Dim grColor As Long
Dim flag As Integer
'Dim maxNumExt As Integer, minNumExt As Integer

Dim mousCol As Long, mousRow As Long
Dim mousCol5 As Long, mousRow5 As Long
Dim Node As Node
Dim quantity  As Long, quantity2 As Long, quantity3 As Long
Public quantity5 As Long
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim tvVes As Double, gridVes As Double, grid2Ves As Double 'веса горизонт. размеров

Dim tbSeries As Recordset
Dim tbKlass As Recordset
Dim typeId As Integer
Dim beShift As Boolean
'Dim QP() As Double
'Dim VN() As String

'Grid4
Const frNomNom = 1
Const frNomName = 2
Const frEdIzm = 3
Const frOstat = 4

'список изделий или номенклатур(Grid3)
'Const gpNN = 0
Const gpName = 1
Const gpSize = 2
Const gpDescript = 3
Const gpId = 0 ' спрятан

'Grid2
Const dnNomNom = 1
Const dnNomName = 2
Const dnEdIzm = 3
Const dnVesEd = 4
Const dnCenaEd = 5
Const dnQuant = 6
Const dnSumm = 7
Const dnVes = 8

'номенклатура по группе или изделию(Grid)
Const nkNomer = 1
Const nkName = 2
Const nkEdIzm = 3
Const nkCurOstat = 4
Const nkDostup = 5
Const nkQuant = 6

Dim buntColumn As Integer

Dim Dragging As Boolean
Dim DraggingX As Single, DraggingY As Single

Private Function adjustGirdMoneyColWidth() As Long
Dim J As Integer
Dim ret As Long
' Расширить колонки для рублей
    For J = 1 To Grid2.Cols - 1
        If Grid2.colWidth(J) > 0 _
            And (InStr(Grid2.TextMatrix(0, J), "Цена") > 0 Or Grid2.TextMatrix(0, J) = "Сумма") _
        Then
            If sessionCurrency = CC_RUBLE Then
                Grid2.colWidth(J) = Grid2.colWidth(J) * ColWidthForRuble
            End If
        End If
        ret = ret + Grid2.colWidth(J)
    Next J
    adjustGirdMoneyColWidth = ret
End Function




Sub nomenkToNNQQ(pQuant As Double, eQuant As Double, prQuant As Double)
Dim J As Integer, leng As Integer

leng = UBound(NN)

    For J = 1 To leng
        If NN(J) = tbNomenk!nomNom Then
            QQ(J) = QQ(J) + pQuant * tbNomenk!quantity
            If eQuant > 0 Then _
                QQ2(J) = QQ2(J) + eQuant * tbNomenk!quantity
            If prQuant > 0 Then _
                QQ3(J) = QQ3(J) + prQuant * tbNomenk!quantity
            Exit Sub
        End If
    Next J
    leng = leng + 1
    ReDim Preserve NN(leng): NN(leng) = tbNomenk!nomNom
    ReDim Preserve QQ(leng): QQ(leng) = pQuant * tbNomenk!quantity
    ReDim Preserve QQ2(leng): QQ2(leng) = eQuant * tbNomenk!quantity
'    QQ2(leng) = 0: If eQuant > 0 Then QQ2(leng) = eQuant * tbNomenk!quantity
    ReDim Preserve QQ3(leng): QQ3(leng) = prQuant * tbNomenk!quantity
    

End Sub

Private Sub cmExel_Click()
'    GridToExcel Grid, laGrid1.Caption
    GridToExcel Grid, laGrid.Caption

End Sub

Private Sub cmExel2_Click()
    GridToExcel Grid2, laGrid2.Caption
End Sub

Private Sub cmExit_Click()
'If Regim = "ostat" Then
    Unload Me
'ElseIf checkRowsQuant Then
'    Unload Me
'End If
    
End Sub

Private Sub cmHide_Click()
Dim I As Integer
If quantity = 0 Then Exit Sub
For I = Grid.row To Grid.RowSel
    If Grid.Rows > 2 Then
        Grid.RemoveItem Grid.row
        quantity = quantity - 1
    End If
Next I
Grid.SetFocus
Grid_EnterCell
End Sub

Private Sub cmSel_Click() '<Добавить>
'Dim befColor As Long, il As Long, nl As Long, n As Integer, str As String

'laQuant.Visible = False
If Regim = "ostat" Or Regim = "products" Then Exit Sub
If beNaklads() Then Exit Sub

dostupOstatkiToGrid

tbQuant.Enabled = True
laQuant.Enabled = True
  tbQuant.Text = 1
  tbQuant.SelLength = 1
  tbQuant.SetFocus
  cmSel.Enabled = False
End Sub



Sub dostupOstatkiToGrid(Optional reg As String)
Dim s As Double, sum As Double, rr As Long, il As Long

Me.MousePointer = flexHourglass
'If numExt = 254 Then
laGrid4.Caption = "Доступные остатки"

clearGrid Grid4
Grid4.FormatString = "|<Номер|<Описание|<Ед.измерения|Oстатки"
Grid4.colWidth(0) = 0
Grid4.colWidth(frNomNom) = 870
Grid4.colWidth(frNomName) = 4485
Grid4.colWidth(frEdIzm) = 645
Grid4.colWidth(frOstat) = 885

nomencOstatkiToGrid 1

Grid4.Visible = True
EN1:
Me.MousePointer = flexDefault
gridFrame.Visible = True
gridFrame.ZOrder

End Sub

Public Function nomencOstatkiToGrid(row As Long) As Double
Dim s As Double, str As String, z As Double, str2 As String

'Ф.остатки
sql = "SELECT nomName, Ed_Izmer2, perList From sGuideNomenk " & _
"WHERE (((nomNom)='" & gNomNom & "'));"
'MsgBox sql
byErrSqlGetValues "##144", sql, str, str2, tmpSng
If row > 0 Then
    Grid4.TextMatrix(row, frNomNom) = gNomNom
    Grid4.TextMatrix(row, frNomName) = str
    Grid4.TextMatrix(row, frEdIzm) = str2
End If


'AA: вычисляем доступные остатки
FO = PrihodRashod("+", -1001) - PrihodRashod("-", -1001)
    
sql = "SELECT Sum(quantity) AS Sum_quantity, " & _
"Sum(Sum_quant) AS Sum_Sum_quant From wCloseNomenk" & _
" WHERE (((nomNom)='" & gNomNom & "'));"
If Not byErrSqlGetValues("##145", sql, z, s) Then myBase.Close: End
nomencOstatkiToGrid = FO - (z - s) ' минус, что несписано

nomencOstatkiToGrid = nomencOstatkiToGrid / tmpSng

If row > 0 Then _
    Grid4.TextMatrix(row, frOstat) = Round(nomencOstatkiToGrid, 2)

End Function

Private Sub Form_Load()
Dim str As String, I As Integer, delta As Double


controlEnable False
laQuant.Visible = False
laQuant.Caption = "лист"
splLeftH.Visible = False

Frame.Visible = False
    

If Regim = "products" Then
    loadSeria
    laBegin = "В классификаторе выберите (кликом Mouse) серию готовых изделий"
    Grid3.FormatString = "|<Номер|<Размер|<Описание"
    Grid3.colWidth(gpId) = 0
    Grid3.colWidth(gpName) = 1300
    Grid3.colWidth(gpSize) = 1080
    Grid3.colWidth(gpDescript) = 4085
    laGrid.Left = Grid3.Left

Else
    laBegin = "В классификаторе выберите (кликом Mouse) группу, при этом " & _
    "откроется таблица, где будет представлена вся номенклатура этой группы." & _
    vbCrLf & "      Выберите в этой таблице требуемую позицию и нажмите <Добавить>." & _
    vbCrLf & vbCrLf & "При необходимости повторите эти действия для " & _
    "других позиций."
    loadKlass
End If


noClick = False
msgBilo = False
Grid.FormatString = "|<Номер|<Описание|<Ед.изм|Ф.остатки|Д.остатки|Кол-во"
Grid2.FormatString = "|<Номер|<Описание|<Ед.измерения|вес.ед|Цена за ед.|" & _
                     "кол-во|Сумма|вес"
Grid.colWidth(0) = 0
Grid.colWidth(nkNomer) = 900
Grid.colWidth(nkEdIzm) = 630 'ostat
Grid.colWidth(nkCurOstat) = 0
Grid.colWidth(nkDostup) = 700
Grid.colWidth(nkQuant) = 0
cmExel.Visible = False
laBegin.Visible = True

If Regim = "ostat" Or Regim = "products" Then
    Me.Caption = "Ведомость остатков"
    If Regim = "products" Then
        Me.Caption = Me.Caption & " по готовым изделиям"
        Grid.colWidth(nkQuant) = 700
    End If
    If Regim = "ostat" Then
        Grid.colWidth(nkQuant) = 0
    End If
    
    cmExel.Visible = True
    cmHide.Visible = True
    Grid.colWidth(nkName) = 3510
    Grid.colWidth(nkCurOstat) = 810
    Grid.colWidth(nkDostup) = 800
    cmSel.Visible = False
    tbQuant.Visible = False
    laQuant.Visible = False
    laGrid2.Visible = False
    Grid2.Visible = False
    splRightV.Visible = Grid2.Visible
    laGrid.Visible = False
    Grid.Width = 7000 '6230
    Me.Width = Grid.Width + 2527
    cmExit.Left = Me.Width - cmExit.Width - 200
    Grid2.Width = 0 'для Resize
    GoTo EN1
ElseIf Regim = "" Or Regim = "closeZakaz" Then
    cmExel2.Visible = True
End If

gSeriaId = 0 'необходим  для добавления класса

Grid2.colWidth(0) = 0
Grid2.colWidth(dnNomNom) = 0 '900
'Grid2.ColWidth(dnNomName) =  в Resize
Grid2.colWidth(dnEdIzm) = 435
Grid2.colWidth(dnCenaEd) = 510
Grid2.colWidth(dnQuant) = 600
Grid2.colWidth(dnSumm) = 660
Grid2.colWidth(dnVesEd) = 495
Grid2.colWidth(dnVes) = 600

quantity2 = 0
loadPredmeti ' сюда попадаем только из предметов заказа
Dim Grid2Width As Long: Grid2Width = adjustGirdMoneyColWidth()
Grid2.Width = Grid2Width + 500
Me.Width = Grid.Left + Grid.Width + Grid2Width + 800

If quantity2 > 0 Then
    str = "Редактирование"
Else
    str = "Формирование"
End If
Me.Caption = str & " предметов к заказу № " & numDoc
If Regim = "closeZakaz" Then
    Me.Caption = "Предметы к заказу № " & numDoc
    laBegin.Caption = "Это закрытый заказ. Редактирование предметов невозможно."
    tv.Enabled = False
    cmSel.Enabled = False
    laGrid2.Enabled = False
    cmExel2.Visible = False
End If
EN1:
oldHeight = Me.Height
oldWidth = Me.Width
tvVes = tv.Width / (tv.Width + Grid.Width + Grid2.Width)
gridVes = Grid.Width / (tv.Width + Grid.Width + Grid2.Width)
grid2Ves = Grid2.Width / (tv.Width + Grid.Width + Grid2.Width)
isLoad = True
End Sub

Function loadPredmeti() As Double

Dim s As Double, sum As Double, sumVes As Double, quant As Double

MousePointer = flexHourglass
Grid2.Visible = False
splRightV.Visible = Grid2.Visible
quantity2 = 0
clearGrid Grid2
sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.ed_Izmer2, " & _
"sGuideNomenk.Size, sGuideNomenk.cod, sGuideNomenk.perList, sDMCrez.quantity, " & _
"sDMCrez.intQuant, sGuideNomenk.VES " & _
"FROM sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom " & _
"Where (((sDMCrez.numDoc) = " & numDoc & ")) ORDER BY sGuideNomenk.nomNom;"

'MsgBox sql

Set tbNomenk = myOpenRecordSet("##118", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  sum = 0: sumVes = 0
  While Not tbNomenk.EOF
    quantity2 = quantity2 + 1
    Grid2.TextMatrix(quantity2, dnNomName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid2.TextMatrix(quantity2, dnNomNom) = tbNomenk!nomNom
    Grid2.TextMatrix(quantity2, dnEdIzm) = tbNomenk!ed_Izmer2
    Grid2.TextMatrix(quantity2, dnVesEd) = tbNomenk!VES

    Grid2.TextMatrix(quantity2, dnCenaEd) = Round(rated(tbNomenk!intQuant, orderRate), 2)
    quant = Round(tbNomenk!quantity / tbNomenk!perList, 2)
    Grid2.TextMatrix(quantity2, dnQuant) = quant
    s = Round(tbNomenk!VES * quant, 3)
    Grid2.TextMatrix(quantity2, dnVes) = s
    sumVes = sumVes + s
    
    s = quant * tbNomenk!intQuant
    Grid2.TextMatrix(quantity2, dnSumm) = Round(rated(s, orderRate), 2)
    sum = sum + s
    Grid2.AddItem ""
    tbNomenk.MoveNext
  Wend
  'Grid2.RemoveItem quantity2 + 1
End If
tbNomenk.Close
EN1:
Grid2.Visible = True
splRightV.Visible = True

If quantity2 > 0 Then
    Grid2.TextMatrix(quantity2 + 1, dnQuant) = "Итого:"
    Grid2.row = quantity2 + 1: Grid2.col = dnSumm
    Grid2.Text = Round(rated(sum, orderRate), 2)
    Grid2.CellFontBold = True
    
    Grid2.col = dnVes
    Grid2.Text = Round(sumVes, 3)
    Grid2.CellFontBold = True
    
    Grid2.row = 1: Grid2.col = 1
    On Error Resume Next
    If Me.isLoad Then
        Grid2.SetFocus
    Else
        Grid2.TabIndex = 0
    End If
End If
loadPredmeti = sum

MousePointer = flexDefault


End Function


Private Sub Form_Resize()
Dim h As Integer, w As Integer, Left As Long

If Not isLoad Then Exit Sub
If Me.WindowState = vbMinimized Then Exit Sub
If Me.WindowState = vbMaximized And Me.Width > cDELLwidth Then 'экран DELL
    Grid2.colWidth(dnNomName) = 5220
    Grid.colWidth(nkName) = 5670
Else
    Grid2.colWidth(dnNomName) = 1230 '2340
    Grid.colWidth(nkName) = 2820
End If
On Error Resume Next
h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
tv.Height = tv.Height + h

tv.Width = tv.Width + w * tvVes
If Regim <> "products" Then
    Grid.Height = tv.Height
    splLeftH.Visible = False
End If

Grid.Left = Grid.Left + w * tvVes
laGrid1.Left = Grid.Left
laBegin.Left = tv.Left + tv.Width + 100
laBegin.Top = tv.Top

Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w * gridVes

Grid2.Left = Grid2.Left + w * (tvVes + gridVes)
laGrid2.Left = Grid2.Left
Grid2.Height = Grid2.Height + h
Grid2.Width = Grid2.Width + w * grid2Ves

splLeftV.Top = Grid.Top
splLeftV.Left = tv.Left + tv.Width + 15
splLeftV.Height = tv.Height

splLeftH.Top = Grid3.Top + Grid3.Height + 5
splLeftH.Left = splLeftV.Left + splLeftV.Width
splLeftH.Width = Grid3.Width

splRightV.Top = splLeftV.Top
splRightV.Left = splLeftH.Left + splLeftH.Width
splRightV.Height = splLeftV.Height

cmSel.Top = cmSel.Top + h
cmSel.Left = cmSel.Left + w
tbQuant.Top = tbQuant.Top + h
tbQuant.Left = tbQuant.Left + w
laQuant.Top = laQuant.Top + h
laQuant.Left = laQuant.Left + w
cmExit.Top = cmExit.Top + h
cmExit.Left = cmExit.Left + w
cmExel2.Top = cmExel2.Top + h
cmExel2.Left = cmExel2.Left + w
cmExel.Top = cmExel.Top + h
cmHide.Top = cmHide.Top + h
Grid3.Left = Grid.Left
Grid3.Width = Grid.Width
laGrid.Left = Grid3.Left

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
End If
End Sub



Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If quantity = 0 Then Exit Sub
'в Grid недопустима сортировка т.к. цветовая группировка
End Sub

Private Sub Grid_DblClick()
Dim il As Long, curRow As Long

grColor = Grid.CellBackColor
If grColor = &H88FF88 Then
    If MsgBox("Если Вы хотите просмотреть список всех заказов, под " & _
    "которые была зарезервирована эта номенклатура, нажмите <Да>.", _
    vbYesNo Or vbDefaultButton2, "Посмотреть, кто резервировал? '" & _
    gNomNom & "' ?") = vbYes Then
        Report.Regim = "whoRezerved"
        Report.Show vbModal
    End If
Else
    cmSel_Click
End If
End Sub

Private Sub Grid_EnterCell()
Dim f As String, d As String

If quantity = 0 Or Grid.col = buntColumn Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col


'If opProduct.value Then
'    gProductId = Grid.TextMatrix(mousRow, gpId)
'    gProduct = Grid.TextMatrix(mousRow, gpName)
'Else
gNomNom = Grid.TextMatrix(mousRow, nkNomer)

'End If
 
Grid.CellBackColor = vbYellow
If mousCol = nkDostup Then
    f = Grid.TextMatrix(mousRow, nkCurOstat)
    d = Grid.TextMatrix(mousRow, nkDostup)
    If d < f Then Grid.CellBackColor = &H88FF88
End If

End Sub

Private Sub Grid_GotFocus()
'If opProduct.value Then rightORleft "l"
'rightORleft "l"
cmHide.Enabled = True
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid_DblClick
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyEscape Then Grid.CellBackColor = Grid.BackColor
If KeyCode = vbKeyEscape Then Grid_EnterCell
End Sub

Private Sub Grid_LeaveCell()
    If Grid.col <> 0 And Grid.col <> buntColumn Then Grid.CellBackColor = Grid.BackColor
End Sub


Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.colWidth(Grid.MouseCol)
End Sub

Private Sub Grid2_Click()
mousCol2 = Grid2.MouseCol
mousRow2 = Grid2.MouseRow

If mousRow2 = 0 Then
    Grid2.CellBackColor = Grid.BackColor
'    SortCol Grid2, mousCol
    trigger = Not trigger
    Grid2.Sort = 9
    
    Grid2.row = 1    ' только чтобы снять выделение
Else
    If quantity2 = 0 Then Exit Sub
End If
'Grid_EnterCell

End Sub

Private Sub Grid2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Grid2.MouseRow = 0 Then
    If Shift = 2 Then MsgBox "ColWidth = " & Grid2.colWidth(Grid2.MouseCol)
ElseIf Button = 2 And quantity2 <> 0 Then
    Grid2.row = Grid2.MouseRow
    Grid2.col = dnNomNom
    gNomNom = Grid2.Text
    Grid2.SetFocus
    Grid2.CellBackColor = vbButtonFace
    Me.PopupMenu mnContext
'    noClick = True
End If
End Sub

Private Sub Grid2_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    If Row1 = Grid2.Rows - 1 Then
        Cmp = 1: Exit Sub
    End If
    If Row2 = Grid2.Rows - 1 Then
        Cmp = -1: Exit Sub
    End If
    If Grid2.TextMatrix(Row1, mousCol2) < Grid2.TextMatrix(Row2, mousCol2) Then
        Cmp = -1
    ElseIf Grid2.TextMatrix(Row1, mousCol2) > Grid2.TextMatrix(Row2, mousCol2) Then
        Cmp = 1
    Else
        Cmp = 0
    End If
    If (trigger) Then Cmp = -Cmp
End Sub

Private Sub Grid2_DblClick()
If mousRow2 = 0 Then Exit Sub
If Grid2.CellBackColor = &H88FF88 Then textBoxInGridCell tbMobile, Grid2

End Sub

Private Sub Grid2_EnterCell()

mousRow2 = Grid2.row
mousCol2 = Grid2.col

If mousCol2 = dnSumm Or mousCol2 = dnCenaEd Then
    Grid2.CellBackColor = &H88FF88
Else
    Grid2.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid2_DblClick

End Sub

Private Sub Grid2_LeaveCell()
If Grid2.col <> 0 Then Grid2.CellBackColor = Grid2.BackColor

End Sub



Private Sub Grid3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Grid3.MouseRow = 0 And Shift = 2 Then MsgBox "ColWidth = " & Grid3.colWidth(Grid3.MouseCol)
    
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
        MsgBox "ColWidth = " & Grid4.colWidth(Grid4.MouseCol)

End Sub


Private Sub mnDel_Click()
Dim pQuant As Double, I As Integer ', str  As String

If beNaklads() Then Exit Sub

If MsgBox("Вы хотите удалить позицию '" & gNomNom & _
"'", vbYesNo Or vbDefaultButton2, "Подтвердите удаление") = vbNo Then Exit Sub

sql = "DELETE From sDMCrez WHERE (((numDoc)=" & gNzak & _
") AND ((nomNom)='" & gNomNom & "'));"
'MsgBox sql

If myExecute("##348", sql) = 0 Then
    Orders.Grid.TextMatrix(Orders.Grid.row, orZakazano) = Round(rated(loadPredmeti, orderRate), 2)
End If
'Grid2.SetFocus ' здесь не срабатывает

End Sub

'
Sub controlEnable(EN As Boolean)
If Not EN Then ' только гасим
    Grid.Visible = False
    splLeftH.Visible = Grid.Visible
End If
cmSel.Enabled = EN
End Sub


Function nomenkToDMC(delta As Double, Optional noOpen As String = "") As Boolean
Dim s As Double

nomenkToDMC = False

If noOpen = "" Then
    
    If Not lockSklad Then Exit Function
    
'    s = nomencOstatkiToGrid(1) - delta ' одновременно обновляем таблицу
'AA: If s < -0.005 Then 'в 2х местах
'        If MsgBox("Дефицит товара '" & gNomNom & "' по подразделению '" & _
'        sDocs.getGridColSour() & "'составит (" & s & "), продолжить?", _
'        vbOKCancel Or vbDefaultButton2, "Подтвердите") = vbCancel Then GoTo EN1
'    End If
    
    Set tbDMC = myOpenRecordSet("##123", "sDMC", dbOpenTable)
    If tbDMC Is Nothing Then GoTo EN1
    tbDMC.index = "NomDoc"
End If

tbDMC.Seek "=", numDoc, numExt, gNomNom
If tbDMC.NoMatch Then
    tbDMC.AddNew
    tbDMC!numDoc = numDoc
    tbDMC!numExt = 254
    tbDMC!nomNom = gNomNom
    tbDMC!quant = Round(delta, 2)
Else
    tbDMC.Edit
    tbDMC!quant = Round(tbDMC!quant + delta, 2)
End If
tbDMC.Update
    
If noOpen = "" Then tbDMC.Close

'корректируем остатки(для межскладских не корректирует)
If Not ostatCorr(delta) Then MsgBox "Не прошла коррекция остатков. " & _
     "Сообщите администратору.", , "Error 83" '##83
nomenkToDMC = True

EN1:
If noOpen = "" Then lockSklad "un"
End Function

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
        
        If Grid.Width - DraggingShift > 100 Then
        Else
            Exit Sub
        End If
            Grid.Width = Grid.Width - DraggingShift
            tv.Width = tv.Width + DraggingShift
        splLeftV.Left = splLeftV.Left + DraggingShift
        Grid.Left = Grid.Left + DraggingShift
        laGrid.Left = Grid.Left
        laGrid.Width = Grid.Width
        Grid3.Left = Grid.Left
        Grid3.Width = Grid.Width
        laBegin.Left = Grid.Left
        If laBegin.Width > DraggingShift Then _
            laBegin.Width = Grid.Width
        splLeftH.Left = Grid.Left
        splLeftH.Width = Grid.Width
    End If
End Sub


Private Sub splLeftV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
        If Not (Grid2.Width - DraggingShift > 100) Then
            Exit Sub
        End If
            
        Grid3.Width = Grid3.Width + DraggingShift
        Grid.Width = Grid3.Width
        splRightV.Left = splRightV.Left + DraggingShift
        splLeftH.Width = Grid3.Width
        Grid2.Left = Grid2.Left + DraggingShift
        Grid2.Width = Grid2.Width - DraggingShift
        laGrid2.Left = Grid2.Left
        laGrid2.Width = Grid2.Width
    End If
End Sub

Private Sub splRightV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim c As Double, s As Double, str As String

If KeyCode = vbKeyReturn Then
    If Not isNumericTbox(tbMobile, 0) Then Exit Sub
    If mousCol2 = dnSumm Then
        s = tuneCurencyAndGranularity(tbMobile.Text, orderRate, sessionCurrency, Grid2.TextMatrix(mousRow2, dnQuant))
        c = s / CDbl(Grid2.TextMatrix(mousRow2, dnQuant)) 'не округлять
    Else 'dnCenaEd
        c = tuneCurencyAndGranularity(tbMobile.Text, orderRate, sessionCurrency, 1)
        s = c * CDbl(Grid2.TextMatrix(mousRow2, dnQuant))
    End If
    sql = "UPDATE sDMCrez SET intQuant = " & c & " WHERE (((numDoc)=" & _
    gNzak & ") AND ((nomNom)='" & Grid2.TextMatrix(mousRow2, dnNomNom) & "'));"
    
    If myExecute("##205", sql) = 0 Then
        Grid2.TextMatrix(mousRow2, dnCenaEd) = Round(rated(c, orderRate), 2)
        Grid2.TextMatrix(mousRow2, dnSumm) = Round(rated(s, orderRate), 2)
        s = getOrdered(gNzak)
        Orders.Grid.TextMatrix(Orders.Grid.row, orZakazano) = Round(rated(s, orderRate), 2)
        Orders.Grid.TextMatrix(Orders.Grid.row, orOtgrugeno) = getShipped(gNzak)
        Grid2.TextMatrix(Grid2.Rows - 1, dnSumm) = Round(rated(s, orderRate), 2)
    End If
    
    lbHide
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub

Sub lbHide()
tbMobile.Visible = False
Grid2.Enabled = True
Grid2.SetFocus
Grid2_EnterCell

End Sub

Function deficitAndNoIgnore(delta As Double) As Boolean
Dim s As Double, il As Long


deficitAndNoIgnore = False
s = nomencOstatkiToGrid(il) - delta ' одновременно обновляем таблицу
If s < -0.005 Then
    If MsgBox("Дефицит товара '" & gNomNom & "' в доступных остатках" & " составит (" & _
    s & "), продолжить?", vbOKCancel Or vbDefaultButton2, "Подтвердите") _
    = vbOK Then Exit Function
    deficitAndNoIgnore = True
End If
End Function

Private Sub tbQuant_KeyDown(KeyCode As Integer, Shift As Integer)
Dim s As Double ', str As String
'Dim i As Integer, NN2() As String
Dim per As Double ', delta As Double

If KeyCode = vbKeyReturn Then

  If Not isNumericTbox(tbQuant, 0.01) Then Exit Sub

  If Not lockSklad Then Exit Sub
  
  sql = "SELECT perList From sGuideNomenk WHERE (((nomNom)='" & gNomNom & "'));"
  If Not byErrSqlGetValues("W##346", sql, per) Then GoTo AA
    
  s = nomencOstatkiToGrid(1) - tbQuant.Text ' одновременно обновляем таблицу
  If s < -0.005 Then 'в 2х местах
    If MsgBox("Дефицит товара '" & gNomNom & "' в доступных остатках " & _
    "составит (" & s & "), продолжить?", vbOKCancel Or vbDefaultButton2, _
    "Подтвердите") = vbCancel Then GoTo AA
  End If
        
  per = per * tbQuant.Text
  If Not nomenkToDMCrez(per) Then
AA: lockSklad "un"
    Grid.SetFocus
    GoTo ESC
  End If
    
  lockSklad "un"

    Orders.Grid.TextMatrix(Orders.Grid.row, orZakazano) = Round(rated(loadPredmeti, orderRate), 2)
  
  Grid.SetFocus
  GoTo ES2

ElseIf KeyCode = vbKeyEscape Then
ESC: tbQuant.Text = ""
    Grid.SetFocus
ES2: gridFrame.Visible = False
    tbQuant.Enabled = False
    laQuant.Enabled = False
    cmSel.Enabled = True
End If

End Sub
'при delta < 0 - возм. удаление
Function nomenkToDMCrez(delta As Double, Optional mov As String = "") As Boolean
Dim s As Double, I As Integer

nomenkToDMCrez = False

' выписанные межскладские не резервируем
strWhere = " WHERE (((numDoc)=" & numDoc & ") AND ((nomNom)='" & gNomNom & "'));"

If mov = "" Then mov = "rez"

sql = "SELECT quantity FROM sDMC" & mov & strWhere
If Not byErrSqlGetValues("W##423", sql, s) Then Exit Function

If s = 0 Then ' такой ном-ры нет
  If delta >= 0.01 Then
    sql = "INSERT INTO sDMC" & mov & " ( numDoc, nomNom, quantity ) " & _
    "SELECT " & numDoc & ", '" & gNomNom & "', " & delta & ";"
    If myExecute("##195", sql) <> 0 Then Exit Function
  End If
Else
    delta = s + delta

    If Round(delta, 2) > 0 Then
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

Private Sub Timer1_Timer()
'biloG3Enter_Cell = False
Timer1.Enabled = False
End Sub

Private Sub tv_AfterLabelEdit(Cancel As Integer, NewString As String)
' If Not flseriaAdd Then
'ValueToTableField "##115", "'" & NewString & "'", "sProducts", "seriaName", "bySeriaId"
gSeriaId = Mid$(tv.SelectedItem.Key, 2)
ValueToTableField "##115", "'" & NewString & "'", "sGuideSeries", "seriaName", "bySeriaId"
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
    Grid3.ZOrder
Else '                обе присутствуют
    Dim gHeightMax As Double, g3HeightMax As Double
    Dim fullHeight As Double
    
    gHeightMax = Grid.Rows * (Grid.CellHeight + 13) + 95
    g3HeightMax = Grid3.Rows * (Grid3.CellHeight + 13) + 95
    fullHeight = gHeightMax + g3HeightMax + laGrid.Height
    
    If fullHeight > formHeight Then
        Dim normalGrid As Double, normalGird3 As Double
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
        Else
            Grid3.Height = g3HeightMax
            laGrid.Top = Grid3.Top + Grid3.Height
            Grid.Top = laGrid.Top + laGrid.Height
            Grid.Height = tv.Top + tv.Height - Grid.Top
        End If
        
        
    Else
        Grid3.Height = g3HeightMax
        laGrid.Top = Grid3.Top + g3HeightMax
        Grid.Height = gHeightMax
        Grid.Top = laGrid.Top + laGrid.Height
        Grid.Height = tv.Top + tv.Height - Grid.Top
    End If
    If mousRow3 < Grid3.Rows Then
        If Not Grid3.RowIsVisible(mousRow3) Then rowViem mousRow3, Grid3
    End If
    splLeftH.Top = Grid3.Top + Grid3.Height
    splLeftH.Width = Grid3.Width
    splLeftH.Left = Grid3.Left
    
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


Sub setNameColWidth()
    If Me.WindowState = vbMaximized And Me.Width > cDELLwidth Then 'экран DELL
        Grid3.colWidth(gpDescript) = 5055
        Grid.colWidth(nkName) = 4350 + 900
    ElseIf Regim = "ostat" Then
        Grid3.colWidth(gpDescript) = 3495
    Else
        Grid.colWidth(nkName) = 2100 + 900
        Grid3.colWidth(gpDescript) = 2200
    End If

End Sub


Sub loadKlass()
Dim Key As String, pKey As String, k() As String, pK()  As String
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

 ReDim k(0): ReDim pK(0): ReDim NN(0): iErr = 0
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
  For I = 1 To UBound(k())
    If k(I) <> "" Then
        On Error GoTo ERR2 ' назначить еще проход
        Set Node = tv.Nodes.Add(pK(I), tvwChild, k(I), NN(I))
        On Error GoTo 0
        k(I) = ""
        Node.Sorted = True
    End If
NXT:
  Next I
Wend
tv.Nodes.Item("k0").Expanded = True
Exit Sub
ERR1:
 iErr = iErr + 1: bilo = True
 ReDim Preserve k(iErr): ReDim Preserve pK(iErr): ReDim Preserve NN(iErr)
 k(iErr) = Key: pK(iErr) = pKey: NN(iErr) = tbKlass!klassName
 Resume Next

ERR2: bilo = True: Resume NXT

End Sub

Sub loadSeria()
Dim Key As String, pKey As String, k() As String, pK()  As String
Dim I As Integer, iErr As Integer
bilo = False
sql = "SELECT sGuideSeries.*  From sGuideSeries ORDER BY sGuideSeries.seriaId;"
Set tbSeries = myOpenRecordSet("##110", sql, dbOpenForwardOnly)
If tbSeries Is Nothing Then myBase.Close: End
If Not tbSeries.BOF Then
 tv.Nodes.Clear
 Set Node = tv.Nodes.Add(, , "k0", "Справочник по сериям")
 Node.Sorted = True
 
 ReDim k(0): ReDim pK(0): ReDim NN(0): iErr = 0
 While Not tbSeries.EOF
    If tbSeries!seriaId = 0 Then GoTo NXT1
    Key = "k" & tbSeries!seriaId
    pKey = "k" & tbSeries!parentSeriaId
    On Error GoTo ERR1 ' назначить второй проход
    Set Node = tv.Nodes.Add(pKey, tvwChild, Key, tbSeries!seriaName)
    On Error GoTo 0
    Node.Sorted = True
NXT1:
    tbSeries.MoveNext
 Wend
End If
tbSeries.Close

While bilo ' необходимы еще проходы
  bilo = False
  For I = 1 To UBound(k())
    If k(I) <> "" Then
        On Error GoTo ERR2 ' назначить еще проход
        Set Node = tv.Nodes.Add(pK(I), tvwChild, k(I), NN(I))
        On Error GoTo 0
        k(I) = ""
        Node.Sorted = True
    End If
NXT:
  Next I
Wend
tv.Nodes.Item("k0").Expanded = True
Exit Sub
ERR1:
 iErr = iErr + 1: bilo = True
 ReDim Preserve k(iErr): ReDim Preserve pK(iErr): ReDim Preserve NN(iErr)
 k(iErr) = Key: pK(iErr) = pKey: NN(iErr) = tbSeries!seriaName
 Resume Next

ERR2: bilo = True: Resume NXT

End Sub


Sub loadProductNomenk(ByVal v_productId As Integer)
Dim s As Double, grBef As String, Left As Integer

Dragging = False

Me.MousePointer = flexHourglass
buntColumn = nkQuant

quantity = 0
Grid.Visible = False
clearGrid Grid

sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xGroup, " & _
"sGuideNomenk.Size, sGuideNomenk.cod, " & _
" sGuideNomenk.nomName, sGuideNomenk.ed_Izmer  " & _
"FROM sGuideNomenk INNER JOIN sProducts ON sGuideNomenk.nomNom = sProducts.nomNom " & _
"WHERE (((sProducts.ProductId)=" & v_productId & ")) ORDER BY sProducts.xGroup DESC;"
'MsgBox sql
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
        
    gNomNom = tbNomenk!nomNom
    Grid.TextMatrix(quantity, nkNomer) = gNomNom
    Grid.TextMatrix(quantity, nkName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid.TextMatrix(quantity, nkEdIzm) = tbNomenk!ed_Izmer
'    ReDim Preserve QP(quantity): QP(quantity) = tbNomenk!quantity
        Grid.TextMatrix(quantity, nkQuant) = tbNomenk!quantity
    'доступные остатки:
    Grid.TextMatrix(quantity, nkDostup) = Round(nomencOstatkiToGrid(-1), 2)
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
splLeftH.Visible = Grid.Visible
Grid.ZOrder
Me.MousePointer = flexDefault

End Sub

Sub loadKlassNomenk()
Dim il As Long, r As Double, strWhere As String
Dim beg As Double, prih As Double, rash As Double, oldNow As Double

Dragging = False


If tv.SelectedItem.Key = "all" Then
    strWhere = ""
    quantity = 0
Else
  strWhere = "WHERE (((klassId)=" & gKlassId & "))"
End If

laGrid1.Caption = "Номенклатура из группы '" & tv.SelectedItem.Text & "'"

Me.MousePointer = flexHourglass

If beShift Then
    Grid.AddItem ""
Else
    quantity = 0
    Grid.Visible = False
    splLeftH.Visible = Grid.Visible
    clearGrid Grid
End If


sql = "SELECT nomNom, nomName, Size, cod, ed_Izmer, ed_Izmer2, nowOstatki " & _
"From sGuideNomenk " & strWhere & " order by nomnom"
'"WHERE (((sGuideNomenk.klassId)=" & gKlassId & "));"

Set tbNomenk = myOpenRecordSet("##103", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
 tbNomenk.MoveFirst
 While Not tbNomenk.EOF
'    beg = tbNomenk!begOstatki
    gNomNom = tbNomenk!nomNom
    quantity = quantity + 1
    Grid.TextMatrix(quantity, nkNomer) = gNomNom
    Grid.TextMatrix(quantity, nkName) = tbNomenk!cod & " " & _
        tbNomenk!nomName & " " & tbNomenk!Size
    Grid.TextMatrix(quantity, nkEdIzm) = tbNomenk!ed_Izmer2
    
    If Regim = "ostat" Or Regim = "" Then
        'доступные остатки без дробной части (nomencOstatkiToGrid выдает perList в tmpSng):
        Grid.TextMatrix(quantity, nkDostup) = Round(nomencOstatkiToGrid(-1) - 0.4999, 0)
        Grid.TextMatrix(quantity, nkCurOstat) = Round(FO / tmpSng, 2) 'FO из nomencOstatkiToGrid
    End If
    Grid.AddItem ""
    tbNomenk.MoveNext
 Wend
 If quantity > 0 Then Grid.RemoveItem quantity + 1
End If
tbNomenk.Close

Grid.Visible = True
'splLeftH.Visible = Grid.Visible
EN1:
Me.MousePointer = flexDefault
End Sub

    
Private Sub tv_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer, str As String
If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
'    tv_NodeClick tv.SelectedItem
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
controlEnable True
gKlassId = Mid$(tv.SelectedItem.Key, 2)
If Regim = "products" Then
    gSeriaId = Mid$(tv.SelectedItem.Key, 2)
    loadSeriaProduct
    Grid3.Visible = True
    Grid.Visible = False
    splLeftH.Visible = Grid.Visible
    'laGrid.Visible = True
    cmSel.Enabled = False
    gridOrGrid3Hide "grid"
    Grid.colWidth(nkCurOstat) = 0
Else
    loadKlassNomenk
End If
'    Grid.Visible = True
Grid_EnterCell
On Error Resume Next
Grid.SetFocus

'biloG3Enter_Cell = False

End Sub

Sub loadSeriaProduct()
Dim il As Long, strWhere As String

If tv.SelectedItem.Key = "k0" Then
    gSeriaId = 0
    Grid3.Visible = False
    Exit Sub
End If

Me.MousePointer = flexHourglass
laGrid1.Caption = "Список готовых изделий по серии '" & tv.SelectedItem.Text & "'"

quantity3 = 0
Grid3.Visible = False
clearGrid Grid3
il = 0

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
    
    'Grid3.TextMatrix(quantity3, gpNN) = quantity3
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
EN1:
Me.MousePointer = flexDefault

End Sub

