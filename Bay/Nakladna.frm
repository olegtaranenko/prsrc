VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Nakladna 
   BackColor       =   &H8000000A&
   Caption         =   "Предметы "
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmClose 
      Caption         =   "Списать"
      Height          =   315
      Left            =   2820
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmSostav 
      Caption         =   "Состав изд."
      Height          =   315
      Left            =   4380
      TabIndex        =   14
      Top             =   5160
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox tbMobile2 
      Height          =   315
      Left            =   780
      TabIndex        =   12
      Text            =   "tbMobile2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7329
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
   Begin VB.Label Label1 
      Caption         =   "Заказ №"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   60
      Width           =   735
   End
   Begin VB.Label laControl 
      Caption         =   "Контроль: "
      Height          =   195
      Left            =   6180
      TabIndex        =   18
      Top             =   5220
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label laControl2 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6975
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label laOt 
      Caption         =   "От:"
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   480
      Width           =   255
   End
   Begin VB.Label laDate 
      Height          =   195
      Left            =   7020
      TabIndex        =   13
      Top             =   0
      Width           =   1155
   End
   Begin VB.Label laSignatura 
      BackColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   6780
      TabIndex        =   11
      Top             =   300
      Width           =   1395
   End
   Begin VB.Label laPerson 
      Caption         =   "Исполнитель:"
      Height          =   195
      Left            =   5700
      TabIndex        =   10
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label laFirm 
      Caption         =   "laFirm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3600
      TabIndex        =   9
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label laPlatel 
      Caption         =   "Плательщик:"
      Height          =   195
      Left            =   2520
      TabIndex        =   8
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label laDest 
      Caption         =   "Сборка"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label laKomu 
      Caption         =   "Кому:"
      Height          =   195
      Left            =   3060
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label laSours 
      Caption         =   "laSours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   540
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label laDocNum 
      Caption         =   "laDocNum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
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
Dim secondNaklad As String, beSUO As Boolean ' была листовая ном-ра

Const nkNomNom = 1
Const nkNomName = 2
Const nkEdIzm = 3
Const nkTreb = 4
Const nkClos = 5
'Const nkEtap = 6
'Const nkEClos = 7
Const nkQuant = 6
'Const nkIntEdIzm = 9
'Const nkIntQuant = 10


Private Sub cmExel_Click()
Dim str As String
str = laDocNum.Caption
GridToExcel Grid2, "Накладная № " & str
End Sub

Private Sub cmExit_Click()

Unload Me

End Sub

Private Sub cmPrint_Click()
laDate.Caption = Format(Now(), "dd.mm.yy hh:nn")
Me.PrintForm
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

Private Sub Form_Load()
Dim str As String
Dim i As Integer, delta As Integer 'notBay As Long,

oldHeight = Me.Height
oldWidth = Me.Width
'Const nkNomNom = 1
'Const nkNomName = 2
'Const nkEdIzm = 3
'Const nkTreb = 4
'Const nkClos = 5
'Const nkQuant = 8

Grid2.MergeRow(0) = True
Grid2.FormatString = "|<Номер|<Описание|<Ед.измерения|Затребовано по заказу|Отпущено|кол-во"
Grid2.ColWidth(0) = 0
Grid2.ColWidth(nkNomNom) = 945
Grid2.ColWidth(nkNomName) = 4500 '5265
Grid2.ColWidth(nkEdIzm) = 645
Grid2.ColWidth(nkQuant) = 735
'размеры некот. колонок обнуляются также и по результататм загрузки (см. loadToGrid)
Grid2.ColWidth(nkTreb) = 630
Grid2.ColWidth(nkClos) = 855
'Grid2.ColWidth(nkEtap) = 780
'Grid2.ColWidth(nkEClos) = 765
'Grid2.ColWidth(nkIntEdIzm) = 700
'Grid2.ColWidth(nkIntQuant) = 700

cmExit.Caption = "Выход"
secondNaklad = ""
Me.Caption = "Предметы к заказу."
cmSostav.Visible = True
laDocNum.Caption = gNzak

MousePointer = flexHourglass

laPlatel.Visible = False
laFirm.Visible = False
'If numExt <> 254 Then  'к заказу
'    sql = "SELECT Orders.numOrder, GuideFirms.Name " & _
'    "FROM GuideFirms INNER JOIN Orders ON GuideFirms.FirmId = Orders.FirmId " & _
'    "WHERE (((Orders.numOrder)=" & numDoc & "));"
'    notBay = 0
'    byErrSqlGetValues "W##170", sql, notBay, str
'    If notBay > 0 Then GoTo AA ' заказ не на продажу
    
    sql = "SELECT BayGuideFirms.Name " & _
    "FROM BayGuideFirms INNER JOIN BayOrders ON BayGuideFirms.FirmId = " & _
    "BayOrders.FirmId WHERE (((BayOrders.numOrder)=" & gNzak & "));"
    If byErrSqlGetValues("W##426", sql, str) Then
AA:     laPlatel.Visible = True
        laFirm.Visible = True
        laFirm.Caption = str
    End If
'End If

loadToGrid
i = 350 + (Grid2.CellHeight + 17) * quantity2
delta = i - Grid2.Height
Me.Height = Me.Height + delta

MousePointer = flexDefault
End Sub
'ind=1 м.б. только при Regim = ""
Sub loadToGrid()
Dim i As Integer, s As Double, s2 As Double, str As String, str2 As String


'ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): QQ2(0) = 0: ReDim QQ3(0)

laSours.Caption = "Склад1"
'If Not sProducts.zakazNomenkToNNQQ Then GoTo EN1

Grid2.Visible = False
quantity2 = 0
clearGrid Grid2
beSUO = False
'sDMCrez.curQuant, sDMCrez.intQuant,
sql = "SELECT sDMCrez.quantity, sDMCrez.curQuant, sGuideNomenk.nomNom, " & _
"sGuideNomenk.nomName, sGuideNomenk.perList, sGuideNomenk.Size, " & _
"sGuideNomenk.ed_Izmer2, sGuideNomenk.cod, sGuideNomenk.obrez " & _
"FROM sGuideNomenk INNER JOIN sDMCrez ON sGuideNomenk.nomNom = sDMCrez.nomNom " & _
"WHERE (((sDMCrez.numDoc)=" & gNzak & "));"
'MsgBox sql
Set tbNomenk = myOpenRecordSet("##129", sql, dbOpenForwardOnly)
While Not tbNomenk.EOF
        quantity2 = quantity2 + 1
        If tbNomenk!perList > 1 Then Grid2.TextMatrix(quantity2, 0) = "Да" 'обрезная
        Grid2.TextMatrix(quantity2, nkNomNom) = tbNomenk!nomNom
        Grid2.TextMatrix(quantity2, nkNomName) = tbNomenk!cod & " " & _
            tbNomenk!nomName & " " & tbNomenk!Size
        Grid2.TextMatrix(quantity2, nkEdIzm) = tbNomenk!ed_Izmer2
        Grid2.TextMatrix(quantity2, nkTreb) = Round(tbNomenk!quantity / tbNomenk!perList, 2)
'        Grid2.TextMatrix(quantity2, nkEtap) = Round(QQ2(i) - QQ3(i), 2)
            
        sql = "SELECT Sum(quant) AS Sum_quant From sDMC WHERE " & _
        "(((sDMC.numDoc)=" & gNzak & ") AND ((sDMC.nomNom)='" & tbNomenk!nomNom & "'));"
'Debug.Print sql
        If byErrSqlGetValues("##194", sql, s) Then
                Grid2.TextMatrix(quantity2, nkClos) = Round(s / tbNomenk!perList, 2)
'                Grid2.TextMatrix(quantity2, nkEClos) = Round(s - QQ3(i), 2)
        End If

        Grid2.TextMatrix(quantity2, nkQuant) = Round(tbNomenk!curQuant)

'        If tbNomenk!perList <> 1 Then 'для обрезной доп. колонка для целых
'            beSUO = True
'            Grid2.TextMatrix(quantity2, nkIntEdIzm) = tbNomenk!ed_Izmer2
'        End If
'            s = 0: s2 = 0
'              sql = "SELECT curQuant, intQuant from sDMCrez " & _
'              "WHERE (((numDoc)=" & gNzak & ") AND ((nomNom)='" & NN(i) & "'));"
'              byErrSqlGetValues "##362", sql, s, s2
'              If s > 0 Then _
'                Grid2.TextMatrix(quantity2, nkQuant) = s
'              If s2 > 0 Then _
'                Grid2.TextMatrix(quantity2, nkIntQuant) = s2
        Grid2.AddItem ""
        tbNomenk.MoveNext
Wend
tbNomenk.Close
If quantity2 > 0 Then
    Grid2.RemoveItem quantity2 + 1
End If
'If ind = 0 Then
'  If QQ2(0) = 0 Then  'если не этапный убираем колонки
'    Grid2.ColWidth(nkEtap) = 0'
'    Grid2.ColWidth(nkEClos) = 0
'  ElseIf dostup <> "" Then ' для менеджеров оставляем в любом случае
'    Grid2.ColWidth(nkTreb) = 0
'    Grid2.ColWidth(nkClos) = 0
'  End If
'  If Not beSUO Then
'    Grid2.ColWidth(nkIntEdIzm) = 0
'    Grid2.ColWidth(nkIntQuant) = 0
'  End If
'End If
Dim sum  As Long
sum = 0
For i = 0 To Grid2.Cols - 1
    sum = sum + Grid2.ColWidth(i)
Next i
sum = sum + 680 '650
If sum < 8300 Then sum = 8300
Me.Width = sum
Grid2.col = nkQuant
EN1:
Grid2.Visible = True
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
Grid2.Height = Grid2.Height + h
Grid2.Width = Grid2.Width + w

cmPrint.Top = cmPrint.Top + h
cmExel.Top = cmExel.Top + h
cmSostav.Top = cmSostav.Top + h
cmClose.Top = cmClose.Top + h
cmClose.Left = cmClose.Left + w
laControl.Top = laControl.Top + h
laControl2.Top = laControl2.Top + h
laControl.Left = laControl.Left + w
laControl2.Left = laControl2.Left + w
cmExit.Top = cmExit.Top + h
cmExit.Left = cmExit.Left + w
End Sub

Private Sub Form_Unload(Cancel As Integer)
Regim = "" 'нужно для lbInside_LostFocus
End Sub

Private Sub Grid2_DblClick()
Dim str As String, per As Double, ed_Izmer As String

If Not Grid2.CellBackColor = &H88FF88 Then Exit Sub
  
gNomNom = Grid2.TextMatrix(mousRow2, nkNomNom)
textBoxInGridCell tbMobile2, Grid2
End Sub

Private Sub Grid2_EnterCell()
mousRow2 = Grid2.row
mousCol2 = Grid2.col

If mousCol2 = nkQuant Then
    Grid2.CellBackColor = &H88FF88
Else
    Grid2.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Grid2_DblClick
End Sub

Private Sub Grid2_LeaveCell()
Grid2.CellBackColor = Grid2.BackColor

End Sub

Private Sub Grid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid2.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid2.ColWidth(Grid2.MouseCol)

End Sub

Sub lbHide2()
tbMobile2.Visible = False
'gridFrame.Visible = False
Grid2.Enabled = True
Grid2.SetFocus
Grid2_EnterCell
End Sub

Private Sub tbMobile2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim delta As Double, quant As Double, s As Double, str As String

If KeyCode = vbKeyReturn Then
  
    quant = Grid2.TextMatrix(mousRow2, nkTreb)
    quant = Round(quant - Grid2.TextMatrix(mousRow2, nkClos), 2)
    
    If Not isNumericTbox(tbMobile2, 0, quant) Then Exit Sub
    
    quant = Round(tbMobile2.Text, 2)
    str = "cur"

sql = "UPDATE sDMCrez SET curQuant = " & quant & _
" WHERE (((numDoc)=" & gNzak & ") AND ((nomNom)='" & _
Grid2.TextMatrix(mousRow2, nkNomNom) & "'));"
'MsgBox sql
If myExecute("##363", sql) = 0 Then
    If quant = 0 Then
        Grid2.TextMatrix(mousRow2, mousCol2) = ""
    Else
        Grid2.TextMatrix(mousRow2, mousCol2) = quant
    End If
End If
lbHide2
Grid2.SetFocus

ElseIf KeyCode = vbKeyEscape Then
NN:  lbHide2
End If

End Sub


