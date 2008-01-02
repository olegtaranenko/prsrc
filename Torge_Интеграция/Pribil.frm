VERSION 5.00
Begin VB.Form Pribil 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Реализация"
   ClientHeight    =   4128
   ClientLeft      =   552
   ClientTop       =   9336
   ClientWidth     =   12456
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4128
   ScaleWidth      =   12456
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      Caption         =   "ПМ"
      Height          =   2415
      Left            =   7800
      TabIndex        =   36
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton cmDetailPM 
         Caption         =   "Детализация"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   38
         Top             =   1620
         Width           =   1215
      End
      Begin VB.CommandButton cmStatPM 
         Caption         =   "Статистика"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   37
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label laMaterialsPM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   372
         Left            =   50
         TabIndex        =   51
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label laRealizPM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   372
         Left            =   50
         TabIndex        =   40
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label laClearPM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   39
         Top             =   1140
         Width           =   1200
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "ММ"
      Height          =   2415
      Left            =   9240
      TabIndex        =   46
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton cmStatMM 
         Caption         =   "Статистика"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   34
         Top             =   1980
         Width           =   1215
      End
      Begin VB.CommandButton cmDetailMM 
         Caption         =   "Детализация"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   47
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label laClearMM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   50
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label laMaterialsMM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   49
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label laRealizMM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   48
         Top             =   300
         Width           =   1200
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Аналит"
      Height          =   2415
      Left            =   10680
      TabIndex        =   41
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton cmDetailAN 
         Caption         =   "Детализация"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   42
         Top             =   1620
         Width           =   1215
      End
      Begin VB.CommandButton cmStatAN 
         Caption         =   "Статистика"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   35
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label laRealizAn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   45
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label laClearAn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   44
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label laMaterialsAn 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   43
         Top             =   720
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmStat1 
      Caption         =   "Статистика"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3600
      TabIndex        =   33
      Top             =   3540
      Width           =   1215
   End
   Begin VB.CommandButton cmStat2 
      Caption         =   "Статистика"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5040
      TabIndex        =   32
      Top             =   3540
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Caption         =   " Услуги"
      Height          =   2415
      Left            =   2100
      TabIndex        =   27
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton cmStat3 
         Caption         =   "Статистика"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   31
         Top             =   1980
         Width           =   1215
      End
      Begin VB.CommandButton cmDetail3 
         Caption         =   "Детализация"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   28
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label laRealiz1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   30
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label laUslug 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   29
         Top             =   300
         Width           =   1200
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Итого"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6420
      TabIndex        =   23
      Top             =   1560
      Width           =   1275
      Begin VB.Label laRealiz 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   50
         TabIndex        =   26
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label laClear 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   50
         TabIndex        =   25
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label laMaterials 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   50
         TabIndex        =   24
         Top             =   720
         Width           =   1200
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Продажа"
      Height          =   2415
      Left            =   4980
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton cmDetail2 
         Caption         =   "Детализация"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   19
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label laMaterials2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   22
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label laClear2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   21
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label laRealiz2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   20
         Top             =   300
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Материалы не под заказы:"
      Height          =   795
      Left            =   3540
      TabIndex        =   15
      Top             =   660
      Width           =   2715
      Begin VB.CommandButton cmDetail 
         Caption         =   "Детализация"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label laOther 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Товары"
      Height          =   2415
      Left            =   3540
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
      Begin VB.CommandButton cmDetail1 
         Caption         =   "Детализация"
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label laProduct 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   14
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label laMaterials1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   13
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label laClear1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   50
         TabIndex        =   12
         Top             =   1140
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   552
      Left            =   -60
      TabIndex        =   3
      Top             =   0
      Width           =   8715
      Begin VB.TextBox tbStartDate 
         Height          =   285
         Left            =   960
         MaxLength       =   8
         TabIndex        =   7
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox tbEndDate 
         Height          =   285
         Left            =   1980
         MaxLength       =   8
         TabIndex        =   6
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton cmManag 
         Caption         =   "Применить"
         Height          =   315
         Left            =   2880
         TabIndex        =   5
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmSave 
         Caption         =   "Записать в журнал Х.О. "
         Enabled         =   0   'False
         Height          =   315
         Left            =   6300
         TabIndex        =   4
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label laPeriod 
         Caption         =   "Период  с "
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   795
      End
      Begin VB.Label laPo 
         Caption         =   "по"
         Height          =   195
         Left            =   1785
         TabIndex        =   8
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Реализац. - материалы:"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1932
   End
   Begin VB.Label laHMaterials 
      Caption         =   "Материалы под заказы:"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   2340
      Width           =   1992
   End
   Begin VB.Label laHRealiz 
      Caption         =   "Реализация:"
      Height          =   192
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   972
   End
End
Attribute VB_Name = "Pribil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pDateWhere As String, nDateWhere As String, uDateWhere As String
Public bDateWhere As String, mDateWhere As String
Public statistic As String, ventureId As String

Dim begDateHron As Date ' Начало ведения хронологии

Private Sub cmDetail_Click()
Me.MousePointer = flexHourglass
Report.param1 = laOther.Caption
Report.Regim = "mat"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmDetail1_Click()
Me.MousePointer = flexHourglass
'Report.param1 = laRealiz1.Caption
Report.param1 = laProduct.Caption
Report.param2 = laMaterials1.Caption
Report.Regim = ""
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmDetail2_Click()
Me.MousePointer = flexHourglass
Report.param1 = laRealiz2.Caption
Report.param2 = laMaterials2.Caption
Report.Regim = "bay"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmDetail3_Click()
Me.MousePointer = flexHourglass
'Report.param1 = laRealiz1.Caption
Report.param1 = laUslug.Caption
'Report.param2 = laMaterials1.Caption
Report.Regim = "uslug"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmDetailAN_Click()
Me.MousePointer = flexHourglass
statistic = False
ventureId = 3
Report.Regim = "venture"
Report.Show vbModal
Me.MousePointer = flexDefault
End Sub

Private Sub cmDetailMM_Click()
Me.MousePointer = flexHourglass
statistic = False
ventureId = 2
Report.Regim = "venture"
Report.Show vbModal
Me.MousePointer = flexDefault
End Sub

Private Sub cmDetailPM_Click()
Me.MousePointer = flexHourglass
statistic = False
ventureId = 1
Report.Regim = "venture"
Report.Show vbModal
Me.MousePointer = flexDefault
End Sub

Private Sub cmManag_Click() 'кнопка "применить" из отчета "Реализация"
Dim oborot As Single, dohod As Single, s2 As Single, s As Single
Dim ventureMat() As Single, ventureRealiz() As Single
Dim mat As Single, realiz As Single



Me.MousePointer = flexHourglass

ReDim ventureMat(2)
ReDim ventureRealiz(2)


strWhere = getWhereByDateBoxes(Me, "outDate", begDateHron) ' между
If strWhere = "error" Then GoTo EN1
If strWhere <> "" Then strWhere = " WHERE " & strWhere
pDateWhere = strWhere

'хронология по изделиям
sql = "SELECT r.*, p.cenaEd" _
    & ", isnull(o.ventureid, 1) as ventureid" _
    & " FROM xPredmetyByIzdeliaOut r " _
    & " JOIN xPredmetyByIzdelia p ON (p.prExt = r.prExt) AND (p.prId = r.prId) AND (p.numOrder = r.numOrder)" _
    & " join orders o on r.numorder = o.numorder and p.numorder = o.numorder" _
    & strWhere


' Debug.Print sql

Set tbProduct = myOpenRecordSet("##306", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then GoTo EN1
oborot = 0: dohod = 0:

If Not tbProduct.BOF Then
 While Not tbProduct.EOF
    gNzak = tbProduct!numOrder
    gProductId = tbProduct!prId
    prExt = tbProduct!prExt
    mat = getProductNomenkSum * tbProduct!quant
    realiz = tbProduct!cenaEd * tbProduct!quant
    
    oborot = oborot + mat
    dohod = dohod + realiz
    ventureMat(tbProduct!ventureId - 1) = ventureMat(tbProduct!ventureId - 1) + mat
    ventureRealiz(tbProduct!ventureId - 1) = ventureRealiz(tbProduct!ventureId - 1) + realiz
    
    tbProduct.MoveNext
 Wend
End If
tbProduct.Close

'хронология по отд.номеклатурам
strWhere = getWhereByDateBoxes(Me, "outDate", begDateHron) ' между
nDateWhere = strWhere
If strWhere <> "" Then strWhere = " WHERE " & strWhere

sql = "SELECT r.quant, n.cost, n.perList, p.cenaEd " _
    & " , isnull(o.ventureid, 1) as ventureid" _
    & " FROM xPredmetyByNomenkOut r " _
    & " JOIN xPredmetyByNomenk p ON p.nomNom = r.nomNom AND p.numOrder = r.numOrder" _
    & " join orders o on r.numorder = o.numorder" _
    & " JOIN sGuideNomenk n ON n.nomNom = r.nomNom " _
    & strWhere

'Debug.Print sql

Set tbNomenk = myOpenRecordSet("##307", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    mat = tbNomenk!cost * tbNomenk!quant / tbNomenk!perList
    realiz = tbNomenk!cenaEd * tbNomenk!quant
    
    oborot = oborot + mat
    dohod = dohod + realiz
    ventureMat(tbNomenk!ventureId - 1) = ventureMat(tbNomenk!ventureId - 1) + mat
    ventureRealiz(tbNomenk!ventureId - 1) = ventureRealiz(tbNomenk!ventureId - 1) + realiz
    tbNomenk.MoveNext
  Wend
End If
tbNomenk.Close
    
laProduct.Caption = Format(Round(dohod, 2), "## ##0.00")
laMaterials1.Caption = Format(Round(oborot, 2), "## ##0.00")
laClear1.Caption = Format(Round(dohod - oborot, 2), "## ##0.00")


' ------------------ услуги -------------------
strWhere = getWhereByDateBoxes(Me, "u.outDate", begDateHron) ' между
If strWhere <> "" Then strWhere = " WHERE " & strWhere

uDateWhere = strWhere
sql = "SELECT Sum(u.quant) AS Sum_quant " _
    & ", isnull(o.ventureid, 1) as ventureid" _
    & " from xUslugOut u" _
    & " join orders o on u.numorder = o.numorder" _
    & strWhere _
    & " group by isnull(o.ventureid, 1)"
    
    
Debug.Print sql

s2 = 0
Set tbNomenk = myOpenRecordSet("##380", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    realiz = tbNomenk!sum_quant
    s2 = s2 + realiz
    ventureRealiz(tbNomenk!ventureId - 1) = ventureRealiz(tbNomenk!ventureId - 1) + realiz
    dohod = dohod + realiz
    tbNomenk.MoveNext
  Wend
End If
tbNomenk.Close

laUslug.Caption = Format(Round(s2, 2), "## ##0.00")
laRealiz1.Caption = laUslug.Caption



' ------------------ Отгружено по  заказам продаж -------------------
strWhere = getWhereByDateBoxes(Me, "xDate", begDateHron) ' между
bDateWhere = strWhere

If strWhere <> "" Then strWhere = " WHERE " & strWhere

sql = "SELECT Sum(sDMC.quant*sDMCrez.intQuant/n.perList) AS bSum " _
    & " , isnull(BayOrders .ventureid, 1) as ventureid " _
    & " FROM sGuideNomenk n " _
    & " INNER JOIN ((BayOrders " _
    & " INNER JOIN sDocs ON BayOrders.numOrder = sDocs.numDoc) " _
    & " INNER JOIN (sDMC INNER JOIN sDMCrez ON sDMC.nomNom = sDMCrez.nomNom) ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc) AND (BayOrders.numOrder = sDMCrez.numDoc)) ON n.nomNom = sDMC.nomNom " _
    & strWhere _
    & " group by isnull(BayOrders.ventureid, 1)"


Debug.Print sql

s = 0
Set tbNomenk = myOpenRecordSet("##431", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    realiz = tbNomenk!bSum
    s = s + realiz
    ventureRealiz(tbNomenk!ventureId - 1) = ventureRealiz(tbNomenk!ventureId - 1) + realiz
    dohod = dohod + realiz
    tbNomenk.MoveNext
  Wend
End If
tbNomenk.Close

laRealiz2.Caption = Format(s, "## ##0.00")

sql = "SELECT Sum(n.cost*sDMC.quant/n.perList) AS cena " _
    & " , isnull(BayOrders .ventureid, 1) as ventureid " _
    & " FROM sGuideNomenk n INNER JOIN ((sDocs INNER JOIN " & _
    "BayOrders ON sDocs.numDoc = BayOrders.numOrder) INNER JOIN sDMC ON " & _
    "(sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc)) ON " & _
    "n.nomNom = sDMC.nomNom" & strWhere _
    & " group by isnull(BayOrders.ventureid, 1)"

s2 = 0
Set tbNomenk = myOpenRecordSet("##431", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    mat = tbNomenk!Cena
    s2 = s2 + mat
    ventureMat(tbNomenk!ventureId - 1) = ventureMat(tbNomenk!ventureId - 1) + mat
    oborot = oborot + mat
    tbNomenk.MoveNext
  Wend
End If
tbNomenk.Close

'MsgBox sql
'If byErrSqlGetValues("##430", sql, s2) Then
'    s2 = Round(s2, 2)
'    laMaterials2.Caption = Format(s2, "## ##0.00")
'    oborot = Round(oborot + s2, 2)
'End If

laMaterials2.Caption = Format(Round(s2, 2), "## ##0.00")
laClear2.Caption = Format(Round(s, 2) - Round(s2, 2), "## ##0.00")

laRealiz.Caption = Format(Round(dohod, 2), "## ##0.00")
laMaterials.Caption = Format(Round(oborot, 2), "## ##0.00")
laClear.Caption = Format(Round(dohod - oborot, 2), "## ##0.00")

laMaterialsPM.Caption = Format(Round(ventureMat(0), 2), "## ##0.00")
laMaterialsMM.Caption = Format(Round(ventureMat(1), 2), "## ##0.00")
laMaterialsAn.Caption = Format(Round(ventureMat(2), 2), "## ##0.00")

laRealizPM.Caption = Format(Round(ventureRealiz(0), 2), "## ##0.00")
laRealizMM.Caption = Format(Round(ventureRealiz(1), 2), "## ##0.00")
laRealizAn.Caption = Format(Round(ventureRealiz(2), 2), "## ##0.00")

laClearPM.Caption = Format(Round(ventureRealiz(0), 2) - Round(ventureMat(0), 2), "## ##0.00")
laClearMM.Caption = Format(Round(ventureRealiz(1), 2) - Round(ventureMat(1), 2), "## ##0.00")
laClearAn.Caption = Format(Round(ventureRealiz(2), 2) - Round(ventureMat(2), 2), "## ##0.00")


'
' ------------------ материалы не под заказ -------------------
'
strWhere = getWhereByDateBoxes(Me, "sDocs.xDate", begDateHron) ' между
If strWhere <> "" Then strWhere = strWhere & " AND "
mDateWhere = strWhere & "(sDocs.numExt = 254) AND (sDocs.destId >-1000) And (sDocs.destId < 0)"


sql = "SELECT Sum(sDMC.quant*n.cost/n.perList) AS sum " & _
"FROM sGuideSource INNER JOIN (sGuideNomenk n INNER JOIN (sDocs INNER JOIN sDMC ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc)) ON n.nomNom = sDMC.nomNom) ON sGuideSource.sourceId = sDocs.destId " & _
"WHERE " & mDateWhere

Debug.Print sql

If byErrSqlGetValues("##404", sql, s) Then
    laOther.Caption = Format(s, "## ##0.00")
End If



EN1:
Me.MousePointer = flexDefault
cmSave.Enabled = True
cmDetail.Enabled = True
cmDetail1.Enabled = True
cmDetail2.Enabled = True
cmDetail3.Enabled = True
cmStat1.Enabled = True
cmStat2.Enabled = True
cmStat3.Enabled = True

cmDetailPM.Enabled = True
cmDetailMM.Enabled = True
cmDetailAN.Enabled = True
cmStatPM.Enabled = True
cmStatMM.Enabled = True
cmStatAN.Enabled = True


End Sub

Sub addRowToBook(sum As String)
Dim mig As Single, str As String

' As Integer,  As Integer,  As Integer,
'detailId As Integer, purposeId As Integer, KredDebitor
tbDocs.AddNew
mig = Timer
While (Timer - mig < 1#): DoEvents: Wend  ' Дата в сек. д.б. уникальной

tmpStr = Format(Now(), "dd/mm/yy hh:nn:ss")
tbDocs!xDate = CDate(tmpStr)
tbDocs!UEsumm = sum
tbDocs!m = AUTO.cbM.Text
tbDocs!debit = debit
tbDocs!subDebit = subDebit
tbDocs!kredit = kredit
tbDocs!subKredit = subKredit
str = tbStartDate.Text & " - " & tbEndDate.Text
tbDocs!note = str
tbDocs!purposeId = purposeId
'tbDocs!detailId = detailId

On Error GoTo ER1
tbDocs.Update
Journal.addRowToGrid sum, str
Exit Sub
ER1:
errorCodAndMsg "378" '##378

'MsgBox Error, , "Ошибка 378-" & Err & ":  " '##378

End Sub

Private Sub cmMaterials_Click()

End Sub

Private Sub cmRealiz_Click()

End Sub

Private Sub cmSave_Click() '$$4
Dim str As String, str2 As String, i As Integer
cmSave.Enabled = False
'If Not (IsNumeric(laRealiz.Caption) And IsNumeric(laRealiz.Caption)) Then

str = "В Настройках не заданы параметры записи для поля '"
str2 = "' в журнал Х.О., поэтому эта запись произведена НЕ будет."

Set tbDocs = myOpenRecordSet("##376", "yBook", dbOpenTable) 'dbOpenForwardOnly)
'If tbDocs Is Nothing Then Exit Sub

strWhere = "SELECT 1,Debit, subDebit, Kredit, subKredit, pId as purposeId " & _
"From yGuidePurpose WHERE (((auto)='"

sql = strWhere & "r'));"
If Not byErrSqlGetValues("W##401", sql, i, debit, subDebit, kredit, subKredit, _
purposeId) Then GoTo EN1
If i = 0 Then 'err WHERE
    MsgBox str & laHRealiz.Caption & str2, , "Предупреждение"
    GoTo EN1
Else
    addRowToBook laRealiz.Caption
End If

sql = strWhere & "z'));"
If Not byErrSqlGetValues("W##402", sql, i, debit, subDebit, kredit, subKredit, _
purposeId) Then GoTo EN1
If i = 0 Then 'err WHERE
    MsgBox str & laHMaterials.Caption & str2, , "Предупреждение"
    GoTo EN1
Else
    addRowToBook laMaterials.Caption
End If

sql = strWhere & "m'));"
If Not byErrSqlGetValues("W##403", sql, i, debit, subDebit, kredit, subKredit, _
purposeId) Then GoTo EN1
If i = 0 Then 'err WHERE
    MsgBox str & laOther.Caption & str2, , "Предупреждение"
Else
    addRowToBook laOther.Caption
End If
EN1:
tbDocs.Close
Unload Me
On Error Resume Next
Journal.Grid.SetFocus
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmStat1_Click()
Me.MousePointer = flexHourglass
'Report.param1 = laRealiz1.Caption
Report.param1 = laProduct.Caption
Report.param2 = laMaterials1.Caption
Report.Regim = "relizStatistic"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmStat2_Click()
Me.MousePointer = flexHourglass
Report.param1 = laRealiz2.Caption
Report.param2 = laMaterials2.Caption
Report.Regim = "bayStatistic"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmStat3_Click()

Me.MousePointer = flexHourglass
Report.param1 = laUslug.Caption
Report.Regim = "uslugStatistic"
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmStatAN_Click()
Me.MousePointer = flexHourglass
statistic = True
ventureId = 3
Report.Regim = "venture"
Report.Show vbModal
Me.MousePointer = flexDefault
End Sub

Private Sub cmStatMM_Click()
Me.MousePointer = flexHourglass
statistic = True
ventureId = 2
Report.Regim = "venture"
Report.Show vbModal
Me.MousePointer = flexDefault
End Sub

Private Sub cmStatPM_Click()
Me.MousePointer = flexHourglass
statistic = True
ventureId = 1
Report.Regim = "venture"
Report.Show vbModal
Me.MousePointer = flexDefault
End Sub

Private Sub Form_Load()
tbEndDate.Text = Format(CurDate, "dd/mm/yy")
begDateHron = "01.09.03" '
tbStartDate.Text = "01." & Format(CurDate, "mm/yy")
'tbStartDate.Text = "01.09.03"
'tbStartDate.Text = "01.10.07"
End Sub
'для отчета Прибыль
Function getProductNomenkSum() As Variant
Dim i As Integer, j As Integer, gr() As String, sum As Single

getProductNomenkSum = Null
'вариантная ном-ра изделия

'sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup " & _
"FROM sProducts INNER JOIN xVariantNomenc ON (sProducts.nomNom = " & _
"xVariantNomenc.nomNom) AND (sProducts.ProductId = xVariantNomenc.prId) " & _
"WHERE (((xVariantNomenc.numOrder)=" & numDoc & ") AND (" & _
"(xVariantNomenc.prId)=" & gProductId & ") AND ((xVariantNomenc.prExt)=" & prExt & "));"
'!!!
sql = "SELECT sProducts.xgroup, n.cost*sProducts.quantity" & _
"/n.perList AS sum " & _
"FROM (sGuideNomenk n INNER JOIN sProducts ON n.nomNom = " & _
"sProducts.nomNom) INNER JOIN xVariantNomenc ON (sProducts.nomNom = " & _
"xVariantNomenc.nomNom) AND (sProducts.ProductId = xVariantNomenc.prId) " & _
"WHERE (((xVariantNomenc.numOrder)=" & gNzak & ") AND (" & _
"(xVariantNomenc.prId)=" & gProductId & ") AND ((xVariantNomenc.prExt)=" & prExt & "));"

'MsgBox sql
Set tbNomenk = myOpenRecordSet("##192", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
ReDim gr(0): i = 0: sum = 0
While Not tbNomenk.EOF
    i = i + 1
    sum = sum + tbNomenk!sum
    ReDim Preserve gr(i): gr(i) = tbNomenk!xgroup
    tbNomenk.MoveNext
Wend
tbNomenk.Close
    
'НЕвариантная ном-ра изделия
'sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup " & _
"From sProducts WHERE (((sProducts.ProductId)=" & gProductId & "));"
'cost-цена фактическая !!!
sql = "SELECT sProducts.xgroup, n.cost*sProducts.quantity" & _
"/n.perList AS sum " & _
"FROM sGuideNomenk n INNER JOIN sProducts ON n.nomNom = sProducts.nomNom " & _
"WHERE (((sProducts.ProductId)=" & gProductId & "));"

'MsgBox sql
Set tbNomenk = myOpenRecordSet("##177", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
While Not tbNomenk.EOF
    For j = 1 To UBound(gr) ' если группа состоит из одной ном-ры, то она
        If gr(j) = tbNomenk!xgroup Then GoTo NXT ' НЕвариантна, т.к. не
    Next j                                      ' не попала в xVariantNomenc
    i = i + 1
    sum = sum + tbNomenk!sum
NXT: tbNomenk.MoveNext
Wend
tbNomenk.Close
getProductNomenkSum = sum
End Function

