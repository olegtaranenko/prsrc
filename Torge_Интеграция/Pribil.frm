VERSION 5.00
Begin VB.Form Pribil 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Реализация"
   ClientHeight    =   7284
   ClientLeft      =   552
   ClientTop       =   9336
   ClientWidth     =   11424
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7284
   ScaleWidth      =   11424
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ckSaleNomenk 
      BackColor       =   &H8000000A&
      Caption         =   "Учет Номенклат. =>"
      Height          =   192
      Left            =   3960
      TabIndex        =   68
      Top             =   5880
      Width           =   1932
   End
   Begin VB.CheckBox ckStatistic 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Caption         =   "<= Стат-ка/Затраты"
      CausesValidation=   0   'False
      Height          =   192
      Left            =   360
      TabIndex        =   46
      Top             =   5880
      Width           =   1932
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   480
      TabIndex        =   31
      Top             =   1080
      Width           =   10692
      Begin VB.CommandButton cmDetailPM 
         Caption         =   "Петровские Мастерские"
         Enabled         =   0   'False
         Height          =   432
         Left            =   120
         TabIndex        =   32
         Top             =   100
         Width           =   1215
      End
      Begin VB.Label laPmResultTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   9360
         TabIndex        =   53
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label laPmResultMain 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   6720
         TabIndex        =   52
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label laPmAddCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   8040
         TabIndex        =   51
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label laPmMainCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   5400
         TabIndex        =   50
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label laMaterialsPM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   2760
         TabIndex        =   45
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label laRealizPM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   1440
         TabIndex        =   34
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label laClearPM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   4080
         TabIndex        =   33
         Top             =   180
         Width           =   1200
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   480
      TabIndex        =   40
      Top             =   1680
      Width           =   10692
      Begin VB.CommandButton cmDetailMM 
         Caption         =   "Маркмастер"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label laMmResultTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   9360
         TabIndex        =   62
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMmResultMain 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   6720
         TabIndex        =   61
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMmAddCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   8040
         TabIndex        =   60
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMmMainCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   5400
         TabIndex        =   59
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laClearMM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   4080
         TabIndex        =   44
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMaterialsMM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   2760
         TabIndex        =   43
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laRealizMM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   1440
         TabIndex        =   42
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   480
      TabIndex        =   35
      Top             =   2160
      Width           =   10692
      Begin VB.CommandButton cmDetailAN 
         Caption         =   "Аналитика"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label laAnResultTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   9360
         TabIndex        =   66
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laAnResultMain 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   6720
         TabIndex        =   65
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laAnAddCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   8040
         TabIndex        =   64
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laAnMainCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   5400
         TabIndex        =   63
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laRealizAn 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   1440
         TabIndex        =   39
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laClearAn 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   4080
         TabIndex        =   38
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMaterialsAn 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   2760
         TabIndex        =   37
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   480
      TabIndex        =   27
      Top             =   3240
      Width           =   5412
      Begin VB.CommandButton cmDetail3 
         Caption         =   "Услуги"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   1212
      End
      Begin VB.Label laRealiz1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   4080
         TabIndex        =   30
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laUslug 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   1440
         TabIndex        =   29
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   480
      TabIndex        =   23
      Top             =   2640
      Width           =   10692
      Begin VB.CommandButton cmItogo 
         Caption         =   "Всего"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label laTotalResultTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   312
         Left            =   9360
         TabIndex        =   58
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laTotalResultMain 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   312
         Left            =   6720
         TabIndex        =   57
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laTotalAddCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   312
         Left            =   8040
         TabIndex        =   56
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laTotalMainCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   312
         Left            =   5400
         TabIndex        =   55
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laRealiz 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   312
         Left            =   1440
         TabIndex        =   26
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laClear 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   312
         Left            =   4080
         TabIndex        =   25
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMaterials 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   312
         Left            =   2760
         TabIndex        =   24
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   480
      TabIndex        =   18
      Top             =   4440
      Width           =   5412
      Begin VB.CommandButton cmSales 
         Caption         =   " Продажа"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label laMaterials2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   2760
         TabIndex        =   22
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laClear2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   4080
         TabIndex        =   21
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laRealiz2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   1440
         TabIndex        =   20
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Материалы не под заказы:"
      Height          =   552
      Left            =   480
      TabIndex        =   15
      Top             =   5040
      Width           =   5352
      Begin VB.CommandButton cmDetail 
         Caption         =   "Списания"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1212
      End
      Begin VB.Label laOther 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   2760
         TabIndex        =   17
         Top             =   120
         Width           =   1152
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   480
      TabIndex        =   10
      Top             =   3840
      Width           =   5412
      Begin VB.CommandButton cmDetail1 
         Caption         =   "Товары"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label laProduct 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   1440
         TabIndex        =   14
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMaterials1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   2760
         TabIndex        =   13
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laClear1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   4080
         TabIndex        =   12
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   552
      Left            =   60
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
         BackStyle       =   0  'Transparent
         Caption         =   "Период  с "
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   795
      End
      Begin VB.Label laPo 
         BackStyle       =   0  'Transparent
         Caption         =   "по"
         Height          =   195
         Left            =   1785
         TabIndex        =   8
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Line Line15 
      X1              =   6120
      X2              =   6000
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line20 
      X1              =   120
      X2              =   240
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line19 
      X1              =   120
      X2              =   120
      Y1              =   6000
      Y2              =   2880
   End
   Begin VB.Line Line18 
      X1              =   120
      X2              =   240
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line17 
      X1              =   240
      X2              =   240
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line16 
      X1              =   6000
      X2              =   6000
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line14 
      X1              =   6120
      X2              =   6120
      Y1              =   4800
      Y2              =   6000
   End
   Begin VB.Line Line13 
      X1              =   6120
      X2              =   6000
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line12 
      X1              =   5880
      X2              =   6000
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line11 
      X1              =   240
      X2              =   600
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line10 
      X1              =   6000
      X2              =   5760
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line9 
      X1              =   3840
      X2              =   3840
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line8 
      X1              =   3840
      X2              =   6000
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line7 
      X1              =   3840
      X2              =   6000
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line3 
      X1              =   6000
      X2              =   6000
      Y1              =   3840
      Y2              =   5640
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line5 
      X1              =   2400
      X2              =   2400
      Y1              =   5760
      Y2              =   6120
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   2400
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   2400
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   240
      Y1              =   960
      Y2              =   4920
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Окончат. результат"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   9840
      TabIndex        =   54
      Top             =   720
      Width           =   1152
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Основные затраты:"
      Height          =   432
      Left            =   5880
      TabIndex        =   49
      Top             =   720
      Width           =   1212
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Вспомогат. затраты"
      Height          =   372
      Left            =   8520
      TabIndex        =   48
      Top             =   720
      Width           =   1152
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Прибыль.- Основные"
      Height          =   372
      Left            =   7200
      TabIndex        =   47
      Top             =   720
      Width           =   1212
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Прибыль"
      Height          =   252
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      Width           =   972
   End
   Begin VB.Label laHMaterials 
      BackStyle       =   0  'Transparent
      Caption         =   "Материалы"
      Height          =   252
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   1032
   End
   Begin VB.Label laHRealiz 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Реализация:"
      Height          =   192
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   1092
   End
End
Attribute VB_Name = "Pribil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pDateWhere As String, nDateWhere As String, uDateWhere As String
Public bDateWhere As String, mDateWhere As String, costsDateWhere As String
Public StartDate As String, EndDate As String
Public statistic As String, ventureId As String

Dim begDateHron As Date ' Начало ведения хронологии

Private Sub ckSaleNomenk_Click()
    If ckStatistic.value = 1 Then
        ckStatistic.value = 0
    End If

End Sub

Private Sub ckStatistic_Click()
    If ckSaleNomenk.value = 1 Then
        ckSaleNomenk.value = 0
    End If
End Sub

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
If ckStatistic.value = 1 Then
    Report.Regim = "relizStatistic"
Else
    Report.Regim = ""
End If

Report.Show vbModal
Me.MousePointer = flexDefault

End Sub


Private Sub cmDetail3_Click()
Me.MousePointer = flexHourglass
'Report.param1 = laRealiz1.Caption
Report.param1 = laUslug.Caption
'Report.param2 = laMaterials1.Caption
If ckStatistic.value = 1 Then
    Report.Regim = "uslugStatistic"
Else
    Report.Regim = "uslug"
End If
Report.Show vbModal
Me.MousePointer = flexDefault

End Sub

Private Sub cmDetailAN_Click()
    Me.MousePointer = flexHourglass
    Report.param2 = laAnMainCosts.Caption
    Report.param1 = laAnAddCosts.Caption
    setVentureRegim
    ventureId = 3
    Report.Show vbModal
    Me.MousePointer = flexDefault
End Sub

Private Sub cmDetailMM_Click()
    Me.MousePointer = flexHourglass
    Report.param2 = laMmMainCosts.Caption
    Report.param1 = laMmAddCosts.Caption
    setVentureRegim
    ventureId = 2
    Report.Show vbModal
    Me.MousePointer = flexDefault
End Sub

Private Sub cmDetailPM_Click()
    Me.MousePointer = flexHourglass
    Report.param2 = laPmMainCosts.Caption
    Report.param1 = laPmAddCosts.Caption
    setVentureRegim
    ventureId = 1
    Report.Show vbModal
    Me.MousePointer = flexDefault
End Sub

Private Sub setVentureRegim()
    If ckStatistic.value = 1 Then
        Report.Regim = "ventureZatrat"
    Else
        Report.Regim = "venture"
    End If
End Sub
Private Sub cmManag_Click() 'кнопка "применить" из отчета "Реализация"
Dim oborot As Single, dohod As Single, s2 As Single, s As Single
Dim ventureMat() As Single, ventureRealiz() As Single
Dim mainCosts() As Single, addCosts() As Single
Dim mat As Single, realiz As Single
Dim mainCostsTotal As Single, addCostsTotal As Single



Me.MousePointer = flexHourglass

ReDim ventureMat(2)
ReDim ventureRealiz(2)


If isDateTbox(tbStartDate) Then
    StartDate = Format(tmpDate, "yyyy-mm-dd")
Else
    StartDate = "null"
End If

If isDateTbox(tbEndDate) Then
    EndDate = Format(tmpDate, "yyyy-mm-dd") & " 11:59:59 PM'"
Else
    EndDate = "null"
End If

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

ReDim mainCosts(2)
ReDim addCosts(2)

'
' ------------------ основные и вспомогательные затраты -------------------
'
strWhere = getWhereByDateBoxes(Me, "xDate", begDateHron)
costsDateWhere = strWhere


sql = "select sum(uesumm) as total, ventureid, is_main_costs" _
& " from ybook b" _
& " join shiz s on s.id = b.id_shiz" _
& " where " & costsDateWhere _
& " and s.is_main_costs is not null" _
& " group by ventureid, is_main_costs"

'Debug.Print sql

Set tbNomenk = myOpenRecordSet("##431", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    If tbNomenk!is_main_costs = 1 Then
        mainCosts(tbNomenk!ventureId - 1) = CSng(tbNomenk!total)
        mainCostsTotal = mainCostsTotal + tbNomenk!total
    Else
        addCosts(tbNomenk!ventureId - 1) = CSng(tbNomenk!total)
        addCostsTotal = addCostsTotal + tbNomenk!total
    End If
    tbNomenk.MoveNext
  Wend
End If
tbNomenk.Close

laTotalMainCosts.Caption = Format(Round(mainCostsTotal, 2), "## ##0.00")
laTotalResultMain.Caption = Format(Round(dohod - oborot - mainCostsTotal, 2), "## ##0.00")
laTotalAddCosts.Caption = Format(Round(addCostsTotal, 2), "## ##0.00")
laTotalResultTotal.Caption = Format(Round(dohod - mainCostsTotal - addCostsTotal, 2), "## ##0.00")

laPmMainCosts.Caption = Format(Round(mainCosts(0), 2), "## ##0.00")
laPmResultMain.Caption = Format(Round(ventureRealiz(0) - ventureMat(0) - mainCosts(0), 2), "## ##0.00")
laPmAddCosts.Caption = Format(Round(addCosts(0), 2), "## ##0.00")
laPmResultTotal.Caption = Format(Round(ventureRealiz(0) - mainCosts(0) - addCosts(0), 2), "## ##0.00")

laMmMainCosts.Caption = Format(Round(mainCosts(1), 2), "## ##0.00")
laMmResultMain.Caption = Format(Round(ventureRealiz(1) - ventureMat(1) - mainCosts(1), 2), "## ##0.00")
laMmAddCosts.Caption = Format(Round(addCosts(1), 2), "## ##0.00")
laMmResultTotal.Caption = Format(Round(ventureRealiz(1) - mainCosts(1) - addCosts(1), 2), "## ##0.00")

laAnMainCosts.Caption = Format(Round(mainCosts(2), 2), "## ##0.00")
laAnResultMain.Caption = Format(Round(ventureRealiz(2) - ventureMat(2) - mainCosts(2), 2), "## ##0.00")
laAnAddCosts.Caption = Format(Round(addCosts(2), 2), "## ##0.00")
laAnResultTotal.Caption = Format(Round(ventureRealiz(2) - mainCosts(2) - addCosts(2), 2), "## ##0.00")

EN1:
Me.MousePointer = flexDefault
cmSave.Enabled = True
cmDetail.Enabled = True
cmDetail1.Enabled = True
cmSales.Enabled = True
cmItogo.Enabled = True
'cmSaleNomenk.Enabled = True

cmDetail3.Enabled = True

cmDetailPM.Enabled = True
cmDetailMM.Enabled = True
cmDetailAN.Enabled = True



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
tbDocs!uesumm = sum
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

Private Sub cmSales_Click()
Me.MousePointer = flexHourglass
Report.param1 = laRealiz2.Caption
Report.param2 = laMaterials2.Caption
If ckStatistic.value = 1 Then
    Report.Regim = "bayStatistic"
Else
    If ckSaleNomenk.value = 1 Then
        Report.Regim = "bayNomenk"
    Else
        Report.Regim = "bay"
    End If
End If
Report.Show vbModal
Me.MousePointer = flexDefault

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


Private Sub Form_Load()
tbEndDate.Text = Format(CurDate, "dd/mm/yy")
begDateHron = "01.09.03" '
tbStartDate.Text = "01." & Format(CurDate, "mm/yy")
'tbStartDate.Text = "01.09.03"
'tbStartDate.Text = "01.12.07"
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

