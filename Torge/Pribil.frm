VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Pribil 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����������"
   ClientHeight    =   5784
   ClientLeft      =   552
   ClientTop       =   9336
   ClientWidth     =   11376
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5784
   ScaleWidth      =   11376
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   240
      TabIndex        =   32
      Top             =   1680
      Width           =   11052
      Begin VB.CommandButton cmDetailPM 
         Caption         =   "���������� ����������"
         Enabled         =   0   'False
         Height          =   432
         Left            =   120
         TabIndex        =   31
         Top             =   96
         Width           =   1215
      End
      Begin VB.Label laPmResultTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   9840
         TabIndex        =   53
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label laPmResultMain 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   6240
         TabIndex        =   52
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label laPmAddCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   7440
         TabIndex        =   51
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label laPmMainCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   5040
         TabIndex        =   50
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label laPmturnCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   8640
         TabIndex        =   72
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label laMaterialsPM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   2640
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
         Left            =   3840
         TabIndex        =   33
         Top             =   180
         Width           =   1200
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   240
      TabIndex        =   40
      Top             =   2280
      Width           =   11052
      Begin VB.CommandButton cmDetailMM 
         Caption         =   "����������"
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
         Left            =   9840
         TabIndex        =   62
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMmResultMain 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   6240
         TabIndex        =   61
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMmAddCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   7440
         TabIndex        =   60
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMmMainCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   5040
         TabIndex        =   59
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMmturnCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   8640
         TabIndex        =   73
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laClearMM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   3840
         TabIndex        =   44
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMaterialsMM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   2640
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
      Left            =   240
      TabIndex        =   35
      Top             =   2760
      Width           =   11052
      Begin VB.CommandButton cmDetailAN 
         Caption         =   "���������"
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
         Left            =   9840
         TabIndex        =   66
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laAnResultMain 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   6240
         TabIndex        =   65
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laAnAddCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   7440
         TabIndex        =   64
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laAnMainCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   5040
         TabIndex        =   63
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laAnturnCosts 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   8640
         TabIndex        =   74
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
         Left            =   3840
         TabIndex        =   38
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laMaterialsAn 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   2640
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
      Left            =   240
      TabIndex        =   27
      Top             =   3720
      Width           =   5412
      Begin VB.CommandButton cmDetail3 
         Caption         =   "������"
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
         Left            =   3840
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
      Left            =   240
      TabIndex        =   23
      Top             =   3240
      Width           =   11052
      Begin VB.CommandButton cmItogo 
         Caption         =   "�����"
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
         Left            =   9840
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
         Left            =   6240
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
         Left            =   7440
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
         Left            =   5040
         TabIndex        =   55
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laTotalturnCosts 
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
         Left            =   8640
         TabIndex        =   75
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
         Left            =   3840
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
         Left            =   2640
         TabIndex        =   24
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   240
      TabIndex        =   18
      Top             =   4680
      Width           =   5412
      Begin VB.CommandButton cmSales 
         Caption         =   " �������"
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
         Left            =   2640
         TabIndex        =   22
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laClear2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   3840
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
      Caption         =   "��������� �� ��� ������:"
      Height          =   552
      Left            =   240
      TabIndex        =   15
      Top             =   5160
      Width           =   5352
      Begin VB.CommandButton cmDetail 
         Caption         =   "��������"
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
         Left            =   2640
         TabIndex        =   17
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   240
      TabIndex        =   10
      Top             =   4200
      Width           =   5412
      Begin VB.CommandButton cmDetail1 
         Caption         =   "������"
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
         Left            =   2640
         TabIndex        =   13
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label laClear1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   3840
         TabIndex        =   12
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1332
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   10992
      Begin VB.OptionButton rbNomenk 
         BackColor       =   &H8000000A&
         Caption         =   "�� ������������"
         Height          =   252
         Left            =   4800
         TabIndex        =   71
         Top             =   960
         Width           =   1812
      End
      Begin VB.OptionButton rbDetailMode 
         BackColor       =   &H8000000A&
         Caption         =   "�� ��������"
         Height          =   252
         Left            =   3240
         TabIndex        =   70
         Top             =   960
         Width           =   1452
      End
      Begin VB.OptionButton rbOrders 
         BackColor       =   &H8000000A&
         Caption         =   "�� �������"
         Height          =   252
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1332
      End
      Begin VB.OptionButton rbStatistic 
         BackColor       =   &H8000000A&
         Caption         =   "�� ������"
         Height          =   252
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   1452
      End
      Begin VB.ComboBox cbDateShift 
         Height          =   288
         ItemData        =   "Pribil.frx":0000
         Left            =   6000
         List            =   "Pribil.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   68
         ToolTipText     =   "�������� ��� ���� ��"
         Top             =   180
         Width           =   2412
      End
      Begin VB.TextBox tbStartDate 
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   1
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox tbEndDate 
         Height          =   285
         Left            =   2340
         MaxLength       =   8
         TabIndex        =   2
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton cmManag 
         Caption         =   "���������"
         Height          =   315
         Left            =   3480
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   372
         Left            =   5640
         TabIndex        =   77
         ToolTipText     =   "��������� ������������ �������� ���� �� ���������� ������"
         Top             =   120
         Width           =   312
         _ExtentX        =   550
         _ExtentY        =   656
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label laDetailMode 
         BackColor       =   &H8000000A&
         Caption         =   "����� �����������"
         Height          =   252
         Left            =   240
         TabIndex        =   69
         Top             =   650
         Width           =   2412
      End
      Begin VB.Label laPeriod 
         BackStyle       =   0  'Transparent
         Caption         =   "������  � "
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   795
      End
      Begin VB.Label laPo 
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   192
         Left            =   2028
         TabIndex        =   9
         Top             =   240
         Width           =   432
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��������� ��������"
      Height          =   372
      Left            =   8880
      TabIndex        =   76
      Top             =   1320
      Width           =   1152
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   10080
      TabIndex        =   54
      Top             =   1440
      Width           =   1152
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      Height          =   312
      Left            =   5160
      TabIndex        =   49
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "������ �������"
      Height          =   372
      Left            =   7560
      TabIndex        =   48
      Top             =   1320
      Width           =   1152
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      Height          =   252
      Left            =   6480
      TabIndex        =   47
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      Height          =   252
      Left            =   4080
      TabIndex        =   6
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label laHMaterials 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "���������"
      Height          =   252
      Left            =   2880
      TabIndex        =   5
      Top             =   1440
      Width           =   1152
   End
   Begin VB.Label laHRealiz 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "����������:"
      Height          =   192
      Left            =   1680
      TabIndex        =   0
      Top             =   1440
      Width           =   1212
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
Public statistic As String, ventureId As String
Dim flagSync As Boolean

Dim begDateHron As Date ' ������ ������� ����������


Private Sub cmSales_Click()
    Report.param1 = laRealiz2.Caption
    Report.param2 = laMaterials2.Caption
    
    If rbOrders.Value = True Then
        Report.Regim = "bay"
        Report.Sortable = True
    ElseIf rbStatistic.Value = True Then
        Report.Regim = "bayStatistic"
        Report.Sortable = True
    ElseIf rbNomenk.Value = True Then
        Report.Regim = "bayNomenk"
    Else
        Exit Sub
    End If
    
    Set Report.Caller = Me
    Report.Show vbModal

End Sub


Private Sub cmDetail_Click()
    Report.param1 = laOther.Caption
    Report.Regim = "mat"
    Report.Sortable = True
    Set Report.Caller = Me
    Report.Show vbModal

End Sub

Private Sub cmDetail1_Click()
    'Report.param1 = laRealiz1.Caption
    Set Report.Caller = Me
    Report.param1 = laProduct.Caption
    Report.param2 = laMaterials1.Caption
    
    If rbOrders.Value = True Then
        Report.Regim = ""
    ElseIf rbStatistic.Value = True Then
        Report.Regim = "relizStatistic"
    ElseIf rbNomenk.Value = True Then
        Report.Regim = "relizNomenk"
    Else
        Exit Sub
    End If
    
    Report.Sortable = True
    Set Report.Caller = Me
    Report.Show vbModal

End Sub


Private Sub cmDetail3_Click()
    Report.param1 = laUslug.Caption
    If rbOrders.Value = True Then
        Report.Regim = "uslug"
    ElseIf rbStatistic.Value = True Then
        Report.Regim = "uslugStatistic"
    Else
        Exit Sub
    End If
        
    Report.Sortable = True
    Set Report.Caller = Me
    Report.Show vbModal

End Sub

Private Sub cmDetailAN_Click()
    Report.param2 = laAnMainCosts.Caption
    Report.param1 = laAnAddCosts.Caption
    setVentureRegim
    ventureId = 3
    Report.Sortable = True
    Set Report.Caller = Me
    Report.Show vbModal
End Sub

Private Sub cmDetailMM_Click()
    Report.param2 = laMmMainCosts.Caption
    Report.param1 = laMmAddCosts.Caption
    setVentureRegim
    ventureId = 2
    Report.Sortable = True
    Set Report.Caller = Me
    Report.Show vbModal
End Sub

Private Sub cmDetailPM_Click()
    Report.param2 = laPmMainCosts.Caption
    Report.param1 = laPmAddCosts.Caption
    setVentureRegim
    ventureId = 1
    Report.Sortable = True
    Set Report.Caller = Me
    Report.Show vbModal
End Sub

Private Sub cmItogo_Click()
    Report.param2 = laTotalMainCosts.Caption
    Report.param1 = laTotalAddCosts.Caption
    Report.param3 = laTotalturnCosts.Caption
    setVentureRegim
    ventureId = 0
    Report.Sortable = True
    Set Report.Caller = Me
    Report.Show vbModal
End Sub

Private Sub setVentureRegim()
    If rbOrders.Value = True Then
        Report.Regim = "venture"
    ElseIf rbDetailMode.Value = True Then
        Report.Regim = "ventureZatrat"
    Else
        Exit Sub
    End If
    
End Sub


Private Sub cmManag_Click() '������ "���������" �� ������ "����������"
Dim oborot As Single, dohod As Single, s2 As Single, s As Single
Dim ventureMat() As Single, ventureRealiz() As Single
Dim mainCosts() As Single, addCosts() As Single, turnCosts() As Single
Dim mat As Single, realiz As Single
Dim mainCostsTotal As Single, addCostsTotal As Single, turnCostsTotal As Single



Me.MousePointer = flexHourglass

ReDim ventureMat(2)
ReDim ventureRealiz(2)

setStartEndDates tbStartDate, tbEndDate

strWhere = getWhereByDateBoxes(Me, "outDate", begDateHron) ' �����
If strWhere = "error" Then GoTo EN1
If strWhere <> "" Then strWhere = " WHERE " & strWhere
pDateWhere = strWhere

'���������� �� ��������
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
    gNzak = tbProduct!Numorder
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

'���������� �� ���.������������
strWhere = getWhereByDateBoxes(Me, "outDate", begDateHron) ' �����
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


' ------------------ ������ -------------------
strWhere = getWhereByDateBoxes(Me, "u.outDate", begDateHron) ' �����
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
    realiz = tbNomenk!Sum_quant
    s2 = s2 + realiz
    ventureRealiz(tbNomenk!ventureId - 1) = ventureRealiz(tbNomenk!ventureId - 1) + realiz
    dohod = dohod + realiz
    tbNomenk.MoveNext
  Wend
End If
tbNomenk.Close

laUslug.Caption = Format(Round(s2, 2), "## ##0.00")
laRealiz1.Caption = laUslug.Caption



' ------------------ ��������� �� ������� ������ -------------------
strWhere = getWhereByDateBoxes(Me, "outDate", begDateHron) ' �����
bDateWhere = strWhere

sql = "select sum(cenaed * quant) as bSum, sum(costEd * quant) as cSum, isnull(ventureid, 1) as venture_id" _
    & " from itemWallShip " _
    & " where type = 8 and " & strWhere _
    & " group by venture_id "

'Debug.Print sql

s = 0
s2 = 0
Set tbNomenk = myOpenRecordSet("##431", sql, dbOpenForwardOnly)
If tbNomenk Is Nothing Then GoTo EN1
If Not tbNomenk.BOF Then
  While Not tbNomenk.EOF
    realiz = tbNomenk!bSum
    mat = tbNomenk!cSum
    s = s + realiz
    s2 = s2 + mat
    oborot = oborot + mat
    ventureRealiz(tbNomenk!venture_Id - 1) = ventureRealiz(tbNomenk!venture_Id - 1) + realiz
    ventureMat(tbNomenk!venture_Id - 1) = ventureMat(tbNomenk!venture_Id - 1) + mat
    dohod = dohod + realiz
    tbNomenk.MoveNext
  Wend
End If
tbNomenk.Close

laRealiz2.Caption = Format(s, "## ##0.00")
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
' ------------------ ��������� �� ��� ����� -------------------
'
strWhere = getWhereByDateBoxes(Me, "sDocs.xDate", begDateHron) ' �����
If strWhere <> "" Then strWhere = strWhere & " AND "
mDateWhere = strWhere & "(sDocs.numExt = 254) AND (sDocs.destId >-1000) And (sDocs.destId < 0)"


sql = "SELECT Sum(sDMC.quant*n.cost/n.perList) AS sum " & _
"FROM sGuideSource INNER JOIN (sGuideNomenk n INNER JOIN (sDocs INNER JOIN sDMC ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc)) ON n.nomNom = sDMC.nomNom) ON sGuideSource.sourceId = sDocs.destId " & _
"WHERE " & mDateWhere

'Debug.Print sql

If byErrSqlGetValues("##404", sql, s) Then
    laOther.Caption = Format(s, "## ##0.00")
End If

ReDim mainCosts(2)
ReDim turnCosts(2)
ReDim addCosts(2)

'
' ------------------ �������� � ��������������� ������� -------------------
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
    ElseIf tbNomenk!is_main_costs = 2 Then
        turnCosts(tbNomenk!ventureId - 1) = CSng(tbNomenk!total)
        turnCostsTotal = turnCostsTotal + tbNomenk!total
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
laTotalturnCosts.Caption = Format(Round(turnCostsTotal, 2), "## ##0.00")
laTotalResultTotal.Caption = Format(Round(dohod - mainCostsTotal - addCostsTotal - turnCostsTotal, 2), "## ##0.00")

laPmMainCosts.Caption = Format(Round(mainCosts(0), 2), "## ##0.00")
laPmResultMain.Caption = Format(Round(ventureRealiz(0) - ventureMat(0) - mainCosts(0), 2), "## ##0.00")
laPmAddCosts.Caption = Format(Round(addCosts(0), 2), "## ##0.00")
laPmturnCosts.Caption = Format(Round(turnCosts(0), 2), "## ##0.00")
laPmResultTotal.Caption = Format(Round(ventureRealiz(0) - mainCosts(0) - addCosts(0) - turnCosts(0), 2), "## ##0.00")

laMmMainCosts.Caption = Format(Round(mainCosts(1), 2), "## ##0.00")
laMmResultMain.Caption = Format(Round(ventureRealiz(1) - ventureMat(1) - mainCosts(1), 2), "## ##0.00")
laMmAddCosts.Caption = Format(Round(addCosts(1), 2), "## ##0.00")
laMmturnCosts.Caption = Format(Round(turnCosts(1), 2), "## ##0.00")
laMmResultTotal.Caption = Format(Round(ventureRealiz(1) - mainCosts(1) - addCosts(1) - turnCosts(1), 2), "## ##0.00")

laAnMainCosts.Caption = Format(Round(mainCosts(2), 2), "## ##0.00")
laAnResultMain.Caption = Format(Round(ventureRealiz(2) - ventureMat(2) - mainCosts(2), 2), "## ##0.00")
laAnAddCosts.Caption = Format(Round(addCosts(2), 2), "## ##0.00")
laAnturnCosts.Caption = Format(Round(turnCosts(2), 2), "## ##0.00")
laAnResultTotal.Caption = Format(Round(ventureRealiz(2) - mainCosts(2) - addCosts(2) - turnCosts(2), 2), "## ##0.00")

EN1:
Me.MousePointer = flexDefault
flagSync = True

    If rbDetailMode.Value = True Then
        rbDetailMode_Click
    ElseIf rbNomenk.Value = True Then
        rbNomenk_Click
    ElseIf rbOrders.Value = True Then
        rbOrders_Click
    ElseIf rbStatistic.Value = True Then
        rbStatistic_Click
    End If

End Sub

Sub addRowToBook(sum As String)
Dim mig As Single, str As String

' As Integer,  As Integer,  As Integer,
'detailId As Integer, purposeId As Integer, KredDebitor
tbDocs.AddNew
mig = Timer
While (Timer - mig < 1#): DoEvents: Wend  ' ���� � ���. �.�. ����������

tmpStr = Format(Now(), "dd/mm/yy hh:nn:ss")
tbDocs!xDate = CDate(tmpStr)
tbDocs!uesumm = sum
tbDocs!m = AUTO.cbM.Text
tbDocs!debit = debit
tbDocs!subDebit = subDebit
tbDocs!kredit = kredit
tbDocs!subKredit = subKredit
str = tbStartDate.Text & " - " & tbEndDate.Text
tbDocs!Note = str
tbDocs!purposeId = purposeId
'tbDocs!detailId = detailId

On Error GoTo ER1
tbDocs.Update
Journal.addRowToGrid sum, str
Exit Sub
ER1:
errorCodAndMsg "378" '##378

'MsgBox Error, , "������ 378-" & Err & ":  " '##378

End Sub


Private Sub Form_Load()
tbEndDate.Text = Format(CurDate, "dd/mm/yy")
begDateHron = "01.09.03" '
tbStartDate.Text = "01." & Format(CurDate, "mm/yy")
'tbStartDate.Text = "01.09.03"
'tbStartDate.Text = "01.12.07"
End Sub
'��� ������ �������
Function getProductNomenkSum() As Variant
Dim I As Integer, J As Integer, gr() As String, sum As Single

getProductNomenkSum = Null
'���������� ���-�� �������

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
ReDim gr(0): I = 0: sum = 0
While Not tbNomenk.EOF
    I = I + 1
    sum = sum + tbNomenk!sum
    ReDim Preserve gr(I): gr(I) = tbNomenk!xgroup
    tbNomenk.MoveNext
Wend
tbNomenk.Close
    
'������������ ���-�� �������
'sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup " & _
"From sProducts WHERE (((sProducts.ProductId)=" & gProductId & "));"
'cost-���� ����������� !!!
sql = "SELECT sProducts.xgroup, n.cost*sProducts.quantity" & _
"/n.perList AS sum " & _
"FROM sGuideNomenk n INNER JOIN sProducts ON n.nomNom = sProducts.nomNom " & _
"WHERE (((sProducts.ProductId)=" & gProductId & "));"

'MsgBox sql
Set tbNomenk = myOpenRecordSet("##177", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
While Not tbNomenk.EOF
    For J = 1 To UBound(gr) ' ���� ������ ������� �� ����� ���-��, �� ���
        If gr(J) = tbNomenk!xgroup Then GoTo NXT ' �����������, �.�. ��
    Next J                                      ' �� ������ � xVariantNomenc
    I = I + 1
    sum = sum + tbNomenk!sum
NXT: tbNomenk.MoveNext
Wend
tbNomenk.Close
getProductNomenkSum = sum
End Function


Private Sub rbDetailMode_Click()
    If Not flagSync Then
        disableAll
        Exit Sub
    End If
    cmDetailPM.Enabled = True
    cmDetailMM.Enabled = True
    cmDetailAN.Enabled = True
    cmItogo.Enabled = True
    cmDetail3.Enabled = False
    cmDetail1.Enabled = False
    cmSales.Enabled = False
    cmDetail.Enabled = False
End Sub

Private Sub rbNomenk_Click()
    If Not flagSync Then
        disableAll
        Exit Sub
    End If
    cmDetailPM.Enabled = False
    cmDetailMM.Enabled = False
    cmDetailAN.Enabled = False
    cmItogo.Enabled = False
    cmDetail3.Enabled = False
    cmDetail1.Enabled = True
    cmSales.Enabled = True
    cmDetail.Enabled = False
End Sub

Private Sub rbOrders_Click()
    If Not flagSync Then
        disableAll
        Exit Sub
    End If
    cmDetailPM.Enabled = True
    cmDetailMM.Enabled = True
    cmDetailAN.Enabled = True
    cmItogo.Enabled = True
    cmDetail3.Enabled = True
    cmDetail1.Enabled = True
    cmSales.Enabled = True
    cmDetail.Enabled = True
End Sub

Private Sub rbStatistic_Click()
    If Not flagSync Then
        disableAll
        Exit Sub
    End If
    cmDetailPM.Enabled = False
    cmDetailMM.Enabled = False
    cmDetailAN.Enabled = False
    cmItogo.Enabled = False
    cmDetail3.Enabled = True
    cmDetail1.Enabled = True
    cmSales.Enabled = True
    cmDetail.Enabled = False
End Sub

Private Sub tbStartDate_Change()
    flagSync = False
End Sub

Private Sub disableAll()
    cmDetailPM.Enabled = False
    cmDetailMM.Enabled = False
    cmDetailAN.Enabled = False
    cmItogo.Enabled = False
    cmDetail3.Enabled = False
    cmDetail1.Enabled = False
    cmSales.Enabled = False
    cmDetail.Enabled = False
End Sub


Private Sub UpDownChange(ByVal upDirection As Integer)
Dim startDate As Date, endDate As Date, shiftPeriod As String
Dim startDateNew As Date, endDateNew As Date
Dim dataAddInterval As String, dataAddNumber As Integer, shiftStrategy As Integer
Const ssDefault As Integer = 1
Const ssMonth   As Integer = 2

    If cbDateShift.ListIndex = -1 Then
        cbDateShift.ListIndex = 0
    End If

    If isDateTbox(tbStartDate) Then
        startDate = tmpDate
    Else
        Exit Sub
    End If
    If isDateTbox(tbEndDate) Then
        endDate = tmpDate
    Else
        Exit Sub
    End If
    
    shiftPeriod = cbDateShift.Text
    
    If shiftPeriod = "�����" Then
        dataAddInterval = "m"
        dataAddNumber = 1
        shiftStrategy = ssMonth
    ElseIf shiftPeriod = "���" Then
        dataAddInterval = "yyyy"
        dataAddNumber = 1
        shiftStrategy = ssMonth
    ElseIf shiftPeriod = "�������" Then
        dataAddInterval = "m"
        dataAddNumber = 3
        shiftStrategy = ssMonth
    ElseIf shiftPeriod = "������" Then
        dataAddInterval = "ww"
        dataAddNumber = 1
        shiftStrategy = ssDefault
    ElseIf shiftPeriod = "����" Then
        dataAddInterval = "d"
        dataAddNumber = 1
        shiftStrategy = ssDefault
    End If

    If shiftStrategy = ssDefault Then
        startDateNew = DateAdd(dataAddInterval, dataAddNumber * upDirection, startDate)
        endDateNew = DateAdd(dataAddInterval, dataAddNumber * upDirection, endDate)
    ElseIf shiftStrategy = ssMonth Then
        startDateNew = DateAdd(dataAddInterval, dataAddNumber * upDirection, startDate)
        Dim endDateNext As Date
        endDateNext = DateAdd("d", 1, endDate)
        endDateNext = DateAdd(dataAddInterval, dataAddNumber * upDirection, endDateNext)
        endDateNew = DateAdd("d", -1, endDateNext)
    End If
    
    tbStartDate.Text = Format(startDateNew, "dd.mm.yy")
    tbEndDate.Text = Format(endDateNew, "dd.mm.yy")

End Sub

Private Sub UpDown1_DownClick()
    UpDownChange -1

End Sub

Private Sub UpDown1_UpClick()
    UpDownChange 1
End Sub
