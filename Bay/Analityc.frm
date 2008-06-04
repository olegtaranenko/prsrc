VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Analityc 
   Caption         =   "Параметры запроса"
   ClientHeight    =   9192
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5364
   LinkTopic       =   "Form1"
   ScaleHeight     =   9192
   ScaleWidth      =   5364
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Группировки ..."
      Height          =   1332
      Left            =   240
      TabIndex        =   21
      Top             =   1800
      Width           =   4932
      Begin VB.TextBox tbTop 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   192
         Left            =   2040
         TabIndex        =   33
         Text            =   "10"
         Top             =   960
         Width           =   252
      End
      Begin VB.CheckBox ckTop 
         Caption         =   "Только первые "
         Height          =   252
         Left            =   360
         TabIndex        =   32
         Top             =   960
         Width           =   1572
      End
      Begin VB.ComboBox cbGroupByRow 
         Height          =   288
         ItemData        =   "Analityc.frx":0000
         Left            =   360
         List            =   "Analityc.frx":0013
         TabIndex        =   30
         Text            =   "Фирмы"
         Top             =   480
         Width           =   2052
      End
      Begin VB.ComboBox cbGroupByColumn 
         Height          =   288
         ItemData        =   "Analityc.frx":0053
         Left            =   2640
         List            =   "Analityc.frx":0069
         TabIndex        =   22
         Text            =   "Месяцы"
         Top             =   480
         Width           =   2172
      End
      Begin VB.Label Label4 
         Caption         =   "позиций"
         Height          =   252
         Left            =   2400
         TabIndex        =   34
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label3 
         Caption         =   "... по строкам"
         Height          =   252
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "... по столбцам"
         Height          =   252
         Left            =   2400
         TabIndex        =   23
         Top             =   240
         Width           =   1452
      End
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   4080
      TabIndex        =   25
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton cmApply 
      Caption         =   "Применить"
      Height          =   315
      Left            =   360
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Выбор периода"
      Height          =   1212
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   4932
      Begin VB.CheckBox ckStartDate 
         Caption         =   " "
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   195
      End
      Begin VB.CheckBox ckEndDate 
         Caption         =   " "
         Height          =   315
         Left            =   360
         TabIndex        =   18
         Top             =   600
         Width           =   200
      End
      Begin VB.CommandButton cmDayRight 
         Caption         =   ">"
         Height          =   252
         Left            =   2520
         TabIndex        =   17
         Top             =   600
         Width           =   372
      End
      Begin VB.CommandButton cmDayLeft 
         Caption         =   "<"
         Height          =   252
         Left            =   2040
         TabIndex        =   16
         Top             =   600
         Width           =   372
      End
      Begin VB.ComboBox cbDateShift 
         Height          =   288
         ItemData        =   "Analityc.frx":009A
         Left            =   3120
         List            =   "Analityc.frx":00B0
         TabIndex        =   15
         Text            =   "на год"
         Top             =   600
         Width           =   1692
      End
      Begin MSComCtl2.DTPicker tbStartDate 
         Height          =   288
         Left            =   720
         TabIndex        =   27
         Top             =   240
         Width           =   1092
         _ExtentX        =   1926
         _ExtentY        =   508
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   39599
      End
      Begin MSComCtl2.DTPicker tbEndDate 
         Height          =   288
         Left            =   720
         TabIndex        =   28
         Top             =   600
         Width           =   1092
         _ExtentX        =   1926
         _ExtentY        =   508
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   39599
      End
      Begin VB.Label Label5 
         Caption         =   "Одновременно сдвинуть даты"
         Height          =   252
         Left            =   1920
         TabIndex        =   35
         Top             =   360
         Width           =   2772
      End
      Begin VB.Label Label2 
         Caption         =   "C"
         Height          =   192
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   180
      End
      Begin VB.Label laPo 
         Caption         =   "по"
         Height          =   192
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   180
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Дополнительные условия"
      Height          =   4572
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   4932
      Begin VB.CheckBox ckKriteriumOborud 
         Caption         =   "Выбор оборудования"
         Height          =   252
         Left            =   120
         TabIndex        =   26
         Top             =   3840
         Width           =   3252
      End
      Begin VB.CheckBox ckKriteriumRegion 
         Caption         =   "Выбор региона"
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   3252
      End
      Begin MSComctlLib.TreeView tvMat 
         Height          =   1332
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   4332
         _ExtentX        =   7641
         _ExtentY        =   2350
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         Enabled         =   0   'False
      End
      Begin VB.CheckBox ckKriteriumMat 
         Caption         =   "Выбор групп материалов"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3252
      End
      Begin VB.CheckBox cbOborud 
         Caption         =   "Механика"
         Height          =   252
         Index           =   0
         Left            =   3000
         TabIndex        =   1
         Top             =   4200
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CheckBox cbOborud 
         Caption         =   "Сублимация"
         Height          =   252
         Index           =   2
         Left            =   1440
         TabIndex        =   2
         Top             =   4200
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.CheckBox cbOborud 
         Caption         =   "Лазер"
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   4200
         Value           =   1  'Checked
         Width           =   900
      End
      Begin MSComctlLib.TreeView tvRegion 
         Height          =   1332
         Left            =   360
         TabIndex        =   13
         Top             =   2400
         Width           =   4332
         _ExtentX        =   7641
         _ExtentY        =   2350
         _Version        =   393217
         Indentation     =   529
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         Enabled         =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Предустановленные фильтры"
      Height          =   1332
      Left            =   240
      TabIndex        =   4
      Top             =   7680
      Width           =   4932
      Begin VB.CommandButton cmFilterApply 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   252
         Left            =   3720
         Picture         =   "Analityc.frx":00F3
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Удалить Фильтр"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   252
      End
      Begin VB.ComboBox cbFilters 
         Height          =   288
         Left            =   120
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   360
         Width           =   3492
      End
      Begin VB.TextBox txFilterName 
         Height          =   288
         Left            =   120
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   840
         Width           =   3492
      End
      Begin VB.CommandButton cmFilterAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   252
         Left            =   3720
         Picture         =   "Analityc.frx":04DD
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Сохранить фильтр"
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   252
      End
      Begin VB.CommandButton cmFilterDelete 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   252
         Left            =   4080
         Picture         =   "Analityc.frx":08AC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Удалить Фильтр"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   252
      End
   End
End
Attribute VB_Name = "Analityc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbKlass As Recordset
Dim Node As Node

Private Sub cbOborud_Click(Index As Integer)
    checkDirtyFilterCommads
End Sub

Private Sub ckEndDate_Click()
    If ckEndDate.value = 1 Then
        tbEndDate.Enabled = True
    Else
        tbEndDate.Enabled = False
    End If
End Sub

Private Sub ckKriteriumMat_Click()
    checkDirtyFilterCommads
    If ckKriteriumMat.value = 1 Then
        tvMat.Enabled = True
    Else
        tvMat.Enabled = False
    End If
End Sub

Private Sub ckKriteriumRegion_Click()
    checkDirtyFilterCommads
    If ckKriteriumRegion.value = 1 Then
        tvRegion.Enabled = True
    Else
        tvRegion.Enabled = False
    End If
End Sub

Private Sub ckKriteriumOborud_Click()
Dim I As Integer

    checkDirtyFilterCommads
    If ckKriteriumOborud.value = 1 Then
        For I = 0 To 2
            cbOborud(I).Enabled = True
        Next I
    Else
        For I = 0 To 2
            cbOborud(I).Enabled = False
        Next I
    End If
End Sub


Private Sub ckStartDate_Click()
    If ckStartDate.value = 1 Then
        tbStartDate.Enabled = True
    Else
        tbStartDate.Enabled = False
    End If
End Sub

Private Sub cmExit_Click()
    Unload Me
End Sub

Private Sub checkDirtyFilterCommads()

End Sub


Sub loadRegions()
Dim key As String, pKey As String, K() As String, pK()  As String


    sql = "call wf_territory_catalog"
    Set tbKlass = myOpenRecordSet("##loadRegions", sql, dbOpenForwardOnly)
    If tbKlass Is Nothing Then myBase.Close: End
    
    If Not tbKlass.BOF Then
        tvRegion.Nodes.Clear
        While Not tbKlass.EOF
            key = "k" & tbKlass!RegionId
            If Not IsNull(tbKlass!territoryId) Then
                pKey = "k" & tbKlass!territoryId
                Set Node = tvRegion.Nodes.Add(pKey, tvwChild, key, tbKlass!region)
            Else
                Set Node = tvRegion.Nodes.Add(, , key, tbKlass!region)
            End If
            
            tbKlass.MoveNext
        Wend
    End If
    tbKlass.Close

End Sub


Sub loadKlass()
Dim key As String, pKey As String, K() As String, pK()  As String
    sql = "call wf_klass_catalog"
    Set tbKlass = myOpenRecordSet("##loadKlasss", sql, dbOpenForwardOnly)
    If tbKlass Is Nothing Then myBase.Close: End
    
    If Not tbKlass.BOF Then
        tvMat.Nodes.Clear
'        Set Node = tvMat.Nodes.Add(, , "k0", "Все регионы")
'        Node.Sorted = True
        While Not tbKlass.EOF
            key = "k" & tbKlass!KlassId
            If Not IsNull(tbKlass!parentKlassId) And (tbKlass!parentKlassId <> 0) Then
                pKey = "k" & tbKlass!parentKlassId
                Set Node = tvMat.Nodes.Add(pKey, tvwChild, key, tbKlass!KlassName)
            Else
                Set Node = tvMat.Nodes.Add(, , key, tbKlass!KlassName)
            End If
            
            tbKlass.MoveNext
        Wend
    End If
    tbKlass.Close

End Sub

Private Sub Form_Load()
    loadKlass
    loadRegions
End Sub

Private Sub tvMat_NodeCheck(ByVal Node As MSComctlLib.Node)
    checkDirtyFilterCommads
    If Not Node.Child Is Nothing Then
        setRecursiveNodeChecked Node.Child, Node.Checked
    End If
End Sub


Private Sub setRecursiveNodeChecked(ByRef root As Node, value As Boolean)
Dim NextNode As Node


    root.Checked = value
    Set NextNode = root.Next
    If Not NextNode Is Nothing Then
        setRecursiveNodeChecked NextNode, value
    End If
    If Not root.Child Is Nothing Then
        setRecursiveNodeChecked root.Child, value
    End If
End Sub

Private Sub setRecursiveParent(ByRef root As Node, value As Boolean)
    root.Checked = value
    If Not root.Parent Is Nothing Then
        setRecursiveParent root.Parent, value
    End If
End Sub



Private Sub tvRegion_NodeCheck(ByVal Node As MSComctlLib.Node)
    checkDirtyFilterCommads
    If Not Node.Child Is Nothing Then
        setRecursiveNodeChecked Node.Child, Node.Checked
    End If

End Sub
