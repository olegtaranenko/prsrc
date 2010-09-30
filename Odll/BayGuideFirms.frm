VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BayGuideFirms 
   BackColor       =   &H8000000A&
   Caption         =   "Справочник фирм по продажам"
   ClientHeight    =   8184
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8184
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5472
      Index           =   1
      Left            =   360
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   6072
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   12
         Left            =   2400
         TabIndex        =   56
         Top             =   4440
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   11
         Left            =   2400
         TabIndex        =   54
         Top             =   4080
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   10
         Left            =   2400
         TabIndex        =   52
         Top             =   3720
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   9
         Left            =   2400
         TabIndex        =   50
         Top             =   3360
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   8
         Left            =   2400
         TabIndex        =   48
         Top             =   3000
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   7
         Left            =   2400
         TabIndex        =   46
         Top             =   2640
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   6
         Left            =   2400
         TabIndex        =   44
         Top             =   2280
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   5
         Left            =   2400
         TabIndex        =   42
         Top             =   1920
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   4
         Left            =   2400
         TabIndex        =   40
         Top             =   1560
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   3
         Left            =   2400
         TabIndex        =   38
         Top             =   1200
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   2
         Left            =   2400
         TabIndex        =   36
         Top             =   840
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   1
         Left            =   2400
         TabIndex        =   34
         Top             =   480
         Width           =   3492
      End
      Begin VB.TextBox tbInfo 
         Height          =   288
         Index           =   0
         Left            =   2400
         TabIndex        =   32
         Top             =   120
         Width           =   3492
      End
      Begin VB.CommandButton cmOk 
         Caption         =   "Ok"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   4920
         Width           =   915
      End
      Begin VB.CommandButton cmCancel 
         Caption         =   "Cancel"
         Height          =   315
         Index           =   1
         Left            =   5040
         TabIndex        =   29
         Top             =   4920
         Width           =   795
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Доставка:"
         Height          =   252
         Index           =   12
         Left            =   0
         TabIndex        =   55
         Tag             =   "Supplier"
         Top             =   4440
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Сайт:"
         Height          =   252
         Index           =   11
         Left            =   0
         TabIndex        =   53
         Tag             =   "HomePage"
         Top             =   4080
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "E-Mail:"
         Height          =   252
         Index           =   10
         Left            =   0
         TabIndex        =   51
         Tag             =   "Email"
         Top             =   3720
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Факс:"
         Height          =   252
         Index           =   9
         Left            =   0
         TabIndex        =   49
         Tag             =   "Fax"
         Top             =   3360
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Телефон:"
         Height          =   252
         Index           =   8
         Left            =   0
         TabIndex        =   47
         Tag             =   "Phone"
         Top             =   3000
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Контактное лицо:"
         Height          =   252
         Index           =   7
         Left            =   0
         TabIndex        =   45
         Tag             =   "FIO"
         Top             =   2640
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Система налогообложения:"
         Height          =   252
         Index           =   6
         Left            =   0
         TabIndex        =   43
         Tag             =   "Steuer"
         Top             =   2280
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Число сотрудников:"
         Height          =   252
         Index           =   5
         Left            =   0
         TabIndex        =   41
         Tag             =   "KvoSotr"
         Top             =   1920
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Руководитель:"
         Height          =   252
         Index           =   4
         Left            =   0
         TabIndex        =   39
         Tag             =   "Director"
         Top             =   1560
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Год основания:"
         Height          =   252
         Index           =   3
         Left            =   0
         TabIndex        =   37
         Tag             =   "InceptionYear"
         Top             =   1200
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Деятельность:"
         Height          =   252
         Index           =   2
         Left            =   0
         TabIndex        =   35
         Tag             =   "Delat"
         Top             =   840
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Фактический адрес:"
         Height          =   252
         Index           =   1
         Left            =   0
         TabIndex        =   33
         Tag             =   "Address"
         Top             =   480
         Width           =   2292
      End
      Begin VB.Label lbInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Город:"
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   31
         Tag             =   "City"
         Top             =   120
         Width           =   2292
      End
   End
   Begin VB.ListBox lbBayStatus 
      CausesValidation=   0   'False
      Height          =   240
      Left            =   1440
      TabIndex        =   27
      Top             =   2760
      Visible         =   0   'False
      Width           =   1872
   End
   Begin VB.ListBox lbTools 
      CausesValidation=   0   'False
      Height          =   240
      Left            =   1440
      TabIndex        =   26
      Top             =   2400
      Visible         =   0   'False
      Width           =   1872
   End
   Begin VB.ListBox lbOborud 
      Height          =   1584
      ItemData        =   "BayGuideFirms.frx":0000
      Left            =   600
      List            =   "BayGuideFirms.frx":001C
      TabIndex        =   25
      Top             =   2400
      Visible         =   0   'False
      Width           =   792
   End
   Begin VB.CommandButton cmLoad 
      Caption         =   "Обновить"
      Height          =   315
      Left            =   180
      TabIndex        =   24
      Top             =   7680
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3675
      Index           =   0
      Left            =   6720
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   4755
      Begin VB.TextBox tbType 
         Height          =   2835
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   240
         Width           =   4515
      End
      Begin VB.CommandButton cmCancel 
         Caption         =   "Cancel"
         Height          =   315
         Index           =   0
         Left            =   3600
         TabIndex        =   23
         Top             =   3240
         Width           =   795
      End
      Begin VB.CommandButton cmOk 
         Caption         =   "Ok"
         Height          =   315
         Index           =   0
         Left            =   660
         TabIndex        =   22
         Top             =   3240
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Вид  деятельности"
         Height          =   255
         Left            =   1620
         TabIndex        =   21
         Top             =   0
         Width           =   1515
      End
   End
   Begin VB.ListBox lbRegion 
      Height          =   4272
      ItemData        =   "BayGuideFirms.frx":005E
      Left            =   3000
      List            =   "BayGuideFirms.frx":0065
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   2592
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Exel"
      Height          =   315
      Left            =   6600
      TabIndex        =   17
      Top             =   7680
      Width           =   1215
   End
   Begin VB.ComboBox cbM 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   540
      Width           =   1635
   End
   Begin VB.TextBox tbInform 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   555
      Width           =   8835
   End
   Begin VB.Timer Timer1 
      Left            =   3540
      Top             =   7740
   End
   Begin VB.CommandButton cmAllOrders 
      Caption         =   "Отч.""Все заказы фирмы"
      Height          =   315
      Left            =   7020
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmNoClose 
      Caption         =   "Отчет ""Незакрытые заказы""  "
      Height          =   315
      Left            =   4380
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmNoCloseFiltr 
      Caption         =   "Фильтр""Незакрытые заказы"""
      Height          =   315
      Left            =   9180
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox lbM 
      Height          =   240
      Left            =   420
      TabIndex        =   12
      Top             =   1860
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2460
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox tbMobile 
      Height          =   285
      Left            =   600
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "tbMobile"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   7680
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmFind 
      Caption         =   "Поиск"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox tbFind 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmSel 
      Caption         =   "Выбрать"
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   11100
      TabIndex        =   10
      Top             =   7680
      Width           =   675
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6435
      Left            =   180
      TabIndex        =   5
      Top             =   1020
      Width           =   11715
      _ExtentX        =   20659
      _ExtentY        =   11345
      _Version        =   393216
      MergeCells      =   2
      AllowUserResizing=   1
      FormatString    =   " "
   End
   Begin VB.Label Label2 
      Caption         =   "Фильтр:"
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   600
      Width           =   735
   End
   Begin VB.Label laQuant 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   315
      Left            =   5520
      TabIndex        =   14
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label laHeadQ 
      Caption         =   "Число записей:"
      Height          =   195
      Left            =   4260
      TabIndex        =   13
      Top             =   7740
      Width           =   1215
   End
End
Attribute VB_Name = "BayGuideFirms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Regim As String ' режим окна
Public mousRow As Long    '
Public mousCol As Long    '
Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Dim quantity As Integer 'количество найденных фирм
Dim pos As Long 'поиция поиска

Const cEmpty = "пустой менеджер"





Private Sub cbM_Click()
loadGuide
On Error Resume Next ' требуется при вызове из Load
Grid.SetFocus
End Sub

Private Sub cmAdd_Click()
If Grid.TextMatrix(Grid.Rows - 1, bfId) <> "" Then Grid.AddItem ("")

'Grid.col = bfId ' чтобы наверняка было соб.EnterCell по Grid.col = bfNazwFirm
Grid.row = Grid.Rows - 1
Grid.col = bfNazwFirm 'название
Grid.SetFocus
textBoxInGridCell tbMobile, Grid
End Sub

Private Sub cmAllOrders_Click()
Me.MousePointer = flexHourglass
Report.Regim = "allFromFirms"
Report.Show vbModal
Grid.SetFocus
Me.MousePointer = flexDefault


End Sub

Private Sub cmCancel_Click(Index As Integer)
    Frame(Index).Visible = False
    tbInform.Text = Grid.TextMatrix(mousRow, mousCol)
    Grid.Enabled = True
    Grid.SetFocus
End Sub

Private Sub cmDel_Click()
Dim strId As String, I As Integer

If MsgBox("По кнопке <Да> вся информация по фирме будет безвозвратно " & _
"удалена из базы!", vbYesNo, "Удалить Фирму?") = vbNo Then Exit Sub

strId = Grid.TextMatrix(mousRow, bfId)
'sql = "SELECT FirmGuide.FirmId, FirmGuide.Name  From FirmGuide " & _
'"WHERE (((FirmGuide.FirmId)=" & strId & "));"
'Set tbFirms = myOpenRecordSet("##67", sql, dbOpenDynaset)
'If tbFirms Is Nothing Then Exit Sub
'On Error GoTo ERR1
'tbFirms.MoveFirst
'tbFirms.Delete
'tbFirms.Close

sql = "DELETE FROM BayGuideFirms WHERE FirmId = " & strId
I = myExecute("##67", sql, -198)
If I = -2 Then
    MsgBox "У этой фирмы есть заказы. Перед ее удалением необходимо " & _
    "в этих заказах выбать другую фирму, либо удалить эти заказы", , _
    "Удаление невозможно!"
    GoTo EN1
ElseIf I <> 0 Then
    GoTo EN1
End If

quantity = quantity - 1
If quantity = 0 Then
    clearGridRow Grid, mousRow
Else
    Grid.RemoveItem mousRow
End If
'Grid.SetFocus
'Exit Sub'

'ERR1:
'If Err = 3200 Then
'    MsgBox "У этой фирмы есть заказы. Перед ее удалением необходимо " & _
'    "в этих заказах выбать другую фирму, либо удалить эти заказы", , _
'    "Удаление невозможно!"
'Else
'    MsgBox Error, , "Ошибка 352-" & Err & ":  " '##352
'End If
EN1:
Grid.SetFocus
End Sub

Private Sub cmExel_Click()
GridToExcel Grid, "Справочник сторонних организаций (" & cbM.Text & ")"

End Sub

Private Sub cmExit_Click()
    Unload Me
End Sub

Public Sub cmFind_Click()
'Static pos As Long
pos = findExValInCol(Grid, tbFind.Text, bfNazwFirm, pos)
If pos > 0 Then
    cmSel.Enabled = True
    cmFind.Caption = "Далее"
    Grid.SetFocus
Else
    cmSel.Enabled = False
    tbFind.SetFocus
End If
pos = pos + 1

End Sub

Sub lbHide(Optional noGrid As String)
Dim I As Integer
tbMobile.Visible = False
lbM.Visible = False
lbRegion.Visible = False
lbOborud.Visible = False
lbTools.Visible = False
lbBayStatus.Visible = False
For I = 0 To Frame.UBound
    Frame(I).Visible = False
Next I
Grid.Enabled = True
If noGrid <> "" Then Exit Sub
Grid.SetFocus
Grid_EnterCell
End Sub

Private Function isBayInformationFull() As String
Dim I As Integer, res As String
Dim infoEmpty As Boolean, infoFull As Boolean
    infoEmpty = True
    infoFull = True
    For I = 0 To lbInfo.UBound
        Dim fieldName As String
        fieldName = lbInfo(I).Tag
        If IsNull(tbFirms(fieldName)) Or tbFirms(fieldName) = "" Then
            If infoFull Then
                infoFull = False
            End If
        Else
            If infoEmpty Then
                infoEmpty = False
            End If
        End If
    Next I
    If Not infoFull And Not infoEmpty Then
        isBayInformationFull = "Не полная"
    ElseIf infoFull Then
        isBayInformationFull = "Есть"
    Else
        isBayInformationFull = "Отсутствует"
    End If
End Function

Private Function makeBayInformation() As String
Dim I As Integer, res As String
    For I = 0 To lbInfo.UBound
        Dim fieldName As String
        fieldName = lbInfo(I).Tag
        If Not (fieldName = "FIO" Or fieldName = "Director") Then
            addBayInformation tbFirms(fieldName), res
        End If
    Next I
    makeBayInformation = res
End Function

Private Sub addBayInformation(newValue As Variant, ByRef buf As String)
'Dim Delim As String
Const Delim = " " & "+++" & " "

    Dim str As String
    If Not IsNull(newValue) Then
        str = newValue
    Else
        str = ""
    End If
    If str <> "" Then
        If "" <> buf Then
            buf = buf & Delim
        End If
        buf = buf & str
    End If

End Sub



Private Sub BayInfoInit()
Dim I As Integer
    For I = 0 To lbInfo.UBound
        Dim fieldName As String
        fieldName = lbInfo(I).Tag
        safeFieldInfo tbFirms(fieldName), I
    Next I
End Sub


Private Sub safeFieldInfo(newValue As Variant, Index As Integer)
    Dim str As String
    If Not IsNull(newValue) Then
        str = newValue
    Else
        str = ""
    End If
    tbInfo(Index).Text = str
    tbInfo(Index).Tag = str

End Sub



Sub loadGuide()
Dim I As Long, strWhere As String, str As String

Me.MousePointer = flexHourglass
Grid.Visible = False
clearGrid Grid
strWhere = Trim(tbFind.Text)
If Not strWhere = "" Then
    strWhere = "(f.Name) = '" & strWhere & "'"
End If
str = ""
If cbM.ListIndex > 0 Then str = "(f.ManagId) = " & _
    manId(cbM.ListIndex - 1)
If strWhere <> "" And str <> "" Then
    strWhere = strWhere & " AND " & str
Else
    strWhere = strWhere & str
End If
If strWhere <> "" Then strWhere = "Where " & strWhere
'MsgBox "strWhere = " & strWhere
quantity = 0

sql = "SELECT f.*, isnull(r.region, '') as region, isnull(u.oborud, '') as oborudName" _
& ", isnull(bs.BayStatusName, '') as BayStatusName, f.Director " _
& " FROM BayGuideFirms f " _
& " left join BayRegion r on r.regionid = f.regionid" _
& " left join GuideOborud u on u.oborudId = f.oborudId" _
& " left join GuideBayStatus bs on bs.bayStatusId = f.bayStatusId" _
& " " & strWhere _
& " ORDER BY Name"
'MsgBox sql
'Debug.Print sql
Set tbFirms = myOpenRecordSet("##15", sql, dbOpenForwardOnly)
If tbFirms Is Nothing Then GoTo EN1

If Not tbFirms.BOF Then
  While Not tbFirms.EOF
    If tbFirms!FirmId = 0 Then GoTo AA
    quantity = quantity + 1
    Grid.TextMatrix(quantity, bfId) = tbFirms!FirmId
    Grid.TextMatrix(quantity, bfNazwFirm) = tbFirms!Name
    Grid.TextMatrix(quantity, bfM) = Manag(tbFirms!ManagId)
    fieldToCol tbFirms!region, bfRegion
    Grid.TextMatrix(quantity, bfBayInfo) = isBayInformationFull
    fieldToCol tbFirms!tools, bfTools
    fieldToCol tbFirms!BayStatusName, bfBayStatus
    
    fieldToCol tbFirms!Director, bfDirector
    fieldToCol tbFirms!FIO, bfFIO
    fieldToCol tbFirms!year01, bf2001 '$$3
    fieldToCol tbFirms!year02, bf2002 '
    fieldToCol tbFirms!year03, bf2003 '
    fieldToCol tbFirms!year04, bf2004

    
    fieldToCol tbFirms!Type, bfType
    fieldToCol tbFirms!OborudName, bfOborud
    fieldToCol tbFirms!Sale, bfSale
    
    'fieldToCol tbFirms!Otklik, bfOtklik
    'fieldToCol tbFirms!Fax, bfFax
    'fieldToCol tbFirms!Email, bfEmail
    'fieldToCol tbFirms!Pass, bfPass
    'fieldToCol tbFirms!xLogin, bfLogin
    'fieldToCol tbFirms!Phone, bfTlf
    Grid.AddItem ("")
AA: tbFirms.MoveNext
  Wend
  If quantity > 0 Then Grid.RemoveItem (quantity + 1)
End If
tbFirms.Close
EN1:
Grid.Visible = True
laQuant.Caption = quantity
Me.MousePointer = flexDefault

End Sub

Sub fieldToCol(Field As Variant, col As Long)
    If Not IsNull(Field) Then Grid.TextMatrix(quantity, col) = Field
End Sub

Private Sub cmLoad_Click()
loadGuide
Grid.SetFocus
End Sub

Private Sub cmNoClose_Click()
Me.MousePointer = flexHourglass
Report.Regim = "FromFirms"
Report.Show vbModal
Grid.SetFocus
Me.MousePointer = flexDefault

End Sub

Private Sub cmNoCloseFiltr_Click()
Dim str As String
str = Grid.TextMatrix(mousRow, bfNazwFirm)
Unload Me
Orders.loadFirmOrders str

End Sub

Private Sub cmOk_Click(Index As Integer)
If Index = 0 Then
    If ValueToTableField("##353", "'" & tbType.Text & "'", "FirmGuide", "Type", "byFirmId") = 0 Then
        Grid.TextMatrix(mousRow, bfType) = tbType.Text
    End If
ElseIf Index = 1 Then
    Dim fieldName As String, changed As Boolean
    Dim I As Integer
    For I = 0 To lbInfo.UBound
        fieldName = lbInfo(I).Tag
        If tbInfo(I).Text <> tbInfo(I).Tag Then
            ValueToTableField "##353.1", "'" & tbInfo(I).Text & "'", "FirmGuide", fieldName, "byFirmId"
            changed = True
        End If
    Next I
    If changed Then
        Dim FirmId As String
        FirmId = Grid.TextMatrix(mousRow, bfId)
        sql = "select * from BayGuideFirms where firmId = " & FirmId
        Set tbFirms = myOpenRecordSet("##15", sql, dbOpenForwardOnly)
        If tbFirms Is Nothing Then End
        Grid.TextMatrix(mousRow, bfBayInfo) = isBayInformationFull
        tbFirms.Close
        
        tbInform.Text = Grid.TextMatrix(mousRow, bfBayInfo)
    End If
    
End If
Frame(Index).Visible = False
Grid.Enabled = True
Grid.SetFocus
End Sub

Private Sub cmSel_Click()
Dim sqlReq As String, FirmId As String, DNM As String

    Orders.Grid.Text = Grid.Text

    gNzak = Orders.Grid.TextMatrix(Orders.Grid.row, orNomZak)
    visits "-", "firm" ' уменьщаем посещения у старой фирмы, если она была
    FirmId = Grid.TextMatrix(Grid.row, bfId)
    ValueToTableField "##20", FirmId, "Orders", "FirmId"
    visits "+", "firm" ' увеличиваем посещения у новой фирмы

    DNM = Format(Now(), "dd.mm.yy hh:nn") & vbTab & Orders.cbM.Text & " " & gNzak ' именно vbTab
    On Error Resume Next ' в некот.ситуациях один из Open logFile дает Err: файл уже открыт
    Open logFile For Append As #2
    Print #2, DNM & " фирма=" & Grid.Text
    Close #2
    refreshTimestamp gNzak
    
Unload Me

'Orders.SetFocus
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then lbHide
End Sub

Private Sub Form_Load()
Dim I As Integer
quantity = 0
pos = 0
oldHeight = Me.Height
oldWidth = Me.Width

'Grid.FormatString = "|< Название  фирмы|^ M|<Оборудование|<Регион|Скидки в %|" & _
"<Контакт|<Отклик|200x|2001|2002|2004|<Конт.лицо|<Телефон|<Факс|<e-mail|<Вид деятельности" & _
"|<Логин|<Пароль|<Технологии|<Статус|>Id"

Grid.FormatString = "Id|< Название  фирмы|^ M|<Регион|<Информация о компании|<Оборудование|<Статус" & _
"|<Руководитель|<Конт.лицо|200x|2001|2002|2004" & _
"|<Вид деятельности|<Обор.старое|Скидки в %"

Grid.MergeRow(0) = True
Grid.ColWidth(0) = 0
Grid.ColWidth(bfNazwFirm) = 2730
Grid.ColWidth(bfM) = 330
Grid.ColWidth(bfRegion) = 1140
Grid.ColWidth(bfBayInfo) = 800
Grid.ColWidth(bfTools) = 1300
Grid.ColWidth(bfBayStatus) = 1300
Grid.ColWidth(bfFIO) = 1410
Grid.ColWidth(bfDirector) = 1410
            'bf2001
Grid.TextMatrix(0, bf2002) = Format(lastYear - 2, "0000") '$$3
Grid.TextMatrix(0, bf2003) = Format(lastYear - 1, "0000")
Grid.TextMatrix(0, bf2004) = Format(lastYear, "0000")

Grid.ColWidth(bfOborud) = 735
Grid.ColWidth(bfSale) = 480
Grid.ColWidth(bfType) = 1100

cbM.AddItem "все менеджеры"
lbM.AddItem "not"
For I = 0 To Orders.lbM.ListCount - 1
    If I < Orders.lbM.ListCount - 1 Then lbM.AddItem Orders.lbM.List(I)
    If Orders.lbM.List(I) = "" Then
        cbM.AddItem cEmpty
    Else
        cbM.AddItem "менеджер " & Orders.lbM.List(I)
    End If
Next I
cbM.ListIndex = 0
lbM.Height = lbM.Height + 195 * (lbM.ListCount - 1)

'Me.Caption = "Справочник сторонних организаций"

sql = "SELECT Region FROM BayRegion ORDER BY Region"
Set tbGuide = myOpenRecordSet("##349", sql, dbOpenForwardOnly)
'tbGuide.Index = "Region"
'If Not tbGuide Is Nothing Then
  While Not tbGuide.EOF
    lbRegion.AddItem tbGuide!region
    tbGuide.MoveNext
  Wend
  tbGuide.Close
'End If

If tbFind.Text <> "" Then cmFind.Enabled = True
If Regim = "fromContext" Then 'из Orders
    tbFind.Text = Orders.Grid.Text
    tbFind.SelStart = 0
    tbFind.SelLength = Len(BayGuideFirms.tbFind.Text)
    cmSel.Visible = True
    GoTo AA
ElseIf Regim = "fromFindFirm" Then
'    tbFind.Text = FindFirm.lb.Text
    tbFind.SelStart = 0
    tbFind.SelLength = Len(BayGuideFirms.tbFind.Text)
'    cmSel.Visible = FindFirm.cmSelect.Visible
ElseIf Regim = "fromMenu" Then 'по <F11> из Orders
    cmLoad.Visible = True
AA: If Orders.tbEnable.Visible Then
        cmNoClose.Visible = True
        cmAllOrders.Visible = True
        cmNoCloseFiltr.Visible = True
    End If
End If
cmAdd.Visible = True
cmDel.Visible = True

Set Table = myOpenRecordSet("##72", "bayRegion", dbOpenForwardOnly)
If Table Is Nothing Then myBase.Close: End

initListbox "select * from GuideTool ", Me.lbTools, "ToolId", "ToolName", 1

initListbox "select * from GuideBayStatus", Me.lbBayStatus, "bayStatusId", "BayStatusName"

'loadGuide не надо, т.к. при загрузке cbM_Click
Timer1.Interval = 100
Timer1.Enabled = True

End Sub

Private Sub Form_Resize()
Dim H As Integer, W As Integer

If Me.WindowState = vbMinimized Then Exit Sub
On Error Resume Next

lbHide "noGrid"
W = Me.Width - oldWidth
oldWidth = Me.Width
H = Me.Height - oldHeight
oldHeight = Me.Height

Grid.Height = Grid.Height + H
Grid.Width = Grid.Width + W
cmSel.Top = cmSel.Top + H
cmExit.Top = cmExit.Top + H
cmExit.Left = cmExit.Left + W
cmDel.Top = cmDel.Top + H
cmAdd.Top = cmAdd.Top + H
laHeadQ.Top = laHeadQ.Top + H
laQuant.Top = laQuant.Top + H
cmExel.Top = cmExel.Top + H
cmLoad.Top = cmLoad.Top + H
tbInform.Width = Grid.Width - (tbInform.Left - Grid.Left + 100)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Me.Visible = False
'Orders.Enabled = True
'Orders.SetFocus
End Sub


Private Sub Grid_Click()
mousCol = Grid.MouseCol
mousRow = Grid.MouseRow
If quantity = 0 Then Exit Sub
If Grid.MouseRow = 0 Then
    Grid.CellBackColor = Grid.BackColor
    If mousCol = bf2004 Then
        SortCol Grid, mousCol, "numeric"
    Else
        SortCol Grid, mousCol
    End If
    Grid.row = 1    ' только чтобы снять выделение
    Grid_EnterCell
End If

End Sub

Private Sub Grid_DblClick()
Dim I As Integer

If Grid.CellBackColor = vbYellow Then Exit Sub

gFirmId = Grid.TextMatrix(mousRow, bfId)

If mousCol = bfM Then
    listBoxInGridCell lbM, Grid, "select"
ElseIf mousCol = bfBayStatus Then
    listBoxInGridCell lbBayStatus, Grid, "select"
ElseIf mousCol = bfTools Then
    setFirmTools CInt(gFirmId)
    listBoxInGridCell lbTools, Grid, "select"
ElseIf mousCol = bfOborud Then
    listBoxInGridCell lbOborud, Grid, "select"
ElseIf mousCol = bfRegion Then ' Регион
'        For I = 0 To lbRegion.ListCount - 1 '
'            If Grid.Text = lbRegion.List(I) Then
''                noClick = True
'                lbRegion.ListIndex = I 'вызывает ложное onClick
''                noClick = False
'                Exit For
'            End If
'        Next I
'    lbRegion.Visible = True
'    lbRegion.ZOrder
'    lbRegion.SetFocus
'    Grid.Enabled = False 'иначе курсор по ней бегает
    listBoxInGridCell lbRegion, Grid, "select"
ElseIf mousCol = bfType Then
    tbType.Text = Grid.Text
'    tbType.SelLength = Len(tbType.Text)
    Frame(0).Visible = True
    Frame(0).ZOrder
    tbType.SetFocus
ElseIf mousCol = bfBayInfo Then
    Dim FirmId As String
    FirmId = Grid.TextMatrix(mousRow, bfId)
    Grid.Enabled = False
    sql = "select * from BayGuideFirms where firmId = " & FirmId
    Set tbFirms = myOpenRecordSet("##15", sql, dbOpenForwardOnly)
    If tbFirms Is Nothing Then End
    BayInfoInit
    
    If Grid.CellTop + Frame(1).Height < Grid.Height Then
        Frame(1).Top = Grid.CellTop + Grid.Top + Grid.CellHeight
    Else
        Frame(1).Top = Grid.CellTop + Grid.Top - Frame(1).Height '+ Grid.CellHeight
    End If
    Dim shiftRight As Long
    If Frame(1).Top < 0 Then
        Frame(1).Top = 0
        shiftRight = Grid.CellWidth
    End If
    
    Frame(1).Left = Grid.CellLeft + Grid.Left + shiftRight
    If Frame(1).Left + Frame(1).Width > Me.Width Then
        Frame(1).Left = Me.Width - Frame(1).Width
    End If
    tbFirms.Close
    
    Frame(1).Visible = True
    Frame(1).ZOrder
Else
    textBoxInGridCell tbMobile, Grid
'ElseIf cmSel.Enabled Then
'    If cmSel.Visible Then cmSel_Click
End If
End Sub


Private Sub Grid_EnterCell()
If noClick Then Exit Sub
mousRow = Grid.row
mousCol = Grid.col
'If quantity = 0 Or Regim = "F7" Then Exit Sub
'If Regim = "F7" Then Exit Sub
If mousCol = bfNazwFirm Then
    cmSel.Enabled = True
    cmDel.Enabled = True
Else
    cmSel.Enabled = False
    cmDel.Enabled = False
End If

If mousCol = bfNazwFirm Or mousCol = bfFIO Then
    tbInform.MaxLength = 80
    tbMobile.MaxLength = 80
ElseIf mousCol = bfType Then
    tbInform.MaxLength = 255
ElseIf mousCol = bfBayInfo Or mousCol = bfRegion Then
    tbInform.MaxLength = 0
'    tbMobile.MaxLength = 255
'ElseIf mousCol = bfTlf Or mousCol = bfSale Or _
mousCol = bfFax Or mousCol = bfEmail Then
    'tbInform.MaxLength = 50
    'tbMobile.MaxLength = 50
Else
    tbInform.MaxLength = 10
    tbMobile.MaxLength = 10
End If
tbInform.Text = Grid.TextMatrix(mousRow, mousCol)

If mousCol >= bf2004 And mousCol <> bfType Or Grid.TextMatrix(mousRow, bfId) = "" Then
    Grid.CellBackColor = vbYellow
    tbInform.Locked = True
Else
    Grid.CellBackColor = &H88FF88 ' бл.зел
'   Grid.CellBackColor = &H8888FF    ' бл.кр
    If mousCol = bfM Or mousCol = bfRegion Or mousCol = bfType Then
        tbInform.Locked = True
    Else
        tbInform.Locked = False
    End If
End If

End Sub

Private Sub Grid_GotFocus()
Grid_EnterCell
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
        Grid_DblClick
ElseIf KeyCode = vbKeyF4 Then
    'If mousCol = bfLogin Or mousCol = bfPass Then _
    '            textBoxInGridCell tbMobile, Grid
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If
End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then _
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
End Sub

Private Sub lbBayStatus_DblClick()
Dim str As String
str = lbBayStatus.ItemData(lbBayStatus.ListIndex)
If str = "0" Then
    str = "null"
End If
ValueToTableField "##lbBayStatus", str, "FirmGuide", "bayStatusId", "byFirmId"
Grid.TextMatrix(mousRow, bfBayStatus) = lbBayStatus.Text
lbHide
End Sub

Private Sub lbBayStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbBayStatus_DblClick
End Sub

Private Sub lbOborud_DblClick()

If lbOborud.Text = "" Then
    sql = "update FirmGuide set Oborudid = null " _
    & " where FirmGuide.firmId = " & gFirmId
Else
    sql = "update FirmGuide set Oborudid = u.Oborudid " _
    & " from GuideOborud u" _
    & " where FirmGuide.firmId = " & gFirmId _
    & " and u.Oborud = '" & lbOborud.Text & "'"
End If
myExecute "##354", sql

Grid.TextMatrix(mousRow, bfOborud) = lbOborud.Text

lbHide

End Sub

Private Sub lbOborud_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbOborud_DblClick
End Sub

Private Sub lbRegion_DblClick()
'Region

If lbRegion.Text = "" Then
    sql = "update FirmGuide set regionid = null " _
    & " where FirmGuide.firmId = " & gFirmId
Else
    sql = "update FirmGuide set regionid = r.regionid " _
    & " from bayregion r" _
    & " where FirmGuide.firmId = " & gFirmId _
    & " and r.region = '" & lbRegion.Text & "'"
End If

myExecute "##354", sql

Grid.TextMatrix(mousRow, bfRegion) = lbRegion.Text
lbHide
cmAdd.Enabled = True
cmExit.Enabled = True
cbM.Enabled = True
cmExel.Enabled = True
End Sub

Private Sub lbRegion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbRegion_DblClick

End Sub

Private Sub lbM_DblClick()
Dim str As String

If lbM.ListIndex = 0 Then
    str = "14" ' not
Else
    str = manId(lbM.ListIndex - 1)
End If
ValueToTableField "##355", str, "FirmGuide", "ManagId", "byFirmId"
Grid.TextMatrix(mousRow, bfM) = lbM.Text
lbHide

End Sub

Private Sub lbM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then lbM_DblClick
End Sub

Private Sub setFirmTools(FirmId As Integer)
Dim I As Integer, myToolId As Integer

For I = 0 To lbTools.ListCount - 1
    If Left(lbTools.List(I), 2) = "* " Then
        lbTools.List(I) = Mid(lbTools.List(I), 3)
    End If
Next I



sql = "select toolId from FirmTools where firmId = " & CStr(FirmId)
Set tbFirms = myOpenRecordSet("##15.0", sql, dbOpenForwardOnly)
If tbFirms Is Nothing Then End

If Not tbFirms.BOF Then
  While Not tbFirms.EOF
    myToolId = tbFirms!toolId
    For I = 0 To lbTools.ListCount - 1
        If lbTools.ItemData(I) = myToolId Then
            lbTools.List(I) = "* " & lbTools.List(I)
            Exit For
        End If
    Next I
    tbFirms.MoveNext
  Wend
End If
tbFirms.Close

End Sub
Private Sub lbTools_DblClick()
Dim myToolId As Integer, myTool As String
    myTool = lbTools.List(lbTools.ListIndex)
    myToolId = lbTools.ItemData(lbTools.ListIndex)
    If Left(myTool, 2) = "* " Then
        sql = "delete from FirmTools where ToolId = " & CStr(myToolId) & " and firmId =" & CStr(gFirmId)
    Else
        sql = "insert into FirmTools (ToolId, FirmId) select " & CStr(myToolId) & "," & CStr(gFirmId)
    End If
    myExecute "##15.1", sql
    
    sql = "select tools from bayGuideFirms where FirmId = " & gFirmId
    byErrSqlGetValues "##15.2", sql, myTool
    Grid.TextMatrix(mousRow, bfTools) = myTool
    lbHide
End Sub

Private Sub lbTools_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lbTools_DblClick
End Sub

Private Sub tbFind_Change()
If tbFind.Text <> "" Then cmFind.Enabled = True
End Sub

Private Sub tbFind_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmFind_Click
End Sub


Private Sub tbInform_GotFocus()
    tbInform.SelStart = Len(tbInform.Text)
End Sub

Private Sub tbInform_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tbMobile.Text = tbInform.Text
    tbMobile_KeyDown vbKeyReturn, 0
ElseIf KeyCode = vbKeyEscape Then
    Grid.SetFocus
End If

End Sub

Private Sub tbMobile_Change()
tbInform.Text = tbMobile.Text
End Sub

Private Sub tbMobile_DblClick()
lbHide
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String, I As Integer, strId As String
If KeyCode = vbKeyReturn Then
 str = Trim(tbMobile.Text)
 gFirmId = Grid.TextMatrix(mousRow, bfId)

 If mousCol = bfNazwFirm Then
   strId = Grid.TextMatrix(mousRow, bfId)
'   On Error GoTo ERR1
   If strId = "" Then
    wrkDefault.BeginTrans
    
    sql = "insert into FirmGuide (name, werkId) values ('" & str & "', 1); select @@identity;"
    I = byErrSqlGetValues("##50", sql, gFirmId)
    wrkDefault.CommitTrans
    
    
    Grid.TextMatrix(mousRow, bfId) = gFirmId
    quantity = quantity + 1
'    cep = True ' запускаем цепь посл.ввода еще 2х полей
'    cmAdd.Enabled = False
'    cmDel.Enabled = False
'    cmExit.Enabled = False
'    cbM.Enabled = False
'    cmExel.Enabled = False
    Grid.TextMatrix(mousRow, mousCol) = str
    Grid.TextMatrix(mousRow, bfM) = "not"
    lbHide
    Grid.col = bfM: mousCol = bfM
    listBoxInGridCell lbM, Grid
    Exit Sub
   Else
    sql = "UPDATE FirmGuide SET Name = '" & str & "' WHERE FirmId = " & strId
    I = myExecute("##356", sql, -196)
    If I <> 0 Then GoTo ERR0:
    
   End If
   On Error GoTo 0
' ElseIf mousCol = bfSale Then
'    ValueToTableField "##66", "'" & str & "'", "FirmGuide", "Sale", "byFirmId"
' ElseIf mousCol = bfKontakt Then
'    ValueToTableField "##66", "'" & str & "'", "FirmGuide", "Kontakt", "byFirmId"
 'ElseIf mousCol = bfOtklik Then
 '   ValueToTableField "##66", "'" & str & "'", "FirmGuide", "Otklik", "byFirmId"
 ElseIf mousCol = bfFIO Then
    ValueToTableField "##66", "'" & str & "'", "FirmGuide", "FIO", "byFirmId"
 ElseIf mousCol = bfDirector Then
    ValueToTableField "##66", "'" & str & "'", "FirmGuide", "Director", "byFirmId"
' ElseIf mousCol = bfFax Then
'    ValueToTableField "##66", "'" & str & "'", "FirmGuide", "Fax", "byFirmId"
' ElseIf mousCol = bfEmail Then
'    ValueToTableField "##66", "'" & str & "'", "FirmGuide", "Email", "byFirmId"
' ElseIf mousCol = bfLogin Then
'    If str <> "" And str <> Grid.TextMatrix(mousRow, bfLogin) Then
'        If existValueInTableFielf(str, "FirmGuide", "xLogin") Then
'            MsgBox "Такой логин уже есть.", , "Недопустимое значение"
'            tbMobile.SelStart = Len(tbMobile.Text)
'            tbMobile.SetFocus
'            Exit Sub
'        End If
'    End If
'    ValueToTableField "##66", "'" & str & "'", "FirmGuide", "xLogin", "byFirmId"
' ElseIf mousCol = bfPass Then
'    ValueToTableField "##66", "'" & str & "'", "FirmGuide", "Pass", "byFirmId"
End If
'GoTo AA
'ElseIf KeyCode = vbKeyEscape Then
AA:
 Grid.TextMatrix(mousRow, mousCol) = str
 lbHide
End If
Exit Sub

ERR0: If I = -2 Then
        MsgBox "Такая фирма уже есть", , "Ошибка!"
        tbMobile.SetFocus
      Else
ERR1:   lbHide
      End If
End Sub

Private Sub tbType_Change()
tbInform.Text = tbType.Text
End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False
 If Regim = "fromMenu" Then  'по F11 из Orders
    tbFind.SetFocus
 Else ' из контекстного меню
    cmFind.Caption = "Поиск"
   'cmFind_Click
    If findValInCol(Grid, tbFind.Text, bfNazwFirm) Then
        Grid.SetFocus
    Else
        If tbFind.Visible Then tbFind.SetFocus
    End If
 End If

End Sub
