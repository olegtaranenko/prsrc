VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Report 
   BackColor       =   &H8000000A&
   Caption         =   "�����"
   ClientHeight    =   8184
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11808
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8184
   ScaleWidth      =   11808
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbReserveTerm 
      Height          =   288
      ItemData        =   "Report.frx":0000
      Left            =   3600
      List            =   "Report.frx":0016
      TabIndex        =   11
      Text            =   "-- ���� ������� --"
      Top             =   7440
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.ComboBox cbAnormal 
      Height          =   288
      ItemData        =   "Report.frx":0083
      Left            =   3600
      List            =   "Report.frx":0099
      TabIndex        =   10
      Text            =   "-- ������� �������� --"
      Top             =   7800
      Width           =   2052
   End
   Begin VB.CheckBox ckSubtitle 
      BackColor       =   &H8000000A&
      Caption         =   "������������"
      Height          =   192
      Left            =   2040
      TabIndex        =   9
      Top             =   7860
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "������"
      Height          =   315
      Left            =   5700
      TabIndex        =   4
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "�����"
      Height          =   315
      Left            =   10980
      TabIndex        =   2
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "������ � Exel"
      Height          =   315
      Left            =   6600
      TabIndex        =   1
      Top             =   7800
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7455
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   13145
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label laRecCount 
      BackColor       =   &H8000000A&
      Caption         =   "����� �������:"
      Height          =   192
      Left            =   3780
      TabIndex        =   8
      Top             =   7800
      Width           =   1332
   End
   Begin VB.Label laCount 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label laRecSum 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   60
      TabIndex        =   6
      Top             =   7800
      Width           =   1100
   End
   Begin VB.Label laSum 
      BackColor       =   &H8000000A&
      Caption         =   "�����:"
      Height          =   192
      Left            =   1260
      TabIndex        =   5
      Top             =   7860
      Width           =   696
   End
   Begin VB.Label laHeader 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Regim As String, param1 As String, param2 As String, param3 As String
Public Caller As Form


Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����
Public nCols As Integer ' ����� ���-�� �������
Public mousRow As Long
Public mousCol As Long
Dim quantity As Long
Dim Cena()  As Single
Dim isLoad As Boolean


Const rrNumOrder = 1
Const rrDate = 2
Const rrFirm = 3
Const rrProduct = 4
Const rrMater = 5
Const rrReliz = 6

'��������� ��� whoReserved
Const rtNomZak = 1
Const rtReserv = 2
Const rtCeh = 3
Const rtData = 4
Const rtMen = 5
Const rtStatus = 6
Const rtFirma = 7
Const rtProduct = 8
Const rtZakazano = 9
Const rtOplacheno = 10

Const rzZatratName = 1
Const rzMainCosts = 2
Const rzAddCosts = 3

Const zdDate = 1
Const zdSumm = 2
Const zdProvodka = 3
Const zdAgent = 4
Const zdNazn = 5
Const zdUtochn = 6

Const sbnNomnom = 1
Const sbnNomnam = 2
Const sbnEdizm = 3
Const sbnPrice = 4
Const sbnSaled = 5
Const sbnSumma = 6

Const riMaxSklad = 1
Const riFactSklad = 2
Const riIncomplete = 3
Const riGoodsInWay = 4
Const riGoodsDebts = 5
Const riKonto = 6
Const riCash = 7
Const riCommonDebts = 8
Const riDebitor = 9
Const riKreditor = 10
Const riTotals = 11

Const CT_NUMBER = "numeric"
Const CT_DATE = "date"
Const CT_STRING = ""
Const CT_EMPTY = "empty"
Const CT_CUSTOM = "custom"
Const CT_SCHET = "schet"

Const TI_BIS_30 = 1
Const TI_1_2 = 2
Const TI_2_4 = 3
Const TI_4_6 = 4
Const TI_SEIT_6 = 5

Public setSortable As Boolean
    '��������� ���������� ������� ������������� ������� ������ �������

Public Sortable As Boolean
    '��������� ���������� - ����� ��� ��� ����� �������������.

Public Subtitle As Boolean
    ' true ���� � ������ ������������ ������������(������ ������������ � �.�.)

Public emptyColIndex As Integer
    '����������, �� ������ ������ ������� ������ �������������

Public groupIdColIndex As Integer
    '����������, �� ������ ������� ����� ���������� ����� ������ ��� ����������

Public subtitleColIndex As Integer
    '����� ������������ �� ������� - ��� ���� ������� ��� ����������
    
Public numSortSecondColIndex As Integer
    '������� ������ ����������, ���� ���������� �������� � ��� �������������
    
Public numSortThirdColIndex As Integer
    '������� ������� ����������, ���� ���������� �������� � ��� �������������
    
Dim aNormal As String
    '��� ������ ��������� ������ - ����� �������� ����������

Dim colType As String
    '���������� ��� ������� ����������.
    
Public whoRezervedIndex As Integer
    ' ��� ��������������� ������������ ���������� ������ ��������������
    





'otlaDwkdh - ���������� ����, ����� �����


'���� col <> "" - �����������, ����� �������
Sub laControl(Optional col As String = "")
    If col <> "" And Grid.col <> rrFirm Then GoTo AA
    If InStr(Regim, "tatistic") Then
       laSum.Caption = "���-�� ����:"
       If col = "" Then laRecSum.Caption = Grid.Rows - 1
    Else
AA:
       laSum.Caption = "�����:"
    End If
    
End Sub

Sub fitFormToGrid()
Dim i As Long, delta As Long

    i = 350 + (Grid.CellHeight + 17) * Grid.Rows
    delta = i - Grid.Height
    
    If Me.Height + delta > (Screen.Height - 900) Then _
        delta = (Screen.Height - 900) - Me.Height
    Me.Height = Me.Height + delta
    
    'Grid.Height = i
    delta = 0
    For i = 0 To Grid.Cols - 1
        delta = delta + Grid.ColWidth(i)
    Next i
    Me.Width = delta + 700

End Sub



Private Sub cbAnormal_Click()

    reloadSklad Grid, True

End Sub

Private Sub cbReserveTerm_Click()
    reloadSklad Grid, True
End Sub

Private Sub ckSubtitle_Click()
    reloadSklad Grid, False

End Sub

Private Sub reloadSklad(aGrid As MSFlexGrid, forceReload As Boolean)
Dim p_rowid As Integer


    If Not isLoad Then Exit Sub
    
    MousePointer = flexHourglass
    aGrid.Visible = False

    
    ' ����������� �������
    If forceReload Or ckSubtitle.value = 1 Then
        clearGrid Grid
        If Regim = "aReportDetail" Then
            ' ��������� ������
            If cbAnormal.ListIndex = 0 Or cbAnormal.ListIndex = -1 Then
                aNormal = "0"
            ElseIf cbAnormal.ListIndex = 1 Then
                aNormal = "null"
            Else
                aNormal = CStr(cbAnormal.ListIndex - 1)
            End If
        
            sqlRowDetail(1) = "call wf_nomenk_areport(" & aNormal & ")"
            p_rowid = CInt(param1)
            aReportDetail (p_rowid)
        ElseIf Regim = "reservedAll" Then
            Dim v_res_term As Integer
            v_res_term = cbReserveTerm.ListIndex
            If cbReserveTerm.ListIndex = -1 Then
                v_res_term = 0
            End If
            reservedAll v_res_term
        End If
    End If

    If ckSubtitle.value = 0 Then
        '������� ������������
        removeSubtitles aGrid
    End If
    
    MousePointer = flexDefault
    aGrid.Visible = True
    

End Sub

Public Sub removeSubtitles(aGrid As MSFlexGrid)
Dim i As Integer, maxrows As Integer
 
    i = aGrid.Rows - 1
    Do
        If aGrid.TextMatrix(i, emptyColIndex) = "" And Not (aGrid.Rows <= 2 And i = 1) Then
            aGrid.RemoveItem (i)
        End If
        i = i - 1
    Loop While (i > 0)
    
End Sub


Private Sub cmExel_Click()
GridToExcel Grid, laHeader.Caption
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmPrint_Click()
Me.PrintForm

End Sub

Private Sub Form_Load()
Dim prevDate As Date, prevNom As Long

oldHeight = Me.Height
oldWidth = Me.Width
Me.Caller.MousePointer = flexHourglass
cbAnormal.Visible = False

If Regim = "subDetail" Then
    laHeader.Caption = "����������� ���� " & param3 & "  �� �������� �� " & _
    Left$(param2, 8) & " �� ������ �" & gNzak
    subDetail
ElseIf Regim = "subDetailMat" Then
    laHeader.Caption = "����������� �����" & param3 & " �� ��������� �" & gNzak
    subDetail
ElseIf Regim = "aReport" Then
    laHeader.Caption = "����� � �� " & Format(Now(), "dd.mm.yy")
    aReport
ElseIf Regim = "aReportDetail" Then
    If param3 <> "" Then
        laHeader.Caption = "����������� ������ """ & param3 & """ �� ������ � "
    Else
        laHeader.Caption = aRowText(1) ' ��������� ������
        cbAnormal.Visible = True
    End If
    emptyColIndex = 1
    groupIdColIndex = 0
    subtitleColIndex = 2
    numSortSecondColIndex = 0 ' �� ������ ������
    numSortThirdColIndex = 2 ' �� �������� ������������
    Subtitle = arowSubtitle(param1)
    aReportDetail (param1)
    
ElseIf Regim = "whoRezerved" Then
    clearGrid Me.Grid
    quantity = 0
    whoRezerved whoRezervedIndex
    
ElseIf Regim = "reservedAll" Then
    cbReserveTerm.Visible = True
    reservedAll 0
ElseIf Regim = "" Then '����� ���������� - ������ ������������
    laHeader.Caption = "����������� ���� " & param2 & "(���������) � " & _
    param1 & "(����������) �� ����� �������� ������� ������������."
    relizDetail
ElseIf Regim = "relizStatistic" Then '����� ���������� - ������ ������������
    laHeader.Caption = "����������� ���� " & param2 & "(���������) � " & _
    param1 & "(����������) �� ������."
    relizDetail "statistic"
ElseIf Regim = "relizNomenk" Then
    laHeader.Caption = "���������� �� ��������� ������������ " & param2
    byNomenk "reliz"
ElseIf Regim = "uslug" Then '����� ���������� - ������ ������������
    laHeader.Caption = "����������� ����� " & param1 & "(������)" & _
    " �� ����� �������� ������� ������������."
    uslugDetail
ElseIf Regim = "uslugStatistic" Then '����� ���������� - ������ ������������
    laHeader.Caption = "����������� ����� " & param1 & "(������)" & _
    " �� ����� �������� ������� ������������."
    uslugDetail "statistic"
ElseIf Regim = "bay" Then '����� ���������� - ������ ������
    laHeader.Caption = "����������� ���� " & param2 & "(���������) � " & _
    param1 & "(����������) �� ����� �������� ��� ������ ������."
    relizDetailBay
ElseIf Regim = "bayNomenk" Then
    laHeader.Caption = "���������� �� ��������� ������������ " & param2
    byNomenk "saled"
ElseIf Regim = "bayStatistic" Then '����� ���������� - ������ ������
    laHeader.Caption = "����������� ���� " & param2 & "(���������) � " & _
    param1 & "(����������) �� ������."
    relizDetailBay "statistic"
ElseIf Regim = "mat" Then '����� ���������� - ��������� �� ��� ������
    laHeader.Caption = "����������� ����� " & _
    param1 & " �� ����� �������� ���������� �� ��� ������."
    relizDetailMat
ElseIf Regim = "venture" Then '����������� � ���������� �� ������������
    laHeader.Caption = "����������� ���� " & param2 & "(���������) � " & _
    param1 & "(����������) �� """ & Pribil.ventureId & """"
    ventureReport Pribil.ventureId
ElseIf Regim = "ventureZatrat" Then '���������� �� ������������
    Sortable = False
    laHeader.Caption = "����������� ���� �� ������ ������. ���������� """ & Pribil.ventureId & """"
    ventureZatrat Pribil.ventureId
ElseIf Regim = "ventureZatratDetail" Then '����������� �� ������������
    Sortable = True
    laHeader.Caption = "����������� ���� �� ������ ������ """ & Grid.TextMatrix(mousRow, rzZatratName) & """"
    ventureZatratDetail Pribil.ventureId, Grid.TextMatrix(Grid.row, 0)
End If

laControl

    If Subtitle Then
        ckSubtitle.Visible = True
        ckSubtitle.value = 1
    Else
        ckSubtitle.Visible = False
        ckSubtitle.value = 0
    End If
    

If InStr(Regim, "tatistic") Then
    trigger = False
    SortCol Grid, rrReliz, "numeric"
End If
fitFormToGrid
Me.Caller.MousePointer = flexDefault
isLoad = True
End Sub

Sub byNomenk(saled As String)
Dim restrictDate As Variant
Dim historyStart As Variant
Dim groupklassid As Integer

    
'select trim(n.cod + ' ' + nomname + ' ' + n.size) as name, s.quant, s.sm, s.nomnom, o.ord, k.klassname, k.klassid, n.cost, n.ed_izmer2
    Grid.FormatString = "|<����� ���.|<��������|�� ���.|>�-��|>����|>���� ����.|>�����|>����� ����."
    Grid.ColWidth(0) = 0
    Grid.ColWidth(sbnNomnom) = 1000
    Grid.ColWidth(sbnNomnam) = 4000
    Grid.ColWidth(sbnEdizm) = 600
    Grid.ColWidth(sbnPrice) = 700
    Grid.ColWidth(sbnSaled) = 1000
    Grid.ColWidth(sbnSumma) = 1000
    'Grid.ColWidth() =

    sql = "call wf_nomenk_" & saled & "(convert(datetime, " & startDate & "), convert(datetime, " & endDate & "))"
    'Debug.Print sql
    
    Set tbOrders = myOpenRecordSet("##vnt_det", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    quantity = 0 ': sum = 0
    groupklassid = 0
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF
            quantity = quantity + 1
            If groupklassid <> tbOrders!klassid Then
                Grid.AddItem ""
                Grid.AddItem Chr(9) & tbOrders!outtype & Chr(9) & tbOrders!klassName
                quantity = quantity + 2
                Grid.row = Grid.Rows - 1
                'quantity = quantity + 1
                Grid.col = sbnNomnom
                Grid.CellFontBold = True
                Grid.col = sbnNomnam
                Grid.CellFontBold = True
                groupklassid = tbOrders!klassid
            End If
            
            Grid.AddItem Chr(9) & tbOrders!nomnom _
                & Chr(9) & tbOrders!Name _
            & Chr(9) & tbOrders!ed_Izmer2 _
            & Chr(9) & Format(tbOrders!quant, "## ##0.00") _
            & Chr(9) & Format(tbOrders!cost, "## ##0.00") _
            & Chr(9) & Format(tbOrders!cenaTotal / tbOrders!quant, "## ##0.00") _
 _
            & Chr(9) & Format(tbOrders!cost * tbOrders!quant, "## ##0.00") _
            & Chr(9) & Format(tbOrders!cenaTotal, "## ##0.00")
            
            tbOrders.MoveNext
        Wend
    End If
    tbOrders.Close

'    Grid.row = quantity + 1
'    Grid.col = rzMainCosts: Grid.CellFontBold = True
'    Grid.col = rzAddCosts: Grid.CellFontBold = True
'    Grid.TextMatrix(quantity + 1, rzMainCosts) = Format(param2, "## ##0.00")
'    Grid.TextMatrix(quantity + 1, rzAddCosts) = Format(param1, "## ##0.00")

    

End Sub

Sub ventureZatratDetail(ventureId As Integer, id_shiz As String)
Dim sum As Single

    Grid.FormatString = "|����|>�����|��������|�����|����������|���������"
    Grid.ColWidth(0) = 0
    Grid.ColWidth(zdDate) = 850
    Grid.ColWidth(zdSumm) = 1000
    Grid.ColWidth(zdProvodka) = 1200
    Grid.ColWidth(zdAgent) = 2300
    Grid.ColWidth(zdNazn) = 3000
    Grid.ColWidth(zdUtochn) = 3000
    'Grid.ColWidth() =
    
    sql = "select xdate, uesumm, b.debit + '-' + b.subdebit + ' => ' + b.kredit + '-' + b.subkredit as provodka" _
    & " , k.name, p.pdescript as nazn, b.descript as utochn" _
    & " from ybook b" _
    & " join shiz s on s.id = b.id_shiz" _
    & " join ydebkreditor k on k.id = b.kreddebitor" _
    & " join yguidepurpose p on p.debit = b.debit and p.subdebit = b.subdebit and p.kredit = b.kredit and p.subkredit = b.subkredit and p.pid = b.purposeid" _
    & " where" _
    & " ventureid = " & ventureId & " and id_shiz = " & param1
 
    If Pribil.costsDateWhere <> "" Then
        sql = sql & " and " & Pribil.costsDateWhere
    End If
    sql = sql & " order by xdate, provodka, uesumm desc"

    'Debug.Print sql
    
    Set tbOrders = myOpenRecordSet("##vnt_det", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    quantity = 0: sum = 0
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF
            quantity = quantity + 1
            Grid.TextMatrix(quantity, zdDate) = tbOrders!xDate
            Grid.TextMatrix(quantity, zdSumm) = Format(tbOrders!uesumm, "## ##0.00")
            Grid.TextMatrix(quantity, zdProvodka) = tbOrders!provodka
            If Not IsNull(tbOrders!Name) Then Grid.TextMatrix(quantity, zdAgent) = tbOrders!Name
            If Not IsNull(tbOrders!nazn) Then Grid.TextMatrix(quantity, zdNazn) = tbOrders!nazn
            If Not IsNull(tbOrders!utochn) Then Grid.TextMatrix(quantity, zdUtochn) = tbOrders!utochn
            
            Grid.AddItem ""
            tbOrders.MoveNext
        Wend
    End If
    tbOrders.Close
    Grid.row = quantity + 1
    Grid.col = rzMainCosts: Grid.CellFontBold = True
    Grid.col = rzAddCosts: Grid.CellFontBold = True
    Grid.TextMatrix(quantity + 1, rzMainCosts) = Format(param2, "## ##0.00")
    Grid.TextMatrix(quantity + 1, rzAddCosts) = Format(param1, "## ##0.00")

End Sub


Sub ventureZatrat(ventureId As Integer)
Dim sum As Single

    Grid.FormatString = "|������������|>���.�������|>�����.�������"
    Grid.ColWidth(0) = 0
    Grid.ColWidth(rzZatratName) = 3600
    Grid.ColWidth(rzMainCosts) = 1500
    Grid.ColWidth(rzAddCosts) = 1500
    
    sql = "select sum(uesumm) as sm, ventureid, is_main_costs, s.nm as nm, b.id_shiz" _
    & " from ybook b" _
    & " join shiz s on s.id = b.id_shiz" _
    & " where " _
    & "     ventureid = " & ventureId _
    & " and s.is_main_costs is not null "
 
    If Pribil.costsDateWhere <> "" Then
        sql = sql & " and " & Pribil.costsDateWhere
    End If
    sql = sql & " group by ventureid, is_main_costs, nm, b.id_shiz" _
    & " order by nm"

    'Debug.Print sql
    
    Set tbOrders = myOpenRecordSet("##vnt_det", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    quantity = 0: sum = 0
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF
            quantity = quantity + 1
            Grid.TextMatrix(quantity, rzZatratName) = tbOrders!nm
            Grid.TextMatrix(quantity, 0) = tbOrders!id_shiz
            
            If tbOrders!is_main_costs = 1 Then
                Grid.TextMatrix(quantity, rzMainCosts) = Format(tbOrders!sm, "## ##0.00")
            Else
                Grid.TextMatrix(quantity, rzAddCosts) = Format(tbOrders!sm, "## ##0.00")
            End If
            Grid.AddItem ""
            tbOrders.MoveNext
        Wend
    End If
    tbOrders.Close
    Grid.row = quantity + 1
    Grid.col = rzMainCosts: Grid.CellFontBold = True
    Grid.col = rzAddCosts: Grid.CellFontBold = True
    Grid.TextMatrix(quantity + 1, rzMainCosts) = Format(param2, "## ##0.00")
    Grid.TextMatrix(quantity + 1, rzAddCosts) = Format(param1, "## ##0.00")
End Sub

Sub ventureReport(ventureId As Integer)
Dim sum As Single

    Grid.FormatString = "|�����|<����|<�����|<�����������|>�������|>����������"
    Grid.ColWidth(0) = 250
    Grid.ColWidth(rrNumOrder) = 885
    Grid.ColWidth(rrDate) = 765
    Grid.ColWidth(rrFirm) = 3855
    Grid.ColWidth(rrProduct) = 0
    Grid.ColWidth(rrMater) = 1005
    Grid.ColWidth(rrReliz) = 1005
    
    sql = "SELECT * from orderWallShip where ventureid = " & ventureId
    If Pribil.nDateWhere <> "" Then
        sql = sql & " and " & Pribil.nDateWhere
    End If
    sql = sql & " order by outdate"
    Set tbOrders = myOpenRecordSet("##vnt_det", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    quantity = 0: sum = 0
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF
            quantity = quantity + 1
            Grid.TextMatrix(quantity, rrNumOrder) = tbOrders!numorder
            Grid.TextMatrix(quantity, rrDate) = Format(tbOrders!outDate, "dd/mm/yy hh:nn:ss")
            Grid.TextMatrix(quantity, rrFirm) = tbOrders!firmName
            'Grid.TextMatrix(quantity, rrProduct) = tbOrders!Text
            If tbOrders!Type = 1 Then
                Grid.TextMatrix(quantity, 0) = "p"
            ElseIf tbOrders!Type = 2 Then
                Grid.TextMatrix(quantity, 0) = "n"
            ElseIf tbOrders!Type = 3 Then
                Grid.TextMatrix(quantity, 0) = "w"
            ElseIf tbOrders!Type = 4 Then
                Grid.TextMatrix(quantity, 0) = "u"
            ElseIf tbOrders!Type = 8 Then
                Grid.TextMatrix(quantity, 0) = "b"
            End If
            
            Grid.TextMatrix(quantity, rrMater) = Format(tbOrders!costTotal, "## ##0.00")
            Grid.TextMatrix(quantity, rrReliz) = Format(tbOrders!cenaTotal, "## ##0.00")
            Grid.AddItem ""
            tbOrders.MoveNext
        Wend
    End If
    tbOrders.Close

End Sub

' �������� ��� ����������������� ������������ ������ ������� � ��������������
' �� ���� ������ ������ �� ������ �
Sub reservedAll(p_term_index As Integer)

Dim groupklassid As Integer, rowStr As String
Dim p_days_start As Integer, p_days_end As Integer
    laHeader.Caption = "������ ����������������� ������������"
    Grid.FormatString = "#|<����� ���.|<��������|�� ���.|>>�-�� �����.|>����� �����."
    gridAutoWidth Grid
    
    If p_term_index = 0 Then
        p_days_start = 10000
        p_days_end = 0
    ElseIf p_term_index = 1 Then
        p_days_start = 30
        p_days_end = 0
    ElseIf p_term_index = 2 Then
        p_days_start = 60
        p_days_end = 30
    ElseIf p_term_index = 3 Then
        p_days_start = 120
        p_days_end = 60
    ElseIf p_term_index = 4 Then
        p_days_start = 180
        p_days_end = 120
    ElseIf p_term_index = 5 Then
        p_days_start = 10000
        p_days_end = 120
    End If
    
    sql = "call wf_nomenk_resered_all (" & p_days_start & ", " & p_days_end & ")"
    
    Set tbOrders = myOpenRecordSet("##reserved_all", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF
            quantity = quantity + 1
            If groupklassid <> tbOrders!klassid Then
                Grid.AddItem tbOrders!ord
                Grid.AddItem tbOrders!ord & Chr(9) & Chr(9) & tbOrders!klassName
                quantity = quantity + 2
                Grid.row = Grid.Rows - 1
                'quantity = quantity + 1
                Grid.col = sbnNomnom
                Grid.CellFontBold = True
                Grid.col = sbnNomnam
                Grid.CellFontBold = True
                groupklassid = tbOrders!klassid
            End If
            rowStr = tbOrders!ord _
                & Chr(9) & tbOrders!nomnom _
                & Chr(9) & tbOrders!Name _
                & Chr(9) & tbOrders!ed_Izmer2 _
                & Chr(9) & Format(tbOrders!quant, "## ##0.00") _
                & Chr(9) & Format(tbOrders!sm, "## ##0.00") _
                & Chr(9) & Format(tbOrders!sm, "## ##0.00")

            Grid.AddItem rowStr
            tbOrders.MoveNext
        Wend
    End If
    tbOrders.Close
    If quantity > 1 Then
        Grid.RemoveItem (1)
    End If
End Sub

Sub whoRezerved(p_term_index As Integer)
Dim groupklassid As Integer, rowStr As String
Dim p_days_start As Integer, p_days_end As Integer

    Grid.Visible = False
    
    laHeader.Caption = "������ �������, ���. ������������� ���-�� '" & gNomNom & "' [" & "]."
    
    Grid.FormatString = "|<� ������|>���-��|^��� |^���� |^ �|������" & _
    "|<�������� �����|<�������|>��������|>�����������"
    Grid.ColWidth(0) = 0
    'Grid.ColWidth(rtNomZak) =
    Grid.ColWidth(rtReserv) = 765
    Grid.ColWidth(rtCeh) = 765
    Grid.ColWidth(rtData) = 1600
    'Grid.ColWidth(rtMen) =
    Grid.ColWidth(rtStatus) = 930
    Grid.ColWidth(rtFirma) = 3270
    Grid.ColWidth(rtProduct) = 1950
    'Grid.ColWidth(rtZakazano) =
    Grid.ColWidth(rtOplacheno) = 810

    If p_term_index = 0 Then
        p_days_start = 10000
        p_days_end = 0
    ElseIf p_term_index = 1 Then
        p_days_start = 30
        p_days_end = 0
    ElseIf p_term_index = 2 Then
        p_days_start = 60
        p_days_end = 30
    ElseIf p_term_index = 3 Then
        p_days_start = 120
        p_days_end = 60
    ElseIf p_term_index = 4 Then
        p_days_start = 180
        p_days_end = 120
    ElseIf p_term_index = 5 Then
        p_days_start = 10000
        p_days_end = 120
    End If
    
    
    sql = "call wf_order_reserved ('" & gNomNom & "', " & p_days_start & ", " & p_days_end & ")"
    
    Set tbOrders = myOpenRecordSet("##350", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF

            quantity = quantity + 1
            Grid.TextMatrix(quantity, rtNomZak) = tbOrders!numorder
            Grid.TextMatrix(quantity, rtReserv) = Format(tbOrders!quant, "# ##0.00")
            If Not IsNull(tbOrders!ceh) Then _
                Grid.TextMatrix(quantity, rtCeh) = tbOrders!ceh
            
            Grid.TextMatrix(quantity, rtData) = tbOrders!date1
            If Not IsNull(tbOrders!Manager) Then _
                Grid.TextMatrix(quantity, rtMen) = tbOrders!Manager
            
            If Not IsNull(tbOrders!Status) Then _
                Grid.TextMatrix(quantity, rtStatus) = tbOrders!Status
            
            If Not IsNull(tbOrders!client) Then _
                Grid.TextMatrix(quantity, rtFirma) = tbOrders!client
            
            If Not IsNull(tbOrders!note) Then _
                Grid.TextMatrix(quantity, rtProduct) = tbOrders!note
            
            If Not IsNull(tbOrders!sm_zakazano) Then _
                Grid.TextMatrix(quantity, rtZakazano) = Format(tbOrders!sm_zakazano, "# ##0.00")
                
            If Not IsNull(tbOrders!sm_paid) Then _
                Grid.TextMatrix(quantity, rtOplacheno) = Format(tbOrders!sm_paid, "# ##0.00")
                
            Grid.AddItem ""
            tbOrders.MoveNext
        Wend
    End If
  tbOrders.Close

laCount.Caption = quantity
'laRecSum.Caption = Round(sum, 2)
If quantity > 0 Then
    Grid.RemoveItem quantity + 1
End If
trigger = False
Grid.Visible = True
Me.MousePointer = flexDefault

End Sub

Sub gridAutoWidth(pGrid As MSFlexGrid)
Dim i As Integer
Dim colHeaderText As String

    For i = 0 To pGrid.Cols - 1
        colHeaderText = pGrid.TextMatrix(0, i)
        If colHeaderText = "" Then
            pGrid.ColWidth(i) = 0
        ElseIf InStr(1, colHeaderText, "����� ���.") Then
            pGrid.ColWidth(i) = 1200
        ElseIf InStr(1, colHeaderText, "��������", vbTextCompare) Then
            pGrid.ColWidth(i) = 3500
        ElseIf InStr(1, colHeaderText, "���", vbTextCompare) Then
            pGrid.ColWidth(i) = 500
        ElseIf InStr(1, colHeaderText, "����", vbTextCompare) Then
            pGrid.ColWidth(i) = 600
        ElseIf InStr(1, colHeaderText, "�-��", vbTextCompare) Then
            pGrid.ColWidth(i) = 1000
        ElseIf InStr(1, colHeaderText, "�����", vbTextCompare) Then
            pGrid.ColWidth(i) = 1150
        ElseIf InStr(1, colHeaderText, "��������", vbTextCompare) Then
            pGrid.ColWidth(i) = 1300
        ElseIf InStr(1, colHeaderText, "����", vbTextCompare) Then
            pGrid.ColWidth(i) = 900
        ElseIf InStr(1, colHeaderText, "�����", vbTextCompare) Then
            pGrid.ColWidth(i) = 1200
        ElseIf InStr(1, colHeaderText, "������", vbTextCompare) Then
            pGrid.ColWidth(i) = 1000
        ElseIf InStr(1, colHeaderText, "�����", vbTextCompare) Then
            pGrid.ColWidth(i) = 1000
        ElseIf InStr(1, colHeaderText, "���.", vbTextCompare) Then
            pGrid.ColWidth(i) = 230
        ElseIf InStr(1, colHeaderText, "���", vbTextCompare) Then
            pGrid.ColWidth(i) = 430
        ElseIf InStr(1, colHeaderText, "-", vbTextCompare) Then
            pGrid.ColWidth(i) = 0
        ElseIf InStr(1, colHeaderText, "#", vbTextCompare) Then
            pGrid.ColWidth(i) = 300
        End If
    Next i
    

End Sub

Sub aReportDetail(p_rowid As Integer)
Dim rowStr As String
    Grid.TextMatrix(1, 0) = 0
    Sortable = aRowSortable(p_rowid)
    sql = sqlRowDetail(p_rowid)
    Grid.FormatString = rowFormatting(p_rowid)
    
    gridAutoWidth Grid
    
    'Debug.Print sql
    Set tbOrders = myOpenRecordSet("##vnt_det", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF
            quantity = quantity + 1
If p_rowid = 1 Then
        Dim groupklassid  As Integer, headerRowIndex As Long
        Dim groupQtyFact As Single, groupQtyMax As Single
        Dim groupFact As Single, groupMax As Single
            If groupklassid <> tbOrders!klassid Then
                Grid.AddItem tbOrders!klassOrdered
                Grid.AddItem tbOrders!klassOrdered & Chr(9) & Chr(9) & tbOrders!klassName

                quantity = quantity + 2
                Grid.row = Grid.Rows - 1
                'quantity = quantity + 1
                Dim i As Integer
                For i = sbnNomnom To 8
                    Grid.col = i
                    Grid.CellFontBold = True
                Next i
                groupklassid = tbOrders!klassid
                groupQtyFact = 0: groupQtyMax = 0
                groupFact = 0: groupMax = 0
                headerRowIndex = Grid.Rows - 1

            End If
            Dim qty_max As String
            If tbOrders!mark = "Used" Then
                qty_max = Format(tbOrders!qty_max, "## ##0.00")
                groupQtyMax = groupQtyMax + tbOrders!qty_max
            Else
                qty_max = "-0"
            End If
            
            groupQtyFact = groupQtyFact + tbOrders!qty_fact:
            groupFact = groupFact + tbOrders!qty_fact * tbOrders!cost
            groupMax = groupMax + tbOrders!qty_max * tbOrders!cost
            
            If headerRowIndex > 1 Then
                Grid.TextMatrix(headerRowIndex, 5) = Format(groupQtyFact, "## ##0.00")
                Grid.TextMatrix(headerRowIndex, 6) = Format(groupQtyMax, "## ##0.00")
                Grid.TextMatrix(headerRowIndex, 7) = Format(groupFact, "## ##0.00")
                Grid.TextMatrix(headerRowIndex, 8) = Format(groupMax, "## ##0.00")
            End If

            
            rowStr = tbOrders!klassOrdered _
                & Chr(9) & tbOrders!nomnom _
                & Chr(9) & tbOrders!Name _
                & Chr(9) & tbOrders!ed_Izmer2 _
                & Chr(9) & Format(tbOrders!cost, "## ##0.00") _
                & Chr(9) & Format(tbOrders!qty_fact, "## ##0.00") _
                & Chr(9) & qty_max _
                & Chr(9) & Format(tbOrders!qty_fact * tbOrders!cost, "## ##0.00") _
                & Chr(9) & Format(tbOrders!qty_max * tbOrders!cost, "## ##0.00")
            
ElseIf p_rowid = 2 Then
            rowStr = tbOrders!scope _
                & Chr(9) & tbOrders!numorder _
                & Chr(9) & tbOrders!firmName _
                & Chr(9) & tbOrders!ceh _
                & Chr(9) & tbOrders!Manag _
                & Chr(9) & Format(tbOrders!date2, "dd.mm.yy hh:nn") _
                & Chr(9) & Format(tbOrders!sm_processed, "## ##0.00") _


ElseIf p_rowid = 7 Then
            rowStr = tbOrders!Type _
                & Chr(9) & tbOrders!numorder _
                & Chr(9) & Format(tbOrders!Name, "## ##0.00") _
                & Chr(9) & Format(tbOrders!d, "## ##0.00") _
                & Chr(9) & Format(tbOrders!k, "## ##0.00") _

Else
            ' ������ �� ������ (��������)
                Dim v_purpose As String, v_cherez As String, v_note As String
                If IsNull(tbOrders!purpose) Then
                    v_purpose = ""
                Else
                    v_purpose = tbOrders!purpose
                End If
                
                If IsNull(tbOrders!cherez) Then
                    v_cherez = ""
                Else
                    v_cherez = tbOrders!cherez
                End If
                If IsNull(tbOrders!note) Then
                    v_note = ""
                Else
                    v_note = tbOrders!note
                End If
                rowStr = _
                     Chr(9) & tbOrders!provodka _
                    & Chr(9) & tbOrders!xDate _
                    & Chr(9) & Format(tbOrders!debit, "## ##0.00") _
                    & Chr(9) & Format(tbOrders!kredit, "## ##0.00") _
                    & Chr(9) & v_cherez _
                    & Chr(9) & v_purpose _
                    & Chr(9) & v_note _

End If
            
            Grid.AddItem rowStr
            tbOrders.MoveNext
        Wend
    End If
    tbOrders.Close
    
    If p_rowid = 1 And Not Subtitle Then
        removeSubtitles Grid
    End If

End Sub


Sub aReport()
Dim s As Single, k As Single, d As Single, sumD As Single, sumK As Single
Dim s2 As Single
Dim rowid As Integer

ReDim sqlRowDetail(11)
ReDim aRowText(11)
ReDim rowFormatting(11)
ReDim aRowSortable(11)
ReDim arowSubtitle(11)
    
Grid.FormatString = "||>�����       |>������       |>������       "
Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 4000
Grid.ColWidth(2) = 1360
Grid.ColWidth(3) = 1360
Grid.ColWidth(4) = 1360
'������� ����������� � ��������� ��� �� ����� ���-�� ���-� !!!

' ��������� ���� ��� ������� �� ������.
setStartEndDates Journal.tbStartDate, Journal.tbEndDate

'--------------------
sumD = 0: sumK = 0
rowid = 1
aRowText(rowid) = "C����: �����������(+)/������������(-)"
Grid.TextMatrix(1, 1) = aRowText(rowid)
sql = "SELECT Sum((if mark = 'Used' then zakup else nowOstatki/perlist endif) * cost) as max_sklad" _
    & " , Sum(cost * nowOstatki / perList) as fact_sklad" _
    & " FROM sGuideNomenk"
    
sqlRowDetail(rowid) = "call wf_nomenk_areport (" & aNormal & ")"
aRowSortable(rowid) = True
arowSubtitle(rowid) = True

'Debug.Print sql
rowFormatting(rowid) = "#|<����� ���.|<��������|�� ���.|>����|>�-�� ����|>�-�� ����|>�����.����|>����� ����."
byErrSqlGetValues "##387", sql, k, d
Grid.TextMatrix(1, 0) = ""
Grid.TextMatrix(1, 1) = aRowText(rowid)
Grid.TextMatrix(1, 2) = Format(Round(d, 2), "## ##0.00")
Grid.TextMatrix(1, 3) = Format(Round(k, 2), "## ##0.00")
Grid.TextMatrix(1, 4) = Format(Round(k - d, 4), "## ##0.00")
sumK = sumK + k
sumD = sumD + d
rowid = rowid + 1


'--------------------
' ����� ��������� � ��� ������������� ���-�� �� ���������� ������� !!!
    
sql = "select sum(sm_processed) as rest from orderBranProc"
   
byErrSqlGetValues "##386", sql, s

aRowSortable(rowid) = True
aRowText(rowid) = "������������� ������������"
rowFormatting(rowid) = "-|<� ������|<�������� �����|<���|���.|����� ����.��.|>����� �����."
sqlRowDetail(rowid) = " select r.scope, r.numorder, r.firmname, sm_processed" _
    & ", ceh, manag, date2" _
    & " from orderBranProc r" _
    & " order by numorder"

Grid.AddItem "0" & Chr(9) & aRowText(rowid) & Chr(9) & Format(Round(s, 2), "## ##0.00")
rowid = rowid + 1
sumD = sumD + s


'--------------------
s = Round(schetOstat("60"), 2)
If s < 0 Then k = -s: s = 0 Else k = 0

aRowSortable(rowid) = False
aRowText(rowid) = "������ � ���� / �������� ������"
rowFormatting(rowid) = "|��������|����|>�����|>������|����/�� ����|����������|���������"
sqlRowDetail(rowid) = "call wf_a_report_goods(" & startDate & ")"

Grid.AddItem rowid & Chr(9) & aRowText(rowid) & Chr(9) & Format(s, "## ##0.00") & Chr(9) & Format(k, "## ##0.00")
rowid = rowid + 1
sumD = sumD + s
sumK = sumK + k

'--------------------
s = schetOstat("51", "03")
s = s + schetOstat("51", "04")
s = Round(s + schetOstat("51", "05"), 2)
If s < 0 Then k = -s: s = 0 Else k = 0
aRowSortable(rowid) = False
aRowText(rowid) = "�/����"
rowFormatting(rowid) = rowFormatting(rowid - 1)
sqlRowDetail(rowid) = "call wf_a_report_konto(" & startDate & ")"
Grid.AddItem rowid & Chr(9) & aRowText(rowid) & Chr(9) & Format(s, "## ##0.00") & Chr(9) & Format(k, "## ##0.00")
rowid = rowid + 1
sumD = sumD + s
sumK = sumK + k

'--------------------
s = schetOstat("50", "01")
s = s + schetOstat("50", "02")
s = Round(s + schetOstat("50", "05"), 2)
If s < 0 Then k = -s: s = 0 Else k = 0
aRowSortable(rowid) = False
aRowText(rowid) = "�����"
rowFormatting(rowid) = rowFormatting(rowid - 1)
sqlRowDetail(rowid) = "call wf_a_report_kassa(" & startDate & ")"
Grid.AddItem rowid & Chr(9) & aRowText(rowid) & Chr(9) & s & Chr(9) & Format(k, "## ##0.00")
rowid = rowid + 1
sumD = sumD + s
sumK = sumK + k

'--------------------
s = Round(schetOstat("57"), 2)
If s < 0 Then k = -s: s = 0 Else k = 0
aRowSortable(rowid) = False
aRowText(rowid) = "�����"
rowFormatting(rowid) = rowFormatting(rowid - 1)
sqlRowDetail(rowid) = "call wf_a_report_debts(" & startDate & ")"
Grid.AddItem rowid & Chr(9) & aRowText(rowid) & Chr(9) & Format(s, "## ##0.00") & Chr(9) & Format(k, "## ##0.00")
rowid = rowid + 1
sumD = sumD + s
sumK = sumK + k


'--------------------
d = 0: k = 0
sql = "select sum(k) as k, sum(d) as d from vDebitorKreditor"
byErrSqlGetValues "##392", sql, k, d

aRowSortable(rowid) = True
aRowText(rowid) = "�������� / ���������"
sqlRowDetail(rowid) = "select type, numorder, name, k, d from vDebitorKreditor order by numorder"
rowFormatting(rowid) = "-|<� ������|<�������� �����|>����� ������|>����� �������"
Grid.AddItem rowid & Chr(9) & aRowText(rowid) & Chr(9) & Format(d, "## ##0.00") & Chr(9) & Format(k, "## ##0.00")
rowid = rowid + 1
sumD = sumD + d
sumK = sumK + k

'--------------------
aRowSortable(rowid) = False
aRowText(rowid) = "                                       �����:"
Grid.AddItem rowid & Chr(9) & aRowText(rowid) & _
Chr(9) & Format(sumD, "## ##0.00") & Chr(9) & Format(sumK, "## ##0.00") & Chr(9) & Format(Round(sumD - sumK, 2), "## ##0.00")
rowid = rowid + 1
Grid.row = Grid.Rows - 1
Grid.col = 1: Grid.CellFontBold = True
Grid.col = 2: Grid.CellFontBold = True
Grid.col = 3: Grid.CellFontBold = True
Grid.col = 4: Grid.CellFontBold = True

End Sub

Function schetOstat(schet As String, Optional subSchet As String)
Dim d As Single, k As Single

schetOstat = 0
If subSchet <> "" Then
    sql = "SELECT begDebit, begKredit From yGuideSchets" _
        & " where number = '" & schet & "' and subnumber = '" & subSchet & "'"
Else
    sql = "SELECT Sum(begDebit) AS Sum_begDebit, Sum(begKredit) AS Sum_begKredit " _
        & "From yGuideSchets GROUP BY number HAVING number = '" & schet & "'"
End If

If Not byErrSqlGetValues("W##389", sql, d, k) Then GoTo EN1 '$$4 � ����� ������ ����� �.� �� ����
schetOstat = d - k

d = 0: k = 0
sql = "SELECT Sum(UEsumm) AS Sum_UEsumm from yBook " & _
"WHERE Debit =" & schet & ""
If subSchet <> "" Then
    sql = sql & " and subdebit = '" & subSchet & "'"
End If

If Not byErrSqlGetValues("##390", sql, d) Then GoTo EN1

sql = "SELECT Sum(UEsumm) AS Sum_UEsumm from yBook " & _
"WHERE Kredit =" & schet & ""
If subSchet <> "" Then
    sql = sql & " and subkredit = '" & subSchet & "'"
End If
If Not byErrSqlGetValues("##391", sql, k) Then GoTo EN1
schetOstat = schetOstat + d - k

EN1:
End Function

Sub subDetail()
Dim str As String, i As Integer, delta As Integer, ed_izm As String
Dim str2 As String, str3 As String

Grid.FormatString = "|<�����|<��������|>���-�� � ����� |>���-�� �����|" & _
"<��.���������|>����|>�����|>����������"
Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 1500
Grid.ColWidth(2) = 3840
Grid.ColWidth(3) = 765
Grid.ColWidth(4) = 720
Grid.ColWidth(5) = 420
Grid.ColWidth(6) = 1080

strWhere = "20" & Mid$(param2, 7, 2) & "-" & Mid$(param2, 4, 2) & "-" & _
Left$(param2, 2) & Mid$(param2, 9)

If param1 = "p" Or param1 = "w" Then '����  ���.�������
  sql = "SELECT r.prId, r.prExt, " & _
  "r.quant, sGuideProducts.prName, " & _
  "sGuideProducts.prDescript, p.cenaEd " & _
  "FROM sGuideProducts INNER JOIN (xPredmetyByIzdelia p INNER JOIN xPredmetyByIzdeliaOut r ON (p.prExt = r.prExt) AND (p.prId = r.prId) AND (p.numOrder = r.numOrder)) ON sGuideProducts.prId = p.prId " & _
  "WHERE (((r.numOrder)=" & gNzak & ") AND " & _
  "((r.outDate) ='" & strWhere & "'));"
  
  Set tbProduct = myOpenRecordSet("##382", sql, dbOpenForwardOnly)
  If tbProduct Is Nothing Then Exit Sub

    
  While Not tbProduct.EOF
    Grid.AddItem _
        Chr(9) & tbProduct!prName _
        & Chr(9) & tbProduct!prDescript _
        & Chr(9) & "<--�������" _
        & Chr(9) & tbProduct!quant _
        & Chr(9) & "��." _
        & Chr(9) & "(" & Format(tbProduct!cenaEd, "## ##0.00") & ")" _
        & Chr(9) _
        & Chr(9) & Format(tbProduct!quant * tbProduct!cenaEd, "## ##0.00")
        
    Grid.row = Grid.Rows - 1: Grid.col = 1: Grid.CellFontBold = True
    Grid.col = 2: Grid.CellFontBold = True
    ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): ReDim QQ3(0): ReDim Cena(0)
    gProductId = tbProduct!prId
    prExt = tbProduct!prExt
    If Not productNomenkToNNQQ(1, tbProduct!quant) Then GoTo NXT
    For i = 1 To UBound(NN)
      sql = "SELECT nomName, ed_izmer, Size, cod From sGuideNomenk WHERE nomNom='" & NN(i) & "'"
      byErrSqlGetValues "##333", sql, str, ed_izm, str2, str3
      Grid.AddItem _
        Chr(9) & Format(NN(i), "## ##0.00") _
      & Chr(9) & str3 & " " & str & " " & str2 _
      & Chr(9) & Format(QQ(i), "## ##0.00") _
      & Chr(9) & Format(QQ2(i), "## ##0.00") _
      & Chr(9) & ed_izm _
      & Chr(9) & Format(Cena(i), "## ##0.00") _
      & Chr(9) & Format(QQ3(i), "## ##0.00")
    Next i
    Grid.AddItem ""
NXT:
    tbProduct.MoveNext
  Wend
  tbProduct.Close
End If

If param1 = "n" Or param1 = "w" Then
  sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.cost, " & _
  "sGuideNomenk.ed_izmer, sGuideNomenk.Size, sGuideNomenk.cod, " & _
  "sGuideNomenk.perList, xPredmetyByNomenk.cenaEd, xPredmetyByNomenkOut.quant " & _
  "FROM sGuideNomenk INNER JOIN (xPredmetyByNomenk INNER JOIN xPredmetyByNomenkOut ON (xPredmetyByNomenk.nomNom = xPredmetyByNomenkOut.nomNom) AND (xPredmetyByNomenk.numOrder = xPredmetyByNomenkOut.numOrder)) ON sGuideNomenk.nomNom = xPredmetyByNomenk.nomNom " & _
  "WHERE (((xPredmetyByNomenkOut.numOrder)=" & gNzak & ") AND " & _
  "((xPredmetyByNomenkOut.outDate) =  '" & strWhere & "'));"
  
  Set tbNomenk = myOpenRecordSet("##383", sql, dbOpenDynaset)
  If tbNomenk Is Nothing Then Exit Sub
  While Not tbNomenk.EOF
    Grid.AddItem _
          Chr(9) & tbNomenk!nomnom _
        & Chr(9) & tbNomenk!cod & " " & tbNomenk!nomName & " " & tbNomenk!Size _
        & Chr(9) & "<--������������" _
        & Chr(9) & tbNomenk!quant _
        & Chr(9) & tbNomenk!ed_izmer _
        & Chr(9) & Format(tbNomenk!cost / tbNomenk!perList, "## ##0.00") & " (" & Format(tbNomenk!cenaEd, "## ##0.00") & ")" _
        & Chr(9) & Format(tbNomenk!quant * tbNomenk!cost / tbNomenk!perList, "## ##0.00") _
        & Chr(9) & Format(tbNomenk!quant * tbNomenk!cenaEd, "## ##0.00")
    
    Grid.row = Grid.Rows - 1: Grid.col = 1: Grid.CellFontBold = True
    Grid.col = 2: Grid.CellFontBold = True
    Grid.AddItem ""
    tbNomenk.MoveNext
  Wend
  tbNomenk.Close
End If

If param1 = "b" Then
  Grid.ColWidth(3) = 0
'  sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.nomName, sGuideNomenk.cost, " & _
  "sGuideNomenk.ed_izmer2, sGuideNomenk.Size, sGuideNomenk.cod, " & _
  "sGuideNomenk.perList, sDMC.quant, sDMCrez.intQuant,  sDMCrez.numDoc " & _
  "FROM sGuideNomenk INNER JOIN ((BayOrders INNER JOIN sDocs ON BayOrders.numOrder = sDocs.numDoc) INNER JOIN (sDMC INNER JOIN sDMCrez ON sDMC.nomNom = sDMCrez.nomNom) ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc) AND (BayOrders.numOrder = sDMCrez.numDoc)) ON sGuideNomenk.nomNom = sDMC.nomNom " & _
  "WHERE (((sDMCrez.numDoc)=" & gNzak & ") AND " & _
  "((dateformat(sDocs.xDate, 'yyyy-mm-dd hh:nn:ss')) = '" & strWhere & "'));"
sql = "select po.outDate, o.numOrder, po.nomnom, r.intQuant AS cenaed, po.quant, n.cost as costEd, trim(n.cod + ' ' + nomname + ' ' + n.size) as name, n.ed_izmer2 " _
    & " from bayorders o" _
    & " join sDMCrez r on r.numDoc = o.numorder" _
    & " join baynomenkout po on po.numorder = o.numorder and po.nomnom = r.nomnom" _
    & " join sguidenomenk n on n.nomnom = po.nomnom" _
    & " WHERE r.numDoc = " & gNzak _
    & " AND dateformat(po.outDate, 'yyyy-mm-dd hh:nn:ss') = '" & strWhere & "'"
  
'  Debug.Print sql
  
  Set tbNomenk = myOpenRecordSet("##432", sql, dbOpenDynaset)
  If tbNomenk Is Nothing Then Exit Sub
  
  While Not tbNomenk.EOF '!!!
    Grid.AddItem _
        Chr(9) & tbNomenk!nomnom _
        & Chr(9) & tbNomenk!Name _
        & Chr(9) & "<--������������" _
        & Chr(9) & Format(tbNomenk!quant, "## ##0.00") _
        & Chr(9) & tbNomenk!ed_Izmer2 _
        & Chr(9) & Format(tbNomenk!costed, "## ##0.00") _
        & Chr(9) & Format(tbNomenk!quant * tbNomenk!costed, "## ##0.00") _
        & Chr(9) & Format(tbNomenk!quant * tbNomenk!cenaEd, "## ##0.00")
    tbNomenk.MoveNext
  Wend
  tbNomenk.Close
End If

If param1 = "m" Then
  Grid.ColWidth(3) = 0
  Grid.ColWidth(8) = 0
  sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.cod, sGuideNomenk.nomName, " & _
  "sGuideNomenk.Size, sDMC.quant, sGuideNomenk.cost, sGuideNomenk.perList, " & _
  "sGuideNomenk.ed_Izmer2 " & _
  "FROM sGuideNomenk INNER JOIN sDMC ON sGuideNomenk.nomNom = sDMC.nomNom " & _
  "GROUP BY sGuideNomenk.nomNom, sGuideNomenk.cod, sGuideNomenk.nomName, sGuideNomenk.Size, sDMC.quant, sGuideNomenk.cost, sGuideNomenk.perList, sGuideNomenk.ed_Izmer2, sDMC.numDoc " & _
  "HAVING (((sDMC.numDoc)=" & gNzak & "));"

  Set tbNomenk = myOpenRecordSet("##435", sql, dbOpenDynaset)
  If tbNomenk Is Nothing Then Exit Sub
  While Not tbNomenk.EOF '!!!
    Grid.AddItem _
          Chr(9) & tbNomenk!nomnom _
        & Chr(9) & tbNomenk!cod & " " & tbNomenk!nomName & " " & tbNomenk!Size _
        & Chr(9) _
        & Chr(9) & Format(tbNomenk!quant / tbNomenk!perList, "## ##0.00") _
        & Chr(9) & tbNomenk!ed_Izmer2 _
        & Chr(9) & Format(tbNomenk!cost, "## ##0.00") _
        & Chr(9) & Format(tbNomenk!quant * tbNomenk!cost / tbNomenk!perList, "## ##0.00")
    
    tbNomenk.MoveNext
  Wend
  tbNomenk.Close
End If

End Sub


Sub nomenkToNNQQ(pQuant As Single, eQuant As Single, prQuant As Single)
Dim j As Integer, leng As Integer

leng = UBound(NN)

    For j = 1 To leng
        If NN(j) = tbNomenk!nomnom Then
            QQ(j) = QQ(j) + pQuant * tbNomenk!quantity
            If eQuant > 0 Then _
                QQ2(j) = QQ2(j) + eQuant * tbNomenk!quantity
            If prQuant > 0 Then _
                QQ3(j) = QQ3(j) + prQuant * tbNomenk!quantity
            Exit Sub
        End If
    Next j
    leng = leng + 1
    ReDim Preserve NN(leng): NN(leng) = tbNomenk!nomnom
    ReDim Preserve Cena(leng): Cena(leng) = tbNomenk!cost
    ReDim Preserve QQ(leng): QQ(leng) = pQuant * tbNomenk!quantity
    ReDim Preserve QQ2(leng): QQ2(leng) = eQuant * tbNomenk!quantity
    ReDim Preserve QQ3(leng): QQ3(leng) = prQuant * tbNomenk!quantity
    

End Sub
'�����( �� ������-��) ��� ����������� ������������(���������� ������)
Function otgruzNomenk() As Single
Dim i As Integer
otgruzNomenk = 0

ReDim NN(0): ReDim QQ(0): ReDim QQ2(0): QQ2(0) = 0: ReDim QQ3(0)

'���-�� �������� � ������ �������
sql = "SELECT r.* " & _
"FROM xPredmetyByIzdeliaOut r INNER JOIN Orders ON r.numOrder = Orders.numOrder " & _
"WHERE (((Orders.StatusId)<6));"

Set tbProduct = myOpenRecordSet("##384", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Function

While Not tbProduct.EOF
    gNzak = tbProduct!numorder
    gProductId = tbProduct!prId
    prExt = tbProduct!prExt
    productNomenkToNNQQ 0, tbProduct!quant '2
    tbProduct.MoveNext
Wend
tbProduct.Close

'��������� ���-�� �������
sql = "SELECT sGuideNomenk.nomNom, sGuideNomenk.cost, sGuideNomenk.perList, " & _
"xPredmetyByNomenkOut.quant as quantity FROM (xPredmetyByNomenkOut INNER JOIN sGuideNomenk ON xPredmetyByNomenkOut.nomNom = sGuideNomenk.nomNom) INNER JOIN Orders ON xPredmetyByNomenkOut.numOrder = Orders.numOrder " & _
"WHERE (((Orders.StatusId)<6));"
Set tbNomenk = myOpenRecordSet("##385", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
While Not tbNomenk.EOF
  Dim str As String: str = tbNomenk!nomnom
  nomenkToNNQQ 0, 0, tbNomenk!cost / tbNomenk!perList '!!!
  tbNomenk.MoveNext
Wend
tbNomenk.Close

otgruzNomenk = 0
For i = 1 To UBound(NN)
    otgruzNomenk = otgruzNomenk + QQ3(i)
Next i

End Function

'� QQ3 ������������� ��������� ������-�� ���-�� � ��������� �� ���.��.���!!!
'����� ���-�� ���� ReDim NN(0): ReDim QQ(0): ReDim QQ2(0) : ReDim QQ3(0):QQ2(0)=0 - �� �.�����
Function productNomenkToNNQQ(pQuant As Single, eQuant As Single) As Boolean
Dim i As Integer, gr() As String

productNomenkToNNQQ = False
'ReDim NN(0): ReDim QQ(0)

'���������� ���-�� �������
sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup, " & _
"sGuideNomenk.cost, sGuideNomenk.perList " & _
"FROM sGuideNomenk INNER JOIN (sProducts INNER JOIN xVariantNomenc ON (sProducts.nomNom = xVariantNomenc.nomNom) AND (sProducts.ProductId = xVariantNomenc.prId)) ON sGuideNomenk.nomNom = sProducts.nomNom " & _
"WHERE (((xVariantNomenc.numOrder)=" & gNzak & ") AND (" & _
"(xVariantNomenc.prId)=" & gProductId & ") AND ((xVariantNomenc.prExt)=" & prExt & "));"
'MsgBox sql
Set tbNomenk = myOpenRecordSet("##192", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
ReDim gr(0): i = 0
While Not tbNomenk.EOF
    nomenkToNNQQ pQuant, eQuant, eQuant * tbNomenk!cost / tbNomenk!perList '!!!
    i = i + 1
    ReDim Preserve gr(i): gr(i) = tbNomenk!xgroup
    tbNomenk.MoveNext
Wend
tbNomenk.Close
    
'������������ ���-�� �������
'sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup " & _
"From sProducts WHERE (((sProducts.ProductId)=" & gProductId & "));"
sql = "SELECT sProducts.nomNom, sProducts.quantity, sProducts.xgroup, " & _
"sGuideNomenk.cost, sGuideNomenk.perList " & _
"FROM sGuideNomenk INNER JOIN sProducts ON sGuideNomenk.nomNom = sProducts.nomNom " & _
"WHERE (((sProducts.ProductId)=" & gProductId & "));"
'MsgBox sql
Set tbNomenk = myOpenRecordSet("##177", sql, dbOpenDynaset)
If tbNomenk Is Nothing Then Exit Function
While Not tbNomenk.EOF
    For i = 1 To UBound(gr) ' ���� ������ ������� �� ����� ���-��, �� ���
        If gr(i) = tbNomenk!xgroup Then GoTo NXT ' �����������, �.�. ��
    Next i                                      ' �� ������ � xVariantNomenc
    nomenkToNNQQ pQuant, eQuant, eQuant * tbNomenk!cost / tbNomenk!perList '!!!
NXT: tbNomenk.MoveNext
Wend
tbNomenk.Close

productNomenkToNNQQ = True
End Function
Sub relizDetailMat()
Dim r As Single ', typ As String, prevTyp As String

sql = "SELECT sDocs.numDoc, sDocs.xDate, sGuideSource.sourceName, " & _
"sGuideSource_1.sourceName AS destName, sDocs.Note, " & _
"Sum([sDMC].[quant]*[sGuideNomenk].[cost]/[sGuideNomenk].[perList]) AS cSum " & _
"FROM sGuideNomenk INNER JOIN (sGuideSource AS sGuideSource_1 INNER JOIN ((sGuideSource INNER JOIN sDocs ON sGuideSource.sourceId = sDocs.sourId) INNER JOIN sDMC ON (sDocs.numExt = sDMC.numExt) AND (sDocs.numDoc = sDMC.numDoc)) ON sGuideSource_1.sourceId = sDocs.destId) ON sGuideNomenk.nomNom = sDMC.nomNom " & _
"WHERE (" & Pribil.mDateWhere & ") " & _
"GROUP BY sDocs.numDoc, sDocs.xDate, sGuideSource.sourceName, " & _
"sGuideSource_1.sourceName, sDocs.Note ORDER BY sDocs.numDoc;"
'MsgBox sql
Set tbProduct = myOpenRecordSet("##434", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Sub
Grid.FormatString = "|<��������� �|<����|<������|<����|<����������|<���������"
Grid.ColWidth(0) = 0
Grid.ColWidth(rrNumOrder) = 930
Grid.ColWidth(rrDate) = 765
Grid.ColWidth(rrFirm) = 1300
Grid.ColWidth(rrProduct) = 1300
Grid.ColWidth(rrMater) = 1035
Grid.ColWidth(rrReliz) = 1035
quantity = 0
While Not tbProduct.EOF
    quantity = quantity + 1
    Grid.TextMatrix(quantity, 0) = "m"
    Grid.TextMatrix(quantity, rrNumOrder) = tbProduct!numDoc
    Grid.TextMatrix(quantity, rrDate) = Format(tbProduct!xDate, "dd/mm/yy hh:nn:ss")
    Grid.TextMatrix(quantity, rrFirm) = tbProduct!SourceName
    Grid.TextMatrix(quantity, rrProduct) = tbProduct!destName
    Grid.TextMatrix(quantity, rrMater) = tbProduct!note
    Grid.TextMatrix(quantity, rrReliz) = Format(tbProduct!cSum, "0.00") ' ����� ��� ����.����������� � ��������� �� �����
    Grid.AddItem ""
    tbProduct.MoveNext
Wend
tbProduct.Close

End Sub


Sub relizDetailBay(Optional statistic As String = "")

strWhere = Pribil.bDateWhere


If statistic = "" Then
    sql = "select d.numorder, d.outdate as xdate, f.name as name, d.costTotal, d.cenaTotal" _
        & " from orderWallShip d" _
        & " join bayorders o on o.numorder = d.numorder" _
        & " join bayguidefirms f on o.firmid = f.firmid" _
        & " WHERE d.type = 8 and " _
        & strWhere _
        & " ORDER BY o.numOrder, outDate"
Else
    sql = "select count(*) as numOrder, f.name as name, sum(d.costTotal) as costTotal, sum(d.cenaTotal) as cenaTotal" _
        & " from orderWallShip d" _
        & " join bayorders o on o.numorder = d.numorder" _
        & " join bayguidefirms f on o.firmid = f.firmid" _
        & " WHERE d.type = 8 and " _
        & strWhere _
        & " group BY f.Name order by cenaTotal desc"
End If

Debug.Print sql
Set tbProduct = myOpenRecordSet("##433", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Sub
If statistic = "" Then
    Grid.FormatString = "|<�����|<����|<�����||>���������|>����������"
    Grid.ColWidth(rrDate) = 765
Else
    Grid.FormatString = "|<����p��||<�����||>���������|>����������"
    Grid.ColWidth(rrDate) = 0
End If
Grid.ColWidth(0) = 0
Grid.ColWidth(rrNumOrder) = 885
Grid.ColWidth(rrFirm) = 3855
Grid.ColWidth(rrProduct) = 0
Grid.ColWidth(rrReliz) = 1005
Grid.ColWidth(rrMater) = 1005

quantity = 0
While Not tbProduct.EOF
  gNzak = tbProduct!numorder
  Grid.AddItem _
    Chr(9) & tbProduct!numorder _
    & Chr(9) _
    & Chr(9) & tbProduct!Name _
    & Chr(9) _
    & Chr(9) & Format(tbProduct!costTotal, "## ##0.00") _
    & Chr(9) & Format(tbProduct!cenaTotal, "## ##0.00") _
  
  quantity = quantity + 1
  If statistic = "" Then
    Grid.TextMatrix(quantity + 1, 0) = "b"
    Grid.TextMatrix(quantity + 1, rrDate) = Format(tbProduct!xDate, "dd/mm/yy hh:nn:ss")
  End If
  tbProduct.MoveNext
Wend
tbProduct.Close

End Sub

Sub uslugDetail(Optional statistic As String = "")
'Dim prevDate As Date, prevNom As Long, prevReliz As Single, prevMater As Single
Dim prevName As String, cSum As Single, prevNom As Long

'strWhere = Pribil.bDateWhere
'If strWhere <> "" Then strWhere = "HAVING ((" & strWhere & ")) "
If statistic = "" Then
    strWhere = " ORDER BY u.numOrder, u.outDate;"
Else
    strWhere = " ORDER BY GuideFirms.Name, u.numOrder;"
End If

sql = "SELECT u.numOrder, u.outDate, " & _
"u.quant, 1 AS cenaEd, GuideFirms.Name, Orders.Product " & _
"FROM GuideFirms INNER JOIN (Orders INNER JOIN xUslugOut u ON Orders.numOrder = u.numOrder) ON GuideFirms.FirmId = Orders.FirmId " & _
Pribil.uDateWhere & strWhere
'" ORDER BY xUslugOut.numOrder, xUslugOut.outDate;"

'MsgBox sql
Set tbProduct = myOpenRecordSet("##383", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Sub
If statistic = "" Then
    Grid.FormatString = "|�����|<����|<�����|<�������||>����������"
    Grid.ColWidth(rrDate) = 765
    Grid.ColWidth(rrProduct) = 2400
Else
    Grid.FormatString = "|�������||<�����|||>����������"
    Grid.ColWidth(rrDate) = 0
    Grid.ColWidth(rrProduct) = 0
End If
Grid.ColWidth(0) = 0
Grid.ColWidth(rrNumOrder) = 885
Grid.ColWidth(rrFirm) = 3855
Grid.ColWidth(rrReliz) = 1005
Grid.ColWidth(rrMater) = 0 '1005

'prevDate = 0: prevNom = 0: quantity = 0: prevReliz = 0: prevMater = 0
quantity = 0: prevName = "$$$$#####@@@@"
While Not tbProduct.EOF
  gNzak = tbProduct!numorder
  If statistic = "" Or tbProduct!Name <> prevName Then
  'If 1 = 1 Then
    quantity = quantity + 1
    If statistic = "" Then
        Grid.TextMatrix(quantity, rrNumOrder) = gNzak
    Else
        Grid.TextMatrix(quantity, rrNumOrder) = 1
    End If
    Grid.TextMatrix(quantity, rrDate) = Format(tbProduct!outDate, "dd/mm/yy hh:nn:ss")
    Grid.TextMatrix(quantity, rrFirm) = tbProduct!Name
    Grid.TextMatrix(quantity, rrProduct) = tbProduct!Product
    cSum = tbProduct!cenaEd * tbProduct!quant
    Grid.TextMatrix(quantity, rrReliz) = Format(cSum, "0.00")
    Grid.AddItem ""
  Else ' ������ ��� ����������
    If prevNom <> gNzak Then _
      Grid.TextMatrix(quantity, rrNumOrder) = 1 + Grid.TextMatrix(quantity, rrNumOrder)
    cSum = cSum + tbProduct!cenaEd * tbProduct!quant
    Grid.TextMatrix(quantity, rrReliz) = Format(cSum, "0.00") ' ����� ��� ����.����������� � ��������� �� �����
  End If
  prevName = tbProduct!Name
  prevNom = gNzak
  tbProduct.MoveNext
Wend
tbProduct.Close

End Sub

Sub relizDetail(Optional statistic As String = "")
Dim prevDate As Date, prevNom As Long, prevReliz As Single, prevMater As Single
Dim m As Single, r As Single, typ As String, prevTyp As String, prevName As String


If statistic = "" Then
'    strWhere = " ORDER BY r.numOrder, r.outDate;"
    strWhere = " ORDER BY 1, 2;"
Else
'    strWhere = " ORDER BY GuideFirms.Name, r.numOrder;"
    strWhere = " ORDER BY 8, 1;"
End If
sql = "SELECT r.numOrder, r.outDate, " & _
"r.prId, r.prExt, -1 AS costI, " & _
"r.quant, p.cenaEd, f.Name, o.Product " & _
"FROM (GuideFirms f INNER JOIN Orders o ON f.FirmId = o.FirmId) INNER JOIN (xPredmetyByIzdelia p INNER JOIN xPredmetyByIzdeliaOut r ON (p.prExt = r.prExt) AND (p.prId = r.prId) AND (p.numOrder = r.numOrder)) ON o.numOrder = p.numOrder " & _
Pribil.pDateWhere & _
" UNION ALL SELECT pno.numOrder, pno.outDate, " & _
"-1 AS prId, -1 AS prExt, n.cost/n.perList as costI, " & _
"pno.quant, pn.cenaEd, f.Name, o.Product " & _
"FROM sGuideNomenk n INNER JOIN ((GuideFirms f INNER JOIN Orders o ON f.FirmId = o.FirmId) " & _
"INNER JOIN (xPredmetyByNomenk pn INNER JOIN xPredmetyByNomenkOut pno ON " & _
"(pn.nomNom = pno.nomNom) AND (pn.numOrder = pno.numOrder)) ON o.numOrder = pn.numOrder) ON n.nomNom = pn.nomNom " & _
" where " & Pribil.nDateWhere & strWhere

'Debug.Print sql
Set tbProduct = myOpenRecordSet("##381", sql, dbOpenForwardOnly)
If tbProduct Is Nothing Then Exit Sub
Grid.FormatString = "|�����|<����|<�����|<�������|>���������|>����������"
If statistic = "" Then
    Grid.FormatString = "|�����|<����|<�����|<�������|>���������|>����������"
    Grid.ColWidth(0) = 300
    Grid.ColWidth(rrDate) = 765
    Grid.ColWidth(rrProduct) = 2400
Else
    Grid.FormatString = "|�������||<�����||>���������|>����������"
    Grid.ColWidth(0) = 0
    Grid.ColWidth(rrDate) = 0
    Grid.ColWidth(rrProduct) = 0
End If
Grid.ColWidth(rrNumOrder) = 885
Grid.ColWidth(rrFirm) = 3855
Grid.ColWidth(rrReliz) = 1005
Grid.ColWidth(rrMater) = 1005

prevDate = 0: prevNom = 0: quantity = 0: prevReliz = 0: prevMater = 0
While Not tbProduct.EOF
    
  gNzak = tbProduct!numorder
  If tbProduct!costI = -1 Then ' ������� �������
        gProductId = tbProduct!prId
        prExt = tbProduct!prExt
        m = Pribil.getProductNomenkSum * tbProduct!quant
        typ = "p"
        GoTo AA
'  ElseIf tbProduct!costI = -2 Then ' ������
'        m = 0: typ = "u"
'        GoTo AA
  Else ' ���.���-��
        typ = "n"
        m = tbProduct!costI * tbProduct!quant
AA:     r = tbProduct!cenaEd * tbProduct!quant
  End If
'If gNzak = "3102201" Then
'    gNzak = gNzak
'End If
  If statistic = "" Then
      bilo = (prevNom <> gNzak Or prevDate <> tbProduct!outDate)
  Else
      bilo = (prevName <> tbProduct!Name)
  End If
'  bilo = True
  If bilo Then
'  If prevNom <> gNzak Or prevDate <> tbProduct!outDate Then
    quantity = quantity + 1
    If statistic = "" Then
        Grid.TextMatrix(quantity, rrNumOrder) = gNzak
    Else
        Grid.TextMatrix(quantity, rrNumOrder) = 1 '���-�� �������
    End If
    Grid.TextMatrix(quantity, rrDate) = Format(tbProduct!outDate, "dd/mm/yy hh:nn:ss")
    Grid.TextMatrix(quantity, rrFirm) = tbProduct!Name
    Grid.TextMatrix(quantity, rrProduct) = tbProduct!Product
    Grid.AddItem ""
    prevReliz = r
    prevMater = m
    prevTyp = typ
  Else ' ��� ������ � ��� �� ������� � � ��� �� ����� - ���� �������� � ������� � ������.�����������
    If statistic <> "" And prevNom <> gNzak Then _
        Grid.TextMatrix(quantity, rrNumOrder) = 1 + Grid.TextMatrix(quantity, rrNumOrder)
    prevReliz = r + prevReliz
    prevMater = m + prevMater
    If typ <> prevTyp Then prevTyp = "w" '����� �� �.�."u"
  End If
  Grid.TextMatrix(quantity, 0) = prevTyp
  Grid.TextMatrix(quantity, rrReliz) = Format(prevReliz, "0.00")
  Grid.TextMatrix(quantity, rrMater) = Format(prevMater, "0.00")
  prevNom = gNzak: prevDate = tbProduct!outDate
  prevName = tbProduct!Name
  tbProduct.MoveNext
Wend
tbProduct.Close

End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

If WindowState = vbMinimized Then Exit Sub
On Error Resume Next

h = Me.Height - oldHeight
oldHeight = Me.Height
w = Me.Width - oldWidth
oldWidth = Me.Width
Grid.Height = Grid.Height + h
Grid.Width = Grid.Width + w

laSum.Top = laSum.Top + h
laRecSum.Top = laRecSum.Top + h
laHeader.Width = laHeader.Width + w
cmExel.Top = cmExel.Top + h
cmExit.Top = cmExit.Top + h
cmPrint.Top = cmPrint.Top + h
cbAnormal.Top = cbAnormal.Top + h
cbReserveTerm.Top = cbAnormal.Top


cmExit.Left = cmExit.Left + w
cmExel.Left = cmExit.Left - 50 - cmExel.Width
cmPrint.Left = cmExel.Left - 50 - cmPrint.Width
cbAnormal.Left = ckSubtitle.Left + ckSubtitle.Width + 50
cbReserveTerm.Left = cbAnormal.Left




ckSubtitle.Top = laSum.Top
laRecCount.Top = laSum.Top

End Sub

Private Function determineColType(colIndex As Long) As String
Dim rowIndex As Long, cellText As String
Dim asNumber As Integer, asString As Integer, asEmpty As Integer, asDate As Integer, asUnknown As Integer, asSchet As Integer

    For rowIndex = 2 To Grid.Rows
        cellText = Grid.TextMatrix(rowIndex - 1, colIndex)
        If IsNumeric(cellText) Then
            asNumber = asNumber + 1
        ElseIf IsDate(cellText) Then
            asDate = asDate + 1
        ElseIf cellText = "" Then
            asEmpty = asEmpty + 1
        ElseIf IsDate(cellText) Then
            asDate = asDate + 1
        ElseIf InStr(cellText, "=>") > 1 Then
            asSchet = asSchet + 1
        ElseIf Len(cellText) > 1 Then
            asString = asString + 1
        End If
    Next rowIndex
    
    Dim totalRows As Integer
    totalRows = Grid.Rows - asEmpty - 1
    If totalRows = 0 Then
        determineColType = CT_EMPTY
    ElseIf asNumber / totalRows > 0.9 Then
        determineColType = CT_NUMBER
    ElseIf asDate / totalRows > 0.9 Then
        determineColType = CT_DATE
    ElseIf asSchet / totalRows > 0.9 Then
        determineColType = CT_SCHET
    Else
        determineColType = CT_STRING
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Sortable = False
    Subtitle = False
    isLoad = False
    cbReserveTerm.Visible = False
End Sub

Private Sub Grid_Click()

    mousCol = Grid.MouseCol
    mousRow = Grid.MouseRow

'Grid.CellBackColor = Grid.BackColor
    If Sortable And mousRow = 0 Then
        Grid.MousePointer = flexHourglass
        colType = determineColType(mousCol)
        'MsgBox "Type of the determened column's type is: '" & colType & "'"
        
        Static ascSort As Integer, dscSort As Integer
        If Not ckSubtitle.value = 1 Then
            If colType = CT_STRING Then
                ascSort = 5
                dscSort = 6
            Else
                ascSort = 9
                dscSort = 9
            End If
            Grid.col = mousCol
            Grid.ColSel = mousCol
            If trigger Then
                Grid.Sort = dscSort
            Else
                Grid.Sort = ascSort
            End If
        Else
            ' ���� ����������� � �������������� - ������ ������ - ����������
                Grid.Sort = 9
        End If
        trigger = Not trigger

        Grid.MousePointer = flexDefault

        If colType <> CT_NUMBER Then
            'Grid.row = 1    ' ������ ����� ����� ���������
        End If
    End If

End Sub

Private Sub Grid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    Static sortAsc As Boolean
    Dim cell_1, cell_2 As String, ord1 As Integer, ord2 As Integer, empty1 As String, empty2 As String
    Dim date1, date2 As Date
    Dim num1, num2 As Single
    
    
    cell_1 = Grid.TextMatrix(Row1, mousCol)
    cell_2 = Grid.TextMatrix(Row2, mousCol)
    If ckSubtitle.value = 1 Then
        empty1 = Grid.TextMatrix(Row1, emptyColIndex)
        empty2 = Grid.TextMatrix(Row2, emptyColIndex)
        ord1 = CInt(Grid.TextMatrix(Row1, groupIdColIndex))
        ord2 = CInt(Grid.TextMatrix(Row2, groupIdColIndex))
        
        If ord1 <> ord2 Then
            Cmp = Sgn(ord1 - ord2)
        Else
            If empty1 = "" And empty2 = "" Then
                ' ����� ������������ �� �������
                If (Grid.TextMatrix(Row1, subtitleColIndex) = "") Then
                    Cmp = -1
                ElseIf Grid.TextMatrix(Row1, subtitleColIndex) <> "" Then
                    Cmp = 1
                Else
                    Cmp = 0
                End If
            ElseIf empty1 = "" Then
                Cmp = -1
            ElseIf empty2 = "" Then
                Cmp = 1
            Else
                If colType = CT_DATE Then
                    date1 = CDate(cell_1)
                    date2 = CDate(cell_2)
                    If date1 > date2 Then
                        Cmp = 1
                    ElseIf date1 < date2 Then
                        Cmp = -1
                    Else
                        Cmp = 0
                    End If
                ElseIf colType = CT_NUMBER Then
                    num1 = Round(CSng(cell_1), 5)
                    num2 = Round(CSng(cell_2), 5)
                    Cmp = Sgn(num1 - num2)
                ElseIf colType = CT_STRING Then
                    empty1 = Grid.TextMatrix(Row1, mousCol)
                    empty2 = Grid.TextMatrix(Row2, mousCol)
                    If empty1 > empty2 Then
                        Cmp = 1
                    ElseIf empty1 < empty2 Then
                        Cmp = -1
                    Else
                        Cmp = 0
                    End If
                End If
                If trigger Then Cmp = -Cmp
            End If
        End If
        
    Else
        If colType = CT_DATE Then
            
            If Not IsDate(cell_1) And Not IsDate(cell_2) Then
                Cmp = 0
                Exit Sub
            ElseIf Not IsDate(cell_1) Then
                Cmp = 1
                Exit Sub
            ElseIf Not IsDate(cell_2) Then
                Cmp = -1
                Exit Sub
            End If
            
            date1 = CDate(cell_1)
            date2 = CDate(cell_2)
            If date1 > date2 Then
                Cmp = 1
            ElseIf date1 < date2 Then
                Cmp = -1
            Else
                Cmp = 0
            End If
        ElseIf colType = CT_NUMBER Then
            If Not IsNumeric(cell_1) And Not IsNumeric(cell_2) Then
                Cmp = 0
                Exit Sub
            ElseIf Not IsNumeric(cell_1) Then
                Cmp = 1
                Exit Sub
            ElseIf Not IsNumeric(cell_2) Then
                Cmp = -1
                Exit Sub
            End If
            
            num1 = Round(CSng(cell_1), 2)
            num2 = Round(CSng(cell_2), 2)
            If num1 > num2 Then
                Cmp = 1
            ElseIf num1 < num2 Then
                Cmp = -1
            Else
                Cmp = secondarySorting(Row1, Row2)
            End If
        End If
    If trigger Then Cmp = -Cmp
    End If
End Sub

Private Function secondarySorting(ByVal Row1 As Long, ByVal Row2 As Long)
Dim str1 As String, str2 As String
Dim num1 As Single, num2 As Single
    str1 = Grid.TextMatrix(Row1, numSortSecondColIndex)
    str2 = Grid.TextMatrix(Row2, numSortSecondColIndex)
    If Not IsNumeric(str1) Or Not IsNumeric(str2) Then
        secondarySorting = 0
        Exit Function
    End If
    
    num1 = Round(CSng(Grid.TextMatrix(Row1, numSortSecondColIndex)), 2)
    num2 = Round(CSng(Grid.TextMatrix(Row2, numSortSecondColIndex)), 2)
    If num1 <> num2 Then
        secondarySorting = Sgn(num1 - num2) '������ �� �����������
    Else
        str1 = Grid.TextMatrix(Row1, numSortThirdColIndex)
        str2 = Grid.TextMatrix(Row2, numSortThirdColIndex)
        If (str1 < str2) Then
            secondarySorting = -1 '������ �� �����������
        ElseIf str1 > str2 Then
            secondarySorting = 1
        Else
            secondarySorting = 0
        End If
    End If

End Function
Private Sub Grid_DblClick()
    Dim str As String
    Dim Report2 As New Report
    Set Report2.Caller = Me
    

    If Grid.CellBackColor <> &H88FF88 Then Exit Sub
    
    gNzak = Grid.TextMatrix(mousRow, rrNumOrder)
    If Grid.TextMatrix(mousRow, 0) = "u" Then
        MsgBox "����� �" & gNzak & " �� �������� ���������, ������� ����� �� �� " & _
        "��������������!", , "��������������"
        Exit Sub
    End If
        
    Report2.param1 = Grid.TextMatrix(mousRow, 0) '
    Report2.param2 = Grid.TextMatrix(mousRow, rrDate)
    
    If Regim = "ventureZatrat" Then
        Report2.Regim = "ventureZatratDetail"
        If Grid.TextMatrix(mousRow, 2) <> "" Then
            Report2.param2 = Grid.TextMatrix(mousRow, 2)
        Else
            Report2.param2 = Grid.TextMatrix(mousRow, 3)
        End If
    ElseIf Regim = "mat" Then
    
        Report2.Regim = "subDetailMat"
        str = Grid.TextMatrix(mousRow, rrReliz)
    '    If MsgBox("�� ������ ���������� ������, ������� �������� ����� " & str _
        , vbDefaultButton2 Or vbYesNo, "����������?") = vbNo Then Exit Sub
    ElseIf Regim = "aReport" Then
        Report2.Regim = "aReportDetail"
        str = aRowText(mousRow)
        Report2.param1 = CStr(mousRow)
        Report2.param2 = CStr(mousRow)
        If mousRow = 1 Then
            Report2.cbAnormal.Visible = True
        End If
    ElseIf Regim = "reservedAll" Then
        Report2.Regim = "whoRezerved"
        Set Report2.Caller = Me
        Report2.Sortable = True
        gNomNom = Grid.TextMatrix(mousRow, 1)
        Report2.param1 = CStr(mousRow)
        Report2.param2 = CStr(mousRow)
        Report2.whoRezervedIndex = cbReserveTerm.ListIndex
        If Report2.whoRezervedIndex = -1 Then
            Report2.whoRezervedIndex = 0
        End If
    Else
        Report2.Regim = "subDetail"
        str = Grid.TextMatrix(mousRow, rrMater) & " � " & Grid.TextMatrix(mousRow, rrReliz)
    '    If MsgBox("�� ������ ���������� ������, ������� �������� ����� " & str _
        , vbDefaultButton2 Or vbYesNo, "����������?") = vbNo Then Exit Sub
    End If
    Report2.param3 = str
    
    Report2.Show vbModal
End Sub


Private Sub Grid_EnterCell()
If Not (Regim = "" _
    Or Regim = "bay" Or Regim = "mat" _
    Or Regim = "venture" Or Regim = "ventureZatrat" _
    Or Regim = "aReport" Or Regim = "reservedAll" _
) Then Exit Sub
mousRow = Grid.row
If mousRow = 0 Then Exit Sub
mousCol = Grid.col

If IsNull(sqlRowDetail) Then
    mousRow = mousRow
End If

Dim isReportDetail As Boolean

isReportDetail = False
If Regim = "aReport" Then
    If UBound(sqlRowDetail) > 0 Then
        If sqlRowDetail(mousRow) <> "" Then
            isReportDetail = True
        End If
    End If
End If

If (mousCol = rrReliz Or (mousCol = rrMater And Regim <> "mat") _
    Or (Regim = "ventureZatrat" And Grid.col >= rzMainCosts) _
    Or isReportDetail Or (Regim = "reservedAll" And Grid.TextMatrix(mousRow, 1) <> "") _
    ) _
Then
   Grid.CellBackColor = &H88FF88
Else
   Grid.CellBackColor = vbYellow
End If

End Sub

Private Sub Grid_LeaveCell()
Grid.CellBackColor = Grid.BackColor

End Sub


Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Grid.MouseRow = 0 And Shift = 2 Then
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
Else
'ElseIf Grid.col = rrReliz Or Grid.col = rrMater Then
    laControl "col"
    laRecSum.Caption = Round(sumInGridCol(Grid, Grid.col), 2)
End If
End Sub

