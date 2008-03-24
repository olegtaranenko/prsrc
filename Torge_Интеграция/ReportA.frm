VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ReportA 
   BackColor       =   &H8000000A&
   Caption         =   "Отчет"
   ClientHeight    =   8184
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11808
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8184
   ScaleWidth      =   11808
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmManag 
      Caption         =   "Применить"
      Height          =   315
      Left            =   2640
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmMonthRight 
      Caption         =   ">"
      Height          =   252
      Left            =   6720
      TabIndex        =   11
      Top             =   120
      Width           =   372
   End
   Begin VB.CommandButton cmMonthLeft 
      Caption         =   "<"
      Height          =   252
      Left            =   5520
      TabIndex        =   10
      Top             =   120
      Width           =   372
   End
   Begin VB.CommandButton cmDayRight 
      Caption         =   ">"
      Height          =   252
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   372
   End
   Begin VB.CommandButton cmDayLeft 
      Caption         =   "<"
      Height          =   252
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   372
   End
   Begin VB.TextBox tbStartDate 
      Height          =   285
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   5
      Top             =   120
      Width           =   912
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "Печать"
      Height          =   315
      Left            =   5700
      TabIndex        =   3
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   10980
      TabIndex        =   2
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "Печать в Exel"
      Height          =   315
      Left            =   6600
      TabIndex        =   1
      Top             =   7800
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7212
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   11652
      _ExtentX        =   20553
      _ExtentY        =   12721
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Месяц"
      Height          =   192
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   552
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "День"
      Height          =   192
      Left            =   4440
      TabIndex        =   9
      Top             =   120
      Width           =   432
   End
   Begin VB.Label laPeriod 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Отчет на дату"
      Height          =   192
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   1272
   End
   Begin VB.Label laCount 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   6600
      TabIndex        =   4
      Top             =   6720
      Width           =   495
   End
End
Attribute VB_Name = "ReportA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Caller As Form


Dim oldHeight As Integer, oldWidth As Integer ' нач размер формы
Public nCols As Integer ' общее кол-во колонок
Public mousRow As Long
Public mousCol As Long
Dim quantity As Long
Dim Cena()  As Single
Dim isLoad As Boolean
Dim isToday As Boolean


'otlaDwkdh - отладочная база, дебаг режим


'если col <> "" - проверяется, какая колонка
Sub laControl(Optional col As String = "")
    
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


Private Sub cmDayLeft_Click()
    Dim dt As Date, effectiveDt As Date
    dt = dateBasic2Sybase(tbStartDate.Text)
    dt = DateAdd("d", -1, dt)
    effectiveDt = dt
    showDay dt, -1

End Sub

Private Sub cmDayRight_Click()
    Dim dt As Date
    dt = dateBasic2Sybase(tbStartDate.Text)
    dt = DateAdd("d", 1, dt)
    showDay dt, 1

End Sub

Private Sub cmExel_Click()
GridToExcel Grid
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub


Private Sub showDay(dt As Date, Optional toFuture As Integer = 0)
Dim searchNotEmptyDay As Date
Dim doRefresh As Boolean

    MousePointer = flexHourglass
    isLoad = False
    
    sql = "call wf_nearest_day('" & Format(dt, "yyyymmdd") & "', " & toFuture & ")"

    If byErrSqlGetValues("## wf_nearest_day", sql, searchNotEmptyDay) Then
        If Format(searchNotEmptyDay, "yyyymmdd") <> "20000101" Then
            doRefresh = True
            dt = searchNotEmptyDay
        End If
    End If
    
    If doRefresh Then
        If Format(dt, "yyyymmdd") = Format(Now(), "yyyymmdd") Then
            isToday = True
        Else
            isToday = False
        End If
        tbStartDate.Text = Format(dt, "dd.mm.yy")
        
        clearGrid Grid
        aReport dt
    End If
    
    isLoad = True
    MousePointer = flexDefault

End Sub

Private Sub cmManag_Click()
    Dim dt As Date
    dt = dateBasic2Sybase(tbStartDate.Text)
    showDay dt, 0

End Sub

Private Sub cmMonthLeft_Click()
    Dim dt As Date
    dt = dateBasic2Sybase(tbStartDate.Text)
    dt = DateAdd("m", -1, dt)
    showDay dt, -1

End Sub

Private Sub cmMonthRight_Click()
    Dim dt As Date
    dt = dateBasic2Sybase(tbStartDate.Text)
    dt = DateAdd("m", 1, dt)
    showDay dt, 1
End Sub

Private Sub cmPrint_Click()
Me.PrintForm

End Sub

Private Sub Form_Load()
Dim prevDate As Date, prevNom As Long

Me.Caller.MousePointer = flexHourglass

oldHeight = Me.Height
oldWidth = Me.Width
Grid.FormatString = "||>Актив       |>Пассив       |>Баланс       "
Grid.Cols = Grid.Cols + 6

Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 4000
Grid.ColWidth(2) = 1360
Grid.ColWidth(3) = 1360
Grid.ColWidth(4) = 1360
Grid.ColWidth(5) = 0
Grid.ColWidth(6) = 0
Grid.ColWidth(7) = 0
Grid.ColWidth(8) = 0
Grid.ColWidth(9) = 0
Grid.ColWidth(10) = 0


    isToday = True
    showDay Now()

fitFormToGrid
Me.Caller.MousePointer = flexDefault
isLoad = True
End Sub


Sub aReport(aDay As Date)
Dim s As Single, k As Single, d As Single, sumD As Single, sumK As Single
Dim s2 As Single
Dim rowid As Integer
Dim dateEnd As Date, dateStart As Date



dateStart = DateAdd("m", -1, aDay)


sql = "call wf_areport_retrieve('" & Format(aDay, "yyyymmdd") & "', '" & Format(dateStart, "yyyymmdd") & "')"
    Set tbOrders = myOpenRecordSet("##vnt_det", sql, dbOpenForwardOnly)
    If tbOrders Is Nothing Then Exit Sub
    quantity = 0 ': sum = 0
    If Not tbOrders.BOF Then
        While Not tbOrders.EOF
            quantity = quantity + 1
            
            Dim balans As Integer, BalansSum As String
            
            If Not IsNull(tbOrders!balans) Then
                BalansSum = Format((tbOrders!debit - tbOrders!kredit) * tbOrders!balans, "## ##0.00")
            Else
                BalansSum = ""
            End If
            Grid.AddItem tbOrders!row_id _
                & Chr(9) & tbOrders!row_descr _
                & Chr(9) & Format(tbOrders!debit, "## ##0.00") _
                & Chr(9) & Format(tbOrders!kredit, "## ##0.00") _
                & Chr(9) & BalansSum _
                & Chr(9) & tbOrders!detailSql _
                & Chr(9) & tbOrders!col_formatting _
                & Chr(9) & tbOrders!restorable _
                & Chr(9) & tbOrders!Sortable _
                & Chr(9) & tbOrders!Subtitle _
            
            sumD = sumD + tbOrders!debit
            sumK = sumK + tbOrders!kredit
            tbOrders.MoveNext
        Wend
    End If
    tbOrders.Close
    Grid.RemoveItem 1
    Grid.AddItem CStr(quantity + 1) & Chr(9) & "                                       ИТОГО:" & _
        Chr(9) & Format(sumD, "## ##0.00") _
        & Chr(9) & Format(sumK, "## ##0.00") _
        & Chr(9) & Format(sumD - sumK, "## ##0.00")
Grid.row = Grid.Rows - 1
Grid.col = 1: Grid.CellFontBold = True
Grid.col = 2: Grid.CellFontBold = True
Grid.col = 3: Grid.CellFontBold = True
Grid.col = 4: Grid.CellFontBold = True

End Sub

Public Function getDetailParameter(ByVal row_id As Long, ByVal name As String) As String
    If name = "description" Then
        getDetailParameter = Grid.TextMatrix(row_id, 1)
    ElseIf name = "detailSql" Then
        getDetailParameter = Grid.TextMatrix(row_id, 5)
    ElseIf name = "col_formatting" Then
        getDetailParameter = Grid.TextMatrix(row_id, 6)
    ElseIf name = "restorable" Then
        getDetailParameter = Grid.TextMatrix(row_id, 7)
    ElseIf name = "sortable" Then
        getDetailParameter = Grid.TextMatrix(row_id, 8)
    ElseIf name = "subtitle" Then
        getDetailParameter = Grid.TextMatrix(row_id, 9)
    ElseIf name = "balans" Then
        getDetailParameter = Grid.TextMatrix(row_id, 10)
    End If

End Function


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
cmPrint.Top = Grid.Top + Grid.Height + 50
cmExel.Top = cmPrint.Top
cmExit.Top = cmPrint.Top

cmExit.Left = Grid.Width - cmExit.Width - 50
cmExel.Left = cmExit.Left - cmExel.Width - 250
cmPrint.Left = cmExel.Left - cmPrint.Width - 100


End Sub


Private Sub Grid_Click()

    mousCol = Grid.MouseCol
    mousRow = Grid.MouseRow

'Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_DblClick()

    If Grid.CellBackColor <> &H88FF88 Then Exit Sub
    
    ReDim sqlRowDetail(1)
    ReDim aRowText(1)
    ReDim rowFormatting(1)
    ReDim aRowSortable(1)
    ReDim arowSubtitle(1)
    
    Dim Report2 As New Report
    Set Report2.Caller = Me
    
    Report2.Regim = "aReportDetail"
    Report2.param1 = Grid.TextMatrix(mousRow, 0)
    Report2.param2 = Grid.TextMatrix(mousRow, 1) 'description
    
    sqlRowDetail(1) = getDetailParameter(mousRow, "detailSql")
    aRowText(1) = getDetailParameter(mousRow, "description")
    rowFormatting(1) = getDetailParameter(mousRow, "col_formatting")
    aRowSortable(1) = getDetailParameter(mousRow, "sortable")
    arowSubtitle(1) = getDetailParameter(mousRow, "subtitle")
    
    
    Report2.Show vbModal
End Sub


Private Sub Grid_EnterCell()
    mousRow = Grid.row
    
    If Grid.TextMatrix(mousRow, 5) = "" Or Not isToday Then
        Grid.CellBackColor = vbYellow
    Else
        Grid.CellBackColor = &H88FF88
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
End If
End Sub

