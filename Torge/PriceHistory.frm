VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PriceHistory 
   Caption         =   "История изменения цены"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   6360
      TabIndex        =   0
      Text            =   "tbMobile"
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   6480
      TabIndex        =   1
      Top             =   5880
      Width           =   915
   End
   Begin VB.CommandButton cmPrint 
      Caption         =   "Печать"
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   5880
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9340
      _Version        =   393216
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      Enabled         =   -1  'True
      MergeCells      =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lbName 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lbCod 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbNomnom 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "PriceHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldHeight As Integer, oldWidth As Integer
Dim mousCol As Long, mousRow As Long
Dim isLoad As Boolean

Const phTimestamp = 0
Const phCost = 2
Const phCostNew = 3
Const phDate = 1
Const phChangedBy = 4
Const phNote = 5



Sub lbHide()
    tbMobile.Visible = False
    Grid.Enabled = True
    On Error Resume Next
    Grid.SetFocus
    Grid_EnterCell
End Sub

Private Sub cmExcel_Click()
    GridToExcel Grid, "Карта движения по номенклатуре '" & gNomNom & "' по предприятиям"
End Sub

Private Sub cmExit_Click()
    Unload Me
End Sub
    
Private Sub Form_Load()
Dim I As Integer, sz As Integer
    sql = "select * from sguidenomenk where nomnom = '" & gNomNom & "'"
    Set Table = myOpenRecordSet("##234.1", sql, dbOpenForwardOnly)
    While Not Table.EOF
        lbNomnom.Caption = gNomNom
        lbName.Caption = Table!Nomname
        If Not IsNull(Table!cod) Then
            lbCod.Caption = Table!cod
            Table.MoveNext
        End If
    Wend
    Table.Close
    
    Grid.Visible = False
    oldHeight = Me.Height
    oldWidth = Me.Width

    sql = "select ph.change_date as dat, ph.cost, m.manag " _
        & " from spricehistory ph" _
        & " left join guidemanag m on m.managid = changed_by_id" _
        & " where nomnom = '" & gNomNom & "'" _
        & "     union " _
        & " select convert(datetime, '21000101') as dat, cost, null" _
        & " from sguidenomenk " _
        & " where nomnom = '" & gNomNom & "'" _
        & " order by 1"
    
    Set Table = myOpenRecordSet("##234.1", sql, dbOpenForwardOnly)
    If Table Is Nothing Then End
    
    Grid.FormatString = "|^Дата|>Цена пред|>Цена новая|^Мен.|Примечание"
    Grid.ColWidth(0) = 0
    Grid.ColWidth(phDate) = 1000
    Grid.ColWidth(phCost) = 1000
    Grid.ColWidth(phCostNew) = 1000
    Grid.ColWidth(phChangedBy) = 400
    
    Grid.Rows = 1
    
    While Not Table.EOF
        If Grid.Rows > 1 Then
            Grid.TextMatrix(Grid.Rows - 1, phCostNew) = Format(Table!cost, "#.00")
        End If
        Grid.AddItem CStr(Table!dat)
        Grid.TextMatrix(Grid.Rows - 1, phDate) = Format(Table!dat, "dd-mm-yyyy")
        Grid.TextMatrix(Grid.Rows - 1, phCost) = Format(Table!cost, "#.00#")
        If Not IsNull(Table!Manag) Then
            Grid.TextMatrix(Grid.Rows - 1, phChangedBy) = Table!Manag
        End If
        Table.MoveNext
    Wend
    Grid.RemoveItem (Grid.Rows - 1)
    Table.Close
    
    
    Grid.Visible = True
    isLoad = True
End Sub
Private Sub Form_Resize()
    Dim H As Integer, W As Integer
    
    If WindowState = vbMinimized Then Exit Sub
    On Error Resume Next
    H = Me.Height - oldHeight
    oldHeight = Me.Height
    W = Me.Width - oldWidth
    oldWidth = Me.Width
    Grid.Height = Grid.Height + H
    Grid.Width = Grid.Width + W
    cmExit.Top = cmExit.Top + H
    cmExit.Left = Grid.Left + Grid.Width - cmExit.Width
    
    cmPrint.Top = cmPrint.Top + H
    cmPrint.Left = cmExit.Left - 50 - cmPrint.Width
    
    'cmExcel.Top = cmExcel.Top + h
    
End Sub
Private Sub Grid_Click()
    mousCol = Grid.MouseCol
    mousRow = Grid.MouseRow
    'If quantity = 0 Then Exit Sub

End Sub


Private Sub Grid_DblClick()
    If Grid.CellBackColor = &H88FF88 Then
        If mousCol = phNote Then
            textBoxInGridCell tbMobile, Grid
        End If
    End If

End Sub

Private Sub Grid_EnterCell()
Dim isVentureOrder As Boolean
    mousRow = Grid.row
    mousCol = Grid.col

    If mousCol = 0 Then Exit Sub

    If _
            mousCol = phNote _
    Then
        Grid.CellBackColor = &H88FF88
    Else
        Grid.CellBackColor = vbYellow
    End If
    
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        mousCol = Grid.col
        mousRow = Grid.row
        Grid_DblClick
    End If

End Sub

Private Sub Grid_LeaveCell()
    If Grid.col <> 0 Then Grid.CellBackColor = Grid.BackColor
End Sub


Private Sub lbCod_Change()
    lbName.Left = lbCod.Left + lbCod.Width + 100
End Sub


Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String
    If KeyCode = vbKeyReturn Then
        str = tbMobile.Text
        If mousCol = phNote Then
                lbHide
        End If
    ElseIf KeyCode = vbKeyEscape Then
        lbHide
    End If

End Sub
