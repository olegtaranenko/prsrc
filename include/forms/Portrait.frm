VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Portrait 
   BackColor       =   &H8000000A&
   Caption         =   "�����"
   ClientHeight    =   8184
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8184
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmPrint 
      Caption         =   "������"
      Height          =   315
      Left            =   2760
      TabIndex        =   6
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "�����"
      Height          =   315
      Left            =   10980
      TabIndex        =   4
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmExel 
      Caption         =   "������ � Excel"
      Height          =   315
      Left            =   3780
      TabIndex        =   3
      Top             =   7800
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7212
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11652
      _ExtentX        =   20553
      _ExtentY        =   12721
      _Version        =   393216
      MergeCells      =   2
      AllowUserResizing=   1
   End
   Begin VB.Label laHeader 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   432
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   11772
   End
   Begin VB.Label laRecCount 
      Caption         =   "����� �������:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   7830
      Width           =   1335
   End
   Begin VB.Label laCount 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   7800
      Width           =   615
   End
End
Attribute VB_Name = "Portrait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldHeight As Integer, oldWidth As Integer ' ��� ������ �����
Public nCols As Integer ' ����� ���-�� �������
Public mousRow As Long
Public mousCol As Long
Public mode As String, byRowId As Integer, byColumnId As Integer
Public filterId As Integer



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
    oldHeight = Me.Height
    oldWidth = Me.Width
    If filterId <= 0 Then Exit Sub
    If mode = "portrait" Then
        
    ElseIf mode = "detail" Then
        LoadTableDetail
    End If
End Sub

Private Sub LoadTableDetail()
Dim rownum As Integer

    sql = "call n_exec_filter( " & filterId & ", " & byRowId & ", " & byColumnId & ")"
'    Debug.Print sql
    Set Table = myOpenRecordSet("##Results.1", sql, dbOpenDynaset)
    If Table Is Nothing Then
        Table.Close
        MsgBox "������ ��� �������� ������ �� ����", vbCritical
        Exit Sub
    End If
    If Table.BOF Then
        Table.Close
        MsgBox "����� �� �������� ������", vbExclamation
        Exit Sub
    End If
    
    clearGrid Me.Grid
    Me.Grid.FormatString = "|����� ���.|���� ���.|>��������|>��������|>�-��.���.|>����� ���."
    rownum = 1
    Table.MoveFirst
    Dim I As Integer ' ����� �������
    While Not Table.EOF
        
        I = 1
        If rownum > 1 Then
            Grid.AddItem ""
        End If
        Grid.TextMatrix(rownum, I) = Table!Numorder: I = I + 1
        Grid.TextMatrix(rownum, I) = Table!inDate: I = I + 1
        Grid.TextMatrix(rownum, I) = Format(Table!orderOrdered, "# ###.00"): I = I + 1
        Grid.TextMatrix(rownum, I) = Format(Table!orderPaid, "# ###.00"): I = I + 1
        Grid.TextMatrix(rownum, I) = Format(Table!materialQty, "# ###.00"): I = I + 1
        Grid.TextMatrix(rownum, I) = Format(Table!materialSaled, "# ###.00"): I = I + 1
        
        Table.MoveNext
        rownum = rownum + 1
        
    Wend
    Table.Close

End Sub

Private Sub Form_Resize()
    Dim H As Integer, W As Integer
    
    If Me.WindowState = vbMinimized Then Exit Sub
    On Error Resume Next
    
    H = Me.Height - oldHeight
    oldHeight = Me.Height
    W = Me.Width - oldWidth
    oldWidth = Me.Width
    Grid.Height = Grid.Height + H
    Grid.Width = Grid.Width + W
    laRecCount.Top = laRecCount.Top + H
    laCount.Top = laCount.Top + H
    laHeader.Width = laHeader.Width + W
    cmExel.Top = cmExel.Top + H
    cmPrint.Top = cmPrint.Top + H
    cmExit.Top = cmExit.Top + H
    cmExit.Left = cmExit.Left + W
End Sub

Private Sub Grid_Click()
    mousCol = Grid.MouseCol
    mousRow = Grid.MouseRow
    If mousRow = 0 Then
        Grid.CellBackColor = Grid.BackColor
        If mousCol = 0 Then Exit Sub
        If mousCol > 3 Then
            SortCol Grid, mousCol, "numeric"
        Else
            SortCol Grid, mousCol
        End If
        Grid_LeaveCell
    End If
    
End Sub

Private Sub Grid_LeaveCell()
    Grid.CellBackColor = Grid.BackColor
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Grid.MouseRow = 0 And Shift = vbKeyShift Then
        MsgBox "ColWidth = " & Grid.ColWidth(Grid.MouseCol)
    End If
End Sub

