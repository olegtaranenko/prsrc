VERSION 5.00
Begin VB.Form AUTO 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ������������� �����"
   ClientHeight    =   1836
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4704
   ControlBox      =   0   'False
   Icon            =   "AUTO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1836
   ScaleWidth      =   4704
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbM 
      Height          =   315
      ItemData        =   "AUTO.frx":030A
      Left            =   2520
      List            =   "AUTO.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmExit2 
      Caption         =   "�����"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmBook 
      Caption         =   "������ ���.��������"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmSklad 
      Caption         =   "�����"
      Enabled         =   0   'False
      Height          =   375
      Left            =   300
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox tbEnable 
      BackColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   4020
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer Timer1 
      Left            =   300
      Top             =   720
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1860
      TabIndex        =   0
      Top             =   1260
      Width           =   795
   End
   Begin VB.Label laManag 
      Caption         =   "��������:"
      Height          =   195
      Left            =   1620
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label laHour 
      Caption         =   "23:00"
      Height          =   255
      Left            =   3180
      TabIndex        =   4
      Top             =   780
      Width           =   435
   End
   Begin VB.Label laDate 
      Caption         =   "������� �����:  12.02.03  12:"
      Height          =   255
      Left            =   900
      TabIndex        =   3
      Top             =   780
      Width           =   2295
   End
   Begin VB.Label laInform 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "AUTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isOn As Boolean, isSinhro As Boolean ', baseIsOpen As Boolean
Sub dummy()
'Dim IsEmpty, Numorder, StatusId, Rollback, Outdatetime, p_numOrder, tbWorktime, Left, RemoveItem, J, Value, X, Y, Table, IL, Name, L, Equip, Worktime, ManagId, ColWidth, Index, W, K, Visible, Field, WerkId
Dim IsEmpty, Numorder, StatusId, Rollback, Outdatetime, p_numOrder, tbWorktime, Left, RemoveItem, J, Value, X, Y, Table, IL, Name, L, Equip, Worktime, ManagId, ColWidth, Index, W, K, Visible, Field, WerkId
End Sub
Sub nextWindow()
Dim str As String
Dim I As Integer

    
    cmSklad.Visible = True
    cmBook.Visible = True
    cmExit2.Visible = True
    Me.Caption = "����"
    laManag.Visible = True
    cbM.Visible = True
    
    laInform.Visible = False
    laDate.Visible = False
    laHour.Visible = False
    cmExit.Visible = False
    tbEnable.Visible = False
    

On Error GoTo 0
begDate = getSystemField("begDate")
If IsNull(begDate) Then End

'Set wrkDefault = DBEngine.Workspaces(0) ' ��� ���-�� ����������

'sql = "SELECT GuideManag.ManagId, GuideManag.Manag, GuideManag.listNum " & _
'      "From GuideManag ORDER BY GuideManag.listNum;"
'Set Table = myOpenRecordSet("##03", sql, dbOpenForwardOnly)
'If Table Is Nothing Then End'

'While Not Table.EOF
'    If IsNumeric(Table!listNum) Then
'        str = Table!Manag
'        If str <> "" Then cbM.AddItem str, Table!listNum
'        manId(Table!listNum) = Table!ManagId
'        Manag(Table!ManagId) = str
'    End If
'    Table.MoveNext
'Wend
'Table.Close

'*********************************************************************$$7
sql = "SELECT * From GuideManag WHERE Manag <>'not'  ORDER BY forSort;"
Set Table = myOpenRecordSet("##03", sql, dbOpenForwardOnly)
If Table Is Nothing Then myBase.Close: End
I = 0: ReDim manId(0):
Dim imax As Integer: imax = 0: ReDim Manag(0)
While Not Table.EOF
    If LCase(Table!ForSort) <> "unused" Then
        str = Table!Manag
        If Table!ManagId <> 0 Then cbM.AddItem str
        manId(I) = Table!ManagId
        If imax < Table!ManagId Then
            imax = Table!ManagId
            ReDim Preserve Manag(imax)
        End If
        Manag(Table!ManagId) = str
        I = I + 1
        ReDim Preserve manId(I):
    End If
    Table.MoveNext
Wend
Table.Close



ReDim Status(0)
Set Table = myOpenRecordSet("##05", "GuideStatus", dbOpenForwardOnly)
If Table Is Nothing Then myBase.Close: End
'Table.MoveFirst
I = 0 '���� ������
While Not Table.EOF
    If I < Table!StatusId Then
        I = Table!StatusId
        ReDim Preserve Status(I)
    End If
    Status(Table!StatusId) = Table!Status
    Table.MoveNext
Wend
Table.Close

'Exit Sub $$2
'ERRb:
'MsgBox "�� ������� ������������ � ���� '" & mainTitle & "'", , "Error 168" '##168
'End
    
End Sub

Private Sub cbM_Click()
    cmSklad.Enabled = True
    cmBook.Enabled = True
    sql = "set @manager = '" & cbM.Text & "'"
    If myExecute("##1.2", sql, 0) = 0 Then End

End Sub

Private Sub cmBook_Click()
Journal.Show 'vbModal

End Sub

Private Sub cmExit_Click()
If Not isOn Then
    End
ElseIf LCase(tbEnable.Text) = "arh" Or LCase(tbEnable.Text) = "���" Then
    dostup = "b"
    GoTo AA
ElseIf LCase(tbEnable.Text) = "arc" Or LCase(tbEnable.Text) = "���" Then
    dostup = "a"
AA: 'Unload Me
    nextWindow
Else
    End
End If

End Sub

Private Sub cmExit2_Click()
Unload Me
End Sub

Private Sub cmSklad_Click()
Documents.Show

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbCtrlMask And KeyCode = vbKeyL And isOn Then
    tbEnable.Text = ""
    tbEnable.Visible = True
    tbEnable.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim otladSet As Variant

otladSet = getEffectiveSetting("otlad")

If Not IsEmpty(otladSet) Then
    dostup = "a"
    nextWindow
    cbM.ListIndex = cbM.ListCount - 1
    cmSklad.Enabled = True
    cmBook.Enabled = True
    Me.BackColor = otladColor
    Exit Sub
Else
    dostup = getEffectiveSetting("dostup")
End If

isOn = True
isSinhro = True
laDate.Caption = ""
laHour.Caption = ""
tmpDate = Now()

On Error GoTo ERRs
'If Command() = "a" Then '�������� ��� �. ��������  winXP
If Dir$("C:\WINDOWS\net.exe") = "" Then '�� �����
    Shell "C:\WINDOWS\system32\net time \\server /SET /YES", vbHide ' winXP
Else
    Shell "C:\WINDOWS\net time \\server /WORKGROUP:JOBSHOP /SET /YES", vbHide
End If
On Error GoTo 0


Timer1.Interval = 100 ' 0.1 c
Timer1.Enabled = True

Exit Sub
ERRs:
isSinhro = False
Resume Next

End Sub


Private Sub Form_Unload(Cancel As Integer)
'If baseIsOpen Then myBase.Close
If Documents.isLoad Then Unload Documents
If Journal.isLoad Then Unload Journal
myBase.Close

End Sub

Private Sub tbEnable_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then cmExit_Click
End Sub

Private Sub Timer1_Timer()
Dim I As Long
Static sek As Integer

'If sek > 32000 Then End
sek = sek + 1
If sek = 10 Then
  If isSinhro Then
    I = DateDiff("s", tmpDate, Now()) - 1
    laInform.Caption = "������������� ������ ������� (���������  " & I & " ���.)"
  Else
    laInform.Caption = "������� �� ������ ���������������� ����!"
  End If
  laDate.Caption = "������� �����:   " & Format(Now(), "dd.mm.yy hh:")
End If

If sek >= 10 Then _
    laHour.Caption = Format(Now(), "nn:ss")

If sek > 100 Then
    Timer1.Enabled = False
    isOn = False
'    tbEnable.Visible = False
End If

End Sub
