VERSION 5.00
Begin VB.Form AUTO 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Синхронизация часов"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ControlBox      =   0   'False
   Icon            =   "AUTO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4710
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
      Caption         =   "Выход"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmBook 
      Caption         =   "Журнал хоз.операций"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmSklad 
      Caption         =   "Склад"
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
      Caption         =   "Менеджер:"
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
      Caption         =   "Текущее время:  12.02.03  12:"
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

Sub nextWindow()
Dim str  As String
    cmSklad.Visible = True
    cmBook.Visible = True
    cmExit2.Visible = True
    Me.Caption = "Вход"
    laManag.Visible = True
    cbM.Visible = True
    
    laInform.Visible = False
    laDate.Visible = False
    laHour.Visible = False
    cmExit.Visible = False
    tbEnable.Visible = False
    
Dim i As Integer

'If InStr(otlad, ":\") > 0 Then '$$2
'    str = "\"
'    If Right$(otlad, 1) = "\" Then str = ""
'    baseNamePath = otlad & str & "dlsricN.mdb"
'    webProducts = otlad & str & "Products"
'    mainTitle = "    " & baseNamePath
'    webNomenks = otlad & str & "Nomenks."
If otlad = "work" Or otlad = "" Then '
'    mainTitle = ""
'   baseNamePath = "\\Server\D\!INSTAL!\EPILOG\RADIUS.V20\dlsricN.mdb"
'    webProducts = "\\SERVER\C\WebServers\home\petmas.ru\mirror\files\Products."
'    webNomenks = "\\Server\C\WebServers\home\petmas.ru\mirror\files\Nomenks."
    webProducts = cfg.ProductsPath '$$2
    webNomenks = cfg.NomenksPath '$$2
'    mainTitle = "         " & base(cfg.curBaseInd) '$$2
    cfg.baseOpen cfg.curBaseInd  '$$2
'ElseIf otlad = "otlad" Then
'    baseNamePath = "\\Server\D\!INSTAL!\EPILOG\RADIUS.V20\pitchN.mdb"
'    webProducts = "Products."
'    mainTitle = "Учебная"
'    webNomenks = "Nomenks."
Else '"otlaD"
'   baseNamePath = "D:\VB_DIMA\dlsricN.mdb"
    webProducts = "D:\VB_DIMA\Products"
    webNomenks = "D:\VB_DIMA\Nomenks."
'    mainTitle = baseNamePath
'    mainTitle = "    otlad"
    cfg.baseOpen '"C:\VB_DIMA\dlsricN.mdb"
End If

'On Error GoTo ERRb $$2
'                                                                                                                                                                            Set myBase = OpenDatabase(baseNamePath, False, False, ";PWD=play")
'Set myBase = OpenDatabase(baseNamePath)

On Error GoTo 0
begDate = getSystemField("begDate")
If IsNull(begDate) Then End

'Set wrkDefault = DBEngine.Workspaces(0) ' для орг-ии транзакций

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
i = 0: ReDim manId(0):
Dim imax As Integer: imax = 0: ReDim Manag(0)
While Not Table.EOF
    If LCase(Table!ForSort) <> "unused" Then
        str = Table!Manag
        If Table!ManagId <> 0 Then cbM.AddItem str
        manId(i) = Table!ManagId
        If imax < Table!ManagId Then
            imax = Table!ManagId
            ReDim Preserve Manag(imax)
        End If
        Manag(Table!ManagId) = str
        i = i + 1
        ReDim Preserve manId(i):
    End If
    Table.MoveNext
Wend
Table.Close



ReDim Status(0)
Set Table = myOpenRecordSet("##05", "GuideStatus", dbOpenForwardOnly)
If Table Is Nothing Then myBase.Close: End
'Table.MoveFirst
i = 0 'макс индекс
While Not Table.EOF
    If i < Table!StatusId Then
        i = Table!StatusId
        ReDim Preserve Status(i)
    End If
    Status(Table!StatusId) = Table!Status
    Table.MoveNext
Wend
Table.Close

'Exit Sub $$2
'ERRb:
'MsgBox "Не удалось подключиться к базе '" & mainTitle & "'", , "Error 168" '##168
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
ElseIf LCase(tbEnable.Text) = "arh" Or LCase(tbEnable.Text) = "фкр" Then
    dostup = "b"
    GoTo AA
ElseIf LCase(tbEnable.Text) = "arc" Or LCase(tbEnable.Text) = "фкс" Then
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

'If InStr(Command(), ":\") > 0 Then '$$2
'    otlad = Command()
'Else
If Len(Command()) > 4 Then
    otlad = Left$(Command(), 4)
End If


'If InStr(Command(), "       wkdh") <> 0 Then
If Right$(Command(), 4) = "wkdh" Then
    dostup = "a"
    nextWindow
    cbM.ListIndex = cbM.ListCount - 1
    cmSklad.Enabled = True
    cmBook.Enabled = True
    Exit Sub
End If

isOn = True
isSinhro = True
laDate.Caption = ""
laHour.Caption = ""
tmpDate = Now()

On Error GoTo ERRs
'If Command() = "a" Then 'временно это б. означать  winXP
If Dir$("C:\WINDOWS\net.exe") = "" Then 'не файла
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
Dim i As Long
Static sek As Integer

'If sek > 32000 Then End
sek = sek + 1
If sek = 10 Then
  If isSinhro Then
    i = DateDiff("s", tmpDate, Now()) - 1
    laInform.Caption = "Синхронизация прошла успешно (коррекция  " & i & " сек.)"
  Else
    laInform.Caption = "Система не смогла синхронизировать часы!"
  End If
  laDate.Caption = "Текущее время:   " & Format(Now(), "dd.mm.yy hh:")
End If

If sek >= 10 Then _
    laHour.Caption = Format(Now(), "nn:ss")

If sek > 100 Then
    Timer1.Enabled = False
    isOn = False
'    tbEnable.Visible = False
End If

End Sub
