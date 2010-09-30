VERSION 5.00
Begin VB.Form FindFirm 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Поиск по фирмам"
   ClientHeight    =   5040
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmNoCloseFiltr 
      Caption         =   "Фильтр""Незакрытые заказы"""
      Enabled         =   0   'False
      Height          =   315
      Left            =   840
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmFiltr 
      Caption         =   "Фильтр"
      Height          =   315
      Left            =   3180
      TabIndex        =   8
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmNext 
      Caption         =   "Далее"
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmNoClose 
      Caption         =   "Отчет ""Незакрытые заказы""  "
      Enabled         =   0   'False
      Height          =   315
      Left            =   840
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   4800
      TabIndex        =   5
      Top             =   4680
      Width           =   675
   End
   Begin VB.CommandButton cmAllOrders 
      Caption         =   "Отч.""Все заказы фирмы"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3420
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmGuide 
      Caption         =   "В справочник"
      Height          =   315
      Left            =   4260
      TabIndex        =   3
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmSelect 
      Caption         =   "Выбор"
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   4500
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.ListBox lb 
      Height          =   3696
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   5415
   End
   Begin VB.TextBox tb 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2310
   End
End
Attribute VB_Name = "FindFirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim firmsId() As Integer
Public Regim As String
Public FirmId As String
Public idWerk As Integer

Dim pos As Integer, oldWord As String


Private Sub cmAllOrders_Click()
FirmId = firmsId(lb.ListIndex)
Report.Regim = "allOrders"
Report.Show vbModal
End Sub

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmGuide_Click()
bilo = cmSelect.Visible
tmpStr = lb.Text

Me.MousePointer = flexHourglass
If idWerk = 1 Then
    BayGuideFirms.Regim = "fromFindFirm"
    BayGuideFirms.tbFind.Text = tmpStr
    BayGuideFirms.cmSel.Visible = bilo
    BayGuideFirms.Show vbModal
Else
    GuideFirms.Regim = "fromFindFirm"
    GuideFirms.tbFind.Text = tmpStr
    GuideFirms.cmSel.Visible = bilo
    GuideFirms.Show vbModal
End If

Me.MousePointer = flexDefault
Unload Me

End Sub

Private Sub cmNext_Click()
Dim I As Integer, wordLen As Integer, word As String

pos = pos + 1
word = LCase(tb.Text)
oldWord = word
wordLen = Len(word)
For I = pos To lb.ListCount - 1
    If InStr(LCase(lb.List(I)), word) > 0 Then
        lb.Selected(I) = True
        pos = I
        cmSelect.Enabled = True
        Exit Sub
    End If
    pos = -1
Next I


End Sub

Private Sub cmNoClose_Click()
Me.MousePointer = flexHourglass
FirmId = firmsId(lb.ListIndex)
Report.Regim = "Orders"
Report.Show vbModal
Me.MousePointer = flexDefault
End Sub

Private Sub cmNoCloseFiltr_Click()
Dim str As String
str = lb.Text
Unload Me
Orders.loadFirmOrders str
End Sub

Private Sub cmSelect_Click()
Dim sqlReq As String, DNM As String

If Regim = "edit" Then
    Orders.Grid.Text = lb.Text

    gNzak = Orders.Grid.TextMatrix(Orders.Grid.row, orNomZak)
    visits "-", "firm" ' уменьщаем посещения у старой фирмы, если она была
    FirmId = firmsId(lb.ListIndex)
    ValueToTableField "##20", FirmId, "Orders", "FirmId"
    visits "+", "firm" ' увеличиваем посещения у новой фирмы

    DNM = Format(Now(), "dd.mm.yy hh:nn") & vbTab & Orders.cbM.Text & " " & gNzak ' именно vbTab
    On Error Resume Next ' в некот.ситуациях один из Open logFile дает Err: файл уже открыт
    Open logFile For Append As #2
    Print #2, DNM & " фирма=" & lb.Text
    Close #2
ElseIf Regim = "fromFiltr" Then
    Filtr.lbFirm.AddItem lb.Text, 0
    Filtr.lbFirm.Selected(0) = True
End If
refreshTimestamp gNzak

Unload Me

End Sub

Private Sub cmFiltr_Click()
Dim I As Integer, wordLen As Integer, word As String
Dim tmpId() As Integer, tmpName() As String, ti As Integer

If cmFiltr.Caption = "Фильтр" Then
 myRedim tmpId, 100
 myRedim tmpName, 100
 cmNext.Enabled = False
 
 
 word = LCase(tb.Text)
 oldWord = word
 wordLen = Len(word)
 ti = -1
 For I = 0 To lb.ListCount - 1
    If InStr(LCase(lb.List(I)), word) > 0 Then
        ti = ti + 1
        myRedim tmpId, ti
        myRedim tmpName, ti
        tmpName(ti) = lb.List(I)
        tmpId(ti) = firmsId(I)
    End If
 Next I
 lb.Clear
 For I = 0 To ti
    lb.AddItem tmpName(I), I
    firmsId(I) = tmpId(I)
 Next I
 cmFiltr.Caption = "Обновить"
Else
 cmNext.Enabled = True
 tb.Text = ""
 lb.Clear
 loadFirms
 cmFiltr.Caption = "Фильтр"
End If

End Sub

Private Sub Form_Load()
isFindFirm = True
loadFirms
Me.Caption = Werk(idWerk) & " - Поиск по фирмам"
End Sub

Sub loadFirms()
Dim I As Integer, Name

sql = "SELECT FirmId, Name From FirmGuide " _
& "Where WerkId = " & idWerk _
& "ORDER BY Name"
Set tbFirms = myOpenRecordSet("##70", sql, dbOpenForwardOnly)
If tbFirms Is Nothing Then Exit Sub
myRedim firmsId, 1000
I = 0
If Not tbFirms.BOF Then
  While Not tbFirms.EOF
    If tbFirms!FirmId = 0 Then GoTo NXT
    lb.AddItem tbFirms!Name
    myRedim firmsId, I + 1
    firmsId(I) = tbFirms!FirmId
    I = I + 1
NXT:
    tbFirms.MoveNext
  Wend
End If
tbFirms.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
isFindFirm = False
End Sub

Private Sub lb_Click()
cmSelect.Enabled = True
    cmNoClose.Enabled = True
    cmAllOrders.Enabled = True
    cmNoCloseFiltr.Enabled = True
End Sub

Private Sub lb_DblClick()
If cmSelect.Enabled = True And cmSelect.Visible = True Then cmSelect_Click
End Sub

Private Sub tb_Change()
Dim I As Integer, wordLen As Integer, word As String

word = LCase(tb.Text)
wordLen = Len(word)
If Left$(word, Len(oldWord)) <> oldWord Then pos = 0
oldWord = word
For I = pos To lb.ListCount - 1
    If LCase(Left$(lb.List(I), wordLen)) = word Then
        lb.Selected(I) = True
        pos = I
        cmSelect.Enabled = True
        cmNoClose.Enabled = True
        cmAllOrders.Enabled = True
        cmNoCloseFiltr.Enabled = True
        Exit Sub
    End If
Next I

End Sub

