VERSION 5.00
Begin VB.Form FindFirm 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Расширенный поиск по всем фирмам из справочника"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmFiltr 
      Caption         =   "Фильтр"
      Height          =   315
      Left            =   3180
      TabIndex        =   5
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmNext 
      Caption         =   "Далее"
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmExit 
      Cancel          =   -1  'True
      Caption         =   "Выход"
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      Top             =   4380
      Width           =   675
   End
   Begin VB.CommandButton cmSelect 
      Caption         =   "Выбор"
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   4380
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.ListBox lb 
      Height          =   3765
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
Public firmId As String
Dim pos As Integer, oldWord As String
Dim tbFirms As Recordset

Private Sub cmExit_Click()
Unload Me
End Sub

Private Sub cmNext_Click()
Dim i As Integer, wordLen As Integer, word As String

pos = pos + 1
word = LCase(tb.Text)
oldWord = word
wordLen = Len(word)
For i = pos To lb.ListCount - 1
    If InStr(LCase(lb.List(i)), word) > 0 Then
        lb.Selected(i) = True
        pos = i
        cmSelect.Enabled = True
        Exit Sub
    End If
    pos = -1
Next i


End Sub


Private Sub cmNoCloseFiltr_Click()

End Sub

Private Sub cmSelect_Click()
Dim sqlReq As String, DNM As String

If Regim = "edit" Then
    If Journal.valueToBookField("##357", CStr(firmsId(lb.ListIndex)), _
    "KredDebitor") Then
        Journal.Grid.Text = lb.Text
        If jKassaReport.isLoad Then jKassaReport.laInform.Visible = True
    End If
End If
    
Unload Me

End Sub

Private Sub cmFiltr_Click()
Dim i As Integer, wordLen As Integer, word As String
Dim tmpId() As Integer, tmpName() As String, ti As Integer

If cmFiltr.Caption = "Фильтр" Then
 myRedim tmpId, 100
 myRedim tmpName, 100
 cmNext.Enabled = False
 
 
 word = LCase(tb.Text)
 oldWord = word
 wordLen = Len(word)
 ti = -1
 For i = 0 To lb.ListCount - 1
    If InStr(LCase(lb.List(i)), word) > 0 Then
        ti = ti + 1
        myRedim tmpId, ti
        myRedim tmpName, ti
        tmpName(ti) = lb.List(i)
        tmpId(ti) = firmsId(i)
    End If
 Next i
 lb.Clear
 For i = 0 To ti
    lb.AddItem tmpName(i), i
    firmsId(i) = tmpId(i)
 Next i
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
loadFirms
End Sub

Sub loadFirms()
Dim i As Integer

sql = "SELECT GuideFirms.FirmId, GuideFirms.Name From GuideFirms " & _
"ORDER BY GuideFirms.Name;"
Set tbFirms = myOpenRecordSet("##70", sql, dbOpenForwardOnly)
If tbFirms Is Nothing Then Exit Sub
myRedim firmsId, 1000
i = 0
If Not tbFirms.BOF Then
  While Not tbFirms.EOF
    If tbFirms!firmId = 0 Then GoTo NXT
    lb.AddItem tbFirms!Name
    myRedim firmsId, i + 1
    firmsId(i) = tbFirms!firmId
    i = i + 1
NXT:
    tbFirms.MoveNext
  Wend
End If
tbFirms.Close
End Sub

Private Sub lb_Click()
cmSelect.Enabled = True
End Sub

Private Sub lb_DblClick()
If cmSelect.Enabled = True And cmSelect.Visible = True Then cmSelect_Click
End Sub

Private Sub tb_Change()
Dim i As Integer, wordLen As Integer, word As String

word = LCase(tb.Text)
wordLen = Len(word)
If Left$(word, Len(oldWord)) <> oldWord Then pos = 0
oldWord = word
For i = pos To lb.ListCount - 1
    If LCase(Left$(lb.List(i), wordLen)) = word Then
        lb.Selected(i) = True
        pos = i
        cmSelect.Enabled = True
        Exit Sub
    End If
Next i

End Sub

