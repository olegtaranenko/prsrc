VERSION 5.00
Begin VB.Form FindFirmComtex 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Поиск фирмы плательщика по бух. базе Комтех"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmFiltr 
      Caption         =   "Фильтр"
      Height          =   315
      Left            =   8520
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmNext 
      Caption         =   "Далее"
      Height          =   315
      Left            =   7560
      TabIndex        =   3
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmSelect 
      Caption         =   "Выбор"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tb 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2910
   End
End
Attribute VB_Name = "FindFirmComtex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim firmsId() As Integer
Public Regim As String
Public firmId As String
Dim pos As Integer, oldWord As String



Private Sub cmExit_Click()
Unload Me
End Sub



End Sub

Private Sub cmSelect_Click()
Dim sqlReq As String, DNM As String

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
End Sub

Sub loadFirms()
Dim I As Integer

sql = "SELECT GuideFirms.FirmId, GuideFirms.Name From GuideFirms " & _
"ORDER BY GuideFirms.Name;"
Set tbFirms = myOpenRecordSet("##70", sql, dbOpenForwardOnly)
If tbFirms Is Nothing Then Exit Sub
myRedim firmsId, 1000
I = 0
If Not tbFirms.BOF Then
  While Not tbFirms.EOF
    If tbFirms!firmId = 0 Then GoTo NXT
'    lb.AddItem tbFirms!name
    myRedim firmsId, I + 1
    firmsId(I) = tbFirms!firmId
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

End Sub

