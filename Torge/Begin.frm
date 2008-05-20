VERSION 5.00
Begin VB.Form Begin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Склад"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmProdukt 
      Caption         =   "Справочник готовых изделий"
      Height          =   435
      Left            =   900
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmNomenk 
      Caption         =   "Справочник по номенклатуре"
      Height          =   435
      Left            =   900
      TabIndex        =   0
      Top             =   420
      Width           =   3015
   End
End
Attribute VB_Name = "Begin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmNomenk_Click()
Nomenklatura.Regim = ""
Nomenklatura.Show vbModal
End Sub

Private Sub cmProdukt_Click()

Products.Show vbModal

End Sub

Private Sub Form_Load()
cmdLine = Command()
If cmdLine = "otlad" Then
    baseNamePath = "D:\VB_DIMA\torge.mdb"
    Me.Caption = "Склад     " & baseNamePath
Else
    baseNamePath = "\\Server\D\!INSTAL!\EPILOG\RADIUS.V20\torge.mdb"
    Me.Caption = "Склад"
End If
Set myBase = OpenDatabase(baseNamePath)
End Sub
