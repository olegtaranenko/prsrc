VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Results 
   Caption         =   "Результаты анализа"
   ClientHeight    =   7812
   ClientLeft      =   48
   ClientTop       =   588
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   7812
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5292
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11892
      _ExtentX        =   20976
      _ExtentY        =   9335
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11892
      _ExtentX        =   20976
      _ExtentY        =   656
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Количество"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Сумма"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public filterId As Integer
Public applyTriggered As Boolean
Public StartDate As Date
Public endDate As Date



Private Sub Form_Activate()
    If applyTriggered Then
        Debug.Print "filterId = " & filterId
        LoadTable
    End If
End Sub


Private Sub Form_Resize()

    Grid.left = 100
    Grid.Width = Me.Width - 300
    TabStrip1.Top = 100
    TabStrip1.Width = Grid.Width
    TabStrip1.left = Grid.left
    Grid.Top = TabStrip1.Top + TabStrip1.Height
    Grid.Height = Me.Height - Grid.Top - 1200

End Sub

Private Sub TabStrip1_Click()
Dim currentTab As Tabs
Dim curTab As Variant

'   Debug.Print TabStrip1.SelectedItem.index
    Set curTab = TabStrip1.SelectedItem

End Sub


Private Sub LoadTable()
    applyTriggered = False
    sql = "call n_exec_filter( '" & Format(StartDate, "yyyymmdd") & "', '" & _
            Format(endDate, "yyyymmdd") & "', " & filterId & ")"
    Set table = myOpenRecordSet("##Results.1", sql, dbOpenDynaset)
    If table Is Nothing Then Exit Sub
    If table.BOF Then
        fatalError "Ошибка при загрузки данных из базы"
    End If
    
    table.MoveFirst
    While Not table.EOF
        table.MoveNext
    End
    table.Close
            
End Sub
