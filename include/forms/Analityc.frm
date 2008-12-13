VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Analityc 
   Caption         =   "Параметры запроса"
   ClientHeight    =   9276
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   5268
   LinkTopic       =   "Form1"
   ScaleHeight     =   9276
   ScaleWidth      =   5268
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView tvColumns 
      Height          =   2532
      Left            =   1440
      TabIndex        =   36
      Top             =   1680
      Visible         =   0   'False
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   4466
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   0
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
   Begin VB.CommandButton cmColumns 
      Caption         =   "Выбор столбцов"
      Height          =   315
      Left            =   1440
      TabIndex        =   34
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Frame Frame4 
      Caption         =   "Группировки ..."
      Height          =   1212
      Left            =   240
      TabIndex        =   19
      Top             =   1800
      Width           =   4932
      Begin VB.TextBox tbTop 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   192
         Left            =   2040
         TabIndex        =   31
         Text            =   "10"
         Top             =   840
         Width           =   252
      End
      Begin VB.CheckBox ckTop 
         Caption         =   "Только первые "
         Enabled         =   0   'False
         Height          =   252
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Width           =   1572
      End
      Begin VB.ComboBox cbGroupByRow 
         Height          =   288
         ItemData        =   "Analityc.frx":0000
         Left            =   360
         List            =   "Analityc.frx":0013
         TabIndex        =   28
         Text            =   "Фирмы"
         Top             =   480
         Width           =   2052
      End
      Begin VB.ComboBox cbGroupByColumn 
         Height          =   288
         ItemData        =   "Analityc.frx":0053
         Left            =   2640
         List            =   "Analityc.frx":0069
         TabIndex        =   20
         Text            =   "Месяцы"
         Top             =   480
         Width           =   2172
      End
      Begin VB.Label Label4 
         Caption         =   "позиций"
         Height          =   252
         Left            =   2400
         TabIndex        =   32
         Top             =   840
         Width           =   852
      End
      Begin VB.Label lbTop 
         Caption         =   "... по строкам"
         Enabled         =   0   'False
         Height          =   252
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "... по столбцам"
         Height          =   252
         Left            =   2400
         TabIndex        =   21
         Top             =   240
         Width           =   1452
      End
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   4320
      TabIndex        =   23
      Top             =   1320
      Width           =   852
   End
   Begin VB.CommandButton cmApply 
      Caption         =   "Применить"
      Height          =   315
      Left            =   240
      TabIndex        =   22
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Выбор периода"
      Height          =   1092
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   4932
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   492
         Left            =   3360
         TabIndex        =   35
         ToolTipText     =   "Позволяет одновременно сдвинуть даты на одинаковый период"
         Top             =   480
         Width           =   312
         _ExtentX        =   550
         _ExtentY        =   868
         _Version        =   393216
         Enabled         =   0   'False
      End
      Begin VB.CheckBox ckStartDate 
         Caption         =   " "
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   195
      End
      Begin VB.CheckBox ckEndDate 
         Caption         =   " "
         Height          =   315
         Left            =   1800
         TabIndex        =   16
         Top             =   600
         Width           =   200
      End
      Begin VB.ComboBox cbDateShift 
         Enabled         =   0   'False
         Height          =   288
         ItemData        =   "Analityc.frx":00A5
         Left            =   3720
         List            =   "Analityc.frx":00BE
         TabIndex        =   15
         Text            =   "год"
         ToolTipText     =   "Выбор периода сдвига даты"
         Top             =   600
         Width           =   1092
      End
      Begin MSComCtl2.DTPicker tbStartDate 
         Height          =   288
         Left            =   360
         TabIndex        =   25
         Top             =   600
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   39599
      End
      Begin MSComCtl2.DTPicker tbEndDate 
         Height          =   288
         Left            =   2040
         TabIndex        =   26
         Top             =   600
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   39599
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "сдвинуть обе даты на"
         Height          =   192
         Left            =   3000
         TabIndex        =   37
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Начиная с даты"
         Height          =   192
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1296
      End
      Begin VB.Label laPo 
         AutoSize        =   -1  'True
         Caption         =   "по дату"
         Height          =   192
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   624
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Дополнительные условия"
      Height          =   4572
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   4932
      Begin VB.CheckBox ckKriteriumFirms 
         Caption         =   "Выбор фирм(ы)"
         Height          =   252
         Left            =   2880
         TabIndex        =   38
         Top             =   2040
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.CheckBox ckKriteriumNoOborud 
         Caption         =   "Без оборудования"
         Height          =   252
         Left            =   2640
         TabIndex        =   33
         Top             =   3840
         Width           =   2052
      End
      Begin VB.CheckBox ckKriteriumOborud 
         Caption         =   "Выбор оборудования"
         Height          =   252
         Left            =   120
         TabIndex        =   24
         Top             =   3840
         Width           =   2052
      End
      Begin VB.CheckBox ckKriteriumRegion 
         Caption         =   "Выбор региона"
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1692
      End
      Begin MSComctlLib.TreeView tvMat 
         Height          =   1332
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   4332
         _ExtentX        =   7641
         _ExtentY        =   2350
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         Enabled         =   0   'False
      End
      Begin VB.CheckBox ckKriteriumMat 
         Caption         =   "Выбор групп материалов"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3252
      End
      Begin VB.CheckBox cbOborud 
         Caption         =   "Механика"
         Enabled         =   0   'False
         Height          =   252
         Index           =   3
         Left            =   3600
         TabIndex        =   1
         Top             =   4200
         Width           =   1140
      End
      Begin VB.CheckBox cbOborud 
         Caption         =   "Сублимация"
         Enabled         =   0   'False
         Height          =   252
         Index           =   2
         Left            =   1800
         TabIndex        =   2
         Top             =   4200
         Width           =   1380
      End
      Begin VB.CheckBox cbOborud 
         Caption         =   "Лазер"
         Enabled         =   0   'False
         Height          =   252
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   4200
         Width           =   900
      End
      Begin MSComctlLib.TreeView tvRegion 
         Height          =   1332
         Left            =   360
         TabIndex        =   13
         Top             =   2400
         Width           =   4332
         _ExtentX        =   7641
         _ExtentY        =   2350
         _Version        =   393217
         Indentation     =   529
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
         Enabled         =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Предустановленные фильтры"
      Height          =   1332
      Left            =   240
      TabIndex        =   4
      Top             =   7800
      Width           =   4932
      Begin VB.CommandButton cmFilterApply 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   252
         Left            =   3720
         Picture         =   "Analityc.frx":00F8
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Применить Фильтр"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   252
      End
      Begin VB.ComboBox cbFilters 
         Height          =   288
         ItemData        =   "Analityc.frx":04E2
         Left            =   120
         List            =   "Analityc.frx":04E4
         TabIndex        =   8
         Top             =   360
         Width           =   3492
      End
      Begin VB.TextBox txFilterName 
         Height          =   288
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   3492
      End
      Begin VB.CommandButton cmFilterAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   252
         Left            =   3720
         Picture         =   "Analityc.frx":04E6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Сохранить фильтр"
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   252
      End
      Begin VB.CommandButton cmFilterDelete 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   252
         Left            =   4080
         Picture         =   "Analityc.frx":08B5
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Удалить Фильтр"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   252
      End
   End
End
Attribute VB_Name = "Analityc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public managId As String
Public applicationType As String


Dim tbKlass As Recordset
Dim Node As Node
Dim columnsVisible As Boolean

Dim flagInitFilter As Boolean



Private Sub cbGroupByColumn_Click()
    checkDirtyFilterCommads
End Sub

Private Sub cbGroupByRow_Change()
    initByColumnList cbGroupByRow.ItemData(cbGroupByRow.ListIndex)
    If cbGroupByColumn.ListCount > 0 Then
        cbGroupByColumn.ListIndex = 0
    End If
    checkDirtyFilterCommads
End Sub


Private Sub cbGroupByRow_Click()
    cbGroupByRow_Change
End Sub

Private Sub cbOborud_Click(index As Integer)
    checkDirtyFilterCommads
End Sub


Private Sub ckEndDate_Click()
    If ckEndDate.value = 1 Then
        tbEndDate.Enabled = True
    Else
        tbEndDate.Enabled = False
    End If
    checkUpDown
End Sub


Private Sub ckKriteriumFirms_Click()
    checkDirtyFilterCommads
    If ckKriteriumMat.value = 1 Then
'        tvMat.Enabled = True
    Else
'        tvMat.Enabled = False
    End If
End Sub

Private Sub ckKriteriumMat_Click()
    checkDirtyFilterCommads
    If ckKriteriumMat.value = 1 Then
        tvMat.Enabled = True
    Else
        tvMat.Enabled = False
    End If
End Sub


Private Sub ckKriteriumNoOborud_Click()
    checkDirtyFilterCommads
    If ckKriteriumNoOborud.value = 1 Then
        If ckKriteriumOborud.value = 1 Then
            ckKriteriumOborud.value = 0
        End If
    End If
End Sub

Private Sub ckKriteriumRegion_Click()
    checkDirtyFilterCommads
    If ckKriteriumRegion.value = 1 Then
        tvRegion.Enabled = True
        ckKriteriumFirms.Enabled = True
    Else
        tvRegion.Enabled = False
        ckKriteriumFirms.Enabled = False
    End If
End Sub


Private Sub ckKriteriumOborud_Click()
Dim i As Integer

    checkDirtyFilterCommads
    If ckKriteriumOborud.value = 1 Then
        If ckKriteriumNoOborud.value = 1 Then
            ckKriteriumNoOborud.value = 0
        End If
        For i = 1 To 3
            cbOborud(i).Enabled = True
        Next i
    Else
        For i = 1 To 3
            cbOborud(i).Enabled = False
        Next i
    End If
End Sub


Private Sub ckStartDate_Click()
    If ckStartDate.value = 1 Then
        tbStartDate.Enabled = True
    Else
        tbStartDate.Enabled = False
    End If
    checkUpDown
End Sub
Private Sub checkUpDown()
    If tbStartDate.Enabled And tbEndDate.Enabled Then
        UpDown1.Enabled = True
        cbDateShift.Enabled = True
    Else
        UpDown1.Enabled = False
        cbDateShift.Enabled = False
    End If
End Sub
Private Sub cmApply_Click()
Dim filterId As Integer

    Results.left = Me.left + Me.Width
    Results.Top = Me.Top
    Results.filterId = submitFilter("")
    Results.applyTriggered = True
    Results.managId = managId
    If ckStartDate.value = 1 Then
        Results.startDate = tbStartDate.value
    Else
        Results.startDate = Empty
    End If
    If ckEndDate.value = 1 Then
        Results.endDate = tbEndDate.value
    Else
        Results.endDate = Empty
    End If
    Results.Show , Me

End Sub


Private Sub cmExit_Click()
    Unload Me
End Sub

Private Sub checkDirtyFilterCommads()
    If Not flagInitFilter Then
        If txFilterName <> "" Then
            cmFilterAdd.Enabled = True
        End If
    End If

End Sub


Sub loadRegions()
Dim Key As String, pKey As Variant, k() As String, pK()  As String


    sql = "call wf_territory_catalog (1)"
    Set tbKlass = myOpenRecordSet("##loadRegions", sql, dbOpenForwardOnly)
    If tbKlass Is Nothing Then myBase.Close: End
    
    If Not tbKlass.BOF Then
        tvRegion.Nodes.Clear
        While Not tbKlass.EOF
            Key = "k" & tbKlass!regionId
            If Not IsNull(tbKlass!territoryId) Then
                pKey = "k" & tbKlass!territoryId
            Else
                pKey = Null
            End If
            
            If IsNull(pKey) Then
                Set Node = tvRegion.Nodes.Add(, , Key, tbKlass!region)
            Else
                If Not existsInTreeview(tvRegion, Key) Then
                    Set Node = tvRegion.Nodes.Add(pKey, tvwChild, Key, tbKlass!region)
                End If
            End If
            
            If Not IsNull(tbKlass!firmId) Then
                Set Node = tvRegion.Nodes.Add(Key, tvwChild, "f" & tbKlass!firmId, tbKlass!FirmName)
            End If
                
            
            tbKlass.MoveNext
        Wend
    End If
    tbKlass.Close

End Sub


Sub loadKlass()
Dim Key As String, pKey As String, k() As String, pK()  As String
    sql = "call wf_klass_catalog"
    Set tbKlass = myOpenRecordSet("##loadKlasss", sql, dbOpenForwardOnly)
    If tbKlass Is Nothing Then myBase.Close: End
    
    If Not tbKlass.BOF Then
        tvMat.Nodes.Clear
'        Set Node = tvMat.Nodes.Add(, , "k0", "Все регионы")
'        Node.Sorted = True
        While Not tbKlass.EOF
            Key = "k" & tbKlass!r_KlassId
            If Not IsNull(tbKlass!r_parentKlassId) And (tbKlass!r_parentKlassId <> 0) Then
                pKey = "k" & tbKlass!r_parentKlassId
                Set Node = tvMat.Nodes.Add(pKey, tvwChild, Key, tbKlass!r_KlassName)
            Else
                Set Node = tvMat.Nodes.Add(, , Key, tbKlass!r_KlassName)
            End If
            
            tbKlass.MoveNext
        Wend
    End If
    tbKlass.Close

End Sub


Private Sub cleanTree(ByRef tView As TreeView)
Dim currentNode As Node
Dim i As Integer, nCount As Integer
Dim enabledFlag As Boolean

    enabledFlag = tView.Enabled
    tView.Enabled = True
    
    nCount = tView.Nodes.Count
    For i = 1 To nCount
        Set currentNode = tView.Nodes(i)
        If currentNode.checked Then
            currentNode.checked = False
            currentNode.Expanded = False
        End If
    Next i

    tView.Enabled = enabledFlag
End Sub

Private Sub cleanOborud()
Dim currentOborud As CheckBox
Dim i As Integer, nCount As Integer

    'nCount = UBound(cbOborud)
    For i = 1 To 3
        Set currentOborud = cbOborud(i)
        currentOborud.value = 0
    Next i
End Sub

Private Sub cleanFilterWindows()
    ckKriteriumMat.value = 0
    cleanTree tvMat
    ckKriteriumRegion.value = 0
    cleanTree tvRegion
    
    ckKriteriumOborud.value = 0
    ckKriteriumNoOborud.value = 0
    cleanOborud
End Sub


Sub initFilter(filterName As String, personal As Integer)

Dim filterId As Integer, byRowId As Integer, byColumnId As Integer

    flagInitFilter = True
    txFilterName.Text = cbFilters.Text
    cmFilterAdd.Enabled = False

    cleanFilterWindows
    
    sql = "select id, byrowId, bycolumnId from nFilter where name = '" & filterName & "' and personal = " & personal
    
    byErrSqlGetValues "W#initFilter.1", sql, filterId, byRowId, byColumnId
    If filterId = 0 Then
        'Нет еще такого фильтра
        GoTo done
    End If
    
    If setListIndexByItemDataValue(cbGroupByRow, byRowId) Then
        cbGroupByRow_Change
        setListIndexByItemDataValue cbGroupByColumn, byColumnId
    End If
    
    
    sql = " select " _
        & " i.id as itemId, p.id as paramId, isActive as isActive, itemType, paramType, paramClass, intValue, charValue" _
        & " from nItem i" _
        & "  left join nItemType it  on i.itemTypeId  = it.id" _
        & "  left join nParam p      on p.itemId      = i.id" _
        & "  left join nParamType pt on p.paramTypeId = pt.id" _
        & "  where i.filterid = " & filterId
    
    Set table = myOpenRecordSet("##initFilter.2", sql, dbOpenForwardOnly)
    If table Is Nothing Then myBase.Close: End
    ckStartDate.value = 0
    tbStartDate.value = Now() - 365
    tbStartDate.Enabled = False
    ckEndDate.value = 0
    tbEndDate.value = Now()
    tbEndDate.Enabled = False
    
    While Not table.EOF
        If table!itemType = "materials" Then
            If table!isActive = 1 Then
                ckKriteriumMat.value = 1
            Else
                ckKriteriumMat.value = 0
            End If
            If Not IsNull(table!intValue) Then
                Set Node = tvMat.Nodes("k" & table!intValue)
                expandParents Node
                Node.checked = True
            End If
            
        End If
        
        If table!itemType = "regions" Then
            If table!isActive = 1 Then
                ckKriteriumRegion.value = 1
            Else
                ckKriteriumRegion.value = 0
            End If
            Set Node = tvRegion.Nodes("k" & table!intValue)
            Node.checked = True
            expandParents Node
        End If
        
        If table!itemType = "oborudItems" Then
            If table!isActive = 1 And ckKriteriumNoOborud.value = 0 Then
                ckKriteriumOborud.value = 1
            Else
                ckKriteriumRegion.value = 0
            End If
            cbOborud(table!intValue).value = 1
        End If
        
        If table!itemType = "noOboruds" Then
            ckKriteriumNoOborud.value = 1
        End If
        
        If table!itemType = "filterPeriod" Then
            If table!paramType = "periodStart" Then
                ckStartDate.value = 1
                tbStartDate.value = table!charValue
                tbStartDate.Enabled = True
            End If
            If table!paramType = "periodEnd" Then
                ckEndDate.value = 1
                tbEndDate.value = table!charValue
                tbStartDate.Enabled = True
            End If
        End If
        
        table.MoveNext
    Wend
    table.Close
    
done:
    flagInitFilter = False

End Sub


Private Function prepareFilter(filterName As String, personal As Integer, byRowId As Integer, byColumnId As Integer) As Integer
Dim exists As Integer, result As Integer

    sql = "select id from nFilter " _
        & "where name = '" & filterName & "' and personal = " & personal
    byErrSqlGetValues "W#prepareFilter", sql, result
    
    If result <> 0 Then
        sql = "delete from nItem i from nFilter f " _
            & " where f.name = '" & filterName & "'" _
            & " and i.filterId = f.id and f.personal = " & personal
        myExecute "W#deleteFilter", sql, -1
        sql = "update nFilter set byrowid = " & byRowId & ", byColumnId = " & byColumnId _
            & "where name = '" & filterName & "' and personal = " & personal
        myExecute "W#updateFilter", sql, -1
    Else
        sql = "select n_insertFilter ('" & filterName & "', '" & managId & "', " & personal _
            & ", " & byRowId & ", " & byColumnId & ")"
        byErrSqlGetValues "W#clearFilter", sql, result
    End If
    
    prepareFilter = result

End Function


Private Function saveFilterItem(filterId As Integer, itemName As String, value As Variant) As Integer
Dim result As Integer
    
    sql = "select n_insertItem (" _
    & filterId _
    & ", '" & itemName & "', "
    If IsNumeric(value) Then
        sql = sql & CStr(value)
    Else
        sql = sql & "'" & CStr(value) & "'"
    End If
    sql = sql & ")"
    byErrSqlGetValues "##insertFilterItem", sql, result
    saveFilterItem = result
End Function


Private Function saveFilterParam(itemId As Integer, paramName As String, value As Variant) As Integer
Dim result As Integer
    
    sql = "select n_insertParam(" _
        & itemId _
        & ",'" & paramName & "', "
        If IsNumeric(value) Then
            sql = sql & CStr(value) & ", null"
        Else
            sql = sql & "null, '" & CStr(value) & "'"
        End If
        sql = sql & ")"
    
    'Debug.Print sql
    byErrSqlGetValues "##insertFilterParam", sql, result
    saveFilterParam = result
    
End Function


Private Function submitFilter(filterName As String) As Integer
Dim hasCheckedMat As Boolean, hasCheckedReg As Boolean, hasOborud As Boolean
Dim itemId As Integer
Dim filterId As Integer
Dim personal As Integer

    ' проверяем группы материалов
    hasCheckedMat = getCheckedInTree(tvMat)
    If Not hasCheckedMat And ckKriteriumMat.value = 1 Then
        MsgBox "Не выбрано ни одной группы материалов. " _
        & vbCr & "Нужно выбрать хотя бы одну или отключить критерий по материалом", vbExclamation, "Неправильный выбор параметров"
        Exit Function
    End If
    
    ' проверяем регионы
    hasCheckedReg = getCheckedInTree(tvRegion)
    If Not hasCheckedReg And ckKriteriumRegion.value = 1 Then
        MsgBox "Не выбрано ни одного региона. " _
        & vbCr & "Нужно выбрать хотя бы один или отключить критерий по регионам", vbExclamation, "Неправильный выбор параметров"
        Exit Function
    End If

    hasOborud = getOborudItems
    
    If Not hasOborud And ckKriteriumOborud.value = 1 Then
        MsgBox "Не выбрано никакого типа оборудования. " _
        & vbCr & "Нужно выбрать хотя бы один или отключить критерий по оборудования", vbExclamation, "Неправильный выбор параметров"
        Exit Function
    End If

    If txFilterName.Text = "" Then
        filterName = managId
        personal = 1
    Else
        filterName = txFilterName.Text
        personal = 0
    End If

    Dim indexRow As Integer, indexColumn As Integer
    indexRow = cbGroupByRow.ListIndex
    If indexRow = -1 Then
        indexRow = 0
    End If
    
    indexColumn = cbGroupByColumn.ListIndex
    If indexColumn = -1 Then
        indexColumn = 0
    End If
    
    filterId = prepareFilter(filterName, personal, cbGroupByRow.ItemData(indexRow), cbGroupByColumn.ItemData(indexColumn))
    
    If hasCheckedMat Then
        itemId = saveFilterItem(filterId, "materials", ckKriteriumMat.value)
        saveParamsOfTree tvMat, itemId, "klassId"
    End If
    If hasCheckedReg Then
        itemId = saveFilterItem(filterId, "regions", ckKriteriumRegion.value)
        saveParamsOfTree tvRegion, itemId, "regionId"
    End If

    If hasOborud Then
        itemId = saveFilterItem(filterId, "oborudItems", ckKriteriumOborud.value)
        Dim i As Integer
        For i = 1 To 3
            If cbOborud(i).value Then
                saveFilterParam itemId, "oborudItemId", i
            End If
        Next i
    End If
    
    If ckKriteriumNoOborud.value = 1 Then
        itemId = saveFilterItem(filterId, "noOboruds", 1)
    End If
    
    If ckStartDate.value = 1 Or ckEndDate.value = 1 Then
        itemId = saveFilterItem(filterId, "filterPeriod", 1)
        If ckStartDate.value = 1 Then
            saveFilterParam itemId, "periodStart", tbStartDate.value
        End If
        If ckEndDate.value = 1 Then
            saveFilterParam itemId, "periodEnd", tbEndDate.value
        End If
    End If
    
    submitFilter = filterId
End Function


Private Sub cmFilterAdd_Click()
    submitFilter txFilterName.Text
    cmFilterAdd.Enabled = False
    
End Sub


Private Sub cmFilterApply_Click()
    initFilter CStr(cbFilters.Text), 0

End Sub


Private Sub cmColumns_Click()
Dim columnDefs() As columnDef
    
    
    If Not columnsVisible Then
        tvColumns.Visible = Not tvColumns.Visible
    End If
    columnsVisible = False
    
    If tvColumns.Visible Then
        tvColumns.SetFocus
        initColumns columnDefs, 0, managId, , cbGroupByRow.ItemData(cbGroupByRow.ListIndex), cbGroupByColumn.ItemData(cbGroupByColumn.ListIndex)
        initColumnTree columnDefs
    End If
    'Debug.Print "cmColumns_Click, tvColumns.Visible = " & tvColumns.Visible
End Sub



Private Sub initColumnTree(ByRef headerList() As columnDef)
Dim i As Integer, anySaved As Boolean, aNode As Node


    tvColumns.Nodes.Clear
    
    For i = 0 To UBound(headerList)
        If headerList(i).hidden <> 1 Then
            Set aNode = tvColumns.Nodes.Add(, , "c" & headerList(i).columnId, headerList(i).nameRu)
            If headerList(i).saved Then
                aNode.checked = True
            End If
        End If
    Next i
    
End Sub

' Используем инверсную логику: наличие записи в таблице означает исключение столбца.

Private Sub persistColumnSelect(checked As Boolean, columnId As Integer, managId As String, byRow As Integer, byColumn As Integer)
    If Not checked Then
        sql = "insert into nHeaderColumnSelected (managId, templateId, columnId) select " _
        & "'" & managId & "'" _
        & ", a.templateId " _
        & ", " & columnId _
        & " from nAnalys a " _
        & " where a.byrow = " & byRow _
        & " and a.bycolumn = " & byColumn
    Else
        sql = "delete from nHeaderColumnSelected hc " _
        & " from nAnalys a " _
        & " where a.byrow = " & byRow _
        & " and a.bycolumn = " & byColumn _
        & " and hc.managId = '" & managId & "'" _
        & " and hc.templateId = a.templateId " _
        & " and hc.columnId = " & columnId
    End If
    
    'Debug.Print sql
    myExecute "W#ColumnSelect", sql, 0

End Sub




Private Sub tvColumns_NodeCheck(ByVal Node As MSComctlLib.Node)
    persistColumnSelect Node.checked, CInt(Mid(Node.Key, 2)), managId _
        , cbGroupByRow.ItemData(cbGroupByRow.ListIndex) _
        , cbGroupByColumn.ItemData(cbGroupByColumn.ListIndex)

    
End Sub

Private Sub tvColumns_LostFocus()
    tvColumns.Visible = False
    'Debug.Print "tvColumns_LostFocus, tvColumns.Visible = " & tvColumns.Visible
    columnsVisible = True
End Sub


Private Sub Form_Load()
Dim i As Integer

    For i = 1 To cbDateShift.ListCount - 1
        cbDateShift.ItemData(i) = i
    Next i
    
    loadKlass
    loadRegions

    Set table = myOpenRecordSet("W#72", "select * from nFilter where personal != 1", dbOpenForwardOnly)
    If table Is Nothing Then myBase.Close: End
    cbFilters.Text = ""
    
    cbFilters.AddItem ""
    While Not table.EOF
        cbFilters.AddItem "" & table!Name & ""
        table.MoveNext
    Wend
    table.Close
    
    'проинициализировать листбокс группировки по горизонтали
    initByRowList
    
    Dim currentFilterId As Integer, filterName As String, personal As Integer
    currentFilterId = getEffectiveSetting("CurrentFilter", 0)
    
    If currentFilterId > 0 And currentFilterId < cbFilters.ListCount Then
        cbFilters.ListIndex = currentFilterId
        filterName = cbFilters.List(currentFilterId)
        personal = 0
    End If
    
    If filterName = "" Then
        filterName = managId
        personal = 1
    End If
    
    initFilter filterName, personal
    
End Sub

Private Sub populateAxeList(ByRef table As Recordset, cb As ComboBox)
    
    'проинициализировать комбобокс с доступными группировками по одной из оси (по строкам или по столбцам)
    
    If table Is Nothing Then
        myBase.Close: End
    End If
    
    cb.Clear
    Dim i As Integer
    i = 0
    While Not table.EOF
        cb.AddItem table!Name_ru
        cb.ItemData(i) = table!id
        i = i + 1
        table.MoveNext
    Wend
    table.Close
    
End Sub

Private Sub initByColumnList(byRowId As Integer)
    sql = "select * from nAnalysCategory c" _
        & " where c.bycolumn_flag != 0" _
        & " and exists (select 1 from nAnalys a where a.byrow = " & byRowId & " and c.id = a.bycolumn)" _
        & " order by c.bycolumn_flag"
        
    Set table = myOpenRecordSet("W#initByRowList", sql, dbOpenForwardOnly)
    'Debug.Print sql
    populateAxeList table, cbGroupByColumn
    
End Sub


Private Sub initByRowList()
    
    sql = "select * from nAnalysCategory ac where byrow_flag = 1 " _
        & " and exists (select 1 from nAnalys a where a.byrow = ac.id and application = '" & applicationType & "')"
    
    Set table = myOpenRecordSet("W#initByRowList", sql, dbOpenForwardOnly)
    If table Is Nothing Then
        myBase.Close: End
    End If
    
    populateAxeList table, cbGroupByRow

End Sub



Private Sub tvMat_NodeCheck(ByVal Node As MSComctlLib.Node)
    checkDirtyFilterCommads
    If Not Node.Child Is Nothing Then
        setRecursiveNodeChecked Node.Child, Node.checked
    End If
End Sub


Private Function getOborudItems() As Boolean
Dim i As Integer

    For i = 1 To 3
        If cbOborud(i).value = 1 Then
            getOborudItems = True
            Exit Function
        End If
    Next i

End Function


Private Function getCheckedInTree(tView As TreeView) As Boolean
Dim currentNode As Node
Dim i As Integer

    getCheckedInTree = False
    For i = 1 To tView.Nodes.Count
        Set currentNode = tView.Nodes(i)
        If currentNode.checked Then
            getCheckedInTree = True
            Exit Function
        End If
    Next i
    
End Function


Private Sub saveParamsOfTree(tView As TreeView, itemId As Integer, paramName As String)
Dim currentNode As Node
Dim i As Integer, nCount As Integer

    nCount = tView.Nodes.Count
    For i = 1 To nCount
        Set currentNode = tView.Nodes(i)
        If currentNode.checked Then
            saveFilterParam itemId, paramName, CInt(Mid(currentNode.Key, 2))
        End If
    Next i
    
End Sub


Private Sub setRecursiveNodeChecked(ByRef root As Node, value As Boolean)
Dim NextNode As Node

    root.checked = value
    Set NextNode = root.Next
    If Not NextNode Is Nothing Then
        setRecursiveNodeChecked NextNode, value
    End If
    If Not root.Child Is Nothing Then
        setRecursiveNodeChecked root.Child, value
    End If
End Sub


Private Sub setRecursiveParent(ByRef root As Node, value As Boolean)
    root.checked = value
    If Not root.Parent Is Nothing Then
        setRecursiveParent root.Parent, value
    End If
End Sub


Private Sub expandParents(ByRef aNode As Node)
    
    If Not aNode.Parent Is Nothing Then
        aNode.Parent.Expanded = True
        expandParents aNode.Parent
    End If
End Sub


Private Sub tvRegion_NodeCheck(ByVal Node As MSComctlLib.Node)
    checkDirtyFilterCommads
    If Not Node.Child Is Nothing Then
        setRecursiveNodeChecked Node.Child, Node.checked
    End If

End Sub

Private Sub txFilterName_Change()
    cmFilterAdd.Enabled = True
End Sub

Private Function setListIndexByItemDataValue(ByRef cb As ComboBox, ByVal itemDataValue As Integer) As Boolean

Dim i As Integer

    setListIndexByItemDataValue = True
    For i = 0 To cb.ListCount - 1
        If cb.ItemData(i) = itemDataValue Then
            cb.ListIndex = i
            Exit Function
        End If
    Next i
    setListIndexByItemDataValue = False
End Function


Private Sub UpDownChange(ByVal upDirection As Integer)
Dim startDate As Date, endDate As Date, shiftPeriod As Integer
Dim dataAddInterval As String, dataAddNumber As Integer

    startDate = tbStartDate.value
    endDate = tbEndDate.value
    If cbDateShift.ListIndex = -1 Then
        cbDateShift.ListIndex = 0
    End If
    shiftPeriod = cbDateShift.ItemData(cbDateShift.ListIndex)
    
    If shiftPeriod = 1 Then
        dataAddInterval = "m"
        dataAddNumber = 6
    ElseIf shiftPeriod = 2 Then
        dataAddInterval = "q"
        dataAddNumber = 1
    ElseIf shiftPeriod = 3 Then
        dataAddInterval = "m"
        dataAddNumber = 1
    ElseIf shiftPeriod = 4 Then
        dataAddInterval = "d"
        dataAddNumber = 10
    ElseIf shiftPeriod = 5 Then
        dataAddInterval = "w"
        dataAddNumber = 1
    ElseIf shiftPeriod = 6 Then
        dataAddInterval = "d"
        dataAddNumber = 1
    ElseIf shiftPeriod = 7 Then
        dataAddInterval = "d"
        dataAddNumber = 1
    Else
        dataAddInterval = "yyyy"
        dataAddNumber = 1
    End If

    tbStartDate.value = DateAdd(dataAddInterval, dataAddNumber * upDirection, startDate)
    tbEndDate.value = DateAdd(dataAddInterval, dataAddNumber * upDirection, endDate)

End Sub

Private Sub UpDown1_DownClick()
    UpDownChange -1
End Sub

Private Sub UpDown1_UpClick()
    UpDownChange 1
End Sub

