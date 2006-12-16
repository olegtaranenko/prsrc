VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form VentureOrder 
   Caption         =   "Накладные между предприятиями"
   ClientHeight    =   5904
   ClientLeft      =   168
   ClientTop       =   816
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5904
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmRecalc 
      Caption         =   "Перерасчет"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   25
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lbBlock 
      Height          =   432
      ItemData        =   "VentureOrder.frx":0000
      Left            =   4800
      List            =   "VentureOrder.frx":000A
      TabIndex        =   24
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox ckCumulative 
      Caption         =   "Сводные"
      Height          =   255
      Left            =   4680
      TabIndex        =   23
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmExcel2 
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   9120
      TabIndex        =   22
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmExcel 
      Caption         =   "Печать в Excel"
      Height          =   315
      Left            =   4440
      TabIndex        =   21
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox tbProcent 
      Height          =   285
      Left            =   10680
      TabIndex        =   20
      Text            =   "10.0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox tbMobile2 
      Height          =   315
      Left            =   7080
      TabIndex        =   16
      Text            =   "tbMobile2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox tbMobile 
      Height          =   315
      Left            =   1800
      TabIndex        =   15
      Text            =   "tbMobile"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lbVenture 
      Height          =   240
      ItemData        =   "VentureOrder.frx":0014
      Left            =   3000
      List            =   "VentureOrder.frx":0016
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmExit 
      Caption         =   "Выход"
      Height          =   315
      Left            =   10800
      TabIndex        =   13
      Top             =   5400
      Width           =   915
   End
   Begin VB.CheckBox ckStartDate 
      Caption         =   " "
      Height          =   315
      Left            =   900
      TabIndex        =   10
      Top             =   120
      Width           =   195
   End
   Begin VB.TextBox tbStartDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Text            =   "01.11.04"
      Top             =   120
      Width           =   735
   End
   Begin VB.CheckBox ckEndDate 
      Caption         =   " "
      Height          =   315
      Left            =   2460
      TabIndex        =   8
      Top             =   120
      Width           =   200
   End
   Begin VB.TextBox tbEndDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmLoad 
      Caption         =   "Загрузить"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmAdd 
      Caption         =   "Добавить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmDel2 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7740
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmDel 
      Caption         =   "Удалить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   5400
      Width           =   915
   End
   Begin VB.CommandButton cmAdd2 
      Caption         =   "Добавить"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6300
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5775
      _ExtentX        =   10181
      _ExtentY        =   8065
      _Version        =   393216
      MergeCells      =   2
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4575
      Left            =   6000
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   5835
      _ExtentX        =   10287
      _ExtentY        =   8065
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label laProcent 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Процент (для вновь создаваемых зачетов)"
      Height          =   195
      Left            =   7335
      TabIndex        =   19
      Top             =   180
      Width           =   3270
   End
   Begin VB.Label laGrid 
      Caption         =   "Реестр документов"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label laGrid2 
      Height          =   195
      Left            =   6000
      TabIndex        =   17
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label laPeriod 
      Caption         =   "Период с  "
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   180
      Width           =   795
   End
   Begin VB.Label laPo 
      Caption         =   "по"
      Height          =   195
      Left            =   2100
      TabIndex        =   11
      Top             =   180
      Width           =   180
   End
   Begin VB.Menu mpGrid 
      Caption         =   "mpGrid"
      Visible         =   0   'False
      Begin VB.Menu mnRecalc 
         Caption         =   "Пересчитать"
      End
   End
   Begin VB.Menu mpGrid2 
      Caption         =   "mpGrid2"
      Visible         =   0   'False
      Begin VB.Menu mnNomHistory 
         Caption         =   "Карточка движения"
      End
      Begin VB.Menu mnAddToHistory 
         Caption         =   "Добавить к карточке движения"
      End
   End
   Begin VB.Menu mpGrid3 
      Caption         =   "mpGrid3"
      Visible         =   0   'False
      Begin VB.Menu mnFilter 
         Caption         =   "Наложить фильтр"
      End
   End
End
Attribute VB_Name = "VentureOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public isLoad As Boolean
Dim mousCol As Long, mousRow As Long
Dim mousCol2 As Long, mousRow2 As Long
Dim oldHeight As Integer, oldWidth As Integer, buttonWidth As Integer ' нач размер формы
Dim quantity  As Long
Dim sum As Single
Dim addedId As Integer

Dim ventureIds() As Integer
Dim ventureRests() As Variant
Dim fromArr() As Variant
Dim venturePairFilter() As Variant

Dim defaultVentureIndex As Integer

Dim quantity2  As Long
Private trigger As Boolean
Private trigger2 As Boolean
Private procentChanged As Boolean

Const dcIdJmat = 0
Const dcDate = 1
Const dcNumDoc = 2 'id
Const dcNDate = 3
Const dcSour = 4
Const dcDest = 5
Const dcTermFrom = 6
Const dcTermTo = 7
Const dcSumma = 8
Const dcProcent = 9
Const dcComtec = 10
Const dcNote = 11
Const dcInvalid = 12
Const dcCount = 13

'Grid2
Const dnId = 0
Const dnIdMat = 1
Const dnNomNom = 2
Const dnNomName = 3
Const dnEdIzm = 4
Const dnQuant = 5
Const dnCosted = 6
Const dnSumma = 7
Const dnCostFact = 8

Sub lbHide()
    lbVenture.Visible = False
    lbBlock.Visible = False
    tbMobile.Visible = False
    Grid.Enabled = True
    On Error Resume Next
    Grid.SetFocus
    Grid_EnterCell
End Sub
Sub lbHide2()
    tbMobile2.Visible = False
    Grid2.Enabled = True
    On Error Resume Next
    Grid2.SetFocus
    Grid2_EnterCell
End Sub

Function loadDocNomenk(ByVal sdvId As String) As Boolean
Dim msgOst As String, docProcent As Single, summa_fact As Single, fullNomName As String

    loadDocNomenk = True ' не надо отката - пока
    msgOst = ""
    Me.MousePointer = flexHourglass
    Grid2.Visible = False

    gDocDate = Grid.TextMatrix(mousRow, dcDate)
    laGrid2.Caption = "Номенклатура по документу '" & numDoc & "' (всего " & Grid.TextMatrix(mousRow, dcCount) & " поз.)"

    quantity2 = 0
    clearGrid Grid2
    If Grid.TextMatrix(mousRow, dcProcent) = "" Then
        docProcent = 0
    Else
        docProcent = CSng(Grid.TextMatrix(mousRow, dcProcent))
    End If
    sql = _
        " select cod, m.id, m.nomnom, round(isnull(m.quant, 0) / n.perlist, 2) as quant, m.id_mat" _
        & "     , isnull(m.costed, 0.00) as costed , n.nomname, n.ed_izmer2" _
        & "     , round(n.cena1, 2) as cost_fact, n.size" _
        & " from sdmcventure m" _
        & " left join sguidenomenk n on m.nomnom = n.nomnom" _
        & " where sdv_id = " & sdvId _
        & " order by n.nomname"

    Set tbNomenk = myOpenRecordSet("##118.1", sql, dbOpenForwardOnly)
    If tbNomenk Is Nothing Then Exit Function
    If Not tbNomenk.BOF Then
        While Not tbNomenk.EOF
            quantity2 = quantity2 + 1
            Grid2.TextMatrix(quantity2, dnId) = tbNomenk!id
            If Not IsNull(tbNomenk!id_mat) Then
                Grid2.TextMatrix(quantity2, dnIdMat) = tbNomenk!id_mat
            End If
            Grid2.TextMatrix(quantity2, dnNomName) = tbNomenk!nomName
            If Not IsNull(tbNomenk!cod) Then
                fullNomName = tbNomenk!cod & " "
            End If
            fullNomName = fullNomName & tbNomenk!nomName
            If Not IsNull(tbNomenk!Size) Then
                fullNomName = fullNomName & " " & tbNomenk!Size
            End If
          
            Grid2.TextMatrix(quantity2, dnNomName) = fullNomName
            Grid2.TextMatrix(quantity2, dnEdIzm) = tbNomenk!ed_Izmer2
            Grid2.TextMatrix(quantity2, dnNomNom) = tbNomenk!nomnom
            Grid2.TextMatrix(quantity2, dnQuant) = tbNomenk!quant
            Grid2.TextMatrix(quantity2, dnCosted) = Round(tbNomenk!costed, 2)
            Grid2.TextMatrix(quantity2, dnSumma) = Round(tbNomenk!quant * tbNomenk!costed, 2)
            summa_fact = Round(tbNomenk!cost_fact * tbNomenk!quant * (1 + docProcent / 100), 2)
            If Abs(tbNomenk!cost_fact * tbNomenk!quant * (1 + docProcent / 100) - tbNomenk!quant * tbNomenk!costed) > 0.01 Then
                Grid2.col = dnNomName
                
                Grid2.row = Grid2.Rows - 1
                Grid2.CellForeColor = vbRed
            End If
            
            If Not IsNull(tbNomenk!cost_fact) Then
                Grid2.TextMatrix(quantity2, dnCostFact) = tbNomenk!cost_fact
            Else
                Grid2.TextMatrix(quantity2, dnCostFact) = "?"
            End If

            Grid2.AddItem ""
            tbNomenk.MoveNext
        Wend
        Grid2.RemoveItem quantity2 + 1
    End If
    Grid2.col = 0
    tbNomenk.Close
    Grid2.Visible = True
    Me.MousePointer = flexDefault
End Function


Sub loadVentureOrders(Optional reg As String)

Dim prevRow As Long, rowIndex As Long
Dim dstRow As Long, i As Integer, j As Integer, k As Integer, itogTxt As String, rowIsFiltered As Integer, rowisAdded As Integer
Dim filterIsApplied As Integer


    prevRow = -1
    Grid.Visible = False
    cmRecalc.Enabled = False
    
    Me.MousePointer = flexHourglass
    If reg = "" Then
        gridIsLoad = False
        quantity = 0
        clearGrid Grid
    End If
    
    If reg = "" Then
        strWhere = getWhereByDateBoxes(Me, "v.nDate", CDate("01.10.2000"))
    ElseIf reg = "add" And addedId <> 0 Then
        strWhere = " v.id = " & addedId
        addedId = 0
    ElseIf reg = "update" Then
        strWhere = " v.id = " & Grid.TextMatrix(mousRow, dcNumDoc)
    End If

    
    sql = " select " _
        & "     v.id" _
        & "   , v.xDate, v.nDate, v.termFrom, v.termTo" _
        & "   , isnull(v.id_jmat, cum.id_jmat) as id_jmat" _
        & "   , d.ventureName as dest" _
        & "   , s.ventureName as sour" _
        & "   , v.procent " _
        & "   , v.note " _
        & "   , v.Invalid " _
        & "   , sum(round(isnull(m.quant, 0) * isnull(m.costed, 0) / isnull(n.perlist, 1), 2)) as summa" _
        & "   , count(*) as cnt " _
        & "   , s.ventureId as sourId" _
        & "   , d.ventureId as destId" _
        & " from sdocsventure v" _

        
    sql = sql _
        & "     left join sDmcVenture m on m.sdv_id = v.id " _
        & "     left join sguidenomenk n on n.nomnom = m.nomnom" _
        & "     left join guideVenture d on d.ventureId = v.dstVentureId" _
        & "     left join guideVenture s on s.ventureId = v.srcVentureid" _
        & "     left join sdocsventure cum on cum.id = v.cumulative_id" _
        & "     where v.cumulative_id is "
        
    If ckCumulative.value = 0 Then
        sql = sql & " not "
    End If
        
    sql = sql & " null "
    If strWhere <> "" Then
        sql = sql & " and " & strWhere
    End If
        
    sql = sql & " group by " _
        & "    v.id, v.xDate, v.nDate, v.termFrom, v.termTo" _
        & "   , isnull(v.id_jmat, cum.id_jmat), dest, sour, v.procent, v.note" _
        & "   , v.Invalid, sourId, destId " _
        & " order by v.nDate, dest, sour"

 
    'Debug.Print sql
 
    Set tbDocs = myOpenRecordSet("##176.2", sql, dbOpenForwardOnly)
    If tbDocs Is Nothing Then End
    
    resetRests
    
    If reg <> "update" Then
        rowisAdded = 1
    Else
        rowIndex = mousRow
    End If
    
    filterIsApplied = 0
    
    For i = 1 To UBound(ventureIds)
        For j = 1 To UBound(ventureIds)
            If venturePairFilter(i)(j) = 1 Then
                filterIsApplied = 1
            End If
            
        Next j
    Next i
    
    If Not tbDocs.BOF Then
        While Not tbDocs.EOF
            ' Сюда добавить проверку на то, что строка не отфильтровывается
            i = vlIndex(tbDocs!sourId)
            j = vlIndex(tbDocs!destId)
            If filterIsApplied = 1 And venturePairFilter(i)(j) = 0 Then
                rowIsFiltered = 1
            Else
                rowIsFiltered = 0
            End If
            
            
            If rowIsFiltered = 0 Then
                If rowisAdded = 1 Then
                    Grid.AddItem ""
                    quantity = quantity + 1
                    rowIndex = quantity
                End If
                LoadDate Grid, rowIndex, dcDate, tbDocs!xDate, "dd.mm.yy"
                If Not IsNull(tbDocs!nDate) Then
                    LoadDate Grid, rowIndex, dcNDate, tbDocs!nDate, "dd.mm.yy"
                End If
                If Not IsNull(tbDocs!termFrom) Then
                    LoadDate Grid, rowIndex, dcTermFrom, tbDocs!termFrom, "dd.mm.yy"
                Else
                    Grid.TextMatrix(rowIndex, dcTermFrom) = "от начала"
                End If
                
                If Not IsNull(tbDocs!termTo) Then
                    LoadDate Grid, rowIndex, dcTermTo, tbDocs!termTo, "dd.mm.yy"
                Else
                    Grid.TextMatrix(rowIndex, dcTermTo) = "наст.вр"
                End If
                
                If Not IsNull(tbDocs!id_Jmat) Then
                    Grid.TextMatrix(rowIndex, dcIdJmat) = tbDocs!id_Jmat
                    Grid.TextMatrix(rowIndex, dcComtec) = "Да"
                End If
                Grid.TextMatrix(rowIndex, dcNumDoc) = tbDocs!id
                If Not IsNull(tbDocs!Sour) Then
                    Grid.TextMatrix(rowIndex, dcSour) = tbDocs!Sour
                End If
                If Not IsNull(tbDocs!Dest) Then
                    Grid.TextMatrix(rowIndex, dcDest) = tbDocs!Dest
                End If
                
                If Not IsNull(tbDocs!procent) Then
                    Grid.TextMatrix(rowIndex, dcProcent) = tbDocs!procent
                End If
                
                
                If Not IsNull(tbDocs!note) Then
                    Grid.TextMatrix(rowIndex, dcNote) = tbDocs!note
                End If
                
                If Not IsNull(tbDocs!Invalid) Then
                    Grid.TextMatrix(rowIndex, dcInvalid) = tbDocs!Invalid
                    If tbDocs!Invalid = "1" Then
                        cmRecalc.Enabled = True
                    End If
                End If
                
                Grid.TextMatrix(rowIndex, dcSumma) = Round(tbDocs!summa, 2)
                Grid.TextMatrix(rowIndex, dcCount) = tbDocs!cnt
                
            End If
            ventureRests(i)(j) = ventureRests(i)(j) + tbDocs!summa
            tbDocs.MoveNext
        Wend
        
    End If
    tbDocs.Close
'    rowViem quantity, Grid
    
    If quantity > 0 Then
        If reg = "" Then Grid.RemoveItem quantity + 1
        Grid.col = 1
        gridIsLoad = True
        If reg <> "update" Then Grid.row = quantity
        Grid.col = 2      'вызов loadDocNomenk

        On Error Resume Next
        Grid.SetFocus
        cmDel.Enabled = True
        Grid2.Visible = True
        cmAdd2.Enabled = True
        If reg = "" Then
            
            For i = 1 To UBound(ventureIds)
                For j = 1 To UBound(ventureIds)
                    If i <> j Then
'                        venturePairFilter(i)(j) = 1
                        Grid.AddItem ""
                        dstRow = Grid.Rows - 1
                        Grid.MergeRow(dstRow) = True
                        Grid.row = dstRow
                        
                        For k = 1 To Grid.Cols - 1
                            Grid.col = k
                            If k < dcSumma Then
                                itogTxt = lbVenture.List(i - 1) & " => " & lbVenture.List(j - 1)
                            Else
                                itogTxt = Format(ventureRests(i)(j), "#0.00")
                            End If
                            
                            Grid.Text = itogTxt
                            Grid.CellAlignment = flexAlignRightCenter
                            Grid.CellFontBold = True
                            If venturePairFilter(i)(j) = 1 Then
                                Grid.CellForeColor = vbRed
                            End If
                        Next k
'                        Grid.TextMatrix(dstRow, dcSumma) = Format(ventureRests(i)(j), "#0.00")

                    End If
                Next j
            Next i
        End If
    
    Else
        cmDel.Enabled = False
        Grid2.Visible = False
        cmAdd2.Enabled = False
    End If

    Grid.Visible = True
    Me.MousePointer = flexDefault
    gridIsLoad = True
    
End Sub

Private Sub resetRests()
Dim i As Integer, j As Integer

    For i = 1 To UBound(ventureIds)
        For j = 1 To UBound(ventureIds)
            ventureRests(i)(j) = CSng(0)
'            venturePairFilter(i)(j) = 0
        Next j
    Next i
    
End Sub

Private Function vlIndex(ventureId)
Dim i As Integer

    For i = 1 To UBound(ventureIds)
        If ventureId = ventureIds(i) Then
            vlIndex = i
            Exit Function
        End If
    Next i
    vlIndex = defaultVentureIndex
    
End Function


Private Sub ckCumulative_Click()
    If Not gridIsLoad Then Exit Sub
    If ckCumulative.value = 1 Then
        Grid.ColWidth(dcComtec) = 400
    Else
        Grid.ColWidth(dcComtec) = 0
    End If
    loadVentureOrders
End Sub

Private Sub ckEndDate_Click()
    If ckEndDate.value = 1 Then
        tbEndDate.Enabled = True
    Else
        tbEndDate.Enabled = False
    End If
    
    If tbEndDate.Text = "" Then
        tbEndDate.Text = Format(Now, "dd.mm.yy")
    End If
    
    If ckEndDate.value = 1 And ckStartDate.value = 1 Then
        cmAdd.Enabled = True
    Else
        cmAdd.Enabled = False
    End If
End Sub

Private Sub ckStartDate_Click()
    If ckStartDate.value = 1 Then
        tbStartDate.Enabled = True
    Else
        tbStartDate.Enabled = False
    End If
    
    If ckEndDate.value = 1 And ckStartDate.value = 1 Then
        cmAdd.Enabled = True
    Else
        cmAdd.Enabled = False
    End If
End Sub

Private Sub cmAdd_Click()
Dim queryTimeout As Variant

    On Error GoTo finally
    queryTimeout = myBase.queryTimeout

    If tbStartDate.Enabled = True And tbEndDate.Enabled = True Then
        If isDateEmpty(tbStartDate) And isDateEmpty(tbEndDate) Then
            If MsgBox("Нажмите ОК, если вы действительно хотите создать накладные взаимозачета за период с " _
                & Format(tbStartDate.Text, "dd.mm.yyyy") _
                & " по " & Format(tbEndDate.Text, "dd.mm.yyyy") _
                , vbOK Or vbDefaultButton2, "Подтвердите") = vbOK _
            Then
                sql = "call ivo_generate(" _
                    & tbProcent.Text _
                    & ", convert(date, '" & Format(tbStartDate.Text, "yyyymmdd") & "')" _
                    & ", convert(date, '" & Format(tbEndDate.Text, "yyyymmdd") & "')"
    
                sql = sql _
                    & " ) "
                
                myBase.queryTimeout = 600
                If myExecute("##ivo_generate", sql, 0) = 0 Then
                    wrkDefault.CommitTrans
                    loadVentureOrders
                Else
                    wrkDefault.Rollback
                End If
            End If
            
        Else
            Exit Sub
        End If
    End If
    
finally:
    myBase.queryTimeout = queryTimeout

End Sub

Private Sub cmDel_Click()
Dim prevCol As Long
Dim delSumma As Single

    If MsgBox("Удалить документ № '" & Grid.TextMatrix(mousRow, dcNumDoc) & _
    "', Вы уверены?", vbYesNo Or vbDefaultButton2, "Подтвердите удаление") _
    = vbNo Then GoTo EN1
    
    delSumma = CSng(Grid.TextMatrix(mousRow, dcSumma))
    
    sql = "DELETE From sDocsVenture WHERE id = " & Grid.TextMatrix(mousRow, dcNumDoc)
    'MsgBox sql
    If myExecute("##del sDocVenture", sql) = 0 Then
        loadVentureOrders
        wrkDefault.CommitTrans
    Else
        wrkDefault.Rollback
    End If
EN1:

End Sub


Private Sub cmDel2_Click()
Dim deletedSum As Single

    If MsgBox("Нажмите OK, если вы действительно хотите удалить позицию ", vbOKCancel, "Подтверждение") <> vbOK Then Exit Sub
    deletedSum = CSng(Grid2.TextMatrix(mousRow2, dnSumma))
    If vo_deleteNomnom(Grid2.TextMatrix(mousRow2, dnNomNom), Grid.TextMatrix(mousRow, dcNumDoc)) Then
        loadDocNomenk Grid.TextMatrix(mousRow, dcNumDoc)
        On Error Resume Next
        Grid2.SetFocus
        cmLoad.Caption = "Обновить"
        ' откорректировать сумму по накладной
        ' сохраняем, какая сумма у накладной была
        ' получаем сумму по удаленной позиции
        ' и итоги за период по предприятию
    End If
    
End Sub

Private Sub cmExcel_Click()
    GridToExcel Grid, "Накладные взаимозачета за период с "
End Sub

Private Sub cmExcel2_Click()
    GridToExcel Grid2, "Содержание накладной взаимозачета № " & " за "
End Sub

Private Sub cmExit_Click()
    Unload Me
End Sub

Private Sub cmLoad_Click()
    prevRow = -1
    loadVentureOrders
    cmLoad.Caption = "Загрузить"
End Sub


Private Sub cmRecalc_Click()
    sql = "call ivo_validate( " & tbProcent.Text & ")"
    myBase.queryTimeout = 600
    Me.MousePointer = flexHourglass
    If myExecute("##cmRecalc_click", sql, 0) = 0 Then
        wrkDefault.CommitTrans
        loadVentureOrders
    Else
        wrkDefault.Rollback
    End If
    Me.MousePointer = flexDefault
    cmLoad_Click
    cmRecalc.Enabled = False
End Sub

Private Sub Form_Load()
Dim sz As Integer, docProcent As Single, i As Integer


    buttonWidth = cmAdd.Width
    oldHeight = Me.Height
    oldWidth = Me.Width
    isLoad = True
    
    sql = "select ivo_procent from system"
    byErrSqlGetValues "##333.1", sql, docProcent
    
    tbProcent.Text = docProcent
    

    sql = "SELECT ventureId, ventureName, rusAbbrev, s.id_analytic_default " _
        & " From GuideVenture v" _
        & " left join system s on v.id_analytic = s.id_analytic_default " _
        & "WHERE id_analytic is not null order by ventureName"
    Set Table = myOpenRecordSet("##144", sql, dbOpenForwardOnly)
    If Table Is Nothing Then End
    ReDim ventureIds(0)
    ReDim ventureRests(0)
    ReDim venturePairFilter(0)
    
    While Not Table.EOF
        lbVenture.AddItem Table!ventureName
        lbVenture.ItemData(lbVenture.ListCount - 1) = Table!ventureId
        sz = UBound(ventureIds)
        ReDim Preserve ventureIds(sz + 1)
        ventureIds(sz + 1) = Table!ventureId

        If Not IsNull(Table!id_analytic_default) Then
            defaultVentureIndex = sz + 1
        End If
        
        Table.MoveNext
    Wend
    
    ReDim Preserve ventureRests(sz + 1)
    
    ReDim Preserve venturePairFilter(sz + 1)
    
    For i = 1 To UBound(ventureIds)
        fromArr = Array()
        ReDim Preserve fromArr(sz + 1)
        ventureRests(i) = fromArr
        
        fromArr = Array()
        ReDim Preserve fromArr(sz + 1)
        venturePairFilter(i) = fromArr
        
    Next i
        
    Table.Close
    lbVenture.Height = lbVenture.ListCount * 205 + 50

    Grid.FormatString = "|<Сформировано|<№ Док-та|Дата накл.|<Откуда|<Куда|<Учет: с|<   по|>Сумма|%|Комтех|<Примечание||"
    Grid.ColWidth(dcIdJmat) = 0
    Grid.ColWidth(dcDate) = 0
    Grid.ColWidth(dcNumDoc) = 500
    'Grid.ColWidth(dcM) = 300
    Grid.ColWidth(dcSour) = 500
    Grid.ColWidth(dcDest) = 500
    Grid.ColWidth(dcNDate) = 800
    Grid.ColWidth(dcTermFrom) = 800
    Grid.ColWidth(dcTermTo) = 800
    Grid.ColWidth(dcSumma) = 900
    Grid.ColWidth(dcProcent) = 400
    Grid.ColWidth(dcComtec) = 0
    Grid.ColWidth(dcNote) = 1530
    Grid.ColWidth(dcInvalid) = 0
    Grid.ColWidth(dcCount) = 0


    Grid2.FormatString = "||<Номер|<Название|<Ед.измерения|Кол-во|>По цене|Сумма|<Цена Факт."
    Grid2.ColWidth(dnId) = 0
    Grid2.ColWidth(dnIdMat) = 0
    Grid2.ColWidth(dnNomNom) = 945
    Grid2.ColWidth(dnNomName) = 2400
    Grid2.ColWidth(dnEdIzm) = 435
    Grid2.ColWidth(dnQuant) = 435
    Grid2.ColWidth(dnCosted) = 660
    Grid2.ColWidth(dnSumma) = 800
    Grid2.ColWidth(dnCostFact) = 660

End Sub

Private Sub Form_Resize()
Dim h As Integer, w As Integer

    If WindowState = vbMinimized Then Exit Sub
    On Error Resume Next
    h = Me.Height - oldHeight
    oldHeight = Me.Height
    w = Me.Width - oldWidth
    oldWidth = Me.Width
    Grid.Height = Grid.Height + h
    Grid.Width = Grid.Width + w / 2
    laGrid.Width = Grid.Width
    ckCumulative.Left = Grid.Left + Grid.Width - ckCumulative.Width

    Grid2.Height = Grid2.Height + h
    Grid2.Width = Grid2.Width + w / 2
    Grid2.Left = Grid2.Left + w / 2
    laGrid2.Left = Grid2.Left
    laGrid2.Width = Grid2.Width
    laProcent.Left = Grid2.Left
    tbProcent.Left = laProcent.Left + laProcent.Width + 50
    cmRecalc.Left = tbProcent.Left + tbProcent.Width + 50

    cmLoad.Top = cmLoad.Top + h
    cmAdd.Top = cmLoad.Top
    cmDel.Top = cmLoad.Top
    cmExcel.Top = cmLoad.Top
    cmRecalc.Top = tbProcent.Top ' - tbProcent.Height + cmRecalc.Height
    
    cmAdd2.Top = cmLoad.Top
    cmDel2.Top = cmLoad.Top
    cmExit.Top = cmLoad.Top
    cmExcel2.Top = cmLoad.Top

    cmLoad.Width = buttonWidth
    cmAdd.Width = buttonWidth
    cmDel.Width = buttonWidth
    cmExcel.Width = buttonWidth * 1.2
    
    cmAdd2.Width = buttonWidth
    cmDel2.Width = buttonWidth
    cmExit.Width = buttonWidth
    cmExcel2.Width = buttonWidth * 1.2
    
    
resize:
    cmLoad.Left = Grid.Left
    cmAdd.Left = cmLoad.Left + cmLoad.Width + 50
    cmDel.Left = Grid.Left + Grid.Width - cmDel.Width
    cmExcel.Left = cmDel.Left - cmExcel.Width - 50
    
    cmAdd2.Left = Grid2.Left
    cmDel2.Left = cmAdd2.Left + cmAdd2.Width + 50
    cmExit.Left = Grid2.Left + Grid.Width - cmExit.Width
    cmExcel2.Left = cmExit.Left - cmExcel2.Width - 50
    
    If cmExcel.Left < cmAdd.Left + cmAdd.Width + 20 Then
        cmLoad.Width = cmLoad.Width * 0.9
        cmAdd.Width = cmAdd.Width * 0.9
        cmDel.Width = cmDel.Width * 0.9
        cmExcel.Width = cmExcel.Width * 0.9
        
        
        cmAdd2.Width = cmAdd2.Width * 0.9
        cmDel2.Width = cmDel2.Width * 0.9
        cmExit.Width = cmExit.Width * 0.9
        cmExcel2.Width = cmExcel2.Width * 0.9
        GoTo resize
    Else
        
    End If
    
    
End Sub

Private Sub Grid_Click()
    mousCol = Grid.MouseCol
    mousRow = Grid.MouseRow
    If mousRow = 0 And Not (mousCol = dcProcent) Then
        Grid.CellBackColor = Grid.BackColor
    '    SortCol Grid2, mousCol
        trigger = Not trigger
        Grid.Sort = 9
        Grid.row = 1    ' только чтобы снять выделение
    End If
End Sub

Private Sub Grid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Dim val1 As Variant, val2 As Variant

    
    If Not IsNumeric(Grid.TextMatrix(Row1, dcNumDoc)) And Not IsNumeric(Grid.TextMatrix(Row2, dcNumDoc)) Then
        Cmp = 0: Exit Sub
    End If

    If Not IsNumeric(Grid.TextMatrix(Row1, dcNumDoc)) Then
        Cmp = 1: Exit Sub
    End If
    If Not IsNumeric(Grid.TextMatrix(Row2, dcNumDoc)) Then
        Cmp = -1: Exit Sub
    End If
    
    
    If mousCol = dcDate Or mousCol = dcNDate Then
        val1 = CDate(Grid.TextMatrix(Row1, mousCol))
        val2 = CDate(Grid.TextMatrix(Row2, mousCol))
    ElseIf mousCol = dcProcent Or mousCol = dcSumma Then
        val1 = CSng(Grid.TextMatrix(Row1, mousCol))
        val2 = CSng(Grid.TextMatrix(Row2, mousCol))
    ElseIf mousCol = dcNumDoc Then
        val1 = CInt(Grid.TextMatrix(Row1, mousCol))
        val2 = CInt(Grid.TextMatrix(Row2, mousCol))
    Else
        val1 = Grid.TextMatrix(Row1, mousCol)
        val2 = Grid.TextMatrix(Row2, mousCol)
    End If
        
    If val1 < val2 Then
        Cmp = -1
    ElseIf val1 > val2 Then
        Cmp = 1
    Else
        Cmp = 0
    End If
    If (trigger) Then Cmp = -Cmp

End Sub

Private Sub Grid_DblClick()
    If Grid.CellBackColor = &H88FF88 Then
        If mousCol = dcSour Then
            listBoxInGridCell lbVenture, Grid, "select"
        ElseIf mousCol = dcComtec Then
            listBoxInGridCell lbBlock, Grid, Grid.TextMatrix(mousRow, dcComtec)
        ElseIf mousCol = dcDest Then
            listBoxInGridCell lbVenture, Grid, "select"
        ElseIf mousCol = dcTermFrom Or mousCol = dcTermTo Then
            If MsgBox("Изменение периода действия накладной приведет к перерасчету ее содержания. " & _
            "Нажмите <Да> если вы действительно хотите это сделать. " _
            , vbYesNo Or vbDefaultButton2, "Подтвердите изменение " & _
            "Даты!") = vbYes Then textBoxInGridCell tbMobile, Grid
        ElseIf mousCol = dcNote Then
            textBoxInGridCell tbMobile, Grid
        Else
            textBoxInGridCell tbMobile, Grid
        End If
    End If
    
End Sub

Private Sub Grid_EnterCell()
    If quantity = 0 Or Not gridIsLoad Then
        cmDel.Enabled = False
        Exit Sub
    End If
    
    mousRow = Grid.row
    mousCol = Grid.col
    If IsNumeric(Grid.TextMatrix(mousRow, dcNumDoc)) Then
        numDoc = CLng(Grid.TextMatrix(mousRow, dcNumDoc))
    Else
        clearGrid Grid2
        quantity2 = 0
        numDoc = 0
        laGrid2.Caption = ""
        prevRow = -1
    End If
    If prevRow <> mousRow And gridIsLoad And numDoc <> 0 Then
        prevRow = mousRow
        loadDocNomenk Grid.TextMatrix(mousRow, dcNumDoc)
    End If
    If mousCol = 0 Then Exit Sub
    
    
    If _
        (Grid.TextMatrix(mousRow, dcComtec) = "" Or mousCol = dcNote Or mousCol = dcComtec) And _
        (mousRow <= quantity) And _
        (mousCol >= dcProcent Or ckCumulative.value <> 1) And _
        (mousCol >= dcNDate _
            And mousCol <> dcSumma _
            And (Grid.TextMatrix(mousRow, dcIdJmat) = "" Or mousCol >= dcProcent) _
        ) _
    Then
        Grid.CellBackColor = &H88FF88
    Else
        Grid.CellBackColor = vbYellow
    End If

exit_sub:
    If Grid.TextMatrix(mousRow, dcIdJmat) = "" And mousRow <= quantity Then
        cmDel.Enabled = True
    Else
        cmDel.Enabled = False
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


Sub searchPair(ByVal row As Long, ByRef indexFrom As Integer, ByRef indexTo As Integer)
Dim i As Integer, j As Integer, curIndex As Long

    indexFrom = -1: indexTo = -1
    curIndex = quantity + 1
    For i = 1 To UBound(ventureIds)
        For j = 1 To UBound(ventureIds)
            If i <> j Then
                If curIndex = row Then
                    indexFrom = i: indexTo = j
                    Exit Sub
                End If
                curIndex = curIndex + 1
            End If
        Next j
    Next i
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim indexFrom As Integer, indexTo As Integer

    If Button = vbRightButton Then
        If Grid.MouseRow > 0 Then
            Grid.row = Grid.MouseRow
            Grid.col = Grid.MouseCol
            mousRow = Grid.MouseRow
            mousCol = Grid.MouseCol
            
            If IsNumeric(Grid.TextMatrix(mousRow, dcNumDoc)) Then
                Me.PopupMenu mpGrid  'Пересчитать
            Else
                 'Наложить или снять фильтр по паре
                Call searchPair(Grid.MouseRow, indexFrom, indexTo)
                If venturePairFilter(indexFrom)(indexTo) = 0 Then
                    mnFilter.Caption = "Наложить"
                Else
                    mnFilter.Caption = "Снять"
                End If
                mnFilter.Caption = mnFilter.Caption & " фильтр"
                Me.PopupMenu mpGrid3
            End If
        End If
    End If
End Sub

Private Sub Grid2_Click()
    mousCol2 = Grid2.MouseCol
    mousRow2 = Grid2.MouseRow
    If mousRow2 = 0 Then
        Grid2.CellBackColor = Grid2.BackColor
    '    SortCol Grid2, mousCol
        trigger2 = Not trigger2
        Grid2.Sort = 9
        Grid2_LeaveCell    ' только чтобы снять выделение
    End If
End Sub

Private Sub Grid2_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Dim val1 As Variant, val2 As Variant
    
    If mousCol2 = dnQuant Or mousCol2 = dnCosted Or mousCol2 = dnSumma Or mousCol2 = dnCostFact Then
        val1 = CSng(Grid2.TextMatrix(Row1, mousCol2))
        val2 = CSng(Grid2.TextMatrix(Row2, mousCol2))
    Else
        val1 = Grid2.TextMatrix(Row1, mousCol2)
        val2 = Grid2.TextMatrix(Row2, mousCol2)
    End If
        
    If val1 < val2 Then
        Cmp = -1
    ElseIf val1 > val2 Then
        Cmp = 1
    Else
        Cmp = 0
    End If
    If (trigger2) Then Cmp = -Cmp

End Sub

Private Sub Grid2_DblClick()
    If Grid2.CellBackColor = &H88FF88 Then
        If mousCol2 = dnQuant Or mousCol2 = dnSumma Or mousCol2 = dnCosted Then
            textBoxInGridCell tbMobile2, Grid2
        End If
    End If
End Sub

Private Sub Grid2_EnterCell()
    
    If Grid2.col = 0 Or quantity2 = 0 Or Not gridIsLoad Then Exit Sub
    mousRow2 = Grid2.row
    mousCol2 = Grid2.col
    
    'dmcId = Grid2.TextMatrix(mousRow, dnId)
    If mousCol2 = 0 Then Exit Sub
    
    If (mousCol2 = dnCosted _
        Or mousCol2 = dnSumma _
        Or mousCol2 = dnQuant) _
        And (Grid.TextMatrix(mousRow, dcComtec) = "") _
    Then
        Grid2.CellBackColor = &H88FF88
    Else
        Grid2.CellBackColor = vbYellow
    End If
    If ckCumulative.value = 1 And mousCol2 = dnQuant Then
        Grid2.CellBackColor = vbYellow
    End If
End Sub

Private Sub Grid2_GotFocus()
    cmDel2.Enabled = True
End Sub

Private Sub Grid2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Grid2_DblClick
    End If
    
End Sub

Private Sub Grid2_LeaveCell()
    If Grid2.col > 1 Then
        Grid2.CellBackColor = Grid2.BackColor
    End If
'    mousRow2 = 0
End Sub

Private Sub Grid2_LostFocus()
    'cmDel2.Enabled = False
    Grid2_LeaveCell
End Sub

Private Sub Grid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If Grid2.MouseRow > 0 Then
            If Grid2.row = Grid2.RowSel Then
                Grid2.row = Grid2.MouseRow
                Grid2.col = Grid2.MouseCol
            End If
            If quantity2 > 0 Then
                Me.PopupMenu mpGrid2
            End If
        End If
    End If
End Sub

Private Sub lbBlock_DblClick()
    If lbBlock.Text = Grid.TextMatrix(mousRow, dcComtec) Then
        lbHide
        Exit Sub
    End If
    If Trim(lbBlock.Text) = "Да" Then
        If MsgBox("Нажмите OK, если вы действительно хотите создать взаимозачеты в базе Комтех.", vbOK Or vbDefaultButton2) <> vbOK Then
            lbHide
            Exit Sub
        End If
        Me.MousePointer = flexHourglass
        sql = "call ivo_to_comtex ( " _
            & Grid.TextMatrix(mousRow, dcNumDoc) _
            & ")"
            
'        Debug.Print sql
        
        If myExecute("##126.1", sql, 0) <> 0 Then
            Grid.Text = lbBlock.Text
        End If
    
        Me.MousePointer = flexDefault
    Else
        If MsgBox("Нажмите OK, если вы действительно хотите удалить зачеты из базы Комтех.", vbOK Or vbDefaultButton2) <> vbOK Then
            lbHide
            Exit Sub
        End If
        Me.MousePointer = flexHourglass
    
        sql = "call ivo_comtex_delete ( " _
            & Grid.TextMatrix(mousRow, dcNumDoc) _
            & ")"
            
        If myExecute("##126.2", sql, 0) = 0 Then
            Grid.Text = lbBlock.Text
            'loadVentureOrders "update"
        End If
    
        Me.MousePointer = flexDefault
    End If
    loadVentureOrders "update"
    lbHide
End Sub

Private Sub lbBlock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lbBlock_DblClick
    ElseIf KeyCode = vbKeyEscape Then
        lbHide
    End If

End Sub

Private Sub lbVenture_DblClick()
Dim partName As String

    If mousCol = dcSour Then
        partName = "src"
    Else
        partName = "dst"
    End If

    sql = "UPDATE sDocsVenture SET " & partName & "VentureId = " _
        & lbVenture.ItemData(lbVenture.ListIndex) & " WHERE id = " & Grid.TextMatrix(mousRow, dcNumDoc)
    If myExecute("##126", sql) = 0 Then _
        Grid.Text = lbVenture.Text
    
    lbHide

End Sub

Private Sub lbVenture_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lbVenture_DblClick
ElseIf KeyCode = vbKeyEscape Then
    lbHide
End If

End Sub


Private Sub lbVenture_LostFocus()
    lbHide
End Sub

Private Sub mnFilter_Click()
Dim fromIndex As Integer, toIndex As Integer

    ' Перерисовать
    Call searchPair(mousRow, fromIndex, toIndex)
    If venturePairFilter(fromIndex)(toIndex) = 0 Then
        venturePairFilter(fromIndex)(toIndex) = 1
        Call foreColorGridRow(Grid, mousRow, vbRed, mousCol)
    Else
        venturePairFilter(fromIndex)(toIndex) = 0
        Call foreColorGridRow(Grid, mousRow, Grid.ForeColor, mousCol)
    End If
    
    
End Sub

Private Sub mnNomHistory_Click()
Dim selectedRows As Integer
Dim i As Integer
Dim curRow As Integer, startRow As Integer, stopRow As Integer
    
    selectedRows = Abs(Grid2.row - Grid2.RowSel) + 1
    ReDim DMCnomNom(selectedRows)
    
    If Grid2.row >= Grid2.RowSel Then
        startRow = Grid2.RowSel
        stopRow = Grid2.row
    Else
        startRow = Grid2.row
        stopRow = Grid2.RowSel
    End If
    
    i = 0
    curRow = Grid2.row
    For curRow = startRow To stopRow
        DMCnomNom(i + 1) = Grid2.TextMatrix(curRow, dnNomNom)
        i = i + 1
    Next curRow
    
    Me.MousePointer = flexHourglass
    VentureHistory.tbStartDate.Text = Grid.TextMatrix(mousRow, dcTermFrom)
    VentureHistory.tbEndDate.Text = Grid.TextMatrix(mousRow, dcTermTo)
    VentureHistory.isLoad = False
    VentureHistory.ckCumulative.value = Me.ckCumulative.value
    VentureHistory.ckPerList.value = 1
    VentureHistory.fillGrid
    VentureHistory.Show
    Me.MousePointer = flexDefault
    
End Sub

Private Sub mnAddToHistory_Click()

Dim i As Integer, str As String, newLen As Integer
Dim j As Integer, l As Long
Dim length As Integer
Dim aStep As Integer

    length = UBound(DMCnomNom)
    newLen = length
    If Grid2.row >= Grid2.RowSel Then
        aStep = -1
    Else
        aStep = 1
    End If
    For i = Grid2.row To Grid2.RowSel Step aStep
        str = Grid2.TextMatrix(i, dnNomNom)
        For j = 1 To length ' может этот эл-т был уже добавлен
            If DMCnomNom(j) = str Then GoTo NXT
        Next j
        newLen = newLen + 1
        ReDim Preserve DMCnomNom(newLen)
        DMCnomNom(newLen) = str ' чтобы корректно работал перерасчет Карты после правки в документе
NXT:
    Next i
    
    Me.MousePointer = flexHourglass
    VentureHistory.Show
    Me.MousePointer = flexDefault

End Sub



Private Sub mnRecalc_Click()
Dim i As Integer, preserve_yn As Boolean
Dim curRow As Long, curCol As Long, curTopRow As Long

    If MsgBox("Нажмите Да если вы хотите пересчитать накладную с новыми значениеми", vbDefaultButton2 Or vbYesNo, "") <> vbOK Then Exit Sub
    Grid2.Visible = False
    Grid2.col = dnNomName
    preserve_yn = False
    
    For i = 1 To Grid2.Rows
        Grid2.row = i
        If Grid2.ForeColor = vbRed Then
            preserve_yn = True
            Exit For
        End If
    Next i
    
    If preserve_yn Then
        If MsgBox("Нажмите Да если вы хотите сохранить значения введенные вручную", vbDefaultButton2 Or vbYesNo, "") <> vbOK Then
        End If
    End If
    
End Sub

Private Sub tbMobile_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String
Dim TermField As String
Dim fldValue As String

    If KeyCode = vbKeyReturn Then
        str = tbMobile.Text
        
        If mousCol = dcTermFrom Or mousCol = dcTermTo Or mousCol = dcNDate Then
            If tbMobile.Text = "" Then
                fldValue = "null"
            Else
                If Not isDateTbox(tbMobile) Then
                    Exit Sub
                Else
                    str = "'" & Format(tmpDate, "yyyy-mm-dd") & "'"
                    fldValue = "'" & Format(tmpDate, "yyyymmdd") & "'"
                End If
            End If
            
            If mousCol = dcTermFrom Then
                TermField = "TermFrom"
            ElseIf mousCol = dcTermTo Then
                TermField = "TermTo"
            ElseIf mousCol = dcNDate Then
                TermField = "nDate"
            End If
            
            If tbMobile.Text = "" Then
                If mousCol = dcTermFrom Then
                    str = "нач.учёта"
                ElseIf mousCol = dcTermTo Then
                    str = "кон.учета"
                ElseIf mousCol = dcNDate Then
                    str = ""
                End If
            Else
                str = tbMobile.Text
            End If
            
            sql = "UPDATE sDocsVenture SET " & TermField & " = " & fldValue
            
            
        ElseIf mousCol = dcProcent Then
            sql = "UPDATE sDocsVenture SET procent = " & tbMobile.Text
        ElseIf mousCol = dcNote Then
            sql = "UPDATE sDocsVenture SET note = '" & tbMobile.Text & "'"
        End If
            
        sql = sql _
                & " WHERE id = " & numDoc
        
'        Debug.Print sql
        If myExecute("##119", sql) <> 0 Then GoTo EN1
        
        
        If mousCol = dcProcent Then
            Grid.TextMatrix(mousRow, dcProcent) = str
            loadDocNomenk Grid.TextMatrix(mousRow, dcNumDoc)
        End If
        Grid.TextMatrix(mousRow, mousCol) = str
        lbHide
    ElseIf KeyCode = vbKeyEscape Then
EN1:
        lbHide
    End If

End Sub

Private Sub tbMobile2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim prevSummaItem As Double, prevSummaDoc As Double

    If KeyCode = vbKeyReturn Then
        If Not isNumericTbox(tbMobile2) Then Exit Sub
        
        prevSummaItem = CDbl(Grid2.TextMatrix(mousRow2, dnSumma))
        
        If mousCol2 = dnCosted Then
            Grid2.TextMatrix(mousRow2, dnSumma) = Round(CDbl(tbMobile2.Text) * CDbl(Grid2.TextMatrix(mousRow2, dnQuant)), 2)
            Grid2.TextMatrix(mousRow2, dnCosted) = Round(CDbl(tbMobile2.Text), 2)
        ElseIf mousCol2 = dnQuant Then
            Grid2.TextMatrix(mousRow2, dnSumma) = Round(CDbl(Grid2.TextMatrix(mousRow2, dnCosted)) * CDbl(tbMobile2.Text), 2)
            Grid2.TextMatrix(mousRow2, dnQuant) = Round(CDbl(tbMobile2.Text), 2)
        ElseIf mousCol2 = dnSumma Then
            Grid2.TextMatrix(mousRow2, dnCosted) = Round(CDbl(tbMobile2.Text) / CDbl(Grid2.TextMatrix(mousRow2, dnQuant)), 2)
            Grid2.TextMatrix(mousRow2, dnSumma) = Round(CDbl(tbMobile2.Text), 2)
        End If
        
        If mousCol2 = dnCosted Or mousCol2 = dnSumma Then
            sql = "update sdmcventure set costed = " _
                & Grid2.TextMatrix(mousRow2, dnCosted) _
                & " where id = " & Grid2.TextMatrix(mousRow2, dnId)
        ElseIf mousCol2 = dnQuant Then
            sql = "update sdmcventure set quant = " _
                & Grid2.TextMatrix(mousRow2, dnQuant) _
                & " where id = " & Grid2.TextMatrix(mousRow2, dnId)
        End If
        If myExecute("##119", sql) <> 0 Then GoTo EN1
        Grid2.TextMatrix(mousRow2, mousCol2) = tbMobile2.Text
        prevSummaDoc = CDbl(Grid.TextMatrix(mousRow, dcSumma))
        Grid.TextMatrix(mousRow, dcSumma) = prevSummaDoc - prevSummaItem + CDbl(Grid2.TextMatrix(mousRow2, dnSumma))
      
        lbHide2
        
    ElseIf KeyCode = vbKeyEscape Then
EN1:
        lbHide2
    End If
End Sub

Private Sub tbProcent_Change()
    procentChanged = True
End Sub

Private Sub tbProcent_LostFocus()
    If procentChanged Then
        sql = "update system set ivo_procent = " & tbProcent.Text
        If myExecute("##update ivo_procent", sql, 0) <> 0 Then
            ' unexpected error
        End If
    End If
End Sub

Private Sub tbProcent_Validate(Cancel As Boolean)
    If IsNumeric(tbProcent.Text) Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub
