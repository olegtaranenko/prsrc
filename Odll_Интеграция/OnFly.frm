VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form OnFly 
   Caption         =   "Создание готового изделия на лету"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lbVariants 
      Height          =   255
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox tbQty 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7320
      TabIndex        =   1
      Text            =   "0"
      Top             =   240
      Width           =   585
   End
   Begin MSComCtl2.UpDown udQty 
      Height          =   285
      Left            =   7905
      TabIndex        =   2
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "tbQty"
      BuddyDispid     =   196611
      OrigLeft        =   7920
      OrigTop         =   360
      OrigRight       =   8175
      OrigBottom      =   675
      Max             =   1000000
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox tbField 
      CausesValidation=   0   'False
      Height          =   315
      Index           =   0
      Left            =   6720
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox tbField 
      CausesValidation=   0   'False
      Height          =   315
      Index           =   1
      Left            =   5400
      TabIndex        =   9
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox tbField 
      CausesValidation=   0   'False
      Height          =   315
      Index           =   2
      Left            =   4080
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CheckBox ckWeb 
      Caption         =   "web"
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox tbField 
      CausesValidation=   0   'False
      Height          =   315
      Index           =   3
      Left            =   5760
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox tbField 
      CausesValidation=   0   'False
      Height          =   315
      Index           =   4
      Left            =   4080
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox tbField 
      CausesValidation=   0   'False
      Height          =   315
      Index           =   5
      Left            =   4080
      TabIndex        =   7
      Top             =   3840
      Width           =   4095
   End
   Begin VB.CommandButton cmCreate 
      Caption         =   "Создать"
      Height          =   315
      Left            =   5760
      TabIndex        =   0
      Top             =   5400
      Width           =   1155
   End
   Begin VB.CommandButton cmCancel 
      Caption         =   "Отмена"
      Height          =   315
      Left            =   7200
      TabIndex        =   11
      Top             =   5400
      Width           =   1035
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   2820
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4974
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grid5 
      Height          =   1695
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   2990
      _Version        =   393216
      AllowBigSelection=   0   'False
      MergeCells      =   2
      AllowUserResizing=   1
   End
   Begin VB.Label Label7 
      Caption         =   "Цена продажи:"
      Height          =   195
      Left            =   6720
      TabIndex        =   13
      Top             =   4440
      Width           =   1275
   End
   Begin VB.Label Label6 
      Caption         =   "Время:"
      Height          =   195
      Left            =   5400
      TabIndex        =   14
      Top             =   4440
      Width           =   795
   End
   Begin VB.Label Label5 
      Caption         =   "Размер:"
      Height          =   195
      Left            =   4080
      TabIndex        =   15
      Top             =   4440
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "Номер:"
      Height          =   195
      Left            =   5760
      TabIndex        =   16
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Код изделия:"
      Height          =   195
      Left            =   4080
      TabIndex        =   17
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "Название изделия:"
      Height          =   195
      Left            =   4080
      TabIndex        =   18
      Top             =   3600
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Выберите группу или шаблон:"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Width           =   2355
   End
   Begin VB.Label laGrid5 
      Caption         =   "Выберите количество изделий, которое будет создано из данной  номенклатуры:"
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   6315
   End
End
Attribute VB_Name = "OnFly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ofName = 0
Private Const ofNomnom = 1
Private Const ofNomName = 2
Private Const ofEdizm = 3
Private Const ofQty = 4
Private Const ofPerList = 5

Private Const tbCena4 = 0
Private Const tbVremaObr = 1
Private Const tbSize = 2
Private Const tbSortNom = 3
Private Const tbPrName = 4
Private Const tbDescript = 5

Private cFldNames(tbCena4 To tbDescript) As String

Private usedNomnom() As String
Private usedQty() As Single
Private usedWhole() As Integer

' Массив содержащий доступных значений количества нового изделия в заказе
' Должен быть кратными значением для всех штучных изделий
Private possibleQty() As Integer
Private qtyChange As Boolean
Private manualChange As Boolean
Private fieldChanged As Boolean

Private prNames() As Variant
Private prSortNoms() As Variant
Private prSizes() As Variant
Private prCosts() As Variant
Private prTimes() As Variant

Private recieveField As Integer
Private blockFocus As Boolean
Private selectedNode As MSComctlLib.Node



Private Function checkUniqueness(csql As String) As Boolean
Dim ret As Integer
    byErrSqlGetValues "##610", csql, ret
    If ret = 0 Then
        checkUniqueness = True
    Else
        checkUniqueness = False
    End If
End Function

Private Function mapping(selector() As Variant, arg As Integer) As Variant
    mapping = selector(arg)
End Function

Private Sub cmCancel_Click()
    Unload Me
End Sub
Private Function validateForm() As Boolean
Dim I As Integer, ret As Boolean

    For I = tbField.LBound To tbField.UBound
        tbField_Validate I, ret
        If ret Then
            MsgBox "Значение " & tbField(I).Text & " не уникально ", , "Ошибка"
            tbField(I).SetFocus
            validateForm = False
            Exit Function
        End If
    Next I
    validateForm = True
    
End Function


Private Sub dbConvertSelected(newProductId As Integer, qty As Integer, Optional cenaEd As Single)
Dim I As Integer

    ' Сначала удаляем предметы из состава заказа
    'On Error GoTo fail
    
'    wrkDefault.BeginTrans
    For I = 1 To UBound(selectedItems)
        getIdFromGrid5Row sProducts, selectedItems(I)
'        nomnomOrPrid = sProducts.Grid5.TextMatrix(selectedItems(I), prId)
'        ext = CInt(sProducts.Grid5.TextMatrix(selectedItems(I), prExt))
        If sProducts.Grid5.TextMatrix(selectedItems(I), prType) = "изделие" Then
            sql = "delete from xPredmetyByIzdelia where numorder = " & gNzak & " and prId = '" & gProductId & "' and prExt = " & prExt
            If myExecute("##605." & selectedItems(I), sql) <> 0 Then GoTo fail
        Else
            sql = "delete from xPredmetyByNomenk where numorder = " & gNzak & " and nomnom = '" & gNomNom & "'"
            If myExecute("##604." & selectedItems(I), sql) <> 0 Then GoTo fail
        End If
    Next I
    
    sql = "insert into xpredmetybyizdelia (numOrder, prId, prExt, quant, cenaEd) " _
        & " select " & gNzak & ", " & newProductId & ", 0 ," & qty & ", " & cenaEd & ";"
    myExecute "##606", sql
'    wrkDefault.CommitTrans
    GoTo finally
fail:
    'wrkDefault.rollback
    errorCodAndMsg "Ошибка при преобразовании в изделие на лету"
    GoTo finally
finally:

End Sub


Private Function dbAddIzdelie( _
    prSeriaId As Integer, prName As String, prDescript As String _
    , Optional prSortNom As String, Optional prWeb As Boolean _
    , Optional prSize As String, Optional prTime As Single _
    , Optional prCost As Single _
) As Integer
Dim fields As String, values As String
Dim prId As Integer
Const comma = ", "
    
    sql = " select max(prId) + 1  from sguideproducts "
        
    If byErrSqlGetValues("##602.1", sql, dbAddIzdelie) = 0 Then
        Exit Function
    End If
    
    
    fields = "prSeriaId, prName, prDescript"
    
    values = prSeriaId _
        & comma & "'" & prName & "'" _
        & comma & "'" & prDescript & "'"
    
    If prSortNom <> "" Then
        fields = fields & comma & "SortNom"
        values = values & comma & "'" & prSortNom & "'"
    End If
    
    If prWeb Then
        fields = fields & comma & "web"
        values = values & comma & "'web'"
    End If
    
    If prSize <> "" Then
        fields = fields & comma & "prSize"
        values = values & comma & "'" & prSize & "'"
    End If
    
    If prTime <> 0 Then
        fields = fields & comma & "VremObr"
        values = values & comma & prTime
    End If
    If prCost <> 0 Then
        fields = fields & comma & "Cena4"
        values = values & comma & "'" & prCost & "'"
    End If
    ' Не делать транзакции здесь.
'    Debug.Print fields
    
    sql = "insert into sguideProducts (prId, " & fields & ") " _
        & " select " & CStr(dbAddIzdelie) & " , " _
        & values

    Debug.Print sql
        
    myExecute "##602.2", sql
    
End Function


Private Function selectedSeriaId() As Integer
Dim key As String
        
    key = selectedNode.key
    If Left(key, 1) = "p" Then
        key = selectedNode.Parent.key
    End If
    
    selectedSeriaId = CStr(Mid(key, 2))
    
End Function

Private Sub cmCreate_Click()
Dim I As Integer, newProductId As Integer, qty As Single

    If (Not validateForm) Then Exit Sub
    If MsgBox("Подтвердите что вы действительно хотите создать новое изделия", vbYesNo, "Новое изделие " & tbField(tbPrName).Text) = vbYes Then
        On Error GoTo rollback
        wrkDefault.BeginTrans
        sProducts.convertToIzdelie = True
        ' sProducts.mnDel_Click
        
        newProductId = dbAddIzdelie(selectedSeriaId _
            , tbField(tbPrName).Text _
            , tbField(tbDescript).Text _
            , tbField(tbSortNom).Text _
            , IIf(ckWeb.value = 1, True, False) _
            , tbField(tbSize).Text _
            , tbField(tbVremaObr).Text _
            , tbField(tbCena4).Text _
        )
        
        For I = 1 To UBound(usedNomnom)
            qty = usedQty(I) / CSng(tbQty.Text)
            sql = "insert into sproducts( productid, nomnom, quantity)" _
                & " values( " & newProductId & ", '" & usedNomnom(I) & "', " & qty _
                & ")"
            myExecute "##602.3", sql
        Next I
        
        sProducts.tbQuant.Text = Me.tbQty.Text
        gProductId = newProductId
        prExt = 0
        dbConvertSelected newProductId, CInt(tbQty.Text), CSng(tbField(tbCena4).Text)
        
        wrkDefault.CommitTrans
        
        Unload Me
        GoTo finally
rollback:
        wrkDefault.rollback
        errorCodAndMsg "Новое изделие"
        MsgBox "При создании нового изделия произошла ошибка", , "Обратитесь к администратору"
        GoTo finally
        
    End If
finally:
    'sProducts.convertToIzdelie = False
    
End Sub

Private Sub Form_Load()
    cFldNames(tbPrName) = "prName"
    cFldNames(tbSortNom) = "SortNom"
    
    Grid5.FormatString = "|<Код|<Описание|<Ед.измерения" & _
    "|Кол-во|Штучн."
    Grid5.ColWidth(ofName) = 0
    Grid5.ColWidth(ofNomnom) = 900
    Grid5.ColWidth(ofNomName) = 5200
    Grid5.ColWidth(ofEdizm) = 650
    Grid5.ColWidth(ofQty) = 650
    Grid5.ColWidth(ofPerList) = 500
    
    loadSelected sProducts.Grid5, selectedItems
    loadSeria tv
    loadTemplates tv
    
    
End Sub
Private Sub loadTemplates(ByRef p_tv As TreeView)
Dim key As String, pKey As String
    
'    sql = _
     " select * from sguideproducts g " _
    & "    where not exists (select 1 from sproducts p where p.productId = g.prId)"
    
    sql = _
        "select count(*), max(prId) as prId, prDescript, prSeriaId from sguideproducts group by prDescript, prSeriaId " _
        & " order by prSeriaId, prDescript"

    Set tbSeries = myOpenRecordSet("##601.1", sql, 0)
    If tbSeries Is Nothing Then myBase.Close: End
    If Not tbSeries.BOF Then
        
        While Not tbSeries.EOF
            key = "p" & tbSeries!prId
            pKey = "k" & tbSeries!prSeriaId
            Set Node = p_tv.Nodes.Add(pKey, tvwChild, key, _
                tbSeries!prDescript)
            Node.Sorted = True
            Node.Bold = True
            'Node.EnsureVisible
            
            
            tbSeries.MoveNext
        Wend
        tbSeries.Close
        
    End If
    
End Sub
Private Sub append(item As Variant)
Dim sz As Integer
    
    sz = UBound(usedNomnom)
    ReDim Preserve usedNomnom(sz + 1)
    usedNomnom(sz + 1) = item
    
End Sub

Private Function existsUsed(p_nomnom As String) As Long
Dim I As Long
    For I = 1 To UBound(usedNomnom)
        If usedNomnom(I) = p_nomnom Then
            existsUsed = I: Exit Function
        End If
    Next I
    existsUsed = Empty
End Function

Private Sub putNomenkToGrid(Grd As MSFlexGrid, p_nomnom As String _
    , p_qty As Single, p_edIzm As String, p_Size As String, p_cod As String _
    , p_perList As Single, p_NomName As String)
Dim v_name As String
Dim v_rowUsed As Long
Dim v_qty As Single

    v_rowUsed = existsUsed(p_nomnom)
    
    If IsEmpty(v_rowUsed) Or v_rowUsed = 0 Then
        Grd.AddItem ""
        Grd.TextMatrix(Grd.Rows - 1, ofNomnom) = p_nomnom
        If p_cod <> "" Then v_name = p_cod & " "
        
        v_name = v_name & p_NomName & " " & p_Size
        Grd.TextMatrix(Grd.Rows - 1, ofNomName) = v_name
        Grd.TextMatrix(Grd.Rows - 1, ofEdizm) = p_edIzm
        Grd.TextMatrix(Grd.Rows - 1, ofQty) = p_qty
        If p_perList = 1 Then
            Grd.TextMatrix(Grd.Rows - 1, ofPerList) = "Да"
        End If
        append p_nomnom
        appendQty p_qty, p_perList
    Else
        usedQty(v_rowUsed) = usedQty(v_rowUsed) + p_qty
'        v_qty = CSng(Grd.TextMatrix(v_rowUsed, ofQty))
'        Grd.TextMatrix(v_rowUsed, ofQty) = CStr(v_qty + p_qty)
    End If
    
    
End Sub

Private Sub loadSelected(Grd As MSFlexGrid, selectedItems() As Long)
Dim I As Integer
Dim sz As Integer

' Инициализация массива кодов используемой номенклатуры
    ReDim usedNomnom(0)
    ReDim usedQty(0)
    ReDim usedWhole(0)

' Очистить таблицу
    Grid5.Rows = 1
    
    ' По всем выбранным позициям
    sz = UBound(selectedItems)
    For I = 1 To sz
        
        If Grd.TextMatrix(selectedItems(I), prType) = "изделие" Then
        
            getIdFromGrid5Row sProducts, CLng(selectedItems(I))
            loadProduct _
                  gProductId _
                , prExt _
                , CSng(Grd.TextMatrix(selectedItems(I), prQuant))
        Else
            loadNomenk _
                  Grd.TextMatrix(selectedItems(I), prId) _
                , Grd.TextMatrix(selectedItems(I), prQuant)
        End If
        
    Next I
    
    findAllPossibleWholeQtys
    
    'По умолчанию установить максимальное
    If UBound(possibleQty) > 0 Then
        udQty.max = possibleQty(UBound(possibleQty))
        tbQty.Text = udQty.max
    Else
        If UBound(usedWhole) > 0 Then
            udQty.max = 1
        End If
        tbQty.Text = 1
    End If
End Sub


Private Sub loadNomenk(p_nomnom As String, p_qty As Single)
' Отдельная номенклатура
    sql = _
            "select n.nomnom, n.ed_izmer, n.size, n.cod, n.perList, n.nomName" _
        & " from sguidenomenk n where nomnom = '" & p_nomnom & "'"

    Set tbNomenk = myOpenRecordSet("##onFly.2", sql, dbOpenForwardOnly)
    If tbNomenk Is Nothing Then Exit Sub
    If Not tbNomenk.BOF Then
        While Not tbNomenk.EOF
            putNomenkToGrid Grid5, tbNomenk!nomNom, p_qty, tbNomenk!ed_Izmer, tbNomenk!Size, tbNomenk!cod, tbNomenk!perList, tbNomenk!nomName
            tbNomenk.MoveNext
        Wend
        tbNomenk.Close
        
    End If
        

End Sub


Private Sub loadProduct(ByVal p_productId As Integer, ByVal p_prext As Integer, p_qty As Single)
Dim p_numOrder As Long


    p_numOrder = CLng(gNzak)
' Номенклатура по изделиям (включая вариантные)
    sql = _
            "   select p.nomnom, n.ed_izmer, n.size, n.cod, n.perList, n.nomName, p.quantity " _
        & "   from sproducts p " _
        & "     join xpredmetybyizdelia i on numorder = " & CStr(p_numOrder) & " and i.prid = p.productid and i.prext = " & CStr(p_prext) _
        & "     join sguidenomenk n on n.nomnom = p.nomnom " _
        & "   where p.productid = " & CStr(p_productId) _
        & "     and not exists (select 1 from   sguidevariant v where p.xgroup != '' and v.xgroup = p.xgroup and v.productid = p.productid and c > 1) " _
        & "   union " _
        & "   select v.nomnom, n.ed_izmer, n.size, n.cod, n.perList, n.nomName, p.quantity " _
        & "   from xvariantnomenc v" _
        & "     join sguidenomenk n on n.nomnom = v.nomnom" _
        & "     join sproducts p on p.productid = v.prid and v.nomnom = p.nomnom" _
        & "   where v.prid = " & CStr(p_productId) & " and v.prext = " & CStr(p_prext) _
        & "        and v.numorder = " & CStr(p_numOrder)
    
    'Debug.Print sql
    
    Set tbNomenk = myOpenRecordSet("##onFly.1", sql, dbOpenForwardOnly)
    If tbNomenk Is Nothing Then Exit Sub
    If Not tbNomenk.BOF Then
        While Not tbNomenk.EOF
            putNomenkToGrid Grid5, tbNomenk!nomNom, p_qty * tbNomenk!quantity, tbNomenk!ed_Izmer, tbNomenk!Size, tbNomenk!cod, tbNomenk!perList, tbNomenk!nomName
            tbNomenk.MoveNext
        Wend
    End If
    tbNomenk.Close

End Sub
Private Sub setAllQty(Optional p_product_qty As Integer = 1)
Dim I As Integer
    ' Изменить удельную номенклатуру на единицу изделия
    For I = 1 To UBound(usedQty)
        Grid5.TextMatrix(I, ofQty) = Round(usedQty(I) / p_product_qty, 2)
    Next I

End Sub

Private Sub appendQty(p_qty As Single, p_perList As Single)

    ReDim Preserve usedQty(UBound(usedQty) + 1)
    usedQty(UBound(usedQty)) = p_qty
    If (p_perList = 1) Then
        ReDim Preserve usedWhole(UBound(usedWhole) + 1)
        usedWhole(UBound(usedWhole)) = p_qty
    End If
    
End Sub

' Проверка, является od делителем qty
Private Function tryOD(ByVal qty As Integer, od As Integer) As Boolean
Dim unround As Single

    tryOD = False
    unround = qty / od
    If (Round(unround, 0) = unround) Then
        tryOD = True
        Exit Function
    End If
    
End Function


Private Sub findAllPossibleWholeQtys()
    

Dim I As Integer, j As Integer
Dim imax As Integer
Dim isPossible As Boolean


    ReDim possibleQty(0)
    possibleQty(0) = 1
    ' найти максимум
    imax = 0
    For I = 1 To UBound(usedWhole)
        If imax < usedWhole(I) Then imax = usedWhole(I)
    Next I
    
    
    'для всех штучных изделий...
    For I = 2 To imax
        ' проверить на делитель
        isPossible = True
        For j = 1 To UBound(usedWhole)
            If Not tryOD(usedWhole(j), I) Then
                isPossible = False
                Exit For
            End If
        Next j
        If isPossible Then
            ReDim Preserve possibleQty(UBound(possibleQty) + 1)
            possibleQty(UBound(possibleQty)) = I
        End If
        
    Next I
    
    
End Sub

Private Function ArrangePossibleWhole(tryQty As Integer)
Dim I As Integer
Dim cur As Integer, prev As Integer
    ' Используем тот факт, что в possibleQty лежит отсортированный
    ' список возможных значений
    
    ArrangePossibleWhole = tryQty
    If UBound(usedWhole) = 0 Then Exit Function
    
    prev = 1
    For I = 0 To UBound(possibleQty)
        cur = possibleQty(I)
        If tryQty = cur Then Exit Function
        If tryQty < cur Then
            ArrangePossibleWhole = prev
            Exit Function
        End If
        prev = cur
    Next I
    ArrangePossibleWhole = cur 'макимум
    
End Function
Private Sub forceQtyChange()
Dim chkQty As Integer
    If IsNumeric(tbQty.Text) Then
        chkQty = CInt(tbQty.Text)
    End If
    If chkQty > 0 Then
        setAllQty (chkQty)
    Else
        tbQty.Text = 1
    End If
End Sub


Private Sub lbVariants_DblClick()
    tbField(recieveField).Text = lbVariants.List(lbVariants.ListIndex)
    lbVariants.Visible = False
'    blockFocus = True
    tbField(recieveField).SetFocus
End Sub

Private Sub lbVariants_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        lbVariants_DblClick
    ElseIf KeyAscii = vbKeyEscape Then
        lbVariants.Visible = False
'        blockFocus = True
        tbField(recieveField).SetFocus
    End If
End Sub

Private Sub lbVariants_LostFocus()
    lbVariants.Visible = False
End Sub

Private Sub tbField_Change(Index As Integer)
Dim I As Integer
    If Index = tbPrName Then
        For I = 1 To UBound(prNames)
            If prNames(I) = tbField(Index).Text Then
                initFields I
                Exit For
            End If
        Next I
    End If

    If Not manualChange Then Exit Sub
    fieldChanged = True
    manualChange = False
End Sub


Private Function isNewItem(lb As listBox, itm As Variant, up As Integer) As Boolean
Dim I As Integer
    isNewItem = True
    For I = 1 To up
        If lb.List(I - 1) = CStr(itm) Then
            isNewItem = False
            Exit Function
        End If
    Next I
    
End Function


Private Function makeUniqueListbox(arr() As Variant, tbIndex As Integer) As Integer

Dim I As Integer

    While lbVariants.ListCount > 0
        lbVariants.removeItem (lbVariants.ListCount - 1)
    Wend
    makeUniqueListbox = -1
    For I = 1 To UBound(arr)
        If isNewItem(lbVariants, arr(I), I) Then
           lbVariants.AddItem CStr(arr(I))
        End If
    Next I
    
    For I = 0 To lbVariants.ListCount - 1
        If lbVariants.List(I) = tbField(tbIndex) Then
            makeUniqueListbox = I
        End If
    Next I

    'lbVariants.Height = currentItems * 250
End Function


Private Sub showVariants(Index As Integer)
Dim noShow As Boolean, lbIndex As Integer
    
    If blockFocus Then
        blockFocus = False
        Exit Sub
    End If
    noShow = False
    recieveField = Index
    If UBound(prNames) > 1 Then
        'clearUniqueListbox
        If Index = tbCena4 Then
            lbIndex = makeUniqueListbox(prCosts, Index)
        ElseIf Index = tbVremaObr Then
            lbIndex = makeUniqueListbox(prTimes, Index)
        ElseIf Index = tbSize Then
            lbIndex = makeUniqueListbox(prSizes, Index)
        ElseIf Index = tbSortNom Then
            lbIndex = makeUniqueListbox(prSortNoms, Index)
        ElseIf Index = tbPrName Then
            lbIndex = makeUniqueListbox(prNames, Index)
        Else
            noShow = True
        End If
        If Not noShow And lbVariants.ListCount > 1 Then
            lbVariants.Width = tbField(Index).Width
            lbVariants.Left = tbField(Index).Left
            lbVariants.Top = tbField(Index).Top
            lbVariants.Height = lbVariants.ListCount * 205 + 50
            If Me.Height - lbVariants.Top - 500 < lbVariants.Height Then
                lbVariants.Height = Me.Height - lbVariants.Top - 500
            End If
            
            lbVariants.ListIndex = lbIndex
            lbVariants.Visible = True
'            blockFocus = True
            lbVariants.SetFocus
        End If
    End If
    
End Sub

Private Sub tbField_DblClick(Index As Integer)
    showVariants Index
End Sub

Private Sub tbField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    manualChange = True
    If (Shift = vbCtrlMask And KeyCode = vbKeyDown) Or KeyCode = vbKeyReturn Then
        showVariants Index
    End If
    
End Sub

Private Sub tbField_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    manualChange = False
End Sub

Private Sub tbField_LostFocus(Index As Integer)
    fieldChanged = False
    If Not blockFocus Then
        'lbVariants.Visible = False
    End If
End Sub

Private Sub tbField_Validate(Index As Integer, Cancel As Boolean)

    Cancel = False
    If Index = tbPrName Then
        If Not checkUniqueness("select count(*) from sguideproducts where " _
            & " prName = '" & tbField(Index).Text & "'" _
        ) Then
            Cancel = True
        End If
    ElseIf Index = tbSortNom Then
        If Not checkUniqueness("select count(*) from sguideproducts where " _
            & " prSeriaId = " & selectedSeriaId & " and SortNom = '" & tbField(tbSortNom).Text & "'" _
        ) Then
            Cancel = True
        End If
    End If

'    If fieldChanged And tbField(Index).Text = "" Then
'        MsgBox "Поле не должно быть пустым", , "Ошибка"
'        tbField(Index).SetFocus
'        Cancel = True
'    End If
End Sub

Private Sub tbQty_Change()

    If qtyChange = True Then Exit Sub
    forceQtyChange
End Sub

Private Sub tbQty_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyReturn Then
        If IsNumeric(tbQty.Text) Then
            If UBound(usedWhole) = 0 Then
                forceQtyChange
            Else
                tbQty.Text = ArrangePossibleWhole(CInt(tbQty.Text))
            End If
        Else
            tbQty.Text = 1
        End If
    Else
        qtyChange = True
    End If
End Sub

Private Sub tbQty_KeyUp(KeyCode As Integer, Shift As Integer)
    qtyChange = False
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
Dim key As String
    
    key = Node.key
    Set selectedNode = Node
    If key = "k0" Then Exit Sub
    fieldChanged = False
    If Left(key, 1) = "p" Then
        fillDataWithTemplate CLng(Mid(key, 2))
    Else
        clearData
    End If

End Sub
Private Sub clearData()
    tbField(tbCena4).Text = ""
    tbField(tbVremaObr).Text = ""
    tbField(tbSize).Text = ""
    ckWeb.value = False
    tbField(tbSortNom).Text = ""
    tbField(tbPrName).Text = ""
    tbField(tbDescript).Text = ""
    
End Sub
Private Sub fillDataWithTemplate(p_productId As Long)
Dim sz As Integer


'    sql = _
     " select * from sguideproducts g " _
    & "    where g.prId = " & CStr(p_productId)
    
    sql = _
        "select l.* from sguideproducts l " _
        & "join sguideproducts p on p.prSeriaId = l.prSeriaId and p.prDescript = l.prDescript and p.prid = " & CStr(p_productId) _
        & "order by l.sortnom "

    ReDim prNames(0)
    ReDim prSortNoms(0)
    ReDim prSizes(0)
    ReDim prCosts(0)
    ReDim prTimes(0)

    Set tbSeries = myOpenRecordSet("##601.1", sql, 0)
    If tbSeries Is Nothing Then myBase.Close: End
    If Not tbSeries.BOF Then
        
        While Not tbSeries.EOF
            sz = UBound(prNames) + 1
            ReDim Preserve prNames(sz)
            ReDim Preserve prSortNoms(sz)
            ReDim Preserve prSizes(sz)
            ReDim Preserve prCosts(sz)
            ReDim Preserve prTimes(sz)
            
            prCosts(sz) = tbSeries!Cena4
            prTimes(sz) = tbSeries!VremObr
            prSizes(sz) = tbSeries!prSize
            'ckWeb.value = False
            prSortNoms(sz) = tbSeries!sortNom
            prNames(sz) = tbSeries!prName
            tbField(tbDescript).Text = tbSeries!prDescript 'одна для всех
            tbSeries.MoveNext
        Wend
    End If
    tbSeries.Close
    initFields sz
'    tbField(tbDescript).Text = prDescript(index)
End Sub

Private Sub initFields(Index As Integer)
    tbField(tbCena4).Text = prCosts(Index)
    tbField(tbVremaObr).Text = prTimes(Index)
    tbField(tbSize).Text = prSizes(Index)
    tbField(tbSortNom).Text = prSortNoms(Index)
    tbField(tbPrName).Text = prNames(Index)
    
End Sub
Private Function nextInPossible(current As Integer, up As Boolean)
Dim I As Integer
    nextInPossible = current
    
    If UBound(possibleQty) = 0 Then Exit Function
    
    If Not up Then
        For I = 0 To UBound(possibleQty)
            If possibleQty(I) > current Then Exit For
        Next I
        nextInPossible = possibleQty(I - 1)
    Else
        For I = UBound(possibleQty) To 0 Step -1
            If possibleQty(I) < current Then Exit For
        Next I
        'If I < 0 Then
            nextInPossible = possibleQty(I + 1)
        'End If
    End If
    
End Function

Private Sub udQty_DownClick()
    tbQty.Text = nextInPossible(CInt(tbQty.Text), False)
End Sub


Private Sub udQty_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    qtyChange = True
End Sub


Private Sub udQty_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    qtyChange = False
    forceQtyChange
End Sub

Private Sub udQty_UpClick()
    tbQty.Text = nextInPossible(CInt(tbQty.Text), True)
End Sub
