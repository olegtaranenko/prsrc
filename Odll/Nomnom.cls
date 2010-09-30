VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Nomnom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public nomName As String
Public cod As String
Public Size As String
Public ves As Variant ' за единицу продажную (лист)

Public edizm1 As String  ' дм
Public Edizm2 As String  ' лист
Public perlist As Single
Private m_quantInProduct 'Double

Public Function IsInited() As Boolean
    IsInited = perlist = 0
End Function

Public Function getEdizm(whole As Integer) As String
    If whole = 1 Then
        getEdizm = Edizm2
    Else
        getEdizm = edizm1
    End If
End Function

Public Function getVesEd(whole As Integer) As Double
    If whole = 1 Then
        getVesEd = ves
    Else
        getVesEd = ves / perlist
    End If
    getVesEd = calcRounded(getVesEd)
End Function



Public Function getCenaEd(cenaEd As Double, whole As Integer) As Double
    If whole = 1 Then
        getCenaEd = cenaEd * perlist
    Else
        getCenaEd = cenaEd
    End If
    getCenaEd = calcRounded(getCenaEd)
End Function

Public Function getCenaEdRevert(cenaEd As Double, whole As Integer) As Double
    If whole = 1 Then
        getCenaEdRevert = cenaEd / perlist
    Else
        getCenaEdRevert = cenaEd
    End If
    getCenaEdRevert = calcRounded(getCenaEdRevert)
End Function

' Возвращает уже округленное значение до 2-х или до 3-х цифр (если очень маленькое)
' нужно если вывод в лисах
Public Function getQuantity(quant As Double, whole As Integer) As Double
    If whole = 1 Then
        getQuantity = quant / perlist
    Else
        getQuantity = quant
    End If
    getQuantity = calcRounded(getQuantity)
End Function

Public Function getQuantityRevert(quant As Double, whole As Integer) As Double
    If whole = 1 Then
        getQuantityRevert = quant * perlist
    Else
        getQuantityRevert = quant
    End If
    getQuantityRevert = calcRounded(getQuantityRevert)
End Function

Public Sub retrieveFromDb(nom As String, hasRecordSet As Boolean)
Dim sql As String

    If hasRecordSet Then
        nomName = tbNomenk!nomName
        perlist = tbNomenk!perlist
        edizm1 = tbNomenk!ed_Izmer
        Edizm2 = tbNomenk!ed_Izmer2
        cod = tbNomenk!cod
        Size = tbNomenk!Size
        ves = tbNomenk!ves
    Else
        sql = "SELECT nomName, Ed_Izmer, Ed_Izmer2, perList, cod, size, ves From sGuideNomenk " _
        & "WHERE nomNom ='" & nom & "'"
        byErrSqlGetValues "##144", sql, nomName, edizm1, Edizm2, perlist, cod, Size, ves
    End If

End Sub

Property Get quantInProduct() As Double
    quantInProduct = m_quantInProduct
End Property

Property Let quantInProduct(Value As Double)
    m_quantInProduct = Value
End Property

