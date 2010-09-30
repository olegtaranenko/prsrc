VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NomnomFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private keys() As String
Private values() As Variant
Private cursor As Long

Public Sub evict(Key As String)
    If exists(Key) Then
        Remove Key
    End If
End Sub


Public Function onlyInCache(nom As String) As Nomnom
    If IsEmpty(Item(nom)) Then
        Set onlyInCache = New Nomnom
    Else
        Set onlyInCache = Item(nom)
    End If
End Function


Public Function getNomnom(nom As String, Optional hasRecordSet As Boolean = False) As Variant
    If Not exists(nom) Then
        Dim Nomnom1 As New Nomnom
        Nomnom1.retrieveFromDb nom, hasRecordSet
        Add nom, Nomnom1
        Set getNomnom = Nomnom1
    Else
        If IsEmpty(Item(nom)) Then
            Set getNomnom = New Nomnom
        Else
            Set getNomnom = Item(nom)
        End If
    End If
    
End Function


Property Get Item(Key As String) As Variant
    Item = Empty
    If positionCursor(Key) Then
        If Not IsEmpty(values(cursor)) Then
            Set Item = values(cursor)
        End If
    End If
End Property


Property Get Count() As Long
    Count = UBound(keys)
End Property



Public Function exists(Key As String) As Boolean
    exists = False
    If positionCursor(Key) Then
        exists = True
    End If

End Function


Public Sub Add(Key As String, Value As Variant)
    If positionCursor(Key) Then
        Error -1
    Else
        Dim myCount As Long
        myCount = Count
        ReDim Preserve keys(myCount + 1)
        ReDim Preserve values(myCount + 1)
        keys(myCount) = Key
        Set values(myCount) = Value
        cursor = myCount
    End If
    
End Sub

Public Function Remove(Key As String) As Variant
    Remove = Empty
    If positionCursor(Key) Then
        Set Remove = values(cursor)
        Dim I As Long, myCount As Long
        myCount = Count
        For I = cursor + 1 To myCount
            keys(cursor) = keys(cursor + 1)
            If IsEmpty(values(cursor + 1)) Then
                values(cursor) = Empty
            Else
                Set values(cursor) = values(cursor + 1)
            End If
        Next I
        ReDim Preserve keys(myCount - 1)
        ReDim Preserve values(myCount - 1)
    Else
        'Error -1
    End If
End Function

Public Sub RemoveAll()
    ReDim keys(0)
    ReDim values(0)
    cursor = -1
End Sub

Private Function positionCursor(Key As String) As Boolean
Dim I As Long, startLooking As Long
Dim LookBack As Boolean
    
    
    positionCursor = False
    
    If cursor > 1 Then
        startLooking = cursor - 1
        LookBack = True
    Else
        startLooking = cursor
        LookBack = False
    End If
    
    If startLooking < 0 Then
        startLooking = 0
    End If
    
    For I = startLooking To UBound(keys)
        If Key = keys(I) Then
            cursor = I
            positionCursor = True
            Exit Function
        End If
    Next I
    
    If LookBack Then
        For I = startLooking - 1 To 0 Step -1
            If Key = keys(I) Then
                cursor = I
                positionCursor = True
                Exit Function
            End If
        Next I
    End If
End Function

Private Sub Class_Initialize()
    ReDim keys(0)
    ReDim values(0)
    cursor = -1
End Sub
