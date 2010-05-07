VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZakazVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Numorder As Long
Public werkId As Integer
Public equipId As Integer
Public StatusId As Integer
Public outTime As Variant ' Integer (may be null)

Public Worktime As Double
Public Outdatetime As Variant
Public statusEquipID As Variant
Public statusInCeh As Variant

Public DateRS As Variant
Public DateTimeMO As Variant
Public workTimeMO As Variant
Public StatM As Variant
Public StatO As Variant
Public nevip As Variant
Public urgent As String 'Variant
Public m_lastModified As Variant ' Date
Public m_lastModifiedEquip As Variant ' Date
Public m_lastManagId As Variant 'String
Public m_lastManagEquipId As Variant 'String

Private done As Boolean

Public Property Get lastModified() As Variant
    If IsNull(m_lastModified) And IsNull(m_lastModifiedEquip) Then
        lastModified = Null
    Else
        If IsNull(m_lastModified) Then
            lastModified = m_lastModifiedEquip
        ElseIf IsNull(m_lastModifiedEquip) Then
            lastModified = m_lastModified
        Else
            If m_lastModified > m_lastModifiedEquip Then
                lastModified = m_lastModified
            Else
                lastModified = m_lastModifiedEquip
            End If
        End If
    End If
End Property

Public Property Get lastManagId() As Variant
    If IsNull(m_lastManagId) And IsNull(m_lastManagEquipId) Then
        lastManagId = Null
    Else
        If IsNull(m_lastManagId) Then
            lastManagId = m_lastManagEquipId
        ElseIf IsNull(m_lastManagEquipId) Then
            lastManagId = m_lastManagId
        Else
            If m_lastManagId > m_lastManagEquipId Then
                lastManagId = m_lastManagId
            Else
                lastManagId = m_lastManagEquipId
            End If
        End If
    End If
End Property

Public Property Get lastManag() As String
    Dim vlmid As Variant
    vlmid = lastManagId
    lastManag = getManagById(vlmid)
End Property



Public Sub incrementFromDb()
Dim resetToCurrentEquip As Boolean
    done = False
    resetToCurrentEquip = False
    If tbOrders.BOF Then
        Exit Sub
    End If
    If tbOrders!Numorder <> Me.Numorder Then
        Me.clean
        Me.initFromDb
    Else
        Dim currentLastModified As Date
        If Not IsNull(tbOrders!lastModified) Then
            currentLastModified = tbOrders!lastModified
            If IsNull(Me.m_lastModified) Or (Not IsNull(Me.m_lastModified) And Me.m_lastModified < currentLastModified) Then
                Me.m_lastModified = currentLastModified
                resetToCurrentEquip = True
            End If
        End If
        If Not IsNull(tbOrders!lastModifiedEquip) Then
            currentLastModified = tbOrders!lastModifiedEquip
            If IsNull(Me.m_lastModified) Or (Not IsNull(Me.m_lastModified) And Me.m_lastModified < currentLastModified) Then
                m_lastModifiedEquip = currentLastModified
                resetToCurrentEquip = True
            End If
        End If
    End If
    
    If resetToCurrentEquip Then
        werkId = tbOrders!werkId
        equipId = tbOrders!equipId
        If tbOrders!presentationFormat = 1 Then
        ElseIf tbOrders!presentationFormat = 0 Then
            outTime = tbOrders!outTime
            If IsNull(outTime) Then
                Outdatetime = tbOrders!Outdatetime
            Else
                Outdatetime = DateAdd("h", outTime, tbOrders!Outdatetime)
            End If
            Worktime = tbOrders!Worktime
            statusEquipID = tbOrders!statusEquipID
            statusInCeh = tbOrders!statusInCeh
            DateRS = tbOrders!DateRS
            DateTimeMO = tbOrders!DateTimeMO
            workTimeMO = tbOrders!workTimeMO
            StatM = tbOrders!StatM
            StatO = tbOrders!StatO
            nevip = tbOrders!nevip
            If Not IsNull(tbOrders!urgent) Then
                urgent = tbOrders!urgent
            Else
                urgent = ""
            End If
        End If
        m_lastManagEquipId = tbOrders!lastManagEquipId
        m_lastManagId = tbOrders!lastManagId

        done = True
    End If
    
End Sub

Public Sub initFromDb()

    If tbOrders.BOF Then
        Exit Sub
    End If

    Numorder = tbOrders!Numorder
    If Not IsNull(tbOrders!werkId) Then
        werkId = tbOrders!werkId
    End If
    If Not IsNull(tbOrders!equipId) Then
        equipId = tbOrders!equipId
    End If
    StatusId = tbOrders!StatusId
    m_lastModified = tbOrders!lastModified
    m_lastManagId = tbOrders!lastManagId
    
    If tbOrders!presentationFormat = 1 Then
        m_lastModifiedEquip = tbOrders!lastModifiedEquip
        m_lastManagEquipId = tbOrders!lastManagEquipId
    ElseIf tbOrders!presentationFormat = 0 Then
        outTime = tbOrders!outTime
        If IsNull(outTime) Then
            Outdatetime = tbOrders!Outdatetime
        Else
            Outdatetime = DateAdd("h", outTime, tbOrders!Outdatetime)
        End If
        Worktime = tbOrders!Worktime
        statusEquipID = tbOrders!statusEquipID
        statusInCeh = tbOrders!statusInCeh
        DateRS = tbOrders!DateRS
        DateTimeMO = tbOrders!DateTimeMO
        workTimeMO = tbOrders!workTimeMO
        StatM = tbOrders!StatM
        StatO = tbOrders!StatO
        nevip = tbOrders!nevip
        If Not IsNull(tbOrders!urgent) Then
            urgent = tbOrders!urgent
        Else
            urgent = ""
        End If
    End If
    done = True
End Sub

Public Property Get inited() As Boolean
    inited = done
End Property

Public Sub clean()
    Numorder = 0
    werkId = 0
    equipId = 0
    StatusId = -1
    outTime = 0
    
    Worktime = 0
    Outdatetime = Null
    statusEquipID = Null
    statusInCeh = Null
    
    DateRS = Null
    DateTimeMO = Null
    workTimeMO = Null
    StatM = Null
    StatO = Null
    nevip = Null
    urgent = ""
    Me.m_lastModified = Null
    Me.m_lastManagId = 0
    Me.m_lastModifiedEquip = Null
    Me.m_lastManagEquipId = 0
End Sub

