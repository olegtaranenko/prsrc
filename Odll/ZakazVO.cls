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
Public WerkId As Integer
Public EquipId As Integer
Public FirmId As Integer        ' заказчик
Public StatusId As Integer
Public Outtime As Variant       ' Integer Orders
Public StatusInCeh As Variant   ' OrdersInCeh
Public Outdatetime As Variant   ' Orders

Public Remark As Variant ' nullable String, Orders
Public Transport As Variant ' nullable String, Orders


Public Worktime As Double       ' OrdersEquip накапливается по всем оборудованиям
Public statusEquipID As Variant ' OrdersEquip

Public DateRS As Variant        'Date  Orders
Public DateTimeMO As Variant    'Date  OrdersInCeh
Public WorktimeMO As Variant    'Double OrdersEquip накапливается по всем оборудованиям
Public StatM As Variant         'String OrdersInCeh
Public StatO As Variant         'String OrdersEquip
Public nevip As Variant         'Double OrdersInCeh
Public urgent As String         'Variant OrdersInCeh
Public equipStatusSync As Integer     ' 1 - статусы по оборудованиям рассогласованы, 0 - согласованы

Public m_lastModified As Variant        ' Date
Public m_lastModifiedEquip As Variant   ' Date
Public m_lastModifiedInCeh As Variant   ' Date
Public m_lastManagId As Variant         'String
Public m_lastManagEquipId As Variant    'String
Public m_lastManagInCehId As Variant    'String

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
        If Not IsNull(m_lastManagId) And Not IsEmpty(m_lastManagId) Then
            lastManagId = m_lastManagId
        Else
            lastManagId = m_lastManagEquipId
        End If
    End If
End Property


Public Property Get lastManag() As String
    Dim vlmid As Variant
    vlmid = lastManagId
    If Not IsNull(vlmid) Then
        lastManag = Manag(vlmid)
    End If
End Property


Public Sub incrementFromDb()

Dim resetToCurrentEquip As Boolean
Dim currentLastModified As Date
    
    If tbOrders.BOF Then
        Exit Sub
    End If
    
    done = False
    resetToCurrentEquip = False
    
        
    Numorder = tbOrders!Numorder
        
        If Not IsNull(tbOrders!lastModified) Then
            currentLastModified = tbOrders!lastModified
            If IsNull(Me.m_lastModified) Or (Not IsNull(Me.m_lastModified) And Me.m_lastModified < currentLastModified) Then
                Me.m_lastModified = currentLastModified
                resetToCurrentEquip = True
            End If
        End If
        
        WerkId = tbOrders!WerkId
        If IsNull(tbOrders!EquipId) Then
            EquipId = 0
        Else
            EquipId = tbOrders!EquipId
        End If
        StatusId = tbOrders!StatusId
        FirmId = tbOrders!FirmId
        
        m_lastManagId = tbOrders!lastManagId
        m_lastManagEquipId = tbOrders!lastManagEquipId
        
' PresentationFormat = 1 сокращенный, для чтения информации по Оборудованию?
        If tbOrders!presentationFormat = 1 Then
            
            If Not IsNull(tbOrders!lastModifiedEquip) Then
                currentLastModified = tbOrders!lastModifiedEquip
                If IsNull(Me.m_lastModified) Or (Not IsNull(Me.m_lastModified) And Me.m_lastModified < currentLastModified) Then
                    m_lastModifiedEquip = currentLastModified
                    resetToCurrentEquip = True
                End If
            End If

            If Not IsNull(tbOrders!Worktime) Then
                Worktime = Worktime + tbOrders!Worktime
            End If
            If Not IsNull(tbOrders!WorktimeMO) Then
                WorktimeMO = WorktimeMO + tbOrders!WorktimeMO
            End If
            If Not IsNull(tbOrders!StatO) Then
                ' комиш, абер лучше, чем просто переписывать
                StatO = StatO & tbOrders!StatO
                'StatO = tbOrders!StatO
            End If

            If Not IsNull(tbOrders!statusEquipID) Then
                If Not IsEmpty(statusEquipID) And statusEquipID <> tbOrders!statusEquipID Then
                    Me.equipStatusSync = 1
                End If
                statusEquipID = tbOrders!statusEquipID
            End If

        ElseIf tbOrders!presentationFormat = 0 Then

            Outtime = tbOrders!Outtime
            If IsNull(Outtime) Then
                Outdatetime = tbOrders!Outdatetime
            Else
                Outdatetime = DateAdd("h", Outtime, tbOrders!Outdatetime)
            End If
            Worktime = tbOrders!Worktime
            statusEquipID = tbOrders!statusEquipID
            StatusInCeh = tbOrders!StatusInCeh
            DateRS = tbOrders!DateRS
            DateTimeMO = tbOrders!DateTimeMO
            If Not IsNull(tbOrders!WorktimeMO) Then
                WorktimeMO = tbOrders!WorktimeMO
            End If
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


Public Sub initFromDb()
    clean
    incrementFromDb
End Sub


Public Sub initFromDb_()

    If tbOrders.BOF Then
        Exit Sub
    End If

    Numorder = tbOrders!Numorder
    If Not IsNull(tbOrders!WerkId) Then
        WerkId = tbOrders!WerkId
    End If
    If Not IsNull(tbOrders!EquipId) Then
        EquipId = tbOrders!EquipId
    End If
    StatusId = tbOrders!StatusId
    m_lastModified = tbOrders!lastModified
    m_lastManagId = tbOrders!lastManagId
    
    If tbOrders!presentationFormat = 1 Then
        m_lastModifiedEquip = tbOrders!lastModifiedEquip
        m_lastManagEquipId = tbOrders!lastManagEquipId
        
    ElseIf tbOrders!presentationFormat = 0 Then
        Outtime = tbOrders!Outtime
        If IsNull(Outtime) Then
            Outdatetime = tbOrders!Outdatetime
        Else
            Outdatetime = DateAdd("h", Outtime, tbOrders!Outdatetime)
        End If
        Worktime = tbOrders!Worktime
        statusEquipID = tbOrders!statusEquipID
        StatusInCeh = tbOrders!StatusInCeh
        DateRS = tbOrders!DateRS
        DateTimeMO = tbOrders!DateTimeMO
        WorktimeMO = tbOrders!WorktimeMO
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

    done = False

    Numorder = 0
    WerkId = 0
    EquipId = 0
    StatusId = -1
    Outtime = 0
    equipStatusSync = 0

    Worktime = 0
    WorktimeMO = 0
    
    Outdatetime = Null
    statusEquipID = Null
    StatusInCeh = Null
    
    DateRS = Null
    DateTimeMO = Null
    StatM = Null
    StatO = Null
    nevip = Null
    urgent = ""
    Me.m_lastModified = Null
    Me.m_lastManagId = 0
    Me.m_lastModifiedEquip = Null
    Me.m_lastManagEquipId = 0
End Sub

