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
Public cehId As Integer
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
Public urgent As Variant

Private done As Boolean


Public Sub initFromDb()

    If tbOrders.BOF Then
        Exit Sub
    End If

    Numorder = tbOrders!Numorder
    cehId = tbOrders!cehId
    StatusId = tbOrders!StatusId
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
    urgent = tbOrders!urgent

    done = True
End Sub

Public Property Get inited() As Boolean
    inited = done
End Property
