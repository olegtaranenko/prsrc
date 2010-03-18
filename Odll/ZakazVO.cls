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

Public numorder As Long
Public cehId As Integer
Public statusId As Integer

Public worktime As Double
Public outdatetime As Variant
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

    numorder = tbOrders!numorder
    cehId = tbOrders!cehId
    statusId = tbOrders!statusId
    worktime = tbOrders!worktime
    outdatetime = tbOrders!outdatetime
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
