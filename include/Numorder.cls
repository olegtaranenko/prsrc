VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Numorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_dat As Date
Private m_num As Integer
Private m_year As Integer
Private m_month As Integer
Private m_day As Integer
Private m_doc As Boolean
Private inited As Boolean

Property Let docs(value As Boolean)
    m_doc = value
End Property



Property Get num() As Integer
    num = m_num
End Property

Property Let num(value As Integer)
    m_num = value
End Property

Property Get val() As Long
    Dim fmt_month As Integer
    Dim i_num As Integer, i_day As Integer, i_month As Integer
    i_num = m_num: i_day = m_day: i_month = m_month
    
    While i_num >= 100
        i_num = i_num - 100
        i_day = i_day + 31
    Wend
    If m_doc Then
        i_month = i_month + 20
    End If
    val = CLng(Format(Right(CStr(m_year), 2), "#0") & Format(i_month, "00") & Format(i_day, "00") & Format(i_num, "00"))
End Property

Property Let val(value As Long)
    Dim prevYearSize As Integer
    Dim str As String
    
    str = CStr(value)
    prevYearSize = -1
    If Len(str) = 7 Then
        prevYearSize = 0
        m_year = "200" & left$(value, 1)
    ElseIf Len(str) = 8 Then
        prevYearSize = 1
        m_year = "20" & left$(value, 2)
    End If
 
    If prevYearSize >= 0 Then
        Dim i_day As Integer, i_num As Integer, i_month As Integer
        
        i_month = Mid$(value, 2 + prevYearSize, 2)
        i_day = Mid$(value, 4 + prevYearSize, 2)
        i_num = Right(value, 2)
        While i_day > 31
            i_num = i_num + 100
            i_day = i_day - 31
        Wend
        
        If i_month > 12 Then
            m_doc = True
            i_month = i_month - 20
        Else
            m_doc = False
        End If
        
        m_num = i_num
        m_day = i_day
        m_month = i_month
        m_dat = DateSerial(m_year, m_month, m_day)
        inited = True
    Else
        inited = False
    End If
End Property


Property Get dat() As Date
    dat = m_dat
End Property

Property Let dat(value As Date)
    m_dat = value
    m_year = DatePart("yyyy", m_dat)
    m_month = DatePart("m", m_dat)
    m_day = DatePart("d", m_dat)
End Property


Public Function nextNum() As Long
    m_num = m_num + 1
    nextNum = val
End Function


Private Sub Class_Initialize()
    dat = Now
    m_num = 0
    inited = True
    m_doc = False
End Sub

Public Function isEmpty() As Boolean
    isEmpty = Not inited
End Function

Public Function isCurrentDay() As Boolean
    isCurrentDay = DateDiff("d", m_dat, Now) = 0
End Function

