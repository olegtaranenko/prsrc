Attribute VB_Name = "CommonUtils"
Option Explicit

Sub fatalError(msg As String)
    MsgBox msg & vbCr & "���������� � ��������������", vbCritical, "����������� ������"
    End
End Sub
