Attribute VB_Name = "Errors"
Option Explicit

Sub fatalError(msg As String)
    MsgBox msg & vbCr & "���������� � ��������������", vbCritical, "����������� ������"
    End
End Sub
