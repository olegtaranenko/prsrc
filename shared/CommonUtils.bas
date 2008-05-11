Attribute VB_Name = "CommonUtils"
Option Explicit

Sub fatalError(msg As String)
    MsgBox msg & vbCr & "Обратитесь к администратору", vbCritical, "Критическая ошибка"
    End
End Sub
