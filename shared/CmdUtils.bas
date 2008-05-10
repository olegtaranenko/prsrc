Attribute VB_Name = "CmdUtils"
Option Explicit

Public parsedArgs() As Variant

Type AliasInfo
    key As String
    value As String
End Type



Function GetCommandLine(Optional MaxArgs)
   'Declare variables.
   Dim C, CmdLine, CmdLnLen, InArg, I, NumArgs
   'See if MaxArgs was provided.
   If IsMissing(MaxArgs) Then MaxArgs = 10
   'Make array of the correct size.
   ReDim ArgArray(MaxArgs)
   NumArgs = 0: InArg = False
   'Get command line arguments.
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   'Go thru command line one character
   'at a time.
   For I = 1 To CmdLnLen
      C = Mid(CmdLine, I, 1)
      'Test for space or tab.
      If (C <> " " And C <> vbTab) Then
         'Neither space nor tab.
         'Test if already in argument.
         If Not InArg Then
         'New argument begins.
         'Test for too many arguments.
            If NumArgs = MaxArgs Then Exit For
            NumArgs = NumArgs + 1
            InArg = True
         End If
         'Concatenate character to current argument.
         ArgArray(NumArgs) = ArgArray(NumArgs) & C
      Else
         'Found a space or tab.
         'Set InArg flag to False.
         InArg = False
      End If
   Next I
   'Resize array just enough to hold arguments.
   ReDim Preserve ArgArray(NumArgs)
   'Return Array in Function name.
   GetCommandLine = ArgArray()
End Function




Function getFullExeName() As String
Dim path As String, exe As String
    exe = getExeName()
    If Not isAbsolute(exe) Then
        getFullExeName = getExePath() & "\" & exe
    Else
        getFullExeName = exe
    End If
End Function

Function getExePath() As String
Dim absolute As Boolean

    getExePath = getParam("path")
    
    If getExePath <> "" Then
        ' check if path is relative or absolute
    End If
    If Not isAbsolute(getExePath) Then
        getExePath = App.path & "\" & getExePath
    End If
    If Len(getExePath) = 0 Then
        getExePath = App.path
    End If
End Function


Function getExeName() As String
Dim hasExt As Boolean
Dim alias(3) As AliasInfo
Dim I As Integer

    alias(1).key = "prior": alias(1).value = "PriorN.exe"
    alias(2).key = "stime": alias(2).value = "stimeN.exe"
    alias(3).key = "rowmat": alias(3).value = "Rowmat_N.exe"
    
    getExeName = getParam("exe")
    For I = 1 To UBound(alias)
        If getExeName = alias(I).key Then
            getExeName = alias(I).value
            Exit Function
        End If
    Next I
    
    If Len(getExeName) > 0 Then
        If Not Right(getExeName, 4) = ".exe" Then
            getExeName = getExeName & ".exe"
        End If
    End If
End Function

Function getParam(key As String) As String
Dim I As Integer
    For I = 1 To UBound(parsedArgs)
        If parsedArgs(I) = "-" & key Then
            getParam = parsedArgs(I + 1)
            Exit Function
        End If
    Next I
End Function

Function isAbsolute(path As String) As Boolean
Dim firstChar As String, secondChar As String

    firstChar = Mid(path, 1, 1)
    secondChar = Mid(path, 2, 1)
    isAbsolute = False
    If Len(path) <= 2 Then
        Exit Function
    End If
    If firstChar = "\" Or secondChar = ":" Then
        isAbsolute = True
    End If
End Function
