Attribute VB_Name = "Common"
Option Explicit

Sub Main()
Dim CmdLine As String, path As String, exe As String
Dim exeHandle As Double

'-- Main application entry point.
Dim myFilename As String
Dim myLoadpath As String
Dim maj As Long, min As Long, rev As Long, bld As Long

parsedArgs = GetCommandLine
    
myFilename = getFullExeName()

'    exeHandle = Shell(exe, vbNormalFocus)
'    MsgBox "after start... " & exe
    

If (Len(myFilename) <= 0) Then
    MsgBox "VerRes requires a filename as a parameter"
    Exit Sub
End If

If (GetDllVersion(myFilename, myLoadpath, maj, min, rev, bld)) Then
    MsgBox "File " & myFilename & " is vesion " & CStr(maj) & _
           "." & CStr(min) & "." & CStr(rev) & "." & CStr(bld), _
           vbInformation, "VerRes"
Else
    MsgBox "Couldn't get the version of " & myFilename, vbExclamation, _
           "VerRes"
End If


End Sub

