Attribute VB_Name = "VerUtils"
Option Explicit
'-- Win32 API Declarations
Private Declare Function LoadLibrary Lib "kernel32" _
        Alias "LoadLibraryA" _
        (ByVal lpLibFileName As String) As Long

Private Declare Function LoadLibraryEx Lib "kernel32" _
        Alias "LoadLibraryExA" _
        (ByVal lpLibFileName As String _
            , ByVal hFile As Long _
            , ByVal dwFlags As Long _
        ) As Long


Private Declare Function GetModuleHandle Lib "kernel32" _
        Alias "GetModuleHandleA" _
        (ByVal lpLibFileName As String) As Long
        
Private Declare Function FindResource Lib "kernel32" _
        Alias "FindResourceA" (ByVal hInstance As Long, _
        ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function FindResourceI Lib "kernel32" _
        Alias "FindResourceA" (ByVal hInstance As Long, _
        ByVal lpName As Long, ByVal lpType As Long) As Long
Private Declare Function LoadResource Lib "kernel32" _
        (ByVal hInstance As Long, ByVal hResInfo As Long) _
        As Long
Private Declare Function LockResource Lib "kernel32" _
        (ByVal hResData As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" _
        Alias "GetModuleFileNameA" (ByVal hModule As Long, _
        ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" _
        (ByVal hLibModule As Long) As Long
Private Declare Function GetLastError Lib "kernel32" _
        () As Long


'-- VB type casting!
Public Declare Sub CopyMemoryFromPointer Lib "kernel32" _
        Alias "RtlMoveMemory" (Destination As Any, _
        ByVal Source As Long, ByVal length As Long)


'--------------Shell API and Constants----------
Private Const WAIT_FAILED = -1&
Private Const WAIT_OBJECT_0 = 0
Private Const WAIT_ABANDONED = &H80&
Private Const WAIT_ABANDONED_0 = &H80&
Private Const WAIT_TIMEOUT = &H102&
Private Const INFINITE = &HFFFFFFFF       '  Infinite timeout
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const SYNCHRONIZE = &H100000
Private Const LOAD_LIBRARY_AS_DATAFILE = &H2

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long


Private Declare Function GetVersion Lib "kernel32" () As Long


Type VS_VERSIONINFO
  wLength As Integer
  wValueLength As Integer
  wType As Integer
  szKey(29) As Byte '-- contains the UNICODE String
                    '-- "VS_VERSION_INFO" and in VB that's (29)!
  '-- Padding1(?) As Byte '-- this is the dynamic element
End Type

'-- This UDT is defined in <windows.h>, and here
'-- is the VB translation
Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long
    dwFileVersionMS As Long
    dwFileVersionLS As Long
    dwProductVersionMS As Long
    dwProductVersionLS As Long
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type



Function GetDllVersion(ByVal supFile As String, _
                    ByRef info As VersionInfo _
                    ) As Boolean
'---------------------------------------------------------------
'   This function uses the supFile parameter to do a dynamic
'   load of the dll name.  Once loaded, the version resource
'   is queried for the file version number.  The version info
'   as well as the full load path are returned to the caller.
'
'   parameters:
'       supFile     [in] String         Dll file name
'       loadpath    [out] String        full path of file
'       maj         [out] integer       File version info
'       min         [out] integer       File version info
'       rev         [out] integer       File version info
'       build       [out] integer       File version info
'---------------------------------------------------------------
Dim hDll As Long
Dim retval As Long

Dim loadpath As String
Dim maj As Long
Dim min As Long
Dim rev As Long
Dim build As Long

maj = -1: min = -1: rev = -1: build = -1

GetDllVersion = False           '-- pessimistic view

Dim errorCode As Long
Dim hFile As Long
hDll = LoadLibraryEx(supFile, hFile, LOAD_LIBRARY_AS_DATAFILE)     '-- try to load the file
errorCode = GetLastError

If (hDll) Then
    '-- get the load path
    Dim tmpPath As String
    tmpPath = String(512, 0)    '-- buffer for API call
    retval = GetModuleFileName(hDll, tmpPath, 511)
    If (retval) Then
        '-- make sure there is a null(0)
        If (InStr(tmpPath, Chr$(0)) > 0) Then
            '-- trim the returned string
            loadpath = Left$(tmpPath, InStr(tmpPath, Chr$(0)) - 1)
        End If
    End If
    '-- find the version resource
    Dim hRes As Long
    hRes = FindResourceI(hDll, 1, 16)
    If (hRes) Then
        Dim hGbl As Long
        hGbl = LoadResource(hDll, hRes)
        If (hGbl) Then
            Dim lpRes As Long
            lpRes = LockResource(hGbl)
            If (lpRes) Then
                '-- lpRes is a memory pointer to file's
                '-- version resource!
                Dim verinfo As VS_VERSIONINFO   '-- make space
                '-- copy what we know of the verinfo UDT
                CopyMemoryFromPointer verinfo, lpRes, _
                   Len(verinfo)
                '-- test if we have a VS_FIXEDFILEINFO struct
                If (verinfo.wValueLength > 0) Then
                    '-- lpRes is the pointer to the locked
                    '-- resource and we need to position just
                    '-- past the known data elements...

                    '-- set the pointer to Padding1(0)
                    lpRes = lpRes + Len(verinfo)

                    '-- Since the actual Padding1 element is
                    '-- unknown in size we must loop and
                    '-- increment the memory pointer until it
                    '-- it is on a 32bit (DWORD) boundry.
                    While ((lpRes And &H4) <> 0)
                        lpRes = lpRes + 1
                    Wend

                    '-- create a variable to hold the fixed
                    '-- version info
                    Dim fInfo As VS_FIXEDFILEINFO
                    
                    '-- copy the fixed file info now
                    CopyMemoryFromPointer fInfo, lpRes, Len(fInfo)
                    
                    '-- extract the version data, and we're done!
                    maj = fInfo.dwFileVersionMS / 65535
                    min = fInfo.dwFileVersionMS And &H7FFF
                    rev = fInfo.dwFileVersionLS / 65535
                    build = fInfo.dwFileVersionLS And &H7FFF
                    GetDllVersion = True    '-- SUCCESS!!!
                    info.maj = maj: info.min = min: info.rev = rev: info.bld = build: info.path = loadpath
                End If
            End If
        End If
    End If
    '-- unload the library instance count...
    FreeLibrary (hDll)
Else
    erro " Ошибка при определении версии файла. GetLastError = " & errorCode
End If
End Function



'Shell and wait for a process to finish

'One of the limitation of using the Shell function is that it is asynchronous. Below are a couple of different methods of shelling processes and waiting until they are finished (or initialised):


'Purpose   :    Shells a process synchronised i.e. Holds execution until application has closed.
'Inputs    :    sCommandLine        =   The Command line to run the application e.g. "Notepad.exe"
'               State               =   The Window State to run of the shelled program (A Long)
'Outputs   :    Returns the Process Handle
'Notes     :    Have noticed side effects. Other applications like Internet Explorer seem to be effected by this.

Function ShellAndHold(sCommandLine As String, Optional lState As Long = vbNormalFocus) As Long
    Dim lRetVal As Long, FileToOpen As String
    
    'Check to see that the file exists
    If FileExists(sCommandLine) Then
        'Add double quotes around the path (otherwise you can't use spaces in the path)
        If Left$(sCommandLine, 1) <> Chr(34) Then
            sCommandLine = Chr(34) & sCommandLine
        End If
        If Right$(sCommandLine, 1) <> Chr(34) Then
            sCommandLine = sCommandLine & Chr(34)
        End If
    End If
    
    'Start the shell
    lRetVal = Shell(sCommandLine, lState)
    'Open the process
    ShellAndHold = OpenProcess(SYNCHRONIZE, False, lRetVal)
    
    'Wait for the process to complete
    lRetVal = WaitForSingleObject(ShellAndHold, INFINITE)
    lRetVal = CloseHandle(ShellAndHold)
End Function


'Purpose   :    Holds execution until application has closed.
'Inputs    :    sFilePath       =   The path to the application to run e.g. "Notepad.exe"
'               [sCommandLine]  =   Any command line arguments
'               [lState]        =   The Window State to run of the shelled program (A Long)
'               [lMaxTimeOut]   =   The maximum amount of time to wait for the process to finish (in secs).
'                                   -1 = infinate
'Outputs   :    Returns the True if failed open a process or complete within the specified timeout.
'Notes     :    Similiar to ShellAndHold, but will not get any 'spiking' effects using this method.


Function ShellAndWait(sFilePath As String, Optional sCommandLine, Optional lState As VbAppWinStyle = vbNormalFocus, Optional lMaxTimeOut As Long = -1) As Boolean
    Dim lRetVal As Long, siStartTime As Single, lProcID As Long

    'Check to see that the file exists
    If FileExists(sFilePath) Then
        'Add double quotes around the path (otherwise you can't use spaces in the path)
        If Left$(sFilePath, 1) <> Chr(34) Then
            sFilePath = Chr(34) & sFilePath
        End If
        If Right$(sFilePath, 1) <> Chr(34) Then
            sFilePath = sFilePath & Chr(34)
        End If
    End If
    
    'Start the shell
    lRetVal = Shell(Trim$(sFilePath + " " + sCommandLine), lState)
    'Open the process
    lProcID = OpenProcess(SYNCHRONIZE, True, lRetVal)
    
    siStartTime = Timer
    Do
        lRetVal = WaitForSingleObject(lProcID, 0)
        If lRetVal = WAIT_OBJECT_0 Then
            'Finished process
            lRetVal = CloseHandle(lProcID)
            ShellAndWait = False
            Exit Do
        ElseIf lRetVal = WAIT_FAILED Then
            lRetVal = CloseHandle(lProcID)
            'Failed to open process
            ShellAndWait = True
            Exit Do
        End If
        Sleep 100
        If lMaxTimeOut > 0 Then
            'Check timeout has not been exceeded
            If siStartTime + lMaxTimeOut < Timer Then
                'Failed, timeout exceeded
                lRetVal = CloseHandle(lProcID)
                ShellAndWait = True
            End If
        End If
    Loop
End Function


'Purpose   :    Holds execution until application has finished opening
'Inputs    :    sCommandLine     =   The Command line to run the application e.g. "Notepad.exe"
'               lState           =   The Window State to run of the shelled program (A Long)
'Outputs   :    Returns the Process Handle
'Notes     :    Use this when you want to wait for an application to finishing opening before proceeding
'               The side effects mentioned in ShellAndHold will be negligible since the most applications
'               load in under 5 seconds.


Function ShellAndWaitReady(sCommandLine As String, Optional lState As Long = vbNormalFocus) As Long
    Dim lhProc As Long
    
    If Left$(sCommandLine, 1) <> Chr(34) Then
        sCommandLine = Chr(34) & sCommandLine
    End If
    If Right$(sCommandLine, 1) <> Chr(34) Then
        sCommandLine = sCommandLine & Chr(34)
    End If
    lhProc = Shell(sCommandLine, lState)
    'Wait for the process to initialize
    Call WaitForInputIdle(lhProc, INFINITE)
    'Return the handle
    ShellAndWaitReady = lhProc
End Function


'Purpose     :  Checks if a file exists
'Inputs      :  sFilePathName                   The path and file name e.g. "C:\Autoexec.bat"
'Outputs     :  Returns True if the file exists


Function FileExists(sFilePathName As String) As Boolean
    
    On Error GoTo ErrFailed
    If Len(sFilePathName) Then
        If (GetAttr(sFilePathName) And vbDirectory) < 1 Then
            'File Exists
            FileExists = True
        End If
    End If
    Exit Function
    
ErrFailed:
    'File Exists
    FileExists = False
    On Error GoTo 0
End Function

'Purpose     :  Converts a File Name and Path to a Path
'Inputs      :  sFilePathName                   The path and file name e.g. "C:\Autoexec.bat"
'Outputs     :  Returns the path


Function PathFileToPath(sFilePathName As String) As String
    Dim ThisChar As Long

    For ThisChar = 0 To Len(sFilePathName) - 1
        If Mid$(sFilePathName, Len(sFilePathName) - ThisChar, 1) = "\" Then
            PathFileToPath = Left$(sFilePathName, Len(sFilePathName) - ThisChar)
            Exit For
        End If
    Next
End Function


Private Function getRepositoryPathFile(exeName As String, ByRef version As VersionInfo) As String
    getRepositoryPathFile = PathFileToPath(version.path) & "\" & version.maj & "\" & version.min & "\" & exeName & ".exe" _
        & "." & version.maj & "." & version.min & "." & version.bld
End Function


Public Sub GetHiLoByte(X As Integer, LoByte As Integer, HiByte As Integer)
    LoByte = X And &HFF&
    HiByte = X \ &H100
End Sub


Public Sub GetHiLoWord(X As Long, LoWord As Integer, HiWord As Integer)
    LoWord = CInt(X And &HFFFF&)
    HiWord = CInt(X \ &H10000)
End Sub


Public Sub printWindowsVersion()
    Dim WinMajor As Integer
    Dim WinMinor As Integer
    Dim DosMajor As Integer
    Dim DosMinor As Integer
    Dim RetLong As Long
    Dim LoWord As Integer
    Dim HiWord As Integer

    RetLong = GetVersion()
    Call GetHiLoWord(RetLong, LoWord, HiWord)

    Call GetHiLoByte(LoWord, WinMajor, WinMinor)
    Call GetHiLoByte(HiWord, DosMinor, DosMajor)

    dbg "Windows version:" & WinMajor & "." & WinMinor
    dbg "DOS version:" & DosMajor & "." & DosMinor

End Sub


