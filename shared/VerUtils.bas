Attribute VB_Name = "VerUtils"
Option Explicit
'-- Win32 API Declarations
Private Declare Function LoadLibrary Lib "kernel32" _
        Alias "LoadLibraryA" _
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


'-- VB type casting!
Public Declare Sub CopyMemoryFromPointer Lib "kernel32" _
        Alias "RtlMoveMemory" (Destination As Any, _
        ByVal Source As Long, ByVal Length As Long)

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
                       ByRef loadpath As String, _
                       ByRef maj As Long, _
                       ByRef min As Long, _
                       ByRef rev As Long, _
                       ByRef build As Long) As Boolean
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

maj = -1: min = -1: rev = -1: build = -1

GetDllVersion = False           '-- pessimistic view
hDll = LoadLibrary(supFile)     '-- try to load the file
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
                End If
            End If
        End If
    End If
    '-- unload the library instance count...
    FreeLibrary (hDll)
End If
End Function

