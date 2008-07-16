Attribute VB_Name = "modWin32Api"
Option Explicit
'Copyright 2003 by AQUA TERRA Consultants

' ##MODULE_NAME modWin32Api
' ##MODULE_DATE January 1, 2003
' ##MODULE_AUTHOR Mark Gray and Jack Kittle of AQUA TERRA CONSULTANTS
' ##MODULE_SUMMARY Declarations for miscellaneous Win32 API functions.

  '##SUMMARY Identifies the ID of the module that is actually executing.
  '##PARAM lpModuleName I Specifies the file name of the module to load.
  '##RETURNS The handle of an already loaded DLL.
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" _
    (ByVal lpModuleName As String) As Long
  '##SUMMARY Retrieves the full path and filename for the executable file containing _
      the specified module.
  '##PARAM hModule I Identifies the module whose executable filename is being requested. _
      If this parameter is NULL, GetModuleFileName returns the path for the file used to _
      create the calling process.
  '##PARAM lpFileName O Points to a buffer that is filled in with the path and filename _
      of the given module.
  '##PARAM nSize I Specifies the length, in characters, of the lpFilename buffer. If the _
      length of the path and filename exceeds this limit, the string is truncated.
  '##RETURNS If the function succeeds, the return value is the length, in characters, _
      of the string copied to the buffer. If the function fails, the return value is zero.
Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" _
    (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
  '##SUMMARY Copies data from a named or anonymous pipe into a buffer without _
    removing it from the pipe. It also returns information about data in the pipe.
  '##PARAM hNamedPipe I The handle to the pipe.
  '##PARAM lpBuffer O Points to the buffer that receives the data read from the pipe.
  '##PARAM nBufferSize I The size of the buffer.
  '##PARAM lpBytesRead O Points to the number of bytes read.
  '##RETURNS If the function succeeds, the return value is nonzero. _
    If the function fails, the return value is zero.
Declare Function WinExec Lib "kernel32" _
    (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
  '##SUMMARY Runs a specified application.
  '##PARAM lpCmdLine I Points to a null-terminated character string that contains the _
    command line (filename plus optional parameters) for the application to be executed..
  '##PARAM uCmdShow I Specifies how a Windows-based application window is to be shown and _
    is used to supply the wShowWindow member of the STARTUPINFO parameter to the CreateProcess function.
  '##RETURNS If the function succeeds, the return value is greater than 31. _
    If the function fails, the return value is one of the following error values: _
    0: The system is out of memory or resources. _
    ERROR_BAD_FORMAT: The .EXE file is invalid (non-Win32 .EXE or error in .EXE image). _
    ERROR_FILE_NOT_FOUND: The specified file was not found. _
    ERROR_PATH_NOT_FOUND: The specified path was not found. .
Declare Function PeekNamedPipe Lib "kernel32" _
    (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
  '##SUMMARY Reads data from a file, starting at the position indicated by the file pointer.
  '##PARAM hFile I Identifies the file to be read.
  '##PARAM lpBuffer O Points to the buffer that receives the data read from the file.
  '##PARAM nNumberOfBytesToRead I Specifies the number of bytes to be read from the file.
  '##PARAM lpNumberOfBytesRead O Points to the number of bytes read.
  '##PARAM lpOverlapped O Points to an OVERLAPPED structure.
  '##RETURNS If the function succeeds, the return value is nonzero. _
    If the return value is nonzero and the number of bytes read is zero, _
    the file pointer was beyond the current end of the file at the time of _
    the read operation. However, if the file was opened with FILE_FLAG_OVERLAPPED _
    and lpOverlapped is not NULL, the return value is FALSE and GetLastError _
    returns ERROR_HANDLE_EOF when the file pointer goes beyond the current end of file. _
    If the function fails, the return value is zero.
Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long


