Attribute VB_Name = "modDeclares"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

' Text API
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

' Get system uptime in milliseconds
Public Declare Sub GetSysTimeMS Lib "KERNEL32.DLL" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)

'For Clear functions
Public Declare Sub ZeroMemory Lib "KERNEL32.DLL" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

