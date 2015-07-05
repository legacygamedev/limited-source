Attribute VB_Name = "modDeclares"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' get window that's displayed
Public Declare Function GetForegroundWindow Lib "user32" () As Long

' used with reading / writing INI files
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

' get system uptime in milliseconds
Public Declare Sub GetSysTimeMS Lib "KERNEL32.DLL" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Text declares
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

' For clear functions
Public Declare Sub ZeroMemory Lib "KERNEL32.DLL" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' Makes richtextboxes transparent!
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
