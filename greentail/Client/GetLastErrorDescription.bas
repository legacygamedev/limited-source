Attribute VB_Name = "Module3"
'***********************************************************************************
'* Name:    GetLastErrorDescription()
'*
'* Purpose: To obtain the readable error description of GetLastError().
'*
'* Author:  Donny Dunning donny@submain.com
'*
'* Usage:   strVariable = GetLastErrorDescription()
'*
'* Return Value:    String value containing the Error Description of the GetLastError,
'*                  error code.
'*
'***********************************************************************************

' This function is useful for displaying the readable error description, while
' utilizing the API.  There is no need to call GetLastError, because the GetLastErrorDescription()
' function passes it as the parameter to retrieve the description for.  Here is a sample error handling
' routine that implements the GetLastErrorDescrtiption() function.

' On Error Goto ErrorHandler
' ... code
'
'   'ErrorHandling routine
' ErrorHandler:
' If Err.Number <> 0 then
'    MsgBox "Error Code: " & GetLastError & vbCrLf & _
'        "Error Description: " & GetLastErrorDescription, vbCritical
' End If

Option Explicit

'** Declares **

'FormatMessage Constants
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Declare Function GetLastError Lib "kernel32" () As Long

Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long

'** End Declares **

Public Function GetLastErrorDescription() As String
Dim sErrorMessage As String
Dim lErrorLen As Long

'Create buffer for Error Message by filling the string
'with spaces.
sErrorMessage = Space$(1024)

'Get the Error Description
lErrorLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, GetLastError, 0, _
    sErrorMessage, Len(sErrorMessage), 0)

    'If FormatMessage is successful, return the message
    If lErrorLen Then
        'Retrieve only the message by using the length returned
        'by FormatMessage
        GetLastErrorDescription = Left(sErrorMessage, lErrorLen)
    Else
        'The error is not defined
        GetLastErrorDescription = "Error (" & GetLastError & ") not defined."
    End If

End Function



