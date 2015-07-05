Attribute VB_Name = "modGeneral"
Option Explicit

Function FileExists(FileName As String) As Boolean
    On Error GoTo HelpMe

    FileExists = (GetAttr(FileName) And vbDirectory) = 0

    Exit Function

HelpMe:
    Call MsgBox("The file specified could not be found.", vbOKOnly, "Error")
End Function
