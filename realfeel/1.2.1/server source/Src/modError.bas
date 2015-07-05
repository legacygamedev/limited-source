Attribute VB_Name = "modError"
Option Explicit

Public Sub ReportError(CurrFile As String, CurrProc As String, ErrNum As Long, ErrDesc As String)
On Error Resume Next
Dim f As Integer, c As Byte

  f = FreeFile 'Get an available file number so we don't mess with what's already open.

  Open App.Path & "\Errors.Log" For Append As #f 'Open the Error Log File.
    Print #f, "-------------------------------"
    Print #f, Date + Time 'Assuming the user of the client has their computer's time set correctly :)
    Print #f, "File: " & CurrFile
    Print #f, "Procedure: " & CurrProc
    Print #f, "Error #" & ErrNum
    Print #f, "Description: " & ErrDesc
    Print #f, ""
  Close #f
  
  ' Add a notice to the server's main page
  Call TextAdd(frmServer.txtText, "Warning: An error has occured at " & Date & " " & Time & "!", True)
  Call TextAdd(frmServer.txtText, "Please see the error report for more details! A copy of the error has been printed for reference.", True)

  ' Make the error noted
  c = MsgBox(Date & " - " & Time & vbCrLf & "File: " & CurrFile & vbCrLf & "Procedure: " & CurrProc & vbCrLf & "Error #" & ErrNum & vbCrLf & "Description: " & ErrDesc & vbCrLf & vbCrLf & "This may or may not be a fatal error! Would you like the server to attempt to continue running?", vbYesNo, "RealFeel Server")
  If c = vbNo Then Call DestroyServer
End Sub

Public Sub SimulateErrors()
Dim n As Long, quit As Byte
For n = 1 To 20
Error n
Next n
quit = MsgBox("Stop throwing errors?", vbYesNo)
If quit = vbYes Then Exit Sub
For n = 21 To 40
Error n
Next n
quit = MsgBox("Stop throwing errors?", vbYesNo)
If quit = vbYes Then Exit Sub
For n = 41 To 60
Error n
Next n
quit = MsgBox("Stop throwing errors?", vbYesNo)
If quit = vbYes Then Exit Sub
For n = 61 To 80
Error n
Next n
quit = MsgBox("Stop throwing errors?", vbYesNo)
If quit = vbYes Then Exit Sub
For n = 81 To 100
Error n
Next n
quit = MsgBox("Stop throwing errors?", vbYesNo)
If quit = vbYes Then Exit Sub
For n = 101 To 120
Error n
Next n
quit = MsgBox("Stop throwing errors?", vbYesNo)
If quit = vbYes Then Exit Sub
For n = 121 To 140
Error n
Next n
quit = MsgBox("Stop throwing errors?", vbYesNo)
If quit = vbYes Then Exit Sub
For n = 141 To 160
Error n
Next n
quit = MsgBox("Stop throwing errors?", vbYesNo)
If quit = vbYes Then Exit Sub
For n = 161 To 180
Error n
Next n
quit = MsgBox("Stop throwing errors?", vbYesNo)
If quit = vbYes Then Exit Sub
For n = 181 To 200
Error n
Next n
End Sub


