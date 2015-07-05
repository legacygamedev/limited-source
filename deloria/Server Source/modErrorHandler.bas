Attribute VB_Name = "modErrorHandler"
Public Sub ReportError(CurrFile As String, CurrProc As String, ErrNum As Long, ErrDesc As String)
  Dim FileNumber As Integer

  FileNumber = FreeFile 'Get an available file number so we don't mess with what's already open.

  Open App.Path & "\Errors.Log" For Append As #FileNumber 'Open the Error Log File.
  Print #FileNumber, "-------------------------------"
  Print #FileNumber, Date + Time 'Assuming the user of the client has their computer's time set correctly :)
  Print #FileNumber, "File: " & CurrFile
  Print #FileNumber, "Procedure: " & CurrProc
  Print #FileNumber, "Error #" & ErrNum
  Print #FileNumber, "Description: " & ErrDesc
  Print #FileNumber, ""
  Close #FileNumber
  
  Call TextAdd(frmServer.txtErrorLog, "-------------------------------", True)
  Call TextAdd(frmServer.txtErrorLog, Date + Time, True)
  Call TextAdd(frmServer.txtErrorLog, "File: " & CurrFile, True)
  Call TextAdd(frmServer.txtErrorLog, "Procedure: " & CurrProc, True)
  Call TextAdd(frmServer.txtErrorLog, "Error #" & ErrNum, True)
  Call TextAdd(frmServer.txtErrorLog, "Description: " & ErrDesc, True)
  Call TextAdd(frmServer.txtErrorLog, "", True)
End Sub
