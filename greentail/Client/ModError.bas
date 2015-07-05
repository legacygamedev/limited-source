Attribute VB_Name = "ModError"
'=========================================================================================
'  VBErrorTrapDemo
'  How to Use:-
'  (1)Add Errfrm.frm, ModError.Bas, ModMSMail.bas _
'  Sysinfo.bas and ErrBitmap.res to your project.
'  (2)Add a refrence to Microsoft Scripting Runtime _
'  from Project --> references
'  (3)Add RichtextBox and winsock controls to your toolbox
'   from Project --> Components
'=========================================================================================
'  Coded By: Deepesh Agarwal
'  Published Date: 29/09/2003
'  WebSite: http://www.deepeshagarwal.tk
'  E-mail: agarwal_deepesh@indiatimes.com
'  Visit my site for Free-Software's like:
'  1). The-AdPolice - Blocks 17000+ adservers to save bandwidth
'  2). Dr. System -  Schedule Computer Maintainence - A must for every computer user
'  3). Service Controller XP (A Must For XP User) - Start,Stop,Pause and change startup type of 2000/XP services with recommended settings for different system config.
'   And Many More........
'=========================================================================================


'Call this function to show the error to the user

Public Sub ErrorHandler()
    Dim ver As String

    ver = App.EXEName & " v " & App.Major & "." & App.Minor & "." & App.Revision
    'Handle Errors according to number here

    'You must not show our error form for
    'non-critical or prev. known genral errors
    'like "Disk Not ready" error.

    Select Case Err.Number
        Case 0
            Unload Errfrm
        Case Else
            'Unhandled Error calling ShowErrorForm to display Error Info Form
            Call ShowErrorForm
    End Select
End Sub 'ErrorHandler()

Public Sub ShowErrorForm()

    Errfrm.Visible = True
    Errfrm.eno.Caption = Err.Number
    Errfrm.esrc.Caption = Err.Source
    Errfrm.edesc.Caption = Err.Description
    Errfrm.edll.Caption = Err.LastDllError
    Errfrm.lver.Caption = ver
    'Log Error to a file
    Call App.StartLogging(App.Path & "\ErrorLog.log", 2)
    'Actually Logging Error in file
    Call App.LogEvent(Err.Description, vbLogEventTypeError)
    'Log error into NT Event Log (If on NT)
    If Left$(GetOS, 11) = "Windows NT " Then
        Call App.StartLogging(App.Path & "\ErrorLog.log", 3)
    End If
    'Actually Logging Event in Event Viewer
    Call App.LogEvent(Err.Description, vbLogEventTypeError)
    'Closing Logging
    Call App.StartLogging(App.Path & "\ErrorLog.log", 1)
    'Show Loading Info message for slow machines
    Errfrm.RTDebug.Text = "Populating debug info, please Wait..."
    Errfrm.SendReport.Enabled = False
    Errfrm.nodonotsend.Enabled = False
    'Get and fill the debug info
    Errfrm.RTDebug.Text = GetSysInfo() 'GetSysInfo() In SysInfo Module
    'save this info as text file in app folder
    'use .Rtf because will not show whole info in textbox or notepad
    Call WriteDebugTofile(App.Path & "\DebugLog.rtf")

    'Done with report re-enable buttons
    Errfrm.SendReport.Enabled = True
    Errfrm.nodonotsend.Enabled = True
End Sub 'ShowErrorForm()

Public Sub WriteDebugTofile(fname As String)
    Dim fsys As New FileSystemObject, rb As String
    Dim FIO As Object
    If fsys.FileExists(fname) Then
        Set FIO = fsys.OpenTextFile(fname, ForAppending, True)
    Else
        Set FIO = fsys.CreateTextFile(fname, False, False)
    End If
    FIO.WriteBlankLines (2)
    FIO.Write (Errfrm.RTDebug.Text)
    Set FIO = Nothing
    Set fsys = Nothing
End Sub
