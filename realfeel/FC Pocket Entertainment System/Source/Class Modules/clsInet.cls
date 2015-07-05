VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Variable information holders
Private Size As Long
Private Retrieved As Long

Public Sub Download(ByVal Address As String, ByVal File As String)
Dim Size As Long, Remaining As Long
Dim f As Integer, choice As Integer
Dim FilePart() As Byte

' Check for the file existing and ask overwrite question
If Exists(App.Path & "\Files\" & File) Then
    choice = MsgBox("The file: " & vbCrLf & vbCrLf & App.Path & "\Files\" & File & vbCrLf & vbCrLf & "already exists!" & vbCrLf & vbCrLf & "Would you like to overwrite this file?", vbYesNo, "FriendCodes Pocket Entertainment System")
    If choice = vbNo Then Exit Sub
    If choice = vbYes Then Call Kill(App.Path & "\Files\" & File)
End If


f = FreeFile

' Set up the download form
frmDownload.lblName.Caption = Address & File
frmDownload.Visible = True

' Make sure the animation boxes are correctly placed
frmDownload.picAnimWalk1.Left = 120
frmDownload.picAnimWalk1.Top = 1080
frmDownload.picAnimWalk2.Left = 120
frmDownload.picAnimWalk2.Top = 1080
frmDownload.picAnimJump.Left = 120
frmDownload.picAnimJump.Top = 1080
frmDownload.picAnimWalk1.Visible = True

' Enable the timer
frmDownload.tmrAnimation.Enabled = True

' Commence downloading the file
frmMenu.cInet.Execute Address & File, "GET"

Do While frmMenu.cInet.StillExecuting
    DoEvents
Loop

' Get the file size
Size = Val(frmMenu.cInet.GetHeader("Content-length"))
Retrieved = 0

Debug.Print "SIZE: " & Size

' Open file to write binary retrieved from internet download
Open App.Path & "\Files\" & File For Binary As #f
    Do Until Retrieved >= Size
        If Size - Retrieved > 1023 Then
            FilePart = frmMenu.cInet.GetChunk(1024, icByteArray)
            Retrieved = Retrieved + 1024
        Else
            FilePart = frmMenu.cInet.GetChunk(Size - Retrieved, icByteArray)
            Retrieved = Size
        End If
     Put #f, , FilePart

    ' Display download data
    If Size > 0 Then
        frmDownload.lblRetrieved.Caption = Retrieved
        frmDownload.lblSize.Caption = Size
    End If
    
    Loop
Close #f

' Disable the timer so it won't lag
frmDownload.tmrAnimation.Enabled = False

' Close the form
Unload frmDownload
End Sub

'Public Function OpenPage(ByVal Address As String) As String
'OpenPage = Inet.OpenURL(Address, icString)
'End Function
