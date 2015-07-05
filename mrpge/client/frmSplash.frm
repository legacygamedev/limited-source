VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6000
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timMoveNext 
      Interval        =   1000
      Left            =   3135
      Top             =   2760
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flaIntro 
      Height          =   5925
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7500
      _cx             =   13229
      _cy             =   10451
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   -1  'True
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
    Call flaIntro.LoadMovie(0, App.Path & "\data\intro.mov")
End Sub

Private Sub Form_Click()
'tickCounter1.Enabled = False
'    Call modGameLogic.Main
'    Unload Me
End Sub

Private Sub Form_Load()
'notInGame = True
'    Call PlayMP3("splash")
'    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'    lblCopy.Caption = "© Copyright 2004, SGN Games."
    

End Sub

Private Sub timMoveNext_Timer()
Debug.Print flaIntro.FrameNum
    If flaIntro.FrameNum >= 299 Then
    flaIntro.Stop
    flaIntro.StopPlay
    Call flaIntro.LoadMovie(0, vbNull)
        Call modGameLogic.Main
        Unload Me
        timMoveNext.Enabled = False
    End If
End Sub
