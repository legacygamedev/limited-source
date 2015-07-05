VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Begin VB.Form frmFlash 
   BorderStyle     =   0  'None
   Caption         =   "Flash Event"
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFlash.frx":0000
   ScaleHeight     =   631
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   857
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Check 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   6000
      Left            =   2280
      TabIndex        =   1
      Top             =   1725
      Width           =   8250
      _cx             =   14552
      _cy             =   10583
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
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   12150
      TabIndex        =   0
      Top             =   240
      Width           =   435
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Private Sub Check_Timer()
    If Flash.CurrentFrame > 0 Then
        If Flash.CurrentFrame >= Flash.TotalFrames - 1 Then
            Flash.FrameNum = 0
            Flash.Stop
            Check.Enabled = False
            Call PutVar(App.Path & "\config.ini", "CONFIG", "Music", frmMirage.chkmusic.Value)
            MusicOn = frmMirage.chkmusic.Value
            Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
            Call PutVar(App.Path & "\config.ini", "CONFIG", "Sound", frmMirage.chkSound.Value)
            SoundOn = frmMirage.chkSound.Value
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
        If FileExist("GUI\flash" & Ending) Then frmFlash.Picture = LoadPicture(App.Path & "\GUI\flash" & Ending)
    Next I
End Sub

Private Sub Label1_Click()
    Flash.FrameNum = 0
    Flash.Stop
    Check.Enabled = False
    Call PutVar(App.Path & "\config.ini", "CONFIG", "Music", frmMirage.chkmusic.Value)
    MusicOn = frmMirage.chkmusic.Value
    Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
    Call PutVar(App.Path & "\config.ini", "CONFIG", "Sound", frmMirage.chkSound.Value)
    SoundOn = frmMirage.chkSound.Value
    Unload Me
End Sub

