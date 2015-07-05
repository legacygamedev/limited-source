VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9c.ocx"
Begin VB.Form frmFlash 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flash Event"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   Icon            =   "frmFlash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Check 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   1560
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   5280
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5730
      _cx             =   10107
      _cy             =   9313
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   5520
      Width           =   5535
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check_Timer()
    If Flash.CurrentFrame > 0 Then
        If Flash.CurrentFrame >= Flash.TotalFrames - 1 Then
            Flash.FrameNum = 0
            Flash.Stop
            Check.Enabled = False
            WriteINI "CONFIG", "Music", frmMirage.chkMusic.Value, App.Path & "\config.ini"
            Call PlayBGM(Trim$(Map(GetPlayerMap(MyIndex)).music))
            WriteINI "CONFIG", "Sound", frmMirage.chksound.Value, App.Path & "\config.ini"
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim Ending As String
    For i = 1 To 3
        If i = 1 Then
            Ending = ".gif"
        End If
        If i = 2 Then
            Ending = ".jpg"
        End If
        If i = 3 Then
            Ending = ".png"
        End If

        If FileExists("GUI\FlashTheatre" & Ending) Then
            frmFlash.Picture = LoadPicture(App.Path & "\GUI\FlashTheatre" & Ending)
        End If
    Next i
End Sub

Private Sub Label1_Click()
    Flash.FrameNum = 0
    Flash.Stop
    Check.Enabled = False
    WriteINI "CONFIG", "Music", frmMirage.chkMusic.Value, App.Path & "\config.ini"
    Call PlayBGM(Trim$(Map(GetPlayerMap(MyIndex)).music))
    WriteINI "CONFIG", "Sound", frmMirage.chksound.Value, App.Path & "\config.ini"
    Unload Me
End Sub

