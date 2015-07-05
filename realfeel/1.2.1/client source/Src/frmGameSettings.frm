VERSION 5.00
Begin VB.Form frmGameSettings 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dual Solace (Game Settings)"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGameSettings.frx":0000
   ScaleHeight     =   3120
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkNPCNames 
      BackColor       =   &H002F3336&
      Caption         =   "Turn on NPC Names"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   285
      TabIndex        =   6
      Top             =   2415
      Width           =   1380
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5295
      Picture         =   "frmGameSettings.frx":3E604
      ScaleHeight     =   330
      ScaleWidth      =   840
      TabIndex        =   5
      Top             =   765
      Width           =   840
   End
   Begin VB.PictureBox picSubmit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4305
      Picture         =   "frmGameSettings.frx":3F4B6
      ScaleHeight     =   375
      ScaleWidth      =   1020
      TabIndex        =   4
      Top             =   735
      Width           =   1020
   End
   Begin VB.HScrollBar scrTempo 
      Height          =   255
      LargeChange     =   500
      Left            =   3000
      Max             =   10000
      Min             =   10
      SmallChange     =   10
      TabIndex        =   3
      Top             =   2400
      Value           =   1000
      Width           =   2175
   End
   Begin VB.HScrollBar scrAudio 
      Height          =   255
      LargeChange     =   500
      Left            =   3000
      Max             =   0
      Min             =   -10000
      TabIndex        =   2
      Top             =   1500
      Width           =   2175
   End
   Begin VB.CheckBox chkMusic 
      BackColor       =   &H002F3336&
      Caption         =   "Turn on music"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   285
      MaskColor       =   &H8000000D&
      TabIndex        =   1
      Top             =   1425
      Width           =   1485
   End
   Begin VB.CheckBox chkSound 
      BackColor       =   &H002F3336&
      Caption         =   "Turn on sound"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   285
      TabIndex        =   0
      Top             =   1950
      Width           =   1425
   End
End
Attribute VB_Name = "frmGameSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PicArray() As VB.PictureBox

Public Sub MakePic(ByVal n As Long)
    Set PicArray(n) = Controls.Add("VB.PictureBox", "PicArray" & CStr(n), Me)
    PicArray(n).Appearance = 0
    PicArray(n).BorderStyle = 0
    PicArray(n).AutoRedraw = True
    PicArray(n).AutoSize = True
    PicArray(n).Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "PicFrame" & n & "Target"))
    PicArray(n).Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "PicFrame" & n & "X"))
    PicArray(n).top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "PicFrame" & n & "Y"))
    PicArray(n).Visible = True
End Sub

Public Sub SetArray(ByVal num As Long)
    ReDim PicArray(1 To num) As VB.PictureBox
End Sub

Private Sub picCancel_Click()
    Unload Me
End Sub

Private Sub picSubmit_Click()
If frmGameSettings.chkMusic.Value = 1 Then
    Call PutVar(App.Path & "\data.dat", "Settings", "Music", "1")
Else
    Call PutVar(App.Path & "\data.dat", "Settings", "Music", "0")
    Call DirectMusic.StopMusic
End If

If frmGameSettings.chkSound = 1 Then
    Call PutVar(App.Path & "\data.dat", "Settings", "Sound", "1")
Else
    Call PutVar(App.Path & "\data.dat", "Settings", "Sound", "0")
End If

If frmGameSettings.chkNPCNames.Value = 1 Then
    Call PutVar(App.Path & "\data.dat", "Settings", "NPC_Names", "1")
Else
    Call PutVar(App.Path & "\data.dat", "Settings", "NPC_Names", "0")
End If

Call PutVar(App.Path & "\data.dat", "Settings", "Audio", CStr(scrAudio.Value))
Call PutVar(App.Path & "\data.dat", "Settings", "Tempo", CStr(CLng(scrTempo.Value) * 10))

'Unload the form
frmGameSettings.Visible = False
Call SendData("NEEDMAP" & SEP_CHAR & "YES" & SEP_CHAR & END_CHAR)
Unload frmGameSettings
End Sub

Private Sub Form_Load()
If GetVar(App.Path & "\data.dat", "Settings", "Music") = 1 Then
    frmGameSettings.chkMusic.Value = 1
Else
    frmGameSettings.chkMusic.Value = 0
End If

If GetVar(App.Path & "\data.dat", "Settings", "Sound") = 1 Then
    frmGameSettings.chkSound.Value = 1
Else
    frmGameSettings.chkSound.Value = 0
End If

If GetVar(App.Path & "\data.dat", "Settings", "NPC_Names") = 1 Then
    frmGameSettings.chkNPCNames.Value = 1
Else
    frmGameSettings.chkNPCNames.Value = 0
End If

'scrAudio.Value = CLng(GetVar(App.Path & "\data.dat", "Settings", "Audio"))
'lblAudio.Caption = GetVar(App.Path & "\data.dat", "Settings", "Audio")
'scrTempo.Value = CLng(GetVar(App.Path & "\data.dat", "Settings", "Tempo"))
'lblTempo.Caption = GetVar(App.Path & "\data.dat", "Settings", "Tempo")

End Sub

Private Sub scrAudio_Change()
Dim n As Byte
'lblAudio.Caption = CInt(100 - Abs(100 * (scrAudio.Value / 10000))) & "%"
End Sub

Private Sub scrTempo_Change()
Dim n As Byte
'lblTempo.Caption = (10 * CLng(scrTempo.Value)) / 1000 & "khz"
End Sub
