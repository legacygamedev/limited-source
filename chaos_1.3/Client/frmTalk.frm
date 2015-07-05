VERSION 5.00
Begin VB.Form frmTalk 
   BorderStyle     =   0  'None
   Caption         =   "NPC Speech"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTalk.frx":0000
   ScaleHeight     =   253
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   360
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   7
      Top             =   720
      Width           =   555
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   15
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   8
         Top             =   15
         Width           =   495
      End
   End
   Begin VB.PictureBox Picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   4800
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5280
      Top             =   480
   End
   Begin VB.Label lblQuit 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label txtActual 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hello"
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
      Height          =   1575
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblChoice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choice 3"
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
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label lblChoice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choice 2"
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
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label lblChoice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choice 1"
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
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1815
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frmTalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public animi As Long

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
Dim sDc As Long

    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
    Next i
    
    sDc = DD_SpriteSurf.GetDC
    With Picsprites
        .Width = DDSD_Sprite.lWidth
        .Height = DDSD_Sprite.lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_SpriteSurf.ReleaseDC(sDc)
End Sub

Private Sub lblQuit_Click()
    Unload frmTalk
End Sub

Private Sub lblChoice_Click(index As Integer)
    If Speech(SpeechConvo1).num(SpeechConvo2).Respond < index + 1 Then Exit Sub
    
    If Speech(SpeechConvo1).num(SpeechConvo2).Responces(index + 1).Exit = 1 Then
        Unload frmTalk
        Exit Sub
    End If
    
    SpeechConvo2 = Speech(SpeechConvo1).num(SpeechConvo2).Responces(index + 1).GoTo
    
    If Speech(SpeechConvo1).num(SpeechConvo2).Script <> 0 Then
        Call SendData("SPEECHSCRIPT" & SEP_CHAR & Speech(SpeechConvo1).num(SpeechConvo2).Script & SEP_CHAR & END_CHAR)
    End If
    
    If Speech(SpeechConvo1).num(SpeechConvo2).Exit = 1 Then
        Unload frmTalk
        Exit Sub
    End If
    
    frmTalk.txtActual.Caption = Speech(SpeechConvo1).num(SpeechConvo2).Text
    frmTalk.txtActual.Left = Picture4.Left + Picture4.Width + 16
    
    If Speech(SpeechConvo1).num(SpeechConvo2).Respond > 0 Then
        frmTalk.lblChoice(0).Caption = Speech(SpeechConvo1).num(SpeechConvo2).Responces(1).Text
    Else
        frmTalk.lblChoice(0).Caption = ""
    End If
        
    If Speech(SpeechConvo1).num(SpeechConvo2).Respond > 1 Then
        frmTalk.lblChoice(1).Caption = Speech(SpeechConvo1).num(SpeechConvo2).Responces(2).Text
    Else
        frmTalk.lblChoice(1).Caption = ""
    End If
        
    If Speech(SpeechConvo1).num(SpeechConvo2).Respond > 2 Then
        frmTalk.lblChoice(2).Caption = Speech(SpeechConvo1).num(SpeechConvo2).Responces(3).Text
    Else
        frmTalk.lblChoice(2).Caption = ""
    End If
End Sub

Private Sub Timer1_Timer()
    If Speech(SpeechConvo1).num(SpeechConvo2).SaidBy = 0 Then
        Picpic.Width = SIZE_X
        Picpic.Height = SIZE_Y
        Picture4.Width = SIZE_X + 4
        Picture4.Height = SIZE_Y + 4
        Call BitBlt(Picpic.hDC, 0, 0, SIZE_X, SIZE_Y, Picsprites.hDC, animi * SIZE_X, Int(Npc(SpeechConvo3).Sprite) * SIZE_Y, SRCCOPY)
    Else
        Picpic.Width = SIZE_X
        Picpic.Height = SIZE_Y
        Picture4.Width = SIZE_X + 4
        Picture4.Height = SIZE_Y + 4
        Call BitBlt(Picpic.hDC, 0, 0, SIZE_X, SIZE_Y, Picsprites.hDC, animi * SIZE_X, Int(Player(MyIndex).Sprite) * SIZE_Y, SRCCOPY)
    End If
End Sub

Private Sub Timer2_Timer()
    animi = animi + 1
    If animi > 4 Then
        animi = 3
    End If
End Sub

