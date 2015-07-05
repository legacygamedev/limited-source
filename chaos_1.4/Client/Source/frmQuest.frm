VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmQuest 
   BorderStyle     =   0  'None
   Caption         =   "NPC Speech"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQuest.frx":0000
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   6000
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6480
      Top             =   240
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   480
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   3
      Top             =   960
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
         TabIndex        =   4
         Top             =   15
         Width           =   495
      End
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2220
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3916
      _Version        =   393217
      BackColor       =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmQuest.frx":3335
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblQuit 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblChoice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
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
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public animi As Long

Private Sub form_load()
Dim I As Long
Dim Ending As String
Dim sDc As Long

    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
    Next I
    
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
    Unload frmQuest
End Sub

Private Sub lblChoice_Click(Index As Integer)
Unload Me
End Sub

Private Sub Timer1_Timer()
Picpic.Width = SIZE_X
Picpic.Height = SIZE_Y
Picture4.Width = SIZE_X + 4
Picture4.Height = SIZE_Y + 4
Call BitBlt(Picpic.hDC, 0, 0, SIZE_X, SIZE_Y, Picsprites.hDC, animi * SIZE_X, Int(Player(MyIndex).Sprite) * SIZE_Y, SRCCOPY)
End Sub

Private Sub Timer2_Timer()
animi = animi + 1
    If animi > 4 Then
        animi = 3
    End If
End Sub
