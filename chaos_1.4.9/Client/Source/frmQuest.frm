VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
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
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Abandon Quest"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "I Choose Not to Accept"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Accept Quest"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
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
      TabIndex        =   4
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
      TabIndex        =   2
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
         TabIndex        =   3
         Top             =   15
         Width           =   495
      End
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2220
      Left            =   1080
      TabIndex        =   5
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
   Begin VB.Label lblChoice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
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
      TabIndex        =   1
      Top             =   0
      Width           =   255
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
      TabIndex        =   0
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

Private Sub cmdNo_Click()
InQuestMenu = 0
Unload Me
End Sub

Private Sub cmdQuit_Click()
InQuestMenu = 0
Call SendData("STOPKILLQUEST" & SEP_CHAR & END_CHAR)
cmdQuit.Visible = False
lblChoice.Visible = True
End Sub

Private Sub cmdYes_Click()
Call SendData("ACCEPTQUEST" & SEP_CHAR & CurrentQuestNum & SEP_CHAR & CurrentQuestNpcNum & SEP_CHAR & END_CHAR)
cmdYes.Visible = False
cmdNo.Visible = False
lblChoice.Visible = True
CurrentQuestNum = 0
CurrentQuestNpcNum = 0
InQuestMenu = 0
End Sub

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
    Unload frmQuest
End Sub

Private Sub lblChoice_Click()
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
