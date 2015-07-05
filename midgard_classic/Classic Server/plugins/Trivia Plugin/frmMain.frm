VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trivia Plugin By Pc"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrNewGame 
      Interval        =   1000
      Left            =   2400
      Top             =   120
   End
   Begin VB.Timer tmrTrivia 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1920
      Top             =   120
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Trivia"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2775
   End
   Begin VB.Frame fraIndex 
      Caption         =   "Bot Index"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2775
      Begin VB.Label lblIndex 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdWarp 
      Caption         =   "Warp"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRun_Click()
    Call SendEmoteMSG("Loading Trivia Questions.")
    Call SetupTrivia
End Sub

Private Sub cmdWarp_Click()
    Call Codes.MoveBotTo(BotIndex, BotMap, BotX, BotY, BotDir)
End Sub

Private Sub tmrNewGame_Timer()
    SecondsTill = SecondsTill - 1

    If SecondsTill = 0 Then
        Call SendEmoteMSG("New game starting in 5 seconds.")
        tmrTrivia.Enabled = True
        tmrNewGame.Enabled = False
    End If
End Sub

Private Sub tmrTrivia_Timer()
    
    SecondsLeft = SecondsLeft - 5
    
    If SecondsLeft = 0 Then
        Call SendEmoteMSG("Round Over. The Answer Was " & CurrentAnswer)
        Call NewGame
        tmrTrivia.Enabled = False
        Exit Sub
    End If
    
    Call SendEmoteMSG("Question: " & CurrentQuestion)
    Call SendEmoteMSG("Time Left: " & SecondsLeft)
    Call SendEmoteMSG("Help: Just type the answer in lowercase and send.")
    
    DoEvents
End Sub
