VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Endieko Server"
   ClientHeight    =   2895
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMain 
      Interval        =   1
      Left            =   7920
      Top             =   0
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   2655
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmLoad.frx":0000
      Top             =   0
      Width           =   8535
   End
   Begin VB.Label lblPlayersOnline 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current Players Online: None"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status: Loading Information..."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Show
    DoEvents
    Call InitServer
End Sub

