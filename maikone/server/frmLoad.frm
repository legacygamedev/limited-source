VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maikone Server"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgLoad 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Status"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.PictureBox Picture1 
         Height          =   265
         Left            =   240
         ScaleHeight     =   210
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   360
         Width           =   4215
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            Caption         =   "Loading server..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   4215
         End
      End
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
