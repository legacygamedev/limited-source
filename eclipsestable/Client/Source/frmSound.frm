VERSION 5.00
Begin VB.Form frmSound 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sound Attribute"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4110
   ControlBox      =   0   'False
   Icon            =   "frmSound.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Choose Sound"
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
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
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   3600
         Width           =   1815
      End
      Begin VB.ListBox lstSound 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    SoundFileName = lstSound.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Call ListSounds(App.Path & "\SFX\", 1)
    lstSound.Text = SoundFileName
End Sub

Private Sub lstSound_Click()
    Call PlaySound(lstSound.Text)
End Sub
