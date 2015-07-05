VERSION 5.00
Begin VB.Form frmNotice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notice Attribute"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4830
   ControlBox      =   0   'False
   Icon            =   "frmNotice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Set Notice"
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.TextBox Text 
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox Title 
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   4
         Top             =   480
         Width           =   4335
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
         Left            =   2640
         TabIndex        =   3
         Top             =   4560
         Width           =   1935
      End
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
         Left            =   2640
         TabIndex        =   2
         Top             =   4200
         Width           =   1935
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
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    NoticeTitle = Title.Text
    NoticeText = Text.Text
    NoticeSound = lstSound.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Call ListSounds(App.Path & "\SFX\", 2)

    Title.Text = NoticeTitle
    Text.Text = NoticeText
    lstSound.Text = NoticeSound
End Sub

Private Sub lstSound_Click()
    Call PlaySound(lstSound.Text)
End Sub
