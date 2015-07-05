VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Main Menu"
   ClientHeight    =   3750
   ClientLeft      =   150
   ClientTop       =   -30
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0442
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox boxNewChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   480
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   29
      Top             =   1440
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtCharName 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   225
         Left            =   0
         MaxLength       =   20
         TabIndex        =   35
         Top             =   240
         Width           =   2760
      End
      Begin VB.OptionButton optMale 
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   600
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton optFemale 
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1080
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   32
         Top             =   480
         Width           =   480
         Begin VB.PictureBox picPic2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   40
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   1080
         TabIndex        =   39
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   360
         TabIndex        =   38
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   360
         TabIndex        =   37
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Character Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   1485
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   4440
         TabIndex        =   31
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   3840
         TabIndex        =   30
         Top             =   960
         Width           =   420
      End
   End
   Begin VB.PictureBox boxChars 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   480
      ScaleHeight     =   1215
      ScaleWidth      =   5055
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ListBox lstChars 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   570
         ItemData        =   "frmMainMenu.frx":71C4
         Left            =   120
         List            =   "frmMainMenu.frx":71C6
         TabIndex        =   24
         Top             =   240
         Width           =   4785
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Logout"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   0
         TabIndex        =   28
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Characters:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   1020
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   3120
         TabIndex        =   26
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   3960
         TabIndex        =   25
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Use"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   960
         Width           =   315
      End
   End
   Begin VB.PictureBox boxNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   480
      ScaleHeight     =   1215
      ScaleWidth      =   5055
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtNewPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   0
         MaxLength       =   20
         TabIndex        =   17
         Top             =   720
         Width           =   2355
      End
      Begin VB.TextBox txtNewName 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   225
         Left            =   0
         MaxLength       =   20
         TabIndex        =   16
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "New Account:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   0
         TabIndex        =   19
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   4440
         TabIndex        =   18
         Top             =   960
         Width           =   585
      End
   End
   Begin VB.PictureBox boxLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   480
      ScaleHeight     =   1215
      ScaleWidth      =   5055
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   225
         Left            =   0
         MaxLength       =   20
         TabIndex        =   11
         Top             =   240
         Width           =   2355
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   0
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   720
         Width           =   2355
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Save Password"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label picConnect 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Login "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   4560
         TabIndex        =   14
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "Account:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   750
      End
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Config"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   5400
      TabIndex        =   41
      Top             =   3540
      Width           =   555
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3375
      Width           =   5535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapping By DarkAngel And BNK-Devilz"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   533
      TabIndex        =   7
      Top             =   2400
      Width           =   4935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Maps Inspired By Deloria"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   533
      TabIndex        =   6
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GFX Owned By Deloria"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   533
      TabIndex        =   5
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Deloria Unofficial Made By BNK-Devilz ""Sean"""
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   533
      TabIndex        =   4
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Deloria Made By Nexemis ""Jeff O'Blenis"""
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   533
      TabIndex        =   3
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label picLogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   2970
      Width           =   1380
   End
   Begin VB.Label picNewAccount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Account"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   2370
      TabIndex        =   1
      Top             =   2970
      Width           =   1380
   End
   Begin VB.Label picQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   900
      TabIndex        =   0
      Top             =   2970
      Width           =   1380
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public animi As Long

Private Sub picAddChar_Click()
Dim Msg As String
Dim i As Long

    If Trim(txtName.Text) <> "" Then
        Msg = Trim(txtName.Text)
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                txtName.Text = ""
                Exit Sub
            End If
        Next i
        
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub Label20_Click()
Dim Msg As String
Dim i As Long, d As Long

    d = 0
    For i = 1 To Len(txtCharName.Text)
        If Mid(txtCharName.Text, i, 1) = " " Then
            d = d + 1
        End If
    Next i
    If d >= 1 Then
        MsgBox "No spaces allowed!"
        Exit Sub
    End If

    If Trim(txtCharName.Text) <> "" Then
        Msg = Trim(txtCharName.Text)
        
        If Len(Trim(txtCharName.Text)) < 3 Then
            MsgBox "Character name must be at least three characters in length."
            Exit Sub
        End If
                
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub Label21_Click()
    If optMale.Value = True Then
        If NCPic = 266 Then
            NCPic = 268
        ElseIf NCPic = 268 Then
            NCPic = 270
        ElseIf NCPic = 270 Then
            NCPic = 277
        ElseIf NCPic = 277 Then
            NCPic = 295
        ElseIf NCPic = 295 Then
            NCPic = 266
        End If
    Else
        If NCPic = 267 Then
            NCPic = 269
        ElseIf NCPic = 269 Then
            NCPic = 273
        ElseIf NCPic = 273 Then
            NCPic = 275
        ElseIf NCPic = 275 Then
            NCPic = 284
        ElseIf NCPic = 284 Then
            NCPic = 267
        End If
    End If
End Sub

Private Sub Label22_Click()
    frmInput.Show vbModal
End Sub

Private Sub optFemale_Click()
    If optFemale.Value = True Then NCPic = 267
End Sub

Private Sub optMale_Click()
    If optMale.Value = True Then NCPic = 266
End Sub

Private Sub Timer1_Timer()
If boxNewChar.Visible = False Then Exit Sub
    picPic2.Left = animi * -PIC_X
    picPic2.Top = NCPic * -PIC_Y
End Sub

Private Sub Form_Load()
    picPic2.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
End Sub

Private Sub Timer2_Timer()
    animi = animi + 1
If animi > 4 Then
    animi = 3
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(frmMainMenu, Button, Shift, x, y)
End Sub

Private Sub Label11_Click()
    If lstChars.List(lstChars.ListIndex) <> ">Free Slot" Then Exit Sub
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub Label12_Click()
Dim Value As Long

    If lstChars.List(lstChars.ListIndex) = ">Free Slot" Then Exit Sub
    Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If Value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

Private Sub Label13_Click()
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Private Sub Label15_Click()
    Call TcpDestroy
    txtName.SetFocus
    txtName.SelStart = Len(txtName.Text)
    boxChars.Visible = False
    picLogin.Enabled = True
    picNewAccount.Enabled = True
End Sub

Private Sub Label16_Click()
    boxNewChar.Visible = False
    boxChars.Visible = True
End Sub

Private Sub Label6_Click()
Dim Msg As String
Dim i As Long
    
    If Trim(txtNewName.Text) <> "" And Trim(txtNewPass.Text) <> "" Then
        Msg = Trim(txtNewName.Text)
                
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                Exit Sub
            End If
        Next i
    
        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

Private Sub picConnect_Click()
    If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
        Call MenuState(MENU_STATE_LOGIN)
        Call WriteINI("CONFIG", "Account", txtName.Text, (App.Path & "\config.ini"))
        If Check1.Value = Checked Then
            Call WriteINI("CONFIG", "Password", txtPassword.Text, (App.Path & "\config.ini"))
        Else
            Call WriteINI("CONFIG", "Password", "", (App.Path & "\config.ini"))
        End If
    End If
End Sub

Private Sub picNewAccount_Click()
    boxLogin.Visible = False
    boxNew.Visible = True
    
    txtNewName.SetFocus
    txtNewName.SelStart = Len(txtNewName.Text)
End Sub

Private Sub picLogin_Click()
    txtName.Text = Trim(ReadINI("CONFIG", "Account", App.Path & "\config.ini"))
    txtPassword.Text = Trim(ReadINI("CONFIG", "Password", App.Path & "\config.ini"))
    boxNew.Visible = False
    boxLogin.Visible = True
    If Trim(txtPassword.Text) <> "" Then
        Check1.Value = Checked
    Else
        Check1.Value = Unchecked
    End If
    
    txtName.SetFocus
    txtName.SelStart = Len(txtName.Text)
End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub

