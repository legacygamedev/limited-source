VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMainMenu 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":08CA
   ScaleHeight     =   5535.041
   ScaleMode       =   0  'User
   ScaleWidth      =   3840.909
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox mnuCredits 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5550
      Left            =   0
      Picture         =   "frmMainMenu.frx":47066
      ScaleHeight     =   5550
      ScaleWidth      =   3900
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3900
      Begin RichTextLib.RichTextBox txtSpecial 
         Height          =   975
         Left            =   1080
         TabIndex        =   27
         ToolTipText     =   "Special Contributers to Playerworlds"
         Top             =   3960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1720
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMainMenu.frx":8D802
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblProgramming2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Lizzy Rognile"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   30
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Image imgCredits 
         Height          =   255
         Left            =   1560
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label lblProgramming 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "James Will"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   29
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label lblGUI1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "James Will"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1520
         TabIndex        =   28
         Top             =   3240
         Width           =   915
      End
   End
   Begin VB.PictureBox mnuNewAccount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5550
      Left            =   0
      Picture         =   "frmMainMenu.frx":8D87E
      ScaleHeight     =   5550
      ScaleWidth      =   3900
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3900
      Begin VB.TextBox txtNewAcctPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   600
         MaxLength       =   20
         PasswordChar    =   "•"
         TabIndex        =   2
         Top             =   4320
         Width           =   2415
      End
      Begin VB.TextBox txtNewAcctName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   240
         Left            =   600
         MaxLength       =   20
         TabIndex        =   1
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Image imgNewAcct 
         Height          =   255
         Index           =   1
         Left            =   2760
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image imgNewAcct 
         Height          =   255
         Index           =   0
         Left            =   360
         Top             =   5040
         Width           =   1215
      End
   End
   Begin VB.PictureBox mnuChars 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5550
      Left            =   0
      Picture         =   "frmMainMenu.frx":D401A
      ScaleHeight     =   5550
      ScaleWidth      =   3900
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3900
      Begin VB.PictureBox picSelChar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1680
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3720
         Width           =   495
      End
      Begin VB.ListBox lstChars 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   870
         ItemData        =   "frmMainMenu.frx":11A7B6
         Left            =   510
         List            =   "frmMainMenu.frx":11A7B8
         TabIndex        =   8
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Image imgChars 
         Height          =   255
         Index           =   3
         Left            =   2400
         Top             =   5040
         Width           =   735
      End
      Begin VB.Image imgChars 
         Height          =   255
         Index           =   2
         Left            =   720
         Top             =   5040
         Width           =   975
      End
      Begin VB.Image imgChars 
         Height          =   255
         Index           =   1
         Left            =   2280
         Top             =   4680
         Width           =   735
      End
      Begin VB.Image imgChars 
         Height          =   255
         Index           =   0
         Left            =   960
         Top             =   4680
         Width           =   495
      End
   End
   Begin VB.PictureBox mnuLogin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5550
      Left            =   0
      Picture         =   "frmMainMenu.frx":11A7BA
      ScaleHeight     =   5550
      ScaleWidth      =   3900
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3900
      Begin VB.TextBox txtLoginName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   240
         Left            =   480
         MaxLength       =   20
         TabIndex        =   6
         Top             =   3810
         Width           =   1455
      End
      Begin VB.TextBox txtLoginPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   480
         MaxLength       =   20
         PasswordChar    =   "•"
         TabIndex        =   5
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CheckBox chkLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   600
         TabIndex        =   4
         Top             =   2950
         Width           =   200
      End
      Begin VB.Image imgLogin 
         Height          =   255
         Index           =   1
         Left            =   2640
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image imgLogin 
         Height          =   255
         Index           =   0
         Left            =   480
         Top             =   5040
         Width           =   1215
      End
   End
   Begin VB.PictureBox mnuNewCharacter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5550
      Left            =   0
      Picture         =   "frmMainMenu.frx":160F56
      ScaleHeight     =   5550
      ScaleWidth      =   3900
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3900
      Begin VB.PictureBox picPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   3000
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2280
         Width           =   480
      End
      Begin VB.OptionButton optFemale 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         Picture         =   "frmMainMenu.frx":1A76F2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3480
         Width           =   855
      End
      Begin VB.OptionButton optMale 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         Picture         =   "frmMainMenu.frx":1AABBE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3480
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.ComboBox cmbClass 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmMainMenu.frx":1AE08A
         Left            =   600
         List            =   "frmMainMenu.frx":1AE08C
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtNewCharName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   600
         MaxLength       =   20
         TabIndex        =   11
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Image imgNewChar 
         Height          =   255
         Index           =   1
         Left            =   2400
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image imgNewChar 
         Height          =   255
         Index           =   0
         Left            =   720
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label lblMAGI 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblDEF 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label lblSP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblSPEED 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label lblMP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label lblSTR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label lblHP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   3840
         Width           =   375
      End
   End
   Begin VB.PictureBox mnuIPConfig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5550
      Left            =   0
      Picture         =   "frmMainMenu.frx":1AE08E
      ScaleHeight     =   5565.041
      ScaleMode       =   0  'User
      ScaleWidth      =   3855.513
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3900
      Begin VB.TextBox txtIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   225
         Left            =   600
         MaxLength       =   20
         TabIndex        =   25
         Top             =   3510
         Width           =   2775
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   600
         MaxLength       =   20
         TabIndex        =   24
         Top             =   4200
         Width           =   2775
      End
      Begin VB.Image imgIPConfig 
         Height          =   255
         Index           =   1
         Left            =   2400
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image imgIPConfig 
         Height          =   255
         Index           =   0
         Left            =   600
         Top             =   5040
         Width           =   735
      End
   End
   Begin VB.Image imgMainMenu 
      Height          =   375
      Index           =   5
      Left            =   480
      Top             =   840
      Width           =   2775
   End
   Begin VB.Image imgMainMenu 
      Height          =   375
      Index           =   4
      Left            =   960
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Image imgMainMenu 
      Height          =   375
      Index           =   3
      Left            =   600
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Image imgMainMenu 
      Height          =   375
      Index           =   2
      Left            =   1080
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Image imgMainMenu 
      Height          =   375
      Index           =   1
      Left            =   600
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Image imgMainMenu 
      Height          =   375
      Index           =   0
      Left            =   720
      Top             =   2280
      Width           =   2535
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ********************************************
' **               rootSource               **
' ********************************************

Private Sub Form_Load()
Dim rec As DXVBLib.RECT
Dim FileName As String

    'Me.Caption = GAME_NAME
    
    ' Allow DirectX to load in background
    Me.Show
    DoEvents

    ' initialize DirectX in the background after the form appears
    If Not InitDirectDraw Then
        MsgBox "Error Initializing DirectX7 - DirectDraw."
        DestroyGame
    End If
        
    '  sets the backbuffer dimensions to picScreen
    frmMainGame.picScreen.Width = DDSD_BackBuffer.lWidth
    frmMainGame.picScreen.Height = DDSD_BackBuffer.lHeight
    
    Call InitDirectMusic
    Call InitDirectSound
    
    

    
    'Call DirectMusic_PlayMidi("main.mid")
    
    If App.PrevInstance = True Then
        MsgBox "Another Playerworlds Client is already running! Please run only one client at a time!", Error
    End If

    FileName = App.Path & DATA_PATH & "config.dat"
    txtIP.Text = Trim$(GameData.IP)       ' GetVar(FileName, "IPCONFIG", "IP")
    txtPort.Text = Trim$(GameData.Port)   ' GetVar(FileName, "IPCONFIG", "PORT")
    
    ' Used for Credits
    Dim result As Long
    With txtSpecial
        result = SetWindowLong(.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
        .SelColor = QBColor(White)
        .SelAlignment = 2
        .SelText = "Jacob Burton" & vbNewLine & "Robin Perris" & vbNewLine & "Dmitry Bromberg" & vbNewLine & "Jon Petros" & vbNewLine & "Liam Stewart" & vbNewLine & "Chris Kremer" & vbNewLine & "Mr. Shannara" & vbNewLine & "OGC Community" & vbNewLine & "PW Community" & vbNewLine & "MS Community"
    End With
    
End Sub

'**********************************
'* Handles Character Menu Buttons *
'**********************************
Private Sub imgChars_Click(Index As Integer)
Dim Value As Long

Select Case Index

    Case 0
        Call MenuState(MENU_STATE_USECHAR)
    
        Exit Sub
    Case 1
        Call MenuState(MENU_STATE_NEWCHAR)
    
        Exit Sub
    Case 2
        Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
        If Value = vbYes Then
            Call MenuState(MENU_STATE_DELCHAR)
        End If
    
        Exit Sub
    Case 3
        Call DestroyTCP
        mnuLogin.Visible = True
        mnuChars.Visible = False
    
        Exit Sub
End Select
End Sub

'**************************
'* Handles Credits Button *
'**************************
Private Sub imgCredits_Click()
    mnuCredits.Visible = False
End Sub

'************************************
'* Handles IP Configuration Buttons *
'************************************
Private Sub imgIPConfig_Click(Index As Integer)
Dim IP, Port As String
Dim FileName As String
Dim fErr As Integer
Dim Texto As String

Select Case Index

    Case 0
        IP = Trim$(txtIP.Text)
        Port = Val(Trim$(txtPort.Text))
        FileName = App.Path & DATA_PATH & "config.dat"
    
        fErr = 0
        If fErr = 0 And Len(Trim$(IP)) = 0 Then
            fErr = 1
            Call MsgBox("Inform a correct IP.", vbCritical, GAME_NAME)
            Exit Sub
        End If
        If fErr = 0 And Port <= 0 Then
            fErr = 1
            Call MsgBox("Inform a correct Port.", vbCritical, GAME_NAME)
            Exit Sub
        End If
        If fErr = 0 Then
            GameData.IP = txtIP.Text
            GameData.Port = txtPort.Text
            Dim F  As Long
            F = FreeFile
            Open FileName For Binary As #F
            Put #F, , GameData
            Close #F
        End If
        mnuIPConfig.Visible = False
        Call DestroyTCP
        Call TcpInit
    
        Exit Sub
    Case 1
        mnuIPConfig.Visible = False
    
        Exit Sub
End Select

End Sub

'*************************
'* Handles Login Buttons *
'*************************
Private Sub imgLogin_Click(Index As Integer)
Dim FileName As String

Select Case Index

    Case 0
        FileName = App.Path & DATA_PATH & "config.dat"
    
        If chkLogin.Value Then
            GameData.SaveLogin = 1
            GameData.Username = txtLoginName.Text
            GameData.Password = Trim$(txtLoginPassword.Text)
        Else
            GameData.SaveLogin = 0
            GameData.Username = vbNullString
            GameData.Password = vbNullString
        End If
    
        Dim F As Long
        F = FreeFile
        Open FileName For Binary As #F
        Put #F, , GameData
        Close #F
    
        Call LoginConnect
    
        Exit Sub
    Case 1
        mnuLogin.Visible = False
    
        Exit Sub
End Select

End Sub

'*****************************
'* Handles Main Menu Buttons *
'*****************************
Private Sub imgMainMenu_Click(Index As Integer)
Select Case Index

    Case 0
        mnuNewAccount.Visible = True
        txtNewAcctName.SetFocus
        
        Exit Sub
    Case 1
        If GameData.SaveLogin = 1 Then
            chkLogin.Value = 1
            txtLoginName.Text = Trim$(GameData.Username)
            txtLoginPassword.Text = Trim$(GameData.Password)
        End If
        
        mnuLogin.Visible = True
        txtLoginName.SetFocus
        
        Exit Sub
    Case 2
        txtIP.Text = Trim$(GameData.IP)
        txtPort.Text = Trim$(GameData.Port)
        mnuIPConfig.Visible = True

        txtIP.SetFocus
        
        Exit Sub
    Case 3
        ' Game Options Here
            
        Exit Sub
    Case 4
        Call DestroyGame
    
        Exit Sub
    Case 5
        mnuCredits.Visible = True
        
        Exit Sub
End Select

End Sub

'************************************
'* Handles New Account Menu Buttons *
'************************************
Private Sub imgNewAcct_Click(Index As Integer)
Select Case Index

    Case 0
        Call NewAccountConnect
        
        Exit Sub
    Case 1
        mnuNewAccount.Visible = False
        
        Exit Sub
End Select

End Sub

'**************************************
'* Handles New Character Menu Buttons *
'**************************************
Private Sub imgNewChar_Click(Index As Integer)

Select Case Index

    Case 0
        Call AddCharClick
    
        Exit Sub
    Case 1
        mnuChars.Visible = True
        mnuNewCharacter.Visible = False
    
        Exit Sub
End Select

End Sub

Private Sub cmbClass_Change()
    'If Class(cmbClass.ListIndex).Sprite = Class(cmbClass.ListIndex).FSprite Then
    '    optMale.Value = True
    '    optMale.Visible = False
    '   optFemale.Value = False
    '    optFemale.Visible = False
    'ElseIf Class(cmbClass.ListIndex).Sprite <> Class(cmbClass.ListIndex).FSprite Then
    '    optMale.Value = True
    '    optMale.Visible = True
    '    optFemale.Value = False
    '    optFemale.Visible = True
    'End If
    
    DrawNewChar
    
End Sub

'**********************************
'* Handles Moving the Form Around *
'**********************************
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub lstChars_Click()
    DrawSelChar lstChars.ListIndex + 1
End Sub

Private Sub mnuChars_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub mnuCredits_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub mnuIPConfig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub mnuLogin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub mnuNewAccount_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub mnuNewCharacter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(frmMainMenu)
End Sub
