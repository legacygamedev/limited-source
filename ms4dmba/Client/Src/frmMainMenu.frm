VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4515
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7515
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
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   4800
      TabIndex        =   5
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblDeleteAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   4800
      TabIndex        =   4
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label lblNewAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   4800
      TabIndex        =   3
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   4800
      TabIndex        =   1
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6480
      TabIndex        =   2
      Top             =   4080
      Width           =   795
   End
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Private Sub Form_Load()
Dim rec As DXVBLib.RECT

    Me.Caption = GAME_NAME
    Me.Picture = LoadPicture(App.Path & "/gfx/interface/Menu.bmp")
    
    ' Allow DirectX to load in background
    Me.Show
    DoEvents

    ' initialize DirectX in the background after the form appears
    If Not InitDirectDraw Then
        MsgBox "Error Initializing DirectX7 - DirectDraw."
        DestroyGame
    End If
        
    '  sets the backbuffer dimensions to picScreen
'    frmMirage.picScreen.Width = DDSD_BackBuffer.lWidth
'    frmMirage.picScreen.Height = DDSD_BackBuffer.lHeight
    
    Call InitDirectMusic
    Call InitDirectSound
    
    'Call DirectMusic_PlayMidi("main.mid")

End Sub

Private Sub lblLogin_Click()
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub lblNewAccount_Click()
    frmNewAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub lblDeleteAccount_Click()
     If MsgBox("You are on the path for a character deletion, are you sure you want to go through with this?", vbYesNo, GAME_NAME) = vbYes Then
        frmDeleteAccount.Visible = True
        Me.Visible = False
    End If
End Sub

Private Sub lblCredits_Click()
    frmCredits.Visible = True
    Me.Visible = False
End Sub

Private Sub lblCancel_Click()
    Call DestroyGame
End Sub

