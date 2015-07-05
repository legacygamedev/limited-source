VERSION 5.00
Object = "{96366485-4AD2-4BC8-AFBF-B1FC132616A5}#2.0#0"; "VBMP.ocx"
Begin VB.Form frmCredits 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Engine Credits"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picCredits 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1785
      ScaleWidth      =   5745
      TabIndex        =   1
      Top             =   1320
      Width           =   5775
      Begin VB.PictureBox picCreditScroll 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8535
         Left            =   240
         ScaleHeight     =   8535
         ScaleWidth      =   5295
         TabIndex        =   2
         Top             =   2160
         Width           =   5295
         Begin VBMP.VBMPlayer VBMPlayer1 
            Height          =   1095
            Left            =   1680
            Top             =   6960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1931
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "VBMP (c)2007 DXGames"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1680
            TabIndex        =   29
            Top             =   8040
            Width           =   1935
         End
         Begin VB.Label lblVerrigan 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Verrigan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   6240
            Width           =   4965
         End
         Begin VB.Label lblMalikona 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Malikona"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   6000
            Width           =   4965
         End
         Begin VB.Label lblCoke 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Coke"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   6480
            Width           =   4965
         End
         Begin VB.Label lblGodSentDeath 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "GodSentDeath"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   5040
            Width           =   4965
         End
         Begin VB.Label lblGlobe 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Globe"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   5520
            Width           =   4965
         End
         Begin VB.Label lblShannara 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Shannara"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   5280
            Width           =   4965
         End
         Begin VB.Label lblIceCream 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Ice Cream Tuesday"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   5760
            Width           =   4965
         End
         Begin VB.Label lblAlanSpike 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "AlanSpike"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   4080
            Width           =   4965
         End
         Begin VB.Label lblPingu 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Pingu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   4320
            Width           =   4965
         End
         Begin VB.Label lblRafiki 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Rafiki"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   4560
            Width           =   4965
         End
         Begin VB.Label lblFrogImmortal 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "FrogImmortal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   4800
            Width           =   4965
         End
         Begin VB.Label lblContributors 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Eclipse - Contributors"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   3840
            Width           =   4965
         End
         Begin VB.Label lblMellowz 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Mellowz"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   3000
            Width           =   4965
         End
         Begin VB.Label lblEmblem 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Emblem"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   3240
            Width           =   4965
         End
         Begin VB.Label lblBraydo25 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Braydo25"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   3480
            Width           =   4965
         End
         Begin VB.Label lblBaron 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Baron"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2040
            Width           =   4965
         End
         Begin VB.Label lblUnnown 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Unnown"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   2280
            Width           =   4965
         End
         Begin VB.Label lblPickle 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Pickle"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2520
            Width           =   4965
         End
         Begin VB.Label lblDemonX 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Demon X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   2760
            Width           =   4965
         End
         Begin VB.Label lblHeyThereJake 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "HeyThereJake"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1560
            Width           =   4965
         End
         Begin VB.Label lblBrutal 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Brutal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1800
            Width           =   4965
         End
         Begin VB.Label lblTopher 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Topher"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   4965
         End
         Begin VB.Label lblYellowMole 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "The Yellow Mole"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   4965
         End
         Begin VB.Label lblMarsh 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Marsh"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   4965
         End
         Begin VB.Label lblDevelopers 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Eclipse - Developers"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   4965
         End
         Begin VB.Label lblDescription 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "This game was built using Eclipse Evolution  http://freemmorpgmaker.com"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   4965
         End
      End
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   5520
      Width           =   5520
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Long
    Dim Ending As String

    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        If FileExists("GUI\Credits" & Ending) Then
            frmCredits.Picture = LoadPicture(App.Path & "\GUI\Credits" & Ending)
        End If
    Next i

    Call VBMPlayer1.SetColors(QBColor(RED), QBColor(WHITE))

    picCreditScroll.Top = 2160

    tmrScroll.Enabled = True
End Sub

Private Sub picCancel_Click()
    tmrScroll.Enabled = False

    frmCredits.Visible = False
    frmMainMenu.Visible = True
    
    Unload Me
End Sub

Private Sub tmrScroll_Timer()
    If picCreditScroll.Top <= -10000 Then
        picCreditScroll.Top = 2160
    Else
        picCreditScroll.Top = picCreditScroll.Top - 25
    End If
End Sub
