VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMainMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mirage Online"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   4530
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtNews 
      Height          =   2775
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4895
      _Version        =   393217
      BackColor       =   4210752
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"frmMainMenu.frx":6F33A
   End
   Begin InetCtlsObjects.Inet InetNews 
      Left            =   0
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   4920
      TabIndex        =   6
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   4920
      TabIndex        =   5
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   4920
      TabIndex        =   4
      Top             =   2880
      Width           =   3105
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Del. Acc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   4920
      TabIndex        =   3
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Acc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Config"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   3600
      Width           =   3135
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim TmpStr As String

    TmpStr = InetNews.OpenURL("http://www.spirea.flphost.com/News.txt")

    txtNews.Text = TmpStr
End Sub
Private Sub picIpConfig_Click()

End Sub

Private Sub Label1_Click()
frmIpconfig.Visible = True
Me.Visible = False
End Sub

Private Sub Label2_Click()
Call GameDestroy
End Sub

Private Sub Label4_Click()
    frmNewAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub Label5_Click()
Dim YesNo As Long

    YesNo = MsgBox("You are on the path for a character deletion, are you sure you want to go through with this?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        frmDeleteAccount.Visible = True
        Me.Visible = False
    End If
End Sub

Private Sub Label6_Click()
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub Label7_Click()
    frmCredits.Visible = True
    Me.Visible = False
End Sub

Private Sub picNewAccount_Click()

End Sub

Private Sub picDeleteAccount_Click()

End Sub

Private Sub picLogin_Click()

End Sub

Private Sub picCredits_Click()

End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub

