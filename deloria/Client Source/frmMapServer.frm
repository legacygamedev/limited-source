VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMapServer 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Map Server"
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "frmMapServer.frx":0000
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox lstDone 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMapServer.frx":20F9A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Left            =   1905
      TabIndex        =   2
      Top             =   1905
      Width           =   1140
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Done"
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
      Left            =   3090
      TabIndex        =   1
      Top             =   1905
      Width           =   1140
   End
End
Attribute VB_Name = "frmMapServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MoveForm(frmMapServer, Button, Shift, X, Y)
End Sub

Private Sub Label1_Click()
    Call SendMapData("cancel" & SEP_CHAR & END_CHAR)
    frmMainMenu.Visible = True
    frmMainMenu.picQuit.Enabled = True
    frmMainMenu.picNewAccount.Enabled = True
    frmMainMenu.picLogin.Enabled = True
    Me.Visible = False
End Sub

Private Sub lblButton_Click()
    frmMainMenu.Visible = True
    frmMainMenu.picQuit.Enabled = True
    frmMainMenu.picNewAccount.Enabled = True
    frmMainMenu.picLogin.Enabled = True
    Me.Visible = False
End Sub
