VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Selection"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtfocus 
      Height          =   390
      Left            =   735
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   6255
      Width           =   465
   End
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1290
      ItemData        =   "frmChars.frx":000C
      Left            =   120
      List            =   "frmChars.frx":000E
      TabIndex        =   0
      Top             =   120
      Width           =   2400
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2640
      TabIndex        =   4
      Top             =   1190
      Width           =   1335
   End
   Begin VB.Label picDelChar 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Delete Char"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label picNewChar 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "New Char"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label picUseChar 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Use Char"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    frmSendGetData.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub lstChars_GotFocus()
    txtFocus.SetFocus
End Sub

Private Sub lstChars_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call picUseChar_Click
        KeyAscii = 0
    End If
End Sub

Private Sub picCancel_Click()
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub picNewChar_Click()
    frmNewChar.txtName.Text = ""
    frmChars.Visible = False
    Call SetStatus("Connected, getting available classes...")
    Call SendGetClasses
End Sub

Private Sub picUseChar_Click()
    'save the last used
    WriteIniValue App.Path & "\Core Files\Configuration.ini", "Account Information", "LastChar", frmChars.lstChars.ListIndex
    
    frmChars.Visible = False
    If ConnectToServer = True Then
        Call SetStatus("Connected, sending char data...")
        Call SendUseChar(frmChars.lstChars.ListIndex + 1)
    End If
End Sub

Private Sub picDelChar_Click()
    Me.Visible = False
    frmDeleteCharacter.Visible = True
End Sub

