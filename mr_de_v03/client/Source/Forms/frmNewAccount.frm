VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mirage Realms"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2460
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2460
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
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
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1995
   End
   Begin VB.TextBox txtName 
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
      Height          =   240
      Left            =   240
      MaxLength       =   20
      TabIndex        =   2
      Top             =   480
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   885
      TabIndex        =   1
      Top             =   1520
      Width           =   1335
   End
   Begin VB.Label picConnect 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Create Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   885
      TabIndex        =   0
      Top             =   1245
      Width           =   1335
   End
End
Attribute VB_Name = "frmNewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    frmMainMenu.Visible = True
    frmNewAccount.Visible = False
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPassword.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Call picConnect_Click
         KeyAscii = 0
     End If
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmNewAccount.Visible = False
End Sub

Private Sub picConnect_Click()
Dim Msg As String
Dim i As Long

    If Trim$(txtName.Text) <> vbNullString And Trim$(txtPassword.Text) <> vbNullString Then
        Msg = Trim$(txtName.Text)
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                Call TcpDestroy
                frmEvent.Visible = True
                frmEvent.lblInformation.Caption = "Account name contains invalid characters. (Err: #4)"
                
                txtName.Text = vbNullString
                Exit Sub
            End If
        Next
    
        frmNewAccount.Visible = False
        If ConnectToServer = True Then
            Call SetStatus("Connected, sending new account information...")
            Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text)
        End If
    End If
End Sub

