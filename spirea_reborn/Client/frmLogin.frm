VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mirage Online (Login)"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   4560
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox RPass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Remember id/pw"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4680
      TabIndex        =   5
      Top             =   3240
      Width           =   200
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   390
      Left            =   4920
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   4920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   300
      Left            =   4800
      TabIndex        =   8
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   300
      Left            =   4800
      TabIndex        =   7
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remember Id/Password"
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   4920
      TabIndex        =   6
      Top             =   3240
      Width           =   1710
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a account name and password.  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public f As Integer

Private Sub Form_Load()
f = FreeFile

Dim n
Dim O
Dim usr
Dim t
Dim pass
On Error GoTo Error
Open "Data.ini" For Input As #f

Input #f, n
Input #f, O
Input #f, usr
Input #f, t
Input #f, pass
Close #f
RPass.Value = 1
txtName.Text = usr
txtPassword.Text = pass
Error:
End Sub

Private Sub Label5_Click()
If RPass.Value = 1 Then
Open "Data.ini" For Output As #f
Print #f, "[SETTINGS]"
Print #f, "Last User= "
Print #f, txtName
Print #f, "Last Password= "
Print #f, txtPassword
Close #f
Else
End If
If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
        Call MenuState(MENU_STATE_LOGIN)
    End If
End Sub

Private Sub Label6_Click()
    frmMainMenu.Visible = True
    frmLogin.Visible = False
End Sub

Private Sub picCancel_Click()

End Sub

Private Sub picConnect_Click()

End Sub



